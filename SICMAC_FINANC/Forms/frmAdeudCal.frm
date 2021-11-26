VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmAdeudCal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADEUDADOS: CALENDARIO DE PAGARES"
   ClientHeight    =   9180
   ClientLeft      =   1740
   ClientTop       =   1245
   ClientWidth     =   12015
   ClipControls    =   0   'False
   Icon            =   "frmAdeudCal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   900
      Left            =   120
      TabIndex        =   44
      Top             =   8280
      Width           =   8220
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6870
         TabIndex        =   56
         Top             =   465
         Width           =   1200
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
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
         Left            =   60
         TabIndex        =   50
         Top             =   465
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1290
         TabIndex        =   49
         Top             =   465
         Width           =   1200
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5625
         TabIndex        =   48
         Top             =   465
         Width           =   1200
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   2535
         TabIndex        =   47
         Top             =   465
         Width           =   1200
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Im&portar"
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
         Left            =   4365
         TabIndex        =   46
         Top             =   465
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "El Orden del Calendario Mi Vivienda son 1. Cuotas No Concesionales 2. Cuotas."
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   105
         TabIndex        =   45
         Top             =   195
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Adeudado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   11715
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   8955
         TabIndex        =   2
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8385
         TabIndex        =   29
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblCodAdeudado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   885
         TabIndex        =   0
         Top             =   225
         Width           =   2730
      End
      Begin VB.Label lblEntidad 
         AutoSize        =   -1  'True
         Caption         =   "Entidad :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   39
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblDescEntidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   885
         TabIndex        =   1
         Top             =   570
         Width           =   7335
      End
   End
   Begin VB.Frame fraIngAdeud 
      Caption         =   "Datos Principales"
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
      ForeColor       =   &H00000080&
      Height          =   2145
      Left            =   120
      TabIndex        =   33
      Top             =   1080
      Width           =   11760
      Begin VB.Frame Frame4 
         Caption         =   "Mi  Vivienda"
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
         Height          =   1935
         Left            =   8040
         TabIndex        =   52
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtConcesional 
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
            Left            =   120
            TabIndex        =   54
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CheckBox chkMiVivienda 
            Caption         =   "Mi vivienda"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Concesional :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   1170
         End
      End
      Begin VB.TextBox txtComision 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6720
         TabIndex        =   27
         Text            =   "0"
         Top             =   1710
         Width           =   690
      End
      Begin VB.Frame frmTpoCuota 
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
         Height          =   1305
         Left            =   4200
         TabIndex        =   17
         Top             =   420
         Width           =   1695
         Begin VB.OptionButton optTpoCuota 
            Caption         =   "Iterativo"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optTpoCuota 
            Caption         =   "Variable"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   630
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.TextBox txtCuota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   720
            TabIndex        =   19
            Text            =   "0"
            Top             =   210
            Width           =   870
         End
         Begin VB.OptionButton optTpoCuota 
            Caption         =   "Fija"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   765
         End
      End
      Begin VB.Frame frmPeriodo 
         Caption         =   "Periodo"
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
         Height          =   945
         Left            =   6000
         TabIndex        =   21
         Top             =   660
         Width           =   1965
         Begin VB.TextBox txtFechaCuota 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1290
            MaxLength       =   3
            TabIndex        =   25
            Text            =   "0"
            Top             =   540
            Width           =   540
         End
         Begin VB.OptionButton OptTpoPeriodo 
            Caption         =   "Fecha Fija"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   24
            Top             =   600
            Width           =   1125
         End
         Begin VB.OptionButton OptTpoPeriodo 
            Caption         =   "Periodo Fijo"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   270
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.TextBox txtPlazoCuotas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1290
            MaxLength       =   5
            TabIndex        =   23
            Text            =   "0"
            Top             =   180
            Width           =   540
         End
      End
      Begin VB.TextBox txtTramo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3210
         MaxLength       =   6
         TabIndex        =   16
         Text            =   "0"
         Top             =   1710
         Width           =   720
      End
      Begin Spinner.uSpinner SpnGracia 
         Height          =   360
         Left            =   945
         TabIndex        =   11
         Top             =   1320
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   635
         Max             =   9999
         MaxLength       =   4
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
      Begin Spinner.uSpinner SpnCuotas 
         Height          =   360
         Left            =   945
         TabIndex        =   10
         Top             =   915
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
         Max             =   999
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
      Begin VB.Frame FraTipoCambio 
         Caption         =   "T.C."
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
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   6000
         TabIndex        =   4
         Top             =   0
         Width           =   1965
         Begin VB.CheckBox chkVac 
            Caption         =   "VAC"
            Height          =   255
            Left            =   1260
            TabIndex        =   43
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtTipoCambio 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   150
            TabIndex        =   5
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.TextBox txtInteres 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   765
      End
      Begin VB.TextBox txtCapital 
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
         Height          =   315
         Left            =   945
         TabIndex        =   3
         Top             =   225
         Width           =   2010
      End
      Begin VB.CheckBox chkInterno 
         Caption         =   "&Plaza Interna"
         Height          =   255
         Left            =   2175
         TabIndex        =   12
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.TextBox txtCuotaPagoK 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   3210
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "0"
         Top             =   1320
         Width           =   720
      End
      Begin VB.CheckBox chkPagoK 
         Caption         =   "Pago K a :"
         Height          =   195
         Left            =   2175
         TabIndex        =   13
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   480
         Left            =   2085
         TabIndex        =   7
         Top             =   465
         Width           =   1980
         Begin VB.OptionButton optPeriodo 
            Caption         =   "&Mensual"
            Height          =   240
            Index           =   1
            Left            =   930
            TabIndex        =   9
            Top             =   165
            Width           =   915
         End
         Begin VB.OptionButton optPeriodo 
            Caption         =   "&Anual "
            Height          =   240
            Index           =   0
            Left            =   105
            TabIndex        =   8
            Top             =   165
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comisión por Cuota                    %"
         Height          =   195
         Left            =   5190
         TabIndex        =   26
         Top             =   1770
         Width           =   2385
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tramo No Concesional :                 %"
         Height          =   195
         Left            =   1500
         TabIndex        =   15
         Top             =   1800
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas :"
         Height          =   195
         Left            =   105
         TabIndex        =   37
         Top             =   975
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Interes :                        %"
         Height          =   195
         Left            =   105
         TabIndex        =   36
         Top             =   615
         Width           =   1770
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Per. Gracia                 Dias"
         Height          =   195
         Left            =   105
         TabIndex        =   35
         Top             =   1395
         Width           =   1875
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Capital :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   34
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Frame framPlanPagos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5010
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   11760
      Begin VB.CommandButton btnexportar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Exportar XLS"
         Height          =   375
         Left            =   10200
         TabIndex        =   72
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtPLargo 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8160
         TabIndex        =   67
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtPCorto 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6000
         TabIndex        =   66
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtDlargo 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3000
         TabIndex        =   64
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtDcorto 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   63
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtTotalcapitalC 
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
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   3720
         Width           =   1830
      End
      Begin VB.TextBox txtTotalInteresC 
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
         Height          =   315
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   3720
         Width           =   1440
      End
      Begin VB.CommandButton cmdNuevaCuota 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   8280
         TabIndex        =   51
         Top             =   3600
         Width           =   960
      End
      Begin Sicmact.FlexEdit fgCronograma 
         Height          =   2925
         Left            =   120
         TabIndex        =   42
         Top             =   300
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5159
         Cols0           =   18
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmAdeudCal.frx":000C
         EncabezadosAnchos=   "300-1200-600-1150-1000-1000-0-0-0-0-0-0-0-1150-1150-1200-1200-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-5-X-X-X-X-X-X-X-13-14-15-X-X"
         ListaControles  =   "0-2-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-L-C-C-C-C-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-2-2-2-2-2-0-0-0-0-0-2-2-2-2-2"
         CantEntero      =   14
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.TextBox txtTotalGeneral 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3720
         Width           =   1995
      End
      Begin VB.TextBox txtTotalInteres 
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
         Height          =   315
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3405
         Width           =   1440
      End
      Begin VB.TextBox txtTotalcapital 
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
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3405
         Width           =   1830
      End
      Begin VB.Label Label17 
         Caption         =   "LP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   71
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "CP"
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
         Left            =   5520
         TabIndex        =   70
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "LP"
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
         Left            =   2640
         TabIndex        =   69
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "CP"
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
         Left            =   600
         TabIndex        =   68
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Pendiente"
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
         Left            =   7320
         TabIndex        =   65
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Desembolso"
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
         Left            =   1800
         TabIndex        =   62
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Interes Con:"
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
         Left            =   3420
         TabIndex        =   60
         Top             =   3720
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Capital Con :"
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
         Left            =   195
         TabIndex        =   59
         Top             =   3720
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Interes :"
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
         Left            =   3420
         TabIndex        =   41
         Top             =   3450
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Capital :"
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
         Left            =   195
         TabIndex        =   40
         Top             =   3450
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   10680
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAdeudCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lnCapital As Currency
Public lbCargaIF As Boolean
Public lbMod As Boolean
Dim lbOk     As Boolean
Public gbGeneraCal As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim oDAdeud As DCaja_Adeudados
Dim oNAdeudCal As nAdeudCal
Dim rs         As ADODB.Recordset

Dim lsIFCod       As String
Dim lsIFDescrip   As String
Dim lnComisionIni As String
Dim lsTpoCta      As String
Dim lnInterno   As Integer
Dim lnCuotaCap  As Integer, ldFecContrato As Date
Dim lnNroCuotas As Integer
Dim lnPerGracia As Currency
Dim lbGetAdeud  As Boolean
Dim lnTasaInt   As Double
Dim lbSoloConsulta As Boolean
Dim fbAgregarCuota As Boolean 'EJVG20121205
Dim lnConcesionado As Currency 'ALPA20130614
Dim objPista As COMManejador.Pista 'ARLO20170217
Dim RutaXLSExport As String     'ANGC20210205
Dim Archivo As String

'Public Sub Inicio(pbCargaIF As Boolean, psIFCod As String, psIFDescrip As String, pnCapital As Currency, pdFecContrato As Date, Optional pnTasaInt As Double = 0, Optional pbGetAdeud As Boolean = False, Optional pbSoloConsulta As Boolean = False)
'Public Sub Inicio(pbCargaIF As Boolean, psIFCod As String, psIFDescrip As String, pnCapital As Currency, pdFecContrato As Date, Optional pnTasaInt As Double = 0, Optional pbGetAdeud As Boolean = False, Optional pbSoloConsulta As Boolean = False, Optional ByVal pbAgregarCuota As Boolean = False) 'EJVG20121205
Public Sub Inicio(pbCargaIF As Boolean, psIFCod As String, psIFDescrip As String, pnCapital As Currency, pdFecContrato As Date, Optional pnTasaInt As Double = 0, Optional pbGetAdeud As Boolean = False, Optional pbSoloConsulta As Boolean = False, Optional ByVal pbAgregarCuota As Boolean = False, Optional ByVal pnConcesionado As Currency = 0) 'ALPA20130614
    lbCargaIF = pbCargaIF
    lsIFCod = psIFCod
    lsIFDescrip = psIFDescrip
    lnCapital = pnCapital
    ldFecContrato = pdFecContrato
    lbGetAdeud = pbGetAdeud
    lnTasaInt = pnTasaInt
    lbSoloConsulta = pbSoloConsulta
    fbAgregarCuota = pbAgregarCuota 'EJVG20121205
    lnConcesionado = pnConcesionado 'ALPA20130614
    Me.Show 1
End Sub
    
Private Sub btnexportar_Click()
    Dim LibroTrabajo As Object
    
    If Trim(lblCodAdeudado.Caption) = "" Then
        Archivo = "SimulacionAdeudados-" & Format(gdFecSis, "YYYYMMDD")
    Else
        Archivo = lblCodAdeudado.Caption
    End If
    
    If Exportar_Excel(App.path & "\FormatoCarta\" & Archivo & ".xls", fgCronograma) Then
        MsgBox " Datos exportados! ", vbOKOnly, "MENSAJE"
        RutaXLSExport = App.path & "\FormatoCarta\" & Archivo & ".xls"

        Set LibroTrabajo = GetObject(RutaXLSExport) 'con el path correspondiente
        LibroTrabajo.Application.Windows(Archivo & ".xls").Visible = True
    End If
End Sub

Private Sub chkInterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkPagoK.SetFocus
    End If
End Sub

Private Sub chkMiVivienda_Click()
    If chkMiVivienda.value = 1 Then
        txtConcesional.Enabled = True
    Else
        txtConcesional.Enabled = False
        txtConcesional.Text = 0#
    End If
End Sub

Private Sub chkPagoK_Click()
    If Me.chkPagoK.value = 1 Then
        Me.txtCuotaPagoK.Enabled = True
    Else
        Me.txtCuotaPagoK.Enabled = False
        Me.txtCuotaPagoK = 0
    End If
End Sub
Private Sub chkPagoK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtCuotaPagoK.Enabled Then
            Me.txtCuotaPagoK.SetFocus
        Else
            Me.txtTramo.SetFocus
        End If
    End If
End Sub

Private Sub chkVac_Click()
If chkVac.value = vbChecked Then
    txtTipoCambio = oDAdeud.CargaIndiceVAC(txtFecha)
    If Val(txtTipoCambio) > 0 And Trim(Mid(lblCodAdeudado, 20, 1)) = "1" Then
        txtCapital = Format(lnCapital / Val(txtTipoCambio), "#,#0.00")
        txtConcesional = Format(lnConcesionado / Val(txtTipoCambio), "#,#0.00") 'ALPA20130614
    End If
    txtCapital.BackColor = vbGreen
    txtTotalGeneral.BackColor = vbGreen
    txtConcesional.BackColor = vbGreen 'ALPA20130614
Else
    txtConcesional = Format(lnConcesionado, "#,#0.00") 'ALPA20130614
    txtCapital = Format(lnCapital, "#,#0.00")
    txtCapital.BackColor = vbWhite
    txtTotalGeneral.BackColor = vbWhite
    txtConcesional.BackColor = vbWhite 'ALPA20130614
    Me.txtTipoCambio = "0.00"
End If
    If nVal(txtConcesional.Text) > 0 Then
        chkMiVivienda.value = 1
    End If
End Sub

Private Sub chkVac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtInteres.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
    If ValidaInterfaz Then
        If lbCargaIF Then
            Me.Hide
        Else
            Unload Me
        End If
        lbOk = True
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiaControles
    lbOk = False
End Sub

Private Sub cmdGenerar_Click()
Dim lnMontoCuota As Double
Dim lnTasaInt As Currency
Dim lVerCuotasCanceladas As Boolean
lVerCuotasCanceladas = False

    If fgCronograma.TextMatrix(1, 1) <> "" Then
        If fgCronograma.TextMatrix(1, 9) = gTpoEstCuotaAdeudCanc Then
           If MsgBox("Calendario ya posee Cuotas Canceladas. " & Chr(10) & _
                   "¿ Desea generar Calendario ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
              lVerCuotasCanceladas = True
           Else
               Exit Sub
           End If
        End If
    End If
    
    If ValidaInterfaz Then
        If optTpoCuota(0).value Then
            lnMontoCuota = nVal(txtCuota)
        End If
        If optTpoCuota(0) Then  ' Case gAdeudTpoCuotaFija
            oNAdeudCal.GeneraCalendarioCuotaFija fgCronograma, CCur(txtCapital), lnMontoCuota, CCur(txtInteres), _
                IIf(Me.optPeriodo(0).value = True, 360, 30), SpnCuotas.Valor, Me.txtPlazoCuotas, txtFecha, IIf(OptTpoPeriodo(0), PeriodoFijo, FechaFija), PrimeraCuota, _
                Val(SpnGracia.Valor), txtFechaCuota, False, , optTpoCuota(0).value And OptTpoPeriodo(1), nVal(txtComision), nVal(txtTramo), nVal(txtConcesional.Text), nVal(txtComision)
        End If
        If optTpoCuota(1) Then  ' Case gAdeudTpoCuotaVariable
            oNAdeudCal.GeneraCalendarioCuotaVariable fgCronograma, CCur(txtCapital), Val(SpnCuotas.Valor), Val(txtPlazoCuotas), _
             IIf(Me.optPeriodo(0).value = True, 360, 30), txtFecha, _
             Val(SpnGracia.Valor), CCur(txtInteres), Val(txtCuotaPagoK), nVal(txtComision), nVal(txtConcesional.Text), nVal(txtComision)
        End If
        If optTpoCuota(2) Then  'ANGC202012
            Dim Interes As Double
            If Me.optPeriodo(0).value = True Then   'TEA            CONVERTIR A TEA A TEM
                Interes = Round(txtInteres.Text / 100#, 8)
                Interes = Round(((1 + Interes) ^ (1 / 12)) - 1, 8)
            Else
                Interes = txtInteres.Text           'TEM
            End If
            oNAdeudCal.GenerarCronogramaIterativo fgCronograma, txtCapital.Text, SpnCuotas.Valor, Interes, SpnGracia.Valor, txtFecha.Text, txtComision.Text
        End If
        
        
        If lVerCuotasCanceladas Then
            VerCuotasCanceladas
        End If
        Me.txtCuota = Format(lnMontoCuota, gsFormatoNumeroView)
        
        txtTotalGeneral = Format(oNAdeudCal.TotalGeneral, gsFormatoNumeroView)
        txtTotalInteres = Format(oNAdeudCal.TotalInteres, gsFormatoNumeroView)
        txtTotalcapital = Format(oNAdeudCal.TotalCapital, gsFormatoNumeroView)
        
        txtTotalInteresC = Format(oNAdeudCal.TotalInteresC, gsFormatoNumeroView)
        txtTotalcapitalC = Format(oNAdeudCal.TotalCapitalC, gsFormatoNumeroView)
        
        Sumatoria
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Genero Calendario "
        Set objPista = Nothing
        '****

    End If
End Sub

Public Sub Sumatoria()
    Dim Inicio, i As Integer
    Inicio = 1
    Dim dCorto, dLargo, PCorto, PLargo As Double
    dCorto = 0
    dLargo = 0
    PCorto = 0
    PLargo = 0
    Dim dFecha, PFecha As Date
    Dim var As Boolean
    var = True
    
    For i = Inicio To fgCronograma.Rows - 1
        If i = 1 Then
            dFecha = CDate(fgCronograma.TextMatrix(i, 1))
            dFecha = DateSerial(Year(dFecha), Month(dFecha), 1)
            dFecha = DateAdd("d", -1, dFecha)
            dFecha = DateAdd("d", 366, dFecha)
        End If
        
        If CDate(fgCronograma.TextMatrix(i, 1)) <= dFecha Then
            dCorto = dCorto + CDbl(fgCronograma.TextMatrix(i, 3))
        Else
            dLargo = dLargo + CDbl(fgCronograma.TextMatrix(i, 3))
        End If
        
        If fgCronograma.TextMatrix(i, 9) = "0" And var Then         'SI LA CUOTA ES PENDIENTE
            PFecha = CDate(fgCronograma.TextMatrix(i, 1))
            PFecha = DateSerial(Year(PFecha), Month(PFecha), 1)
            PFecha = DateAdd("d", -1, PFecha)
            PFecha = DateAdd("d", 366, PFecha)
            var = False
        End If
        
        If var = False Then
            If CDate(fgCronograma.TextMatrix(i, 1)) <= PFecha Then
                PCorto = PCorto + CDbl(fgCronograma.TextMatrix(i, 3))
            Else
                PLargo = PLargo + CDbl(fgCronograma.TextMatrix(i, 3))
            End If
        End If
    Next i
    txtDcorto.Text = Format(dCorto, gsFormatoNumeroView)
    txtDlargo.Text = Format(dLargo, gsFormatoNumeroView)
    txtPCorto.Text = Format(PCorto, gsFormatoNumeroView)
    txtPLargo.Text = Format(PLargo, gsFormatoNumeroView)
    
    If fgCronograma.Rows - 1 > 0 Then
        With fgCronograma
        ' Recorre las filas
            Dim Fila As Integer
            Dim c As Integer
            For Fila = 1 To fgCronograma.Rows - 1
                'Indica la fila y la columna
                .row = Fila
                .Col = 2
                For c = 0 To .Cols - 1
                    If CDbl(fgCronograma.TextMatrix(Fila, 9)) = CDbl(1) Then
                        .Col = c
                        .CellForeColor = vbBlue  'color
                        .CellBackColor = &HC0FFC0
                    End If
                Next
                .Col = 2
            Next
        End With
    End If
End Sub

Private Sub cmdImportar_Click()

    frmAdeudCalMnt.Show 1

End Sub

Private Sub VerCuotasCanceladas()
Dim nCol As Integer
Dim i As Integer
'Verificando datos de Cuotas Canceladas
Set rs = oNAdeudCal.GetCalendarioDatos(Mid(lblCodAdeudado, 4, 13), Left(lblCodAdeudado, 2), Mid(lblCodAdeudado, 18, 7), , True)
If Not rs.EOF Then
    nCol = 0
    For i = 1 To fgCronograma.Rows - 1
        If fgCronograma.TextMatrix(i, 8) = rs!ctpocuota Then
            nCol = i
            Exit For
        End If
    Next
    If nCol = 0 Then
        MsgBox "No se pudo relacionar Pagos anteriores con nuevo Cronograma", vbInformation, "¡Aviso!"
    Else
        Do While Not rs.EOF
            fgCronograma.TextMatrix(i, 9) = rs!cEstado   'gTpoEstCuotaAdeudCanc
            fgCronograma.TextMatrix(i, 11) = rs!nInteresPagado
            fgCronograma.TextMatrix(i, 12) = rs!cMovNro
            i = i + 1
            rs.MoveNext
        Loop
    End If
End If
End Sub

Private Sub cmdImprimir_Click()
Dim oPrevio As PrevioFinan.clsPrevioFinan
Dim oCajaImp As nCajaGenImprimir
Set oPrevio = New PrevioFinan.clsPrevioFinan
Set oCajaImp = New nCajaGenImprimir
Dim lsImpre As String
    If fgCronograma.TextMatrix(1, 0) <> "" Then     ' ANGC2020  VERIFICAR EL PLAZO
        lsImpre = oCajaImp.ImprimeAdeudaCalendario(Me.fgCronograma.GetRsNew(1), CCur(txtCapital), CDate(txtFecha), Val(txtPlazoCuotas), _
                        Val(SpnCuotas.Valor), Val(SpnGracia.Valor), CCur(txtInteres), IIf(Me.optPeriodo(0).value = True, 360, 30), Me.lblCodAdeudado, Trim(Me.lblDescEntidad), nVal(txtTotalcapital), nVal(txtTotalInteres), nVal(txtTotalGeneral))
        oPrevio.Show lsImpre, "Calendario de Pago de Pagarés de Adeudados", False, 66
    Else
        MsgBox "Datos no procesados. Por favor Genere el calendario", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdNuevaCuota_Click()
    Dim Resultado As VbMsgBoxResult
    Dim lbDespuesUltCuotaPagada As Boolean
    Dim i As Long
    Dim lnCuotaNueva As Long
    
    Resultado = MsgBox("Presione SI para agregarla despues de la ultima cuota pagada" & Chr(10) & "Presione NO para agregarla despues de la ultima cuota", vbQuestion + vbYesNoCancel, "Nueva Cuota")
    If Resultado = vbYes Then
        lbDespuesUltCuotaPagada = True
    ElseIf Resultado = vbNo Then
        lbDespuesUltCuotaPagada = False
    Else
        Exit Sub
    End If
    
    If lbDespuesUltCuotaPagada Then 'Buscamos la ultima cuota pagada
        For i = 1 To fgCronograma.Rows - 1
            If Trim(fgCronograma.TextMatrix(i, 9)) = "1" Then
                lnCuotaNueva = CLng(fgCronograma.TextMatrix(i, 2))
            End If
        Next
        lnCuotaNueva = lnCuotaNueva + 1
    Else
        lnCuotaNueva = CLng(fgCronograma.TextMatrix(fgCronograma.Rows - 1, 2)) + 1
    End If
    
        
    fgCronograma.AdicionaFila lnCuotaNueva

    fgCronograma.TextMatrix(lnCuotaNueva, 0) = "-"
    fgCronograma.TextMatrix(lnCuotaNueva, 1) = Format(gdFecSis, "dd/mm/yyyy")
    fgCronograma.TextMatrix(lnCuotaNueva, 2) = lnCuotaNueva
    fgCronograma.TextMatrix(lnCuotaNueva, 3) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 4) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 5) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 6) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 7) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 8) = "2"
    fgCronograma.TextMatrix(lnCuotaNueva, 9) = "2"
    fgCronograma.TextMatrix(lnCuotaNueva, 10) = "0"
    fgCronograma.TextMatrix(lnCuotaNueva, 11) = "0"
    fgCronograma.TextMatrix(lnCuotaNueva, 12) = ""
    'ALPA 20130614************************************************
    fgCronograma.TextMatrix(lnCuotaNueva, 13) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 14) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 15) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 16) = "0.00"
    fgCronograma.TextMatrix(lnCuotaNueva, 17) = "0.00" 'ALPA20130904
    '*************************************************************
    
    For i = (lnCuotaNueva + 1) To fgCronograma.Rows - 1
        fgCronograma.TextMatrix(i, 2) = i
    Next
End Sub

Private Sub cmdSalir_Click()
    lbOk = False
    Unload Me
End Sub

Private Sub fgCronograma_DblClick()
If nVal(fgCronograma.TextMatrix(fgCronograma.row, 9)) = gTpoEstCuotaAdeudCanc Then
'    If MsgBox("Cuota ya Cancelada. ¿ Desea modificar datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
'        fgCronograma.lbEditarFlex = False
'    Else
'        fgCronograma.lbEditarFlex = True
'    End If
    MsgBox "Cuota ya Cancelada. Imposible modificar datos ", vbInformation, "¡Aviso!"
    fgCronograma.lbEditarFlex = False
    Exit Sub
End If
End Sub

Private Sub fgCronograma_EnterCell()
If nVal(fgCronograma.TextMatrix(fgCronograma.row, 9)) = gTpoEstCuotaAdeudCanc Then
    fgCronograma.lbEditarFlex = False
Else
    fgCronograma.lbEditarFlex = True
End If
End Sub

Private Sub fgCronograma_KeyPress(KeyAscii As Integer)
If nVal(fgCronograma.TextMatrix(fgCronograma.row, 9)) = gTpoEstCuotaAdeudCanc Then
    MsgBox "Cuota ya Cancelada. Imposible modificar datos ", vbInformation, "¡Aviso!"
    KeyAscii = 27
    Exit Sub
End If
End Sub

Private Sub fgCronograma_OnCellChange(pnRow As Long, pnCol As Long)
Dim i As Integer
Dim lsTpoCuotaRow As String

If pnCol = 3 Or pnCol = 4 Or pnCol = 5 Or pnCol = 13 Or pnCol = 14 Or pnCol = 15 Then
    lsTpoCuotaRow = fgCronograma.TextMatrix(pnRow, 8)
    'ANGC2021
    For i = pnRow To fgCronograma.Rows - 1
        If optTpoCuota(0) Or optTpoCuota(1) Then
            If (lsTpoCuotaRow = gCGTipoCuotCalIFCuota And fgCronograma.TextMatrix(i, 8) = gCGTipoCuotCalIFCuota) Or lsTpoCuotaRow = gCGTipoCuotCalIFNoConcesional Then
                fgCronograma.TextMatrix(i, 6) = Format(nVal(fgCronograma.TextMatrix(i, 3)) + nVal(fgCronograma.TextMatrix(i, 4)) + nVal(fgCronograma.TextMatrix(i, 5)) + nVal(fgCronograma.TextMatrix(i, 13)) + nVal(fgCronograma.TextMatrix(i, 14)) + nVal(fgCronograma.TextMatrix(i, 15)), gsFormatoNumeroView)
                'ALPA 20130614********************************
                fgCronograma.TextMatrix(i, 16) = Format(nVal(fgCronograma.TextMatrix(i, 3)) + nVal(fgCronograma.TextMatrix(i, 4)) + nVal(fgCronograma.TextMatrix(i, 5)) + nVal(fgCronograma.TextMatrix(i, 13)) + nVal(fgCronograma.TextMatrix(i, 14)) + nVal(fgCronograma.TextMatrix(i, 15)), gsFormatoNumeroView)
                '*********************************************
                If i = 1 Then
                    If nVal(Me.txtTramo) <> 0 Then
                        fgCronograma.TextMatrix(i, 7) = Format(Round(nVal(txtCapital) * nVal(txtTramo) / 100, 2) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        'ALPA 20130614********************************
                        fgCronograma.TextMatrix(i, 17) = Format(Round(nVal(txtCapital) * nVal(txtTramo) / 100, 2) - nVal(fgCronograma.TextMatrix(i, 3)) - -nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        '*********************************************
                    Else
                        fgCronograma.TextMatrix(i, 7) = Format(nVal(txtCapital) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        'ALPA 20130614********************************
                        fgCronograma.TextMatrix(i, 17) = Format(nVal(txtCapital) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        '*********************************************
                    End If
                Else
                    If lsTpoCuotaRow = gCGTipoCuotCalIFNoConcesional And fgCronograma.TextMatrix(i, 2) = "1" Then
                        fgCronograma.TextMatrix(i, 7) = Format(Round(nVal(txtCapital) * (100 - nVal(txtTramo)) / 100, 2) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        'ALPA 20130614********************************
                        fgCronograma.TextMatrix(i, 17) = Format(Round(nVal(txtCapital) * (100 - nVal(txtTramo)) / 100, 2) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        '*********************************************
                    Else
                        fgCronograma.TextMatrix(i, 7) = Format(nVal(fgCronograma.TextMatrix(i - 1, 7)) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        'ALPA 20130614********************************
                        fgCronograma.TextMatrix(i, 17) = Format(nVal(fgCronograma.TextMatrix(i - 1, 7)) - nVal(fgCronograma.TextMatrix(i, 3)) - nVal(fgCronograma.TextMatrix(i, 13)), gsFormatoNumeroView)
                        '*********************************************
                    End If
                End If
            End If
        Else
            fgCronograma.TextMatrix(i, 6) = Format(nVal(fgCronograma.TextMatrix(i, 3)) + nVal(fgCronograma.TextMatrix(i, 4)) + nVal(fgCronograma.TextMatrix(i, 5)) + nVal(fgCronograma.TextMatrix(i, 13)) + nVal(fgCronograma.TextMatrix(i, 14)) + nVal(fgCronograma.TextMatrix(i, 15)), gsFormatoNumeroView)
            fgCronograma.TextMatrix(i, 16) = Format(nVal(fgCronograma.TextMatrix(i, 3)) + nVal(fgCronograma.TextMatrix(i, 4)) + nVal(fgCronograma.TextMatrix(i, 5)) + nVal(fgCronograma.TextMatrix(i, 13)) + nVal(fgCronograma.TextMatrix(i, 14)) + nVal(fgCronograma.TextMatrix(i, 15)), gsFormatoNumeroView)
        End If
    Next
    txtTotalGeneral = Format(fgCronograma.SumaRow(6), gsFormatoNumeroView)
    txtTotalInteres = Format(fgCronograma.SumaRow(4), gsFormatoNumeroView)
    txtTotalcapital = Format(fgCronograma.SumaRow(3), gsFormatoNumeroView)
    txtTotalcapitalC = Format(fgCronograma.SumaRow(13), gsFormatoNumeroView)
    txtTotalInteresC = Format(fgCronograma.SumaRow(14), gsFormatoNumeroView)
    
    Sumatoria
End If
End Sub

Private Sub Form_Load()
CentraForm Me


    Dim oGen As New DGeneral
    Set oGen = Nothing
    
    Set oDAdeud = New DCaja_Adeudados
    Set oNAdeudCal = New nAdeudCal
    oNAdeudCal.Inicio gsFormatoFecha
    LimpiaControles
    
    lblCodAdeudado = lsIFCod
    lblDescEntidad = lsIFDescrip
    If ldFecContrato <> 0 Then
       txtFecha = ldFecContrato
    End If
    txtCapital = Format(lnCapital, gsFormatoNumeroView)
    txtConcesional = Format(lnConcesionado, gsFormatoNumeroView) 'ALPA20130614
    If nVal(txtConcesional.Text) > 0 Then
        chkMiVivienda.value = 1
    End If
    If Not lbCargaIF Then
        txtCapital.Locked = False
    Else
        txtCapital.Locked = True
    End If
    
    Me.lblEntidad.Visible = lbCargaIF
    Me.lblCodAdeudado.Visible = lbCargaIF
    Me.lblDescEntidad.Visible = lbCargaIF
    If Mid(lblCodAdeudado, 20, 1) = "2" Then
        If txtFecha = gdFecSis Then
            txtTipoCambio = gnTipCambio
        Else
            Dim oTC As New nTipoCambio
            txtTipoCambio = Format(oTC.EmiteTipoCambio(txtFecha, TCFijoDia), "#,#0.00####")
            Set oTC = Nothing
        End If
        txtCapital.BackColor = vbGreen
        txtTotalGeneral.BackColor = vbGreen
        txtConcesional.BackColor = vbGreen 'ALPA 20130614
        FraTipoCambio.Enabled = False
        chkVac.Visible = False
    Else
        txtCapital.BackColor = vbWhite
        txtTotalGeneral.BackColor = vbWhite
        Me.FraTipoCambio.Enabled = True
        txtConcesional.BackColor = vbWhite 'ALPA 20130614
    End If
    
    If Not lbCargaIF Then
        cmdAceptar.Visible = False
    End If
    
    If lbGetAdeud Then
'        cmdGenerar.Enabled = False
        Dim oCaja As New NCajaCtaIF
        Set rs = oCaja.GetDatosCtaIf(Mid(lblCodAdeudado, 4, 13), Left(lblCodAdeudado, 2), Mid(lblCodAdeudado, 18, 7))
        If Not rs.EOF Then
            If rs!nTpoCuota = gAdeudTpoCuotaFija Then
                optTpoCuota(0).value = True
            ElseIf rs!nTpoCuota = gAdeudTpoCuotaVariable Then
                optTpoCuota(1).value = True
            Else
                optTpoCuota(2).value = True
            End If
            If rs!cMonedaPago = "2" And Mid(lblCodAdeudado, 20, 1) = "1" Then
                chkVac.value = vbChecked
            End If
            txtPlazoCuotas = rs!nCtaIFPlazo
            SpnCuotas.Valor = rs!nCtaIFCuotas
            chkInterno.value = IIf(rs!cPlaza = 1, vbChecked, vbUnchecked)
            SpnGracia.Valor = rs!nPeriodoGracia
            'Este es el interes Provisionado
            'Debe ir la tasa de interes de CtaIFInteres
            If lnTasaInt = 0 Then
                txtInteres = oCaja.GetCtaIfInteres(Mid(lblCodAdeudado, 4, 13), Left(lblCodAdeudado, 2), Mid(lblCodAdeudado, 18, 7))
            Else
                txtInteres = Format(lnTasaInt, "#,##0.00###")
            End If
            txtTramo = rs!nTramoConcesion
            txtCuotaPagoK = rs!nCuotaPagoCap
            Me.txtFecha = IIf(IsNull(rs!dCtaIFAper), "__/__/____", Format(rs!dCtaIFAper, gsFormatoFechaView))
            Me.txtComision = rs!nComisionCuota
            If optTpoCuota(0) Or optTpoCuota(1) Then 'No para Adeudados
                If rs!nFechaFija > 0 Then
                    OptTpoPeriodo(1).value = True
                    Me.txtFechaCuota = rs!nFechaFija
                Else
                    OptTpoPeriodo(0).value = True
                End If
            Else
                Me.txtFechaCuota = rs!nFechaFija
            End If
        End If
        
        oNAdeudCal.GeneraCalendarioBase fgCronograma, Trim(lblCodAdeudado), nVal(txtCapital), nVal(txtTramo)
        
        If optTpoCuota(0).value = True Then
            If fgCronograma.Rows > 2 Then
                Me.txtCuota = Format(fgCronograma.TextMatrix(2, 6), gsFormatoNumeroView)
            Else
                Me.txtCuota = Format(fgCronograma.TextMatrix(1, 6), gsFormatoNumeroView)
            End If
        End If
        
'    Else
'        oNAdeudCal.GeneraCalendario fgCronograma, CCur(txtCapital), Val(SpnCuotas.Valor), Val(txtPlazoCuotas), _
'           IIf(Me.optPeriodo(0).value = True, 360, 60), gdFecSis, _
'           nVal(SpnGracia.Valor), nVal(txtInteres), nVal(txtCuotaPagoK)
    End If
    txtTotalGeneral = Format(oNAdeudCal.TotalGeneral, gsFormatoNumeroView)
    txtTotalInteres = Format(oNAdeudCal.TotalInteres, gsFormatoNumeroView)
    txtTotalcapital = Format(oNAdeudCal.TotalCapital, gsFormatoNumeroView)
    'ALPA 20130614***************************
    txtTotalcapitalC = Format(oNAdeudCal.TotalCapitalC, gsFormatoNumeroView)
    txtTotalInteresC = Format(oNAdeudCal.TotalInteresC, gsFormatoNumeroView)
    '****************************************
    If lbGetAdeud Then
        Sumatoria
    End If
    If lbSoloConsulta Then
        Me.Frame1.Enabled = False
        Me.Frame2.Enabled = False
        Me.Frame3.Visible = False
        Me.fraIngAdeud.Enabled = False
        Me.fgCronograma.lbEditarFlex = False
        Me.FraTipoCambio.Enabled = False
    End If
    cmdNuevaCuota.Visible = fbAgregarCuota 'EJVG20121205
End Sub

Private Sub LimpiaControles()
    txtInteres = 0
    fraIngAdeud.Enabled = True
    SpnCuotas.Valor = 0
    SpnGracia.Valor = 0
    txtPlazoCuotas = 0
    chkPagoK.value = 0
    txtTotalcapital = 0
    txtTotalGeneral = 0
    txtTotalInteres = 0
    txtTipoCambio = ""
    txtCuotaPagoK = 0
    txtDcorto.Text = "0"
    txtDlargo.Text = "0"
    txtPCorto.Text = "0"
    txtPLargo.Text = "0"
    If txtCapital.Locked = False Then
        txtCapital = 0
    End If
    ReDim ltCalendario(0)
    gbGeneraCal = False
    fgCronograma.row = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
lbOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oNAdeudCal = Nothing
Set oDAdeud = Nothing
End Sub

Private Sub optPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SpnCuotas.SetFocus
    End If
End Sub

Private Sub optTpoCuota_Click(Index As Integer)
    txtCuota = "0"
    Select Case Index
    Case 0:
        txtCuota.Enabled = True
        frmPeriodo.Enabled = True
        Frame2.Enabled = True
    
        txtTramo.Text = "0"
        txtTramo.Enabled = True
        chkMiVivienda = False
        chkMiVivienda.Enabled = True
    Case 1:
        txtCuota = "0"
        txtCuota.Enabled = False
        frmPeriodo.Enabled = True
        Frame2.Enabled = True
    
        txtTramo.Text = "0"
        txtTramo.Enabled = True
        chkMiVivienda = False
        chkMiVivienda.Enabled = True
    Case 2:
        txtCuota = "0"
        txtCuota.Enabled = False
        frmPeriodo.Enabled = False
        txtPlazoCuotas = "30"
    
        txtTramo.Text = "0"
        txtTramo.Enabled = False
        chkMiVivienda = False
        chkMiVivienda.Enabled = False
    End Select
End Sub

Private Sub optTpoCuota_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        txtCuota.SetFocus
    Else
        If optTpoCuota(2) = False Then
            If OptTpoPeriodo(0) Then 'COMENT ANGC2020
                OptTpoPeriodo(0).SetFocus
            Else
                OptTpoPeriodo(1).SetFocus
            End If
        Else
            txtComision.SetFocus
        End If
    End If
End If
End Sub
Private Sub OptTpoPeriodo_Click(Index As Integer)
txtPlazoCuotas = "0"
txtFechaCuota = "0"
If Index = 0 Then
    txtPlazoCuotas.Enabled = True
    txtFechaCuota.Enabled = False
Else
    txtPlazoCuotas = "30"
    txtPlazoCuotas.Enabled = False
    txtFechaCuota.Enabled = True
End If
End Sub

Private Sub OptTpoPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtPlazoCuotas.Enabled Then
        txtPlazoCuotas.SetFocus
    Else
        txtFechaCuota.SetFocus
    End If
End If
End Sub
Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.SpnGracia.SetFocus
    End If
End Sub

Private Sub SpnGracia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkInterno.SetFocus
    End If
End Sub

Private Sub txtcapital_GotFocus()
    fEnfoque txtCapital
End Sub
'ALPA20130614
Private Sub txtConcesional_GotFocus()
    fEnfoque txtConcesional
End Sub
Private Sub txtConcesional_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtConcesional, KeyAscii, 10)
  If KeyAscii = 13 Then
        If Me.chkVac.Enabled Then
            Me.chkVac.SetFocus
        Else
            If txtTipoCambio.Enabled Then
                txtTipoCambio.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtConcesional_LostFocus()
    txtConcesional = Format(txtConcesional, "#,#0.00")
End Sub
'************
Private Sub txtcapital_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCapital, KeyAscii, 10)
  If KeyAscii = 13 Then
        If Me.chkVac.Enabled Then
            Me.chkVac.SetFocus
        Else
            If txtTipoCambio.Enabled Then
                txtTipoCambio.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtcapital_LostFocus()
    txtCapital = Format(txtCapital, "#,#0.00")
End Sub

Private Sub txtComision_GotFocus()
fEnfoque txtComision
End Sub

Private Sub txtComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtComision, KeyAscii, 10, 2)
If KeyAscii = 13 Then
   txtComision = Format(txtComision, gsFormatoNumeroView)
   cmdGenerar.SetFocus
End If
End Sub

Private Sub txtCuotaPagoK_GotFocus()
    fEnfoque txtCuotaPagoK
End Sub

Private Sub txtCuotaPagoK_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtTramo.SetFocus
    End If
End Sub


Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCapital.SetFocus
    End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
If ValidaFecha(txtFecha) = "" Then
    If txtFecha = gdFecSis Then
        txtTipoCambio = gnTipCambio
    Else
        Dim oTC As New nTipoCambio
        txtTipoCambio = Format(oTC.EmiteTipoCambio(txtFecha, TCFijoDia), "#,#0.00####")
        Set oTC = Nothing
    End If
End If
End Sub

Private Sub txtFechaCuota_GotFocus()
fEnfoque txtFechaCuota
End Sub

Private Sub txtFechaCuota_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtComision.SetFocus
End If
End Sub

Private Sub txtInteres_GotFocus()
    fEnfoque txtInteres
End Sub
Private Sub txtInteres_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtInteres, KeyAscii, 14, 6)
    If KeyAscii = 13 Then
        optPeriodo(0).SetFocus
    End If
End Sub
Private Sub txtInteres_LostFocus()
    txtInteres = Format(txtInteres, "#,#0.0000##")
End Sub

Private Sub txtPlazoCuotas_GotFocus()
    fEnfoque txtPlazoCuotas
End Sub
Private Sub txtPlazoCuotas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtComision.SetFocus
    End If
End Sub

Private Sub txtTipoCambio_GotFocus()
    fEnfoque txtTipoCambio
End Sub
Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(Me.txtTipoCambio, KeyAscii, 8, 5)
    If KeyAscii = 13 Then
        txtInteres.SetFocus
    End If
End Sub

Private Sub txtTipoCambio_LostFocus()
    If txtTipoCambio = "" Then
        txtTipoCambio = 0
        Exit Sub
    End If
    If lbCargaIF Then
        If Val(txtTipoCambio) > 0 And Trim(Mid(lblCodAdeudado, 20, 1)) = "1" Then
            txtCapital = Format(lnCapital / Val(txtTipoCambio), "#,#0.00")
        End If
    End If
End Sub

'Valida todos los ingresos de datos en los controles respectivos de la interfaz
Private Function ValidaInterfaz() As Boolean
    On Error GoTo ValidaInterFazErr

    ValidaInterfaz = False
    If Me.lblCodAdeudado = "" And lblCodAdeudado.Visible Then
        MsgBox "Entidad Principal no Válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        Exit Function
    End If
    If Val(txtCapital) = 0 Then
        MsgBox "Capital no Válido", vbInformation, "Aviso"
        If txtCapital.Enabled Then
            txtCapital.SetFocus
        End If
        Exit Function
    End If
    If ValidaFecha(Me.txtFecha) <> "" Then
        MsgBox "Fecha no válida", vbInformation, "Aviso"
        Exit Function
    End If
    If optTpoCuota(2) = False Then              ' angc2021
        If OptTpoPeriodo(0) And Val(txtPlazoCuotas) = 0 Then 'COMENT ANGC2020
            MsgBox "Plazo no Válido", vbInformation, "Aviso"
            txtPlazoCuotas.SetFocus
            Exit Function
        End If
    End If
    If optTpoCuota(2) = False Then              ' angc2021
        If OptTpoPeriodo(1) And Val(txtFechaCuota) = 0 Then
            MsgBox "Indicar día de pago de cada Cuota", vbInformation, "¡Aviso!"
            txtFechaCuota.SetFocus
            Exit Function
        End If
    End If
    If Val(Me.txtInteres) = 0 Then
        MsgBox "Interes no Válido", vbInformation, "Aviso"
        Me.txtInteres.SetFocus
        Exit Function
    End If
    If Me.SpnCuotas.Valor = 0 Then
        MsgBox "Numero de Cuotas no Válido", vbInformation, "Aviso"
        Me.SpnCuotas.SetFocus
        Exit Function
    End If
    If Val(Me.txtCuotaPagoK) = 0 And Me.chkPagoK.value = 1 Then
        MsgBox "Cuota de Capital no Válido", vbInformation, "Aviso"
        Me.txtCuotaPagoK.SetFocus
        Exit Function
    Else
        If Val(Me.txtCuotaPagoK) > Me.SpnCuotas.Valor Then
            MsgBox "Cuota de Capital no puede ser mayor que el N° de Cuotas", vbInformation, "Aviso"
            Me.txtCuotaPagoK.SetFocus
            Exit Function
        End If
    End If
    If optTpoCuota(0).value And nVal(txtCuota) = 0 Then
        MsgBox "La Cuota Fija se calculará!", vbInformation, "¡Aviso!"
    End If
    'ALPA 201306
    If chkMiVivienda.value = 1 Then
        If Trim(txtConcesional.Text) = "" Then
            MsgBox "Ingresar el capital concesional", vbInformation, "Aviso"
            Me.txtConcesional.SetFocus
            Exit Function
        End If
        
        If CDbl(txtCapital.Text) <= CDbl(txtConcesional.Text) Then
            MsgBox "El Capital No Concesional debe ser mayor que el capital concesional", vbInformation, "Aviso"
            Me.txtConcesional.SetFocus
            Exit Function
        End If
    End If
    'END
    ValidaInterfaz = True
    Exit Function
ValidaInterFazErr:
    MsgBox "Error N°[" & Err.Number & "]", vbInformation, "Aviso"
End Function

Public Property Get OK() As Boolean
OK = lbOk
End Property

Public Property Let OK(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property

Private Sub txtTramo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTramo, KeyAscii, 14, 2)
If KeyAscii = 13 Then
    If optTpoCuota(0) Then
        optTpoCuota(0).SetFocus
    Else
        optTpoCuota(1).SetFocus
    End If
End If
End Sub

Public Property Get nCapital() As Currency
nCapital = lnCapital
End Property

Public Property Let nCapital(ByVal vNewValue As Currency)
lnCapital = vNewValue
End Property

Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
    On Error GoTo Error_Handler
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 1 To .Rows - 1
            If Fila = 1 Then
                o_Hoja.Cells(Fila, 1).value = "Contrato"
                o_Hoja.Cells(Fila, 2).value = "Secuencia"
                o_Hoja.Cells(Fila, 3).value = "Cuota"
                o_Hoja.Cells(Fila, 4).value = "F .Venc"
                o_Hoja.Cells(Fila, 5).value = "N dias"
                o_Hoja.Cells(Fila, 6).value = "Moneda"
                o_Hoja.Cells(Fila, 7).value = "Principal"
                o_Hoja.Cells(Fila, 8).value = "Interes"
                o_Hoja.Cells(Fila, 9).value = "Comision"
                o_Hoja.Cells(Fila, 10).value = "Monto a cobrar"
                o_Hoja.Cells(Fila, 11).value = "Principal por Vencer"
                o_Hoja.Cells(Fila, 12).value = "Capitalizacion Int"
                o_Hoja.Cells(Fila, 13).value = "Estado Cuota"
            End If
            'For Columna = 0 To .Cols - 1
            'Next
            o_Hoja.Cells(Fila + 1, 1).value = "'" & Archivo
            o_Hoja.Cells(Fila + 1, 2).value = ""
            o_Hoja.Cells(Fila + 1, 3).value = .TextMatrix(Fila, 2)
            o_Hoja.Cells(Fila + 1, 4).value = Format(.TextMatrix(Fila, 1), "DD/MM/YYYY")
            o_Hoja.Cells(Fila + 1, 5).value = .TextMatrix(Fila, 10)
            o_Hoja.Cells(Fila + 1, 6).value = "Moneda"
            o_Hoja.Cells(Fila + 1, 7).value = .TextMatrix(Fila, 3)
            o_Hoja.Cells(Fila + 1, 8).value = .TextMatrix(Fila, 4)
            o_Hoja.Cells(Fila + 1, 9).value = .TextMatrix(Fila, 5)
            o_Hoja.Cells(Fila + 1, 10).value = .TextMatrix(Fila, 16)
            o_Hoja.Cells(Fila + 1, 11).value = .TextMatrix(Fila, 17)
            o_Hoja.Cells(Fila + 1, 12).value = "0"
            o_Hoja.Cells(Fila + 1, 13).value = .TextMatrix(Fila, 9)
        Next
    End With
    o_Libro.Close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function
' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub
