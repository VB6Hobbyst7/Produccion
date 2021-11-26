VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredSugerencia_NEW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sugerencia del Crédito"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   Icon            =   "frmCredSugerencia_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8190
      TabIndex        =   54
      ToolTipText     =   "Salir"
      Top             =   10200
      Width           =   1090
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Ca&ncelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7050
      TabIndex        =   52
      ToolTipText     =   "Limpiar la Pantalla"
      Top             =   10200
      Width           =   1090
   End
   Begin VB.CommandButton cmdCheckList 
      Caption         =   "CheckList"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4830
      TabIndex        =   188
      ToolTipText     =   "CheckList"
      Top             =   10200
      Width           =   1090
   End
   Begin VB.CommandButton CmdGastos 
      Caption         =   "&Gastos"
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
      Height          =   375
      Left            =   3700
      TabIndex        =   159
      ToolTipText     =   "Generar Plan Gastos"
      Top             =   10200
      Width           =   1090
   End
   Begin VB.CommandButton CmdDesembolsos 
      Caption         =   "Desem&bolsos"
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
      Height          =   375
      Left            =   2370
      TabIndex        =   55
      ToolTipText     =   "Ingresar los Desembolsos Parciales"
      Top             =   10200
      Width           =   1300
   End
   Begin VB.CommandButton CmdCalend 
      Caption         =   "Calen&dario"
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
      Height          =   375
      Left            =   1250
      TabIndex        =   53
      ToolTipText     =   "Mostrar el Calendario de Pagos"
      Top             =   10200
      Width           =   1090
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   375
      Left            =   120
      TabIndex        =   51
      ToolTipText     =   "Grabar Datos de Sugerencia"
      Top             =   10200
      Width           =   1090
   End
   Begin VB.Frame Frame7 
      ForeColor       =   &H80000006&
      Height          =   660
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.CommandButton cmdVerEntidades 
         Caption         =   "Ver Entidades"
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
         Left            =   7560
         TabIndex        =   181
         Top             =   200
         Width           =   1335
      End
      Begin VB.CommandButton CmdCredVig 
         Caption         =   "Creditos Vigentes"
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
         Left            =   5640
         TabIndex        =   158
         ToolTipText     =   "Creditos Vigentes"
         Top             =   200
         Width           =   1815
      End
      Begin VB.CommandButton cmdbuscar 
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
         Height          =   360
         Left            =   3960
         TabIndex        =   2
         ToolTipText     =   "Buscar Credito"
         Top             =   200
         Width           =   1575
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   240
         TabIndex        =   1
         Top             =   165
         Width           =   3660
         _extentx        =   6456
         _extenty        =   741
         texto           =   "Credito"
         enabledcmac     =   -1  'True
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Solicitud"
      TabPicture(0)   =   "frmCredSugerencia_NEW.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "feDeudaComprar"
      Tab(0).Control(1)=   "Frame16"
      Tab(0).Control(2)=   "Frame12"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Sugerencia"
      TabPicture(1)   =   "frmCredSugerencia_NEW.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Inf.Relacionada"
      TabPicture(2)   =   "frmCredSugerencia_NEW.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraVinculosEmpresas"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "fraOpGarantia"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Garantías"
      TabPicture(3)   =   "frmCredSugerencia_NEW.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame20"
      Tab(3).ControlCount=   1
      Begin SICMACT.FlexEdit feDeudaComprar 
         Height          =   2730
         Left            =   -74760
         TabIndex        =   195
         Top             =   5760
         Width           =   8700
         _extentx        =   15346
         _extenty        =   4815
         cols0           =   12
         highlight       =   2
         encabezadosnombres=   $"frmCredSugerencia_NEW.frx":037A
         encabezadosanchos=   "400-2500-1500-800-900-1300-1200-0-0-0-0-0"
         font            =   "frmCredSugerencia_NEW.frx":0407
         fontfixed       =   "frmCredSugerencia_NEW.frx":042F
         columnasaeditar =   "X-X-X-X-X-X-6-X-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0-0-1-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-C-R-R-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
         cantentero      =   12
         textarray0      =   "N°"
         lbflexduplicados=   0   'False
         lbultimainstancia=   -1  'True
         tipobusqueda    =   6
         lbformatocol    =   -1  'True
         lbpuntero       =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   405
         rowheight0      =   300
      End
      Begin VB.Frame Frame16 
         Caption         =   " Datos de Compra de Deuda"
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
         Height          =   3330
         Left            =   -74880
         TabIndex        =   196
         Top             =   5520
         Width           =   8925
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   177
         Top             =   4320
         Visible         =   0   'False
         Width           =   8925
         Begin VB.TextBox txtMontoMivivienda 
            Height          =   375
            Left            =   2760
            TabIndex        =   178
            Text            =   "0.00"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label29 
            Caption         =   "Valor de Venta (MIVIVIENDA)"
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
            Left            =   120
            TabIndex        =   179
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Garantias"
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
         Height          =   3180
         Left            =   -74880
         TabIndex        =   46
         Top             =   600
         Width           =   8880
         Begin VB.CommandButton CmdGarantia 
            Caption         =   "Garantias"
            Height          =   330
            Left            =   5160
            TabIndex        =   47
            Top             =   2640
            Width           =   1515
         End
         Begin SICMACT.FlexEdit FEGarantCred 
            Height          =   2190
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   8595
            _extentx        =   15161
            _extenty        =   3863
            cols0           =   12
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "-Garantia-Gravament-Comercial-Realizacion-Disponible-Titular-Nro Docum-TipoDoc-cNumGarant-Legal-Poliza"
            encabezadosanchos=   "300-3800-1200-1200-1200-1200-3500-1200-0-1500-1800-1500"
            font            =   "frmCredSugerencia_NEW.frx":0455
            fontfixed       =   "frmCredSugerencia_NEW.frx":0481
            columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "C-C-R-R-R-R-L-L-L-C-L-L"
            formatosedit    =   "0-0-2-2-2-2-0-0-0-0-0-0"
            lbultimainstancia=   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   300
            rowheight0      =   300
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   7440
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label LblTotGarantiasGrav 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   3600
            TabIndex        =   50
            Top             =   2640
            Width           =   1200
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Total Garantias Gravament:"
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
            Height          =   195
            Left            =   1200
            TabIndex        =   49
            Top             =   2640
            Width           =   2355
         End
      End
      Begin VB.Frame FraDatos 
         Caption         =   "Datos Sugeridos"
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
         Height          =   8880
         Left            =   120
         TabIndex        =   45
         Top             =   330
         Width           =   8895
         Begin VB.Frame fraGracia 
            Height          =   615
            Left            =   5040
            TabIndex        =   70
            ToolTipText     =   "El interés de gracia ser incluida en la cuota inicial, según periodo."
            Top             =   3975
            Width           =   1920
            Begin VB.CommandButton cmdgracia 
               Caption         =   "-->"
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
               Left            =   2760
               TabIndex        =   76
               Top             =   240
               Width           =   555
            End
            Begin VB.TextBox txtPerGra 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1200
               MaxLength       =   4
               TabIndex        =   75
               Text            =   "0"
               Top             =   240
               Width           =   660
            End
            Begin VB.TextBox TxtTasaGracia 
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
               Height          =   285
               Left            =   1200
               TabIndex        =   74
               Top             =   240
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton optTipoGracia 
               Caption         =   "Gracia en Cuotas"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   73
               Top             =   600
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.OptionButton optTipoGracia 
               Caption         =   "Capitalizar"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   72
               Top             =   600
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.CheckBox chkIncremenK 
               Caption         =   "Incrementa Capital"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1200
               TabIndex        =   71
               Top             =   840
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.Label LblTasaGracia 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   79
               Top             =   240
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.Label lblPerGra 
               AutoSize        =   -1  'True
               Caption         =   "Periodo Días:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   195
               TabIndex        =   77
               Top             =   280
               Width           =   975
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tasa :"
               Height          =   195
               Left            =   1320
               TabIndex        =   78
               Top             =   285
               Visible         =   0   'False
               Width           =   450
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Tipo Pago:           Fecha Pago:"
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
            Left            =   120
            TabIndex        =   80
            Top             =   3980
            Width           =   4875
            Begin VB.ComboBox cmbTipoPer 
               Height          =   315
               ItemData        =   "frmCredSugerencia_NEW.frx":04AF
               Left            =   120
               List            =   "frmCredSugerencia_NEW.frx":04B9
               Style           =   2  'Dropdown List
               TabIndex        =   200
               Top             =   210
               Width           =   1455
            End
            Begin VB.OptionButton opttper 
               Caption         =   "&Periodo Fijo"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   85
               Top             =   210
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton opttper 
               Caption         =   "Fec&ha Fija"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   84
               Top             =   300
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.TextBox TxtDiaFijo2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   4800
               MaxLength       =   2
               TabIndex        =   83
               Text            =   "00"
               Top             =   435
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.TextBox TxtDiaFijo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   4470
               MaxLength       =   2
               TabIndex        =   82
               Top             =   210
               Width           =   330
            End
            Begin VB.CheckBox ChkProxMes 
               Caption         =   "Prox Mes"
               Enabled         =   0   'False
               Height          =   210
               Left            =   4800
               TabIndex        =   81
               Top             =   240
               Visible         =   0   'False
               Width           =   960
            End
            Begin MSMask.MaskEdBox txtFechaFija 
               Height          =   300
               Left            =   2880
               TabIndex        =   86
               ToolTipText     =   "Presione Enter"
               Top             =   210
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblDia2 
               AutoSize        =   -1  'True
               Caption         =   "&Dia 2:"
               Height          =   195
               Left            =   3960
               TabIndex        =   89
               Top             =   600
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.Label LblDia 
               AutoSize        =   -1  'True
               Caption         =   "&Dia:"
               Enabled         =   0   'False
               Height          =   195
               Left            =   4180
               TabIndex        =   88
               Top             =   255
               Width           =   285
            End
            Begin VB.Label lblFechaPago 
               AutoSize        =   -1  'True
               Caption         =   "&Ingresar Fecha:"
               Height          =   195
               Left            =   1740
               TabIndex        =   87
               Top             =   280
               Width           =   1110
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   " Mantenimiento :"
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
            Left            =   1800
            TabIndex        =   193
            Top             =   240
            Width           =   1575
            Begin VB.CommandButton cmdPersona 
               Caption         =   "Persona ..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   240
               TabIndex        =   194
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Formato Eval."
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
            Left            =   120
            TabIndex        =   191
            Top             =   240
            Width           =   1575
            Begin VB.CommandButton cmdEvaluacion 
               Caption         =   "Evaluación ..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   120
               TabIndex        =   192
               ToolTipText     =   "Evaluación del Crédito"
               Top             =   240
               Width           =   1320
            End
         End
         Begin VB.CheckBox chkAutoCalifCPP 
            Caption         =   "Autorización Calificación CPP 6 Meses"
            Height          =   315
            Left            =   4800
            TabIndex        =   190
            Top             =   8175
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.CheckBox chkTasa 
            Caption         =   "[Exoneración Tasa]"
            Height          =   255
            Left            =   4800
            TabIndex        =   184
            Top             =   8160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CheckBox ckcPreferencial 
            Caption         =   "Preferencial"
            Height          =   255
            Left            =   4800
            TabIndex        =   183
            Top             =   7920
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkCSP 
            Caption         =   "Construcción en Sitio Propio"
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
            Left            =   5400
            TabIndex        =   180
            Top             =   2800
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox TxtComenta 
            Height          =   1905
            Left            =   120
            MaxLength       =   380
            MultiLine       =   -1  'True
            TabIndex        =   169
            ToolTipText     =   "Los comentarios serán visualizados en la hoja de aprobación."
            Top             =   6300
            Width           =   4520
         End
         Begin VB.Frame Frame13 
            Caption         =   "Cuota Aprox."
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
            Left            =   6960
            TabIndex        =   166
            Top             =   3980
            Width           =   1815
            Begin VB.Label lblcuota 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   120
               TabIndex        =   167
               ToolTipText     =   "La cuota se calculará después de generar el Calendario."
               Top             =   240
               Width           =   1635
            End
         End
         Begin VB.CommandButton cmdExoneracion 
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
            Height          =   280
            Left            =   8400
            TabIndex        =   165
            ToolTipText     =   "Desembolso en Banco de la Nación"
            Top             =   7155
            Width           =   375
         End
         Begin VB.CheckBox ChkBloqueo 
            Caption         =   "Bloquear Desembolso"
            Height          =   240
            Left            =   4800
            TabIndex        =   164
            Top             =   6960
            Width           =   1905
         End
         Begin VB.CheckBox ChkExoneracion 
            Caption         =   "Autorizaciones no contempladas"
            Height          =   240
            Left            =   4800
            TabIndex        =   163
            Top             =   7200
            Width           =   3465
         End
         Begin VB.CheckBox ChkExoSegDes 
            Caption         =   "Exonerar Seguro Desgravamen"
            Height          =   240
            Left            =   4800
            TabIndex        =   162
            Top             =   7440
            Width           =   2745
         End
         Begin VB.CheckBox chkDesembBcoNac 
            Caption         =   "&Desembolso en Banco de la Nación"
            Height          =   240
            Left            =   4800
            TabIndex        =   161
            Top             =   7680
            Width           =   2985
         End
         Begin VB.CommandButton cmbDesBcoNac 
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
            Height          =   280
            Left            =   8400
            TabIndex        =   160
            ToolTipText     =   "Desembolso en Banco de la Nación"
            Top             =   7680
            Width           =   375
         End
         Begin VB.CheckBox chkGracia 
            Caption         =   "Gracia"
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
            Height          =   195
            Left            =   5150
            TabIndex        =   69
            Top             =   3980
            Width           =   915
         End
         Begin VB.Frame Frame4 
            Caption         =   " Fuentes de Ingreso "
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
            Left            =   120
            TabIndex        =   155
            Top             =   240
            Visible         =   0   'False
            Width           =   975
            Begin VB.CommandButton cmdSeleccionaFuente 
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
               Height          =   300
               Left            =   120
               TabIndex        =   157
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton cmdFuentes 
               Caption         =   "Ftes Ingreso"
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
               Left            =   600
               TabIndex        =   156
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   " Vinculados:"
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
            Left            =   3480
            TabIndex        =   153
            Top             =   240
            Width           =   1335
            Begin VB.CommandButton cmdVinculados 
               Caption         =   "Vinculados"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   120
               TabIndex        =   154
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   " Seguros Optativos "
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
            Height          =   780
            Left            =   5400
            TabIndex        =   148
            Top             =   140
            Width           =   3375
            Begin VB.ComboBox cmbMicroseguro 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   150
               Top             =   380
               Width           =   1455
            End
            Begin VB.ComboBox cmbBancaSeguro 
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   149
               Top             =   380
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Microseguro :"
               Height          =   195
               Left            =   120
               TabIndex        =   152
               Top             =   200
               Width           =   960
            End
            Begin VB.Label lblBancaSeguro 
               AutoSize        =   -1  'True
               Caption         =   "Multiriesgo :"
               Height          =   195
               Left            =   1680
               TabIndex        =   151
               Top             =   200
               Width           =   840
            End
         End
         Begin VB.CommandButton cmdVentasAnual 
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2520
            TabIndex        =   147
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdFlujoCaja 
            Caption         =   "Flujo Caja"
            Height          =   330
            Left            =   6120
            TabIndex        =   146
            Top             =   8880
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame Frame3 
            Caption         =   " Características de Credito "
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
            Height          =   2180
            Left            =   120
            TabIndex        =   118
            Top             =   960
            Width           =   8655
            Begin VB.CommandButton cmbCreditoVerdeDet 
               Caption         =   "E.A"
               Height          =   315
               Left            =   4680
               TabIndex        =   198
               ToolTipText     =   "EcoAhorro"
               Top             =   1440
               Visible         =   0   'False
               Width           =   460
            End
            Begin VB.CommandButton cmbAguaSaneamientoDet 
               Caption         =   "A.S"
               Height          =   315
               Left            =   4200
               TabIndex        =   197
               ToolTipText     =   "Agua y Saneamiento"
               Top             =   1440
               Visible         =   0   'False
               Width           =   460
            End
            Begin VB.ComboBox cmbDatoVivienda 
               Height          =   315
               Left            =   820
               Style           =   2  'Dropdown List
               TabIndex        =   173
               Top             =   1800
               Width           =   3370
            End
            Begin VB.CommandButton cmdActTipoCred 
               Caption         =   "->"
               Height          =   280
               Left            =   8360
               Style           =   1  'Graphical
               TabIndex        =   127
               ToolTipText     =   "Buscar tipo credito"
               Top             =   610
               Width           =   240
            End
            Begin VB.ComboBox cmbInstitucionFinanciera 
               Height          =   315
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   1800
               Width           =   3255
            End
            Begin VB.ComboBox cmbSubTipo 
               Height          =   315
               Left            =   5080
               Style           =   2  'Dropdown List
               TabIndex        =   125
               Top             =   600
               Width           =   3255
            End
            Begin VB.ComboBox cmbTipoCredito 
               Height          =   315
               Left            =   820
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   600
               Width           =   3360
            End
            Begin VB.ComboBox cmbProducto 
               Height          =   315
               Left            =   820
               Style           =   2  'Dropdown List
               TabIndex        =   123
               Top             =   240
               Width           =   3370
            End
            Begin VB.ComboBox cmbSubProd 
               Height          =   315
               Left            =   5080
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   240
               Width           =   3520
            End
            Begin VB.ComboBox cmbDestino 
               Height          =   315
               Left            =   820
               Style           =   2  'Dropdown List
               TabIndex        =   121
               Top             =   1440
               Width           =   3375
            End
            Begin VB.ComboBox cmbCampana 
               Height          =   315
               Left            =   5080
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   1080
               Width           =   3495
            End
            Begin VB.CommandButton cmdLineas 
               Height          =   315
               Left            =   8160
               Style           =   1  'Graphical
               TabIndex        =   119
               ToolTipText     =   "Buscar Lineas de Credito"
               Top             =   1815
               Visible         =   0   'False
               Width           =   390
            End
            Begin SICMACT.TxtBuscar txtBuscarLinea 
               Height          =   345
               Left            =   960
               TabIndex        =   128
               Top             =   2190
               Visible         =   0   'False
               Width           =   1785
               _extentx        =   3149
               _extenty        =   609
               appearance      =   1
               appearance      =   1
               font            =   "frmCredSugerencia_NEW.frx":050C
               appearance      =   1
            End
            Begin Spinner.uSpinner spnNumConCer 
               Height          =   315
               Left            =   960
               TabIndex        =   129
               Top             =   2235
               Visible         =   0   'False
               Width           =   645
               _ExtentX        =   1138
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
            Begin MSMask.MaskEdBox TxtFecDesemb 
               Height          =   315
               Left            =   7420
               TabIndex        =   130
               Top             =   1450
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin Spinner.uSpinner spnNumConMic 
               Height          =   330
               Left            =   3000
               TabIndex        =   131
               Top             =   2355
               Visible         =   0   'False
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
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
            Begin VB.Line Line1 
               BorderColor     =   &H80000003&
               X1              =   855
               X2              =   8520
               Y1              =   975
               Y2              =   975
            End
            Begin VB.Label lblLineaDesc 
               BackColor       =   &H8000000E&
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
               Height          =   315
               Left            =   2760
               TabIndex        =   136
               Top             =   2190
               Visible         =   0   'False
               Width           =   5295
            End
            Begin VB.Label Label14 
               Caption         =   "Consulta Score Microfinanzas:"
               Height          =   375
               Left            =   1800
               TabIndex        =   199
               Top             =   2205
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label lblMsj 
               Caption         =   "Obteniendo tipo de crédito..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   135
               Left            =   5280
               TabIndex        =   145
               Top             =   1920
               Width           =   2895
            End
            Begin VB.Label lblInstitucionFinanciera 
               Caption         =   "Inst. Corp.:"
               Height          =   255
               Left            =   40
               TabIndex        =   144
               Top             =   1860
               Width           =   855
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "-->SubTipo:"
               Height          =   195
               Left            =   4240
               TabIndex        =   143
               Top             =   645
               Width           =   825
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   420
               TabIndex        =   142
               Top             =   645
               Width           =   360
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Categoria:"
               Height          =   195
               Index           =   0
               Left            =   100
               TabIndex        =   141
               Top             =   285
               Width           =   720
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "-->Producto:"
               Height          =   195
               Left            =   4220
               TabIndex        =   140
               Top             =   285
               Width           =   870
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Destino:"
               Height          =   195
               Index           =   1
               Left            =   225
               TabIndex        =   139
               Top             =   1500
               Width           =   585
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Campaña:"
               Height          =   195
               Index           =   2
               Left            =   4365
               TabIndex        =   138
               Top             =   1140
               Width           =   720
            End
            Begin VB.Label Label18 
               Caption         =   "Linea:"
               Height          =   255
               Left            =   480
               TabIndex        =   137
               Top             =   2235
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label7 
               Caption         =   "Consultas Certicom :"
               Height          =   435
               Left            =   120
               TabIndex        =   135
               Top             =   2205
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label lblCondicion 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   300
               Left            =   820
               TabIndex        =   134
               Top             =   1080
               Width           =   3370
            End
            Begin VB.Label Label30 
               Caption         =   "Condición:"
               Height          =   255
               Left            =   45
               TabIndex        =   133
               Top             =   1155
               Width           =   840
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Desembolso:"
               Height          =   195
               Left            =   5740
               TabIndex        =   132
               Top             =   1515
               Width           =   1635
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   " Datos del Credito "
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
            Height          =   630
            Left            =   120
            TabIndex        =   105
            Top             =   3220
            Width           =   8655
            Begin VB.TextBox Txtinteres 
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
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   6240
               TabIndex        =   107
               Top             =   240
               Visible         =   0   'False
               Width           =   920
            End
            Begin VB.TextBox TxtMora 
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
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   7680
               TabIndex        =   106
               Top             =   240
               Visible         =   0   'False
               Width           =   920
            End
            Begin Spinner.uSpinner spnCuotas 
               Height          =   315
               Left            =   2750
               TabIndex        =   108
               Top             =   240
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               Max             =   350
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
            Begin Spinner.uSpinner SpnPlazo 
               Height          =   315
               Left            =   4560
               TabIndex        =   109
               Top             =   240
               Width           =   645
               _ExtentX        =   1138
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
            Begin SICMACT.EditMoney txtMonSug 
               Height          =   315
               Left            =   600
               TabIndex        =   110
               Top             =   235
               Width           =   1500
               _extentx        =   2302
               _extenty        =   556
               font            =   "frmCredSugerencia_NEW.frx":0538
               text            =   "0"
               enabled         =   -1
            End
            Begin VB.TextBox txtInteresTasa 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   315
               Left            =   6240
               TabIndex        =   182
               Top             =   240
               Visible         =   0   'False
               Width           =   920
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Cuotas:"
               Height          =   195
               Left            =   2210
               TabIndex        =   117
               Top             =   300
               Width           =   540
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Monto:"
               Height          =   195
               Left            =   80
               TabIndex        =   116
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Plazo (Dias):"
               Height          =   195
               Left            =   3680
               TabIndex        =   115
               Top             =   300
               Width           =   885
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Tasa Comp.:"
               Height          =   195
               Left            =   5340
               TabIndex        =   114
               Top             =   285
               Width           =   900
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "T.M.:"
               Height          =   195
               Left            =   7275
               TabIndex        =   113
               Top             =   285
               Width           =   375
            End
            Begin VB.Label LblInteres 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   6240
               TabIndex        =   112
               Top             =   240
               Width           =   915
            End
            Begin VB.Label LblMora 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
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
               Height          =   315
               Left            =   7680
               TabIndex        =   111
               Top             =   240
               Width           =   915
            End
         End
         Begin VB.Frame fratipodes 
            Caption         =   "Desembolso"
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
            Height          =   645
            Left            =   120
            TabIndex        =   97
            Top             =   4680
            Width           =   1515
            Begin VB.ComboBox cmbDesemb 
               Height          =   315
               ItemData        =   "frmCredSugerencia_NEW.frx":0564
               Left            =   120
               List            =   "frmCredSugerencia_NEW.frx":056E
               Style           =   2  'Dropdown List
               TabIndex        =   186
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton Optdesemb 
               Caption         =   "&Parcial"
               Height          =   285
               Index           =   1
               Left            =   240
               TabIndex        =   99
               Top             =   330
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.OptionButton Optdesemb 
               Caption         =   "&Total"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   98
               Top             =   225
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VB.Frame FraCalendario 
            Caption         =   "Calendario"
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
            Height          =   645
            Left            =   120
            TabIndex        =   94
            Top             =   4650
            Visible         =   0   'False
            Width           =   1395
            Begin VB.ComboBox cmbTipoCalend 
               Height          =   315
               ItemData        =   "frmCredSugerencia_NEW.frx":05B8
               Left            =   120
               List            =   "frmCredSugerencia_NEW.frx":05C2
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton OptTipoCalend 
               Caption         =   "&Fijo"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   96
               Top             =   225
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.OptionButton OptTipoCalend 
               Caption         =   "Dina&mico"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   95
               Top             =   480
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin VB.Frame FraTpoCalend 
            Caption         =   "Tipos de Calendario"
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
            Height          =   645
            Left            =   1680
            TabIndex        =   90
            Top             =   4680
            Width           =   2985
            Begin VB.CommandButton cmdMIVIVIENDA 
               Caption         =   "Bono Buen Pagador"
               Height          =   375
               Left            =   1320
               TabIndex        =   189
               Top             =   220
               Width           =   1575
            End
            Begin VB.CheckBox ChkMiViv 
               Caption         =   "Mi Vivienda"
               Height          =   360
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Width           =   1440
            End
            Begin VB.CheckBox ChkTrabajadores 
               Caption         =   "Trabajadores y Funcionarios"
               Height          =   240
               Left            =   120
               TabIndex        =   92
               Top             =   840
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.CheckBox ChkCuotaCom 
               Caption         =   "Cuota Comodin"
               Height          =   360
               Left            =   120
               TabIndex        =   91
               Top             =   240
               Visible         =   0   'False
               Width           =   1140
            End
         End
         Begin VB.Frame FraGastos 
            Caption         =   " Gastos Seguro Desgrav. "
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
            Height          =   600
            Left            =   120
            TabIndex        =   64
            Top             =   5385
            Width           =   8660
            Begin VB.ComboBox cboRepDesgrav 
               Height          =   315
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   220
               Width           =   6735
            End
            Begin VB.OptionButton OptTipoGasto 
               Caption         =   "&Fijo"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   66
               Top             =   270
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.OptionButton OptTipoGasto 
               Caption         =   "&Variable"
               Height          =   285
               Index           =   1
               Left            =   360
               TabIndex        =   67
               Top             =   225
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Rep. Desgravamen:"
               Height          =   195
               Left            =   240
               TabIndex        =   68
               Top             =   255
               Width           =   1425
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   " Exposición "
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
            Height          =   680
            Left            =   4815
            TabIndex        =   60
            Top             =   4680
            Width           =   3945
            Begin VB.TextBox txtExpAntMax 
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
               Left            =   2280
               MaxLength       =   9
               TabIndex        =   172
               Top             =   240
               Width           =   1545
            End
            Begin VB.CheckBox chkExpuestoRCC 
               Caption         =   "Expuesto RCC"
               Height          =   195
               Left            =   2280
               TabIndex        =   62
               Top             =   170
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.CheckBox chkVAC 
               Caption         =   "Credito VAC"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2280
               TabIndex        =   61
               Top             =   380
               Width           =   1335
            End
            Begin VB.Label lblExpAntMax 
               AutoSize        =   -1  'True
               Caption         =   "Exposición Anterior Máxima :"
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   285
               Width           =   2025
            End
         End
         Begin VB.Frame frmCtaAho 
            Caption         =   " Desembolso a Terceros "
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
            Height          =   800
            Left            =   4680
            TabIndex        =   56
            Top             =   6120
            Width           =   4095
            Begin VB.CommandButton cmdVerTercero 
               Caption         =   "Titular"
               Height          =   380
               Left            =   3375
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   260
               Width           =   600
            End
            Begin VB.CommandButton cmdCtaAho 
               Caption         =   "Buscar"
               Height          =   380
               Left            =   2685
               TabIndex        =   57
               Top             =   260
               Width           =   620
            End
            Begin SICMACT.ActXCodCta actxCtaAho 
               Height          =   435
               Left            =   -960
               TabIndex        =   58
               Top             =   260
               Width           =   3705
               _extentx        =   6747
               _extenty        =   556
               texto           =   ""
               enabledcmac     =   -1
               enabledcta      =   -1
               enabledprod     =   -1
               enabledage      =   -1
            End
            Begin VB.Label lblTercero 
               Height          =   375
               Left            =   5040
               TabIndex        =   59
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame FraTipoCuota 
            Caption         =   " Tipo Cuota "
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
            Height          =   645
            Left            =   120
            TabIndex        =   100
            Top             =   4650
            Visible         =   0   'False
            Width           =   1635
            Begin VB.ComboBox cmbTipoCuota 
               Height          =   315
               ItemData        =   "frmCredSugerencia_NEW.frx":0606
               Left            =   120
               List            =   "frmCredSugerencia_NEW.frx":0616
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton opttcuota 
               Caption         =   "&Cuota Libre"
               Height          =   255
               Index           =   3
               Left            =   1440
               TabIndex        =   104
               Top             =   450
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.OptionButton opttcuota 
               Caption         =   "D&ecreciente"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   103
               Top             =   465
               Visible         =   0   'False
               Width           =   1180
            End
            Begin VB.OptionButton opttcuota 
               Caption         =   "C&reciente"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   102
               Top             =   225
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.OptionButton opttcuota 
               Caption         =   "&Fija"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   101
               Top             =   225
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.CheckBox chkCuotaBalon 
            Caption         =   "Cuotas con Periodo de Gracia con Pago de Intereses"
            Height          =   195
            Left            =   4800
            TabIndex        =   175
            Top             =   8280
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.TextBox txtCuotaBalon 
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
            Left            =   8400
            MaxLength       =   3
            TabIndex        =   174
            Top             =   8145
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label lblSolModif 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   171
            Top             =   8475
            Width           =   8655
         End
         Begin VB.Label Label9 
            Caption         =   "Solicitud de Modificación del nivel de aprobación:"
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
            TabIndex        =   170
            Top             =   8265
            Width           =   4575
         End
         Begin VB.Label Label3 
            Caption         =   "Comentarios del Analista:"
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
            TabIndex        =   168
            Top             =   6075
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Datos Solicitados "
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
         Height          =   1875
         Left            =   -74880
         TabIndex        =   29
         Top             =   2400
         Width           =   8895
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Analista :"
            Height          =   195
            Left            =   3840
            TabIndex        =   44
            Top             =   1420
            Width           =   645
         End
         Begin VB.Label lblanalista 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4560
            TabIndex        =   43
            Top             =   1380
            Width           =   4095
         End
         Begin VB.Label lblplazosol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7920
            TabIndex        =   42
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Dias) :"
            Height          =   195
            Left            =   6960
            TabIndex        =   41
            Top             =   1020
            Width           =   930
         End
         Begin VB.Label lblcuotassol 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4800
            TabIndex        =   40
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "No Cuotas :"
            Height          =   195
            Left            =   3855
            TabIndex        =   39
            Top             =   1025
            Width           =   840
         End
         Begin VB.Label lbldescre 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1680
            TabIndex        =   38
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Destino Credito :"
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   1420
            Width           =   1170
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   1680
            TabIndex        =   35
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   1025
            Width           =   540
         End
         Begin VB.Label lbltProd 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1680
            TabIndex        =   33
            Top             =   240
            Width           =   6975
         End
         Begin VB.Label Label15 
            Caption         =   "Producto :"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Sub Producto :"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   680
            Width           =   1335
         End
         Begin VB.Label lblSubProd 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1680
            TabIndex        =   30
            Top             =   600
            Width           =   6975
         End
         Begin VB.Label lblmoneda 
            BackColor       =   &H80000004&
            Height          =   300
            Left            =   2835
            TabIndex        =   36
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Datos de Cliente "
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
         Height          =   1650
         Left            =   -74880
         TabIndex        =   18
         Top             =   600
         Width           =   8895
         Begin VB.ComboBox CboPersCiiu 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   990
            Width           =   7380
         End
         Begin VB.CommandButton cmdrelac 
            Caption         =   "&Relaciones"
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   19
            ToolTipText     =   "Mostrar Relaciones del Credito"
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label lblPersCIIU 
            AutoSize        =   -1  'True
            Caption         =   "CIIU :"
            Height          =   195
            Left            =   165
            TabIndex        =   28
            Top             =   1030
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cliente :"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   280
            Width           =   570
         End
         Begin VB.Label LblNat 
            AutoSize        =   -1  'True
            Caption         =   "Doc.Identidad :"
            Height          =   195
            Left            =   165
            TabIndex        =   26
            Top             =   650
            Width           =   1095
         End
         Begin VB.Label lbljur 
            AutoSize        =   -1  'True
            Caption         =   "RUC :"
            Height          =   195
            Left            =   3840
            TabIndex        =   25
            Top             =   645
            Width           =   435
         End
         Begin VB.Label lblcod 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1305
            TabIndex        =   24
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label lblnom 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3120
            TabIndex        =   23
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label LblDni 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1305
            TabIndex        =   22
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label LblRuc 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4485
            TabIndex        =   21
            Top             =   600
            Width           =   1680
         End
      End
      Begin VB.Frame fraVinculosEmpresas 
         Caption         =   "Empresas Vinculadas"
         ForeColor       =   &H00000080&
         Height          =   2775
         Left            =   -74880
         TabIndex        =   14
         Top             =   660
         Width           =   8925
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   5685
            TabIndex        =   16
            Top             =   2280
            Width           =   855
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   6600
            TabIndex        =   15
            Top             =   2280
            Width           =   855
         End
         Begin SICMACT.FlexEdit grdEmpVinculados 
            Height          =   2055
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   7335
            _extentx        =   13070
            _extenty        =   3466
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            visiblepopmenu  =   -1
            encabezadosnombres=   "#-Codigo-Nombre-Relacion-Monto-CtaAbono-P"
            encabezadosanchos=   "250-1700-3500-1500-1200-1800-0"
            font            =   "frmCredSugerencia_NEW.frx":06AF
            font            =   "frmCredSugerencia_NEW.frx":06DB
            font            =   "frmCredSugerencia_NEW.frx":0707
            font            =   "frmCredSugerencia_NEW.frx":0733
            font            =   "frmCredSugerencia_NEW.frx":075F
            fontfixed       =   "frmCredSugerencia_NEW.frx":078B
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-1-X-3-4-5-X"
            textstylefixed  =   4
            listacontroles  =   "0-1-0-3-0-1-0"
            encabezadosalineacion=   "C-L-L-L-R-L-C"
            formatosedit    =   "0-0-0-0-2-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   255
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   -74880
         TabIndex        =   9
         Top             =   4140
         Width           =   7605
         Begin SICMACT.EditMoney txtTasacion 
            Height          =   285
            Left            =   960
            TabIndex        =   10
            Top             =   210
            Width           =   1215
            _extentx        =   2143
            _extenty        =   503
            font            =   "frmCredSugerencia_NEW.frx":07B9
            appearance      =   0
            text            =   "0.00"
            enabled         =   -1
         End
         Begin VB.Label Label6 
            Caption         =   "Tasación:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblComisionEC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   285
            Left            =   4560
            TabIndex        =   12
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label26 
            Caption         =   "Comisión Estruc.Caja:"
            Height          =   255
            Left            =   2880
            TabIndex        =   11
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame fraOpGarantia 
         Caption         =   "Cta. Garantía (Operador)"
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   -74880
         TabIndex        =   4
         Top             =   3420
         Width           =   7605
         Begin SICMACT.TxtBuscar txtCtaGarantia 
            Height          =   345
            Left            =   960
            TabIndex        =   5
            Top             =   250
            Width           =   2475
            _extentx        =   4366
            _extenty        =   609
            appearance      =   1
            appearance      =   1
            font            =   "frmCredSugerencia_NEW.frx":07E5
            appearance      =   1
            editflex        =   -1
         End
         Begin SICMACT.EditMoney txtMontoGarantia 
            Height          =   285
            Left            =   4560
            TabIndex        =   6
            Top             =   280
            Width           =   1215
            _extentx        =   2143
            _extenty        =   503
            font            =   "frmCredSugerencia_NEW.frx":0811
            appearance      =   0
            text            =   "0.00"
            enabled         =   -1
         End
         Begin VB.Label Label27 
            Caption         =   "Monto:"
            Height          =   255
            Left            =   3840
            TabIndex        =   8
            Top             =   310
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   310
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmCredSugerencia_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredSugerencia_NEW
'***     Descripcion:       Realiza el proceso de Sugerencia de un Credito
'***     Creado por:        JUEZ
'***     Fecha-Tiempo:         17/12/2012 10:18:30 AM
'*****************************************************************************************

Option Explicit
Dim RLinea As ADODB.Recordset
'Dim RLinea2 As ADODB.Recordset
Dim MatGracia As Variant
Dim bGraciaGenerada As Boolean
Dim vnTipoGracia As Integer
Dim MatCalend As Variant
Dim MatCalend_2 As Variant
Dim MatDesemb As Variant
Dim MatDesPar() As String
Dim MatrizCal() As String
Dim MatGastos As Variant
Dim nNroTransac As Long
Dim bDesembParcialGenerado As Boolean
Dim nEstadoActual As Integer
Dim vbInicioCargaDatos As Boolean
Dim vnTipoCarga As lSugerTipoActualizacionNew
Dim nTipoGracia As Integer
Dim MatCredVig As Variant
'**** PEAC 20080412
Private oPersona As UPersona_Cli
Private MatFuentes As Variant
'**ALPA**18/04/2008************
Private MatFuentesF As Variant
'***End*************************
Private MatFteFecEval As Variant

'-------------------------------------
Dim nPersFIDIngCliActual As Double
Dim cPersFIMonedaActual As String
'-------------------------------------

Public Enum lSugerTipoActualizacionNew
    lSugerTipoActRegistroNew = 1
    lSugerTipoActConsultarNew = 2
End Enum

'Manejo de Exposicion RCC
Dim bControlRCC As Boolean
Dim nSaldoDisponible As Double
'Actualizacion de Filtros de Lineas de Credito
Dim bBuscarLineas As Boolean

Dim bEsRefinanciado As Boolean 'DAOR 20070407
Dim fnPersPersoneria As Integer 'DAOR 20071218
Dim objPista As COMManejador.Pista
Dim nActualizaTipoCred As Integer
Dim sTipoProdCod As String
Dim sSTipoProdCod As String
Dim nMostrarLineaCred As Integer
Dim nPorcCEC As Double 'BRGO 20111111 Porcentaje de Comisión
Dim nComisionEC As Double 'BRGO 20111111 Monto Total Empresas Afiliadas Ecotaxi}
Dim sPersOperador, sPersOperadorNombre As String 'BRGO 20111111
Dim oTipoCambio As nTipoCambio 'BRGO 20111111
Dim nTC As Double 'BRGO 20111111
Dim bLeasing As Boolean 'ALPA 20111209
Dim lnTasaPeriodoLeasing As Double 'ALPA 20111209
Dim nValorDiaGracia As Integer
Dim fbMicroseguro As Boolean 'WIOR 20120517
Dim fnMicroseguro As Integer 'WIOR 20120517
Dim fbMultiriesgo  As Boolean 'WIOR 20120517
Dim nAgenciaCredEval As Integer 'JUEZ 20120907
Dim nVerifCredEval As Integer 'JUEZ 20120914
Dim bCalendGen As Boolean
Dim bGastosGen As Boolean
Private nNumGastos As Integer
Dim sCliCtaAho As String
Dim rsExonera As ADODB.Recordset
Dim bCorrigeSug As Boolean
Dim bVerDetalle As Boolean
Dim bVerifNivApr As Boolean
Dim lnCSP As Integer 'ALPA 20141126
Dim oRsVerEntidades As ADODB.Recordset 'ALPA20141021***
Dim lnCantidadVerEntidades As Integer 'ALPA20141021***
Dim bCantidadVerEntidadesCmac As Integer 'ALPA20141021***
Dim nLogicoVerEntidades As Integer 'ALPA20141201
Dim lnColocDestino As Integer 'ALPA20141201
Dim lnMostrarCSP As Integer 'ALPA 20141126
Dim lnColocCondicion As Integer '20141230
'ALPA 20150114**************************************
Dim lnTasaInicial As Currency
Dim lnTasaFinal As Currency
Dim lnCampanaId As Integer
Dim lnLogicoBuscarDatos As Integer
Dim lnCliPreferencial As Integer
'***************************************************

'JOEP ERS007-2018 20180210**************************************
Dim lnTasaGraciaInicial As Currency
Dim lnTasaGraciaFinal As Currency
'***************************************************

Dim bCheckList As Boolean 'RECO20150421 ERS010-2015
Dim vMatriz As Variant
'WIOR 20151223 ***
Private fbMIVIVIENDA As Boolean
Private fArrMIVIVIENDA As Variant
Private fbDatosCargados As Boolean
'WIOR FIN ********
Dim fvGravamen() As tGarantiaGravamen 'EJVG20150513
Dim fbSalirCargaDatos As Boolean 'EJVG20151104
'WIOR 20160224 ***
Private fnTasaSegDes As Double
Private fnCantAfiliadosSegDes As Integer
'WIOR ************
Dim fbEsAmpliado As Boolean 'EJVG20160512
Dim fbAutoCalfCPP As Boolean 'RECO20160628 ERS002-2016
Dim fnMontoExpxRiesgoUnico_NEW As Double 'EJVG20160713
Dim fnMontoExpEsteCred_NEW As Double 'EJVG20160713
Dim fbEliminarEvaluacion As Boolean 'EJVG20160713
Dim fvListaCompraDeuda() As TCompraDeuda '**ARLO20171127 ERS070 - 2017
Dim fvListaAguaSaneamiento() As TAguaSaneamiento 'EAAS20180801 ERS054-2018
Dim bValidaCargaSugerenciaAguaSaneamiento As Integer 'EAAS20180801 ERS054-2018
Dim bValidaCargaSugerenciaCreditoVerde As Integer 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim fvListaCreditoVerde() As TCreditoVerde 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nMontoCreditoVariable As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nCentinela As Integer 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nSumaAguaSaneamiento As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nSumaCreditoVerde As Double 'EAAS20191401 SEGUN 018-GM-DI_CMACM
Dim nDestino As Integer '**ARLO20171127 ERS070 - 2017

Dim fbTieneTasaGracia As Boolean 'LUCV20180601, Según ERS022-2018
Dim fbNoTieneTasaGracia As Boolean  'LUCV20180601, Según ERS022-2018
Private MatCalendSegDes As Variant 'LUCV20180601, Según ERS022-2018

'JOEP20181001 CP ERS042 - 2018
Dim cValorIni As String
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018
Dim MontoMin As Currency, MontoMax As Currency
Dim CuotaMin As Integer, CuotaMax As Integer
Dim PlazoMin As Integer, PlazoMax As Integer
Dim GraciaMin As Integer, GraciaMax As Integer
Dim nMatMontoPre As Variant
Dim nMatAporteLleno As Boolean
Dim sCodPers As String
Dim bEntrotxtMontoSeg As Boolean
'JOEP20181001 CP ERS042 - 2018

Public Sub Sugerencia(ByVal pnTipoCarga As lSugerTipoActualizacionNew, Optional ByVal pbLeasing As Boolean = False)
    vnTipoCarga = pnTipoCarga
    bLeasing = pbLeasing
    bCorrigeSug = False
    bVerDetalle = False
    If bLeasing = True Then
        Me.Caption = "Sugerencia de Arrendamiento Financiero"
        ActxCta.texto = "Operación"
        Frame3.Caption = "Datos de la Operación"
    End If
    ReDim vMatriz(3, 0) 'RECO20150603
    Me.Show 1
End Sub

Private Sub HabilitaPermiso()
    Select Case vnTipoCarga
        Case lSugerTipoActRegistroNew
            CmdGrabar.Enabled = True
            CmdCalend.Enabled = True
            CmdGastos.Enabled = True
            'cmdEvaluacion.Enabled = True
            cmdCheckList.Enabled = True
            ActxCta.Enabled = True
            CmdDesembolsos.Enabled = False
            'Cmblincre.Enabled = True
            txtBuscarLinea.Enabled = True
            
            txtMonSug.Enabled = True
            'txtExpAntMax.Enabled = True '*** PEAC 20080412
            spnCuotas.Enabled = True
            'spnPlazo.Enabled = True 'Comentado Por MAVM 25102010
            FraTipoCuota.Enabled = True
            Frame5.Enabled = True
            'fratipodes.Enabled = True 'LUCV20180601, Comentó Según ERS022-2018
            fraGracia.Enabled = True
            fraVinculosEmpresas.Enabled = True 'BRGO 20111103
        Case lSugerTipoActConsultarNew
            CmdGrabar.Enabled = False
            CmdCalend.Enabled = False
            CmdGastos.Enabled = False
            'cmdEvaluacion.Enabled = False
            cmdCheckList.Enabled = False
            ActxCta.Enabled = False
            CmdDesembolsos.Enabled = False
            'Cmblincre.Enabled = False
            txtBuscarLinea.Enabled = False

            txtMonSug.Enabled = False
            'txtExpAntMax.Enabled = False
            spnCuotas.Enabled = False
            SpnPlazo.Enabled = False
            FraTipoCuota.Enabled = False
            Frame5.Enabled = False
            'fratipodes.Enabled = False 'LUCV20180601, Comentó Según ERS022-2018
            fraGracia.Enabled = False
            fraVinculosEmpresas.Enabled = False 'BRGO 20111103
    End Select
End Sub
'JUEZ 20121210 ****************************************************************************************
Public Sub InicioVerDetalleSug(ByVal psCtaCod As String, ByVal pbCorrigeSug As Boolean)
    
    ActxCta.NroCuenta = psCtaCod
    bCorrigeSug = pbCorrigeSug
    bVerDetalle = True
    Call ActxCta_KeyPress(13)

    vbInicioCargaDatos = True
    vnTipoCarga = lSugerTipoActRegistroNew
    Call HabilitaPermiso
    
    ActxCta.Enabled = False
    cmdbuscar.Enabled = False
    cmdcancelar.Enabled = False
        
    FraDatos.Enabled = pbCorrigeSug
    CmdGrabar.Enabled = pbCorrigeSug
    CmdCalend.Enabled = pbCorrigeSug
    CmdGarantia.Enabled = pbCorrigeSug
    CmdDesembolsos.Enabled = False
    CmdGastos.Enabled = pbCorrigeSug
    'cmdEvaluacion.Enabled = pbCorrigeSug
    
    bLeasing = False
    
     If Mid(sSTipoProdCod, 1, 1) = "7" Then
        cmdActTipoCred.Visible = False
     Else
        cmdActTipoCred.Visible = True
     End If
     
     'If nAgenciaCredEval = 0 Then
     '    If sSTipoProdCod = "703" Then
     '       cmdSeleccionaFuente.Enabled = False
     '       cmdFuentes.Enabled = False
     '       'Label13.Enabled = False
     '   Else
     '       cmdSeleccionaFuente.Enabled = True
     '       cmdFuentes.Enabled = True
     '       'Label13.Enabled = True
     '   End If
    'Else
    '    cmdSeleccionaFuente.Enabled = False
    '    cmdFuentes.Enabled = False
    '    'Label13.Enabled = False
    'End If
    
     Frame3.Enabled = True
    'If nAgenciaCredEval = 0 Then
        Me.Show 1
    'Else
    '    If nVerifCredEval = 1 Then
    '        Me.Show 1
    '    End If
    'End If
End Sub
'END JUEZ ********************************************************************************************

'27-12
'Private Sub ImprimirResumenComite()
'Dim oNCredDoc As COMNCredito.NCOMCredDoc
'Dim oPrevio As Previo.clsPrevio
'
'    Set oPrevio = New Previo.clsPrevio
'    Set oNCredDoc = New COMNCredito.NCOMCredDoc
'    oPrevio.Show oNCredDoc.ImprimeResumenComite(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, gsNomCmac), "Resumen de Comite"
'    Set oNCredDoc = Nothing
'    Set oPrevio = Nothing
'End Sub

Public Sub InicioCargaDatos(ByVal psCtaCod As String, Optional ByVal pbLeasing As Boolean = False, Optional ByVal pbLeasingInicio As Boolean = False)
    fbSalirCargaDatos = False 'EJVG20151104
    bValidaCargaSugerenciaAguaSaneamiento = 0 'EAAS20180907
    bValidaCargaSugerenciaCreditoVerde = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    bCorrigeSug = False
    bVerDetalle = False
    ActxCta.NroCuenta = psCtaCod
    If pbLeasingInicio = True Then
        Call ActxCta_KeyPress(13)
    End If
    vbInicioCargaDatos = True
    bLeasing = pbLeasing
    If bLeasing = True Then
        Me.Caption = "Sugerencia de Arrendamiento Financiero"
        Me.Label11.Caption = "Rep."
    End If
    
    'MADM 20100517
    'ALPA 20100609 B2***********************************
     'If Mid(psCtaCod, 6, 3) = "302" Then
     If Mid(sSTipoProdCod, 1, 1) = "7" Then
        cmdActTipoCred.Visible = False
     Else
        cmdActTipoCred.Visible = True
     End If
     
     ''** JUEZ 20120907 ******************************************
       'If sSTipoProdCod = "703" Then
         '**************************************************
       '     cmdSeleccionaFuente.Enabled = False
       '    cmdFuentes.Enabled = False
       '     Label13.Enabled = False
       ' Else
       '     cmdSeleccionaFuente.Enabled = True
       '     cmdFuentes.Enabled = True
       '     Label13.Enabled = True
       ' End If
     
     'If nAgenciaCredEval = 0 Then
     '    If sSTipoProdCod = "703" Then
     '       cmdSeleccionaFuente.Enabled = False
     '       cmdFuentes.Enabled = False
     '       'Label13.Enabled = False
     '   Else
     '       cmdSeleccionaFuente.Enabled = True
     '       cmdFuentes.Enabled = True
     '       'Label13.Enabled = True
     '   End If
    'Else
    '    cmdSeleccionaFuente.Enabled = False
    '    cmdFuentes.Enabled = False
    '    'Label13.Enabled = False
    'End If
    ''** END JUEZ ***********************************************
    'JOEP20190201 CP
    If cmbSubProd.Text <> "" Then
        Call bValiExistCheckList
    End If
    'JOEP20190201 CP
     Frame3.Enabled = True
    'END MADM
    'If nAgenciaCredEval = 0 Then
        Me.Show 1
    'Else
    '    If nVerifCredEval = 1 Then
    '        Me.Show 1
    '    End If
    'End If
End Sub

Private Function ValidaDatosGrabar(ByVal psValorLinea As String) As Boolean
Dim oNCredito As COMNCredito.NCOMCredito
Dim sValor As String
Dim nValor As Double

'Dim bNecesitaPoliza As Boolean
'Dim oNCredito As COMNCredito.NCOMCredito
'WIOR 20120511******************************************************************************
Dim oDPersona As COMDPersona.DCOMPersona
Dim oCreditoBD As COMDCredito.DCOMCredActBD
Dim oCredito As COMDCredito.DCOMCredito
Dim rsPersona As ADODB.Recordset
Dim rsPersonaF As ADODB.Recordset
Dim rsCredito As ADODB.Recordset
Dim rsCreditoBD As ADODB.Recordset
Dim nEdad As Integer
Dim nEdadF As Integer
Dim nTiempo As Double
Dim dFuturo As Date
'WIOR FIN **********************************************************************************
Dim lsMsg As String 'EJVG20160512
Dim i As Integer
Dim bExisteExonera As Boolean
Dim nTipoEval As Integer
''WIOR 201207024 SEGUN OYP-RFC066-2012********************************************************************
'Dim sCodPersonas As String
'Dim oRegPersona As COMDPersona.DCOMPersona
'Set oRegPersona = New COMDPersona.DCOMPersona
'Dim bVinculados As Boolean
'bVinculados = False
'Dim oCodPersonas As COMDPersona.DCOMPersona
'Set oCodPersonas = New COMDPersona.DCOMPersona
'Dim rsCodPersonas As ADODB.Recordset
'Dim rsPersonasVin As ADODB.Recordset
'Dim CantVinculados As Long
'CantVinculados = 0
'Dim Recorrido As Long
'Recorrido = 0
'Dim nRiesgo As Integer
'nRiesgo = 0
'Dim SaldoFinal As Double
''WIOR FIN **********************************************************************************
'Dim nMontoSug As Double 'FRHU 20140329 ERS042-2014 RQ14177-RQ14178
'Dim bVerificaDPF As Boolean 'WIOR 20140726
'bVerificaDPF = False 'WIOR 20140726

    
    '**ARLO20171127 ERS070 - 2017
    nDestino = Trim(Right(Me.cmbDestino.Text, 5)) 'ARLO20171113
        
    If (nDestino = 14 And Not bEsRefinanciado) Then 'ARLO20180322 'NOT ARLO20180322
    
    '**ARLO20180315 INICIO ERS070 - 2017 ANEXO 01
    If UBound(fvListaCompraDeuda) <= 0 Then
            MsgBox "Para elegir este destino (Reestructuración de Pasivos)," & Chr(13) & _
            "necesitas ingresar las IFIS a comprar en la Solicitud del Crédito.", vbInformation, "Alerta"
            ValidaDatosGrabar = False
            Exit Function
    End If
    '**ARLO20180315 FIN ERS070 - 2017 ANEXO 01
    
    Dim nCantidad As Integer
    Dim maxValue As Double
    Dim lvListaCompraDeudaNew(1) As TCompraDeuda
    Dim oDCreditos As COMDCredito.DCOMCreditos
    Dim rsRCC As ADODB.Recordset
    Dim nCantCompraIFIS As Integer
    '**ARLO20180604 INICIO ERS070 - 2017 ANEXO 01
    Dim oTC  As New COMDConstSistema.NCOMTipoCambio
    Dim nTpoC As Double
    Dim nMontoSol, nSaldoComp, nDesem As Double
    nTpoC = CDbl(oTC.EmiteTipoCambio(gdFecSis, TCFijoDia))
    nMontoSol = (val(txtMonSug.Text))

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
    '**ARLO20180604 FIN ERS070 - 2017 ANEXO 01
        If Not bEsRefinanciado Then '**ARLO20180317 ERS070 - 2017 ANEXO 02
            Set oDCreditos = New COMDCredito.DCOMCreditos
            Set rsRCC = oDCreditos.ObtenerCalificacionRCC(Trim(lblcod.Caption), Me.ActxCta.NroCuenta)
            
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("N0000146", Trim(Right(Me.cmbSubProd.Text, 5))) Then
            'If ((CInt(Trim(Right(Me.cmbSubProd.Text, 5)))) <> 704) Then '**ARLO20180317 ERS070 - 2017 ANEXO 02
                If CInt(spnCuotas.valor) > nCantidad Then
                    MsgBox "El número de cuotas debe ser menor o igual a " & nCantidad, vbInformation, "Aviso"
                    spnCuotas.SetFocus
                    ValidaDatosGrabar = False
                    Exit Function
                End If
                If Not (rsRCC.EOF And rsRCC.BOF) Then
                        'If ((CInt(Trim(Right(Me.cmbSubProd.Text, 5)))) <> 704) Then
                            If (rsRCC!Calif_0 <> 100) Then
                            MsgBox "El Cliente no tiene calificación 100% normal. ", vbInformation, "Aviso"
                            Exit Function
                            End If
                        End If
                        'End If
                End If
            End If '**ARLO20180317 ERS070 - 2017 ANEXO 02
            If Not (rsRCC.EOF And rsRCC.BOF) Then
                nCantCompraIFIS = rsRCC!Can_Ents - UBound(fvListaCompraDeuda)
                If (nCantCompraIFIS + 1) > 3 Then '**ARLO20180317 ERS070 - 2017 --ANEXO 02
                MsgBox "El Cliente no cumple con los requisitos de compra de deuda, máximo debe contar con 3 IFIS (incluyendo Caja Maynas).", vbInformation, "Alerta"
                Exit Function
            End If
            If (CInt(Me.txtPerGra) > 30) Then
                MsgBox "El periodo de gracia no debe se mayor que 30 días. ", vbInformation, "Aviso"
                Exit Function
            End If
            '**ARLO20180315 INICIO ERS070 - 2017 ANEXO 01
            Dim Y As Integer
            Dim nTotalCompra As Double

            If Not bEsRefinanciado Then
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
                            Exit Function
                    End If
                End If
            End If
            Set oTC = Nothing
            '**ARLO20180315 FIN ERS070 - 2017 ANEXO 01
        End If '**ARLO20180317 ERS070 - 2017 ANEXO 02
    End If
    '**ARLO20171127 ERS070 - 2017
    
    
    VerificarFechaSistema Me, True 'EJVG20151020 -> validar fecha de sistema, en caso no apaguen sus PC los usuarios les sacará del sistema
If lnColocCondicion <> 4 Then 'JOEP20180301 CP
    'INICIO EAAS20180815
    Dim rsValDesAguaSaneamiento As ADODB.Recordset
    Dim obDCredValDesAguaSaneamiento As COMDCredito.DCOMCredito
    Set obDCredValDesAguaSaneamiento = New COMDCredito.DCOMCredito
    
    Set rsValDesAguaSaneamiento = obDCredValDesAguaSaneamiento.ValidadDestinoConsEmpAguaSaneamiento(CInt(Trim(Right(cmbProducto.Text, 5))), nDestino, lblcod.Caption)
    If (UBound(fvListaAguaSaneamiento) > 0) Then
        If Not (rsValDesAguaSaneamiento.EOF And rsValDesAguaSaneamiento.BOF) Then
            If rsValDesAguaSaneamiento!cMensaje <> "" Then
                MsgBox rsValDesAguaSaneamiento!cMensaje, vbInformation, "No podrá continuar"
                rsValDesAguaSaneamiento.Close
                Set obDCredValDesAguaSaneamiento = Nothing
               ValidaDatosGrabar = False
                Exit Function
            End If
        rsValDesAguaSaneamiento.Close
        Set obDCredValDesAguaSaneamiento = Nothing
        End If
    End If
    If (UBound(fvListaAguaSaneamiento) = 0 And nDestino = 26) Then
                        MsgBox "Ingrese el detalle del destino Agua y saneamiento", vbInformation, "Alerta"
                        ValidaDatosGrabar = False
                        Exit Function
    End If
        Dim nSumaTotalAguaSaneamiento As Double
        nSumaTotalAguaSaneamiento = 0
        Dim ixCD As Integer
        For ixCD = 1 To UBound(fvListaAguaSaneamiento)
            nSumaTotalAguaSaneamiento = nSumaTotalAguaSaneamiento + fvListaAguaSaneamiento(ixCD).nMontoS
        Next
        If (nSumaTotalAguaSaneamiento <> CDbl(txtMonSug.Text) And nDestino = 26) Then
                        MsgBox "La suma de los subdestinos de agua y saneamiento debe ser igual al monto solicitado", vbInformation, "Alerta"
                        
                        ValidaDatosGrabar = False
                        Exit Function
        End If
        
        If (nSumaTotalAguaSaneamiento > CDbl(txtMonSug.Text)) Then
                        MsgBox "La suma de los subdestinos de agua y saneamiento es mayor al monto solicitado", vbInformation, "Alerta"
                        
                        ValidaDatosGrabar = False
                        Exit Function
        End If
         'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
        Dim nSumaTotalCreditoVerde As Double
        nSumaTotalCreditoVerde = 0
        Dim ixCD2 As Integer
        For ixCD2 = 1 To UBound(fvListaCreditoVerde)
            nSumaTotalCreditoVerde = nSumaTotalCreditoVerde + fvListaCreditoVerde(ixCD2).nMontoS
        Next
        If (nSumaTotalCreditoVerde > CDbl(txtMonSug.Text)) Then
                        MsgBox "La suma de los subdestinos de Eco Ahorro es mayor al monto solicitado", vbInformation, "Alerta"
                        
                        ValidaDatosGrabar = False
                        Exit Function
        End If
        If (nSumaTotalAguaSaneamiento + nSumaTotalCreditoVerde > CDbl(txtMonSug.Text)) Then
                        MsgBox "La suma de los subdestinos es mayor al monto solicitado", vbInformation, "Alerta"
                        
                        ValidaDatosGrabar = False
                        Exit Function
        End If
        'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
        If (CInt(Me.txtPerGra) > 30 And nDestino = 26) Then
                MsgBox "El periodo de gracia no debe se mayor que 30 dias para el destino de Agua y Saneamiento. ", vbInformation, "Aviso"
                ValidaDatosGrabar = False
                Exit Function
            End If
'END EAAS20180815

End If 'JOEP20190115 CP
 'JOEP20190115 CP
        If Not CP_ValMensaje(1) Then
            ValidaDatosGrabar = False
            Exit Function
        End If
        
        If Not CP_ValidaRangoFecha(Trim(Right(cmbSubProd.Text, 9)), 3000, Right(cmbTipoPer.Text, 3), txtFechaFija.Text) Then
            txtFechaFija.SetFocus
            ValidaDatosGrabar = False
            Exit Function
        End If
    'JOEP20190114 CP

    'EJVG20160527 ***
    'If bCheckList = False And ActxCta.Prod <> "703" Then 'Comento JOEP20190125 CP
    If cmdCheckList.Enabled = True And bCheckList = False Then 'JOEP20190125 CP
        MsgBox "Debe registrar el CheckList", vbInformation, "Alerta"
        ValidaDatosGrabar = False
        Exit Function
    End If
    
    If frmCtaAho.Enabled = True Then
        Call actxCtaAho_KeyPress(13)
        If CmdCalend.Enabled = False Then
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If
    
'JOEP 201710 Acta201
Dim rsValidaPriFecPago As ADODB.Recordset
Set oCredito = New COMDCredito.DCOMCredito
Set rsValidaPriFecPago = oCredito.ValidaPriFecPago(CDate(TxtFecDesemb.Text), CDate(txtFechaFija.Text))
If Not (rsValidaPriFecPago.EOF And rsValidaPriFecPago.BOF) Then
    If rsValidaPriFecPago!cMensaje <> "" Then
        MsgBox rsValidaPriFecPago!cMensaje, vbInformation, "Aviso"
        ValidaDatosGrabar = False
        rsValidaPriFecPago.Close
        Exit Function
    End If
rsValidaPriFecPago.Close
End If
'JOEP 201710 ACTA201
    
'    If nAgenciaCredEval = 0 Then
'        If sSTipoProdCod <> "703" Then
'            If MatFuentesF(3, 1) <> "" Then
'                If MatFuentesF(3, 1) = "D" Then
'                    nTipoEval = 1
'                Else
'                    nTipoEval = 2
'                End If
'            Else
'                MsgBox "Seleccione una fuente de Ingreso.", vbInformation, "Aviso"
'                ValidaDatosGrabar = False
'                Exit Function
'            End If
'        End If
'    End If

    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000001", sSTipoProdCod) Then     '**END ARLO
    'If sSTipoProdCod = "517" Then
        If Me.grdEmpVinculados.rows < 4 Then
            MsgBox "El registro de Personas/Empresas vinculadas no está completo"
            cmdAgregar.SetFocus
            ValidaDatosGrabar = False
            Exit Function
        End If
        If nComisionEC + CCur(txtTasacion.Text) + CCur(lblComisionEC.Caption) <> CCur(txtMonSug.Text) Then
            MsgBox "La suma total de los montos distribuidos es " & Format(nComisionEC + CCur(txtTasacion.Text) + CCur(lblComisionEC.Caption), "#,##0.00") & " y es diferente al Monto sugerido"
            txtMonSug.SetFocus
            ValidaDatosGrabar = False
            Exit Function
        End If
        If txtCtaGarantia.Text = "" Then
            MsgBox "Debe seleccionar la cuenta de abono de garantía"
            txtCtaGarantia.SetFocus
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If
    If ActxCta.Prod = "152" Or ActxCta.Prod = "252" Or ActxCta.Prod = "352" Or ActxCta.Prod = "452" Or ActxCta.Prod = "552" Then
        MsgBox "No se olvide de asignar el calendario dinamico", vbInformation, "Aviso"
    End If
    Set objProducto = New COMDCredito.DCOMCredito
    If (txtMontoMivivienda.Text = 0# Or txtMontoMivivienda.Text = "") And objProducto.GetResultadoCondicionCatalogo("N0000002", sSTipoProdCod) Then     '**END ARLO
    'If (txtMontoMivivienda.Text = 0# Or txtMontoMivivienda.Text = "") And sSTipoProdCod = "801" Then
        MsgBox "No se olvide de asignar el valor de venta del credito MIVIVIENDA", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtMontoMivivienda.SetFocus
        ValidaDatosGrabar = False
        Exit Function
    End If
    Set objProducto = New COMDCredito.DCOMCredito
    If Round((CDbl(txtMontoMivivienda.Text) - ((CDbl(txtMontoMivivienda.Text) * 1) / 10)), 2) < CDbl(txtMonSug.Text) And objProducto.GetResultadoCondicionCatalogo("N0000003", sSTipoProdCod) Then     '**END ARLO
    'If Round((CDbl(txtMontoMivivienda.Text) - ((CDbl(txtMontoMivivienda.Text) * 1) / 10)), 2) < CDbl(txtMonSug.Text) And sSTipoProdCod = "801" Then
        MsgBox "MIVIVIENDA, no se olvide que el monto de la inicial no debe ser menor al 10% (" & Round((CDbl(txtMonSug.Text) * 10) / 9, 2) & ")", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtMontoMivivienda.SetFocus
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Trim(cmbTipoCredito.Text) = "" Then
        MsgBox "No se olvide de seleccionar el Tipo de Credito", vbInformation, "Aviso"
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Trim(cmbSubTipo.Text) = "" Then
        MsgBox "No se olvide de seleccionar el sub tipo de credito", vbInformation, "Aviso"
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Trim(cmbProducto.Text) = "" Then
        MsgBox "No se olvide de seleccionar el producto", vbInformation, "Aviso"
        cmbProducto.SetFocus
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Trim(cmbSubProd.Text) = "" Then
        MsgBox "No se olvide de seleccionar el sub producto", vbInformation, "Aviso"
        cmbSubProd.SetFocus
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Right(cmbTipoCredito.Text, 3) = gColCredCorpo Then
        If Trim(cmbInstitucionFinanciera.Text) = "" Then
            MsgBox "No se olvide de seleccionar la institucion financiera", vbInformation, "Aviso"
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If
    If Trim(cmbDestino.Text) = "" Then
        MsgBox "No se olvide de seleccionar el destino del credito", vbInformation, "Aviso"
        cmbDestino.SetFocus
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Trim(cmbCampana.Text) = "" Then
        MsgBox "No se olvide de seleccionar la campaña", vbInformation, "Aviso"
        cmbCampana.SetFocus
        ValidaDatosGrabar = False
        Exit Function
    End If
    If Right(cmbTipoCredito.Text, 3) = gColCredHipot Then
        If Trim(cmbDatoVivienda.Text) = "" Then
            MsgBox "No se olvide de seleccionar el Dato de la Vivienda", vbInformation, "Aviso"
            cmbDatoVivienda.SetFocus
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If

    '->***** LUCV20180601, Comentó según ERS022-2018
    Set rsExonera = IIf(frmCredSugExonera.feTiposExonera.rows - 1 > 0, frmCredSugExonera.feTiposExonera.GetRsNew(), Nothing) 'JOEP20200217 Mejora pase de Canlendario interativo
    'Set rsExonera = Nothing 'Comento JOEP20200217 Mejora pase de Canlendario interativo
    '<-***** Fin LUCV20180601

    bExisteExonera = False
    'RECO20160711 ERS002-2016****************************************
'    If ChkExoneracion.value = 1 Then
'        For i = 0 To rsExonera.RecordCount - 1
'            If rsExonera!Solicitar = 1 Then
'                bExisteExonera = True
'                Exit For
'            End If
'            rsExonera.MoveNext
'        Next i
'        If bExisteExonera = False Then
'            MsgBox "Debe escoger un tipo de exoneración, de lo contrario desactivar el check de Aprobacion de Exoneración", vbInformation, "Aviso"
'            ChkExoneracion.SetFocus
'            ValidaDatosGrabar = False
'            Exit Function
'        End If
'    End If
    'RECO FIN *********************************************************
    'END EJVG *******
        
    If Trim(Right(cmbTipoPer, 2)) = "0" Then
        If CDate(CDate(TxtFecDesemb.Text) + CDate(SpnPlazo.valor) + CDate(txtPerGra.Text)) <> CDate(txtFechaFija.Text) Then
            MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
            TxtFecDesemb.SetFocus
            ValidaDatosGrabar = False
            Exit Function
        End If
    Else
        If CDate(txtPerGra.Text) <> "0" Then
            If CDate(CDate(TxtFecDesemb.Text) + CDate(30) + CDate(txtPerGra.Text)) <> CDate(txtFechaFija.Text) Then
                MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
                TxtFecDesemb.SetFocus
                ValidaDatosGrabar = False
                Exit Function
            End If
        End If
    End If

    ValidaDatosGrabar = True
    If CmdDesembolsos.Enabled Then
        If UBound(MatDesPar) = 0 Then
            MsgBox "No se ha Generado el calendario de Desembolsos", vbInformation, "Aviso"
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If
    If UBound(MatrizCal) = 0 Then
        MsgBox "No Se ha Generado el Calendario de Pagos", vbInformation, "Aviso"
        ValidaDatosGrabar = False
        Exit Function
    End If
    
'JOEP ERS007-2018 20180210
 'If ActxCta.Prod = "703" And CInt(Trim(Right(cmbCampana.Text, 5))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TC
 Set objProducto = New COMDCredito.DCOMCredito
 If objProducto.GetResultadoCondicionCatalogo("N0000016", ActxCta.Prod) And CInt(Trim(Right(cmbCampana.Text, 5))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TC
     If Txtinteres = "" Then
     Else
        If Txtinteres >= lnTasaInicial And Txtinteres <= lnTasaFinal Then
        Else
            MsgBox "La T.C: esta fuera del Rango: Mínima : " & Format(lnTasaInicial, "#0.0000") & "  Máxima : " & Format(lnTasaFinal, "#0.0000"), vbInformation, "Aviso"
            Txtinteres.Text = Format(lnTasaFinal, "#0.0000")
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If
 End If
 
'->***** LUCV20180601, Comentó según ERS022-2018
' If chkGracia.value = 1 Then
'        If TxtTasaGracia.Visible Then
'             Set objProducto = New COMDCredito.DCOMCredito
'             If objProducto.GetResultadoCondicionCatalogo("N0000017", ActxCta.Prod) And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TG
'             'If ActxCta.Prod = "703" And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TG
'                If TxtTasaGracia = "" Then
'                Else
'                    If TxtTasaGracia >= lnTasaGraciaInicial And TxtTasaGracia <= lnTasaGraciaFinal Then
'                    Else
'                        MsgBox "La T.G: esta fuera del Rango: Mínima : " & Format(lnTasaGraciaInicial, "#0.0000") & "  Máxima : " & Format(lnTasaGraciaFinal, "#0.0000"), vbInformation, "Aviso"
'                        TxtTasaGracia.Text = Format(lnTasaGraciaFinal, "#0.0000")
'                        ValidaDatosGrabar = False
'                        Exit Function
'                    End If
'                End If
'             End If
'        End If
'End If
'<-***** Fin LUCV20180601
'JOEP ERS007-2018 20180210
    
    '04-05-2006
'    If bBuscarLineas = False Then
'        MsgBox "Debe elegir una nueva Linea de Crédito", vbInformation, "Aviso"
'        ValidaDatosGrabar = False
'        Exit Function
'    End If
'    If txtBuscarLinea.Text = "" Then
'        MsgBox "Debe elegir una Linea de Crédito", vbInformation, "Aviso"
'        ValidaDatosGrabar = False
'        Exit Function
'    End If
    
    'ARCV 30-10-2006
    'If SpnPlazo.valor > 0 And opttper(1).value Then
    If SpnPlazo.valor > 0 And Trim(Right(cmbTipoPer.Text, 2)) = "1" Then
        MsgBox "El Plazo indicado debe ser cero para los Periodos de Fecha Fija", vbInformation, "Mensaje"
        ValidaDatosGrabar = False
        Exit Function
    End If
    '---------------
    'ALPA 20160419********************************************
    If SpnPlazo.valor < 30 And Trim(Right(cmbTipoPer.Text, 5)) = 0 Then
        MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
         ValidaDatosGrabar = False
         Exit Function
    End If
    '*********************************************************
    '*************
    ' CMACICA_CSTS - 26/11/2003 - -------------------------------------------------------------------------------------
    'Valida que para los tipos de Credito CONSUMO - DESCUENTO POR PLANILLA, el Monto MAX. de la Cuota sea el 30%
    'de su ingreso familiar (sueldo neto) de la fuente de ingreso dependiente
    
'ARCV 30-01-2007
'    If CInt(Mid(ActxCta.NroCuenta, 6, 3)) = gColConsuDctoPlan Then
'       If Len(lblcuota.Caption) > 0 Then
'
'            Set oNCredito = New COMNCredito.NCOMCredito
'            Call oNCredito.ValidaDatosSugerencia(sValor, psValorLinea, gdFecSis, CInt(Mid(ActxCta.NroCuenta, 9, 1)), CDbl(lblcuota.Caption), CInt(cPersFIMonedaActual), _
'                                        nPersFIDIngCliActual, ActxCta.NroCuenta, CDbl(txtmonsug.Text), IIf(Optdesemb(0).value, 1, 2))
'
'            If sValor <> "" Then
'                MsgBox sValor, vbInformation, "Aviso"
'                If spnCuotas.Enabled Then spnCuotas.SetFocus
'                ValidaDatosGrabar = False
'                Set oNCredito = Nothing
'                Exit Function
'            End If
'            Set oNCredito = Nothing
'        End If
'   End If
    'EJVG20150705 ***
    ''ARCV 24-01-2007
    'Dim oPol As COMDCredito.DCOMPoliza
    'Set oPol = New COMDCredito.DCOMPoliza
    
    'bNecesitaPoliza = oPol.Poliza_para_Credito(ActxCta.NroCuenta, CDbl(txtMonSug.Text))
    ''ALPA 20120509*************************
    'If Not (sSTipoProdCod = "515" Or sSTipoProdCod = "516") Then
    '    If bNecesitaPoliza Then
    '        MsgBox "El credito necesita registro de Poliza." & vbCrLf & "Monto >= 15000 dolares.", vbInformation, "Mensaje"
    '    End If
    'End If
    'Set oPol = Nothing
    ''------------------
    'END EJVG *******
    '**DAOR 20071218 **********************************************
    If fnPersPersoneria > 1 Then
        If cboRepDesgrav.Text = "" Then
            MsgBox "Necesita seleccionar al representante del seguro de desgravamen"
            cboRepDesgrav.SetFocus
            ValidaDatosGrabar = False
        End If
    End If
    '**************************************************************
'madm 20100513 ---------------------------------------------------------------------------------
'ALPA 20100609 B2*******************
'If ActxCta.Prod <> "302" Then

'If nAgenciaCredEval = 0 Then '** JUEZ 20120907
'    If sSTipoProdCod <> "703" Then
'    '***********************************
'        '**** PEAC 20080412
'        If Not IsArray(MatFuentes) Then
'            MsgBox "Debe Selecionar un Fuente de Ingreso para el Credito", vbInformation, "Aviso"
'            ValidaDatosGrabar = False
'            Exit Function
'        Else
'            If UBound(MatFuentes) = 0 Then
'                MsgBox "Debe Selecionar un Fuente de Ingreso para el Credito", vbInformation, "Aviso"
'                ValidaDatosGrabar = False
'                Exit Function
'            End If
'        End If
'        '**** FIN PEAC
'
'    ''************------------------------------------------------------------
'        'Valida Caducidad de Fuente de Ingreso
'        Dim nPos As Integer
'        Dim rsFteIng As ADODB.Recordset
'        Dim rsFIDep As ADODB.Recordset
'        'Dim rsFIInd As ADODB.Recordset, i As Integer
'        Dim rsFIInd As ADODB.Recordset
'
'        Set oNCredito = New COMNCredito.NCOMCredito
'
'        ReDim MatFteFecEval(0)
'        'Call oNCredito.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, oRelPersCred.TitularPersCod, , cmbFuentes.ListIndex)
'        For i = 0 To UBound(MatFuentes) - 1
'
'            'Call oNCredito.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, oRelPersCred.TitularPersCod, , MatFuentes(i))
'            Call oNCredito.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, Me.lblcod, , MatFuentes(i))
'
'            'Call oPersona.RecuperaFtesdeIngreso(oRelPersCred.TitularPersCod, rsFteIng)
'            Call oPersona.RecuperaFtesdeIngreso(Me.lblcod, rsFteIng)
'
'            Call oPersona.RecuperaFtesIngresoDependiente(MatFuentes(i), rsFIDep)
'            Call oPersona.RecuperaFtesIngresoIndependiente(MatFuentes(i), rsFIInd)
'
'            ReDim Preserve MatFteFecEval(UBound(MatFteFecEval) + 1)
'
'        MatFteFecEval(UBound(MatFteFecEval) - 1) = _
'            oPersona.ObtenerFteIngFecEval(MatFuentes(i), _
'            IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = _
'                gPersFteIngresoTipoDependiente, _
'                oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, _
'                oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1))
'
'            'nPos = oPersona.UbicaPosicionFteIngreso(Trim(Mid(cmbFuentes.Text, 100, 20)), CDate(Trim(Right(cmbFuentes.Text, 20))), MatFuentes(i))
'            If gdFecSis >= oPersona.ObtenerFteIngFecCaducac(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1)) Then
'                MsgBox "Fuente de Ingreso a Caducado Ingrese otra Fuente de Ingreso Actual", vbInformation, "Aviso"
'                'cmbFuentes.SetFocus
'                ValidaDatosGrabar = False
'                Exit Function
'            End If
'            'WIOR 20140321 *****************************************
'            If Trim(Right(cmbTipoCredito.Text, 4)) = "150" Or Trim(Right(cmbTipoCredito.Text, 4)) = "250" Or Trim(Right(cmbTipoCredito.Text, 4)) = "350" Then
'                Dim oDPersonaS As comdpersona.DCOMPersonas
'                Dim rsFteEF As ADODB.Recordset
'
'                Set oDPersonaS = New comdpersona.DCOMPersonas
'                'Set rsFteEF = oDPersonaS.RecuperaFuenteIngEstFinan(Trim(MatFuentesF(1, 1)), CDate(MatFuentesF(2, 1)))
'                Set rsFteEF = oDPersonaS.GetUltSemestreEstadoFinancieroPersona(Trim(lblcod.Caption)) 'FRHU 20150311 ERS013-2015
'
'                If rsFteEF.RecordCount = 0 Then
'                    'MsgBox "La fuente de Ingreso debe tener registrado los Estados Financieros.", vbInformation, "Aviso"
'                    MsgBox "Debe registrar los Estados Financieros del Ultimo Semestre.", vbInformation, "Aviso" 'FRHU 20150311 ERS013-2015
'                    ValidaDatosGrabar = False
'                    Set oDPersonaS = Nothing
'                    Set rsFteEF = Nothing
'                    Exit Function
'                End If
'                Set oDPersonaS = Nothing
'                Set rsFteEF = Nothing
'            End If
'            'WIOR 20140321 *****************************************
'    '        'Valida Fuente de Ingreso de Credito Pyme y Comercial Sea una Fuente de Ingreso Independiente
'    '        If CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEAgro Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEEmp Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEPesq Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColComercEmp Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColComercAgro Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColComercPesq Then
'    '            'nPos = oPersona.UbicaPosicionFteIngreso(Trim(Mid(cmbFuentes.Text, 100, 20)), CDate(Trim(Right(cmbFuentes.Text, 20))))
'    '            If CInt(oPersona.ObtenerFteIngTipo(MatFuentes(i))) <> gPersFteIngresoTipoIndependiente Then
'    '                MsgBox "Debe Seleccionar una Fuente de Ingreso Independiente para Este tipo de Credito", vbInformation, "Aviso"
'    '                'cmbFuentes.SetFocus
'    '                ValidaDatosGrabar = False
'    '                Exit Function
'    '            End If
'    '        End If
'    '
'    '        'Valida Fuente de Ingreso de Credito Consumo Sea una Fuente de Ingreso Dependiente
'    '        If CInt(Right(cmbSubTipo.Text, 3)) = gColConsuDctoPlan Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuPlazoFijo Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColConsCTS Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuUsosDiv Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuPrendario Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColConsuPrestAdm Then
'    '            'nPos = oPersona.UbicaPosicionFteIngreso(Trim(Mid(cmbFuentes.Text, 100, 20)), CDate(Trim(Right(cmbFuentes.Text, 20))))
'    '            If CInt(oPersona.ObtenerFteIngTipo(MatFuentes(i))) <> gPersFteIngresoTipoDependiente Then
'    '                MsgBox "Debe Seleccionar una Fuente de Ingreso Dependiente para Este tipo de Credito", vbInformation, "Aviso"
'    '                'cmbFuentes.SetFocus
'    '                ValidaDatosGrabar = False
'    '                Exit Function
'    '            End If
'    '        End If
'
'            '23/092004:LMMD Desabilitado por recomendaciones de Javier Cabrera
'            'Valida que la Institucion y la Fuente de Igreso sean las mismas
'
'        '    If CInt(Trim(Right(cmbSubTipo.Text, 10))) = gColConsuDctoPlan Then
'        '        If Trim(Mid(cmbFuentes.Text, 100, 20)) <> Trim(Right(cmbInstitucion.Text, 20)) Then
'        '            MsgBox "La Fuente de Ingreso no Pertenece a la Institucion", vbInformation, "Aviso"
'        '            cmbFuentes.SetFocus
'        '            ValidaDatos = False
'        '            Exit Function
'        '        End If
'        '    End If
'
'            'CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------------------
'            'Valida el Monto Total a la Fecha (Otros Prestamos Sistema Financiero + Prestamos CMAC + Monto del Prestamo) para distingur un Credito Mes y un Comercial
'    '        If CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEAgro Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEEmp Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColPYMEPesq Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColComercEmp Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColComercAgro Or _
'    '            CInt(Right(cmbSubTipo.Text, 3)) = gColComercPesq Then
'    '
'    '            nMontoFte = oPersona.ObtenerFteIngCreditosCmact(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1))
'    '            nMontoFte = nMontoFte + oPersona.ObtenerFteIngOtrosCreditos(MatFuentes(i), IIf(oPersona.ObtenerFteIngIngresoTipo(MatFuentes(i)) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(MatFuentes(i)) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(MatFuentes(i)) - 1))
'    '            sMonedaFteCod = oPersona.ObtenerFteIngMoneda(MatFuentes(i))
'    '
'    '            Set oNCredito = New COMNCredito.NCOMCredito
'    '            sValor = oNCredito.ValidaMontoParaTipoNCreditoito(Mid(Right(CmbTipoNCredito.Text, 3), 1, 2), Trim(Right(cmbMoneda.Text, 2)), CDbl(txtMontoSol.Text), sMonedaFteCod, nMontoFte, gdFecSis)
'    '            If sValor <> "" Then
'    '                If MsgBox(sValor & vbCrLf & "Desea continuar", vbInformation + vbQuestion, "Aviso") = vbNo Then
'    '                    'cmbTipoNCredito.SetFocus
'    '                    ValidaDatosGrabar = False
'    '                    Exit Function
'    '                End If
'    '            End If
'    '
'    '            Set oNCredito = Nothing
'    '        End If
'            '------------------------------------------------------------------------------------------------------------------------
'        Next i
'
'    ''**-----------------------------------------------------------------------
'    End If 'END MADM
'End If '** END JUEZ

''ARLO 20170910
'Dim oDCreditos As COMDCredito.DCOMCreditos
'Set oDCreditos = New COMDCredito.DCOMCreditos
'
'If oDCreditos.VerificaClienteCampania(ActxCta.NroCuenta) Then
'MsgBox "Este Crédito pertenece a la Campaña Automático, por favor regístrelo por el SICMACM WEB.  ", vbInformation, "Aviso"
'ValidaDatosGrabar = False
'End If
''ARLO 20170910

'WIOR 201205011******************************************************************************************
Set oDPersona = New COMDPersona.DCOMPersona
Set rsPersona = oDPersona.ObtenerEdadPersona(gdFecSis, Trim(lblcod.Caption))
If Not (rsPersona.BOF And rsPersona.EOF) Then
    nEdad = rsPersona!nEdad
End If
nTiempo = val(Me.spnCuotas.valor) * val(Me.SpnPlazo.valor)
dFuturo = DateAdd("d", nTiempo, gdFecSis)
Set rsPersonaF = oDPersona.ObtenerEdadPersona(dFuturo, Trim(lblcod.Caption))
If Not (rsPersonaF.BOF And rsPersonaF.EOF) Then
    nEdadF = rsPersonaF!nEdad
End If

If Trim(Right(cmbMicroseguro.Text, 4)) <> "0" Or Trim(Right(Me.cmbBancaSeguro.Text, 4)) <> "0" Then
    If nEdad >= 70 Then
        MsgBox "El cliente no puede tener o pasar de 70 años.", vbInformation, "Aviso"
        ValidaDatosGrabar = False
        Exit Function
    ElseIf nEdadF >= 75 Then
        MsgBox "El cliente no puede tener mas de 75 años al finalizar el credito.", vbInformation, "Aviso"
        ValidaDatosGrabar = False
        Exit Function
    End If
End If
'WIOR  FIN***********************************************************************************************

'WIOR 201205021******************************************************************************************
Set oCreditoBD = New COMDCredito.DCOMCredActBD
Set oCredito = New COMDCredito.DCOMCredito
'Si tiene Microseguro
Set rsCredito = oCredito.ObtenerMicroseguro(Trim(Me.ActxCta.NroCuenta))
If rsCredito.RecordCount > 0 Then
    Set rsCreditoBD = oCreditoBD.ObtenerBeneficiariosMicroseguro(Trim(Me.ActxCta.NroCuenta))
    If rsCreditoBD.RecordCount > 0 Then
        If Trim(rsCredito!nTipo) <> Trim(Right(cmbMicroseguro.Text, 4)) Then
            If Trim(Right(cmbMicroseguro.Text, 4)) <> "0" Then
                If MsgBox("Credito tiene " & rsCreditoBD.RecordCount & " Beneficiario(s) en Microseguros, Esta seguro de cambiar el tipo del Microseguro de " & IIf(Trim(rsCredito!nTipo) = "1", "S/. 2.50", "S/. 1.50") & " a " & IIf(Trim(Right(cmbMicroseguro.Text, 4)) = "1", "S/. 2.50", "S/. 1.50") & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                End If
            ElseIf Trim(Right(cmbMicroseguro.Text, 4)) = "0" Then
                If MsgBox("Credito tiene " & rsCreditoBD.RecordCount & " Beneficiario(s) en Microseguros, Esta seguro de quitar Microseguro de la sugerencia, Este proceso eliminara a los beneficiarios?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                End If
            End If
        End If
    Else
        If Trim(rsCredito!nTipo) <> Trim(Right(cmbMicroseguro.Text, 4)) Then
            If Trim(Right(cmbMicroseguro.Text, 4)) <> "0" Then
                If MsgBox("Esta seguro de cambiar el tipo del Microseguro de " & IIf(Trim(rsCredito!nTipo) = "1", "S/. 2.50", "S/. 1.50") & " a " & IIf(Trim(Right(cmbMicroseguro.Text, 4)) = "1", "S/. 2.50", "S/. 1.50") & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                End If
            ElseIf Trim(Right(cmbMicroseguro.Text, 4)) = "0" Then
                If MsgBox("Esta seguro de quitar Microseguro de la sugerencia?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                End If
            End If
        End If
    End If
End If
'Si tiene Multiriesgo
Set rsCredito = oCredito.ObtenerMultiriesgo(Trim(Me.ActxCta.NroCuenta))
If rsCredito.RecordCount > 0 Then
    Set rsCreditoBD = oCreditoBD.ObtenerMueblesMultiriesgo(Trim(Me.ActxCta.NroCuenta))
    If rsCreditoBD.RecordCount > 0 Then
        If Trim(Right(cmbBancaSeguro.Text, 4)) = "0" Then
            If MsgBox("Credito tiene " & rsCreditoBD.RecordCount & " Mueble(s) asignados al Seguro Multiriesgo, Esta seguro de quitarlo de la sugerencia, Este proceso eliminara los muebles?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                ValidaDatosGrabar = False
                Exit Function
            End If
        End If
    Else
        If Trim(Right(cmbBancaSeguro.Text, 4)) = "0" Then
            If MsgBox("Esta seguro de quitar el seguro Multiriesgo de la sugerencia?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                ValidaDatosGrabar = False
                Exit Function
            End If
        End If
    End If
End If
Set rsCreditoBD = Nothing
Set rsCredito = Nothing
'WIOR  FIN***********************************************************************************************
    'EJVG20160511 *** Informe de Riesgos y Comumicar Riesgos
    lsMsg = ""
    Set oNCredito = New COMNCredito.NCOMCredito
    lsMsg = oNCredito.ValidaGarantia(Trim(Me.ActxCta.NroCuenta), gdFecSis, CDbl(txtMonSug.Text), fbEsAmpliado)
    Set oNCredito = Nothing
    If Len(lsMsg) > 0 Then
        MsgBox lsMsg, vbInformation, "Aviso"
        ValidaDatosGrabar = False
        SSTab1.Tab = 3
        Exit Function
    End If
    
    'EJVG20160713 ***
    fbEliminarEvaluacion = False
    GenerarDataExposicionEsteCredito ActxCta.NroCuenta, CDbl(txtMonSug.Text), fnMontoExpEsteCred_NEW 'Seteamos el valor de la nueva exposición
    If NecesitaFormatoEvaluacion(ActxCta.NroCuenta, 2001, CInt(Left(sSTipoProdCod, 1) & "00"), CInt(sSTipoProdCod), fnMontoExpEsteCred_NEW, fbEliminarEvaluacion) Then
        ValidaDatosGrabar = False
        Exit Function
    End If
    'END EJVG *******
    
    If Not GenerarDataExposicionRiesgoUnico(Trim(Me.ActxCta.NroCuenta), Trim(lblcod.Caption), Trim(lblnom.Caption)) Then
        ValidaDatosGrabar = False
        Exit Function
    End If
    
    If Not bEsRefinanciado Then
        If Not EmiteInformeRiesgo(eProcesoEmiteInformeRiesgo.Sugerencia, Trim(Me.ActxCta.NroCuenta), sSTipoProdCod, Trim(Right(cmbSubTipo.Text, 5)), Trim(lblcod.Caption), Trim(lblnom.Caption), CDbl(txtMonSug.Text), fbEsAmpliado, spnCuotas.valor) Then
            ValidaDatosGrabar = False
            Exit Function
        End If
    End If
    'END EJVG *******
    'FRHU 20160802 ERS002-2016
    If Not ValidarExisteNivelAprobacionParaAutorizacion(Trim(Me.ActxCta.NroCuenta)) Then
        ValidaDatosGrabar = False
        Exit Function
    End If
    'FIN FRHU 20160802
    
    'JOEP 20160811 ERS004-2016
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000004", sSTipoProdCod) Then     '**END ARLO
    'If sSTipoProdCod = "704" Then
        Dim obj As COMDCredito.DCOMFormatosEval
        Dim rs As ADODB.Recordset
        Set obj = New COMDCredito.DCOMFormatosEval
        Set rs = New ADODB.Recordset
        
        Set rs = obj.ObtenerCapPagoConConvenio(ActxCta.NroCuenta)
        
        If Not (rs.BOF And rs.EOF) Then ' Verifico si trae datos
            If CCur(MatCalend(1, 2)) > rs!nCapPagoTotal Then
                MsgBox "Cliente no Cumple Condiciones " & Chr(13) & "Cuota " & MatCalend(1, 2) & " > " & Format(rs!nCapPagoTotal, "#,##0.00") & " Cap. Pago ", vbInformation, "Alerta"
                    ValidaDatosGrabar = False
                    Exit Function
            End If
        End If
    End If
    'Fin JOEP 20160811 ERS004-2016
    'APRI20170705 TI-ERS025 2017
    Dim rsPers As ADODB.Recordset
    Dim oPers As COMDPersona.UCOMPersona
    Set oPers = New COMDPersona.UCOMPersona
    Set rsPers = oPers.ObtenerVinculadoRiesgoUnico(Trim(lblcod.Caption), "", 0)
    
        If Not (rsPers.BOF And rsPers.EOF) Then
            If rsPers.RecordCount = 1 Then
               If rsPers!nTotal = 1 Then
                    If MsgBox("El vinculado " & rsPers!cPersNombre & " tiene un crédito que se encuentra en " & rsPers!cEstado & ". ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                    End If
                Else
                    If MsgBox("El vinculado " & rsPers!cPersNombre & " tiene " & rsPers!nTotal & " créditos que se encuentran en mora. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                    End If
                End If
            ElseIf rsPers.RecordCount > 1 Then
                If MsgBox("El cliente tiene vinculados en persona que se encuentra en mora. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                    ValidaDatosGrabar = False
                    Exit Function
                   End If
            End If
        End If
        
    Set oPers = Nothing
    'END APRI
   
 'JOEP20171107 3097-2017-GM Acta193.
Dim rsValDes As ADODB.Recordset
Dim obDCredValDes As COMDCredito.DCOMCredito
Set obDCredValDes = New COMDCredito.DCOMCredito

Set rsValDes = obDCredValDes.ValidadDestinoConsEmp(CInt(Trim(Right(cmbProducto.Text, 5))), CInt(Trim(Right(cmbSubProd.Text, 5))), CInt(Trim(Right(cmbDestino.Text, 5))))

If Not (rsValDes.EOF And rsValDes.BOF) Then
    If rsValDes!cMensaje <> "" Then
        MsgBox rsValDes!cMensaje, vbInformation, "No podrá continuar"
        rsValDes.Close
        Set obDCredValDes = Nothing
        ValidaDatosGrabar = False
        Exit Function
    End If
rsValDes.Close
Set obDCredValDes = Nothing
End If
'JOEP20171107 3097-2017-GM Acta193.
   
'JOEP20190115 CP
If Not CP_CondicionesSeg(Trim(Right(cmbSubProd.Text, 5)), Mid(ActxCta.NroCuenta, 9, 1), txtMonSug.Text, spnCuotas.valor, SpnPlazo.valor, Trim(Right(cmbDestino.Text, 5)), txtPerGra.Text, Right(cmbTipoCredito.Text, 5), 2) Then Exit Function
'JOEP20190115 CP
'->*****LUCV20180601, Agregó según ERS022-2018
If TxtComenta.Text = "" Then
    MsgBox "Registrar los comentarios del analista.", vbInformation, "Alerta"
    ValidaDatosGrabar = False
    SSTab1.Tab = 1
    EnfocaControl TxtComenta
    Exit Function
End If
'<-*****  Fin LUCV20180601
End Function

Private Sub GrabarDatos()
Dim oNCredito As COMNCredito.NCOMCredito
Dim oDCredito As COMDCredito.DCOMCredito 'WIOR 20151223
Dim nTasa As Double
Dim sCad As String
Dim pnTipoCuota As ColocTipoCalend
Dim sError As String
Dim rsDR As ADODB.Recordset
Dim sPersCodRepDesgrav As String 'DAOR 20071207
Dim rsRelEmp As ADODB.Recordset 'BRGO 20111103
'WIOR 20131111 **************************
Dim lnCuotaBalon As Integer
Dim vArrDatos As Variant 'EJVG20150513
Dim bRequierePoliza As Boolean 'EJVG20150602
'WIOR 20160610 ***
Dim rsSobreEndCodigos As ADODB.Recordset
Dim sMensajSobreEnd As String
'WIOR FIN ********
Dim oEval As COMDCredito.DCOMFormatosEval 'EJVG20160714
Dim lnMontoCuota As Double 'EJVG20160714

'Dim oSeguro As COMDCaptaGenerales.DCOMSeguros 'APRI20171122 ERS028-2017

'FRHU 20170915 ERS049-2017
Dim ClsMov As COMNContabilidad.NCOMContFunciones
Dim sMovNroM As String
Set ClsMov = New COMNContabilidad.NCOMContFunciones
sMovNroM = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'FIN FRHU 20170915

If chkCuotaBalon.Visible Then
    If chkCuotaBalon.value = 1 Then
        If Trim(txtCuotaBalon.Text) = "" Or Trim(txtCuotaBalon.Text) = "0" Then
            lnCuotaBalon = 0
        Else
            lnCuotaBalon = CInt(Trim(txtCuotaBalon.Text))
        End If
    Else
        lnCuotaBalon = 0
    End If
Else
    lnCuotaBalon = 0
End If
'WIOR FIN *******************************
    On Error GoTo ErrorGrabarDatos
    If Txtinteres.Visible Then
        'ALPA 20141028********************************************************************
        'nTasa = CDbl(Txtinteres.Text)
        nTasa = CDbl(IIf(chkTasa.value = 1, txtInteresTasa.Text, Txtinteres.Text))
        '*********************************************************************************
    Else
        nTasa = CDbl(LblInteres.Caption)
    End If
    
    sCad = ValidaFecha(TxtFecDesemb.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFecDesemb.SetFocus
        Exit Sub
    End If
    
    'If opttcuota(3).value Then
    If Trim(Right(cmbTipoCuota.Text, 2)) = "3" Then
        pnTipoCuota = gColocCalendCodCL
    Else
        'If opttper(0).value Then 'Si es Periodo Fijo
        If Trim(Right(cmbTipoPer.Text, 2)) = "0" Then
            'If opttcuota(0).value Then
            If Trim(Right(cmbTipoCuota.Text, 2)) = "0" Then
                If CInt(txtPerGra.Text) > 0 Then
                    pnTipoCuota = gColocCalendCodPFCFPG
                Else
                    pnTipoCuota = gColocCalendCodPFCF
                End If
            End If
            'If opttcuota(1).value Then 'Cuota Creciente
            If Trim(Right(cmbTipoCuota.Text, 2)) = "1" Then
                If CInt(txtPerGra.Text) > 0 Then
                    pnTipoCuota = gColocCalendCodPFCCPG
                Else
                    pnTipoCuota = gColocCalendCodPFCC
                End If
            End If
            'If opttcuota(2).value Then 'Cuota Decreciente
            If Trim(Right(cmbTipoCuota.Text, 2)) = "2" Then
                If CInt(txtPerGra.Text) > 0 Then
                    pnTipoCuota = gColocCalendCodPFCDPG
                Else
                    pnTipoCuota = gColocCalendCodPFCD
                End If
            End If
        Else
            'If opttcuota(0).value Then
            If Trim(Right(cmbTipoCuota.Text, 2)) = "0" Then
                If CInt(txtPerGra.Text) > 0 Then
                    pnTipoCuota = gColocCalendCodFFCFPG
                Else
                    pnTipoCuota = gColocCalendCodFFCF
                End If
            End If
            'If opttcuota(1).value Then 'Cuota Creciente
            If Trim(Right(cmbTipoCuota.Text, 2)) = "1" Then
                If CInt(txtPerGra.Text) > 0 Then
                    pnTipoCuota = gColocCalendCodFFCCPG
                Else
                    pnTipoCuota = gColocCalendCodFFCC
                End If
            End If
            'If opttcuota(2).value Then 'Cuota Decreciente
            If Trim(Right(cmbTipoCuota.Text, 2)) = "2" Then
                If CInt(txtPerGra.Text) > 0 Then
                    pnTipoCuota = gColocCalendCodFFCDPG
                Else
                    pnTipoCuota = gColocCalendCodFFCD
                End If
            End If
        End If
    End If
    'If Optdesemb(0).value Then
    If Trim(Right(cmbDesemb.Text, 2)) = "0" Then
        ReDim MatDesPar(1, 2)
        MatDesPar(0, 0) = Format(gdFecSis, "dd/mm/yyyy")
        MatDesPar(0, 1) = Format(txtMonSug.Text, "#0.00")
    End If
    
'    Set oNCredito = New COMNCredito.NCOMCredito
'    sError = oNCredito.SugerenciaCredito(ActxCta.NroCuenta, nEstadoActual, gColocEstSug, gdFecSis, nNroTransac, _
'        Trim(Right(Cmblincre.Text, 20)), nTasa, CDbl(txtmonsug.Text), CInt(spnCuotas.Valor), _
'        CInt(spnPlazo.Valor), pnTipoCuota, CInt(txtDiafijo.Text), IIf(ChkProxMes.value, 1, 0), _
'        IIf(Optdesemb(0).value, gColocTiposDesembolsoTotal, gColocTiposDesembolsoParcial), _
'        IIf(OptTipoCalend(0).value, 0, 1), CInt(txtPerGra.Text), IIf(TxtTasaGracia.Visible, CDbl(TxtTasaGracia.Text), CDbl(LblTasaGracia.Caption)), _
'        vnTipoGracia, 1, MatDesPar, MatrizCal, MatCalend_2, ChkCuotaCom.value, ChkMiViv.value, 2, IIf(OptTipoGasto(0).value, "F", "V"), IIf(TxtMora.Visible, CDbl(TxtMora.Text), CDbl(lblMora.Caption)))
    
    If Not IsArray(MatCredVig) Then
        ReDim MatCredVig(0)
    End If
    Dim lsLineaCred As String
    lsLineaCred = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
    
 '  Verifica si el credito es un credito ampliado
 '  Dim oAmpliado  As COMDCredito.DCOMAmpliacion
 '  Dim bAmpliado As Boolean
    
 '  Set oAmpliado = New COMDCredito.DCOMAmpliacion
 '  bAmpliado = oAmpliado.ValidaCreditoaAmpliar(ActxCta.NroCuenta)
 '  Set oAmpliado = Nothing
    
    
'   sError = oNCredito.SugerenciaCredito(ActxCta.NroCuenta, nEstadoActual, gColocEstSug, gdFecSis, nNroTransac, _
        Trim(lsLineaCred), nTasa, CDbl(txtmonsug.Text), CInt(SpnCuotas.Valor), _
        CInt(SpnPlazo.Valor), pnTipoCuota, CInt(TxtDiaFijo.Text), IIf(ChkProxMes.value, 1, 0), _
        IIf(OptDesemb(0).value, gColocTiposDesembolsoTotal, gColocTiposDesembolsoParcial), _
        IIf(OptTipoCalend(0).value, 0, 1), CInt(TxtPerGra.Text), IIf(TxtTasaGracia.Visible, CDbl(TxtTasaGracia.Text), CDbl(LblTasaGracia.Caption)), _
        vnTipoGracia, 1, MatDesPar, MatrizCal, MatCalend_2, ChkCuotaCom.value, ChkMiViv.value, 2, IIf(OptTipoGasto(0).value, "F", "V"), IIf(TxtMora.Visible, CDbl(TxtMora.Text), CDbl(LblMora.Caption)), MatCredVig, TxtComenta.Text, bAmpliado, IIf(ChkTrabajadores.value = 1, True, False))

'madm 20100512 ------------------------------------------------------------------------------------------------------
'ALPA 20100609 B2***********************
  'If Mid(ActxCta.NroCuenta, 6, 3) <> "302" Then
'If nAgenciaCredEval = 0 Then '** JUEZ 20120907
'  If sSTipoProdCod <> "703" Then
'  '*************************************
'    '**** PEAC 20080412 --------------
    Dim MatFtesSel As Variant
'    Dim i As Integer

'    ReDim MatFtesSel(UBound(MatFuentes), 2)
    ReDim MatFtesSel(0, 2)
        
'    For i = 0 To UBound(MatFuentes) - 1
'        MatFtesSel(i, 0) = oPersona.ObtenerFteIngcNumFuente(MatFuentes(i))
'        MatFtesSel(i, 1) = MatFteFecEval(i)
'        MatFtesSel(i, 2) = MatFuentesF(2, i + 1)
'    Next i
'    '**** FIN PEAC 20080412 --------------
'    End If
'end madm -----------------------------------------------------------------------------------------------------------
'End If '** END JUEZ
    '**DAOR 20071207 *********************************
    If cboRepDesgrav.Enabled = False Then
        sPersCodRepDesgrav = ""
    Else
        sPersCodRepDesgrav = Trim(Right(cboRepDesgrav.Text, 13))
    End If
    '*************************************************
    Set oNCredito = New COMNCredito.NCOMCredito
    
    '*** BRGO 20111103 ************************************
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000005", sSTipoProdCod) Then     '**END ARLO
    'If sSTipoProdCod = "517" Then
        grdEmpVinculados.AdicionaFila
        grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.rows - 1, 1) = sPersOperador
        grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.rows - 1, 2) = sPersOperadorNombre
        grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.rows - 1, 3) = gColRelPersOperGarantia
        grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.rows - 1, 4) = txtMontoGarantia.Text
        grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.rows - 1, 5) = txtCtaGarantia.Text
        Set rsRelEmp = Me.grdEmpVinculados.GetRsNew
    End If
    '******************************************************
    'EJVG20150513 *** Nuevos parametros
    If Not IsArray(fvGravamen) Then
        ReDim fvGravamen(0)
    End If
    
    ReDim vArrDatos(4) 'WIOR 20160610 PASO DE 2 A 3'EJVG20160713 Pasó de 3 a 4
    vArrDatos(0) = IIf(ckcPreferencial.value = 1, 1, 0)
    vArrDatos(1) = fvGravamen
    vArrDatos(2) = bRequierePoliza
    'END EJVG *******
    
    'EJVG20160714 *** Actualizamos Ratios e Indicadores
    Set oEval = New COMDCredito.DCOMFormatosEval
    If IsArray(MatCalend) Then
        If UBound(MatCalend, 1) = 1 Then
            lnMontoCuota = CDbl(MatCalend(0, 3)) + CDbl(MatCalend(0, 4)) + CDbl(MatCalend(0, 5)) + CDbl(MatCalend(0, 6))
        Else
            lnMontoCuota = CDbl(MatCalend(1, 3)) + CDbl(MatCalend(1, 4)) + CDbl(MatCalend(1, 5)) + CDbl(MatCalend(1, 6))
        End If
        oEval.RecalculaIndicadoresyRatiosEvaluacion ActxCta.NroCuenta, lnMontoCuota
        Set oEval = Nothing
    End If
    'END EJVG *******
    
    'WIOR 20160609 ***
    Set vArrDatos(3) = Nothing
    sMensajSobreEnd = ""
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000006", sSTipoProdCod) Then     '**END ARLO
    'If sSTipoProdCod <> "703" Then
    'If IsArray(MatFtesSel) Then
        'If MatFtesSel(0, 0) <> "" Then
        If Left(CInt(Trim(Right(cmbSubTipo.Text, 3))), 1) >= 4 Then 'JOEP 20160919
            Set oDCredito = New COMDCredito.DCOMCredito
            'Set rsSobreEndCodigos = oDCredito.SobreEndObtenerCodigos(ActxCta.NroCuenta, Trim(MatFtesSel(0, 0)))
            Set rsSobreEndCodigos = oDCredito.SobreEndObtenerCodigos(ActxCta.NroCuenta, "")
            
            If Not (rsSobreEndCodigos.EOF And rsSobreEndCodigos.BOF) Then
                If CInt(rsSobreEndCodigos!nCodigoFinal) > 0 Then
                    If CInt(rsSobreEndCodigos!nCodigo1) > 0 Then
                        sMensajSobreEnd = sMensajSobreEnd & "Codigo 1: " & IIf(CInt(rsSobreEndCodigos!nCodigo1) = 1, "Potencial Sobreendeudado", "Sobreendeudado") & " - " & Trim(rsSobreEndCodigos!cCodigo1DetDesc) & Chr(10) & Chr(10)
                    End If
                    
                    If CInt(rsSobreEndCodigos!nCodigo2) > 0 Then
                        sMensajSobreEnd = sMensajSobreEnd & "Codigo 2: " & IIf(CInt(rsSobreEndCodigos!nCodigo2) = 1, "Potencial Sobreendeudado", "Sobreendeudado") & " - " & Trim(rsSobreEndCodigos!cCodigo2DetDesc) & Chr(10) & Chr(10)
                    End If
                    
                    If CInt(rsSobreEndCodigos!nCodigo3) > 0 Then
                        sMensajSobreEnd = sMensajSobreEnd & "Codigo 3: " & IIf(CInt(rsSobreEndCodigos!nCodigo3) = 1, "Potencial Sobreendeudado", "Sobreendeudado") & " - " & Trim(rsSobreEndCodigos!cCodigo3DetDesc) & Chr(10) & Chr(10)
                    End If
                    
                    If CInt(rsSobreEndCodigos!nCodigo4) > 0 Then
                        sMensajSobreEnd = sMensajSobreEnd & "Codigo 4: " & IIf(CInt(rsSobreEndCodigos!nCodigo4) = 1, "Potencial Sobreendeudado", "Sobreendeudado") & " - " & Trim(rsSobreEndCodigos!cCodigo4DetDesc) & Chr(10) & Chr(10)
                    End If
                    
                    If CInt(rsSobreEndCodigos!nCodigo5) > 0 Then
                        sMensajSobreEnd = sMensajSobreEnd & "Codigo 5: " & IIf(CInt(rsSobreEndCodigos!nCodigo5) = 1, "Potencial Sobreendeudado", "Sobreendeudado") & " - " & Trim(rsSobreEndCodigos!cCodigo5DetDesc) & Chr(10) & Chr(10)
                    End If
                    
                    MsgBox "El cliente presenta las siguientes alertas de Sobreendeudado con este crédito: " & Chr(10) & Chr(10) & _
                    sMensajSobreEnd & _
                    "Se procederá a enviar una solicitud para el desbloqueo por Sobreendeudamiento.", vbInformation, "Aviso"
                    Set vArrDatos(3) = rsSobreEndCodigos
                End If
            End If
        End If 'JOEP 20160919
        'End If
    'End If
    Set oDCredito = Nothing 'LUCV20171101 - Observacion Inserción de autorizaciones
    End If
    'WIOR FIN ********
    
    vArrDatos(4) = fnMontoExpEsteCred_NEW 'EJVG2016713
    
    sError = oNCredito.GrabarSugerencia(ActxCta.NroCuenta, nEstadoActual, gColocEstSug, gdFecSis, nNroTransac, _
                Trim(lsLineaCred), nTasa, CDbl(txtMonSug.Text), CInt(spnCuotas.valor), CInt(SpnPlazo.valor), _
                pnTipoCuota, CInt(TxtDiaFijo.Text), IIf(ChkProxMes.value, 1, 0), Trim(Right(cmbDesemb.Text, 2)), _
                Trim(Right(cmbTipoCalend.Text, 2)), CInt(txtPerGra.Text), _
                IIf(fbTieneTasaGracia = True, CDbl(TxtTasaGracia.Text), CDbl(LblTasaGracia.Caption)), _
                vnTipoGracia, IIf(bEsRefinanciado, 0, 1), MatDesPar, MatrizCal, MatCalend_2, ChkCuotaCom.value, ChkMiViv.value, IIf(bEsRefinanciado, 1, 2), IIf(OptTipoGasto(0).value, "F", "V"), _
                IIf(TxtMora.Visible, CDbl(TxtMora.Text), CDbl(LblMora.Caption)), MatCredVig, TxtComenta.Text, False, _
                IIf(ChkTrabajadores.value = 1, True, False), rsDR, VerificaTipoCredito, Trim(Right(CboPersCiiu.Text, 10)), _
                CInt(TxtDiaFijo2.Text), chkIncremenK.value, chkVAC.value, IIf(bControlRCC, IIf(chkExpuestoRCC.value = 1, 2, 1), 0), _
                spnNumConCer.valor, sPersCodRepDesgrav, MatFtesSel, CDbl(txtExpAntMax), spnNumConMic.valor, _
                Right(cmbSubTipo.Text, 3), Trim(Right(cmbInstitucionFinanciera, 3)), Mid(ActxCta.NroCuenta, 4, 2), _
                sSTipoProdCod, rsRelEmp, CDbl(txtTasacion.Text), CDbl(Me.lblComisionEC.Caption), _
                fbMicroseguro, fnMicroseguro, fbMultiriesgo, CInt(IIf(Trim(Right(cmbDatoVivienda.Text, 3)) = "", -1, _
                Trim(Right(cmbDatoVivienda.Text, 3)))), lnCuotaBalon, CCur(txtMontoMivivienda.Text), lnCSP, _
                CDbl(IIf(chkTasa.value = 1, Txtinteres.Text, 0)), vArrDatos) 'WIOR 20120517
        'Se agrego el Numero de consulta Score Microfinanazas Gitu 20-05-2009
        'se agrego el parametro "MatFtesSel" PEAC 20080412
        'Manejo CIIU y Nuevas opciones de Calendario y Cod RCC
        'DAOR 20061216, Numero de Consultas Certicom
        'DAOR 20071207, cPersCod de repre. segu. desgravamen (sPersCodRepDesgrav)
        'WIOR 20120517 - SE AGREGO LOS PAREMETROS fbMicroseguro, fnMicroseguro, fbMultiriesgo
        'JUEZ 20130913 Trim(Right(cmbDatoVivienda.Text, 3))
        'WIOR 20131111 AGREGO lnCuotaBalon
        'ALPA 20140621 SE AGREGO txtMontoMivivienda.Text
        'ALPA 20141127 lnCSP
        'LUCV 20160526 IIf(bEsRefinanciado, 0, 1),IIf(bEsRefinanciado, 1, 2)
        'LUCV20180601, ERS022-2018. Reemplazo par?metro TxtTasaGracia.Visible por fbTieneTasaGracia
    ''*** PEAC 20090126
    nMostrarLineaCred = 0
    
    'INICIO EAAS20180912 SEGUN ERS-054-2018
    Dim oDCreditoAS As COMDCredito.DCOMCredito
    Set oDCreditoAS = New COMDCredito.DCOMCredito
    Call oDCreditoAS.GrabarSugerenciaAguaSaneamiento(fvListaAguaSaneamiento, ActxCta.NroCuenta, CDbl(txtMonSug.Text))
    Set oDCreditoAS = Nothing
    'FIN EAAS20180912 SEGUN ERS-054-2018
    'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
    Dim oDCreditoCV As COMDCredito.DCOMCredito
    Set oDCreditoCV = New COMDCredito.DCOMCredito
    Call oDCreditoCV.GrabarSugerenciaCreditoVerde(fvListaCreditoVerde, ActxCta.NroCuenta, CDbl(txtMonSug.Text))
    Set oDCreditoCV = Nothing
    'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
    'WIOR 20151223 ***
    If sError = "" Then
        Set oDCredito = New COMDCredito.DCOMCredito
        Call oDCredito.EliminarDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstSug)
        
        If fbMIVIVIENDA Then
            If IsArray(fArrMIVIVIENDA) Then
                If Trim(fArrMIVIVIENDA(0)) <> "" Then
                    Call oDCredito.RegistrarDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstSug, CDbl(fArrMIVIVIENDA(0)), CDbl(fArrMIVIVIENDA(1)), _
                    CDbl(fArrMIVIVIENDA(2)), CDbl(fArrMIVIVIENDA(3)), CDbl(fArrMIVIVIENDA(4)), CInt(fArrMIVIVIENDA(5)), CInt(fArrMIVIVIENDA(6)), _
                    CDbl(fArrMIVIVIENDA(7)), CDbl(fArrMIVIVIENDA(8)), CInt(fArrMIVIVIENDA(10)))
                End If
            End If
        End If
        Set oDCredito = Nothing 'LUCV20171101 - Observacion Inserción de autorizaciones
    End If
    'WIOR FIN ********
        
    'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta 'RECO20161020 ERS060-2016 'FRHU20161110 Comentado
     'RECO20161020 ERS060-2016 **********************************************************'FRHU 20161110 Descomentado
     Dim oNCOMColocEval As New NCOMColocEval
     Dim lcMovNro As String

        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'JOEP22052017
        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta 'JOEP22052017

     If Not ValidaExisteRegProceso(ActxCta.NroCuenta, gTpoRegCtrlSugerencia) Then
        'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta
        Call oNCOMColocEval.insEstadosExpediente(ActxCta.NroCuenta, "Sugerencia de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlSugerencia)
        Set oNCOMColocEval = Nothing
     End If
     'RECO FIN **************************************************************************
    'ALPA 20141028*******************
'    If Not (ActxCta.Prod = "703" Or ActxCta.Prod = "513" Or lnColocCondicion = 3) Then
'        Dim oDCredito As COMDCredito.DCOMCredito
'        Set oDCredito = New COMDCredito.DCOMCredito
'        Call oDCredito.GrabarVerEntidadesEndeudamiento(ActxCta.NroCuenta, oRsVerEntidades, gdFecSis, gsCodAge, gsCodUser)
'        Set oDCredito = Nothing
'    End If
    '********************************
    
    Set rsRelEmp = Nothing
    If sError <> "" Then
        MsgBox sError, vbInformation, "Aviso"
        Exit Sub
    Else
    'FRHU 20161025 INCIDENTE: CInt(Trim(Right(Me.CmbCampana.Text, 3)))
    Call oNCredito.GrabarSugerenciaNew(ActxCta.NroCuenta, CInt(Trim(Right(cmbDestino.Text, 2))), CInt(Trim(Right(Me.cmbCampana.Text, 3))), actxCtaAho.NroCuenta, _
            ChkBloqueo.value, ChkExoneracion.value, rsExonera, ChkExoSegDes.value, chkDesembBcoNac.value, frmCredDesembBcoNac.txtCodigo.Text, Right(frmCredDesembBcoNac.cmbAgencia.Text, 4))
    'Call oNCredito.GrabarSugerenciaNew(ActxCta.NroCuenta, CInt(Trim(Right(cmbDestino.Text, 2))), CInt(Trim(Right(Me.cmbCampana.Text, 2))), actxCtaAho.NroCuenta, _
            ChkBloqueo.value, ChkExoneracion.value, rsExonera, ChkExoSegDes.value, chkDesembBcoNac.value, frmCredDesembBcoNac.txtCodigo.Text, Right(frmCredDesembBcoNac.cmbAgencia.Text, 4))
    'FIN FRHU 20161025
    'RECO20150602 ERS023-2015************************************************
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim nIndice As Integer
    If IsArray(vMatriz) Then 'WIOR 20151126
        If UBound(vMatriz, 2) > 0 And Right(cmbBancaSeguro.Text, 1) = "2" Then
            For nIndice = 1 To UBound(vMatriz, 2)
                Call oGarant.RegistroGastoMicroSegMYPE(ActxCta.NroCuenta, vMatriz(1, nIndice), vMatriz(2, nIndice), Format(gdFecSis, "yyyyMMdd"))
            Next
        End If
    End If
    'RECO FIN*****************************************************************
    
    'LUCV20160720, Segun ERS004-2016->*****
     If Not CumpleCriteriosRatios(ActxCta.NroCuenta) Then
        MsgBox "El crédito sugerido no podrá ser aprobado. " & Chr(10) & " - Motivo: No cumple con los criterios de ratios financieros. " & Chr(10) & " - Consideración: Favor revisar la evaluación del crédito.", vbInformation, "Alerta"
     Else
        EvaluarCredito ActxCta.NroCuenta, False, 2001, CInt(Mid(sSTipoProdCod, 1, 1) & "00"), CInt(sSTipoProdCod), fnMontoExpEsteCred_NEW, False, True
     End If
    'Fin LUCV20160720  <-*****
    Set oNCredito = Nothing
    
    'LUCV20170302, Según ANEXO 001-2017
        Set oEval = New COMDCredito.DCOMFormatosEval
        oEval.GrabarAlertasTempranas ActxCta.NroCuenta, "190390", Format(gdFecSis, "yyyyMMdd")
        If oEval.VerificarExisteAlertaTemprana(ActxCta.NroCuenta) Then
            Call frmCredAlertaTemprana.Inicio(ActxCta.NroCuenta)
        End If
        Set oEval = Nothing
    'Fin LUCV20170302
    
    'FRHU 20170915 ERS049-2017
    Dim oLineasM As New COMDCredito.DCOMLineaCredito
    Call oLineasM.InsertarColocHistorialTasaMaxima(ActxCta.NroCuenta, CDbl(lnTasaInicial), CDbl(lnTasaFinal), nTasa, sMovNroM, gColocEstSug, gsCodCargo)
    Set oLineasM = Nothing
    'FIN FRHU 20170915
    
    'FRHU 20160615 ERS002-2016
    Set oDCredito = New COMDCredito.DCOMCredito
    Call oDCredito.RegistraAutorizacionesRequeridas(Format(gdFecSis & " " & GetHoraServer(), "yyyy/mm/dd hh:mm:ss"), gsCodUser, gsCodAge, ActxCta.NroCuenta)
    'If oDCredito.verificarExisteAutorizaciones(ActxCta.NroCuenta) Then 'FRHU 20160803
        'Call frmCredNewNivAutorizaVer.Consultar(ActxCta.NroCuenta)
    'End If
    Set oDCredito = Nothing
    'FIN FRHU
    'RECO20160628 ERS002-2016*************************************************
    If fbAutoCalfCPP Then
        Dim oCredNiv As New COMDCredito.DCOMNivelAprobacion
        Dim oRs As New ADODB.Recordset
        Dim sMovNro As String
        
        Set oRs = oCredNiv.ObtieneDatosNivelAutoCta(ActxCta.NroCuenta, "TIP0013")
        If Not (oRs.EOF And oRs.BOF) Then
            If oRs!cMovNroAuto = "" Then
                sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Call oCredNiv.RegistroAutorizacionManual(ActxCta.NroCuenta, oRs!cAutorizaCod, oRs!cNivAprCod, EstadoAutoExonera.gEstadoPendiente, "", sMovNro, oRs!nPrdEstado)
            End If
        End If
    End If
    'RECO FIN ***************************************************************

'JOEP20190919 ERS042 CP-2018
    Dim objAporteReg_CP As COMDCredito.DCOMCredito
    If IsArray(nMatMontoPre) Then
        Set objAporteReg_CP = New COMDCredito.DCOMCredito
        If UBound(nMatMontoPre) > 0 Then
            If Trim(Right(Me.cmbSubProd.Text, 5)) = objProducto.GetResultadoCondicionCatalogo("N0000020", ActxCta.Prod) Then
                Call objAporteReg_CP.CP_GrabaAporte(Trim(ActxCta.NroCuenta), CCur(nMatMontoPre(1, 1)), CCur(nMatMontoPre(1, 2)), CCur(nMatMontoPre(1, 3)), CCur(nMatMontoPre(1, 4))) 'JOEP20190306 Mejora CP
            Else
                Call objAporteReg_CP.CP_GrabaAporte(Trim(ActxCta.NroCuenta), CCur(nMatMontoPre(1, 1)), CCur(nMatMontoPre(1, 2)), CCur(nMatMontoPre(1, 3)))
            End If
        Else
            Set objAporteReg_CP = New COMDCredito.DCOMCredito
            Call objAporteReg_CP.CP_DeleteAporte(Trim(ActxCta.NroCuenta))
            Set objAporteReg_CP = Nothing
        End If
        Set objAporteReg_CP = Nothing
'Comento JOEP20190306 Mejora CP
'    Else
'        Set objAporteReg_CP = New COMDCredito.DCOMCredito
'        Call objAporteReg_CP.CP_DeleteAporte(Trim(ActxCta.NroCuenta))
'        Set objAporteReg_CP = Nothing
'Comento JOEP20190306 Mejora CP
    End If
    'JOEP20190919 ERS042 CP-2018
    
    'FRHU 20160803 ERS002-2016
    Set oDCredito = New COMDCredito.DCOMCredito
    If oDCredito.verificarExisteAutorizaciones(ActxCta.NroCuenta) Then
        Call frmCredNewNivAutorizaVer.Consultar(ActxCta.NroCuenta)
    End If
    Set oDCredito = Nothing
    'FIN FRHU
    
    'NRLO 20180306 Visita Domiciliaria
    Dim rsValidaVisita As ADODB.Recordset
    Dim obDCredVD As COMDCredito.DCOMCredito
    Set obDCredVD = New COMDCredito.DCOMCredito
    Set rsValidaVisita = obDCredVD.RegistrarVisitaDomiciliaria(ActxCta.NroCuenta)
    If Not (rsValidaVisita.EOF And rsValidaVisita.BOF) Then
        If rsValidaVisita!nEstado <> 1 Then
            MsgBox rsValidaVisita!cMensaje, vbInformation, "Aviso"
        End If
    rsValidaVisita.Close
    Set obDCredVD = Nothing
    End If
    'NRLO 20180306 Visita Domiciliaria END
    
    'APRI20171122 ERS028-2017
    'COMENTADO POR APRI20181121 ERS071-2018 - MEJORA
'    Dim oRD As New ADODB.Recordset
'    Set oSeguro = New COMDCaptaGenerales.DCOMSeguros
'    Set oRD = oSeguro.InsertarObservacionesCredSegDesg(ActxCta.NroCuenta)
'    If Not oRD.EOF Then
'    If oRD!nTotal > 0 Then
'        MsgBox "El crédito requiere exámenes obligatorios para la afiliación del Seg. Desgravamen." & _
'            Chr(13) & Chr(13) & "Por favor comunicarse con el Coordinador de Seguros.", vbInformation, "Aviso"
'    End If
'    End If
'    Set oRD = Nothing
'    Set oSeguro = Nothing
    'END APRI
    'APRI20181121 ERS071-2018 - MEJORA
    Dim oSeguro As COMNCaptaGenerales.NCOMSeguros
    Set oSeguro = New COMNCaptaGenerales.NCOMSeguros
    Dim nTotal As Integer
    nTotal = oSeguro.ObservacionesCredSegDesg(ActxCta.NroCuenta)
    If nTotal > 0 Then
        
     MsgBox "El crédito requiere exámenes obligatorios para la afiliación del Seg. Desgravamen." & _
        Chr(13) & Chr(13) & "Por favor comunicarse con el Coordinador de Seguros.", vbInformation, "Aviso"
        
    End If
    Set oSeguro = Nothing
    'END APRI
    
    If vArrDatos(2) = True Then 'Requiere Poliza
        MsgBox "El crédito requiere Seguro Multiriesgo o PYME, coordinar con la Supervisor de Operaciones", vbInformation, "Aviso"
    End If
    
    'JOEP20180725 ERS034 - 2018
    Dim objCredRiegCamb As COMDCredito.DCOMCredito
    Dim rsRiegCamb As ADODB.Recordset
    Set objCredRiegCamb = New COMDCredito.DCOMCredito
    Set rsRiegCamb = objCredRiegCamb.ValidadRigCambCred(ActxCta.NroCuenta)
    If Not (rsRiegCamb.BOF And rsRiegCamb.EOF) Then
        If rsRiegCamb!nApli = 1 Then
            Call RiesgoCambCredPDF(ActxCta.NroCuenta, gdFecSis, gsCodUser, gsNomAge)
        End If
    End If
    RSClose rsRiegCamb
    Set objCredRiegCamb = Nothing
'JOEP20180725 ERS034 - 2018
        ' ya no se imprime el resumen de comite
        'Call ImprimirResumenComite
        
        ' Imprime el Reporte Consolidado del Credito --------------------
    '    Set oNCredito = New COMNCredito.NCOMCredito
        
    '    If VerificaTipoCredito = "AGRICOLA" Then
    '        prsDR = oNCredito.ImprimeConsolidadoCredAgricola(ActxCta.NroCuenta, "S", gdFecSis)
    '    ElseIf VerificaTipoCredito = "CONSUMO" Then
    '        prsDR = oNCredito.ImprimeConsolidadoConsumo(ActxCta.NroCuenta, "S", gdFecSis)
    '    Else
    '        prsDR = oNCredito.ImprimeConsolidadoCred(ActxCta.NroCuenta, "S", gdFecSis)
    '    End If
    '    Set oNCredito = Nothing
        
    'SE COMENTO PARA EVITAR LA DEMORA POR EL DATAENVIROMENT
    '    With DRSugerencia
    '        '.Orientation ='
    '        Set .DataSource = rsDR
    '        .DataMember = ""
    '        .Inicio ActxCta.NroCuenta, "S", gdFecSis
    '        .Refresh
    '        .Show vbModal
    '    End With

        ' ---------------------------------------------------------------

    End If
    Call LimpiaPantalla
    If bCorrigeSug Then
        Unload Me
    End If
    Exit Sub

ErrorGrabarDatos:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Function VerificaTipoCredito() As String
'Devuelve AGRICOLA si  para creditos agricolas
'Devuelve Comercial si es para pequeñas empresas
'Devuelve consumo si es de consumo
 Dim sTipoCredito As String
 Dim sSubTipoCredito As String
 Dim sCuenta As String
 
 sTipoCredito = Mid(ActxCta.NroCuenta, 6, 1) & "00"
 sSubTipoCredito = Mid(ActxCta.NroCuenta, 6, 1) & Mid(ActxCta.NroCuenta, 7, 2)
'ALP 20100609 B2
' If (sTipoCredito = 100 And sSubTipoCredito = "102") Or _
'    (sTipoCredito = "200" And sSubTipoCredito = "202") Then
'        VerificaTipoCredito = "AGRICOLA"
' ElseIf (sTipoCredito = "300" And sSubTipoCredito = "301") Or _
'         (sTipoCredito = "300" And sSubTipoCredito = "302") Or _
'         (sTipoCredito = "300" And sSubTipoCredito = "303") Or _
'         (sTipoCredito = "300" And sSubTipoCredito = "304") Or _
'         (sTipoCredito = "300" And sSubTipoCredito = "305") Or _
'         (sTipoCredito = "300" And sSubTipoCredito = "320") Then
'          VerificaTipoCredito = "CONSUMO"
' End If
 If (Mid(sSTipoProdCod, 1, 2) = "60") Then
        VerificaTipoCredito = "AGRICOLA"
 ElseIf (Mid(sSTipoProdCod, 1, 2) = "70") Then
          VerificaTipoCredito = "CONSUMO"
 End If
End Function

Public Function SumaDesembolsos() As Double
Dim i As Integer
    SumaDesembolsos = 0
    For i = 0 To UBound(MatDesPar) - 1
        SumaDesembolsos = SumaDesembolsos + MatDesPar(i, 1)
    Next i
End Function


Private Function ValidaDatosCalendario() As Boolean
    'ALPA 20111209********************************
    Dim lsCtaCodLeasing As String
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000018", ActxCta.Prod) Then     '**END ARLO
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
        lsCtaCodLeasing = ActxCta.GetCuenta
    End If
    '*********************************************
    ValidaDatosCalendario = True
    
    'Monto de Prestamo
    If Trim(txtMonSug.Text) = "" Then
        MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
        ValidaDatosCalendario = False
        txtMonSug.SetFocus
        Exit Function
    End If
    If CDbl(txtMonSug.Text) <= 0# Then
        MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
        ValidaDatosCalendario = False
        If txtMonSug.Enabled Then
            txtMonSug.SetFocus
        Else
            If CmdDesembolsos.Enabled Then
                CmdDesembolsos.SetFocus
            End If
        End If
        Exit Function
    End If
    'Interes
    
    If Txtinteres.Visible Then
        If Trim(Txtinteres.Text) = "" Then
            MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
            ValidaDatosCalendario = False
            Txtinteres.SetFocus
            Exit Function
        End If
        If CDbl(Txtinteres.Text) <= 0# Then
            MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
            ValidaDatosCalendario = False
            Txtinteres.SetFocus
            Exit Function
        End If
    Else
        If Trim(LblInteres.Caption) = "" Then
            MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
            ValidaDatosCalendario = False
            Exit Function
        End If
        If CDbl(LblInteres.Caption) <= 0# Then
            MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
            ValidaDatosCalendario = False
            Exit Function
        End If
    End If
    
    'Gracia
    If chkGracia.value = 1 Then 'CInt(TxtPerGra.Text) > 0 Then
        'If TxtTasaGracia.Visible Then 'LUCV20180601, Comentó según ERS022-2018
        If fbTieneTasaGracia Then 'LUCV20180601, Agregú según ERS022-2018
            
            'MAVM 25102010 ***
            If txtPerGra.Text = "0" Then
                MsgBox "Ingrese los Dias de Gracia", vbInformation, "Aviso"
                ValidaDatosCalendario = False
                txtFechaFija.SetFocus
                Exit Function
            End If
            '***
            
            If Trim(TxtTasaGracia.Text) = "" Then
                'MsgBox "Ingrese la Tasa de Interes Gracia", vbInformation, "Aviso" 'Comentó según ERS022-2018
                MsgBox "La tasa de interés de gracia, está vinculado a la T.C, por favor ingresarlo.", vbInformation, "Aviso" 'Agregó según ERS022-2018
                ValidaDatosCalendario = False
                'TxtTasaGracia.SetFocus 'LUCV20180601, Comentó según ERS022-2018
                Exit Function
            End If
            If CDbl(TxtTasaGracia.Text) <= 0# Then
                'MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso" 'LUCV20180601, Comentó según ERS022-2018
                MsgBox "La tasa de interés de gracia, está vinculado a la T.C.", vbInformation, "Aviso"
                ValidaDatosCalendario = False
                'TxtTasaGracia.SetFocus 'LUCV20180601, Comentó según ERS022-2018
                Exit Function
            End If
        '->***** LUCV20180601, Comentó según ERS022-2018
'        'JOEP ERS007-2018 20180210
'            Set objProducto = New COMDCredito.DCOMCredito
'            If objProducto.GetResultadoCondicionCatalogo("N0000019", ActxCta.Prod) And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TG
'            'If ActxCta.Prod = "703" And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TG
'                If TxtTasaGracia = "" Then
'                Else
'                    If TxtTasaGracia >= lnTasaGraciaInicial And TxtTasaGracia <= lnTasaGraciaFinal Then
'                    Else
'                        MsgBox "La T.G: esta fuera del Rango: Mínima : " & Format(lnTasaGraciaInicial, "#0.0000") & "  Máxima : " & Format(lnTasaGraciaFinal, "#0.0000"), vbInformation, "Aviso"
'                        TxtTasaGracia.Text = Format(lnTasaGraciaFinal, "#0.0000")
'                    ValidaDatosCalendario = False
'                    Exit Function
'                    End If
'                End If
'            End If
'        'JOEP ERS007-2018 20180210
        '<-*****Fin LUCV20180601
        Else
            If Trim(LblTasaGracia.Caption) = "" Then
                MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso"
                ValidaDatosCalendario = False
                Exit Function
            End If
            If CDbl(LblTasaGracia.Caption) <= 0# Then
                MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso"
                ValidaDatosCalendario = False
                Exit Function
            End If
        End If
    End If
    
'JOEP ERS007-2018 20180210
  Set objProducto = New COMDCredito.DCOMCredito
  If objProducto.GetResultadoCondicionCatalogo("N0000020", ActxCta.Prod) And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TC
 'If ActxCta.Prod = "703" And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then 'TC
    If Txtinteres = "" Then
    Else
        If Txtinteres >= lnTasaInicial And Txtinteres <= lnTasaFinal Then
        Else
            MsgBox "La T.C: esta fuera del Rango: Mínima : " & Format(lnTasaInicial, "#0.0000") & "  Máxima : " & Format(lnTasaFinal, "#0.0000"), vbInformation, "Aviso"
            Txtinteres.Text = Format(lnTasaFinal, "#0.0000")
            ValidaDatosCalendario = False
            Exit Function
        End If
    End If
 End If
'JOEP ERS007-2018 20180210
    
    'Numero de Cuotas
    If Trim(spnCuotas.valor) = "" Or CInt(spnCuotas.valor) <= 0 Then
        MsgBox "Ingrese el Numero de Cuotas del Prestamo", vbInformation, "Aviso"
        ValidaDatosCalendario = False
        spnCuotas.SetFocus
        Exit Function
    End If
    
    'Plazo de Cuotas
    'If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") And opttper(0).value = True Then
    If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") And Trim(Right(cmbTipoPer.Text, 2)) = "0" Then
        MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
        ValidaDatosCalendario = False
        SpnPlazo.SetFocus
        Exit Function
    End If
        
    'Valida dia de Fecha Fija
    'If opttper(1).value And (Trim(TxtDiaFijo.Text) = "" Or Trim(TxtDiaFijo.Text) = "0" Or Trim(TxtDiaFijo.Text) = "00") Then
    If Trim(Right(cmbTipoPer.Text, 2)) = "1" And (Trim(TxtDiaFijo.Text) = "" Or Trim(TxtDiaFijo.Text) = "0" Or Trim(TxtDiaFijo.Text) = "00") Then
        MsgBox "Ingrese el Dia del Mes que Venceran todas las cuotas", vbInformation, "Aviso"
        ValidaDatosCalendario = False
        'TxtDiaFijo.SetFocus Comentado Por MAVM 25102010
        txtFechaFija.SetFocus 'MAVM 25102010
        Exit Function
    End If
    
    If Trim(TxtDiaFijo.Text) = "" Then
        TxtDiaFijo.Text = "00"
    End If
    'Valida Generacion de Tipos de Periodo de Gracia
    If Trim(txtPerGra.Text) = "" Then
        'MAVM 25102010 ***
        'txtPerGra.Text = "00"
        txtPerGra.Text = "0"
        '***
    End If
    '->***** LUCV20180601, Según ERS022-2018
'    If Len(Trim(lsCtaCodLeasing)) = 0 Then 'ALPA 20111209
'        If Trim(txtPerGra.Text) <> "" Then
'            If CInt(txtPerGra.Text) > 0 Then
'                If Not bGraciaGenerada Then
'                    'Excepto las de las nuevas opciones de Gracia
'                    If optTipoGracia(0).value = False And optTipoGracia(1).value = False Then
'                        ValidaDatosCalendario = False
'                        MsgBox "Seleccione un Tipo de Gracia", vbInformation, "Aviso"
'                        If cmdgracia.Enabled Then cmdgracia.SetFocus
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
'    End If
    '<-***** Fin LUCV20180601

    'Valida Calendario de Desembolsos Parciales
    If CmdDesembolsos.Enabled Then
        If Not bDesembParcialGenerado Then
            MsgBox "Ingrese los Desembolsos del Credito", vbInformation, "Aviso"
            ValidaDatosCalendario = False
            CmdDesembolsos.SetFocus
            Exit Function
        End If
    End If
    
End Function

Private Function DameTipoCuota() As Integer
Dim i As Integer
    For i = 0 To 3
        'If opttcuota(i).value Then
        If CInt(Trim(Right(Me.cmbTipoCuota.Text, 2))) = i Then
            DameTipoCuota = i + 1
            Exit For
        End If
    Next i
End Function

Private Sub LimpiaPantalla()
    bValidaCargaSugerenciaAguaSaneamiento = 0 'EAAS20180912 SEGUN ERS-054-2018
    bValidaCargaSugerenciaCreditoVerde = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    ReDim MatDesPar(0, 0)
    ReDim MatrizCal(0, 0)
    ReDim MatDesemb(0, 0)
    'ActxCta.SetFocusProd
    Call LimpiaControles(Me)
    'Cmblincre.Clear
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
    
    If Not RLinea Is Nothing Then
        RLinea.Close
    End If
    Set RLinea = Nothing
    
    'opttcuota(0).value = True
    cmbTipoCuota.ListIndex = 0
    'opttper(0).value = True
    'cmbTipoPer.ListIndex = 0'Comento JOEP20190125 CP
    cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
    'Optdesemb(0).value = True
    cmbDesemb.ListIndex = 0
    'OptTipoCalend(0).value = True
    cmbTipoCalend.ListIndex = 0
    CmdDesembolsos.Enabled = False
    ChkProxMes.value = 0
    spnCuotas.valor = "0"
    SpnPlazo.valor = "0"
    'opttcuota(0).value = True
    cmbTipoCuota.ListIndex = 0
    opttper(0).value = True
    'OptTipoCalend(0).Enabled = True
    cmbTipoCalend.Enabled = True
    FraDatos.Enabled = False
    cmdrelac.Enabled = False
    CmdGrabar.Enabled = False
    CmdCalend.Enabled = False
    CmdGastos.Enabled = False
    'cmdEvaluacion.Enabled = False
    cmdCheckList.Enabled = False
    ActxCta.Enabled = True
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    CmdCredVig.Enabled = False
    chkGracia.value = 0
    cmdLineas.Enabled = True 'DAOR 20070409
    ChkBloqueo.value = 0
    ChkExoSegDes.value = 0
    ChkExoneracion.value = 0
    chkDesembBcoNac.value = 0
    Call Unload(frmCredVigentes)
    MatCredVig = ""
    'ALPA 20100604 B2 *******************************
    cmbTipoCredito.ListIndex = -1
    cmbSubTipo.Clear
    '************************************************
    txtFechaFija = "__/__/____" 'MAVM 25102010
    SSTab1.TabVisible(2) = False
    'grdEmpVinculados.Recordset = Nothing
    'WIOR 20120510 ******************************************************************
    cmbMicroseguro.ListIndex = IndiceListaCombo(cmbMicroseguro, 0)
    cmbBancaSeguro.ListIndex = IndiceListaCombo(cmbBancaSeguro, 0)
    lblComisionEC.Caption = "0.00"
    'WIOR FIN ***********************************************************************
    Set rsExonera = Nothing
    'cmdCheckList.Enabled = False 'RECO20150415 ERS010-2015
    bCheckList = False 'RECO20150513 *****************************
    FormateaFlex FEGarantCred 'EJVG20150922
    'fbAutoCalfCPP = True 'RECO20160628 ERS002-2016
    fbAutoCalfCPP = False 'FRHU20160915 ERS002-2016
    chkAutoCalifCPP.value = 0 'RECO20160628 ERS002-2016
End Sub

'Private Sub UbicaRegistro(ByVal psLineaCred As String)
'    If RLinea.RecordCount > 0 Then
'        RLinea.MoveFirst
'        Do While Not RLinea.EOF
'            If RLinea!cLineaCred = psLineaCred Then
'                Exit Do
'            End If
'            RLinea.MoveNext
'        Loop
'    End If
'End Sub

Private Sub HabilitaEditarTasa(ByVal pbHabilita As Boolean)
    Txtinteres.Visible = pbHabilita
    Txtinteres.Enabled = pbHabilita
    LblInteres.Enabled = Not pbHabilita
    LblInteres.Visible = Not pbHabilita
End Sub

Private Sub HabilitaFechaFija(ByVal pbHabilita As Boolean)
    'MAVM 25102010 ***
    SpnPlazo.Enabled = Not pbHabilita
    SpnPlazo.valor = IIf(pbHabilita, "0", SpnPlazo.valor)
    '***
    'TxtDiaFijo.Enabled = pbHabilita Comentado Por MAVM 25102010
    'TxtDiaFijo.Text = "00" 'Comentado Por MAVM 25102010
    ChkProxMes.value = 0
    'ChkProxMes.Enabled = pbHabilita 'Comentado Por MAVM 25102010
    'TxtDiaFijo2.Enabled = pbHabilita   'ARCV 30-04-2007
    TxtDiaFijo2.Text = "00"
    'Comentado Por MAVM 25102010 ***
    'If pbHabilita Then
    '    spnPlazo.Enabled = False
    'Else
    '    spnPlazo.Enabled = True
    'End If
    'spnPlazo.valor = "0"
    '***
End Sub

Private Sub HabilitaCuotaLibre(ByVal pbHabilita As Boolean)
    'If opttper(1).value Then
    If Trim(Right(cmbTipoPer.Text, 2)) = "1" Then
        Call HabilitaFechaFija(True)
    Else
        Call HabilitaFechaFija(False)
    End If
    'opttper(0).Enabled = Not pbHabilita
    'opttper(1).Enabled = Not pbHabilita
    cmbTipoPer.Enabled = Not pbHabilita
    'OptTipoCalend(0).Enabled = Not pbHabilita
    'OptTipoCalend(1).Enabled = Not pbHabilita
    cmbTipoCalend.Enabled = Not pbHabilita
    lblPerGra.Enabled = Not pbHabilita
    txtPerGra.Text = "0"
    txtPerGra.Enabled = Not pbHabilita
    'cmdgracia.Enabled = Not pbHabilita 'LUCV20180601, Comentó Según ERS022
End Sub

Private Sub CargaTipoCuota(ByVal pnTipoCuota As String)
    
    Select Case pnTipoCuota
        Case Trim(str(gColocCalendCodPFCF)) 'Periodo Fijo Cuota Fija
            'opttcuota(0).value = True
            cmbTipoCuota.ListIndex = 0
            'opttper(0).value = True
            'cmbTipoPer.ListIndex = 0'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
            Call HabilitaFechaFija(False)
            Call HabilitaCuotaLibre(True)
        Case Trim(str(gColocCalendCodPFCC))  'Periodo Fijo - Cuota Creciente
            'opttcuota(1).value = True
            cmbTipoCuota.ListIndex = 1
            'opttper(0).value = True
            'cmbTipoPer.ListIndex = 0'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
            Call HabilitaFechaFija(False)
            Call HabilitaCuotaLibre(True)
        Case Trim(str(gColocCalendCodPFCD))  'Periodo Fijo - Cuota Decreciente"
            'opttcuota(2).value = True
            cmbTipoCuota.ListIndex = 2
            'opttper(0).value = True
             'cmbTipoPer.ListIndex = 0'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
            Call HabilitaFechaFija(False)
            Call HabilitaCuotaLibre(True)
        Case Trim(str(gColocCalendCodFFCF))  'Fecha Fija - Cuota Fija
            'opttcuota(1).value = True
            cmbTipoCuota.ListIndex = 1
            'opttper(1).value = True
            'cmbTipoPer.ListIndex = 1'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
            Call HabilitaFechaFija(True)
            Call HabilitaCuotaLibre(True)
        Case Trim(str(gColocCalendCodFFCC))  'Fecha Fija - Cuota Creciente
            'opttcuota(2).value = True
            cmbTipoCuota.ListIndex = 2
            'opttper(1).value = True
            'cmbTipoPer.ListIndex = 1'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
            Call HabilitaFechaFija(True)
            Call HabilitaCuotaLibre(True)
        Case Trim(str(gColocCalendCodFFCDPG))      'Fecha Fija - Cuota Decreciente
            'opttcuota(3).value = True
            cmbTipoCuota.ListIndex = 3
            'opttper(1).value = True
            'cmbTipoPer.ListIndex = 1'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
            Call HabilitaFechaFija(True)
            Call HabilitaCuotaLibre(False)
        Case Trim(str(gColocCalendCodCL))
            'opttcuota(4).value = True
            cmbTipoCuota.ListIndex = 4
            'opttper(1).value = True
            'cmbTipoPer.ListIndex = 1'Comento JOEP20190125 CP
            cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
            Call HabilitaCuotaLibre(True)
    End Select
End Sub

Private Sub DeshabilitaTipoDesemb(Optional ByVal pnTipoDesemb As Boolean = True)
    'Optdesemb(0).Enabled = pnTipoDesemb
    'Optdesemb(1).Enabled = pnTipoDesemb
    cmbDesemb.Enabled = pnTipoDesemb
End Sub

Private Sub DeshabilitaTipoCalend(Optional ByVal pnTipoCalend As Boolean = True)
    'OptTipoCalend(0).Enabled = pnTipoCalend
    'OptTipoCalend(1).Enabled = pnTipoCalend
    cmbTipoCalend.Enabled = pnTipoCalend
End Sub

Private Sub DeshabilitaTipoGracia(Optional ByVal pnPeriodoGRacia As Boolean = True)
    lblPerGra.Enabled = pnPeriodoGRacia
    txtPerGra.Enabled = pnPeriodoGRacia
    Label4.Enabled = pnPeriodoGRacia
    TxtTasaGracia.Enabled = pnPeriodoGRacia
    LblTasaGracia.Enabled = pnPeriodoGRacia
    'cmdgracia.Enabled = pnPeriodoGRacia 'LUCV20180601, Comentó Según ERS022-2018
    
    'ARCV 30-04-2007
    'optTipoGracia(0).Enabled = pnPeriodoGRacia
    'optTipoGracia(1).Enabled = pnPeriodoGRacia
    '---
End Sub

Private Sub DeshabilitaTipoPeriodo(Optional ByVal pnPeriodoFijo As Boolean = True, Optional ByVal pnFechaFijo As Boolean = True)
    'opttper(0).Enabled = pnPeriodoFijo
    'opttper(1).Enabled = pnFechaFijo
    cmbTipoPer.Enabled = pnFechaFijo
    'If opttper(1).value Then
    If Trim(Right(cmbTipoPer.Text, 2)) = "1" Then
        TxtDiaFijo.Enabled = pnFechaFijo
        ChkProxMes.Enabled = pnFechaFijo
    End If
End Sub

Public Sub AsignaTipoCalendario(ByVal pnColocCalendCod As ColocTipoCalend)
    If pnColocCalendCod = gColocCalendCodCL Then
        DeshabilitaTipoPeriodo False, False
        DeshabilitaTipoGracia False
        DeshabilitaTipoCalend False
        DeshabilitaTipoDesemb False
        'opttcuota(3).value = True
        cmbTipoCuota.ListIndex = 3
    End If
    If pnColocCalendCod = gColocCalendCodFFCC Or pnColocCalendCod = gColocCalendCodFFCCPG Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        'opttper(1).value = True
        'cmbTipoPer.ListIndex = 1 'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
        'opttcuota(1).value = True
        cmbTipoCuota.ListIndex = 1
    End If
    If pnColocCalendCod = gColocCalendCodFFCD Or pnColocCalendCod = gColocCalendCodFFCDPG Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        'opttper(1).value = True
        'cmbTipoPer.ListIndex = 1 'Comento JOEP20190125 CP
         cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
        'opttcuota(2).value = True
        cmbTipoCuota.ListIndex = 2
    End If
    If pnColocCalendCod = gColocCalendCodFFCF Or pnColocCalendCod = gColocCalendCodFFCFPG Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        'opttper(1).value = True
        'cmbTipoPer.ListIndex = 1 'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
        'opttcuota(0).value = True
        cmbTipoCuota.ListIndex = 0
    End If
    If pnColocCalendCod = gColocCalendCodPFCC Or pnColocCalendCod = gColocCalendCodPFCCPG Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        'opttper(0).value = True
        'cmbTipoPer.ListIndex = 0 'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
        'opttcuota(1).value = True
        cmbTipoCuota.ListIndex = 1
    End If
    If pnColocCalendCod = gColocCalendCodPFCD Or pnColocCalendCod = gColocCalendCodPFCDPG Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        'opttper(0).value = True
        'cmbTipoPer.ListIndex = 0 'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
        'opttcuota(2).value = True
        cmbTipoCuota.ListIndex = 2
    End If
    If pnColocCalendCod = gColocCalendCodPFCF Or pnColocCalendCod = gColocCalendCodPFCFPG Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        opttper(0).value = True
        'opttcuota(0).value = True
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
        cmbTipoCuota.ListIndex = 0
    End If
End Sub

'Private Function UbicaLineaCredito(ByVal psLinea As String) As Integer
'Dim I As Integer
'    UbicaLineaCredito = 0
'    For I = 0 To Cmblincre.ListCount - 1
'        If Trim(Right(Cmblincre.List(I), 20)) = Trim(psLinea) Then
'            UbicaLineaCredito = I
'            Exit Function
'        End If
'    Next I
'End Function

 Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'Dim oCredito As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim R2 As ADODB.Recordset
'Dim oLineas As COMDCredito.DCOMLineaCredito
'Dim oCalend As COMDCredito.DCOMCalendario
Dim nSaldoCapTmp As Double
Dim nTasaCompLinea As Double
Dim nTasaGraciaLinea As Double
Dim sLineaTmp As String
Dim oNCredito As COMNCredito.NCOMCredito
Dim nTasaMora As Double
'Dim oCal As COMDCredito.DCOMCalendario

'Variables pasados por Referencia
Dim bValidaSugerencia As Boolean
Dim rsSuger As ADODB.Recordset
Dim rsCalend As ADODB.Recordset
Dim rsCalend2 As ADODB.Recordset
Dim rsLineas As ADODB.Recordset
Dim rsRelEmp As ADODB.Recordset 'BRGO 20111103
Dim rsRelOtros As ADODB.Recordset 'BRGO 20111103
Dim bRefinanciado As Boolean
Dim nOpcion As Integer
'Dim rsTipoCredito As ADODB.Recordset
Dim rsRepDesgrav As ADODB.Recordset 'DAOR 20071207
Dim rsRel As ADODB.Recordset 'BRGO 20111104
'WIOR 20120517 ******************************************
Dim nMicroseguro As Integer
Dim nMultiriesgo As Integer
'WIOR - FIN *********************************************
Dim rsDatCredEval As ADODB.Recordset 'JUEZ 20120907
'WIOR 20160618 ***
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsSobreEnd As ADODB.Recordset
'WIOR FIN ********

'WIOR 20151223 ***
fbDatosCargados = False
fArrMIVIVIENDA = ""
'WIOR FIN ********

Dim Y As Integer '**ARLO20171127 ERS070 - 2017
Dim rsTpPago As ADODB.Recordset 'JOEP20190114 CP

    On Error GoTo ErrorCargaDatos
    MatCredVig = ""
    
    'If Optdesemb(1).value Then
    '    nOpcion = 1
    'End If
    'If opttcuota(3).value Then
    If Trim(Right(Me.cmbTipoCuota.Text, 2)) = "3" Then
        nOpcion = 3
    End If
    nComisionEC = 0
    lnMostrarCSP = 0
    lnLogicoBuscarDatos = 0
    
    'MARG ERS003-2018--------------------
    Dim oRelPersCred As UCredRelac_Cli
    Dim MatCredRelaciones As Variant
    
    Set oRelPersCred = New UCredRelac_Cli
    oRelPersCred.CargaRelacPersCred (psCtaCod)
    MatCredRelaciones = oRelPersCred.ObtenerMatrizRelaciones
    Dim ofrmSolicitud As frmCredSolicitud
    Dim bPermiteSugerencia As Boolean
    Set ofrmSolicitud = New frmCredSolicitud
    Screen.MousePointer = vbHourglass 'LUCV20180601, Agregó Según ERS022-2018
    'bPermiteSugerencia = ofrmSolicitud.ValidarScoreExperian(psCtaCod, MatCredRelaciones) 'comment by marg 201906
    bPermiteSugerencia = ofrmSolicitud.ValidarScore(psCtaCod, MatCredRelaciones, 2) 'add by marg 201906
    If Not bPermiteSugerencia Then
        Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
        Exit Function
    End If
    'END MARG----------------------------
    
    
    'ARLO 20170910
    Dim oDCreditos As COMDCredito.DCOMCreditos
    Set oDCreditos = New COMDCredito.DCOMCreditos
    
    If oDCreditos.VerificaClienteCampania(ActxCta.NroCuenta) Then
        Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
        MsgBox "Este Crédito pertenece a la Campaña Automático, por favor regístrelo por el SICMACM WEB.  ", vbInformation, "Aviso"
        Exit Function
    End If
    'ARLO 20170910
    
    
    '**ARLO20171113 INICIO ERS070 - 2017
    Dim rsCompraDeuda  As ADODB.Recordset
    Set rsCompraDeuda = oDCreditos.obtieneCompraDeuda(ActxCta.NroCuenta)

    If Not (rsCompraDeuda.EOF And rsCompraDeuda.BOF) Then
    ReDim fvListaCompraDeuda(rsCompraDeuda.RecordCount)

    Me.feDeudaComprar.Clear
    Me.feDeudaComprar.rows = 2
    Me.feDeudaComprar.FormaCabecera
    Me.feDeudaComprar.Enabled = True ''**ARLO20180313
    
    For Y = 1 To rsCompraDeuda.RecordCount
        feDeudaComprar.AdicionaFila
        feDeudaComprar.TextMatrix(Y, 0) = rsCompraDeuda!nItem
        fvListaCompraDeuda(Y).nId = feDeudaComprar.TextMatrix(Y, 0)

        feDeudaComprar.TextMatrix(Y, 1) = rsCompraDeuda!cInstitucion
        fvListaCompraDeuda(Y).sIFINombre = feDeudaComprar.TextMatrix(Y, 1)

        feDeudaComprar.TextMatrix(Y, 2) = rsCompraDeuda!cNroCredito
        fvListaCompraDeuda(Y).sNroCredito = feDeudaComprar.TextMatrix(Y, 2)

        feDeudaComprar.TextMatrix(Y, 3) = rsCompraDeuda!cMoneda
        fvListaCompraDeuda(Y).nmoneda = IIf(rsCompraDeuda!cMoneda = "SOLES", 1, 2)

        feDeudaComprar.TextMatrix(Y, 4) = rsCompraDeuda!nNroCuoPacta
        fvListaCompraDeuda(Y).nNroCuotasPactadas = feDeudaComprar.TextMatrix(Y, 4)

        feDeudaComprar.TextMatrix(Y, 5) = Format(rsCompraDeuda!nSaldoComprar, "#,##0.00") 'ARLO20180314
        fvListaCompraDeuda(Y).nSaldoComprar = feDeudaComprar.TextMatrix(Y, 5)

        feDeudaComprar.TextMatrix(Y, 6) = Format(rsCompraDeuda!nMontoCuoPaga, "#,##0.00") 'ARLO20180314
        fvListaCompraDeuda(Y).nMontoCuota = feDeudaComprar.TextMatrix(Y, 6)

        feDeudaComprar.TextMatrix(Y, 7) = rsCompraDeuda!nDestino
        fvListaCompraDeuda(Y).nDestino = feDeudaComprar.TextMatrix(Y, 7)

        feDeudaComprar.TextMatrix(Y, 8) = rsCompraDeuda!bRecompra
        fvListaCompraDeuda(Y).bRecompra = feDeudaComprar.TextMatrix(Y, 8)

        feDeudaComprar.TextMatrix(Y, 9) = rsCompraDeuda!nMontoDesemb
        fvListaCompraDeuda(Y).nMontoDesembolso = feDeudaComprar.TextMatrix(Y, 9)

        feDeudaComprar.TextMatrix(Y, 10) = rsCompraDeuda!nNroCuoPaga
        fvListaCompraDeuda(Y).nNroCuotasPagadas = feDeudaComprar.TextMatrix(Y, 10)

        feDeudaComprar.TextMatrix(Y, 11) = rsCompraDeuda!dFechaDesemb
        fvListaCompraDeuda(Y).dDesembolso = feDeudaComprar.TextMatrix(Y, 11)

        fvListaCompraDeuda(Y).sIFICod = rsCompraDeuda!cPersCodIfi

        rsCompraDeuda.MoveNext

    Next Y
    
    Else
        ''**ARLO20180313
        Me.feDeudaComprar.Clear
        Me.feDeudaComprar.rows = 2
        Me.feDeudaComprar.FormaCabecera
        Me.feDeudaComprar.Enabled = False
        ''**ARLO20180313
    End If
    '**ARLO20171113 FIN ERS070 - 2017
        
    
    
    'WIOR 20160618 ***
    Set oDCredito = New COMDCredito.DCOMCredito
    Set rsSobreEnd = oDCredito.SobreEndObtenerCodigosRegXCta(psCtaCod)
    If Not (rsSobreEnd.EOF And rsSobreEnd.BOF) Then
        Dim sCodigosPotSobreEnd As String
        Dim sCodigosSobreEnd As String
        sCodigosPotSobreEnd = ""
        sCodigosSobreEnd = ""
        
        sCodigosPotSobreEnd = IIf(CInt(rsSobreEnd!nCodigo1) = 1, "1,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo2) = 1, "2,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo3) = 1, "3,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo4) = 1, "4,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo5) = 1, "5,", "")
        If Len(sCodigosPotSobreEnd) > 0 Then
            sCodigosPotSobreEnd = Mid(sCodigosPotSobreEnd, 1, Len(sCodigosPotSobreEnd) - 1)
        End If
        
        sCodigosSobreEnd = IIf(CInt(rsSobreEnd!nCodigo1) = 2, "1,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo2) = 2, "2,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo3) = 2, "3,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo4) = 2, "4,", "") & _
                                IIf(CInt(rsSobreEnd!nCodigo5) = 2, "5,", "")
        If Len(sCodigosSobreEnd) > 0 Then
            sCodigosSobreEnd = Mid(sCodigosSobreEnd, 1, Len(sCodigosSobreEnd) - 1)
        End If
        
        If CInt(rsSobreEnd!nEstado) = 3 Then
            CargaDatos = False
            If MsgBox("Crédito ya fue desbloqueado por Sobreendeudamiento, Favor de proceder con la Aprobación " & _
                        "o Desea volver a Sugerir(Solitara Nuevamente los Desbloqueos de Sobreendeudamiento)", vbInformation + vbYesNo, "Aviso") = vbNo Then
                cmdCancelar_Click
                Exit Function
            End If
        ElseIf CInt(rsSobreEnd!nEstado) = 1 Then
            CargaDatos = False
            If MsgBox("Crédito ya fue desbloqueado por Potencial Sobreendeudamiento(Codigo(s): " & sCodigosPotSobreEnd & ")" & _
                ", faltando aún el desbloqueo por SobreEndeudamiento(Codigo(s): " & sCodigosSobreEnd & "). Favor de solicitar su Desbloqueo respetivo" & _
                "o Desea volver a Sugerir(Solitara Nuevamente los Desbloqueos de Sobreendeudamiento)", vbInformation + vbYesNo, "Aviso") = vbNo Then
                cmdCancelar_Click
                Exit Function
            End If
        ElseIf CInt(rsSobreEnd!nEstado) = 2 Then
            CargaDatos = False
            If MsgBox("Crédito ya fue desbloqueado por Sobreendeudamiento(Codigo(s): " & sCodigosSobreEnd & ")" & _
                ", faltando aún el desbloqueo por Potencial SobreEndeudamiento(Codigo(s): " & sCodigosPotSobreEnd & "). Favor de solicitar su Desbloqueo respetivo." & _
                "o Desea volver a Sugerir(Solitara Nuevamente los Desbloqueos de Sobreendeudamiento)", vbInformation + vbYesNo, "Aviso") = vbNo Then
                cmdCancelar_Click
                Exit Function
            End If
        End If
    End If
    Set oDCredito = Nothing
    Set rsSobreEnd = Nothing
    'WIOR FIN ********
    
    Set oNCredito = New COMNCredito.NCOMCredito
    Call oNCredito.CargaDatosSugerencia(psCtaCod, nOpcion, Mid(ActxCta.NroCuenta, 6, 3), Mid(ActxCta.NroCuenta, 9, 1), bValidaSugerencia, rsSuger, _
                                        rsCalend, rsCalend2, rsLineas, bRefinanciado, fvListaAguaSaneamiento, fvListaCreditoVerde, nSaldoDisponible, rsRepDesgrav, rsRelEmp, rsRelOtros, _
                                        nMicroseguro, nMultiriesgo, rsDatCredEval, fbEsAmpliado) 'DAOR 20071207, se agregó rsRepDesgrav
                                        'WIOR 20120517 SE AGREGARON LOS PARAMETROS nMicroseguro, nMultiriesgo
                                      'JUEZ 20120908 se agregó rsDatCredEval EAAS fvListaAguaSaneamiento
    'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
     If UBound(fvListaAguaSaneamiento) >= 1 Then
        nSumaAguaSaneamiento = fvListaAguaSaneamiento(1).nSumaAguaSaneamiento
     End If
     If UBound(fvListaCreditoVerde) >= 1 Then
        nSumaCreditoVerde = fvListaCreditoVerde(1).nSumaCreditoVerde
     End If
    'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
    Set oNCredito = Nothing
    
    'Unload frmCredSugExonera 'LUCV20180601, Comentó según ERS022-2018
    Set rsExonera = Nothing
    nLogicoVerEntidades = 0
    bEsRefinanciado = bRefinanciado 'DAOR 20070407
    fnMontoExpEsteCred_NEW = 0 'EJVG20160713
    'If oCredito.ValidaSugAprobacion(psCtaCod) = False Then 'verifica que el credito no se haya aprobado anteriormente
    If bValidaSugerencia = False Then
        ' MsgBox "El credito ya esta aprobado", vbInformation, "AVISO"
        Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
        Exit Function
    End If
    'Set R = oCredito.RecuperaSugerencia(psCtaCod)
    If Not rsSuger.BOF And Not rsSuger.EOF Then
        txtMontoMivivienda.Text = Format(rsSuger!nMontoMiVivienda, "#0.00") 'ALPA20140611
        vnTipoGracia = IIf(IsNull(rsSuger!nTipoGracia), -1, rsSuger!nTipoGracia)
        'Set oCal = New COMDCredito.DCOMCalendario
        'Set R2 = oCal.RecuperaCalendarioPagos(psCtaCod)
        'Set oCal = Nothing
        lnCampanaId = rsSuger!idCampana 'ALPA 20141030**********
        ReDim MatGracia(rsCalend.RecordCount)
        Do While Not rsCalend.EOF
            MatGracia(rsCalend.Bookmark - 1) = Format(rsCalend!nIntGracia, "#0.00")
            rsCalend.MoveNext
        Loop
        'R2.Close
        CargaDatos = True
        'RECO20160628 ERS002-2016 ***********************************************
        If rsSuger!cAutorizaCod <> "" Then
            chkAutoCalifCPP.value = 1
            chkAutoCalifCPP.Enabled = False
        Else
            chkAutoCalifCPP.value = 0
            chkAutoCalifCPP.Enabled = True
        End If
        'RECO FIN ***************************************************************
        'JUEZ 20140610 **********************************************************
        Dim oDPersGen As COMDPersona.DCOMPersGeneral
        Dim RsSector As ADODB.Recordset
        Set oDPersGen = New COMDPersona.DCOMPersGeneral
        Set RsSector = oDPersGen.VerificaSolicitudAutorizacionRiesgos(rsSuger!cPersCod)
        Set oDPersGen = Nothing
        If Not (RsSector.EOF And RsSector.BOF) Then
            If RsSector!nEstado = 0 Then
                'MsgBox "El crédito no puede ser sugerido, pues tiene una solicitud de autorización pendiente en riesgos", vbInformation, "Aviso"
                MsgBox "El crédito supera el porcentaje máximo por sector. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso" 'WIOR 20150714
                CargaDatos = False
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            ElseIf RsSector!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
                CargaDatos = False
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            End If
        End If
        Set RsSector = Nothing
        'END JUEZ ***************************************************************
        
        'JOEP ERS047 20170901 **********************************************************
        Dim oCredZonaG As COMDCredito.DCOMCredito
        Dim RsZonaGeog As ADODB.Recordset
        Set oCredZonaG = New COMDCredito.DCOMCredito
        Set RsZonaGeog = oCredZonaG.VerificaSolicitudAutorizacionZonaGeog(psCtaCod, Mid(ActxCta.NroCuenta, 6, 3))
        Set oCredZonaG = Nothing
        
        If Not (RsZonaGeog.EOF And RsZonaGeog.BOF) Then
            If RsZonaGeog!nEstado = 0 Then
                MsgBox "El crédito supera el porcentaje máximo por Zona Geografica. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso"
                CargaDatos = False
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            ElseIf RsZonaGeog!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
                CargaDatos = False
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            End If
        End If
        Set RsZonaGeog = Nothing
        
        Dim oCredlimiteProd As COMDCredito.DCOMCredito
        Dim RslimitePro As ADODB.Recordset
        Set oCredlimiteProd = New COMDCredito.DCOMCredito
        Set RslimitePro = oCredlimiteProd.VerificaSolicitudAutorizacionProducto(psCtaCod, Mid(ActxCta.NroCuenta, 6, 3))
        Set oCredlimiteProd = Nothing
        If Not (RslimitePro.EOF And RslimitePro.BOF) Then
            If RslimitePro!nEstado = 0 Then
                MsgBox "El crédito supera el porcentaje máximo por Tipo de Producto. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso"
                CargaDatos = False
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            ElseIf RslimitePro!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
                CargaDatos = False
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            End If
        End If
        Set RslimitePro = Nothing
        'END JOEP ERS047 ***************************************************************
        
        'ALPA 20150116***********************************************************
        chkTasa.value = IIf(rsSuger!bExononeracionTasa = 1, 1, 0)
        'ckcPreferencial.value = IIf(rsSuger!bPreferencial = 1, 1, 0)
        chkTasa.value = IIf(rsSuger!bExononeracionTasa = 1, 1, 0)
        'Dim oCliPre As COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
        Dim bValidar As Boolean
        'Set oCliPre = New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
        'bValidar = oCliPre.ValidarClientePreferencial(rsSuger!cPersCod)    'COMENTADO POR ARLO 20170722
        bValidar = False 'ARLO 20170722
        If bValidar Then
            ckcPreferencial.value = 1
            lnCliPreferencial = 1
        Else
            ckcPreferencial.value = 0
            lnCliPreferencial = 0
        End If
        'Set oCliPre = Nothing 'COMENTADO POR ARLO 20170722
        '************************************************************************
        'JUEZ 20120907 **********************************************
        Set oNCredito = New COMNCredito.NCOMCredito
        'nAgenciaCredEval = IIf(oNCredito.ObtieneAgenciaCredEval(gsCodAge) = 1 And Not rsDatCredEval.EOF, 1, 0)
        'Set oNCredito = Nothing
        'nVerifCredEval = rsSuger!nVerifCredEval
        'If nAgenciaCredEval = 1 Then
        '    If nVerifCredEval = 0 Then
        '        CargaDatos = False
        '        Exit Function
        '    End If
        '    MsgBox "Datos registrados en la evaluación: Monto: " & IIf(rsDatCredEval!nmoneda = "1", "S/. ", "$ ") & Format(rsDatCredEval!nMontoCalc, "#,##0.00") & ", TEM: " & Format(rsDatCredEval!nTEMCalc, "#,##0.00") & "%, Cuotas: " & rsDatCredEval!nCuotasCalc, vbInformation, "Referencia"
        'End If
        ''END JUEZ ***************************************************
        bVerifNivApr = True
        If gnAgenciaCredEval = 1 Then
            If bCorrigeSug = False And bVerDetalle = False Then 'VERIFICAR SI SE ESTA APROBANDO EL CREDITO POR NIVELES DE APROBACION
                Dim rsVerif As ADODB.Recordset
                Dim oDNiv As COMDCredito.DCOMNivelAprobacion
                Set oDNiv = New COMDCredito.DCOMNivelAprobacion
                Set rsVerif = oDNiv.RecuperaDatosCredResultado(psCtaCod)
                Set oNCredito = Nothing
                If Not rsVerif.EOF Then
                    MsgBox "No puede modificar este crédito, está siendo verificado por los niveles de aprobación", vbInformation, "Aviso"
                    bVerifNivApr = False
                    CargaDatos = False
                    Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                    Exit Function
                End If
            End If
        End If
        lnColocCondicion = rsSuger!nColocCondicion 'ALPA20141230
        fnMontoExpEsteCred_NEW = rsSuger!nMontoExpCredito 'EJVG20160713
        'ALPA20141127******************
        lnCSP = rsSuger!nCuotaPolizaMivivienda
        If rsSuger!nSTipoCredCod = "853" Then
            chkCSP.Visible = True
            If lnCSP > 0 Then
                chkCSP.value = 1
            End If
        Else
            lnCSP = 0
        End If
        '*************************************
        'BN
        Dim oCredBN As COMDCredito.DCOMCredito
        Dim rsDesBcoNac As ADODB.Recordset
        Set oCredBN = New COMDCredito.DCOMCredito
        Set rsDesBcoNac = oCredBN.RecuperaDatosDesembolsoBcoNac(psCtaCod)
        If Not rsDesBcoNac.BOF And Not rsDesBcoNac.EOF Then
            If rsDesBcoNac!nEstado = 1 Then ' Si ya se generó el archivo para el banco de la nación
                CargaDatos = False
                MsgBox "El Credito es para desembolso en Banco de la Nación y ya forma parte del archivo que será enviado al Banco de la Nación", vbExclamation, "Aviso"
                Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
                Exit Function
            End If
        End If
        Set rsDesBcoNac = Nothing
        Set oCredBN = Nothing
        
        TxtFecDesemb.Text = Format(gdFecSis, "dd/mm/yyyy")
        Txtinteres.Text = "0.00"
        LblInteres.Caption = "0.00"
        TxtTasaGracia.Text = "0.00"
        LblTasaGracia.Caption = "0.00"
        TxtMora.Text = "0.00"
        LblMora.Caption = "0.00"
        nMostrarLineaCred = 1
        ChkCuotaCom.value = IIf(IsNull(rsSuger!bCuotaCom), 0, rsSuger!bCuotaCom)
        sLineaTmp = rsSuger!cLineaCred
        nTasaCompLinea = IIf(IsNull(rsSuger!nTasaComp), 0, rsSuger!nTasaComp)
        nTasaGraciaLinea = IIf(IsNull(rsSuger!nTasaGracia), 0, rsSuger!nTasaGracia)
        nEstadoActual = IIf(IsNull(rsSuger!nPrdEstado), vbNull, rsSuger!nPrdEstado)
        'nPersFIDIngCliActual = IIf(IsNull(rsSuger!nPersFIDIngCli), 0, rsSuger!nPersFIDIngCli)
        'cPersFIMonedaActual = IIf(IsNull(rsSuger!cPersFIMoneda), "", rsSuger!cPersFIMoneda)
        nNroTransac = rsSuger!nTransacc
        spnNumConCer.valor = IIf(IsNull(rsSuger!nNumConCer), 0, rsSuger!nNumConCer) ' DAOR 20061216, Numero de Consultas a la central de riesgos
        If rsSuger!cTipoGasto = "V" Then
            OptTipoGasto(1).value = True
        Else
            OptTipoGasto(0).value = True
        End If
        lblcod.Caption = rsSuger!cPersCod
        lblnom.Caption = rsSuger!cPersNombre
        TxtComenta.Text = IIf(IsNull(rsSuger!cDescripcion), "", rsSuger!cDescripcion)
        lblSolModif.Caption = rsSuger!cSolModif
        LblDni.Caption = Trim(IIf(IsNull(rsSuger!Dni), "", rsSuger!Dni))
        LblRuc.Caption = Trim(IIf(IsNull(rsSuger!Ruc), "", rsSuger!Ruc))
        
        lblMonto = Format(rsSuger!nMontoSol, "#0.00")
        lblcuotassol.Caption = Trim(str(rsSuger!nCuotasSol))
        lblplazosol.Caption = Trim(str(rsSuger!nPlazoSol))
        lbldescre.Caption = Trim(rsSuger!cDestinoDescripcion)
        lblanalista.Caption = PstaNombre(Trim(IIf(IsNull(rsSuger!cAnalista), "", rsSuger!cAnalista)))
        txtMonSug.Text = Format(rsSuger!nMonto, "#0.00")
        '*** PEAC 20080412
        txtExpAntMax.Text = Format(rsSuger!nExpoAntMax, "#0.00")
        'ALPA 20100605 B2******************
        'Carga Tipo de creditos
'        Call Llenar_Combo_con_Recordset(rsTipoCredito, cmbTipoCredito)
'        Call CambiaTamañoCombo(cmbTipoCredito)
        Call CargaInstitucionesFinancieras(gTpoInstFinanc)
        Call CargaDatoVivienda 'JUEZ 20130913
        Call cmbTipoCredito_Click
        lbltProd.Caption = Trim(rsSuger!cTipoProdDescrip)
        lblSubProd.Caption = Trim(rsSuger!cSTipoProdDescrip)
        sTipoProdCod = rsSuger!nTipoProdCod
        sSTipoProdCod = rsSuger!nSTipoProdCod
        cmbProducto.ListIndex = IndiceListaCombo(cmbProducto, rsSuger!nTipoProdCod)
        Call cmbProducto_Click
        cmbSubProd.ListIndex = IndiceListaCombo(cmbSubProd, rsSuger!nSTipoProdCod)
        
'Catalogo JOEP
        sCodPers = ""
        sCodPers = rsSuger!cPersCod
        Call CatalogoCombos(Trim(Right(Me.cmbSubProd.Text, 5)), 1000)
'Catalogo JOEP
        
        cmbTipoCredito.ListIndex = IndiceListaCombo(cmbTipoCredito, rsSuger!nTipoCredCod)
        cmbSubTipo.ListIndex = IndiceListaCombo(cmbSubTipo, rsSuger!nSTipoCredCod)
        cmbInstitucionFinanciera.ListIndex = IndiceListaCombo(cmbInstitucionFinanciera, rsSuger!nTipoInstCorp)
        cmbDatoVivienda.ListIndex = IndiceListaCombo(cmbDatoVivienda, rsSuger!nDatoVivienda) 'JUEZ 20130913
        'Condicion
        lblCondicion.Caption = rsSuger!cColocCondicionDesc
        
        'Destino del Credito
        If lnColocCondicion <> 4 Then
            Call CatalogoCombos(Trim(Right(Me.cmbSubProd.Text, 5)), 2000)
        End If
'Tipo de pago
        Call CatalogoCombos(Trim(Right(Me.cmbSubProd.Text, 5)), 3000)
'Catalogo JOEP
        cmbDestino.ListIndex = IndiceListaCombo(cmbDestino, rsSuger!nColocDestino)
        cmbDestino.Enabled = False 'Catalogo JOEP
        'Campaña
        cmbCampana.ListIndex = IndiceListaCombo(cmbCampana, IIf(IsNull(rsSuger!idCampana), 0, rsSuger!idCampana))
        '**********************************
        spnCuotas.valor = Trim(str(rsSuger!nCuotas))
        Call CargaTipoPago '20200217 Mejora pase de Canlendario interativo
        'ALPA 20111209****************
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000021", ActxCta.Prod) Then     '**END ARLO
        'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            nValorDiaGracia = 0
        End If
        '*****************************
        SpnPlazo.valor = Trim(str(rsSuger!nPlazo))
        txtPerGra.Text = IIf(IsNull(rsSuger!nPeriodoGracia), "0", rsSuger!nPeriodoGracia)
        
        'MAVM 25102010 ***
        'ALPA 20141206**********************
        If txtFechaFija.Text <> "__/__/____" Then
            If SpnPlazo.valor <> "0" Then
                txtFechaFija.Text = CDate(txtFechaFija.Text + CDate(txtPerGra.Text))
            End If
        End If
        '***********************************
        If txtPerGra.Text <> "0" Then
            chkGracia.Enabled = True
        End If
        '***
        
        nTasaMora = rsSuger!nTasaMora
        
        'Muestra el control Spinner para el cobro por consulta Score Microfinanzas
        'solo para creditos MES Gitu 20-05-2009
        
        'If Mid(psCtaCod, 6, 3) = "201" Then
        '->***** LUCV20180601, Comentó según ERS022-2018
'        '**ARLO20180712 ERS042 - 2018
'        Set objProducto = New COMDCredito.DCOMCredito
'        If objProducto.GetResultadoCondicionCatalogo("N0000007", sSTipoProdCod) Then     '**END ARLO
'        'If sSTipoProdCod = "503" Or sSTipoProdCod = "504" Then
'        'Comento JOEP Catalogo
'            'spnNumConMic.Visible = True
'            'Label14.Visible = True
'        'Comento JOEP Catalogo
'            spnNumConMic.valor = IIf(IsNull(rsSuger!nNumConMic), 0, rsSuger!nNumConMic) ' GITU 20090602, Numero de Consultas Score Microfinanzas
'        Else
'            spnNumConMic.Visible = False
'            Label14.Visible = False
'        End If
        '<-***** Fin LUCV20180601
        'End Gitu
        
        'Asigna Tipo de Cuota
        Call AsignaTipoCalendario(IIf(IsNull(rsSuger!nColocCalendCod), gColocCalendCodPFCF, rsSuger!nColocCalendCod))
        'MAVM 25102010 ***
        'TxtDiaFijo.Text = Trim(str(IIf(IsNull(rsSuger!nPeriodoFechaFija), "00", rsSuger!nPeriodoFechaFija)))
        TxtDiaFijo.Text = Format(Trim(str(IIf(IsNull(rsSuger!nPeriodoFechaFija), "00", rsSuger!nPeriodoFechaFija))), "00")
        '***
        If IsNull(rsSuger!nProxMes) Then
            ChkProxMes.value = 0
        Else
            If rsSuger!nProxMes = True Then
                ChkProxMes.value = 1
            Else
                ChkProxMes.value = 0
            End If
        End If
        
'JOEP20190305 CP
    Dim rsAporte As ADODB.Recordset
    Dim objAporte As COMDCredito.DCOMCredito
    Set objAporte = New COMDCredito.DCOMCredito

    Set rsAporte = objAporte.CargaAporte(Trim(ActxCta.NroCuenta))
    If Not (rsAporte.BOF And rsAporte.EOF) Then
        If Trim(Right(cmbSubProd.Text, 9)) = objProducto.GetResultadoCondicionCatalogo("N0000017", Trim(Right(cmbSubProd.Text, 9))) Then 'JOEP20190305
            ReDim nMatMontoPre(1, 4)
                nMatMontoPre(1, 1) = rsAporte!nMonto
                nMatMontoPre(1, 2) = rsAporte!nAporte
                nMatMontoPre(1, 3) = rsAporte!nMontoSoli
                nMatMontoPre(1, 4) = rsAporte!nMontoDisponible
        Else
            ReDim nMatMontoPre(1, 3)
                nMatMontoPre(1, 1) = rsAporte!nMonto
                nMatMontoPre(1, 2) = rsAporte!nAporte
                nMatMontoPre(1, 3) = rsAporte!nMontoSoli
        End If
    Else
        ReDim nMatMontoPre(0)
    End If
'JOEP20190305 CP
        
        'MAVM 25102010 ***
        txtPerGra.Enabled = False
        If (TxtDiaFijo.Text <> "00") And (TxtDiaFijo.Text <> "0") Then
            ChkProxMes.Enabled = False
            If ChkProxMes.value = 0 Then
                txtFechaFija.Text = CDate(TxtDiaFijo.Text & "/" & Month(gdFecSis) & "/" & Year(gdFecSis)) + CDate(txtPerGra.Text)
            Else
                chkGracia.Enabled = True
                If txtPerGra.Text > "0" Then
                    'txtFechaFija.Text = Mid(CDate(rsSuger!dPrdEstado + 30 + txtPerGra.Text), 1, 10)
                    txtFechaFija.Text = Mid(CDate(gdFecSis + 30 + txtPerGra.Text), 1, 10)
                Else
                    If Not Trim(ValidaFecha((TxtDiaFijo.Text & Mid(DateAdd("m", 1, gdFecSis), 3, 8)))) = "" Then
                        MsgBox Trim(ValidaFecha(txtFechaFija.Text)), vbInformation, "Aviso"
                    Else
                        txtFechaFija.Text = CDate(TxtDiaFijo.Text & Mid(DateAdd("m", 1, gdFecSis), 3, 8))
                    End If
                End If
            End If
        End If
        '***
    
        'Tipo de Desembolso
        If IsNull(rsSuger!nTipoDesembolso) Then
            'Optdesemb(0).value = True
            cmbDesemb.ListIndex = 0
            CmdDesembolsos.Enabled = False
            'MatDesemb = ""
        Else
            If rsSuger!nTipoDesembolso = gColocTiposDesembolsoTotal Then
                'Optdesemb(0).value = True
                cmbDesemb.ListIndex = 0
                CmdDesembolsos.Enabled = False
                'MatDesemb = ""
            Else
                'Optdesemb(1).value = True
                cmbDesemb.ListIndex = 1
                txtMonSug.Text = Format(rsSuger!nMonto, "#0.00")
                CmdDesembolsos.Enabled = True
                
                Dim oCalend As COMDCredito.DCOMCalendario
                Dim RDes As ADODB.Recordset
                Set oCalend = New COMDCredito.DCOMCalendario
                Set RDes = oCalend.RecuperaCalendarioDesemb(Me.ActxCta.NroCuenta)
                If RDes.RecordCount > 0 Then
                ReDim MatDesemb(RDes.RecordCount, 2)
                End If
                Do While Not RDes.EOF
                    MatDesemb(RDes.Bookmark - 1, 0) = Format(RDes!dVenc, "dd/mm/yyyy")
                    MatDesemb(RDes.Bookmark - 1, 1) = Format(RDes!nCapital, "#0.00")
                    RDes.MoveNext
                Loop
            End If
        End If
        'Tipo de Calendario
        If IsNull(rsSuger!nCalendDinamico) Then
            'OptTipoCalend(0).value = True
            cmbTipoCalend.ListIndex = 0
        Else
            If rsSuger!nCalendDinamico = 0 Then
                'OptTipoCalend(0).value = True
                cmbTipoCalend.ListIndex = 0
            Else
                'OptTipoCalend(1).value = True
                cmbTipoCalend.ListIndex = 1
            End If
        End If
        ChkMiViv.value = IIf(IsNull(rsSuger!bMiVivienda), 0, rsSuger!bMiVivienda)
        If ChkMiViv.value = 1 Then
            'OptTipoCalend(0).value = True
            'OptTipoCalend(1).Enabled = False
            cmbTipoCalend.ListIndex = 0
        End If
        
        '*** CIUU del Credito*****
        CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, rsSuger!cPersCIIU)
        '***************************
        
        'No aplica ya que recien en el momento de la Sugerencia se indica si es MiVivienda
        'If rsSuger!bMiVivienda = 1 Then
        '    ChkMiViv.Visible = True
        '    FraGastos.Visible = True
        'Else
        '    ChkMiViv.Visible = False
        '    FraGastos.Visible = False
        'End If
        '**********************
        '*** Manejo de las nuevas opciones de Gracia y 2 dias
        '*** dias fijos en el Tipo de Periodo "Fecha Fija"
        TxtDiaFijo2.Text = Format(rsSuger!nDiaFijo2, "00")
        If rsSuger!nTipoGracia = gColocTiposGraciaCapitalizada Then
            optTipoGracia(0).value = True
        End If
        chkIncremenK.value = rsSuger!bIncremGraciaCap
        If rsSuger!nTipoGracia = gColocTiposGraciaEnCuotas Then
            optTipoGracia(1).value = True
        End If
        fraGracia.Enabled = False
        If CInt(txtPerGra.Text) > 0 Then
            chkGracia.value = 1
        Else
            chkGracia.value = 0
        End If
        'Call chkGracia_Click 'Comentado Por MAVM 25102010
        '****************************
        'WIOR 20131108 *************************************************************
        Set oNCredito = New COMNCredito.NCOMCredito
        If CInt(rsSuger!nCuotaBalonCred) = 0 Then
            txtCuotaBalon.Enabled = False
            If oNCredito.AplicaCuotaBalon(gsCodAge, Trim(rsSuger!nSTipoProdCod), CDbl(rsSuger!nMonto), CInt(Mid(psCtaCod, 9, 1))) Then
                chkCuotaBalon.Visible = True
                txtCuotaBalon.Visible = True
                txtCuotaBalon.Text = "0"
                chkCuotaBalon.value = 0
            Else
                chkCuotaBalon.Visible = False
                txtCuotaBalon.Visible = False
            End If
        Else
            chkCuotaBalon.Visible = True
            chkCuotaBalon.value = 1
            txtCuotaBalon.Visible = True
            txtCuotaBalon.Enabled = True
            txtCuotaBalon.Text = Trim(rsSuger!nCuotaBalonCred)
        End If
        'WIOR FIN ******************************************************************
        'Factor VAC
        chkVAC.value = rsSuger!bVAC
        '04-05-2005
        If Mid(psCtaCod, 7, 1) = "3" Then
            chkVAC.value = 1
        End If
        bBuscarLineas = True
        '****************
    '   R.Close
    '   Set R = Nothing
    '   Set oCredito = Nothing
                    
'        'Carga Lineas de Credito EN COMBO
'        Set oLineas = New dLineaCredito
'        Set RLinea = oLineas.RecuperaLineasProducto(Mid(ActxCta.NroCuenta, 6, 3), Mid(ActxCta.NroCuenta, 9, 1))
'        Set oLineas = Nothing
'        Cmblincre.Clear
'        Do While Not RLinea.EOF
'            Cmblincre.AddItem Trim(RLinea!cDescripcion) & Space(50) & Trim(RLinea!cLineaCred)
'            RLinea.MoveNext
'        Loop
'        If Cmblincre.ListCount > 0 Then
'            Cmblincre.ListIndex = UbicaLineaCredito(sLineaTmp)
'        Else
'            Txtinteres.Text = "0.00"
'            lblInteres.Caption = "0.00"
'            TxtTasaGracia.Text = "0.00"
'            LblTasaGracia.Caption = "0.00"
'        End If
'        Call CambiaTamañoCombo(Cmblincre, 300)
        
        'Comentado por DAOR 20070404, se repetía lineas abajo
        'If Txtinteres.Visible Then
        '    Txtinteres.Text = Format(nTasaCompLinea, "#0.00")
        'End If
        'If TxtTasaGracia.Visible Then
        '    TxtTasaGracia.Text = Format(nTasaGraciaLinea, "#0.00")
        'End If
        'If TxtMora.Visible Then
        '    TxtMora.Text = Format(nTasaMora, "#0.00")
        'End If
        
        'Carga Lineas de Credito en arbol ---------------------------------------------------------------------------
        txtBuscarLinea.sTitulo = "Lineas de Crédito"
        txtBuscarLinea.psRaiz = "Lineas de Crédito"

        'Dim RLineaProducto As ADODB.Recordset
        'Set oLineas = New COMDCredito.DCOMLineaCredito
        'Set RLineaProducto = oLineas.RecuperaLineasProductoArbol(Mid(ActxCta.NroCuenta, 6, 3), Mid(ActxCta.NroCuenta, 9, 1))
        'Set oLineas = Nothing
        
        txtBuscarLinea.rs = rsLineas  'RLineaProducto
        'Set RLineaProducto = Nothing
        
        'Comentado por DAOR 20070407
        'txtBuscarLinea.Text = ""
        'lblLineaDesc.Caption = ""
        
        '**DAOR 20070407***********************************
        If Not bRefinanciado Then
            txtBuscarLinea.Text = ""
            lblLineaDesc.Caption = ""
        End If
        '**************************************************
        
        If sLineaTmp <> "" Then
           txtBuscarLinea.Text = IIf(Mid(sLineaTmp, 6, 1) = "1", "CP1-" + sLineaTmp, "LP2-" + sLineaTmp)
           txtBuscarLinea_EmiteDatos
        '--------------------------------------------------------------------------------------------
        Else
           'Comentado por DAOR 20070407
           'TxtInteres.Text = "0.00"
           'lblInteres.Caption = "0.00"
           'TxtTasaGracia.Text = "0.00"
           'LblTasaGracia.Caption = "0.00"
            
            '**DAOR 20070407***********************************
            If Not bRefinanciado Then
                Txtinteres.Text = "0.00"
                LblInteres.Caption = "0.00"
                TxtTasaGracia.Text = "0.00"
                LblTasaGracia.Caption = "0.00"
            End If
            '**************************************************
        End If
        
        If Txtinteres.Visible Then
            Txtinteres.Text = Format(nTasaCompLinea, "#0.00")
        End If
        'If TxtTasaGracia.Visible Then 'LUCV20180601, Comentó según ERS022-2018
        If fbTieneTasaGracia Then 'LUCV20180601, Según ERS022-2018
            TxtTasaGracia.Text = Format(nTasaGraciaLinea, "#0.00")
        End If
        If TxtMora.Visible Then
            TxtMora.Text = Format(nTasaMora, "#0.00")
        End If
        
        '--------------------------------------------------------------------------------------------------------------
        
        'Carga Calendario Desembolso
        'If Optdesemb(1).value Then
        If Trim(Right(cmbDesemb.Text, 2)) = "1" Then
            'Set oCalend = New COMDCredito.DCOMCalendario
            'Set R = oCalend.RecuperaCalendarioDesemb(psCtaCod)
            ReDim MatDesemb(rsCalend2.RecordCount, 2)
            ReDim MatDesPar(rsCalend2.RecordCount, 2)
            Do While Not rsCalend2.EOF
                MatDesemb(rsCalend2.Bookmark - 1, 0) = Format(rsCalend2!dVenc, "dd/mm/yyyy")
                MatDesPar(rsCalend2.Bookmark - 1, 0) = Format(rsCalend2!dVenc, "dd/mm/yyyy")
                MatDesemb(rsCalend2.Bookmark - 1, 1) = Format(rsCalend2!nCapital, "#0.00")
                MatDesPar(rsCalend2.Bookmark - 1, 1) = Format(rsCalend2!nCapital, "#0.00")
                rsCalend2.MoveNext
            Loop
            'R.Close
            'Set R = Nothing
            'Set oCalend = Nothing
        End If

        'Carga Calendario Pagos
        'If opttcuota(3).value Then
        If Trim(Right(cmbTipoCuota.Text, 2)) = "3" Then
            nSaldoCapTmp = CDbl(txtMonSug.Text)
            'Set oCalend = New COMDCredito.DCOMCalendario
            'Set R = oCalend.RecuperaCalendarioPagos(psCtaCod)
            ReDim MatCalend(rsCalend2.RecordCount, 6)
            ReDim MatrizCal(rsCalend2.RecordCount, 6)
            Do While Not rsCalend2.EOF
                'fecha
                MatCalend(rsCalend2.Bookmark - 1, 0) = Format(rsCalend2!dVenc, "dd/mm/yyyy")
                MatrizCal(rsCalend2.Bookmark - 1, 0) = Format(rsCalend2!dVenc, "dd/mm/yyyy")
                'Cuota
                MatCalend(rsCalend2.Bookmark - 1, 1) = Trim(str(rsCalend2!nCuota))
                MatrizCal(rsCalend2.Bookmark - 1, 1) = Trim(str(rsCalend2!nCuota))
                
                'Monto Cuota
                MatCalend(rsCalend2.Bookmark - 1, 2) = Format(rsCalend2!nCapital + rsCalend2!nIntComp + IIf(IsNull(rsCalend2!nIntGracia), 0, rsCalend2!nIntGracia), "#0.00")
                MatrizCal(rsCalend2.Bookmark - 1, 2) = Format(rsCalend2!nCapital + rsCalend2!nIntComp + IIf(IsNull(rsCalend2!nIntGracia), 0, rsCalend2!nIntGracia), "#0.00")
                
                'Capital
                MatCalend(rsCalend2.Bookmark - 1, 3) = Format(rsCalend2!nCapital, "#0.00")
                MatrizCal(rsCalend2.Bookmark - 1, 3) = Format(rsCalend2!nCapital, "#0.00")
                'Interes Compensatorio
                MatCalend(rsCalend2.Bookmark - 1, 4) = Format(rsCalend2!nIntComp, "#0.00")
                MatrizCal(rsCalend2.Bookmark - 1, 4) = Format(rsCalend2!nIntComp, "#0.00")
                'Interes Gracia
                MatCalend(rsCalend2.Bookmark - 1, 5) = Format(IIf(IsNull(rsCalend2!nIntGracia), 0, rsCalend2!nIntGracia), "#0.00")
                MatrizCal(rsCalend2.Bookmark - 1, 5) = Format(IIf(IsNull(rsCalend2!nIntGracia), 0, rsCalend2!nIntGracia), "#0.00")
                
                'Saldo Capital
                nSaldoCapTmp = nSaldoCapTmp - rsCalend2!nCapital
                MatCalend(rsCalend2.Bookmark - 1, 5) = Format(nSaldoCapTmp, "#0.00")
                MatrizCal(rsCalend2.Bookmark - 1, 5) = Format(nSaldoCapTmp, "#0.00")
                
                rsCalend2.MoveNext
            Loop
            'R.Close
            'Set R = Nothing
            'Set oCalend = Nothing
        End If
        
        'Set oNCredito = New COMNCredito.NCOMCredito
        'If oNCredito.EsRefinanciado(psCtaCod) Then
        If bRefinanciado Then
            txtMonSug.Enabled = False
'''''''            cmdLineas.Enabled = False 'COMENTADO X MADM 20110419 ---   DAOR 20070407
        End If
        'Set oNCredito = Nothing
        
        'madm 20100513
       
        Call CargaFuentesIngreso(Me.lblcod)
        
        '**DAOR 20071207 ****************************************
        fnPersPersoneria = rsSuger!nPersPersoneria
        If rsSuger!nPersPersoneria > 1 Then
            cboRepDesgrav.Enabled = True
            Call Llenar_Combo_con_Recordset(rsRepDesgrav, cboRepDesgrav)
            cboRepDesgrav.ListIndex = IndiceListaCombo(cboRepDesgrav, IIf(IsNull(rsSuger!cPersRepDesgrav), "", rsSuger!cPersRepDesgrav))
        Else
            cboRepDesgrav.Clear
            cboRepDesgrav.Enabled = False
        End If
        '********************************************************
        '***BRGO 20111104 ***************************************
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000008", sSTipoProdCod) Then     '**END ARLO
        'If sSTipoProdCod = "517" Then
            Dim i As Integer
'            Dim nPorcCEC As Double
'            Dim nComisionEC As Double
            Dim clsGen As COMDConstSistema.DCOMGeneral
            Dim rsCred As ADODB.Recordset
            Dim oCred As COMDCredito.DCOMCredito
            Dim nPorcGarant As Double
            
            Set oCred = New COMDCredito.DCOMCredito
            Set clsGen = New COMDConstSistema.DCOMGeneral
            SSTab1.TabVisible(2) = True
            Set oTipoCambio = New nTipoCambio
            nTC = oTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoDia)
            
            Set rsRel = clsGen.GetConstante(gColRelacPersInfoGas, "10,13,16,17")
            grdEmpVinculados.CargaCombo rsRel
            If Not rsRelEmp Is Nothing Then
                If Not rsRelEmp.EOF And Not rsRelEmp.BOF Then
                    Set grdEmpVinculados.Recordset = rsRelEmp
                    For i = 1 To Me.grdEmpVinculados.rows - 1
                        Me.grdEmpVinculados.TextMatrix(i, 4) = Format(Me.grdEmpVinculados.TextMatrix(i, 4), "#,000.00")
                        nComisionEC = nComisionEC + Me.grdEmpVinculados.TextMatrix(i, 4)
                    Next
                End If
            End If
            
            Set rsCred = oCred.RecuperaParametro(1030)
            nPorcCEC = rsCred!nParamValor
            If Not rsRelOtros Is Nothing Then
                If Not rsRelOtros.EOF And Not rsRelOtros.BOF Then
                    While Not rsRelOtros.EOF And Not rsRelOtros.BOF
                        If Trim(Right(rsRelOtros!cRelacion, 4)) = 13 Then
                            Me.txtMontoGarantia.Text = Format(rsRelOtros!nMontoAbono, "#,##0.00")
                            Me.txtCtaGarantia.Text = rsRelOtros!cCtaCodAbono
                            sPersOperador = rsRelOtros!cPersCod
                            sPersOperadorNombre = rsRelOtros!Nombre
                        End If
                        If Trim(Right(rsRelOtros!cRelacion, 4)) = 16 Then
                            Me.txtTasacion.Text = Format(rsRelOtros!nMontoAbono, "#,##0.00")
                        End If
                        If Trim(Right(rsRelOtros!cRelacion, 4)) = 17 Then
                            Me.lblComisionEC.Caption = Format(rsRelOtros!nMontoAbono, "#,##0.00")
                        End If
                        rsRelOtros.MoveNext
                    Wend
                End If
            Else
                Set rsCred = oCred.RecuperaParametro(3146)
                Me.txtTasacion.Text = Format(rsCred!nParamValor, "0.00")
                
                Set rsCred = oCred.RecuperaParametro(3143)
                Me.txtMontoGarantia.Text = Format(rsCred!nParamValor * nTC, "0.00")

                lblComisionEC.Caption = Format((nComisionEC + CDec(Me.txtTasacion.Text)) * nPorcCEC, "0.00")
                Set clsGen = Nothing
                sPersOperador = "": sPersOperadorNombre = ""
                Me.txtCtaGarantia.Text = ""
                Me.txtCtaGarantia.Enabled = False
            End If
            
            Set rsRel = Nothing
            Set oCred = Nothing
            Set rsCred = Nothing
            Set rsRelEmp = Nothing
            Set rsRelOtros = Nothing
        End If
        '***END BRGO *****************************************************
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000022", ActxCta.Prod) Then     '**END ARLO
        'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            nValorDiaGracia = 1
            txtFechaFija.Text = Format(rsSuger!dFechaPago, "DD/MM/YYYY")
            lnTasaPeriodoLeasing = (((1 + (rsSuger!nTasaPeriodoLeasing / 100)) ^ (1 / 12)) - 1)
            txtPerGra.Text = DateDiff("d", CDate(CDate(TxtFecDesemb.Text) + CDate(SpnPlazo.valor)), CDate(txtFechaFija.Text))
            If CInt(txtPerGra.Text) > 0 Then
                chkGracia.value = 1
            End If
            If CInt(SpnPlazo.valor) = 30 Then
                'opttper(1).value = True
                'cmbTipoPer.ListIndex = 1 'Comento 'JOEP20190125 CP
                cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
            End If
            Call txtFechaFija_KeyPress(13)
        End If
        
        'MAVM 20120402***
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000009", sSTipoProdCod) Then     '**END ARLO
        'If sSTipoProdCod = "801" Then
            FraTpoCalend.Enabled = True
            fratipodes.Enabled = True 'LUCV20180601, Según ERS022-2018
            ChkMiViv.value = 1
        Else
            FraTpoCalend.Enabled = False
            fratipodes.Enabled = False 'LUCV20180601, Según ERS022-2018
            ChkMiViv.value = 0
        End If
        '***
        
        'PESTAÑA GARANTIAS
        Call LimpiaFlex(FEGarantCred)
        'Dim oDCredito As COMDCredito.DCOMCredito
        Dim RGar As ADODB.Recordset
        Set oDCredito = New COMDCredito.DCOMCredito
        Set RGar = oDCredito.RecuperaGarantiasCredito(ActxCta.NroCuenta)
        Do While Not RGar.EOF
                FEGarantCred.AdicionaFila
                FEGarantCred.RowHeight(RGar.Bookmark) = 280
                FEGarantCred.TextMatrix(RGar.Bookmark, 1) = RGar!cTpoGarDescripcion
                FEGarantCred.TextMatrix(RGar.Bookmark, 2) = Format(RGar!nGravado, "#,#0.00")
                FEGarantCred.TextMatrix(RGar.Bookmark, 3) = Format(RGar!nTasacion, "#,#0.00")
                FEGarantCred.TextMatrix(RGar.Bookmark, 4) = Format(RGar!nRealizacion, "#,#0.00")
                'FEGarantCred.TextMatrix(RGar.Bookmark, 5) = Format(RGar!nPorGravar, "#,#0.00")
                FEGarantCred.TextMatrix(RGar.Bookmark, 5) = Format(RGar!nDisponible, "#,#0.00")
                FEGarantCred.TextMatrix(RGar.Bookmark, 6) = Trim(RGar!cPersNombre)
                FEGarantCred.TextMatrix(RGar.Bookmark, 7) = Trim(RGar!cNroDoc)
                FEGarantCred.TextMatrix(RGar.Bookmark, 8) = Trim(RGar!cTpoDoc)
                FEGarantCred.TextMatrix(RGar.Bookmark, 9) = Trim(RGar!cNumGarant)
                                  
                'MADM 20110506 * Num Garantia
                 'Dim oDGarantia As COMDCredito.DCOMGarantia
                 'Set oDGarantia = New COMDCredito.DCOMGarantia
                 'Dim rsGarantReal As ADODB.Recordset
                 'Set rsGarantReal = oDGarantia.RecuperaGarantiaRealMaxAprobacion(ActxCta.NroCuenta, RGar!cNumGarant)
                 'Set oDGarantia = Nothing

                'If Not (rsGarantReal.EOF Or rsGarantReal.BOF) Then
                '    'FEGarantCred.TextMatrix(rsGarantReal.Bookmark, 10) = Trim(rsGarantReal!nApruebaLegal)
                '    If (rsGarantReal!cNumGarant <> "") And ((DateDiff("d", rsGarantReal!dCertifGravamen, gdFecSis) > 365) Or (DateDiff("d", rsGarantReal!dTasacion, gdFecSis) > 730)) Then
                '         Select Case rsGarantReal!nVerificaLegal
                '            Case 1
                '                FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Pendiente"
                '            Case 2
                '                FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Aprobado"
                '            Case 3
                '                FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Desaprobado"
                '            Case 4
                '                FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Pendiente por Regularizar"
                '            Case 0
                '                FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Pendiente"
                '        End Select
                '        FEGarantCred.TextMatrix(RGar.Bookmark, 11) = Trim(IIf(Len(Trim(rsGarantReal!cnumpoliza)) > 1, IIf(Len(Trim(rsGarantReal!nEstadoPolizaNew)) > 1, rsGarantReal!nEstadoPolizaNew, "No registrado"), "No tiene"))
                '    Else
                '        FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Conforme"
                '        'FEGarantCred.TextMatrix(Rgar.Bookmark, 11) = "No tiene"
                '        FEGarantCred.TextMatrix(RGar.Bookmark, 11) = Trim(IIf(Len(Trim(rsGarantReal!cnumpoliza)) > 1, IIf(Len(Trim(rsGarantReal!nEstadoPolizaNew)) > 1, rsGarantReal!nEstadoPolizaNew, "No registrado"), "No tiene"))
                '    End If
                '    'FEGarantCred.TextMatrix(Rgar.Bookmark, 11) = Trim(rsGarantReal!nEstadoPolizaNew)
                'Else
                '    FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "No aplica"
                '    FEGarantCred.TextMatrix(RGar.Bookmark, 11) = "No tiene"
                'End If
                
                'END MADM
                'EJVG20150515 ***
                FEGarantCred.TextMatrix(RGar.Bookmark, 10) = RGar!cVerificaLegal
                FEGarantCred.TextMatrix(RGar.Bookmark, 11) = RGar!cVerificaPoliza
                'END EJVG *******
                RGar.MoveNext
        Loop
        'RGar.Close
        'Set RGar = Nothing
        Set oDCredito = Nothing
        'rsGarantReal.Close
        'Set rsGarantReal = Nothing
        'nMontoGravado = 0
    
        'Set oCredito = New DCredito
        'nMontoGravado = oCredito.RecuperaMontoGarantiaCredito(ActxCta.NroCuenta, gdFecSis)
        Set oDCredito = New COMDCredito.DCOMCredito
        LblTotGarantiasGrav.Caption = Format(oDCredito.RecuperaMontoGarantiaCredito(ActxCta.NroCuenta, gdFecSis), "#0.00")
        Set oDCredito = Nothing
        '****
        
        Dim oCredExo As COMNCredito.NCOMCredito
        Set oCredExo = New COMNCredito.NCOMCredito
            If oCredExo.ValidaExonSeguroDesgravamen(lblcod, gdFecSis) Then
                'ChkExoneracion.value = 1
                ChkExoSegDes.value = 1 'APRI20171122 ERS028-2017 MEJORA
            End If
        Set oCredExo = Nothing
        
        ChkBloqueo.value = IIf(rsSuger!bBloqueo, 1, 0)
        ChkExoneracion.value = rsSuger!nExoneraReg
        ChkExoSegDes.value = rsSuger!nExoSeguroDes
        chkDesembBcoNac.value = rsSuger!nDesBcoNac
        actxCtaAho.NroCuenta = rsSuger!cCtaCodAho
        'ALPA 20141030**********************************
        'Dim oCredExo As COMNCredito.NCOMCredito
        Set oCredExo = New COMNCredito.NCOMCredito
            If oCredExo.ValidaExoneracion(psCtaCod, "TIP0009") Then
                Txtinteres.Text = rsSuger!nTasaExononeracion
                chkTasa.value = 1
                txtInteresTasa.Text = rsSuger!nTasaInteres
                
            Else
                Txtinteres.Text = rsSuger!nTasaInteres
                chkTasa.value = 0
                txtInteresTasa.Text = 0
                
            End If
        Set oCredExo = Nothing
        '**********************************************
        Call CargarDatosProductoCrediticio 'ALPA20140918
        Call TxtFecDesemb_KeyPress(13) 'ALPA20141209
        Call cmbSubtipo_Click 'ALPA20141209
        lnLogicoBuscarDatos = 1
        'WIOR 20151223 ***
        txtMonSug.Enabled = True
        FraTpoCalend.Enabled = False
        fratipodes.Enabled = False 'LUCV20180601, Según ERS022-2018
        cmbDesemb.Enabled = False 'LUCV20180601, Según ERS022-2018
        ChkMiViv.Enabled = False
        ChkMiViv.value = 0
        cmdMIVIVIENDA.Enabled = False
        
        If fbMIVIVIENDA Then
            txtMonSug.Enabled = False
            FraTpoCalend.Enabled = True
            fratipodes.Enabled = True 'LUCV20180601, Según ERS022-2018
            cmbDesemb.Enabled = True 'LUCV20180601, Según ERS022-2018
            ChkMiViv.Enabled = False
            ChkMiViv.value = 1
            cmdMIVIVIENDA.Enabled = True
        
            Set oDCredito = New COMDCredito.DCOMCredito
            Dim rsMiViv As ADODB.Recordset
            Set rsMiViv = oDCredito.ObtenerDatosNuevoMIVIVIENDA(ActxCta.NroCuenta, gColocEstSug)
            
            If Not (rsMiViv.EOF And rsMiViv.BOF) Then
                ReDim fArrMIVIVIENDA(11)
                fArrMIVIVIENDA(0) = CDbl(rsMiViv!nMontoVivienda)
                fArrMIVIVIENDA(1) = CDbl(rsMiViv!nCuotaInicial)
                fArrMIVIVIENDA(2) = CDbl(rsMiViv!nBonoOtorgado)
                fArrMIVIVIENDA(3) = CDbl(rsMiViv!nMOntoCred)
                fArrMIVIVIENDA(4) = CDbl(rsMiViv!nUIT)
                fArrMIVIVIENDA(5) = CLng(rsMiViv!nDesde)
                fArrMIVIVIENDA(6) = CLng(rsMiViv!nHasta)
                fArrMIVIVIENDA(7) = CDbl(rsMiViv!nBono)
                fArrMIVIVIENDA(8) = CDbl(rsMiViv!nMinCredUIT)
                fArrMIVIVIENDA(9) = 1
                fArrMIVIVIENDA(10) = CInt(rsMiViv!nPeriodoPerdBono)
            End If
            Set oDCredito = Nothing
        End If
        
        fbDatosCargados = True
        'WIOR FIN ********
        'WIOR 20160224 ***
        Set oNCredito = New COMNCredito.NCOMCredito
        fnTasaSegDes = oNCredito.ObtenerTasaSeguroDesg(ActxCta.NroCuenta, gdFecSis, fnCantAfiliadosSegDes)
        Set oNCredito = Nothing
        'WIOR FIN ********
    Else
        CargaDatos = False
        lnLogicoBuscarDatos = 0
        fbDatosCargados = False 'WIOR 20151223
    End If
    'WIOR 20120517 *************************************************************
    cmbMicroseguro.ListIndex = IndiceListaCombo(cmbMicroseguro, nMicroseguro)
    cmbBancaSeguro.ListIndex = IndiceListaCombo(cmbBancaSeguro, nMultiriesgo)
    'WIOR - FIN ****************************************************************
    lnMostrarCSP = 1
    'EAAS20180827 SEGUN ERS-05-2018
    cmbAguaSaneamientoDet.Visible = False
    cmbCreditoVerdeDet.Visible = False 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    bValidaCargaSugerenciaAguaSaneamiento = 1
    bValidaCargaSugerenciaCreditoVerde = 1 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    If (UBound(fvListaCreditoVerde) > 0) Then
        cmbCreditoVerdeDet.Visible = True
    End If
    If (UBound(fvListaAguaSaneamiento) > 0) Then
        cmbAguaSaneamientoDet.Visible = True
    End If
    'EAAS20180827 SEGUN ERS-054-2018
    lblSolModif.Enabled = False 'LUCV20180601, Agregó según ERS022-2018
    Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"
End Function
'ALPA 20150113***********************************
Private Sub CargarDatosProductoCrediticio()
If Trim(ActxCta.Prod) <> "" Then
Dim sCodigo As String
Dim sCtaCodOrigen As String 'DAOR 20070407, para el caso de refinanciados
Dim oLineas As COMDCredito.DCOMLineaCredito
txtBuscarLinea.Text = ""
lblLineaDesc.Caption = ""
sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
Set oLineas = New COMDCredito.DCOMLineaCredito
Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio(sSTipoProdCod, lnCampanaId, Trim(Right((txtBuscarLinea.psDescripcion), 15)), sCodigo, lblLineaDesc, Mid(ActxCta.NroCuenta, 9, 1), CCur(txtMonSug.Text), IIf(ckcPreferencial.value = 1, 1, 0))
Set oLineas = Nothing
       If RLinea.RecordCount > 0 Then
          If txtBuscarLinea.Text = "" Then
            txtBuscarLinea.Text = "XXX"
          End If
          Call CargaDatosLinea
          If txtBuscarLinea.Text = "XXX" Then
            txtBuscarLinea.Text = ""
          End If
          Set objProducto = New COMDCredito.DCOMCredito
          If objProducto.GetResultadoCondicionCatalogo("N0000023", ActxCta.Prod) Then     '**END ARLO
          'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            Txtinteres.Text = lnTasaPeriodoLeasing * 100
            TxtTasaGracia.Text = lnTasaPeriodoLeasing * 100
          End If
       Else
            lnTasaInicial = 0
            lnTasaFinal = 0
            
            'JOEP ERS007-2018 20180210
            lnTasaGraciaInicial = 0
            lnTasaGraciaFinal = 0
            'JOEP ERS007-2018 20180210
            
           'LUCV20180601. Agregó según ERS022-2018
            Txtinteres.Text = Format(lnTasaFinal, "#0.00")
            txtInteresTasa.Text = Format(lnTasaFinal, "#0.00")
        
            Txtinteres.Visible = True
            Txtinteres.Enabled = False
            LblInteres.Enabled = False
            LblInteres.Visible = False
            'Fin LUCV20180601
            
            If nMostrarLineaCred = 0 Then
                MsgBox "No existen Líneas de Crédito con el Plazo seleccionado", vbInformation, "Aviso"
                txtBuscarLinea.Text = ""
                lblLineaDesc = ""
            End If
       End If
End If
End Sub

'Private Sub chkTasa_Click()
'    If chkTasa.value = 0 Then
'        txtInteresTasa.Enabled = False
'        txtInteresTasa.Visible = False
'        txtInteresTasa.Text = 0#
'    Else
'        txtInteresTasa.Enabled = True
'        txtInteresTasa.Visible = True
'        txtInteresTasa.Text = Format(Txtinteres.Text, "#0.0000")
'    End If
'    Call ExoneraTipoTasa
'End Sub
'Private Sub ExoneraTipoTasa()
'    Dim lnFila As Integer
'    Dim rs As ADODB.Recordset
'    Call LimpiaFlex(frmCredSugExonera.feTiposExonera)
'    'Do While Not rs.EOF
'        frmCredSugExonera.feTiposExonera.AdicionaFila
'        lnFila = frmCredSugExonera.feTiposExonera.row
'        frmCredSugExonera.feTiposExonera.TextMatrix(lnFila, 1) = "TIP0009"
'        frmCredSugExonera.feTiposExonera.TextMatrix(lnFila, 2) = "TASA"
'        frmCredSugExonera.feTiposExonera.TextMatrix(lnFila, 3) = "A"
'        frmCredSugExonera.feTiposExonera.TextMatrix(lnFila, 4) = "."
'    'Loop
'    frmCredSugExonera.feTiposExonera.TopRow = 1
'    frmCredSugExonera.feTiposExonera.row = 1
'End Sub
'*****************************************************
Private Sub ActxCta_KeyPress(KeyAscii As Integer)

     If KeyAscii = 13 Then
            If CargaDatos(ActxCta.NroCuenta) Then
                cmdrelac.Enabled = True
                FraDatos.Enabled = True
'                If Cmblincre.Visible Then
'                    Cmblincre.SetFocus
'                End If
                
                CmdGrabar.Enabled = True
                CmdCalend.Enabled = True
                CmdGastos.Enabled = True
                'cmdEvaluacion.Enabled = True
                ActxCta.Enabled = False
                CmdCredVig.Enabled = True
                'ALPA 20091007***********************
                cmdVinculados.Enabled = True
                Frame3.Enabled = True
                '************************************
                cmdCheckList.Enabled = True 'RECO20150415 ERS010-2015
                'Para la busqueda automática
                Call HabilitaPermiso
                
                If txtBuscarLinea.Visible Then
                    If Len(txtBuscarLinea.Text) > 0 Then
                        If CmdGrabar.Enabled Then CmdGrabar.SetFocus
                    Else
                        If txtBuscarLinea.Enabled Then txtBuscarLinea.SetFocus
                    End If
                End If
                
            'MADM 20100517
            'If Mid(ActxCta.NroCuenta, 6, 3) = "302" Then
            '  If sSTipoProdCod = "703" Then
            '    cmdSeleccionaFuente.Enabled = False
            '    cmdFuentes.Enabled = False
            '    Label13.Enabled = False
            'Else
            '    cmdSeleccionaFuente.Enabled = True
            '    cmdFuentes.Enabled = True
            '    Label13.Enabled = True
            'End If
            'END MADM
            
            ''** JUEZ 20120907 ******************************************
            'If nAgenciaCredEval = 0 Then
            '    If sSTipoProdCod = "703" Then
            '        cmdSeleccionaFuente.Enabled = False
            '        cmdFuentes.Enabled = False
            '        'Label13.Enabled = False
            '    Else
            '        cmdSeleccionaFuente.Enabled = True
            '        cmdFuentes.Enabled = True
            '        'Label13.Enabled = True
            '    End If
            'Else
            '    cmdSeleccionaFuente.Enabled = False
            '    cmdFuentes.Enabled = False
            '    'Label13.Enabled = False
            'End If
            ''** END JUEZ ***********************************************
                
            
                'CUSCO
                CboPersCiiu.Enabled = True
                'FRHU 20170517 ACTA-070-2017
                If Trim(LeeConstanteSist(605)) = "1" Then
                    Txtinteres.Locked = True
                    LblInteres.Enabled = False
                End If
                'FIN FRHU 20170517
                
                 'JOEP20180204 CP
                If lnColocCondicion = 4 Then
                    txtMonSug.Enabled = False
                Else
                    txtMonSug.Enabled = True
                End If
                'JOEP20180204 CP
            Else
                cmdrelac.Enabled = False
                FraDatos.Enabled = False
                CmdGrabar.Enabled = False
                CmdCalend.Enabled = False
                CmdGastos.Enabled = False
                'cmdEvaluacion.Enabled = False
                ActxCta.Enabled = True
                CmdCredVig.Enabled = False
                
                CboPersCiiu.Enabled = False
                'ALPA 20091007***********************
                cmdVinculados.Enabled = False
                '************************************
                MsgBox "El Credito No Existe", vbInformation, "Aviso"
                ''JUEZ 20120914 ***************************************************
                'If nAgenciaCredEval = 1 Then
                '    If nVerifCredEval = 0 Then
                '        MsgBox "El Credito no ha sido verificado por el Coordinador de Creditos", vbInformation, "Aviso"
                '    Else
                '        MsgBox "El Credito No Existe", vbInformation, "Aviso"
                '    End If
                'ElseIf bVerifNivApr = False Then
                '    'Unload Me
                'Else
                '    MsgBox "El Credito No Existe", vbInformation, "Aviso"
                'End If
                ''END JUEZ ********************************************************
            End If
     End If
End Sub

Private Sub actxCtaAho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'CmdCalend.Enabled = False FRHU20140212
         CmdCalend.Enabled = False 'FRHU 20140422 - Observacion
        If ValidaCuenta(Me.actxCtaAho.NroCuenta) = False Then Exit Sub
        If Mid(Me.ActxCta.NroCuenta, 9, 1) <> Mid(Me.actxCtaAho.NroCuenta, 9, 1) Then
            MsgBox "Moneda de Cuentas de Credito y Ahorro no coinciden", vbInformation, "Aviso"
            Exit Sub
        End If
        Dim sCta As String
        sCta = actxCtaAho.NroCuenta
        ObtieneDatosCuenta sCta
    End If
End Sub

'ALPA20141126********************************************************************
Private Sub chkCSP_Click()
If chkCSP.value Then
    If lnMostrarCSP = 1 Then
        lnCSP = frmCredPolizaCobrar.MostrarCuotaCobrar(lnCSP, spnCuotas.valor)
    End If
End If
End Sub

'WIOR 20131129 *************************
Private Sub chkCuotaBalon_Click()
If chkCuotaBalon.value = 1 Then
    If CInt(spnCuotas.valor) < 2 Then
        chkCuotaBalon.value = 0
        txtCuotaBalon.Text = "0"
    Else
        txtCuotaBalon.Text = "1"
        txtCuotaBalon.Enabled = True
    End If
Else
    txtCuotaBalon.Enabled = False
    txtCuotaBalon.Text = "0"
End If
End Sub
'WIOR FIN ******************************

Private Sub chkGracia_Click()
txtPerGra.Enabled = False 'MAVM 25102010
If chkGracia.value = 1 Then
    fraGracia.Enabled = True
    '->***** LUCV20180601, Comentó y agregó según ERS022-2018
    'TxtTasaGracia.Visible = True
    'TxtTasaGracia.Enabled = True
    fbTieneTasaGracia = True
    TxtTasaGracia.Enabled = False
    'LblTasaGracia.Visible = False
    fbNoTieneTasaGracia = False
    '<-***** Fin LUCV20180601
    chkIncremenK.value = 0
Else
     fraGracia.Enabled = False
    'txtPerGra.Text = "0.00" Comentado Por MAVM 25102010
    txtPerGra.Text = "0" 'MAVM 25102010
    TxtTasaGracia.Text = "0.00"
    TxtTasaGracia.Visible = False
    TxtTasaGracia.Enabled = False
    
    'LblTasaGracia.Visible = True 'LUCV20180601, Comentó Según ERS022-2018
    fbTieneTasaGracia = False  'LUCV20180601, Agregó según ERS022-2018
    fbNoTieneTasaGracia = True 'LUCV20180601, Agregó según ERS022-2018
    
    'optTipoGracia(0).value = False 'LUCV20180601, Coment? seg?n ERS022-2018
    'optTipoGracia(1).value = False 'LUCV20180601, Coment? seg?n ERS022-2018
    'MAVM 25102010 ***
'    GenerarFechaPago
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000024", ActxCta.Prod) Then     '**END ARLO
     'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
     Else
        GenerarFechaPago
     End If
        'If opttper(1).value = True Then
        If Trim(Right(cmbTipoPer.Text, 2)) = "1" Then
            chkGracia.Enabled = False
        End If
    '***
End If
'cmdgracia.Enabled = True 'LUCV20180601, Comentó según ERS022-2018
End Sub

Private Sub ChkTrabajadores_Click()
    If ChkTrabajadores.value = 1 Then
        FraTipoCuota.Enabled = False
        fratipodes.Enabled = False
        FraCalendario.Enabled = False
        'ChkCuotaCom.Enabled = False 'LUCV20180601, Comentó según ERS022-2018
        ChkMiViv.Enabled = False
        FraGastos.Enabled = False
   Else
        FraTipoCuota.Enabled = True
        fratipodes.Enabled = True
        FraCalendario.Enabled = True
        'ARCV 20-02-2007
        'ChkCuotaCom.Enabled = True
        'ChkMiViv.Enabled = True
        '-------
        FraGastos.Enabled = True
   End If
End Sub
'INICIO EAAS SEGUN ERS-054-2018
Private Sub cmbAguaSaneamientoDet_Click()
'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
nMontoCreditoVariable = CDbl(txtMonSug.Text) - nSumaAguaSaneamiento - nSumaCreditoVerde
nCentinela = 0
If (nMontoCreditoVariable <> CDbl(txtMonSug.Text) Or nMontoCreditoVariable = 0) Then 'EAAS20190410 SEGUN 018-GM-DI_CMACM
nCentinela = 1
End If
'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
Call frmCredAguaSaneamiento.Inicio(fvListaAguaSaneamiento, CInt(Right(cmbProducto.Text, 2)), cmbDestino.Text, nMontoCreditoVariable, ActxCta.NroCuenta, nCentinela, nSumaAguaSaneamiento) 'EAAS20191401 SEGUN 018-GM-DI_CMACM nMontoCreditoVariable nCentinela, nSumaAguaSaneamiento
End Sub

'FIN EAAS SEGUN ERS-054-2018
Private Sub cmbCreditoVerdeDet_Click()
'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
nMontoCreditoVariable = CDbl(txtMonSug.Text) - nSumaAguaSaneamiento - nSumaCreditoVerde
nCentinela = 0
If (nMontoCreditoVariable <> CDbl(txtMonSug.Text) And nMontoCreditoVariable <> 0) Then
nCentinela = 1
End If
'FIN EAAS20191401 SEGUN 018-GM-DI_CMACM
Call frmCredVerde.Inicio(fvListaCreditoVerde, CInt(Right(cmbProducto.Text, 2)), cmbDestino.Text, nMontoCreditoVariable, ActxCta.NroCuenta, nSumaCreditoVerde) 'EAAS20191401 SEGUN 018-GM-DI_CMACM nMontoCreditoVariable nSumaCreditoVerde

End Sub

'Private Sub ckcPreferencial_Click()
'    If lnLogicoBuscarDatos = 1 Then
'        If lnCliPreferencial = 1 Then
'            ckcPreferencial.value = 1
'        Else
'            ckcPreferencial.value = 0
'        End If
'        Call CargarDatosProductoCrediticio
'        Call MostrarLineas
'    End If
'End Sub

'WIOR 20120510*********************************
Private Sub cmbBancaSeguro_Click()
If cmbBancaSeguro.Text <> "" Then
    If Trim(Right(cmbBancaSeguro.Text, 4)) <> "0" Then
        fbMultiriesgo = True
        'frmGarantMultiriesgoMYPE.Inicia lblcod.Caption, lblMonto.Caption, ActxCta.NroCuenta, spnCuotas.valor  'RECO20160405
    Else
        fbMultiriesgo = False
    End If
Else
    fbMultiriesgo = False
End If
End Sub

Private Sub cmbDesBcoNac_Click()
    frmCredDesembBcoNac.Show 1
End Sub

Private Sub cmbDesemb_Click()
    ReDim MatDesPar(0, 0)
    ReDim MatCalend(0, 0)
    ReDim MatDesPar(0, 0)
    ReDim MatrizCal(0, 0)

    If Trim(Right(cmbDesemb.Text, 2)) = "0" Then 'Si Desembolso Total
        CmdDesembolsos.Enabled = False
        txtMonSug.Enabled = True
        'txtMonSug.Text = "0.00"
        'spnCuotas.valor = 30
        spnCuotas.Enabled = True
    Else
        CmdDesembolsos.Enabled = True
        txtMonSug.Enabled = False
        txtMonSug.Text = "0.00"
        spnCuotas.valor = 1
        spnCuotas.Enabled = False
    End If
End Sub

Private Sub CmbDestino_Click()
    'FRHU20140212-RQ14005 se agrego sSTipoProdCod = "706"
    'If cmbDestino.ListIndex = 5 Or (cmbDestino.ListIndex = 6 And (sSTipoProdCod = "510" Or sSTipoProdCod = "511" Or sSTipoProdCod = "706")) Then
    Dim oDCredito As New COMDCredito.DCOMCredito 'FRHU 20171107 : Inicio
    Dim lnDestino As Integer
    If cmbDestino.Text = "" Then
        lnDestino = cmbDestino.ListIndex
    Else
        lnDestino = CInt(Trim(Right(cmbDestino.Text, 2)))
    End If
    If oDCredito.HabilitarDesembolsoTercero(ActxCta.NroCuenta, lnDestino, sSTipoProdCod) Then 'FRHU 20171107 : Fin
        frmCtaAho.Enabled = True
        CmdCalend.Enabled = False
        CmdGastos.Enabled = False
    Else
        frmCtaAho.Enabled = False
        actxCtaAho.Cuenta = ""
        lblTercero = ""
        CmdCalend.Enabled = True
        CmdGastos.Enabled = True
    End If
    
    'INICIO EAAS20180815
        If (bValidaCargaSugerenciaAguaSaneamiento <> 0) Then
            If (UBound(fvListaAguaSaneamiento) > 0) Then
                MsgBox "Se esta cambiando el destino, el detalle agua y saneamiento se limpiará", vbInformation, "Alerta"
                nSumaAguaSaneamiento = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
                nSumaCreditoVerde = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
            End If
        Dim rsValDesAguaSaneamientoC As ADODB.Recordset
        Dim obDCredValDesAguaSaneamientoC As COMDCredito.DCOMCredito
        Set obDCredValDesAguaSaneamientoC = New COMDCredito.DCOMCredito
        
        
        Set rsValDesAguaSaneamientoC = obDCredValDesAguaSaneamientoC.ValidadDestinoConsEmpAguaSaneamiento(CInt(IIf(Trim(Right(cmbProducto.Text, 5)) = "", 0, Trim(Right(cmbProducto.Text, 5)))), CInt(IIf(Trim(Right(cmbDestino.Text, 5)) = "", 0, Trim(Right(cmbDestino.Text, 5)))), IIf(lblcod.Caption = "", 0, lblcod.Caption))
            If Not (rsValDesAguaSaneamientoC.EOF And rsValDesAguaSaneamientoC.BOF) Then
                If rsValDesAguaSaneamientoC!cMensaje = "" Then
                    
                    cmbAguaSaneamientoDet.Visible = True
                    rsValDesAguaSaneamientoC.Close
                    Set obDCredValDesAguaSaneamientoC = Nothing
  
                Else
                cmbAguaSaneamientoDet.Visible = False
                End If

            End If

    
    ReDim fvListaAguaSaneamiento(0)
    'EAAS 20180727 ERS-054-2018
    End If
       'INICIO EAAS20191401 SEGUN 018-GM-DI_CMACM
        If (bValidaCargaSugerenciaCreditoVerde <> 0) Then
            If (UBound(fvListaCreditoVerde) > 0) Then
                MsgBox "Se esta cambiando el destino, el detalle Eco Ahorro se limpiará", vbInformation, "Alerta"
                nSumaAguaSaneamiento = 0
                nSumaCreditoVerde = 0
            End If
        Dim rsValDesCreditoVerdeC As ADODB.Recordset
        Dim obDCredValDesCreditoVerdeC As COMDCredito.DCOMCredito
        Set obDCredValDesCreditoVerdeC = New COMDCredito.DCOMCredito


        Set rsValDesCreditoVerdeC = obDCredValDesCreditoVerdeC.ValidadDestinoConsEmpCreditoVerde(CInt(IIf(Trim(Right(cmbProducto.Text, 5)) = "", 0, Trim(Right(cmbProducto.Text, 5)))), CInt(IIf(Trim(Right(cmbDestino.Text, 5)) = "", 0, Trim(Right(cmbDestino.Text, 5)))), IIf(lblcod.Caption = "", 0, lblcod.Caption))
            If Not (rsValDesCreditoVerdeC.EOF And rsValDesCreditoVerdeC.BOF) Then
                If rsValDesCreditoVerdeC!cMensaje = "" Then

                    cmbCreditoVerdeDet.Visible = True
                    rsValDesCreditoVerdeC.Close
                    Set obDCredValDesCreditoVerdeC = Nothing

                Else
                cmbCreditoVerdeDet.Visible = False
                End If

            End If

    
    ReDim fvListaCreditoVerde(0)
    'EAAS20191401 SEGUN 018-GM-DI_CMACM
    End If
    
End Sub

Private Sub cmbMicroseguro_Click()
If Me.cmbMicroseguro.Text <> "" Then
    If Trim(Right(Me.cmbMicroseguro.Text, 4)) <> "0" Then
        fbMicroseguro = True
        fnMicroseguro = Int(Trim(Right(Me.cmbMicroseguro.Text, 4)))
    Else
        fbMicroseguro = False
        fnMicroseguro = 0
    End If
Else
    fbMicroseguro = False
    fnMicroseguro = 0
End If
End Sub

Private Sub cmbProducto_Click()
    Call CargaSubProducto(Trim(Right(cmbProducto.Text, 3)))
    Call VerificarMIVIVIENDA 'WIOR 20151223
End Sub

Private Sub CargaSubProducto(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubProducto
    Set oCred = New COMDCredito.DCOMCredito
    Set RTemp = oCred.RecuperaSubProductosCrediticios(psTipo, gsCodCargo) 'NAGL 20171121
    Set oCred = Nothing
    cmbSubProd.Clear
    Do While Not RTemp.EOF
        cmbSubProd.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbSubProd, 250)
    Exit Sub
    
ERRORCargaSubProducto:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmbSubProd_Click()
    sSTipoProdCod = Trim(Right(cmbSubProd.Text, 3))
    Call cmbSubtipo_Click
    Call VerificarMIVIVIENDA 'WIOR 20151223
End Sub

'WIOR FIN *****************************************
Private Sub cmbSubtipo_Click()
Dim oLineas As COMDCredito.DCOMLineaCredito
Dim lrsLineas As ADODB.Recordset
Set lrsLineas = New ADODB.Recordset

'Set oLineas = New COMDCredito.DCOMLineaCredito
'Set lrsLineas = oLineas.RecuperaLineasProductoArbol(Right(cmbSubTipo.Text, 3), Mid(ActxCta.NroCuenta, 9, 1), , Mid(ActxCta.NroCuenta, 4, 2), SpnPlazo.valor, CDbl(lblMonto.Caption), spnCuotas.valor)
'Set oLineas = Nothing

'COMENTADO X MADM 20110419 - Refinanciado
'If Not bEsRefinanciado Then
'    Set oLineas = New COMDCredito.DCOMLineaCredito
'    Set lrsLineas = oLineas.RecuperaLineasProductoArbol(Right(cmbSubTipo.Text, 3), Mid(ActxCta.NroCuenta, 9, 1), , Mid(ActxCta.NroCuenta, 4, 2), spnPlazo.valor, CDbl(lblmonto.Caption), spnCuotas.valor)
'    Set oLineas = Nothing
'Else
'    Set oLineas = New COMDCredito.DCOMLineaCredito
'    Set lrsLineas = oLineas.RecuperaLineasCredOrigenRefinanciadoArbol(ActxCta.NroCuenta)
'    Set lrsLineas = oLineas.RecuperaLineasCredOrigenRefinanciadoArbolNew(ActxCta.NroCuenta, Right(cmbSubTipo.Text, 3), Mid(ActxCta.NroCuenta, 9, 1), , Mid(ActxCta.NroCuenta, 4, 2), spnPlazo.valor, CDbl(lblmonto.Caption), spnCuotas.valor)
'    Set oLineas = Nothing
'End If
'txtBuscarLinea.rs = lrsLineas
Call MostrarLineas 'ALPA 20150113**********
'ALPA 20141127*****************************
If Trim(Right(cmbSubTipo.Text, 10)) = "853" Then
    chkCSP.Visible = True
Else
    chkCSP.Visible = False
End If
'******************************************
bCheckList = False 'RECO20150513 *****************************

Call VerificarMIVIVIENDA 'WIOR 20151223
End Sub
Private Sub MostrarLineas()
'    Dim nTasCom As Currency
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
'    nTasCom = Val(Txtinteres.Text)
    Dim oLineas As COMDCredito.DCOMLineaCredito
    Dim lrsLineas As ADODB.Recordset
    Set lrsLineas = New ADODB.Recordset
    
    Set oLineas = New COMDCredito.DCOMLineaCredito
    Set lrsLineas = oLineas.RecuperaLineasProductoArbol(Right(cmbSubTipo.Text, 3), Mid(ActxCta.NroCuenta, 9, 1), , Mid(ActxCta.NroCuenta, 4, 2), IIf(Trim(SpnPlazo.valor) = "", 0, SpnPlazo.valor), CDbl(val(txtMonSug.Text)), IIf(Trim(spnCuotas.valor) = "", 0, spnCuotas.valor), IIf(Trim(Txtinteres.Text) = "", 0, Txtinteres.Text), IIf(Trim(txtPerGra.Text) = "", 0, txtPerGra.Text), gdFecSis, lblcod.Caption)
    Set oLineas = Nothing
    txtBuscarLinea.rs = lrsLineas
End Sub
Private Sub cmbTipoCalend_Click()
    If ChkMiViv.value = 1 Then
        cmbTipoCalend.ListIndex = 0
    End If
End Sub

Private Sub cmbTipoCuota_Click()
    If ChkMiViv.value = 1 Then
        cmbTipoCuota.ListIndex = 0
    End If
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
    If Trim(Right(cmbTipoCuota.Text, 2)) <> "3" Then
        DeshabilitaTipoPeriodo True, True
        DeshabilitaTipoGracia True
        DeshabilitaTipoCalend True
        DeshabilitaTipoDesemb True
        'If Optdesemb(1).value Then
        If Trim(Right(cmbDesemb.Text, 2)) = "1" Then
            CmdDesembolsos.Enabled = True
        End If
        ReDim MatrizCal(0, 0)
    Else
        DeshabilitaTipoPeriodo False, False
        DeshabilitaTipoGracia False
        DeshabilitaTipoPeriodo False, False
        DeshabilitaTipoGracia False
        DeshabilitaTipoCalend False
        DeshabilitaTipoDesemb True
        txtPerGra.Text = "0"
        'opttper(0).value = True
        'cmbTipoPer.ListIndex = 0 'Cometno JOEP20180227 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
    End If
End Sub

Private Sub cmbTipoPer_Click()
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
    
    If ChkMiViv.value = 1 Then
        'cmbTipoPer.ListIndex = 1 'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
    End If
    
    'If Index = 1 Then
    If Trim(Right(cmbTipoPer.Text, 2)) = "1" Then
        HabilitaFechaFija (True)
        'Comentado Por MAVM 25102010 ***
        'If TxtDiaFijo.Enabled Then
        '    TxtDiaFijo.SetFocus
        'End If
        '***
        'No aplica para la fecha fija la Gracia en Cuotas
        'optTipoGracia(1).Enabled = False 'LUCV20180601, Comentó según ERS022-2018
        'optTipoGracia(1).value = False 'LUCV20180601, Comentó según ERS022-2018
        'Activa el ingreso de fecha fija para el calculo de dias de gracia GITU 19-08-2008
        'txtFechaFija.Enabled = True
        'txtFechaFija.SetFocus
        
        'MAVM 25102010 ***
        txtFechaFija.Enabled = True
        TxtDiaFijo.Enabled = False
        TxtDiaFijo2.Enabled = False
        chkGracia.value = 0
        chkGracia.Enabled = False
        fbTieneTasaGracia = False 'LUCV20180601, según ERS022-2018
        fbNoTieneTasaGracia = True 'LUCV20180601, Según ERS022-2018
        txtFechaFija.Text = gdFecSis
        'ALPA 20111209
        'GenerarFechaPago
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000025", ActxCta.Prod) And nValorDiaGracia = 1 Then
        'If (ActxCta.Prod = "515" Or ActxCta.Prod = "516") And nValorDiaGracia = 1 Then
            txtFechaFija.Text = txtFechaFija.Text
        Else
            txtFechaFija.Text = gdFecSis
            GenerarFechaPago
        End If
        '***********************
    Else
        HabilitaFechaFija (False)
        
        'MAVM 25112010 ***
        txtFechaFija.Text = "__/__/____"
        TxtFecDesemb.Text = gdFecSis
        GenerarFechaPago
        'SpnPlazo.SetFocus
        chkGracia.value = 0
        TxtDiaFijo.Text = "00"
        chkGracia.Enabled = True
        '***
    End If
End Sub

Private Sub cmdActTipoCred_Click()
Dim oDCredito As COMDCredito.DCOMCredito
Dim lnTipoCredito As Integer
If Mid(sSTipoProdCod, 1, 2) <> Mid(gColProConsumo, 1, 2) Then
    If nActualizaTipoCred = 0 Then
        lblMsj.Visible = True
        DoEvents
        nActualizaTipoCred = 1
        Set oDCredito = New COMDCredito.DCOMCredito
        lnTipoCredito = oDCredito.ObtenerTipoCreditoxTipificacion(lblcod.Caption)
        cmbTipoCredito.ListIndex = IndiceListaCombo(cmbTipoCredito, lnTipoCredito)
        Call cmbTipoCredito_Click
        Set oDCredito = Nothing
        lblMsj.Visible = False
        DoEvents
    Else
        If MsgBox("El proceso para determinar el tipo de credito ya fue realizado, Desea volver a realizarlo ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            nActualizaTipoCred = 0
            Call cmdActTipoCred_Click
        End If
    End If
End If
End Sub
'*** BRGO 20111103 *********************************************************
Private Sub cmdAgregar_Click()
    If Me.grdEmpVinculados.rows - 1 >= 1 And Me.grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.rows - 1, 5) = "" And Me.grdEmpVinculados.TextMatrix(Me.grdEmpVinculados.row - 1, 1) <> "" Then
        MsgBox "Falta ingresar la Cuenta de Ahorro"
        Exit Sub
    End If
    If grdEmpVinculados.rows <= 4 Then
        grdEmpVinculados.AdicionaFila
        grdEmpVinculados.SetFocus
        grdEmpVinculados.TextMatrix(grdEmpVinculados.row, 4) = "0.00"
        SendKeys "{ENTER}"
        grdEmpVinculados.TipoBusqueda = BuscaPersona
    Else
        MsgBox "El registro de Empresas Vinculadas está completo"
    End If
End Sub
'*** END BRGO **************************************************************
'RECO20150421 ERS010-2015 **************************************
Private Sub cmdCheckList_Click()
'JOEP20190125 CP
    Dim objGarN As COMDCredito.DCOMCredito
    Dim rsGarN As ADODB.Recordset
    Set objGarN = New COMDCredito.DCOMCredito
    Set rsGarN = objGarN.CP_getValCobGar(ActxCta.NroCuenta)
    
    If Not CP_ValMensaje(2) Then Exit Sub
    
    If Trim(Right(cmbSubProd.Text, 9)) <> "703" Then
        'EJVG20160713 ***
        fbEliminarEvaluacion = False
        GenerarDataExposicionEsteCredito ActxCta.NroCuenta, CDbl(txtMonSug.Text), fnMontoExpEsteCred_NEW 'Seteamos el valor de la nueva exposición
        If NecesitaFormatoEvaluacion(ActxCta.NroCuenta, 2001, CInt(Left(sSTipoProdCod, 1) & "00"), CInt(sSTipoProdCod), fnMontoExpEsteCred_NEW, fbEliminarEvaluacion) Then
            Exit Sub
        End If
        'END EJVG *******
    End If
    
    If Not (rsGarN.BOF And rsGarN.EOF) Then
        If rsGarN!cMsg <> "" Then
            MsgBox rsGarN!cMsg, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
'JOEP20190125 CP

    'If frmAdmCheckListDocument.Inicio(ActxCta.NroCuenta, Right(cmbSubTipo.Text, 3), nRegSugerencia) = True Then 'Comento JOEP20181229 CP
    If frmAdmCheckListDocument.Inicio(ActxCta.NroCuenta, Trim(Right(cmbProducto.Text, 9)), Trim(Right(cmbSubProd.Text, 9)), CCur(txtMonSug.Text), Trim(Right(cmbTipoCredito.Text, 9)), nRegSugerencia) = True Then  'JOEP20181229 CP
        bCheckList = True
    Else
        bCheckList = False
    End If
'JOEP20190125 CP
Set objGarN = Nothing
RSClose rsGarN
'JOEP20190125 CP
End Sub
'RECO FIN ******************************************************
Private Sub CmdCredVig_Click()
    MatCredVig = frmCredVigentes.Inicio(lblcod.Caption, lblnom.Caption, ActxCta.NroCuenta, MatCredVig)
End Sub

Private Sub cmdCtaAho_Click()
Dim clsPers As COMDPersona.UCOMPersona 'UPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
  
    sPers = clsPers.sPersCod
    'By Capi 10042008
    sCliCtaAho = clsPers.sPersNombre
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    
    'By Capi 10042008
    If sPers = gsCodPersUser Then
        MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
        Unload Me
        Exit Sub
    End If
    MsgBox "Solo se Considera Cuentas Activas y Afectas a ITF", vbInformation, "Aviso"
    '
    'Set rsPers = clsCap.GetCuentasPersona(sPers, gCapAhorros, , , , , gsCodAge) 'FRHU20140212 RQ14005
    Set rsPers = clsCap.GetCuentasPersona(sPers, gCapAhorros, , , , , "__") 'FRHU20140212 RQ14005
    Set clsCap = Nothing
    'By Capi 10042008
    Dim lbTieneCuentaValida As Boolean
    lbTieneCuentaValida = False
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            'By Capi 10042008
            If rsPers("Exonerada") = "" And sEstado = "ACTIVA" Then
                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
                lbTieneCuentaValida = True
            End If
            rsPers.MoveNext
        Loop
        'By Capi 15042008
        If lbTieneCuentaValida = True Then
            Set clsCuenta = frmCapMantenimientoCtas.inicia
       
            If clsCuenta.sCtaCod <> "" Then
                actxCtaAho.CMAC = Mid(clsCuenta.sCtaCod, 1, 3) 'FRHU20140212 RQ14005
                actxCtaAho.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                actxCtaAho.Prod = Mid(clsCuenta.sCtaCod, 6, 3) 'FRHU20140212 RQ14005
                actxCtaAho.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                lblTercero.Caption = sCliCtaAho
                actxCtaAho.SetFocusCuenta
                SendKeys "{Enter}"
            End If
            Set clsCuenta = Nothing
        Else
            MsgBox "Persona no posee cuentas validas para depositos.", vbInformation, "Aviso"
        End If
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsPers = Nothing
actxCtaAho.SetFocusCuenta
End Sub

Private Sub CmdEliminar_Click()
    Dim nRel As Integer
    Dim nFila As Integer
    nFila = Me.grdEmpVinculados.row
    
    If MsgBox("¿¿Está seguro de eliminar a la empresa de la relación??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        nRel = CInt(Trim(Right(grdEmpVinculados.TextMatrix(nFila, 3), 4))) 'BRGO 20111115
        grdEmpVinculados.EliminaFila grdEmpVinculados.row
        grdEmpVinculados.TipoBusqueda = BuscaPersona
        '*** BRGO 20111115 **************************
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000010", sSTipoProdCod) Then     '**END ARLO
        'If sSTipoProdCod = "517" Then
            If nRel = gColRelPersOperCertif Then
                sPersOperador = "": sPersOperadorNombre = ""
                Me.txtCtaGarantia.Text = ""
                Me.txtMontoGarantia.Text = "0.00"
                Me.txtCtaGarantia.Enabled = False
                Me.txtMontoGarantia.Enabled = False
            End If
            CalcularDatosEmpVinculados
        End If
        '*** END BRGO *******************************
    End If
End Sub

'Private Sub cmdEvaluacion_Click()
'    Dim oDCred As COMDCredito.DCOMCredito
'    Dim oTipoCam As COMDConstSistema.NCOMTipoCambio
'    Dim rs As ADODB.Recordset
'    Dim nFormato As Integer
'    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
'    Set oDCred = New COMDCredito.DCOMCredito
'    Set rs = oDCred.RecuperaSolicitudDatoBasicos(ActxCta.NroCuenta)
'        If rs.RecordCount > 0 Then
'            nFormato = oDCred.AsignarFormato(Mid(Trim(rs!cTpoProdCod), 1, 1) & "00", Trim(rs!cTpoProdCod), CDbl(rs!nMonto) * IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", 1, CDbl(oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia))))
'            Select Case nFormato
'                Case 0: MsgBox "Crédito no se adecua para este Proceso.", vbInformation, "Aviso"
'                Case 1: Call frmCredEvalFormato1.Inicio(ActxCta.NroCuenta, 2)
'                Case 2: Call frmCredEvalFormato2.Inicio(ActxCta.NroCuenta, 2)
'                Case 3: Call frmCredEvalFormato3.Inicio(ActxCta.NroCuenta, 2)
'            End Select
'        Else
'            MsgBox "Crédito no se adecua para este Proceso.", vbInformation, "Aviso"
'        End If
'End Sub
Private Sub cmdEvaluacion_Click()
    If Len(ActxCta.NroCuenta) <> "18" Then
        Exit Sub
    End If
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim nEstado As Integer
    Dim rs As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    Dim nFormEmpr As Boolean
    Dim nProducto As String
    
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rs = oDCOMFormatosEval.RecuperaCredFormEvalProductoEstadoExposicion(ActxCta.NroCuenta)
    nEstado = IIf(IsNull(rs!nPrdEstado), 0, rs!nPrdEstado)
    Set oRs = oDCOMFormatosEval.RecuperaFormatoEvaluacion(ActxCta.NroCuenta)
    If (oRs.EOF And oRs.BOF) Then
    nProducto = Mid(ActxCta.Prod, 1, 1) & "00"
    If ValidaMultiForm(Trim(Right(cmbSubProd.Text, 8))) Then 'JOEP20190226 CP
        'If ValidaMultiForm(nProducto) Then 'Comento JOEP20190226 CP
            If MsgBox("¿Desea utilizar un formato empresarial?", vbYesNo + vbInformation, "Alerta") = vbYes Then
                nFormEmpr = True
            Else
                nFormEmpr = False
            End If
        End If
    End If
    Call EvaluarCredito(ActxCta.NroCuenta, False, nEstado, CInt(Mid(sSTipoProdCod, 1, 1) & "00"), CInt(sSTipoProdCod), fnMontoExpEsteCred_NEW, False, , nFormEmpr)
    'EvaluarCredito ActxCta.NroCuenta, False, nEstado, CInt(Mid(sSTipoProdCod, 1, 1) & "00"), CInt(sSTipoProdCod), fnMontoExpEsteCred_NEW, False
End Sub

Private Sub cmdExoneracion_Click()
    'frmCredSugExonera.Show 1 'RECO20160526 ERS0022016***************
    'RECO20160628 ERS002-2016*****************************************
    'If frmCredExoneraNCNew.inicia(ActxCta.NroCuenta) Then
    If frmCredExoneraNCNew.inicia(ActxCta.NroCuenta, vnTipoCarga) Then 'FRHU 20160820
        ChkExoneracion.value = 1
    Else
        ChkExoneracion.value = 0
    End If
    'RECO FIN *******************************************************
End Sub

'Private Sub cmdEvaluacion_Click()
'Dim nTipoEval As Integer
'nTipoEval = 0
'If MatFuentesF(3, 1) <> "" Then
'    If MatFuentesF(3, 1) = "D" Then
'        nTipoEval = 1
'    Else
'        nTipoEval = 2
'    End If
'Else
'    MsgBox "Seleccione una fuente de Ingreso.", vbInformation, "Aviso"
'    Exit Sub
'End If
'
'Dim rsHojEval As ADODB.Recordset
'Dim rsHojMaq As ADODB.Recordset
'Dim rsCabHojEval As ADODB.Recordset
'
'Dim oNCredito As COMNCredito.NCOMCredito
'Dim oDCredito As COMDCredito.DCOMCredito
'Dim oDPer As New COMDPersona.DCOMPersonas
'Dim nCapaPa As Double
'
'Set oDCredito = New COMDCredito.DCOMCredito
'Set rsHojEval = oDCredito.ReportesHojaEvaluacionRatios(MatFuentesF(1, 1), MatFuentesF(2, 1), nTipoEval)
'Set rsCabHojEval = oDPer.ObtenerDatosDocsPers(lblcod.Caption)
'nCapaPa = 0
'
'If rsCabHojEval.RecordCount = 1 Then
'    If Not rsCabHojEval.BOF = True And Not rsCabHojEval.EOF = True Then
'        rsCabHojEval.MoveFirst
'    End If
'    Do Until rsCabHojEval.EOF
'       ' oCredito.GeneraMatrixEvaluacion(rsHojEval,rsCabHojEval!cPersona,rsCabHojEval!cPersCod,rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis)
'        'Call ImprimeHojaEvaluacionExcelCab(rsHojEval, rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis)
'        Set oNCredito = New COMNCredito.NCOMCredito
'        nCapaPa = Me.lblcuota.Caption  'txtmonsug.Text / spnCuotas.Valor
'        '*** PEAC 20080618 - SE AGREGO UN PARAMETRO PARA EL ANALISTA
'        previo.Show oNCredito.GeneraMatrixEvaluacion(rsHojEval, rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis, gsCodUser, IIf(nCapaPa, nCapaPa, 0), IIf(txtmonsug.Text, txtmonsug.Text, 0), IIf(Txtinteres.Text, Txtinteres.Text, 0), IIf(spnCuotas.valor, spnCuotas.valor, 0), nTipoEval, 0, lblanalista.Caption), "Hoja de Evaluación", True
'    rsCabHojEval.MoveNext
'    Loop
'End If
'
'cmdEvaluacion.Enabled = False
'End Sub

Private Sub cmdFlujoCaja_Click()
Dim oCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim sVariacionTC As String
Dim MatFechas As Variant
Dim MatFluctuac As Variant
Dim i As Integer

If Mid(ActxCta.NroCuenta, 6, 1) <> "1" And Mid(ActxCta.NroCuenta, 6, 1) <> "2" Then
    MsgBox "El flujo de caja se utiliza para creditos Mes o Comercial", vbInformation, "Mensaje"
    Exit Sub
End If

If UBound(MatrizCal) = 0 Then
    MsgBox "Debe generar el Calendario de Pagos", vbInformation, "Mensaje"
    Exit Sub
End If

ReDim MatFechas(UBound(MatrizCal))

For i = 0 To UBound(MatFechas) - 1
    MatFechas(i) = MatrizCal(i, 0)
Next i

Call frmCredSugerenciaFlujo.Inicio(MatFechas)

If frmCredSugerenciaFlujo.nVarMensualTC = 0 Then
    MsgBox "Debe ingresar la Variación Mensual", vbInformation, "Mensaje"
    Exit Sub
End If
Set oCred = New COMDCredito.DCOMCredito
Set rs = oCred.RecuperaDatosFlujoCaja(ActxCta.NroCuenta)
Set oCred = Nothing

ReDim MatFluctuac(UBound(frmCredSugerenciaFlujo.MatMensualPorc), 2)
For i = 0 To UBound(frmCredSugerenciaFlujo.MatMensualPorc) - 1
    MatFluctuac(i, 0) = Format(MatrizCal(i, 0), "mmm-yy")
    MatFluctuac(i, 1) = frmCredSugerenciaFlujo.MatMensualPorc(i + 1) & "%"
Next i

Call ImprimeFlujoCaja(rs, frmCredSugerenciaFlujo.nVarMensualTC, MatFluctuac, frmCredSugerenciaFlujo.nInflacion)
End Sub

Sub ImprimeFlujoCaja(ByVal pRs As ADODB.Recordset, ByVal pnVariacionTC As Double, _
                        ByVal pMatFluctuac As Variant, ByVal pnInflacion As Double)
    
    Dim fs As Scripting.FileSystemObject
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim nLineaInicio As Integer
    Dim nLineas As Integer
    Dim nLineasTemp As Integer
    Dim nLineaFluctuac As Integer
    
    Dim i As Integer
    Dim nTotal As Double
    
    Dim glsArchivo As String
    Dim nVariacionPorcenMes As Double
    Dim nValorTemp As Double
    
    Dim nMontoPrestamo As Double
    Dim nNumFlujos As Integer
    Dim K As Integer
    
    nMontoPrestamo = CDbl(txtMonSug.Text)
    
    glsArchivo = "FlujoCaja_" & pRs!cCtaCod & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.Name = "Datos"

    xlAplicacion.Range("A1:A1").ColumnWidth = 8
    xlAplicacion.Range("B1:B1").ColumnWidth = 20
    xlAplicacion.Range("C1:C1").ColumnWidth = 15
    xlAplicacion.Range("D1:D1").ColumnWidth = 5
    xlAplicacion.Range("E1:E1").ColumnWidth = 20
    xlAplicacion.Range("F1:F1").ColumnWidth = 15
                
    nLineas = 1
    
    xlHoja1.Cells(nLineas, 1) = "FLUJO DE CAJA DE CREDITOS"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 2) = pRs!cPersNombre
    nLineas = nLineas + 1
    nLineaInicio = nLineas
    xlHoja1.Cells(nLineas, 2) = "Fecha Flujo"
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 3) = Format(gdFecSis, "yyyy-mm-dd")
    xlHoja1.Cells(nLineas, 5) = "Monto Propuesto"
    xlHoja1.Cells(nLineas, 6) = Format(txtMonSug.Text, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Disponible (Activo Circulante)"
    xlHoja1.Cells(nLineas, 3) = Format(CStr(pRs!nPersFIActivoDisp), "#0.00")
    xlHoja1.Cells(nLineas, 5) = "Plazo en Meses"
    xlHoja1.Cells(nLineas, 6) = spnCuotas.valor
    
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Cuentas por Cobrar"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFICtasxCobrar, "#0.00")
    xlHoja1.Cells(nLineas, 5) = "Tipo Moneda"
    xlHoja1.Cells(nLineas, 6) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", "SOLES", "DOLARES")

    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Inventario"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIInventarios, "#0.00")
    xlHoja1.Cells(nLineas, 5) = "Tasa de Interes(%)"
    xlHoja1.Cells(nLineas, 6) = IIf(Txtinteres.Text <> "0.00", Txtinteres.Text, LblInteres.Caption)
    xlHoja1.Cells(nLineas, 7) = "Var. Mens. T.C."

    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Activo Fijo"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIActivosFijos, "#0.00")
    xlHoja1.Cells(nLineas, 5) = "Tipo de Cambio"
    xlHoja1.Cells(nLineas, 6) = Format(gnTipCambio, "#0.00")
    xlHoja1.Cells(nLineas, 7) = pnVariacionTC

    'xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).BorderAround 1, xlMedium
    'xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineas, 6)).Borders(xlEdgeBottom).LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineas, 6)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineas, 6)).Borders.Weight = xlMedium
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.Weight = xlMedium
    
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Activo Circulante"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios, "#0.00")
    xlHoja1.Cells(nLineas, 5) = "Cuota"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 5)).Font.Bold = True
    xlHoja1.Cells(nLineas, 6) = Format(MatrizCal(0, 2), "#0.00")

    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Activo Total"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos, "#0.00")
    
    nLineaInicio = nLineas + 2
    xlHoja1.Cells(nLineas + 2, 5) = "FLUCTUACIONES DE VENTAS"
    xlHoja1.Cells(nLineas + 2, 6) = "(+ o -)"
    xlHoja1.Range(xlHoja1.Cells(nLineas + 2, 5), xlHoja1.Cells(nLineas + 2, 6)).Font.Bold = True
    
    nLineas = nLineas + 2
    
    nLineaFluctuac = nLineas + 1 'Marcamos la Linea
    
    nLineasTemp = nLineaFluctuac
    For i = 0 To UBound(pMatFluctuac) - 1
        xlHoja1.Cells(nLineasTemp, 5) = "'" & pMatFluctuac(i, 0)
        xlHoja1.Cells(nLineasTemp, 6) = Mid(pMatFluctuac(i, 1), 1, Len(pMatFluctuac(i, 1)) - 1)
        xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 5), xlHoja1.Cells(nLineasTemp, 5)).HorizontalAlignment = xlLeft
        nLineasTemp = nLineasTemp + 1
    Next i
            
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineasTemp - 1, 6)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineasTemp - 1, 6)).Borders.Weight = xlMedium
            
            
    nLineaFluctuac = nLineasTemp
    nLineas = nLineas + 1
    nLineaInicio = nLineas
    
    xlHoja1.Cells(nLineas, 2) = "Cuentas por pagar Proveedores"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIProveedores, "#0.00")

    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Cuentas por pagar Saldo Bancos"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFICredOtros, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "CMAC"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFICredCMACT, "#0.00")
    
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.Weight = xlMedium
    
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Patrimonio"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIPatrimonio, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Excedente"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIVentas + pRs!nPersFIRecupCtasXCobrar - pRs!nPersFICostoVentas - pRs!nPersFIEgresosOtros + pRs!nPersIngFam - pRs!nPersEgrFam, "#0.00")
    nLineas = nLineas + 2
    nLineaInicio = nLineas
    xlHoja1.Cells(nLineas, 2) = "Ventas Contado"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIVentas, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Recup.Ventas al Cred."
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIRecupCtasXCobrar, "#0.00")
    
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.Weight = xlMedium
    
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Utilidad Bruta de Ventas"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIVentas + pRs!nPersFIRecupCtasXCobrar - pRs!nPersFICostoVentas, "#0.00")
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    nLineas = nLineas + 1
    
    nLineaInicio = nLineas
    xlHoja1.Cells(nLineas, 2) = "Costo de Ventas"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFICostoVentas, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Otros Gastos"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersFIEgresosOtros, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Otros Ingresos familiares"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersIngFam, "#0.00")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Gastos familiares"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!nPersEgrFam, "#0.00")
    
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 2), xlHoja1.Cells(nLineas, 3)).Borders.Weight = xlMedium
    
    nLineas = nLineasTemp + 7
    xlHoja1.Cells(nLineas, 2) = "Activo Fijo"
    xlHoja1.Cells(nLineas, 5) = "Otros Gastos"
    xlHoja1.Cells(nLineas, 6) = pRs!nPersFIEgresosOtros
    xlHoja1.Cells(nLineas, 3) = "0.00"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas + 9, 3)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas + 9, 3)).Borders.Weight = xlMedium
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas + 9, 6)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas + 9, 6)).Borders.Weight = xlMedium
    
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Personal"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "CTS"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Luz Agua Desague"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Transporte"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Sunat"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Contratados"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Utiles de enseñanza"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 5) = "Otros Promotor Contador"
    
    xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(nLineas, 3)).NumberFormat = "#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(1, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
'    xlHoja1.Cells(nLineas, 2) = pRS!cAgeDescripcion
'    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
'    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
'    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(2, 1)).Font.Size = 9
    xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(nLineas, 8)).Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
    xlHoja1.Cells.EntireRow.AutoFit
    
    'xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(nLineasTemp + 1, 8)).Select
    'xlHoja1.Protect
    
    Set xlHoja1 = Nothing
'********************** FICHA DE FLUJOS ********************************************
    
    If Mid(pRs!cCtaCod, 9, 1) = "1" Then   'soles
        nNumFlujos = 1
    Else                'dolares
        nNumFlujos = 3
    End If
    
    For K = 1 To nNumFlujos
        
        nMontoPrestamo = CDbl(txtMonSug.Text) * (1 + ((K - 1) * 10) / 100)
        
        Set xlHoja1 = xlLibro.Worksheets.Add
    
        xlHoja1.Name = "Flujo " & IIf(Mid(pRs!cCtaCod, 9, 1) = "1", "", (K - 1) * 10 & "%")
    
        xlAplicacion.Range("A1:A1").ColumnWidth = 15
        xlAplicacion.Range("B1:Z1").ColumnWidth = 15
                    
        nLineas = 1
        
        xlHoja1.Cells(nLineas, 1) = "PERIODO"
        xlHoja1.Cells(nLineas, 2) = "'" & Format(gdFecSis, "mmm-yy")
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).NumberFormat = "mmm-yy"
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "PERIODO"
        xlHoja1.Cells(nLineas, 2) = "0"
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "T.C."
        xlHoja1.Cells(nLineas, 2) = gnTipCambio
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "INGRESOS"
        xlHoja1.Cells(nLineas, 2) = nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Ventas Contado"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersFIVentas
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Otros Ingresos"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersIngFam
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Financiamiento"
        xlHoja1.Cells(nLineas, 2) = nMontoPrestamo
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "EGRESOS"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Costo Ventas"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersFICostoVentas
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Costo Operativo"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersFIEgresosOtros
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- CUOTA CMACT"
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Gastos Financ."
        xlHoja1.Cells(nLineas, 2) = ""
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "- Amortizac."
        xlHoja1.Cells(nLineas, 2) = ""
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Otros Egresos"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersEgrFam
            
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "INVERSION"
        xlHoja1.Cells(nLineas, 2) = Format(nMontoPrestamo, "#0.00")
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "SALDO"
        xlHoja1.Cells(nLineas, 2) = nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo)
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "SALDO DISPONIBLE"
        xlHoja1.Cells(nLineas, 2) = pRs!nPersFIActivoDisp
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "SALDO ACUMULADO"
        xlHoja1.Cells(nLineas, 2) = nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo) + pRs!nPersFIActivoDisp
        nValorTemp = nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo) + pRs!nPersFIActivoDisp
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "CLIENTE"
        xlHoja1.Cells(nLineas, 2) = pRs!cPersNombre
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, UBound(pMatFluctuac) + 2)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "INCREMENTO DE VENTAS POR PERIODO"
        xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, UBound(pMatFluctuac) + 2)).Merge True
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Rubro/Periodo"
        xlHoja1.Cells(nLineas, 2) = "0"
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Mes"
        xlHoja1.Cells(nLineas, 2) = Format(gdFecSis, "mmm-yy") '"'" & pMatFluctuac(0, 0)
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Incremento"
        xlHoja1.Cells(nLineas, 2) = "0" 'Mid(pMatFluctuac(0, 1), 1, Len(pMatFluctuac(0, 1)) - 1)
        
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = ""
        
        nLineasTemp = nLineas
        
        'Llenamos las columnas de PERIODOS
        For i = 0 To UBound(pMatFluctuac) - 1
            nVariacionPorcenMes = (1 + CDbl(Mid(CStr(pMatFluctuac(i, 1)), 1, Len(CStr(pMatFluctuac(i, 1))) - 1)) / 100)
            nLineas = 1
            xlHoja1.Cells(nLineas, i + 3) = "'" & pMatFluctuac(i, 0)
            xlHoja1.Range(xlHoja1.Cells(nLineas, i + 3), xlHoja1.Cells(nLineas, i + 3)).Font.Bold = True
            
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = i + 1   'Periodo
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = gnTipCambio + (i + 1) * pnVariacionTC   'Tipo Cambio
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = (pRs!nPersFIVentas + pRs!nPersIngFam) * nVariacionPorcenMes 'INGRESOS
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = (pRs!nPersFIVentas) * nVariacionPorcenMes   'Ventas
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = (pRs!nPersIngFam) * nVariacionPorcenMes 'Otros Ingresos
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = ""  'Financiamiento
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = (pRs!nPersFICostoVentas * nVariacionPorcenMes) + (IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", MatrizCal(i, 2), MatrizCal(i, 2) * gnTipCambio * (1 + pnVariacionTC * i))) + (((pnInflacion / 100 + 1) ^ (i + 1)) * pRs!nPersEgrFam) + pRs!nPersFIEgresosOtros 'EGRESOS
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = pRs!nPersFICostoVentas * nVariacionPorcenMes    'Costo Ventas
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = pRs!nPersFIEgresosOtros   'Costo Operativo
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", MatrizCal(i, 2), MatrizCal(i, 2) * gnTipCambio * (1 + pnVariacionTC * i))   'CUOTA CMAC
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", CDbl(MatrizCal(i, 4)) + CDbl(MatrizCal(i, 5)), (CDbl(MatrizCal(i, 4)) + CDbl(MatrizCal(i, 5))) * gnTipCambio * (1 + pnVariacionTC * i)) 'Gastos Financ
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", MatrizCal(i, 3), MatrizCal(i, 3) * gnTipCambio * (1 + pnVariacionTC * i))   'Amortizac
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = ((pnInflacion / 100 + 1) ^ (i + 1)) * pRs!nPersEgrFam 'Otros Egresos
            nLineas = nLineas + 1
            
            xlHoja1.Cells(nLineas, i + 3) = ""
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = ((pRs!nPersFIVentas + pRs!nPersIngFam) * nVariacionPorcenMes) - ((pRs!nPersFICostoVentas * nVariacionPorcenMes) + (IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", MatrizCal(i, 2), MatrizCal(i, 2) * gnTipCambio * (1 + pnVariacionTC * i))) + (((pnInflacion / 100 + 1) ^ (i + 1)) * pRs!nPersEgrFam) + pRs!nPersFIEgresosOtros)     'SALDO
            nLineas = nLineas + 1
            'xlHoja1.Cells(nLineas, i + 3) = prs!nPersFIActivoDisp * nVariacionPorcenMes 'SALDO DISPONIBLE
            xlHoja1.Cells(nLineas, i + 3) = nValorTemp 'SALDO DISPONIBLE
            nLineas = nLineas + 1
            'xlHoja1.Cells(nLineas, i + 3) = (prs!nPersFIVentas + prs!nPersIngFam - (prs!nPersFICostoVentas + prs!nPersFIEgresosOtros + prs!nPersEgrFam) + prs!nPersFIActivoDisp) * nVariacionPorcenMes 'SALDO ACUMULADO
            xlHoja1.Cells(nLineas, i + 3) = nValorTemp + ((pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam)) * nVariacionPorcenMes) 'SALDO ACUMULADO
            nValorTemp = nValorTemp + ((pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam)) * nVariacionPorcenMes)
            nLineas = nLineas + 3
            xlHoja1.Cells(nLineas, i + 3) = i + 1
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = "'" & pMatFluctuac(i, 0)
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, i + 3) = Mid(pMatFluctuac(i, 1), 1, Len(pMatFluctuac(i, 1)) - 1)
        Next i
        
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(nLineas, i + 2)).Borders.LineStyle = 1
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(nLineas, i + 2)).Borders.Weight = xlMedium
        
        nLineas = nLineasTemp
        
        nLineas = nLineas + 2
        nLineaInicio = nLineas
        xlHoja1.Cells(nLineas, 1) = "SUPUESTOS"
        'nLineas = nLineas + 2
        xlHoja1.Cells(nLineas, 5) = "INDICADORES FINANCIEROS"
        xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 5)).Font.Bold = True
            
        xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 6)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Inflación"
        
        xlHoja1.Cells(nLineas, 2) = pnInflacion
        xlHoja1.Cells(nLineas, 5) = "Liquidez"
        
        If pRs!nPersFIProveedores + pRs!nPersFICredOtros + pRs!nPersFICredCMACT <= 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios) / (pRs!nPersFIProveedores + pRs!nPersFICredOtros + pRs!nPersFICredCMACT)
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp < 1 Then
            xlHoja1.Cells(nLineas, 7) = "Los bienes liquidos no cubren los recursos o deudas a corto plazo"
        ElseIf nValorTemp = 1 Then
            xlHoja1.Cells(nLineas, 7) = "Los bienes liquidos son iguales a Los recursos ajenos a corto plazo"
        ElseIf nValorTemp > 1 Then
            xlHoja1.Cells(nLineas, 7) = "Liquidez suficiente para cubrir recursos ajenos a CP"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Costo de Ventas"
        xlHoja1.Cells(nLineas, 2) = CStr(pRs!nPersFICostoVentas / pRs!nPersFIVentas * 100) & "%"
        
        xlHoja1.Cells(nLineas, 5) = "Prueba acida"
        
        If pRs!nPersFIProveedores + pRs!nPersFICredOtros + pRs!nPersFICredCMACT <= 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar) / (pRs!nPersFIProveedores + pRs!nPersFICredOtros + pRs!nPersFICredCMACT)
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp < 1 Then
            xlHoja1.Cells(nLineas, 7) = "El efectivo cubre obligaciones a largo plazo"
        ElseIf nValorTemp = 1 Then
            xlHoja1.Cells(nLineas, 7) = "Los bienes liquidos son iguales a Los recursos ajenos a corto plazo"
        ElseIf nValorTemp > 1 Then
            xlHoja1.Cells(nLineas, 7) = "Liquidez suficiente para cubrir recursos aun a muy corto plazo"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Costo Operativo"
        xlHoja1.Cells(nLineas, 2) = CStr(Format(pRs!nPersFICostoVentas / pRs!nPersFIEgresosOtros * 100, "#0.00")) & "%"
        
        xlHoja1.Cells(nLineas, 5) = "Endeudamiento"
        
        If pRs!nPersFIPatrimonio = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = (pRs!nPersFIProveedores + pRs!nPersFICredOtros + pRs!nPersFICredCMACT) / pRs!nPersFIPatrimonio
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp >= 80 Then
            xlHoja1.Cells(nLineas, 7) = "Alerta, el negocio no posee autonomia"
        ElseIf nValorTemp = 0 Then
            xlHoja1.Cells(nLineas, 7) = "El negocio no tiene endeudamiento"
        ElseIf nValorTemp < 80 Then
            xlHoja1.Cells(nLineas, 7) = "Negocio presenta deudas con terceros"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Prestamo (1)"
        xlHoja1.Cells(nLineas, 2) = Format(nMontoPrestamo, "#0.00")
        
        xlHoja1.Cells(nLineas, 5) = "Periodo de C"
        
        If pRs!nPersFIVentas = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFICtasxCobrar / (pRs!nPersFIVentas / 30)
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp > 25 Then
            xlHoja1.Cells(nLineas, 7) = "Problemas de cobranzas"
        ElseIf nValorTemp = 0 Then
            xlHoja1.Cells(nLineas, 7) = "Negocio vende al contado"
        ElseIf nValorTemp < 25 Then
            xlHoja1.Cells(nLineas, 7) = "Periodo de cobranza aceptable"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Nº de Cuotas"
        xlHoja1.Cells(nLineas, 2) = spnCuotas.valor
        
        xlHoja1.Cells(nLineas, 5) = "Ventas/Invent."
        
        If pRs!nPersFIInventarios = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFIVentas / pRs!nPersFIInventarios
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp <= 0.5 Then
            xlHoja1.Cells(nLineas, 7) = "La mercaderia que existe puede ser obsoleta"
        Else
            xlHoja1.Cells(nLineas, 7) = "La mercaderia del negocio es actual"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Interes"
        xlHoja1.Cells(nLineas, 2) = IIf(Txtinteres.Text = "", Txtinteres.Text, LblInteres.Caption)
        
        xlHoja1.Cells(nLineas, 5) = "Inv./Capital W"
        
        If pRs!nPersFIInventarios = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFIInventarios / (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios - pRs!nPersFIProveedores - pRs!nPersFICredOtros - pRs!nPersFICredCMACT)
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp >= 1.5 Then
            xlHoja1.Cells(nLineas, 7) = "Excesiva inversion en mercaderia"
        Else
            xlHoja1.Cells(nLineas, 7) = "El nivel de inventarios es adecuado"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        xlHoja1.Cells(nLineas, 1) = "Cuota"
        xlHoja1.Cells(nLineas, 2) = Format(MatrizCal(0, 2), "#0.00")
        
        xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 1), xlHoja1.Cells(nLineas, 2)).Borders.LineStyle = 1
        xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 1), xlHoja1.Cells(nLineas, 2)).Borders.Weight = xlMedium
            
        xlHoja1.Cells(nLineas, 5) = "Rotacion"
        
        If pRs!nPersFIInventarios = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFICostoVentas / pRs!nPersFIInventarios
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp <= 0.2 Then
            xlHoja1.Cells(nLineas, 7) = "La mercaderia que existe es de poca rotacion"
        Else
            xlHoja1.Cells(nLineas, 7) = "La mercaderia del negocio se renueva con normalidad"
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
    
        'nLineasTemp = nLineas
        
        xlHoja1.Cells(nLineas, 5) = "Rentabilidad"
        
        If pRs!nPersFIPatrimonio = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = ((nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo)) + pRs!nPersEgrFam) / pRs!nPersFIPatrimonio
            xlHoja1.Cells(nLineas, 6) = CStr((nValorTemp * 100)) & "%"
        End If
        'xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If pRs!nPersFIInventarios = 0 Then
            If nValorTemp <= 0 Then
                xlHoja1.Cells(nLineas, 7) = "La rentabilidad del capital es demasiada baja"
            Else
                xlHoja1.Cells(nLineas, 7) = "La rentabilidad del capital es mayor al costo financiero"
            End If
        Else
            If nValorTemp <= (pRs!nPersFICostoVentas / pRs!nPersFIInventarios) Then
                xlHoja1.Cells(nLineas, 7) = "La rentabilidad del capital es demasiada baja"
            Else
                xlHoja1.Cells(nLineas, 7) = "La rentabilidad del capital es mayor al costo financiero"
            End If
        End If
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        
        xlHoja1.Cells(nLineas, 5) = "Utilidas/Ventas"
        
        If pRs!nPersFIVentas = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = ((nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo)) + pRs!nPersEgrFam) / pRs!nPersFIVentas
            xlHoja1.Cells(nLineas, 6) = CStr((nValorTemp * 100)) & "%"
        End If
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        
        xlHoja1.Cells(nLineas, 5) = "Utilidas/Activo"
        
        If pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = ((nMontoPrestamo + pRs!nPersFIVentas + pRs!nPersIngFam - (pRs!nPersFICostoVentas + pRs!nPersFIEgresosOtros + pRs!nPersEgrFam + nMontoPrestamo)) + pRs!nPersEgrFam) / (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos)
            xlHoja1.Cells(nLineas, 6) = CStr((nValorTemp * 100)) & "%"
        End If
                    
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
        
        xlHoja1.Cells(nLineas, 5) = "Ventas/Activo"
        
        If pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFIVentas / (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos)
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        If nValorTemp >= 2 Then
            xlHoja1.Cells(nLineas, 7) = "La contribucion del activo es significativa"
        ElseIf nValorTemp <= 0.1 Then
            xlHoja1.Cells(nLineas, 7) = "Existe exceso de inversion en activos o mala utilizacion de los mismos"
        ElseIf nValorTemp > 0.1 Then
            xlHoja1.Cells(nLineas, 7) = "La contribucion del activo es aceptable"
        End If
            
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
        
        nLineas = nLineas + 1
    
        xlHoja1.Cells(nLineas, 5) = "Rotación de K"
        
        If pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFIVentas / (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios + pRs!nPersFIActivosFijos)
            xlHoja1.Cells(nLineas, 6) = CStr((nValorTemp * 100)) & "%"
        End If
                    
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
            
        nLineas = nLineas + 1
    
        xlHoja1.Cells(nLineas, 5) = "Tesoreria"
        
        If pRs!nPersFIInventarios = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios - pRs!nPersFIProveedores - pRs!nPersFICredOtros - pRs!nPersFICredCMACT) / pRs!nPersFIInventarios
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
            
        nLineas = nLineas + 1
    
        xlHoja1.Cells(nLineas, 5) = "Rotacion de Kw"
        
        If (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios - pRs!nPersFIProveedores - pRs!nPersFICredOtros - pRs!nPersFICredCMACT) = 0 Then
            xlHoja1.Cells(nLineas, 6) = "0"
            nValorTemp = 0
        Else
            nValorTemp = pRs!nPersFIVentas / (pRs!nPersFIActivoDisp + pRs!nPersFICtasxCobrar + pRs!nPersFIInventarios - pRs!nPersFIProveedores - pRs!nPersFICredOtros - pRs!nPersFICredCMACT)
            xlHoja1.Cells(nLineas, 6) = nValorTemp
        End If
        xlHoja1.Range(xlHoja1.Cells(nLineas, 6), xlHoja1.Cells(nLineas, 6)).NumberFormat = "#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(nLineas, 7), xlHoja1.Cells(nLineas, 10)).Merge True
            
        xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineas, 6)).Borders.LineStyle = 1
        xlHoja1.Range(xlHoja1.Cells(nLineaInicio, 5), xlHoja1.Cells(nLineas, 6)).Borders.Weight = xlMedium
            
        nLineas = nLineas + 1
    
        xlHoja1.Range(xlHoja1.Cells(1, 2), xlHoja1.Cells(nLineasTemp - 1, 30)).NumberFormat = "#,##0.00"
        xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 30)).NumberFormat = "0"
        xlHoja1.Range(xlHoja1.Cells(21, 1), xlHoja1.Cells(21, 30)).NumberFormat = "0"
        'xlHoja1.Range(xlHoja1.Cells(1, 30), xlHoja1.Cells(nLineas, 30)).NumberFormat = "#,##0.00"
        
        'xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(2, 1)).Font.Size = 9
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(nLineas, 30)).Font.Size = 8
        xlHoja1.Cells.EntireColumn.AutoFit
        xlHoja1.Cells.EntireRow.AutoFit
    
        xlHoja1.Cells.Select
        xlHoja1.Protect
    
    Next K
'***********************************************************************************
    
    xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsArchivo
               
    MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsArchivo, vbInformation, "Mensaje"
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
        
    Set xlAplicacion = Nothing

End Sub

Private Sub cmdFuentes_Click()
'    If Not ExisteTitular Then
'        MsgBox "Debe Ingresar el Titular del Credito", vbInformation, "Aviso"
'        cmdRelaciones.SetFocus
'        Exit Sub
'    End If
  
    'Call frmPersona.Inicio(TitularCredito, PersonaActualiza)
    Call frmPersona.Inicio(lblcod, PersonaActualiza)
    'Call CargaFuentesIngreso(TitularCredito)
    Call CargaFuentesIngreso(lblcod)
    'oPersona.PersCodigo = TitularCredito
    oPersona.PersCodigo = lblcod
End Sub

'**** PEAC 20080412
Private Sub CargaFuentesIngreso(ByVal psPersCod As String)
Dim i As Integer
Dim MatFte As Variant

    On Error GoTo ErrorCargaFuentesIngreso
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli
    Call oPersona.RecuperaPersona_Solicitud(psPersCod, gdFecSis)
    
    'oPersona.PersCodigo = TitularCredito
    oPersona.PersCodigo = lblcod
    Exit Sub

ErrorCargaFuentesIngreso:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdGarantia_Click()
    Call frmCredGarantiasCob.Inicio(ActxCta.NroCuenta)
End Sub

Private Sub CmdGastos_Click()

Dim oNGasto As COMNCredito.NCOMGasto
Dim bErrorCargaCalendario As Boolean 'ARCV

'MAVM 20120802 ***
Dim nTramoNoConsMonto As Double
Dim nTramoConsMonto As Double
Dim nTramoNoConsPorcen As Double
Dim nPlazoMiViv As Integer
Dim nPlazoMiVivMax As Integer
'***
Dim nTasa As Double
Dim bQuincenal As Boolean

If Txtinteres.Visible Then
    nTasa = CDbl(Txtinteres.Text)
Else
    nTasa = CDbl(LblInteres.Caption)
End If
Dim vArrDatos As Variant 'LUCV20180601, Según ERS022-2018
'Dim nMontoGarantia As Double
'Dim oGar As COMDCredito.DCOMGarantia

    On Error GoTo ErrorCmdGastos_Click
    If Not bCalendGen Then
        MsgBox "Genere el Calendario de Pagos para poder Definir los Gastos", vbInformation, "Aviso"
        CmdCalend.SetFocus
        Exit Sub
    End If
    
    'MAVM 20120802 ***
    If Me.ChkMiViv.value = 1 Then
        Dim oParam As COMDCredito.DCOMParametro
        Set oParam = New COMDCredito.DCOMParametro
        Call oParam.RecuperaParametrosCalendario(nTramoNoConsPorcen, nPlazoMiViv, nPlazoMiVivMax)
        nTramoNoConsMonto = Format(CDbl(txtMonSug.Text) - nTramoNoConsPorcen, "#0.00")
        nTramoConsMonto = Format(nTramoNoConsPorcen, "#0.00")
    End If
    '***
    
    Dim oCalend As COMDCredito.DCOMCalendario
    Set oCalend = New COMDCredito.DCOMCalendario
    bQuincenal = oCalend.VerificaCalendQuincenal(ActxCta.NroCuenta)
    Set oCalend = Nothing
    
    '***************************************************
    'FALTA LA FUNCION QUE ME PERMITA CAPTURAR EL TIPO DE GARANTIA SEGUN EL INMUEBLE (INDIVIDUAL O MANCOMUNADA)
    '***************************************************
'    Set oNGasto = New COMNCredito.NCOMGasto
'    Set oGar = New COMDCredito.DCOMGarantia
'    nMontoGarantia = oGar.RecuperaMontoGarantiaCredito(ActxCta.NroCuenta)
'    Set oGar = Nothing
 Set oNGasto = New COMNCredito.NCOMGasto
    
    'MatGastos = oNGasto.GeneraCalendarioGastos(MatCalend, MatDesemb, nNumgastos, gdFecSis, _
                ActxCta.NroCuenta, 1, "DE", IIf(OptTipoGasto(0).value, "F", "V"), _
                CDbl(MatCalend(0, 2)), CDbl(TxtMontoApr.Text), , , , , , , , True)   'Solo Mostrar
        'Modify Gitu 06-04-2009 se cambio los parametros de tasas de interes para que envie los valores de los Text's respectivos
        
        'MAVM 20121113 ***
        'MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMontoApr.Text), txtTasInt.Text, CInt(SpnCuotasApr.valor), _
        '                    IIf(CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)) > 0, 30, CInt(SpnPlazo.valor)), CDate(MatDesemb(0, 0)), DameTipoCuota, _
        '                    IIf(OptTPer(0).value, 1, 2), vnTipoGracia, CInt(TxtDiasGracia.Text), _
        '                    TxtTasaGracia.Text, CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
        '                    ChkProxMes.value, MatGracia, ChkMiViv.value, ChkCuotaCom.value, MatCalend_2, _
        '                    MatDesemb, nNumGastos, gdFecSis, _
        '                    ActxCta.NroCuenta, 1, "DE", IIf(OptTipoGasto(0).value, "F", "V"), _
        '                    CDbl(MatCalend(0, 2)), CDbl(TxtMontoApr.Text), , , , , , , , True, _
        '                    2, True, Me.OptDesemb(1).value, MatDesemb, bQuincenal, bErrorCargaCalendario, _
        '                    CInt(TxtDiaFijo2.Text), IIf(chkIncremenK.value = 1, True, False), , _
        '                     gnITFMontoMin, gnITFPorcent, gbITFAplica, ChkExoneracion.value, Mid(ActxCta.NroCuenta, 4, 2), lsTpoProdCod, lsTpoCredCod) 'Solo Mostrar
                             
        If UBound(MatDesemb) = 0 Then
            ReDim MatDesemb(0, 1)
            MatDesemb(0, 0) = TxtFecDesemb.Text
            MatDesemb(0, 1) = Format(txtMonSug.Text, "#0.00")
        End If
                                     
        '->***** LUCV20180601, Agregó Según ERS022-2018
        ReDim vArrDatos(2)
        vArrDatos(0) = CCur(txtMontoMivivienda.Text)
        vArrDatos(1) = MatCalend 'Para enviar el calendario en la generacion de gastos
        '<-***** Fin LUCV20180601
                        
        MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(txtMonSug.Text), nTasa, CInt(spnCuotas.valor), _
                            IIf(CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)) > 0, 30, CInt(SpnPlazo.valor)), CDate(TxtFecDesemb.Text), DameTipoCuota, _
                            IIf(Trim(Right(cmbTipoPer.Text, 2)) = "0", 1, 2), vnTipoGracia, CInt(txtPerGra.Text), _
                            IIf(TxtTasaGracia.Visible, CDbl(TxtTasaGracia.Text), CDbl(LblTasaGracia.Caption)), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                            ChkProxMes.value, MatGracia, ChkMiViv.value, ChkCuotaCom.value, MatCalend_2, _
                            MatDesemb, nNumGastos, gdFecSis, _
                            ActxCta.NroCuenta, 1, "DE", IIf(OptTipoGasto(0).value, "F", "V"), _
                            CDbl(MatCalend(0, 2)), IIf(ChkMiViv.value, nTramoNoConsMonto, CDbl(txtMonSug.Text)), , , , , , , , True, _
                            2, True, Trim(Right(cmbDesemb.Text, 2)), MatDesemb, bQuincenal, bErrorCargaCalendario, _
                            CInt(TxtDiaFijo2.Text), IIf(chkIncremenK.value = 1, True, False), , _
                            lnCSP, gnITFPorcent, gbITFAplica, ChkExoSegDes.value, Mid(ActxCta.NroCuenta, 4, 2), sSTipoProdCod, Trim(Right(cmbTipoCredito.Text, 3)), , , , , , , , , , vArrDatos) 'Solo Mostrar
                            'LUCV20180601, Agregó vArrDatos =>CCur(txtMontoMivivienda.Text). [Por limite de parámetros]
                            'APRI20171122 ERS028-2017 MEJORA -> CAMBIO ChkExoneracion.value A ChkExoSegDes.value
        '***
    'lnCSP--gnITFMontoMin
    bGastosGen = True
    Set oNGasto = Nothing
    
    Call frmCredGastos.MuestraGastos(MatGastos, nNumGastos, True)
    Exit Sub

ErrorCmdGastos_Click:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdLineas_Click()
Dim oLineas As COMDCredito.DCOMLineaCredito
Dim sCtaCod As String
If Trim(cmbSubTipo.Text) <> "" Then
    bBuscarLineas = True
    sCtaCod = ActxCta.NroCuenta
    Set oLineas = New COMDCredito.DCOMLineaCredito
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
    'txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol(Mid(sCtaCod, 6, 3), Mid(sCtaCod, 9, 1), , Mid(sCtaCod, 4, 2), CInt(spnPlazo.Valor), CDbl(txtmonsug.Text), CInt(spnCuotas.Valor))
    'txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol(Trim(Right(cmbSubTipo.Text, 6)), Mid(sCtaCod, 9, 1), , Mid(sCtaCod, 4, 2), CInt(SpnPlazo.valor), CDbl(txtMonSug.Text), CInt(spnCuotas.valor))
    txtBuscarLinea.rs = oLineas.RecuperaLineasProductoArbol(Trim(Right(cmbSubTipo.Text, 6)), Mid(sCtaCod, 9, 1), , Mid(sCtaCod, 4, 2), CInt(SpnPlazo.valor), CDbl(txtMonSug.Text), CInt(spnCuotas.valor), IIf(Trim(Txtinteres.Text) = "", 0, Txtinteres.Text), IIf(Trim(txtPerGra.Text) = "", 0, txtPerGra.Text), gdFecSis)
    Set oLineas = Nothing
Else
    MsgBox "Seleccionar Sub tipo de crédito", vbCritical
End If
End Sub

Private Sub cmdPersona_Click()
'    If Not ExisteTitular Then
'        MsgBox "Debe Ingresar el Titular del Credito", vbInformation, "Aviso"
'        cmdRelaciones.SetFocus
'        Exit Sub
'    End If
  
    'Call frmPersona.Inicio(TitularCredito, PersonaActualiza)
    Call frmPersona.Inicio(lblcod, PersonaActualiza)
    'Call CargaFuentesIngreso(TitularCredito)
    Call CargaFuentesIngreso(lblcod)
    'oPersona.PersCodigo = TitularCredito
    oPersona.PersCodigo = lblcod
End Sub

Private Sub cmdSeleccionaFuente_Click()
    'Call frmCredSolicitud_SelecFtes.Inicio(oPersona.PersCodigo)
    Call frmCredSolicitud_SelecFtes.Inicio(lblcod)
    MatFuentes = frmCredSolicitud_SelecFtes.MatFuentes
    'ALPA***18*04*2008
   
    MatFuentesF = frmCredSolicitud_SelecFtes.MatFuentesF
    '******************
    'RECO20140226 ERS174-2013***************************
    Dim oUPersona_Cli As UPersona_Cli
    Set oUPersona_Cli = New UPersona_Cli
    Dim i As Integer
    'If MatFuentes(0) Is Nothing Then
    On Local Error Resume Next
    
    Dim sValida As String
    sValida = MatFuentesF(3, 1)
    If Err <> 0 Then
    Else
        If MatFuentesF(3, 1) <> "" Then
            For i = 0 To UBound(MatFuentes) - 1
                If oUPersona_Cli.ValidaExisteFICred(oPersona.ObtenerFteIngcNumFuente(MatFuentes(i)), Me.ActxCta.NroCuenta) Then
                    MsgBox "La fuente de ingreso ya se encuentra vinculada a otro crédito, debe registrar una nueva fuente de ingreso para el crédito.", vbCritical, "Aviso"
                    ReDim MatFuentes(0)
                    'ReDim MatFuentesF(0)
                    Exit Sub
                End If
            Next
        End If
    End If
    'RECO FIN*******************************************
End Sub

'Private Sub Command1_Click()
'Dim nTipoEval As Integer
'nTipoEval = 0
'If MatFuentesF(3, 1) <> "" Then
'    If MatFuentesF(3, 1) = "D" Then
'        nTipoEval = 1
'    Else
'        nTipoEval = 2
'    End If
'Else
'    MsgBox "Seleccione una fuente de Ingreso.", vbInformation, "Aviso"
'    Exit Sub
'End If
'
'Dim rsHojEval As ADODB.Recordset
'Dim rsHojMaq As ADODB.Recordset
'Dim rsCabHojEval As ADODB.Recordset
'
'Dim oNCredito As COMNCredito.NCOMCredito
'Dim oDCredito As COMDCredito.DCOMCredito
'Dim oDPer As New comdpersona.DCOMPersonas
'Dim nCapaPa As Double
'
'Set oDCredito = New COMDCredito.DCOMCredito
'Set rsHojEval = oDCredito.ReportesHojaEvaluacionRatios(MatFuentesF(1, 1), MatFuentesF(2, 1), nTipoEval)
'Set rsCabHojEval = oDPer.ObtenerDatosDocsPers(lblcod.Caption)
'nCapaPa = 0
'
'If rsCabHojEval.RecordCount = 1 Then
'    If Not rsCabHojEval.BOF = True And Not rsCabHojEval.EOF = True Then
'        rsCabHojEval.MoveFirst
'    End If
'Do Until rsCabHojEval.EOF
'   ' oCredito.GeneraMatrixEvaluacion(rsHojEval,rsCabHojEval!cPersona,rsCabHojEval!cPersCod,rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis)
'    'Call ImprimeHojaEvaluacionExcelCab(rsHojEval, rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis)
'    Set oNCredito = New COMNCredito.NCOMCredito
'    nCapaPa = Me.lblcuota.Caption  'txtmonsug.Text / spnCuotas.Valor
'    previo.Show oNCredito.GeneraMatrixEvaluacion(rsHojEval, rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis, gsCodUser, IIf(nCapaPa, nCapaPa, 0), IIf(txtMonSug.Text, txtMonSug.Text, 0), IIf(Txtinteres.Text, Txtinteres.Text, 0), IIf(spnCuotas.valor, spnCuotas.valor, 0), nTipoEval, 0), "Hoja de Evaluación", True
'rsCabHojEval.MoveNext
'Loop
'End If
'End Sub

Private Sub cmdVentasAnual_Click()
    Call frmPersona.Inicio(lblcod, PersonaActualiza)
End Sub

Private Sub cmdVerEntidades_Click()
    Dim oCredPersRela As UCredRelac_Cli
    Set oCredPersRela = New UCredRelac_Cli
    Call oCredPersRela.CargaRelacPersCred(ActxCta.NroCuenta)
    oCredPersRela.IniciarMatriz
    Dim sDocumento As String
    sDocumento = oCredPersRela.ObtenerDocumento
    If Len(Trim(ActxCta.NroCuenta)) = 18 Then
        Set oRsVerEntidades = New ADODB.Recordset
        Set oRsVerEntidades = frmCredVerEntidades.VerEntidades(ActxCta.NroCuenta, lblcod.Caption, sDocumento)
        nLogicoVerEntidades = 1
    End If
End Sub

'FRHU20140212 RQ14005
Private Sub cmdVerTercero_Click()
If Len(actxCtaAho.NroCuenta) = 18 Then
    frmCuentaAhorrosDesembolsoTerceros.Inicio (actxCtaAho.NroCuenta)
Else
    MsgBox "Llena todos los campos Correctamente", vbInformation, "Advertencia"
    actxCtaAho.SetFocusCuenta
End If
End Sub
'FIN FRHU20140212 RQ14005
'ALPA 20090929******************************
Private Sub cmdVinculados_Click()
    frmGruposEconomicos.Show 1
End Sub
'*******************************************

Private Sub Form_Unload(Cancel As Integer)
    'EJVG20151104 *** Tendrá que salir por el botón salir->Sino podría modificar el gravamen y/o solicitud de créditos
    If vbInicioCargaDatos Then
        If Not fbSalirCargaDatos Then
            Cancel = 1
            Exit Sub
        End If
    End If
    'END EJVG *******
    Call Unload(frmCredVigentes)
    MatCredVig = ""
    
    vbInicioCargaDatos = False 'EJVG20151104
    fbSalirCargaDatos = False 'EJVG20151104
End Sub

'****BRGO 20111103 - ADECUACIONES PARA EL PRODUCTO ECOTAXI ************
Private Sub grdEmpVinculados_Click()
    Dim nCol As Integer
    If nCol = 1 Then
        grdEmpVinculados.TipoBusqueda = BuscaPersona
    Else
        grdEmpVinculados.TipoBusqueda = BuscaArbol
    End If
End Sub

Private Sub grdEmpVinculados_DblClick()
    Dim nCol, nFila As Integer
    Dim sCodigo As String
    Dim rsCta As New ADODB.Recordset
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
 
    nCol = Me.grdEmpVinculados.Col
    
    If nCol = 1 Then
        grdEmpVinculados.TipoBusqueda = BuscaPersona
    ElseIf nCol = 5 Then
        grdEmpVinculados.TipoBusqueda = BuscaArbol
        nFila = grdEmpVinculados.row
        sCodigo = grdEmpVinculados.TextMatrix(nFila, 1)

        Set rsCta = clsMant.GetCuentasPersona(sCodigo, gCapAhorros, True, , 1)
        Set clsMant = Nothing
        grdEmpVinculados.rsTextBuscar = rsCta
        If rsCta.EOF And rsCta.BOF Then
            grdEmpVinculados.TextMatrix(nFila, 5) = ""
            MsgBox "Persona no posee cuentas de ahorros disponibles. Debe aperturar.", vbInformation, "Aviso"
        End If
        Set rsCta = Nothing
    End If
End Sub

Private Sub grdEmpVinculados_EnterCell()
    Dim nCol As Integer
    nCol = Me.grdEmpVinculados.Col
    If nCol = 1 Then
        grdEmpVinculados.TipoBusqueda = BuscaPersona
    Else
        grdEmpVinculados.TipoBusqueda = BuscaArbol
    End If
End Sub
Private Sub grdEmpVinculados_OnCellChange(pnRow As Long, pnCol As Long)
    CalcularDatosEmpVinculados
End Sub

Private Sub grdEmpVinculados_OnChangeCombo()
    Dim nRel, nRelacFila, nFila, i As Integer
    Dim nMonto As Double
    Dim OpCertificador, OpGarantia As String
    Dim auxRelacion As String
    
    Dim rs As ADODB.Recordset
    Dim oTipoCambio As nTipoCambio
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    
    nFila = grdEmpVinculados.row
    nRelacFila = CInt(Trim(Right(grdEmpVinculados.TextMatrix(nFila, 3), 4)))
    sPersOperador = "": sPersOperadorNombre = ""
    For i = 1 To nFila
        nRel = CInt(Trim(Right(grdEmpVinculados.TextMatrix(i, 3), 4)))
        If nFila - 1 > 0 And i <> nFila Then
            If nRel = CInt(Trim(Right(grdEmpVinculados.TextMatrix(nFila, 3), 4))) Then
                MsgBox "El tipo de relación ya fue ingresado, seleccione otra", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        If nRel = gColRelPersOperCertif Then
            sPersOperador = grdEmpVinculados.TextMatrix(i, 1)
            sPersOperadorNombre = grdEmpVinculados.TextMatrix(i, 2)
        End If
    Next
    If sPersOperador <> "" Then
        Me.txtCtaGarantia.Enabled = True
    Else
        Me.txtCtaGarantia.Text = ""
    End If
    Set rs = oCred.RecuperaParametro("314" & Right(CStr(nRelacFila), 1))
    If Not (rs.EOF And rs.BOF) Then
        nMonto = rs!nParamValor * IIf(nRel = 11 Or nRel = 12, nTC, 1)
    Else
        MsgBox "No tiene un monto predefinido"
        nMonto = 0
    End If
    If Me.grdEmpVinculados.TextMatrix(nFila, 4) = 0 Then
        Me.grdEmpVinculados.TextMatrix(nFila, 4) = Format(nMonto, "#,##0.00")
    End If
    CalcularDatosEmpVinculados
    Set rs = Nothing
    Set oCred = Nothing
End Sub

Private Sub grdEmpVinculados_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim nRel, nFila, nCol As Integer
    Dim nMonto As Double
    Dim sMsg As String
    Dim rs As ADODB.Recordset
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    
    nFila = Me.grdEmpVinculados.row
    nCol = Me.grdEmpVinculados.Col
    If nCol = 1 Then
        'sMsg = oCred.ValidaPersonaCOFIDE(grdEmpVinculados.TextMatrix(nFila, 1))
        If sMsg <> "" Then
            MsgBox sMsg
            grdEmpVinculados.EliminaFila (nFila)
            grdEmpVinculados.TipoBusqueda = BuscaPersona
        End If
    End If
    Set oCred = Nothing
End Sub
Private Sub CalcularDatosEmpVinculados()
    Dim pnCol As Integer
    Dim i As Integer
    pnCol = Me.grdEmpVinculados.Col
    If pnCol = 4 Or pnCol = 3 Then
        nComisionEC = 0
        For i = 1 To Me.grdEmpVinculados.rows - 1
            Me.grdEmpVinculados.TextMatrix(i, 4) = Format(Me.grdEmpVinculados.TextMatrix(i, 4), "#,000.00")
            nComisionEC = nComisionEC + IIf(Me.grdEmpVinculados.TextMatrix(i, 4) = "", 0, Me.grdEmpVinculados.TextMatrix(i, 4))
        Next
        Me.lblComisionEC.Caption = Format((nComisionEC + CDec(Me.txtTasacion.Text)) * nPorcCEC, "0.00")
    End If
End Sub




Private Sub txtCtaGarantia_EmiteDatos()
Dim sCodigo As String
Dim nmoneda As Moneda
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000011", sSTipoProdCod) Then     '**END ARLO
    'If sSTipoProdCod = "517" Then
        sCodigo = sPersOperador
        nmoneda = CLng(Mid(ActxCta.NroCuenta, 9, 1))
        If sCodigo <> "" Then
            Dim rsCta As New ADODB.Recordset
            Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                 Set rsCta = clsMant.GetCuentasPersona(sCodigo, gCapAhorros, True, , nmoneda)
            Set clsMant = Nothing
            txtCtaGarantia.rs = rsCta
            If rsCta.EOF And rsCta.BOF Then
                txtCtaGarantia.Text = ""
                MsgBox "Cliente No Posee cuentas de ahorros disponibles", vbInformation, "Aviso"
            End If
            Set rsCta = Nothing
            txtCtaGarantia.Visible = True
        Else
            MsgBox "Debe ingresar el Operador (Oper.Certif)", vbInformation, "Aviso"
        End If
    End If
End Sub
'WIOR 20131129 **************************************
Private Sub txtCuotaBalon_Change()
ValidaCuotaBalon
End Sub

Private Sub txtCuotaBalon_GotFocus()
fEnfoque txtCuotaBalon
End Sub

Private Sub txtCuotaBalon_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   CmdCalend.SetFocus
End If
End Sub

'WIOR FIN *******************************************

Private Sub TxtFecDesemb_Change()
    bCalendGen = False
    bGastosGen = False
End Sub

Private Sub TxtFecDesemb_GotFocus()
    fEnfoque TxtFecDesemb
End Sub

Private Sub TxtFecDesemb_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        'MAVM 15112010 ***
        'TxtDiasGracia.SetFocus
        GenerarFechaPago
        'txtFechaFija.SetFocus
        'If opttper(0).value = True Then
                If Trim(Right(cmbTipoPer.Text, 2)) = "0" Then
            SpnPlazo.Enabled = True
            ChkProxMes.value = 0
            TxtDiaFijo2.Text = "00"
            TxtDiaFijo.Text = "00"
            TxtDiaFijo.Enabled = False
            ChkProxMes.Enabled = False
            
            txtPerGra.Enabled = False
            txtPerGra.Text = "0"
            TxtTasaGracia.Text = "0.00"
            'optTipoGracia(0).value = False 'LUCV20180601, Comentó segn ERS022-2018
            'optTipoGracia(1).value = False 'LUCV20180601, Comentó según ERS022-2018
            'cmdgracia.Enabled = True 'LUCV20180601, Comentó según ERS022-2018
        Else
            SpnPlazo.Enabled = False
            SpnPlazo.valor = "0"
            
            ChkProxMes.value = 0
            TxtDiaFijo2.Text = "00"
            'optTipoGracia(1).Enabled = False 'LUCV20180601, Comentó según ERS022-2018
            'optTipoGracia(1).value = False 'LUCV20180601, Comentó según ERS022-2018
            
            txtFechaFija.Enabled = True
            TxtDiaFijo.Enabled = False
            TxtDiaFijo2.Enabled = False
            txtPerGra.Enabled = False
            txtPerGra.Text = "0"
            TxtTasaGracia.Text = "0.00"
            
            'optTipoGracia(0).value = False 'LUCV20180601, Comentó según ERS022-2018
            'optTipoGracia(1).value = False 'LUCV20180601, Comentó según ERS022-2018
            'cmdgracia.Enabled = True 'LUCV20180601, Comentó según ERS022-2018
            
            txtFechaFija.Text = gdFecSis
            GenerarFechaPago
        End If
        '***
     End If
End Sub

Private Sub TxtFecDesemb_LostFocus()
    If ValidaFecha(TxtFecDesemb.Text) <> "" Then
        MsgBox ValidaFecha(TxtFecDesemb.Text) & " en la Fecha de desembolso", vbInformation, "Aviso"
        TxtFecDesemb.SetFocus
        Exit Sub
    End If
    If CDate(TxtFecDesemb.Text) < gdFecSis Then
        MsgBox "La Fecha de desembolso no puede ser menor a la Fecha Actual", vbInformation, "Aviso"
        TxtFecDesemb.SetFocus
    End If
End Sub

Private Sub txtTasacion_Change()
    CalcularDatosEmpVinculados
End Sub
'**** END BRGO ***************************************************
'->***** LUCV20180601, Comentó según ERS022-2018
'Private Sub optTipoGracia_Click(Index As Integer)
'If Index = 0 Then
'    chkIncremenK.Visible = True
'Else
'    chkIncremenK.Visible = False
'End If
'cmdgracia.Enabled = False
'End Sub
'<-***** Fin LUCV20180601
Private Sub txtBuscarLinea_EmiteDatos()
Dim sCodigo As String
Dim sCtaCodOrigen As String 'DAOR 20070407, para el caso de refinanciados
Dim oLineas As COMDCredito.DCOMLineaCredito
   
sCodigo = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
'sCodigo = Right(cmbSubTipo.Text, 3) & a
If sCodigo <> "" Then
'    txtBuscarLinea.Text = sCodigo
    'If txtBuscarLinea.psDescripcion <> "" Then lblLineaDesc = txtBuscarLinea.psDescripcion Else lblLineaDesc = ""
    If txtBuscarLinea.psDescripcion <> "" Then lblLineaDesc = Trim(Left(txtBuscarLinea.psDescripcion, Len(txtBuscarLinea.psDescripcion) - 50)) Else lblLineaDesc = lblLineaDesc = ""
        'VERIFICAR
       'Carga Datos de la Linea de Credito seleccionada
       Set oLineas = New COMDCredito.DCOMLineaCredito
       'Comentado por DAOR 20070407
       'Set RLinea = oLineas.RecuperaLineadeCredito(sCodigo)
       'Set RLinea = oLineas.RecuperaLineadeCredito(sCodigo)
       Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio(Trim(Right(cmbSubProd.Text, 10)), lnCampanaId, Trim(Right((txtBuscarLinea.psDescripcion), 15)), sCodigo, lblLineaDesc, Mid(ActxCta.NroCuenta, 9, 1), CCur(txtMonSug.Text), IIf(ckcPreferencial.value = 1, 1, 0))
     
'COMENTADO X MADM 20110419 - Refinanciados
'''       '**DAOR 20070407**************************************************
'''       If Not bEsRefinanciado Then
'''            'Set RLinea = oLineas.RecuperaLineadeCredito(sCodigo)
'''            Set RLinea = oLineas.RecuperaLineadeCredito(sCodigo)
'''       Else
'''            sCtaCodOrigen = Right$(txtBuscarLinea.psDescripcion, 18)
'''            Set RLinea = oLineas.RecuperaLineadeCredOrigenRefinanciado(sCtaCodOrigen, sCodigo)
'''            'Set RLinea = oLineas.RecuperaLineadeCredOrigenRefinanciado(sCtaCodOrigen, Right(cmbSubTipo.Text, 3))
'''
'''       End If
'''       '*****************************************************************
       
       Set oLineas = Nothing
       If RLinea.RecordCount > 0 Then
          'Call CargaDatosLinea
          Set objProducto = New COMDCredito.DCOMCredito
          If objProducto.GetResultadoCondicionCatalogo("N0000026", ActxCta.Prod) Then     '**END ARLO
          'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
            Txtinteres.Text = lnTasaPeriodoLeasing * 100
            TxtTasaGracia.Text = lnTasaPeriodoLeasing * 100
          End If
       Else
'ALPA 20100610 B2*******************
'          MsgBox "No existen Líneas de Crédito con el Plazo seleccionado", vbInformation, "Aviso"
'          txtBuscarLinea.Text = ""
'          lblLineaDesc = ""
            lnTasaInicial = 0
            lnTasaFinal = 0
            
'JOEP ERS007-2018 20180210
            lnTasaGraciaInicial = 0
            lnTasaGraciaFinal = 0
'JOEP ERS007-2018 20180210
            
            If nMostrarLineaCred = 0 Then
                       MsgBox "No existen Líneas de Crédito con el Plazo seleccionado", vbInformation, "Aviso"
                       txtBuscarLinea.Text = ""
                       lblLineaDesc = ""
            End If
'***********************************
       End If
    
    'txtBuscarLinea.Text = Mid(txtBuscarLinea.Text, 5, Len(txtBuscarLinea.Text))
    
Else
    lblLineaDesc = ""
End If
bCalendGen = False 'ALPA20141206
End Sub

Private Sub ChkMiViv_Click()
    If ChkMiViv.value = 1 Then
        'opttcuota(0).value = True
        'opttcuota(1).Enabled = False
        'opttcuota(2).Enabled = False
        cmbTipoCuota.ListIndex = 0
        'opttper(1).value = True
        'opttper(0).Enabled = False
        'cmbTipoPer.ListIndex = 1 'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 1) 'JOEP20190125 CP
        'ChkCuotaCom.Enabled = False 'LUCV20180522, Comentó según ERS022-2018
        
        'OptTipoCalend(0).value = True
        'OptTipoCalend(1).value = False
        'OptTipoCalend(1).Enabled = False
        cmbTipoCalend.ListIndex = 0
        TxtDiaFijo2.Enabled = False
    Else
        'opttcuota(0).value = True
        'opttcuota(1).Enabled = True
        'opttcuota(2).Enabled = True
        cmbTipoCuota.ListIndex = 0
        'opttper(0).value = True
        'opttper(0).Enabled = True
        'cmbTipoPer.ListIndex = 0'Comento JOEP20190125 CP
        cmbTipoPer.ListIndex = IndiceListaCombo(cmbTipoPer, 0) 'JOEP20190125 CP
        'ChkCuotaCom.Enabled = True 'LUCV20180522, Comentó según ERS022-2018
        'OptTipoCalend(0).value = True
        'OptTipoCalend(1).value = False
        'OptTipoCalend(1).Enabled = True
        cmbTipoCalend.ListIndex = 0
        'TxtDiaFijo2.Enabled = True 'ARCV 30-04-2007
    End If
End Sub

Private Sub ChkProxMes_Click()
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
End Sub

Private Function ValidaLineaCredito(ByVal psValor As String) As Boolean
Dim nPlazoTmp As Integer
'Dim oCred As COMNCredito.NCOMCredito

'Vamos a pasarlo por referencia para no hacer doble conexion
'Dim sValor As String

    'Call UbicaRegistro(Trim(Right(Cmblincre.Text, 20)))
    
    ValidaLineaCredito = True
    
    'Valida Tasa Interes Comp.
'    If Txtinteres.Visible Then
'        If CDbl(Txtinteres.Text) < RLinea!nTasaIni Or CDbl(Txtinteres.Text) > RLinea!nTasafin Then
'            MsgBox "La Tasa de Interes No es Permitida por la Linea de Credito", vbInformation, "Aviso"
'            ValidaLineaCredito = False
'            Txtinteres.SetFocus
'            Exit Function
'        End If
'    End If
    
    'Valida Tasa Interes Gracia.
'    If CInt(txtPerGra.Text) > 0 And TxtTasaGracia.Visible Then
'        If CDbl(TxtTasaGracia.Text) < RLinea!nTasaGraciaIni Or CDbl(TxtTasaGracia.Text) > RLinea!nTasaGraciaFin Then
'            'MsgBox "La Tasa de Interes para el Periodo de Gracia, No es Permitida por la Linea de Credito", vbInformation, "Aviso"
'            MsgBox "La Tasa de Interes para el Periodo de Gracia, No es Permitida por la Configuración del Producto", vbInformation, "Aviso"
'            ValidaLineaCredito = False
'            TxtTasaGracia.SetFocus
'            Exit Function
'        End If
'    End If
    
    'Valida Plazo
'    nPlazoTmp = 0
'    If CmdDesembolsos.Enabled Then
'        If UBound(MatDesPar) > 0 Then
'            nPlazoTmp = CDate(MatrizCal(UBound(MatrizCal) - 1, 0)) - CDate(MatDesPar(UBound(MatDesPar) - 1, 0))
'        End If
'    Else
'        nPlazoTmp = CDate(MatrizCal(UBound(MatrizCal) - 1, 0)) - CDate(Format(gdFecSis, "dd/mm/yyyy"))
'    End If
'    'ALPA 20150113*****************************************************************************************
'    'If nPlazoTmp < RLinea!nPlazoMin Or nPlazoTmp > RLinea!nplazomax Then
'    If nPlazoTmp < RLinea!nPlazoMin * 30 Or nPlazoTmp > RLinea!nPlazoMax * 30 Then
'
'        MsgBox "El Plazo del Credito, No es Permitido por la Linea de Credito", vbInformation, "Aviso"
'        ValidaLineaCredito = False
'        Exit Function
'    End If
    '******************************************************************************************************
    'Valida Monto Sugerido
'    If CDbl(txtMonSug.Text) < RLinea!nMontoMin Or CDbl(txtMonSug.Text) > RLinea!nMontoMax Then
'        MsgBox "El Monto del Credito, No es Permitido por la Linea de Credito", vbInformation, "Aviso"
'        ValidaLineaCredito = False
'        If txtMonSug.Enabled Then
'            txtMonSug.SetFocus
'        End If
'        Exit Function
'    End If
    
    'Verifica el Monto del Prestamo (solo es un aviso de Informacion, mas no de restriccion para ser grabado el credito)
    'Set oCred = New COMNCredito.NCOMCredito
    'sValor = oCred.ValidaMontoPrestamo(ActxCta.NroCuenta, CDbl(txtmonsug.Text), gdFecSis)
    If psValor <> "" Then
        MsgBox psValor, vbInformation, "Aviso"
    End If
    
    'Set oCred = Nothing
End Function

'Private Sub Cmblincre_Click()
'    ReDim MatCalend(0, 0)
'    ReDim MatrizCal(0, 0)
'    If Trim(Cmblincre.Text) = "" Then
'        Exit Sub
'    End If
'    Call UbicaRegistro(Trim(Right(Cmblincre.Text, 20)))
'    If RLinea!nTasaIni <> RLinea!nTasafin Then
'        Txtinteres.Visible = True
'        Txtinteres.Enabled = True
'        LblInteres.Enabled = False
'        LblInteres.Visible = False
'        Txtinteres.Text = Format(RLinea!nTasaIni, "#0.0000")
'        TxtTasaGracia.Text = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
'        Txtinteres.ToolTipText = "Minima : " & Format(RLinea!nTasaIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasafin, "#0.0000")
'    Else
'        Txtinteres.Visible = False
'        Txtinteres.Enabled = False
'        LblInteres.Enabled = True
'        LblInteres.Visible = True
'        LblInteres.Caption = Format(RLinea!nTasaIni, "#0.0000")
'        LblTasaGracia.Caption = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
'    End If
'    If RLinea!nTasaGraciaIni <> RLinea!nTasaGraciaFin Then
'        TxtTasaGracia.Visible = True
'        TxtTasaGracia.Enabled = True
'        LblTasaGracia.Enabled = False
'        LblTasaGracia.Visible = False
'        TxtTasaGracia.Text = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
'        TxtTasaGracia.ToolTipText = "Minima : " & Format(RLinea!nTasaGraciaIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaGraciaFin, "#0.0000")
'    Else
'        TxtTasaGracia.Visible = False
'        TxtTasaGracia.Enabled = False
'        LblTasaGracia.Enabled = True
'        LblTasaGracia.Visible = True
'        LblTasaGracia.Caption = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
'    End If
'
'    If RLinea!nTasaMoraIni <> RLinea!nTasaMoraFin Then
'        TxtMora.Visible = True
'        TxtMora.Enabled = True
'        LblMora.Enabled = False
'        LblMora.Visible = False
'        TxtMora.Text = Format(IIf(IsNull(RLinea!nTasaMoraIni), 0, RLinea!nTasaMoraIni), "#0.0000")
'        TxtMora.ToolTipText = "Minima : " & Format(RLinea!nTasaMoraIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaMoraFin, "#0.0000")
'    Else
'        TxtMora.Visible = False
'        TxtMora.Enabled = False
'        LblMora.Enabled = True
'        LblMora.Visible = True
'        LblMora.Caption = Format(IIf(IsNull(RLinea!nTasaMoraIni), 0, RLinea!nTasaMoraIni), "#0.0000")
'    End If
'End Sub

Private Sub CargaDatosLinea()
ReDim MatCalend(0, 0)
ReDim MatrizCal(0, 0)
    
    If Trim(txtBuscarLinea.Text) = "" Then
        Exit Sub
    End If
    'ALPA 20150313**********************
    If RLinea.BOF Or RLinea.EOF Then
        Exit Sub
    End If
    '***********************************
    'ALPA 20150113**************************************
    lnTasaInicial = RLinea!nTasaIni
    lnTasaFinal = RLinea!nTasaFin
    '***************************************************
    
    'JOEP ERS007-2018 20180210**************************************
    lnTasaGraciaInicial = RLinea!nTasaGraciaIni
    lnTasaGraciaFinal = RLinea!nTasaGraciaFin
    '***************************************************
    'Tasa Interés Compensatorio
    If RLinea!nTasaIni <> RLinea!nTasaFin Then
        Txtinteres.Visible = True
        Txtinteres.Enabled = True
        LblInteres.Enabled = False
        LblInteres.Visible = False
        'FRHU 20170519 ACTA 070-2017
        'If Txtinteres.Text >= 0.001 And Txtinteres.Text < RLinea!nTasafin Then
        'Else
            'Txtinteres.Text = Format(RLinea!nTasafin, "#0.0000")
            'txtInteresTasa.Text = Format(RLinea!nTasafin, "#0.0000")
        'End If
        If Trim(LeeConstanteSist(605)) = "0" Then
            If Txtinteres.Text >= 0.001 And Txtinteres.Text < RLinea!nTasaFin Then
            Else
                Txtinteres.Text = Format(RLinea!nTasaFin, "#0.0000")
                txtInteresTasa.Text = Format(RLinea!nTasaFin, "#0.0000")
            End If
        Else
            Txtinteres.Text = Format(RLinea!nTasaFin, "#0.0000")
            txtInteresTasa.Text = Format(RLinea!nTasaFin, "#0.0000")
        End If
        'FIN FRHU 20170519
        'TxtTasaGracia.Text = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
        If Not IsNull(RLinea!nTasaGraciaFin) Then
            If TxtTasaGracia.Text >= 0.001 And TxtTasaGracia.Text < RLinea!nTasaGraciaFin Then
            Else
                TxtTasaGracia.Text = Format(RLinea!nTasaGraciaFin, "#0.0000")
            End If
        Else
            TxtTasaGracia.Text = 0#
        End If
        Txtinteres.ToolTipText = "Minima : " & Format(RLinea!nTasaIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaFin, "#0.0000")
    Else
        Txtinteres.Visible = True
        Txtinteres.Enabled = True
        LblInteres.Enabled = False
        LblInteres.Visible = False
        If Txtinteres.Text >= 0.001 And Txtinteres.Text < RLinea!nTasaFin Then
        Else
            LblInteres.Caption = Format(RLinea!nTasaFin, "#0.0000")
            Txtinteres.Text = Format(RLinea!nTasaFin, "#0.0000")
            txtInteresTasa.Text = Format(RLinea!nTasaFin, "#0.0000")
        End If
        'LblTasaGracia.Caption = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
        If Not IsNull(RLinea!nTasaGraciaFin) Then
            If TxtTasaGracia.Text >= 0.001 And TxtTasaGracia.Text < RLinea!nTasaGraciaFin Then
            Else
                TxtTasaGracia.Text = Format(RLinea!nTasaGraciaFin, "#0.0000")
            End If
        Else
            TxtTasaGracia.Text = 0#
        End If
    End If
    'Tasa Interés Gracia
    If RLinea!nTasaGraciaIni <> RLinea!nTasaGraciaFin Then
        '->***** LUCV20180601, según ERS022-2018
        'TxtTasaGracia.Visible = True
        'TxtTasaGracia.Enabled = True
        fbTieneTasaGracia = True
        '<-***** Fin LUCV20180522.
        
        LblTasaGracia.Enabled = False
        LblTasaGracia.Visible = False
        'TxtTasaGracia.Text = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
        If Not IsNull(RLinea!nTasaGraciaFin) Then
            If TxtTasaGracia.Text >= 0.001 And TxtTasaGracia.Text < RLinea!nTasaGraciaFin Then
            Else
                TxtTasaGracia.Text = Format(RLinea!nTasaGraciaFin, "#0.0000")
            End If
            TxtTasaGracia.ToolTipText = "Minima : " & Format(RLinea!nTasaGraciaIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaGraciaFin, "#0.0000")
        Else
            TxtTasaGracia.Text = 0#
        End If
    Else
        If TxtTasaGracia.Text = "" Then
            '->***** LUCV20180601, Comentó
            'TxtTasaGracia.Visible = False
            'TxtTasaGracia.Enabled = False
            fbTieneTasaGracia = False
            '<-***** Fin LUCV20180601
            
            LblTasaGracia.Enabled = True
            LblTasaGracia.Visible = True
        End If
        'LblTasaGracia.Caption = Format(IIf(IsNull(RLinea!nTasaGraciaIni), 0, RLinea!nTasaGraciaIni), "#0.0000")
        If Not IsNull(RLinea!nTasaGraciaFin) Then
            If TxtTasaGracia.Text >= 0.001 And TxtTasaGracia.Text < RLinea!nTasaGraciaFin Then
            Else
                TxtTasaGracia.Text = Format(RLinea!nTasaGraciaFin, "#0.0000")
            End If
        Else
            TxtTasaGracia.Text = 0#
        End If
    End If
    'Tasa Interés Moratorio
    If RLinea!nTasaMoraIni <> RLinea!nTasaMoraFin Then
        TxtMora.Visible = True
        TxtMora.Enabled = True
        LblMora.Enabled = False
        LblMora.Visible = False
        TxtMora.Text = Format(IIf(IsNull(RLinea!nTasaMoraIni), 0, RLinea!nTasaMoraIni), "#0.0000")
        If Not IsNull(RLinea!nTasaMoraIni) Then
            TxtMora.ToolTipText = "Minima : " & Format(RLinea!nTasaMoraIni, "#0.0000") & "  Maxima : " & Format(RLinea!nTasaMoraFin, "#0.0000")
        End If
    Else
        TxtMora.Visible = False
        TxtMora.Enabled = False
        LblMora.Enabled = False 'LUCV20180601, modificó true a false. según ERS022-2018
        LblMora.Visible = True
        LblMora.Caption = Format(IIf(IsNull(RLinea!nTasaMoraIni), 0, RLinea!nTasaMoraIni), "#0.0000")
    End If
End Sub

'Private Sub Cmblincre_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Txtinteres.Enabled = True Then
'            Txtinteres.SetFocus
'        Else
'            If txtmonsug.Enabled Then
'                txtmonsug.SetFocus
'            End If
'        End If
'
'    End If
'End Sub

Private Sub TxtBuscarLinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Txtinteres.Enabled = True Then
            Txtinteres.SetFocus
        Else
            If txtMonSug.Enabled Then
                txtMonSug.SetFocus
            End If
        End If
        
    End If
End Sub

Private Sub cmdBuscar_Click()
    SpnPlazo.valor = "0" 'MAVM 25102010
    txtFechaFija = "__/__/____" 'MAVM 25102010
    Screen.MousePointer = 11
    nMostrarLineaCred = 0
    bValidaCargaSugerenciaAguaSaneamiento = 0 'EAAS20180912 SEGUN ERS-054-2018
    bValidaCargaSugerenciaCreditoVerde = 0 'EAAS20191401 SEGUN 018-GM-DI_CMACM
    'ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Creditos para Sugerencia de Analista", , , , gsCodAge, bLeasing)
    ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Creditos para Sugerencia de Analista", , , , gsCodAge, bLeasing, , , , gsCodCargo)
    bCheckList = False 'RECO20150513 *****************************
    bEntrotxtMontoSeg = False 'JOEP20190129 CP
    If ActxCta.NroCuenta <> "" Then
        If CargaDatos(ActxCta.NroCuenta) Then
            'MADM 20100513
            'ALPA 20100609 B2**************************************************************************
            'If Mid(ActxCta.NroCuenta, 6, 3) = "302" Then
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("N0000012", sSTipoProdCod) Then     '**END ARLO
            'If sSTipoProdCod <> "517" Then
                Me.SSTab1.TabVisible(2) = False
                Call LimpiaFlex(Me.grdEmpVinculados)
            End If
            Frame3.Enabled = True
            If Mid(sSTipoProdCod, 1, 1) = "7" Then
                cmdActTipoCred.Visible = False
            Else
                cmdActTipoCred.Visible = True
            End If
            
            'If sSTipoProdCod = "703" Then
            '******************************************************************************************
            '    cmdSeleccionaFuente.Enabled = False
            '    cmdFuentes.Enabled = False
            '    Label13.Enabled = False
            'Else
            '    cmdSeleccionaFuente.Enabled = True
            '    cmdFuentes.Enabled = True
            '    Label13.Enabled = True
            'End If
            'END MADM
            
            '** JUEZ 20120907 ******************************************
            'If nAgenciaCredEval = 0 Then
            '    If sSTipoProdCod = "703" Then
            '        cmdSeleccionaFuente.Enabled = False
            '        cmdFuentes.Enabled = False
            '        'Label13.Enabled = False
            '    Else
            '        cmdSeleccionaFuente.Enabled = True
            '        cmdFuentes.Enabled = True
            '        'Label13.Enabled = True
            '    End If
            'Else
            '    cmdSeleccionaFuente.Enabled = False
            '    cmdFuentes.Enabled = False
            '    'Label13.Enabled = False
            'End If
            ''** END JUEZ ***********************************************
            
            cmdrelac.Enabled = True
            FraDatos.Enabled = True
            'Cmblincre.SetFocus
            'txtBuscarLinea.SetFocus
            'spnPlazo.SetFocus
            
            'MAVM 25102010 ***
            TxtDiaFijo.Enabled = False
            '***
            
            CmdCredVig.Enabled = True
            Call HabilitaPermiso
            '07-05-2006
            'If spnPlazo.Enabled Then
            '   spnPlazo.SetFocus
            'Else
            '   CmdCredVig.SetFocus
            'End If
            If CmdGrabar.Enabled = True Then
                CmdGrabar.SetFocus
            End If
            '**************
            CboPersCiiu.Enabled = True 'CUSCO
            CmdGrabar.Enabled = True
            CmdCalend.Enabled = True
            CmdGastos.Enabled = True
            'cmdEvaluacion.Enabled = True
            cmdCheckList.Enabled = True
            'FRHU 20170517 ACTA-070-2017
            If Trim(LeeConstanteSist(605)) = "1" Then
                Txtinteres.Locked = True
                LblInteres.Enabled = False
            End If
            'FIN FRHU 20170517
            'JOEP20180905 Catalogo
                nMatAporteLleno = False
            If cmbSubProd.Text <> "" Then
                Call bValiExistCheckList
            End If
            'JOEP20180905 Catalogo
        Else
            cmdrelac.Enabled = False
            FraDatos.Enabled = False
            CmdGrabar.Enabled = False
            CmdCalend.Enabled = False
            CmdGastos.Enabled = False
            'cmdEvaluacion.Enabled = False
            cmdCheckList.Enabled = False
            ActxCta.Enabled = True
            CmdCredVig.Enabled = False
            CboPersCiiu.Enabled = False 'CUSCO
            'ALPA 20100622****
            Frame3.Enabled = False
            '*****************
            MsgBox "El Credito No Existe", vbInformation, "Aviso"
            ''JUEZ 20120914 ***************************************************
            'If nAgenciaCredEval = 1 Then
            '    If nVerifCredEval = 0 Then
            '        MsgBox "El Credito no ha sido verificado por el Coordinador de Creditos", vbInformation, "Aviso"
            '    Else
            '        MsgBox "El Credito No Existe", vbInformation, "Aviso"
            '    End If
            'ElseIf bVerifNivApr = False Then
            '    'Unload Me
            'Else
            '    MsgBox "El Credito No Existe", vbInformation, "Aviso"
            'End If
            ''END JUEZ ********************************************************
        End If
    Else
        ActxCta.CMAC = gsCodCMAC
        ActxCta.Age = gsCodAge
        ActxCta.SetFocusProd
        ActxCta.Enabled = True
        CmdGrabar.Enabled = False
        CmdCalend.Enabled = False
        CmdGastos.Enabled = False
        'cmdEvaluacion.Enabled = False
        cmdCheckList.Enabled = False 'JOEP20190125 CP
    End If
End Sub

Private Sub cmdCalend_Click()
Dim nTasaInt As Double
Dim i As Integer
Dim CadTmp As String
Dim lsCtaCodLeasing As String
'WIOR 20131111 **************************

'ARLO20190304
If cmbCampana.ListIndex = -1 Then
    MsgBox "Por favor selecione un tipo de campaña", vbInformation, "Aviso"
    cmbCampana.SetFocus
    Exit Sub
End If
'************

Dim lnCuotaBalon As Integer
If chkCuotaBalon.Visible Then
    If chkCuotaBalon.value = 1 Then
        If Trim(txtCuotaBalon.Text) = "" Or Trim(txtCuotaBalon.Text) = "0" Then
            lnCuotaBalon = 0
        Else
            lnCuotaBalon = CInt(Trim(txtCuotaBalon.Text))
        End If
    Else
        lnCuotaBalon = 0
    End If
Else
    lnCuotaBalon = 0
End If
'ALPA 20140612******
Set objProducto = New COMDCredito.DCOMCredito
 If (txtMontoMivivienda.Text = 0# Or txtMontoMivivienda.Text = "") And objProducto.GetResultadoCondicionCatalogo("N0000013", sSTipoProdCod) Then
'If (txtMontoMivivienda.Text = 0# Or txtMontoMivivienda.Text = "") And sSTipoProdCod = "801" Then
     MsgBox "No se olvide de asignar el valor de venta del credito MIVIVIENDA", vbInformation, "Aviso"
     SSTab1.Tab = 0
     txtMontoMivivienda.SetFocus
     Exit Sub
 End If
Set objProducto = New COMDCredito.DCOMCredito
If Round((CDbl(txtMontoMivivienda.Text) - ((CDbl(txtMontoMivivienda.Text) * 1) / 10)), 2) < CDbl(txtMonSug.Text) And objProducto.GetResultadoCondicionCatalogo("N0000014", sSTipoProdCod) Then
'If Round((CDbl(txtMontoMivivienda.Text) - ((CDbl(txtMontoMivivienda.Text) * 1) / 10)), 2) < CDbl(txtMonSug.Text) And sSTipoProdCod = "801" Then
    MsgBox "MIVIVIENDA,no se olvide que el monto de la inicial no debe ser menor al 10% (" & Round((CDbl(txtMonSug.Text) * 10) / 9, 2) & ")", vbInformation, "Aviso"
     txtMontoMivivienda.SetFocus
     SSTab1.Tab = 0
     Exit Sub
 End If
 
 
'WIOR FIN *******************************
    'Validacion para poder generar el calendario Add Gitu 20-08-2008
    'Descomentar cuando esten seguros de los cambios GITU
'    If Not txtFechaFija.Text = "__/__/____" Then
'        If (CDate(txtFechaFija.Text) - gdFecSis) >= 30 Then
'            ChkProxMes.value = 1
'        Else
'            ChkProxMes.value = 0
'        End If
'    Else
'        MsgBox "Debe ingresar la Fecha Fija", vbInformation, "Atencion!"
'        Exit Sub
'    End If
    bCalendGen = True
    
    If Trim(Right(cmbTipoCuota.Text, 2)) = "3" And Trim(Right(cmbDesemb.Text, 2)) = "0" Then
        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = TxtFecDesemb.Text
        MatDesemb(0, 1) = Format(txtMonSug.Text, "#0.00")
    End If
    
        '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000027", ActxCta.Prod) Then     '**END ARLO
    'If ActxCta.Prod = "515" Or ActxCta.Prod = "516" Then
        lsCtaCodLeasing = ActxCta.GetCuenta
    End If
    
    If (DateDiff("d", CDate(TxtFecDesemb.Text), CDate(txtFechaFija.Text)) = 0) And Len(Trim(lsCtaCodLeasing)) = 0 Then
        MsgBox "La Fecha de Pago es la Misma que la Fecha de Desembolso", vbInformation, "Aviso"
        bCalendGen = False
        txtFechaFija.SetFocus
        Exit Sub
    End If
    'ALPA 20141127****************************************
    If chkCSP.value = 1 And lnCSP = 0 Then
        MsgBox "Seleccionar la cuota para generación de concepto de Poliza contra Incendio", vbInformation, "Aviso"
        lnCSP = -1
        chkCSP.value = 0
        Exit Sub
    End If
    '*****************************************************
'JOEP20190114 CP
    
    If Not CP_ValMensaje(2) Then Exit Sub
    If Not CP_CondicionesSeg(Trim(Right(cmbSubProd.Text, 5)), Mid(ActxCta.NroCuenta, 9, 1), txtMonSug.Text, spnCuotas.valor, SpnPlazo.valor, Trim(Right(cmbDestino.Text, 5)), txtPerGra.Text, Trim(Right(cmbTipoCredito.Text, 5)), 2) Then Exit Sub
    If Not CP_ValidaRangoFecha(Trim(Right(cmbSubProd.Text, 9)), 3000, Right(cmbTipoPer.Text, 3), txtFechaFija.Text) Then
        txtFechaFija.SetFocus
        Exit Sub
    End If
    'JOEP20190114 CP
'JOEP 201710 ACta201
Dim rsValidaPriFecPago As ADODB.Recordset
Dim obDCred As COMDCredito.DCOMCredito
Set obDCred = New COMDCredito.DCOMCredito
Set rsValidaPriFecPago = obDCred.ValidaPriFecPago(CDate(TxtFecDesemb.Text), CDate(txtFechaFija.Text))

If Not (rsValidaPriFecPago.EOF And rsValidaPriFecPago.BOF) Then
    If rsValidaPriFecPago!cMensaje <> "" Then
        MsgBox rsValidaPriFecPago!cMensaje, vbInformation, "No podrá continuar"
        rsValidaPriFecPago.Close
        Set obDCred = Nothing
        Exit Sub
    End If
rsValidaPriFecPago.Close
Set obDCred = Nothing
End If
'JOEP 201710 ACta201
    
    
    'WIOR 20160224 ***
    If fnCantAfiliadosSegDes > 0 Then
        MsgBox "Tasa del Seguro Desgravamen: " & Format(fnTasaSegDes, "#0.000000") & "%" & Chr(10) & fnCantAfiliadosSegDes & " Afiliado(s)", vbInformation, "Tasas"
    End If
    'WIOR FIN ********
    
     '*****->LUCV20170915, Creó. Según Modificaciones del reglamento (4.2.-Crédito Refinanciado. (b))
    If Not ValidaPeriodoGracia Then
        txtFechaFija.SetFocus
        chkGracia.value = 0
        Exit Sub
    End If

    'Estimación de la fecha de vencimiento
    Dim dDesembolso As Date
    Dim dFechaVencEstimada As Date
    Dim nDia As Integer
    Dim nMes As Integer
    Dim nAnio As Integer
    Dim nMontoPoliza As Double
    Dim nTasaMensualSegInc As Double
    Dim oNCOMCredito As COMNCredito.NCOMCredito
    Set oNCOMCredito = New NCOMCredito
    Dim rsMontoPoliza As ADODB.Recordset
    Set rsMontoPoliza = New ADODB.Recordset
    i = 1
    
    dDesembolso = CDate(Format(TxtFecDesemb.Text, "dd/mm/yyyy")) + CInt(txtPerGra.Text)
    nDia = Day(dDesembolso): nMes = Month(dDesembolso): nAnio = Year(dDesembolso)
    
    For i = 1 To CInt(spnCuotas.valor)
        If IIf(Trim(Right(cmbTipoPer.Text, 2)) = "0", 1, 2) = PeriodoFijo Then
            dFechaVencEstimada = dDesembolso
            dFechaVencEstimada = dFechaVencEstimada + CInt(SpnPlazo.valor)
        Else
            If i = 1 And (txtPerGra.Text) <> 0 Then
                dFechaVencEstimada = dDesembolso + 30
                nDia = Day(dFechaVencEstimada): nMes = Month(dFechaVencEstimada): nAnio = Year(dFechaVencEstimada)
            Else
                nMes = nMes + 1
                oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                
                'Artificio para cuando el desembolso sea a fines de enero [Febrero]
                If CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio))) - CDate(Format(TxtFecDesemb.Text, "dd/mm/yyyy")) < 5 Then
                    nMes = nMes + 1
                    oNCOMCredito.ValidaFechasFijasCuota nDia, nMes, nAnio
                End If
            End If
            
            dFechaVencEstimada = CDate(Right("0" & Trim(str(nDia)), 2) & "/" & Right("0" & Trim(str(nMes)), 2) & "/" & Trim(str(nAnio)))
        End If
    Next i
     i = 0
    'Obtiene el importe de la cuota de la póliza
    If oNCOMCredito.RequierePoliza(ActxCta.NroCuenta, Right(cmbSubTipo.Text, 3), (val(txtMonSug.Text)), dFechaVencEstimada) Then 'Or Not (rsMontoPoliza.EOF And rsMontoPoliza.BOF) Then
    Set rsMontoPoliza = oNCOMCredito.RecuperaValoresGarantiaPoliza(ActxCta.NroCuenta)
       Do While Not rsMontoPoliza.EOF
            nMontoPoliza = nMontoPoliza + rsMontoPoliza!nPrimaCuota
            nTasaMensualSegInc = nTasaMensualSegInc + rsMontoPoliza!nTasaSeguroInc
            rsMontoPoliza.MoveNext
       Loop
    Else
        nMontoPoliza = 0
        nTasaMensualSegInc = 0
    End If
    Set oNCOMCredito = Nothing
    'Fin LUCV20170915 <-*****
    
    If DameTipoCuota = 1 Or DameTipoCuota = 2 Or DameTipoCuota = 3 Then
            If Txtinteres.Visible Then
                'ALPA 20150113****************************************************************
                'nTasaInt = CDbl(Txtinteres.Text)
                nTasaInt = CDbl(IIf(chkTasa.value = 1, txtInteresTasa.Text, Txtinteres.Text))
                '*****************************************************************************
            Else
                nTasaInt = CDbl(LblInteres.Caption)
            End If
            If ValidaDatosCalendario Then
                'RECO20150520 ERS023-2015*************************
                If Right(cmbBancaSeguro.Text, 1) = "2" Then
                    vMatriz = frmGarantMultiriesgoMYPE.inicia(oPersona.PersCodigo, txtMonSug.Text, ActxCta.NroCuenta, spnCuotas.valor)
                    If UBound(vMatriz, 2) = 0 Then
                        MsgBox "No se ha seleccioando ninguna garantìa para ser coberturada por el seguro Multiriesgos MYPE", vbInformation, "Alerta"
                        Exit Sub
                    End If
                End If
                'RECO FIN*****************************************
                If CmdDesembolsos.Enabled Then
                    If CmdDesembolsos.Enabled Then
                        If UBound(MatDesPar) = 0 Then
                            MsgBox "Ingrese los Desembolsos Parciales", vbInformation, "Aviso"
                            bCalendGen = False
                            CmdDesembolsos.SetFocus
                            Exit Sub
                        End If
                     End If
                     
                    MatCalend = frmCredCalendPagos.Inicio(CDbl(txtMonSug.Text), nTasaInt, CInt(spnCuotas.valor), _
                    CInt(SpnPlazo.valor), CDate(MatDesPar(0, 0)), DameTipoCuota, IIf(Trim(Right(cmbTipoPer.Text, 2)) = "0", 1, 2), _
                    vnTipoGracia, CInt(txtPerGra.Text), CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                    ChkProxMes.value, MatGracia, ChkMiViv.value, ChkCuotaCom.value, MatCalend_2, 1, , True, MatDesPar, , _
                    ActxCta.NroCuenta, , , , , , , lsCtaCodLeasing, , , lnCuotaBalon, , , CCur(txtMontoMivivienda.Text), _
                    lnCSP, fArrMIVIVIENDA, fnTasaSegDes)
                    'WIOR 20131111 AGREGO lnCuotaBalon
                    'WIOR 20151223 AGREGO fArrMIVIVIENDA
                    'LUCV20180601, Agregó fnTasaSegDes según ERS022-2018
                Else
                    'If optTipoGracia(0).value Then vnTipoGracia = gColocTiposGraciaCapitalizada 'LUCV20180601, Comentó según ERS022-2018
                    'If optTipoGracia(1).value Then vnTipoGracia = gColocTiposGraciaEnCuotas 'LUCV20180601, Comentó según ERS022-2018
                    vnTipoGracia = gColocTiposGraciaPrimeraCuota 'LUCV20180601, Agregó según ERS022-2018
                    MatCalend = frmCredCalendPagos.Inicio(CDbl(txtMonSug.Text), nTasaInt, CInt(spnCuotas.valor), _
                        CInt(SpnPlazo.valor), gdFecSis, DameTipoCuota, IIf(Trim(Right(cmbTipoPer.Text, 2)) = "0", 1, 2), vnTipoGracia, _
                        CInt(txtPerGra.Text), CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                        ChkProxMes.value, MatGracia, ChkMiViv.value, ChkCuotaCom.value, MatCalend_2, 1, , , , IIf(ChkTrabajadores.value = 1, True, False), _
                        ActxCta.NroCuenta, , CInt(IIf(Trim(TxtDiaFijo2.Text) = "", "00", TxtDiaFijo2.Text)), _
                        IIf(chkIncremenK.value = 1, True, False), chkGracia.value, txtFechaFija.Text, , lsCtaCodLeasing, , , lnCuotaBalon, _
                        , , CCur(txtMontoMivivienda.Text), lnCSP, fArrMIVIVIENDA, fnTasaSegDes, MatCalendSegDes, ChkExoSegDes.value, nMontoPoliza, nTasaMensualSegInc)
                        'MAVM ADD chkGracia.value, txtFechaFija.Text 25102010
                        'WIOR 20131111 AGREGO lnCuotaBalon
                        'WIOR 20151223 AGREGO fArrMIVIVIENDA
                        'LUCV20180601, Agregó fnTasaSegDes, MatCalendSegDes, ChkExoSegDes.value. Según ERS022-2018
                End If
                If UBound(MatCalend) <> 0 Then
                    Me.lblcuota.Caption = MatCalend(0, 2)
                End If
            Else
                Exit Sub
            End If
    Else
        'Desembolosos parciales
        If CmdDesembolsos.Enabled Then
            If UBound(MatDesPar) = 0 Then
                MsgBox "Ingrese los Desembolsos Parciales", vbInformation, "Aviso"
                bCalendGen = False
                CmdDesembolsos.SetFocus
                Exit Sub
            End If
        End If
        For i = 0 To UBound(MatrizCal) - 1
            CadTmp = MatrizCal(i, 0)
            MatrizCal(i, 0) = MatrizCal(i, 1)
            MatrizCal(i, 1) = CadTmp
        Next i
        If CmdDesembolsos.Enabled Then
'            MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(False, CDate(MatDesPar(UBound(MatDesPar) - 1, 0)), MatrizCal, CDbl(txtMonSug.Text), Trim(Right(cmbDesemb.Text, 2)), CDbl(Txtinteres.Text))
        Else
'            MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(False, gdFecSis, MatrizCal, CDbl(txtMonSug.Text), Trim(Right(cmbDesemb.Text, 2)), IIf(Txtinteres.Visible, CDbl(Txtinteres.Text), LblInteres.Caption))
            ReDim MatDesemb(1, 2)
            MatDesemb(0, 0) = TxtFecDesemb.Text
            MatDesemb(0, 1) = Format(txtMonSug.Text, "#0.00")
        End If
    End If
    
    'Control RCC
    '->**** LUCV20180601, Comentó según ERS022-2018
'    If bControlRCC = True Then
'        'Solo para Mes y Consumo
'        If Mid(ActxCta.NroCuenta, 6, 1) = Mid(gColPYMEEmp, 1, 1) Or _
'         Mid(ActxCta.NroCuenta, 6, 1) = Mid(gColConsuPlazoFijo, 1, 1) Then
'            Call IdentificaExposicionRCC(MatCalend)
'        End If
'    End If
    '<-***** Fin LUCV20180601
    '*********************
    'ALPA 20120413
    'ReDim MatrizCal(UBound(MatCalend), 6)
    'ReDim MatrizCal(UBound(MatCalend), 7) 'LUCV20180601, Comentó según ERS022-2018
    ReDim MatrizCal(UBound(MatCalend), 9) 'LUCV20180601, Según ERS022-2018
        For i = 0 To UBound(MatCalend) - 1
    'LUCV20180601, Según ERS022-2018
'            If Trim(Right(cmbTipoCuota.Text, 2)) = "3" Then
'                MatrizCal(i, 0) = MatCalend(i, 1)
'                MatrizCal(i, 1) = MatCalend(i, 0)
'            Else
'                MatrizCal(i, 0) = MatCalend(i, 0)
'                MatrizCal(i, 1) = MatCalend(i, 1)
'            End If
    ' Fin LUCV20180601
            MatrizCal(i, 0) = MatCalend(i, 0)
            MatrizCal(i, 1) = MatCalend(i, 1)
            MatrizCal(i, 2) = MatCalend(i, 2)
            MatrizCal(i, 3) = MatCalend(i, 3)
            MatrizCal(i, 4) = MatCalend(i, 4)
            MatrizCal(i, 5) = MatCalend(i, 5)
            MatrizCal(i, 6) = MatCalend(i, 6) 'ALPA 20120413
            MatrizCal(i, 7) = MatCalend(i, 8) 'LUCV20180601, según ERS022-2018
            MatrizCal(i, 8) = MatCalend(i, 15) 'LUCV20180601, según ERS022-2018(Poliza Contra Incendio)
            MatrizCal(i, 9) = MatCalend(i, 16) 'LUCV20180601, según ERS022-2018(Poliza Contra Incendio - Gracia)
        Next i
    'If nAgenciaCredEval = 0 Then
    '    cmdEvaluacion.Enabled = True
    'End If
End Sub

Private Sub IdentificaExposicionRCC(ByVal pMatCalend As Variant)
Dim nMontoCuota  As Double
Dim i As Integer
    'control RCC
    If UBound(pMatCalend) <> 0 Then
        If Not IsArray(pMatCalend) Then
            chkExpuestoRCC.value = 0
            Exit Sub
        End If
        'If opttcuota(0).value Or opttcuota(2).value Then  'Fijo o Decreciente
        If Trim(Right(cmbTipoCuota.Text, 2)) = "0" Or Trim(Right(cmbTipoCuota.Text, 2)) = "2" Then
            nMontoCuota = CDbl(pMatCalend(0, 2))
        Else
            'If opttcuota(1).value Then  'Creciente
            If Trim(Right(cmbTipoCuota.Text, 2)) = "1" Then
                nMontoCuota = CDbl(pMatCalend(UBound(pMatCalend) - 1, 2))
            Else    'Cuota Libre
                nMontoCuota = CDbl(pMatCalend(0, 2))
                For i = 1 To UBound(pMatCalend) - 1 'Buscar el maximo
                    If CDbl(pMatCalend(i, 2)) > nMontoCuota Then
                        nMontoCuota = CDbl(pMatCalend(i, 2))
                    End If
                Next i
            End If
        End If
        nMontoCuota = IIf(Mid(ActxCta.NroCuenta, 6, 1) = "1", nMontoCuota, nMontoCuota * gnTipCambio) 'Soles o Dolares
        'Calculo
        chkExpuestoRCC.value = IIf(nSaldoDisponible / nMontoCuota > 100, 1, 0)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
    ReDim MatFuentes(0)
    'Catalogo JOEP
    Call Form_Load
    cmdCheckList.Enabled = False
    bEntrotxtMontoSeg = False
    ReDim nMatMontoPre(0)
    Screen.MousePointer = vbDefault 'LUCV20180601, Agregó según ERS022-2018
    'JOEP20190129 CP
End Sub

Private Sub CmdDesembolsos_Click()
Dim nSumaDesPar As Double
Dim i As Integer
Dim MonDesAnt As Double

    MonDesAnt = CDbl(txtMonSug.Text)
    MatDesPar = frmCredDesembParcial.Inicio(gdFecSis, MatDesPar)
    If UBound(MatDesPar) > 0 Then
        bDesembParcialGenerado = True
        nSumaDesPar = 0
        For i = 0 To UBound(MatDesPar) - 1
            nSumaDesPar = nSumaDesPar + CDbl(MatDesPar(i, 1))
        Next i
        txtMonSug.Text = Format(nSumaDesPar, "#0.00")
        If MonDesAnt <> nSumaDesPar Then
            ReDim MatCalend(0, 0)
            ReDim MatrizCal(0, 0)
        End If
    Else
        bDesembParcialGenerado = False
        ReDim MatCalend(0, 0)
        ReDim MatrizCal(0, 0)
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim sValor As String
    
    If Not ValidaDatosGrabar(sValor) Then Exit Sub

    If fnCantAfiliadosSegDes > 0 Then
        MsgBox "Tasa del Seguro Desgravamen: " & Format(fnTasaSegDes, "#0.000") & "%" & Chr(10) & fnCantAfiliadosSegDes & " Afiliado(s)", vbInformation, "Tasas"
    End If
    
    If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

    'RIRO 20170526 ********
    Dim sValidacion As String
    sValidacion = validaEstado
    If Len(Trim(sValidacion)) > 0 Then
        MsgBox sValidacion, vbInformation, "Aviso"
        Exit Sub
    End If
    'END RIRO ************

    Call GrabarDatos
    Call cmdCancelar_Click
End Sub

'RIRO 20170529 ***
Private Function validaEstado() As String
    Dim sMensaje As String
    Dim objCred As COMDCredito.DCOMCredito
    Dim nEstadoCred As Integer
    sMensaje = ""
    Set objCred = New COMDCredito.DCOMCredito
    nEstadoCred = objCred.RecuperaEstadoCredito(Me.ActxCta.NroCuenta)
    If nEstadoCred <> gColocEstSolic And nEstadoCred <> gColocEstSug Then
        sMensaje = "El crédito posee un estado NO válido"
    End If
    validaEstado = sMensaje
End Function
'END RIRO ****
'->***** LUCV20180601, Comentó Según ERS022-2018
'Private Sub CmdGracia_Click()
'Dim oCredito As COMNCredito.NCOMCredito
'Dim nTasa As Double
'    If Trim(txtPerGra.Text) <> "" Then
'        If CInt(txtPerGra.Text) <= 0 Then
'            MsgBox "Los Dias del Periodo de Gracia debe ser mayor que Cero", vbInformation, "Aviso"
'            'txtPerGra.SetFocus 'Comentado Por MAVM 25102010
'            txtFechaFija.SetFocus 'MAVM 25102010
'            Exit Sub
'        End If
'    End If
'
'    If Trim(LblTasaGracia.Caption) <> "" And LblTasaGracia.Visible Then
'        If CDbl(LblTasaGracia.Caption) <= 0 Then
'            MsgBox "La Linea de Credito No Tiene Definida un Tasa de Gracia", vbInformation, "Aviso"
'            'Cmblincre.SetFocus
'            txtBuscarLinea.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    'MAVM 25102010 ***
'    If CDbl(TxtTasaGracia.Text) <= 0# Then
'        MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso"
'        TxtTasaGracia.SetFocus
'        Exit Sub
'    End If
'    '***
'
'      'JOEP ERS007-2018 20180210
'        Set objProducto = New COMDCredito.DCOMCredito
'        If objProducto.GetResultadoCondicionCatalogo("N0000028", ActxCta.Prod) And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then
'        'If ActxCta.Prod = "703" And Int(Trim(Right(Me.cmbCampana.Text, 3))) = 88 And CInt(Trim(Right(cmbDestino.Text, 5))) = 4 Then
'            If TxtTasaGracia = "" Then
'            Else
'                If TxtTasaGracia >= lnTasaGraciaInicial And TxtTasaGracia <= lnTasaGraciaFinal Then
'                Else
'                    MsgBox "La T.G: esta fuera del Rango: Mínima : " & Format(lnTasaGraciaInicial, "#0.0000") & "  Máxima : " & Format(lnTasaGraciaFinal, "#0.0000"), vbInformation, "Aviso"
'                    TxtTasaGracia.Text = Format(lnTasaGraciaFinal, "#0.0000")
'                    TxtTasaGracia.SetFocus
'                    Exit Sub
'                End If
'            End If
'        End If
'  'JOEP ERS007-2018 20180210
'
'
'    If Txtinteres.Visible Then
'        nTasa = CDbl(Me.TxtTasaGracia.Text)
'    Else
'        'nTasa = CDbl(LblTasaGracia.Caption)
'        nTasa = IIf(CDbl(LblTasaGracia.Caption) = 0, CDbl(Me.TxtTasaGracia.Text), CDbl(LblTasaGracia.Caption))
'    End If
'
'    Set oCredito = New COMNCredito.NCOMCredito
'    MatGracia = frmCredGracia.Inicio(CInt(txtPerGra.Text), oCredito.MontoIntPerDias(nTasa, CInt(txtPerGra.Text), CDbl(txtMonSug.Text)), CInt(spnCuotas.valor), vnTipoGracia, ActxCta.NroCuenta, MatGracia)
'    Set oCredito = Nothing
'    bGraciaGenerada = True
'End Sub
'<-***** Fin LUCV20180601

Private Sub cmdrelac_Click()
Dim oCredPersRela As UCredRelac_Cli
    Set oCredPersRela = New UCredRelac_Cli
    Call oCredPersRela.CargaRelacPersCred(ActxCta.NroCuenta)
    Call frmCredRelaCta.Inicio(oCredPersRela, InicioSolicitud, InicioConsultaForm)
    Set oCredPersRela = Nothing
End Sub

Private Sub cmdsalir_Click()
    fbSalirCargaDatos = False 'EJVG20151104
 'JOEP20190130 CP
 cmdCheckList.Enabled = False
 ReDim nMatMontoPre(0)
  'JOEP20190130 CP
    If vbInicioCargaDatos Then
        fbSalirCargaDatos = True 'EJVG20151104
        Unload Me
        'Unload frmCredGarantCred
        Unload frmGarantiaCobertura 'EJVG20150707
        Unload frmCredSolicitud
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
Dim oTipoCambio As COMDConstSistema.NCOMTipoCambio
Set oTipoCambio = New COMDConstSistema.NCOMTipoCambio
    If Not oTipoCambio.ExisteTipoCambio(Format(gdFecSis, "mm/dd/yyyy")) Then
        Set oTipoCambio = Nothing
        MsgBox "Falta Ingresar el Tipo de Cambio", vbInformation, "Aviso"
        Unload Me
    End If
    Set oTipoCambio = Nothing
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    lblMsj.Visible = False
End Sub

Private Sub Form_Load()
    'ALPA 20100612*******************************
    Dim lrsTipoCredito As ADODB.Recordset
    Dim oCredito As COMDCredito.DCOMCredito
    '********************************************
    'WIOR 20120510*******************************
    Dim oConstante As COMDConstantes.DCOMConstantes
    Dim rsConstante As ADODB.Recordset
    'WIOR FIN ***********************************
    'cmdEvaluacion.Enabled = False
'JOEP20190129 CP
    sCodPers = ""
    nMatAporteLleno = False
    bEntrotxtMontoSeg = False
    ReDim nMatMontoPre(0)
'JOEP20190129 CP
    CentraForm Me
    Dim sMaTem() As String
    ReDim MatGracia(0, 0)
    ReDim MatCalend(0, 0)
    ReDim MatDesemb(0, 0)
    ReDim MatGastos(0, 0)
    ReDim Preserve sMaTem(3, 1)
    bCalendGen = False
    bGastosGen = False
    nNumGastos = 0
    'ReDim Preserve MatFuentesF(3, 1)
    MatFuentesF = sMaTem
    MatFuentesF(3, 1) = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    actxCtaAho.CMAC = gsCodCMAC
    actxCtaAho.Age = gsCodAge
    actxCtaAho.Prod = gCapAhorros
    ReDim MatrizCal(0, 0)
    bDesembParcialGenerado = False
    Call CargarDatosCarga
    Call GetTipCambio(gdFecSis)
    chkVAC.Visible = False
    'Manejo de Operaciones VAC
    If gsProyectoActual = "H" Then
        chkVAC.Visible = True
    End If
    ChkBloqueo.value = 0
    ChkExoneracion.value = 0
    ChkExoSegDes.value = 0
    chkDesembBcoNac.value = 0
    MatCredVig = ""
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredRegistraSugerenciaAnalista
    lblInstitucionFinanciera.Visible = False
    cmbInstitucionFinanciera.Visible = False
    cmbDatoVivienda.Visible = False 'JUEZ 20130913
    nActualizaTipoCred = 0
    'ALPA 20100604 B2**************************************************************
    Set oCredito = New COMDCredito.DCOMCredito
    Set lrsTipoCredito = oCredito.RecuperaTipoCreditos
    Set oCredito = Nothing
    Call Llenar_Combo_con_Recordset(lrsTipoCredito, cmbTipoCredito)
    Call CambiaTamañoCombo(cmbTipoCredito, 250)
    Set lrsTipoCredito = Nothing
    'Producto
    Set oCredito = New COMDCredito.DCOMCredito
    Set lrsTipoCredito = oCredito.RecuperaProductosCrediticios
    Call Llenar_Combo_con_Recordset(lrsTipoCredito, cmbProducto)
    Call CambiaTamañoCombo(cmbProducto, 250)
    Set lrsTipoCredito = Nothing
    'Destino
    Set oConstante = New COMDConstantes.DCOMConstantes
    Set rsConstante = oConstante.RecuperaConstantes(gColocDestino)
    Call Llenar_Combo_con_Recordset(rsConstante, cmbDestino)
    Set rsConstante = Nothing
    'Campañas
    Set oCredito = New COMDCredito.DCOMCredito
    Set rsConstante = oCredito.RecuperaCampanas(gsCodAge)
    Call Llenar_Combo_con_Recordset(rsConstante, cmbCampana)
    Set rsConstante = Nothing
    cmbTipoCuota.ListIndex = 0
    'cmbTipoPer.ListIndex = 0 'Comento JOEP20190114 CP
    cmbTipoPer.ListIndex = -1 'JOEP20190114 CP
    '******************************************************************************
    SSTab1.TabVisible(2) = False 'BRGO 20111104
    If bLeasing = True Then
        Me.Caption = "Sugerencia de Arrendamiento Financiero"
        ActxCta.texto = "Operación"
        Frame3.Caption = "Datos de la Operación"
    End If
    'WIOR 20120510 ******************************************************************
    Set oConstante = New COMDConstantes.DCOMConstantes
    'Combo Microseguro
    Set rsConstante = oConstante.RecuperaConstantes(9992)
    Call Llenar_Combo_con_Recordset(rsConstante, cmbMicroseguro)
    cmbMicroseguro.ListIndex = IndiceListaCombo(cmbMicroseguro, 0)
    Call CambiaTamañoCombo(cmbMicroseguro, 130)
    Set rsConstante = Nothing
    'Combo Multiriesgo
    Set rsConstante = oConstante.RecuperaConstantes(9993)
    Call Llenar_Combo_con_Recordset(rsConstante, cmbBancaSeguro)
    cmbBancaSeguro.ListIndex = IndiceListaCombo(cmbBancaSeguro, 0)
    Call CambiaTamañoCombo(cmbBancaSeguro, 100)
    Set oConstante = Nothing
    Set rsConstante = Nothing
    'WIOR FIN ***********************************************************************
    'WIOR 20120517 ***********************************
    fbMicroseguro = False
    fnMicroseguro = 0
    fbMultiriesgo = False
    'WIOR -FIN ***************************************
    fbAutoCalfCPP = False 'RECO20160628 ERS002-2016
    chkAutoCalifCPP.value = 0 'RECO20160628 ERS002-2016
    'Catalogo JOEP
        cmbProducto.Enabled = False
        cmbSubProd.Enabled = False
        cmdCheckList.Enabled = False
    'Catalogo JOEP
    
End Sub

Private Sub CargarDatosCarga()
Dim oCred As COMDCredito.DCOMCredito
Set oCred = New COMDCredito.DCOMCredito
Dim rsCIUU As ADODB.Recordset
Dim nParamRCC As Double
Call oCred.CargarDatosSugerencia(gsCodCMAC, rsCIUU, nParamRCC)
Set oCred = Nothing

Call CargarCIUU(rsCIUU)
If nParamRCC = 1 Then
    bControlRCC = True
    'chkExpuestoRCC.Visible = True 'LUCV20180601, Comentó según ERS022-2018
Else
    bControlRCC = False
End If
End Sub

Private Sub CargarCIUU(ByVal pRs As ADODB.Recordset)
'Dim oCIUU As COMDPersona.DCOMPersonas
'Dim rsCIIU As ADODB.Recordset

'Set oCIUU = New COMDPersona.DCOMPersonas
'Set rsCIIU = oCIUU.Cargar_CIIU(gsCodCMAC)
'Set oCIUU = Nothing

'Do While Not rsCIIU.EOF
'    CboPersCiiu.AddItem Trim(rsCIIU!cCIIUdescripcion) & Space(100) & Trim(rsCIIU!cCIIUcod)
'    rsCIIU.MoveNext
'Loop
Do While Not pRs.EOF
    CboPersCiiu.AddItem Trim(pRs!cCIIUdescripcion) & Space(100) & Trim(pRs!cCIIUcod)
    pRs.MoveNext
Loop

End Sub

'Private Sub Optdesemb_Click(Index As Integer)
'    ReDim MatDesPar(0, 0)
'    ReDim MatCalend(0, 0)
'    ReDim MatDesPar(0, 0)
'    ReDim MatrizCal(0, 0)
'
'    If Index = 0 Then 'Si Desembolso Total
'        CmdDesembolsos.Enabled = False
'        txtMonSug.Enabled = True
'        txtMonSug.Text = "0.00"
'        spnCuotas.valor = 30
'        spnCuotas.Enabled = True
'    Else
'        CmdDesembolsos.Enabled = True
'        txtMonSug.Enabled = False
'        txtMonSug.Text = "0.00"
'        spnCuotas.valor = 1
'        spnCuotas.Enabled = False
'    End If
'End Sub

'Private Sub opttcuota_Click(Index As Integer)
'    ReDim MatCalend(0, 0)
'    ReDim MatrizCal(0, 0)
'    If Index <> 3 Then
'        DeshabilitaTipoPeriodo True, True
'        DeshabilitaTipoGracia True
'        DeshabilitaTipoCalend True
'        DeshabilitaTipoDesemb True
'        'If Optdesemb(1).value Then
'        If Trim(Right(cmbDesemb.Text, 2)) = "1" Then
'            CmdDesembolsos.Enabled = True
'        End If
'        ReDim MatrizCal(0, 0)
'    Else
'        DeshabilitaTipoPeriodo False, False
'        DeshabilitaTipoGracia False
'        DeshabilitaTipoPeriodo False, False
'        DeshabilitaTipoGracia False
'        DeshabilitaTipoCalend False
'        DeshabilitaTipoDesemb True
'        txtPerGra.Text = "0"
'        opttper(0).value = True
'    End If
'End Sub

'Private Sub opttper_Click(Index As Integer)
'    ReDim MatCalend(0, 0)
'    ReDim MatrizCal(0, 0)
'    If Index = 1 Then
'        HabilitaFechaFija (True)
'        'Comentado Por MAVM 25102010 ***
'        'If TxtDiaFijo.Enabled Then
'        '    TxtDiaFijo.SetFocus
'        'End If
'        '***
'        'No aplica para la fecha fija la Gracia en Cuotas
'        optTipoGracia(1).Enabled = False
'        optTipoGracia(1).value = False
'        'Activa el ingreso de fecha fija para el calculo de dias de gracia GITU 19-08-2008
'        'txtFechaFija.Enabled = True
'        'txtFechaFija.SetFocus
'
'        'MAVM 25102010 ***
'        txtFechaFija.Enabled = True
'        TxtDiaFijo.Enabled = False
'        TxtDiaFijo2.Enabled = False
'        chkGracia.value = 0
'        chkGracia.Enabled = False
'        txtFechaFija.Text = gdFecSis
'        'ALPA 20111209
'        'GenerarFechaPago
'        If (ActxCta.Prod = "515" Or ActxCta.Prod = "516") And nValorDiaGracia = 1 Then
'            txtFechaFija.Text = txtFechaFija.Text
'        Else
'            txtFechaFija.Text = gdFecSis
'            GenerarFechaPago
'        End If
'        '***********************
'    Else
'        HabilitaFechaFija (False)
'
'        'MAVM 25112010 ***
'        txtFechaFija.Text = "__/__/____"
'        GenerarFechaPago
'        SpnPlazo.SetFocus
'        chkGracia.value = 0
'        TxtDiaFijo.Text = "00"
'        chkGracia.Enabled = True
'        '***
'
'        'txtFechaFija.Text = "__/__/____"
'        'optTipoGracia(1).Enabled = True    'ARCV 30-04-2007
'    End If
'    'txtPerGra.Text = ""
'End Sub

Private Sub SpnCuotas_Change()
    'bGraciaGenerada = False 'LUCV20180601, Comentó según ERS022-2018
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
    bBuscarLineas = False
    ValidaCuotaBalon 'WIOR 20131115
    Call MostrarLineas 'ALPA 20140722
     '->***** LUCV20180601, Seg?n ERS022-2018
     Call CargaTipoPago 'JOEP20200217 Mejora pase de Canlendario interativo
'Comento JOEP20200217 Mejora pase de Canlendario interativo
'    If spnCuotas.valor = 1 Then
'        cmbTipoPer.Enabled = False
'        cmbTipoPer.ListIndex = 0
'    Else
'        cmbTipoPer.Enabled = True
'    End If
'Comento JOEP20200217 Mejora pase de Canlendario interativo
    '<-***** Fin LUCV20180601
End Sub

Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
'        If spnPlazo.Enabled Then
'            spnPlazo.SetFocus
'        End If
    '->***** LUCV20180601, Según ERS022-2018
    Call CargaTipoPago 'JOEP20200217 Mejora pase de Canlendario interativo
'Comento JOEP20200217 Mejora pase de Canlendario interativo
'    If spnCuotas.valor = 1 Then
'        cmbTipoPer.Enabled = False
'        cmbTipoPer.ListIndex = 0
'    Else
'        cmbTipoPer.Enabled = True
'    End If
'Comento JOEP20200217 Mejora pase de Canlendario interativo
    '<-***** Fin LUCV20180601
        SendKeys "{Tab}", True
     End If
End Sub

Private Sub SpnPlazo_Change()
    'bGraciaGenerada = False 'LUCV20180601, Comentó según ERS022-2018
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
    bBuscarLineas = False
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000029", ActxCta.Prod) Then     '**END ARLO
    'If (ActxCta.Prod = "515" Or ActxCta.Prod = "516") And nValorDiaGracia = 1 Then
        'txtFechaFija.Text = txtFechaFija.Text
    Else
        GenerarFechaPago 'MAVM 30092010
    End If
    chkGracia.value = 0 'MAVM 30092010
    Call MostrarLineas 'ALPA 20150113
End Sub
'ALPA 20141028***************************
'Private Sub txtInteresTasa_GotFocus()
'    fEnfoque txtInteresTasa
'End Sub

'Private Sub txtInteresTasa_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        KeyAscii = NumerosDecimales(txtInteresTasa, KeyAscii, , 4)
'        Txtinteres.Text = txtInteresTasa.Text
'        Call MostrarLineas
'     End If
'End Sub

'Private Sub txtInteresTasa_LostFocus()
'    If Trim(txtInteresTasa.Text) = "" Then
'        txtInteresTasa.Text = "0.0000"
'    Else
'        'ALPA 20141030************************************
'        If CDbl(txtInteresTasa.Text) >= lnTasaInicial Then
'            txtInteresTasa.Text = lnTasaInicial
'        End If
'        '*************************************************
'        txtInteresTasa.Text = Format(txtInteresTasa.Text, "#0.0000")
'        Txtinteres.Text = txtInteresTasa.Text
'        Call MostrarLineas
'    End If
'End Sub
'****************************************
Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        'opttcuota(0).SetFocus
        'cmbTipoCuota.SetFocus 'LUCV20180601, Comentó según ERS022-2018
        txtFechaFija.SetFocus 'LUCV20180601. Agregó según ERS022-2018
     End If
End Sub


Private Sub TxtDiaFijo_Change()
    'bGraciaGenerada = False 'LUCV20180601, Comentó según ERS022-2018
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
End Sub

Private Sub TxtDiaFijo_GotFocus()
    fEnfoque TxtDiaFijo
End Sub

Private Sub TxtDiaFijo_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        If ChkProxMes.Enabled Then ChkProxMes.SetFocus
     End If
End Sub

Private Sub TxtDiaFijo_LostFocus()
    If Trim(TxtDiaFijo.Text) = "" Then
        TxtDiaFijo.Text = "00"
        ChkProxMes.value = 0
    Else
        TxtDiaFijo.Text = Right("00" & Trim(TxtDiaFijo.Text), 2)
    End If
End Sub

Private Sub TxtDiaFijo2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtExpAntMax_GotFocus()
    fEnfoque txtExpAntMax
End Sub

Private Sub txtExpAntMax_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtExpAntMax, KeyAscii)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtExpAntMax_LostFocus()
    If Trim(txtExpAntMax.Text) = "" Then
        txtExpAntMax.Text = "0.00"
    Else
        txtExpAntMax.Text = Format(txtExpAntMax.Text, "#0.00")
    End If
End Sub

'MAVM 25102010 ***
Private Sub txtFechaFija_KeyPress(KeyAscii As Integer)
'JOEP20190114 CP
If Not CP_ValMensaje(1) Then
    Exit Sub
End If
'JOEP20190114 CP
    If KeyAscii = 13 Then
        If Not Trim(ValidaFecha(txtFechaFija.Text)) = "" Then
            MsgBox Trim(ValidaFecha(txtFechaFija.Text)), vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFechaFija.Text) < CDate(TxtFecDesemb.Text) Then
            MsgBox "Fecha tiene que ser mayor a la fecha del dia", vbInformation, "Atencion!"
            txtFechaFija.SetFocus
            txtFechaFija.Text = CDate(gdFecSis + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
            chkGracia.value = 0
            Exit Sub
        End If
         
'->***** LUCV20170915, Agregó y Comentó Según Modificaciones del reglamento (4.2.-Crédito Refinanciado. (b)) ->*****
      If Not ValidaPeriodoGracia Then
            txtFechaFija.SetFocus
            chkGracia.value = 0
            Exit Sub
       End If
    'JOEP20190114 CP
        If Not CP_ValidaRangoFecha(Trim(Right(cmbSubProd.Text, 9)), 3000, Right(cmbTipoPer.Text, 3), txtFechaFija.Text) Then
            txtFechaFija.SetFocus
            Exit Sub
        End If
        If Not CP_CondicionesSeg(Trim(Right(cmbSubProd.Text, 5)), Mid(ActxCta.NroCuenta, 9, 1), txtMonSug.Text, spnCuotas.valor, SpnPlazo.valor, Trim(Right(cmbDestino.Text, 5)), txtPerGra.Text, Right(cmbTipoCredito.Text, 5), 2) Then Exit Sub
    'JOEP20190114 CP
     '->***** LUCV20180601
      If fbTieneTasaGracia Then
        TxtTasaGracia.Text = Format(Txtinteres.Text, "#0.0000") 'LUCV20180601, según ERS022-2018
      End If
     '<-***** Fin LUCV20180601
    
'        If Trim(Right(cmbTipoPer.Text, 2)) = "0" Then
'
'            If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") Then
'                MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
'                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
'                Exit Sub
'            End If
'
'            If CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) > CDate(txtFechaFija.Text) Then
'                MsgBox "La Fecha de Pago No puede ser Menor que el Plazo", vbInformation, "Aviso"
'                txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text) + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
'                txtFechaFija.SetFocus
'                chkGracia.value = 0
'                Exit Sub
'            End If
'
'            If txtFechaFija.Text > CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) Then
'                chkGracia.Enabled = True
'                chkGracia.value = 1
'                txtPerGra.Text = CInt(CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor))
'            Else
'                txtPerGra.Text = "0"
'                chkGracia.value = 0
'            End If
'            'ALPA 20160419********************************************
'            If SpnPlazo.valor < 30 Then
'                MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            '*********************************************************
'        Else
'            If Month(gdFecSis) = Month(txtFechaFija.Text) And Year(CDate(TxtFecDesemb.Text)) = Year(txtFechaFija.Text) Then
'                ChkProxMes.value = 0
'            Else
'                ChkProxMes.value = 1
'            End If
'
'            If CDate(gdFecSis + 30) < txtFechaFija.Text Then
'                chkGracia.Enabled = True
'                chkGracia.value = 1
'                txtPerGra.Text = CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + 30)
'            Else
'                chkGracia.value = 0
'            End If
'
'            TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaFija.Text)), 2)
'        End If
'<-***** Fin LUCV20170915<-*****
    End If
End Sub
'***

'*****->LUCV20170925, Creó. Según Modificaciones del reglamento (4.2.-Crédito Refinanciado. (b))
Public Function ValidaPeriodoGracia() As Boolean
Dim nPeriodoGraciaRefinanciado As Integer
nPeriodoGraciaRefinanciado = 0
ValidaPeriodoGracia = True

    If Trim(Right(cmbTipoPer.Text, 2)) = "0" Then
        If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") Then
            MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
            If SpnPlazo.Enabled Then SpnPlazo.SetFocus
            ValidaPeriodoGracia = False
        End If
        
        If CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) > CDate(txtFechaFija.Text) Then
            MsgBox "La Fecha de Pago No puede ser Menor que el Plazo", vbInformation, "Aviso"
            txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text) + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
            txtFechaFija.SetFocus
            chkGracia.value = 0
            ValidaPeriodoGracia = False
        End If
        
        If txtFechaFija.Text > CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor) Then
            chkGracia.Enabled = True
            chkGracia.value = 1
            txtPerGra.Text = CInt(CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + SpnPlazo.valor))
            nPeriodoGraciaRefinanciado = txtPerGra.Text
        Else
            txtPerGra.Text = "0"
            chkGracia.value = 0
        End If
        
        If SpnPlazo.valor < 30 Then
            MsgBox "El crédito no cumple con las especificaciones de plazo, no debe ser menor a 30 días", vbInformation, "Aviso"
            ValidaPeriodoGracia = False
        End If

        'JAOR 20200728 COMENTÒ
        'If bEsRefinanciado And nPeriodoGraciaRefinanciado > 30 Then
        'MsgBox "Periodo de gracia no debe ser mayor a 30 días", vbInformation, "Aviso"
        'txtFechaFija.SetFocus
        'txtFechaFija.Text = CDate(gdFecSis + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
        ' chkGracia.value = 0
        '    ValidaPeriodoGracia = False
        'End If
    Else
        ValidaPeriodoGracia = True
        If Month(gdFecSis) = Month(txtFechaFija.Text) And Year(CDate(TxtFecDesemb.Text)) = Year(txtFechaFija.Text) Then
            ChkProxMes.value = 0
        Else
            ChkProxMes.value = 1
        End If

        If CDate(gdFecSis + 30) < txtFechaFija.Text Then
            chkGracia.Enabled = True
            chkGracia.value = 1
            txtPerGra.Text = CDate(txtFechaFija.Text) - CDate(CDate(TxtFecDesemb.Text) + 30)
            nPeriodoGraciaRefinanciado = txtPerGra.Text
        Else
            chkGracia.value = 0
        End If
        TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaFija.Text)), 2)

        'JAOR 20200728 COMENTÒ
        'If bEsRefinanciado And nPeriodoGraciaRefinanciado > 30 Then
        'MsgBox("Periodo de gracia no debe ser mayor a 30 días", vbInformation, "Aviso")
        ' txtFechaFija.Text = CDate(gdFecSis + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
        'chkGracia.value = 0
        ' ValidaPeriodoGracia = False
        'End If
    End If
End Function
'Fin LUCV20170925 <-*****



Private Sub txtFechaFija_LostFocus()
'Add by Gitu 20-08-2008
'Descomentar cuando esten seguros de los cambios GITU
'    If Not txtFechaFija.Text = "__/__/____" Then
'        If CDate(txtFechaFija.Text) < gdFecSis Then
'            MsgBox "Fecha tiene que ser mayor a la fecha del dia", vbInformation, "Atencion!"
'            txtFechaFija.SetFocus
'            Exit Sub
'        End If
'
'        If OptTPer(0).value = True Then
'            If Val(SpnPlazo.Valor) = 0 Then
'                MsgBox "Ingrese Plazo de las cuotas del prestamo", vbInformation, "Aviso!"
'                SpnPlazo.SetFocus
'                Exit Sub
'            End If
'            TxtPerGra.Text = (CDate(txtFechaFija.Text) - gdFecSis) - Val(SpnPlazo.Valor)
'        Else
'            TxtPerGra.Text = (CDate(txtFechaFija.Text) - gdFecSis) - 30
'        End If
'        If Val(TxtPerGra.Text) < 0 Then
'            TxtPerGra.Text = 0
'        End If
'        TxtDiaFijo.Text = Day(CDate(txtFechaFija.Text))
'    End If
'End Gitu
End Sub

Private Sub txtInteres_Change()
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
End Sub

Private Sub txtinteres_GotFocus()
    fEnfoque Txtinteres
End Sub

Private Sub txtinteres_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(Txtinteres, KeyAscii, , 4)
     If KeyAscii = 13 Then
        'If TxtMora.Enabled Or TxtMora.Visible Then
        '    TxtMora.SetFocus
        'End If
        If txtMonSug.Enabled Then
            txtMonSug.SetFocus
        End If
        TxtTasaGracia.Text = Format(Txtinteres.Text, "#0.0000") 'LUCV20180601, Comentó según ERS022-2018
        
     End If
End Sub

Private Sub txtinteres_LostFocus()
    If Trim(Txtinteres.Text) = "" Then
        Txtinteres.Text = "0.0000"
        txtInteresTasa.Text = "0.0000"
    Else
        Txtinteres.Text = Format(Txtinteres.Text, "#0.0000")
        txtInteresTasa.Text = Format(Txtinteres.Text, "#0.0000")
        TxtTasaGracia.Text = Format(Txtinteres.Text, "#0.0000") 'LUCV20180601, Comentó según ERS022-2018
    End If
End Sub

Private Sub txtMora_Change()
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
End Sub

Private Sub txtMora_GotFocus()
    fEnfoque TxtMora
End Sub

Private Sub txtMora_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(TxtMora, KeyAscii, , 4)
     If KeyAscii = 13 Then
        If txtMonSug.Enabled Then
            txtMonSug.SetFocus
        End If
     End If
End Sub

Private Sub txtMora_LostFocus()
    If Trim(TxtMora.Text) = "" Then
        TxtMora.Text = "0.0000"
    Else
        TxtMora.Text = Format(TxtMora.Text, "#0.0000")
    End If
End Sub

Private Sub txtmonsug_Change()
    'bGraciaGenerada = False 'LUCV20180601, Comentó según ERS022-2018
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
    bBuscarLineas = False
    'WIOR 20131115 *************************************************************
    If Len(Trim(ActxCta.NroCuenta)) = 18 Then
        Dim oNCredito As COMNCredito.NCOMCredito
        Dim nMontoSug As Double
        Set oNCredito = New COMNCredito.NCOMCredito
        
        If txtMonSug.Text = "" Or txtMonSug.Text = "." Then
            nMontoSug = 0
        Else
            nMontoSug = CDbl(txtMonSug.Text)
        End If
        If oNCredito.AplicaCuotaBalon(gsCodAge, sSTipoProdCod, nMontoSug, CInt(Mid(ActxCta.NroCuenta, 9, 1))) Then
            If Not (chkCuotaBalon.Visible And chkCuotaBalon.Enabled) Then
                chkCuotaBalon.Visible = True
                chkCuotaBalon.value = 0
                txtCuotaBalon.Visible = True
                txtCuotaBalon.Text = "0"
            End If
        Else
            txtCuotaBalon.Text = "0"
            chkCuotaBalon.Visible = False
            txtCuotaBalon.Visible = False
        End If
        Set oNCredito = Nothing
    End If
    'WIOR FIN ******************************************************************
    'ALPA 20150113**************************************************************
    Call CargarDatosProductoCrediticio
    Call MostrarLineas
    '***************************************************************************
End Sub

Private Sub txtmonsug_GotFocus()
'JOEP20190129 CP
If bEntrotxtMontoSeg = False Then
    Call CP_AporteSeg
End If
'JOEP20190129 CP
    fEnfoque txtMonSug
bEntrotxtMontoSeg = True 'JOEP20190129 CP
End Sub

Private Sub txtmonsug_KeyPress(KeyAscii As Integer)
     'KeyAscii = NumerosDecimales(txtMonSug, KeyAscii)
'JOEP20190129 CP
    Call CP_AporteSeg
    bEntrotxtMontoSeg = True
'JOEP20190129 CP
     If KeyAscii = 13 Then
        spnCuotas.SetFocus
     End If
End Sub

Private Sub txtmonsug_LostFocus()
    If Trim(txtMonSug.Text) = "" Then
        txtMonSug.Text = "0.00"
    Else
        txtMonSug.Text = Format(txtMonSug.Text, "#0.00")
    End If
bEntrotxtMontoSeg = False 'JOEP20190129 CP
End Sub

Private Sub TxtPerGra_Change()
    bGraciaGenerada = False
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
End Sub

Private Sub TxtPerGra_GotFocus()
    fEnfoque txtPerGra
End Sub

Private Sub TxtPerGra_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        If TxtTasaGracia.Visible Then
            TxtTasaGracia.SetFocus
        Else
            CmdCalend.SetFocus
        End If
     End If
End Sub

Private Sub txtPerGra_LostFocus()
    If Trim(txtPerGra.Text) = "" Then
        txtPerGra.Text = "0"
    Else
        txtPerGra.Text = Format(txtPerGra.Text, "#0")
    End If
End Sub
'->***** LUCV20180601, Comentó según ERS022-2018
'Private Sub TxtTasaGracia_Change()
'    bGraciaGenerada = False
'    ReDim MatCalend(0, 0)
'    ReDim MatrizCal(0, 0)
'End Sub
'
'Private Sub TxtTasaGracia_GotFocus()
'    fEnfoque TxtTasaGracia
'End Sub
'
'Private Sub TxtTasaGracia_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        cmdgracia.SetFocus
'     End If
'End Sub
'
'Private Sub TxtTasaGracia_LostFocus()
'    If Trim(TxtTasaGracia.Text) = "" Then
'        TxtTasaGracia.Text = "0"
'    Else
'        TxtTasaGracia.Text = Format(TxtTasaGracia.Text, "#0.0000")
'
'    End If
'End Sub
'<-***** Fin LUCV20180601

Private Sub cmbTipoCredito_Click()
    Call CargaSubTipoCredito(Trim(Right(cmbTipoCredito.Text, 3)))
    If Right(cmbTipoCredito.Text, 3) = gColCredCorpo Then
        lblInstitucionFinanciera.Visible = True
        cmbInstitucionFinanciera.Visible = True
    'JUEZ 20130913 **************************************************
        cmbDatoVivienda.Visible = False
        lblInstitucionFinanciera.Caption = "Inst. Corp."
    ElseIf Right(cmbTipoCredito.Text, 3) = gColCredHipot Then
        lblInstitucionFinanciera.Visible = True
        cmbInstitucionFinanciera.Visible = False
        cmbDatoVivienda.Visible = True
        lblInstitucionFinanciera.Caption = "Datos Viv."
    'END JUEZ *******************************************************
    Else
        lblInstitucionFinanciera.Visible = False
        cmbInstitucionFinanciera.Visible = False
        cmbDatoVivienda.Visible = False 'JUEZ 20130913
    End If
    Call VerificarMIVIVIENDA 'WIOR 20151223
End Sub

Private Sub CargaInstitucionesFinancieras(ByVal psTipo As String)
Dim oCons As COMDConstantes.DCOMConstantes
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaInstitucionesFinancieras
    Set oCons = New COMDConstantes.DCOMConstantes
    Set RTemp = oCons.RecuperaConstantes(psTipo)
    Set oCons = Nothing
    cmbInstitucionFinanciera.Clear
    Do While Not RTemp.EOF
        cmbInstitucionFinanciera.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbInstitucionFinanciera, 250)
    Exit Sub
    
ERRORCargaInstitucionesFinancieras:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub CargaSubTipoCredito(ByVal psTipo As String)
Dim oCred As COMDCredito.DCOMCredito
Dim ssql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaSubTipoCredito
    Set oCred = New COMDCredito.DCOMCredito
    'Set RTemp = oCred.RecuperaSubTipoCrediticios(psTipo, gsCodCargo)  'NAGL 20171121
    Set RTemp = oCred.RecuperaSubTipoCrediticios(psTipo, gsCodCargo, Trim(Right(cmbSubProd.Text, 9))) 'JOEP20190114 CP
    Set oCred = Nothing
    cmbSubTipo.Clear
    Do While Not RTemp.EOF
        cmbSubTipo.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbSubTipo, 250)
    Exit Sub
    
ERRORCargaSubTipoCredito:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

'MAVM 28102010 ***
Private Sub GenerarFechaPago()
    'If opttper(0).value = True Then
    If Trim(Right(cmbTipoPer.Text, 2)) = "0" Then
        txtFechaFija.Text = CDate(CDate(TxtFecDesemb.Text) + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
    End If
    If opttper(1).value = True Then
        txtFechaFija.Text = CDate(TxtFecDesemb.Text)
        TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaFija.Text)), 2)
                                    
        If Month(gdFecSis) = Month(CDate(txtFechaFija.Text)) And Year(CDate(TxtFecDesemb.Text)) = Year(txtFechaFija.Text) Then
            ChkProxMes.value = 0
        Else
            ChkProxMes.value = 1
        End If
    End If
End Sub
'***

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim ssql As String, sMoneda As String, sPersona As String
Dim nRow As Long

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing

If Not (rsCta.EOF And rsCta.BOF) Then
    CmdCalend.Enabled = True
    If txtMonSug.Enabled Then
        txtMonSug.SetFocus
    Else
        If spnCuotas.Enabled = True Then
            spnCuotas.SetFocus
        Else
            SpnPlazo.SetFocus
        End If
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    lblTercero.Caption = ""
    actxCtaAho.SetFocusCuenta
End If
End Sub

Private Function ValidaCuenta(ByVal psCtaCod As String) As Boolean
    If Len(Trim(psCtaCod)) <> 18 Then
        MsgBox "Ingrese un Numero de Cuenta correcta", vbInformation, "Aviso"
        ValidaCuenta = False
    Else
        ValidaCuenta = True
    End If
End Function

'JUEZ 20130913 ********************************************************
Private Sub CargaDatoVivienda()
Dim oCons As COMDConstantes.DCOMConstantes
Dim ssql As String
Dim RTemp As ADODB.Recordset
    Set oCons = New COMDConstantes.DCOMConstantes
    Set RTemp = oCons.RecuperaConstantes(3040)
    Set oCons = Nothing
    cmbDatoVivienda.Clear
    Do While Not RTemp.EOF
        cmbDatoVivienda.AddItem RTemp!cConsDescripcion & Space(500) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Exit Sub
End Sub
'END JUEZ *************************************************************

'WIOR 20131115 ********************************************************
Private Sub ValidaCuotaBalon()
Dim valor As Integer
Dim valorCB As Integer

If txtCuotaBalon.Visible And chkCuotaBalon.Visible Then
    If chkCuotaBalon.value = 0 Then Exit Sub
    
    If CInt(spnCuotas.valor) < 2 Then
        chkCuotaBalon.value = 0
        txtCuotaBalon.Text = "0"
        Exit Sub
    End If
    
    If spnCuotas.valor = 0 Or spnCuotas.valor = "" Then
        valor = 0
    Else
        valor = CInt(spnCuotas.valor) - 1
    End If
    
    If txtCuotaBalon.Text = "0" Or txtCuotaBalon.Text = "" Then
        valorCB = 0
    Else
        valorCB = CInt(txtCuotaBalon.Text)
    End If
    
    
    If valor < valorCB Then
        txtCuotaBalon.Text = valor
    End If
End If

End Sub
'WIOR FIN *************************************************************


Private Sub txtMontoMivivienda_Change()
    Dim nMontoMV As Double
    If Len(Trim(ActxCta.NroCuenta)) = 18 Then
        
        If txtMontoMivivienda.Text = "" Or txtMontoMivivienda.Text = "." Then
            nMontoMV = 0
        Else
            If IsNumeric(txtMontoMivivienda.Text) Then
                nMontoMV = CDbl(txtMontoMivivienda.Text)
            Else
                nMontoMV = 0
            End If
        End If
    End If
End Sub

Private Sub txtMontoMivivienda_GotFocus()
    fEnfoque txtMontoMivivienda
End Sub

Private Sub txtMontoMivivienda_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtMontoMivivienda, KeyAscii)
End Sub

Private Sub txtMontoMivivienda_LostFocus()
    If Trim(txtMontoMivivienda.Text) = "" Then
        txtMontoMivivienda.Text = "0.00"
    Else
        txtMontoMivivienda.Text = Format(txtMontoMivivienda.Text, "#0.00")
    End If
End Sub

'ALPA 20141113***********************************************
'Private Function ValidarVerEntidades() As String
'        Dim oCredito As COMDCredito.DCOMCredito
'        Set oCredito = New COMDCredito.DCOMCredito
'        Dim nCantidadVerEntidades As Integer
'        Dim nCantidadCreditos As Integer
'        Dim nPorMorAgenc As Double
'        Dim nExcedenteFI As Double
'        Dim nPorExcedenteFI As Double
'
'        Dim nCalificacionNormal As Integer
'        ValidarVerEntidades = ""
'        If nLogicoVerEntidades = 0 Then
'            MsgBox "Tiene que verificar la cantidad de entidades en <<Ver entidades>>"
'            ValidarVerEntidades = "XXX"
'            Exit Function
'        End If
'        If Not (oRsVerEntidades.BOF And oRsVerEntidades.EOF) Then
'            nCantidadVerEntidades = 0
'            bCantidadVerEntidadesCmac = 0
'            lnCantidadVerEntidades = 0
'            oRsVerEntidades.MoveFirst
'            Do While Not oRsVerEntidades.EOF
'                If oRsVerEntidades!bAnulacion = 0 Then
'                    lnCantidadVerEntidades = lnCantidadVerEntidades + 1
'                    If oRsVerEntidades!codigo = 109 Then
'                    bCantidadVerEntidadesCmac = bCantidadVerEntidadesCmac + 1
'                    End If
'                End If
'                oRsVerEntidades.MoveNext
'            Loop
'        End If
'        If lnCantidadVerEntidades = 0 Then
'            lnCantidadVerEntidades = 1
'        ElseIf bCantidadVerEntidadesCmac = 0 Then
'            lnCantidadVerEntidades = lnCantidadVerEntidades + 1
'        End If
'        If bCantidadVerEntidadesCmac = 0 Then
'            bCantidadVerEntidadesCmac = 1
'        End If
'
'        nCantidadCreditos = oCredito.CantidadCreditosVerEntidades(Trim(Me.ActxCta.NroCuenta)) + IIf(lnColocCondicion = 4 Or lnColocCondicion = 5, 0, 1)
'        'If Not ((lnCantidadVerEntidades <= 4 And nCantidadCreditos <= 2) Or ((lnCantidadVerEntidades >= 2 And lnCantidadVerEntidades <= 3) And nCantidadCreditos <= 3) Or (lnCantidadVerEntidades = 1 And bCantidadVerEntidadesCmac = 0 And nCantidadCreditos <= 3) Or (bCantidadVerEntidadesCmac = lnCantidadVerEntidades And bCantidadVerEntidadesCmac >= 1 And nCantidadCreditos <= 4)) Then
'        If Not ((lnCantidadVerEntidades <= 4 And nCantidadCreditos <= 2) Or _
'                (lnCantidadVerEntidades = 4 And nCantidadCreditos = 2) Or _
'               ((lnCantidadVerEntidades >= 2 And lnCantidadVerEntidades <= 3) And nCantidadCreditos <= 3) Or _
'               (lnCantidadVerEntidades = 1 And bCantidadVerEntidadesCmac = 0 And nCantidadCreditos <= 3) Or _
'               (bCantidadVerEntidadesCmac = lnCantidadVerEntidades And bCantidadVerEntidadesCmac >= 1 And nCantidadCreditos <= 4)) Then
'            MsgBox "Cantidad de créditos no esta permitida, el cliente tiene actualmente " & nCantidadCreditos & " créditos en la caja y debe a " & lnCantidadVerEntidades & " entidad(es) financiera(s)", vbInformation, "¡Aviso!"
'            ValidarVerEntidades = "XXX"
'            Exit Function
'        End If
'        nPorMorAgenc = oCredito.ObtenerPorcentajeMoraAgencia(gdFecSis, ActxCta.Age)
'        nExcedenteFI = oCredito.ObtenerExcedenteFI(MatFuentesF(1, 1))
'
'        If nExcedenteFI <= 0 Then
'            MsgBox "El excente no es valido, favor editar la fuente de ingreso", vbInformation, "¡Aviso!"
'            ValidarVerEntidades = "XXX"
'            Exit Function
'        End If
'        'nPorExcedenteFI = Round(lblcuota / nExcedenteFI, 2) * 100
'        If nExcedenteFI = 0 Then
'            nPorExcedenteFI = 0
'        Else
'            nPorExcedenteFI = Round(lblcuota / nExcedenteFI, 2) * 100
'        End If
'        'If (nPorExcedenteFI <= 0) Or Not ((lnCantidadVerEntidades = 0) Or (lnCantidadVerEntidades = 4 And ((nPorExcedenteFI <= 60 Or nPorExcedenteFI <= 50))) Or ((lnCantidadVerEntidades >= 2 And lnCantidadVerEntidades <= 3) And (nPorExcedenteFI <= 65 Or nPorExcedenteFI <= 60)) Or (lnCantidadVerEntidades = 1 And bCantidadVerEntidadesCmac = 0 And (nPorExcedenteFI <= 80 Or nPorExcedenteFI <= 70)) Or (lnCantidadVerEntidades = 1 And bCantidadVerEntidadesCmac = 1 And (nPorExcedenteFI <= 80 Or nPorExcedenteFI <= 70))) Then
'        If (nPorExcedenteFI <= 0) Or Not ((lnCantidadVerEntidades = 0) Or _
'                                          (lnCantidadVerEntidades = 4 And ((nPorExcedenteFI <= 60 Or nPorExcedenteFI <= 50))) Or _
'                                          ((lnCantidadVerEntidades >= 2 And lnCantidadVerEntidades <= 3) And (nPorExcedenteFI <= 65 Or nPorExcedenteFI <= 60)) Or _
'                                          (lnCantidadVerEntidades = 1 And bCantidadVerEntidadesCmac = 0 And (nPorExcedenteFI <= 80 Or nPorExcedenteFI <= 70)) Or _
'                                          (bCantidadVerEntidadesCmac = lnCantidadVerEntidades And bCantidadVerEntidadesCmac >= 1 And (nPorExcedenteFI <= 80 Or nPorExcedenteFI <= 70))) Then
'            MsgBox "Excedente no esta Permitido, el porcentaje del excendete es " & nPorExcedenteFI & "% ", vbInformation, "¡Aviso!"
'            ValidarVerEntidades = "XXX"
'            Exit Function
'        End If
'        'If (nPorExcedenteFI <= 0) Or Not (((nPorExcedenteFI <= 60 Or nPorExcedenteFI <= 65 Or nPorExcedenteFI <= 80) And nPorMorAgenc <= 4.5) Or ((nPorExcedenteFI <= 50 Or nPorExcedenteFI <= 60 Or nPorExcedenteFI <= 70) And nPorMorAgenc > 4.5)) Then
'         If (nPorExcedenteFI <= 0) Or _
'                Not (((nPorExcedenteFI <= 60 Or nPorExcedenteFI <= 65 Or nPorExcedenteFI <= 80) And nPorMorAgenc <= 4.5) _
'                            Or ((nPorExcedenteFI <= 50 Or nPorExcedenteFI <= 60 Or nPorExcedenteFI <= 70) And nPorMorAgenc > 4.5)) Then
'            MsgBox "La mora y el excedente son: " & nPorMorAgenc & "% y " & nPorExcedenteFI & "% respectivamente", vbInformation, "¡Aviso!"
'            ValidarVerEntidades = "XXX"
'            Exit Function
'        End If
'
'        If Right(Trim(CmbDestino.Text), 5) = 14 Then
'            If (lnCantidadVerEntidades > 3) Then
'                MsgBox "No se puede tener mas de 3 deudas en el Sistemas Financiero, actualmente tiene " & lnCantidadVerEntidades, vbInformation, "¡Aviso!"
'                ValidarVerEntidades = "XXX"
'                Exit Function
'            End If
'            nCalificacionNormal = oCredito.ObtenerCalificacionNormal(IIf(Len(Trim(LblDni.Caption)) = 0, LblRuc.Caption, LblDni.Caption))
'            If Not (nCalificacionNormal = 100 Or nCalificacionNormal = -1) Then
'                MsgBox "La calificación Normal debe ser 100% ", vbInformation, "¡Aviso!"
'                ValidarVerEntidades = "XXX"
'                Exit Function
'            End If
'            If CInt(txtPerGra.Text) > 30 Then
'                MsgBox "La gracia no está permitida ", vbInformation, "¡Aviso!"
'                ValidarVerEntidades = "XXX"
'                Exit Function
'            End If
'        End If
'End Function
 '*********************************************************************************
'WIOR 20151223 ***
Private Sub VerificarMIVIVIENDA()
Dim oDCredito As COMDCredito.DCOMCredito
Dim sTpoProd As String
Dim sTpoCred As String
Set oDCredito = New COMDCredito.DCOMCredito

sTpoProd = ""
If (cmbSubProd.ListIndex = -1) Then
    sTpoProd = ""
Else
    sTpoProd = Trim(Right(cmbSubProd.Text, 4))
End If

sTpoCred = ""
If (cmbSubTipo.ListIndex = -1) Then
    sTpoCred = ""
Else
    sTpoCred = Trim(Right(cmbSubTipo.Text, 4))
End If

fbMIVIVIENDA = oDCredito.EsCredMIVIVENDA(sTpoProd, sTpoCred, 2)
'txtMonSug.Enabled = True'comento JOEP20180204 CP
'JOEP20180204 CP
If lnColocCondicion = 4 Then
    txtMonSug.Enabled = False
Else
    txtMonSug.Enabled = True
End If
'JOEP20180204 CP
FraTpoCalend.Enabled = False
fratipodes.Enabled = False 'LUCV20180601, Según ERS022-2018
cmbDesemb.Enabled = False 'LUCV20180601, Según ERS022-2018
ChkMiViv.Enabled = False
ChkMiViv.value = 0
cmdMIVIVIENDA.Enabled = False

If fbDatosCargados Then
    If fbMIVIVIENDA Then
        Call frmCredMiViviendaDatos.Inicio(Trim(ActxCta.NroCuenta), fArrMIVIVIENDA)
        If IsArray(fArrMIVIVIENDA) Then
            txtMonSug.Text = CDbl(fArrMIVIVIENDA(3))
        Else
            txtMonSug.Text = "0.00"
        End If
        Call txtmonsug_Change
        Call txtmonsug_KeyPress(13)
        Call txtmonsug_LostFocus
        txtMonSug.Enabled = False
        FraTpoCalend.Enabled = True
        fratipodes.Enabled = True 'LUCV20180601, Según ERS022-2018
        cmbDesemb.Enabled = True 'LUCV20180601, Según ERS022-2018
        ChkMiViv.Enabled = False
        ChkMiViv.value = 1
        cmdMIVIVIENDA.Enabled = True
    End If
End If
End Sub

Private Sub cmdMIVIVIENDA_Click()
Call VerificarMIVIVIENDA
End Sub
'WIOR FIN ********
'RECO20160628 ERS002-2016 *********************************
Private Sub chkAutoCalifCPP_Click()
    If chkAutoCalifCPP.value = 1 Then
        fbAutoCalfCPP = True
    Else
        fbAutoCalfCPP = False
    End If
End Sub
'RECO FIN *************************************************
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

'**ARLO20171127 INICIO ERS070 - 2017
Private Sub feDeudaComprar_DblClick()
    
    Dim frm As frmCredCompraDeudaDet
    Dim lvCompraDeuda As TCompraDeuda
    Dim lvTemp() As TCompraDeuda
    Dim bOK As Boolean
    Dim Index As Integer
    
    If feDeudaComprar.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = feDeudaComprar.row
    lvTemp = fvListaCompraDeuda 'Temporal para no modificar el actual array
    
    lvCompraDeuda = fvListaCompraDeuda(Index)
    Set frm = New frmCredCompraDeudaDet
    
    bOK = frm.Modificar(lvCompraDeuda, Index, lvTemp, IIf(Trim(Right(Me.cmbSubProd.Text, 5)) = "", 0, CInt(Trim(Right(Me.cmbSubProd.Text, 5))))) 'ARLO20180319
'    If bOK Then
'        fvListaCompraDeuda(Index) = lvCompraDeuda
'        ModificaFila Index, lvCompraDeuda
'    End If
    Set frm = Nothing
End Sub
'**ARLO20171127 FIN ERS070 - 2017

'Catalogo JOEP
Private Sub CatalogoCombos(ByVal pnSubProd As Integer, ByVal pnParCod As Integer)
Dim objCatalogoComb As COMDCredito.DCOMCredito
Dim rsCatalogoComb As ADODB.Recordset
Dim i As Integer
Set objCatalogoComb = New COMDCredito.DCOMCredito
Set rsCatalogoComb = objCatalogoComb.getCatalogoCombo(pnSubProd, pnParCod)
    If Not (rsCatalogoComb.BOF And rsCatalogoComb.EOF) Then
        Select Case pnParCod
            Case 1000
                Call Llenar_Combo_con_Recordset(rsCatalogoComb, cmbTipoCredito)
            Case 2000
                Call Llenar_Combo_con_Recordset(rsCatalogoComb, cmbDestino)
            Case 3000
                cmbTipoPer.Clear
                For i = 1 To rsCatalogoComb.RecordCount
                    cmbTipoPer.AddItem rsCatalogoComb!cConsDescripcion & Space(100) & rsCatalogoComb!nConsValor
                    cmbTipoPer.ItemData(cmbTipoPer.NewIndex) = "" & rsCatalogoComb!nConsValor
                    rsCatalogoComb.MoveNext
                Next i
        End Select
    Else
        Select Case pnParCod
            Case 1000
                cmbTipoCredito.Clear
            Case 2000
                cmbDestino.Clear
            Case 3000
                cmbTipoPer.Clear
        End Select
    End If
End Sub

Private Function CP_ValidaRangoFecha(ByVal nCodProd As String, ByVal nParCod As Integer, ByVal nTpPago As Integer, ByVal dFechaPago As String) As Boolean
Dim obj As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
On Error GoTo ErrorCP_ValidaRangoFecha
CP_ValidaRangoFecha = True
    Set obj = New COMDCredito.DCOMCredito
    Set rs = obj.getCatalogoRangoFecha(nCodProd, nParCod, nTpPago, dFechaPago)
        
    If Not (rs.BOF And rs.EOF) Then
        If rs!mensaje <> "" Then
            MsgBox rs!mensaje, vbInformation, "AVISO"
            CP_ValidaRangoFecha = False
        End If
    End If
Set obj = Nothing
RSClose rs
    Exit Function
ErrorCP_ValidaRangoFecha:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Function CP_ValMensaje(ByVal pnTpMsg As Integer) As Boolean
    CP_ValMensaje = True
If pnTpMsg = 1 Then
    If cmbTipoPer.Text = "" Then
        MsgBox "Seleccione el Tipo de Periodo", vbInformation, "Aviso"
        If cmbTipoPer.Enabled = True Then
            cmbTipoPer.SetFocus
        End If
        CP_ValMensaje = False
    End If
ElseIf pnTpMsg = 2 Then
    If cmbTipoCredito.Text = "" Then
        MsgBox "Seleccione el Tipo de Crédito", vbInformation, "Aviso"
        CP_ValMensaje = False
        If cmbTipoCredito.Enabled = True Then
            cmbTipoCredito.SetFocus
        End If
    End If
End If
End Function

Private Function CP_AporteSeg() As Boolean
Dim objAporte As COMDCredito.DCOMCredito
Dim rsAporte As ADODB.Recordset
Set objAporte = New COMDCredito.DCOMCredito

CP_AporteSeg = True

Set rsAporte = objAporte.CargaAporte(Trim(ActxCta.NroCuenta))
If Not (rsAporte.BOF And rsAporte.EOF) Then

'If nMatAporteLleno = False Then
'    ReDim nMatMontoPre(1, 3)
'    nMatMontoPre(1, 1) = rsAporte!nMonto
'    nMatMontoPre(1, 2) = rsAporte!nAporte
'    nMatMontoPre(1, 3) = Format((rsAporte!nMonto - (Format(rsAporte!nMonto, "#,#00.00") * (Format(rsAporte!nAporte, "#0") / 100))), "#,#00.00")
'End If
        
    If IsArray(nMatMontoPre) Then
        If UBound(nMatMontoPre) > 0 Then
            frmCredMontoPresupuestado.Inicio rsAporte!nAporte, nMatMontoPre, Right(cmbDestino.Text, 9), Trim(Right(cmbSubProd.Text, 9)), MontoMin, 0, Mid(Trim(ActxCta.NroCuenta), 9, 1)
                If IsArray(nMatMontoPre) Then
                        txtMonSug.Text = Format(nMatMontoPre(1, 3), "0.00")
                        CP_AporteSeg = False
                        nMatAporteLleno = True
                Else
                    txtMonSug.Text = "0.00"
                    nMatAporteLleno = False
                End If
        Else
            nMatAporteLleno = False
        End If
    Else
        nMatAporteLleno = False
    End If
Else
    Set nMatMontoPre = Nothing
    CP_AporteSeg = True
    nMatAporteLleno = False
End If
Set objAporte = Nothing
RSClose rsAporte
End Function

Private Sub bValiExistCheckList()
Dim rsActivar As ADODB.Recordset
Dim pnActivar As Integer
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito
Set rsActivar = oCredito.getActivarCheckList(Right(cmbSubProd.Text, 9))
If Not (rsActivar.BOF And rsActivar.EOF) Then
    If rsActivar!Activar = 1 Then
        cmdCheckList.Enabled = True
    Else
        cmdCheckList.Enabled = False
    End If
End If
Set oCredito = Nothing
RSClose rsActivar
End Sub

Private Function CP_CondicionesSeg(ByVal cCodProd As String, ByVal nMonedad As String, ByVal pnMonto As String, _
ByVal pnCuota As String, ByVal pnPlazo As String, ByVal pnDestino As String, ByVal pnPeriodo As String, ByVal pnTpCredito As String, ByVal pnModulo As Integer) As Boolean
    Dim oDCred As COMDCredito.DCOMCredito
    Dim rsDefaut As ADODB.Recordset
    Set oDCred = New COMDCredito.DCOMCredito
    CP_CondicionesSeg = True
    lnCampanaId = CInt(Trim(Right(Me.cmbCampana.Text, 3))) 'arlo20200429
    
    Set rsDefaut = oDCred.CatalogoCondicionSeg(Trim(ActxCta.NroCuenta), cCodProd, nMonedad, pnMonto, pnCuota, pnPlazo, pnDestino, pnPeriodo, pnTpCredito, pnModulo, lnCampanaId) 'arlo20200429 add lnCampanaId
        
    If Not (rsDefaut.BOF And rsDefaut.EOF) Then
        If rsDefaut!mensaje <> "" Then
            MsgBox rsDefaut!mensaje & rsDefaut!ValorMax, vbInformation, "Aviso"
            CP_CondicionesSeg = False
        End If
    End If
End Function
'Catalogo JOEP
'LUCV20180601, Según ERS022-2018
Private Sub TxtComenta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
End Sub

Private Sub cmbCampana_Click() 'Agregó evento para el cambio de campaña. Recomendación JOEP
    Call MostrarLineas
    If Trim(Right(Me.cmbCampana.Text, 3)) <> "" Then
        lnCampanaId = CInt(Trim(Right(Me.cmbCampana.Text, 3)))
    End If
    Call CargarDatosProductoCrediticio
    If lnTasaFinal = 0 Then
       MsgBox "No existen valores de las Tasas de Intereses del Producto." _
              & Chr(10) & "Coordinar la configuración de Tasas con el Dpto. responsable.", vbInformation, "Aviso"
    End If
End Sub
'Fin LUCV20180601

'20200217 Mejora pase de Canlendario interativo
Public Sub CargaTipoPago()
Dim i As Integer
    If spnCuotas.valor = 1 Then
        cmbTipoPer.Enabled = False
        For i = 0 To cmbTipoPer.ListCount
            If Right(cmbTipoPer.List(i), 1) = 0 Then
                cmbTipoPer.ListIndex = i
            End If
        Next i
    Else
        cmbTipoPer.Enabled = True
    End If
End Sub
'20200217 Mejora pase de Canlendario interativo
