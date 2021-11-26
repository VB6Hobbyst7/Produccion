VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormatoSinConvenio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Créditos - Evaluación - Formato Mi Cash Personal."
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   Icon            =   "frmCredFormEvalFormatoSinConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   1935
      Left            =   9315
      TabIndex        =   89
      Top             =   1440
      Width           =   1930
      Begin VB.CommandButton cmdVerCar 
         Caption         =   "&Ver CAR"
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
         Left            =   120
         TabIndex        =   93
         Top             =   1400
         Width           =   1700
      End
      Begin VB.CommandButton cmdInformeVista 
         Caption         =   "&Informe de Visita"
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
         Left            =   120
         TabIndex        =   92
         Top             =   960
         Width           =   1700
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Hoja Evaluación"
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
         Left            =   120
         TabIndex        =   91
         Top             =   520
         Width           =   1700
      End
      Begin VB.CommandButton cmdMNME 
         Caption         =   "MN - ME"
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
         Left            =   120
         TabIndex        =   90
         Top             =   150
         Visible         =   0   'False
         Width           =   1700
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   9315
      TabIndex        =   85
      Top             =   240
      Width           =   1930
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   345
         Left            =   120
         TabIndex        =   88
         Top             =   720
         Width           =   1700
      End
      Begin VB.CommandButton cmdActualizarSinConvenio 
         Caption         =   "&Guardar"
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
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
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
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   1700
      End
   End
   Begin TabDlg.SSTab SSTabInfoNego 
      Height          =   3900
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   9200
      _ExtentX        =   16245
      _ExtentY        =   6879
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Información del Cliente"
      TabPicture(0)   =   "frmCredFormEvalFormatoSinConvenio.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFechaIngreso"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ActXCodCtaSinConv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGiro"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmCredFormEvalFormatoSinConvenio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame frmCredFormEvalFormatoSinConvenio 
         Caption         =   "Datos del Empleador"
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   2950
         Width           =   9015
         Begin VB.TextBox txtNombreEmpleador 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2520
            TabIndex        =   42
            Top             =   360
            Width           =   6375
         End
         Begin SICMACT.TxtBuscar txtCodPers 
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
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
            EnabledText     =   0   'False
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2000
         Left            =   120
         TabIndex        =   19
         Top             =   880
         Width           =   9015
         Begin VB.OptionButton optTipoAportacion 
            Caption         =   "N/A"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   3
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optTipoAportacion 
            Caption         =   "ONP"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   2
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox txtNDependientes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2220
            TabIndex        =   4
            Top             =   1620
            Width           =   615
         End
         Begin VB.OptionButton optTipoAportacion 
            Caption         =   "AFP"
            Height          =   255
            Index           =   1
            Left            =   2220
            TabIndex        =   1
            Top             =   945
            Width           =   700
         End
         Begin VB.TextBox txtCliente 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2220
            TabIndex        =   20
            Top             =   160
            Width           =   6680
         End
         Begin MSMask.MaskEdBox txtFechaDeuda 
            Height          =   300
            Left            =   7750
            TabIndex        =   21
            Top             =   915
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin SICMACT.EditMoney txtCuota 
            Height          =   300
            Left            =   4500
            TabIndex        =   22
            Top             =   1240
            Width           =   735
            _ExtentX        =   1296
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
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
         End
         Begin Spinner.uSpinner spnAno 
            Height          =   315
            Left            =   2220
            TabIndex        =   23
            Top             =   555
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
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
            ForeColor       =   8421504
         End
         Begin Spinner.uSpinner spnMes 
            Height          =   315
            Left            =   3480
            TabIndex        =   24
            Top             =   555
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
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
            ForeColor       =   8421504
         End
         Begin SICMACT.EditMoney txtMonSolicitado 
            Height          =   300
            Left            =   2220
            TabIndex        =   25
            Top             =   1240
            Width           =   1215
            _ExtentX        =   2143
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
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtExpCredito 
            Height          =   300
            Left            =   7750
            TabIndex        =   26
            Top             =   1320
            Width           =   1125
            _ExtentX        =   1984
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
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
         End
         Begin SICMACT.EditMoney txtUltDeuda 
            Height          =   300
            Left            =   7750
            TabIndex        =   27
            Top             =   555
            Width           =   1125
            _ExtentX        =   1984
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
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. De dependientes:"
            Height          =   195
            Left            =   600
            TabIndex        =   40
            Top             =   1660
            Width           =   1605
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha último endeudamiento RCC:"
            Height          =   195
            Left            =   5280
            TabIndex        =   37
            Top             =   960
            Width           =   2460
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último endeudamiento RCC:"
            Height          =   195
            Left            =   5760
            TabIndex        =   36
            Top             =   600
            Width           =   1995
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4260
            TabIndex        =   35
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   2980
            TabIndex        =   34
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Solicitado:"
            Height          =   195
            Left            =   950
            TabIndex        =   33
            Top             =   1300
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Aportación:"
            Height          =   195
            Left            =   800
            TabIndex        =   32
            Top             =   940
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Antigüedad en actual Empleo:"
            Height          =   195
            Left            =   80
            TabIndex        =   31
            Top             =   580
            Width           =   2130
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   1680
            TabIndex        =   30
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exposición con este Crédito:"
            Height          =   195
            Left            =   5760
            TabIndex        =   29
            Top             =   1320
            Width           =   2010
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Cuotas :"
            Height          =   195
            Left            =   3720
            TabIndex        =   28
            Top             =   1320
            Width           =   810
         End
      End
      Begin VB.TextBox txtGiro 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4640
         TabIndex        =   18
         Top             =   555
         Width           =   4455
      End
      Begin SICMACT.ActXCodCta ActXCodCtaSinConv 
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito"
      End
      Begin MSMask.MaskEdBox txtFechaIngreso 
         Height          =   300
         Left            =   7900
         TabIndex        =   83
         Top             =   0
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Evaluación al:"
         Height          =   195
         Left            =   6120
         TabIndex        =   84
         Top             =   60
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actividad:"
         Height          =   195
         Left            =   3900
         TabIndex        =   39
         Top             =   600
         Width           =   705
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5500
      Left            =   0
      TabIndex        =   43
      Top             =   3930
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9710
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   6579
      BackColor       =   0
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ingresos y Egresos"
      TabPicture(0)   =   "frmCredFormEvalFormatoSinConvenio.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Propuesta del Crédito"
      TabPicture(1)   =   "frmCredFormEvalFormatoSinConvenio.frx":0342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Comentarios y Referidos"
      TabPicture(2)   =   "frmCredFormEvalFormatoSinConvenio.frx":035E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame9 
         Caption         =   "Comentarios"
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
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   11055
         Begin VB.TextBox txtComentario 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   10815
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Referidos"
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
         Height          =   2775
         Left            =   -74880
         TabIndex        =   68
         Top             =   2520
         Width           =   11055
         Begin VB.CommandButton cmdQuitar2 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1320
            TabIndex        =   70
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   2280
            Width           =   1095
         End
         Begin SICMACT.FlexEdit feReferidos 
            Height          =   1695
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   10800
            _ExtentX        =   19050
            _ExtentY        =   2990
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   1
            EncabezadosNombres=   "N°-Nombre-DNI-Telef.-Comentarios-DNI-Aux"
            EncabezadosAnchos=   "400-3800-1100-1100-4000-0-0"
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
            EncabezadosAlineacion=   "R-L-L-L-L-L-C"
            FormatosEdit    =   "3-1-0-0-1-0-3"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   0
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            TipoBusPersona  =   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Propuesta del Credito"
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
         Height          =   5145
         Left            =   -74880
         TabIndex        =   8
         Top             =   320
         Width           =   11175
         Begin VB.TextBox txtColaGarantias 
            Height          =   550
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   3740
            Width           =   10815
         End
         Begin VB.TextBox txtEntornoCliente 
            Height          =   550
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   520
            Width           =   10815
         End
         Begin VB.TextBox txtGiroNegocio 
            Height          =   550
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   1350
            Width           =   10815
         End
         Begin VB.TextBox txtExpCrediticia 
            Height          =   550
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   2150
            Width           =   10815
         End
         Begin VB.TextBox txtFormNegocio 
            Height          =   550
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   2950
            Width           =   10815
         End
         Begin VB.TextBox txtImpactoMismo 
            Height          =   550
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   4500
            Width           =   10815
         End
         Begin MSMask.MaskEdBox txtdFechaVisita 
            Height          =   345
            Left            =   9720
            TabIndex        =   9
            Top             =   170
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label45 
            Caption         =   "Fecha de Visita:"
            Height          =   300
            Left            =   8520
            TabIndex        =   67
            Top             =   170
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Sobre el Entorno Familiar del Cliente o Representante"
            Height          =   300
            Left            =   240
            TabIndex        =   66
            Top             =   340
            Width           =   4695
         End
         Begin VB.Label Label43 
            Caption         =   "Sobre la Actividad y o Giro, y la Ubicación del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   65
            Top             =   1150
            Width           =   4095
         End
         Begin VB.Label Label42 
            Caption         =   "Sobre la Experiencia Crediticia"
            Height          =   300
            Left            =   240
            TabIndex        =   64
            Top             =   1950
            Width           =   4215
         End
         Begin VB.Label Label41 
            Caption         =   "Sobre la Consistencia de la Información y la Formalidad del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   63
            Top             =   2750
            Width           =   6255
         End
         Begin VB.Label Label40 
            Caption         =   "Sobre los Colaterales o Garantías"
            Height          =   300
            Left            =   240
            TabIndex        =   62
            Top             =   3550
            Width           =   3975
         End
         Begin VB.Label Label39 
            Caption         =   "Sobre el Destino y el Impacto del Mismo"
            Height          =   300
            Left            =   240
            TabIndex        =   61
            Top             =   4320
            Width           =   4575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Flujo de Caja Familiar"
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
         Height          =   3750
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   11055
         Begin SICMACT.FlexEdit fgEgresos 
            Height          =   3015
            Left            =   5640
            TabIndex        =   7
            Top             =   480
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   5318
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   1
            EncabezadosNombres=   "Index-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-400-3000-1400-0"
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-1-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-C"
            FormatosEdit    =   "0-3-0-2-2"
            TextArray0      =   "Index"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            TipoBusPersona  =   2
         End
         Begin SICMACT.FlexEdit fgIngresos 
            Height          =   3015
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   5318
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "Index-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-400-2700-1500-0"
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
            ListaControles  =   "0-0-0-1-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-C"
            FormatosEdit    =   "0-3-0-2-2"
            TextArray0      =   "Index"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label Label20 
            Caption         =   "Egresos :"
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
            Left            =   5640
            TabIndex        =   80
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Ingresos :"
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
            TabIndex        =   60
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Propuesta del Credito"
         Height          =   6495
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   9375
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   840
            Width           =   9015
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   1800
            Width           =   9015
         End
         Begin VB.TextBox txtCrediticia 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   2760
            Width           =   9015
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   3720
            Width           =   9015
         End
         Begin VB.TextBox txtGarantias 
            Height          =   735
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   4680
            Width           =   9015
         End
         Begin VB.TextBox txtSustentoIncreVenta 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   5640
            Width           =   9015
         End
         Begin MSMask.MaskEdBox txtFechaVista 
            Height          =   345
            Left            =   7920
            TabIndex        =   51
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label34 
            Caption         =   "Fecha de Vista:"
            Height          =   300
            Left            =   6720
            TabIndex        =   58
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "Sobre el Entorno Familiar del Cliente o Representante"
            Height          =   300
            Left            =   240
            TabIndex        =   57
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label27 
            Caption         =   "Sobre el Giro y la Ubicacion del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   56
            Top             =   1560
            Width           =   4095
         End
         Begin VB.Label Label30 
            Caption         =   "Sobre la Experiencia Crediticia"
            Height          =   300
            Left            =   240
            TabIndex        =   55
            Top             =   2520
            Width           =   4215
         End
         Begin VB.Label Label31 
            Caption         =   "Sobre la Consistencia de la Informacion y la Formalidad del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   54
            Top             =   3480
            Width           =   6255
         End
         Begin VB.Label Label32 
            Caption         =   "Sobre los Colaterales o Garantias"
            Height          =   300
            Left            =   240
            TabIndex        =   53
            Top             =   4440
            Width           =   3975
         End
         Begin VB.Label Label33 
            Caption         =   "Sobre el Destino y el Impacto del Mismo"
            Height          =   300
            Left            =   240
            TabIndex        =   52
            Top             =   5400
            Width           =   4575
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1095
         Left            =   135
         TabIndex        =   72
         Top             =   4300
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   1931
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         ForeColor       =   -2147483635
         TabCaption(0)   =   "Ratios e Indicadores"
         TabPicture(0)   =   "frmCredFormEvalFormatoSinConvenio.frx":037A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame Frame4 
            Height          =   680
            Left            =   120
            TabIndex        =   73
            Top             =   320
            Width           =   10815
            Begin SICMACT.EditMoney txtRatioCapPago 
               Height          =   300
               Left            =   1680
               TabIndex        =   77
               Top             =   255
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
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
               BackColor       =   -2147483643
               ForeColor       =   8421504
               Text            =   "0"
            End
            Begin SICMACT.EditMoney txtRatioIngNeto 
               Height          =   300
               Left            =   9600
               TabIndex        =   78
               Top             =   255
               Width           =   975
               _ExtentX        =   1720
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
               BackColor       =   -2147483643
               ForeColor       =   8421504
               Text            =   "0"
            End
            Begin SICMACT.EditMoney txtRatioExcedente 
               Height          =   300
               Left            =   9600
               TabIndex        =   79
               Top             =   250
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
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
               BackColor       =   -2147483643
               ForeColor       =   8421504
               Text            =   "0"
            End
            Begin VB.Label lblCapPagAceptable 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Aceptable"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   2760
               TabIndex        =   81
               Top             =   360
               Width           =   750
            End
            Begin VB.Label Label16 
               Caption         =   "Ingreso Neto :"
               Height          =   300
               Left            =   8520
               TabIndex        =   75
               Top             =   300
               Width           =   1005
            End
            Begin VB.Label lblCapPag 
               Caption         =   "Capacidad de Pago :"
               Height          =   300
               Left            =   120
               TabIndex        =   74
               Top             =   300
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.Label lblCapPagoCritico 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Critico"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   165
               Left            =   2760
               TabIndex        =   82
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label17 
               Caption         =   "Excedente :"
               Height          =   180
               Left            =   8520
               TabIndex        =   76
               Top             =   300
               Visible         =   0   'False
               Width           =   960
            End
         End
      End
   End
End
Attribute VB_Name = "frmCredFormEvalFormatoSinConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalFormatoSinConvenio                 '
'** Descripción : Formulario para evaluación de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creación    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'
Option Explicit
Dim gsOpeCod As String

Dim aux As ADODB.Recordset
Dim rsFeDatGastoFam As ADODB.Recordset
Dim rsFeDatOtrosIng As ADODB.Recordset
'FIN Para Cargar en la Grilla

Dim MatReferidos As Variant
Dim fsCtaCod As String
Dim fnTipoRegMant As Integer

'Para Cargar en la Cabecera
Dim fsGiroNego As String
Dim fsCliente As String
Dim fnAnio As Integer
Dim fnMes As Integer
Dim fnMontoDeudaSbs As Double
Dim fdFechaDeudaSbs As Date
Dim fnMontSolicitado As Double
Dim fnCuota As Integer
Dim fnExpCredito As Double
Dim fdFechaActual As Date
'FIN Para Cargar en la Cabecera

Dim TipoAportacion As Integer
Dim pMtrBoletaPago As Variant
Dim pMtrReciboHono As Variant
Dim pMtrNegocio As Variant
Dim pnIngNegocio As Double
Dim pnEgrVenta As Double
Dim pnMargBruto As Double
Dim pnIngNeto As Double
Dim pMtrIfis As Variant
Dim nTotalEgresos As Double
Dim nFormato As Integer
Dim fbGrabar As Boolean
Dim fnCuotaProp As Currency
Dim lnColocCondi As Integer
Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agregó segun correo: RUSI
Dim nEstado As Integer

Dim fnTipoPermiso As Integer

Dim objPista As COMManejador.Pista

Dim pnOp As Long
Dim i As Integer, lnFila As Integer

Dim lcMovNro As String 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
Dim MatIfiNoSupervisadaGastoFami As Variant 'CTI320200110 ERS003-2020. Agregó
Dim fbImprimirVB As Boolean 'CTI320200110 ERS003-2020
Dim pnMontoOtrasIfis As Double
Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                     ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer, _
                     Optional ByVal pbImprimirVB As Boolean = False) As Boolean
    
    gsOpeCod = ""
    lcMovNro = "" 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
      
    fnTipoRegMant = psTipoRegMant
    fsCtaCod = psCtaCod
    nEstado = pnEstado
    fbGrabar = False
    nFormato = pnFormato
    fbImprimirVB = pbImprimirVB 'CTI3ERS0032020
    'lblCapPag.Visible = IIf(pnEstado = 2001, True, False)
    
    If nEstado = 2001 Then
        If lnColocCondi <> 4 Then
            cmdInformeVista.Enabled = True
            cmdVerCar.Enabled = True
            cmdImprimir.Enabled = True
        End If
    Else
    'Se inicialiaza los Botones como no editable
    cmdInformeVista.Enabled = False
    cmdVerCar.Enabled = False
    cmdImprimir.Enabled = False
    'FIN Se inicialiaza los Botones como no editable
    End If
    
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    
     Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
     Dim rsDCredito As ADODB.Recordset
                   
     Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
     Set rsDCredito = oDCOMFormatosEval.RecuperarDatosConsumoSinConvenio(psCtaCod) ' Recuperar Datos Basico
                    
     lnColocCondi = rsDCredito!nColocCondicion
     fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses   'Si tiene evaluacion registrada 6 meses (LUCV20171115, agregó según correo: RUSI)
     
    If lnColocCondi = 4 Then
        SSTab1.TabEnabled(1) = False
    Else
        SSTab1.TabEnabled(1) = True
    End If
    
 '(3: Analista, 2: Coordinador, 1: JefeAgencia)
 fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
 
If CargaControlesTipoPermiso(fnTipoPermiso) Then

            If fnTipoRegMant = 1 Then
                TipoAportacion = 0
                Call CargarFlexEdit
                
                If Not (rsDCredito.EOF And rsDCredito.BOF) Then
                
                    If (rsDCredito!cActiGiro) = "" Then
                        MsgBox ("Por favor, actualizar los datos del cliente. " & Chr(13) & "(Actividad o Giro del negocio)"), vbInformation, "Alerta"
                        Exit Function
                    End If
                
                    fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
                    fsCliente = Trim(rsDCredito!cPersNombre)
                    fnAnio = rsDCredito!nAnio
                    fnMes = rsDCredito!nMes
                    fnMontoDeudaSbs = rsDCredito!nMontoUltimaDeudaSBS
                    fdFechaDeudaSbs = rsDCredito!dFechaUltimaDeudaSBS
                    fnMontSolicitado = rsDCredito!nMonto
                    fnCuota = rsDCredito!nCuotas
                    fnExpCredito = rsDCredito!nExpoCred
                    fdFechaActual = rsDCredito!dFechaActual
                
                    ActXCodCtaSinConv.NroCuenta = fsCtaCod
                    txtGiro.Text = fsGiroNego
                    txtCliente.Text = fsCliente
                    spnAno.valor = fnAnio
                    spnMes.valor = fnMes
                    txtUltDeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
                    
                   If fdFechaDeudaSbs = "01/01/1900" Then '26
                    txtFechaDeuda.Text = "__/__/____"
                   Else
                    txtFechaDeuda.Text = fdFechaDeudaSbs
                   End If
                    
                    txtMonSolicitado.Text = Format(fnMontSolicitado, "#,##0.00")
                    txtCuota.Text = Format(fnCuota, "0#")
                    txtExpCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
                    txtFechaIngreso.Text = Format(fdFechaActual, "dd/MM/yyyy")
                
                    cmdGuardar.Visible = True
                    cmdActualizarSinConvenio.Visible = False
                
                    Call Registro
                
                End If
                
            ElseIf fnTipoRegMant = 2 Then
                
                If fnTipoRegMant = 2 And Mantenimineto(IIf(fnTipoRegMant = 2, False, True)) = False Then
                   MsgBox "No Cuenta con Registros", vbInformation, "Aviso"
                   Exit Function
                End If
                
                cmdGuardar.Visible = False
                cmdActualizarSinConvenio.Visible = True
                
                'cmdVerCar.Enabled = False
                'cmdInformeVista.Enabled = False
                'cmdImprimir.Enabled = False
                
                Call Calcular(1)
                Call Registro
                
                'RECO20160728 ************
                If pnEstado = 2001 Or pnEstado = 2002 Then
                    Call CargaRatios(psCtaCod)
                End If
                'RECO FIN ****************
                
            ElseIf fnTipoRegMant = 3 Then
            
                Call Mantenimineto(IIf(fnTipoRegMant = 3, False, True))
                Call Consultar
                Call Calcular(1)
                Call CargaRatios(psCtaCod)
                
                'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                gsOpeCod = gCredConsultarEvaluacionCred
                lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 7 - Consumo Sin Convenio", fsCtaCod, gCodigoCuenta
                Set objPista = Nothing
                'Fin LUCV20181220
                
                 'lblCapPag.Visible = True
                 'txtRatioCapPago.Visible = True
                 
            'Activar los boton y textbox y lbl de ratios
                If pnEstado = 2001 Or pnEstado = 2002 Then
                    cmdInformeVista.Enabled = True
                    cmdImprimir.Enabled = True
                    cmdVerCar.Enabled = True
                    
                    lblCapPag.Visible = True
                    txtRatioCapPago.Visible = True
                End If
                
            End If
Else
Unload Me
Exit Function
        'Me.Show 1
End If

    
    
    'LUCV Agrego *****
    fbGrabar = False
    If Not pbImprimir Then
        If fbImprimirVB Then
            Call cmdActualizarSinConvenio_Click
        End If
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
    'Fin, LUCV *****
    
End Function

Private Sub cmdAgregarRef_Click()

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

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdActualizarSinConvenio_Click()

    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim oDCred As COMDCredito.DCOMFormatosEval
    Dim ActualizarDatos As Boolean
    Dim rsIngresos As ADODB.Recordset
    Dim rsEgresos As ADODB.Recordset
        
If Validar Then
    
    gsOpeCod = gCredMantenimientoEvaluacionCred
    Set objPista = New COMManejador.Pista
    Set oDCred = New COMDCredito.DCOMFormatosEval
    
    Set rsIngresos = IIf(fgIngresos.rows - 1 > 0, fgIngresos.GetRsNew(), Nothing)
    Set rsEgresos = IIf(fgEgresos.rows - 1 > 0, fgEgresos.GetRsNew(), Nothing)
   
    'If lnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
    'Flex a Matriz Referidos **********->
        ReDim MatReferidos(feReferidos.rows - 1, 6)
            For i = 1 To feReferidos.rows - 1
                MatReferidos(i, 0) = feReferidos.TextMatrix(i, 0)
                MatReferidos(i, 1) = feReferidos.TextMatrix(i, 1)
                MatReferidos(i, 2) = feReferidos.TextMatrix(i, 2)
                MatReferidos(i, 3) = feReferidos.TextMatrix(i, 3)
                MatReferidos(i, 4) = feReferidos.TextMatrix(i, 4)
                MatReferidos(i, 5) = feReferidos.TextMatrix(i, 5)
            Next i
    Else
        ReDim MatReferidos(0)
    End If
    'Fin Referidos
    If Not fbImprimirVB Then
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    End If
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
       
    
    ActualizarDatos = oNCOMFormatosEval.ActualizarConsumoSinConvenio_InfCliente(fsCtaCod, 7, txtGiro.Text, spnAno.valor, spnMes.valor, txtUltDeuda.Text, TipoAportacion, _
                                                                                IIf(txtFechaDeuda.Text = "__/__/____", "01/01/1900", txtFechaDeuda.Text), txtMonSolicitado.Text, txtCuota.Text, txtExpCredito.Text, txtNDependientes.Text, txtCodPers.Text, _
                                                                                CDate(txtFechaIngreso.Text), rsIngresos, rsEgresos, pMtrBoletaPago, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto, pMtrNegocio, _
                                                                                pMtrReciboHono, pMtrIfis, IIf(txtdFechaVisita.Text = "__/__/____", CDate(gdFecSis), txtdFechaVisita.Text), txtEntornoCliente.Text, txtGiroNegocio.Text, txtExpCrediticia.Text, _
                                                                                txtFormNegocio.Text, txtColaGarantias.Text, txtImpactoMismo.Text, txtComentario.Text, MatReferidos, lnColocCondi, MatIfiNoSupervisadaGastoFami)

    Call oDCred.RecalculaIndicadoresyRatiosEvaluacion(fsCtaCod)
    
    'JOEP20180725 ERS034-2018
            Call EmiteFormRiesgoCamCred(fsCtaCod)
    'JOEP20180725 ERS034-2018
    
    If ActualizarDatos Then
        fbGrabar = True
            'LUCV20181220
            'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato Sin Convenio", fsCtaCod, gCodigoCuenta
            'If fnTipoRegMant = 1 Then
            '    MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            'Else
            '    MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
            'End If
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 7 - Consumo Sin Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
            If Not fbImprimirVB Then
                MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
            End If
            Dim objCredito As COMDCredito.DCOMCredito
            Set objCredito = New COMDCredito.DCOMCredito
            Call objCredito.ActualizarEstadoxVB(fsCtaCod, 1)
            'Fin LUCV20181220
                
            cmdActualizarSinConvenio.Enabled = False
            cmdGuardar.Visible = False
            
            If lnColocCondi <> 4 Then
                cmdInformeVista.Enabled = True
            End If
            
            If (nEstado = 2001) Then
                
                If lnColocCondi <> 4 Then
                    cmdVerCar.Enabled = True
                End If
                    cmdImprimir.Enabled = True
            End If
            
            If nEstado = 2001 Or nEstado = 2002 Then
                Call CargaRatios(fsCtaCod)
            End If
            
            If fbImprimirVB Then
               cmdActualizarSinConvenio.Enabled = True
               fbImprimirVB = False
            End If
                                   
    Else
        MsgBox "Hubo errores al grabar la información", vbError, "Error"
    End If
    
End If

End Sub

Private Sub Cmdguardar_Click()

    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim oDCred As COMDCredito.DCOMFormatosEval
    Dim GrabarDatos As Boolean
    Dim rsIngresos As ADODB.Recordset
    Dim rsEgresos As ADODB.Recordset

If Validar Then
    
    gsOpeCod = gCredRegistrarEvaluacionCred
    Set objPista = New COMManejador.Pista
    Set oDCred = New COMDCredito.DCOMFormatosEval
    
    Set rsIngresos = IIf(fgIngresos.rows - 1 > 0, fgIngresos.GetRsNew(), Nothing)
    Set rsEgresos = IIf(fgEgresos.rows - 1 > 0, fgEgresos.GetRsNew(), Nothing)
   
    'If lnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
    'Flex a Matriz Referidos **********->
        ReDim MatReferidos(feReferidos.rows - 1, 6)
            For i = 1 To feReferidos.rows - 1
                MatReferidos(i, 0) = feReferidos.TextMatrix(i, 0)
                MatReferidos(i, 1) = feReferidos.TextMatrix(i, 1)
                MatReferidos(i, 2) = feReferidos.TextMatrix(i, 2)
                MatReferidos(i, 3) = feReferidos.TextMatrix(i, 3)
                MatReferidos(i, 4) = feReferidos.TextMatrix(i, 4)
                MatReferidos(i, 5) = feReferidos.TextMatrix(i, 5)
            Next i
    Else
        ReDim MatReferidos(0)
    End If
    'Fin Referidos
    
    If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
                                                                            
    GrabarDatos = oNCOMFormatosEval.GuardarConsumoSinConvenio_InfCliente(fsCtaCod, 7, txtGiro.Text, spnAno.valor, spnMes.valor, txtUltDeuda.Text, TipoAportacion, _
                                                                          IIf(txtFechaDeuda.Text = "__/__/____", "01/01/1900", txtFechaDeuda.Text), txtMonSolicitado.Text, txtCuota.Text, txtExpCredito.Text, txtNDependientes.Text, txtCodPers.Text, CDate(txtFechaIngreso.Text), _
                                                                          rsIngresos, rsEgresos, pMtrBoletaPago, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto, pMtrNegocio, pMtrReciboHono, pMtrIfis, _
                                                                          IIf(txtdFechaVisita.Text = "__/__/____", CDate(gdFecSis), txtdFechaVisita.Text), txtEntornoCliente.Text, txtGiroNegocio.Text, txtExpCrediticia.Text, _
                                                                          txtFormNegocio.Text, txtColaGarantias.Text, txtImpactoMismo.Text, txtComentario.Text, MatReferidos, lnColocCondi, MatIfiNoSupervisadaGastoFami)
            
        Call oDCred.RecalculaIndicadoresyRatiosEvaluacion(fsCtaCod)
        
        'JOEP20180725 ERS034-2018
            Call EmiteFormRiesgoCamCred(fsCtaCod)
        'JOEP20180725 ERS034-2018
        
        If GrabarDatos Then
                
            fbGrabar = True
            'RECO20161020 ERS060-2016 **********************************************************
            Dim oNCOMColocEval As New NCOMColocEval
            Dim lcMovNro As String
            
            If Not ValidaExisteRegProceso(fsCtaCod, gTpoRegCtrlEvaluacion) Then
               lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
               objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato Sin Convenio", fsCtaCod, gCodigoCuenta
               Call oNCOMColocEval.insEstadosExpediente(fsCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
               Set oNCOMColocEval = Nothing
            End If
            'RECO FIN **************************************************************************
            'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato Sin Convenio", fsCtaCod, gCodigoCuenta 'RECO20161020 ERS060-2016
            
            If fnTipoRegMant = 1 Then
                objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 7 - Consumo Sin Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Else
                objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 7 - Consumo Sin Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
            End If
                
                cmdGuardar.Enabled = False
                cmdActualizarSinConvenio.Visible = False
            
            If lnColocCondi <> 4 Then
                cmdInformeVista.Enabled = True
            End If
                
            If (nEstado = 2001) Then
            
                If lnColocCondi <> 4 Then
                    cmdVerCar.Enabled = True
                End If
                
                    cmdImprimir.Enabled = True
                    
            End If
            
'            If nEstado = 2001 Or nEstado = 2002 Then
'                Call CargaRatios(fsCtaCod)
'            End If
                
        Else
            MsgBox "Hubo errores al grabar la información", vbError, "Error"
        End If
        
End If

End Sub

Private Sub cmdImprimir_Click()
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsImformeVisitaConsumoConConvenio As ADODB.Recordset
    
    Dim rsMostrarIngresos As ADODB.Recordset
    Dim rsMostrarEgresos As ADODB.Recordset
    
    Dim rsMostrarCuotasIfis As ADODB.Recordset
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsImformeVisitaConsumoConConvenio = New ADODB.Recordset
    'Set rsImformeVisitaConsumoConConvenio = oDCOMFormatosEval.MostrarDatosInformeVisitaFormatoSinConvenio(fsCtaCod)
    Set rsImformeVisitaConsumoConConvenio = oDCOMFormatosEval.MostrarFormatoSinConvenioInfVisCabecera(fsCtaCod, nFormato)
    Set rsMostrarIngresos = oDCOMFormatosEval.MostrarIngresos(fsCtaCod, nFormato)
    Set rsMostrarEgresos = oDCOMFormatosEval.MostrarEgresos(fsCtaCod, nFormato)
    
    Set rsMostrarCuotasIfis = oDCOMFormatosEval.MostrarCuotasIfis(fsCtaCod, nFormato, 7022)
    
    Dim a As Integer
    Dim B As Integer
    Dim nFila As Integer
    Dim nFila1 As Integer
    
    Dim n As Currency
    Dim c As Integer
    Dim Total As Currency
    a = 50
    B = 29

    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Hoja de Evaluacion Nº " & fsCtaCod
    oDoc.Title = "Hoja de Evaluacion Nº " & fsCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoSinConvenio_HojaEvaluacion" & fsCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
    If Not (rsImformeVisitaConsumoConConvenio.BOF Or rsImformeVisitaConsumoConConvenio.EOF) Then

    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical

    '---------- cabecera ---------------
    oDoc.WImage 45, 45, 45, 113, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(rsImformeVisitaConsumoConConvenio!cAgeDescripcion), "F2", 7.5, hLeft

    oDoc.WTextBox 40, 30, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight
    oDoc.WTextBox 60, 440, 10, 490, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
    oDoc.WTextBox 70, 440, 10, 490, "ANALISTA: " & UCase(rsImformeVisitaConsumoConConvenio!cUser), "F2", 7.5, hLeft
    
    oDoc.WTextBox 65, 100, 10, 400, "HOJA DE EVALUACION", "F2", 12, hCenter
    oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & fsCtaCod, "F2", 7.5, hLeft
    oDoc.WTextBox 90, 440, 10, 300, "MONEDA:" & IIf(Mid(fsCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersCod), "F2", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersNombre), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 440, 10, 200, "DNI: " & Trim(rsImformeVisitaConsumoConConvenio!cPersDni) & "   ", "F2", 7.5, hLeft
    oDoc.WTextBox 110, 440, 10, 200, "RUC: " & Trim(IIf(rsImformeVisitaConsumoConConvenio!cPersRuc = "-", Space(11), rsImformeVisitaConsumoConConvenio!cPersRuc)), "F2", 7.5, hLeft

    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 120, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 130, 55, 1, 160, "Ingresos", "F2", 7.5, hjustify
    oDoc.WTextBox 140, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    
    a = 0
    n = 0
   
    If Not (rsMostrarIngresos.BOF And rsMostrarIngresos.EOF) Then
        For i = 1 To rsMostrarIngresos.RecordCount
        oDoc.WTextBox 150 + a, 55, 1, 160, rsMostrarIngresos!nCodIngr, "F1", 7.5, hjustify
        oDoc.WTextBox 150 + a, 80, 1, 300, rsMostrarIngresos!cConsDescripcion, "F1", 7.5, hjustify
        oDoc.WTextBox 150 + a, 170, 1, 160, Format(rsMostrarIngresos!nMonto, "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
        n = n + rsMostrarIngresos!nMonto
        rsMostrarIngresos.MoveNext
        Next i
        oDoc.WTextBox 150 + a, 250, 1, 160, "Total", "F2", 7.5, hjustify
        oDoc.WTextBox 150 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
    End If
    
    '--------------------------------------------------------------------------------------------------------------------------
    'A = A + 10
    a = a + 10
    oDoc.WTextBox 150 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 150 + a, 55, 1, 190, "Detalle de Sueldo", "F2", 7.5, hjustify
    a = a + 10
    oDoc.WTextBox 150 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 150 + a, 55, 1, 160, "Año", "F2", 7.5, hjustify
    oDoc.WTextBox 150 + a, 130, 1, 160, "Mes", "F2", 7.5, hjustify
    oDoc.WTextBox 150 + a, 170, 1, 160, "Monto", "F2", 7.5, hRight
    a = a + 10
    a = 10
    n = 0
    c = 0
    Total = 0
    If IsArray(pMtrBoletaPago) Then
        For i = 1 To UBound(pMtrBoletaPago, 2)
            oDoc.WTextBox 260 + a, 55, 1, 160, pMtrBoletaPago(1, i), "F1", 7.5, hjustify
            oDoc.WTextBox 260 + a, 130, 1, 160, Format(pMtrBoletaPago(2, i), "0#"), "F1", 7.5, hjustify
            oDoc.WTextBox 260 + a, 170, 1, 160, pMtrBoletaPago(3, i), "F1", 7.5, hRight
                n = n + pMtrBoletaPago(3, i)
                a = a + 10
        Next i
            c = UBound(pMtrBoletaPago, 2)
            
        If c > 0 Then
            Total = n / c
            oDoc.WTextBox 260 + a, 250, 1, 160, "Total", "F2", 7.5, hjustify
            oDoc.WTextBox 260 + a, 170, 1, 160, Format(Total, "#,##0.00"), "F2", 7.5, hRight
        End If
        
    End If

   '--------------------------------------------------------------------------------------------------------------------------
    a = a + 10
    a = a + 10
    oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 260 + a, 55, 1, 160, "Detalle de Otros Negocios", "F2", 7.5, hjustify
    a = a + 10
    oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    
    If pnIngNegocio > 0 Then
        
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Ventas y Costos", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        
        oDoc.WTextBox 260 + a, 55, 1, 160, "Ingresos del Negocio", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, Format(pnIngNegocio, "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Egresos po Venta", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, Format(pnEgrVenta, "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Margen Bruto", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, Format(pnMargBruto, "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
        
        oDoc.WTextBox 260 + a, 55, 1, 160, "Gastos del Negocio", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        
        'A = 0
        n = 0
        If IsArray(pMtrNegocio) Then
            For i = 1 To UBound(pMtrNegocio, 2)
                oDoc.WTextBox 260 + a, 55, 1, 500, pMtrNegocio(0, i), "F1", 7.5, hjustify
                oDoc.WTextBox 260 + a, 80, 1, 500, pMtrNegocio(1, i), "F1", 7.5, hjustify
                oDoc.WTextBox 260 + a, 170, 1, 160, pMtrNegocio(2, i), "F1", 7.5, hRight
                    n = n + pMtrNegocio(2, i)
                    a = a + 10
            Next i
                oDoc.WTextBox 260 + a, 250, 1, 160, "Total", "F2", 7.5, hjustify
                oDoc.WTextBox 260 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
        End If
        
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Resumen", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        
        oDoc.WTextBox 260 + a, 55, 1, 160, "Ingreso Neto(Negocio)", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, Format(pnIngNeto, "#,##0.00"), "F1", 7.5, hRight
    
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    'A = A + 10
    a = a + 10
    a = a + 10
    oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 260 + a, 55, 1, 160, "Detalle de Recibo de Honorarios", "F2", 7.5, hjustify
    a = a + 10
    oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 260 + a, 55, 1, 160, "Año", "F2", 7.5, hjustify
    oDoc.WTextBox 260 + a, 130, 1, 160, "Mes", "F2", 7.5, hjustify
    oDoc.WTextBox 260 + a, 170, 1, 160, "Monto", "F2", 7.5, hRight
    a = a + 10
    n = 0
    n = 0
    c = 0
    Total = 0
    If IsArray(pMtrReciboHono) Then
        For i = 1 To UBound(pMtrReciboHono, 2)
            oDoc.WTextBox 260 + a, 55, 1, 500, pMtrReciboHono(1, i), "F1", 7.5, hjustify
            oDoc.WTextBox 260 + a, 130, 1, 160, Format(pMtrReciboHono(2, i), "0#"), "F1", 7.5, hjustify
            oDoc.WTextBox 260 + a, 170, 1, 160, pMtrReciboHono(3, i), "F1", 7.5, hRight
                n = n + pMtrReciboHono(3, i)
                a = a + 10
        Next i
            c = UBound(pMtrReciboHono, 2)

        If c > 0 Then
            Total = n / c
            oDoc.WTextBox 260 + a, 250, 1, 160, "Total", "F2", 7.5, hjustify
            oDoc.WTextBox 260 + a, 170, 1, 160, Format(Total, "#,##0.00"), "F2", 7.5, hRight
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    'A = A + 10
    a = a + 10
    oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 260 + a, 55, 1, 190, "Egresos", "F2", 7.5, hjustify
    a = a + 10
    oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    'A = 0
    n = 0

    If Not (rsMostrarEgresos.BOF And rsMostrarEgresos.EOF) Then
        For i = 1 To rsMostrarEgresos.RecordCount
        oDoc.WTextBox 260 + a, 55, 1, 160, rsMostrarEgresos!nCodGasto, "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 80, 1, 300, rsMostrarEgresos!cConsDescripcion, "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, Format(rsMostrarEgresos!nMonto, "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
        n = n + rsMostrarEgresos!nMonto
        rsMostrarEgresos.MoveNext
        Next i
        oDoc.WTextBox 260 + a, 100, 1, 160, "Total", "F2", 7.5, hRight
        oDoc.WTextBox 260 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
    End If
    
    '--------------------------------------------------------------------------------------------------------------------------

If a >= 540 Then
    oDoc.NewPage A4_Vertical
    
    oDoc.WImage 45, 45, 45, 113, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(rsImformeVisitaConsumoConConvenio!cAgeDescripcion), "F2", 7.5, hLeft

    oDoc.WTextBox 40, 30, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight
    oDoc.WTextBox 60, 440, 10, 490, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
    oDoc.WTextBox 70, 440, 10, 490, "ANALISTA: " & UCase(rsImformeVisitaConsumoConConvenio!cUser), "F2", 7.5, hLeft
    
    oDoc.WTextBox 65, 100, 10, 400, "HOJA DE EVALUACION", "F2", 12, hCenter
    oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & fsCtaCod, "F2", 7.5, hLeft
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersCod), "F2", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersNombre), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 440, 10, 200, "DNI: " & Trim(rsImformeVisitaConsumoConConvenio!cPersDni) & "   ", "F2", 7.5, hLeft
    oDoc.WTextBox 110, 440, 10, 200, "RUC: " & Trim(IIf(rsImformeVisitaConsumoConConvenio!cPersRuc = "-", Space(11), rsImformeVisitaConsumoConConvenio!cPersRuc)), "F2", 7.5, hLeft
        
    a = 0
    oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 130 + a, 55, 1, 160, "Detalle Cuotas Ifis", "F2", 7.5, hjustify
    a = a + 10
    oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    n = 0
    If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
        For i = 1 To rsMostrarCuotasIfis.RecordCount
        oDoc.WTextBox 130 + a, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 80, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 170, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
        n = n + rsMostrarCuotasIfis!nMonto
        rsMostrarCuotasIfis.MoveNext
        Next i
        oDoc.WTextBox 130 + a, 100, 1, 160, "Total", "F2", 7.5, hRight
        oDoc.WTextBox 130 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
    End If
        
    a = a + 10
    oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 130 + a, 55, 1, 160, "Ratios e Indicadores", "F2", 7.5, hjustify
    a = a + 10
    oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    a = a + 10
    oDoc.WTextBox 130 + a, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
    oDoc.WTextBox 130 + a, 175, 1, 160, txtRatioCapPago.Text, "F1", 7.5, hRight
    If rsImformeVisitaConsumoConConvenio!cTpIngr = "I" Then
    oDoc.WTextBox 130 + a, 330, 1, 160, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hRight
    Else
    oDoc.WTextBox 130 + a, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hRight
    End If
    a = a + 10
    oDoc.WTextBox 130 + a, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
    oDoc.WTextBox 130 + a, 170, 1, 160, txtRatioIngNeto.Text, "F1", 7.5, hRight
    a = a + 10
    oDoc.WTextBox 130 + a, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
    oDoc.WTextBox 130 + a, 170, 1, 160, txtRatioExcedente.Text, "F1", 7.5, hRight
Else
    
    If a >= 470 Then
        oDoc.NewPage A4_Vertical
    
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsImformeVisitaConsumoConConvenio!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 30, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight
        oDoc.WTextBox 60, 440, 10, 490, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
        oDoc.WTextBox 70, 440, 10, 490, "ANALISTA: " & UCase(rsImformeVisitaConsumoConConvenio!cUser), "F2", 7.5, hLeft
        
        oDoc.WTextBox 65, 100, 10, 400, "HOJA DE EVALUACION", "F2", 12, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & fsCtaCod, "F2", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersCod), "F2", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersNombre), "F2", 7.5, hLeft
        oDoc.WTextBox 100, 440, 10, 200, "DNI: " & Trim(rsImformeVisitaConsumoConConvenio!cPersDni) & "   ", "F2", 7.5, hLeft
        oDoc.WTextBox 110, 440, 10, 200, "RUC: " & Trim(IIf(rsImformeVisitaConsumoConConvenio!cPersRuc = "-", Space(11), rsImformeVisitaConsumoConConvenio!cPersRuc)), "F2", 7.5, hLeft
            
        a = 0
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Detalle Cuotas Ifis", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        n = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For i = 1 To rsMostrarCuotasIfis.RecordCount
            oDoc.WTextBox 130 + a, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
            oDoc.WTextBox 130 + a, 80, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
            oDoc.WTextBox 130 + a, 170, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
            a = a + 10
            n = n + rsMostrarCuotasIfis!nMonto
            rsMostrarCuotasIfis.MoveNext
            Next i
            oDoc.WTextBox 130 + a, 100, 1, 160, "Total", "F2", 7.5, hRight
            oDoc.WTextBox 130 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
        End If
            
        a = a + 10
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Ratios e Indicadores", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 175, 1, 160, txtRatioCapPago.Text, "F1", 7.5, hRight
        If rsImformeVisitaConsumoConConvenio!cTpIngr = "I" Then
        oDoc.WTextBox 130 + a, 330, 1, 160, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hRight
        Else
        oDoc.WTextBox 130 + a, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hRight
        End If
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 170, 1, 160, txtRatioIngNeto.Text, "F1", 7.5, hRight
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 170, 1, 160, txtRatioExcedente.Text, "F1", 7.5, hRight
    ElseIf a >= 520 Then
        oDoc.NewPage A4_Vertical
    
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsImformeVisitaConsumoConConvenio!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 30, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight
        oDoc.WTextBox 60, 440, 10, 490, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
        oDoc.WTextBox 70, 440, 10, 490, "ANALISTA: " & UCase(rsImformeVisitaConsumoConConvenio!cUser), "F2", 7.5, hLeft
        
        oDoc.WTextBox 65, 100, 10, 400, "HOJA DE EVALUACION", "F2", 12, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & fsCtaCod, "F2", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersCod), "F2", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersNombre), "F2", 7.5, hLeft
        oDoc.WTextBox 100, 440, 10, 200, "DNI: " & Trim(rsImformeVisitaConsumoConConvenio!cPersDni) & "   ", "F2", 7.5, hLeft
        oDoc.WTextBox 110, 440, 10, 200, "RUC: " & Trim(IIf(rsImformeVisitaConsumoConConvenio!cPersRuc = "-", Space(11), rsImformeVisitaConsumoConConvenio!cPersRuc)), "F2", 7.5, hLeft
            
        a = 0
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Detalle Cuotas Ifis", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        n = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For i = 1 To rsMostrarCuotasIfis.RecordCount
            oDoc.WTextBox 130 + a, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
            oDoc.WTextBox 130 + a, 80, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
            oDoc.WTextBox 130 + a, 170, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
            a = a + 10
            n = n + rsMostrarCuotasIfis!nMonto
            rsMostrarCuotasIfis.MoveNext
            Next i
            oDoc.WTextBox 130 + a, 100, 1, 160, "Total", "F2", 7.5, hRight
            oDoc.WTextBox 130 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
        End If
            
        a = a + 10
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Ratios e Indicadores", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 130 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 175, 1, 160, txtRatioCapPago.Text, "F1", 7.5, hRight
        If rsImformeVisitaConsumoConConvenio!cTpIngr = "I" Then
        oDoc.WTextBox 130 + a, 330, 1, 160, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hRight
        Else
        oDoc.WTextBox 130 + a, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hRight
        End If
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 170, 1, 160, txtRatioIngNeto.Text, "F1", 7.5, hRight
        a = a + 10
        oDoc.WTextBox 130 + a, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
        oDoc.WTextBox 130 + a, 170, 1, 160, txtRatioExcedente.Text, "F1", 7.5, hRight
    Else
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Detalle Cuotas Ifis", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        n = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For i = 1 To rsMostrarCuotasIfis.RecordCount
            oDoc.WTextBox 260 + a, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
            oDoc.WTextBox 260 + a, 80, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
            oDoc.WTextBox 260 + a, 170, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
            a = a + 10
            n = n + rsMostrarCuotasIfis!nMonto
            rsMostrarCuotasIfis.MoveNext
            Next i
            oDoc.WTextBox 260 + a, 100, 1, 160, "Total", "F2", 7.5, hRight
            oDoc.WTextBox 260 + a, 170, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
        End If
            
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Ratios e Indicadores", "F2", 7.5, hjustify
        a = a + 10
        oDoc.WTextBox 260 + a, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 175, 1, 160, txtRatioCapPago.Text, "F1", 7.5, hRight
        If rsImformeVisitaConsumoConConvenio!cTpIngr = "I" Then
        oDoc.WTextBox 260 + a, 330, 1, 160, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hRight
        Else
        oDoc.WTextBox 260 + a, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hRight
        End If
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, txtRatioIngNeto.Text, "F1", 7.5, hRight
        a = a + 10
        oDoc.WTextBox 260 + a, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
        oDoc.WTextBox 260 + a, 170, 1, 160, txtRatioExcedente.Text, "F1", 7.5, hRight
    End If
End If
    
    oDoc.PDFClose
    oDoc.Show
    
    Else
        MsgBox "Los Datos de Hoja de Evaluacion se mostrara despues de GRABAR la Sugerencia", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdInformeVista_Click()

    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(fsCtaCod)
       
    Me.cmdInformeVista.Enabled = False
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atención"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes
    Me.cmdInformeVista.Enabled = True
    RSClose rsInfVisita
End Sub

Private Sub cmdQuitar2_Click()

    If MsgBox("Esta Seguro de Eliminar  a " & feReferidos.TextMatrix(feReferidos.row, 1) & " del Registro?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    feReferidos.EliminaFila (feReferidos.row)
    
End Sub

'JOEP20180725 ERS034-2018
Private Sub cmdMNME_Click()
    Call frmCredFormEvalCredCel.Inicio(ActXCodCtaSinConv.NroCuenta, 11)
End Sub
'JOEP20180725 ERS034-2018

Private Sub cmdVerCar_Click()
    
    Dim oCred As COMNCredito.NCOMFormatosEval
    Dim oDCredSbs As COMDCredito.DCOMFormatosEval
    Dim R As ADODB.Recordset
    Dim lcDNI, lcRUC As String

    Dim RSbs, RDatFin1, RCap As ADODB.Recordset

    Set oCred = New COMNCredito.NCOMFormatosEval
    Call oCred.RecuperaDatosInformeComercial(ActXCodCtaSinConv.NroCuenta, R)
    Set oCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lcDNI = Trim(R!dni_deudor)
    lcRUC = Trim(R!ruc_deudor)
    
    Set oDCredSbs = New COMDCredito.DCOMFormatosEval
        Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC)
        Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActXCodCtaSinConv.NroCuenta, nFormato)
        
    Set oDCredSbs = Nothing
    
    Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActXCodCtaSinConv.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1)
    
    RSClose R
End Sub

Private Sub feReferidos_OnCellChange(pnRow As Long, pnCol As Long)

    Select Case pnCol
    
    Case 2
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
            
        Else
            MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
        
    Case 3
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Telefono Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            
        Else
            MsgBox "Telefono Mal Ingresado", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
        
    Case 5
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    End Select
    
End Sub

Private Sub feReferidos_RowColChange()

    If feReferidos.col = 1 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.col = 2 Then
        feReferidos.MaxLength = "8"
    ElseIf feReferidos.col = 3 Then
        feReferidos.MaxLength = "9"
    ElseIf feReferidos.col = 4 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.col = 5 Then
        feReferidos.MaxLength = "8"
    End If
    
End Sub

'Para activar el Boton Buscar en la Grilla Egresos
Private Sub fgEgresos_Click()

    If fgEgresos.col = 3 Then
        If CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = 5 _
        Or CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            fgEgresos.ListaControles = "0-0-0-1-0"
        Else
            fgEgresos.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
      Case 7
        Me.fgEgresos.ColumnasAEditar = "X-X-X-X"
      Case 8
        Me.fgEgresos.ColumnasAEditar = "X-X-X-X"
      Case Else
        Me.fgEgresos.ColumnasAEditar = "X-X-X-3"
    End Select

'Se usa solo para consultar al cliente
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
            Case 5, gCodCuotaIfiNoSupervisadaGastoFami  'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                Me.fgEgresos.ColumnasAEditar = "X-X-X-3-X"
            Case 1, 2, 3, 4, 6, 7, 8
                Me.fgEgresos.ColumnasAEditar = "X-X-X-X-X"
        End Select
    End If
    
End Sub

Private Sub fgEgresos_EnterCell()

    If fgEgresos.col = 3 Then
        If CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = 5 _
            Or CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            fgEgresos.ListaControles = "0-0-0-1-0"
        Else
            fgEgresos.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
      Case 7
        Me.fgEgresos.ColumnasAEditar = "X-X-X-X"
      Case 8
        Me.fgEgresos.ColumnasAEditar = "X-X-X-X"
      Case Else
        Me.fgEgresos.ColumnasAEditar = "X-X-X-3"
    End Select

'Se usa solo para consultar al cliente
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
            Case 5, gCodCuotaIfiNoSupervisadaGastoFami  'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                Me.fgEgresos.ColumnasAEditar = "X-X-X-3-X"
            Case 1, 2, 3, 4, 6, 7, 8
                Me.fgEgresos.ColumnasAEditar = "X-X-X-X-X"
        End Select
    End If
    
End Sub


Private Sub fgEgresos_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
Dim pnTotal As Currency
    If fgEgresos.col = 3 Then
        If CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = 5 Then
            If fgEgresos.TextMatrix(fgEgresos.row, 3) = "0.00" Then
                Set pMtrIfis = Nothing
                frmCredFormEvalCuotasIfis.Inicio fgEgresos.TextMatrix(5, 3), pnTotal, pMtrIfis, fgEgresos.TextMatrix(5, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami 'CTI320200110 ERS003-2020, Agregó: fgEgresos.TextMatrix(5, 2)
            Else
                frmCredFormEvalCuotasIfis.Inicio fgEgresos.TextMatrix(5, 3), pnTotal, pMtrIfis, fgEgresos.TextMatrix(5, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami 'CTI320200110 ERS003-2020, Agregó: fgEgresos.TextMatrix(5, 2)
            End If
            pnOp = pnTotal
            psCodigo = Format(pnTotal, "#,##0.00")
        ElseIf CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            If fgEgresos.TextMatrix(fgEgresos.row, 3) = "0.00" Then
                Set MatIfiNoSupervisadaGastoFami = Nothing
                frmCredFormEvalCuotasIfis.Inicio fgEgresos.TextMatrix(gCodCuotaIfiNoSupervisadaGastoFami, 3), pnTotal, MatIfiNoSupervisadaGastoFami, fgEgresos.TextMatrix(5, 2), gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami
            Else
                frmCredFormEvalCuotasIfis.Inicio fgEgresos.TextMatrix(gCodCuotaIfiNoSupervisadaGastoFami, 3), pnTotal, MatIfiNoSupervisadaGastoFami, fgEgresos.TextMatrix(5, 2), gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami
            End If
            pnOp = pnTotal
            psCodigo = Format(pnTotal, "#,##0.00")
        End If
    End If
End Sub

Private Sub fgEgresos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.fgEgresos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{Tab}"
        fgEgresos.SetFocus
        Exit Sub
    End If
End Sub

Private Sub fgEgresos_RowColChange()

    If fgEgresos.col = 3 Then
        fgEgresos.AvanceCeldas = Vertical
    Else
        fgEgresos.AvanceCeldas = Horizontal
    End If
    
    If fgEgresos.col = 3 Then
            If CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = 5 _
                Or CInt(fgEgresos.TextMatrix(fgEgresos.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            fgEgresos.ListaControles = "0-0-0-1-0"
            Else
            fgEgresos.ListaControles = "0-0-0-0-0"
            End If
    End If
    
    Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
      Case 7
        Me.fgEgresos.ColumnasAEditar = "X-X-X-X"
      Case 8
        Me.fgEgresos.ColumnasAEditar = "X-X-X-X"
      Case Else
        Me.fgEgresos.ColumnasAEditar = "X-X-X-3"
    End Select
    
    'Se usa solo para consultar al cliente
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
            Case 5, gCodCuotaIfiNoSupervisadaGastoFami  'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                Me.fgEgresos.ColumnasAEditar = "X-X-X-3-X"
            Case 1, 2, 3, 4, 6, 7, 8
                Me.fgEgresos.ColumnasAEditar = "X-X-X-X-X"
        End Select
    End If
    
End Sub

Private Sub fgEgresos_OnCellChange(pnRow As Long, pnCol As Long)
    
    Select Case pnCol
    Case 3
        If IsNumeric(fgEgresos.TextMatrix(pnRow, pnCol)) Then
            
            Call Calcular(2)
        Else
            fgEgresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End Select
    
    Select Case pnRow
        Case 1, 2, 3, 4, 6
            If IsNumeric(fgEgresos.TextMatrix(pnRow, pnCol)) Then
               Select Case CCur(fgEgresos.TextMatrix(pnRow, pnCol))
                Case Is >= 0
                    Case Else
                        MsgBox "Monto mal Ingresado", vbInformation, "Alerta"
                        fgEgresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                        Exit Sub
                End Select
            Else
                fgEgresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
            End If
    End Select

End Sub
'FIN Para activar el Boton Buscar en la Grilla Egresos

'Para activar el Boton Buscar en la Grilla Ingresos
Private Sub fgIngresos_Click()
    If fgIngresos.col = 3 Then
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            Else
                fgIngresos.ListaControles = "0-0-0-0-0"
            End If
    End If
   
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
            Case 1, 5, 6
                Me.fgIngresos.ColumnasAEditar = "X-X-X-3-X"
            Case 2, 3, 4
                Me.fgIngresos.ColumnasAEditar = "X-X-X-X-X"
        End Select
    End If
 
End Sub

Private Sub fgIngresos_EnterCell()
        If fgIngresos.col = 3 Then
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            Else
                fgIngresos.ListaControles = "0-0-0-0-0"
            End If
        End If
    
        If fnTipoRegMant = 3 Then
           Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
               Case 1, 5, 6
                   Me.fgIngresos.ColumnasAEditar = "X-X-X-3-X"
               Case 2, 3, 4
                   Me.fgIngresos.ColumnasAEditar = "X-X-X-X-X"
               End Select
        End If
 
End Sub

Private Sub fgIngresos_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)

Dim pnTotal As Double

    If fgIngresos.col = 3 Then
    
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                
                If fgIngresos.TextMatrix(fgIngresos.row, 3) = "0.00" Then
                   Set pMtrBoletaPago = Nothing
                       frmCredFormEvalBoletaPago.Inicio pnTotal, pMtrBoletaPago
                Else
                       frmCredFormEvalBoletaPago.Inicio pnTotal, pMtrBoletaPago
                End If
                       psCodigo = Format(pnTotal, "#,##0.00")
                
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                
                If fgIngresos.TextMatrix(fgIngresos.row, 3) = "0.00" Then
                    Set pMtrNegocio = Nothing
                        frmCredFormEvalNegocio.Inicio fsCtaCod, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto, pMtrNegocio
                Else
                        frmCredFormEvalNegocio.Inicio fsCtaCod, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto, pMtrNegocio
                End If
                    psCodigo = Format(pnIngNeto, "#,##0.00")
                
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                                    
                If fgIngresos.TextMatrix(fgIngresos.row, 3) = "0.00" Then
                    Set pMtrReciboHono = Nothing
                        frmCredFormEvalReciboHonorarios.Inicio pnTotal, pMtrReciboHono
                Else
                        frmCredFormEvalReciboHonorarios.Inicio pnTotal, pMtrReciboHono
                End If
                    psCodigo = Format(pnTotal, "#,##0.00")
                    
            End If
        End If

End Sub

Private Sub fgIngresos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
        
    Editar = Split(Me.fgIngresos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{Tab}"
        Exit Sub
    End If
    
End Sub

Private Sub fgIngresos_RowColChange()

'Se usa para que se direccione hacia abajo en una sola Columna
    If fgIngresos.col = 3 Then
        fgIngresos.AvanceCeldas = Vertical
    Else
        fgIngresos.AvanceCeldas = Horizontal
    End If
    
'Se usa para activar el control en la celda indicada
    If fgIngresos.col = 3 Then
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            Else
                fgIngresos.ListaControles = "0-0-0-0-0"
            End If
    End If

'Se usa solo cuando entra a consultar el credito
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
            Case 1, 5, 6
                Me.fgIngresos.ColumnasAEditar = "X-X-X-3-X"
            Case 2, 3, 4
                Me.fgIngresos.ColumnasAEditar = "X-X-X-X-X"
            End Select
    End If
    
End Sub

Private Sub fgIngresos_OnCellChange(pnRow As Long, pnCol As Long)

    If pnCol = 3 Then
        If IsNumeric(fgIngresos.TextMatrix(pnRow, pnCol)) Then
            Call Calcular(1)
        Else
            fgIngresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End If
    
    Select Case pnRow
    Case 2, 3, 4
        If IsNumeric(fgIngresos.TextMatrix(pnRow, pnCol)) Then
           Select Case CCur(fgIngresos.TextMatrix(pnRow, pnCol))
            Case Is >= 0
                Case Else
                    MsgBox "Monto Incorrecto", vbInformation, "Alerta"
                    fgIngresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                    Exit Sub
            End Select
        Else
            fgIngresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End Select
    
End Sub

Private Sub CargarFlexEdit() 'Cargar los Datos en el FlexEdit
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    
   Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(7, fsCtaCod, _
                                                     aux, _
                                                     rsFeDatGastoFam, _
                                                     rsFeDatOtrosIng)
'Otros Ingresos
    fgIngresos.Clear
    fgIngresos.FormaCabecera
    'fgIngresos.Rows = 3
    Call LimpiaFlex(fgIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            fgIngresos.AdicionaFila
            lnFila = fgIngresos.row
            fgIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
            fgIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
            fgIngresos.TextMatrix(lnFila, 3) = Format(rsFeDatOtrosIng!nMonto, "#,##0.00")
            rsFeDatOtrosIng.MoveNext
            
            Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
                'Case 1, 5, 6 'CTI320200110 ERS003-2020. Comentó
                Case 1, 6 'CTI320200110 ERS003-2020. Agregó
                    fgIngresos.BackColorRow (&HC0FFFF)
                'CTI320200110 ERS003-2020. Agregó
                Case 4
                    If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
                        Me.fgIngresos.RowHeight(lnFila) = 1
                    End If
                Case 5
                    If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
                        Me.fgIngresos.RowHeight(lnFila) = 1
                    Else
                        fgIngresos.BackColorRow (&HC0FFFF)
                    End If
                'Fin CTI320200110
            End Select
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing
    
'Gastos Familiares
    fgEgresos.Clear
    fgEgresos.FormaCabecera
    'fgEgresos.Rows = 3
    Call LimpiaFlex(fgEgresos)
        Do While Not rsFeDatGastoFam.EOF
            fgEgresos.AdicionaFila
            lnFila = fgEgresos.row
            fgEgresos.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            fgEgresos.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            fgEgresos.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,##0.00")
            rsFeDatGastoFam.MoveNext
            
            Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
                Case 5, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                    fgEgresos.BackColorRow (&HC0FFFF)
                Case 7
                    fgEgresos.ForeColorRow vbBlack, True
                Case 8
                    fgEgresos.ForeColorRow vbBlack, True
            End Select
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
    
End Sub

Private Sub Form_Load()

SSTab1.Tab = 0
CentraForm Me

lblCapPagoCritico.Visible = False
lblCapPagAceptable.Visible = False

'JOEP20180725 ERS034-2018
    If fnTipoRegMant = 3 Then
        If Not ConsultaRiesgoCamCred(fsCtaCod) Then
            cmdMNME.Visible = True
        End If
    End If
'JOEP20180725 ERS034-2018

End Sub

Private Sub optTipoAportacion_Click(Index As Integer)
    'Tipo de Aportacion
    '1:AFP ; 2:ONP , 3:N/A
    
    TipoAportacion = Index
    
End Sub

' Para Buscar al Empleador
Private Sub txtCodPers_EmiteDatos()
    Dim oDPersonaS As COMDPersona.DCOMPersonas
    Dim sPersCod As String
    Dim oRS As ADODB.Recordset
    
    If Trim(txtCodPers.Text) = "" Then Exit Sub
              
       sPersCod = Trim(txtCodPers.Text)
       
       Set oDPersonaS = New COMDPersona.DCOMPersonas
       Set oRS = oDPersonaS.BuscaCliente(sPersCod, BusquedaCodigo)
       Set oDPersonaS = Nothing
                                   
       If Not oRS.EOF And Not oRS.BOF Then
        txtNombreEmpleador.Text = oRS!cPersNombre
       End If
        RSClose oRS
       
End Sub
'FIN Para Buscar al Empleador

Public Function Mantenimineto(ByVal pbMantenimiento As Boolean) As Boolean
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsMantenimientoSinConvenio As ADODB.Recordset
    Dim rsMantenimientoSinConvenioPropuestaCredito As ADODB.Recordset
    Dim rsMantenimientoSinConvenioReferidos As ADODB.Recordset
    Dim rsMantenimientoSinConvenioIngresos As ADODB.Recordset
    Dim rsMantenimientoSinConvenioEgresos As ADODB.Recordset
    Dim rsDatIfiNoSupervisadaGastoFami As ADODB.Recordset 'CTI3 ERS0032020
    Dim rsMantenimientoSinConvenioBoletaPago As ADODB.Recordset
    Dim rsMantenimientoSinConvenioReciboHonorarios As ADODB.Recordset
    Dim rsMantenimientoSinConvenioCuotasIfis As ADODB.Recordset
    Dim rsMantenimientoSinConvenioGastoNegocio As ADODB.Recordset
    Dim rsMantenimientoSinConvenioGastoNegocioVentasCosto As ADODB.Recordset
    
    Mantenimineto = False

    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Set rsMantenimientoSinConvenio = New ADODB.Recordset
    Set rsMantenimientoSinConvenioPropuestaCredito = New ADODB.Recordset
    Set rsMantenimientoSinConvenioReferidos = New ADODB.Recordset
    Set rsMantenimientoSinConvenioIngresos = New ADODB.Recordset
    Set rsMantenimientoSinConvenioEgresos = New ADODB.Recordset
    Set rsMantenimientoSinConvenioBoletaPago = New ADODB.Recordset
    Set rsDatIfiNoSupervisadaGastoFami = New ADODB.Recordset
     
    Set rsMantenimientoSinConvenio = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenio(fsCtaCod)
    Set rsMantenimientoSinConvenioPropuestaCredito = oDCOMFormatosEval.RecuperarConsumoSinConvenioPropuestaCredito(fsCtaCod, nFormato)
    Set rsMantenimientoSinConvenioReferidos = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioReferidos(fsCtaCod)
    Set rsMantenimientoSinConvenioIngresos = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioIngresos(fsCtaCod)
    Set rsMantenimientoSinConvenioEgresos = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioEgresos(fsCtaCod)
    
    Set rsMantenimientoSinConvenioBoletaPago = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioBoletaPago(fsCtaCod)
    Set rsMantenimientoSinConvenioReciboHonorarios = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioReciboHonorarios(fsCtaCod)
    Set rsDatIfiNoSupervisadaGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami) 'CTI320200110 ERS003-2020. Agregó
    
    pnMontoOtrasIfis = 0
    If Not fbImprimirVB Then
        'Set rsMantenimientoSinConvenioCuotasIfis = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioCuotasIfis(fsCtaCod)
        Set rsMantenimientoSinConvenioCuotasIfis = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami) 'CTI320200110 ERS003-2020. Agregó
    Else
        Set rsMantenimientoSinConvenioCuotasIfis = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami)
        If Not (rsMantenimientoSinConvenioCuotasIfis.BOF Or rsMantenimientoSinConvenioCuotasIfis.EOF) Then
            For i = 1 To rsMantenimientoSinConvenioCuotasIfis.RecordCount
               pnMontoOtrasIfis = pnMontoOtrasIfis + rsMantenimientoSinConvenioCuotasIfis!nMonto
               rsMantenimientoSinConvenioCuotasIfis.MoveNext
            Next i
            rsMantenimientoSinConvenioCuotasIfis.MoveFirst
        End If
    End If
    'Carga de rsDatIfiNoSupervisadaGastoFami -> Matrix

    ReDim MatIfiNoSupervisadaGastoFami(rsDatIfiNoSupervisadaGastoFami.RecordCount, 4)
    Dim j As Integer
    j = 0
    Do While Not rsDatIfiNoSupervisadaGastoFami.EOF
        MatIfiNoSupervisadaGastoFami(j, 0) = rsDatIfiNoSupervisadaGastoFami!nNroCuota
        MatIfiNoSupervisadaGastoFami(j, 1) = rsDatIfiNoSupervisadaGastoFami!CDescripcion
        MatIfiNoSupervisadaGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiNoSupervisadaGastoFami!nMonto), 0, rsDatIfiNoSupervisadaGastoFami!nMonto), "#0.00")
        rsDatIfiNoSupervisadaGastoFami.MoveNext
    j = j + 1
    Loop
    rsDatIfiNoSupervisadaGastoFami.Close
    Set rsDatIfiNoSupervisadaGastoFami = Nothing
    'Fin CTI320200110 ERS003-2020
    
    Set rsMantenimientoSinConvenioGastoNegocio = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioGastoNegocio(fsCtaCod)
    Set rsMantenimientoSinConvenioGastoNegocioVentasCosto = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioGastoNegocioVentasCosto(fsCtaCod)
    
    If Not (rsMantenimientoSinConvenio.BOF And rsMantenimientoSinConvenio.EOF) Then
        
        ActXCodCtaSinConv.NroCuenta = rsMantenimientoSinConvenio!cCtaCod
        txtGiro.Text = rsMantenimientoSinConvenio!cActividad
        txtCliente.Text = rsMantenimientoSinConvenio!cPersNombreClie
        spnAno.valor = rsMantenimientoSinConvenio!nAntgAnios
        spnMes.valor = rsMantenimientoSinConvenio!nAntgMes
        txtUltDeuda.Text = Format(rsMantenimientoSinConvenio!nUltEndeSBS, "#,##0.00")
        optTipoAportacion(rsMantenimientoSinConvenio!nTipoAportacion).value = 1
        
        If rsMantenimientoSinConvenio!dUltEndeuSBS = "01/01/1900" Then
            txtFechaDeuda.Text = "__/__/____"
        Else
            txtFechaDeuda.Text = rsMantenimientoSinConvenio!dUltEndeuSBS
        End If
        
        txtMonSolicitado.Text = Format(rsMantenimientoSinConvenio!nMontoSol, "#,##0.00")
        txtCuota.Text = rsMantenimientoSinConvenio!nNumCuotas
        txtExpCredito.Text = Format(rsMantenimientoSinConvenio!nExposiCred, "#,##0.00")
        txtNDependientes.Text = Format(rsMantenimientoSinConvenio!nNunDepend, "0#")
        txtCodPers.Text = rsMantenimientoSinConvenio!cPersCodEmpleado
        txtNombreEmpleador.Text = rsMantenimientoSinConvenio!cPersNombre
        txtFechaIngreso.Text = Format(rsMantenimientoSinConvenio!dFecEval, "dd/mm/yyyy")
        
        txtComentario.Text = rsMantenimientoSinConvenio!cComentario
        
        Mantenimineto = True
    End If
        RSClose rsMantenimientoSinConvenio
        
    If lnColocCondi <> 4 Then
        If Not (rsMantenimientoSinConvenioPropuestaCredito.BOF And rsMantenimientoSinConvenioPropuestaCredito.EOF) Then
            txtdFechaVisita.Text = rsMantenimientoSinConvenioPropuestaCredito!dFecVisita
            txtEntornoCliente.Text = rsMantenimientoSinConvenioPropuestaCredito!cEntornoFami
            txtGiroNegocio.Text = rsMantenimientoSinConvenioPropuestaCredito!cGiroUbica
            txtExpCrediticia.Text = rsMantenimientoSinConvenioPropuestaCredito!cExpeCrediticia
            txtFormNegocio.Text = rsMantenimientoSinConvenioPropuestaCredito!cFormalNegocio
            txtColaGarantias.Text = rsMantenimientoSinConvenioPropuestaCredito!cColateGarantia
            txtImpactoMismo.Text = rsMantenimientoSinConvenioPropuestaCredito!cDestino
        Mantenimineto = True
        End If
    End If
        RSClose rsMantenimientoSinConvenioPropuestaCredito
    
    If Not (rsMantenimientoSinConvenioReferidos.EOF And rsMantenimientoSinConvenioReferidos.BOF) Then
        Do While Not rsMantenimientoSinConvenioReferidos.EOF
            feReferidos.AdicionaFila
            lnFila = feReferidos.row
            
            feReferidos.TextMatrix(lnFila, 1) = rsMantenimientoSinConvenioReferidos!cNombre
            feReferidos.TextMatrix(lnFila, 2) = rsMantenimientoSinConvenioReferidos!cDniNom
            feReferidos.TextMatrix(lnFila, 3) = rsMantenimientoSinConvenioReferidos!cTelf
            feReferidos.TextMatrix(lnFila, 4) = rsMantenimientoSinConvenioReferidos!cReferido
            feReferidos.TextMatrix(lnFila, 5) = rsMantenimientoSinConvenioReferidos!cDNIRef
                    
            rsMantenimientoSinConvenioReferidos.MoveNext
        Loop
            rsMantenimientoSinConvenioReferidos.Close
            Set rsMantenimientoSinConvenioReferidos = Nothing
    End If
    
    If Not (rsMantenimientoSinConvenioIngresos.EOF And rsMantenimientoSinConvenioIngresos.BOF) Then
        FormateaFlex fgIngresos
        Do While Not rsMantenimientoSinConvenioIngresos.EOF
            fgIngresos.AdicionaFila
            lnFila = fgIngresos.row
            
            fgIngresos.TextMatrix(lnFila, 1) = rsMantenimientoSinConvenioIngresos!nCodIngr
            fgIngresos.TextMatrix(lnFila, 2) = rsMantenimientoSinConvenioIngresos!cConsDescripcion
            fgIngresos.TextMatrix(lnFila, 3) = Format(rsMantenimientoSinConvenioIngresos!nMonto, "#,##0.00")
                           
            rsMantenimientoSinConvenioIngresos.MoveNext
            
            Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
                'Case 1, 5, 6 'CTI320200110 ERS003-2020. Comentó
                Case 1, 6 'CTI320200110 ERS003-2020. Agregó
                        fgIngresos.BackColorRow (&HC0FFFF)
                'CTI320200110 ERS003-2020. Agregó
                Case 4
                    If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
                        Me.fgIngresos.RowHeight(lnFila) = 1
                    End If
                Case 5
                    If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
                        Me.fgIngresos.RowHeight(lnFila) = 1
                    Else
                        fgIngresos.BackColorRow (&HC0FFFF)
                    End If
                'Fin CTI320200110
            End Select
        Loop
            rsMantenimientoSinConvenioIngresos.Close
            Set rsMantenimientoSinConvenioIngresos = Nothing
    End If
    
    If Not (rsMantenimientoSinConvenioEgresos.EOF And rsMantenimientoSinConvenioEgresos.BOF) Then
    FormateaFlex fgEgresos
        Do While Not rsMantenimientoSinConvenioEgresos.EOF
            fgEgresos.AdicionaFila
            lnFila = fgEgresos.row
            
            fgEgresos.TextMatrix(lnFila, 1) = rsMantenimientoSinConvenioEgresos!nCodGasto
            fgEgresos.TextMatrix(lnFila, 2) = rsMantenimientoSinConvenioEgresos!cConsDescripcion
            fgEgresos.TextMatrix(lnFila, 3) = Format(rsMantenimientoSinConvenioEgresos!nMonto, "#,##0.00")
            
            If rsMantenimientoSinConvenioEgresos!nCodGasto = 5 And fbImprimirVB = True Then
                fgEgresos.TextMatrix(lnFila, 3) = Format(pnMontoOtrasIfis, "#,##0.00")
            End If
                           
            rsMantenimientoSinConvenioEgresos.MoveNext
            
            Select Case CInt(fgEgresos.TextMatrix(fgEgresos.row, 1))
                Case 5, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                    Me.fgEgresos.BackColorRow (&HC0FFFF)
                Case 7
                    Me.fgEgresos.ForeColorRow vbBlack, True
                Case 8
                    Me.fgEgresos.ForeColorRow vbBlack, True
            End Select
            
        Loop
            rsMantenimientoSinConvenioEgresos.Close
            Set rsMantenimientoSinConvenioEgresos = Nothing
        
    End If
                        
    ReDim pMtrBoletaPago(4, 0) 'ACTA Nº 112-2018 JOEP20180614
        For i = 1 To (rsMantenimientoSinConvenioBoletaPago.RecordCount)
            ReDim Preserve pMtrBoletaPago(4, i) 'ACTA Nº 112-2018 JOEP20180614
              pMtrBoletaPago(1, i) = rsMantenimientoSinConvenioBoletaPago!nAnio
              pMtrBoletaPago(2, i) = rsMantenimientoSinConvenioBoletaPago!nMes
              pMtrBoletaPago(3, i) = Format(rsMantenimientoSinConvenioBoletaPago!nMontoBruto, "#,##0.00") 'ACTA Nº 112-2018 JOEP20180614
              pMtrBoletaPago(4, i) = Format(rsMantenimientoSinConvenioBoletaPago!nMonto, "#,##0.00")
              rsMantenimientoSinConvenioBoletaPago.MoveNext
        Next i
        RSClose rsMantenimientoSinConvenioBoletaPago
        
 '------------------------------------------------------
     ReDim pMtrReciboHono(3, 0)
        For i = 1 To (rsMantenimientoSinConvenioReciboHonorarios.RecordCount)
            ReDim Preserve pMtrReciboHono(3, i)
              pMtrReciboHono(1, i) = rsMantenimientoSinConvenioReciboHonorarios!nAnio
              pMtrReciboHono(2, i) = rsMantenimientoSinConvenioReciboHonorarios!nMes
              pMtrReciboHono(3, i) = Format(rsMantenimientoSinConvenioReciboHonorarios!nMonto, "#,##0.00")
              rsMantenimientoSinConvenioReciboHonorarios.MoveNext
        Next i
        RSClose rsMantenimientoSinConvenioReciboHonorarios
        
      ReDim pMtrIfis(rsMantenimientoSinConvenioCuotasIfis.RecordCount, 4)
        i = 0
        Do While Not rsMantenimientoSinConvenioCuotasIfis.EOF
            pMtrIfis(i, 0) = rsMantenimientoSinConvenioCuotasIfis!nNroCuota
            pMtrIfis(i, 1) = rsMantenimientoSinConvenioCuotasIfis!CDescripcion
            pMtrIfis(i, 2) = Format(IIf(IsNull(rsMantenimientoSinConvenioCuotasIfis!nMonto), 0, rsMantenimientoSinConvenioCuotasIfis!nMonto), "#0.00")
            rsMantenimientoSinConvenioCuotasIfis.MoveNext
            i = i + 1
        Loop
        RSClose rsMantenimientoSinConvenioCuotasIfis
        
      ReDim pMtrNegocio(3, 0)
        For i = 1 To (rsMantenimientoSinConvenioGastoNegocio.RecordCount)
            ReDim Preserve pMtrNegocio(3, i)
              pMtrNegocio(0, i) = rsMantenimientoSinConvenioGastoNegocio!nCodGasto
              pMtrNegocio(1, i) = rsMantenimientoSinConvenioGastoNegocio!cConsDescripcion
              pMtrNegocio(2, i) = Format(rsMantenimientoSinConvenioGastoNegocio!nMonto, "#,##0.00")
              rsMantenimientoSinConvenioGastoNegocio.MoveNext
        Next i
        RSClose rsMantenimientoSinConvenioGastoNegocio
        
    If Not (rsMantenimientoSinConvenioGastoNegocioVentasCosto.EOF And rsMantenimientoSinConvenioGastoNegocioVentasCosto.BOF) Then
        For i = 1 To (rsMantenimientoSinConvenioGastoNegocioVentasCosto.RecordCount)
                       
            pnIngNegocio = Format(rsMantenimientoSinConvenioGastoNegocioVentasCosto!nIngNegocio, "#,##0.00")
            pnEgrVenta = Format(rsMantenimientoSinConvenioGastoNegocioVentasCosto!nEgrVenta, "#,##0.00")
            pnMargBruto = Format(rsMantenimientoSinConvenioGastoNegocioVentasCosto!nMargBruto, "#,##0.00")
            pnIngNeto = Format(rsMantenimientoSinConvenioGastoNegocioVentasCosto!nIngNetoNegocio, "#,##0.00")
                           
            rsMantenimientoSinConvenioGastoNegocioVentasCosto.MoveNext
        Next i
        RSClose rsMantenimientoSinConvenioGastoNegocioVentasCosto
    End If
End Function

Public Sub Calcular(ByVal pnTipo As Integer)
    Dim nTotalGastos As Currency
    Dim nIngNeta As Currency
    Dim nExcedente As Currency
    Dim nCapPago As Currency
    
    nTotalGastos = 0
    nIngNeta = 0
    nExcedente = 0
    nCapPago = 0
    
    nTotalGastos = fgEgresos.SumaRow(3) - val(Replace(fgEgresos.TextMatrix(5, 3), ",", "")) - val(Replace(fgEgresos.TextMatrix(7, 3), ",", "")) - val(Replace(fgEgresos.TextMatrix(8, 3), ",", ""))
    nIngNeta = SumarCampo(fgIngresos, 3)
    nExcedente = SumarCampo(fgIngresos, 3) - nTotalGastos
    
    txtRatioIngNeto.Text = Format(nIngNeta, "#,##0.00")
    'txtRatioExcedente.Text = Format(nExcedente, "#,##0.00") 'CTI320200110 ERS003-2020. Comentó
    txtRatioExcedente.Text = Format(0, "#,##0.00") 'CTI320200110 ERS003-2020. Agregó
    
End Sub

'Validar Controles TAB Propuesta de Credito
Private Sub txtdFechaVisita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtEntornoCliente
    End If
End Sub

Private Sub txtdFechaVisita_LostFocus()
'If Not IsDate(txtdFechaVisita) Then
'    MsgBox "Verifique Dia,Mes,Año , Fecha Incorrecta", vbInformation, "Aviso"
'    'txtdFechaVisita.SetFocus
'End If
End Sub

Private Sub txtEntornoCliente_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtGiroNegocio
    End If
End Sub

Private Sub txtGiroNegocio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtExpCrediticia
    End If
End Sub

Private Sub txtExpCrediticia_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtFormNegocio
    End If
End Sub

Private Sub txtFormNegocio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
     If KeyAscii = 13 Then
        EnfocaControl txtColaGarantias
    End If
End Sub

Private Sub txtColaGarantias_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtImpactoMismo
    End If
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAgregarRef
    End If
End Sub

'Validar Control
Private Function Validar() As Boolean
Dim i As Integer
Dim j As Integer
Dim nCon As Currency
Dim lsFecha As String
Dim lsMensajeIfi As String 'LUCV20161115
Validar = True

    'Informacion del Cliente
    If TipoAportacion = 0 Then
        MsgBox "Seleccione Tipo de Aportacion", vbInformation, "Aviso"
        SSTabInfoNego.Tab = 0
        optTipoAportacion(1).SetFocus
        Validar = False
        Exit Function
    End If
    If txtNDependientes.Text = "" Then
        MsgBox "Ingrese N° de Dependientes", vbInformation, "Aviso"
        SSTabInfoNego.Tab = 0
        txtNDependientes.SetFocus
        Validar = False
        Exit Function
    End If
    'If txtNombreEmpleador.Text = "" Then 'comentado por JGPA20181129
    If (Me.txtCodPers.Text = "" And Me.txtNombreEmpleador.Text = "") Or (Me.txtCodPers.Text = "" And Me.txtNombreEmpleador.Text <> "") Then 'JGPA20181129 ACTA N° 192 - 2018
        MsgBox "Ingrese Dato del Empleador", vbInformation, "Aviso"
        SSTabInfoNego.Tab = 0
        txtCodPers.SetFocus
        Validar = False
        Exit Function
    End If
    
    'Ingresos
    nCon = 0
    For i = 1 To fgIngresos.rows - 1
        nCon = nCon + fgIngresos.TextMatrix(i, 3)
    Next i
    If nCon = 0 Then
        MsgBox "Ingrese Datos en INGRESOS", vbInformation, "Aviso"
        SSTab1.Tab = 0
        fgIngresos.SetFocus
        Validar = False
        Exit Function
    End If
    
    'Egresos
    nCon = 0
    For i = 1 To fgEgresos.rows - 1
        If fgEgresos.TextMatrix(i, 1) = 7 Or fgEgresos.TextMatrix(i, 1) = 8 Then
        Else
            nCon = nCon + fgEgresos.TextMatrix(i, 3)
        End If
    Next i
    If nCon = 0 Then
        MsgBox "Ingrese Datos en EGRESOS", vbInformation, "Aviso"
        SSTab1.Tab = 0
        fgEgresos.SetFocus
        Validar = False
        Exit Function
    End If
    
    If lnColocCondi <> 4 Then
        'Propuesta de Credito
        If txtdFechaVisita.Text = "__/__/____" Then
            MsgBox "Ingrese Fecha de Visita", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtdFechaVisita.SetFocus
            Validar = False
            Exit Function
        End If
            
        lsFecha = ValidaFecha(txtdFechaVisita)
        If Len(lsFecha) > 0 Then
            MsgBox lsFecha, vbInformation, "Aviso"
            SSTab1.Tab = 1
            EnfocaControl txtdFechaVisita
            fEnfoque txtdFechaVisita
            Validar = False
            Exit Function
        End If
        
        If txtEntornoCliente.Text = "" Then
            MsgBox "Ingrese Sobre el Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtEntornoCliente.SetFocus
            Validar = False
            Exit Function
        End If
        
        If txtGiroNegocio.Text = "" Then
            MsgBox "Ingrese Sobre el Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtGiroNegocio.SetFocus
            Validar = False
            Exit Function
        End If
         
         If txtExpCrediticia.Text = "" Then
            MsgBox "Ingrese Sobre la Experiencia Crediticia", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtExpCrediticia.SetFocus
            Validar = False
            Exit Function
        End If
           
           If txtFormNegocio.Text = "" Then
            MsgBox "Ingrese Sobre la Consistencia de la Informacion y la Formalidad del Negocio", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtFormNegocio.SetFocus
            Validar = False
            Exit Function
        End If
           
           If txtColaGarantias.Text = "" Then
            MsgBox "Ingrese Sobre los Colaterales o Garantias", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtColaGarantias.SetFocus
            Validar = False
            Exit Function
        End If
           
           If txtImpactoMismo.Text = "" Then
            MsgBox "Ingrese Sobre el Destino y el Impacto del Mismo", vbInformation, "Aviso"
            SSTab1.Tab = 1
            txtImpactoMismo.SetFocus
            Validar = False
            Exit Function
        End If
    End If
    
     'If lnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
        If Not fbTieneReferido6Meses Then
        'Comentario y Referidos
        If txtComentario.Text = "" Then
            MsgBox "Ingrese Datos en Comentarios", vbInformation, "Aviso"
            SSTab1.Tab = 2
            txtComentario.SetFocus
            Validar = False
            Exit Function
        End If
            
        If feReferidos.TextMatrix(1, 1) = "" Then
            MsgBox "Ingrese Datos en Referidos", vbInformation, "Aviso"
            SSTab1.Tab = 2
            feReferidos.SetFocus
            Validar = False
            Exit Function
        End If
        
        If feReferidos.rows - 1 < 2 Then
            MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
            SSTab1.Tab = 2
            cmdAgregarRef.SetFocus
            Validar = False
            Exit Function
        End If
        
        For i = 1 To feReferidos.rows - 1
            If feReferidos.TextMatrix(i, 2) = 0 Then
                MsgBox "DNI Incorrecto", vbInformation, "Alerta"
                SSTab1.Tab = 2
                feReferidos.SetFocus
                Validar = False
                Exit Function
            ElseIf feReferidos.TextMatrix(i, 3) = 0 Then
                MsgBox "Telefono Incorrecto", vbInformation, "Alerta"
                SSTab1.Tab = 2
                feReferidos.SetFocus
                Validar = False
                Exit Function
            End If
        Next i
                
        For i = 1 To feReferidos.rows - 1 'Verfica ambos DNI que no sean iguales
            For j = 1 To feReferidos.rows - 1
                If i <> j Then
                    If feReferidos.TextMatrix(i, 2) = feReferidos.TextMatrix(j, 2) Then
                        MsgBox "No se puede ingresar el mismo DNI mas de una vez en los referidos", vbInformation, "Alerta"
                        SSTab1.Tab = 2
                        feReferidos.SetFocus
                        Validar = False
                        Exit Function
                    End If
                End If
            Next
        Next
        
    End If

    'LUCV20161115, Agregó->Según ERS068-2016
     If Not ValidaIfiExisteCompraDeuda(fsCtaCod, pMtrIfis, Nothing, lsMensajeIfi, MatIfiNoSupervisadaGastoFami) Or Len(Trim(lsMensajeIfi)) > 0 Then
         MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
         SSTab1.Tab = 0
         Validar = False
         Exit Function
     End If
End Function
Private Sub txtImpactoMismo_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTab1.Tab = 2 'LUCV20171115, Agregó segun correo: RUSI
        If Not fbTieneReferido6Meses Then
            txtComentario.SetFocus
        End If
    End If
End Sub

Private Sub txtNDependientes_KeyPress(KeyAscii As Integer)
txtNDependientes.MaxLength = "2"
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        txtCodPers.SetFocus
    End If
End Sub

'RECO20160728 *******************************
Private Sub CargaRatios(ByVal psCtaCod As String)
    Dim oDCOMFormatosEval As New COMDCredito.DCOMFormatosEval
    Dim rsRatios As New ADODB.Recordset 'RECO20160728
    Dim rsRatiosAceptableCritico As ADODB.Recordset
        
    Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(fsCtaCod)
    
    Set rsRatios = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod) 'RECO20160728
    If Not (rsRatios.EOF And rsRatios.BOF) Then
        txtRatioCapPago.Text = rsRatios!nCapPagNeta * 100 & "%"
        txtRatioIngNeto.Text = Format(rsRatios!nIngreNeto, "#,##0.00")
        If (CDbl(rsRatios!nExceMensual) > 0) Then 'CTI320200110 ERS003-2020. Agregó condicional
            txtRatioExcedente.Text = Format(rsRatios!nExceMensual, "#,##0.00")
        End If
        
        If rsRatios!nCapPagNeta > 0 Then
            txtRatioCapPago.Visible = True
            txtRatioIngNeto.Visible = True
            txtRatioExcedente.Visible = True
            
            lblCapPag.Visible = True
                
            If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                    lblCapPagoCritico.Visible = False
                    lblCapPagAceptable.Visible = True
                Else
                    lblCapPagAceptable.Visible = False
                    lblCapPagoCritico.Visible = True
                End If
            End If
        End If
    End If
    
    RSClose rsRatiosAceptableCritico
    RSClose rsRatios
    
End Sub
'RECO FIN ***********************************

Private Function Registro()

    'si el cliente es nuevo-> se aciva referido obligatorio
    'If lnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
        
        feReferidos.Enabled = True
        cmdAgregarRef.Enabled = True
        cmdQuitar2.Enabled = True
        txtComentario.Enabled = True
    Else
        
        feReferidos.Enabled = False
        Frame9.Enabled = False
        Frame10.Enabled = False
        cmdAgregarRef.Enabled = False
        cmdQuitar2.Enabled = False
        txtComentario.Enabled = False
    End If
    'CTI3 ERS0032020******************************************
    If fnTipoRegMant = 1 Then
        'Carga de rsDatIfiGastoFami -> Matrix
        Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        
        Dim rsMantenimientoSinConvenioCuotasIfis As ADODB.Recordset
        'Set rsMantenimientoSinConvenioCuotasIfis = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioCuotasIfis(fsCtaCod)
        Set rsMantenimientoSinConvenioCuotasIfis = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami)
        ReDim pMtrIfis(rsMantenimientoSinConvenioCuotasIfis.RecordCount, 4)
        i = 0
        Do While Not rsMantenimientoSinConvenioCuotasIfis.EOF
            pMtrIfis(i, 0) = rsMantenimientoSinConvenioCuotasIfis!nNroCuota
            pMtrIfis(i, 1) = rsMantenimientoSinConvenioCuotasIfis!CDescripcion
            pMtrIfis(i, 2) = Format(IIf(IsNull(rsMantenimientoSinConvenioCuotasIfis!nMonto), 0, rsMantenimientoSinConvenioCuotasIfis!nMonto), "#0.00")
            rsMantenimientoSinConvenioCuotasIfis.MoveNext
            i = i + 1
        Loop
       
        rsMantenimientoSinConvenioCuotasIfis.Close
        Set rsMantenimientoSinConvenioCuotasIfis = Nothing
    End If
    '*********************************************************
End Function

Public Sub Consultar()
    
    optTipoAportacion(1).Enabled = False
    optTipoAportacion(2).Enabled = False
    optTipoAportacion(3).Enabled = False
    txtNDependientes.Enabled = False
    txtCodPers.Enabled = False
        
    lblCapPag.Visible = False
    txtRatioCapPago.Visible = False
    
    txtRatioIngNeto.Enabled = False
    txtRatioExcedente.Enabled = False

    txtdFechaVisita.Enabled = False
    txtEntornoCliente.Enabled = False
    txtGiroNegocio.Enabled = False
    txtExpCrediticia.Enabled = False
    txtFormNegocio.Enabled = False
    txtColaGarantias.Enabled = False
    txtImpactoMismo.Enabled = False

    txtComentario.Enabled = False
    feReferidos.Enabled = False

    cmdAgregarRef.Enabled = False
    cmdQuitar2.Enabled = False

    cmdInformeVista.Enabled = False
    cmdVerCar.Enabled = False
    cmdImprimir.Enabled = False
    cmdActualizarSinConvenio.Enabled = False
    cmdGuardar.Enabled = False
    
End Sub

'*
Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = False
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean)

    optTipoAportacion(1).Enabled = pbHabilitaA
    optTipoAportacion(2).Enabled = pbHabilitaA
    optTipoAportacion(3).Enabled = pbHabilitaA
    txtNDependientes.Enabled = pbHabilitaA
    txtCodPers.Enabled = pbHabilitaA
    
    fgIngresos.Enabled = pbHabilitaA
    fgEgresos.Enabled = pbHabilitaA
    
    txtdFechaVisita.Enabled = pbHabilitaA
    txtEntornoCliente.Enabled = pbHabilitaA
    txtGiroNegocio.Enabled = pbHabilitaA
    txtExpCrediticia.Enabled = pbHabilitaA
    txtFormNegocio.Enabled = pbHabilitaA
    txtColaGarantias.Enabled = pbHabilitaA
    txtImpactoMismo.Enabled = pbHabilitaA
    
    txtComentario.Enabled = pbHabilitaA
          
    cmdGuardar.Enabled = pbHabilitaA
    cmdActualizarSinConvenio.Enabled = pbHabilitaA
    
End Function
