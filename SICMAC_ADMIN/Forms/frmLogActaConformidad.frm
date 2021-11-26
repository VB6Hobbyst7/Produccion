VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogActaConformidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acta de Conformidad Digital"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   Icon            =   "frmLogActaConformidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabActivaBien 
      Height          =   3525
      Left            =   45
      TabIndex        =   12
      Top             =   4440
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   6218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Acta"
      TabPicture(0)   =   "frmLogActaConformidad.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame20"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancelar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdConforme"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame18"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame17"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame17 
         Caption         =   "Moneda"
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
         Left            =   5040
         TabIndex        =   44
         Top             =   480
         Width           =   1185
         Begin VB.TextBox txtMoneda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   95
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Doc. Referencia"
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
         Left            =   6240
         TabIndex        =   43
         Top             =   480
         Width           =   1785
         Begin VB.TextBox txtDocReferencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "N° Acta"
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
         Left            =   8040
         TabIndex        =   42
         Top             =   480
         Width           =   1815
         Begin VB.TextBox txtActaNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdConforme 
         Caption         =   "&Conforme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   8600
         TabIndex        =   8
         Top             =   2580
         Width           =   1290
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   8600
         TabIndex        =   9
         Top             =   2930
         Width           =   1290
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Área"
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
         Left            =   40
         TabIndex        =   39
         Top             =   480
         Width           =   4930
         Begin VB.TextBox txtSubAreaDescripcion 
            Height          =   285
            Left            =   2520
            MaxLength       =   235
            TabIndex        =   5
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtAreaAgeNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1250
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtAreaAgeCod 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label67 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   41
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Área:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Datos del Proveedor"
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
         Height          =   1035
         Left            =   40
         TabIndex        =   30
         Top             =   1155
         Width           =   9810
         Begin VB.TextBox txtProveedorCtaInstitucionCod 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtProveedorCtaInstitucionNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7000
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   600
            Width           =   2735
         End
         Begin VB.TextBox txtProveedorCtaMoneda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtProveedorCtaNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtProveedorNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtProveedorCod 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtProveedorDocNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtProveedorDocTpo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label73 
            Caption         =   "Institución:"
            Height          =   255
            Left            =   4800
            TabIndex        =   38
            Top             =   615
            Width           =   855
         End
         Begin VB.Label Label72 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   3000
            TabIndex        =   37
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label71 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   620
            Width           =   615
         End
         Begin VB.Label Label70 
            Caption         =   "N° Doc.:"
            Height          =   255
            Left            =   7320
            TabIndex        =   35
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label69 
            Caption         =   "Tipo Doc.:"
            Height          =   255
            Left            =   5280
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label68 
            Caption         =   "Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Datos de Compra"
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
         Height          =   1215
         Left            =   40
         TabIndex        =   10
         Top             =   2280
         Width           =   8505
         Begin VB.TextBox txtCompraObservacion 
            Height          =   405
            Left            =   1080
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   720
            Width           =   7335
         End
         Begin VB.TextBox txtCompraDescripcion 
            Height          =   405
            Left            =   1080
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   240
            Width           =   7335
         End
         Begin VB.Label Label76 
            Caption         =   "Observa.:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label79 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   17
         Top             =   3465
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
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
         Left            =   -73185
         TabIndex        =   16
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   15
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label lblSolesAho 
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
         Left            =   -70815
         TabIndex        =   14
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   -67680
         TabIndex        =   13
         Top             =   3375
         Width           =   2145
      End
   End
   Begin TabDlg.SSTab TabBuscar 
      Height          =   4355
      Left            =   45
      TabIndex        =   18
      Top             =   40
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   7673
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "frmLogActaConformidad.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Bienes/Servicios a dar Conformidad"
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
         Height          =   3190
         Left            =   40
         TabIndex        =   26
         Top             =   1080
         Width           =   9825
         Begin VB.CommandButton cmdCancelarDarConformidad 
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   1800
            TabIndex        =   4
            Top             =   2790
            Width           =   1050
         End
         Begin VB.CommandButton cmdDarConformidad 
            Caption         =   "&Dar Conformidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   120
            TabIndex        =   3
            Top             =   2790
            Width           =   1650
         End
         Begin Sicmact.FlexEdit feContrato 
            Height          =   2475
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   9600
            _ExtentX        =   16933
            _ExtentY        =   4366
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-cCtaContCod--N° de Pago-Fecha de Pago-Moneda-Monto-Tipo-Estado"
            EncabezadosAnchos=   "0-0-450-1000-1200-1000-1200-1800-1800"
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
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-2-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-4-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-C-C-C-R-C-L"
            FormatosEdit    =   "0-0-0-0-0-0-2-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   7
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit feOrden 
            Height          =   2535
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   9615
            _ExtentX        =   16933
            _ExtentY        =   4366
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-nMovNro-nMovItem-cCtaContCod--Ag.Destino-Objeto-Descripcion-Unidad-Solicitado-P.Unitario-SubTotal"
            EncabezadosAnchos=   "0-0-0-0-450-1000-1400-2000-1000-900-1100-1100"
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
            ColumnasAEditar =   "X-X-X-X-4-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-4-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-L-C-C-R-R"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-2-0-2-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   7
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Documento Origen"
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
         Left            =   40
         TabIndex        =   24
         Top             =   380
         Width           =   9825
         Begin VB.ComboBox cboTpoDocOrigen 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "&Mostrar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Left            =   8640
            TabIndex        =   2
            Top             =   230
            Width           =   1050
         End
         Begin Sicmact.TxtBuscar txtDocumentoCod 
            Height          =   315
            Left            =   3600
            TabIndex        =   1
            Top             =   240
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin VB.Label lblDocumentoNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5400
            TabIndex        =   45
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Doc. Origen:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   -67680
         TabIndex        =   23
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label9 
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
         Left            =   -70815
         TabIndex        =   22
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   21
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
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
         Left            =   -73185
         TabIndex        =   20
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   19
         Top             =   3465
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmLogActaConformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'** Nombre : frmLogActaConformidad
'** Descripción : Registro de Acta de Conformidad creado segun ERS062-2013
'** Creación : EJVG, 20131009 09:00:00 AM
'*************************************************************************
Option Explicit
Dim gsOpeCod As String
Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double
Dim fnTpoDocOrigen As Integer
Dim fnMovNroOCS As Long
Dim fsNContrato As String
Dim fnDocTpo As Integer
Dim fsAreaAgeCod As String
Dim fnMoneda As Moneda

Dim fRsOCompra As New ADODB.Recordset
Dim fRsOServicio As New ADODB.Recordset
Dim fRsContratoCompra As New ADODB.Recordset
Dim fRsContratoServicio As New ADODB.Recordset
Dim fsCtaContCodProv As String
Dim fsCtaContCodOC As String, fsCtaContCodOS As String
Dim fnTpoCambio As Currency
Dim fbGraboActa As Boolean

Private Sub Form_Load()
    fsAreaAgeCod = gsCodArea & Right(gsCodAge, 2)
    fnFormTamanioIni = 4815
    fnFormTamanioActiva = 8490
    Height = fnFormTamanioIni
    CargaControles
    CargaVariables
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer) 'Habilitar KeyPreview=True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not fbGraboActa Then
        If MsgBox("¿Desea salir sin grabar el Acta de Conformidad?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
           Cancel = 1
           Exit Sub
        End If
    End If
    Set fRsOCompra = Nothing
    Set fRsOServicio = Nothing
    Set fRsContratoCompra = Nothing
    Set fRsContratoServicio = Nothing
End Sub
Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    gsOpeCod = psOpeCod
    fnMoneda = Mid(gsOpeCod, 3, 1)
    Caption = UCase(psOpeDesc)
    Show 1
End Sub
Private Sub CargaControles()
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Set rs = oLog.ListaTpoDocOrigenActaConformidad("1,2,3,4")

    cboTpoDocOrigen.Clear
    CargaCombo rs, cboTpoDocOrigen, , 1, 0
    
    Set rs = Nothing
    Set oLog = Nothing
End Sub
Private Sub CargaVariables()
    Dim odoc As New DOperacion
    Dim oConstSist As New NConstSistemas
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    
    fbGraboActa = False
    If gbBitTCPonderado Then
        fnTpoCambio = gnTipCambioPonderado
    Else
        fnTpoCambio = gnTipCambioC
    End If
    
    If fnMoneda = gMonedaNacional Then
        fsCtaContCodOC = oConstSist.LeeConstSistema(gsLogCtasContOCMN)
        fsCtaContCodOS = oConstSist.LeeConstSistema(gsLogCtasContOSMN)
    Else
        fsCtaContCodOC = oConstSist.LeeConstSistema(gsLogCtasContOCME)
        fsCtaContCodOS = oConstSist.LeeConstSistema(gsLogCtasContOSME)
    End If

    Set rs = odoc.CargaOpeCta(gsOpeCod, "H")
    fsCtaContCodProv = rs!cCtaContCod

    Set fRsOCompra = oLog.ListaOrdenCompraxActaConformidad(fsAreaAgeCod, fsCtaContCodOC)
    Set fRsOServicio = oLog.ListaOrdenServicioxActaConformidad(fsAreaAgeCod, fsCtaContCodOS)
    'Set fRsContratoCompra = oLog.ListaContratoxActaConformidad(fsAreaAgeCod, 1, fnMoneda) 'Comentado PASIERS0772014
    'Set fRsContratoServicio = oLog.ListaContratoxActaConformidad(fsAreaAgeCod, 2, fnMoneda) 'Comentado PASIERS0772014
    
    Set rs = Nothing
    Set oLog = Nothing
    Set oConstSist = Nothing
    Set odoc = Nothing
End Sub
Private Sub cboTpoDocOrigen_Click()
    Dim lnTpoDoc As Integer
    On Error GoTo ErrCboTpoDocOrigen
    
    Screen.MousePointer = 11
    cancela_busqueda_actual
    If Trim(Right(cboTpoDocOrigen.Text, 4)) <> "" Then
        lnTpoDoc = CInt(Trim(Right(cboTpoDocOrigen.Text, 4)))
        If lnTpoDoc = LogTipoDocOrigenActaConformidad.OrdenCompra Or lnTpoDoc = LogTipoDocOrigenActaConformidad.OrdenServicio Then
            feOrden.Visible = True
            feContrato.Visible = False
        ElseIf lnTpoDoc = LogTipoDocOrigenActaConformidad.ContratoCompra Or lnTpoDoc = LogTipoDocOrigenActaConformidad.ContratoServicio Then
            feOrden.Visible = False
            feContrato.Visible = True
        End If
    End If
    Screen.MousePointer = 0
    fnTpoDocOrigen = lnTpoDoc
    Exit Sub
ErrCboTpoDocOrigen:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cboTpoDocOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDocumentoCod.SetFocus
    End If
End Sub
Private Sub txtCompraDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCompraObservacion.SetFocus
    End If
End Sub
Private Sub txtCompraObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdConforme.SetFocus
    End If
End Sub
Private Sub txtDocumentoCod_EmiteDatos()
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Dim oBusca As New frmLogActaConformidadBusca
    Dim lsDato As String, lsDocNro As String, lsDocNombre As String
    cancela_busqueda_actual
    If fnTpoDocOrigen <> 0 Then
        If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Then
            Set rs = fRsOCompra.Clone
        ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
            Set rs = fRsOServicio.Clone
        ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Then
            Set rs = fRsContratoCompra.Clone
        ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
            Set rs = fRsContratoServicio.Clone
        End If
        oBusca.Inicio fnTpoDocOrigen, lsDato, lsDocNro, lsDocNombre, rs

        txtDocumentoCod.Text = lsDocNro
        lblDocumentoNombre.Caption = lsDocNombre
        If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
            fnMovNroOCS = CLng(IIf(lsDato = "", 0, lsDato))
        ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
            fsNContrato = lsDato
        End If
    End If
    Set oBusca = Nothing
    Set rs = Nothing
    Set oLog = Nothing
    If lsDato <> "" Then
        cmdMostrar.SetFocus
    End If
End Sub
Private Sub cmdMostrar_Click()
    If Not validaSeleccionDocumento Then Exit Sub
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Dim row As Long
    
    On Error GoTo ErrMostrar
    Screen.MousePointer = 11
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        LimpiaFlex feOrden
        If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Then
            Set rs = oLog.ListaOrdenCompraDetxActaConformidad(fnMovNroOCS, fsAreaAgeCod, fsCtaContCodOC)
        Else
            Set rs = oLog.ListaOrdenservicioDetxActaConformidad(fnMovNroOCS, fsAreaAgeCod, fsCtaContCodOS)
        End If
        Do While Not rs.EOF
            feOrden.AdicionaFila
            row = feOrden.row
            feOrden.TextMatrix(row, 1) = rs!nMovNro
            feOrden.TextMatrix(row, 2) = rs!nMovItem
            feOrden.TextMatrix(row, 3) = rs!cCtaContCod
            feOrden.TextMatrix(row, 5) = rs!cAgeCod
            feOrden.TextMatrix(row, 6) = rs!cObjeto
            feOrden.TextMatrix(row, 7) = rs!cDescripcion
            feOrden.TextMatrix(row, 8) = rs!cUnidad
            feOrden.TextMatrix(row, 9) = rs!nSolicitado
            feOrden.TextMatrix(row, 10) = Format(rs!nPrecioUnitario, gsFormatoNumeroView)
            feOrden.TextMatrix(row, 11) = Format(rs!nSubTotal, gsFormatoNumeroView)
            rs.MoveNext
        Loop
        feOrden.SetFocus
        SendKeys "{Right}"
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        LimpiaFlex feContrato
        Set rs = oLog.ListaContratoDetxActaConformidad(fsNContrato)
        Do While Not rs.EOF
            feContrato.AdicionaFila
            row = feContrato.row
            feContrato.TextMatrix(row, 1) = rs!cCtaContCod
            feContrato.TextMatrix(row, 3) = rs!nNPago
            feContrato.TextMatrix(row, 4) = Format(rs!dFecPago, gsFormatoFechaView)
            feContrato.TextMatrix(row, 5) = rs!cMoneda
            feContrato.TextMatrix(row, 6) = Format(rs!nMonto, gsFormatoNumeroView)
            feContrato.TextMatrix(row, 7) = rs!cTipo
            feContrato.TextMatrix(row, 8) = rs!cEstado
            rs.MoveNext
        Loop
        feContrato.SetFocus
        SendKeys "{Right}"
    End If
    Screen.MousePointer = 0
    Set rs = Nothing
    Set oLog = Nothing
    Exit Sub
ErrMostrar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOrden_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feOrden.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub feOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDarConformidad.SetFocus
    End If
End Sub
Private Sub feContrato_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feContrato.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub feContrato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDarConformidad.SetFocus
    End If
End Sub
Private Sub feContrato_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim row As Long
    Dim bNOCheck As Boolean
    'Validamos que las cuotas a pagar sean próximas
    If pnRow > 0 Then
        If feContrato.TextMatrix(pnRow, 2) = "." Then
            'Verifica que las anteriores esten chequeadas
            bNOCheck = False
            For row = 1 To pnRow - 1
                If feContrato.TextMatrix(row, 2) <> "." Then
                    bNOCheck = True
                    Exit For
                End If
            Next
            If bNOCheck Then
                feContrato.TextMatrix(pnRow, 2) = ""
                Exit Sub
            End If
        Else
            'Deseleccionamos cuotas en adelante
            For row = pnRow + 1 To Me.feContrato.Rows - 1
                feContrato.TextMatrix(row, 2) = ""
            Next
        End If
    End If
End Sub
Private Sub cmdDarConformidad_Click()
    Dim row As Long
    Dim bSelecciona As Boolean
    Dim oLog As DLogGeneral
    Dim oMov As DMov
    Dim rs As New ADODB.Recordset
    Dim lsCorrelativo As String
    
    On Error GoTo ErrcmdDarConformidad
    If Not validaSeleccionDocumento Then Exit Sub
    '*** Valida seleccion de Registros
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        For row = 1 To feOrden.Rows - 1
            If feOrden.TextMatrix(row, 4) = "." Then
                bSelecciona = True
                Exit For
            End If
        Next
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        For row = 1 To feContrato.Rows - 1
            If feContrato.TextMatrix(row, 2) = "." Then
                bSelecciona = True
                Exit For
            End If
        Next
    End If
    If Not bSelecciona Then
        MsgBox "Ud. debe seleccionar los registros a dar Conformidad", vbInformation, "Aviso"
        Exit Sub
    End If
    '*** Valida tenga cuentas contables
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        For row = 1 To feOrden.Rows - 1
            If feOrden.TextMatrix(row, 4) = "." Then
                If Len(Trim(feOrden.TextMatrix(row, 3))) = 0 Then
                    MsgBox "El Objeto " & feOrden.TextMatrix(row, 7) & Chr(10) & "no tiene Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                    feOrden.TopRow = row
                    feOrden.row = row
                    feOrden.col = 3
                    Exit Sub
                End If
            End If
        Next
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        For row = 1 To feContrato.Rows - 1
            If feContrato.TextMatrix(row, 2) = "." Then
                If Len(Trim(feContrato.TextMatrix(row, 1))) = 0 Then
                    MsgBox "La Cuota N° " & Format(feContrato.TextMatrix(row, 3), "00") & " del contrato " & fsNContrato & Chr(10) & "no tiene Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                    feContrato.TopRow = row
                    feContrato.row = row
                    feContrato.col = 3
                    Exit Sub
                End If
            End If
        Next
    End If
    '***
    
    Screen.MousePointer = 11
    
    Height = fnFormTamanioActiva
    CentraForm Me
    TabBuscar.Enabled = False
    LimpiarDatosActaConformidad
    
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Then
        fnDocTpo = LogTipoActaConformidad.gActaRecepcionBienes
        TabActivaBien.TabCaption(0) = "Acta de Bien"
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        fnDocTpo = LogTipoActaConformidad.gActaConformidadServicio
        TabActivaBien.TabCaption(0) = "Acta de Servicio"
    End If

    Set oLog = New DLogGeneral
    Set oMov = New DMov
    Set rs = New ADODB.Recordset
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        Set rs = oLog.OCSxActaConformidad(fnMovNroOCS)
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        Set rs = oLog.ContratoxActaConformidad(fsNContrato)
    End If
    
    lsCorrelativo = oMov.GetCorrelativoActaConformidad(fnDocTpo, Right(gsCodAge, 2), CStr(Year(gdFecSis)))
    If Not rs.EOF Then
        EstablecerDatosActaConformidad rs!cAreaAgeCod, rs!cAreaAgeDesc, rs!cMoneda, rs!cDocReferencia, lsCorrelativo, rs!cProveedorCod, rs!cProveedorNombre, rs!cDocTpo, rs!cDocNro, rs!cCtaCodAhorro, IIf(rs!cCtaCodAhorro <> "", rs!cMoneda, ""), rs!cInstitucionCod, rs!cInstitucionNombre, rs!cMovDesc
    End If
    txtSubAreaDescripcion.SetFocus
    
    Screen.MousePointer = 0
    Set rs = Nothing
    Set oLog = Nothing
    Set oMov = Nothing
    Exit Sub
ErrcmdDarConformidad:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelarDarConformidad_Click()
    cboTpoDocOrigen.ListIndex = -1
    cancela_busqueda_actual
End Sub
Private Sub txtSubAreaDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCompraDescripcion.Visible And txtCompraDescripcion.Enabled Then
            txtCompraDescripcion.SetFocus
        End If
    End If
End Sub
Private Sub cmdConforme_Click()
    Dim oLog As NLogGeneral
    Dim oAsiento As NContImprimir
    Dim oPrevio As clsPrevio
    Dim lnMovNro As Long
    Dim lsNroActaConformidad As String
    Dim DatosOrden() As TActaConformidadOrden
    Dim DatosContrato() As TActaConformidadContrato
    Dim Index As Integer, indexMat As Integer, indexObj As Integer
    Dim lsSubCta As String
    Dim lsMovNro As String

    On Error GoTo ErrCmdConforme
    
    If Len(Trim(txtAreaAgeCod.Text)) = 0 Then
        MsgBox "La presente conformidad no cuenta con Área Agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtActaNro.Text)) = 0 Then
        MsgBox "No se ha conseguido el correlativo del Nro de Acta", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtProveedorCod.Text)) = 0 Then
        MsgBox "No se cuenta con Proveedor en la presente operación", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtCompraDescripcion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar una descripción", vbInformation, "Aviso"
        txtCompraDescripcion.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtCompraObservacion.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la observación respectiva", vbInformation, "Aviso"
        txtCompraObservacion.SetFocus
        Exit Sub
    End If

    lsNroActaConformidad = txtActaNro.Text
    indexMat = 0
    ReDim DatosOrden(indexMat)
    ReDim DatosContrato(indexMat)
    
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        For Index = 1 To feOrden.Rows - 1
            If feOrden.TextMatrix(Index, 4) = "." Then
                indexMat = indexMat + 1
                ReDim Preserve DatosOrden(indexMat)
                DatosOrden(indexMat).nMovItem = CInt(feOrden.TextMatrix(Index, 2))
                DatosOrden(indexMat).sCtaContCod = CStr(Trim(feOrden.TextMatrix(Index, 3)))
                DatosOrden(indexMat).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 6)))
                DatosOrden(indexMat).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 7)))
                DatosOrden(indexMat).nCantidad = Val(feOrden.TextMatrix(Index, 9))
                DatosOrden(indexMat).nTotal = feOrden.TextMatrix(Index, 11)
            End If
        Next
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        For Index = 1 To feContrato.Rows - 1
            If feContrato.TextMatrix(Index, 2) = "." Then
                indexMat = indexMat + 1
                ReDim Preserve DatosContrato(indexMat)
                DatosContrato(indexMat).nNPago = CInt(feContrato.TextMatrix(Index, 3))
                DatosContrato(indexMat).sCtaContCod = CStr(Trim(feContrato.TextMatrix(Index, 1)))
                DatosContrato(indexMat).sDescripcion = "CUOTA N° " & Format(DatosContrato(indexMat).nNPago, "00") & " DE " & UCase(feContrato.TextMatrix(Index, 7))
                DatosContrato(indexMat).nMonto = feContrato.TextMatrix(Index, 6)
            End If
        Next
    End If
    
    If indexMat = 0 Then
        MsgBox "No existen Items a dar conformidad", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Then
        fsCtaContCodProv = fsCtaContCodOC
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        fsCtaContCodProv = fsCtaContCodOS
    End If
    
    If fsCtaContCodProv = "" Then
        MsgBox "No se ha definido cuenta contable de Proveedor, consulte al Dpto. de TI", vbInformation, "Aviso"
        Exit Sub
    End If
        
    If MsgBox("¿Esta seguro de guardar el Acta de Conformidad Digital?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oLog = New NLogGeneral
        
    Screen.MousePointer = 11
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        lnMovNro = oLog.GrabarActaConformidad_Orden(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, fnTpoDocOrigen, fnDocTpo, txtDocReferencia.Text, _
                                                    txtAreaAgeCod.Text, Trim(txtSubAreaDescripcion.Text), fnMoneda, lsNroActaConformidad, txtProveedorCod.Text, _
                                                    Trim(Me.txtCompraDescripcion.Text), Trim(txtCompraObservacion.Text), DatosOrden, fsCtaContCodProv, fnTpoCambio, _
                                                    lsMovNro, fnMovNroOCS)
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        lnMovNro = oLog.GrabarActaConformidad_Contrato(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, fnTpoDocOrigen, fnDocTpo, fsNContrato, _
                                                    txtAreaAgeCod.Text, Trim(txtSubAreaDescripcion.Text), fnMoneda, lsNroActaConformidad, txtProveedorCod.Text, _
                                                    Trim(Me.txtCompraDescripcion.Text), Trim(txtCompraObservacion.Text), DatosContrato, fsCtaContCodProv, fnTpoCambio, _
                                                    lsMovNro)
    End If
    Screen.MousePointer = 0
    
    If lnMovNro = 0 Then
        MsgBox "Ha ocurrido un error al registrar el Acta de Conformidad", vbCritical, "Aviso"
        oLog = Nothing
        Exit Sub
    End If
    fbGraboActa = True
    
    Set oAsiento = New NContImprimir
    Set oPrevio = New clsPrevio
        
    MsgBox "Se ha registrado el Acta de Conformidad Nro. " & lsNroActaConformidad & " con éxito", vbInformation, "Aviso"
    ImprimeActaConformidadPDF lnMovNro
    'oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, Caption), Caption, True
    
    cmdCancelarDarConformidad_Click
    cmdCancelar_Click
    
    Set oAsiento = Nothing
    Set oPrevio = Nothing
    Set oLog = Nothing
    
    If MsgBox("¿Desea registrar otra Acta de Conformidad?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        CargaVariables
    Else
        Unload Me
    End If
    Exit Sub
ErrCmdConforme:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    Height = fnFormTamanioIni
    TabBuscar.Enabled = True
End Sub
Private Function validaSeleccionDocumento() As Boolean
    validaSeleccionDocumento = True
    If cboTpoDocOrigen.ListIndex = -1 Then
        MsgBox "Ud. primero debe de seleccionar el Tipo de Documento Origen", vbInformation, "Aviso"
        validaSeleccionDocumento = False
        cboTpoDocOrigen.SetFocus
        Exit Function
    End If
    If Len(Trim(txtDocumentoCod.Text)) = 0 Then
        MsgBox "Ud. debe de seleccionar Documento", vbInformation, "Aviso"
        validaSeleccionDocumento = False
        txtDocumentoCod.SetFocus
        Exit Function
    End If
End Function
Private Sub cancela_busqueda_actual()
    Height = fnFormTamanioIni
    txtDocumentoCod.Text = ""
    lblDocumentoNombre.Caption = ""
    Call FormateaFlex(feOrden)
    Call FormateaFlex(feContrato)
End Sub
Private Sub LimpiarDatosActaConformidad()
    txtAreaAgeCod.Text = ""
    txtAreaAgeNombre.Text = ""
    txtSubAreaDescripcion.Text = ""
    txtMoneda.Text = ""
    txtDocReferencia.Text = ""
    txtActaNro.Text = ""
    txtProveedorCod.Text = ""
    txtProveedorNombre.Text = ""
    txtProveedorDocTpo.Text = ""
    txtProveedorDocNro.Text = ""
    txtProveedorCtaNro.Text = ""
    txtProveedorCtaMoneda.Text = ""
    txtProveedorCtaInstitucionCod.Text = ""
    txtProveedorCtaInstitucionNombre.Text = ""
    txtCompraDescripcion.Text = ""
    txtCompraObservacion.Text = ""
End Sub
Private Sub EstablecerDatosActaConformidad(Optional ByVal psAreaAgeCod As String = "", Optional ByVal psAreaAgeNombre As String = "", _
                                            Optional ByVal psMoneda As String = "", Optional ByVal psDocReferencia As String = "", _
                                            Optional ByVal psActaNro As String = "", Optional ByVal psProveedorCod As String = "", _
                                            Optional ByVal psProveedorNombre As String = "", Optional ByVal psProveedorDocTpo As String = "", _
                                            Optional ByVal psProveedorDocNro As String = "", Optional ByVal psProveedorCtaNro As String = "", _
                                            Optional ByVal psProveedorCtaMoneda As String = "", Optional ByVal psProveedorCtaInstitucionCod As String = "", _
                                            Optional ByVal psProveedorCtaInstitucionNombre As String = "", Optional ByVal psCompraDescripcion As String = "", _
                                            Optional ByVal psCompraObservacion As String = "")
    If psAreaAgeCod <> "" Then
        txtAreaAgeCod.Text = psAreaAgeCod
    End If
    If psAreaAgeNombre <> "" Then
        txtAreaAgeNombre.Text = psAreaAgeNombre
    End If
    If psMoneda <> "" Then
        txtMoneda.Text = psMoneda
    End If
    If psDocReferencia <> "" Then
        txtDocReferencia.Text = psDocReferencia
    End If
    If psActaNro <> "" Then
        txtActaNro.Text = psActaNro
    End If
    If psProveedorCod <> "" Then
        txtProveedorCod.Text = psProveedorCod
    End If
    If psProveedorNombre <> "" Then
        txtProveedorNombre.Text = psProveedorNombre
    End If
    If psProveedorDocTpo <> "" Then
        txtProveedorDocTpo.Text = psProveedorDocTpo
    End If
    If psProveedorDocNro <> "" Then
        txtProveedorDocNro.Text = psProveedorDocNro
    End If
    If psProveedorCtaNro <> "" Then
        txtProveedorCtaNro.Text = psProveedorCtaNro
    End If
    If psProveedorCtaMoneda <> "" Then
        txtProveedorCtaMoneda.Text = psProveedorCtaMoneda
    End If
    If psProveedorCtaInstitucionCod <> "" Then
        txtProveedorCtaInstitucionCod.Text = psProveedorCtaInstitucionCod
    End If
    If psProveedorCtaInstitucionNombre <> "" Then
        txtProveedorCtaInstitucionNombre.Text = psProveedorCtaInstitucionNombre
    End If
    If psCompraDescripcion <> "" Then
        txtCompraDescripcion.Text = psCompraDescripcion
    End If
    If psCompraObservacion <> "" Then
        txtCompraObservacion.Text = psCompraObservacion
    End If
End Sub

