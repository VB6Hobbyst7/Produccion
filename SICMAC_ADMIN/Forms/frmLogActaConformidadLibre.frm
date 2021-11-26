VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogActaConformidadLibre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acta de Conformidad Digital Libre"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   Icon            =   "frmLogActaConformidadLibre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit feObj 
      Height          =   1575
      Left            =   10080
      TabIndex        =   55
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Id-Objeto Orden-CtaContCod-CtaContDesc-Filtro-CodObjeto"
      EncabezadosAnchos=   "0-400-800-800-800-800-800"
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
      ColumnasAEditar =   "X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin TabDlg.SSTab TabActivaBien 
      Height          =   3525
      Left            =   40
      TabIndex        =   13
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
      TabPicture(0)   =   "frmLogActaConformidadLibre.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdConforme"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame20"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
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
         TabIndex        =   36
         Top             =   2235
         Width           =   8505
         Begin VB.TextBox txtCompraDescripcion 
            Height          =   405
            Left            =   1080
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   240
            Width           =   7335
         End
         Begin VB.TextBox txtCompraObservacion 
            Height          =   405
            Left            =   1080
            MaxLength       =   1225
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   720
            Width           =   7335
         End
         Begin VB.Label Label79 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label76 
            Caption         =   "Observa.:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   855
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
         TabIndex        =   22
         Top             =   1155
         Width           =   9810
         Begin VB.TextBox txtProveedorDocTpo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtProveedorDocNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtProveedorNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtProveedorCtaNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtProveedorCtaMoneda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtProveedorCtaInstitucionNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7000
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   600
            Width           =   2735
         End
         Begin VB.TextBox txtProveedorCtaInstitucionCod 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   600
            Width           =   1335
         End
         Begin Sicmact.TxtBuscar txtProveedorCod 
            Height          =   315
            Left            =   1080
            TabIndex        =   8
            Top             =   240
            Width           =   1220
            _ExtentX        =   2143
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
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin VB.Label Label68 
            Caption         =   "Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label69 
            Caption         =   "Tipo Doc.:"
            Height          =   255
            Left            =   5280
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label70 
            Caption         =   "N° Doc.:"
            Height          =   255
            Left            =   7320
            TabIndex        =   33
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label71 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   620
            Width           =   615
         End
         Begin VB.Label Label72 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   3000
            TabIndex        =   31
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label73 
            Caption         =   "Institución:"
            Height          =   255
            Left            =   4800
            TabIndex        =   30
            Top             =   615
            Width           =   855
         End
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
         TabIndex        =   18
         Top             =   480
         Width           =   4930
         Begin VB.TextBox txtAreaAgeNombre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   1290
         End
         Begin VB.TextBox txtSubAreaDescripcion 
            Height          =   285
            Left            =   3000
            MaxLength       =   235
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin Sicmact.TxtBuscar txtAreaAgeCod 
            Height          =   315
            Left            =   600
            TabIndex        =   5
            Top             =   240
            Width           =   900
            _ExtentX        =   1588
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
         End
         Begin VB.Label Label1 
            Caption         =   "Área:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   270
            Width           =   375
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
            Left            =   2880
            TabIndex        =   20
            Top             =   240
            Width           =   135
         End
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
         TabIndex        =   12
         Top             =   2940
         Width           =   1290
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
         TabIndex        =   11
         Top             =   2580
         Width           =   1290
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
         TabIndex        =   16
         Top             =   480
         Width           =   1815
         Begin VB.TextBox txtActaNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   1575
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
         TabIndex        =   15
         Top             =   480
         Width           =   1785
         Begin VB.TextBox txtDocReferencia 
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
      End
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
         TabIndex        =   14
         Top             =   480
         Width           =   1185
         Begin VB.TextBox txtMoneda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   975
         End
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
         TabIndex        =   43
         Top             =   3375
         Width           =   2145
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
         TabIndex        =   42
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   41
         Top             =   3465
         Width           =   765
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
         TabIndex        =   40
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   39
         Top             =   3465
         Width           =   525
      End
   End
   Begin TabDlg.SSTab TabBuscar 
      Height          =   4355
      Left            =   40
      TabIndex        =   44
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
      TabPicture(0)   =   "frmLogActaConformidadLibre.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
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
         TabIndex        =   47
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
         Begin VB.Label Label3 
            Caption         =   "Tipo Doc. Origen:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   270
            Width           =   1335
         End
      End
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
         TabIndex        =   45
         Top             =   1080
         Width           =   9825
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "&Quitar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   8840
            TabIndex        =   4
            Top             =   2790
            Width           =   885
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   7920
            TabIndex        =   3
            Top             =   2790
            Width           =   885
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
            TabIndex        =   1
            Top             =   2790
            Width           =   1650
         End
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
            TabIndex        =   2
            Top             =   2790
            Width           =   1050
         End
         Begin Sicmact.FlexEdit feOrden 
            Height          =   2535
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   9615
            _ExtentX        =   16933
            _ExtentY        =   4366
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Ag.Destino-Objeto-Descripcion-Solicitado-P.Unitario-SubTotal-CtaContCod"
            EncabezadosAnchos=   "0-1000-1400-3500-900-1100-1100-0"
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
            ColumnasAEditar =   "X-1-2-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-1-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-R-L"
            FormatosEdit    =   "0-0-0-0-3-2-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   53
         Top             =   3465
         Width           =   525
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
         TabIndex        =   52
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   51
         Top             =   3465
         Width           =   765
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
         TabIndex        =   50
         Top             =   3375
         Width           =   2145
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
         TabIndex        =   49
         Top             =   3375
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmLogActaConformidadLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** Nombre : frmLogActaConformidadLibre
'** Descripción : Registro de Acta de Conformidad Libre creado segun ERS062-2013
'** Creación : EJVG, 20131009 09:00:00 AM
'*******************************************************************************
Option Explicit

Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double
Dim fnTpoDocOrigen As Integer
Dim fnMovNroOCS As Long
Dim fnDocTpo As Integer
Dim fsAreaAgeCod As String
Dim fnMoneda As Integer
Dim fRsAgencia As New ADODB.Recordset
Dim fRsServicio As New ADODB.Recordset
Dim fRsCompra As New ADODB.Recordset
Dim fsCtaContCodProv As String

Dim fnTpoCambio As Currency
Dim fbGraboActa As Boolean

Private Sub Form_Load()
    fsAreaAgeCod = gsCodArea & Right(gsCodAge, 2)
    fnFormTamanioIni = 4815
    fnFormTamanioActiva = 8475
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
        Else
            Set fRsAgencia = Nothing
            Set fRsCompra = Nothing
            Set fRsServicio = Nothing
        End If
    End If
End Sub
Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    gsOpeCod = psOpeCod
    fnMoneda = Mid(gsOpeCod, 3, 1)
    Caption = UCase(psOpeDesc)
    Show 1
End Sub
Private Sub CargaControles()
    Dim oLog As New DLogGeneral
    Dim oArea As New DActualizaDatosArea
    Dim rs As New ADODB.Recordset
    
    Set rs = oLog.ListaTpoDocOrigenActaConformidad("5,6")
    cboTpoDocOrigen.Clear
    CargaCombo rs, cboTpoDocOrigen, , 1, 0
    txtAreaAgeCod.rs = oArea.GetAgenciasAreas
    
    Set rs = Nothing
    Set oArea = Nothing
    Set oLog = Nothing
End Sub
Private Sub CargaVariables()
    Dim odoc As New DOperacion
    Dim oArea As New DActualizaDatosArea
    Dim oALmacen As New DLogAlmacen
    Dim rs As New ADODB.Recordset
    
    fbGraboActa = False
    If gbBitTCPonderado Then
        fnTpoCambio = gnTipCambioPonderado
    Else
        fnTpoCambio = gnTipCambioC
    End If

    Set rs = odoc.CargaOpeCta(gsOpeCod, "H")
    fsCtaContCodProv = rs!cCtaContCod
    
    Set fRsAgencia = oArea.GetAgencias(, , True)
    Set fRsCompra = oALmacen.GetBienesAlmacen(, "11','12','13")
    Set fRsServicio = OrdenServicio()

    Set rs = Nothing
    Set oArea = Nothing
    Set oALmacen = Nothing
    Set odoc = Nothing
End Sub
Private Sub cboTpoDocOrigen_Click()
    Dim lnTpoDoc As Integer
    cancela_busqueda_actual
    If Trim(Right(cboTpoDocOrigen.Text, 4)) <> "" Then
        lnTpoDoc = CInt(Trim(Right(cboTpoDocOrigen.Text, 4)))
    End If
    fnTpoDocOrigen = lnTpoDoc
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
        feOrden.lbUltimaInstancia = True
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
        feOrden.lbUltimaInstancia = False
    End If
End Sub
Private Sub cboTpoDocOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregar.SetFocus
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
Private Sub feOrden_RowColChange()
    If feOrden.col = 1 Then
        feOrden.rsTextBuscar = fRsAgencia
    ElseIf feOrden.col = 2 Then
        If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
            feOrden.rsTextBuscar = fRsCompra
        ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
            feOrden.rsTextBuscar = fRsServicio
        End If
    End If
End Sub
Private Sub feOrden_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    If psDataCod <> "" Then
        If pnCol = 2 Then
            If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
                AsignaObjetosSer psDataCod
            End If
        End If
        If pnCol = 1 Or pnCol = 2 Then
            '*** Si esta vacio el campo de la cuenta contable y si ya eligió agencia y objeto
            If Len(Trim(feOrden.TextMatrix(pnRow, 1))) <> 0 And Len(Trim(feOrden.TextMatrix(pnRow, 2))) <> 0 Then
                feOrden.TextMatrix(pnRow, 7) = DameCtaCont(feOrden.TextMatrix(pnRow, 2), 0, Trim(feOrden.TextMatrix(pnRow, 1)))
            End If
            '***
        End If
    End If
End Sub
Private Sub feOrden_OnCellChange(pnRow As Long, pnCol As Long)
    On Error GoTo ErrfeOrden_OnCellChange
    
    If feOrden.TextMatrix(1, 0) <> "" Then
        If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
            If pnCol = 4 Or pnCol = 5 Then
                feOrden.TextMatrix(pnRow, 6) = Format(Val(feOrden.TextMatrix(pnRow, 4)) * feOrden.TextMatrix(pnRow, 5), gsFormatoNumeroView)
            End If
        End If
    End If
    Exit Sub
ErrfeOrden_OnCellChange:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdDarConformidad_Click()
    Dim row As Long
    Dim bSelecciona As Boolean
    Dim oMov As DMov
    Dim lsCorrelativo As String
    
    On Error GoTo ErrcmdDarConformidad
    If Not validaBusquedaLibre Then Exit Sub
    If Not validaIngresoRegistros Then Exit Sub
    
    Screen.MousePointer = 11
    Height = fnFormTamanioActiva
    CentraForm Me
    TabBuscar.Enabled = False
    LimpiarDatosActaConformidad
    
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
        fnDocTpo = LogTipoActaConformidad.gActaRecepcionBienes
        TabActivaBien.TabCaption(0) = "Acta de Bien"
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
        fnDocTpo = LogTipoActaConformidad.gActaConformidadServicio
        TabActivaBien.TabCaption(0) = "Acta de Servicio"
    End If

    Set oMov = New DMov
    lsCorrelativo = oMov.GetCorrelativoActaConformidad(fnDocTpo, Right(gsCodAge, 2), CStr(Year(gdFecSis)))
    EstablecerDatosActaConformidad , , , , lsCorrelativo
    txtMoneda.Text = IIf(fnMoneda = 1, "SOLES", "DOLARES")
    
    Screen.MousePointer = 0
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
Private Sub cmdAgregar_Click()
    If Not validaBusquedaLibre Then Exit Sub
    If feOrden.TextMatrix(1, 0) <> "" Then
        If Not validaIngresoRegistros Then Exit Sub
    End If
    feOrden.AdicionaFila
    
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
        feOrden.ColumnasAEditar = "X-1-2-X-4-5-X-X"
        feOrden.TextMatrix(feOrden.row, 4) = "0"
        feOrden.TextMatrix(feOrden.row, 5) = "0.00"
        feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
        feOrden.ColumnasAEditar = "X-1-2-X-X-X-6-X"
    End If
    feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    feOrden.col = 2
    feOrden.SetFocus
    feOrden_RowColChange
End Sub
Private Sub cmdQuitar_Click()
    feOrden.EliminaFila feOrden.row
End Sub
Private Sub txtAreaAgeCod_EmiteDatos()
    txtAreaAgeNombre.Text = ""
    If txtAreaAgeCod.Text <> "" Then
        txtAreaAgeNombre.Text = txtAreaAgeCod.psDescripcion
    End If
End Sub

Private Sub txtDocReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtProveedorCod.SetFocus
    End If
End Sub

Private Sub txtProveedorCod_EmiteDatos()
    Dim oLog As New DLogGeneral
    Dim oProv As New DLogProveedor
    Dim rs As New ADODB.Recordset
    Dim bExite As Boolean
    
    txtProveedorNombre.Text = ""
    txtProveedorDocTpo.Text = ""
    txtProveedorDocNro.Text = ""
    txtProveedorCtaNro.Text = ""
    txtProveedorCtaMoneda.Text = ""
    txtProveedorCtaInstitucionCod.Text = ""
    txtProveedorCtaInstitucionNombre.Text = ""
    
    bExite = oProv.IsExisProveedor(txtProveedorCod.Text)
    If Not bExite Then
        MsgBox "El proveedor seleccionado no se encuentra registrado en la base de proveedores" & Chr(10) & "del Departamento de Logística, esto sera necesario para el Pago", vbInformation, "Aviso"
    End If
    txtProveedorNombre.Text = txtProveedorCod.psDescripcion
    Set rs = oLog.GetProveedorxActaConformidadLibre(txtProveedorCod.Text, fnMoneda)
    
    If Not RSVacio(rs) Then
        EstablecerDatosActaConformidad , , , , , , , rs!cDocTpo, rs!cDocNro, rs!cIFiCtaCod, IIf(rs!cIFiCtaCod = "", "", rs!cMoneda), rs!cIFiCod, rs!cIFiNombre
    End If
    
    Set rs = Nothing
    Set oProv = Nothing
    Set oLog = Nothing
End Sub
Private Sub txtSubAreaDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCompraDescripcion.Visible And txtCompraDescripcion.Enabled Then
            txtDocReferencia.SetFocus
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
    Dim Index As Integer, indexObj As Integer
    Dim lsMovNro As String
    Dim lsSubCta As String
    
    On Error GoTo ErrCmdConforme
    
    If Len(Trim(txtAreaAgeCod.Text)) = 0 Then
        MsgBox "La presente conformidad no cuenta con Área Agencia", vbInformation, "Aviso"
        txtAreaAgeCod.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtDocReferencia.Text)) = 0 Then
        MsgBox "Ud. debe de ingresar un número de documento de referencia", vbInformation, "Aviso"
        txtDocReferencia.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtActaNro.Text)) = 0 Then
        MsgBox "No se ha conseguido el correlativo del Nro de Acta", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtProveedorCod.Text)) = 0 Then
        MsgBox "Ud. debe de seleccionar el Proveedor para la presente Acta de Conformidad", vbInformation, "Aviso"
        txtProveedorCod.SetFocus
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
    ReDim DatosOrden(Index)

    For Index = 1 To feOrden.Rows - 1
        ReDim Preserve DatosOrden(Index)
        If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
            DatosOrden(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
        ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
            lsSubCta = ""
            For indexObj = 1 To feObj.Rows - 1
                If feObj.TextMatrix(indexObj, 1) = feOrden.TextMatrix(Index, 0) Then
                    lsSubCta = lsSubCta & feObj.TextMatrix(indexObj, 5)
                End If
            Next
            DatosOrden(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 2)))) & lsSubCta
        End If
        DatosOrden(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
        DatosOrden(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
        DatosOrden(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
        DatosOrden(Index).nTotal = feOrden.TextMatrix(Index, 6)
    Next
   
    If UBound(DatosOrden) = 0 Then
        MsgBox "No existen Items a dar conformidad", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
        fsCtaContCodProv = "25" & Mid(gsOpeCod, 3, 1) & "601"
    ElseIf fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.ServicioLibre Then
        fsCtaContCodProv = "25" & Mid(gsOpeCod, 3, 1) & "60202"
    End If
    If fsCtaContCodProv = "" Then
        MsgBox "No se ha definido cuenta contable de Proveedor, consulte al Dpto. de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de guardar el Acta de Conformidad Digital Libre?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oLog = New NLogGeneral
    
    Screen.MousePointer = 11
    lnMovNro = oLog.GrabarActaConformidad_Orden(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsOpeCod, fnTpoDocOrigen, fnDocTpo, Trim(txtDocReferencia.Text), txtAreaAgeCod.Text, Trim(txtSubAreaDescripcion.Text), fnMoneda, lsNroActaConformidad, txtProveedorCod.Text, Trim(Me.txtCompraDescripcion.Text), Trim(txtCompraObservacion.Text), DatosOrden, fsCtaContCodProv, fnTpoCambio, lsMovNro)
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
Private Sub cancela_busqueda_actual()
    Height = fnFormTamanioIni
    Call FormateaFlex(feOrden)
    Call FormateaFlex(feObj)
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
Private Function validaBusquedaLibre() As Boolean
    validaBusquedaLibre = True
    If cboTpoDocOrigen.ListIndex = -1 Then
        MsgBox "Ud. primero debe de seleccionar el Tipo de Documento Origen", vbInformation, "Aviso"
        validaBusquedaLibre = False
        cboTpoDocOrigen.SetFocus
        Exit Function
    End If
End Function
Private Function validaIngresoRegistros() As Boolean
    Dim i As Long, J As Long
    Dim col As Integer
    Dim Columnas() As String
    Dim lsColumnas As String
    
    lsColumnas = "1,2,6"
    Columnas = Split(lsColumnas, ",")
        
    validaIngresoRegistros = True
    If feOrden.TextMatrix(1, 0) <> "" Then
        For i = 1 To feOrden.Rows - 1
            For J = 1 To feOrden.Cols - 1
                For col = 0 To UBound(Columnas)
                    If J = Columnas(col) Then
                        If Len(Trim(feOrden.TextMatrix(i, J))) = 0 And feOrden.ColWidth(J) <> 0 Then
                            MsgBox "Ud. debe especificar el campo " & feOrden.TextMatrix(0, J), vbInformation, "Aviso"
                            validaIngresoRegistros = False
                            feOrden.TopRow = i
                            feOrden.row = i
                            feOrden.col = J
                            feOrden_RowColChange
                            Exit Function
                        End If
                    End If
                Next
            Next
            If IsNumeric(feOrden.TextMatrix(i, 6)) Then
                If CCur(feOrden.TextMatrix(i, 6)) <= 0 Then
                    MsgBox "El Importe Total debe ser mayor a cero", vbInformation, "Aviso"
                    validaIngresoRegistros = False
                    feOrden.TopRow = i
                    feOrden.row = i
                    feOrden.col = 6
                    Exit Function
                End If
            Else
                MsgBox "El Importe Total debe ser númerico", vbInformation, "Aviso"
                validaIngresoRegistros = False
                feOrden.TopRow = i
                feOrden.row = i
                feOrden.col = 6
                Exit Function
            End If
            If fnTpoDocOrigen = LogTipoDocOrigenActaConformidad.CompraLibre Then
                If Len(Trim(feOrden.TextMatrix(i, 7))) = 0 Then
                    MsgBox "El Objeto " & feOrden.TextMatrix(i, 3) & Chr(10) & "no tiene configurado Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                    feOrden.TopRow = i
                    feOrden.row = i
                    feOrden.col = 2
                    validaIngresoRegistros = False
                    Exit Function
                End If
            End If
        Next
    Else
        MsgBox "Ud. debe agregar los Bienes/Servicios a dar Conformidad", vbInformation, "Aviso"
        validaIngresoRegistros = False
    End If
End Function
Private Function OrdenServicio() As ADODB.Recordset
    Dim oCon As New DConecta
    Dim sSqlO As String
    Dim lnMoneda As Integer
    If fnMoneda <> 0 Then
        oCon.AbreConexion
        sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
              & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
              & "WHERE b.cCtaContCod = a.cCtaContCod AND (a.cOpeCod='" & IIf(fnMoneda = 1, "501207", "502207") & "' AND (a.cOpeCtaDH='D'))"
        Set OrdenServicio = oCon.CargaRecordSet(sSqlO)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function
Private Sub AsignaObjetosSer(ByVal sCtaCod As String)
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As New DConecta
    Dim oCtaCont As New DCtaCont
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim oRHAreas As New DActualizaDatosArea
    Dim oCtaIf As New NCajaCtaIF
    Dim oEfect As New Defectivo
    Dim oDescObj As New ClassDescObjeto
    Dim oContFunct As New NContFunciones
    Dim lsRaiz As String, lsFiltro As String, sSql As String
        
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
    EliminaObjeto feOrden.row

    sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSql)
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOk Then
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                            Else
                                EliminaObjeto feOrden.row
                                Exit Do
                            End If
                        Else
                            AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If

    Set rs = Nothing
    Set rs1 = Nothing
    Set oDescObj = Nothing
    Set oCon = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
    Set oContFunct = Nothing
    Set oContFunct = Nothing
    Exit Sub
End Sub
Private Sub AdicionaObjeto(ByVal pnItem As Integer, ByVal psCtaObjOrden As String, ByVal psCodigo As String, ByVal psDesc As String, ByVal psFiltro As String, ByVal psObjetoCod As String)
    feObj.AdicionaFila
    feObj.TextMatrix(feObj.row, 1) = pnItem
    feObj.TextMatrix(feObj.row, 2) = psCtaObjOrden
    feObj.TextMatrix(feObj.row, 3) = psCodigo
    feObj.TextMatrix(feObj.row, 4) = psDesc
    feObj.TextMatrix(feObj.row, 5) = psFiltro
    feObj.TextMatrix(feObj.row, 6) = psObjetoCod
End Sub
Private Sub EliminaObjeto(ByVal pnItem As Integer)
    Dim i As Long
    Dim bEncuentra As Boolean
    If feObj.TextMatrix(1, 0) <> "" Then
        For i = 1 To feObj.Rows - 1
            If Val(feObj.TextMatrix(i, 1)) = pnItem Then
                bEncuentra = True
                Exit For
            End If
        Next
    End If
    If bEncuentra Then
        feObj.EliminaFila i
        EliminaObjeto pnItem
    End If
End Sub
Private Function DameCtaCont(ByVal psObjeto As String, nNiv As Integer, psAgeCod As String) As String
    Dim oCon As New DConecta
    Dim oForm As New frmLogOCompra
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    
    sSql = oForm.FormaSelect(gsOpeCod, psObjeto, 0, psAgeCod)
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    If Not rs.EOF Then
        DameCtaCont = rs!cObjetoCod
    End If
    Set rs = Nothing
    Set oForm = Nothing
    Set oCon = Nothing
End Function
