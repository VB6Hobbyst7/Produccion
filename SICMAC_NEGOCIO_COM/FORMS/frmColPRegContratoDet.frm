VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPRegContratoDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Registrar Contrato"
   ClientHeight    =   9768
   ClientLeft      =   996
   ClientTop       =   1260
   ClientWidth     =   9540
   Icon            =   "frmColPRegContratoDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9768
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fr_CampPrendario 
      Height          =   855
      Left            =   120
      TabIndex        =   90
      Top             =   8280
      Width           =   9375
      Begin VB.ComboBox cboCampPrendario 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtCampPrendario 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Label5 
         Caption         =   "Campaña: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   550
      Left            =   240
      TabIndex        =   81
      Top             =   6280
      Width           =   9075
      Begin VB.Label lblSaldCapA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7080
         TabIndex        =   88
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Sald. Cap Acumulado:"
         Height          =   255
         Left            =   5400
         TabIndex        =   87
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label lblMontoMax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4080
         TabIndex        =   86
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lblDeudaSBS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1200
         TabIndex        =   84
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Deuda SBS:"
         Height          =   255
         Left            =   180
         TabIndex        =   83
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Monto Max. Otorgado:"
         Height          =   255
         Left            =   2400
         TabIndex        =   82
         Top             =   200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   69
      Top             =   7560
      Width           =   9375
      Begin SICMACT.TxtBuscar txtBuscarLinea 
         Height          =   345
         Left            =   1200
         TabIndex        =   70
         Top             =   240
         Width           =   1545
         _ExtentX        =   2731
         _ExtentY        =   614
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label18 
         Caption         =   "Linea Cred."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLineaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   3000
         TabIndex        =   71
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdImpVolTas 
      Caption         =   "&Volante de Tas."
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   57
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   52
      Top             =   6845
      Width           =   9375
      Begin VB.CommandButton cmdVerRetasacion 
         Caption         =   "Ver"
         Height          =   350
         Left            =   8640
         TabIndex        =   68
         Top             =   180
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox ChkAnterior 
         Caption         =   "Contrato Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   6000
         Picture         =   "frmColPRegContratoDet.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Buscar ..."
         Top             =   180
         Visible         =   0   'False
         Width           =   420
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   2160
         TabIndex        =   53
         Top             =   180
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6583
         _ExtentY        =   656
         Texto           =   "Credito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblCredRetasado 
         Caption         =   "CRÉDITO RETASADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6600
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.Frame fraContenedor 
      Height          =   6880
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   0
      Width           =   9375
      Begin VB.Frame fr_TasaEspecial 
         Height          =   1035
         Left            =   5160
         TabIndex        =   94
         Top             =   5260
         Width           =   2055
         Begin VB.CheckBox chkTasaEspeci 
            Caption         =   "Tasa Especial (TEM)"
            Height          =   375
            Left            =   120
            TabIndex        =   97
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtTasaEspeci 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   96
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdSolTasa 
            Caption         =   "..."
            Height          =   255
            Left            =   1440
            TabIndex        =   95
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "%"
            Height          =   255
            Left            =   1200
            TabIndex        =   98
            Top             =   480
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1035
         Left            =   120
         TabIndex        =   75
         Top             =   5260
         Width           =   4995
         Begin VB.ComboBox cboTasador 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtHolograma 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   78
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblTasador 
            Caption         =   "Tasador:"
            Height          =   255
            Left            =   2520
            TabIndex        =   77
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblHolograma 
            Caption         =   "N° Holograma:"
            Height          =   255
            Left            =   180
            TabIndex        =   76
            Top             =   200
            Width           =   1095
         End
      End
      Begin VB.Frame fraPiezasDet 
         Caption         =   "Detalle de Piezas"
         Height          =   2055
         Left            =   120
         TabIndex        =   49
         Top             =   1580
         Width           =   7095
         Begin VB.CommandButton cmdPiezaEliminar 
            Caption         =   "-"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6720
            TabIndex        =   5
            Top             =   1080
            Width           =   345
         End
         Begin VB.CommandButton CmdPiezaAgregar 
            Caption         =   "+"
            Height          =   495
            Left            =   6720
            TabIndex        =   3
            Top             =   480
            Width           =   345
         End
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   1695
            Left            =   120
            TabIndex        =   4
            Top             =   225
            Width           =   6555
            _ExtentX        =   11557
            _ExtentY        =   2985
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   2
            EncabezadosNombres=   "Num-Pzas-Material-PBruto-PNeto-Tasac-Descripcion-Item"
            EncabezadosAnchos=   "400-450-1030-650-650-700-2500-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-X-6-X"
            ListaControles  =   "0-0-3-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-R-R-L-C"
            FormatosEdit    =   "0-3-1-2-2-2-0-3"
            TextArray0      =   "Num"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   408
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   975
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   3600
         Width           =   7035
         Begin VB.ComboBox cboPlazo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPRegContratoDet.frx":040C
            Left            =   5580
            List            =   "frmColPRegContratoDet.frx":040E
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   1125
         End
         Begin VB.TextBox txtPiezas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   16
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtMontoPrestamo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5580
            MaxLength       =   11
            TabIndex        =   7
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label lblOroBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   50
            Top             =   240
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
            TabIndex        =   43
            Top             =   615
            Width           =   1095
         End
         Begin VB.Label lblOroNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1200
            TabIndex        =   42
            Top             =   615
            Width           =   1035
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Oro Bruto  (gr)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Oro Neto  (gr)"
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Piezas"
            Height          =   210
            Index           =   2
            Left            =   2520
            TabIndex        =   26
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo  (dias)"
            Height          =   255
            Index           =   8
            Left            =   4560
            TabIndex        =   25
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tasación "
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   24
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Prestamo"
            Height          =   255
            Index           =   9
            Left            =   4530
            TabIndex        =   23
            Top             =   615
            Width           =   1335
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   570
         Index           =   7
         Left            =   420
         TabIndex        =   44
         Top             =   1980
         Width           =   5415
         Begin VB.Label lblEtiqueta 
            Caption         =   "Kilataje (gr)"
            Height          =   195
            Index           =   22
            Left            =   3120
            TabIndex        =   48
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   23
            Left            =   300
            TabIndex        =   47
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblOroPrestamo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   4080
            TabIndex        =   46
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblOroPrestamoPorcen 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   45
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Linea de Credito"
         ForeColor       =   &H8000000D&
         Height          =   540
         Index           =   3
         Left            =   165
         TabIndex        =   35
         Top             =   2640
         Width           =   5370
         Begin VB.CommandButton cmdLineaCredito 
            Caption         =   "..."
            Height          =   285
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   180
            Width           =   375
         End
         Begin VB.Label lblLineaCredito 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   180
            Width           =   4695
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Cliente(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1425
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   7050
         Begin VB.ComboBox cboTipcta 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPRegContratoDet.frx":0410
            Left            =   5040
            List            =   "frmColPRegContratoDet.frx":041D
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   1830
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   180
            TabIndex        =   0
            Top             =   990
            Width           =   825
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   345
            Left            =   1080
            TabIndex        =   1
            Top             =   990
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComctlLib.ListView lstCliente 
            Height          =   795
            Left            =   90
            TabIndex        =   36
            Top             =   180
            Width           =   6795
            _ExtentX        =   11980
            _ExtentY        =   1397
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo del Cliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre / Razón Social del Cliente"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dirección"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Teléfono"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Ciudad"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Zona"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Doc.Civil"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Nro.Doc.Civil"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Doc.Tributario"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Nro.Doc.Tributario"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tipo de contrato"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   34
            Top             =   1080
            Width           =   1245
         End
      End
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   735
         Index           =   2
         Left            =   180
         TabIndex        =   29
         Top             =   4540
         Width           =   7035
         Begin VB.Label lblNetoRecibir 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5640
            TabIndex        =   41
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblInteres 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1200
            TabIndex        =   40
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblFechaVencimiento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3360
            TabIndex        =   39
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Neto Recibir"
            Height          =   255
            Index           =   20
            Left            =   4680
            TabIndex        =   32
            Top             =   250
            Width           =   945
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Interes"
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   31
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fec.Vencim."
            Height          =   255
            Index           =   15
            Left            =   2400
            TabIndex        =   30
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Kilataje"
         Height          =   1500
         Index           =   5
         Left            =   5760
         TabIndex        =   17
         Top             =   3840
         Visible         =   0   'False
         Width           =   1350
         Begin VB.TextBox txt21k 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   525
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1140
            Width           =   720
         End
         Begin VB.TextBox txt18k 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   525
            MaxLength       =   10
            TabIndex        =   14
            Top             =   840
            Width           =   720
         End
         Begin VB.TextBox txt16k 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   540
            MaxLength       =   10
            TabIndex        =   13
            Top             =   540
            Width           =   720
         End
         Begin VB.TextBox txt14k 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   540
            MaxLength       =   10
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "21 K"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   21
            Top             =   1140
            Width           =   465
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "18 K"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "16 K"
            Height          =   255
            Index           =   12
            Left            =   135
            TabIndex        =   19
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "14 K"
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   18
            Top             =   255
            Width           =   495
         End
      End
      Begin VB.Label lblSegPrenExter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   7320
         TabIndex        =   89
         Top             =   5800
         Width           =   1875
      End
      Begin VB.Label lblTipoForGarPigno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   7320
         TabIndex        =   80
         Top             =   4400
         Width           =   1815
      End
      Begin VB.Label lblClienteTpo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   7320
         TabIndex        =   74
         Top             =   5200
         Width           =   1875
      End
      Begin VB.Label lblCalificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   7320
         TabIndex        =   73
         Top             =   5500
         Width           =   1875
      End
      Begin VB.Label lblNotaPigAdjCli 
         Caption         =   "Nota: El oro neto ha sido castigado un 15% por tener más de 3 créditos prendarios adjudicados en los últimos 24 meses."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7320
         TabIndex        =   66
         Top             =   3000
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblCalificacionPerdida 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7320
         TabIndex        =   65
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label lblCalificacionDudoso 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7320
         TabIndex        =   64
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Label lblCalificacionDeficiente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7320
         TabIndex        =   63
         Top             =   1320
         Width           =   1875
      End
      Begin VB.Label lblCalificacionPotencial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7320
         TabIndex        =   62
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label lblPorcentajeTasa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7320
         TabIndex        =   61
         Top             =   2640
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "T.E.M."
         Height          =   255
         Left            =   7320
         TabIndex        =   60
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblCalificacionNormal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7320
         TabIndex        =   59
         Top             =   620
         Width           =   1875
      End
      Begin VB.Label lblTituloCalificacion 
         Caption         =   "Última Calificación Según SBS - RCC "
         Height          =   375
         Left            =   7320
         TabIndex        =   58
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdPrevio 
      Caption         =   "Hoja Informativa"
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   9240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   9240
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   360
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   402
      _ExtentY        =   635
      _Version        =   393217
      TextRTF         =   $"frmColPRegContratoDet.frx":0458
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "frmColPRegContratoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE CONTRATO.
'Archivo:  frmColPRegContrato.frm
'LAYG   :  01/05/2001.
'Resumen:  Nos permite registrar un contrato pignoraticio

Option Explicit
'******** Parametros de Colocaciones
Dim fnPorcentajePrestamo As Double
Dim fnImpresionesContrato As Double
Dim fnMinPesoOro As Double
Dim fnMaxMontoPrestamo1 As Double
Dim fnRangoPreferencial As Double

Dim fnTasaInteresAdelantado As Double ' Cambia si es tasa preferencial
'Dim fnTasaInteresVencido As Double

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
'Dim lista As ListItem
Dim lstTmpCliente As MSComctlLib.ListItem
Dim lsIteSel As String
Dim vContAnte As Boolean
Dim gColocLineaCredPig As String
Dim lsSQL As String

Dim vSolTasa As Boolean 'MACM 23032021
Dim vTasaForm As Boolean 'MACM 26032021
Dim vNuevaSol As Boolean 'MACM 26032021
Dim cc As String 'MACM 20210323
Dim vEstadoSolTasa As Integer 'MACM 24032021
Dim lnTasaSol As Currency 'MACM 25032021
Dim pcCtaCod As String 'MACM 24032021

Dim objPista As COMManejador.Pista

Dim fsColocLineaCredPig  As String

'***Agregado por ELRO el 20120103, según Acta N° 002-2012/TI-D
Dim fbMalCalificacion As Boolean
'***Fin Agregado por ELRO*************************************
'************RECO20131121 ERS158******************************
Dim nTpoCliente As Integer
Dim lnPesoNetoDesc As Double
'*****************END RECO************************************
'******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
Dim fbClienteCPP As Boolean
'************END RECO*****************************************
Dim fsPlazoSist As String 'RECO20150421
Dim sLineaTmp As String 'ALPA20150617**
Dim RLinea As ADODB.Recordset 'ALPA20150617**
Dim lnTasaInicial As Currency 'ALPA20150617**
Dim lnTasaFinal As Currency 'ALPA20150617**
Dim lnTasaGracia As Currency
Dim lnTasaCompes As Currency
Dim lnTasaMorato As Currency
Dim lsTextoCalif As String 'PEAC20170727
Dim fscPerCod As String 'ARLO20171205
Dim fnTpoClinte As Integer 'ARLO20171205
Dim fnPorcentaje As Double 'ARLO20171205

Dim nDiasSegmento As Integer 'JOEP20180220
Dim nCantAdjuSegmento As Integer 'JOEP20180220
Dim fcSegmento As String 'JOEP20180220
Dim nVerificaVencidos As Integer 'JOEP pig pase
Dim nCalificacionPotencialCPP As Double 'JOEP pig pase
Dim nDiasSinAdj As Integer 'JOEP pig pase
Dim nSegmentoAnt As String 'JOEP pig pase
Dim bErrHolog As Boolean 'JUCS TI ERS063-2017
Dim nValorUIT As Currency 'CROB20180611 ERS076-2017
Dim rsRecuperaConstantes As ADODB.Recordset 'ADD PTI1 06-03-2019
Dim rsRecuperaSegDR As ADODB.Recordset 'add pti1 06-03-2019
Dim rsRecuperapsConfDR As ADODB.Recordset 'add pti1 06-03-2019
Dim sCpersTem As String 'add pti1 07-03-2019
Dim lsFecVenc As String ' PEAC 20190515
Dim ldVarFecVencimiento As Date ' PEAC 20190515
Dim lnDiasDomFer As Integer ' PEAC 20190520
Dim gnDeudaSBS, gnMontoMax, gnSaldAcum As Double 'APRI20190515 SATI

Dim vArrayDatosSegPred As Variant 'JOEP20210513 Segmentacion Predario externo
Dim lnPlazonew As Integer 'PEAC 20210914
Dim nCampana As Integer 'RIRO 20210922 campana prendario
Dim loColPFunc As COMDColocPig.DCOMColPFunciones 'PEAC 20211021

'Procedimiento para cargar los valores a los campos txt
Private Sub CalculaCostosAsociados()
Dim loCostos As COMNColoCPig.NCOMColPCalculos
Dim loTasaInt As COMDColocPig.DCOMColPCalculos

Set loCostos = New COMNColoCPig.NCOMColPCalculos
    
    '*** PEAC 20170123 - VALIDA EL CORRECTO INGRESO DEL MONTO
    If Not IsNumeric(txtMontoPrestamo.Text) Then
        txtMontoPrestamo.Text = "0.00"
'        Exit Sub
    End If
    vPrestamo = CDbl(txtMontoPrestamo.Text)
    Set loTasaInt = New COMDColocPig.DCOMColPCalculos
'    If vPrestamo >= fnRangoPreferencial Then
'        fnTasaInteresAdelantado = loTasaInt.dObtieneTasaInteres("01011130502", "01")
'    Else
        'fnTasaInteresAdelantado = loTasaInt.dObtieneTasaInteres(fsColocLineaCredPig, "01")
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
            '***Modificado por ELRO el 20120104, según Acta N° 002-2012/TI-D
            'If fbMalCalificacion = False Then
            '    fnTasaInteresAdelantado = loTasaInt.dObtieneTasaInteres(fsColocLineaCredPig, "1")
            '    lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
            'Else
            '    fnTasaInteresAdelantado = gnTasaIntCrePigCliMalCal
            '    lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
            'End If
            '***Fin Modificado por ELRO*************************************
        'MACM 20210323
        If vSolTasa And vEstadoSolTasa = 2 Then
            lblPorcentajeTasa = lnTasaSol & " %"
            fnTasaInteresAdelantado = lnTasaSol
        Else
            If fbClienteCPP = False Then
                'fnTasaInteresAdelantado = loTasaInt.dObtieneTasaInteres(fsColocLineaCredPig, "1")
                'ALPA 20150620******************************************
                fnTasaInteresAdelantado = lnTasaInicial
                lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
            Else
                'ALPA 20150620******************************************
                'Dim oClases As New clases.NConstSistemas
                'fnTasaInteresAdelantado = gnTasaIntCrePigCliMalCal
                'fnTasaInteresAdelantado = oClases.LeeConstSistema(453)
                fnTasaInteresAdelantado = lnTasaFinal
                lblPorcentajeTasa = CStr(fnTasaInteresAdelantado) & " %"
            End If
            '********END RECO**********************************************
        End If
'    End If
    Set loTasaInt = Nothing
   
'*** PEAC 20190515
    lnDiasDomFer = 0
    Dim loPigFunc As COMDColocPig.DCOMColPFunciones
    Set loPigFunc = New COMDColocPig.DCOMColPFunciones
    
    Dim rsFecVenvFeri As ADODB.Recordset
    Set rsFecVenvFeri = New ADODB.Recordset
    Set rsFecVenvFeri = loPigFunc.dObtieneFechaVencimientoFeriado(Format(DateAdd("d", val(cboPlazo.Text), gdFecSis), "yyyyMMdd"), gsCodAge, Format(gdFecSis, "yyyyMMdd"))
    ldVarFecVencimiento = rsFecVenvFeri!dNuevaFecVenc
    lnDiasDomFer = rsFecVenvFeri!nCuentaDomFer
    lnPlazonew = DateDiff("d", gdFecSis, ldVarFecVencimiento) 'PEAC 20210914

'*** FIN PEAC
   
    'vPlazo = Val(cboPlazo.Text)
    'Cálculo valores
    'vCostoTasacion = Val(lblValorTasacion.Caption) * fnTasaTasacion
    vCostoTasacion = loCostos.nCalculaCostoTasacion(val(lblValorTasacion.Caption), fnTasaTasacion)
    
    'vCostoCustodia = loCostos.nCalculaCostoCustodia(val(lblValorTasacion.Caption), fnTasaCustodia, val(cboPlazo.Text))
    vCostoCustodia = loCostos.nCalculaCostoCustodia(val(lblValorTasacion.Caption), fnTasaCustodia, lnPlazonew) 'PEAC 20210914
    
        
    '*** PEAC 20080806 *************************************
    'vInteres = loCostos.nCalculaInteresAdelantado(Val(txtMontoPrestamo.Text), fnTasaInteresAdelantado, Val(cboPlazo.Text))
    'vInteres = loCostos.nCalculaInteresAlVencimiento(val(txtMontoPrestamo.Text), fnTasaInteresAdelantado, val(cboPlazo.Text))
    
    'vInteres = loCostos.nCalculaInteresAlVencimiento(val(txtMontoPrestamo.Text), fnTasaInteresAdelantado, val(cboPlazo.Text) + lnDiasDomFer) '*** PEAC 20190708 se agrego "lnDiasDomFer"
    vInteres = loCostos.nCalculaInteresAlVencimiento(val(txtMontoPrestamo.Text), fnTasaInteresAdelantado, lnPlazonew + lnDiasDomFer) '*** PEAC 20210914 se agrego "lnPlazonew"
    
    '*** FIN PEAC ******************************************
    
    'COMENTADO POR PEAC 20070813
    'vInteres = 0
    
    'vImpuesto = (vCostoTasacion + vInteres + vCostoCustodia) * pTasaImpuesto
    
    '*** PEAC 20071207 - Solo para mostrar en el contrato el interes que pagará *************
    'vImpuesto = loCostos.nCalculaImpuestoDesembolso(vCostoTasacion, vInteres, vCostoCustodia, fnTasaImpuesto)
    
    vImpuesto = loCostos.nCalculaImpuestoDesembolso(vCostoTasacion, 0, vCostoCustodia, fnTasaImpuesto)
    '****************************************************************************************
    
    '*** PEAC 20071207 - NO CONSIDERAR INTERES PARA CALCULO SOLO PARA MOSTRAR EN EL CONTATO
    'vNetoPagar = Val(txtMontoPrestamo.Text) - vCostoTasacion - vCostoCustodia - vInteres - vImpuesto
    vNetoPagar = val(txtMontoPrestamo.Text) - vCostoTasacion - vCostoCustodia - vImpuesto
    '****************************************************************************************

Set loCostos = Nothing

'Muestra los Resultados
'Me.lblFechaVencimiento = Format(DateAdd("d", val(cboPlazo.Text), gdFecSis), "dd/mm/yyyy") '*** PEAC 20190515
Me.lblFechaVencimiento = Format(ldVarFecVencimiento, "dd/mm/yyyy") '*** PEAC 20190515

'COMENTÓ APRI20190515 SATI
'Me.lblCostoTasacion = Format(vCostoTasacion, "#0.00")
'Me.lblCostoCustodia = Format(vCostoCustodia, "#0.00")
Me.lblInteres = Format(vInteres, "#0.00")
'Me.lblImpuesto = Format(vImpuesto, "#0.00")
Me.lblNetoRecibir = Format(vNetoPagar, "#0.00")

End Sub

'Función para calcular el valor de tasación
' Calcula en base al precio del oro en el mercado

Private Function ValorTasacion() As Double

If val(txt14k.Text) >= 0 And val(txt16k.Text) >= 0 And val(txt18k.Text) >= 0 And val(txt21k.Text) >= 0 Then
   ValorTasacion = (val(txt14k.Text) * fnPrecioOro14) + (val(txt16k.Text) * fnPrecioOro16) + (val(txt18k.Text) * fnPrecioOro18) + (val(txt21k.Text) * fnPrecioOro21)
Else
   MsgBox " No se ha ingresado correctamente el Kilataje ", vbInformation, " Aviso "
End If

End Function

'Inicializa las variables del formulario
Private Sub Limpiar()
' Asignar  Variables Globales al Nro de Contrato
    'AXCodCta.Visible = False
    vContAnte = False
    'AXCodCta.Text = Right(Trim(gsCodAge), 2) & gsConPignor & varMoneda & varNroCorrelativo & varDigitoChequeo
    lblOroBruto.Caption = Format(0, "#0.00")
    lblOroNeto.Caption = Format(0, "#0.00")
    txtPiezas.Text = Format(0, "#0")
    'cboPlazo.ListIndex = 0 'RECO20150421
    lblValorTasacion.Caption = Format(0, "#0.00")
    txtMontoPrestamo.Text = Format(0, "#0.00")
    'txtDescLote.Text = ""
    Me.lblFechaVencimiento = ""
'COMENTÓ APRI20190515 SATI
'    Me.lblCostoTasacion = Format(0, "#0.00")
'    Me.lblCostoCustodia = Format(0, "#0.00")
'    Me.lblImpuesto = Format(0, "#0.00")
'END APRI
    Me.lblInteres = Format(0, "#0.00")
    'vContrato = ""
    Me.lblNetoRecibir.Caption = Format(0, "#0.00")
    txt14k.Text = Format(0, "#0.00")
    txt16k.Text = Format(0, "#0.00")
    txt18k.Text = Format(0, "#0.00")
    txt21k.Text = Format(0, "#0.00")
    Me.lblOroPrestamo.Caption = ""
    Me.lblOroPrestamoPorcen.Caption = ""
    'txtNroContrato.Text = ""
    lstCliente.ListItems.Clear
    cboTipcta.ListIndex = 0
    FEJoyas.Clear
    FEJoyas.rows = 2
    FEJoyas.FormaCabecera
    lnJoyas = 0
    Me.AXCodCta.CMAC = ""
    Me.AXCodCta.Prod = ""
    Me.AXCodCta.Age = ""
    Me.AXCodCta.Cuenta = ""
    Me.AXCodCta.Enabled = False
    Me.cmdBuscar.Enabled = False
    lblCalificacion.Caption = "" 'ARLO ERS082-2017
    Me.lblClienteTpo.Caption = "" 'ARLO ERS082-2017 ---AGREGADO DESDE LA 60
    'Seg. Prendario Externo JOEP20210422
    lblSegPrenExter.Caption = ""
    Set vArrayDatosSegPred = Nothing
    'Seg. Prendario Externo JOEP20210422
    txtHolograma.Text = "" 'JUCS TI ERS 063-2017
    txtTasaEspeci.Text = "" 'MACM ERS005-2021
    chkTasaEspeci.value = 0 'MACM 20210323
    cboTasador.ListIndex = -1 ' JUCS TI ERS 063-2017
    txtHolograma.Enabled = False 'JUCS TI ERS 063-2017
    cboTasador.Enabled = False ' JUCS TI ERS 063-2017
    Me.lblTipoForGarPigno.Caption = "" 'CROB20180611 ERS076-2017
    'APRI20190515 SATI
    lblDeudaSBS.Caption = Format(0, "#0.00")
    lblMontoMax.Caption = Format(0, "#0.00")
    lblSaldCapA.Caption = Format(0, "#0.00")
    'END APRI
    
     'JOEP20210914 campana prendario
    txtCampPrendario.Text = ""
    cboCampPrendario.ListIndex = -1
    CmdPrevio.top = 8280
    cmdImpVolTas.top = 8280
    cmdGrabar.top = 8280
    cmdCancelar.top = 8280
    cmdSalir.top = 8280
    fr_CampPrendario.Visible = False
    frmColPRegContratoDet.Height = 9150
    'JOEP20210914 campana prendario
End Sub

'Función que calcula el total de kilatajes
Private Function SumaKilataje() As Double
If val(txt14k.Text) >= 0 And val(txt16k.Text) >= 0 And val(txt18k.Text) >= 0 And val(txt21k.Text) >= 0 Then
   SumaKilataje = val(txt14k.Text) + val(txt16k.Text) + val(txt18k.Text) + val(txt21k.Text)
Else
   MsgBox " No se ha ingresado correctamente el Kilataje ", vbInformation, " Aviso "
End If
End Function

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
    'Me.cmdGrabar.Enabled = True 'JUEZ 20130717
End Sub

Private Sub cboTipCta_Click()
If lstCliente.ListItems.count = 1 Then
    cboTipcta.ListIndex = 0
ElseIf lstCliente.ListItems.count >= 2 And cboTipcta.ListIndex = 0 Then
    cboTipcta.ListIndex = 1
End If
End Sub

Private Sub cboTipcta_KeyPress(KeyAscii As Integer)
CmdPiezaAgregar.SetFocus
End Sub


Private Sub ChkAnterior_Click()
   If ChkAnterior.value = 1 Then
        Me.AXCodCta.Visible = True
        Me.cmdBuscar.Visible = True
        Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
        Me.AXCodCta.Age = ""
        Me.AXCodCta.Cuenta = ""
        Me.AXCodCta.Enabled = True
        Me.cmdBuscar.Enabled = True
    Else
        Me.AXCodCta.Visible = False
        Me.cmdBuscar.Visible = False
        Me.AXCodCta.Age = ""
        Me.AXCodCta.Cuenta = ""
        lblCredRetasado.Visible = False 'RECO20120823 ERS074-2014
        cmdVerRetasacion.Visible = False 'RECO20120823 ERS074-2014
        'Limpiar
    End If
End Sub
'MACM 20210323
Private Sub ChkTasaEspeci_Click()

    If chkTasaEspeci.value = 1 Then
    'JOEP20210914 Campana Prendario
        fr_CampPrendario.Enabled = False
        cboCampPrendario.ListIndex = -1
        txtCampPrendario.Text = ""
    'JOEP20210914 Campana Prendario
        vSolTasa = True
        vTasaForm = True
        txtTasaEspeci.Enabled = True
        txtTasaEspeci = ""

        If vEstadoSolTasa = 0 Then
            vNuevaSol = True
        End If
    Else
    'JOEP20210914 Campana Prendario
        Call ActivaCampanaPrendario(fscPerCod)
    'JOEP20210914 Campana Prendario
        vSolTasa = False
        vTasaForm = False
        txtTasaEspeci.Enabled = False
        txtTasaEspeci = ""
        'vNuevaSol = False
    End If
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstCance

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        Me.AXCodCta.Enabled = True
        AXCodCta.SetFocusCuenta
    Else
        Me.AXCodCta.CMAC = ""
        Me.AXCodCta.Prod = ""
        Me.AXCodCta.Age = ""
        Me.AXCodCta.Cuenta = ""
        Me.AXCodCta.Enabled = False
    End If
Set loCuentas = Nothing

fscPerCod = lsPersCod '**ARLO20180131

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

'Permite cancelar el proceso actual
Private Sub cmdCancelar_Click()
    Limpiar
    chkTasaEspeci.Enabled = True 'MACM 20210323
    txt14k.Enabled = False
    txt16k.Enabled = False
    txt18k.Enabled = False
    txt21k.Enabled = False
    txtPiezas.Enabled = False
    'cboPlazo.Enabled = False RECO20140208 ERS002
    lblValorTasacion.Enabled = False
    txtMontoPrestamo.Enabled = False
   ' txtDescLote.Enabled = False
    cboPlazo.ListIndex = 0
    cboTipcta.Enabled = False
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = False
    Me.AXCodCta.Enabled = False
    Me.AXCodCta.CMAC = ""
    Me.AXCodCta.Prod = ""
    Me.AXCodCta.Age = ""
    Me.AXCodCta.Cuenta = ""
    Me.cmdBuscar.Enabled = False
    '***Agregado porel 20120720, según Acta N° 002-2012/TI-D
    fbMalCalificacion = False
    lblCalificacionNormal = ""
    lblCalificacionPotencial = ""
    lblCalificacionDeficiente = ""
    lblCalificacionDudoso = ""
    lblCalificacionPerdida = ""
    lblPorcentajeTasa = ""
    ''*****RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
    fbClienteCPP = False
    '********END RECO*******************************************
    '***Fin Agregado por ELRO el 20120720*******************
    '***Agregado por ELRO el 20120720, según OYP-RFC076-2012
    lblNotaPigAdjCli.Visible = False
    '***Fin Agregado por ELRO el 20120720*******************
    lblCredRetasado.Visible = False  'RECO20120823 ERS074-2014
    cmdVerRetasacion.Visible = False 'RECO20120823 ERS074-2014
    txtBuscarLinea.Text = ""
    lblLineaDesc.Caption = ""
    lblCalificacion.Caption = "" 'ARLO ERS082-2017
    Me.lblClienteTpo.Caption = "" 'ARLO ERS082-2017 ---AGREGADO DESDE LA 60
    Me.lblTipoForGarPigno.Caption = "" 'CROB20180611 ERS076-2017
    'Seg. Prendario Externo JOEP20210422
    lblSegPrenExter.Caption = ""
    Set vArrayDatosSegPred = Nothing
    'Seg. Prendario Externo JOEP20210422
    '**** add pti1 07-03-2019
    fscPerCod = ""
    sCpersTem = ""
    ChkAnterior.value = 0
    Me.AXCodCta.Visible = False
    Me.cmdBuscar.Visible = False
    Me.AXCodCta.Age = ""
    Me.AXCodCta.Cuenta = ""
    lblCredRetasado.Visible = False
    cmdVerRetasacion.Visible = False
    cmdEliminar.Visible = False
    '***** fin pti1

'JOEP20210913 Camapaña Prendario
    fr_CampPrendario.Enabled = False
    fr_TasaEspecial.Enabled = False
    'JOEP20210913 Camapaña Prendario
End Sub

'Permite actualizar los datos en la base de datos
Private Sub CmdGrabar_Click()

Dim pbTran As Boolean
Dim lsCtaReprestamo As String

Dim lrPersonas As New ADODB.Recordset
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
Dim lnTasaPreferencial As Currency 'MACM 20210323
Dim lsMovNr As String 'MACM 20210323
Dim lsTipoContrato As String
Dim lsLote As String
Dim ln14k As Double, ln16k As Double, ln18k As Double, ln21k As Double
Dim lnIntAdelantado As Currency, lnCostoTasac As Currency, lnCostoCustodia As Currency, lnImpuesto As Currency
Dim lrJoyas  As New ADODB.Recordset

Dim loRegPig As COMNColoCPig.NCOMColPContrato
Dim loRegImp As COMNColoCPig.NCOMColPImpre
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsContrato As String
Dim loPrevio As previo.clsprevio
Dim lnNumImp As Integer
Dim oPers  As COMDPersona.UCOMPersona 'WIOR 20140123
Dim pnITF As Double

Dim lsCadImprimir As String
Dim lsmensaje As String

Dim objPigLimite As COMDColocPig.DCOMColPContrato 'JOEP ERS047

'AMDO 20130906 TI-ERS112-2013
Dim HojaResumen1 As String
Dim HojaResumen2 As String
'END AMDO20130906

'WIOR 20140123 *********************
Dim bTrabajadoVinc As Boolean
bTrabajadoVinc = False
'WIOR FIN **************************

'*** PEAC 20080811
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim pnMovNro As Long
Dim loVistoElectronico As SICMACT.frmVistoElectronico
Set loVistoElectronico = New SICMACT.frmVistoElectronico
Dim lnPorcentajeCred As Double 'MACM 20210323
Dim lnTasaN As Double 'MACM 20210323
Dim rsDatosPrintPDF As Recordset '*** PEAC 20161128

'JOEP20180412 Pig Pase
Dim objPigSalCatSalCap As COMDColocPig.DCOMColPContrato
'JOEP20180412 Pig Pase

'On Error GoTo ControlError
pbTran = False
'MACM 20210323 INICIO
If vSolTasa And vTasaForm And vEstadoSolTasa = 2 Or vEstadoSolTasa = 0 Then
        If vEstadoSolTasa = 0 And Me.chkTasaEspeci.value = 1 Then
            lnPorcentajeCred = val(Replace(Me.lblPorcentajeTasa.Caption, "%", "")) 'ANGC20211119 'val(Left(Me.lblPorcentajeTasa.Caption, 6))
            lnTasaN = val(Me.txtTasaEspeci)
                
            If lnTasaN > lnPorcentajeCred Then
                MsgBox "La tasa solicitada no puede superar el " & lnPorcentajeCred & "%", vbInformation, "Aviso"
                Exit Sub
            ElseIf lnTasaN = lnPorcentajeCred Then
                MsgBox "La tasa solicitada no puede ser igual al " & lnPorcentajeCred & "%", vbInformation, "Aviso"
                Exit Sub
            ElseIf lnTasaN = 0 Then
                MsgBox "La tasa solicitada no puede ser cero", vbInformation, "Aviso"
                Exit Sub
            End If
                
            MsgBox "Se realizará una solicitud de tasa preferencial", vbInformation, "Aviso"
            'Genera Mov Nro
            Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNr = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set loContFunct = Nothing
                
            Dim objCargo As COMDCredito.DCOMNivelAprobacion
            Set objCargo = New COMDCredito.DCOMNivelAprobacion
                
            Call objCargo.GetActualizarColocacPermisoEstado(pcCtaCod, 4)
            Call objCargo.GetActualizarColocacPermisoAprobacion(pcCtaCod, lsMovNr, "", gsCodPersUser, "", "", lnTasaN, 0, 1, 1)
            Set objCargo = Nothing
            cmdCancelar_Click
            Exit Sub
        Else
            Set loRegImp = New COMNColoCPig.NCOMColPImpre
            Dim rsPig As ADODB.Recordset
            Dim rsPigJoyas As ADODB.Recordset
            Dim rsPigPers As ADODB.Recordset
            Dim rsPigCostos As ADODB.Recordset
            Dim rsPigDet As ADODB.Recordset
            Dim rsPigTasas As ADODB.Recordset
            Dim rsPigCosNot As ADODB.Recordset
            'stp_sel_ObtieneCosNotarialPigno rsPigCosNot
            Call loRegImp.RecuperaDatosHojaResumenPigno(pcCtaCod, rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
            
            Call CargaHojaResumenPignoPDF(rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
            'INICIO EAAS20190516 SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
            Pig_ContratosAutomaticos rsPigPers, rsPig, lsContrato
            
            cmdCancelar_Click
            Exit Sub
        End If
Else
    If ValidaDatosGrabar = False Then Exit Sub
    
    'Asigno los valores a los parametros
    'ALPA 20101005***
    If Me.lstCliente.ListItems.count > 1 Then
        MsgBox "Crédito tienen mas de un titular. Si desea continuar debe eliminar a uno de ellos"
        cmdEliminar.Visible = True 'add pti1 07-03-2019
        Exit Sub
    End If
    '*****************
    Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.lstCliente)
    'txtDescLote = " " & vbCr
    
    lnMontoPrestamo = CCur(txtMontoPrestamo.Text)
    lnNetoRecibir = CCur(lblNetoRecibir.Caption)
    lnPlazo = val(cboPlazo.Text)
    lsFechaVenc = Format$(Me.lblFechaVencimiento, "mm/dd/yyyy")
    lnOroBruto = val(lblOroBruto.Caption)
    lnOroNeto = val(lblOroNeto.Caption)
    lnPiezas = val(txtPiezas.Text)
    lnValTasacion = CCur(lblValorTasacion.Caption)
    lsTipoContrato = Switch(cboTipcta.ListIndex = 0, "0", cboTipcta.ListIndex = 1, "1", cboTipcta.ListIndex = 2, "2")
    lsLote = "" 'txtDescLote.Text
    ln14k = val(txt14k.Text)
    ln16k = val(txt16k.Text)
    ln18k = val(txt18k.Text)
    ln21k = val(txt21k.Text)
    'COMENTÓ APRI20190515 SATI
    lnIntAdelantado = CCur(Me.lblInteres.Caption) 'JOEP20190806 Descomentar
    'lnCostoTasac = CCur(Me.lblCostoTasacion.Caption)
    'lnCostoCustodia = CCur(Me.lblCostoCustodia.Caption)
    'lnImpuesto = CCur(Me.lblImpuesto.Caption)
    'END APRI
    lnTasaPreferencial = val(txtTasaEspeci.Text) 'MACM 20210323
	'lnPorcentajeCred = val(Left(Me.lblPorcentajeTasa.Caption, 6)) 'MACM 20210323
    lnPorcentajeCred = val(Replace(Me.lblPorcentajeTasa.Caption, "%", "")) 'ANGC20211119
    Set lrJoyas = FEJoyas.GetRsNew
    
    'MACM 20210323
    If vSolTasa = True Then
        If lnTasaPreferencial > lnPorcentajeCred Then
            MsgBox "La tasa solicitada no puede superar el " & lnPorcentajeCred & "%", vbInformation, "Aviso"
            Exit Sub
        ElseIf lnTasaPreferencial = lnPorcentajeCred Then
            MsgBox "La tasa solicitada no puede ser igual al " & lnPorcentajeCred & "%", vbInformation, "Aviso"
            Exit Sub
        ElseIf lnTasaPreferencial = 0 Then
            MsgBox "La tasa solicitada no puede ser cero", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    'JUCS TI ERS 063-2017
    If val(txtHolograma.Text) = 0 Then
        MsgBox "Debe registrar un numero de holograma para este contrato pignoraticio", vbInformation, "Aviso"
        txtHolograma.SetFocus
        Exit Sub
        ElseIf val(txtHolograma.Text) <> 0 Then
        Call VerificaHolograma
    End If
    'Si validación del holograma ha devueldo algun error entonces salir sino continuar con el registro
    If bErrHolog = True Then
    Exit Sub
    End If
    'FIN JUCS ERS 063-2017
    
    'Validar ingreso de Joyas
    
    If ValidarMsh = True Then Exit Sub
    
        '*** PEAC 20080811 ******************************************************
            '*** el codigo de operacio falta definir para reg de contrato pig por miestras se puso 150100
            lbResultadoVisto = loVistoElectronico.Inicio(1, gColRegistraContratoPig, lrPersonas(0))
            If Not lbResultadoVisto Then
                Exit Sub
            End If
        '*** FIN PEAC ************************************************************
    'WIOR 20140123 *******************************************
    Set oPers = New COMDPersona.UCOMPersona
        If oPers.fgVerificaEmpleadoVincualdo(lrPersonas(0)) Then
            MsgBox "Este es un Crédito Vinculado...Pariente de Empleado", vbInformation, "Aviso"
            bTrabajadoVinc = True
            MsgBox "Se realizará Automáticamente una solicitud de Saldo al Área de Administración de Créditos", vbInformation, "Aviso"
        End If
    Set oPers = Nothing
    'WIOR FIN ************************************************
    'ALPA 20150617************************************************************
        Dim oCredPers As COMDCredito.DCOMCredito
        Set oCredPers = New COMDCredito.DCOMCredito
        Dim oRsVal As ADODB.Recordset
        Set oRsVal = New ADODB.Recordset
        Set oRsVal = oCredPers.RecValidaProcentajeCredito(gdFecSis, gsCodAge, CDbl(txtMontoPrestamo.Text), sLineaTmp)
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
    'FRHU ERS077-2015 20151204
    Do While Not lrPersonas.EOF
        Call VerSiClienteActualizoAutorizoSusDatos(lrPersonas("cPersCod"))
        lrPersonas.MoveNext
    Loop
    lrPersonas.MoveFirst
    'FIN FRHU ERS077-2015 20151204
    '*************************************************************************
    
    '**ARLO20171204
    Dim nEgresos As Double
    Dim nIngresos As Double
    Dim nVerifica As Integer
    Dim R As ADODB.Recordset
    Dim RB As ADODB.Recordset
    Dim RDatFin As ADODB.Recordset
    Dim oCred As COMNCredito.NCOMCredDoc
    Dim oCredB As COMNCredito.NCOMCredDoc
    Dim oDCredC As COMDCredito.DCOMCredito
    'SIEMPRE QUE SUPERE EL 80% DE LA TASACION
    
    fnPorcentaje = val(txtMontoPrestamo.Text) / val(lnValTasacion)
    
    If (fnPorcentaje > 0.8) Then
        Call frmColPformatoEval.Inicio(val(lblInteres.Caption), val(lblNetoRecibir.Caption), fnTpoClinte)
        nEgresos = val(Replace(frmColPformatoEval.txtEgresos, ",", ""))
        nIngresos = val(Replace(frmColPformatoEval.txtIngresos, ",", ""))
        nVerifica = val(Replace(frmColPformatoEval.txtVerfica, ",", ""))
        If nVerifica = 0 Then
            Exit Sub
        End If
    End If
    '**************
    
    If MsgBox("¿Grabar Contrato Prestamo Pignoraticio ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        'Genera Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        
        Set loRegPig = New COMNColoCPig.NCOMColPContrato
            If fbMalCalificacion = False Then
            
                lsContrato = loRegPig.nRegistraContratoPignoraticioDetalle(gsCodCMAC & gsCodAge, gMonedaNacional, _
                lrPersonas, fnTasaInteresAdelantado, lnMontoPrestamo, lsFechaHoraGrab, lnPlazo, _
                lsFechaVenc, lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lsTipoContrato, _
                lsLote, ln14k, ln16k, ln18k, ln21k, lsMovNro, lnIntAdelantado, lnCostoTasac, _
                lnCostoCustodia, lnImpuesto, lrJoyas, pnMovNro, , , bTrabajadoVinc, , , , , , sLineaTmp, lnTasaGracia, lnTasaMorato, cboTasador.Text, CLng(txtHolograma.Text)) 'WIOR 20140203 AGREGO bTrabajadoVinc/ JUCS AGREGÓ cboTasador TI-ERS 063-2017
            
            
                'JUCS TI ERS 063-2017
                Dim conta As Integer
                Dim oPig As COMDColocPig.DCOMColPContrato
                Dim nMovimiento As String
                Dim contador As String
                Set oPig = New COMDColocPig.DCOMColPContrato
                Set R = oPig.ObtieneContador(gsCodAge)
                nMovimiento = R!cMovNro
                contador = R!nContador + 1
                Call oPig.ActualizaContador(contador, nMovimiento)
                '**ARLO20171207
            
                loRegPig.RegistroFormatoEvalPig lsContrato, nIngresos, nEgresos, lsMovNro, fnTpoClinte, nCantAdjuSegmento, nDiasSegmento, fcSegmento, nDiasSinAdj, nSegmentoAnt, vArrayDatosSegPred 'Agrego JOEP20180220 , fnTpoClinte, nCantAdju, nDiasSegmentacion, fcSegmento: joep pig pase nDiasSinAdj,nSegmentoAnt
                
                'JOEP20210913 campana Prendario
                Set oPig = New COMDColocPig.DCOMColPContrato
                If cboCampPrendario.Text <> "" Then
                    oPig.CampPrenRegCampCred lsContrato, Right(cboCampPrendario.Text, 3), txtCampPrendario.Text, Replace(lblPorcentajeTasa, "%", ""), 0, 0, 0, lsMovNro, 0, 0
                End If
                'JOEP20210913 campana Prendario
                
                If (fnPorcentaje > 0.8) Then
                    Set oCred = New COMNCredito.NCOMCredDoc
                    Call oCred.RecuperaDatosInformeComercial(lsContrato, R)
                    Set oCred = Nothing
                    
                    If R.EOF And R.BOF Then
                    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
                    Exit Sub
                    End If
                    
                    Set oCredB = New COMNCredito.NCOMCredDoc
                    Call oCredB.RecuperaDatosBalance(lsContrato, RB)
                    Set oCredB = Nothing
        
                    Set oDCredC = New COMDCredito.DCOMCredito
                    Set RDatFin = oDCredC.RecuperaDatosFinan(lsContrato)
                    Set oDCredC = Nothing
        
                    'Call ImprimeInformeComercial02(lsContrato, gsNomAge, gsCodUser, R, RB, RDatFin) ---COMENTADO POR ARLO20171218
                    Call ImprimeInformeComercialPig(lsContrato, gsNomAge, gsCodUser, R, RB, RDatFin) '---AGREGADO DESDE LA 60
                End If
                
                Set oCred = Nothing
                '***********************************
            Else
    
    '            lsContrato = loRegPig.nRegistraContratoPignoraticioDetalle(gsCodCMAC & gsCodAge, gMonedaNacional, _
    '            lrPersonas, fnTasaInteresAdelantado, lnMontoPrestamo, lsFechaHoraGrab, lnPlazo, _
    '            lsFechaVenc, lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lsTipoContrato, _
    '            lsLote, ln14k, ln16k, ln18k, ln21k, lsMovNro, lnIntAdelantado, lnCostoTasac, _
    '            lnCostoCustodia, lnImpuesto, lrJoyas, pnMovNro, fbMalCalificacion, gnTasaIntVencidoCrePigCliMalCal)
                'MsgBox ""
            End If
    
            pbTran = False
        Set loRegPig = Nothing
        
        '' *** PEAC 20090126
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , lsContrato, gCodigoCuenta
        '****** MACM 2021-03-18
        If lnTasaPreferencial <= 0 Then
        Else
            Dim objCargos As COMDCredito.DCOMNivelAprobacion
            Set objCargos = New COMDCredito.DCOMNivelAprobacion
            Call objCargos.GetActualizarColocacPermisoAprobacion(lsContrato, lsMovNro, "", gsCodPersUser, "", "", lnTasaPreferencial, 0, 1, 1)
            Set objCargos = Nothing
        End If
        '****** FIN MACM 2021-03-18
        MsgBox "Se ha generado Contrato Nro " & lsContrato, vbInformation, "Aviso"
        
        'JOEP ERS047 20170904
        Set objPigLimite = New COMDColocPig.DCOMColPContrato
        If objPigLimite.VerificaLimiteZonaGeo(gsCodAge, CDbl(lnMontoPrestamo), gMonedaNacional) Then
            MsgBox "El Crédito supera el porcentaje máximo por Zona Geográfica; se podrá continuar, pero no se podrá desembolsar si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation, "Aviso"
            Call objPigLimite.RegistroLimiteZonaGeog(lsContrato, CDbl(lnMontoPrestamo))
        Else
            Call objPigLimite.DeleteLimiteTpProducto(lsContrato, Mid(lsContrato, 6, 3), 1)
        End If
        Set objPigLimite = Nothing
        '====
        Set objPigLimite = New COMDColocPig.DCOMColPContrato
        If objPigLimite.VerificaLimiteTpProducto(Mid(lsContrato, 6, 3), lnMontoPrestamo, gMonedaNacional) Then
            MsgBox "El Crédito supera el porcentaje máximo por Producto Pignoraticio; se podrá continuar, pero no se podrá desembolsar si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation, "Aviso"
            Call objPigLimite.RegistroLimiteTpProducto(lsContrato, Mid(lsContrato, 6, 3), CDbl(lnMontoPrestamo))
        Else
            Call objPigLimite.DeleteLimiteTpProducto(lsContrato, Mid(lsContrato, 6, 3), 2)
        End If
        Set objPigLimite = Nothing
        'JOEP ERS047 20170904
    
        'JOEP20180412 Pig Pase
         Set objPigSalCatSalCap = New COMDColocPig.DCOMColPContrato
         'Set rsDatosSalCartSalCap = New ADODB.Recordset
         Call objPigSalCatSalCap.ObtieneSaldoCartSaldoCap(lsContrato)
         Set objPigSalCatSalCap = Nothing
        'JOEP20180412 Pig Pase
        If vSolTasa And Not vTasaForm Then
            If MsgBox("Desea Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                Call loRegImp.RecuperaDatosHojaResumenPigno(lsContrato, rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
                Call CargaHojaResumenPignoPDF(rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
                    
                'INICIO EAAS20190516 SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
                Pig_ContratosAutomaticos rsPigPers, rsPig, lsContrato
                
            End If
        Else
            If vTasaForm Then
                '***Agregado porel 20120104, según Acta N° 002-2012/TI-D
                fbMalCalificacion = False
                lblCalificacionNormal = ""
                lblCalificacionPotencial = ""
                lblCalificacionDeficiente = ""
                lblCalificacionDudoso = ""
                lblCalificacionPerdida = ""
                lblPorcentajeTasa = ""
                '***Fin Agregado por ELRO*************************************
                '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
                fbClienteCPP = False
                '*****************************END RECO************************
                
                '*** PEAC 20161220
                Set loPrevio = Nothing
                Set loRegPig = Nothing
                Limpiar
                '*** FIN PEAC
            Else
                If MsgBox("Desea Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    Set loRegImp = New COMNColoCPig.NCOMColPImpre
                    'Call loRegImp.RecuperaDatosHojaResumenPigno(lsContrato, rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
                    'Call CargaHojaResumenPignoPDF(rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot)
                    
                    'JOEP20210927 campana prendario
                    nCampana = 0
                    Call loRegImp.RecuperaDatosHojaResumenPigno(lsContrato, rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot, nCampana)
                    Call CargaHojaResumenPignoPDF(rsPig, rsPigJoyas, rsPigPers, rsPigCostos, rsPigDet, rsPigTasas, rsPigCosNot, nCampana)
                    'JOEP20210927 campana prendario
                    
                    'INICIO EAAS20190516 SEGUN Memorándum Nº 756-2019-GM-DI/CMACM
                    Pig_ContratosAutomaticos rsPigPers, rsPig, lsContrato
        
                    loVistoElectronico.RegistraVistoElectronico (pnMovNro)
                    '***Agregado porel 20120104, según Acta N° 002-2012/TI-D
                    fbMalCalificacion = False
                    lblCalificacionNormal = ""
                    lblCalificacionPotencial = ""
                    lblCalificacionDeficiente = ""
                    lblCalificacionDudoso = ""
                    lblCalificacionPerdida = ""
                    lblPorcentajeTasa = ""
                    '***Fin Agregado por ELRO*************************************
                    '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
                    fbClienteCPP = False
                    '*****************************END RECO************************
                    
                    '*** PEAC 20161220
                    Set loPrevio = Nothing
                    Set loRegPig = Nothing
                    Limpiar
                    '*** FIN PEAC
                    
                End If
            End If
        End If
    End If
End If
'*** PEAC 20161220
'Set loPrevio = Nothing
'Set loRegPig = Nothing
'Limpiar
'*** FIN PEAC

Exit Sub

ControlError:   ' Rutina de control de errores.
    'Verificar que se halla iniciado transaccion y la cierra
    'If pbTran Then dbCmact.RollbackTrans
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    Limpiar
End Sub

'**DAOR 20070115, Imprimir Bolante de Tasación
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
    lnOroNeto = val(lblOroNeto.Caption)
    lnPiezas = val(txtPiezas.Text)
    lnValTasacion = CCur(lblValorTasacion.Caption)
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



Private Sub cmdLineaCredito_Click()
    'frmColPLineaCreditoSelecciona.Show 1
End Sub

Private Sub CmdPiezaAgregar_Click()
'*****Agregado MPBR
    vTasaForm = False 'MACM 26032021
    If lstCliente.ListItems.count = 0 Then
       MsgBox "No puede agregar items, si no figura(n) cliente(s) en el contrato.", vbOKOnly + vbInformation, "Atención"
       Exit Sub
    End If

    '*** PEAC 20170509 - NO HAY LIMITE PARA REGISTRAR PIEZAS
'    If FEJoyas.Rows <= 20 Then
        If lnJoyas > 1 And FEJoyas.TextMatrix(FEJoyas.row, 6) = "" Then
            MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
            FEJoyas.SetFocus
            Exit Sub
        Else
            If lnJoyas = 1 And FEJoyas.TextMatrix(FEJoyas.row, 6) = "" Then
                MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
                Exit Sub
            Else
                lnJoyas = lnJoyas + 1
                FEJoyas.AdicionaFila
                If FEJoyas.rows >= 2 Then
                   cmdPiezaEliminar.Enabled = True
                End If
                
                'Dim loConst As COMDConstantes.DCOMConstantes comentado por pti1 07-03-2019
                Dim lrMaterial As New ADODB.Recordset
                'Set loConst = New COMDConstantes.DCOMConstantes comentado por pti1 07-03-2019
                
                FEJoyas.Col = 2
                'Set lrMaterial = loConst.RecuperaConstantes(gColocPMaterialJoyas, , "C.cConsDescripcion") 'comentado por pti1 06-03-2019
                Set lrMaterial = rsRecuperaConstantes.Clone 'add pti1 06-03-2019
                FEJoyas.CargaCombo lrMaterial
                Set lrMaterial = Nothing
                FEJoyas.Col = 1
                FEJoyas.SetFocus
            End If
        End If
'    Else
'        CmdPiezaAgregar.Enabled = False
'        MsgBox "Sólo puede ingresar como máximo veinte piezas", vbInformation, "Aviso"
'    End If
End Sub

Private Sub cmdPiezaEliminar_Click()
    FEJoyas.EliminaFila FEJoyas.row
    If FEJoyas.rows <= 20 Then
        CmdPiezaAgregar.Enabled = True
    End If
    If lnJoyas > 0 Then 'add pti1
    lnJoyas = lnJoyas - 1
    'lnJoyas = lnJoyas + 1 'comentado pot pti1 07-03-2019
    SumaColumnas
    Call CargarDatosProductoCrediticio
    Call MostrarLineas
    Call ObtenerCalificacion(fscPerCod) 'ARLO20180131
    AsignarTipoFormalidadGarantia (txtMontoPrestamo.Text) 'CROB20180611
    End If
'    comentado por  pti1 pti1 07-03-2019
'    lnJoyas = lnJoyas + 1
'    SumaColumnas
'    Call CargarDatosProductoCrediticio
'    Call MostrarLineas
'    Call ObtenerCalificacion(fscPerCod) 'ARLO20180131
'    AsignarTipoFormalidadGarantia (txtMontoPrestamo.Text) 'CROB20180611
'    fin comentado
End Sub

Private Sub CmdPrevio_Click()
    Dim pbTran As Boolean
    Dim lsCtaReprestamo As String
    Dim oImp As New COMNColoCPig.NCOMColPImpre

    Dim lrPersonas As New ADODB.Recordset
    Dim lsMovNro As String
    Dim lsFechaHoraGrab As String
    Dim lnMontoPrestamo As Currency
    Dim lnPlazo As Integer
    Dim lsFechaVenc As String
    Dim lnOroBruto As Double
    Dim lnOroNeto As Double
    Dim lnPiezas As Integer
    Dim lnValTasacion As Currency
    Dim lsTipoContrato As String
    Dim lsLote As String
    Dim ln14k As Double, ln16k As Double, ln18k As Double, ln21k As Double
    Dim lnIntAdelantado As Currency, lnCostoTasac As Currency, lnCostoCustodia As Currency, lnImpuesto As Currency

    Dim loRegPig As COMNColoCPig.NCOMColPContrato

    Dim loContFunct As COMNContabilidad.NCOMContFunciones

    Dim lsContrato As String
    Dim loPrevio As previo.clsprevio

    Dim lsCadImprimir As String
    On Error GoTo ControlError
    pbTran = False

    If ValidaDatosGrabar = False Then Exit Sub

    'Asigno los valores a los parametros

    Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.lstCliente)
    'txtDescLote = fgEliminaEnters(txtDescLote) & vbCr

    lnMontoPrestamo = CCur(txtMontoPrestamo.Text)
    lnPlazo = val(cboPlazo.Text)
    lsFechaVenc = Format$(Me.lblFechaVencimiento, "mm/dd/yyyy")
    lnValTasacion = CCur(lblValorTasacion.Caption)
    lsTipoContrato = Switch(cboTipcta.ListIndex = 0, "I", cboTipcta.ListIndex = 1, "O", cboTipcta.ListIndex = 2, "Y")
    lnIntAdelantado = CCur(Me.lblInteres.Caption)
'COMENTÓ APRI20190515 SATI
    'lnCostoTasac = CCur(Me.lblCostoTasacion.Caption)
    'lnCostoCustodia = CCur(Me.lblCostoCustodia.Caption)
    'lnImpuesto = CCur(Me.lblImpuesto.Caption)


    'Genera Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing

    lsFechaHoraGrab = Format(Now, "dd/mm/yyyy")     'hora

    lsContrato = "INFORMATIVA"


    lsCadImprimir = oImp.PrintHojaInformativa(lsContrato, lrPersonas, fnTasaInteresAdelantado, _
        lnMontoPrestamo, lsFechaHoraGrab, Format(lsFechaVenc, "mm/dd/yyyy"), lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, _
        lnPiezas, lsLote, ln14k, ln16k, ln18k, ln21k, lnIntAdelantado, lnCostoTasac, lnCostoCustodia, lnImpuesto, gsCodUser)
    Set oImp = Nothing
    
    Set loPrevio = New previo.clsprevio
        loPrevio.PrintSpool sLpt, lsCadImprimir, False

    Do While True
        If MsgBox("Reimprimir Hoja Informativa ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            loPrevio.PrintSpool sLpt, lsCadImprimir, False
        Else
            Set loPrevio = Nothing
            Set loRegPig = Nothing
            Exit Do
        End If
    Loop
    
    Set loPrevio = Nothing
    Set loRegPig = Nothing

    Exit Sub

ControlError:   ' Rutina de control de errores.
    'Verificar que se halla iniciado transaccion y la cierra
    'If pbTran Then dbCmact.RollbackTrans
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    'Limpiar


End Sub
'MACM 22-03-2021
Private Sub cmdSolTasa_Click()

Dim lsSolicitudes As New ADODB.Recordset
Dim lsPers As COMDPersona.DCOMPersonas
Set lsPers = New COMDPersona.DCOMPersonas
'Set lsPers = New ADODB.Recordset
Dim lrDocPersona As New ADODB.Recordset
Dim ncPersCod As String
Dim lsCodDeudorRCC As String
pcCtaCod = ""
vSolTasa = False
vNuevaSol = False
vEstadoSolTasa = 4

frmColPSolTasaPendiente.Inicio cc
pcCtaCod = cc
        
If pcCtaCod <> "" Then
            'MsgBox "Nombre del framework " & Me.lblOroBruto, vbCritical, "Aviso"
        
        Dim lsSolicitud As COMDCredito.DCOMNivelAprobacion
        Set lsSolicitud = New COMDCredito.DCOMNivelAprobacion
        Set lsSolicitudes = lsSolicitud.RecuperaSolTasaPendienteCuenta(pcCtaCod)
        vEstadoSolTasa = lsSolicitudes!nEstado
        ncPersCod = lsSolicitudes!cPersCod
        lnTasaSol = lsSolicitudes!nTasaApr
        If lsSolicitudes!nTasaApr <> 0 Then
            chkTasaEspeci.value = 1
            chkTasaEspeci.Enabled = False
            txtTasaEspeci.Enabled = False
            txtTasaEspeci.Text = Format(lsSolicitudes!nTasaApr, "#0.0000")
            txtHolograma.Text = lsSolicitudes!nHolograma
            cboTipcta.Enabled = False
            'Me.cboTasador.AddItem (lsSolicitudes!tasador)
            Me.cboTasador.Text = lsSolicitudes!Tasador
            CmdPiezaAgregar.Enabled = False
            cmdPiezaEliminar.Enabled = False
            Me.cmdGrabar.Enabled = False
            vSolTasa = True
            FEJoyas.Enabled = False
            Me.cmdCancelar.SetFocus
            Me.lstCliente.Enabled = False
            Me.ChkAnterior.Enabled = False
        Else
            txtTasaEspeci.Enabled = True
            txtTasaEspeci.Text = ""
            txtHolograma.Text = ""
            txtTasaEspeci.Enabled = True
            CmdPiezaAgregar.Enabled = True
            cmdPiezaEliminar.Enabled = True
            chkTasaEspeci.Enabled = True
            cboTipcta.Enabled = True
            cmdGrabar.Enabled = True
            vSolTasa = False
            FEJoyas.Enabled = False
            Me.lstCliente.Enabled = True
            Me.ChkAnterior.Enabled = True
        End If
   
        Set lrDocPersona = lsPers.RecuperaDatosPersona_Basic(ncPersCod)
        If vEstadoSolTasa = 1 Then
            MsgBox "El crédito tiene una solicitud de Tasa preferencial pendiente", vbInformation, "Aviso"
        ElseIf vEstadoSolTasa = 0 Then
            MsgBox "La solicitud de Tasa preferencial del " & lsSolicitudes!nTasaApr & "% Fue Rechazada", vbInformation, "Aviso"
            'Exit Sub
            chkTasaEspeci.value = 0
            chkTasaEspeci.Enabled = True
            txtTasaEspeci.Enabled = False
        
        Else
            Me.cmdAgregar.Enabled = False
            MsgBox "Su solicitud de tasa preferencial fue aprobada", vbInformation, "Aviso"
        End If
        Set lsSolicitud = Nothing
        Call calificacionSBS(ncPersCod, IIf(lrDocPersona!nPersPersoneria = "1", True, False), IIf(lrDocPersona!nPersPersoneria = "1", Trim(lrDocPersona!Dni), Trim(lrDocPersona!Ruc)), lsCodDeudorRCC = "")
        Call BuscaContrato(pcCtaCod)
End If
End Sub
'Finaliza la ejecusión del formulario
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdVerRetasacion_Click()
    frmColPHistorialRetasacion.Inicio (AXCodCta.NroCuenta)
End Sub

Private Sub FEJoyas_Click()
Dim loConst As COMDConstantes.DCOMConstantes
Dim lrMaterial As New ADODB.Recordset
Set loConst = New COMDConstantes.DCOMConstantes
If sCpersTem <> "" Then
Select Case FEJoyas.Col
Case 2
    'Set lrMaterial = loConst.RecuperaConstantes(gColocPMaterialJoyas, , "C.cConsDescripcion") 'COMENTADO POR PTI1 06-03-2019
    Set lrMaterial = rsRecuperaConstantes.Clone
    FEJoyas.CargaCombo lrMaterial 'comentado por pti1 06-03-2019
    Set lrMaterial = Nothing 'comentado por pti1 06-03-2019
Case 5
    FEJoyas.SetFocus
    'RECO20140706******************************************
    MsgBox "No se permite seleccionar la celda, presione la tecla 'enter' en la celda anterior.", vbInformation, "Aviso"
    'RECO END**********************************************
    FEJoyas.Col = 6
    FEJoyas.row = FEJoyas.row
    SendKeys "{Enter}"
End Select

Set loConst = Nothing
Else
 MsgBox "Por favor agregar un cliente, presione el boton 'agregar' ", vbInformation, "Aviso"
End If
End Sub


Private Sub FEJoyas_GotFocus()
Select Case FEJoyas.Col
Case 5
    FEJoyas.SetFocus
    FEJoyas.Col = 6
    FEJoyas.row = FEJoyas.row
    SendKeys "{Enter}"
End Select
End Sub

Private Sub FEJoyas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FEJoyas.Col = 1 Then
        If FEJoyas.TextMatrix(FEJoyas.row, 7) <> "" Then
            CmdPiezaAgregar.SetFocus
            CmdPiezaAgregar_Click
        End If
    End If
    If FEJoyas.Col = 5 Then
        KeyAscii = 0
    End If
End If
If KeyAscii = 22 Then
    KeyAscii = 0
End If


End Sub

Private Sub FEJoyas_LostFocus()
    Dim i As Integer
    For i = 1 To FEJoyas.rows - 1
        'Call FEJoyas_OnCellChange(FEJoyas.row, 4)
        Call CalcularValorTasacion(i, 4)
    Next
    'RECO20140220 INC************************
    'FEJoyas.SetFocus
    'FEJoyas.Col = 6
    'FEJoyas.row = FEJoyas.row
    'SendKeys "{Enter}"
    'RECO FIN*********************************
End Sub

Private Sub FEJoyas_OnCellChange(pnRow As Long, pnCol As Long)
Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
Dim lnPOro As Double

''*** PEAC 20170131
'If Len(Trim(FEJoyas.TextMatrix(FEJoyas.row, 6))) >= 260 Then
'    MsgBox "Solo se permite 260 caracteres", vbInformation, "Aviso"
''    Cancel = False
'    Exit Sub
'End If


'******RECO 20131126*****************
Dim loColContrato As COMDColocPig.DCOMColPContrato
Dim lnValorPOro As Double
Dim lnMatOro As Integer
Dim loDR As ADODB.Recordset


Set loColContrato = New COMDColocPig.DCOMColPContrato
Set loDR = New ADODB.Recordset
'******END RECO**********************
    If FEJoyas.Col = 3 Then
        If FEJoyas.TextMatrix(FEJoyas.row, 3) = "" Then
            MsgBox "Ingrese un Peso Bruto Correcto", vbInformation, "Aviso"
            Exit Sub
        End If
        'lnPesoNetoDesc = CDbl(FEJoyas.TextMatrix(FEJoyas.row, 3)) * 0.1 'COMENTADO APRI20170623 segun SATI TIC1706220007
        'lnPesoNetoDesc = CDbl(FEJoyas.TextMatrix(FEJoyas.row, 3)) - lnPesoNetoDesc 'COMENTADO APRI20170623 segun SATI TIC1706220007
        FEJoyas.TextMatrix(FEJoyas.row, 4) = FEJoyas.TextMatrix(FEJoyas.row, 3) 'lnPesoNetoDesc
    End If
    
    If FEJoyas.Col = 4 Then     'Peso Neto

    'RECO***********
        'lnPesoNetoDesc = CDbl(FEJoyas.TextMatrix(FEJoyas.row, 3)) * 0.1 'COMENTADO APRI20170623 segun SATI TIC1706220007
        If FEJoyas.TextMatrix(FEJoyas.row, 3) <> "" Then
            lnPesoNetoDesc = CDbl(FEJoyas.TextMatrix(FEJoyas.row, 3)) '- lnPesoNetoDesc 'COMENTADO APRI20170623 segun SATI TIC1706220007
        End If
    'END RECO*******
        If FEJoyas.TextMatrix(FEJoyas.row, 4) <> "" Then
            If CCur(FEJoyas.TextMatrix(FEJoyas.row, 4)) < 0 Then
                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(FEJoyas.row, 4) = 0
            Else
                If CCur(FEJoyas.TextMatrix(FEJoyas.row, 4)) > lnPesoNetoDesc Then
                    'MsgBox "Peso Neto " & CCur(FEJoyas.TextMatrix(psFila, 4)) & " debe ser menor a peso neto base " & lnPesoNetoDesc, vbInformation, "Aviso" 'COMENTADO APRI20170623 segun SATI TIC1706220007
                    MsgBox "Peso Neto " & CCur(FEJoyas.TextMatrix(FEJoyas.row, 4)) & " no debe ser mayor al Peso Bruto " & lnPesoNetoDesc, vbInformation, "Aviso"
                    FEJoyas.TextMatrix(FEJoyas.row, 4) = lnPesoNetoDesc
                Else
                    'CalculaTasacion
                        Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, val(Left(FEJoyas.TextMatrix(FEJoyas.row, 2), 2)), 1) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2)
                        '********RECO 20131126 ERS158****************
                        'lnMatOro = Left(FEJoyas.TextMatrix(FEJoyas.row, 2), 2) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2)
                        lnMatOro = IIf(Left(FEJoyas.TextMatrix(FEJoyas.row, 2), 2) = "", 0, Left(FEJoyas.TextMatrix(FEJoyas.row, 2), 2)) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2) 'Arlo20171214 ERS082-2017
                           
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
                        '********END RECO****************************
                        If lnPOro <= 0 Then
                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
                            Exit Sub
                        End If
                        Set loColPCalculos = Nothing
                        'Calcula el Valor de Tasacion
                        '********RECO 20131126 ERS158**********
                        FEJoyas.TextMatrix(FEJoyas.row, 5) = Format$(val(FEJoyas.TextMatrix(FEJoyas.row, 4) * lnValorPOro), "#####.00")
                        'FEJoyas.TextMatrix(FEJoyas.row, 5) = Format$(val(FEJoyas.TextMatrix(FEJoyas.row, 4) * lnPOro), "#####.00")
                        '********END RECO**********************
                End If
            End If
        End If
    '
    End If
    If FEJoyas.Col = 6 Then     'Descripcion

        If FEJoyas.TextMatrix(FEJoyas.row, 6) <> "" Then
            'cboPlazo.Enabled = False 'True RECO20140208 ERS002
        End If
        
    End If
    SumaColumnas
    If FEJoyas.Col = 4 Then
    Call txtMontoPrestamo_KeyPress(13) 'ALPA 20140616*****************
    End If

Call ObtenerCalificacion(fscPerCod)

'COMENTADO POR ARLO20180131
''Inicio **ARLO20171204 ERS082-2017
'Dim fcSegmento As String
'Dim pSegDR As ADODB.Recordset
'Dim psConfDR  As ADODB.Recordset
'Dim psVeriDR As ADODB.Recordset
'Dim psSegConDR As ADODB.Recordset
'
'
'Dim loPigContrato As COMDColocPig.DCOMColPContrato
'Set loPigContrato = New COMDColocPig.DCOMColPContrato
'
'Set psVeriDR = loPigContrato.dVerificaTpoCliente(fscPerCod)
'    If Not (psVeriDR.BOF And psVeriDR.EOF) Then
'        fnTpoClinte = 2 'CLIENTE RECURRENTE
'        Me.lblClienteTpo.Caption = "Cliente Recurrente" 'ARLO20171218 ---AGREGADO DESDE LA 60
'    Else
'        fnTpoClinte = 1 'CLIENTE NUEVO
'        Me.lblClienteTpo.Caption = "Cliente Nuevo" 'ARLO20171218    --AGREGADO DESDE LA 60
'    End If
'
'Set psConfDR = loPigContrato.dVerificarConfSegmentacion(fnTpoClinte)
'
'
'
'Set pSegDR = loPigContrato.dVerificarSegmentacion(fscPerCod)
'
'If Not (pSegDR.BOF And pSegDR.EOF) Then '---INCIO MODIFICADO DESDE LA 60
'
'    If (fnTpoClinte = 2) Then
'            Do While Not psConfDR.EOF
'
'                If (pSegDR!nCantAdjuticados) > 0 Then
'                    fcSegmento = "D1"
'                    Exit Do
'                Else
'                    If (psConfDR!cSubSegmento) = "A1" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And pSegDR!nDias > psConfDR!nDiasDesde And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "A1"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "A2" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And pSegDR!nDias >= psConfDR!nDiasDesde And pSegDR!nDias <= psConfDR!nDiasHasta And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "A2"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "A3" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And pSegDR!nDias < psConfDR!nDiasHasta And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "A3"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "B1" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!nDias > psConfDR!nDiasDesde And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "B1"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "B2" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!nDias >= psConfDR!nDiasDesde And pSegDR!nDias <= psConfDR!nDiasHasta And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "B2"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "B3" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!nDias < psConfDR!nDiasHasta And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "B3"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "C1" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!nDias >= psConfDR!nDiasDesde And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "C1"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "C2" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!nDias >= psConfDR!nDiasDesde And pSegDR!nDias < psConfDR!nDiasHasta And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "C2"
'                            Exit Do
'                        End If
'                    ElseIf (psConfDR!cSubSegmento) = "C3" Then
'                        If (val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!nDias < psConfDR!nDiasHasta And psConfDR!nCantAdjudicado = 0) Then
'                            fcSegmento = "C3"
'                            Exit Do
'                        End If
'                    End If
'                End If
'
'                psConfDR.MoveNext
'
'            Loop
'    Else
'
'        Do While Not psConfDR.EOF
'
'
'                If (psConfDR!cSubSegmento) = "A3" Then
'                    If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And pSegDR!RC18 = 100 And pSegDR!RC17 = 100 _
'                    And pSegDR!RC16 = 100 And pSegDR!RC15 = 100 And pSegDR!RC14 = 100 And pSegDR!RC13 = 100 And pSegDR!RC12 = 100 And pSegDR!RC11 = 100 _
'                    And pSegDR!RC10 = 100 And pSegDR!RC09 = 100 And pSegDR!RC08 = 100 And pSegDR!RC07 = 100) Then
'                        fcSegmento = "A3"
'                        Exit Do
'                End If
'                ElseIf (psConfDR!cSubSegmento) = "B3" Then
'                    If (val(Replace(Me.lblValorTasacion, ",", "")) >= psConfDR!nMontoTasaDesde And val(Me.lblValorTasacion) <= psConfDR!nMontoTasaHasta And pSegDR!RC18 = 100 And pSegDR!RC17 = 100 _
'                    And pSegDR!RC16 = 100 And pSegDR!RC15 = 100 And pSegDR!RC14 = 100 And pSegDR!RC13 = 100 And pSegDR!RC12 = 100 And pSegDR!RC11 = 100 _
'                    And pSegDR!RC10 = 100 And pSegDR!RC09 = 100 And pSegDR!RC08 = 100 And pSegDR!RC07 = 100) Then
'                        fcSegmento = "B3"
'                        Exit Do
'                    End If
'                ElseIf (psConfDR!cSubSegmento) = "C3" Then
'                    If (val(Replace(Me.lblValorTasacion, ",", "")) <= psConfDR!nMontoTasaHasta And pSegDR!RC18 = 100 And pSegDR!RC17 = 100 _
'                    And pSegDR!RC16 = 100 And pSegDR!RC15 = 100 And pSegDR!RC14 = 100 And pSegDR!RC13 = 100 And pSegDR!RC12 = 100 And pSegDR!RC11 = 100 _
'                    And pSegDR!RC10 = 100 And pSegDR!RC09 = 100 And pSegDR!RC08 = 100 And pSegDR!RC07 = 100) Then
'                        fcSegmento = "C3"
'                        Exit Do
'                    End If
'                ElseIf (psConfDR!cSubSegmento) = "D1" Then
'                    fcSegmento = "D1"
'                End If
'
'
'            psConfDR.MoveNext
'
'        Loop
'
'    End If
'
'End If
'
''---FIN MODIFICADO DESDE LA 60
'
'Set psSegConDR = loPigContrato.dObtieneValorSegmentacion(fcSegmento)
'
'If Not (psSegConDR.BOF And psSegConDR.EOF) Then
'
'        fnPorcentajePrestamo = (psSegConDR!nMontoTasa) / 100
'        lblCalificacion.Caption = "Cliente " + fcSegmento
'Else
'    MsgBox "No existe configuración para el cliente", vbInformation, "Aviso"
'    fnPorcentajePrestamo = 0
'    lblCalificacion.Caption = ""
'    Exit Sub
'End If
'lblCalificacion.Caption = "Cliente " + fcSegmento
'Fin ARLO20171204 ERS082-2017**************
'****ARLO 20180131
End Sub

Private Sub FEJoyas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(FEJoyas.ColumnasAEditar, "-")
    If pnCol = 1 And FEJoyas.TextMatrix(pnRow, pnCol) = "0" Then
        Cancel = False
        MsgBox "Cero (0) no es un valor válido", vbInformation, "Aviso"
        SendKeys "{Tab}"
    End If
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    
End Sub

'Inicializa el formulario
Private Sub Form_Load()
    CargaParametros
    Limpiar
    lblNetoRecibir.ForeColor = pColPriEgreso
    lblNetoRecibir.BackColor = pColFonSoles
    
    'fsColocLineaCredPig = "0101113050101"
    'MAVM 20100609 BAS II
    fsColocLineaCredPig = "0101117550101"
     
    lsTextoCalif = "Dudoso, Pérdida o Deficiente" '*** PEAC 20170727
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gPigRegistrarContrato
    Call CargaPlazo 'RECO20140421
    nVerificaVencidos = 0 'JOEP20180412 Pig
    nCalificacionPotencialCPP = 0 'JOEP20180412 Pig Pase
    Call obtieneTasador 'JUCS11122017
    'JOEP20210913 Campaña Prendario
    CargaCampPrendario
    fr_CampPrendario.Enabled = False
    fr_TasaEspecial.Enabled = False
    
    CmdPrevio.top = 8280
    cmdImpVolTas.top = 8280
    cmdGrabar.top = 8280
    cmdCancelar.top = 8280
    cmdSalir.top = 8280
    frmColPRegContratoDet.Height = 9150
'JOEP20210913 Campaña Prendario
End Sub
'JUCS TI-ERS 063-2017
Private Sub obtieneTasador()
Dim oCred As COMDColocPig.DCOMColPContrato
Dim R As ADODB.Recordset
   On Error GoTo Error
    Set oCred = New COMDColocPig.DCOMColPContrato
     Set R = oCred.obtieneTasador(gsCodAge)
    Do While Not R.EOF
            cboTasador.AddItem Trim(R!Usuario)
            R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCred = Nothing
    cboTasador.ListIndex = -1
    Exit Sub
Error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub 'FIN JUCS

Private Sub Reutilizables() 'ADD PTI1 06-03-2019
Dim loConst As COMDConstantes.DCOMConstantes
Dim loPigContrato As COMDColocPig.DCOMColPContrato


   On Error GoTo Error
    Set loConst = New COMDConstantes.DCOMConstantes
    Set loPigContrato = New COMDColocPig.DCOMColPContrato
    
    Set rsRecuperaConstantes = New ADODB.Recordset
    Set rsRecuperaSegDR = New ADODB.Recordset
    
    Set rsRecuperaConstantes = loConst.RecuperaConstantes(gColocPMaterialJoyas, , "C.cConsDescripcion")
    Set rsRecuperaSegDR = loPigContrato.dVerificarSegmentacion(Me.lstCliente.ListItems.Item(1))
    If Not (rsRecuperaSegDR.BOF And rsRecuperaSegDR.EOF) Then
        If rsRecuperaSegDR!nDias > 90 Then
            fnTpoClinte = 2 'CLIENTE RECURRENTE
            Me.lblClienteTpo.Caption = "Cliente Recurrente"
        Else
            fnTpoClinte = 1 'CLIENTE NUEVO
            Me.lblClienteTpo.Caption = "Cliente Nuevo"
        End If
    End If

    Set rsRecuperapsConfDR = loPigContrato.dVerificarConfSegmentacion(fnTpoClinte)
    
    Set loConst = Nothing
    Set loPigContrato = Nothing
    Exit Sub
Error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub 'FIN PTI1
Private Sub txtHolograma_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
       If txtHolograma.Text <> "" Then
        Call VerificaHolograma
       Else
        MsgBox "Debe registrar un numero de holograma para este contrato pignoraticio", vbInformation, "Aviso"
        txtHolograma.SetFocus
        Exit Sub
       End If
    Else
        Exit Sub
    End If
End Sub
Private Sub VerificaHolograma()
Dim oPig As COMDColocPig.DCOMColPContrato
Dim R As New ADODB.Recordset
Dim h As Boolean
Dim HologramaRegistro As Long
On Error GoTo Error
Set oPig = New COMDColocPig.DCOMColPContrato
Set R = oPig.ObtenerHologramaActivo(gsCodAge)
HologramaRegistro = CLng(txtHolograma.Text)
     If txtHolograma <> "" Then
      h = oPig.HologDuplicado(HologramaRegistro, gsCodAge) 'APRI20190515 SATI
     End If
'Validación de registro de Hologramas
      If R.BOF And R.EOF Then
        MsgBox " No se encuentra rango activo para esta agencia, comunique a su supervisor para su registro correspondiente.", vbInformation, " Aviso "
        txtHolograma = ""
        bErrHolog = True
        Exit Sub
      Else
            If h = True Then
                    MsgBox "No se puede registrar este número de holograma porque ya fue registrado.", vbQuestion, "Aviso"
                    txtHolograma = ""
                    bErrHolog = True
                    Exit Sub
            Else
                If HologramaRegistro <> 0 Then
                    If R!HologIni <= CDbl(HologramaRegistro) And R!HologFin >= CDbl(HologramaRegistro) Then
                         'MsgBox "Correcto ", vbInformation, "Aviso"  'aqui registraremos los hologramas
                         bErrHolog = False
                    Else
                        MsgBox "El monto ingresado no se encuentra dentro del rango establecido por el supervisor, verifique que el número de su holograma sea el correcto", vbInformation, "Aviso"
                        txtHolograma = ""
                        txtHolograma.SetFocus
                        bErrHolog = True
                    End If
                 End If
            End If
        
      End If
    Set R = Nothing
    Set oPig = Nothing
    'Set h = Nothing
    Exit Sub
Error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'JUCS TI ERS 063-2017********************************************
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumeros = KeyAscii
    If KeyAscii = 13 Then SoloNumeros = KeyAscii
End Function
'FIN JUCS *********************************************************

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
    '***MARG ERS046-2016***AGREGADO 20161109***
    gsOpeCod = ""
    '***END MARG*******************************
End Sub

'Valida el campo txtmontoprestamo
Private Sub txtMontoPrestamo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoPrestamo, KeyAscii, 10, 2)
    If KeyAscii = 13 Then
       CalculaCostosAsociados
       ' Calcula Porcentaje de Prestamo en Oro
       If val(lblOroNeto) = 0 Then
            MsgBox " Oro Neto no debe ser CERO ", vbInformation, " Aviso "
            'txt14k.SetFocus
       ElseIf val(txtMontoPrestamo) > val(Format(fnPorcentajePrestamo * val(lblValorTasacion.Caption), "#0.00")) Then
            MsgBox " Monto de Préstamo debe ser Menor o igual al " & (fnPorcentajePrestamo * 100) & "% (" & Format(fnPorcentajePrestamo * val(lblValorTasacion.Caption), "#,#0.00") & ") del Valor de Tasacion ", vbInformation, " Aviso "
            txtMontoPrestamo.SetFocus
       ElseIf val(txtMontoPrestamo) = 0 Or Len(Trim(txtMontoPrestamo)) = "" Then
            MsgBox " Ingrese Monto prestado ", vbInformation, " Aviso "
            txtMontoPrestamo.SetFocus
      ' ElseIf Val(txtMontoPrestamo) > fnMaxMontoPrestamo1 Then
      '      If MsgBox("Monto de Préstamo es Mayor al permitido, Avise al Administrador " & vbCr & _
      '      " Desea continuar con el Préstamo ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
      '          lblOroPrestamo.Caption = Format(Val(lblOroNeto.Caption) * Val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * Val(lblValorTasacion.Caption)), "#0.00")
      '          lblOroPrestamoPorcen.Caption = Format((Val(lblOroPrestamo) * 100) / Val(lblOroNeto), "#0")
      '          If vContAnte = True Then
      '              cmdGrabar.Enabled = True
      '              cmdGrabar.SetFocus
      '          Else
      '              txtDescLote.Enabled = True
      '              txtDescLote.SetFocus
      '          End If
      '      Else
      '          txtMontoPrestamo.SetFocus
      '      End If
       Else
            lblOroPrestamo.Caption = Format(val(lblOroNeto.Caption) * val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * val(lblValorTasacion.Caption)), "#0.00")
            lblOroPrestamoPorcen.Caption = Format((val(lblOroPrestamo) * 100) / val(lblOroNeto), "#0")
            If vContAnte = True Then
                cmdGrabar.Enabled = True
                cmdGrabar.SetFocus
            Else
                cmdGrabar.Enabled = True
                cmdGrabar.SetFocus
                'txtDescLote.Enabled = True
                'txtDescLote.SetFocus
            End If
            txtMontoPrestamo.Text = Format(txtMontoPrestamo.Text, "###0.00")
       End If
        AsignarTipoFormalidadGarantia (txtMontoPrestamo.Text) 'CROB20180611
        Call CargarDatosProductoCrediticio
        Call MostrarLineas
    End If
End Sub
Private Sub txtMontoPrestamo_Validate(Cancel As Boolean)
  CalculaCostosAsociados
   ' Calcula Porcentaje de Prestamo en Oro
   If val(lblOroNeto) = 0 Then
        MsgBox " Oro Neto no debe ser CERO ", vbInformation, " Aviso "
        Cancel = True
   
   ElseIf val(txtMontoPrestamo) > val(Format(fnPorcentajePrestamo * val(lblValorTasacion.Caption), "#0.00")) Then
'        MsgBox " Monto de Prestamo debe ser Menor al 60 % del Valor de Tasacion ", vbInformation, " Aviso "
        MsgBox " Monto de Préstamo debe ser Menor o igual al " & (fnPorcentajePrestamo * 100) & "% (" & Format(fnPorcentajePrestamo * val(lblValorTasacion.Caption), "#,#0.00") & ") del Valor de Tasacion ", vbInformation, " Aviso " '*** PEAC 20170126
        Cancel = True
   ElseIf val(lblOroNeto) = 0 Then
        MsgBox " Ingrese cantidad de ORO ", vbInformation, " Aviso "
        Cancel = True
   ElseIf val(txtMontoPrestamo) = 0 Or Len(Trim(txtMontoPrestamo)) = "" Then
        MsgBox " Ingrese Monto prestado ", vbInformation, " Aviso "
        Cancel = True
   'ElseIf Val(txtMontoPrestamo) > fnMaxMontoPrestamo1 Then
   '     If MsgBox("Monto de Préstamo es Mayor al permitido, Avise al Administrador " & vbCr & _
   '     " Desea continuar con el Préstamo ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
   '         lblOroPrestamo.Caption = Format(Val(lblOroNeto.Caption) * Val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * Val(lblValorTasacion.Caption)), "#0.00")
   '         lblOroPrestamoPorcen.Caption = Format((Val(lblOroPrestamo) * 100) / Val(lblOroNeto), "#0")
   '     Else
   '         Cancel = True
   '     End If
   Else
        lblOroPrestamo.Caption = Format(val(lblOroNeto.Caption) * val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * val(lblValorTasacion.Caption)), "#0.00")
        lblOroPrestamoPorcen.Caption = Format((val(lblOroPrestamo) * 100) / val(lblOroNeto), "#0")
   End If
End Sub

'Valida el campo txtPiezas
Private Sub txtPiezas_GotFocus()
    fEnfoque txtPiezas
End Sub

Private Sub txtPiezas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If fgTextoNum(txtPiezas) Then
            'cboPlazo.Enabled = False 'True RECO20140208 ERS002
            cboPlazo.SetFocus
        Else
            MsgBox " Ingrese número de piezas ", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub txtPiezas_Validate(Cancel As Boolean)
    'If Not TextoNum(txtPiezas) Then
    '    MsgBox " Ingrese número de piezas ", vbInformation, " Aviso "
    '    Cancel = True
    'End If
End Sub

'Valida el campo cboplazo
Private Sub cboPlazo_Click()
'********Modificacion MPBR
If val(txtMontoPrestamo.Text) = 0 And lstCliente.ListItems.count > 0 Then
    CalculaPrestamo
    txtMontoPrestamo.SetFocus
    Call txtMontoPrestamo_KeyPress(13)
End If
If val(txtMontoPrestamo.Text) <> 0 Then
    CalculaCostosAsociados
End If
End Sub
'*****Agregado MPBR
Private Sub CalculaPrestamo()
    vPrestamo = val(lblValorTasacion.Caption) * fnPorcentajePrestamo
    txtMontoPrestamo.Text = Format(vPrestamo, "#0.00")
    AsignarTipoFormalidadGarantia (vPrestamo) 'CROB20180611
    txtMontoPrestamo.Enabled = True
End Sub

Private Sub cboplazo_KeyPress(KeyAscii As Integer)
    'ValorTasacion = Val(lblValorTasacion.Caption)
    vPrestamo = val(lblValorTasacion.Caption) * fnPorcentajePrestamo
    txtMontoPrestamo.Text = Format(vPrestamo, "#0.00")
    AsignarTipoFormalidadGarantia (vPrestamo) 'CROB20180611
    txtMontoPrestamo.Enabled = True
    txtMontoPrestamo.SetFocus
End Sub

' Valida el campo txtvalortasación
Private Sub lblValorTasacion_Change()
If val(txtMontoPrestamo) <> 0 Then
    cmdGrabar.Enabled = False
    txtMontoPrestamo = Format(0, "#0.00")
End If
End Sub

Private Sub CmdEliminar_Click()
On Error GoTo ControlError
    Dim i As Integer, J As Integer
    If lstCliente.ListItems.count = 0 Then
       MsgBox "No existen datos, imposible eliminar", vbInformation, "Aviso"
       cmdEliminar.Enabled = False
       cboTipcta.Enabled = False
       lstCliente.SetFocus
       Exit Sub
    Else
       For i = 1 To lstCliente.ListItems.count
           If lstCliente.ListItems.Item(i) = lsIteSel Then
              lstCliente.ListItems.Remove (i)
              Exit For
           End If
       Next i
    End If
    lstCliente.SetFocus
    If lstCliente.ListItems.count = 0 Then
        lblOroBruto = Format(0, "#0.00")
        cboTipcta.Enabled = False
        cmdEliminar.Enabled = False
    ElseIf lstCliente.ListItems.count = 1 Then
        cboTipcta.ListIndex = 0
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite buscar un cliente por nombre y/o documento
Private Sub cmdAgregar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim lnSolicitudApr As Integer 'MACM 20210323
Dim ls As String
Dim loColPFunc As COMDColocPig.DCOMColPFunciones
'*********RECO20131120 ERS158 **************************
Dim loPigContrato As COMDColocPig.DCOMColPContrato
'*********END RECO**************************************
'On Error GoTo ControlError
Set loPers = New COMDPersona.UCOMPersona
vEstadoSolTasa = 4 'MACM 20210323
Me.lstCliente.Enabled = True
'***Agregado por ELRO el 20120103, según Acta N° 002-2012/TI-D
'Dim oDCOMCreditos As DCOMCreditos
'Set oDCOMCreditos = New DCOMCreditos
'Dim rsCalificacionSBS As ADODB.Recordset
'Set rsCalificacionSBS = New ADODB.Recordset
Dim lsCodDeudorRCC As String
'***Fin Agregado por ELRO*************************************

'*********RECO20131120 ERS158 **************************
Set loPigContrato = New COMDColocPig.DCOMColPContrato
'***************END RECO********************************
Set loPers = frmBuscaPersona.Inicio

vSolTasa = False

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    sCpersTem = loPers.sPersCod 'add pti1 07-03-2019
    
    '** APRI20170622 Verifica edad del cliente
    If loPers.sPersEdad < 18 Then
     MsgBox "El Cliente es menor de edad, no se le puede otorgar un Crédito Pignoraticio.", vbInformation, "Aviso"
           Exit Sub
    End If
    '** END APRI20170622
    '** Verifica que no este en lista
    For liFil = 1 To Me.lstCliente.ListItems.count
        If lsPersCod = Me.lstCliente.ListItems.Item(liFil).Text Then
           MsgBox " Cliente Duplicado ", vbInformation, "Aviso"
           Exit Sub
        End If
    Next liFil
    '** Maximo Nro de Clientes = 4
    If Me.lstCliente.ListItems.count > 4 Then
           MsgBox " Maximo Nro de Clientes ==> 4 ", vbInformation, "Aviso"
           Exit Sub
    End If
    'Verifica si es Empleado '(ICA-2004/02/02)
    If loPers.fgVerificaEmpleado(lsPersCod) = True Then
        MsgBox "El Cliente tambien es empleado de la CMAC,  " & _
               "NO puede tener creditos Pignoraticios", vbInformation, "Aviso"
        Exit Sub
    End If
    
    '***Agregado por ELRO el 20120104, según Acta N° 002-2012/TI-D
    If verificarCreditosCanceladosCastigados(lsPersCod, gdFecSis) Then
        'MsgBox "El Cliente tiene algún Credito Cancelado Castigado por la CMACM", vbInformation, "Aviso"
        MsgBox "El Cliente no esta sujeto a crédito por tener un historial crediticio Judicial y/o Castigado", vbInformation, "Aviso" 'APRI20190515 SATI
        Exit Sub
    End If
    '***Fin Agregado por ELRO
    
    'JUEZ 20130717 *********************************************
    Dim oPREDA As COMDPersona.DCOMPersonas
    Set oPREDA = New COMDPersona.DCOMPersonas
    If oPREDA.VerificarPersonaPREDA(lsPersCod, 1) Then
        MsgBox "El cliente " & loPers.sPersNombre & " es un cliente PREDA no sujeto de Crédito, consultar a Coordinador de Producto Agropecuario", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oPREDA = Nothing
    'END JUEZ **************************************************
    
    'MACM 2021-03-19 VERIFICA SOL DE TASA PREFERENCIAL
    Dim objSolicitud As COMDCredito.DCOMNivelAprobacion
    Set objSolicitud = New COMDCredito.DCOMNivelAprobacion
    
    If objSolicitud.GetCargarColocacPignoAprobacionPersona(lsPersCod, 1) Then
        MsgBox "El cliente cuenta con solicitudes de tasa preferencial pendiente", vbInformation, "Aviso"
    End If
    If objSolicitud.GetCargarColocacPignoAprobacionPersona(lsPersCod, 0) Then
        MsgBox "El cliente cuenta con solicitudes de tasa preferencial rechazado", vbInformation, "Aviso"
    End If
    'MACM FIN SOL DE TASA PREFERENCIAL
    
    '***PEAC 20080115
    Call BuscaCreditosVencidos(lsPersCod, gdFecSis)
     
    fscPerCod = lsPersCod 'ARLO20171205
    'Set lstTmpCliente = "Lista"
    Set lstTmpCliente = lstCliente.ListItems.Add(, , lsPersCod)
        lstTmpCliente.SubItems(1) = loPers.sPersNombre
        lstTmpCliente.SubItems(2) = loPers.sPersDireccDomicilio
        lstTmpCliente.SubItems(3) = loPers.sPersTelefono
        
        'lstTmpCliente.SubItems(4) = Trim(ClienteCiudad(loPers.sPersZona))
        'lstTmpCliente.SubItems(5) = Trim(ClienteZona(vCodZona))
        lstTmpCliente.SubItems(6) = gPersIdDNI
        lstTmpCliente.SubItems(7) = loPers.sPersIdnroDNI
        'lstTmpCliente.SubItems(8) = TipoDoTr(RegPersona!cTidotr & "")
        lstTmpCliente.SubItems(9) = loPers.sPersIdnroRUC
    
        Set loColPFunc = New COMDColocPig.DCOMColPFunciones
            lstTmpCliente.SubItems(4) = Trim(loColPFunc.dObtieneNombreZonaPersona(loPers.sPersCod))
            'lstTmpCliente.SubItems(5) = Trim(loColPFunc.dObtieneCiudadZona(loPers.svCodZona))
        Set loColPFunc = Nothing
    '******Agregado MPBR
    cmdEliminar.Enabled = True
    '** Comentado por DAOR 20070806 ***********************
    cboTipcta.Enabled = True
    
    If vSolTasa Then 'MACM 20210324
        txtHolograma.Enabled = False 'JUCS TI ERS 063-2017
        cboTasador.Enabled = False ' JUCS TI ERS 063-2017
        cboTipcta.Enabled = False 'MACM 20210324
    Else
        txtHolograma.Enabled = True 'JUCS TI ERS 063-2017
        cboTasador.Enabled = True ' JUCS TI ERS 063-2017
        cboTipcta.Enabled = True 'MACM 20210324
        CmdPiezaAgregar.Enabled = True 'MACM 20210324
        CmdPiezaAgregar.SetFocus 'MACM 20210324
    End If
    
'    txtHolograma.Enabled = True 'JUCS TI ERS 063-2017
'    cboTasador.Enabled = True ' JUCS TI ERS 063-2017
    
    Call calificacionSBS(lsPersCod, IIf(loPers.sPersPersoneria = "1", True, False), IIf(loPers.sPersPersoneria = "1", Trim(loPers.sPersIdnroDNI), Trim(loPers.sPersIdnroRUC)), lsCodDeudorRCC = "")
    '***Agregado por ELRO el 20120103, según Acta N° 002-2012/TI-D
    
'End If 'comentado por pti1

'*************TORE ERS054***********************************
    'Call MostrarObservacionesRetasacion(lsPersCod, "", 1) 'TORE RFC1811260001 -> Comentado
    'frmColPObservacionesRetasacion.Observaciones lsPersCod, "", 1
'*************END TORE***********************************



If lstCliente.ListItems.count >= 1 And cboTipcta.ListIndex = 0 Then
    cboTipcta.ListIndex = 1
    '**DAOR 20070806 *************************
    cmdAgregar.Enabled = False
    '*****************************************
End If
'*************RECO 20131121 ERS158*********************************
If lstCliente.ListItems.count >= 1 And cboTipcta.ListIndex = 0 Then
    Dim poDR As ADODB.Recordset
    Set poDR = New ADODB.Recordset
    
    
    '*** PEAC 20161216 - esta calif de cliente para otro pase
    'este será cambiado
    Set poDR = loPigContrato.dVerificarCredPignoAdjudicado(loPers.sPersCod)
    If Not (poDR.BOF And poDR.EOF) Then
        nTpoCliente = 1
    Else
        Set poDR = Nothing
        Set poDR = loPigContrato.dVerificarCredPignoDesembolso(loPers.sPersCod)
        If Not (poDR.BOF And poDR.EOF) Then
            nTpoCliente = 2
        Else
            nTpoCliente = 1
        End If
        
    End If
    
    Call Reutilizables 'pti1 06-03-2019
    'este será puesto en producción
'    Set poDR = loPigContrato.dVerificarCredPignoTipoCliente(loPers.sPerscod)
'    If Not (poDR.BOF And poDR.EOF) Then
'        nTpoCliente = poDR!nTipoCliente
'    Else
'        nTpoCliente = 1
'    End If
    
    '*** FIN PEAC
    
    Set loPers = Nothing
End If
'*************END RECO*********************************************
'cboPlazo.Enabled = False 'True RECO20140208 ERS002
Exit Sub
End If 'add pti1

'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'TORE ERS054-2017
'Private Sub MostrarObservacionesRetasacion(ByVal pscPersCod As String, _
'                                            ByVal psCtaCod As String, _
'                                            ByVal pnTpoProceso As Integer)
'frmColPObservacionesRetasacion.Observaciones pscPersCod, psCtaCod, pnTpoProceso
'End Sub
'MACM 20210719
Private Function calificacionSBS(ByVal cPersCod As String, ByVal bPersoneria As Boolean, ByVal nPersoneria As String, Optional nCodDeudorRCC As String)
    Dim oDCOMCreditos As DCOMCreditos
    Set oDCOMCreditos = New DCOMCreditos
    Dim rsCalificacionSBS As ADODB.Recordset
    Set rsCalificacionSBS = New ADODB.Recordset
    Dim loColPFunc As COMDColocPig.DCOMColPFunciones
fr_TasaEspecial.Enabled = True
'***Agregado por ELRO el 20120103, según Acta N° 002-2012/TI-D
    Set rsCalificacionSBS = oDCOMCreditos.DatosPosicionClienteCalificacionSBS(bPersoneria, nPersoneria, nCodDeudorRCC)
    If Not rsCalificacionSBS.BOF And Not rsCalificacionSBS.EOF Then
        lblTituloCalificacion = "Última Calificación Según SBS - RCC " & Format(rsCalificacionSBS!Fec_Rep, "dd/mm/yyyy")
        fbMalCalificacion = False
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        fbClienteCPP = False
        '****************************END RECO*************************
        lblCalificacionNormal = "Normal " & rsCalificacionSBS!nNormal & " %"
        lblCalificacionPotencial = "Potencial " & rsCalificacionSBS!nPotencial & " %"
        lblCalificacionDeficiente = "Deficiente " & rsCalificacionSBS!nDeficiente & " %"
        lblCalificacionDudoso = "Dudoso " & rsCalificacionSBS!nDudoso & " %"
        lblCalificacionPerdida = "Perdida " & rsCalificacionSBS!nPerdido & " %"
        
        nCalificacionPotencialCPP = rsCalificacionSBS!nPotencial 'JOEP20180412 Pig Pase
        Call ActivaCampanaPrendario(cPersCod) 'JOEP20210913 Campaña Prendario
        '*** PEAC 20170727
        'RECO20140527 RFC1405270001***********************************************************
        'If CDbl(rsCalificacionSBS!nDudoso) <> 0 Or CDbl(rsCalificacionSBS!nPerdido) <> 0 Or CDbl(rsCalificacionSBS!nDeficiente) Then
        'If CDbl(rsCalificacionSBS!nDudoso) <> 0 Or CDbl(rsCalificacionSBS!nPerdido) <> 0 Then
        'RECO20140527 FIN*********************************************************************
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        'MsgBox "Clientes con calificación Deficiente,Dudoso y Perdida no pueden ser atendidos ", vbCritical, "Aviso"
        'Call cmdCancelar_Click
        'Exit Sub
        '*******************END RECO**********************************
        'fbMalCalificacion = True
        'End If

        If fVerificaSiPasaCalificacion(CDbl(rsCalificacionSBS!nNormal), CDbl(rsCalificacionSBS!nPotencial), CDbl(rsCalificacionSBS!nDeficiente), CDbl(rsCalificacionSBS!nDudoso), CDbl(rsCalificacionSBS!nPerdido)) = False Then
            MsgBox "Clientes con calificación " & lsTextoCalif & " no pueden ser atendidos.", vbCritical, "Aviso"
            Call cmdCancelar_Click
            Exit Function
        End If
        '*** FIN PEAC
        
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        If CDbl(rsCalificacionSBS!nPotencial) <> 0 Or CDbl(rsCalificacionSBS!nDeficiente) <> 0 Then
            fbClienteCPP = True = True
        End If
        '*************END RECO*****************************************
    Else
        lblTituloCalificacion = "Última Calificación Según SBS - RCC "
        fbMalCalificacion = False
        lblCalificacionNormal = "No Registrado"
        lblCalificacionPotencial = "No Registrado"
        lblCalificacionDeficiente = "No Registrado"
        lblCalificacionDudoso = "No Registrado"
        lblCalificacionPerdida = "No Registrado"
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        fbClienteCPP = False
        '****************************END RECO*************************
        nCalificacionPotencialCPP = 0 'JOEP20180412 Pig Pase
    End If
    '***Fin Agregado por ELRO*************************************
    
    Call BuscaInfoClientePig(cPersCod) 'PEAC 20211021
    
'   PEAC 20211021 - Se envio al procedimeinto "BuscaInfoClientePig"
'    'APRI20190515 SATI
'    Dim rsVerInfor As ADODB.Recordset
'    Set rsVerInfor = New ADODB.Recordset
'    Set loColPFunc = New COMDColocPig.DCOMColPFunciones
'    Set rsVerInfor = loColPFunc.dObtieneInfoClientePig(cPersCod)
'    If Not rsVerInfor.BOF And Not rsVerInfor.EOF Then
'        gnDeudaSBS = rsVerInfor!nDeudaSBS
'        gnMontoMax = rsVerInfor!nMontoMax
'        gnSaldAcum = rsVerInfor!nCapAcumulado
'        lblDeudaSBS.Caption = Format(gnDeudaSBS, "###,##0.00")
'        lblMontoMax.Caption = Format(gnMontoMax, "###,##0.00")
'        lblSaldCapA.Caption = Format(gnSaldAcum, "###,##0.00")
'    End If
'    Set rsVerInfor = Nothing
'    Set loColPFunc = Nothing
'    'END APRI

End Function

'***PEAC 20211021
Private Sub BuscaInfoClientePig(ByVal pscPersCod As String)

    Dim rsVerInfor As ADODB.Recordset
    Set rsVerInfor = New ADODB.Recordset
    Set loColPFunc = New COMDColocPig.DCOMColPFunciones
    Set rsVerInfor = loColPFunc.dObtieneInfoClientePig(pscPersCod)
    If Not rsVerInfor.BOF And Not rsVerInfor.EOF Then
        gnDeudaSBS = rsVerInfor!nDeudaSBS
        gnMontoMax = rsVerInfor!nMontoMax
        gnSaldAcum = rsVerInfor!nCapAcumulado
        lblDeudaSBS.Caption = Format(gnDeudaSBS, "###,##0.00")
        lblMontoMax.Caption = Format(gnMontoMax, "###,##0.00")
        lblSaldCapA.Caption = Format(gnSaldAcum, "###,##0.00")
    End If
    Set rsVerInfor = Nothing
    Set loColPFunc = Nothing
   
End Sub


'***PEAC 20080115
Private Sub BuscaCreditosVencidos(pcCodcli, pdFecSis)

Dim oPol As COMDCredito.DCOMCredDoc
Dim rs As ADODB.Recordset
Set oPol = New COMDCredito.DCOMCredDoc

Set rs = oPol.RecuperaCredPigVigVen(pcCodcli, pdFecSis)
 
If rs.RecordCount > 0 Then
    Call frmCredListaVigAtraso.Inicio(rs)
End If
   
End Sub

'Controla el desplazamiento dentro del ListView, indica numero de fila
Private Sub lstCliente_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lsIteSel = Item
End Sub

'Carga los Parametros
Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDColocPig.DCOMColPCalculos

    fnTasaInteresAdelantado = loParam.dObtieneTasaInteres(fsColocLineaCredPig, "1")
    
    'pTasaInteresVencido = (ReadParametros("10105") + 1) ^ 12 - 1
    
    fnTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
    fnTasaTasacion = loParam.dObtieneColocParametro(gConsColPTasaTasacion)
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)

    fnRangoPreferencial = loParam.dObtieneColocParametro(3019)
    'fnPorcentajePrestamo = loParam.dObtieneColocParametro(gConsColPPorcentajePrestamo) 'COMENTADO POR ARLO20171204
    fnImpresionesContrato = loParam.dObtieneColocParametro(gConsColPNroImpresionesContrato)
    fnMaxMontoPrestamo1 = loParam.dObtieneColocParametro(gConsColPLim1MontoPrestamo)
    
Set loParam = Nothing
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    fnJoyasDet = loConstSis.LeeConstSistema(109)
    fsPlazoSist = loConstSis.LeeConstSistema(503) 'RECO20150421
Set loConstSis = Nothing
End Sub

Private Function ValidaDatosGrabar() As Boolean
Dim lbOk As Boolean
Dim i As Integer 'JOEP20180509
lbOk = True
If lstCliente.ListItems.count <= 0 Then
    MsgBox "Falta ingresar el cliente" & vbCr & _
    " Cancele operación ", , " Aviso "
    lbOk = False
    Exit Function
End If
'If Len(Trim(fgEliminaEnters(txtDescLote.Text))) = 0 Then
'    MsgBox " No se ha llenado la descripción de la pieza ", vbInformation, " Aviso "
'    txtDescLote.Enabled = True
'    txtDescLote.SetFocus
'    lbOk = False
'    Exit Function
'End If
' Valida que OroBruto >= OroNeto
If val(lblOroBruto) < val(lblOroNeto) Then
    MsgBox " Oro Neto debe ser menor o igual a Oro Bruto ", vbInformation, " Aviso "
    lbOk = False
    Exit Function
End If
' Monto de Prestamo < 60% de Valor de Tasacion

If val(txtMontoPrestamo.Text) > val(Format(fnPorcentajePrestamo * val(lblValorTasacion.Caption), "#0.00")) Then
'    MsgBox " Monto de Prestamo debe ser menor al 60 % del Valor de Tasacion ", vbInformation, " Aviso "
    MsgBox " Monto de Préstamo debe ser Menor o igual al " & (fnPorcentajePrestamo * 100) & "% (" & Format(fnPorcentajePrestamo * val(lblValorTasacion.Caption), "#,#0.00") & ") del Valor de Tasacion ", vbInformation, " Aviso " '*** PEAC 20170126
    txtMontoPrestamo.SetFocus
    lbOk = False
    Exit Function
End If
If Trim(txtMontoPrestamo.Text) = "" Then
    MsgBox " Falta ingresar Monto de Prestamo " & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
    txtMontoPrestamo.SetFocus
    lbOk = False
    Exit Function
End If
' llena las tipos de Kilatajes
'JOEP20180509
    For i = 1 To FEJoyas.rows - 1
        If FEJoyas.TextMatrix(i, 6) = "" Then
            MsgBox " No ha ingresado el Detalle de las Joyas" & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
            lbOk = False
            Exit Function
        End If
    Next i
'JOEP20180509


    If FEJoyas.rows < 1 Then
        MsgBox " No ha ingresado el Detalle de las Joyas" & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
        cmdAgregar.SetFocus
        lbOk = False
        Exit Function
    Else
        FEJoyas.row = 0
        txt14k.Text = 0: txt16k.Text = 0: txt18k.Text = 0: txt21k.Text = 0
        Do While FEJoyas.row < FEJoyas.rows - 1
            Select Case val(Right(FEJoyas.TextMatrix(FEJoyas.row + 1, 2), 3))
                Case 14
                    txt14k.Text = val(txt14k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
                Case 16
                    txt16k.Text = val(txt16k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
                Case 18
                    txt18k.Text = val(txt18k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
                Case 21
                    txt21k.Text = val(txt21k.Text) + val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
            End Select
            FEJoyas.row = FEJoyas.row + 1
        Loop
    End If
If Trim(txtBuscarLinea.Text) = "" Then
    MsgBox " No existe ninguna Línea de Crédito que se adecúe al crédito pignoraticio " & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
    lbOk = False
    Exit Function
End If

'**ARLO20171204
Dim nContador As Double
Dim nTotal As Double
Me.FEJoyas.row = 0
Do While Me.FEJoyas.row < Me.FEJoyas.rows - 1
    nContador = val(FEJoyas.TextMatrix(FEJoyas.row + 1, 4))
    nTotal = nTotal + nContador
    Me.FEJoyas.row = Me.FEJoyas.row + 1
Loop
If (nTotal < 2) Then
    MsgBox "El peso neto deber ser mayor o igual a 2 gramos", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If
'***************

ValidaDatosGrabar = lbOk
End Function

'TODOCOMPLETA VERIFICAR DLL PIG  *******************************************************
'***************************************************************************************
Private Sub SumaColumnas()
Dim i As Integer
'Dim loPigCalculos As NPigCalculos
Dim lnPiezasT As Integer, lnPBrutoT As Double, lnPNetoT As Double, lnTasacT As Double
'***Agregado por ELRO el 20120719, según OYP-RFC076-2012
Dim lbNroPigAdlCli As Boolean
Dim lsPersCod As String
lsPersCod = lstCliente.ListItems.Item(1).Text
'***Fin Agregado por ELRO el 20120719*******************


    lnPiezasT = 0: lnPBrutoT = 0:       lnPNetoT = 0:       lnTasacT = 0 ':         lnPrestamoT = 0
    'Total Piezas
    lnPiezasT = FEJoyas.SumaRow(1)
    txtPiezas.Text = Format$(lnPiezasT, "##")

    'PESO BRUTO
    lnPBrutoT = FEJoyas.SumaRow(3)
    lblOroBruto.Caption = Format$(lnPBrutoT, "######.00")

    'PESO NETO
    lnPNetoT = FEJoyas.SumaRow(4)
    
    'Tasacion - RIRO 20130702
    lnTasacT = FEJoyas.SumaRow(5)
    
    '***Modificado por ELRO el 20120719, según OYP-RFC076-2012
    'lblOroNeto.Caption = Format$(lnPNetoT, "######.00")
    lbNroPigAdlCli = devolverPignoraticiosAdjudicadosCliente(lsPersCod, gdFecSis)
    If lbNroPigAdlCli = False Then
        lblOroNeto.Caption = Format$(lnPNetoT, "######.00")
        lblNotaPigAdjCli.Visible = False
        lblOroNeto.ForeColor = &H80000008
    Else
        lblOroNeto.Caption = Format$(lnPNetoT - (lnPNetoT * 0.15), "######.00")
        lblNotaPigAdjCli.Visible = True
        lblOroNeto.ForeColor = &HFF&
        lnTasacT = lnTasacT - lnTasacT * 0.15 ' RIRO 20130702
    End If
    '***Fin Modificado por ELRO el 20120719*******************

    
    '***Modificado por ELRO el 20120719, según OYP-RFC076-2012
    'lblValorTasacion.Caption = Format$(lnTasacT, "######.00")
''    If lbNroPigAdlCli = False Then
        lblValorTasacion.Caption = Format$(lnTasacT, "######.00")
''    Else
''        lblValorTasacion.Caption = Format$(lnTasacT - (lnTasacT * 0.15), "######.00")
''    End If
    '***Fin Modificado por ELRO el 20120719*******************

  '****Modificacion MPBR
    cboPlazo.ListIndex = 0
    
    If vSolTasa Then
        fnPorcentajePrestamo = 1
    End If
    If vEstadoSolTasa = 0 Then
        fnPorcentajePrestamo = 1
    End If
    
    vPrestamo = val(lblValorTasacion.Caption) * fnPorcentajePrestamo
    txtMontoPrestamo.Text = Format(vPrestamo, "#0.00")
    txtMontoPrestamo.Enabled = True
    If vSolTasa Then
        txtMontoPrestamo.Enabled = False
    Else
        If vEstadoSolTasa = 0 Then
            txtMontoPrestamo.Enabled = False
        Else
            txtMontoPrestamo.Enabled = True
            If CDbl(txtMontoPrestamo.Text) > 0 Then '** Modificado Por DAOR 20070115
                cmdGrabar.Enabled = True
                cmdImpVolTas.Enabled = True
            End If
        End If
    End If
    
    
    CalculaCostosAsociados
    'txtMontoPrestamo.Text = Format(0, "#0.00")


    'Case Else
       ' lnPNetoT = FEJoyas.SumaRow(3)
       ' lblOroNeto.Caption = Format$(lnPNetoT, "######.00")

       ' lnTasacT = FEJoyas.SumaRow(4)
       ' lblValorTasacion.Caption = Format$(lnTasacT, "######.00")

        'lnPrestamoT = FEJoyas.SumaRow(11)
        'LblPrestamo.Caption = Format$(lnPrestamoT, "######.00")

        'lnPrestamoT = FEJoyas.SumaRow(11)
        'LblPrestamo.Caption = Format$(lnPrestamoT, "######.00")

    'End Select

    'If LblPrestamo <> "" Then
    '    txtPrestamo = CCur(LblPrestamo)
    'End If

End Sub

'------------------------ Cargar Contrato Cancelado ----------------------------
Private Function MuestraCredPig(ByVal psNroContrato As String) As Boolean
Dim lrCredPig As ADODB.Recordset
Dim lrCredPigCostos As ADODB.Recordset
Dim lrCredPigPersonas As ADODB.Recordset
Dim lrCredPigJoyasDet As ADODB.Recordset
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnJoyasDet As Integer
Dim lnPrestamo As Double 'MACM 20212303
Dim lnInteres As Double 'MACM 20212303
Dim contador As Long 'WIOR 20130926
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
    MuestraCredPig = True
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
    
        Set lrCredPig = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psNroContrato)
        Set lrCredPigCostos = loMuestraContrato.dObtieneDatosCreditoPignoraticioCostos(psNroContrato)
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psNroContrato, True)
    Set loMuestraContrato = Nothing
        
    If lrCredPig.BOF And lrCredPig.EOF Then
        lrCredPig.Close
        Set lrCredPig = Nothing
        Set lrCredPigPersonas = Nothing
        MsgBox " No se encuentra el Credito Pignoraticio " & psNroContrato, vbInformation, " Aviso "
        MuestraCredPig = False
        Exit Function
    Else
       
        Me.lblOroBruto.Caption = lrCredPig!nOroBruto
        Me.lblOroNeto.Caption = lrCredPig!nOroNeto
        
        CargaDatosComboTipo Me.cboTipcta, Trim(lrCredPig!cTipCta)
        
        Me.txtPiezas.Text = lrCredPig!npiezas
        Me.lblValorTasacion.Caption = lrCredPig!nTasacion
        Me.txtMontoPrestamo.Text = lrCredPig!nMontoCol
        
        CargaDatosComboPlazo Me.cboPlazo, Trim(lrCredPig!nPlazo)
        
        Me.lblInteres.Caption = lrCredPig!nTasaInteres
        Me.lblFechaVencimiento.Caption = Format(lrCredPig!dvenc, "dd/mm/yyyy")
        
        lrCredPig.Close
        Set lrCredPig = Nothing
        
        'Mostrar Costos e Impuesto
'COMENTÓ APRI20190515 SATI
'        If lrCredPigCostos.EOF And lrCredPigCostos.BOF Then
'             'lrCredPig.Close
'             'Set lrCredPig = Nothing
'             Me.lblCostoCustodia = Format(0, "0.00")
'             Me.lblCostoTasacion = Format(0, "0.00")
'             Me.lblImpuesto = Format(0, "0.00")
'        Else
'            Do Until lrCredPigCostos.EOF
'                If lrCredPigCostos!nPrdConceptoCod = gColPConceptoCodTasacion Then
'                    Me.lblCostoTasacion = Format(lrCredPig!nMonto, "0.00")
'                ElseIf lrCredPigCostos!nPrdConceptoCod = gColPConceptoCodCustodia Then
'                    Me.lblCostoCustodia = Format(lrCredPig!nMonto, "0.00")
'                ElseIf lrCredPigCostos!nPrdConceptoCod = gColPConceptoCodImpuesto Then
'                    Me.lblImpuesto = Format(lrCredPig!nMonto, "0.00")
'                End If
'                lrCredPigCostos.MoveNext
'            Loop
'        End If
        
'        If Trim(Me.lblCostoTasacion.Caption) = "" Then
'            Me.lblCostoTasacion = Format(0, "0.00")
'        ElseIf Trim(Me.lblCostoCustodia.Caption) = "" Then
'            Me.lblCostoCustodia = Format(0, "0.00")
'        ElseIf Trim(Me.lblImpuesto.Caption) = "" Then
'            Me.lblImpuesto = Format(0, "0.00")
'        End If
'END APRI
        '-------------------------------------------------------------------------------------
        
        'vNetoPagar = val(txtMontoPrestamo.Text) - val(lblCostoTasacion) - val(lblCostoCustodia) - val(Me.lblInteres) - val(Me.lblImpuesto)
        vNetoPagar = val(txtMontoPrestamo.Text) - val(Me.lblInteres)  'APRI20190515 SATI
         
        ' Mostrar Clientes
        If fgMostrarClientes(Me.lstCliente, lrCredPigPersonas) = False Then
            'MsgBox " No se encuentra Datos de Clientes de Contrato " & psNroContrato, vbInformation, " Aviso " 'JUEZ 20130717
            MuestraCredPig = False
            Exit Function
        End If
        sCpersTem = Me.lstCliente.ListItems.Item(1) 'add pti1
        Call Reutilizables 'add pti1 07-03-2019
        
        'RECO20140115 INC1401130010*******************************************
        ObtenerTipoCiente (Me.lstCliente.ListItems.Item(1))
        ObtenerCalificacionRCC (AXCodCta.NroCuenta)
        'ObtenerCalificacionRCC
        'RECO END***********************************************
        lrCredPigPersonas.Close
        Set lrCredPigPersonas = Nothing
        
        '*** PEAC 2017-07-27 - si calificacion no cumple entonces cancela
        If lstCliente.ListItems.count <= 0 Then
            Exit Function
        End If
        '*** FIN PEAC
        
        'Para mostrar el Detalle de las joyas
        Set loConstSis = New COMDConstSistema.NCOMConstSistema
            lnJoyasDet = loConstSis.LeeConstSistema(109)
        Set loConstSis = Nothing
        If lnJoyasDet = 1 Then
        
            If MostrarJoyasDet(lrCredPigJoyasDet) = False Then
                MsgBox " No se encuentra Datos de Joyas de Contrato " & psNroContrato, vbInformation, " Aviso "
                MuestraCredPig = False
                Exit Function
            End If
        End If
        
       'WIOR 20130926 **********************
        If FEJoyas.Visible Then
        FEJoyas.SetFocus
            For contador = 1 To FEJoyas.rows - 1
                FEJoyas.row = contador
                FEJoyas.Col = 4
                Call FEJoyas_OnCellChange(contador, 4)
            Next
        End If
        'WIOR FIN ***************************
        If vSolTasa Then
            txtHolograma.Enabled = False
            cboTasador.Enabled = False
            Me.txtMontoPrestamo = Format(lnPrestamo, "0.00")
            Me.lblInteres = Format(lnInteres, "0.00")
            Me.lblNetoRecibir = Format(lnPrestamo, "0.00")
        Else
            If vEstadoSolTasa = 0 Then
                Me.txtHolograma.Enabled = False
                Me.cboTasador.Enabled = False
                Me.txtMontoPrestamo.Enabled = False
            Else
                txtHolograma.Enabled = True
                cboTasador.Enabled = True
            End If
        End If
        'add pti1 07-03-2019
'        txtHolograma.Enabled = True
'        cboTasador.Enabled = True
        'end pti1
            
    End If
Exit Function

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Function

Private Function MostrarJoyasDet(ByVal prJoyas As ADODB.Recordset) As Boolean

    Dim i As Integer

    i = 1
    If prJoyas.BOF And prJoyas.EOF Then
        MsgBox " Error al mostrar datos del cliente ", vbCritical, " Aviso "
        MostrarJoyasDet = False
    Else
        Me.FEJoyas.rsFlex = prJoyas
        MostrarJoyasDet = True
    End If
End Function

Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lrValida As ADODB.Recordset
Dim lbOk As Boolean
Dim lbCan As Boolean
Dim lsmensaje As String

Dim loCredPContrato As New COMNColoCPig.NCOMColPContrato 'RECO20120823 ERS074-2014
'On Error GoTo ControlError

    'Valida Contrato si esta Cancelado
    Set loValContrato = New COMNColoCPig.NCOMColPValida
        lbCan = loValContrato.ValidaCredCancelado(Trim(psNroContrato))
        If Not vSolTasa Then 'MACM 20210323
            If vEstadoSolTasa = 0 Then
            Else
                If lbCan = False Then
                     MsgBox "Su estado de este credito no es Cancelado", vbInformation, "Aviso"
                     Set loValContrato = Nothing
                     Exit Sub
                End If
            End If
        End If
    Set loValContrato = Nothing
    'Muestra Datos
    lbOk = MuestraCredPig(psNroContrato)
    If lbOk = False Then
        AXCodCta.SetFocusCuenta
        Exit Sub
    End If
    'MACM 24032021
    If vSolTasa Then
        If vEstadoSolTasa = 2 Or vEstadoSolTasa = 0 Then
            Me.cmdGrabar.Enabled = True
        Else
            Me.cmdGrabar.Enabled = False 'MACM 23032021
            Me.cmdAgregar.Enabled = False 'MACM 23032021
        End If
    Else
        If vEstadoSolTasa = 0 Then
            Me.cmdGrabar.Enabled = True
            Me.cmdAgregar.Enabled = True 'MACM 23032021
        Else
            Me.cmdGrabar.Enabled = False 'JUEZ 20130717
            Me.cmdAgregar.Enabled = False 'MACM 23032021
        End If
    End If
    'Me.cmdGrabar.Enabled = True 'JUEZ 20130717
    
    Set lrValida = Nothing
    'RECO20120823 ERS074-2014*************************************
    If loCredPContrato.ObtieneHistorialCredRetasacion(AXCodCta.NroCuenta) = True Then
        lblCredRetasado.Visible = True
        cmdVerRetasacion.Visible = True
         If cmdVerRetasacion.Visible Then
            frmColPHistorialRetasacion.Inicio (AXCodCta.NroCuenta)
        End If
    End If
    'RECO FIN ****************************************************
    
    '/** INI PEAC 20211022 **/
    Dim oPig As COMDColocPig.DCOMColPContrato
    Dim RSPERSCOD As Recordset
    Set RSPERSCOD = New ADODB.Recordset
    Set oPig = New COMDColocPig.DCOMColPContrato
    Set RSPERSCOD = oPig.dObtieneDatosCreditoPignoraticio(AXCodCta.NroCuenta)
    '/** FIN PEAC 20211022 **/
    
    Call ObtenerCalificacion(fscPerCod) 'ARLO201831
    
    Call BuscaInfoClientePig(RSPERSCOD!cPersCod) 'PEAC 20211021
    
    'TORE ERS054-2017
    'Call MostrarObservacionesRetasacion("", AXCodCta.NroCuenta, 2) 'TORE 20190614 -> RFC1811260001 :Comentado
    'End TORE
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

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

Public Function ValidarMsh() As Boolean
    Dim nFilas As Integer
    Dim i As Integer
    nFilas = FEJoyas.rows
    For i = 0 To nFilas - 1
        If FEJoyas.TextMatrix(i, 1) = "" Then
            ValidarMsh = True
            MsgBox "Ingrese el detalle de Joyas", vbInformation, "Aviso"
            Exit Function
        End If
    Next
End Function

Public Sub ImprimeHojaResumenPig()
    Dim loPrevio As previo.clsprevio
    Dim lsCadImprimir  As String
    Dim lsCartaModelo As String
  
    lsCadImprimir = ""
    rtfCartas.Filename = App.Path & "\FormatoCarta\HojaResumenPig.txt"
     
    lsCartaModelo = rtfCartas.Text
    lsCartaModelo = Replace(lsCartaModelo, "<<CUENTAS>>", fsContrato, , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<CLIENTE>>", lstCliente.ListItems.Item(1).SubItems(1), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<TASADOR>>", "VICTOR TORRE BLANCA", , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<TASACION>>", CDbl(lblValorTasacion.Caption), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<INTERESCOMPENSATORIO>>", CDbl(lblInteres.Caption), , 2, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<MONTOPRESTAMO>>", CDbl(txtMontoPrestamo.Text), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<MONTONETO>>", CDbl(lblNetoRecibir.Caption), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<TIPOPERIODO>>", cboPlazo.Text, , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<FECHADESEMBOLSO>>", Format(gdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
    lsCadImprimir = lsCadImprimir & lsCartaModelo
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Sobrante de Remate", True
    Set loPrevio = Nothing
    
End Sub

'***Agregado por ELRO el 20120104, según Acta N° 002-2012/TI-D
Private Function verificarCreditosCanceladosCastigados(pcCodcli, pdFecSis) As Boolean

Dim oDCOMCredDoc As COMDCredito.DCOMCredDoc
Dim rsCreditosCanceladosCastigados As ADODB.Recordset
Set oDCOMCredDoc = New COMDCredito.DCOMCredDoc

Set rsCreditosCanceladosCastigados = oDCOMCredDoc.recuperarCreditosCanceladoCastigado(pcCodcli, pdFecSis)

If Not rsCreditosCanceladosCastigados.BOF And Not rsCreditosCanceladosCastigados.EOF Then
    verificarCreditosCanceladosCastigados = True
Else
    verificarCreditosCanceladosCastigados = False
End If

End Function
'***Agregado por ELRO el 20120719, según OYP-RFC076-2012
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
'***Agregado por ELRO el 20120719***********************
'RECO 20140114 INC1401130010***********************************************************
Public Sub ObtenerTipoCiente(ByVal psPersCod As String)
    lnPesoNetoDesc = 0
    Dim loPigContrato As COMDColocPig.DCOMColPContrato
    Set loPigContrato = New COMDColocPig.DCOMColPContrato
    Dim poDR As ADODB.Recordset
    Set poDR = New ADODB.Recordset

    '*** PEAC 20161216
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

'    Set poDR = loPigContrato.dVerificarCredPignoTipoCliente(loPers.sPerscod)
'    If Not (poDR.BOF And poDR.EOF) Then
'        nTpoCliente = poDR!nTipoCliente
'    Else
'        nTpoCliente = 1
'    End If
    '*** FIN PEAC
    
    'Set loPers = Nothing
End Sub

Public Sub ObtenerCalificacionRCC(ByVal psCtaCod As String)
Dim rsCalificacionSBS As ADODB.Recordset
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Set rsCalificacionSBS = New ADODB.Recordset
Set loPersContrato = New COMDColocPig.DCOMColPContrato
Dim lrPersContrato As ADODB.Recordset
Dim oDCOMCreditos As DCOMCreditos
Set oDCOMCreditos = New DCOMCreditos

Set lrPersContrato = loPersContrato.ObtieneDatosPersonaXCredito(psCtaCod)

If Not (lrPersContrato.BOF And lrPersContrato.EOF) Then
    Set rsCalificacionSBS = oDCOMCreditos.DatosPosicionClienteCalificacionSBS(IIf(lrPersContrato!nPersPersoneria = 1, True, False), _
                                                                              IIf(lrPersContrato!nPersPersoneria = 1, Trim(lrPersContrato!NroDNI), Trim(lrPersContrato!NroRuc)), "")
    If Not rsCalificacionSBS.BOF And Not rsCalificacionSBS.EOF Then
        lblTituloCalificacion = "Última Calificación Según SBS - RCC " & Format(rsCalificacionSBS!Fec_Rep, "dd/mm/yyyy")
        fbMalCalificacion = False
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        fbClienteCPP = False
        '****************************END RECO*************************
        lblCalificacionNormal = "Normal " & rsCalificacionSBS!nNormal & " %"
        lblCalificacionPotencial = "Potencial " & rsCalificacionSBS!nPotencial & " %"
        lblCalificacionDeficiente = "Deficiente " & rsCalificacionSBS!nDeficiente & " %"
        lblCalificacionDudoso = "Dudoso " & rsCalificacionSBS!nDudoso & " %"
        lblCalificacionPerdida = "Perdida " & rsCalificacionSBS!nPerdido & " %"
        
        nCalificacionPotencialCPP = rsCalificacionSBS!nPotencial 'JOEP20180412 Pig Pase
        
        If CDbl(rsCalificacionSBS!nDudoso) <> 0 Or CDbl(rsCalificacionSBS!nPerdido) <> 0 Then
            '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
                MsgBox "Clientes con calificación dudoso o perdida no pueden ser atendidos ", vbCritical, "Aviso"
                Call cmdCancelar_Click
                Exit Sub
            '*******************END RECO**********************************
            'fbMalCalificacion = True
        End If
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        If CDbl(rsCalificacionSBS!nPotencial) <> 0 Or CDbl(rsCalificacionSBS!nDeficiente) <> 0 Then
            fbClienteCPP = True = True
        End If
        '*************END RECO*****************************************
    Else
        lblTituloCalificacion = "Última Calificación Según SBS - RCC "
        fbMalCalificacion = False
        lblCalificacionNormal = "No Registrado"
        lblCalificacionPotencial = "No Registrado"
        lblCalificacionDeficiente = "No Registrado"
        lblCalificacionDudoso = "No Registrado"
        lblCalificacionPerdida = "No Registrado"
        
        nCalificacionPotencialCPP = 0 'JOEP20180412 Pig Pase
        
        '******RECO20131213 MEMORANDUM N° 2918-2013-GM-DI/CMAC********
        fbClienteCPP = False
        '****************************END RECO*************************
    End If
End If

End Sub
'END RECO******************************************************************************
'RECO20140220 INC**************************************
Public Sub CalcularValorTasacion(ByVal psFila As Integer, ByVal psColumna As Integer)
Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
Dim lnPOro As Double
'******RECO 20131126*****************
Dim loColContrato As COMDColocPig.DCOMColPContrato
Dim lnValorPOro As Double
Dim lnMatOro As Integer
Dim loDR As ADODB.Recordset


Set loColContrato = New COMDColocPig.DCOMColPContrato
Set loDR = New ADODB.Recordset
'******END RECO**********************
    If FEJoyas.TextMatrix(psFila, 0) = "" Then
        Exit Sub
    End If
    'If (FEJoyas.TextMatrix(psFila, 1) = "" Or FEJoyas.TextMatrix(psFila, 2) = "" Or FEJoyas.TextMatrix(psFila, 3) = "" Or _
    FEJoyas.TextMatrix(psFila, 4) = "" Or FEJoyas.TextMatrix(psFila, 5) = "" Or FEJoyas.TextMatrix(psFila, 6) = "") Then
    'If (FEJoyas.TextMatrix(psFila, 5) = "" Or FEJoyas.TextMatrix(psFila, 6) = "") Then
    '    MsgBox "Los valores no pueden ser vacios", vbCritical, "Aviso"
    '    Exit Sub
    'End If
    If FEJoyas.TextMatrix(psFila, 5) <> "" Then
    If psColumna = 4 Then     'Peso Neto

    'RECO***********
        'lnPesoNetoDesc = CDbl(FEJoyas.TextMatrix(psFila, 3)) * 0.1 'COMENTADO APRI20170623 segun SATI TIC1706220007
        lnPesoNetoDesc = CDbl(FEJoyas.TextMatrix(psFila, 3)) '- lnPesoNetoDesc
    'END RECO*******
        If FEJoyas.TextMatrix(psFila, 4) <> "" Then
            If CCur(FEJoyas.TextMatrix(psFila, 4)) < 0 Then
                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(psFila, 4) = 0
            Else
                If CCur(FEJoyas.TextMatrix(psFila, 4)) > lnPesoNetoDesc Then
                    'MsgBox "Peso Neto " & CCur(FEJoyas.TextMatrix(psFila, 4)) & " debe ser menor a peso neto base " & lnPesoNetoDesc, vbInformation, "Aviso" 'COMENTADO APRI20170623 segun SATI TIC1706220007
                    MsgBox "Peso Neto " & CCur(FEJoyas.TextMatrix(FEJoyas.row, 4)) & " no debe ser mayor al Peso Bruto " & lnPesoNetoDesc, vbInformation, "Aviso"
                    FEJoyas.TextMatrix(psFila, 4) = lnPesoNetoDesc
                Else
                    'CalculaTasacion
                        Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, val(Left(FEJoyas.TextMatrix(psFila, 2), 2)), 1) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2)
                        '********RECO 20131126 ERS158****************
                        lnMatOro = Left(FEJoyas.TextMatrix(psFila, 2), 2) 'APRI 20170408  CAMBIO de Right (X,3) -> Left (X,2)
                           
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
                        '********END RECO****************************
                        If lnPOro <= 0 Then
                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
                            Exit Sub
                        End If
                        Set loColPCalculos = Nothing
                        'Calcula el Valor de Tasacion
                        '********RECO 20131126 ERS158**********
                        FEJoyas.TextMatrix(psFila, 5) = Format$(val(FEJoyas.TextMatrix(psFila, 4) * lnValorPOro), "#####.00")
                        'FEJoyas.TextMatrix(FEJoyas.row, 5) = Format$(val(FEJoyas.TextMatrix(FEJoyas.row, 4) * lnPOro), "#####.00")
                        '********END RECO**********************
                End If
            End If
        End If
        
    End If
    End If
    If psColumna = 6 Then     'Descripcion

        If FEJoyas.TextMatrix(psFila, 6) <> "" Then
            'cboPlazo.Enabled = False 'True RECO20140208 ERS002
        End If
        
    End If
    SumaColumnas
End Sub
'FIN RECO**********************************************
'RECO20150421 *****************************************
Private Sub CargaPlazo()
    Dim i As Integer
    Dim sPlazo As String
    For i = 1 To Len(fsPlazoSist)
        If Mid(fsPlazoSist, i, 1) <> "," Then
            sPlazo = sPlazo & Mid(fsPlazoSist, i, 1)
        Else
            cboPlazo.AddItem (sPlazo)
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
'Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio("705", "0", Trim(Right((txtBuscarLinea.psDescripcion), 15)), sLineaTmp, lblLineaDesc, "1", CCur(IIf((txtMontoPrestamo.Text) = "", 0, txtMontoPrestamo.Text)), 0) 'arlo20200307 comentó
Set RLinea = oLineas.RecuperaLineadeCreditoProductoCrediticio("709", "0", fscPerCod, sLineaTmp, lblLineaDesc, "1", CCur(IIf((txtMontoPrestamo.Text) = "", 0, txtMontoPrestamo.Text)), 0) 'arlo20200307
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
    If val(txtMontoPrestamo.Text) > 0 Then
    Dim oLineas As COMDCredito.DCOMLineaCredito
    Dim lrsLineas As ADODB.Recordset
    Set lrsLineas = New ADODB.Recordset
    
    Set oLineas = New COMDCredito.DCOMLineaCredito
    Set lrsLineas = oLineas.RecuperaLineasProductoArbol("755", "1", , gsCodAge, 30, CDbl(txtMontoPrestamo.Text), 1, , 0, gdFecSis)
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

'*** PEAC 20170727
Private Function fVerificaSiPasaCalificacion(ByVal pnNormal As Double, ByVal pnPotencial As Double, ByVal pnDeficiente As Double, ByVal pnDudoso As Double, ByVal pnPerdido As Double) As Boolean
    fVerificaSiPasaCalificacion = True
    
    If (pnDudoso + pnPerdido + pnDeficiente) > 0 Then
        fVerificaSiPasaCalificacion = False
    End If
    
End Function

'***ARLO20180131 -INICIO
Private Sub ObtenerCalificacion(ByVal fscPerCod As String)

'Inicio **ARLO20171204 ERS082-2017
Dim pSegDR As ADODB.Recordset
Dim psConfDR  As ADODB.Recordset
Dim psVeriDR As ADODB.Recordset
Dim psSegConDR As ADODB.Recordset

Dim loPigContrato As COMDColocPig.DCOMColPContrato
Set loPigContrato = New COMDColocPig.DCOMColPContrato

Set pSegDR = rsRecuperaSegDR.Clone 'add pti1 07-03-2019
Set psConfDR = rsRecuperapsConfDR.Clone 'add pti1 07-03-2019

If Not (pSegDR.BOF And pSegDR.EOF) Then '---INCIO MODIFICADO DESDE LA 60
    nSegmentoAnt = "" 'JOEP pig pase
    fcSegmento = "" 'JOEP20200124 Se agrego Variable segun observacion
    nDiasSegmento = pSegDR!nDias 'Agrego JOEP20180220
    nCantAdjuSegmento = pSegDR!nCantAdjuticados 'Agrego JOEP20180220
    nDiasSinAdj = pSegDR!DiasSinAdj 'JOEP pig pase
    nVerificaVencidos = pSegDR!CredVenc 'JOEP pig pase

'JOEP20190611 Mejora en la evaluacion de segmentacion
    Dim rsResulEvalSeg As ADODB.Recordset
    Set rsResulEvalSeg = loPigContrato.ObtieneResulSeg(fscPerCod, CCur(lblValorTasacion), nCalificacionPotencialCPP, pSegDR!nDias, pSegDR!nCantAdjuticados, pSegDR!DiasSinAdj, nVerificaVencidos, fnTpoClinte, pSegDR!ObtEmpresa, pSegDR!RC18, pSegDR!RC17, pSegDR!RC16, pSegDR!RC15, pSegDR!RC14, pSegDR!RC13, pSegDR!RC12, pSegDR!RC11, pSegDR!RC10, pSegDR!RC09, pSegDR!RC08, pSegDR!RC07)
        
        If Not (rsResulEvalSeg.BOF And rsResulEvalSeg.EOF) Then
            If rsResulEvalSeg!cMensaje = "" Then
                fnPorcentajePrestamo = rsResulEvalSeg!nPorctPres
                lblCalificacion.Caption = rsResulEvalSeg!cVerCalif
                fcSegmento = rsResulEvalSeg!cSegmento 'JOEP20200124 Se agrego Variable segun observacion
                nSegmentoAnt = rsResulEvalSeg!cSegmentoAnt 'JOEP20200124 Se agrego Variable segun observacion
                
                'Seg. Prendario Externo JOEP20210422
                Call SegPrendarioExterno(fscPerCod)
                                
                ReDim vArrayDatosSegPred(17)
                vArrayDatosSegPred(0) = pSegDR!CredVenc
                vArrayDatosSegPred(1) = pSegDR!ObtEmpresa
                vArrayDatosSegPred(2) = Format(nCalificacionPotencialCPP, "#0.00")
                vArrayDatosSegPred(3) = pSegDR!RC18
                vArrayDatosSegPred(4) = pSegDR!RC17
                vArrayDatosSegPred(5) = pSegDR!RC16
                vArrayDatosSegPred(6) = pSegDR!RC15
                vArrayDatosSegPred(7) = pSegDR!RC14
                vArrayDatosSegPred(8) = pSegDR!RC13
                vArrayDatosSegPred(9) = pSegDR!RC12
                vArrayDatosSegPred(10) = pSegDR!RC11
                vArrayDatosSegPred(11) = pSegDR!RC10
                vArrayDatosSegPred(12) = pSegDR!RC09
                vArrayDatosSegPred(13) = pSegDR!RC08
                vArrayDatosSegPred(14) = pSegDR!RC07
                vArrayDatosSegPred(15) = pSegDR!dFechaRCC
                vArrayDatosSegPred(16) = Mid(lblSegPrenExter, 9, 20)
                'Seg. Prendario Externo JOEP20210422
                                
            Else
                MsgBox rsResulEvalSeg!cMensaje, vbInformation, "Aviso"
                fnPorcentajePrestamo = 0
                lblCalificacion.Caption = ""
                'Seg. Prendario Externo JOEP20210422
                lblSegPrenExter.Caption = ""
                Set vArrayDatosSegPred = Nothing
                'Seg. Prendario Externo JOEP20210422
            End If
        End If
    Set loPigContrato = Nothing
    RSClose rsResulEvalSeg
'JOEP20190611 Mejora en la evaluacion de segmentacion

End If
End Sub
'**ARLO20183101

'CROB20180611 begin
Private Sub AsignarTipoFormalidadGarantia(ByVal nMontoPrestamo As Currency)
    Dim nValorUIT_3 As Currency
    Dim nValorUIT_7 As Currency
    Dim nValorPrestamo As Currency
    Dim oDNiv As COMDColocPig.DCOMColPContrato
    
    If nValorUIT = 0 Then
        Set oDNiv = New COMDColocPig.DCOMColPContrato
        nValorUIT = oDNiv.ObtenerUITdelAnio(2001) 'Valor UIT
        If nValorUIT = 0 Then
            MsgBox "No existe UIT vigente", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    nValorUIT_3 = nValorUIT * 3
    nValorUIT_7 = nValorUIT * 7
    nValorPrestamo = nMontoPrestamo
    
    If nValorPrestamo = 0 Then
        lblTipoForGarPigno.Caption = ""
    ElseIf nValorPrestamo <= nValorUIT_3 Then
        lblTipoForGarPigno.Caption = "Garantía Firma simple"
    ElseIf nValorPrestamo > (nValorUIT_3 + 0.1) And nValorPrestamo <= nValorUIT_7 Then
        lblTipoForGarPigno.Caption = "Garantía Firma legalizada"
    ElseIf nValorPrestamo > nValorUIT_7 Then
        lblTipoForGarPigno.Caption = "Garantía Inscripción ante RRPP"
    End If
End Sub
'CROB20180611 end

'Seg. Prendario Externo JOEP20210422
Private Sub SegPrendarioExterno(ByVal pcPersCod As String)
    Dim objSP As COMDCredito.DCOMCreditos
    Dim rsSegPred As ADODB.Recordset
    
    lblSegPrenExter.Caption = ""
    
    Set objSP = New COMDCredito.DCOMCreditos
    Set rsSegPred = objSP.DataSegPrendarioExterno(pcPersCod)
    If Not (rsSegPred.EOF And rsSegPred.BOF) Then
        lblSegPrenExter.Font = "6dp"
        lblSegPrenExter.Caption = "Cliente " & rsSegPred!cSegmento
    End If
    
    Set objSP = Nothing
    RSClose rsSegPred
End Sub
'Seg. Prendario Externo JOEP20210422
'MACM 17032021 INICIO
Private Sub txtTasaEspeci_Change()
    ReDim MatCalend(0, 0)
    ReDim MatrizCal(0, 0)
End Sub

Private Sub txtTasaEspeci_GotFocus()
    fEnfoque txtTasaEspeci
End Sub

Private Sub txtTasaEspeci_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTasaEspeci, KeyAscii, , 4)
     If KeyAscii = 13 Then
     End If
End Sub

Private Sub txtTasaEspeci_LostFocus()
    If Trim(txtTasaEspeci.Text) = "" Then
        txtTasaEspeci.Text = "0.0000"
    Else
        txtTasaEspeci.Text = Format(txtTasaEspeci.Text, "#0.0000")
    End If
End Sub
'MACM 17032021 FIN


'JOEP20210913 Campana Prendario
Private Sub CargaCampPrendario()
    Dim objCampPre As COMDConstantes.DCOMConstantes
    Dim rsCampPre As ADODB.Recordset
    Set objCampPre = New COMDConstantes.DCOMConstantes
   
    Set rsCampPre = objCampPre.RecuperaConstantes(100)
    
    If Not (rsCampPre.BOF And rsCampPre.EOF) Then
        Call Llenar_Combo_con_Recordset(rsCampPre, cboCampPrendario)
        Call CambiaTamañoCombo(cboCampPrendario)
    End If
   
    Set objCampPre = Nothing
    RSClose rsCampPre
End Sub

Private Sub cboCampPrendario_Click()
    Dim oRCampPre As COMDColocPig.DCOMColPContrato
    Dim rsRCampPre As ADODB.Recordset
    Set oRCampPre = New COMDColocPig.DCOMColPContrato
    
    If FEJoyas.TextMatrix(FEJoyas.row, 5) = "" Then
        cboCampPrendario.ListIndex = 0
        Exit Sub
    End If
    
    If cboCampPrendario.Text <> "" Then
        Set rsRCampPre = oRCampPre.CampPrendarioDescCampa(Right(cboCampPrendario.Text, 3))
        If Trim(Right(cboCampPrendario.Text, 3)) <> 0 Then
            If Not (rsRCampPre.BOF And rsRCampPre.EOF) Then
                txtCampPrendario.Text = rsRCampPre!cDescripcion
                fr_TasaEspecial.Enabled = False
                
                If Right(cboCampPrendario.Text, 2) = 2 Then
                    fsPlazoSist = "60,"
                    cboPlazo.Clear
                    CargaPlazo
                    cboPlazo.ListIndex = 0
                    Call CalculaCostosAsociados
                Else
                    Call CargaParametros
                    cboPlazo.Clear
                    CargaPlazo
                    cboPlazo.ListIndex = 0
                    Call CalculaCostosAsociados
                End If
            Else
                fr_TasaEspecial.Enabled = True
                txtCampPrendario.Text = ""
            End If
        Else
            Call CargaParametros
            cboPlazo.Clear
            CargaPlazo
            cboPlazo.ListIndex = 0
            Call CalculaCostosAsociados
                    
            fr_TasaEspecial.Enabled = True
            txtCampPrendario.Text = ""
        End If
    End If
    Set oRCampPre = Nothing
    RSClose rsRCampPre
End Sub

Private Sub ActivaCampanaPrendario(ByVal pcPersCod As String)
    Dim obCamPr As COMDColocPig.DCOMColPContrato
    Dim rsCamPr As ADODB.Recordset
    Set obCamPr = New COMDColocPig.DCOMColPContrato
        
    cboCampPrendario.ListIndex = -1
    txtCampPrendario.Text = ""
    Set rsCamPr = obCamPr.CampPrendarioActivaCampana(pcPersCod)
    If rsCamPr!nPase = 1 Then
        fr_CampPrendario.Enabled = True
        fr_CampPrendario.Visible = True
        
        CmdPrevio.top = 9120
        cmdImpVolTas.top = 9120
        cmdGrabar.top = 9120
        cmdCancelar.top = 9120
        cmdSalir.top = 9120
                
        frmColPRegContratoDet.Height = 9975

    Else
        fr_CampPrendario.Enabled = False
        fr_CampPrendario.Visible = False
        
        CmdPrevio.top = 8280
        cmdImpVolTas.top = 8280
        cmdGrabar.top = 8280
        cmdCancelar.top = 8280
        cmdSalir.top = 8280
        frmColPRegContratoDet.Height = 9150

    End If
        
    Set obCamPr = Nothing
    RSClose rsCamPr
End Sub

'JOEP20210913 Campana Prendario
