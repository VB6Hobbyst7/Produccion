VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColPRegContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Registrar Contrato"
   ClientHeight    =   5985
   ClientLeft      =   990
   ClientTop       =   1260
   ClientWidth     =   7275
   Icon            =   "frmColPRegContrato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   35
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6075
      TabIndex        =   14
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3525
      TabIndex        =   13
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   5370
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   7110
      Begin VB.Frame fraPiezasDet 
         Caption         =   "Detalle de Piezas"
         Height          =   1875
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   6795
         Begin VB.CommandButton cmdEliminarJ 
            Caption         =   "Eliminar"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5730
            TabIndex        =   59
            Top             =   1500
            Width           =   765
         End
         Begin VB.CommandButton CmdAgregarJ 
            Caption         =   "A&gregar"
            Height          =   315
            Left            =   4920
            TabIndex        =   58
            Top             =   1500
            Width           =   795
         End
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   1215
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   6570
            _ExtentX        =   11589
            _ExtentY        =   2143
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   2
            EncabezadosNombres=   "Num-Material-PBruto-PNeto-Tasacion-Descripcion-p-Item"
            EncabezadosAnchos=   "350-1030-650-650-900-2000-0-0"
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-X-5-X-X"
            ListaControles  =   "0-3-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-C-L-R-C"
            FormatosEdit    =   "0-1-2-2-4-0-4-3"
            TextArray0      =   "Num"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   570
         Index           =   7
         Left            =   120
         TabIndex        =   50
         Top             =   2640
         Width           =   5415
         Begin VB.Label lblEtiqueta 
            Caption         =   "Kilataje (gr)"
            Height          =   195
            Index           =   22
            Left            =   3120
            TabIndex        =   54
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   23
            Left            =   300
            TabIndex        =   53
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   4080
            TabIndex        =   52
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   51
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
         TabIndex        =   38
         Top             =   2640
         Width           =   5370
         Begin VB.CommandButton cmdLineaCredito 
            Caption         =   "..."
            Height          =   285
            Left            =   4860
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   180
            Width           =   4695
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Cliente(s)"
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
         Height          =   1425
         Index           =   6
         Left            =   135
         TabIndex        =   36
         Top             =   120
         Width           =   6810
         Begin VB.ComboBox cboTipcta 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPRegContrato.frx":030A
            Left            =   4800
            List            =   "frmColPRegContrato.frx":0317
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   990
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
            Width           =   825
         End
         Begin MSComctlLib.ListView lstCliente 
            Height          =   765
            Left            =   90
            TabIndex        =   39
            Top             =   180
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   1349
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
               Size            =   8.25
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
            Left            =   3480
            TabIndex        =   37
            Top             =   1080
            Width           =   1245
         End
      End
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   915
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   4320
         Width           =   6795
         Begin VB.Label lblNetoRecibir 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5580
            TabIndex        =   47
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label lblImpuesto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3420
            TabIndex        =   46
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label lblInteres 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1260
            TabIndex        =   45
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label lblFechaVencimiento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   5580
            TabIndex        =   44
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblCostoTasacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1260
            TabIndex        =   43
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblCostoCustodia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3420
            TabIndex        =   42
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Neto Recibir"
            Height          =   255
            Index           =   20
            Left            =   4590
            TabIndex        =   34
            Top             =   540
            Width           =   945
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Tasac."
            Height          =   255
            Index           =   19
            Left            =   180
            TabIndex        =   33
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Custod."
            Height          =   255
            Index           =   18
            Left            =   2430
            TabIndex        =   32
            Top             =   180
            Width           =   990
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Interes"
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   31
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Impuesto "
            Height          =   255
            Index           =   17
            Left            =   2430
            TabIndex        =   30
            Top             =   540
            Width           =   855
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fec.Vencim."
            Height          =   255
            Index           =   15
            Left            =   4590
            TabIndex        =   29
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Descripcion Lote"
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
         Height          =   1095
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   3060
         Width           =   6795
         Begin VB.TextBox txtDescLote 
            Enabled         =   0   'False
            Height          =   765
            Left            =   120
            MaxLength       =   254
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   210
            Width           =   6555
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   1065
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Top             =   1560
         Width           =   5415
         Begin VB.CommandButton cmdPiezasDet 
            Caption         =   "..."
            Height          =   255
            Left            =   2520
            TabIndex        =   56
            Top             =   120
            Width           =   255
         End
         Begin VB.ComboBox cboPlazo 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmColPRegContrato.frx":0352
            Left            =   4140
            List            =   "frmColPRegContrato.frx":035F
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   1125
         End
         Begin VB.TextBox txtOroBruto 
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
            Left            =   1440
            MaxLength       =   7
            TabIndex        =   3
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtPiezas 
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
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   9
            Top             =   420
            Width           =   1095
         End
         Begin VB.TextBox txtMontoPrestamo 
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
            Left            =   4140
            MaxLength       =   11
            TabIndex        =   11
            Top             =   735
            Width           =   1125
         End
         Begin VB.Label lblValorTasacion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   49
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblOroNeto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4140
            TabIndex        =   48
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Oro Bruto  (gr)"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Oro Neto  (gr)"
            Height          =   210
            Index           =   10
            Left            =   2880
            TabIndex        =   25
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Piezas"
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Plazo  (dias)"
            Height          =   255
            Index           =   8
            Left            =   2880
            TabIndex        =   23
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Valor Tasación "
            Height          =   255
            Index           =   3
            Left            =   255
            TabIndex        =   22
            Top             =   735
            Width           =   1215
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Monto Prestamo"
            Height          =   255
            Index           =   9
            Left            =   2850
            TabIndex        =   21
            Top             =   735
            Width           =   1335
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Kilataje"
         Height          =   1500
         Index           =   5
         Left            =   5580
         TabIndex        =   15
         Top             =   1560
         Width           =   1350
         Begin VB.TextBox txt21k 
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
            Left            =   525
            MaxLength       =   5
            TabIndex        =   8
            Top             =   1140
            Width           =   720
         End
         Begin VB.TextBox txt18k 
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
            Left            =   525
            MaxLength       =   5
            TabIndex        =   7
            Top             =   840
            Width           =   720
         End
         Begin VB.TextBox txt16k 
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
            Left            =   540
            MaxLength       =   5
            TabIndex        =   6
            Top             =   540
            Width           =   720
         End
         Begin VB.TextBox txt14k 
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
            Left            =   540
            MaxLength       =   5
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "21 K"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   19
            Top             =   1140
            Width           =   465
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "18 K"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "16 K"
            Height          =   255
            Index           =   12
            Left            =   135
            TabIndex        =   17
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "14 K"
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   16
            Top             =   255
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmColPRegContrato"
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

Dim fnTasaInteresAdelantado As Double
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
Dim lstTmpCliente As ListItem
Dim lsIteSel As String
Dim vContAnte As Boolean

'Procedimiento para cargar los valores a los campos txt
Private Sub CalculaCostosAsociados()
Dim loCostos As COMNColoCPig.NCOMColPCalculos

Set loCostos = New COMNColoCPig.NCOMColPCalculos

    vPrestamo = Val(txtMontoPrestamo.Text)
    'vPlazo = Val(cboPlazo.Text)
    'Cálculo valores
    'vCostoTasacion = Val(lblValorTasacion.Caption) * fnTasaTasacion
    vCostoTasacion = loCostos.nCalculaCostoTasacion(Val(lblValorTasacion.Caption), fnTasaTasacion)
    vCostoCustodia = loCostos.nCalculaCostoCustodia(Val(lblValorTasacion), fnTasaCustodia, Val(cboPlazo.Text))
    vInteres = loCostos.nCalculaInteresAdelantado(Val(txtMontoPrestamo), fnTasaInteresAdelantado, Val(cboPlazo.Text))
    'vImpuesto = (vCostoTasacion + vInteres + vCostoCustodia) * pTasaImpuesto
    vImpuesto = loCostos.nCalculaImpuestoDesembolso(vCostoTasacion, vInteres, vCostoCustodia, fnTasaImpuesto)
    
    vNetoPagar = Val(txtMontoPrestamo.Text) - vCostoTasacion - vCostoCustodia - vInteres - vImpuesto

Set loCostos = Nothing

'Muestra los Resultados
Me.lblFechaVencimiento = Format(DateAdd("d", Val(cboPlazo.Text), gdFecSis), "dd/mm/yyyy")
Me.lblCostoTasacion = Format(vCostoTasacion, "#0.00")
Me.lblCostoCustodia = Format(vCostoCustodia, "#0.00")
Me.lblInteres = Format(vInteres, "#0.00")
Me.lblImpuesto = Format(vImpuesto, "#0.00")
Me.lblNetoRecibir = Format(vNetoPagar, "#0.00")

End Sub

'Función para calcular el valor de tasación
' Calcula en base al precio del oro en el mercado
Private Function ValorTasacion() As Double
If Val(txt14k.Text) >= 0 And Val(txt16k.Text) >= 0 And Val(txt18k.Text) >= 0 And Val(txt21k.Text) >= 0 Then
   ValorTasacion = (Val(txt14k.Text) * fnPrecioOro14) + (Val(txt16k.Text) * fnPrecioOro16) + (Val(txt18k.Text) * fnPrecioOro18) + (Val(txt21k.Text) * fnPrecioOro21)
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
    txtOroBruto.Text = Format(0, "#0.00")
    LblOroNeto.Caption = Format(0, "#0.00")
    txtPiezas.Text = Format(0, "#0")
    cboPlazo.ListIndex = 0
    lblValorTasacion.Caption = Format(0, "#0.00")
    txtMontoPrestamo.Text = Format(0, "#0.00")
    txtDescLote.Text = ""
    Me.lblFechaVencimiento = ""
    Me.lblCostoTasacion = Format(0, "#0.00")
    Me.lblCostoCustodia = Format(0, "#0.00")
    Me.lblImpuesto = Format(0, "#0.00")
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
End Sub

'Función que calcula el total de kilatajes
Private Function SumaKilataje() As Double
If Val(txt14k.Text) >= 0 And Val(txt16k.Text) >= 0 And Val(txt18k.Text) >= 0 And Val(txt21k.Text) >= 0 Then
   SumaKilataje = Val(txt14k.Text) + Val(txt16k.Text) + Val(txt18k.Text) + Val(txt21k.Text)
Else
   MsgBox " No se ha ingresado correctamente el Kilataje ", vbInformation, " Aviso "
End If
End Function

Private Sub cboTipCta_Click()
If lstCliente.ListItems.Count = 1 Then
    cboTipcta.ListIndex = 0
    txtOroBruto.Enabled = True
ElseIf lstCliente.ListItems.Count >= 2 And cboTipcta.ListIndex = 0 Then
    cboTipcta.ListIndex = 1
    txtOroBruto.Enabled = True
End If
End Sub

Private Sub cboTipcta_KeyPress(KeyAscii As Integer)
txtOroBruto.SetFocus
End Sub

Private Sub CmdAgregarJ_Click()
    If FEJoyas.Rows <= 20 Then
        If lnJoyas > 1 And FEJoyas.TextMatrix(FEJoyas.Row, 5) = "" Then
            MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
            Exit Sub
        Else
            If lnJoyas = 1 And FEJoyas.TextMatrix(FEJoyas.Row, 5) = "" Then
                MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
                Exit Sub
            Else
                lnJoyas = lnJoyas + 1
                FEJoyas.AdicionaFila
                If FEJoyas.Rows >= 2 Then
                   cmdEliminarJ.Enabled = True
                End If
            End If
        End If
    Else
        CmdAgregarJ.Enabled = False
        MsgBox "Sólo puede ingresar como máximo veinte piezas", vbInformation, "Aviso"
    End If
End Sub

'Permite cancelar el proceso actual
Private Sub cmdCancelar_Click()
    Limpiar
    txtOroBruto.Enabled = False
    txt14k.Enabled = False
    txt16k.Enabled = False
    txt18k.Enabled = False
    txt21k.Enabled = False
    txtPiezas.Enabled = False
    cboPlazo.Enabled = False
    lblValorTasacion.Enabled = False
    txtMontoPrestamo.Enabled = False
    txtDescLote.Enabled = False
    cboPlazo.ListIndex = 0
    cboTipcta.Enabled = False
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = False
    'Me.AXCodCta.Visible = False
    'cmdContAnterior.Enabled = True
End Sub

Private Sub cmdEliminarJ_Click()
    FEJoyas.EliminaFila FEJoyas.Row
    If FEJoyas.Rows <= 20 Then
        CmdAgregarJ.Enabled = True
    End If
    lnJoyas = lnJoyas + 1
    SumaColumnas
    cboPlazo_Click
    
End Sub

'Permite actualizar los datos en la base de datos
Private Sub cmdGrabar_Click()

Dim pbTran As Boolean
Dim lsCtaReprestamo As String

Dim lrPersonas As ADODB.Recordset
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
Dim loRegImp As COMNColoCPig.NCOMColPImpre
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsContrato As String
Dim loPrevio As Previo.clsPrevio

Dim lsmensaje As String
Dim lsCadImprimir As String

'On Error GoTo ControlError
pbTran = False

If ValidaDatosGrabar = False Then Exit Sub

'Asigno los valores a los parametros

Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.lstCliente)
txtDescLote = fgEliminaEnters(txtDescLote) & vbCr

lnMontoPrestamo = CCur(txtMontoPrestamo.Text)
lnPlazo = Val(cboPlazo.Text)
lsFechaVenc = Format$(Me.lblFechaVencimiento, "mm/dd/yyyy")
lnOroBruto = Val(txtOroBruto.Text)
lnOroNeto = Val(LblOroNeto.Caption)
lnPiezas = Val(txtPiezas.Text)
lnValTasacion = CCur(lblValorTasacion.Caption)
lsTipoContrato = Switch(cboTipcta.ListIndex = 0, "I", cboTipcta.ListIndex = 1, "O", cboTipcta.ListIndex = 2, "Y")
lsLote = txtDescLote.Text
ln14k = Val(txt14k.Text)
ln16k = Val(txt16k.Text)
ln18k = Val(txt18k.Text)
ln21k = Val(txt21k.Text)
lnIntAdelantado = CCur(Me.lblInteres.Caption)
lnCostoTasac = CCur(Me.lblCostoTasacion.Caption)
lnCostoCustodia = CCur(Me.lblCostoCustodia.Caption)
lnImpuesto = CCur(Me.lblImpuesto.Caption)
          
If MsgBox("¿Grabar Contrato Prestamo Pignoraticio ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    'Genera Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    
    Set loRegPig = New COMNColoCPig.NCOMColPContrato
        lsContrato = loRegPig.nRegistraContratoPignoraticio(gsCodCMAC & gsCodAge, gMonedaNacional, _
            lrPersonas, fnTasaInteresAdelantado, lnMontoPrestamo, lsFechaHoraGrab, lnPlazo, _
            lsFechaVenc, lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lsTipoContrato, _
            lsLote, ln14k, ln16k, ln18k, ln21k, lsMovNro, lnIntAdelantado, lnCostoTasac, lnCostoCustodia, lnImpuesto)
        pbTran = False
    Set loRegPig = Nothing
    MsgBox "Se ha generado Contrato Nro " & lsContrato, vbInformation, "Aviso"

    If MsgBox("Imprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loRegImp = New COMNColoCPig.NCOMColPImpre
            lsCadImprimir = loRegImp.nPrintContratoPignoraticio(lsContrato, False, lrPersonas, fnTasaInteresAdelantado, _
                lnMontoPrestamo, lsFechaHoraGrab, Format(lsFechaVenc, "mm/dd/yyyy"), lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, _
                lnPiezas, lsLote, ln14k, ln16k, ln18k, ln21k, lnIntAdelantado, lnCostoTasac, lnCostoCustodia, lnImpuesto, gsCodUser, lsmensaje)
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation, "Aviso"
                    Exit Sub
                End If
        Set loRegImp = Nothing
        Set loPrevio = New Previo.clsPrevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, False
        
            Do While True
                If MsgBox("Reimprimir Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False
                Else
                    Set loPrevio = Nothing
                    Set loRegPig = Nothing
                    Exit Do
                End If
            Loop
    End If
End If
Limpiar

Exit Sub

ControlError:   ' Rutina de control de errores.
    'Verificar que se halla iniciado transaccion y la cierra
    'If pbTran Then dbCmact.RollbackTrans
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    Limpiar
End Sub

Private Sub cmdLineaCredito_Click()
    'frmColPLineaCreditoSelecciona.Show 1
End Sub

'Finaliza la ejecusión del formulario
Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub FEJoyas_Click()
Dim loConst As COMDConstantes.DCOMConstantes
Dim lrMaterial As ADODB.Recordset
Set loConst = New COMDConstantes.DCOMConstantes

Select Case FEJoyas.Col
Case 1
    Set lrMaterial = loConst.RecuperaConstantes(gColocPigMaterial, , "C.cConsDescripcion")
    FEJoyas.CargaCombo lrMaterial
    Set lrMaterial = Nothing

End Select

Set loConst = Nothing

End Sub

Private Sub FEJoyas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FEJoyas.Col = 1 Then
        If FEJoyas.TextMatrix(FEJoyas.Row, 9) <> "" Then
            CmdAgregarJ.SetFocus
            CmdAgregarJ_Click
        End If
    End If
End If
End Sub

'Inicializa el formulario
Private Sub Form_Load()
    CargaParametros
    Limpiar
    lblNetoRecibir.ForeColor = pColPriEgreso
    lblNetoRecibir.BackColor = pColFonSoles
End Sub

'Valida el campo txt14k
Private Sub txt14k_GotFocus()
fEnfoque txt14k
End Sub

Private Sub txt14k_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txt14k, KeyAscii, 6, 2)
If KeyAscii = 13 Then
    txt14k.Text = IIf(Val(txt14k.Text) = 0, Format(0, "#0.00"), Format(txt14k.Text, "#0.00"))
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Then
        MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt14k)
        txt14k.Text = Format(0, "#0.00")
        fEnfoque txt14k
    Else
        txt16k.Enabled = True
        txt16k.SetFocus
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End If
End Sub

Private Sub txt14k_Validate(Cancel As Boolean)
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Then
        'MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt14k)
        txt14k.Text = Format(0, "#0.00")
        'CalculaValorTasacion ' Calcula valor Tasacion
        lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
        txt14k.SetFocus
        Cancel = True
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End Sub

'Valida el campo txt16k
Private Sub txt16k_GotFocus()
fEnfoque txt16k
End Sub

Private Sub txt16k_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txt16k, KeyAscii, 6, 2)
If KeyAscii = 13 Then
    txt16k.Text = IIf(Val(txt16k.Text) = 0, Format(0, "#0.00"), Format(txt16k.Text, "#0.00"))
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Then
        MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt16k)
        txt16k.Text = Format(0, "#0.00")
        fEnfoque txt16k
    Else
        txt18k.Enabled = True
        txt18k.SetFocus
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End If
End Sub

Private Sub txt16k_Validate(Cancel As Boolean)
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Then
        'MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt16k)
        txt16k.Text = Format(0, "#0.00")
        'CalculaValorTasacion ' Calcula valor Tasacion
        lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
        Cancel = True
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End Sub

'Valida el campo txt18k
Private Sub txt18k_GotFocus()
fEnfoque txt18k
End Sub

Private Sub txt18k_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txt18k, KeyAscii, 6, 2)
If KeyAscii = 13 Then
    txt18k.Text = IIf(Val(txt18k.Text) = 0, Format(0, "#0.00"), Format(txt18k.Text, "#0.00"))
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Then
        MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt18k)
        txt18k.Text = Format(0, "#0.00")
        fEnfoque txt18k
    Else
        txt21k.Enabled = True
        txt21k.SetFocus
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End If
End Sub

Private Sub txt18k_Validate(Cancel As Boolean)
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Then
        'MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt18k)
        txt18k.Text = Format(0, "#0.00")
        'CalculaValorTasacion ' Calcula valor Tasacion
        lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
        Cancel = True
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End Sub

'Valida el campo txt21k
Private Sub txt21k_GotFocus()
fEnfoque txt21k
End Sub

Private Sub txt21k_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txt21k, KeyAscii, 6, 2)
If KeyAscii = 13 Then
    txt21k.Text = IIf(Val(txt21k) = 0, Format(0, "#0.00"), Format(txt21k.Text, "#0.00"))
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If (Val(txtOroBruto) < Val(LblOroNeto)) Or (Val(LblOroNeto) = 0) Then
        MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt21k)
        txt21k.Text = Format(0, "#0.00")
        txt14k.SetFocus
    ElseIf Val(LblOroNeto) < fnMinPesoOro Then
        MsgBox "Peso de Oro Neto debe ser mayor a Peso mínimo de oro", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt21k)
        txt21k.Text = Format(0, "#0.00")
        txt14k.SetFocus
    Else
        txtPiezas.Enabled = True
        txtPiezas.SetFocus
    End If
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End If
End Sub

Private Sub txt21k_Validate(Cancel As Boolean)
    LblOroNeto.Caption = Format(SumaKilataje, "#0.00")
    If Val(txtOroBruto.Text) < Val(LblOroNeto.Caption) Or Val(LblOroNeto) = 0 Then
        'MsgBox " El peso de Oro Bruto debe ser Mayor " & vbCr _
        & " que el peso del Oro Neto ", vbInformation, " Aviso "
        LblOroNeto = Val(LblOroNeto) - Val(txt21k)
        txt21k.Text = Format(0, "#0.00")
        'CalculaValorTasacion ' Calcula valor Tasacion
        lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
        If lblValorTasacion.Enabled = True Then
            Cancel = True
        End If
    End If
    'CalculaValorTasacion ' Calcula valor Tasacion
    lblValorTasacion.Caption = Format(ValorTasacion, "#0.00")
End Sub

'Valida el campo txtdesclote
Private Sub txtDescLote_GotFocus()
If Len(Trim(fgEliminaEnters(txtDescLote))) > 0 And Val(txtOroBruto) > 0 And _
Val(lblValorTasacion) > 0 And Val(txtMontoPrestamo) > 0 And Val(lblOroPrestamo) > 0 Then
    cmdGrabar.Enabled = True
End If
End Sub
Private Sub txtDescLote_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
KeyAscii = fgIntfLineas(txtDescLote, KeyAscii, 14)
If Len(Trim(fgEliminaEnters(txtDescLote))) > 0 And Val(txtOroBruto) > 0 And _
Val(lblValorTasacion) > 0 And Val(txtMontoPrestamo) > 0 And Val(lblOroPrestamo) > 0 Then
    cmdGrabar.Enabled = True
End If
End Sub

'Valida el campo txtmontoprestamo
Private Sub txtMontoPrestamo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoPrestamo, KeyAscii, 10, 2)
If KeyAscii = 13 Then
   CalculaCostosAsociados
   ' Calcula Porcentaje de Prestamo en Oro
   If Val(LblOroNeto) = 0 Then
        MsgBox " Oro Neto no debe ser CERO ", vbInformation, " Aviso "
        txt14k.SetFocus
   ElseIf Val(txtMontoPrestamo) > Val(Format(fnPorcentajePrestamo * Val(lblValorTasacion.Caption), "#0.00")) Then
        MsgBox " Monto de Préstamo debe ser Menor al 60 % del Valor de Tasacion ", vbInformation, " Aviso "
        txtMontoPrestamo.SetFocus
   ElseIf Val(txtMontoPrestamo) = 0 Or Len(Trim(txtMontoPrestamo)) = "" Then
        MsgBox " Ingrese Monto prestado ", vbInformation, " Aviso "
        txtMontoPrestamo.SetFocus
   ElseIf Val(txtMontoPrestamo) > fnMaxMontoPrestamo1 Then
        If MsgBox("Monto de Préstamo es Mayor al permitido, Avise al Administrador " & vbCr & _
        " Desea continuar con el Préstamo ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
            lblOroPrestamo.Caption = Format(Val(LblOroNeto.Caption) * Val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * Val(lblValorTasacion.Caption)), "#0.00")
            lblOroPrestamoPorcen.Caption = Format((Val(lblOroPrestamo) * 100) / Val(LblOroNeto), "#0")
            If vContAnte = True Then
                cmdGrabar.Enabled = True
                cmdGrabar.SetFocus
            Else
                txtDescLote.Enabled = True
                txtDescLote.SetFocus
            End If
        Else
            txtMontoPrestamo.SetFocus
        End If
   Else
        lblOroPrestamo.Caption = Format(Val(LblOroNeto.Caption) * Val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * Val(lblValorTasacion.Caption)), "#0.00")
        lblOroPrestamoPorcen.Caption = Format((Val(lblOroPrestamo) * 100) / Val(LblOroNeto), "#0")
        If vContAnte = True Then
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        Else
            txtDescLote.Enabled = True
            txtDescLote.SetFocus
        End If
   End If
End If
End Sub
Private Sub txtMontoPrestamo_Validate(Cancel As Boolean)
  CalculaCostosAsociados
   ' Calcula Porcentaje de Prestamo en Oro
   If Val(LblOroNeto) = 0 Then
        MsgBox " Oro Neto no debe ser CERO ", vbInformation, " Aviso "
        Cancel = True
   ElseIf Val(txtMontoPrestamo) > Val(Format(fnPorcentajePrestamo * Val(lblValorTasacion.Caption), "#0.00")) Then
        MsgBox " Monto de Prestamo debe ser Menor al 60 % del Valor de Tasacion ", vbInformation, " Aviso "
        Cancel = True
   ElseIf Val(LblOroNeto) = 0 Then
        MsgBox " Ingrese cantidad de ORO ", vbInformation, " Aviso "
        Cancel = True
   ElseIf Val(txtMontoPrestamo) = 0 Or Len(Trim(txtMontoPrestamo)) = "" Then
        MsgBox " Ingrese Monto prestado ", vbInformation, " Aviso "
        Cancel = True
   ElseIf Val(txtMontoPrestamo) > fnMaxMontoPrestamo1 Then
        If MsgBox("Monto de Préstamo es Mayor al permitido, Avise al Administrador " & vbCr & _
        " Desea continuar con el Préstamo ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
            lblOroPrestamo.Caption = Format(Val(LblOroNeto.Caption) * Val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * Val(lblValorTasacion.Caption)), "#0.00")
            lblOroPrestamoPorcen.Caption = Format((Val(lblOroPrestamo) * 100) / Val(LblOroNeto), "#0")
        Else
            Cancel = True
        End If
   Else
        lblOroPrestamo.Caption = Format(Val(LblOroNeto.Caption) * Val(txtMontoPrestamo.Text) / (fnPorcentajePrestamo * Val(lblValorTasacion.Caption)), "#0.00")
        lblOroPrestamoPorcen.Caption = Format((Val(lblOroPrestamo) * 100) / Val(LblOroNeto), "#0")
   End If
End Sub

'Valida el campo txtorobruto
Private Sub txtOroBruto_GotFocus()
fEnfoque txtOroBruto
End Sub
Private Sub txtOroBruto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtOroBruto, KeyAscii, 7, 2)
If KeyAscii = 13 And Val(txtOroBruto.Text) > 0 And Val(txtOroBruto) >= fnMinPesoOro Then
    txtOroBruto.Text = Format(txtOroBruto.Text, "##0.00")
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    txt14k.Enabled = True
    txt14k.SetFocus
End If
End Sub

Private Sub txtOroBruto_LostFocus()
If Val(LblOroNeto) <> 0 Then
    If Len(Trim(txtOroBruto)) = 0 Then
        MsgBox " Oro bruto no debe ser Cero ", vbInformation, " Aviso "
        txtOroBruto.SetFocus
    ElseIf Val(txtOroBruto) < fnMinPesoOro Then
        MsgBox " Oro bruto debe ser Mayor al peso mínimo de oro ", vbInformation, " Aviso "
        txtOroBruto.SetFocus
    End If
    If Val(LblOroNeto) > Val(txtOroBruto) Then
        LblOroNeto.Caption = Format(0, "#0.00")
        txt14k.Text = Format(0, "#0.00")
        txt16k.Text = Format(0, "#0.00")
        txt18k.Text = Format(0, "#0.00")
        txt21k.Text = Format(0, "#0.00")
        lblValorTasacion.Caption = Format(0, "#0.00")
        lblOroPrestamoPorcen.Caption = ""
        lblOroPrestamo.Caption = ""
    End If
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
            cboPlazo.Enabled = True
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
If Val(txtMontoPrestamo.Text) <> 0 Then
    CalculaCostosAsociados
End If
End Sub

Private Sub cboplazo_KeyPress(KeyAscii As Integer)
    'ValorTasacion = Val(lblValorTasacion.Caption)
    vPrestamo = Val(lblValorTasacion.Caption) * fnPorcentajePrestamo
    txtMontoPrestamo.Text = Format(vPrestamo, "#0.00")
    txtMontoPrestamo.Enabled = True
    txtMontoPrestamo.SetFocus
End Sub

' Valida el campo txtvalortasación
Private Sub lblValorTasacion_Change()
If Val(txtMontoPrestamo) <> 0 Then
    cmdGrabar.Enabled = False
    txtMontoPrestamo = Format(0, "#0.00")
End If
End Sub

Private Sub cmdeliminar_Click()
On Error GoTo ControlError
    Dim i As Integer, j As Integer
    If lstCliente.ListItems.Count = 0 Then
       MsgBox "No existen datos, imposible eliminar", vbInformation, "Aviso"
       cmdEliminar.Enabled = False
       cboTipcta.Enabled = False
       lstCliente.SetFocus
       Exit Sub
    Else
       For i = 1 To lstCliente.ListItems.Count
           If lstCliente.ListItems.Item(i) = lsIteSel Then
              lstCliente.ListItems.Remove (i)
              Exit For
           End If
       Next i
    End If
    lstCliente.SetFocus
    If lstCliente.ListItems.Count = 0 Then
        txtOroBruto = Format(0, "#0.00")
        txtOroBruto.Enabled = False
        cboTipcta.Enabled = False
        cmdEliminar.Enabled = False
    ElseIf lstCliente.ListItems.Count = 1 Then
        cboTipcta.ListIndex = 0
        txtOroBruto.Enabled = True
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite buscar un cliente por nombre y/o documento
Private Sub CmdAgregar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim ls As String
Dim loColPFunc As COMDColocPig.DCOMColPFunciones
On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    '** Verifica que no este en lista
    For liFil = 1 To Me.lstCliente.ListItems.Count
        If lsPersCod = Me.lstCliente.ListItems.Item(liFil).Text Then
           MsgBox " Cliente Duplicado ", vbInformation, "Aviso"
           Exit Sub
        End If
    Next liFil
    '** Maximo Nro de Clientes = 4
    If Me.lstCliente.ListItems.Count > 4 Then
           MsgBox " Maximo Nro de Clientes ==> 4 ", vbInformation, "Aviso"
           Exit Sub
    End If
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
        
    cboTipcta.Enabled = True
End If

Set loPers = Nothing

If lstCliente.ListItems.Count >= 1 And cboTipcta.ListIndex = 0 Then
    cboTipcta.ListIndex = 1
    txtOroBruto.Enabled = True
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
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

    fnTasaInteresAdelantado = loParam.dObtieneTasaInteres("01011130501", "01")
    'pTasaInteresVencido = (ReadParametros("10105") + 1) ^ 12 - 1
    
    fnTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
    fnTasaTasacion = loParam.dObtieneColocParametro(gConsColPTasaTasacion)
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    
    fnPrecioOro14 = loParam.dObtieneColocParametro(gConsColPPrecioOro14)
    fnPrecioOro16 = loParam.dObtieneColocParametro(gConsColPPrecioOro16)
    fnPrecioOro18 = loParam.dObtieneColocParametro(gConsColPPrecioOro18)
    fnPrecioOro21 = loParam.dObtieneColocParametro(gConsColPPrecioOro21)
    fnPorcentajePrestamo = loParam.dObtieneColocParametro(gConsColPPorcentajePrestamo)
    fnImpresionesContrato = loParam.dObtieneColocParametro(gConsColPNroImpresionesContrato)
    fnMaxMontoPrestamo1 = loParam.dObtieneColocParametro(gConsColPLim1MontoPrestamo)
    
Set loParam = Nothing
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    fnJoyasDet = loConstSis.LeeConstSistema(109)
Set loConstSis = Nothing
End Sub

Private Function ValidaDatosGrabar() As Boolean
Dim lbOk As Boolean
lbOk = True
If lstCliente.ListItems.Count <= 0 Then
    MsgBox "Falta ingresar el cliente" & vbCr & _
    " Cancele operación ", , " Aviso "
    lbOk = False
    Exit Function
End If
If Len(Trim(fgEliminaEnters(txtDescLote.Text))) = 0 Then
    MsgBox " No se ha llenado la descripción de la pieza ", vbInformation, " Aviso "
    txtDescLote.Enabled = True
    txtDescLote.SetFocus
    lbOk = False
    Exit Function
End If
' Valida que OroBruto >= OroNeto
If Val(txtOroBruto) < Val(LblOroNeto) Then
    MsgBox " Oro Neto debe ser menor o igual a Oro Bruto ", vbInformation, " Aviso "
    txtOroBruto.Enabled = True
    txtOroBruto.SetFocus
    lbOk = False
    Exit Function
End If
' Monto de Prestamo < 60% de Valor de Tasacion
If Val(txtMontoPrestamo.Text) > Val(Format(fnPorcentajePrestamo * Val(lblValorTasacion.Caption), "#0.00")) Then
    MsgBox " Monto de Prestamo debe ser menor al 60 % del Valor de Tasacion ", vbInformation, " Aviso "
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
ValidaDatosGrabar = lbOk
End Function

'Private Function HabilitaControles(ByVal CmdGrabar As Boolean, ByVal txtOroBruto As Boolean, ByVal txtOro14 As Boolean, ByVal txtOro16 As Boolean, _
'        ByVal txtoro18 As Boolean, ByVal txtOro21 As Boolean, ByVal txtPiezas As Boolean, ByVal cboPlazo As Boolean)
'
'    CmdGrabar.Enabled = False
'    txtOroBruto.Enabled = False
'    txt14k.Enabled = False
'    txt16k.Enabled = False
'    txt18k.Enabled = False
'    txt21k.Enabled = False
'    txtPiezas.Enabled = False
'    cboPlazo.Enabled = False
'    lblValorTasacion.Enabled = False
'    txtMontoPrestamo.Enabled = False
'    txtDescLote.Enabled = False
'    cboTipcta.Enabled = False
'    CmdAgregar.Enabled = True
'    'cmdContAnterior.Enabled = True
'    txtOroBruto.Enabled = False
'    txt14k.Enabled = False
'    txt16k.Enabled = False
'    txt18k.Enabled = False
'    txt21k.Enabled = False
'    txtPiezas.Enabled = False
'    cboPlazo.Enabled = False
'    lblValorTasacion.Enabled = False
'    txtMontoPrestamo.Enabled = False
'    txtDescLote.Enabled = False
'    cboPlazo.ListIndex = 0
'    cboTipcta.Enabled = False
'    CmdAgregar.Enabled = True
'    CmdEliminar.Enabled = False
'End Function

'TODOCOMPLETA ***********************************************************************
'************************************************************************************
Private Sub SumaColumnas()
'Dim i As Integer
'Dim loPigCalculos As NPigCalculos
'Dim lnPBrutoT As Double, lnPNetoT As Double, lnTasacT As Double
'    lnPBrutoT = 0:      lnPNetoT = 0:       lnTasacT = 0 ':         lnPrestamoT = 0
'    Select Case FEJoyas.Col
'
'    Case 3      'PESO BRUTO
'        lnPBrutoT = FEJoyas.SumaRow(3)
'        txtOroBruto.Text = Format$(lnPBrutoT, "######.00")
'        'TxtTotalB.Text =txtorobrutPBruto.Caption
'
'    Case 4      'PESO NETO
'        lnPNetoT = FEJoyas.SumaRow(4)
'        lblOroNeto.Caption = Format$(lnPNetoT, "######.00")
'
'        lnTasacT = FEJoyas.SumaRow(5)
'        lblValorTasacion.Caption = Format$(lnTasacT, "######.00")
'
'        'lnPrestamoT = FEJoyas.SumaRow(11)
'        'LblPrestamo.Caption = Format$(lnPrestamoT, "######.00")
'
'    Case Else
'        lnPNetoT = FEJoyas.SumaRow(3)
'        lblOroNeto.Caption = Format$(lnPNetoT, "######.00")
'
'        lnTasacT = FEJoyas.SumaRow(4)
'        lblValorTasacion.Caption = Format$(lnTasacT, "######.00")
'
'        'lnPrestamoT = FEJoyas.SumaRow(11)
'        'LblPrestamo.Caption = Format$(lnPrestamoT, "######.00")
'
'        'lnPrestamoT = FEJoyas.SumaRow(11)
'        'LblPrestamo.Caption = Format$(lnPrestamoT, "######.00")
'
'    End Select
'
'    'If LblPrestamo <> "" Then
'    '    txtPrestamo = CCur(LblPrestamo)
'    'End If
    
End Sub
