VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogIngAlmacen 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   Icon            =   "frmLogIngAlmacen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
     Begin VB.Frame FraNIngresoLibre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   120
      TabIndex        =   43
      Top             =   5220
      Width           =   2175
      Begin VB.CheckBox ChkNIngresoLibre 
         Height          =   210
         Left            =   1800
         TabIndex        =   44
         Top             =   155
         Width           =   255
      End
      Begin VB.Label lblNIngreso 
         Caption         =   "Nota de Ingreso Libre"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdEliminarDoc 
      Caption         =   "E&liminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11610
      TabIndex        =   38
      Top             =   2535
      Width           =   1020
   End
   Begin VB.CommandButton cmdAgregarDoc 
      Caption         =   "Ag&regar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10515
      TabIndex        =   37
      Top             =   2535
      Width           =   1020
   End
   Begin Sicmact.TxtBuscar txtAlmacen 
      Height          =   330
      Left            =   1365
      TabIndex        =   17
      Top             =   1095
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
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
   Begin Sicmact.FlexEdit FlexDoc 
      Height          =   990
      Left            =   7275
      TabIndex        =   0
      Top             =   1515
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   1746
      Cols0           =   6
      HighLight       =   1
      RowSizingMode   =   1
      EncabezadosNombres=   "#-OK-Documento-Fecha-Serie-Numero"
      EncabezadosAnchos=   "290-400-1800-800-400-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-2-3-4-5"
      TextStyleFixed  =   3
      ListaControles  =   "0-4-3-2-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-L"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   0
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   285
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraProveedor 
      Appearance      =   0  'Flat
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   15
      TabIndex        =   8
      Top             =   1425
      Width           =   7215
      Begin Sicmact.TxtBuscar txtProveedor 
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
      Begin VB.Label lblProveedorNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   105
         TabIndex        =   10
         Top             =   660
         Width           =   6795
      End
   End
   Begin VB.Frame framCont 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   15
      TabIndex        =   11
      Top             =   2820
      Width           =   12630
      Begin Sicmact.FlexEdit FlexSerie 
         Height          =   2235
         Left            =   8700
         TabIndex        =   26
         Top             =   180
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   3942
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Serie-id-idx-IGV-Valor-Val-Serie Fisica"
         EncabezadosAnchos=   "300-1200-0-0-700-900-0-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-7"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-R-R-C-L"
         FormatosEdit    =   "0-0-0-0-2-2-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1155
         TabIndex        =   16
         Top             =   2835
         Width           =   1020
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   15
         Top             =   2835
         Width           =   1020
      End
      Begin Sicmact.FlexEdit FlexDetalle 
         Height          =   2235
         Left            =   60
         TabIndex        =   14
         Top             =   180
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   3942
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Descripción-Cantidad-Precio Unit-IGV-Total-CtaCnt"
         EncabezadosAnchos=   "300-1200-3000-800-900-800-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name = "Tahoma"
Size = 6.75
Charset = 0
Weight = 400
Underline = 0   'False
Italic = 0   'False
Strikethrough = 0   'False
EndProperty
ColumnasAEditar = "X-1-X-3-4-5-X-X"
TextStyleFixed = 3
ListaControles = "0-1-0-0-0-0-0-0"
BackColorControl = -2147483643
BackColorControl = -2147483643
BackColorControl = -2147483643
EncabezadosAlineacion = "C-L-L-R-R-R-R-C"
FormatosEdit = "0-0-0-3-2-2-2-0"
TextArray0 = "#"
lbPuntero = -1  'True
lbBuscaDuplicadoText = -1  'True
Appearance = 0
ColWidth0 = 300
RowHeight0 = 300
ForeColorFixed = -2147483630
End
      Begin VB.Frame fraComentario 
         Appearance      =   0  'Flat
         Caption         =   "Comentario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   765
         Left            =   2265
         TabIndex        =   12
         Top             =   2775
         Width           =   10275
         Begin VB.TextBox txtComentario 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   75
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   225
            Width           =   10125
         End
      End
      Begin VB.Label lblTotalIGV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6195
         TabIndex        =   39
         Top             =   2445
         Width           =   975
      End
      Begin VB.Label lblTotalG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7155
         TabIndex        =   36
         Top             =   2445
         Width           =   1200
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4905
         TabIndex        =   35
         Top             =   2445
         Width           =   3435
      End
      Begin VB.Label lblTotalGDet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   10950
         TabIndex        =   41
         Top             =   2460
         Width           =   1005
      End
      Begin VB.Label lblTotalIGVDet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   10110
         TabIndex        =   40
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label lblTotDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Detalle"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8865
         TabIndex        =   42
         Top             =   2460
         Width           =   3090
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11625
      TabIndex        =   7
      Top             =   6480
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1245
      TabIndex        =   6
      Top             =   6480
      Width           =   1020
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2340
      TabIndex        =   5
      Top             =   6480
      Width           =   1020
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   1020
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   285
      Left            =   11355
      TabIndex        =   2
      Top             =   45
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Sicmact.TxtBuscar txtOCompra 
      Height          =   330
      Left            =   6330
      TabIndex        =   20
      Top             =   1095
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   582
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
   Begin VB.Frame fraMoneda 
      Appearance      =   0  'Flat
      Caption         =   "Moneda"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   8805
      TabIndex        =   23
      Top             =   -45
      Width           =   1095
      Begin VB.OptionButton optDolares 
         Appearance      =   0  'Flat
         Caption         =   "&Dolares"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   495
         Width           =   945
      End
      Begin VB.OptionButton optSoles 
         Appearance      =   0  'Flat
         Caption         =   "&Soles"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin Sicmact.TxtBuscar txtNotaIng 
      Height          =   330
      Left            =   1365
      TabIndex        =   27
      Top             =   735
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   582
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
   Begin VB.Frame fraCambio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   9945
      TabIndex        =   30
      Top             =   270
      Visible         =   0   'False
      Width           =   2670
      Begin VB.Label lblCompraG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         TabIndex        =   34
         Top             =   135
         Width           =   600
      End
      Begin VB.Label lblFijoG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         TabIndex        =   33
         Top             =   135
         Width           =   600
      End
      Begin VB.Label llCompra 
         Caption         =   "Compra"
         Height          =   225
         Left            =   1230
         TabIndex        =   32
         Top             =   135
         Width           =   555
      End
      Begin VB.Label lblFijo 
         Caption         =   "Fijo"
         Height          =   225
         Left            =   105
         TabIndex        =   31
         Top             =   135
         Width           =   435
      End
   End
   Begin VB.Label lblNotaIng 
      Caption         =   "Nota de Ingreso :"
      Height          =   210
      Left            =   150
      TabIndex        =   29
      Top             =   795
      Width           =   1230
   End
   Begin VB.Label lblNotaIngG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4020
      TabIndex        =   28
      Top             =   750
      Width           =   8595
   End
   Begin VB.Label lblOrdenCompra 
      Caption         =   "O/C :"
      Height          =   210
      Left            =   5880
      TabIndex        =   22
      Top             =   1155
      Width           =   1125
   End
   Begin VB.Label lblOCompra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8460
      TabIndex        =   21
      Top             =   1110
      Width           =   4170
   End
   Begin VB.Label lblAlmacenG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2565
      TabIndex        =   18
      Top             =   1110
      Width           =   3225
   End
   Begin VB.Label lblAlmacen 
      Caption         =   "Almacen :"
      Height          =   210
      Left            =   150
      TabIndex        =   19
      Top             =   1155
      Width           =   1125
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10740
      TabIndex        =   3
      Top             =   75
      Width           =   660
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota de Ingreso : 2001-00001"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   8700
   End
End
Attribute VB_Name = "frmLogIngAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsOpeCod As String
Dim lnMovNroG As Long
Dim lnMovNroOPG As Long
Dim lbIngreso As Boolean
Dim lbConfirma As Boolean
Dim lsCaptionG As String
Dim lbReporte As Boolean
Dim lbMantenimiento As Boolean
Dim lbExtorno As Boolean
Dim lbGrabar As Boolean
Dim lnPorProIni As Currency

Dim rs As ADODB.Recordset 'anps
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************
'CHK NOTA INGRESO LIBRE - ANPS
Private Sub ChkNIngresoLibre_Click()
    If ChkNIngresoLibre.value = 1 Then 'ACTIVA CHECK
        Me.FlexDetalle.lbEditarFlex = True
        FlexDetalle.AdicionaFila
        Me.txtOCompra.Enabled = False
        Me.cmdAgregar.Visible = True
        Me.cmdEliminar.Visible = True
    Else
        Me.FlexDetalle.lbEditarFlex = False
        LimpiaFlex Me.FlexDetalle
        Me.txtOCompra.Enabled = True
        Me.cmdAgregar.Visible = False
        Me.cmdEliminar.Visible = False
    End If
End Sub
'FIN CHK NOTA INGRESO LIBRE - ANPS

Private Sub cmdAgregar_Click()
    If ChkNIngresoLibre.value <> 1 Then
        If Me.txtProveedor.Text = "" Then
            MsgBox "Debe elegir una  persona responsable.", vbInformation, "Aviso"
        Me.txtProveedor.SetFocus
            Exit Sub
        End If
    End If

    If Me.FlexDetalle.TextMatrix(1, 1) <> "" Then
        If FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) <> "" Or FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3) <> "" Then
            Me.FlexDetalle.AdicionaFila , , True
        End If
    Else
        Me.FlexDetalle.AdicionaFila
    End If
    Me.FlexDetalle.SetFocus
End Sub

Private Sub cmdAgregarDoc_Click()
    Me.FlexDoc.AdicionaFila
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdEliminar_Click()
    Dim I As Integer
    Dim lnEncontrar As Integer
    Dim lnContador As Integer
    
    If MsgBox("Desea Eliminar la fila, si ha incluido numeros de serie para este producto se perderan.", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

    For I = 1 To CInt(Me.FlexSerie.Rows - 1)
        If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
            lnContador = lnContador + 1
        End If
    Next I
    
    I = 0
    While lnEncontrar < lnContador
        I = I + 1
        If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
            Me.FlexSerie.EliminaFila I
            lnEncontrar = lnEncontrar + 1
            I = I - 1
        End If
    Wend
    
    For I = 1 To Me.FlexSerie.Rows - 1
        If IsNumeric(FlexSerie.TextMatrix(I, 3)) Then
            If FlexSerie.TextMatrix(I, 3) > Me.FlexDetalle.row Then
                FlexSerie.TextMatrix(I, 3) = Trim(Str(CInt(FlexSerie.TextMatrix(I, 3)) - 1))
                FlexSerie.TextMatrix(I, 0) = FlexSerie.TextMatrix(I, 3)
                FlexSerie.TextMatrix(I, 2) = FlexSerie.TextMatrix(I, 3)
            End If
        End If
    Next I

    Me.FlexDetalle.EliminaFila Me.FlexDetalle.row
            'anps
    Dim lnI As Integer
    Dim lnTotal As Currency
    Dim lnTotalIGV As Currency
    lnTotal = 0
    lnTotalIGV = 0
    For lnI = 1 To Me.FlexDetalle.Rows - 1
        If IsNumeric(FlexDetalle.TextMatrix(lnI, 5)) And IsNumeric(FlexDetalle.TextMatrix(lnI, 6)) Then
            lnTotal = lnTotal + CCur(FlexDetalle.TextMatrix(lnI, 6))
            lnTotalIGV = lnTotalIGV + CCur(FlexDetalle.TextMatrix(lnI, 5))
        End If
    Next lnI

    Me.lblTotalG.Caption = Format(lnTotal, "#,##0.00")
    Me.lblTotalIGV.Caption = Format(lnTotalIGV, "#,##0.00")
    'anps fin
End Sub

Private Sub cmdEliminarDoc_Click()
    Me.FlexDoc.EliminaFila Me.FlexDoc.row
End Sub

Private Sub cmdGrabar_Click()
    If Not Valida Then Exit Sub
    Dim I As Integer
    Dim oAlmacen As DMov
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnItem As Integer
    Dim lsBSCod As String
    Dim lsDocNI As String
    Dim ldFechaOC As Date
    Set oAlmacen = New DMov
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    Dim lsOpeCodLocal As String
    Dim lsCtaCont As String
    Dim lnContador As Long
    Dim lsAgePagare As String
    Dim lsAlmacenTmp As String
    Dim lsTipoMovAlm As String
    Dim lnStock As Currency
    Dim lnCosProm As Currency
    Dim lsEsNuevo As Boolean
    Dim lsUniMed As String
    Dim lsDocumento As String
    Dim lsMoneda As String
    Dim lnCostoTotal As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
            
  'JEOM  Validación para Ingreso por Orden Compra/Servicio x Almacén
  '-----------------------------------------------------------------
   If gsopecod = 591101 Or gsopecod = 591102 Or gsopecod = 591103 Then
        If Len(Trim(txtAlmacen.Text)) = 1 Then
           lsAlmacenTmp = "0" & txtAlmacen.Text
        Else
           lsAlmacenTmp = Trim(txtAlmacen.Text)
        End If
        If gsCodAge <> lsAlmacenTmp Then
           MsgBox "Código de Almacén diferente Código de Agencia Usuario", vbInformation, "Aviso"
           Exit Sub
        End If
    End If
   '-------------------------------------------------------------------
   'FIN JEOM
    
    'GITU
    lsTipoMovAlm = "E"  'para extorno
    
    If Left(gsopecod, 4) = "5911" Then 'Ingresos
        lsTipoMovAlm = "I"
    ElseIf Left(gsopecod, 4) = "5912" Then
        lsTipoMovAlm = "S"
    End If
    
    If Me.optSoles.value = True Then
        lsMoneda = "1"
    End If
    
    If Me.optDolares.value = True Then
        lsMoneda = "2"
    End If
    
    If MsgBox("Desea Grabar los cambios Realizados ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lbGrabar = True
    
    If lbExtorno Then
        
        oAlmacen.BeginTrans
          'Inserta Mov
          lsMovNro = oAlmacen.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
          oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabRechazado, gMovFlagEliminado
          
          lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
          oAlmacen.InsertaMovGasto lnMovNro, Me.txtProveedor.Text, ""
          
          If lnMovNroG <> 0 Then  'Para Modificados
             oAlmacen.ActualizaMov lnMovNroG, , gMovEstContabNoContable, gMovFlagModificado   'Modificado
             oAlmacen.ActualizaMov oAlmacen.GetnMovNroRef(lnMovNroG, gnAlmaIngXCompras), , , gMovFlagEliminado  'Cambia el modificado a vigente
             oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
             oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
             
             'Gitu Extorno
             
             Set rs = oOpe.GetPrimerIngreso(Me.txtAlmacen.Text)
             Set rsA = oOpe.GetMaestroAlmacen(Me.txtAlmacen.Text)
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsBSCod = Me.FlexDetalle.TextMatrix(I, 1)
                rs.Filter = "cBSCod = '" & lsBSCod & "'"
              
                If rs.RecordCount > 1 Then
                    rsA.Filter = "cBSCod = '" & lsBSCod & "'"
                    lnStock = rsA!nStock - Me.FlexDetalle.TextMatrix(I, 3)
                    lnCostoTotal = rsA!nCostoTotal - Me.FlexDetalle.TextMatrix(I, 6)
                    '****Modificado por ALPA
                    '****13/03/2008
                    If lnStock <> 0 Then
                    lnCosProm = lnCostoTotal / lnStock
                    Else
                    lnCosProm = 0
                    End If
                    '*****Fin de modificación
                    oAlmacen.ActualizaMasterAlmacen Me.FlexDetalle.TextMatrix(I, 1), lnStock, Me.FlexDetalle.TextMatrix(I, 4), lnCosProm, lnCostoTotal, txtAlmacen.Text, gdFecSis, lsTipoMovAlm, lsMovNro
                    oAlmacen.EliminaKardexAlmacen lsBSCod, lnMovNroG, Me.txtAlmacen.Text
                Else
                    oAlmacen.EliminaMasterAlmacen lsBSCod, Me.txtAlmacen.Text
                    oAlmacen.EliminaKardexAlmacen lsBSCod, lnMovNroG, Me.txtAlmacen.Text
                End If
             Next I
             'Fin Gitu
          End If
          
          If lnMovNroOPG <> 0 And lnMovNroG = 0 Then 'Para Modificados
             oAlmacen.InsertaMovRef lnMovNro, lnMovNroOPG
          End If
        oAlmacen.CommitTrans
        
        cmdImprimir_Click
        Unload Me
        Exit Sub
    
    ElseIf lbMantenimiento Then
        lsOpeCodLocal = GetOpeMov(lnMovNroG)
        
        oAlmacen.BeginTrans
          'Inserta Mov
          lsMovNro = oAlmacen.GeneraMovNro(CDate(Me.mskFecha), Right(gsCodAge, 2), gsCodUser)
          oAlmacen.InsertaMov lsMovNro, lsOpeCodLocal, Me.txtComentario.Text, 20
          
          lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
          oAlmacen.InsertaMovGasto lnMovNro, Me.txtProveedor.Text, ""
          
          If lnMovNroG <> 0 Then  'Para Modificados
             oAlmacen.ActualizaMov lnMovNroG, , , gMovFlagModificado 'Modificado
             oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
             oAlmacen.InsertaMovRefAnt lnMovNro, lnMovNroG
             oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
          End If
          
          If lnMovNroOPG <> 0 And lnMovNroG = 0 Then 'Para Modificados
             oAlmacen.InsertaMovRef lnMovNro, lnMovNroOPG
          End If
          
          lsDocNI = Right(Me.lblTit, 13)
          
          'Inserta Documentos
          lsAgePagare = ""
          For I = 1 To Me.FlexDoc.Rows - 1
            If Me.FlexDoc.TextMatrix(I, 1) <> "" Then
                oAlmacen.InsertaMovDoc lnMovNro, Trim(Right(Me.FlexDoc.TextMatrix(I, 2), 8)), Trim(Me.FlexDoc.TextMatrix(I, 4)) & "-" & Me.FlexDoc.TextMatrix(I, 5), Format(CDate(Me.FlexDoc.TextMatrix(I, 3)), gsFormatoFecha)
                If Trim(Right(Me.FlexDoc.TextMatrix(I, 2), 5)) Then
                   lsAgePagare = Right(Me.FlexDoc.TextMatrix(I, 4), 2)
                End If
            End If
          Next I
          
          If lsAgePagare = "" Then lsAgePagare = Right(gsCodAge, 2)
          
          oAlmacen.InsertaMovDoc lnMovNro, 42, lsDocNI, Format(CDate(Me.mskFecha.Text), gsFormatoFecha)
          
          For I = 1 To Me.FlexDetalle.Rows - 1
            lsBSCod = Me.FlexDetalle.TextMatrix(I, 1)
            oAlmacen.InsertaMovBS lnMovNro, I, txtAlmacen.Text, Me.FlexDetalle.TextMatrix(I, 1)
            oAlmacen.InsertaMovCant lnMovNro, I, Me.FlexDetalle.TextMatrix(I, 3)
            If Me.optSoles.value Then
                oAlmacen.InsertaMovCta lnMovNro, I, Me.FlexDetalle.TextMatrix(I, 7), Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)), "#0.00")
                oAlmacen.InsertaMovOtrosItem lnMovNro, I, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5)), "#0.00"), ""
            Else
                oAlmacen.InsertaMovCta lnMovNro, I, Me.FlexDetalle.TextMatrix(I, 7), Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * CCur(Me.lblCompraG.Caption), "#0.00")
                oAlmacen.InsertaMovOtrosItem lnMovNro, I, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5)), "#0.00"), ""
            End If
          Next I
          
          lnContador = I
         
          If lsOpeCod = gnAlmaIngXAdjudicacion Then
             'Ctas provision
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), True)
                oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * lnPorProIni)
                oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(I, 5)), ""
                lnContador = lnContador + 1
             Next I
             'Ctas Pendientes para amortizar Creditos
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oAlmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7))
                oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * -1)
                oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(I, 5)), ""
                lnContador = lnContador + 1
             Next I
             'Ctas Pendientes para provicion
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), False)
                oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(I, 6)) * lnPorProIni) * -1)
                oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(I, 5)), ""
                lnContador = lnContador + 1
             Next I
          ElseIf lsOpeCod = gnAlmaIngXDacionPago Then
             'Ctas provision
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), True)
                oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6) * lnPorProIni))
                oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(I, 5)), ""
                lnContador = lnContador + 1
             Next I
             'Ctas Pendientes para amortizar Creditos
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = Replace(oAlmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7)), "AG", lsAgePagare)
                oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * -1)
                oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(I, 5)), ""
                lnContador = lnContador + 1
             Next I
             'Ctas Pendientes para provicion
             For I = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), False)
                oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(I, 6)) * lnPorProIni) * -1)
                oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(I, 5)), ""
                lnContador = lnContador + 1
             Next I
          End If
          
          If FlexSerie.TextMatrix(1, 1) <> "" Then
            For I = 1 To Me.FlexSerie.Rows - 1
              If InStr(1, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(I, 3)), 2), "[S]") <> 0 Then
                 oAlmacen.InsertaMovBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(I, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(I, 3)), 1), Me.FlexSerie.TextMatrix(I, 1), Me.FlexSerie.TextMatrix(I, 4), Me.FlexSerie.TextMatrix(I, 5), Me.FlexSerie.TextMatrix(I, 7)
              End If
            Next I
          End If
          
          If Me.optDolares.value Then
            oAlmacen.GeneraMovME lnMovNro, lsMovNro
          End If
        oAlmacen.CommitTrans
        'ARLO 20160126 ***
        Dim lsPalabra As String, lsOpe
        If (gsopecod = 591101) Then
        lsPalabra = "Registrado"
        lsOpe = 1
        ElseIf (gsopecod = 591102) Then
        lsPalabra = "Confirmado"
        lsOpe = 1
        ElseIf (gsopecod = 591103) Then
        lsOpe = 1
        lsPalabra = "Confirmado Ingreso Por transferencia"
        ElseIf (gsopecod = 591301) Then
        lsOpe = 2
        lsPalabra = "Modificado"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsOpe, "Se ha " & lsPalabra & " la Nota de Ingreso N° : " & lsDocNI
        Set objPista = Nothing
        '***
        cmdImprimir_Click
        Unload Me
        Exit Sub
    End If
    
    
    If Me.txtOCompra.Text <> "" Then
        ldFechaOC = oOpe.GetFechaDoc(Me.txtOCompra.Text, "90")
    End If
    
    oAlmacen.BeginTrans
      'Inserta Mov
      lsMovNro = oAlmacen.GeneraMovNro(CDate(Me.mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
      If lsOpeCod = gnAlmaIngXCompras Then
          oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabNoContable
      ElseIf lsOpeCod = gnAlmaExtornoXIngreso Then 'Rechazo
           oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, 21
      ElseIf Left(lsOpeCod, 4) = Left(gnAlmaIngXComprasConfirma, 4) Then
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, 20
      Else
        oAlmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabNoContable
      End If
      
      lnMovNro = oAlmacen.GetnMovNro(lsMovNro)
      oAlmacen.InsertaMovGasto lnMovNro, Me.txtProveedor.Text, ""
      
      If lnMovNroG <> 0 Then  'Para Modificados
        If gsopecod = "591103" Then '*** PEAC 20110714 SI ES UN INGRESO POR TRANSFERENCIA EL ESTADO GRABA COMO LEIDO para considerar en el kardex E INVENTARIO
            oAlmacen.ActualizaMov lnMovNroG, , , 4 'LEIDO
        Else
            oAlmacen.ActualizaMov lnMovNroG, , , gMovFlagModificado 'Modificado
        End If
         oAlmacen.InsertaMovRef lnMovNro, lnMovNroG
         oAlmacen.EliminaMovBSSerieparaActualizar lnMovNroG
      End If
      
      If lnMovNroOPG <> 0 And lnMovNroG = 0 Then 'Para Modificados
         oAlmacen.InsertaMovRef lnMovNro, lnMovNroOPG
      End If
      
      If lsOpeCod <> gnAlmaIngXComprasConfirma And Left(lsOpeCod, 4) = Left(gnAlmaIngXCompras, 4) Then  ' Genera documento
         lsDocNI = oAlmacen.GeneraDocNro(42, gMonedaExtranjera, Year(gdFecSis))
      Else
         lsDocNI = Right(Me.lblTit, 13)
      End If
      
      'Inserta Documentos
      lsAgePagare = ""
      For I = 1 To Me.FlexDoc.Rows - 1
        If Me.FlexDoc.TextMatrix(I, 1) <> "" Then
            oAlmacen.InsertaMovDoc lnMovNro, Trim(Right(Me.FlexDoc.TextMatrix(I, 2), 8)), Trim(Me.FlexDoc.TextMatrix(I, 4)) & "-" & Me.FlexDoc.TextMatrix(I, 5), Format(CDate(Me.FlexDoc.TextMatrix(I, 3)), gsFormatoFecha)
            If Trim(Right(Me.FlexDoc.TextMatrix(I, 2), 5)) Then
                lsAgePagare = Right(Me.FlexDoc.TextMatrix(I, 4), 2)
            End If
        End If
      Next I
      
      If lsAgePagare = "" Then lsAgePagare = Right(gsCodAge, 2)
      'Gitu
      If lsOpeCod <> "591101" Then
        Set rs = oOpe.GetMaestroAlmacen(Me.txtAlmacen.Text)
      End If
      
      For I = 1 To Me.FlexDetalle.Rows - 1
          If lsOpeCod <> "591101" And Me.FlexDetalle.TextMatrix(I, 3) > 0 Then
             lsBSCod = Me.FlexDetalle.TextMatrix(I, 1)
             'lsEsNuevo = oOpe.GetNuevoBienAlmacen(lsBsCod, txtAlmacen.Text)
             rs.Filter = "cBSCod = '" & lsBSCod & "'"
             If rs.RecordCount = 0 Then
                lsEsNuevo = True
             Else
                lsEsNuevo = False
             End If
             If lsEsNuevo Then
                lnCosProm = Me.FlexDetalle.TextMatrix(I, 6) / Me.FlexDetalle.TextMatrix(I, 3)
                lnStock = Me.FlexDetalle.TextMatrix(I, 3)
                lnCostoTotal = Me.FlexDetalle.TextMatrix(I, 6)
                lsUniMed = Mid(Me.FlexDetalle.TextMatrix(I, 2), InStr(Me.FlexDetalle.TextMatrix(I, 2), "(") + 1, InStr(Me.FlexDetalle.TextMatrix(I, 2), "[") - InStr(Me.FlexDetalle.TextMatrix(I, 2), "(") - 1)
                oAlmacen.InsertaMasterAlmacen lsBSCod, lnStock, Me.FlexDetalle.TextMatrix(I, 4), lnCosProm, lnCostoTotal, lsUniMed, txtAlmacen.Text, lsMoneda, lsMovNro, gdFecSis, lsTipoMovAlm
             Else
                lnStock = rs!nStock + Me.FlexDetalle.TextMatrix(I, 3)
                lnCostoTotal = rs!nCostoTotal + Me.FlexDetalle.TextMatrix(I, 6)
                'lnCosProm = lnCostoTotal / lnStock
                '****EJVG 20111122
                If lnStock <> 0 Then
                    lnCosProm = lnCostoTotal / lnStock
                Else
                    lnCosProm = 0
                End If
                'END ************
                oAlmacen.ActualizaMasterAlmacen Me.FlexDetalle.TextMatrix(I, 1), lnStock, Me.FlexDetalle.TextMatrix(I, 4), lnCosProm, lnCostoTotal, txtAlmacen.Text, gdFecSis, lsTipoMovAlm, lsMovNro
             End If
             oAlmacen.InsertaKardexAlmacen lnMovNro, I, lsBSCod, lsDocNI, Me.FlexDetalle.TextMatrix(I, 3), Me.FlexDetalle.TextMatrix(I, 4), lnStock, lnCosProm, lnCostoTotal, txtAlmacen.Text, lsTipoMovAlm, lsMovNro
         End If
      Next I
      'Fin Gitu
      oAlmacen.InsertaMovDoc lnMovNro, 42, lsDocNI, Format(CDate(Me.mskFecha.Text), gsFormatoFecha)
      
      If lsOpeCod = "562403" Or lsOpeCod = "561403" Then
         oAlmacen.InsertaMovDoc lnMovNro, 33, Me.txtOCompra.Text, Format(ldFechaOC, gsFormatoFecha)
      End If

    For I = 1 To Me.FlexDetalle.Rows - 1
        lsBSCod = Me.FlexDetalle.TextMatrix(I, 1)
        oAlmacen.InsertaMovBS lnMovNro, I, txtAlmacen.Text, Me.FlexDetalle.TextMatrix(I, 1)
        oAlmacen.InsertaMovCant lnMovNro, I, Me.FlexDetalle.TextMatrix(I, 3)
        oAlmacen.InsertaMovCta lnMovNro, I, Me.FlexDetalle.TextMatrix(I, 7), Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)))
        oAlmacen.InsertaMovOtrosItem lnMovNro, I, gcCtaIGV, Format(CCur(IIf(Me.FlexDetalle.TextMatrix(I, 5) = "", 0, Me.FlexDetalle.TextMatrix(I, 5)))), ""
      Next I
    'GUARDA NOTA DE INGRESO LIBRE ANPS
    If Me.ChkNIngresoLibre.value = 1 Then
                For i = 1 To Me.FlexDetalle.Rows - 1
                  oAlmacen.InsertaMovCotizacDet lnMovNro, i, FlexDetalle.TextMatrix(i, 2), Format(CDate(Me.mskFecha.Text), gsFormatoFecha)
                Next i
          End If
      'FIN DEL GUARDADO ANPS
      lnContador = I
      
      If lsOpeCod = gnAlmaIngXAdjudicacion Then
         'Ctas provision
         For I = 1 To Me.FlexDetalle.Rows - 1
            lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), True)
            oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * lnPorProIni)
            oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5))), ""
            lnContador = lnContador + 1
         Next I
         'Ctas Pendientes para amortizar Creditos
         For I = 1 To Me.FlexDetalle.Rows - 1
            lsCtaCont = oAlmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7))
            oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * -1)
            oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5))), ""
            lnContador = lnContador + 1
         Next I
         'Ctas Pendientes para provicion
         For I = 1 To Me.FlexDetalle.Rows - 1
            lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), False)
            oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(I, 6)) * lnPorProIni) * -1)
            oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5))), ""
            lnContador = lnContador + 1
         Next I
      ElseIf lsOpeCod = gnAlmaIngXDacionPago Then
         'Ctas provision
         For I = 1 To Me.FlexDetalle.Rows - 1
            lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), True)
            oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6) * lnPorProIni))
            oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5))), ""
            lnContador = lnContador + 1
         Next I
         'Ctas Pendientes para amortizar Creditos
         For I = 1 To Me.FlexDetalle.Rows - 1
            lsCtaCont = Replace(oAlmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7)), "AG", lsAgePagare)
            oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(I, 6)) * -1)
            oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5))), ""
            lnContador = lnContador + 1
         Next I
         'Ctas Pendientes para provicion
         For I = 1 To Me.FlexDetalle.Rows - 1
            lsCtaCont = oAlmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(I, 7), False)
            oAlmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(I, 6)) * lnPorProIni) * -1)
            oAlmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(I, 5))), ""
            lnContador = lnContador + 1
         Next I
      End If

      If FlexSerie.TextMatrix(1, 1) <> "" Then
        For I = 1 To Me.FlexSerie.Rows - 1
          If InStr(1, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(I, 3)), 2), "[S]") <> 0 Then
            oAlmacen.InsertaMovBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(I, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(I, 3)), 1), Me.FlexSerie.TextMatrix(I, 1), Me.FlexSerie.TextMatrix(I, 4), Me.FlexSerie.TextMatrix(I, 5), Me.FlexSerie.TextMatrix(I, 7)
          End If
        Next I
      End If
      
      If Me.optDolares.value Then
        oAlmacen.GeneraMovME lnMovNro, lsMovNro
      End If
    oAlmacen.CommitTrans
    cmdImprimir_Click

    MsgBox "Se guardo exitosamente" 'ANPS
    Call limpiar 'ANPS
    'ARLO 20160126 ***
    If (gsopecod = 591101) Then
        lsPalabra = "Registrado"
        End If
        If (gsopecod = 591102) Then
        lsPalabra = "Confirmado"
        ElseIf (gsopecod = 591103) Then
        lsPalabra = "Confirmado Ingreso Por transferencia"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se ha " & lsPalabra & " la Nota de Ingreso N° : " & lsDocNI
        Set objPista = Nothing
        '***
     ' Unload Me anps comentado
End Sub
Private Function limpiar()
    LimpiaFlex Me.FlexDetalle
ChkNIngresoLibre.Enabled = True
    ChkNIngresoLibre.value = 0
    Me.txtComentario.Text = ""
    Me.lblTotalG.Caption = "0.00"
    Me.lblTotalIGV.Caption = "0.00"
    LimpiaFlex Me.FlexDoc
LimpiaFlex Me.FlexSerie
    Me.txtOCompra.Enabled = True
    Me.txtProveedor.Text = ""
    Me.lblProveedorNombre.Caption = ""
    Me.txtOCompra.Text = ""
    Me.lblOCompra.Caption = ""
    Me.lblTotalIGVDet.Caption = "0.00"
    Me.lblTotalGDet.Caption = "0.00"
    Call Form_Load()
End Function
Private Sub cmdImprimir_Click()
    'If Me.txtNotaIng.Text = "" Then
    '    MsgBox "Debe elegir una nota de ingreso.", vbInformation, "Aviso"
    '    Me.txtNotaIng.SetFocus
    '    Exit Sub
    'End If
    
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    Dim lsCadena As String
    Dim lsCadenaSerie As String
    Dim lsDocNom As String * 20
    Dim lsDocFec As String * 12
    Dim lsDocNum As String * 20
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsItem As String * 5
    Dim lsCodigo As String * 15
    Dim lsNombre As String * 50
    Dim lsUnidad As String * 10
    Dim lsCantidad As String * 10
    Dim lsPrecio As String * 15
    Dim lsTotal As String * 15
    Dim I As Long
    Dim J As Long
    
    lsCadena = ""
    lsCadena = lsCadena & oImpresora.gPrnCondensadaON
     
    lsCadena = lsCadena & CabeceraPagina1(lblTit.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(45) & lblAlmacenG.Caption & oImpresora.gPrnSaltoLinea
    If txtOCompra.Text <> "" Then
       lsCadena = lsCadena & Space(45) & "ORDEN COMPRA: " & txtOCompra.Text & oImpresora.gPrnSaltoLinea
    End If
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTextoCadena(Me.txtComentario.Text, 105, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Me.fraProveedor.Caption & " : " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
      lsCadena = lsCadena & Space(5) & "MOTIVO DE INGRESO : " & IIf(Me.ChkNIngresoLibre.value = 1, "INGRESO DE NOTA LIBRE", Me.Caption) & oImpresora.gPrnSaltoLinea 'ANPS
    lsCadena = lsCadena & Space(5) & "      DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
   
    For J = 1 To Me.FlexDoc.Rows - 1
        If Me.FlexDoc.TextMatrix(J, 1) <> "" Then
             lsDocNom = Left(Me.FlexDoc.TextMatrix(J, 2), 20)
             RSet lsDocNum = Trim(Me.FlexDoc.TextMatrix(J, 4)) & "-" & Me.FlexDoc.TextMatrix(J, 5)
             RSet lsDocFec = Format(CDate(Me.FlexDoc.TextMatrix(J, 3)), gsFormatoFechaView)
             lsCadena = lsCadena & Space(5) & "      " & lsDocNom & lsDocNum & lsDocFec & oImpresora.gPrnSaltoLinea
             lnItem = lnItem + 1
             If lnItem > 35 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina1(lblTit.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & JustificaTextoCadena(Me.txtComentario.Text, 105, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & Space(5) & Me.fraProveedor.Caption & " : " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & Space(5) & "MOTIVO DE INGRESO : " & IIf(Me.ChkNIngresoLibre.value = 1, "INGRESO DE NOTA LIBRE", Me.Caption) & oImpresora.gPrnSaltoLinea 'ANPS
                lsCadena = lsCadena & Space(5) & "      DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
             End If
        
        End If
    Next J
    
    lsCadena = lsCadena & Encabezado1("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;CANTIDAD;10; ;5;PRECIO;17; ;5;TOTAL;10; ;3;", lnItem)

    For I = 1 To Me.FlexDetalle.Rows - 1
        lsItem = Format(I, "0000")
        lsCodigo = Me.FlexDetalle.TextMatrix(I, 1)
        lsNombre = Me.FlexDetalle.TextMatrix(I, 2)
        lsCantidad = Me.FlexDetalle.TextMatrix(I, 3)
        RSet lsPrecio = Format(Me.FlexDetalle.TextMatrix(I, 4), "#,#00.00")
        RSet lsTotal = Format(Me.FlexDetalle.TextMatrix(I, 6), "#,#00.00")
        
        If Me.FlexDetalle.TextMatrix(I, 3) = "0.00" Or Me.FlexDetalle.TextMatrix(I, 3) = "0" Then
        Else
           lsCadena = lsCadena & Space(5) & "  " & lsItem & lsCodigo & "  " & lsNombre & " " & lsCantidad & "  " & lsPrecio & "  " & lsTotal & oImpresora.gPrnSaltoLinea
        End If
        
        lnItem = lnItem + 1
        If InStr(1, Me.FlexDetalle.TextMatrix(I, 2), "[S]") <> 0 Then
            lsItem = ""
            lsCodigo = ""
            lsCadenaSerie = ""
            For J = 1 To Me.FlexSerie.Rows - 1
                If Me.FlexSerie.TextMatrix(J, 3) = Me.FlexDetalle.TextMatrix(I, 0) Then
                    If lsCadenaSerie = "" Then
                        lsCadenaSerie = Me.FlexSerie.TextMatrix(J, 1)
                    Else
                        lsCadenaSerie = lsCadenaSerie & " / " & Me.FlexSerie.TextMatrix(J, 1)
                        lnItem = lnItem + 1
                    End If
                    If J Mod 3 = 0 Then
                        lsCadena = lsCadena & Space(5) & lsItem & lsCodigo & "    " & lsCadenaSerie & oImpresora.gPrnSaltoLinea
                        lsCadenaSerie = ""
                    End If
                End If
            Next J
            lsCadena = lsCadena & Space(5) & lsItem & lsCodigo & "    " & lsCadenaSerie & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
        End If
        
        If lnItem > 44 Then
             lnItem = 0
             lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
             lsCadena = lsCadena & CabeceraPagina1(lblTit.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & Space(5) & Me.fraProveedor.Caption & " : " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(5) & "MOTIVO DE INGRESO : " & IIf(Me.ChkNIngresoLibre.value = 1, "INGRESO DE NOTA LIBRE", Me.Caption) & oImpresora.gPrnSaltoLinea 'ANPS
             lsCadena = lsCadena & Space(5) & "      DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
            
             For J = 1 To Me.FlexDoc.Rows - 1
                 If Me.FlexDoc.TextMatrix(J, 1) <> "" Then
                      lsDocNom = Left(Me.FlexDoc.TextMatrix(J, 2), 20)
                      RSet lsDocNum = Trim(Me.FlexDoc.TextMatrix(J, 4)) & "-" & Me.FlexDoc.TextMatrix(J, 5)
                      RSet lsDocFec = Format(CDate(Me.FlexDoc.TextMatrix(J, 3)), gsFormatoFechaView)
                      lsCadena = lsCadena & Space(5) & "      " & lsDocNom & lsDocNum & lsDocFec & oImpresora.gPrnSaltoLinea
                 End If
             Next J
             
             lsCadena = lsCadena & Encabezado1("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;CANTIDAD;10; ;5;PRECIO;17; ;5;TOTAL;10; ;3;", lnItem)
        End If
    Next I
      
    lsCadena = lsCadena & Space(4) & String(120, "=") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
      
      
    lsCadena = lsCadena & Space(5) & "----------------------------          ----------------------------          --------------------------" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "    FIRMA ALMACENERO                          Logistcia                                Vo Bo          " & oImpresora.gPrnSaltoLinea
    
    oPrevio.Show lsCadena, Caption, True, 66
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FlexDetalle_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim lnCorrelativoIni  As Currency
    Dim lbCorrelativo As Boolean
    Dim lnI As Long
    
    If pnCol = 1 And (lsOpeCod <> gnAlmaIngXCompras And lsOpeCod <> gnAlmaIngXComprasConfirma) Then
        If lsOpeCod = gnAlmaIngXDacionPago Or lsOpeCod = gnAlmaIngXAdjudicacion Then
            Me.FlexDetalle.TextMatrix(Me.FlexDetalle.row, 4) = Format(oAlmacen.GetPrePromedio("1", Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 1), "#0.00")
        ElseIf lsOpeCod = gnAlmaIngXEmbargo Then
            Me.FlexDetalle.TextMatrix(Me.FlexDetalle.row, 4) = Format(oAlmacen.GetPrePromedio("1", Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 2), "#0.00")
        Else
            Me.FlexDetalle.TextMatrix(Me.FlexDetalle.row, 4) = Format(oAlmacen.GetPrePromedio("1", Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), 0), "#0.00")
        End If
    End If
    
    If pnCol = 1 Then
        If Me.FlexDetalle.TextMatrix(pnRow, pnCol) = "" Then Exit Sub
        lbCorrelativo = oAlmacen.VerfBSCorrela(Me.FlexDetalle.TextMatrix(pnRow, pnCol))
        
        If lbCorrelativo Then
            lnCorrelativoIni = oAlmacen.GetBSCorrelaIni(Me.FlexDetalle.TextMatrix(pnRow, pnCol))
            
            If GetUltCodVig(Me.FlexDetalle.TextMatrix(pnRow, pnCol), pnRow) > lnCorrelativoIni Then
                lnCorrelativoIni = GetUltCodVig(Me.FlexDetalle.TextMatrix(pnRow, pnCol), pnRow)
            End If
            
            If Not IsNumeric(FlexSerie.TextMatrix(lnI, 3)) Then Exit Sub
            For lnI = 1 To Me.FlexSerie.Rows - 1
                If FlexSerie.TextMatrix(lnI, 3) = pnRow Then
                    FlexSerie.TextMatrix(lnI, 1) = Trim(Str(Year(gdFecSis))) & "-" & Format(lnCorrelativoIni, "00000000")
                    lnCorrelativoIni = lnCorrelativoIni + 1
                End If
            Next lnI
        
        End If
    End If
    
End Sub

Private Sub FlexDetalle_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim lnCorrelativoIni  As Currency
    Dim lbCorrelativo As Boolean
    Dim lnI As Long
    Dim lnUltimo As Long
    Dim nIGVAcum As Currency
    Dim nTotAcum As Currency
	 'ANPS VALIDACION DE CHECK Y DATOS
    '------------------------------------------
    If pnCol = 1 Then
        If Not IsNumeric(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1)) Then
        If Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1) = "" Then
        MsgBox ("Ingresar valor correcto")
        LimpiaFlex Me.FlexDetalle
         Exit Sub
        Else
        MsgBox ("Ingresar valores correctos")
         LimpiaFlex Me.FlexDetalle
         
         Exit Sub
         End If
        End If
     End If
    

    If Me.ChkNIngresoLibre.value = 1 Then
        Dim Total As Currency
        Total = 0
        If pnCol = 3 Or pnCol = 4 Then
        If pnCol = 3 Then
                If CCur(IIf(FlexDetalle.TextMatrix(pnRow, 3) = "", 0, FlexDetalle.TextMatrix(pnRow, 3))) = 0 Then
                    MsgBox "Cantidad no debe ser menor que Cero...!", vbCritical, "Aviso"
                    Exit Sub
                End If
        ElseIf pnCol = 4 Then
            If CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 4) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 4))) = 0 Then
                    MsgBox "Precio no debe ser menor que Cero...!", vbCritical, "Aviso"
                    Exit Sub
              End If
        End If
        'anps
            If CCur(IIf(FlexDetalle.TextMatrix(pnRow, 3) = "", 0, FlexDetalle.TextMatrix(pnRow, 3))) >= 0 Or CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 4) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 4))) >= 0 Then
              Total = CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 4) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 4))) + CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 5) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 5)))
             Me.FlexDetalle.TextMatrix(pnRow, 6) = Total * CCur(IIf(FlexDetalle.TextMatrix(pnRow, 3) = "", 0, FlexDetalle.TextMatrix(pnRow, 3)))
            Exit Sub
            Else
             MsgBox "Valor no debe ser menor que Cero...!", vbCritical, "Aviso"
            End If
        End If
        If pnCol = 5 And CCur(IIf(FlexDetalle.TextMatrix(pnRow, 3) = "", 0, FlexDetalle.TextMatrix(pnRow, 3))) >= 0 And CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 4) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 4))) >= 0 Then
            Total = CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 4) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 4))) + CCur(IIf(Me.FlexDetalle.TextMatrix(pnRow, 5) = "", 0, Me.FlexDetalle.TextMatrix(pnRow, 5)))
             Me.FlexDetalle.TextMatrix(pnRow, 6) = Total * CCur(IIf(FlexDetalle.TextMatrix(pnRow, 3) = "", 0, FlexDetalle.TextMatrix(pnRow, 3)))
            Exit Sub
        End If
        Exit Sub
    End If
    '------------------------------------------
    'FIN ANPS
    
    If pnCol = 6 Or pnCol = 3 Then
        If IsNumeric(FlexDetalle.TextMatrix(pnRow, 6)) And IsNumeric(FlexDetalle.TextMatrix(pnRow, 3)) Then
            If CCur(FlexDetalle.TextMatrix(pnRow, 3)) <> 0 Then
                FlexDetalle.TextMatrix(pnRow, 4) = Format(CCur(FlexDetalle.TextMatrix(pnRow, 6)) / CCur(FlexDetalle.TextMatrix(pnRow, 3)), "0.00")
            End If
            FlexDetalle.TextMatrix(pnRow, 5) = Format(CCur(FlexDetalle.TextMatrix(pnRow, 6) * gnIGV), "0.00")

'            If CCur(FlexDetalle.TextMatrix(pnRow, 3)) <> 0 Then
'                FlexDetalle.TextMatrix(pnRow, 6) = FlexDetalle.TextMatrix(pnRow, 3) * FlexDetalle.TextMatrix(pnRow, 4)
'                FlexDetalle.TextMatrix(pnRow, 4) = Format(CCur(FlexDetalle.TextMatrix(pnRow, 6)) / CCur(FlexDetalle.TextMatrix(pnRow, 3)), "0.00")
'                Else
'                FlexDetalle.TextMatrix(pnRow, 6) = 0
'                'FlexDetalle.TextMatrix(pnRow, 4) = 0
'
'            End If
'            FlexDetalle.TextMatrix(pnRow, 5) = Format(CCur(FlexDetalle.TextMatrix(pnRow, 6) * gnIGV), "0.00")
'
            If InStr(1, Me.FlexDetalle.TextMatrix(pnRow, 2), "[S]") = 0 Then Exit Sub
            nIGVAcum = 0
            nTotAcum = 0
            For lnI = 1 To Me.FlexSerie.Rows - 1
                If IsNumeric(FlexSerie.TextMatrix(lnI, 3)) Then
                    If FlexSerie.TextMatrix(lnI, 3) = pnRow Then
                        
                        If Me.FlexDetalle.TextMatrix(pnRow, 3) = 0 Then
                                FlexSerie.TextMatrix(lnI, 4) = 0
                                FlexSerie.TextMatrix(lnI, 5) = 0
                           Else
                                FlexSerie.TextMatrix(lnI, 4) = Round(Me.FlexDetalle.TextMatrix(pnRow, 5) / Me.FlexDetalle.TextMatrix(pnRow, 3), 2)
                                FlexSerie.TextMatrix(lnI, 5) = Round(Me.FlexDetalle.TextMatrix(pnRow, 6) / Me.FlexDetalle.TextMatrix(pnRow, 3), 2)
                            
                        End If
                        nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(lnI, 4)
                        nTotAcum = nTotAcum + FlexSerie.TextMatrix(lnI, 5)
                        lnUltimo = lnI
                    End If
                End If
            Next lnI
            If IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4)) Then
                If nIGVAcum <> Me.FlexDetalle.TextMatrix(pnRow, 5) Then
                    FlexSerie.TextMatrix(lnUltimo, 4) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) - (nIGVAcum - Me.FlexDetalle.TextMatrix(pnRow, 5))
                End If
            End If
            If IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5)) Then
                If nTotAcum <> Me.FlexDetalle.TextMatrix(pnRow, 6) Then
                    FlexSerie.TextMatrix(lnUltimo, 5) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) - (nTotAcum - Me.FlexDetalle.TextMatrix(pnRow, 6))
                End If
            End If
            SumaTotalDetalle pnRow
        End If
    ElseIf pnCol = 5 Then
        If InStr(1, Me.FlexDetalle.TextMatrix(pnRow, 2), "[S]") <= 0 Then Exit Sub
        
        nIGVAcum = 0
        
        For lnI = 1 To Me.FlexSerie.Rows - 1
            If IsNumeric(FlexSerie.TextMatrix(lnI, 3)) Then
                If FlexSerie.TextMatrix(lnI, 3) = pnRow Then
                    FlexSerie.TextMatrix(lnI, 4) = Round(Me.FlexDetalle.TextMatrix(pnRow, 5) / Me.FlexDetalle.TextMatrix(pnRow, 3), 2)
                    nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(lnI, 4)
                    lnUltimo = lnI
                End If
            End If
        Next lnI
        
        If IsNumeric(FlexSerie.TextMatrix(lnUltimo, 4)) Then
            If nIGVAcum <> Me.FlexDetalle.TextMatrix(pnRow, 5) Then
                FlexSerie.TextMatrix(lnUltimo, 4) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) - (nIGVAcum - Me.FlexDetalle.TextMatrix(pnRow, 5))
            End If
        End If
        SumaTotalDetalle pnRow
    End If
End Sub

Private Sub FlexDetalle_RowColChange()
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim lsCtaCnt As String
    Dim I As Integer
    Dim lnContador As Integer
    Dim lnEncontrar As Integer
    Dim lnI As Integer
    Dim lnTotal As Currency
    Dim lnTotalIGV As Currency
    Dim lbCorrelativo  As Boolean
    Dim lnCorrelativoIni As Long
    
    Dim nIGVAcum As Currency
    Dim nTotAcum As Currency
    
    Dim lnItems As Long

    'VALIDACION DE DATOS ANPS
     If Not IsNumeric(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1)) Then
            MsgBox ("Ingresar valores correctos. Agregar nuevamente")
             LimpiaFlex Me.FlexDetalle
             Exit Sub
      End If
     'FIN ANPS    
	
    lnTotal = 0
    lnTotalIGV = 0
    For lnI = 1 To Me.FlexDetalle.Rows - 1
        If IsNumeric(FlexDetalle.TextMatrix(lnI, 5)) And IsNumeric(FlexDetalle.TextMatrix(lnI, 6)) Then
            lnTotal = lnTotal + CCur(FlexDetalle.TextMatrix(lnI, 6))
            lnTotalIGV = lnTotalIGV + CCur(FlexDetalle.TextMatrix(lnI, 5))
        End If
    Next lnI
    
    Me.lblTotalG.Caption = Format(lnTotal, "#,##0.00")
    Me.lblTotalIGV.Caption = Format(lnTotalIGV, "#,##0.00")
    
    If InStr(1, Me.FlexDetalle.TextMatrix(FlexDetalle.row, 2), "[S]") <> 0 Or Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1) = "" Then
        lnContador = 0
        If Me.FlexDetalle.TextMatrix(FlexDetalle.row, 3) = "" Then
            Exit Sub
        End If
        
        For I = 1 To CInt(Me.FlexSerie.Rows - 1)
            If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
                lnContador = lnContador + 1
                FlexSerie.RowHeight(I) = 285
            Else
                FlexSerie.RowHeight(I) = 0
                'FlexSerie.RowHeight(I) = 285
            End If
        Next I
        
        SumaTotalDetalle FlexDetalle.row
        
        If lnContador <> CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 3)) Then
            I = 0
            lnEncontrar = 0
            While lnEncontrar < lnContador
                I = I + 1
                If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
                    Me.FlexSerie.EliminaFila I
                    lnEncontrar = lnEncontrar + 1
                    I = I - 1
                End If
            Wend
            
            'For lnItems = 1 To Me.FlexDetalle.Rows - 1
                lnItems = FlexDetalle.row
                If Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1) = Me.FlexDetalle.TextMatrix(lnItems, 1) And FlexDetalle.row <> lnItems Then
                    lnContador = CLng(Me.FlexDetalle.TextMatrix(lnItems, 3))
                    lnEncontrar = 0
                    I = 0
                    While lnEncontrar < lnContador
                        I = I + 1
                        If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(lnItems, 0) Then
                            Me.FlexSerie.EliminaFila I
                            lnEncontrar = lnEncontrar + 1
                            I = I - 1
                        End If
                    Wend
                End If
            'Next lnItems
            
            lbCorrelativo = False
            
            If Not lbMantenimiento Then
                
                lbCorrelativo = oAlmacen.VerfBSCorrela(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1))
                
                If lbCorrelativo Then
                    lnCorrelativoIni = oAlmacen.GetBSCorrelaIni(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1))
                End If
                'If GetUltCodVig(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1)) > lnCorrelativoIni Then
                '    lnCorrelativoIni = GetUltCodVig(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1))
                'End If
            Else
                'If lnContador > CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 3)) Then
                '   lbCorrelativo = oALmacen.VerfBSCorrela(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1))
                '   lnCorrelativoIni = oALmacen.GetBSCorrelaIniMov(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), lnMovNroG)
                'End If
            End If
            
            'For lnItems = 1 To Me.FlexDetalle.Rows - 1
            
                If Not IsNumeric(Me.FlexDetalle.TextMatrix(lnItems, 5)) Or Not IsNumeric(Me.FlexDetalle.TextMatrix(lnItems, 6)) Then Exit Sub
                nIGVAcum = 0
                nTotAcum = 0
                lnItems = FlexDetalle.row
                If Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1) = Me.FlexDetalle.TextMatrix(lnItems, 1) Then
                    nIGVAcum = 0
                    For I = 1 To CLng(Me.FlexDetalle.TextMatrix(lnItems, 3))
                        If Me.FlexSerie.TextMatrix(1, 3) = "" Then
                            Me.FlexSerie.AdicionaFila
                        Else
                            Me.FlexSerie.AdicionaFila , , True
                        End If
                        
                        If lbCorrelativo Then
                            FlexSerie.TextMatrix(FlexSerie.Rows - 1, 1) = Trim(Str(Year(gdFecSis))) & "-" & Format(lnCorrelativoIni, "00000000")
                            lnCorrelativoIni = lnCorrelativoIni + 1
                        End If
                        
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 0) = Me.FlexDetalle.TextMatrix(lnItems, 0)
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 2) = Me.FlexDetalle.TextMatrix(lnItems, 0)
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 3) = Me.FlexDetalle.TextMatrix(lnItems, 0)
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) = Round(Me.FlexDetalle.TextMatrix(lnItems, 5) / Me.FlexDetalle.TextMatrix(lnItems, 3), 2)
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) = Round(Me.FlexDetalle.TextMatrix(lnItems, 6) / Me.FlexDetalle.TextMatrix(lnItems, 3), 2)
                        nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4)
                        nTotAcum = nTotAcum + FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5)
                    Next I
                    
                    If nIGVAcum <> Me.FlexDetalle.TextMatrix(lnItems, 5) Then
                        If IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4)) Then
                            FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) - (nIGVAcum - Me.FlexDetalle.TextMatrix(lnItems, 5))
                        End If
                    End If
                    If nTotAcum <> Me.FlexDetalle.TextMatrix(lnItems, 6) Then
                        If IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5)) Then
                            FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) - (nTotAcum - Me.FlexDetalle.TextMatrix(lnItems, 6))
                        End If
                    End If
                    
                End If
            'Next lnItems
        
        End If
    Else
        For I = 1 To CInt(Me.FlexSerie.Rows - 1)
            FlexSerie.RowHeight(I) = 0
        Next I
    End If
    
    If lbMantenimiento Or lbExtorno Then
        lsCtaCnt = GetCtaCntBS(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), GetOpeMov(lnMovNroG), Format(Me.txtAlmacen.Text, "00"))
    Else
        lsCtaCnt = GetCtaCntBS(Me.FlexDetalle.TextMatrix(FlexDetalle.row, 1), lsOpeCod, Format(Me.txtAlmacen.Text, "00"))
    End If
    'EJVG 20111031 ****************************************
    If lsCtaCnt = "" Then 'ANPS
        MsgBox "No se a especificado la cuenta contable para este Bien de Consumo", vbCritical, "Aviso"
     For I = 1 To CInt(Me.FlexSerie.Rows - 1)
            If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
                lnContador = lnContador + 1
            End If
        Next I

        I = 0
        While lnEncontrar < lnContador
            I = I + 1
            If FlexSerie.TextMatrix(I, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.row, 0) Then
                Me.FlexSerie.EliminaFila I
            lnEncontrar = lnEncontrar + 1
                I = I - 1
            End If
    Wend
    
    For I = 1 To Me.FlexSerie.Rows - 1
            If IsNumeric(FlexSerie.TextMatrix(I, 3)) Then
                If FlexSerie.TextMatrix(I, 3) > Me.FlexDetalle.row Then
                    FlexSerie.TextMatrix(I, 3) = Trim(Str(CInt(FlexSerie.TextMatrix(I, 3)) - 1))
                    FlexSerie.TextMatrix(I, 0) = FlexSerie.TextMatrix(I, 3)
                    FlexSerie.TextMatrix(I, 2) = FlexSerie.TextMatrix(I, 3)
                End If
            End If
        Next I

        Me.FlexDetalle.EliminaFila Me.FlexDetalle.row
    Exit Sub
    End If ' FIN ANPS
    If lsOpeCod = gnAlmaIngXTransferencia Then
        For I = 1 To FlexDetalle.Rows - 1
            FlexDetalle.TextMatrix(I, 7) = lsCtaCnt
        Next
    Else
        FlexDetalle.TextMatrix(FlexDetalle.row, 7) = lsCtaCnt
    End If
    'END **************************************************
    'FlexDetalle.TextMatrix(FlexDetalle.Row, 7) = lsCtaCnt
    
    Set oAlmacen = Nothing
End Sub

Private Sub FlexDoc_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If pnCol = 4 Then
        FlexDoc.TextMatrix(pnRow, pnCol) = Format(FlexDoc.TextMatrix(pnRow, pnCol), "000")
    ElseIf pnCol = 5 Then
        FlexDoc.TextMatrix(pnRow, pnCol) = Format(FlexDoc.TextMatrix(pnRow, pnCol), "0000000")
    End If
End Sub

Private Sub FlexDoc_RowColChange()
    If FlexDoc.TextMatrix(FlexDoc.row, 1) = "" And FlexDoc.col <> 1 Then
        FlexDoc.lbEditarFlex = False
    Else
        FlexDoc.lbEditarFlex = True
    End If
    
End Sub



Private Sub FlexSerie_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim nIGVAcum As Currency
    Dim nTotAcum As Currency
    Dim lnI As Integer
    Dim lnUltimo As Integer
    
    If pnCol = 4 Then
        nIGVAcum = 0
        lnUltimo = -1
        
        If CCur(Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 5)) < CCur(FlexSerie.TextMatrix(pnRow, pnCol)) Then
            MsgBox "El valor no puede ser mayor que el total.", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
        
        For lnI = 1 To Me.FlexSerie.Rows - 1
            If Me.FlexSerie.TextMatrix(pnRow, 3) = Me.FlexSerie.TextMatrix(lnI, 3) And pnRow <> lnI Then
                Me.FlexSerie.TextMatrix(lnI, 4) = Round((Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 5) - FlexSerie.TextMatrix(pnRow, pnCol)) / (Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 3) - 1), 2)
                nIGVAcum = nIGVAcum + Me.FlexSerie.TextMatrix(lnI, 4)
                lnUltimo = lnI
            End If
        Next lnI
        
        
        nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(pnRow, pnCol)
        
        If lnUltimo = -1 Then
            lnUltimo = pnRow
        End If
        
        If nIGVAcum <> Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 5) Then
            Me.FlexSerie.TextMatrix(lnUltimo, 4) = Me.FlexSerie.TextMatrix(lnUltimo, 4) + ((Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 5) - nIGVAcum))
        End If
    
    ElseIf pnCol = 5 Then
        nTotAcum = 0
    
        lnUltimo = -1
        
        If CCur(Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 6)) < CCur(FlexSerie.TextMatrix(pnRow, pnCol)) Then
            MsgBox "El valor no puede ser mayor que el total.", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
        
        For lnI = 1 To Me.FlexSerie.Rows - 1
            If Me.FlexSerie.TextMatrix(pnRow, 3) = Me.FlexSerie.TextMatrix(lnI, 3) And pnRow <> lnI Then
                Me.FlexSerie.TextMatrix(lnI, 5) = Round((Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 6) - FlexSerie.TextMatrix(pnRow, pnCol)) / (Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 3) - 1), 2)
                nTotAcum = nTotAcum + Me.FlexSerie.TextMatrix(lnI, 5)
                lnUltimo = lnI
            End If
        Next lnI
        
        
        nTotAcum = nTotAcum + FlexSerie.TextMatrix(pnRow, pnCol)
        
        If lnUltimo = -1 Then
            lnUltimo = pnRow
        End If
        
        If nTotAcum <> Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 6) Then
            Me.FlexSerie.TextMatrix(lnUltimo, 5) = Me.FlexSerie.TextMatrix(lnUltimo, 5) + ((Me.FlexDetalle.TextMatrix(Me.FlexSerie.TextMatrix(pnRow, 3), 6) - nTotAcum))
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim oGen As DLogGeneral
    Set oGen = New DLogGeneral
    Dim oMov As DMov
    Set oMov = New DMov
        
    lbGrabar = False
    GetTipCambio gdFecSis, Not gbBitCentral
    
    lnMovNroG = 0
    lnMovNroOPG = 0
    
    lnPorProIni = oGen.CargaParametro(5000, 1005) / 100
    
    Me.lblFijoG.Caption = Format(gnTipCambio, "#.###")
    Me.lblCompraG.Caption = Format(gnTipCambioC, "#.###")
    
    Caption = lsCaptionG
    
    If Mid(lsOpeCod, 3, 1) = gcMNDig Then
        Me.optDolares.value = False
        Me.optSoles.value = True
    Else
        Me.optDolares.value = True
        Me.optSoles.value = False
    End If
    
    '*** PEAC 20110712
    If lsOpeCod = "591103" Then
        Me.txtNotaIng.Visible = lbIngreso
        Me.lblNotaIng.Visible = lbIngreso
        Me.lblNotaIngG.Visible = lbIngreso
        Me.lblNotaIng.Caption = "Guia Remisión:"
    Else
        Me.txtNotaIng.Visible = Not lbIngreso
        Me.lblNotaIng.Visible = Not lbIngreso
        Me.lblNotaIngG.Visible = Not lbIngreso
    End If
    '*** FIN PEAC
    
    Me.cmdImprimir.Visible = Not lbIngreso
    
    Me.FlexDoc.CargaCombo oDoc.GetDocOpe(lsOpeCod, False)
    
    Me.mskFecha = Format(gdFecSis, gsFormatoFechaView)
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    'Me.txtAlmacen.Text = "1"
    Me.txtAlmacen.Text = gsCodAge
    Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
        
    If lbMantenimiento Or lbExtorno Then
        Me.txtNotaIng.rs = oDoc.GetNotaIngreso("20", lsOpeCod)
        txtNotaIng.Visible = True
        Me.txtOCompra.Visible = False
        Me.lblOCompra.Visible = False
        Me.lblOrdenCompra.Visible = False
        'Exit Sub
    Else
        'If lsOpeCod = gnAlmaIngXComprasConfirma And Not lbReporte Then Me.txtNotaIng.rs = oDoc.GetNotaIngreso("13", lsOpeCod)
        If lsOpeCod = gnAlmaIngXComprasConfirma And Not lbReporte Then Me.txtNotaIng.rs = oDoc.GetNotaIngreso("13", lsOpeCod, , CInt(Right(gsCodAge, 2))) 'EJVG20140320
        If lsOpeCod = "591103" And Not lbReporte Then Me.txtNotaIng.rs = oDoc.GetGuiaSalidaPorTransf(Mid(lsOpeCod, 3, 1), Right(gsCodAge, 2)) '*** PEAC 20110712
    End If
        
    If lbIngreso Then Me.lblTit.Caption = "Nota de Ingreso : " & oMov.GeneraDocNro(42, gMonedaExtranjera, Year(gdFecSis))
    
    If Not lbConfirma Then
        If lbIngreso Then
            If Not (lsOpeCod = gnAlmaIngXCompras) Then
                Me.txtOCompra.Visible = False
                Me.lblOCompra.Visible = False
                Me.lblOrdenCompra.Visible = False
                
                Me.fraProveedor.Caption = "Persona"
                'JEOM
                '************************************
                Dim oPersona As UPersona
                Set oPersona = New UPersona
                If gsCodPersUser <> "" Then
                    oPersona.ObtieneClientexCodigo gsCodPersUser
                    Me.txtProveedor.Text = gsCodPersUser
                    Me.lblProveedorNombre.Caption = oPersona.sPersNombre
                End If
                Dim rsAge As ADODB.Recordset
                Set rsAge = New ADODB.Recordset
                
               Set rsAge = oDoc.GetAlmacenesUsuarioAlmacen(gsCodAge)
                Me.txtAlmacen.Text = rsAge(0)
                Me.lblAlmacenG.Caption = rsAge(1)
                '***********************************
                'FIN JEOM
                
            End If
        Else
            If lsOpeCod <> "561603" And lsOpeCod <> "562603" Then
                Me.txtOCompra.Visible = False
                Me.lblOCompra.Visible = False
                Me.lblOrdenCompra.Visible = False
            End If
            Me.txtOCompra.Enabled = False
        End If
    Else
        Me.txtOCompra.Enabled = False
        Me.txtAlmacen.Enabled = False
        Me.FlexDoc.lbEditarFlex = False
        Me.FlexDetalle.lbEditarFlex = False
        Me.FlexSerie.lbEditarFlex = False
        Me.cmdAgregar.Visible = False
        Me.cmdEliminar.Visible = False
        
        If lsOpeCod = gnAlmaIngXComprasConfirma Then
            Me.cmdGrabar.Caption = "&Confirma"
        ElseIf lsOpeCod = gnAlmaMantXIngreso Then
            Me.cmdGrabar.Caption = "&Grabar"
        Else
            Me.cmdGrabar.Caption = "&Rechaza"
        End If
    End If
    
    If lsOpeCod = gnAlmaIngXOtrosMotivos Or lsOpeCod = gnAlmaIngXEmbargo Or lsOpeCod = gnAlmaIngXDacionPago Or lsOpeCod = gnAlmaIngXAdjudicacion Then
        Me.FlexDetalle.lbEditarFlex = True
        Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X"
        Me.fraMoneda.Enabled = True
    ElseIf lsOpeCod = gnAlmaMantXIngreso Then
        Me.FlexDetalle.lbEditarFlex = True
        Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X"
        Me.fraMoneda.Enabled = True
        Me.cmdAgregar.Visible = True
        Me.cmdEliminar.Visible = True
    ElseIf lsOpeCod = gnAlmaIngXCompras Then 'EJVG20111110
        Me.FlexDetalle.lbEditarFlex = False
        Me.txtAlmacen.Enabled = False
        Me.txtProveedor.Enabled = False
    ElseIf lsOpeCod = gnAlmaIngXComprasConfirma Then
        Me.txtProveedor.Enabled = False
    End If
    
    'If lsOpeCod = gnAlmaIngXEmbargo Then
    '    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienEmbargado)
    'ElseIf lsOpeCod = gnAlmaIngXDacionPago Then
    '    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnAlmaIngXDacionPago)
    'ElseIf lsOpeCod = gnAlmaIngXAdjudicacion Then
    '    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienAdjudicado)
    'Else
        Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "")
    'End If
    
    If lbMantenimiento Then
        Me.FlexSerie.lbEditarFlex = True
    End If
    
    
    If gsopecod = "591103" Then '*** PEAC 20110712 'anps
         Me.FlexDetalle.lbEditarFlex = False
         Me.txtAlmacen.Enabled = False
         Me.txtProveedor.Enabled = False
         Me.cmdAgregar.Visible = False
         Me.cmdEliminar.Visible = False
         Me.ChkNIngresoLibre.Visible = False
         Me.FraNIngresoLibre.Visible = False
         Me.lblNIngreso.Visible = False
    End If
    
    If lsOpeCod = 591101 Then
    cmdAgregar.Visible = False
    cmdEliminar.Visible = False
    End If
    If lsOpeCod = 591102 Then 'anps
         Me.ChkNIngresoLibre.Visible = False
         Me.FraNIngresoLibre.Visible = False
         Me.lblNIngreso.Visible = False
    End If
    
    
    'If lsOpeCod = gnAlmaIngXCompras Then Me.txtOCompra.rs = oDoc.GetOrdenesCompra("16")
    If lsOpeCod = gnAlmaIngXCompras Then Me.txtOCompra.rs = oDoc.listarOrdenesCompra(16, gsCodAge) 'EJVG20111110
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not lbGrabar Then
        If MsgBox("Desea Salir sin grabar? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.cmdGrabar.Enabled Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub mskFecha_LostFocus()
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Debe ingresar una fecha correcta", vbInformation, "Aviso"
        mskFecha_GotFocus
        mskFecha.SetFocus
    End If
End Sub

Private Sub optDolares_Click()
    If optDolares.value Then
        Me.fraCambio.Visible = True
    Else
        Me.fraCambio.Visible = False
    End If
End Sub

Private Sub optSoles_Click()
    If optDolares.value Then
        Me.fraCambio.Visible = True
    Else
        Me.fraCambio.Visible = False
    End If
End Sub

Private Sub txtAlmacen_EmiteDatos()
    Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
End Sub

Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 300
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtNotaIng_EmiteDatos()
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim oOpe As DOperaciones
   Set oOpe = New DOperaciones
   Dim lnItemAnt As String
   Dim I As Integer
   Dim lnMovNroAnt As Long
   
   Me.lblNotaIngG.Caption = Left(txtNotaIng.psDescripcion, 50)
   
    Me.lblTotalIGVDet.Caption = Format(0, "#,##0.00")
    Me.lblTotalGDet.Caption = Format(0, "#,##0.00")
    
    If Me.lblNotaIngG.Caption <> "" Then
        If Right(txtNotaIng.psDescripcion, 1) = gcMNDig Then
            Me.optSoles.value = True
            Me.optDolares.value = False
        Else
            Me.optSoles.value = False
            Me.optDolares.value = True
        End If
        
        If InStr(1, Me.lblNotaIngG.Caption, "]") Then
        
        End If
        
        lnMovNroG = Mid(txtNotaIng.psDescripcion, InStr(1, txtNotaIng.psDescripcion, "[") + 1, InStr(1, txtNotaIng.psDescripcion, "]") - InStr(1, txtNotaIng.psDescripcion, "[") - 1)
        If lbReporte Or lbMantenimiento Or lbExtorno Then
            Set rs = oOpe.GetDetNotaIngresoReporte(Me.txtNotaIng.Text, "20", "42")
        Else
            If gsopecod = "591103" Then
                'Set rs = oOpe.GetDetGuiaRemisionTransf(Me.txtNotaIng.Text)
                Set rs = oOpe.GetDetGuiaRemisionTransf(Me.txtNotaIng.Text, Right(gsCodAge, 2)) 'EJVG20140321
            Else
                Set rs = oOpe.GetDetNotaIngreso(Me.txtNotaIng.Text)
            End If
        End If
        
        Me.mskFecha.Text = Format(rs!dDocFecha, gsFormatoFechaView)
        Me.txtProveedor.Text = rs!cPersCod
        Me.lblProveedorNombre.Caption = rs!cPersNombre
        Me.txtAlmacen.Text = rs!nMovBsOrden
        Me.lblAlmacenG.Caption = rs!cAlmDescripcion
        
        If gsopecod = "591103" Then
            Me.txtComentario.Text = "INGRESO POR TRANSFERENCIA SEGUN GUIA/R " & Me.txtNotaIng.Text
        Else
            Me.txtComentario.Text = rs!cMovDesc
        End If
        
        FlexDetalle.Clear
        FlexDetalle.FormaCabecera
        FlexDetalle.Rows = 2
        
        lnItemAnt = 0
        
        FlexSerie.Clear
        FlexSerie.Rows = 2
        FlexSerie.FormaCabecera
        
        lnMovNroAnt = rs!nMovNro
        lnMovNroOPG = lnMovNroAnt
        While Not rs.EOF
            If lnItemAnt <> rs!nMovItem Then
                Me.FlexDetalle.AdicionaFila
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) = rs!cBSCod
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 2) = rs!cBSDescripcion
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3) = rs!nMovCant
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = Format(rs!nMovImporte / IIf(rs!nMovCant = 0, 1, rs!nMovCant), "#,##0.00")
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = rs!nMovOtroImporte
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) = rs!nMovImporte
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 7) = rs!cCtaContCod
            End If
            
            If InStr(1, rs!cBSDescripcion, "[S]") <> 0 Then
                Me.FlexSerie.AdicionaFila
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 0) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 1) = rs!cSerie & ""
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 2) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 3) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 4) = rs!nIGV
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 5) = rs!nValor
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 7) = rs!SerieFisica
            End If
            
            lnItemAnt = rs!nMovItem
            rs.MoveNext
        Wend
        
        FlexDoc.Clear
        FlexDoc.Rows = 2
        FlexDoc.FormaCabecera
        
        Set rs = oOpe.GetOpeDoc(lnMovNroAnt, 0)
        
        While Not rs.EOF
            If rs!nDocTpo = 33 Then
                Me.txtOCompra.Text = rs!cDocNro
                Me.lblOCompra.Caption = Me.lblNotaIngG
            ElseIf rs!nDocTpo = 42 Then
                Me.lblTit.Caption = "Nota de Ingreso : " & rs!cDocNro
            Else
                Me.FlexDoc.AdicionaFila
                Me.FlexDoc.TextMatrix(FlexDoc.Rows - 1, 1) = "1"
                FlexDoc.TextMatrix(FlexDoc.Rows - 1, 2) = rs!Descrip
                FlexDoc.TextMatrix(FlexDoc.Rows - 1, 3) = Format(rs!dDocFecha, gsFormatoFechaView)
                FlexDoc.TextMatrix(FlexDoc.Rows - 1, 4) = Left(rs!cDocNro, 3)
                FlexDoc.TextMatrix(FlexDoc.Rows - 1, 5) = Right(rs!cDocNro, Len(rs!cDocNro) - 4)
            End If
            
            rs.MoveNext
        Wend
        FlexDetalle_RowColChange
    End If
End Sub

Private Sub txtOCompra_EmiteDatos()
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim I As Integer
    Dim lbCorrelativo As Boolean
    Dim lnCorrelativoIni As Long
    Dim oAlmacen As DLogAlmacen
    Dim lb
    Set oAlmacen = New DLogAlmacen
    
    Dim nIGVAcum As Currency
    Dim nTotAcum   As Currency
    'anps desactivo check
    ChkNIngresoLibre.Enabled = False
    '  FIN ANPS
    Me.lblOCompra.Caption = Trim(Left(Me.txtOCompra.psDescripcion, 65))
    
    Me.lblTotalIGVDet.Caption = Format(0, "#,##0.00")
    Me.lblTotalGDet.Caption = Format(0, "#,##0.00")
    
    If Me.txtOCompra.psDescripcion <> "" Then
        If Right(Me.txtOCompra.psDescripcion, 1) = gcMNDig Then
            Me.optSoles.value = True
            Me.optDolares.value = False
        Else
            Me.optSoles.value = True
            Me.optDolares.value = False
        End If
        
        'Set rs = oOpe.GetDetOC(Me.txtOCompra.Text)
        Set rs = oOpe.GetDetalleOCxAgencia(Me.txtOCompra.Text, gsCodAge) 'EJVG20111110
        Me.FlexSerie.Rows = 2
        Me.FlexSerie.Clear
        Me.FlexSerie.FormaCabecera
        
        lnMovNroOPG = rs!nMovNro
        
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        Dim oProv As DLogProveedor
        Set oProv = New DLogProveedor
        Dim lsPersCod  As String
        
'        lsPersCod = rs!cPersCod
'        Set rs1 = oProv.GetProveedorRUC(lsPersCod)
'        txtProveedor.Text = rs1!cPersIDnro
        
        Me.txtProveedor.Text = rs!cPersCod
        Me.lblProveedorNombre.Caption = rs!cPersNombre
        Me.txtComentario.Text = rs!cMovDesc
        FlexDetalle.Clear
        FlexDetalle.FormaCabecera
        FlexDetalle.Rows = 2
        While Not rs.EOF
            Me.FlexDetalle.AdicionaFila
            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) = rs!cBSCod
            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 2) = rs!cBSDescripcion
            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3) = rs!nMovCant
            If gbBitIGVCredFiscal Then
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = Format((rs!nMovImporte / (1 + gnIGV)) / rs!nMovCantT, "#0.00")
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = Format((rs!nMovCant * ((rs!nMovImporte / (1 + gnIGV)) / rs!nMovCantT)) * gnIGV, "#0.00")
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) = Format(rs!nMovCant * ((rs!nMovImporte / (1 + gnIGV)) / rs!nMovCantT), "#0.00")
            Else
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 4) = Format(rs!nMovImporte / rs!nMovCantT, "#0.00")
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) = Format((rs!nMovCant * (rs!nMovImporte / rs!nMovCantT)) * gnIGV, "#0.00")
                Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) = Format(rs!nMovCant * (rs!nMovImporte / rs!nMovCantT), "#0.00")
            End If
            Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 7) = rs!cCtaContCod
            
            If InStr(1, rs!cBSDescripcion, "[S]") <> 0 Then
            
            lbCorrelativo = False
            
            If Not lbMantenimiento Then
                lbCorrelativo = oAlmacen.VerfBSCorrela(rs!cBSCod)
                
                If lbCorrelativo Then
                    lnCorrelativoIni = oAlmacen.GetBSCorrelaIni(rs!cBSCod)
                    
                    If GetUltCodVig(rs!cBSCod) > lnCorrelativoIni Then
                        lnCorrelativoIni = GetUltCodVig(rs!cBSCod)
                    End If
                End If
            End If
                
                nIGVAcum = 0
                nTotAcum = 0
                For I = 1 To CInt(rs!nMovCant)
                    If Me.FlexSerie.TextMatrix(1, 3) = "" Then
                        Me.FlexSerie.AdicionaFila
                    Else
                        Me.FlexSerie.AdicionaFila , , True
                    End If
                    Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 0) = FlexDetalle.Rows - 1
                    Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 2) = FlexDetalle.Rows - 1
                    Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 3) = FlexDetalle.Rows - 1
                    'Reparto de IGV y TOTAL
                    Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 4) = Round(Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) / Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3), 2)
                    Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 5) = Round(Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) / Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3), 2)
                    
                    If lbCorrelativo Then
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 1) = Trim(Str(Year(gdFecSis))) & "-" & Format(lnCorrelativoIni, "00000000")
                        lnCorrelativoIni = lnCorrelativoIni + 1
                    End If
                    
                    nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4)
                    nTotAcum = nTotAcum + FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5)
                Next I
                
                If nIGVAcum <> Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5) Then
                    If Not IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4)) Then
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) = "0"
                    Else
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4) - (nIGVAcum - Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 5))
                    End If
                End If
                If nTotAcum <> Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6) Then
                    If Not IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5)) Then
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) = "0"
                    Else
                        FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) = FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5) - (nTotAcum - Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 6))
                    End If
                End If
                
            End If
            
            rs.MoveNext
        Wend
    End If
    
    FlexDetalle_RowColChange
End Sub

Private Sub txtProveedor_EmiteDatos()
    Me.lblProveedorNombre.Caption = txtProveedor.psDescripcion
End Sub

Public Sub Ini(psOpeCod As String, psCaption As String, Optional pbIngreso As Boolean = True, Optional pbConfirma As Boolean = False, Optional pbMantenimiento As Boolean = False, Optional pbExtorno As Boolean = False)
    lsOpeCod = psOpeCod
    lbIngreso = pbIngreso
    lbConfirma = pbConfirma
    lsCaptionG = psCaption
    lbReporte = False
    lbMantenimiento = pbMantenimiento
    lbExtorno = pbExtorno
    
    Me.Show 1
End Sub

Private Function Valida() As Boolean
    Dim I As Integer
    Dim J As Integer
    Dim lnContador As Integer
    Dim oSal As DLogAlmacen
    Set oSal = New DLogAlmacen
    
    
    If Not lbIngreso Then
        If Me.txtNotaIng.Text = "" Then
            MsgBox "Debe ingresar una Nota de Ingreso valida.", vbInformation, "Aviso"
            Me.txtNotaIng.SetFocus
            Valida = False
            Exit Function
        End If
    
        If oSal.CierreMesLogistica(CDate(Me.mskFecha.Text), txtAlmacen.Text) Then
            MsgBox "No se puede modificar el documento, la fecha que se desea ingresar es una fecha anterior a un cierre de almacen, lo que modificaria los reportes de cierre.", vbInformation, "Aviso"
            mskFecha.SetFocus
            Valida = False
            Exit Function
        End If
    
    End If
    
    If Me.txtAlmacen.Text = "" Then
        MsgBox "Debe ingresar un almacen valido.", vbInformation, "Aviso"
        Me.txtAlmacen.SetFocus
        Valida = False
        Exit Function
    End If
    
    If lsOpeCod = "562403" Or lsOpeCod = "561403" Then
        If Me.txtOCompra.Text = "" Then
            MsgBox "Debe ingresar una orden de compra valida.", vbInformation, "Aviso"
            Me.txtOCompra.SetFocus
            Valida = False
            Exit Function
        End If
    End If
    
    For I = 1 To Me.FlexDoc.Rows - 1
        If Me.FlexDoc.TextMatrix(I, 1) <> "" Then
            If Not IsDate(Me.FlexDoc.TextMatrix(I, 3)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
                FlexDoc.col = 3
                FlexDoc.row = I
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            'LARI20200311: ACTUALMENTE LAS SERIES DE LOS COMPROBANTES SON ALFANÚMERICOS SEGÚN CAMBIO PARA FACTURACIÓN ELECTRÓNICA
            'ElseIf Not IsNumeric(Me.FlexDoc.TextMatrix(I, 4)) Or Me.FlexDoc.TextMatrix(I, 4) = 0 Then
            ElseIf Me.FlexDoc.TextMatrix(I, 4) = "" Then
                MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
                FlexDoc.col = 4
                FlexDoc.row = I
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDoc.TextMatrix(I, 5)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
                FlexDoc.col = 5
                FlexDoc.row = I
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            ElseIf Val(Me.FlexDoc.TextMatrix(I, 5)) = 0 Then
                MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
                FlexDoc.col = 5
                FlexDoc.row = I
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            End If
        End If
    Next I
    
    If Me.FlexDetalle.TextMatrix(1, 1) = "" Then
        MsgBox "Debe ingresar por lo menos un producto.", vbInformation, "Aviso"
        'Me.cmdAgregar.SetFocus
        Valida = False
        Exit Function
    End If
    
    For I = 1 To Me.FlexDetalle.Rows - 1
        If Me.FlexDetalle.TextMatrix(I, 1) = "" Then
            MsgBox "Debe ingresar un bien valido para el registro " & I & " .", vbInformation, "Aviso"
            FlexDetalle.col = 1
            FlexDetalle.row = I
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf InStr(1, Me.FlexDetalle.TextMatrix(I, 2), "[S]") <> 0 Then
            lnContador = 0
            
            For J = 1 To CInt(Me.FlexSerie.Rows - 1)
                If FlexSerie.TextMatrix(J, 2) = Me.FlexDetalle.TextMatrix(I, 0) And FlexSerie.TextMatrix(J, 1) <> "" Then
                    lnContador = lnContador + 1
                End If
            Next J
        
            If lnContador <> CInt(Me.FlexDetalle.TextMatrix(I, 3)) Then
                MsgBox "Debe ingresar una numeros serie valida para el registro " & I & " .", vbInformation, "Aviso"
                FlexDetalle.col = 1
                FlexDetalle.row = I
                FlexSerie.row = lnContador + 1
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(I, 3)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
                FlexDetalle.col = 3
                FlexDetalle.row = I
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(I, 4)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
                FlexDetalle.col = 4
                FlexDetalle.row = I
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            End If
            
            If Not lbMantenimiento And lsOpeCod <> gnAlmaIngXComprasConfirma Then
                For J = 1 To CInt(Me.FlexSerie.Rows - 1)
                    If FlexSerie.TextMatrix(J, 3) = Me.FlexDetalle.TextMatrix(I, 0) Then
                        If (Not lbIngreso And Not VerfBSSerieMov(Me.FlexDetalle.TextMatrix(I, 1), FlexSerie.TextMatrix(J, 1), lnMovNroG)) Or (VerfBSSerie(Me.FlexDetalle.TextMatrix(I, 1), FlexSerie.TextMatrix(J, 1), "0") And lbIngreso) Then
                            MsgBox "El bien ya fue ingresado o no ha sido descargado de almacen, para el registro " & I & " .", vbInformation, "Aviso"
                            FlexSerie.col = 1
                            FlexSerie.row = J
                            FlexDetalle.row = CInt(FlexSerie.TextMatrix(FlexSerie.row, 3))
                            FlexDetalle_RowColChange
                            Me.FlexSerie.SetFocus
                            Valida = False
                            Exit Function
                        End If
                    End If
                Next J
            End If
        
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(I, 3)) Then
            MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
            FlexDetalle.col = 3
            FlexDetalle.row = I
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(I, 4)) Then
            MsgBox "Debe ingresar un valor valido para el registro " & I & " .", vbInformation, "Aviso"
            FlexDetalle.col = 4
            FlexDetalle.row = I
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(I, 6)) Then
            MsgBox "No se ha definido Cta Contable para el registro (Defina una cuenta contable para este producto) " & I & " .", vbInformation, "Aviso"
            FlexDetalle.col = 4
            FlexDetalle.row = I
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        End If
    Next I
    
    If txtComentario.Text = "" Then
        Valida = False
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        txtComentario.SetFocus
        Exit Function
    End If
    
    Valida = True
End Function

Public Sub InicioRep(psDocNro As String)
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    lbReporte = True
    lsOpeCod = gnAlmaIngXComprasConfirma
    lbMantenimiento = False
    lbExtorno = False
    Me.txtNotaIng.rs = oDoc.GetNotaIngresoReporte("20", gnAlmarReporteMovNotIng, , psDocNro)
    Me.cmdAgregar.Visible = False
    Me.cmdEliminar.Visible = False
    Me.FlexDetalle.lbEditarFlex = False
    Me.FlexDoc.lbEditarFlex = False
    Me.FlexSerie.lbEditarFlex = False
    Me.cmdGrabar.Visible = False
    Me.cmdCancelar.Visible = False
    Me.txtProveedor.Enabled = False
    Me.txtAlmacen.Enabled = False
    Me.cmdImprimir.Visible = True
    Me.cmdImprimir.Enabled = True
    Me.Show 1
End Sub

Private Function GetUltCodVig(psBSCod As String, Optional pnItem As Long = -1) As Long
    Dim lnI As Long
    Dim lnJ As Long
    Dim lnValor As Long
    
    lnValor = 0
    
    If pnItem = -1 Then
        For lnI = 1 To Me.FlexDetalle.Rows - 1
            If Me.FlexDetalle.TextMatrix(lnI, 1) = psBSCod Then
                For lnJ = 1 To Me.FlexSerie.Rows - 1
                    If Me.FlexDetalle.TextMatrix(lnI, 0) = FlexSerie.TextMatrix(lnJ, 3) Then
                       If FlexSerie.TextMatrix(lnJ, 1) <> "" Then
                            If CLng(Mid(FlexSerie.TextMatrix(lnJ, 1), 6, 8)) > lnValor Then
                               lnValor = CLng(Mid(FlexSerie.TextMatrix(lnJ, 1), 6, 8))
                            End If
                        Else
                            lnValor = 0
                        End If
                    End If
                Next lnJ
            End If
        Next lnI
    Else
        For lnI = 1 To Me.FlexDetalle.Rows - 1
            If Me.FlexDetalle.TextMatrix(lnI, 1) = psBSCod Then
                For lnJ = 1 To Me.FlexSerie.Rows - 1
                    If Me.FlexDetalle.TextMatrix(lnI, 0) = FlexSerie.TextMatrix(lnJ, 3) Then
                       If FlexSerie.TextMatrix(lnJ, 1) <> "" Then
                            If CLng(Mid(FlexSerie.TextMatrix(lnJ, 1), 6, 8)) > lnValor And FlexSerie.TextMatrix(lnJ, 3) <> pnItem Then
                               lnValor = CLng(Mid(FlexSerie.TextMatrix(lnJ, 1), 6, 8))
                            End If
                        Else
                            lnValor = 0
                        End If
                    End If
                Next lnJ
            End If
        Next lnI
    End If
    
    GetUltCodVig = lnValor + 1
End Function

Private Sub SumaTotalDetalle(pnRow As Long)
    Dim nIGVAcum As Currency
    Dim nTotAcum  As Currency
    Dim lnI As Long
    
    nIGVAcum = 0
    nTotAcum = 0
    
    For lnI = 1 To Me.FlexSerie.Rows - 1
        If IsNumeric(FlexSerie.TextMatrix(lnI, 3)) Then
            If FlexSerie.TextMatrix(lnI, 3) = pnRow Then
                nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(lnI, 4)
                nTotAcum = nTotAcum + FlexSerie.TextMatrix(lnI, 5)
            End If
        End If
    Next lnI
                
                
    Me.lblTotalIGVDet.Caption = Format(nIGVAcum, "#,##0.00")
    Me.lblTotalGDet.Caption = Format(nTotAcum, "#,##0.00")
End Sub

