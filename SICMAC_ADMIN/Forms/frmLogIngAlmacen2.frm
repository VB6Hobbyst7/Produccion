VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogIngAlmacen2 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   225
   ClientTop       =   1455
   ClientWidth     =   11475
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
   ForeColor       =   &H80000010&
   Icon            =   "frmLogIngAlmacen2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      TabIndex        =   40
      Top             =   1800
      Width           =   4635
      Begin VB.TextBox txtComentario 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   360
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
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
         Left            =   135
         TabIndex        =   49
         Top             =   150
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      TabIndex        =   36
      Top             =   1800
      Width           =   6615
      Begin VB.CommandButton cmdAgregarDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Picture         =   "frmLogIngAlmacen2.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   180
         Width           =   825
      End
      Begin VB.CommandButton cmdEliminarDoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Picture         =   "frmLogIngAlmacen2.frx":114C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   600
         Width           =   825
      End
      Begin Sicmact.FlexEdit FlexDoc 
         Height          =   945
         Left            =   930
         TabIndex        =   39
         Top             =   180
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   1667
         Cols0           =   6
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-OK-Documento-Fecha-Serie-Numero"
         EncabezadosAnchos=   "290-400-2000-800-600-1200"
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
   End
   Begin VB.Frame fraProveedor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   11295
      Begin Sicmact.TxtBuscar txtProveedor 
         Height          =   330
         Left            =   1020
         TabIndex        =   31
         Top             =   0
         Width           =   2580
         _ExtentX        =   4551
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
         Left            =   3600
         TabIndex        =   35
         Top             =   0
         Width           =   5430
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   0
         TabIndex        =   34
         Top             =   60
         Width           =   885
      End
      Begin VB.Label lblDNI 
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
         Left            =   9960
         TabIndex        =   33
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
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
         Left            =   9240
         TabIndex        =   32
         Top             =   60
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9060
      TabIndex        =   27
      Top             =   -60
      Width           =   2355
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   540
      End
   End
   Begin Sicmact.TxtBuscar txtAlmacen 
      Height          =   330
      Left            =   1140
      TabIndex        =   8
      Top             =   1080
      Width           =   795
      _ExtentX        =   1402
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
   Begin VB.Frame framCont 
      Height          =   3495
      Left            =   60
      TabIndex        =   4
      Top             =   2910
      Width           =   11355
      Begin VB.CommandButton cmdEliminar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   990
         Picture         =   "frmLogIngAlmacen2.frx":198E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1980
         Width           =   840
      End
      Begin VB.CommandButton cmdAgregar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   120
         Picture         =   "frmLogIngAlmacen2.frx":21D0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1980
         Width           =   840
      End
      Begin Sicmact.FlexEdit FlexDetalle 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   2990
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Descripción-Cantidad-Precio Unit-IGV-Total-CtaCnt-ProvCont-UltSubasta-Comentario"
         EncabezadosAnchos=   "300-1200-2500-700-800-800-1000-0-900-900-2000"
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
         ColumnasAEditar =   "X-1-X-3-X-5-6-X-8-9-10"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-R-R-C-R-R-L"
         FormatosEdit    =   "0-0-0-2-2-2-2-0-2-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FlexSerie 
         Height          =   1140
         Left            =   7620
         TabIndex        =   17
         Top             =   1940
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   2011
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Serie-id-idx-IGV-Valor-Val"
         EncabezadosAnchos=   "300-1200-0-0-900-900-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-R-R-C"
         FormatosEdit    =   "0-0-0-0-2-2-0"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1600
         Left            =   4560
         TabIndex        =   42
         Top             =   1830
         Width           =   3015
         Begin VB.Line Line1 
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            X1              =   120
            X2              =   2820
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL ...."
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
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   48
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total items"
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
            Left            =   180
            TabIndex        =   47
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total IGV"
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
            Index           =   0
            Left            =   180
            TabIndex        =   46
            Top             =   420
            Width           =   825
         End
         Begin VB.Label lblTotalIGV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   1260
            TabIndex        =   45
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label lblTotalTot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   1260
            TabIndex        =   44
            Top             =   1260
            Width           =   1440
         End
         Begin VB.Label lblTotalG 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   1260
            TabIndex        =   43
            Top             =   720
            Width           =   1440
         End
         Begin VB.Shape shpMarco 
            BackColor       =   &H00DDFFFE&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   330
            Index           =   0
            Left            =   1200
            Top             =   285
            Width           =   1635
         End
         Begin VB.Shape shpMarco 
            BackColor       =   &H00DDFFFE&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   330
            Index           =   1
            Left            =   1200
            Top             =   645
            Width           =   1635
         End
         Begin VB.Shape shpMarco 
            BackColor       =   &H00DDFFFE&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FF8080&
            Height          =   330
            Index           =   2
            Left            =   120
            Top             =   1185
            Width           =   2715
         End
      End
      Begin VB.Label lblTotalIGVDet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFE&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   9180
         TabIndex        =   25
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblTotalGDet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DDFFFE&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   10140
         TabIndex        =   26
         Top             =   3120
         Width           =   795
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00DDFFFE&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   9120
         Top             =   3075
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00DDFFFE&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   10080
         Top             =   3075
         Width           =   915
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   7620
         Top             =   3060
         Width           =   3615
      End
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
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Top             =   6480
      Width           =   1100
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
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
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
      Height          =   375
      Left            =   9060
      TabIndex        =   1
      Top             =   6480
      Width           =   1155
   End
   Begin Sicmact.TxtBuscar txtOCompra 
      Height          =   330
      Left            =   4380
      TabIndex        =   11
      Top             =   1080
      Width           =   2040
      _ExtentX        =   3598
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6780
      TabIndex        =   14
      Top             =   -60
      Width           =   2175
      Begin VB.OptionButton optDolares 
         Caption         =   "&Dolares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   285
         Width           =   945
      End
      Begin VB.OptionButton optSoles 
         Caption         =   "&Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   285
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin Sicmact.TxtBuscar txtNotaIng 
      Height          =   330
      Left            =   1140
      TabIndex        =   18
      Top             =   720
      Width           =   2580
      _ExtentX        =   4551
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3840
      TabIndex        =   20
      Top             =   -60
      Visible         =   0   'False
      Width           =   2850
      Begin VB.Label lblCompraG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2040
         TabIndex        =   24
         Top             =   255
         Width           =   600
      End
      Begin VB.Label lblFijoG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   540
         TabIndex        =   23
         Top             =   255
         Width           =   645
      End
      Begin VB.Label llCompra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compra"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   22
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblFijo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fijo"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Label lblNotaIngG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Top             =   720
      Width           =   7695
   End
   Begin VB.Label lblOrdenCompra 
      AutoSize        =   -1  'True
      Caption         =   "O/C :"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   1140
      Width           =   465
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
      Height          =   315
      Left            =   6420
      TabIndex        =   12
      Top             =   1080
      Width           =   4995
   End
   Begin VB.Label lblAlmacenG 
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
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label lblAlmacen 
      AutoSize        =   -1  'True
      Caption         =   "Almacen"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nota de Ingreso: 2001-00001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DDFFFE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   555
      Left            =   60
      Top             =   30
      Width           =   3675
   End
End
Attribute VB_Name = "frmLogIngAlmacen2"
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
Dim HayAdj As Boolean

Private Sub cmdAgregar_Click()
    If Me.txtProveedor.Text = "" Then
        MsgBox "Debe elegir una  persona responsable.", vbInformation, "Aviso"
        Me.txtProveedor.SetFocus
        Exit Sub
    End If
    
    If Me.FlexDetalle.TextMatrix(1, 1) <> "" Then
        If FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 1) <> "" Or FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 3) <> "" Then
            Me.FlexDetalle.AdicionaFila , , True
        End If
    Else
        Me.FlexDetalle.AdicionaFila
    End If
    Me.FlexSerie.AdicionaFila
    Me.FlexDetalle.SetFocus
End Sub

Private Sub cmdAgregarDoc_Click()
    Me.FlexDoc.AdicionaFila
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    Dim i As Integer
    Dim lnEncontrar As Integer
    Dim lnContador As Integer
    
    If MsgBox("Desea Eliminar la fila, si ha incluido numeros de serie para este producto se perderan.", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

    For i = 1 To CInt(Me.FlexSerie.Rows - 1)
        If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0) Then
            lnContador = lnContador + 1
        End If
    Next i
    
    i = 0
    While lnEncontrar < lnContador
        i = i + 1
        If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0) Then
            Me.FlexSerie.EliminaFila i
            lnEncontrar = lnEncontrar + 1
            i = i - 1
        End If
    Wend
    
    For i = 1 To Me.FlexSerie.Rows - 1
        If IsNumeric(FlexSerie.TextMatrix(i, 3)) Then
            If FlexSerie.TextMatrix(i, 3) > Me.FlexDetalle.Row Then
                FlexSerie.TextMatrix(i, 3) = Trim(Str(CInt(FlexSerie.TextMatrix(i, 3)) - 1))
                FlexSerie.TextMatrix(i, 0) = FlexSerie.TextMatrix(i, 3)
                FlexSerie.TextMatrix(i, 2) = FlexSerie.TextMatrix(i, 3)
            End If
        End If
    Next i

    Me.FlexDetalle.EliminaFila Me.FlexDetalle.Row
End Sub

Private Sub cmdEliminarDoc_Click()
    Me.FlexDoc.EliminaFila Me.FlexDoc.Row
End Sub

Private Sub cmdGrabar_Click()
    If Not Valida Then Exit Sub
    Dim i As Integer
    Dim oALmacen As DMov
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnItem As Integer
    Dim lsBsCod As String
    Dim lsDocNI As String
    Dim ldFechaOC As Date
    Set oALmacen = New DMov
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    Dim lsOpeCodLocal As String
    Dim lsCtaCont As String
    Dim lnContador As Long
    Dim lsAgePagare As String
    Dim oConn As DConecta
    
    Set oConn = New DConecta
    
    If MsgBox("Desea Grabar los cambios Realizados ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
    lbGrabar = True
    
    If lbExtorno Then
        
        oALmacen.BeginTrans
          'Inserta Mov
          lsMovNro = oALmacen.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
          oALmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabRechazado, gMovFlagEliminado
          
          lnMovNro = oALmacen.GetnMovNro(lsMovNro)
          oALmacen.InsertaMovGasto lnMovNro, Me.txtProveedor.Text, ""
          
          If lnMovNroG <> 0 Then  'Para Modificados
             oALmacen.ActualizaMov lnMovNroG, , gMovEstContabNoContable, gMovFlagModificado   'Modificado
             oALmacen.ActualizaMov GetnMovNroRef(lnMovNroG, gnAlmaIngXCompras), , , gMovFlagEliminado  'Cambia el modificado a vigente
             oALmacen.InsertaMovRef lnMovNro, lnMovNroG
             oALmacen.EliminaMovBSSerieparaActualizar lnMovNroG
          End If
          
          If lnMovNroOPG <> 0 And lnMovNroG = 0 Then 'Para Modificados
             oALmacen.InsertaMovRef lnMovNro, lnMovNroOPG
          End If
        oALmacen.CommitTrans
        
        cmdImprimir_Click
        Unload Me
        Exit Sub
    
    ElseIf lbMantenimiento Then
        lsOpeCodLocal = GetOpeMov(lnMovNroG)
        
        oALmacen.BeginTrans
          'Inserta Mov
          lsMovNro = oALmacen.GeneraMovNro(CDate(Me.mskFecha), Right(gsCodAge, 2), gsCodUser)
          oALmacen.InsertaMov lsMovNro, lsOpeCodLocal, Me.txtComentario.Text, 20
          
          lnMovNro = oALmacen.GetnMovNro(lsMovNro)
          oALmacen.InsertaMovGasto lnMovNro, Me.txtProveedor.Text, ""
          
          If lnMovNroG <> 0 Then  'Para Modificados
             oALmacen.ActualizaMov lnMovNroG, , , gMovFlagModificado 'Modificado
             oALmacen.InsertaMovRef lnMovNro, lnMovNroG
             oALmacen.InsertaMovRefAnt lnMovNro, lnMovNroG
             oALmacen.EliminaMovBSSerieparaActualizar lnMovNroG
          End If
          
          If lnMovNroOPG <> 0 And lnMovNroG = 0 Then 'Para Modificados
             oALmacen.InsertaMovRef lnMovNro, lnMovNroOPG
          End If
          
          lsDocNI = Right(Me.lblTit, 13)
          
          'Inserta Documentos
          lsAgePagare = ""
          For i = 1 To Me.FlexDoc.Rows - 1
            If Me.FlexDoc.TextMatrix(i, 1) <> "" Then
                oALmacen.InsertaMovDoc lnMovNro, Trim(Right(Me.FlexDoc.TextMatrix(i, 2), 8)), Trim(Me.FlexDoc.TextMatrix(i, 4)) & "-" & Me.FlexDoc.TextMatrix(i, 5), Format(CDate(Me.FlexDoc.TextMatrix(i, 3)), gsFormatoFecha)
                If Trim(Right(Me.FlexDoc.TextMatrix(i, 2), 5)) Then
                   lsAgePagare = Right(Me.FlexDoc.TextMatrix(i, 4), 2)
                End If
            End If
          Next i
          
          If lsAgePagare = "" Then lsAgePagare = Right(gsCodAge, 2)
          
          oALmacen.InsertaMovDoc lnMovNro, 42, lsDocNI, Format(CDate(Me.mskFecha.Text), gsFormatoFecha)
          
          For i = 1 To Me.FlexDetalle.Rows - 1
            lsBsCod = Me.FlexDetalle.TextMatrix(i, 1)
            oALmacen.InsertaMovBS lnMovNro, i, txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1)
            oALmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 3)
            If Me.optSoles.value Then
                oALmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 7), Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)), "#0.00")
                oALmacen.InsertaMovOtrosItem lnMovNro, i, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5)), "#0.00"), ""
            Else
                oALmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 7), Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * CCur(Me.lblCompraG.Caption), "#0.00")
                oALmacen.InsertaMovOtrosItem lnMovNro, i, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5)), "#0.00"), ""
            End If
          Next i
          
          lnContador = i
         
          If lsOpeCod = gnAlmaIngXAdjudicacion Then
             'Ctas provision
             For i = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), True)
                oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * lnPorProIni)
                oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(i, 5)), ""
                lnContador = lnContador + 1
             Next i
                         
             'Ctas Pendientes para amortizar Creditos
             For i = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oALmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7))
                oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * -1)
                oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(i, 5)), ""
                lnContador = lnContador + 1
             Next i
             'Ctas Pendientes para provicion
             For i = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), False)
                oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(i, 6)) * lnPorProIni) * -1)
                oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(i, 5)), ""
                lnContador = lnContador + 1
             Next i
          ElseIf lsOpeCod = gnAlmaIngXDacionPago Then
             'Ctas provision
             For i = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), True)
                oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6) * lnPorProIni))
                oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(i, 5)), ""
                lnContador = lnContador + 1
             Next i
             'Ctas Pendientes para amortizar Creditos
             For i = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = Replace(oALmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7)), "AG", lsAgePagare)
                oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * -1)
                oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(i, 5)), ""
                lnContador = lnContador + 1
             Next i
             'Ctas Pendientes para provicion
             For i = 1 To Me.FlexDetalle.Rows - 1
                lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), False)
                oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(i, 6)) * lnPorProIni) * -1)
                oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, CCur(Me.FlexDetalle.TextMatrix(i, 5)), ""
                lnContador = lnContador + 1
             Next i
          End If
          
          If FlexSerie.TextMatrix(1, 1) <> "" Then
            For i = 1 To Me.FlexSerie.Rows - 1
              If InStr(1, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 2), "[S]") <> 0 Then
                 oALmacen.InsertaMovBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1), Me.FlexSerie.TextMatrix(i, 4), Me.FlexSerie.TextMatrix(i, 5), lsMovNro
              End If
            Next i
          End If
          
          If Me.optDolares.value Then
            oALmacen.GeneraMovME lnMovNro, lsMovNro
          End If
        oALmacen.CommitTrans
        
        cmdImprimir_Click
        Unload Me
        Exit Sub
    End If
    
    
    If Me.txtOCompra.Text <> "" Then
        ldFechaOC = oOpe.GetFechaDoc(Me.txtOCompra.Text, "33")
    End If
    
    oALmacen.BeginTrans
      'Inserta Mov
      lsMovNro = oALmacen.GeneraMovNro(CDate(Me.mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
      
      If lsOpeCod = gnAlmaIngXCompras Then
         oALmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabNoContable
      ElseIf lsOpeCod = gnAlmaExtornoXIngreso Then 'Rechazo
         oALmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, 21
      ElseIf Left(lsOpeCod, 4) = Left(gnAlmaIngXComprasConfirma, 4) Then
         oALmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, 20
      Else
         oALmacen.InsertaMov lsMovNro, lsOpeCod, Me.txtComentario.Text, gMovEstContabNoContable
      End If
      
      lnMovNro = oALmacen.GetnMovNro(lsMovNro)
      oALmacen.InsertaMovGasto lnMovNro, Me.txtProveedor.Text, ""
      
      If lnMovNroG <> 0 Then  'Para Modificados
         oALmacen.ActualizaMov lnMovNroG, , , gMovFlagModificado 'Modificado
         oALmacen.InsertaMovRef lnMovNro, lnMovNroG
         oALmacen.EliminaMovBSSerieparaActualizar lnMovNroG
      End If
      
      If lnMovNroOPG <> 0 And lnMovNroG = 0 Then 'Para Modificados
         oALmacen.InsertaMovRef lnMovNro, lnMovNroOPG
      End If
      
      If lsOpeCod <> gnAlmaIngXComprasConfirma And Left(lsOpeCod, 4) = Left(gnAlmaIngXCompras, 4) Then  ' Genera documento
         lsDocNI = oALmacen.GeneraDocNro(42, gMonedaExtranjera, Year(gdFecSis))
      Else
         lsDocNI = Right(Me.lblTit, 13)
      End If
      
      'Inserta Documentos
      lsAgePagare = ""
      For i = 1 To Me.FlexDoc.Rows - 1
        If Me.FlexDoc.TextMatrix(i, 1) <> "" Then
            oALmacen.InsertaMovDoc lnMovNro, Trim(Right(Me.FlexDoc.TextMatrix(i, 2), 8)), Trim(Me.FlexDoc.TextMatrix(i, 4)) & "-" & Me.FlexDoc.TextMatrix(i, 5), Format(CDate(Me.FlexDoc.TextMatrix(i, 3)), gsFormatoFecha)
            If Trim(Right(Me.FlexDoc.TextMatrix(i, 2), 5)) Then
                lsAgePagare = Right(Me.FlexDoc.TextMatrix(i, 4), 2)
            End If
        End If
      Next i
      
      If lsAgePagare = "" Then lsAgePagare = Right(gsCodAge, 2)
      
      oALmacen.InsertaMovDoc lnMovNro, 42, lsDocNI, Format(CDate(Me.mskFecha.Text), gsFormatoFecha)
      
      If lsOpeCod = "562403" Or lsOpeCod = "561403" Then
         oALmacen.InsertaMovDoc lnMovNro, 33, Me.txtOCompra.Text, Format(ldFechaOC, gsFormatoFecha)
      End If
      
      'ESÑM ------------------------------------------------------------------------
      If lsOpeCod = gnAlmaIngXAdjudicacion Then
         If Not oConn.AbreConexion Then
            MsgBox "No se puede grabar datos del Bien Adjudicado..." + Space(10), vbInformation
         End If
      End If
      '-----------------------------------------------------------------------------
      For i = 1 To Me.FlexDetalle.Rows - 1
          lsBsCod = Me.FlexDetalle.TextMatrix(i, 1)
          oALmacen.InsertaMovBS lnMovNro, i, txtAlmacen.Text, Me.FlexDetalle.TextMatrix(i, 1)
          oALmacen.InsertaMovCant lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 3)
          oALmacen.InsertaMovCta lnMovNro, i, Me.FlexDetalle.TextMatrix(i, 7), Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)))
          oALmacen.InsertaMovOtrosItem lnMovNro, i, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
          
          'ESÑM --------------------------------------------------------------------
          oConn.Ejecutar "INSERT INTO BSAdjudicados (nMovNro,nMovItem,cPersCod,cBSCod,nProvCont,nUltimaSub,cComentario) " & _
                         " VALUES (" & lnMovNro & "," & i & ",'" & txtProveedor.Text & "','" & lsBsCod & "'," & VNumero(Me.FlexDetalle.TextMatrix(i, 8)) & "," & VNumero(Me.FlexDetalle.TextMatrix(i, 9)) & ",'" & Me.FlexDetalle.TextMatrix(i, 10) & "') "
          '-------------------------------------------------------------------------
      Next i
      oConn.CierraConexion
      
      lnContador = i
      
      If lsOpeCod = gnAlmaIngXAdjudicacion Then
         'Ctas provision
         For i = 1 To Me.FlexDetalle.Rows - 1
             lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), True)
             oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * lnPorProIni)
             oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
             lnContador = lnContador + 1
         Next i
         'Ctas Pendientes para amortizar Creditos
         For i = 1 To Me.FlexDetalle.Rows - 1
             lsCtaCont = oALmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7))
             oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * -1)
             oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
             lnContador = lnContador + 1
         Next i
         'Ctas Pendientes para provicion
         For i = 1 To Me.FlexDetalle.Rows - 1
             lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), False)
             oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(i, 6)) * lnPorProIni) * -1)
             oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
             lnContador = lnContador + 1
         Next i
      ElseIf lsOpeCod = gnAlmaIngXDacionPago Then
         'Ctas provision
         For i = 1 To Me.FlexDetalle.Rows - 1
             lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), True)
             oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6) * lnPorProIni))
             oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
             lnContador = lnContador + 1
         Next i
         'Ctas Pendientes para amortizar Creditos
         For i = 1 To Me.FlexDetalle.Rows - 1
             lsCtaCont = Replace(oALmacen.GetOpeCtaCta(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7)), "AG", lsAgePagare)
             oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format(CCur(Me.FlexDetalle.TextMatrix(i, 6)) * -1)
             oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
             lnContador = lnContador + 1
         Next i
         'Ctas Pendientes para provicion
         For i = 1 To Me.FlexDetalle.Rows - 1
             lsCtaCont = oALmacen.GetOpeCtaCtaOtro(lsOpeCod, "", Me.FlexDetalle.TextMatrix(i, 7), False)
             oALmacen.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Format((CCur(Me.FlexDetalle.TextMatrix(i, 6)) * lnPorProIni) * -1)
             oALmacen.InsertaMovOtrosItem lnMovNro, lnContador, gcCtaIGV, Format(CCur(Me.FlexDetalle.TextMatrix(i, 5))), ""
             lnContador = lnContador + 1
         Next i
      End If

      If FlexSerie.TextMatrix(1, 1) <> "" Then
        For i = 1 To Me.FlexSerie.Rows - 1
          If InStr(1, Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 2), "[S]") <> 0 Then
             oALmacen.InsertaMovBSSerie lnMovNro, CInt(Me.FlexSerie.TextMatrix(i, 3)), Me.FlexDetalle.TextMatrix(CInt(Me.FlexSerie.TextMatrix(i, 3)), 1), Me.FlexSerie.TextMatrix(i, 1), Me.FlexSerie.TextMatrix(i, 4), Me.FlexSerie.TextMatrix(i, 5), lsMovNro
          End If
        Next i
      End If
      
      If Me.optDolares.value Then
         oALmacen.GeneraMovME lnMovNro, lsMovNro
      End If
    oALmacen.CommitTrans
    cmdImprimir_Click
    
    Unload Me
End Sub

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
    Dim i As Long, j As Long
    Dim lsTMP As String * 10
    Dim lsTMC As String * 8
    Dim lsNumeros As String
    
    Dim mStr() As String, nLineas As Integer, nLinea As Integer
        
    lsCadena = ""
    'lsCadena = lsCadena & CabeceraPagina(lblTit.Caption, lnPagina, lnItem,gsNomAge , gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CabeceraPagina(lblTit.Caption, lnPagina, lnItem, UCase(lblAlmacenG.Caption), gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTextoCadena(Me.txtComentario.Text, 105, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JIZQ(UCase(Me.Label3.Caption), 18) & "  " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "MOTIVO DE INGRESO : " & Me.Caption & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "DOCUMENTOS : --------------------------------------------" & oImpresora.gPrnSaltoLinea
   
    For j = 1 To Me.FlexDoc.Rows - 1
        If Me.FlexDoc.TextMatrix(j, 1) <> "" Then
             lsDocNom = Left(Me.FlexDoc.TextMatrix(j, 2), 20)
             RSet lsDocNum = Trim(Me.FlexDoc.TextMatrix(j, 4)) & "-" & Me.FlexDoc.TextMatrix(j, 5)
             RSet lsDocFec = Format(CDate(Me.FlexDoc.TextMatrix(j, 3)), gsFormatoFechaView)
             lsCadena = lsCadena & lsDocNom & lsDocNum & lsDocFec & oImpresora.gPrnSaltoLinea
             lnItem = lnItem + 1
             If lnItem > 35 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                'lsCadena = lsCadena & CabeceraPagina(lblTit.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & CabeceraPagina(lblTit.Caption, lnPagina, lnItem, UCase(lblAlmacenG.Caption), gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & JustificaTextoCadena(Me.txtComentario.Text, 105, 5) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & JIZQ(UCase(Me.Label3.Caption), 18) & "  " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & "MOTIVO DE INGRESO : " & Me.Caption & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & "DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
             End If
        End If
    Next j

    'If lsOpeCod = gnAlmaIngXAdjudicacion Then
    If HayAdj Then
       lsCadena = lsCadena & Encabezado("ITEM;4;CODIGO;9; ;7;DESCRIPCION;18; ;25;CANTIDAD;10; ;1;PRECIO UNIT;15; ;1;VALOR-ADJ;10; ;1;PROV-CONT;10;ULT-SUB;10;", lnItem)
    Else
       lsCadena = lsCadena & Encabezado("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;CANTIDAD;10; ;5;PRECIO;17; ;5;TOTAL;10; ;3;", lnItem)
    End If

    For i = 1 To Me.FlexDetalle.Rows - 1
        lsItem = Format(i, "0000")
        lsCodigo = Me.FlexDetalle.TextMatrix(i, 1)
        lsNombre = Me.FlexDetalle.TextMatrix(i, 2)

        'ESÑM - Solo para adjudicados ------------------------------------------------
        'If lsOpeCod = gnAlmaIngXAdjudicacion Then
        If HayAdj Then
           lsTMP = ""
           lsNumeros = ""
           RSet lsTMC = Format(Me.FlexDetalle.TextMatrix(i, 3), "#,##0.00")
           lsNumeros = lsNumeros + lsTMC
           RSet lsTMP = Format(Me.FlexDetalle.TextMatrix(i, 4), "##,##0.00")
           lsNumeros = lsNumeros + Space(5) + lsTMP
           RSet lsTMP = Format(Me.FlexDetalle.TextMatrix(i, 6), "##,##0.00")
           lsNumeros = lsNumeros + lsTMP
           RSet lsTMP = Format(Me.FlexDetalle.TextMatrix(i, 8), "##,##0.00")
           lsNumeros = lsNumeros + lsTMP
           RSet lsTMP = Format(Me.FlexDetalle.TextMatrix(i, 9), "##,##0.00")
           lsNumeros = lsNumeros + lsTMP
           lsCadena = lsCadena & "  " & lsItem & lsCodigo & "  " & JIZQ(UCase(lsNombre), 40) & " " & lsNumeros & oImpresora.gPrnSaltoLinea
           
           If Len(Trim(Me.FlexDetalle.TextMatrix(i, 10))) > 0 Then
              mStr = AjustaTexto(Me.FlexDetalle.TextMatrix(i, 10), nAncho36)
           Else
              ReDim mStr(1)
              mStr(1) = Me.FlexDetalle.TextMatrix(i, 10)
           End If
           lsCadena = lsCadena & Space(24) & mStr(1) & oImpresora.gPrnSaltoLinea
           nLineas = UBound(mStr)
           For nLinea = 2 To nLineas
               lsCadena = lsCadena & Space(24) & mStr(nLinea) & oImpresora.gPrnSaltoLinea
           Next
       '-----------------------------------------------------------------------------------
       Else
           'ORIGINAL
           lsCantidad = Me.FlexDetalle.TextMatrix(i, 3)
           RSet lsPrecio = Format(Me.FlexDetalle.TextMatrix(i, 4), "#,#00.00")
           RSet lsTotal = Format(Me.FlexDetalle.TextMatrix(i, 6), "#,#00.00")
           lsCadena = lsCadena & "  " & lsItem & lsCodigo & "  " & lsNombre & " " & lsCantidad & "  " & lsPrecio & "  " & lsTotal & oImpresora.gPrnSaltoLinea
        End If
        
        lnItem = lnItem + 1
        
        If InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
            lsItem = ""
            lsCodigo = ""
            lsCadenaSerie = ""
            For j = 1 To Me.FlexSerie.Rows - 1
                If Me.FlexSerie.TextMatrix(j, 3) = Me.FlexDetalle.TextMatrix(i, 0) Then
                    If lsCadenaSerie = "" Then
                        lsCadenaSerie = Me.FlexSerie.TextMatrix(j, 1)
                    Else
                        lsCadenaSerie = lsCadenaSerie & " / " & Me.FlexSerie.TextMatrix(j, 1)
                        lnItem = lnItem + 1
                    End If
                    If j Mod 3 = 0 Then
                        lsCadena = lsCadena & lsItem & lsCodigo & "    " & lsCadenaSerie & oImpresora.gPrnSaltoLinea
                        lsCadenaSerie = ""
                    End If
                End If
            Next j
            lsCadena = lsCadena & lsItem & lsCodigo & "    " & lsCadenaSerie & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
        End If
        
        If lnItem > 44 Then
             lnItem = 0
             lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
             lsCadena = lsCadena & CabeceraPagina(lblTit.Caption, lnPagina, lnItem, gsNomAge, gsEmpresa, CDate(Me.mskFecha.Text), Mid(lsOpeCod, 3, 1)) & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & JIZQ(UCase(Me.Label3.Caption), 18) & "  " & PstaNombre(Me.lblProveedorNombre.Caption) & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & "MOTIVO DE INGRESO : " & Me.Caption & oImpresora.gPrnSaltoLinea
             lsCadena = lsCadena & "DOCUMENTOS : -----------------------------------------" & oImpresora.gPrnSaltoLinea
             For j = 1 To Me.FlexDoc.Rows - 1
                 If Me.FlexDoc.TextMatrix(j, 1) <> "" Then
                      lsDocNom = Left(Me.FlexDoc.TextMatrix(j, 2), 20)
                      RSet lsDocNum = Trim(Me.FlexDoc.TextMatrix(j, 4)) & "-" & Me.FlexDoc.TextMatrix(j, 5)
                      RSet lsDocFec = Format(CDate(Me.FlexDoc.TextMatrix(j, 3)), gsFormatoFechaView)
                      lsCadena = lsCadena & lsDocNom & lsDocNum & lsDocFec & oImpresora.gPrnSaltoLinea
                 End If
             Next j
             lsCadena = lsCadena & Encabezado("ITEM;5;CODIGO;9; ;10;DESCRIPCION;15; ;30;CANTIDAD;10; ;5;PRECIO;17; ;5;TOTAL;10; ;3;", lnItem)
        End If
    Next i
      
    lsCadena = lsCadena & String(121, "=") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "----------------------------          ----------------------------          --------------------------" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "         ALMACEN                              LOGISTICA                              Vo Bo          " & oImpresora.gPrnSaltoLinea
    oPrevio.Show lsCadena, Caption, True, 66
    
End Sub



Private Sub FlexDetalle_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lnCorrelativoIni  As Currency
    Dim lbCorrelativo As Boolean
    Dim lnI As Long
    
    If pnCol = 1 And (lsOpeCod <> gnAlmaIngXCompras And lsOpeCod <> gnAlmaIngXComprasConfirma) Then
        If lsOpeCod = gnAlmaIngXDacionPago Or lsOpeCod = gnAlmaIngXAdjudicacion Then
            Me.FlexDetalle.TextMatrix(Me.FlexDetalle.Row, 4) = Format(oALmacen.GetPrePromedio("1", Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), 1), "#0.00")
        ElseIf lsOpeCod = gnAlmaIngXEmbargo Then
            Me.FlexDetalle.TextMatrix(Me.FlexDetalle.Row, 4) = Format(oALmacen.GetPrePromedio("1", Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), 2), "#0.00")
        Else
            Me.FlexDetalle.TextMatrix(Me.FlexDetalle.Row, 4) = Format(oALmacen.GetPrePromedio("1", Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), 0), "#0.00")
        End If
    End If
    
    If pnCol = 1 Then
        If Me.FlexDetalle.TextMatrix(pnRow, pnCol) = "" Then Exit Sub
        lbCorrelativo = oALmacen.VerfBSCorrela(Me.FlexDetalle.TextMatrix(pnRow, pnCol))
        
        If lbCorrelativo Then
            lnCorrelativoIni = oALmacen.GetBSCorrelaIni(Me.FlexDetalle.TextMatrix(pnRow, pnCol))
            
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
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lnCorrelativoIni  As Currency
    Dim lbCorrelativo As Boolean
    Dim lnI As Long
    Dim lnUltimo As Long
    Dim nIGVAcum As Currency
    Dim nTotAcum As Currency
    
    If pnCol = 6 Or pnCol = 3 Then
        If IsNumeric(FlexDetalle.TextMatrix(pnRow, 6)) And IsNumeric(FlexDetalle.TextMatrix(pnRow, 3)) Then
            If CCur(FlexDetalle.TextMatrix(pnRow, 3)) <> 0 Then
                FlexDetalle.TextMatrix(pnRow, 4) = Format(CCur(FlexDetalle.TextMatrix(pnRow, 6)) / CCur(FlexDetalle.TextMatrix(pnRow, 3)), "0.00")
            End If
            FlexDetalle.TextMatrix(pnRow, 5) = Format(CCur(FlexDetalle.TextMatrix(pnRow, 6) * gnIGV), "0.00")
            
            If InStr(1, Me.FlexDetalle.TextMatrix(pnRow, 2), "[S]") = 0 Then Exit Sub
            nIGVAcum = 0
            nTotAcum = 0
            For lnI = 1 To Me.FlexSerie.Rows - 1
                If IsNumeric(FlexSerie.TextMatrix(lnI, 3)) Then
                    If FlexSerie.TextMatrix(lnI, 3) = pnRow Then
                        FlexSerie.TextMatrix(lnI, 4) = Round(Me.FlexDetalle.TextMatrix(pnRow, 5) / Me.FlexDetalle.TextMatrix(pnRow, 3), 2)
                        FlexSerie.TextMatrix(lnI, 5) = Round(Me.FlexDetalle.TextMatrix(pnRow, 6) / Me.FlexDetalle.TextMatrix(pnRow, 3), 2)
                        nIGVAcum = nIGVAcum + FlexSerie.TextMatrix(lnI, 4)
                        nTotAcum = nTotAcum + FlexSerie.TextMatrix(lnI, 5)
                        lnUltimo = lnI
                    End If
                End If
            Next lnI
            
            If lnUltimo <> 0 Then
                If IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 4)) Then
                    If nIGVAcum <> Me.FlexDetalle.TextMatrix(pnRow, 5) Then
                        FlexSerie.TextMatrix(lnUltimo, 4) = FlexSerie.TextMatrix(lnUltimo, 4) - (nIGVAcum - Me.FlexDetalle.TextMatrix(pnRow, 5))
                    End If
                End If
            
                If IsNumeric(FlexSerie.TextMatrix(FlexSerie.Rows - 1, 5)) Then
                    If nTotAcum <> Me.FlexDetalle.TextMatrix(pnRow, 6) Then
                        FlexSerie.TextMatrix(lnUltimo, 5) = FlexSerie.TextMatrix(lnUltimo, 5) - (nTotAcum - Me.FlexDetalle.TextMatrix(pnRow, 6))
                    End If
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
                FlexSerie.TextMatrix(lnUltimo, 4) = FlexSerie.TextMatrix(lnUltimo, 4) - (nIGVAcum - Me.FlexDetalle.TextMatrix(pnRow, 5))
            End If
        End If
        SumaTotalDetalle pnRow
    End If
End Sub

Private Sub FlexDetalle_RowColChange()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCtaCnt As String
    Dim i As Integer
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
            
    lnTotal = 0
    lnTotalIGV = 0
    For lnI = 1 To Me.FlexDetalle.Rows - 1
        If IsNumeric(FlexDetalle.TextMatrix(lnI, 5)) And IsNumeric(FlexDetalle.TextMatrix(lnI, 6)) Then
            lnTotal = lnTotal + CCur(FlexDetalle.TextMatrix(lnI, 6))
            lnTotalIGV = lnTotalIGV + CCur(FlexDetalle.TextMatrix(lnI, 5))
        End If
    Next lnI
    
    'Me.lblTotalG.Caption = Format(lnTotal, "#,##0.00")
    'Me.lblTotalIGV.Caption = Format(lnTotalIGV, "#,##0.00")
    
    If InStr(1, Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 2), "[S]") <> 0 Or Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1) = "" Then
        lnContador = 0
        If Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 3) = "" Then
            Exit Sub
        End If
        
        '------------------------------------------------------------------------------
        'ESÑM: Se ha puesto comentario a este bloque ----------------------------------
        '------------------------------------------------------------------------------
        For i = 1 To CInt(Me.FlexSerie.Rows - 1)
            If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0) Then
                lnContador = lnContador + 1
                FlexSerie.RowHeight(i) = 285
            Else
                FlexSerie.RowHeight(i) = 0
                'FlexSerie.RowHeight(I) = 285
            End If
        Next i
        '------------------------------------------------------------------------------
        
        SumaTotalDetalle FlexDetalle.Row
        
        If lnContador <> CInt(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 3)) Then
            i = 0
            lnEncontrar = 0
            While lnEncontrar < lnContador
                i = i + 1
                If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 0) Then
                    Me.FlexSerie.EliminaFila i
                    lnEncontrar = lnEncontrar + 1
                    i = i - 1
                End If
            Wend
            
            'For lnItems = 1 To Me.FlexDetalle.Rows - 1
                lnItems = FlexDetalle.Row
                If Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1) = Me.FlexDetalle.TextMatrix(lnItems, 1) And FlexDetalle.Row <> lnItems Then
                    lnContador = CLng(Me.FlexDetalle.TextMatrix(lnItems, 3))
                    lnEncontrar = 0
                    i = 0
                    While lnEncontrar < lnContador
                        i = i + 1
                        If FlexSerie.TextMatrix(i, 3) = Me.FlexDetalle.TextMatrix(lnItems, 0) Then
                            Me.FlexSerie.EliminaFila i
                            lnEncontrar = lnEncontrar + 1
                            i = i - 1
                        End If
                    Wend
                End If
            'Next lnItems
            
            lbCorrelativo = False
            
            If Not lbMantenimiento Then
                
                lbCorrelativo = oALmacen.VerfBSCorrela(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1))
                
                If lbCorrelativo Then
                    lnCorrelativoIni = oALmacen.GetBSCorrelaIni(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1))
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
                   lnItems = FlexDetalle.Row
                   If Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1) = Me.FlexDetalle.TextMatrix(lnItems, 1) Then
                      For i = 1 To CLng(Me.FlexDetalle.TextMatrix(lnItems, 3))
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
                     Next i
                    
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
        For i = 1 To CInt(Me.FlexSerie.Rows - 1)
            FlexSerie.RowHeight(i) = 0
        Next i
        Me.lblTotalIGVDet.Caption = Format(0, "#,##0.00")
        Me.lblTotalGDet.Caption = Format(0, "#,##0.00")
    End If
    
    If lbMantenimiento Or lbExtorno Then
    '    lsCtaCnt = GetCtaCntBS(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), GetOpeMov(lnMovNroG))
    Else
    '    lsCtaCnt = GetCtaCntBS(Me.FlexDetalle.TextMatrix(FlexDetalle.Row, 1), lsOpeCod)
    End If
    FlexDetalle.TextMatrix(FlexDetalle.Row, 7) = lsCtaCnt
    Set oALmacen = Nothing
End Sub

Private Sub FlexDoc_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If pnCol = 4 Then
        FlexDoc.TextMatrix(pnRow, pnCol) = Format(FlexDoc.TextMatrix(pnRow, pnCol), "000")
    ElseIf pnCol = 5 Then
        FlexDoc.TextMatrix(pnRow, pnCol) = Format(FlexDoc.TextMatrix(pnRow, pnCol), "0000000")
    End If
End Sub

Private Sub FlexDoc_RowColChange()
    If FlexDoc.TextMatrix(FlexDoc.Row, 1) = "" And FlexDoc.Col <> 1 Then
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
    Dim odoc As DOperaciones
    Set odoc = New DOperaciones
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oGen As DLogGeneral
    Set oGen = New DLogGeneral
    Dim oMov As DMov
    Set oMov = New DMov
    
    HayAdj = False
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
    
    'Me.lblNotaIng.Visible = Not lbIngreso
    Me.lblNotaIngG.Visible = Not lbIngreso
    Me.txtNotaIng.Visible = Not lbIngreso
    Me.cmdImprimir.Visible = Not lbIngreso
    
    Me.FlexDoc.CargaCombo odoc.GetDocOpe(lsOpeCod, False)
    
    Me.mskFecha = Format(gdFecSis, gsFormatoFechaView)
    
    Me.txtAlmacen.rs = odoc.GetAlmacenes
    Me.txtAlmacen.Text = "1"
    Me.lblAlmacenG.Caption = Me.txtAlmacen.psDescripcion
        
    If lbMantenimiento Or lbExtorno Then
        Me.txtNotaIng.rs = odoc.GetNotaIngreso("20", lsOpeCod)
        txtNotaIng.Visible = True
        Me.txtOCompra.Visible = False
        Me.lblOCompra.Visible = False
        Me.lblOrdenCompra.Visible = False
        'Exit Sub
    Else
        If lsOpeCod = gnAlmaIngXComprasConfirma And Not lbReporte Then Me.txtNotaIng.rs = odoc.GetNotaIngreso("13", lsOpeCod)
    End If
        
    If lbIngreso Then Me.lblTit.Caption = "Nota de Ingreso : " & oMov.GeneraDocNro(42, gMonedaExtranjera, Year(gdFecSis))
    
    If Not lbConfirma Then
        If lbIngreso Then
            If Not (lsOpeCod = gnAlmaIngXCompras) Then
                Me.txtOCompra.Visible = False
                Me.lblOCompra.Visible = False
                Me.lblOrdenCompra.Visible = False
                Me.Label3.Caption = "Persona"
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
        
        'Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X"
        
        'Modificacion  ESÑM ----------------------------------------------
        If lsOpeCod = gnAlmaIngXAdjudicacion Then
           HayAdj = True
           Label3.Caption = "Cliente :"
           Me.FlexDetalle.ColWidth(6) = 1000: Me.FlexDetalle.TextMatrix(0, 6) = "Val. Adjud."
           Me.FlexDetalle.ColWidth(8) = 1000
           Me.FlexDetalle.ColWidth(9) = 1000
           Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X-7-8-9"
        Else
           Me.FlexDetalle.ColWidth(8) = 0
           Me.FlexDetalle.ColWidth(9) = 0
           Me.FlexDetalle.ColWidth(9) = 0
           Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X"
        End If
        Me.fraMoneda.Enabled = True
        
    ElseIf lsOpeCod = gnAlmaMantXIngreso Then
       'Modificacion  ESÑM ----------------------------------------------
        If lsOpeCod = gnAlmaIngXAdjudicacion Then
           HayAdj = True
           Label3.Caption = "Cliente :"
           Me.FlexDetalle.ColWidth(6) = 1000: Me.FlexDetalle.TextMatrix(0, 6) = "Val. Adjud."
           Me.FlexDetalle.ColWidth(8) = 1000
           Me.FlexDetalle.ColWidth(9) = 1000
           Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X-7-8-9"
        Else
           Me.FlexDetalle.lbEditarFlex = True
           Me.FlexDetalle.ColWidth(2) = 5505
           Me.FlexDetalle.ColWidth(3) = 1000
           Me.FlexDetalle.ColWidth(4) = 1000
           Me.FlexDetalle.ColWidth(8) = 0
           Me.FlexDetalle.ColWidth(9) = 0
           Me.FlexDetalle.ColWidth(10) = 0
           Me.FlexDetalle.ColumnasAEditar = "X-1-X-3-X-5-6-X"
           Me.fraMoneda.Enabled = True
           Me.cmdAgregar.Visible = True
           Me.cmdEliminar.Visible = True
        End If
    End If
    
    'If lsOpeCod = gnAlmaIngXEmbargo Then
    '    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienEmbargado)
    'ElseIf lsOpeCod = gnAlmaIngXDacionPago Then
    '    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnAlmaIngXDacionPago)
    'ElseIf lsOpeCod = gnAlmaIngXAdjudicacion Then
    '    Me.FlexDetalle.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienAdjudicado)
    'Else
        Me.FlexDetalle.rsTextBuscar = oALmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "")
    'End If
    
    If lbMantenimiento Then
        Me.FlexSerie.lbEditarFlex = True
    End If
    
    If lsOpeCod = gnAlmaIngXCompras Then Me.txtOCompra.rs = odoc.GetOrdenesCompra("16")
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
   Dim rs As ADODB.Recordset, i As Integer
   Dim oOpe As DOperaciones
   Dim oAdj As DLogAlmacen
   Dim lnItemAnt As String
   Dim lnMovNroAnt As Long
   
   Set oOpe = New DOperaciones
   Set oAdj = New DLogAlmacen
   Set rs = New ADODB.Recordset
   Dim rt As New ADODB.Recordset
   
   HayAdj = False
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
            Set rs = oOpe.GetDetNotaIngreso(Me.txtNotaIng.Text)
        End If
        
        Me.mskFecha.Text = Format(rs!dDocFecha, gsFormatoFechaView)
        Me.txtProveedor.Text = rs!cPersCod
        Me.lblProveedorNombre.Caption = rs!cPersNombre
        Me.txtAlmacen.Text = rs!nMovBsOrden
        Me.lblAlmacenG.Caption = rs!cAlmDescripcion
        Me.txtComentario.Text = rs!cMovDesc
        
        FlexDetalle.Clear
        FlexDetalle.FormaCabecera
        If rs!cOpeCod = gnAlmaIngXAdjudicacion Then
           HayAdj = True
           Label3.Caption = "Cliente :"
           Me.FlexDetalle.ColWidth(2) = 2505
           Me.FlexDetalle.ColWidth(6) = 1000: Me.FlexDetalle.TextMatrix(0, 6) = "Val. Adjud."
           Me.FlexDetalle.ColWidth(8) = 1000
           Me.FlexDetalle.ColWidth(9) = 1000
        Else
           Label3.Caption = "Proveedor"
           Me.FlexDetalle.ColWidth(2) = 5505
           Me.FlexDetalle.ColWidth(3) = 1000
           Me.FlexDetalle.ColWidth(4) = 1000
           Me.FlexDetalle.ColWidth(8) = 0
           Me.FlexDetalle.ColWidth(9) = 0
           Me.FlexDetalle.ColWidth(10) = 0
        End If
        
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
                If rs!cOpeCod = gnAlmaIngXAdjudicacion Then
                   HayAdj = True
                   Set rt = oAdj.DatosAdjudicados(rs!nMovNro, rs!nMovItem)
                   Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 8) = rt!nProvCont
                   Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 9) = rt!nUltimaSub
                   Me.FlexDetalle.TextMatrix(FlexDetalle.Rows - 1, 10) = rt!cComentario
                End If
            End If

            If InStr(1, rs!cBSDescripcion, "[S]") <> 0 Then
                Me.FlexSerie.AdicionaFila
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 0) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 1) = rs!cSerie & ""
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 2) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 3) = rs!nMovItem
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 4) = rs!nIGV
                Me.FlexSerie.TextMatrix(Me.FlexSerie.Rows - 1, 5) = rs!nValor
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
        'lblTotalTot.Caption = FNumero(VNumero(lblTotalG.Caption) + VNumero(lblTotalIGV.Caption))
    End If
End Sub

Private Sub txtOCompra_EmiteDatos()
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim i As Integer
    Dim lbCorrelativo As Boolean
    Dim lnCorrelativoIni As Long
    Dim oALmacen As DLogAlmacen
    Dim lb
    Set oALmacen = New DLogAlmacen
    
    Dim nIGVAcum As Currency
    Dim nTotAcum   As Currency
    
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
        
        Set rs = oOpe.GetDetOC(Me.txtOCompra.Text)
        Me.FlexSerie.Rows = 2
        Me.FlexSerie.Clear
        Me.FlexSerie.FormaCabecera
        
        lnMovNroOPG = rs!nMovNro
        
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
                lbCorrelativo = oALmacen.VerfBSCorrela(rs!cBSCod)
                
                If lbCorrelativo Then
                    lnCorrelativoIni = oALmacen.GetBSCorrelaIni(rs!cBSCod)
                    
                    If GetUltCodVig(rs!cBSCod) > lnCorrelativoIni Then
                        lnCorrelativoIni = GetUltCodVig(rs!cBSCod)
                    End If
                End If
            End If
                
                nIGVAcum = 0
                nTotAcum = 0
                For i = 1 To CInt(rs!nMovCant)
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
                Next i
                
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
    Me.lblDNI = Trim(txtProveedor.sPersNroDoc)
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
    Dim i As Integer
    Dim j As Integer
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
    
    For i = 1 To Me.FlexDoc.Rows - 1
        If Me.FlexDoc.TextMatrix(i, 1) <> "" Then
            If Not IsDate(Me.FlexDoc.TextMatrix(i, 3)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDoc.Col = 3
                FlexDoc.Row = i
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            'ElseIf Not IsNumeric(Me.FlexDoc.TextMatrix(i, 4)) Or Me.FlexDoc.TextMatrix(i, 4) = 0 Then
            ElseIf Not IsNumeric(Me.FlexDoc.TextMatrix(i, 4)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDoc.Col = 4
                FlexDoc.Row = i
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDoc.TextMatrix(i, 5)) Then
                MsgBox "[Numero de Documento] debe ser un valor numérico en el Item (" & i & ")...", vbInformation, "Aviso"
                FlexDoc.Col = 5
                FlexDoc.Row = i
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            ElseIf Val(Me.FlexDoc.TextMatrix(i, 5)) = 0 Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDoc.Col = 5
                FlexDoc.Row = i
                Me.FlexDoc.SetFocus
                Valida = False
                Exit Function
            End If
        End If
    Next i
    
    If Me.FlexDetalle.TextMatrix(1, 1) = "" Then
        MsgBox "Debe ingresar por lo menos un producto.", vbInformation, "Aviso"
        Me.cmdAgregar.SetFocus
        Valida = False
        Exit Function
    End If
    
    For i = 1 To Me.FlexDetalle.Rows - 1
        If Me.FlexDetalle.TextMatrix(i, 1) = "" Then
            MsgBox "Debe ingresar un bien valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.Col = 1
            FlexDetalle.Row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf InStr(1, Me.FlexDetalle.TextMatrix(i, 2), "[S]") <> 0 Then
            lnContador = 0
            
            For j = 1 To CInt(Me.FlexSerie.Rows - 1)
                If FlexSerie.TextMatrix(j, 2) = Me.FlexDetalle.TextMatrix(i, 0) And FlexSerie.TextMatrix(j, 1) <> "" Then
                    lnContador = lnContador + 1
                End If
            Next j
        
            If lnContador <> CInt(Me.FlexDetalle.TextMatrix(i, 3)) Then
                MsgBox "Debe ingresar una numeros serie valida para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.Col = 1
                FlexDetalle.Row = i
                FlexSerie.Row = lnContador + 1
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 3)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.Col = 3
                FlexDetalle.Row = i
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 4)) Then
                MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
                FlexDetalle.Col = 4
                FlexDetalle.Row = i
                Me.FlexDetalle.SetFocus
                Valida = False
                Exit Function
            End If
            
            If Not lbMantenimiento And lsOpeCod <> gnAlmaIngXComprasConfirma Then
                For j = 1 To CInt(Me.FlexSerie.Rows - 1)
                    If FlexSerie.TextMatrix(j, 3) = Me.FlexDetalle.TextMatrix(i, 0) Then
                        If (Not lbIngreso And Not VerfBSSerieMov(Me.FlexDetalle.TextMatrix(i, 1), FlexSerie.TextMatrix(j, 1), lnMovNroG)) Or (VerfBSSerie(Me.FlexDetalle.TextMatrix(i, 1), FlexSerie.TextMatrix(j, 1), "0") And lbIngreso) Then
                            MsgBox "El bien ya fue ingresado o no ha sido descargado de almacen, para el registro " & i & " .", vbInformation, "Aviso"
                            FlexSerie.Col = 1
                            FlexSerie.Row = j
                            FlexDetalle.Row = CInt(FlexSerie.TextMatrix(FlexSerie.Row, 3))
                            FlexDetalle_RowColChange
                            Me.FlexSerie.SetFocus
                            Valida = False
                            Exit Function
                        End If
                    End If
                Next j
            End If
        
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 3)) Then
            MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.Col = 3
            FlexDetalle.Row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 4)) Then
            MsgBox "Debe ingresar un valor valido para el registro " & i & " .", vbInformation, "Aviso"
            FlexDetalle.Col = 4
            FlexDetalle.Row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        ElseIf Not IsNumeric(Me.FlexDetalle.TextMatrix(i, 6)) Then
            MsgBox "No se ha definido Cta Contable para el registro (Defina una cuenta contable para este producto) " & i & " .", vbInformation, "Aviso"
            FlexDetalle.Col = 4
            FlexDetalle.Row = i
            Me.FlexDetalle.SetFocus
            Valida = False
            Exit Function
        End If
    Next i
    
    If txtComentario.Text = "" Then
        Valida = False
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        txtComentario.SetFocus
        Exit Function
    End If
    
    Valida = True
End Function

Public Sub InicioRep(psDocNro As String)
    Dim odoc As DOperaciones
    Set odoc = New DOperaciones
    lbReporte = True
    lsOpeCod = gnAlmaIngXComprasConfirma
    lbMantenimiento = False
    lbExtorno = False
    Me.txtNotaIng.rs = odoc.GetNotaIngresoReporte("20", gnAlmarReporteMovNotIng, , psDocNro)
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
    
    For lnI = 1 To Me.FlexDetalle.Rows - 1
        nIGVAcum = nIGVAcum + VNumero(FlexDetalle.TextMatrix(lnI, 5))
        nTotAcum = nTotAcum + VNumero(FlexDetalle.TextMatrix(lnI, 6))
    Next lnI
                
    Me.lblTotalIGV.Caption = Format(nIGVAcum, "#,##0.00")
    Me.lblTotalG.Caption = Format(nTotAcum, "#,##0.00")
    Me.lblTotalTot.Caption = Format(nIGVAcum + nTotAcum, "#,##0.00")
    
    nIGVAcum = 0
    nTotAcum = 0
    
    For lnI = 1 To Me.FlexSerie.Rows - 1
        If FlexSerie.TextMatrix(lnI, 0) = pnRow Then
           nIGVAcum = nIGVAcum + VNumero(FlexSerie.TextMatrix(lnI, 4))
           nTotAcum = nTotAcum + VNumero(FlexSerie.TextMatrix(lnI, 5))
        End If
    Next lnI
    Me.lblTotalIGVDet.Caption = Format(nIGVAcum, "#,##0.00")
    Me.lblTotalGDet.Caption = Format(nTotAcum, "#,##0.00")
    
End Sub

