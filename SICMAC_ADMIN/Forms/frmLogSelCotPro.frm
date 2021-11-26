VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.OCX"
Begin VB.Form frmLogSelCotPro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7035
   ClientLeft      =   450
   ClientTop       =   1260
   ClientWidth     =   10875
   Icon            =   "frmLogSelCotPro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstPropuesta 
      Height          =   4500
      Left            =   135
      TabIndex        =   9
      Top             =   585
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   7938
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Técnica"
      TabPicture(0)   =   "frmLogSelCotPro.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEtiqueta(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEtiqueta(11)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fgeParEcoTot"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fgeParTecTot"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fgeParEco"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fgeParTec"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Económica"
      TabPicture(1)   =   "frmLogSelCotPro.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgeBS"
      Tab(1).Control(1)=   "fgeTot"
      Tab(1).Control(2)=   "lblEtiqueta(2)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Garantía de Oferta"
      TabPicture(2)   =   "frmLogSelCotPro.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtGarNro"
      Tab(2).Control(1)=   "txtComentario"
      Tab(2).Control(2)=   "fgeGarantia"
      Tab(2).Control(3)=   "lblEtiqueta(4)"
      Tab(2).Control(4)=   "lblEtiqueta(3)"
      Tab(2).Control(5)=   "lblEtiqueta(0)"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txtGarNro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70335
         MaxLength       =   25
         TabIndex        =   24
         Top             =   810
         Width           =   2205
      End
      Begin VB.TextBox txtComentario 
         Enabled         =   0   'False
         Height          =   2610
         Left            =   -70350
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1515
         Width           =   5370
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3015
         Left            =   -74895
         TabIndex        =   11
         Top             =   615
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   5318
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Precio-Total-Cant.Prov-Prec.Prov-Total Prov"
         EncabezadosAnchos=   "400-0-2000-650-900-900-900-900-900-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-7-8-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeParTec 
         Height          =   1710
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3016
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "400-0-1700-550-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X"
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-R-L"
         FormatosEdit    =   "0-0-0-3-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeParEco 
         Height          =   1710
         Left            =   5355
         TabIndex        =   14
         Top             =   630
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3016
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "400-0-1700-550-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X"
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-R-L"
         FormatosEdit    =   "0-0-0-3-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeParTecTot 
         Height          =   915
         Left            =   180
         TabIndex        =   17
         Top             =   2025
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   1614
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "400-0-1700-550-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-L-L-R-L"
         FormatosEdit    =   "0-0-0-3-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   -2147483624
      End
      Begin Sicmact.FlexEdit fgeParEcoTot 
         Height          =   915
         Left            =   5355
         TabIndex        =   18
         Top             =   2025
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   1614
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "400-0-1700-550-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-L-L-R-L"
         FormatosEdit    =   "0-0-0-3-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   -2147483624
      End
      Begin Sicmact.FlexEdit fgeTot 
         Height          =   990
         Left            =   -74895
         TabIndex        =   10
         Top             =   3330
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   1746
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "------Total---Total Prov"
         EncabezadosAnchos=   "400-0-2000-650-900-900-900-900-900-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2-2-2-2"
         AvanceCeldas    =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeGarantia 
         Height          =   3315
         Left            =   -74625
         TabIndex        =   20
         Top             =   825
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   5847
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Opc"
         EncabezadosAnchos=   "400-0-2500-400"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3"
         ListaControles  =   "0-0-0-5"
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Observacion"
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
         Height          =   210
         Index           =   4
         Left            =   -70245
         TabIndex        =   23
         Top             =   1275
         Width           =   1545
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Documento"
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
         Height          =   210
         Index           =   3
         Left            =   -70245
         TabIndex        =   22
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo de garantía"
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
         Height          =   210
         Index           =   0
         Left            =   -74550
         TabIndex        =   19
         Top             =   525
         Width           =   1545
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Económicos :"
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
         Height          =   210
         Index           =   11
         Left            =   5640
         TabIndex        =   16
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Técnicos :"
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
         Height          =   210
         Index           =   10
         Left            =   330
         TabIndex        =   15
         Top             =   375
         Width           =   1425
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Detalle"
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
         Index           =   2
         Left            =   -74820
         TabIndex        =   12
         Top             =   405
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdCot 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   8625
      TabIndex        =   3
      Top             =   5880
      Width           =   1290
   End
   Begin VB.CommandButton cmdCot 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   8625
      TabIndex        =   2
      Top             =   5370
      Width           =   1290
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8625
      TabIndex        =   4
      Top             =   6435
      Width           =   1290
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   75
      Top             =   6435
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeCot 
      Height          =   1605
      Left            =   300
      TabIndex        =   1
      Top             =   5370
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   2831
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Cotización-Codigo-Proveedor-Dirección-Ok"
      EncabezadosAnchos=   "400-800-0-3500-0-350"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-5"
      ListaControles  =   "0-0-0-0-0-5"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
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
      TipoBusqueda    =   2
      sTitulo         =   ""
   End
   Begin VB.Label lblProvee 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3330
      TabIndex        =   8
      Top             =   225
      Width           =   4245
   End
   Begin VB.Label lblProvEti 
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
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   3480
      TabIndex        =   7
      Top             =   30
      Width           =   960
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Proceso de selección"
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
      Height          =   210
      Index           =   1
      Left            =   405
      TabIndex        =   6
      Top             =   30
      Width           =   1950
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Cotizaciones"
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
      Height          =   210
      Index           =   5
      Left            =   375
      TabIndex        =   5
      Top             =   5160
      Width           =   1185
   End
End
Attribute VB_Name = "frmLogSelCotPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim psTpoPro As String

Public Sub Inicio(ByVal psFormTpo As String, ByVal psTipoProp As String)
'Presentacion de propuestas [1] o evaluacion [2]
psFrmTpo = psFormTpo
'Propuesta Técnica [1] - Económica [2] - Garantía Oferta [3]
psTpoPro = psTipoProp
Me.Show 1
End Sub

Private Sub cmdCot_Click(Index As Integer)
    Dim clsDGnral As DLogGeneral
    Dim clsDMov  As DLogMov
    Dim sSelNro As String, sSelCotNro As String, sSelTraNro As String, sBSCod As String
    Dim nSelNro As Integer, nSelCotNro As Integer, nSelTraNro As Integer
    Dim sActualiza As String, sProvee As String
    Dim nCont As Integer, nSum As Integer, nResult As Integer
    Dim nCantidad As Currency, nPrecio As Currency
    Dim nParCod As Integer, nGarTpo As Integer
    
    Select Case Index
        Case 0:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                cmdCot(0).Enabled = False
                cmdCot(1).Enabled = False
                txtSelNro.Text = ""
                Call Limpiar
                Call CargaTxtSelNro
            End If
        Case 1:
            'GRABACION
            sSelNro = txtSelNro.Text
            nSelCotNro = fgeCot.TextMatrix(fgeCot.Row, 1)
            If psFrmTpo = "1" Then
                'PROPUESTA
                If psTpoPro = "1" Then
                    'TECNICA
                    'GRABAR
                    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                        Set clsDGnral = New DLogGeneral
                        sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDGnral = Nothing
                        
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDMov = New DLogMov
                        
                        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoCotizacion
                        nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                        nSelNro = clsDMov.GetnMovNro(sSelNro)
                        clsDMov.InsertaMovRef nSelTraNro, nSelNro
                        
                        For nCont = 1 To fgeParTec.Rows - 1
                            nParCod = fgeParTec.TextMatrix(nCont, 1)
                            nCantidad = Val(fgeParTec.TextMatrix(nCont, 3))
                            clsDMov.ActualizaSelCotPar nSelNro, nSelCotNro, 1, _
                                nParCod, nCantidad, sActualiza
                        Next
                        
                        For nCont = 1 To fgeParEco.Rows - 1
                            nParCod = fgeParEco.TextMatrix(nCont, 1)
                            nCantidad = Val(fgeParEco.TextMatrix(nCont, 3))
                            clsDMov.ActualizaSelCotPar nSelNro, nSelCotNro, 2, _
                                nParCod, nCantidad, sActualiza
                        Next
                        
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            fgeBS.lbEditarFlex = False
                            cmdCot(0).Enabled = False
                            cmdCot(1).Enabled = False
                            Call CargaTxtSelNro
                        Else
                            MsgBox "Error al grabar la información", vbInformation, " Aviso "
                        End If
                    End If
                ElseIf psTpoPro = "2" Then
                    'ECONOMICA

                    If CCur(IIf(fgeTot.TextMatrix(1, 9) = "", 0, fgeTot.TextMatrix(1, 9))) = 0 Then
                        MsgBox "Falta determinar cantidades y precios", vbInformation, " Aviso"
                        Exit Sub
                    End If
                    'Verifica si cantidades son iguales
                    For nCont = 1 To fgeBS.Rows - 1
                        If CCur(IIf(fgeBS.TextMatrix(nCont, 4) = "", 0, fgeBS.TextMatrix(nCont, 4))) <> CCur(IIf(fgeBS.TextMatrix(nCont, 7) = "", 0, fgeBS.TextMatrix(nCont, 7))) Then
                            nSum = nSum + 1
                        End If
                    Next
                    If nSum > 0 Then
                        If MsgBox("Cantidades ingresadas son diferentes a las solicitadas " & vbCr & "¿ Deseas continuar con la grabación ? ", vbQuestion + vbYesNo, " Aviso ") = vbNo Then
                            Exit Sub
                        End If
                    End If
                    'GRABAR
                    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                        Set clsDGnral = New DLogGeneral
                        sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDGnral = Nothing
                        
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDMov = New DLogMov
                        
                        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoCotizacion
                        nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                        nSelNro = clsDMov.GetnMovNro(sSelNro)
                        clsDMov.InsertaMovRef nSelTraNro, nSelNro
                        
                        For nCont = 1 To fgeBS.Rows - 1
                            sBSCod = fgeBS.TextMatrix(nCont, 1)
                            nCantidad = CCur(IIf(fgeBS.TextMatrix(nCont, 7) = "", 0, fgeBS.TextMatrix(nCont, 7)))
                            nPrecio = CCur(IIf(fgeBS.TextMatrix(nCont, 8) = "", 0, fgeBS.TextMatrix(nCont, 8)))
                            clsDMov.ActualizaSelCotDetalle nSelNro, nSelCotNro, sBSCod, _
                                nCantidad, nPrecio, sActualiza
                        Next
                        
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            fgeBS.lbEditarFlex = False
                            cmdCot(0).Enabled = False
                            cmdCot(1).Enabled = False
                            Call CargaTxtSelNro
                        Else
                            MsgBox "Error al grabar la información", vbInformation, " Aviso "
                        End If
                    End If
                ElseIf psTpoPro = "3" Then
                    'GARANTIA DE SERIEDAD
                    'GRABAR
                    txtGarNro.Text = Replace(txtGarNro.Text, "'", " ", , , vbTextCompare)
                    txtComentario.Text = Replace(txtComentario.Text, "'", " ", , , vbTextCompare)
                    For nCont = 1 To fgeGarantia.Rows - 1
                        If fgeGarantia.TextMatrix(nCont, 3) = "." Then
                            nGarTpo = Val(fgeGarantia.TextMatrix(nCont, 1))
                            Exit For
                        End If
                    Next
                    If nGarTpo = 0 Then
                        MsgBox "Falta determinar el tipo de Garantía", vbInformation, " Aviso "
                        Exit Sub
                    End If
                    
                    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                        Set clsDGnral = New DLogGeneral
                        sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDGnral = Nothing
                        
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDMov = New DLogMov
                        
                        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoCotizacion
                        nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                        nSelNro = clsDMov.GetnMovNro(sSelNro)
                        clsDMov.InsertaMovRef nSelTraNro, nSelNro
                        
                        clsDMov.ActualizaSelCotizacion nSelNro, nSelCotNro, _
                            nGarTpo, txtGarNro.Text, txtComentario.Text, sActualiza
                        
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            fgeGarantia.lbEditarFlex = False
                            'txtGarNro.Enabled = False
                            'txtComentario.Enabled = False
                            cmdCot(0).Enabled = False
                            cmdCot(1).Enabled = False
                            Call CargaTxtSelNro
                        Else
                            MsgBox "Error al grabar la información", vbInformation, " Aviso "
                        End If
                    End If
                End If
            ElseIf psFrmTpo = "2" Then
                'EVALUACION
                For nCont = 1 To fgeCot.Rows - 1
                    If fgeCot.TextMatrix(nCont, 5) = "." Then
                        sSelCotNro = fgeCot.TextMatrix(nCont, 1)
                        sProvee = fgeCot.TextMatrix(nCont, 3)
                        Exit For
                    End If
                Next
                If sSelNro = "" Or sSelCotNro = "" Then
                    MsgBox "Falta seleccionar una cotización ", vbInformation, " Aviso"
                    Exit Sub
                End If
                
                'EVALUACION TECNICA Y ECONOMICA
                If MsgBox("¿ Estás seguro de Adjudicar a " & sProvee & vbCr & " el proceso de Selección ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Inserta MOV - MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoProcAdju
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion - Inserta la cotizacion Adjudicada
                    clsDMov.ActualizaSeleccionAdju nSelNro, nSelCotNro, sActualiza
                        
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgeCot.lbEditarFlex = False
                        cmdCot(0).Enabled = False
                        cmdCot(1).Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                End If
            End If
        Case Else
            MsgBox "Tipo de comando no reconocido", vbInformation, " Aviso"
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgeBS_OnCellChange(pnRow As Long, pnCol As Long)
    If psFrmTpo = "1" And psTpoPro = "2" Then
        If pnCol = 7 Or pnCol = 8 Then
            fgeBS.TextMatrix(pnRow, 9) = Format(CCur(IIf(fgeBS.TextMatrix(pnRow, 7) = "", 0, fgeBS.TextMatrix(pnRow, 7))) * CCur(IIf(fgeBS.TextMatrix(pnRow, 8) = "", 0, fgeBS.TextMatrix(pnRow, 8))), "#,##0.00")
            fgeTot.TextMatrix(1, 9) = Format(fgeBS.SumaRow(9), "#,##0.00")
        End If
    End If
End Sub

Private Sub fgeCot_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDGnral As DLogGeneral
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim nSelNro As Integer, nSelCotNro As Integer
    Dim nCont  As Integer
    Dim clsDMov As DLogMov
    
    Set clsDMov = New DLogMov
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    
    If psFrmTpo = "1" Then
        'PROPUESTAS
        'Verifica que siempre este por lo menos UNO
        If fgeCot.TextMatrix(1, 1) = "" Then
            Exit Sub
        End If
        lblProvee.Caption = fgeCot.TextMatrix(fgeCot.Row, 3)
        nSelNro = clsDMov.GetnMovNro(txtSelNro.Text)
        nSelCotNro = Val(fgeCot.TextMatrix(fgeCot.Row, 1))
        If psTpoPro = "1" Then
            'TECNICA
            fgeParEco.lbEditarFlex = True
            fgeParTec.lbEditarFlex = True
            cmdCot(0).Enabled = True
            cmdCot(1).Enabled = True            'Muestra datos
            Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, nSelNro, 1, nSelCotNro)
            If rs.RecordCount > 0 Then
                Set fgeParTec.Recordset = rs
            End If
            Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, nSelNro, 2, nSelCotNro)
            If rs.RecordCount > 0 Then
                Set fgeParEco.Recordset = rs
            End If
        ElseIf psTpoPro = "2" Then
            'ECONOMICA
            fgeBS.lbEditarFlex = True
            cmdCot(0).Enabled = True
            cmdCot(1).Enabled = True
            'Muestra datos
            Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegistroCotiza, nSelNro, nSelCotNro)
            If rs.RecordCount > 0 Then
                Set fgeBS.Recordset = rs
            End If
            'Color en las COLUMNAS
            For nCont = 1 To fgeBS.Rows - 1
                fgeBS.Row = nCont
                fgeBS.Col = 7
                fgeBS.CellForeColor = vbBlue '&HC0FFC0 '(verde)
                fgeBS.Col = 8
                fgeBS.CellForeColor = vbBlue '&HC0FFC0 '(verde)
            Next
            fgeTot.TextMatrix(1, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
            fgeTot.TextMatrix(1, 9) = Format(fgeBS.SumaRow(9), "#,##0.00")
        ElseIf psTpoPro = "3" Then
            'Garantía de seriedad de Oferta
            
            txtGarNro.Enabled = True
            txtComentario.Enabled = True
            fgeGarantia.lbEditarFlex = True
            cmdCot(0).Enabled = True
            cmdCot(1).Enabled = True
            
            Set rs = clsDGnral.CargaConstante(gPersGarantia)
            If rs.RecordCount > 0 Then
                Set fgeGarantia.Recordset = rs
            End If
            
            Set rs = clsDAdq.CargaSelCotiza(SelCotGarantia, nSelNro, nSelCotNro)
            txtGarNro.Text = rs!cLogSelCotGarNro
            txtComentario.Text = rs!cLogSelCotGarComentario
            For nCont = 1 To fgeGarantia.Rows - 1
                If rs!nLogSelCotGarTpo = Val(fgeGarantia.TextMatrix(nCont, 1)) Then
                    fgeGarantia.TextMatrix(nCont, 3) = 1
                    Exit For
                End If
            Next
            
        End If
    ElseIf psFrmTpo = "2" Then
        'Solo intercambia entre GARANTIAS DE SERIEDAD DE OFERTA
        nSelNro = clsDMov.GetnMovNro(txtSelNro.Text)
        nSelCotNro = Val(fgeCot.TextMatrix(fgeCot.Row, 1))
        
        'Set rs = clsDGnral.CargaConstante(gPersGarantia)
        'If rs.RecordCount > 0 Then
        '    Set fgeGarantia.Recordset = rs
        'End If
        'Set clsDGnral = Nothing
        
        Set rs = clsDAdq.CargaSelCotiza(SelCotGarantia, nSelNro, nSelCotNro)
        If rs.RecordCount > 0 Then
            txtGarNro.Text = rs!cLogSelCotGarNro
            txtComentario.Text = rs!cLogSelCotGarComentario
            fgeGarantia.TextMatrix(1, 0) = "1"
            fgeGarantia.TextMatrix(1, 1) = rs!nLogSelCotGarTpo
            fgeGarantia.TextMatrix(1, 2) = rs!cConsDescripcion
            
            'For nCont = 1 To fgeGarantia.Rows - 1
            '    If rs!nLogSelCotGarTpo = Val(fgeGarantia.TextMatrix(nCont, 1)) Then
            '        fgeGarantia.TextMatrix(nCont, 3) = 1
            '        Exit For
            '    End If
            'Next
        End If
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
    Set clsDGnral = Nothing
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If psFrmTpo = "1" Then
        If psTpoPro = "1" Then
            Me.Caption = "Propuesta Técnica"
            sstPropuesta.TabVisible(1) = False
            sstPropuesta.TabVisible(2) = False
        ElseIf psTpoPro = "2" Then
            Me.Caption = "Propuesta Económica"
            sstPropuesta.TabVisible(0) = False
            sstPropuesta.TabVisible(2) = False
        ElseIf psTpoPro = "3" Then
            Me.Caption = "Garantía de seriedad de Oferta"
            sstPropuesta.TabVisible(0) = False
            sstPropuesta.TabVisible(1) = False
        End If
        fgeCot.ListaControles = "0-0-0-0-0-0"
        fgeCot.EncabezadosAnchos = "400-800-0-3500-0-0"
    ElseIf psFrmTpo = "2" Then
        Me.Caption = "Evaluación de Propuestas"
        lblProvEti.Visible = False
        lblProvee.Visible = False
        
        fgeParTec.lbEditarFlex = False
        fgeParEco.lbEditarFlex = False
        fgeParTec.EncabezadosAnchos = "400-0-1700-550-700"
        fgeParEco.EncabezadosAnchos = "400-0-1700-550-700"
        fgeParTecTot.EncabezadosAnchos = "400-0-1700-550-700"
        fgeParEcoTot.EncabezadosAnchos = "400-0-1700-550-700"
        
        fgeGarantia.EncabezadosAnchos = "400-0-2500-0"
        fgeGarantia.lbEditarFlex = False
        txtGarNro.Enabled = False
        txtComentario.Enabled = False
    Else
        MsgBox "Tipo de formulario no reconocido", vbInformation, " Aviso"
    End If
    
    Call CargaTxtSelNro
End Sub

Private Sub CargaTxtSelNro()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    
    Set rs = clsDAdq.CargaSeleccion(SelTodosEstadoParaProAdj)
    If rs.RecordCount > 0 Then
        txtSelNro.EditFlex = True
        txtSelNro.rs = rs
        txtSelNro.EditFlex = False
    Else
        txtSelNro.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sAdqNro As String
    Dim sBSCod As String, sLogSelCotNroAnt As String
    Dim nCont As Integer, nCont2 As Integer, pColIni As Integer
    Dim nSelNro As Integer, nSelCotNro As Integer
    Dim nMax As Currency, nMin As Currency, nMon As Currency
    Dim nPun As Currency, nTot As Currency
    Dim clsDMov As DLogMov
    
    If txtSelNro.Ok = False Then Exit Sub
    
    Set clsDMov = New DLogMov
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    Call Limpiar
    
    nSelNro = clsDMov.GetnMovNro(txtSelNro.Text)
    
    Set rs = clsDAdq.CargaSelCotiza(SelCotNroCotPersona, nSelNro)
    If rs.RecordCount > 0 Then
        Set fgeCot.Recordset = rs
        If psFrmTpo = "1" Then
            'PROPUESTA
            Call fgeCot_OnRowChange(fgeCot.Row, fgeCot.Col)
        ElseIf psFrmTpo = "2" Then
            'EVALUACION
            '********************************************************************************
            'TECNICA
            Set rs = clsDAdq.CargaSelParametro(nSelNro, 1)
            If rs.RecordCount > 0 Then
                Set fgeParTec.Recordset = rs
                'Carga parametros de los proveedores
                CargaParPro fgeParTec, fgeParTecTot, fgeCot, nSelNro, 1
                'Carga Puntajes
                CargaPuntaje fgeParTec, fgeParTecTot, nSelNro, 1
            End If
            Set rs = clsDAdq.CargaSelParametro(nSelNro, 2)
            If rs.RecordCount > 0 Then
                Set fgeParEco.Recordset = rs
                'Carga parametros de los proveedores
                CargaParPro fgeParEco, fgeParEcoTot, fgeCot, nSelNro, 2
                'Carga Puntajes
                CargaPuntaje fgeParEco, fgeParEcoTot, nSelNro, 1
            End If
            '********************************************************************************
            'ECONOMICA
            'Set rs = clsDAdq.CargaSelDetalle(nSelNro)
            Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegEvalua, nSelNro, 0)
            If rs.RecordCount > 0 Then
                fgeCot.lbEditarFlex = True
                cmdCot(0).Enabled = True
                cmdCot(1).Enabled = True
                
                'Base izquierda del FLEX
                Set fgeBS.Recordset = rs
                fgeTot.TextMatrix(1, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
                'Carga los detalles
                Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegEvaluaTodos, nSelNro, 0)
                pColIni = 7
                fgeBS.Cols = (fgeBS.Cols - 3) + ((fgeCot.Rows - 1) * 3)
                fgeTot.Cols = fgeBS.Cols
                
                For nCont = 1 To fgeCot.Rows - 1
                    fgeBS.TextMatrix(0, pColIni) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Cantidad"
                    fgeBS.TextMatrix(0, pColIni + 1) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Precio"
                    fgeBS.TextMatrix(0, pColIni + 2) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Total"
                    
                    fgeTot.TextMatrix(0, pColIni + 2) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Total"
                    pColIni = pColIni + 3
                Next
                
                For nCont = 1 To fgeBS.Rows - 1
                    sBSCod = fgeBS.TextMatrix(nCont, 1)
                    pColIni = 7
                    rs.MoveFirst
                    sLogSelCotNroAnt = rs!nLogSelCotNro
                    Do While Not rs.EOF
                        If sLogSelCotNroAnt <> rs!nLogSelCotNro Then
                            pColIni = pColIni + 3
                            sLogSelCotNroAnt = rs!nLogSelCotNro
                        End If
                        If sBSCod = rs!cBSCod Then
                            fgeBS.Row = nCont
                            fgeBS.Col = pColIni
                            fgeBS.CellForeColor = vbBlue
                            fgeBS.Col = pColIni + 1
                            fgeBS.CellForeColor = vbBlue
                            fgeBS.TextMatrix(nCont, pColIni) = Format(rs!nLogSelCotDetCantidad, "#,##0.00")
                            fgeBS.TextMatrix(nCont, pColIni + 1) = Format(rs!nLogSelCotDetPrecio, "#,##0.00")
                            fgeBS.TextMatrix(nCont, pColIni + 2) = Format(rs!Total, "#,##0.00")
                        End If
                        rs.MoveNext
                    Loop
                Next
            End If
            pColIni = 9
            For nCont = 1 To (fgeCot.Rows - 1)
                fgeTot.TextMatrix(1, pColIni) = Format(fgeBS.SumaRow(pColIni), "#,##0.00")
                pColIni = pColIni + 3
            Next
            
            'PARA INTERCAMBIO DE GARANTIAS
            Call fgeCot_OnRowChange(fgeCot.Row, fgeCot.Col)
        End If
    End If
End Sub

Private Sub CargaParPro(pfgePar As FlexEdit, pfgeParTot As FlexEdit, pfgeCot As FlexEdit, _
ByVal pnSelNro As Long, ByVal pnParTpo As Long)
    Dim nCol As Integer, nCont As Integer, nCont2 As Integer
    Dim nSelCotNro As Integer
    Dim rs As ADODB.Recordset
    Dim clsDAdq As DLogAdquisi
    
    Set rs = New ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    For nCont = 1 To pfgeCot.Rows - 1
        'Aumenta una Columna
        nCol = pfgePar.Cols + 1
        pfgePar.Cols = nCol
        pfgePar.ColWidth(nCol - 1) = 400
        pfgePar.TextMatrix(0, nCol - 1) = " P" & nCont
        pfgeParTot.Cols = nCol
        pfgeParTot.ColWidth(nCol - 1) = 400
        
        nSelCotNro = Val(pfgeCot.TextMatrix(nCont, 1))
        Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, pnSelNro, pnParTpo, nSelCotNro)
        Do While Not rs.EOF
            For nCont2 = 1 To pfgePar.Rows - 1
                If Val(pfgePar.TextMatrix(nCont2, 1)) = rs!nLogSelParNro Then
                    pfgePar.TextMatrix(nCont2, nCol - 1) = rs!nLogSelCotParValor
                    Exit For
                End If
            Next
            rs.MoveNext
        Loop
    Next
End Sub

Private Sub CargaPuntaje(pfgePar As FlexEdit, pfgeParTot As FlexEdit, _
ByVal pnSelNro As Long, ByVal pnParTpo As Long)
    Dim nCont As Integer, nCont2 As Integer
    Dim nMax As Currency, nMin As Currency
    Dim nPun As Currency, nMon As Currency, nTot As Currency
    Dim rs As ADODB.Recordset
    Dim clsDAdq As DLogAdquisi
    
    Set rs = New ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    
    For nCont = 1 To pfgePar.Rows - 1
        nMax = 0 'pfgePar.TextMatrix(nCont, 5)
        nMin = 999 'pfgePar.TextMatrix(nCont, 5)
        For nCont2 = 5 To pfgePar.Cols - 1
            nMon = Val(pfgePar.TextMatrix(nCont, nCont2))
            If nMax < nMon Then nMax = nMon
            If nMin > nMon And nMon <> 0 Then nMin = nMon
        Next
        If Right(pfgePar.TextMatrix(nCont, 4), 1) = "1" Then
            'Directa
            For nCont2 = 5 To pfgeParTot.Cols - 1
                nTot = Val(pfgeParTot.TextMatrix(1, nCont2))
                nMon = Val(pfgePar.TextMatrix(nCont, nCont2))
                nPun = Val(pfgePar.TextMatrix(nCont, 3))
                If nMax > 0 Then
                    pfgeParTot.TextMatrix(1, nCont2) = Round(nTot + ((nMon * nPun) / nMax), 2)
                End If
            Next
        ElseIf Right(pfgePar.TextMatrix(nCont, 4), 1) = "2" Then
            'Inversa
            For nCont2 = 5 To pfgeParTot.Cols - 1
                nTot = Val(pfgeParTot.TextMatrix(1, nCont2))
                nMon = Val(pfgePar.TextMatrix(nCont, nCont2))
                nPun = Val(pfgePar.TextMatrix(nCont, 3))
                If nMon > 0 Then
                    pfgeParTot.TextMatrix(1, nCont2) = Round(nTot + ((nMin * nPun) / nMon), 2)
                End If
            Next
        Else
            'Rangos
            For nCont2 = 5 To pfgeParTot.Cols - 1
                nTot = Val(pfgeParTot.TextMatrix(1, nCont2))
                nMon = Val(pfgePar.TextMatrix(nCont, nCont2))
                nPun = Val(pfgePar.TextMatrix(nCont, 1))
                
                Set rs = clsDAdq.CargaSelParDetalle(pnSelNro, pnParTpo, nPun)
                Do While Not rs.EOF
                    If nMon >= rs!nLogSelParDetIni And nMon <= rs!nLogSelParDetFin Then
                        pfgeParTot.TextMatrix(1, nCont2) = Round((nTot + rs!nLogSelParDetPuntaje), 2)
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
            Next
        End If
    Next

End Sub

Private Sub Limpiar()
    lblProvee.Caption = ""
    
    fgeParTec.Clear
    fgeParTec.FormaCabecera
    fgeParTec.Rows = 2
    fgeParTec.Cols = 5
    fgeParTecTot.Clear
    fgeParTecTot.FormaCabecera
    fgeParTecTot.Rows = 2
    fgeParTecTot.Cols = 5
    fgeParTecTot.TextMatrix(1, 0) = "="
    fgeParTecTot.TextMatrix(1, 1) = "."
    If psFrmTpo <> "1" Then fgeParTecTot.TextMatrix(1, 2) = "PUNTAJES"
    
    fgeParEco.Clear
    fgeParEco.FormaCabecera
    fgeParEco.Rows = 2
    fgeParEco.Cols = 5
    fgeParEcoTot.Clear
    fgeParEcoTot.FormaCabecera
    fgeParEcoTot.Rows = 2
    fgeParEcoTot.Cols = 5
    fgeParEcoTot.TextMatrix(1, 0) = "="
    fgeParEcoTot.TextMatrix(1, 1) = "."
    If psFrmTpo <> "1" Then fgeParEcoTot.TextMatrix(1, 2) = "PUNTAJES"
    
    fgeCot.Clear
    fgeCot.FormaCabecera
    fgeCot.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeBS.Cols = 10
    fgeTot.Clear
    fgeTot.FormaCabecera
    fgeTot.Rows = 2
    fgeTot.Cols = 10
    fgeTot.TextMatrix(1, 0) = "="
    fgeTot.TextMatrix(1, 1) = "."
    fgeTot.TextMatrix(1, 2) = "T O T A L E S"
    
    
    txtGarNro.Text = ""
    txtComentario.Text = ""
    fgeGarantia.Clear
    fgeGarantia.FormaCabecera
    fgeGarantia.Rows = 2
End Sub
