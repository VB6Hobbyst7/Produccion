VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogSelConsol 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6735
   ClientLeft      =   -15
   ClientTop       =   1500
   ClientWidth     =   11730
   Icon            =   "frmLogSelConsol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdq 
      Enabled         =   0   'False
      Height          =   390
      Left            =   9510
      TabIndex        =   1
      Top             =   5055
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9495
      TabIndex        =   0
      Top             =   5715
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   10965
      Top             =   5805
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin TabDlg.SSTab sstSeleccion 
      Height          =   6465
      Left            =   60
      TabIndex        =   2
      Top             =   105
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   13
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Resolución"
      TabPicture(0)   =   "frmLogSelConsol.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblEtiqueta(9)"
      Tab(0).Control(1)=   "lblEtiqueta(1)"
      Tab(0).Control(2)=   "lblEtiqueta(2)"
      Tab(0).Control(3)=   "lblEtiqueta(3)"
      Tab(0).Control(4)=   "fgeSelTpo"
      Tab(0).Control(5)=   "dtpResFec"
      Tab(0).Control(6)=   "fgeAutoriza"
      Tab(0).Control(7)=   "txtResNro"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Comité"
      TabPicture(1)   =   "frmLogSelConsol.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEtiqueta(6)"
      Tab(1).Control(1)=   "fgeComite"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Detalle"
      TabPicture(2)   =   "frmLogSelConsol.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fgeBSTotal"
      Tab(2).Control(1)=   "fgeBS"
      Tab(2).Control(2)=   "fraMoneda"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Cronograma"
      TabPicture(3)   =   "frmLogSelConsol.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgeCronograma"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Referencias"
      TabPicture(4)   =   "frmLogSelConsol.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblEtiqueta(16)"
      Tab(4).Control(1)=   "lblEtiqueta(15)"
      Tab(4).Control(2)=   "lblEtiqueta(7)"
      Tab(4).Control(3)=   "fgeSisAdj"
      Tab(4).Control(4)=   "txtValReferencia"
      Tab(4).Control(5)=   "txtCostoBase"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Parámetros"
      TabPicture(5)   =   "frmLogSelConsol.frx":0396
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "lblEtiqueta(10)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lblEtiqueta(11)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblTotTec"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "lblTotEco"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lblEtiqueta(14)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "lblEtiqueta(4)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "fgeParEco"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "fgeParTec"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "fraParEcoRango"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "fraParTecRango"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Publicación"
      TabPicture(6)   =   "frmLogSelConsol.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fgePublica"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Cotización"
      TabPicture(7)   =   "frmLogSelConsol.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblEtiqueta(19)"
      Tab(7).Control(1)=   "fgeCot"
      Tab(7).Control(2)=   "fgeSel"
      Tab(7).Control(3)=   "fgePro"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Bases"
      TabPicture(8)   =   "frmLogSelConsol.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblEtiqueta(20)"
      Tab(8).Control(1)=   "fgePostor"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Consultas"
      TabPicture(9)   =   "frmLogSelConsol.frx":0406
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lblObserva(0)"
      Tab(9).Control(1)=   "lblEtiqueta(21)"
      Tab(9).Control(2)=   "rtfConsulta"
      Tab(9).Control(3)=   "fgeConsulta"
      Tab(9).ControlCount=   4
      TabCaption(10)  =   "Absolución"
      TabPicture(10)  =   "frmLogSelConsol.frx":0422
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lblEtiqueta(22)"
      Tab(10).Control(1)=   "lblObserva(1)"
      Tab(10).Control(2)=   "rtfAbsolucion"
      Tab(10).Control(3)=   "fgeAbsolucion"
      Tab(10).ControlCount=   4
      TabCaption(11)  =   "Observación"
      TabPicture(11)  =   "frmLogSelConsol.frx":043E
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "lblEtiqueta(23)"
      Tab(11).Control(1)=   "lblObserva(2)"
      Tab(11).Control(2)=   "rtfObservacion"
      Tab(11).Control(3)=   "fgeObservacion"
      Tab(11).ControlCount=   4
      TabCaption(12)  =   "Evaluación"
      TabPicture(12)  =   "frmLogSelConsol.frx":045A
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lblEtiqueta(0)"
      Tab(12).Control(1)=   "fgeEvaParEco"
      Tab(12).Control(2)=   "fgeEvaParTec"
      Tab(12).Control(3)=   "fgeEvaEcoTot"
      Tab(12).Control(4)=   "fgeEvaEcoPos"
      Tab(12).Control(5)=   "fgeEvaEco"
      Tab(12).ControlCount=   6
      Begin VB.Frame fraMoneda 
         Caption         =   "Moneda "
         Enabled         =   0   'False
         Height          =   1710
         Left            =   -74805
         TabIndex        =   53
         Top             =   1155
         Width           =   1260
         Begin VB.TextBox txtTipCambio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   150
            TabIndex        =   56
            Top             =   1230
            Width           =   945
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Soles"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   55
            Top             =   285
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Dólares"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   54
            Top             =   570
            Width           =   900
         End
         Begin VB.Label lblTipCambio 
            Caption         =   "Tipo Cambio"
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
            Left            =   75
            TabIndex        =   57
            Top             =   945
            Width           =   1065
         End
      End
      Begin VB.TextBox txtResNro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69885
         MaxLength       =   20
         TabIndex        =   44
         Top             =   1170
         Width           =   1890
      End
      Begin VB.TextBox txtCostoBase 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72150
         TabIndex        =   8
         Top             =   1320
         Width           =   1305
      End
      Begin VB.TextBox txtValReferencia 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72150
         TabIndex        =   7
         Top             =   1935
         Width           =   1305
      End
      Begin VB.Frame fraParTecRango 
         Caption         =   "Rangos"
         Height          =   2010
         Left            =   135
         TabIndex        =   5
         Top             =   3285
         Visible         =   0   'False
         Width           =   3705
         Begin Sicmact.FlexEdit fgeParTecRango 
            Height          =   1695
            Left            =   255
            TabIndex        =   6
            Top             =   225
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   2990
            Rows            =   5
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-De-A-Puntos"
            EncabezadosAnchos=   "400-800-800-800"
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
            ColumnasAEditar =   "X-1-2-3"
            ListaControles  =   "0-0-0-0"
            EncabezadosAlineacion=   "C-R-R-C"
            FormatosEdit    =   "0-3-3-3"
            CantDecimales   =   0
            TextArray0      =   "Item"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraParEcoRango 
         Caption         =   "Rangos"
         Height          =   2010
         Left            =   3930
         TabIndex        =   3
         Top             =   3285
         Visible         =   0   'False
         Width           =   3720
         Begin Sicmact.FlexEdit fgeParEcoRango 
            Height          =   1695
            Left            =   255
            TabIndex        =   4
            Top             =   210
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   2990
            Rows            =   5
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-De-A-Puntos"
            EncabezadosAnchos=   "400-800-800-800"
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
            ColumnasAEditar =   "X-1-2-3"
            ListaControles  =   "0-0-0-0"
            EncabezadosAlineacion=   "C-R-R-C"
            FormatosEdit    =   "0-3-3-3"
            CantDecimales   =   0
            TextArray0      =   "Item"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin Sicmact.FlexEdit fgeComite 
         Height          =   3405
         Left            =   -74715
         TabIndex        =   9
         Top             =   1380
         Width           =   7230
         _ExtentX        =   12753
         _ExtentY        =   6006
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Codigo-Area"
         EncabezadosAnchos=   "400-1200-2700-0-2500"
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
         EncabezadosAlineacion=   "C-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgePublica 
         Height          =   2955
         Left            =   -74655
         TabIndex        =   10
         Top             =   1485
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   5212
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha Inicial-Fecha Final"
         EncabezadosAnchos=   "400-1100-2500-1300-1300"
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
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgeParTec 
         Height          =   1710
         Left            =   135
         TabIndex        =   11
         Top             =   1275
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   3016
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "350-0-1700-600-700"
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
         EncabezadosAlineacion=   "C-L-L-R-L"
         FormatosEdit    =   "0-0-0-3-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeParEco 
         Height          =   1710
         Left            =   3945
         TabIndex        =   12
         Top             =   1275
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   3016
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "350-0-1700-600-700"
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
         EncabezadosAlineacion=   "C-L-L-R-L"
         FormatosEdit    =   "0-0-0-3-0"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeSisAdj 
         Height          =   2325
         Left            =   -74385
         TabIndex        =   13
         Top             =   2850
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   4101
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Opc"
         EncabezadosAnchos=   "400-0-3000-0"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeCronograma 
         Height          =   3270
         Left            =   -74640
         TabIndex        =   14
         Top             =   1365
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   5768
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Descripción-Chk-Fecha Inicial-Fecha Final"
         EncabezadosAnchos=   "400-0-3000-0-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgePro 
         Height          =   1755
         Left            =   -74820
         TabIndex        =   15
         Top             =   4515
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   3096
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Nombre-Dirección"
         EncabezadosAnchos=   "400-0-2500-3200"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeSel 
         Height          =   855
         Left            =   -74820
         TabIndex        =   16
         Top             =   3405
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   1508
         Cols0           =   6
         ScrollBars      =   0
         AllowUserResizing=   3
         EncabezadosNombres=   ".-----"
         EncabezadosAnchos=   "400-0-3000-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "."
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   -2147483624
      End
      Begin Sicmact.FlexEdit fgeCot 
         Height          =   2355
         Left            =   -74820
         TabIndex        =   17
         Top             =   1065
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   4154
         Cols0           =   6
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Bien/Servicio-Unidad-Cantidad-Precio"
         EncabezadosAnchos=   "400-0-3000-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-L-R-R"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgePostor 
         Height          =   2955
         Left            =   -74595
         TabIndex        =   18
         Top             =   1440
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   5212
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1100-2500-1000-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-L-C-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
         CellBackColor   =   16777215
      End
      Begin Sicmact.FlexEdit fgeConsulta 
         Height          =   1860
         Left            =   -74700
         TabIndex        =   19
         Top             =   1410
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   3281
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1100-3000-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-L-C-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
         CellBackColor   =   16777215
      End
      Begin RichTextLib.RichTextBox rtfConsulta 
         Height          =   2040
         Left            =   -74715
         TabIndex        =   20
         Top             =   3645
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   3598
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelConsol.frx":0476
      End
      Begin Sicmact.FlexEdit fgeAbsolucion 
         Height          =   1860
         Left            =   -74700
         TabIndex        =   21
         Top             =   1410
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   3281
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1100-3000-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-L-C-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
         CellBackColor   =   16777215
      End
      Begin RichTextLib.RichTextBox rtfAbsolucion 
         Height          =   2040
         Left            =   -74715
         TabIndex        =   22
         Top             =   3645
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   3598
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelConsol.frx":04E4
      End
      Begin Sicmact.FlexEdit fgeObservacion 
         Height          =   1860
         Left            =   -74700
         TabIndex        =   23
         Top             =   1410
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   3281
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1100-3000-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-L-C-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
         CellBackColor   =   16777215
      End
      Begin RichTextLib.RichTextBox rtfObservacion 
         Height          =   2040
         Left            =   -74715
         TabIndex        =   24
         Top             =   3645
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   3598
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelConsol.frx":0552
      End
      Begin Sicmact.FlexEdit fgeAutoriza 
         Height          =   3225
         Left            =   -70920
         TabIndex        =   45
         Top             =   1965
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   5689
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cAreaCod-Area-Opc"
         EncabezadosAnchos=   "400-0-2700-0"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin MSComCtl2.DTPicker dtpResFec 
         Height          =   300
         Left            =   -73590
         TabIndex        =   46
         Top             =   1200
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   23527425
         CurrentDate     =   36783
         MaxDate         =   401768
         MinDate         =   36526
      End
      Begin Sicmact.FlexEdit fgeSelTpo 
         Height          =   3225
         Left            =   -74670
         TabIndex        =   47
         Top             =   1980
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   5689
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Opc"
         EncabezadosAnchos=   "400-0-2700-0"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3660
         Left            =   -73470
         TabIndex        =   52
         Top             =   1245
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   6456
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-PrecioProm-Sub Total-Opc"
         EncabezadosAnchos=   "400-0-1850-650-900-900-1000-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-C"
         FormatosEdit    =   "0-0-0-0-2-2-2-0"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeBSTotal 
         Height          =   705
         Left            =   -73470
         TabIndex        =   58
         Top             =   4590
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   1244
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-PrecioProm-Sub Total"
         EncabezadosAnchos=   "400-0-1850-650-900-900-1000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeEvaEco 
         Height          =   2220
         Left            =   -74520
         TabIndex        =   63
         Top             =   1080
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3916
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Precio-Total"
         EncabezadosAnchos=   "400-0-2500-650-900-900-900"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeEvaEcoPos 
         Height          =   855
         Left            =   -74520
         TabIndex        =   64
         Top             =   5505
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1508
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Nombre-Dirección"
         EncabezadosAnchos=   "400-0-2500-3200"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin Sicmact.FlexEdit fgeEvaEcoTot 
         Height          =   915
         Left            =   -74520
         TabIndex        =   62
         Top             =   2985
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1614
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "------Total"
         EncabezadosAnchos=   "400-0-2500-650-900-900-900"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2"
         AvanceCeldas    =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   -2147483624
      End
      Begin Sicmact.FlexEdit fgeEvaParTec 
         Height          =   1305
         Left            =   -74520
         TabIndex        =   66
         Top             =   3945
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2302
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Valor-Tipo"
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
      Begin Sicmact.FlexEdit fgeEvaParEco 
         Height          =   1305
         Left            =   -70875
         TabIndex        =   67
         Top             =   3945
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2302
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Valor-Tipo"
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
      Begin VB.Label lblEtiqueta 
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
         Height          =   210
         Index           =   0
         Left            =   -74415
         TabIndex        =   65
         Top             =   5295
         Width           =   1065
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Total :"
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
         Left            =   1260
         TabIndex        =   59
         Top             =   3030
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Autorización"
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
         Left            =   -70845
         TabIndex        =   51
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Número :"
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
         Index           =   2
         Left            =   -70875
         TabIndex        =   50
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha :"
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
         Left            =   -74565
         TabIndex        =   49
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo Proceso de Selección :"
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
         Index           =   9
         Left            =   -74610
         TabIndex        =   48
         Top             =   1740
         Width           =   2505
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Responsables"
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
         Index           =   25
         Left            =   -74490
         TabIndex        =   43
         Top             =   885
         Width           =   1395
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Cronograma :"
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
         Index           =   24
         Left            =   -74385
         TabIndex        =   42
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Total :"
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
         Index           =   14
         Left            =   5025
         TabIndex        =   41
         Top             =   3015
         Width           =   705
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Total :"
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
         Index           =   6
         Left            =   -72690
         TabIndex        =   40
         Top             =   3015
         Width           =   855
      End
      Begin VB.Label lblTotEco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5970
         TabIndex        =   39
         Top             =   2970
         Width           =   660
      End
      Begin VB.Label lblTotTec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2175
         TabIndex        =   38
         Top             =   2970
         Width           =   660
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
         Left            =   4230
         TabIndex        =   37
         Top             =   1035
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
         Left            =   285
         TabIndex        =   36
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo bases :"
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
         Index           =   7
         Left            =   -74355
         TabIndex        =   35
         Top             =   1335
         Width           =   1170
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Valor de referencia :"
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
         Index           =   15
         Left            =   -74370
         TabIndex        =   34
         Top             =   1950
         Width           =   2235
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Sistema de Adjudicación :"
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
         Index           =   16
         Left            =   -74385
         TabIndex        =   33
         Top             =   2580
         Width           =   2370
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Proveedores "
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
         Index           =   19
         Left            =   -74535
         TabIndex        =   32
         Top             =   4305
         Width           =   1245
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Postores "
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
         Index           =   20
         Left            =   -74370
         TabIndex        =   31
         Top             =   1170
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Postores "
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
         Index           =   21
         Left            =   -74610
         TabIndex        =   30
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label lblObserva 
         Caption         =   "Consulta"
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
         Left            =   -74595
         TabIndex        =   29
         Top             =   3405
         Width           =   1140
      End
      Begin VB.Label lblObserva 
         Caption         =   "Absolución"
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
         Left            =   -74580
         TabIndex        =   28
         Top             =   3405
         Width           =   1140
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Postores "
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
         Index           =   22
         Left            =   -74610
         TabIndex        =   27
         Top             =   1155
         Width           =   1485
      End
      Begin VB.Label lblObserva 
         Caption         =   "Observación"
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
         Index           =   2
         Left            =   -74595
         TabIndex        =   26
         Top             =   3405
         Width           =   1140
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Postores "
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
         Index           =   23
         Left            =   -74610
         TabIndex        =   25
         Top             =   1155
         Width           =   1485
      End
   End
   Begin Sicmact.FlexEdit fgeProceso 
      Height          =   4410
      Left            =   8625
      TabIndex        =   60
      Top             =   345
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   7779
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Seleccion-Resolución-Estado"
      EncabezadosAnchos=   "350-0-1200-1000"
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
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L"
      FormatosEdit    =   "0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      Appearance      =   0
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Procesos"
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
      Left            =   8685
      TabIndex        =   61
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmLogSelConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim paTecRango() As Currency
Dim paEcoRango() As Currency

Public Sub Inicio(ByVal psFormTpo As String)
    psFrmTpo = psFormTpo
    Me.Show 1
End Sub

Private Sub cmdAdq_Click()
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    Dim sSelTraNro As String, sActualiza As String
    Dim nSelNro As Long, nSelTraNro As Long
    Dim nResult As Integer
    
    'Verifica que siempre este por lo menos UNO
    If fgeProceso.TextMatrix(1, 1) = "" Then Exit Sub
    nSelNro = fgeProceso.TextMatrix(fgeProceso.Row, 1)
    If nSelNro = 0 Then Exit Sub
    
    If psFrmTpo = "2" Then
        'RECHAZO
        If MsgBox("¿ Estás seguro de Rechazar este Proceso de Selección " & vbCr & " cuya resolución es " & fgeProceso.TextMatrix(fgeProceso.Row, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoRechazado
            nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
            clsDMov.InsertaMovRef nSelTraNro, nSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoRechazado, sActualiza
            
            'Libera los requerimientos relacionados
            clsDMov.ActualizaReqDetMes nSelNro, sActualiza
            
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            
            If nResult = 0 Then
                cmdAdq.Enabled = False
                Call CargaProcesos
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    ElseIf psFrmTpo = "3" Then
        'DESIERTO
        If MsgBox("¿ Estás seguro de declarar Desierto este Proceso de Selección " & vbCr & " cuya resolución es " & fgeProceso.TextMatrix(fgeProceso.Row, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoDesierto
            nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
            clsDMov.InsertaMovRef nSelTraNro, nSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoDesierto, sActualiza
            
            'Libera los requerimientos relacionados
            clsDMov.ActualizaReqDetMes nSelNro, sActualiza
            
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            
            If nResult = 0 Then
                cmdAdq.Enabled = False
                
                Call CargaProcesos
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    ElseIf psFrmTpo = "4" Then
        'ACEPTACION
        If MsgBox("¿ Estás seguro de Aceptar este Proceso de Selección " & vbCr & " cuya resolución es " & fgeProceso.TextMatrix(fgeProceso.Row, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoAceptado
            nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
            clsDMov.InsertaMovRef nSelTraNro, nSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoAceptado, sActualiza
            
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            
            If nResult = 0 Then
                cmdAdq.Enabled = False
                Call CargaProcesos
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    ElseIf psFrmTpo = "5" Then
        'CONSENTIMIENTO
        If MsgBox("¿ Estás seguro de dar el consentimiento a este Proceso de Selección " & vbCr & " cuya resolución es " & fgeProceso.TextMatrix(fgeProceso.Row, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoConsentimiento
            nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
            clsDMov.InsertaMovRef nSelTraNro, nSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoConsentimiento, sActualiza
            
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            
            If nResult = 0 Then
                cmdAdq.Enabled = False
                Call CargaProcesos
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'''Private Sub fgeCotEco_OnRowChange(pnRow As Long, pnCol As Long)
'''    Dim clsDAdq As DLogAdquisi
'''    Dim rs As ADODB.Recordset
'''    Dim sSelCotNro As String
'''
'''    'ECONOMICA
'''    sSelCotNro = fgeCotEco.TextMatrix(fgeCotEco.Row, 1)
'''    'Muestra datos
'''    Set clsDAdq = New DLogAdquisi
'''    Set rs = New ADODB.Recordset
'''    Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegistroCotiza, sSelCotNro)
'''    If rs.RecordCount > 0 Then
'''        Set fgeBSEco.Recordset = rs
'''        fgeBSEco.AdicionaFila
'''        fgeBSEco.BackColorRow &HC0FFFF
'''        fgeBSEco.TextMatrix(fgeBSEco.Row, 2) = "T O T A L  R E F E R E N C I A L"
'''        fgeBSEco.TextMatrix(fgeBSEco.Row, 9) = Format(fgeBSEco.SumaRow(9), "#,##0.00")
'''    End If
'''    Set rs = Nothing
'''    Set clsDAdq = Nothing
'''End Sub

Private Sub fgeAbsolucion_OnRowChange(pnRow As Long, pnCol As Long)
rtfAbsolucion.Text = fgeAbsolucion.TextMatrix(pnRow, 5)
End Sub

Private Sub fgeConsulta_OnRowChange(pnRow As Long, pnCol As Long)
rtfConsulta.Text = fgeConsulta.TextMatrix(pnRow, 4)
End Sub

Private Sub fgeObservacion_OnRowChange(pnRow As Long, pnCol As Long)
rtfObservacion.Text = fgeObservacion.TextMatrix(pnRow, 6)
End Sub

Private Sub fgeParEco_OnRowChange(pnRow As Long, pnCol As Long)
    If Right(Trim(fgeParEco.TextMatrix(fgeParEco.Row, 4)), 1) = "3" Then
        fraParEcoRango.Visible = True
        CargaRango (2)
    Else
        fraParEcoRango.Visible = False
    End If
End Sub

Private Sub fgeParTec_OnRowChange(pnRow As Long, pnCol As Long)
    If Right(Trim(fgeParTec.TextMatrix(fgeParTec.Row, 4)), 1) = "3" Then
        fraParTecRango.Visible = True
        CargaRango (1)
    Else
        fraParTecRango.Visible = False
    End If
End Sub

Private Sub fgeProceso_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDReq As DLogRequeri
    Dim clsDAdq As DLogAdquisi
    Dim rsPro As ADODB.Recordset, rs As ADODB.Recordset
    Dim nSelNro As Integer
    Dim nCont As Integer, nCont2 As Integer
    Dim clsDGnral As DLogGeneral

    
    Set clsDGnral = New DLogGeneral
    Set clsDReq = New DLogRequeri
    'Verifica que siempre este por lo menos UNO
    If fgeProceso.TextMatrix(1, 1) = "" Then
        cmdAdq.Enabled = False
        Exit Sub
    End If
    
    If psFrmTpo = "2" Or psFrmTpo = "3" Or psFrmTpo = "4" Or psFrmTpo = "5" Then
        cmdAdq.Enabled = True
    End If
    'nSelNro = clsDGnral.GetnMovNro(fgeProceso.TextMatrix(fgeProceso.Row, 1))
    nSelNro = fgeProceso.TextMatrix(fgeProceso.Row, 1)
    'Actualiza Detalle de Proceso
    Set clsDAdq = New DLogAdquisi
    Set rsPro = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rsPro = clsDAdq.CargaSeleccion(SelUnRegistro, nSelNro)
    If rsPro.RecordCount > 0 Then
        With rsPro
            Call Limpiar
            txtResNro.Text = !cLogSelResNro
            dtpResFec.Value = Format(!dLogSelRes, gsFormatoFechaView)
            'lblSisAdj.Caption = !cConsDescripcion
            fgeSelTpo.AdicionaFila
            fgeSelTpo.TextMatrix(1, 1) = !nLogSelTpo
            fgeSelTpo.TextMatrix(1, 2) = !cConsDescSelTpo
            fgeAutoriza.AdicionaFila
            fgeAutoriza.TextMatrix(1, 1) = !cAreaCod
            fgeAutoriza.TextMatrix(1, 2) = !cAreaDescripcion
            'Carga Detalle
            Set rs = clsDReq.CargaSelDetalle(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgeBS.Recordset = rs
                'Total
                fgeBSTotal.BackColorRow &HC0FFFF
                fgeBSTotal.TextMatrix(1, 0) = "="
                fgeBSTotal.TextMatrix(1, 2) = "T O T A L "
                fgeBSTotal.TextMatrix(1, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
            End If
            
            If !nLogSelMoneda = gMonedaNacional Then
                optMoneda(0).Value = True
            ElseIf !nLogSelMoneda = gMonedaExtranjera Then
                optMoneda(1).Value = True
            Else
                optMoneda(0).Value = False
                optMoneda(1).Value = False
            End If
            txtTipCambio.Text = Format(!nLogSelTipCambio, "##.000")
            txtCostoBase.Text = Format(!nLogSelCostoBase, "#0.00")
            txtValReferencia.Text = Format(!nLogSelValorRefe, "#0.00")
            
            fgeSisAdj.AdicionaFila
            fgeSisAdj.TextMatrix(1, 1) = !nLogSelSisAdj
            fgeSisAdj.TextMatrix(1, 2) = !cConsDescSisAdj
            
            'Muestra Comite
            Set rs = clsDAdq.CargaSelComite(nSelNro)
            If rs.RecordCount > 0 Then Set fgeComite.Recordset = rs
            
            'Muestra Cronograma
            Set rs = clsDAdq.CargaSelCronograma(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgeCronograma.Recordset = rs
            End If
            
            'Muestra parámetros ingresados
            Set rs = clsDAdq.CargaSelParametro(nSelNro, 1)
            If rs.RecordCount > 0 Then
                fgeParTec.lbEditarFlex = True
                Set fgeParTec.Recordset = rs
                'Rango
                ReDim paTecRango(3, fgeParTec.Rows, fgeParTecRango.Rows)
                Set rs = clsDAdq.CargaSelParDetalle(nSelNro, 1)
                If rs.RecordCount > 0 Then
                    For nCont = 1 To fgeParTec.Rows - 1
                        If Right(Trim(fgeParTec.TextMatrix(nCont, 4)), 1) = 3 Then
                            rs.MoveFirst
                            nCont2 = 0
                            Do While Not rs.EOF
                                If rs!nLogSelParNro = nCont Then
                                    nCont2 = nCont2 + 1
                                    If rs!nLogSelParDetNro = nCont2 Then
                                        paTecRango(3, nCont, rs!nLogSelParDetNro) = rs!nLogSelParDetPuntaje
                                        paTecRango(1, nCont, rs!nLogSelParDetNro) = rs!nLogSelParDetIni
                                        paTecRango(2, nCont, rs!nLogSelParDetNro) = rs!nLogSelParDetFin
                                    End If
                                End If
                                rs.MoveNext
                            Loop
                        End If
                    Next
                End If
                Call fgeParTec_OnRowChange(fgeParTec.Row, fgeParTec.Col)
                lblTotTec.Caption = fgeParTec.SumaRow(3)
            End If
            Set rs = clsDAdq.CargaSelParametro(nSelNro, 2)
            If rs.RecordCount > 0 Then
                fgeParEco.lbEditarFlex = True
                Set fgeParEco.Recordset = rs
                'Rango
                ReDim paEcoRango(3, fgeParEco.Rows, fgeParEcoRango.Rows)
                Set rs = clsDAdq.CargaSelParDetalle(nSelNro, 2)
                If rs.RecordCount > 0 Then
                    For nCont = 1 To fgeParEco.Rows - 1
                        If Right(Trim(fgeParEco.TextMatrix(nCont, 4)), 1) = 3 Then
                            rs.MoveFirst
                            nCont2 = 0
                            Do While Not rs.EOF
                                If rs!nLogSelParNro = nCont Then
                                    nCont2 = nCont2 + 1
                                    If rs!nLogSelParDetNro = nCont2 Then
                                        paEcoRango(3, nCont, rs!nLogSelParDetNro) = rs!nLogSelParDetPuntaje
                                        paEcoRango(1, nCont, rs!nLogSelParDetNro) = rs!nLogSelParDetIni
                                        paEcoRango(2, nCont, rs!nLogSelParDetNro) = rs!nLogSelParDetFin
                                    End If
                                End If
                                rs.MoveNext
                            Loop
                        End If
                    Next
                End If
                Call fgeParEco_OnRowChange(fgeParEco.Row, fgeParEco.Col)
                lblTotEco.Caption = fgeParEco.SumaRow(3)
            End If
            'Muestra Publicación
            Set rs = clsDAdq.CargaSelPublica(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgePublica.Recordset = rs
            End If
            'Muestra Cotizaciones
            Set rs = clsDAdq.CargaSelDetalle(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgeCot.Recordset = rs
            End If
            Set rs = clsDAdq.CargaSelCotiza(SelCotPersona, nSelNro)
            If rs.RecordCount > 0 Then
                Set fgePro.Recordset = rs
                Call CargaSelBSPro
            End If
            
            'Muestra Postores
            Set rs = clsDAdq.CargaSelPostor(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgePostor.Recordset = rs
            End If
            'Muestra Consulta
            Set rs = clsDAdq.CargaSelPostor(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgeConsulta.Recordset = rs
            End If
            'Muestra Absolucion
            Set rs = clsDAdq.CargaSelPostor(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgeAbsolucion.Recordset = rs
            End If
            
            'Muestra Observación
            Set rs = clsDAdq.CargaSelPostor(nSelNro)
            If rs.RecordCount > 0 Then
                Set fgeObservacion.Recordset = rs
            End If
            
            'EVALUACION
            If !nLogSelCotNro <> 0 Then
                Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegistro, nSelNro, !nLogSelCotNro)
                If rs.RecordCount > 0 Then
                    Set fgeEvaEco.Recordset = rs
                    fgeEvaEcoTot.BackColorRow &HC0FFFF
                    fgeEvaEcoTot.TextMatrix(1, 0) = "="
                    fgeEvaEcoTot.TextMatrix(1, 2) = "T O T A L "
                    fgeEvaEcoTot.TextMatrix(1, 6) = Format(fgeEvaEco.SumaRow(6), "#,##0.00")
                End If
                
                Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, nSelNro, 1, !nLogSelCotNro)
                If rs.RecordCount > 0 Then Set fgeEvaParTec.Recordset = rs
                
                Set rs = clsDAdq.CargaSelCotPar(SelCotParRegistro, nSelNro, 2, !nLogSelCotNro)
                If rs.RecordCount > 0 Then Set fgeEvaParEco.Recordset = rs
            
                Set rs = clsDAdq.CargaSelCotiza(SelCotPersona, nSelNro, !nLogSelCotNro)
                If rs.RecordCount > 0 Then Set fgeEvaEcoPos.Recordset = rs
            End If
        End With
    End If
    Set clsDAdq = Nothing
    Set rs = Nothing
    Set rsPro = Nothing
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    sstSeleccion.TabVisible(12) = False
    If psFrmTpo = "1" Then
        sstSeleccion.TabVisible(12) = True
        Me.Caption = "Consulta de Proceso de Selección"
    ElseIf psFrmTpo = "2" Then
        Me.Caption = "Cancelación de Proceso de Selección"
        cmdAdq.Caption = "&Rechazar"
        cmdAdq.Visible = True
    ElseIf psFrmTpo = "3" Then
        Me.Caption = "Proceso de Selección Desierto"
        cmdAdq.Caption = "&Desierto"
        cmdAdq.Visible = True
    ElseIf psFrmTpo = "4" Then
        sstSeleccion.TabVisible(12) = True
        Me.Caption = "Aceptación del Proceso de Selección"
        cmdAdq.Caption = "&Aceptar"
        cmdAdq.Visible = True
    ElseIf psFrmTpo = "5" Then
        sstSeleccion.TabVisible(12) = True
        Me.Caption = "Consentimiento Aceptación del Proceso de Selección"
        cmdAdq.Caption = "&Aceptar"
        cmdAdq.Visible = True
    Else
        MsgBox "Tipo de formulario no reconocido", vbInformation, " Aviso"
        Exit Sub
    End If
    Call CargaProcesos
End Sub

Private Sub CargaProcesos()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    If psFrmTpo = "1" Then
        'CONSULTA
        Set rs = clsDAdq.CargaSeleccion(SelTodosGnral)
    ElseIf psFrmTpo = "2" Then
        'CANCELACION
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstadoParaCancelar)
    ElseIf psFrmTpo = "3" Then
        'DESIERTO
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstadoParaDesierto)
    ElseIf psFrmTpo = "4" Then
        'ACEPTACION
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstadoParaAceptar)
    ElseIf psFrmTpo = "5" Then
        'CONSENTIMIENTO
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstadoParaConsenti)
    End If
    If rs.RecordCount > 0 Then
        Set fgeProceso.Recordset = rs
        Call fgeProceso_OnRowChange(fgeProceso.Row, fgeProceso.Col)
    Else
        fgeProceso.Clear
        fgeProceso.FormaCabecera
        fgeProceso.Rows = 2
        fgeProceso.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub


Private Sub CargaSelBSPro()
    Dim clsDProv As DLogProveedor
    Dim sPersCod As String, sBSCod As String
    Dim rs As ADODB.Recordset
    Dim nCont As Integer, nCont2 As Integer, pColCot As Integer
    
    pColCot = 5
    Set clsDProv = New DLogProveedor
    Set rs = New ADODB.Recordset
    
    fgeSel.TextMatrix(1, 0) = "="
    fgeSel.TextMatrix(1, 2) = "SELECCION DE PROVEEDORES"
    fgeSel.Enabled = False
    
    For nCont = 1 To fgePro.Rows - 1
        fgeCot.Cols = fgeCot.Cols + 1
        fgeCot.ColWidth(nCont + pColCot) = 400
        fgeCot.TextMatrix(0, nCont + pColCot) = "P" & Format(nCont, "00")
        
        fgeSel.Cols = fgeCot.Cols
        fgeSel.ColWidth(nCont + pColCot) = 400
        fgeSel.TextMatrix(0, nCont + pColCot) = "P" & Format(nCont, "00")
        fgeSel.TextMatrix(1, nCont + pColCot) = "  X"
        
        sPersCod = fgePro.TextMatrix(nCont, 1)
        'Carga Productos del Proveedor
        Set rs = clsDProv.CargaProveedorBS(ProBSBienServicio, sPersCod)
        If Not rs.EOF Then
            Do While Not rs.EOF
                For nCont2 = 1 To fgeCot.Rows - 1
                    sBSCod = fgeCot.TextMatrix(nCont2, 1)
                    If sBSCod = rs!Codigo Then
                        fgeCot.TextMatrix(nCont2, nCont + pColCot) = "  X"
                        Exit For
                    End If
                Next
                rs.MoveNext
            Loop
        End If
    Next
End Sub

Private Sub Limpiar()
    Dim nCont As Integer
    dtpResFec.Value = gdFecSis
    txtResNro.Text = ""
    fgeAutoriza.Clear
    fgeAutoriza.FormaCabecera
    fgeAutoriza.Rows = 2
    fgeSelTpo.Clear
    fgeSelTpo.FormaCabecera
    fgeSelTpo.Rows = 2
    
    fgeComite.Clear
    fgeComite.FormaCabecera
    fgeComite.Rows = 2
        
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeBSTotal.TextMatrix(1, 6) = ""
        
    txtCostoBase.Text = ""
    txtValReferencia.Text = ""
    txtTipCambio.Text = ""
    lblTotEco.Caption = ""
    lblTotTec.Caption = ""
    fgeSisAdj.Clear
    fgeSisAdj.FormaCabecera
    fgeSisAdj.Rows = 2
    fgeParTec.Clear
    fgeParTec.FormaCabecera
    fgeParTec.Rows = 2
    fgeParEco.Clear
    fgeParEco.FormaCabecera
    fgeParEco.Rows = 2
    fraParEcoRango.Visible = False
    fraParTecRango.Visible = False
    fgeParTecRango.Clear
    fgeParTecRango.FormaCabecera
    fgeParTecRango.Rows = 5
    fgeParEcoRango.Clear
    fgeParEcoRango.FormaCabecera
    fgeParEcoRango.Rows = 5
    fgeCronograma.Clear
    fgeCronograma.FormaCabecera
    fgeCronograma.Rows = 2
        
    fgePublica.Clear
    fgePublica.FormaCabecera
    fgePublica.Rows = 2
        
    fgeCot.Clear
    fgeCot.FormaCabecera
    fgeCot.Rows = 2
    fgeSel.Clear
    fgeSel.FormaCabecera
    fgeSel.Rows = 2
    fgePro.Clear
    fgePro.FormaCabecera
    fgePro.Rows = 2
        
    fgePostor.Clear
    fgePostor.FormaCabecera
    fgePostor.Rows = 2
        
    fgeConsulta.Clear
    fgeConsulta.FormaCabecera
    fgeConsulta.Rows = 2
        
    fgeAbsolucion.Clear
    fgeAbsolucion.FormaCabecera
    fgeAbsolucion.Rows = 2
    
    fgeEvaEco.Clear
    fgeEvaEco.FormaCabecera
    fgeEvaEco.Rows = 2
    fgeEvaEcoPos.Clear
    fgeEvaEcoPos.FormaCabecera
    fgeEvaEcoPos.Rows = 2
    fgeEvaEcoTot.Clear
    fgeEvaEcoTot.FormaCabecera
    fgeEvaEcoTot.Rows = 2
    fgeEvaParTec.Clear
    fgeEvaParTec.FormaCabecera
    fgeEvaParTec.Rows = 2
    fgeEvaParEco.Clear
    fgeEvaParEco.FormaCabecera
    fgeEvaParEco.Rows = 2
End Sub

Private Sub CargaRango(ByVal Index As Integer)
Dim nCont As Integer
Dim nCol As Integer
If Index = 1 Then
    'Rango Técnico
    If paTecRango(1, fgeParTec.Row, 1) > 0 Or paTecRango(2, fgeParTec.Row, 1) > 0 Then
        fgeParTecRango.Rows = UBound(paTecRango, 3)
        For nCont = 1 To fgeParTecRango.Rows - 1
            fgeParTecRango.TextMatrix(nCont, 0) = nCont
        Next
        For nCol = 1 To fgeParTecRango.Cols - 1
            For nCont = 1 To UBound(paTecRango, 3) - 1
                fgeParTecRango.TextMatrix(nCont, nCol) = paTecRango(nCol, fgeParTec.Row, nCont)
            Next
        Next
    Else
        fgeParTecRango.Clear
        fgeParTecRango.FormaCabecera
        fgeParTecRango.Rows = 5
        For nCont = 1 To fgeParTecRango.Rows - 1
            fgeParTecRango.TextMatrix(nCont, 0) = nCont
        Next
    End If
Else
    'Rango Económico
    If paEcoRango(1, fgeParEco.Row, 1) > 0 Or paEcoRango(2, fgeParEco.Row, 1) > 0 Then
        fgeParEcoRango.Rows = UBound(paEcoRango, 3)
        For nCont = 1 To fgeParEcoRango.Rows - 1
            fgeParEcoRango.TextMatrix(nCont, 0) = nCont
        Next
        For nCol = 1 To fgeParEcoRango.Cols - 1
            For nCont = 1 To UBound(paEcoRango, 3) - 1
                fgeParEcoRango.TextMatrix(nCont, nCol) = paEcoRango(nCol, fgeParEco.Row, nCont)
            Next
        Next
    Else
        fgeParEcoRango.Clear
        fgeParEcoRango.FormaCabecera
        fgeParEcoRango.Rows = 5
        For nCont = 1 To fgeParEcoRango.Rows - 1
            fgeParEcoRango.TextMatrix(nCont, 0) = nCont
        Next
    End If
End If
End Sub

