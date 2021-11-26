VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogSelInicio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5685
   ClientLeft      =   840
   ClientTop       =   2640
   ClientWidth     =   9840
   Icon            =   "frmLogSelInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   3150
      TabIndex        =   9
      Top             =   5160
      Width           =   1290
   End
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   1650
      TabIndex        =   8
      Top             =   5160
      Width           =   1290
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   4650
      TabIndex        =   5
      Top             =   5160
      Width           =   1290
   End
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   6135
      TabIndex        =   4
      Top             =   5160
      Width           =   1290
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8085
      TabIndex        =   0
      Top             =   5160
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   5220
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin TabDlg.SSTab sstSeleccion 
      Height          =   4305
      Left            =   135
      TabIndex        =   3
      Top             =   705
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   7594
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Resolución"
      TabPicture(0)   =   "frmLogSelInicio.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraResolu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comité"
      TabPicture(1)   =   "frmLogSelInicio.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgeComite"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblEtiqueta(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Bases"
      TabPicture(2)   =   "frmLogSelInicio.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fgeBS"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraBase"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Parámetros"
      TabPicture(3)   =   "frmLogSelInicio.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgeParTec"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fgeParEco"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblComenta"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblEtiqueta(14)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblEtiqueta(13)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblTotEco"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblTotTec"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lblEtiqueta(11)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblEtiqueta(10)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Publicación"
      TabPicture(4)   =   "frmLogSelInicio.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdPubli(0)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdPubli(1)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fgePublica"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lblEtiqueta(12)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CommandButton cmdPubli 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -71715
         TabIndex        =   39
         Top             =   3825
         Width           =   1155
      End
      Begin VB.CommandButton cmdPubli 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -70140
         TabIndex        =   38
         Top             =   3840
         Width           =   1155
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3330
         Left            =   -71730
         TabIndex        =   26
         Top             =   825
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   5874
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-PrecioProm-Sub Total"
         EncabezadosAnchos=   "400-0-2000-650-900-900-1000"
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
         RowHeight0      =   285
      End
      Begin VB.Frame fraBase 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   -74895
         TabIndex        =   19
         Top             =   345
         Width           =   9390
         Begin VB.TextBox txtCostoBase 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1470
            TabIndex        =   20
            Top             =   180
            Width           =   1305
         End
         Begin Sicmact.FlexEdit fgeAdjudica 
            Height          =   2130
            Left            =   60
            TabIndex        =   21
            Top             =   1680
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   3757
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-Codigo-Descripción-Ok"
            EncabezadosAnchos=   "400-0-2000-350"
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
            ListaControles  =   "0-0-0-4"
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
            RowHeight0      =   285
         End
         Begin Sicmact.TxtBuscar txtAdqNro 
            Height          =   285
            Left            =   4890
            TabIndex        =   22
            Top             =   150
            Width           =   2895
            _ExtentX        =   5106
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
            Left            =   225
            TabIndex        =   25
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Sistema Adjudicación :"
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
            Left            =   105
            TabIndex        =   24
            Top             =   1380
            Width           =   1950
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Adquisición :"
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
            Index           =   8
            Left            =   3165
            TabIndex        =   23
            Top             =   210
            Width           =   1425
         End
      End
      Begin VB.Frame fraResolu 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3840
         Left            =   105
         TabIndex        =   10
         Top             =   345
         Width           =   9420
         Begin VB.TextBox txtResNro 
            Height          =   285
            Left            =   5580
            MaxLength       =   20
            TabIndex        =   12
            Top             =   120
            Width           =   1890
         End
         Begin Sicmact.FlexEdit fgeAutoriza 
            Height          =   2790
            Left            =   180
            TabIndex        =   11
            Top             =   900
            Width           =   9030
            _ExtentX        =   15928
            _ExtentY        =   4921
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-cAreaCod-Area-cPersCod-Nombre-Ok"
            EncabezadosAnchos=   "400-0-3000-0-3000-350"
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
            ListaControles  =   "0-0-0-0-0-4"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "Item"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   285
         End
         Begin MSComCtl2.DTPicker dtpResFec 
            Height          =   300
            Left            =   1230
            TabIndex        =   13
            Top             =   150
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            Format          =   49545217
            CurrentDate     =   36783
            MaxDate         =   401768
            MinDate         =   36526
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
            Left            =   255
            TabIndex        =   16
            Top             =   210
            Width           =   780
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
            Left            =   4590
            TabIndex        =   15
            Top             =   165
            Width           =   870
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
            Left            =   210
            TabIndex        =   14
            Top             =   660
            Width           =   1245
         End
      End
      Begin Sicmact.FlexEdit fgeComite 
         Height          =   3300
         Left            =   -74715
         TabIndex        =   17
         Top             =   735
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   5821
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cAreaCod-Area-cPersCod-Nombre-Ok"
         EncabezadosAnchos=   "400-0-3500-0-4000-350"
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
         ListaControles  =   "0-0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeParTec 
         Height          =   1965
         Left            =   -74790
         TabIndex        =   29
         Top             =   735
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   3466
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje"
         EncabezadosAnchos=   "400-0-2800-800"
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
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-R"
         FormatosEdit    =   "0-0-0-3"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeParEco 
         Height          =   1965
         Left            =   -70185
         TabIndex        =   30
         Top             =   735
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   3466
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje"
         EncabezadosAnchos=   "400-0-2800-800"
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
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-R"
         FormatosEdit    =   "0-0-0-3"
         CantDecimales   =   0
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgePublica 
         Height          =   2955
         Left            =   -74715
         TabIndex        =   32
         Top             =   810
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   5212
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha Inicial-Fecha Final"
         EncabezadosAnchos=   "400-1700-3500-1300-1300"
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
         ColumnasAEditar =   "X-1-X-3-4"
         ListaControles  =   "0-1-0-2-2"
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
         TipoBusPersona  =   1
      End
      Begin VB.Label lblComenta 
         Caption         =   "Nota."
         ForeColor       =   &H8000000D&
         Height          =   750
         Left            =   -74565
         TabIndex        =   37
         Top             =   3240
         Width           =   4305
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
         Left            =   -67860
         TabIndex        =   36
         Top             =   2745
         Width           =   855
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
         Index           =   13
         Left            =   -72480
         TabIndex        =   35
         Top             =   2745
         Width           =   855
      End
      Begin VB.Label lblTotEco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -66975
         TabIndex        =   34
         Top             =   2685
         Width           =   840
      End
      Begin VB.Label lblTotTec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -71595
         TabIndex        =   33
         Top             =   2685
         Width           =   840
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Publicaciones :"
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
         Index           =   12
         Left            =   -74625
         TabIndex        =   31
         Top             =   570
         Width           =   1530
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
         Left            =   -70110
         TabIndex        =   28
         Top             =   465
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
         Left            =   -74685
         TabIndex        =   27
         Top             =   465
         Width           =   1425
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Areas"
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
         Left            =   -74595
         TabIndex        =   18
         Top             =   480
         Width           =   1245
      End
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
      Index           =   5
      Left            =   330
      TabIndex        =   7
      Top             =   405
      Width           =   870
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
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
      Left            =   360
      TabIndex        =   2
      Top             =   90
      Width           =   555
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   45
      Width           =   3705
   End
End
Attribute VB_Name = "frmLogSelInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim pbNuevo As Boolean
Dim pnTotPar As Currency

Public Sub Inicio(ByVal psFormTpo As String)
psFrmTpo = psFormTpo
Me.Show 1
End Sub

Private Sub cmdAdq_Click(Index As Integer)
    Dim sSelNro As String, sSelTraNro As String, sAdqNro As String, sActualiza As String
    Dim sAreaCod As String, sPersCod As String, sResNro As String, sAdjuCod As String
    Dim nCodPar As Integer
    Dim nCont As Integer, nSum As Integer, nResult As Integer
    Dim nTotEco As Currency, nTotTec As Currency
    Dim nCostoBase As Currency
    Dim dPubIni As Date, dPubFin As Date
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    
    Select Case Index
        Case 0:
            'NUEVO
            pbNuevo = True
            'sstSeleccion.Enabled = True
            fraResolu.Enabled = True
            txtSelNro.Enabled = False
            Set clsDGnral = New DLogGeneral
            txtSelNro.Text = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            Call Limpiar
            fgeAutoriza.lbEditarFlex = True
            'Carga los Flex
            Call CargaAreaAutoriza
            cmdAdq(0).Enabled = False
            cmdAdq(1).Enabled = False
            cmdAdq(2).Enabled = True
            cmdAdq(3).Enabled = True
        Case 1:
            'EDITAR
            If Len(txtSelNro.Text) > 0 Then
                pbNuevo = False
                txtSelNro.Enabled = False
                'sstSeleccion.Enabled = True
                fraResolu.Enabled = True
                cmdAdq(0).Enabled = False
                cmdAdq(1).Enabled = False
                cmdAdq(2).Enabled = True
                cmdAdq(3).Enabled = True
                
            End If
        Case 2:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                pbNuevo = True
                Call Limpiar
                'sstSeleccion.Enabled = False
                fraResolu.Enabled = False
                txtSelNro.Enabled = True
                txtSelNro.Text = ""
                cmdAdq(0).Enabled = True
                cmdAdq(1).Enabled = False
                cmdAdq(2).Enabled = False
                cmdAdq(3).Enabled = False
            End If
        Case 3:
            'GRABAR
            sSelNro = txtSelNro.Text
            If psFrmTpo = "1" Then
                'Inicio
                For nCont = 1 To fgeAutoriza.Rows - 1
                    If fgeAutoriza.TextMatrix(nCont, 5) = "." Then
                        sAreaCod = fgeAutoriza.TextMatrix(nCont, 1)
                        sPersCod = fgeAutoriza.TextMatrix(nCont, 3)
                        nSum = nSum + 1
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta determinar la persona responsable", vbInformation, "Aviso"
                    Exit Sub
                End If
                
                sResNro = Trim(txtResNro.Text)
                If Trim(sResNro) = "" Then
                    MsgBox "Falta ingresar el número de Resolución", vbInformation, " Aviso "
                    Exit Sub
                End If
                If sAreaCod = "" Or sPersCod = "" Then
                    MsgBox "Determine el Area y Persona responsable", vbInformation, "Aviso"
                    Exit Sub
                End If
            ElseIf psFrmTpo = "2" Then
                'Comité
                For nCont = 1 To fgeComite.Rows - 1
                    If fgeComite.TextMatrix(nCont, 5) = "." Then
                        nSum = nSum + 1
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta determinar el comite responsable", vbInformation, "Aviso"
                    Exit Sub
                End If
            ElseIf psFrmTpo = "3" Then
                'Base
                sAdqNro = txtAdqNro.Text
                nCostoBase = CCur(IIf(txtCostoBase.Text = "", 0, txtCostoBase.Text))
                If nCostoBase <= 0 Then
                    MsgBox "Falta ingresar el Costo Base", vbInformation, " Aviso "
                    Exit Sub
                End If
                If sAdqNro = "" Then
                    MsgBox "Falta determinar la adquisición", vbInformation, " Aviso "
                    Exit Sub
                End If
                For nCont = 1 To fgeAdjudica.Rows - 1
                    If fgeAdjudica.TextMatrix(nCont, 3) = "." Then
                        nSum = nSum + 1
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta determinar el sistema de adjudicación", vbInformation, "Aviso"
                    Exit Sub
                End If
            ElseIf psFrmTpo = "4" Then
                'Parametros
                nTotEco = CCur(IIf(lblTotEco.Caption = "", 0, lblTotEco.Caption))
                nTotTec = CCur(IIf(lblTotTec.Caption = "", 0, lblTotTec.Caption))
                If nTotTec <> pnTotPar Then
                    MsgBox "Suma de parámetros técnicos debe ser igual a " & pnTotPar, vbInformation, " Aviso"
                    Exit Sub
                End If
                If nTotEco <> pnTotPar Then
                    MsgBox "Suma de parámetros económicos debe ser igual a " & pnTotPar, vbInformation, " Aviso"
                    Exit Sub
                End If
            
            ElseIf psFrmTpo = "5" Then
                'Publicaciones
                For nCont = 1 To fgePublica.Rows - 1
                    If fgePublica.TextMatrix(nCont, 1) = "" Then
                        MsgBox "Falta el publicador en el Item : " & nCont, vbInformation, "Aviso"
                        Exit Sub
                    End If
                    If fgePublica.TextMatrix(nCont, 3) = "" Or fgePublica.TextMatrix(nCont, 4) = "" Then
                        MsgBox "Falta el determinar las fechas en el Item : " & nCont, vbInformation, "Aviso"
                        Exit Sub
                    End If
                Next
            Else
                MsgBox "Tipo de Formulario ¡ No Reconocido !", vbInformation, " Aviso "
                Exit Sub
            End If
            
            If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                If psFrmTpo = "1" Then
                    'INICIO
                    If pbNuevo Then
                        sSelTraNro = sSelNro
                    Else
                        Set clsDGnral = New DLogGeneral
                        sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDGnral = Nothing
                    End If
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    If pbNuevo Then
                        'Grabación de MOV - MOVREF
                        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelRegistro)), "", Trim(Str(gLogSelEstadoInicioRes))
                        clsDMov.InsertaMovRef sSelTraNro, sSelNro
                        
                        'Inserta LogAdquisicion
                        clsDMov.InsertaSeleccion sSelNro, dtpResFec.Value, sResNro, _
                            sAreaCod, sPersCod, sActualiza
                    Else
                        'Grabación de MOV -MOVREF
                        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelModifica)), "", Trim(Str(gLogSelEstadoInicioRes))
                        clsDMov.InsertaMovRef sSelTraNro, sSelNro
                        
                        'Actualiza LogSelección
                        clsDMov.ActualizaSeleccion sSelNro, dtpResFec.Value, sResNro, _
                            sAreaCod, sPersCod, sActualiza
                    End If
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Enabled = True
                        cmdAdq(1).Enabled = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        fraResolu.Enabled = False
                        txtSelNro.Enabled = True
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                
                ElseIf psFrmTpo = "2" Then
                    'COMITE
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoComite))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", "", "", _
                        sActualiza, gLogSelEstadoComite
                    
                    For nCont = 1 To fgeComite.Rows - 1
                        If fgeComite.TextMatrix(nCont, 5) = "." Then
                            'Inserta LogSelComite
                            sAreaCod = fgeComite.TextMatrix(nCont, 1)
                            sPersCod = fgeComite.TextMatrix(nCont, 3)
                            clsDMov.InsertaSelComite sSelNro, sAreaCod, sPersCod, _
                                sActualiza
                        End If
                    Next
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(1).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        txtSelNro.Enabled = True
                        fgeComite.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf psFrmTpo = "3" Then
                    'BASE
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoBases))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    For nCont = 1 To fgeAdjudica.Rows - 1
                        If fgeAdjudica.TextMatrix(nCont, 3) = "." Then
                            sAdjuCod = fgeAdjudica.TextMatrix(nCont, 1)
                            Exit For
                        End If
                    Next
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionBase sSelNro, nCostoBase, sAdqNro, sAdjuCod, _
                        sActualiza
                    
                    clsDMov.ActualizaAdquisicion sAdqNro, gLogAdqEstadoBase, sActualiza
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(1).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        txtSelNro.Enabled = True
                        fraBase.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf psFrmTpo = "4" Then
                    'PARAMETRO
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoParametro))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", "", "", _
                        sActualiza, gLogSelEstadoParametro
                        
                    For nCont = 1 To fgeParTec.Rows - 1
                        nCodPar = Val(fgeParTec.TextMatrix(nCont, 1))
                        nSum = CCur(IIf(fgeParTec.TextMatrix(nCont, 3) = "", 0, fgeParTec.TextMatrix(nCont, 3)))
                        If nSum > 0 Then
                            clsDMov.InsertaSelParametro sSelNro, "1", nCodPar, nSum, sActualiza
                        End If
                    Next
                    
                    For nCont = 1 To fgeParEco.Rows - 1
                        nCodPar = Val(fgeParEco.TextMatrix(nCont, 1))
                        nSum = CCur(IIf(fgeParEco.TextMatrix(nCont, 3) = "", 0, fgeParEco.TextMatrix(nCont, 3)))
                        If nSum > 0 Then
                            clsDMov.InsertaSelParametro sSelNro, "2", nCodPar, nSum, sActualiza
                        End If
                    Next
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgeParTec.lbEditarFlex = False
                        fgeParEco.lbEditarFlex = False
                        cmdAdq(0).Visible = False
                        cmdAdq(1).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        txtSelNro.Enabled = True
                        fraBase.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf psFrmTpo = "5" Then
                    'PUBLICACION
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoPublicacion))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", "", "", _
                        sActualiza, gLogSelEstadoPublicacion
                    
                    For nCont = 1 To fgePublica.Rows - 1
                        sPersCod = fgePublica.TextMatrix(nCont, 1)
                        dPubIni = fgePublica.TextMatrix(nCont, 3)
                        dPubFin = fgePublica.TextMatrix(nCont, 4)
                        clsDMov.InsertaSelPublica sSelNro, sPersCod, dPubIni, dPubFin, sActualiza
                    Next
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgePublica.lbEditarFlex = False
                        cmdAdq(0).Visible = False
                        cmdAdq(1).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        cmdPubli(0).Enabled = False
                        cmdPubli(1).Enabled = False
                        txtSelNro.Enabled = True
                        fraBase.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                End If
            End If
        Case Else
            MsgBox "Comando no reconocido", vbInformation, " Aviso"
    End Select
End Sub

Private Sub cmdPubli_Click(Index As Integer)
    Dim nBSRow As Integer
    'Botones de comandos del detalle de bienes/servicios
    If Index = 0 Then
        'Agregar en Flex
        fgePublica.AdicionaFila
        fgePublica.SetFocus
    ElseIf Index = 1 Then
        'Eliminar en Flex
        nBSRow = fgePublica.Row
        If MsgBox("¿ Estás seguro de eliminar " & fgePublica.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            fgePublica.EliminaFila nBSRow
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgeParEco_OnCellChange(pnRow As Long, pnCol As Long)
    lblTotEco.Caption = fgeParEco.SumaRow(3)
End Sub
Private Sub fgeParTec_OnCellChange(pnRow As Long, pnCol As Long)
    lblTotTec.Caption = fgeParTec.SumaRow(3)
End Sub

Private Sub fgePublica_OnCellChange(pnRow As Long, pnCol As Long)
Dim dIni As Date, dFin As Date
If pnCol = 3 Then
    fgePublica.TextMatrix(pnRow, 4) = ""
ElseIf pnCol = 4 Then
    If fgePublica.TextMatrix(pnRow, 3) = "" Then
        fgePublica.TextMatrix(pnRow, 4) = ""
        MsgBox "Fecha ingresar fecha inicial", vbInformation, " Aviso"
        Exit Sub
    End If
    If fgePublica.TextMatrix(pnRow, 4) = "" Then
        fgePublica.TextMatrix(pnRow, 3) = ""
        Exit Sub
    End If
    dIni = fgePublica.TextMatrix(pnRow, 3)
    dFin = fgePublica.TextMatrix(pnRow, 4)
    If DateDiff("d", dIni, dFin) < 0 Then
        fgePublica.TextMatrix(pnRow, 4) = ""
        MsgBox "Fecha debe ser mayor a la inicial", vbInformation, " Aviso"
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
    Dim clsDGnral As DLogGeneral
    
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    dtpResFec.Value = gdFecSis
    pbNuevo = True
    If psFrmTpo = "1" Then
        'INICIO
        Me.Caption = "Inicio del Proceso de Selección"
        sstSeleccion.TabVisible(1) = False
        sstSeleccion.TabVisible(2) = False
        sstSeleccion.TabVisible(3) = False
        sstSeleccion.TabVisible(4) = False
        cmdAdq(0).Enabled = True
        
        'Carga FLEX de txtNroSel
        Call CargaTxtSelNro
    ElseIf psFrmTpo = "2" Then
        'COMITE
        Me.Caption = "Ingreso de Comite Especial u Organo Encargado"
        sstSeleccion.TabVisible(2) = False
        sstSeleccion.TabVisible(3) = False
        sstSeleccion.TabVisible(4) = False
        cmdAdq(0).Visible = False
        cmdAdq(1).Visible = False
        'Carga FLEX de txtNroSel
        Call CargaTxtSelNro
    ElseIf psFrmTpo = "3" Then
        'BASE
        Me.Caption = "Registro de Base, valor referencial y sistema de adjudicación"
        sstSeleccion.TabVisible(3) = False
        sstSeleccion.TabVisible(4) = False
        cmdAdq(0).Visible = False
        cmdAdq(1).Visible = False
        fgeComite.EncabezadosAnchos = "400-0-3500-0-4000-0"
        'Carga FLEX de txtNroSel
        Call CargaTxtSelNro
    ElseIf psFrmTpo = "4" Then
        'PARAMETRO
        Me.Caption = "Registro de Parámetros de Evaluación"
        sstSeleccion.TabVisible(4) = False
        fgeComite.EncabezadosAnchos = "400-0-3500-0-4000-0"
        fgeAdjudica.EncabezadosAnchos = "400-0-2000-0"
        'OJO. En cargado de valor debe utilizarse las variables
        'Valor de máxima suma de parámetros de
        Set clsDGnral = New DLogGeneral
        pnTotPar = clsDGnral.CargaParametro(5000, 1001)
        Set clsDGnral = Nothing

        lblComenta.Caption = "El valor máximo de la suma de los parámetros que" & vbCr & _
            "intervendrán en el proceso de selección es : " & pnTotPar & vbCr & _
            "ya sea en los parámetros técnicos o económicos"
        Call CargaTxtSelNro
    ElseIf psFrmTpo = "5" Then
        'PUBLICACION
        Me.Caption = "Registro de Publicaciones"
        lblComenta.Visible = False
        fgeParTec.lbEditarFlex = False
        fgeParEco.lbEditarFlex = False
        fgeComite.EncabezadosAnchos = "400-0-3500-0-4000-0"
        fgeAdjudica.EncabezadosAnchos = "400-0-2000-0"
        Call CargaTxtSelNro
    Else
        MsgBox "Tipo formulario no reconocido", vbInformation, " Aviso "
    End If
End Sub

Private Sub CargaAreaAutoriza()
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaArea(AreaAutoriza)
    If rs.RecordCount > 0 Then
        Set fgeAutoriza.Recordset = rs
    End If
    Set rs = Nothing
End Sub

Private Sub CargaAreaComite()
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaArea(AreaComite)
    If rs.RecordCount > 0 Then
        Set fgeComite.Recordset = rs
    End If
    Set rs = Nothing
End Sub

Private Sub CargaSistAdjudica()
    Dim clsDGnral As DLogGeneral
    Dim rs As ADODB.Recordset
    Set clsDGnral = New DLogGeneral
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaConstante(gLogSelSisAdj)
    If rs.RecordCount > 0 Then
        Set fgeAdjudica.Recordset = rs
    End If
    Set rs = Nothing
End Sub

Private Sub CargaTxtSelNro()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    If psFrmTpo = "1" Or psFrmTpo = "2" Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoInicioRes)
    ElseIf psFrmTpo = "3" Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoComite)
    ElseIf psFrmTpo = "4" Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoBases)
    ElseIf psFrmTpo = "5" Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoParametro)
    End If
    If rs.RecordCount > 0 Then
        txtSelNro.rs = rs
    Else
        txtSelNro.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub Limpiar()
    'Limpiar FLEX
    If psFrmTpo = "1" Then
        dtpResFec.Value = gdFecSis
        txtResNro.Text = ""
        fgeAutoriza.Clear
        fgeAutoriza.FormaCabecera
        fgeAutoriza.Rows = 2
    ElseIf psFrmTpo = "2" Then
        fgeComite.Clear
        fgeComite.FormaCabecera
        fgeComite.Rows = 2
    ElseIf psFrmTpo = "3" Then
        txtCostoBase.Text = ""
        txtAdqNro.Text = ""
        fgeBS.Clear
        fgeBS.FormaCabecera
        fgeBS.Rows = 2
    ElseIf psFrmTpo = "4" Then
        lblTotEco.Caption = ""
        lblTotTec.Caption = ""
        fgeParTec.Clear
        fgeParTec.FormaCabecera
        fgeParTec.Rows = 2
        fgeParEco.Clear
        fgeParEco.FormaCabecera
        fgeParEco.Rows = 2
    ElseIf psFrmTpo = "5" Then
        fgePublica.Clear
        fgePublica.FormaCabecera
        fgePublica.Rows = 2
    End If
    
End Sub

Private Sub txtAdqNro_EmiteDatos()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sAdqNro As String
    If txtAdqNro.Ok = False Then
        Exit Sub
    End If
    
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    sAdqNro = txtAdqNro.Text
    If Trim(sAdqNro) <> "" Then
        Set clsDAdq = New DLogAdquisi
        Set rs = New ADODB.Recordset
        Set rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegistro, sAdqNro)
        If rs.RecordCount > 0 Then
            Set fgeBS.Recordset = rs
            fgeBS.AdicionaFila
            fgeBS.BackColorRow &HC0FFFF
            fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L  R E F E R E N C I A L"
            fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
        Else
            txtAdqNro.Text = ""
        End If
    End If
End Sub

Private Sub txtCostoBase_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCostoBase, KeyAscii, 8, 2)
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim clsDGnral As DLogGeneral
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sSelNro As String
    Dim nCont As Integer
    'Al determinar una seleccion, cargarla !
    If txtSelNro.Ok = False Then
        Exit Sub
    End If
    Call Limpiar
    sSelNro = txtSelNro.Text
    If Trim(sSelNro) <> "" Then
        Set clsDAdq = New DLogAdquisi
        Set rs = New ADODB.Recordset
        Set rs = clsDAdq.CargaSeleccion(SelUnRegistro, sSelNro)
        If rs.RecordCount > 0 Then
            Call CargaAreaAutoriza
            With rs
                dtpResFec.Value = Format(!dLogSelRes, "dd/mm/yyyy")
                txtResNro.Text = !cLogSelResNro
                'Busca la persona responsable
                For nCont = 1 To fgeAutoriza.Rows - 1
                    If fgeAutoriza.TextMatrix(nCont, 1) = !cAreaCod And _
                    fgeAutoriza.TextMatrix(nCont, 3) = !cPersCod Then
                        fgeAutoriza.Row = nCont
                        fgeAutoriza.TextMatrix(nCont, 5) = 1
                        Exit For
                    End If
                Next
                If psFrmTpo >= "4" Then
                    fgeAdjudica.Clear
                    fgeAdjudica.FormaCabecera
                    fgeAdjudica.Rows = 2
                    'Muestra Datos de Base Ingresada
                    txtCostoBase.Text = Format(!nLogSelCostoBase, "#0.0")
                    txtAdqNro.Text = !cLogAdqNro
                    fgeAdjudica.AdicionaFila
                    fgeAdjudica.TextMatrix(1, 1) = !cLogSelSisAdj
                    fgeAdjudica.TextMatrix(1, 2) = !cConsDescripcion
                    
                    Set rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegistro, !cLogAdqNro)
                    If rs.RecordCount > 0 Then
                        Set fgeBS.Recordset = rs
                        fgeBS.AdicionaFila
                        fgeBS.BackColorRow &HC0FFFF
                        fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L  R E F E R E N C I A L"
                        fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
                    End If
                End If
            End With
            If psFrmTpo = "1" Then
                'Inicio de Selección
                cmdAdq(0).Enabled = True
                cmdAdq(1).Enabled = True
                cmdAdq(2).Enabled = False
                cmdAdq(3).Enabled = False
            ElseIf psFrmTpo = "2" Then
                'Activar para ingreso de comite
                cmdAdq(0).Visible = False
                cmdAdq(1).Visible = False
                cmdAdq(2).Enabled = True
                cmdAdq(3).Enabled = True
                fgeComite.Enabled = True
                Call CargaAreaComite
            ElseIf psFrmTpo = "3" Then
                'Activar para ingreso de comite
                fraBase.Enabled = True
                cmdAdq(0).Visible = False
                cmdAdq(1).Visible = False
                cmdAdq(2).Enabled = True
                cmdAdq(3).Enabled = True
                fgeComite.Enabled = True
                'Muestra comite ingresado
                Set rs = clsDAdq.CargaSelComite(sSelNro)
                If rs.RecordCount > 0 Then
                    Set fgeComite.Recordset = rs
                End If
                Call CargaSistAdjudica
                'Carga txtAdjNro
                Set rs = clsDAdq.CargaAdquisicion(AdqTodosEstado, "", "", gLogAdqEstadoInicio)
                'If rs.RecordCount > 0 Then
                    txtAdqNro.rs = rs
                'End If
            ElseIf psFrmTpo = "4" Then
                'Activar para ingreso de comite
                fraBase.Enabled = False
                cmdAdq(0).Visible = False
                cmdAdq(1).Visible = False
                cmdAdq(2).Enabled = True
                cmdAdq(3).Enabled = True
                fgeComite.Enabled = True
                'Muestra comite ingresado
                Set rs = clsDAdq.CargaSelComite(sSelNro)
                If rs.RecordCount > 0 Then
                    Set fgeComite.Recordset = rs
                End If
                'Carga Parametros
                Set clsDGnral = New DLogGeneral
                Set fgeParTec.Recordset = clsDGnral.CargaConstante(gLogSelParTec)
                Set fgeParEco.Recordset = clsDGnral.CargaConstante(gLogSelParEco)
                Set clsDGnral = Nothing
                fgeParTec.lbEditarFlex = True
                fgeParEco.lbEditarFlex = True
            ElseIf psFrmTpo = "5" Then
                fgePublica.lbEditarFlex = True
                fraBase.Enabled = False
                cmdAdq(0).Visible = False
                cmdAdq(1).Visible = False
                cmdAdq(2).Enabled = True
                cmdAdq(3).Enabled = True
                cmdPubli(0).Enabled = True
                cmdPubli(1).Enabled = True
                'Muestra comite ingresado
                Set rs = clsDAdq.CargaSelComite(sSelNro)
                If rs.RecordCount > 0 Then
                    Set fgeComite.Recordset = rs
                End If
                'Muestra parámetros ingresados
                Set rs = clsDAdq.CargaSelParametro(sSelNro, "1")
                If rs.RecordCount > 0 Then
                    Set fgeParTec.Recordset = rs
                    Call fgeParTec_OnCellChange(fgeParTec.Row, fgeParTec.Col)
                End If
                Set rs = clsDAdq.CargaSelParametro(sSelNro, "2")
                If rs.RecordCount > 0 Then
                    Set fgeParEco.Recordset = rs
                    Call fgeParEco_OnCellChange(fgeParEco.Row, fgeParEco.Col)
                End If
                
            End If
        End If
    End If
End Sub
