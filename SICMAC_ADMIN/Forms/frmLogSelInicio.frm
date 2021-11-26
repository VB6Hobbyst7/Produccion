VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogSelInicio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6690
   ClientLeft      =   195
   ClientTop       =   1515
   ClientWidth     =   11265
   Icon            =   "frmLogSelInicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   2925
      TabIndex        =   6
      Top             =   6240
      Width           =   1290
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   105
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
      Left            =   4710
      TabIndex        =   3
      Top             =   6240
      Width           =   1290
   End
   Begin VB.CommandButton cmdAdq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   6405
      TabIndex        =   2
      Top             =   6240
      Width           =   1290
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9000
      TabIndex        =   0
      Top             =   6225
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   90
      Top             =   6225
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin TabDlg.SSTab sstSeleccion 
      Height          =   5685
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   10028
      _Version        =   393216
      Tabs            =   12
      Tab             =   8
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Resolución"
      TabPicture(0)   =   "frmLogSelInicio.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraResolu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comité"
      TabPicture(1)   =   "frmLogSelInicio.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEtiqueta(6)"
      Tab(1).Control(1)=   "fgeComite"
      Tab(1).Control(2)=   "cmdComi(0)"
      Tab(1).Control(3)=   "cmdComi(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Detalle"
      TabPicture(2)   =   "frmLogSelInicio.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBase"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cronograma"
      TabPicture(3)   =   "frmLogSelInicio.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblEtiqueta(0)"
      Tab(3).Control(1)=   "fgeCronograma"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Referencias"
      TabPicture(4)   =   "frmLogSelInicio.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblEtiqueta(7)"
      Tab(4).Control(1)=   "lblEtiqueta(15)"
      Tab(4).Control(2)=   "lblEtiqueta(16)"
      Tab(4).Control(3)=   "fgeSisAdj"
      Tab(4).Control(4)=   "txtCostoBase"
      Tab(4).Control(5)=   "txtValReferencia"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Parámetros"
      TabPicture(5)   =   "frmLogSelInicio.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblComenta"
      Tab(5).Control(1)=   "lblEtiqueta(14)"
      Tab(5).Control(2)=   "lblEtiqueta(13)"
      Tab(5).Control(3)=   "lblTotEco"
      Tab(5).Control(4)=   "lblTotTec"
      Tab(5).Control(5)=   "lblEtiqueta(11)"
      Tab(5).Control(6)=   "lblEtiqueta(10)"
      Tab(5).Control(7)=   "fgeParEco"
      Tab(5).Control(8)=   "fgeParTec"
      Tab(5).Control(9)=   "fraParTecRango"
      Tab(5).Control(10)=   "fraParEcoRango"
      Tab(5).ControlCount=   11
      TabCaption(6)   =   "Publicación"
      TabPicture(6)   =   "frmLogSelInicio.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdPubli(1)"
      Tab(6).Control(1)=   "cmdPubli(0)"
      Tab(6).Control(2)=   "fgePublica"
      Tab(6).Control(3)=   "lblEtiqueta(12)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Cotizaciones"
      TabPicture(7)   =   "frmLogSelInicio.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fgePro"
      Tab(7).Control(1)=   "fgeSel"
      Tab(7).Control(2)=   "fgeCot"
      Tab(7).Control(3)=   "lblEtiqueta(19)"
      Tab(7).Control(4)=   "lblEtiqueta(18)"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "Entrega Bases"
      TabPicture(8)   =   "frmLogSelInicio.frx":03EA
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "lblEtiqueta(20)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "fgePostor"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "cmdPostor(1)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "cmdPostor(0)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).ControlCount=   4
      TabCaption(9)   =   "Consultas"
      TabPicture(9)   =   "frmLogSelInicio.frx":0406
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "cmdConsulta"
      Tab(9).Control(1)=   "fgeConsulta"
      Tab(9).Control(2)=   "rtfConsulta"
      Tab(9).Control(3)=   "lblObserva(0)"
      Tab(9).Control(4)=   "lblEtiqueta(21)"
      Tab(9).ControlCount=   5
      TabCaption(10)  =   "Absolución"
      TabPicture(10)  =   "frmLogSelInicio.frx":0422
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "cmdAbsolucion"
      Tab(10).Control(1)=   "fgeAbsolucion"
      Tab(10).Control(2)=   "rtfAbsolucion"
      Tab(10).Control(3)=   "lblEtiqueta(22)"
      Tab(10).Control(4)=   "lblObserva(1)"
      Tab(10).ControlCount=   5
      TabCaption(11)  =   "Observaciones"
      TabPicture(11)  =   "frmLogSelInicio.frx":043E
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "cmdObservacion"
      Tab(11).Control(1)=   "fgeObservacion"
      Tab(11).Control(2)=   "rtfObservacion"
      Tab(11).Control(3)=   "lblEtiqueta(23)"
      Tab(11).Control(4)=   "lblObserva(2)"
      Tab(11).ControlCount=   5
      Begin VB.CommandButton cmdObservacion 
         Caption         =   "Guarda &Observación"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -67305
         TabIndex        =   77
         Top             =   4740
         Width           =   1785
      End
      Begin VB.CommandButton cmdAbsolucion 
         Caption         =   "Guarda &Absolución"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -67305
         TabIndex        =   76
         Top             =   4740
         Width           =   1785
      End
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "Guarda Con&sulta"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -67305
         TabIndex        =   71
         Top             =   4740
         Width           =   1785
      End
      Begin VB.CommandButton cmdPostor 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   4305
         TabIndex        =   66
         Top             =   4305
         Width           =   1155
      End
      Begin VB.CommandButton cmdPostor 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   5880
         TabIndex        =   65
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Frame fraParEcoRango 
         Caption         =   "Rangos"
         Height          =   2010
         Left            =   -68910
         TabIndex        =   56
         Top             =   2970
         Visible         =   0   'False
         Width           =   4230
         Begin Sicmact.FlexEdit fgeParEcoRango 
            Height          =   1695
            Left            =   510
            TabIndex        =   57
            Top             =   225
            Width           =   3255
            _ExtentX        =   5741
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
      Begin VB.Frame fraParTecRango 
         Caption         =   "Rangos"
         Height          =   2010
         Left            =   -74370
         TabIndex        =   54
         Top             =   2970
         Visible         =   0   'False
         Width           =   4230
         Begin Sicmact.FlexEdit fgeParTecRango 
            Height          =   1695
            Left            =   510
            TabIndex        =   55
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
      Begin VB.TextBox txtValReferencia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72150
         TabIndex        =   50
         Top             =   1635
         Width           =   1305
      End
      Begin VB.TextBox txtCostoBase 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72150
         TabIndex        =   46
         Top             =   1020
         Width           =   1305
      End
      Begin VB.CommandButton cmdPubli 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -69900
         TabIndex        =   34
         Top             =   4230
         Width           =   1155
      End
      Begin VB.CommandButton cmdPubli 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -71475
         TabIndex        =   33
         Top             =   4215
         Width           =   1155
      End
      Begin VB.CommandButton cmdComi 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -69600
         TabIndex        =   21
         Top             =   4695
         Width           =   1155
      End
      Begin VB.CommandButton cmdComi 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -71175
         TabIndex        =   20
         Top             =   4680
         Width           =   1155
      End
      Begin VB.Frame fraBase 
         BorderStyle     =   0  'None
         Height          =   4590
         Left            =   -74955
         TabIndex        =   15
         Top             =   660
         Width           =   10905
         Begin Sicmact.FlexEdit fgeBS 
            Height          =   3660
            Left            =   4470
            TabIndex        =   31
            Top             =   405
            Width           =   6420
            _ExtentX        =   11324
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-7"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-4"
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
         Begin VB.Frame fraMoneda 
            Caption         =   "Moneda "
            Height          =   1320
            Left            =   45
            TabIndex        =   25
            Top             =   3120
            Width           =   2430
            Begin VB.TextBox txtTipCambio 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1335
               TabIndex        =   28
               Top             =   720
               Width           =   945
            End
            Begin VB.OptionButton optMoneda 
               Caption         =   "Soles"
               Height          =   195
               Index           =   0
               Left            =   225
               TabIndex        =   27
               Top             =   285
               Value           =   -1  'True
               Width           =   750
            End
            Begin VB.OptionButton optMoneda 
               Caption         =   "Dólares"
               Height          =   195
               Index           =   1
               Left            =   1095
               TabIndex        =   26
               Top             =   285
               Width           =   900
            End
            Begin VB.Label lblTipCambio 
               Caption         =   "Tipo Cambio :"
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
               Left            =   135
               TabIndex        =   29
               Top             =   750
               Width           =   1260
            End
         End
         Begin Sicmact.FlexEdit fgeBSMes 
            Height          =   4035
            Left            =   2565
            TabIndex        =   22
            Top             =   405
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   7117
            Cols0           =   4
            HighLight       =   2
            AllowUserResizing=   1
            EncabezadosNombres=   "Mes-Código-Descripción-Opc"
            EncabezadosAnchos=   "350-0-1000-350"
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
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-4"
            EncabezadosAlineacion=   "R-L-L-C"
            FormatosEdit    =   "0-0-0-0"
            CantEntero      =   6
            CantDecimales   =   1
            AvanceCeldas    =   1
            TextArray0      =   "Mes"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   345
            RowHeight0      =   300
         End
         Begin Sicmact.FlexEdit fgeAdq 
            Height          =   2715
            Left            =   45
            TabIndex        =   24
            Top             =   405
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   4789
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-Adquisición-Area-Periodo-Moneda-Estado-Opc"
            EncabezadosAnchos=   "400-750-0-0-600-0-350"
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
            ColumnasAEditar =   "X-X-X-X-X-X-6"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-4"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-C-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "Item"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbOrdenaCol     =   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
         Begin Sicmact.FlexEdit fgeBSTotal 
            Height          =   705
            Left            =   4470
            TabIndex        =   30
            Top             =   3735
            Width           =   6420
            _ExtentX        =   11324
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
         Begin VB.Label lblEtiqueta 
            Caption         =   "Items"
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
            Index           =   17
            Left            =   4560
            TabIndex        =   53
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Meses"
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
            Left            =   2655
            TabIndex        =   23
            Top             =   165
            Width           =   720
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Req.Consolidado"
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
            Left            =   135
            TabIndex        =   16
            Top             =   150
            Width           =   1575
         End
      End
      Begin VB.Frame fraResolu 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4335
         Left            =   -74820
         TabIndex        =   7
         Top             =   780
         Width           =   10005
         Begin VB.TextBox txtResNro 
            Height          =   285
            Left            =   5940
            MaxLength       =   20
            TabIndex        =   9
            Top             =   75
            Width           =   1890
         End
         Begin Sicmact.FlexEdit fgeAutoriza 
            Height          =   3225
            Left            =   4890
            TabIndex        =   8
            Top             =   885
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   5689
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Item-cAreaCod-Area-Opc"
            EncabezadosAnchos=   "400-0-3500-0"
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
            Left            =   1335
            TabIndex        =   10
            Top             =   105
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            Format          =   65273857
            CurrentDate     =   36783
            MaxDate         =   401768
            MinDate         =   36526
         End
         Begin Sicmact.FlexEdit fgeSelTpo 
            Height          =   3225
            Left            =   255
            TabIndex        =   17
            Top             =   885
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   5689
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
            ColumnasAEditar =   "X-X-X-3"
            ListaControles  =   "0-0-0-5"
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
            Left            =   315
            TabIndex        =   18
            Top             =   645
            Width           =   2505
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
            Left            =   360
            TabIndex        =   13
            Top             =   165
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
            Left            =   4950
            TabIndex        =   12
            Top             =   120
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
            Left            =   4965
            TabIndex        =   11
            Top             =   645
            Width           =   1245
         End
      End
      Begin Sicmact.FlexEdit fgeComite 
         Height          =   3405
         Left            =   -74580
         TabIndex        =   19
         Top             =   1140
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   6006
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Codigo-Area"
         EncabezadosAnchos=   "400-1200-3500-0-3500"
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
         ColumnasAEditar =   "X-1-X-3-X"
         ListaControles  =   "0-1-0-1-0"
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
         Left            =   -74475
         TabIndex        =   35
         Top             =   1200
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5212
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha Inicial-Fecha Final"
         EncabezadosAnchos=   "400-1700-4000-1300-1300"
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
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgeParTec 
         Height          =   1710
         Left            =   -74715
         TabIndex        =   37
         Top             =   975
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3016
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "400-0-2500-800-1000"
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
         ColumnasAEditar =   "X-X-X-3-4"
         ListaControles  =   "0-0-0-0-3"
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
         Left            =   -69360
         TabIndex        =   38
         Top             =   975
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3016
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Descripción-Puntaje-Tipo"
         EncabezadosAnchos=   "400-0-2500-800-1000"
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
         ColumnasAEditar =   "X-X-X-3-4"
         ListaControles  =   "0-0-0-0-3"
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
      Begin Sicmact.FlexEdit fgeSisAdj 
         Height          =   2325
         Left            =   -74385
         TabIndex        =   51
         Top             =   2550
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
         ColumnasAEditar =   "X-X-X-3"
         ListaControles  =   "0-0-0-5"
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
         Left            =   -74550
         TabIndex        =   52
         Top             =   1245
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   5768
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Descripción-Chk-Fecha Inicial-Fecha Final"
         EncabezadosAnchos=   "400-0-5000-0-1300-1300"
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
         ColumnasAEditar =   "X-X-X-3-4-5"
         ListaControles  =   "0-0-0-4-2-2"
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
         Height          =   1275
         Left            =   -74760
         TabIndex        =   58
         Top             =   4335
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   2249
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Codigo-Nombre-Dirección"
         EncabezadosAnchos=   "400-0-2500-3500"
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
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeSel 
         Height          =   855
         Left            =   -74760
         TabIndex        =   59
         Top             =   3255
         Width           =   10545
         _ExtentX        =   18600
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
         Left            =   -74760
         TabIndex        =   60
         Top             =   915
         Width           =   10545
         _ExtentX        =   18600
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
         Left            =   660
         TabIndex        =   63
         Top             =   1140
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   5212
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1400-4000-1000-0-0-0"
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
         ColumnasAEditar =   "X-1-X-3-X-X-X"
         ListaControles  =   "0-1-0-2-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-L-C-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   1
         CellBackColor   =   16777215
      End
      Begin Sicmact.FlexEdit fgeConsulta 
         Height          =   4035
         Left            =   -74685
         TabIndex        =   67
         Top             =   1125
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   7117
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1400-3500-0-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0"
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
         Height          =   3540
         Left            =   -68880
         TabIndex        =   69
         Top             =   1110
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   6244
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelInicio.frx":045A
      End
      Begin Sicmact.FlexEdit fgeAbsolucion 
         Height          =   4035
         Left            =   -74685
         TabIndex        =   72
         Top             =   1125
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   7117
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1400-3500-0-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0"
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
         Height          =   3540
         Left            =   -68880
         TabIndex        =   73
         Top             =   1110
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   6244
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelInicio.frx":04DC
      End
      Begin Sicmact.FlexEdit fgeObservacion 
         Height          =   4035
         Left            =   -74685
         TabIndex        =   78
         Top             =   1125
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   7117
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Fecha-Consulta-Absolucion-Observacion"
         EncabezadosAnchos=   "400-1400-3500-0-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0"
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
         Height          =   3540
         Left            =   -68880
         TabIndex        =   79
         Top             =   1110
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   6244
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         MaxLength       =   4000
         TextRTF         =   $"frmLogSelInicio.frx":055E
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
         TabIndex        =   81
         Top             =   855
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
         Left            =   -68745
         TabIndex        =   80
         Top             =   840
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
         TabIndex        =   75
         Top             =   855
         Width           =   1485
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
         Left            =   -68745
         TabIndex        =   74
         Top             =   840
         Width           =   1140
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
         Left            =   -68745
         TabIndex        =   70
         Top             =   840
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
         Index           =   21
         Left            =   -74610
         TabIndex        =   68
         Top             =   855
         Width           =   1485
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
         Left            =   780
         TabIndex        =   64
         Top             =   870
         Width           =   855
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
         Left            =   -74475
         TabIndex        =   62
         Top             =   4125
         Width           =   1245
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Bienes/Servicios"
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
         Index           =   18
         Left            =   -74475
         TabIndex        =   61
         Top             =   705
         Width           =   1560
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
         TabIndex        =   49
         Top             =   2280
         Width           =   2370
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
         TabIndex        =   48
         Top             =   1650
         Width           =   2235
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
         TabIndex        =   47
         Top             =   1035
         Width           =   1170
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
         Left            =   -74565
         TabIndex        =   45
         Top             =   720
         Width           =   1425
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
         Left            =   -69075
         TabIndex        =   44
         Top             =   735
         Width           =   1425
      End
      Begin VB.Label lblTotTec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -71805
         TabIndex        =   43
         Top             =   2670
         Width           =   840
      End
      Begin VB.Label lblTotEco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -66465
         TabIndex        =   42
         Top             =   2670
         Width           =   840
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
         Left            =   -72690
         TabIndex        =   41
         Top             =   2715
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
         Index           =   14
         Left            =   -67350
         TabIndex        =   40
         Top             =   2715
         Width           =   855
      End
      Begin VB.Label lblComenta 
         Caption         =   $"frmLogSelInicio.frx":05E0
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   -73065
         TabIndex        =   39
         Top             =   5085
         Visible         =   0   'False
         Width           =   7680
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
         Left            =   -74385
         TabIndex        =   36
         Top             =   960
         Width           =   1530
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
         Index           =   0
         Left            =   -74385
         TabIndex        =   32
         Top             =   945
         Width           =   1170
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
         Index           =   6
         Left            =   -74490
         TabIndex        =   14
         Top             =   885
         Width           =   1395
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
      TabIndex        =   5
      Top             =   150
      Width           =   870
   End
End
Attribute VB_Name = "frmLogSelInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pnFrmTpo As String
Dim pbFrmConsul As Boolean
Dim pnTotPar As Currency
Dim paTecRango() As Currency
Dim paEcoRango() As Currency
Dim pnMinCot As Currency
Dim clsDGnral As DLogGeneral


Public Sub Inicio(ByVal psFormTpo As String, Optional ByVal pbConsulta As Boolean = False)
pnFrmTpo = psFormTpo
pbFrmConsul = pbConsulta
Me.Show 1
End Sub

Private Sub cmdAbsolucion_Click()
    Dim clsDMov As DLogMov
    Dim sSelNro As String, sPersCod As String
    Dim sActualiza As String
    Set clsDMov = New DLogMov
    
    rtfAbsolucion.Text = Replace(rtfAbsolucion.Text, "'", " ", , , vbTextCompare)
    
    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
        sSelNro = txtSelNro.Text
        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        
        sPersCod = fgeAbsolucion.TextMatrix(fgeAbsolucion.Row, 1)
        
        'Actualiza la Absolucion
        clsDMov.ActualizaSelPostor SelPosAbsolucion, clsDMov.GetnMovNro(sSelNro), _
            sPersCod, rtfAbsolucion.Text, sActualiza
        
        cmdAbsolucion.Enabled = False
        fgeAbsolucion.TextMatrix(fgeAbsolucion.Row, 5) = rtfAbsolucion.Text
    End If
End Sub

Private Sub cmdAdq_Click(Index As Integer)
    Dim nSelNro As Long, nSelTraNro As Long
    Dim sSelNro As String, sSelTraNro As String, sActualiza As String
    Dim sAreaCod As String, sPersCod As String, sResNro As String, sBSCod As String
    Dim sConReq As String, sConMes  As String, sConBS As String
    Dim nParCod As Integer, nParTpo As Integer, nSelTpo As Integer, nSisAdj As Integer
    Dim nCont As Integer, nCont2 As Integer, nSum As Integer, nSum2 As Integer, nResult As Integer
    Dim nTotEco As Currency, nTotTec As Currency
    Dim nCostoBase As Currency, nValorRefe As Currency
    Dim nPrecio As Currency, nCantidad As Currency, nTipCambio As Currency
    Dim nMoneda As Integer
    Dim dPubIni As Date, dPubFin As Date, dPostor As Date
    Dim clsDMov As DLogMov
    
    Select Case Index
        Case 0:
            'NUEVO
            fraResolu.Enabled = True
            txtSelNro.Enabled = False
            txtSelNro.Text = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            
            Call Limpiar
            fgeSelTpo.EncabezadosAnchos = "400-0-3000-400"
            fgeAutoriza.EncabezadosAnchos = "400-0-3500-400"
            fgeAutoriza.lbEditarFlex = True
            fgeSelTpo.lbEditarFlex = True
            'Carga los Flex
            Call CargaSelTpo
            Call CargaAreaAutoriza
            cmdAdq(0).Enabled = False
            cmdAdq(2).Enabled = True
            cmdAdq(3).Enabled = True
        Case 1:
            'EDITAR
        Case 2:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call Limpiar
                fraResolu.Enabled = False
                txtSelNro.Enabled = True
                txtSelNro.Text = ""
                cmdAdq(0).Enabled = True
                cmdAdq(2).Enabled = False
                cmdAdq(3).Enabled = False
            End If
        Case 3:
            'GRABAR
            sSelNro = txtSelNro.Text
            'VALIDACIONES
            If pnFrmTpo = 1 Then
                'Inicio
                'Area Responsable
                sAreaCod = ""
                For nCont = 1 To fgeAutoriza.Rows - 1
                    If fgeAutoriza.TextMatrix(nCont, 3) = "." Then
                        sAreaCod = fgeAutoriza.TextMatrix(nCont, 1)
                        Exit For
                    End If
                Next
                If Len(Trim(sAreaCod)) = 0 Then
                    MsgBox "Falta determinar el area responsable", vbInformation, "Aviso"
                    Exit Sub
                End If
                'Tipo de proceso de Selección
                nSelTpo = 0
                For nCont = 1 To fgeSelTpo.Rows - 1
                    If fgeSelTpo.TextMatrix(nCont, 3) = "." Then
                        nSelTpo = Val(fgeSelTpo.TextMatrix(nCont, 1))
                        Exit For
                    End If
                Next
                If nSelTpo = 0 Then
                    MsgBox "Falta determinar el Tipo de proceso de Selección", vbInformation, "Aviso"
                    Exit Sub
                End If
                'Número de Resolución
                sResNro = Trim(txtResNro.Text)
                If Trim(sResNro) = "" Then
                    MsgBox "Falta ingresar el número de Resolución", vbInformation, " Aviso "
                    Exit Sub
                End If
            ElseIf pnFrmTpo = 2 Then
                'Comité
                If fgeComite.TextMatrix(1, 0) = "" Then
                    MsgBox "Falta determinar el comite ", vbInformation, "Aviso"
                    Exit Sub
                End If
                For nCont = 1 To fgeComite.Rows - 1
                    If fgeComite.TextMatrix(nCont, 1) = "" Then
                        MsgBox "Falta completar la información del comite ", vbInformation, "Aviso"
                        Exit Sub
                    End If
                Next
            ElseIf pnFrmTpo = 3 Then
                'DETALLE
                'Requerimiento
                nSum = 0
                For nCont = 1 To fgeAdq.Rows - 1
                    If fgeAdq.TextMatrix(nCont, 6) = "." Then
                        nSum = nSum + 1
                        Exit For
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta determinar el requerimiento (consolidado)", vbInformation, " Aviso "
                    Exit Sub
                End If
                'Meses
                nSum = 0
                For nCont = 1 To fgeBSMes.Rows - 1
                    If fgeBSMes.TextMatrix(nCont, 3) = "." Then
                        nSum = nSum + 1
                        Exit For
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta determinar el(los) mes(es)", vbInformation, " Aviso "
                    Exit Sub
                End If
                'Items
                nSum = 0
                For nCont = 1 To fgeBS.Rows - 1
                    If fgeBS.TextMatrix(nCont, 7) = "." Then
                        nSum = nSum + 1
                        Exit For
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta determinar el(los) item(s)", vbInformation, " Aviso "
                    Exit Sub
                End If
                'Total
                If (Val(fgeBSTotal.TextMatrix(1, 6)) = 0) Then
                    MsgBox "No existen cantidades en el(los) item(s) señalado(s)", vbInformation, " Aviso "
                    Exit Sub
                End If
                'Otros
                If Val(txtTipCambio.Text) <= 0 Then
                    MsgBox "Falta determinar el Tipo de Cambio", vbInformation, " Aviso "
                    Exit Sub
                End If
            ElseIf pnFrmTpo = 4 Then
                'Cronograma
                For nCont = 1 To fgeCronograma.Rows - 1
                    If fgeCronograma.TextMatrix(nCont, 3) = "." Then
                        If IsDate(fgeCronograma.TextMatrix(nCont, 4)) And IsDate(fgeCronograma.TextMatrix(nCont, 5)) Then
                            If DateDiff("d", fgeCronograma.TextMatrix(nCont, 4), fgeCronograma.TextMatrix(nCont, 5)) < 0 Then
                                MsgBox "Fechas no validas en el Item " & Str(nCont), vbInformation, " Aviso "
                                Exit Sub
                            End If
                        Else
                            MsgBox "Falta ingresar Fecha(s), en el Item " & Str(nCont), vbInformation, " Aviso "
                            Exit Sub
                        End If
                    End If
                Next
                'Referencias
                nCostoBase = CCur(IIf(txtCostoBase.Text = "", 0, txtCostoBase.Text))
                If nCostoBase <= 0 Then
                    MsgBox "Falta ingresar el Costo Base", vbInformation, " Aviso "
                    Exit Sub
                End If
                
                nValorRefe = CCur(IIf(txtValReferencia.Text = "", 0, txtValReferencia.Text))
                If nValorRefe <= 0 Then
                    MsgBox "Falta ingresar el Valor de Referencia", vbInformation, " Aviso "
                    Exit Sub
                End If
                
                nSisAdj = 0
                For nCont = 1 To fgeSisAdj.Rows - 1
                    If fgeSisAdj.TextMatrix(nCont, 3) = "." Then
                        nSisAdj = Val(fgeSisAdj.TextMatrix(nCont, 1))
                        Exit For
                    End If
                Next
                If nSisAdj = 0 Then
                    MsgBox "Falta determinar el Sistema de Adjudicación", vbInformation, "Aviso"
                    Exit Sub
                End If
                
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
            
            ElseIf pnFrmTpo = 5 Then
                'Publicaciones
                For nCont = 1 To fgePublica.Rows - 1
                    If fgePublica.TextMatrix(nCont, 1) = "" Then
                        MsgBox "Falta el publicador en el Item : " & nCont, vbInformation, "Aviso"
                        Exit Sub
                    End If
                    If fgePublica.TextMatrix(nCont, 3) = "" Or fgePublica.TextMatrix(nCont, 4) = "" Then
                        MsgBox "Falta el determinar la(s) fecha(s) en el Item : " & nCont, vbInformation, "Aviso"
                        Exit Sub
                    End If
                Next
            ElseIf pnFrmTpo = 6 Then
                'Cotizaciones
                'Verifica que siempre este por lo menos UNO
                nSum = 0
                For nCont = 6 To fgeSel.Cols - 1
                    If fgeSel.TextMatrix(1, nCont) <> "" Then nSum = nSum + 1
                Next
                If nSum = 0 Then
                    MsgBox "Falta seleccionar a los Provedores a enviar cotizaciones", vbInformation, " Aviso"
                    Exit Sub
                End If
                If nSum < pnMinCot Then
                    MsgBox "Mínimo de cotizaciones a generar debe ser : " & pnMinCot, vbInformation, " Aviso"
                    Exit Sub
                End If
            ElseIf pnFrmTpo = 7 Then
                'Postores
                'Verifica que siempre este por lo menos UNO
                nSum = 0
                For nCont = 1 To fgePostor.Cols - 1
                    If fgePostor.TextMatrix(1, nCont) <> "" Then
                        nSum = nSum + 1
                        Exit For
                    End If
                Next
                If nSum = 0 Then
                    MsgBox "Falta ingrezar a los Postores ", vbInformation, " Aviso"
                    Exit Sub
                End If
                
                For nCont = 1 To fgePostor.Rows - 1
                    If fgePostor.TextMatrix(nCont, 1) = "" Or fgePostor.TextMatrix(nCont, 3) = "" Then
                        MsgBox "Falta ingrezar Postor y/o Fecha en el Item " & nCont, vbInformation, " Aviso"
                        Exit Sub
                    End If
                Next
            ElseIf pnFrmTpo = 8 Then
                'CONSULTA
            ElseIf pnFrmTpo = 9 Then
                'ABSOLUCION
            ElseIf pnFrmTpo = 10 Then
                'OBSERVACION
            Else
                MsgBox "Tipo de Formulario ¡ No Reconocido !", vbInformation, " Aviso "
                Exit Sub
            End If
            
            If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                If pnFrmTpo = 1 Then
                    'INICIO
                    sSelTraNro = sSelNro
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    'Grabación de MOV - MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelRegistro)), "", gLogSelEstadoInicioRes
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = nSelTraNro
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Inserta LogAdquisicion
                    clsDMov.InsertaSeleccion nSelNro, dtpResFec.Value, sResNro, _
                        sAreaCod, nSelTpo, sActualiza
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Enabled = True
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        fraResolu.Enabled = False
                        txtSelNro.Enabled = True
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                
                ElseIf pnFrmTpo = 2 Then
                    'COMITE
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoComite
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoComite, sActualiza
                    
                    For nCont = 1 To fgeComite.Rows - 1
                        'Inserta LogSelComite
                        sPersCod = fgeComite.TextMatrix(nCont, 1)
                        sAreaCod = fgeComite.TextMatrix(nCont, 3)
                        clsDMov.InsertaSelComite nSelNro, sAreaCod, sPersCod, _
                            sActualiza
                    Next
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        cmdComi(0).Enabled = False
                        cmdComi(1).Enabled = False
                        txtSelNro.Enabled = True
                        fgeComite.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf pnFrmTpo = 3 Then
                    'DETALLE
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoBases
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Requerimientos a Consolidar
                    sConReq = ""
                    For nCont = 1 To fgeAdq.Rows - 1
                        If fgeAdq.TextMatrix(nCont, 6) = "." Then
                            sConReq = sConReq & "','" & fgeAdq.TextMatrix(nCont, 1)
                        End If
                    Next
                    sConReq = Mid(sConReq, 4)
                    'Meses a Consolidar
                    sConMes = ""
                    For nCont = 1 To fgeBSMes.Rows - 1
                        If fgeBSMes.TextMatrix(nCont, 3) = "." Then
                            sConMes = sConMes & "," & fgeBSMes.TextMatrix(nCont, 0)
                        End If
                    Next
                    sConMes = Mid(sConMes, 2)
                    'Items a Consolidar
                    sConBS = ""
                    For nCont = 1 To fgeBS.Rows - 1
                        If fgeBS.TextMatrix(nCont, 7) = "." Then
                            sConBS = sConBS & "','" & fgeBS.TextMatrix(nCont, 1)
                        End If
                    Next
                    sConBS = Mid(sConBS, 4)
                    
                    nMoneda = IIf(optMoneda(0).Value = True, gMonedaNacional, gMonedaExtranjera)
                    nTipCambio = Val(txtTipCambio.Text)
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionDetalle nSelNro, nMoneda, nTipCambio, sActualiza
                    
                    'Actualiza LogReqDetMes (Consolidado)
                    clsDMov.ActualizaReqDetMesConsol sConReq, sConMes, sConBS, nSelNro, sActualiza
                    
                    'Inserta el Detalle de los Bienes/Servicios
                    For nCont = 1 To fgeBS.Rows - 1
                        If fgeBS.TextMatrix(nCont, 7) = "." Then
                            sConBS = fgeBS.TextMatrix(nCont, 1)
                            nPrecio = CCur(IIf(fgeBS.TextMatrix(nCont, 5) = "", 0, fgeBS.TextMatrix(nCont, 5)))
                            nCantidad = CCur(IIf(fgeBS.TextMatrix(nCont, 4) = "", 0, fgeBS.TextMatrix(nCont, 4)))
                            clsDMov.InsertaSelDetalle nSelNro, sConBS, _
                                  nPrecio, nCantidad, sActualiza
                        End If
                    Next
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        txtSelNro.Enabled = True
                        fraBase.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf pnFrmTpo = 4 Then
                    'REFERENCIA (PARAMETRO)
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoParametro
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionReferencia nSelNro, nCostoBase, nValorRefe, nSisAdj, sActualiza
                    'Inserta Cronograma
                    For nCont = 1 To fgeCronograma.Rows - 1
                        If fgeCronograma.TextMatrix(nCont, 3) = "." Then
                            clsDMov.InsertaSelCronograma nSelNro, fgeCronograma.TextMatrix(nCont, 1), _
                                Format(fgeCronograma.TextMatrix(nCont, 4), gsFormatoFechaView), _
                                Format(fgeCronograma.TextMatrix(nCont, 5), gsFormatoFechaView), sActualiza
                        End If
                    Next
                    'Inserta LogSelParametro - Técnico
                    For nCont = 1 To fgeParTec.Rows - 1
                        nParCod = Val(fgeParTec.TextMatrix(nCont, 1))
                        nSum = CCur(IIf(fgeParTec.TextMatrix(nCont, 3) = "", 0, fgeParTec.TextMatrix(nCont, 3)))
                        If nSum > 0 Then
                            nParTpo = Right(Trim(fgeParTec.TextMatrix(nCont, 4)), 1)
                            clsDMov.InsertaSelParametro nSelNro, 1, nParCod, nSum, nParTpo, sActualiza
                            If nParTpo = 3 Then
                                'Por Rangos
                                For nCont2 = 1 To 5
                                    If paTecRango(3, nCont, nCont2) <> 0 And paTecRango(1, nCont, nCont2) <> 0 And paTecRango(2, nCont, nCont2) <> 0 Then
                                        clsDMov.InsertaSelParDetalle nSelNro, 1, nParCod, nCont2, _
                                            paTecRango(3, nCont, nCont2), _
                                            paTecRango(1, nCont, nCont2), _
                                            paTecRango(2, nCont, nCont2)
                                    End If
                                Next
                            End If
                        End If
                    Next
                    'Inserta LogSelParametro - Económico
                    For nCont = 1 To fgeParEco.Rows - 1
                        nParCod = Val(fgeParEco.TextMatrix(nCont, 1))
                        nSum = CCur(IIf(fgeParEco.TextMatrix(nCont, 3) = "", 0, fgeParEco.TextMatrix(nCont, 3)))
                        If nSum > 0 Then
                            nParTpo = Right(Trim(fgeParEco.TextMatrix(nCont, 4)), 1)
                            clsDMov.InsertaSelParametro nSelNro, 2, nParCod, nSum, nParTpo, sActualiza
                            If nParTpo = 3 Then
                                'Por Rangos
                                For nCont2 = 1 To 5
                                    If paEcoRango(3, nCont, nCont2) <> 0 And paEcoRango(1, nCont, nCont2) <> 0 And paEcoRango(2, nCont, nCont2) <> 0 Then
                                        clsDMov.InsertaSelParDetalle nSelNro, 2, nParCod, nCont2, paEcoRango(3, nCont, nCont2), paEcoRango(1, nCont, nCont2), paEcoRango(2, nCont, nCont2)
                                    End If
                                Next
                            End If
                        End If
                    Next
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgeCronograma.lbEditarFlex = False
                        fgeSisAdj.lbEditarFlex = False
                        fgeParTec.lbEditarFlex = False
                        fgeParEco.lbEditarFlex = False
                        txtCostoBase.Enabled = False
                        txtValReferencia.Enabled = False
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        txtSelNro.Enabled = True
                        fraBase.Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf pnFrmTpo = 5 Then
                    'PUBLICACION
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoPublicacion
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoPublicacion, sActualiza
                    
                    For nCont = 1 To fgePublica.Rows - 1
                        sPersCod = fgePublica.TextMatrix(nCont, 1)
                        dPubIni = fgePublica.TextMatrix(nCont, 3)
                        dPubFin = fgePublica.TextMatrix(nCont, 4)
                        clsDMov.InsertaSelPublica nSelNro, sPersCod, dPubIni, dPubFin, sActualiza
                    Next
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgePublica.lbEditarFlex = False
                        cmdAdq(0).Visible = False
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
                ElseIf pnFrmTpo = 6 Then
                    'COTIZACION
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV -MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoCotizacion
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoCotizacion, sActualiza
                    
                    nCont = 0: nSum = 0
                    For nCont2 = 6 To fgeSel.Cols - 1
                        nSum = nSum + 1
                        If fgeSel.TextMatrix(1, nCont2) <> "" Then
                            nCont = nCont + 1   'GeneraCotiza(sSelNro, nCont)
                            sPersCod = fgePro.TextMatrix(nSum, 1)
                            'Inserta LogSelCotiza
                            clsDMov.InsertaSelCotiza nSelNro, nCont, sPersCod, sActualiza
                            For nSum2 = 1 To fgeCot.Rows - 1
                                sBSCod = fgeCot.TextMatrix(nSum2, 1)
                                'Inserta LogSelCotDetalle
                                clsDMov.InsertaSelCotDetalle nSelNro, nCont, sBSCod, 0, 0, sActualiza
                            Next
                            'Inserta LogSelCotPar
                            'Tecnico
                            For nSum2 = 1 To fgeParTec.Rows - 1
                                nParCod = fgeParTec.TextMatrix(nSum2, 1)
                                clsDMov.InsertaSelCotPar nSelNro, nCont, 1, nParCod, 0, sActualiza
                            Next
                            'Economico
                            For nSum2 = 1 To fgeParEco.Rows - 1
                                nParCod = fgeParEco.TextMatrix(nSum2, 1)
                                clsDMov.InsertaSelCotPar nSelNro, nCont, 2, nParCod, 0, sActualiza
                            Next
                        End If
                    Next
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        fgeSel.Enabled = False
                        txtSelNro.Enabled = True
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                ElseIf pnFrmTpo = 7 Then
                    'ENTREGA DE BASES
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV -MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoRegBase
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoRegBase, sActualiza
                    
                    For nCont = 1 To fgePostor.Rows - 1
                        sPersCod = fgePostor.TextMatrix(nCont, 1)
                        dPostor = fgePostor.TextMatrix(nCont, 3)
                        clsDMov.InsertaSelPostor nSelNro, sPersCod, dPostor, sActualiza
                    Next
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        fgePostor.lbEditarFlex = False
                        cmdPostor(0).Enabled = False
                        cmdPostor(1).Enabled = False
                        txtSelNro.Enabled = True
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If

                ElseIf pnFrmTpo = 8 Then
                    'CONSULTA DE BASES
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV -MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoConsulta
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoConsulta, sActualiza
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        
                        rtfConsulta.Locked = True
                        cmdConsulta.Enabled = False
                        txtSelNro.Enabled = True
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                
                ElseIf pnFrmTpo = 9 Then
                    'ABSOLUCION DE CONSULTA DE BASES
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV -MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoAbsolucion
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoAbsolucion, sActualiza
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        
                        rtfAbsolucion.Locked = True
                        cmdAbsolucion.Enabled = False
                        txtSelNro.Enabled = True
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If

                ElseIf pnFrmTpo = 10 Then
                    'OBSERVACION
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Grabación de MOV -MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", gLogSelEstadoObservacion
                    nSelTraNro = clsDMov.GetnMovNro(sSelTraNro)
                    nSelNro = clsDMov.GetnMovNro(sSelNro)
                    clsDMov.InsertaMovRef nSelTraNro, nSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionEstado nSelNro, gLogSelEstadoObservacion, sActualiza
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        cmdAdq(0).Visible = False
                        cmdAdq(2).Enabled = False
                        cmdAdq(3).Enabled = False
                        
                        rtfObservacion.Locked = True
                        cmdObservacion.Enabled = False
                        txtSelNro.Enabled = True
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

Private Sub cmdComi_Click(Index As Integer)
    Dim nBSRow As Integer
    'Botones de comandos para agregar/eliminar personas del comite
    If Index = 0 Then
        'Agregar en Flex
        fgeComite.AdicionaFila
        fgeComite.SetFocus
    ElseIf Index = 1 Then
        'Eliminar en Flex
        nBSRow = fgeComite.Row
        If fgeComite.TextMatrix(1, 0) <> "" Then
            If MsgBox("¿ Estás seguro de eliminar " & fgeComite.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                fgeComite.EliminaFila nBSRow
            End If
        Else
            MsgBox "No existe registro a eliminar", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub cmdConsulta_Click()
    Dim clsDMov As DLogMov
    Dim sSelNro As String, sPersCod As String
    Dim sActualiza As String
    Set clsDMov = New DLogMov
    
    rtfConsulta.Text = Replace(rtfConsulta.Text, "'", " ", , , vbTextCompare)
    
    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
        sSelNro = txtSelNro.Text
        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        
        sPersCod = fgeConsulta.TextMatrix(fgeConsulta.Row, 1)
        
        'Actualiza la Consulta
        clsDMov.ActualizaSelPostor SelPosConsulta, clsDMov.GetnMovNro(sSelNro), _
            sPersCod, rtfConsulta.Text, sActualiza
        
        cmdConsulta.Enabled = False
        fgeConsulta.TextMatrix(fgeConsulta.Row, 4) = rtfConsulta.Text
    End If
End Sub

Private Sub cmdObservacion_Click()
    Dim clsDMov As DLogMov
    Dim sSelNro As String, sPersCod As String
    Dim sActualiza As String
    Set clsDMov = New DLogMov
    
    rtfObservacion.Text = Replace(rtfObservacion.Text, "'", " ", , , vbTextCompare)
    
    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
        sSelNro = txtSelNro.Text
        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        
        sPersCod = fgeObservacion.TextMatrix(fgeObservacion.Row, 1)
        
        'Actualiza la Consulta
        clsDMov.ActualizaSelPostor SelPosObservacion, clsDMov.GetnMovNro(sSelNro), _
            sPersCod, rtfObservacion.Text, sActualiza
        
        cmdObservacion.Enabled = False
        fgeObservacion.TextMatrix(fgeObservacion.Row, 6) = rtfObservacion.Text
    End If
End Sub

Private Sub cmdPostor_Click(Index As Integer)
    Dim nBSRow As Integer
    'Botones de comandos para agregar/eliminar publicaciones
    If Index = 0 Then
        'Agregar en Flex
        fgePostor.AdicionaFila
        fgePostor.SetFocus
    ElseIf Index = 1 Then
        'Eliminar en Flex
        nBSRow = fgePostor.Row
        If fgePostor.TextMatrix(1, 0) <> "" Then
            If MsgBox("¿ Estás seguro de eliminar " & fgePostor.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                fgePostor.EliminaFila nBSRow
            End If
        Else
            MsgBox "No existe registro a eliminar", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub cmdPubli_Click(Index As Integer)
    Dim nBSRow As Integer
    'Botones de comandos para agregar/eliminar publicaciones
    If Index = 0 Then
        'Agregar en Flex
        fgePublica.AdicionaFila
        fgePublica.SetFocus
    ElseIf Index = 1 Then
        'Eliminar en Flex
        nBSRow = fgePublica.Row
        If fgePublica.TextMatrix(1, 0) <> "" Then
            If MsgBox("¿ Estás seguro de eliminar " & fgePublica.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                fgePublica.EliminaFila nBSRow
            End If
        Else
            MsgBox "No existe registro a eliminar", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Set clsDGnral = Nothing
    Unload Me
End Sub

Private Sub fgeAbsolucion_OnRowChange(pnRow As Long, pnCol As Long)
    rtfAbsolucion.Text = fgeAbsolucion.TextMatrix(pnRow, 5)
    cmdAbsolucion.Enabled = False
End Sub

Private Sub fgeAdq_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim clsDReq As DLogRequeri
    Dim rs As ADODB.Recordset
    Dim sReqNroAll As String, sMeses As String
    Dim nCont As Integer, nMoneda As Integer
    
    sReqNroAll = ""
    For nCont = 1 To fgeAdq.Rows - 1
        If fgeAdq.TextMatrix(nCont, 6) = "." Then
            sReqNroAll = sReqNroAll & clsDGnral.GetnMovNro(Trim(fgeAdq.TextMatrix(nCont, 1))) & ","
        End If
    Next
    'Carga detalle requerimientos para aprobación CONSOLIDADO
    If Len(Trim(sReqNroAll)) > 0 Then
        'Verifica el Tipo de Cambio
        If Val(txtTipCambio.Text) = 0 Then
            MsgBox "Por Favor ingrese el Tipo de Cambio," & vbCr & "que será necesario cuando halla cambio de moneda", vbInformation, " Aviso"
            Exit Sub
        End If
        
        sReqNroAll = Left(sReqNroAll, Len(sReqNroAll) - 1)
        Set clsDReq = New DLogRequeri
        Set rs = New ADODB.Recordset
        sMeses = MesCheck
        'Carga Consolidación de Requerimiento
        Set rs = clsDReq.CargaReqDetalle(ReqDetTodosSelec, sReqNroAll, , sMeses, IIf(optMoneda(0).Value = True, True, False), Val(txtTipCambio.Text))
        If rs.RecordCount > 0 Then
            fgeBS.EncabezadosAnchos = "400-0-1850-650-900-900-1000-350"
            Set fgeBS.Recordset = rs
            fgeBSTotal.TextMatrix(1, 6) = ""
        End If
        Set rs = Nothing
    Else
        fgeBS.Clear
        fgeBS.FormaCabecera
        fgeBS.Rows = 2
        fgeBSTotal.TextMatrix(1, 6) = ""
    End If
    
End Sub

Private Sub fgeBS_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim nCont As Integer
    Dim nSum As Currency
    If pnCol = 7 Then
        For nCont = 1 To fgeBS.Rows - 1
            If fgeBS.TextMatrix(nCont, 7) = "." Then
                nSum = nSum + CCur(IIf(fgeBS.TextMatrix(nCont, 6) = "", 0, fgeBS.TextMatrix(nCont, 6)))
            End If
        Next
        fgeBSTotal.BackColorRow &HC0FFFF
        fgeBSTotal.TextMatrix(1, 0) = "="
        fgeBSTotal.TextMatrix(1, 2) = "T O T A L "
        fgeBSTotal.TextMatrix(1, 6) = Format(nSum, "#,##0.00")
    End If
End Sub


Private Sub fgeBSMes_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Call fgeAdq_OnCellCheck(fgeAdq.Row, fgeAdq.Col)
End Sub

Private Function MesCheck() As String
    Dim sMesAll As String
    Dim nCont As Integer
    For nCont = 1 To fgeBSMes.Rows - 1
        If fgeBSMes.TextMatrix(nCont, 3) = "." Then
            sMesAll = sMesAll & Trim(fgeBSMes.TextMatrix(nCont, 0)) & ","
        End If
    Next
    If Len(Trim(sMesAll)) > 0 Then
        sMesAll = Left(sMesAll, Len(sMesAll) - 1)
    Else
        sMesAll = "0"
    End If
    MesCheck = sMesAll
End Function

Private Sub fgeComite_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If pnCol = 1 Then
    'Carga el area de la persona ingresada
    Set rs = clsDGnral.CargaArea(AreaPersona, psDataCod)
    If rs.RecordCount > 0 Then
        fgeComite.TextMatrix(pnRow, 3) = Trim(rs!cAreaCod)
        fgeComite.TextMatrix(pnRow, 4) = Trim(rs!cAreaDescripcion)
    End If
End If
End Sub

Private Sub fgeConsulta_OnRowChange(pnRow As Long, pnCol As Long)
    rtfConsulta.Text = fgeConsulta.TextMatrix(pnRow, 4)
    cmdConsulta.Enabled = False
End Sub

Private Sub fgeCronograma_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If fgeCronograma.TextMatrix(pnRow, pnCol) = "" Then
    fgeCronograma.TextMatrix(pnRow, pnCol + 1) = ""
    fgeCronograma.TextMatrix(pnRow, pnCol + 2) = ""
End If
End Sub

Private Sub fgeObservacion_OnRowChange(pnRow As Long, pnCol As Long)
    rtfObservacion.Text = fgeObservacion.TextMatrix(pnRow, 6)
    cmdObservacion.Enabled = False
End Sub

Private Sub fgeParEco_GotFocus()
Call fgeParEco_OnRowChange(fgeParEco.Row, fgeParEco.Col)
End Sub
Private Sub fgeParEco_OnCellChange(pnRow As Long, pnCol As Long)
    lblTotEco.Caption = fgeParEco.SumaRow(3)
End Sub
Private Sub fgeParEco_OnChangeCombo()
    If Right(Trim(fgeParEco.TextMatrix(fgeParEco.Row, 4)), 1) = "3" Then
        fraParEcoRango.Visible = True
        CargaRango (2)
    Else
        fraParEcoRango.Visible = False
    End If
End Sub
Private Sub fgeParEco_OnRowChange(pnRow As Long, pnCol As Long)
    If Right(Trim(fgeParEco.TextMatrix(pnRow, 4)), 1) = "3" Then
        fraParEcoRango.Visible = True
        CargaRango (2)
    Else
        fraParEcoRango.Visible = False
    End If
End Sub
Private Sub fgeParEcoRango_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    paEcoRango(pnCol, fgeParEco.Row, pnRow) = fgeParEcoRango.TextMatrix(pnRow, pnCol)
End Sub

Private Sub fgeParTec_GotFocus()
    Call fgeParTec_OnChangeCombo
End Sub
Private Sub fgeParTec_OnCellChange(pnRow As Long, pnCol As Long)
    lblTotTec.Caption = fgeParTec.SumaRow(3)
End Sub
Private Sub fgeParTec_OnChangeCombo()
    If Right(Trim(fgeParTec.TextMatrix(fgeParTec.Row, 4)), 1) = "3" Then
        fraParTecRango.Visible = True
        CargaRango (1)
    Else
        fraParTecRango.Visible = False
    End If
End Sub
Private Sub fgeParTec_OnRowChange(pnRow As Long, pnCol As Long)
    If Right(Trim(fgeParTec.TextMatrix(pnRow, 4)), 1) = "3" Then
        fraParTecRango.Visible = True
        CargaRango (1)
    Else
        fraParTecRango.Visible = False
    End If
End Sub
Private Sub fgeParTecRango_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    paTecRango(pnCol, fgeParTec.Row, pnRow) = fgeParTecRango.TextMatrix(pnRow, pnCol)
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
    Dim nCont As Integer
    Set clsDGnral = New DLogGeneral
    Call CentraForm(Me)
    For nCont = 0 To sstSeleccion.Tabs - 1
        sstSeleccion.TabVisible(nCont) = False
    Next
    If pbFrmConsul = True Then
        cmdAdq(0).Visible = False
        cmdAdq(2).Visible = False
        cmdAdq(3).Visible = False
    End If
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    'lblAreaDes.Caption = Usuario.AreaNom
    dtpResFec.Value = gdFecSis
    If pnFrmTpo = 1 Then
        'INICIO
        Me.Caption = "Proceso de Selección : Resolución o acuerdo de Inicio"
        cmdAdq(0).Enabled = True
        sstSeleccion.TabVisible(0) = True
        'Carga FLEX de txtNroSel
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 2 Then
        'COMITE
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Conformación del Comité u Organo Encargado"
            For nCont = 0 To pnFrmTpo - 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Resolución o acuerdo de Inicio : Consulta"
            sstSeleccion.TabVisible(0) = True
        End If
        cmdAdq(0).Visible = False
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 3 Then
        'DETALLE
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Registro del detalle "
            For nCont = 0 To pnFrmTpo - 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Conformación del Comité u Organo Encargado : Consulta "
            For nCont = 0 To pnFrmTpo - 2
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        cmdAdq(0).Visible = False
        fgeComite.EncabezadosAnchos = "400-0-3500-0-4000-0"
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 4 Then
        'REFERENCIAS (CRONOGRAMA, REFERENCIA, PARAMETRO)
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Bases o términos de Referencia"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Registro del detalle : Consulta"
            For nCont = 0 To pnFrmTpo - 2
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        fgeComite.EncabezadosAnchos = "400-0-3500-0-4000-0"
        fgeSelTpo.EncabezadosAnchos = "400-0-2000-0"
        'OJO. En cargado de valor debe utilizarse las variables
        'Valor de máxima suma de parámetros de
        pnTotPar = clsDGnral.CargaParametro(5000, 1001)
        lblComenta.Visible = True
        lblComenta.Caption = "NOTA: El valor máximo de la suma de los parámetros que " & _
            "intervendrán en el proceso de selección es : " & pnTotPar & vbCr & _
            "ya sea en los parámetros técnicos o económicos"
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 5 Then
        'PUBLICACION
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Publicaciones"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Bases o términos de Referencia : Consulta"
            For nCont = 0 To pnFrmTpo
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 6 Then
        'COTIZACION
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Cotizaciones"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Publicaciones : Consulta"
            For nCont = 0 To pnFrmTpo
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        'OJO. En cargado de valor debe utilizarse las variables
        'Valor de máxima suma de parámetros de
        pnMinCot = clsDGnral.CargaParametro(5000, 1002)
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 7 Then
        'ENTREGA DE BASES
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Entrega de Bases"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Cotizaciones : Consulta"
            For nCont = 0 To pnFrmTpo
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 8 Then
        'CONSULTA DE BASES
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Consulta de Bases"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Entrega de Bases : Consulta"
            For nCont = 0 To pnFrmTpo
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 9 Then
        'ABSOLUCION DE CONSULTAS
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Absolución de Consultas"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Consulta de Bases : Consulta"
            For nCont = 0 To pnFrmTpo
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 10 Then
        'OBSERVACION DE LAS BASES
        If pbFrmConsul = False Then
            Me.Caption = "Proceso de Selección : Observación de las Bases"
            For nCont = 0 To pnFrmTpo + 1
                sstSeleccion.TabVisible(nCont) = True
            Next
        Else
            Me.Caption = "Proceso de Selección : Absolución de Consultas : Consulta"
            For nCont = 0 To pnFrmTpo
                sstSeleccion.TabVisible(nCont) = True
            Next
        End If
        Call CargaTxtSelNro
    ElseIf pnFrmTpo = 11 Then
        'Consulta
        Me.Caption = "Proceso de Selección : Observación de las Bases : Consulta"
        For nCont = 0 To pnFrmTpo
            sstSeleccion.TabVisible(nCont) = True
        Next
        Call CargaTxtSelNro
    Else
        MsgBox "Tipo formulario no reconocido", vbInformation, " Aviso "
    End If
End Sub

Private Sub CargaAreaAutoriza(Optional ByVal psAreaAutorizada As String = "")
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaArea(AreaAutoriza, psAreaAutorizada)
    If rs.RecordCount > 0 Then
        Set fgeAutoriza.Recordset = rs
    End If
    Set rs = Nothing
End Sub
Private Sub CargaSelTpo(Optional ByVal pnProcSelec As Integer = 0)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = clsDGnral.CargaConstante(gLogSelTpo, , pnProcSelec)
    If rs.RecordCount > 0 Then
        Set fgeSelTpo.Recordset = rs
    End If
    Set rs = Nothing
End Sub

'Private Sub CargaReferencias()
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    '******************************************
'    'NOTA. cambiar numero por constante global
'    'Carga Cronograma
'    Set rs = clsDGnral.CargaConstante(5054)
'    If rs.RecordCount > 0 Then
'        Set fgeCronograma.Recordset = rs
'    End If
'    Set rs = Nothing
'End Sub

Private Sub CargaTxtSelNro()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    If pnFrmTpo = 2 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoInicioRes)
    ElseIf pnFrmTpo = 3 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoComite)
    ElseIf pnFrmTpo = 4 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoBases)
    ElseIf pnFrmTpo = 5 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoParametro)
    ElseIf pnFrmTpo = 6 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoPublicacion)
    ElseIf pnFrmTpo = 7 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoCotizacion)
    ElseIf pnFrmTpo = 8 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoRegBase)
    ElseIf pnFrmTpo = 9 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoConsulta)
    ElseIf pnFrmTpo = 10 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoAbsolucion)
    ElseIf pnFrmTpo = 11 Then
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, 0, gLogSelEstadoObservacion)
    End If
    
    If pnFrmTpo <> 1 Or (pnFrmTpo = 1 And pbFrmConsul = True) Then
        If rs.RecordCount > 0 Then
            txtSelNro.EditFlex = True
            txtSelNro.rs = rs
            txtSelNro.EditFlex = False
        Else
            txtSelNro.Enabled = False
        End If
        Set rs = Nothing
        Set clsDAdq = Nothing
    End If
End Sub

Private Sub Limpiar()
    Dim nCont As Integer
    'Limpiar FLEX
    If pnFrmTpo >= 1 Then
        dtpResFec.Value = gdFecSis
        txtResNro.Text = ""
        fgeAutoriza.EncabezadosAnchos = "400-0-3500-0"
        fgeAutoriza.Clear
        fgeAutoriza.FormaCabecera
        fgeAutoriza.Rows = 2
        fgeSelTpo.EncabezadosAnchos = "400-0-3000-0"
        fgeSelTpo.Clear
        fgeSelTpo.FormaCabecera
        fgeSelTpo.Rows = 2
    End If
    If pnFrmTpo >= 2 Then
        fgeComite.Clear
        fgeComite.FormaCabecera
        fgeComite.Rows = 2
        cmdComi(0).Enabled = False
        cmdComi(1).Enabled = False
    End If
    If pnFrmTpo >= 3 Then
        fgeAdq.Clear
        fgeAdq.FormaCabecera
        fgeAdq.Rows = 2
        
        fgeBSMes.Clear
        fgeBSMes.FormaCabecera
        fgeBSMes.Rows = 2
        
        fgeBS.Clear
        fgeBS.FormaCabecera
        fgeBS.Rows = 2
        fgeBSTotal.TextMatrix(1, 6) = ""
    End If
    If pnFrmTpo >= 4 Then
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
    End If
    If pnFrmTpo >= 5 Then
        fgePublica.Clear
        fgePublica.FormaCabecera
        fgePublica.Rows = 2
    End If
    If pnFrmTpo >= 6 Then
        fgeCot.Clear
        fgeCot.FormaCabecera
        fgeCot.Rows = 2
        fgeSel.Clear
        fgeSel.FormaCabecera
        fgeSel.Rows = 2
        fgePro.Clear
        fgePro.FormaCabecera
        fgePro.Rows = 2
    End If
    If pnFrmTpo >= 7 Then
        fgePostor.Clear
        fgePostor.FormaCabecera
        fgePostor.Rows = 2
    End If
    If pnFrmTpo >= 8 Then
        fgeConsulta.Clear
        fgeConsulta.FormaCabecera
        fgeConsulta.Rows = 2
    End If
    If pnFrmTpo >= 9 Then
        fgeAbsolucion.Clear
        fgeAbsolucion.FormaCabecera
        fgeAbsolucion.Rows = 2
    End If
    If pnFrmTpo >= 10 Then
        fgeAbsolucion.Clear
        fgeAbsolucion.FormaCabecera
        fgeAbsolucion.Rows = 2
    End If

End Sub

Private Sub optMoneda_Click(Index As Integer)
Call fgeAdq_OnCellCheck(fgeAdq.Row, fgeAdq.Col)
End Sub

Private Sub rtfAbsolucion_Change()
cmdAbsolucion.Enabled = True
End Sub
Private Sub rtfConsulta_Change()
cmdConsulta.Enabled = True
End Sub
Private Sub rtfObservacion_Change()
cmdObservacion.Enabled = True
End Sub

Private Sub txtCostoBase_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCostoBase, KeyAscii, 8, 2)
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim clsDReq As DLogRequeri
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sSelNro As String
    Dim nCont As Integer, nCont2 As Integer
    'Al determinar una seleccion, cargarla !
    If txtSelNro.OK = False Then
        Exit Sub
    End If
    Call Limpiar
    sSelNro = txtSelNro.Text
    If Trim(sSelNro) <> "" Then
        Set clsDReq = New DLogRequeri
        Set clsDAdq = New DLogAdquisi
        Set rs = New ADODB.Recordset
        Set rs = clsDAdq.CargaSeleccion(SelUnRegistro, clsDGnral.GetnMovNro(sSelNro))
        If rs.RecordCount > 0 Then
            With rs
                dtpResFec.Value = Format(!dLogSelRes, gsFormatoFechaView)
                txtResNro.Text = !cLogSelResNro
                'LLena los Flex
                fgeSelTpo.EncabezadosAnchos = "400-0-3000-0"
                fgeSelTpo.AdicionaFila
                fgeSelTpo.TextMatrix(1, 1) = !nLogSelTpo
                fgeSelTpo.TextMatrix(1, 2) = !cConsDescSelTpo
                fgeAutoriza.EncabezadosAnchos = "400-0-3500-0"
                fgeAutoriza.AdicionaFila
                fgeAutoriza.TextMatrix(1, 1) = !cAreaCod
                fgeAutoriza.TextMatrix(1, 2) = !cAreaDescripcion
                If pnFrmTpo >= 4 Then
                    If !nLogSelMoneda = gMonedaNacional Then
                        optMoneda(0).Value = True
                    ElseIf !nLogSelMoneda = gMonedaExtranjera Then
                        optMoneda(1).Value = True
                    Else
                        optMoneda(0).Value = False
                        optMoneda(1).Value = False
                    End If
                    txtTipCambio.Text = Format(!nLogSelTipCambio, "##.000")
                    'Carga Detalle
                    fgeBS.ListaControles = "0-0-0-0-0-0-0-0"
                    Set rs = clsDReq.CargaSelDetalle(clsDGnral.GetnMovNro(sSelNro))
                    If rs.RecordCount > 0 Then
                        Set fgeBS.Recordset = rs
                        'Total
                        fgeBSTotal.BackColorRow &HC0FFFF
                        fgeBSTotal.TextMatrix(1, 0) = "="
                        fgeBSTotal.TextMatrix(1, 2) = "T O T A L "
                        fgeBSTotal.TextMatrix(1, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
                        txtValReferencia.Text = fgeBSTotal.TextMatrix(1, 6)
                    End If
                    If pnFrmTpo > 4 Then
                        'Muestra Datos de Base Ingresada
                        txtCostoBase.Text = Format(!nLogSelCostoBase, "#0.0")
                        fgeSisAdj.ListaControles = "0-0-0-0"
                        fgeSisAdj.AdicionaFila
                        fgeSisAdj.TextMatrix(1, 1) = !nLogSelSisAdj
                        fgeSisAdj.TextMatrix(1, 2) = !cConsDescSisAdj
                    End If
                End If
            End With
            If pnFrmTpo >= 1 Then
                'Inicio de Selección
                cmdAdq(0).Enabled = True
                cmdAdq(2).Enabled = False
                cmdAdq(3).Enabled = False
            End If
            If pnFrmTpo >= 2 Then
                'Activar para ingreso de comite
                cmdAdq(0).Visible = False
                cmdAdq(2).Enabled = True
                cmdAdq(3).Enabled = True
                fgeComite.Enabled = True
                cmdComi(0).Enabled = True
                cmdComi(1).Enabled = True
            End If
            If pnFrmTpo >= 3 Then
                'Detalle
                cmdComi(0).Enabled = False
                cmdComi(1).Enabled = False
                fgeBSMes.Enabled = True
                fgeBS.Enabled = True
                fraBase.Enabled = True
                'Muestra comite ingresado
                Set rs = clsDAdq.CargaSelComite(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgeComite.Recordset = rs
                End If
                If pnFrmTpo = 3 Then
                    'Carga Meses
                    fgeBSMes.rsFlex = clsDGnral.CargaConstante(gMeses)
                    For nCont = 1 To fgeBSMes.Rows - 1
                        fgeBSMes.TextMatrix(nCont, 3) = 1
                    Next
                    'Carga Req. Consolidados
                    Set rs = clsDReq.CargaRequerimiento(gLogReqTipoConsolidado, ReqTodosFlexSeleccion, "")
                    If rs.RecordCount > 0 Then
                        fgeAdq.rsFlex = rs
                    End If
                End If
            End If
            If pnFrmTpo >= 4 Then
                'Referencias
                fraMoneda.Enabled = False
                If pnFrmTpo = 4 Then
                    txtValReferencia.Enabled = True
                    txtCostoBase.Enabled = True
                    fgeSisAdj.lbEditarFlex = True
                    fgeCronograma.lbEditarFlex = True
                    'Carga Cronograma
                    fgeCronograma.EncabezadosAnchos = "400-0-5000-400-1300-1300"
                    Set fgeCronograma.Recordset = clsDGnral.CargaConstante(gLogSelCro)
                    'Carga Sistema Adjudicación
                    fgeSisAdj.EncabezadosAnchos = "400-0-3000-400"
                    Set fgeSisAdj.Recordset = clsDGnral.CargaConstante(gLogSelSisAdj)
                    'Carga Parametros
                    Set fgeParTec.Recordset = clsDGnral.CargaConstante(gLogSelParTec)
                    Set fgeParEco.Recordset = clsDGnral.CargaConstante(gLogSelParEco)
                
                    ReDim paTecRango(3, fgeParTec.Rows, fgeParTecRango.Rows)
                    ReDim paEcoRango(3, fgeParEco.Rows, fgeParEcoRango.Rows)
                
                    fgeParTec.CargaCombo clsDGnral.CargaConstante(gLogSelParTpo, False)
                    fgeParEco.CargaCombo clsDGnral.CargaConstante(gLogSelParTpo, False)
                
                    fgeParTec.lbEditarFlex = True
                    fgeParEco.lbEditarFlex = True
                End If
            End If
            If pnFrmTpo >= 5 Then
                'Publicación
                fgePublica.lbEditarFlex = True
                cmdPubli(0).Enabled = True
                cmdPubli(1).Enabled = True
                txtCostoBase.Enabled = False
                txtValReferencia.Enabled = False
                'Muestra Cronograma
                fgeCronograma.lbEditarFlex = False
                fgeCronograma.ListaControles = "0-0-0-0-2-2"
                Set rs = clsDAdq.CargaSelCronograma(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgeCronograma.Recordset = rs
                End If
                'Muestra parámetros ingresados
                fgeParTec.lbEditarFlex = False
                Set rs = clsDAdq.CargaSelParametro(clsDGnral.GetnMovNro(sSelNro), 1)
                If rs.RecordCount > 0 Then
                    Set fgeParTec.Recordset = rs
                    'Rango
                    ReDim paTecRango(3, fgeParTec.Rows, fgeParTecRango.Rows)
                    Set rs = clsDAdq.CargaSelParDetalle(clsDGnral.GetnMovNro(sSelNro), 1)
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
                    Call fgeParTec_OnCellChange(fgeParTec.Row, fgeParTec.Col)
                End If
                fgeParEco.lbEditarFlex = False
                Set rs = clsDAdq.CargaSelParametro(clsDGnral.GetnMovNro(sSelNro), 2)
                If rs.RecordCount > 0 Then
                    Set fgeParEco.Recordset = rs
                    'Rango
                    ReDim paEcoRango(3, fgeParEco.Rows, fgeParEcoRango.Rows)
                    Set rs = clsDAdq.CargaSelParDetalle(clsDGnral.GetnMovNro(sSelNro), 2)
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
                    Call fgeParEco_OnCellChange(fgeParEco.Row, fgeParEco.Col)
                End If
            End If
            If pnFrmTpo >= 6 Then
                'Cotización
                cmdPubli(0).Enabled = False
                cmdPubli(1).Enabled = False
                'Muestra Publicación
                fgePublica.lbEditarFlex = False
                Set rs = clsDAdq.CargaSelPublica(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgePublica.Recordset = rs
                End If
                If pnFrmTpo = 6 Then
                    'Muestra detalle de Bienes/Servicios
                    fgeSel.Enabled = True
                    Set rs = clsDAdq.CargaSelDetalle(clsDGnral.GetnMovNro(sSelNro))
                    If rs.RecordCount > 0 Then
                        Set fgeCot.Recordset = rs
                        Call TransObj(rs)
                    End If
                End If
            End If
            If pnFrmTpo >= 7 Then
                'Entrega de Bases
                'Muestra Cotizaciones
                Set rs = clsDAdq.CargaSelDetalle(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgeCot.Recordset = rs
                End If
                
                Set rs = clsDAdq.CargaSelCotiza(SelCotPersona, clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgePro.Recordset = rs
                End If
                
                Call CargaSelBSPro
                
                If pnFrmTpo = 7 Then
                    fgePostor.lbEditarFlex = True
                    cmdPostor(0).Enabled = True
                    cmdPostor(1).Enabled = True
                End If
            End If
            If pnFrmTpo >= 8 Then
                'Consulta de Bases
                'Muestra Postores
                Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgePostor.Recordset = rs
                End If
                
                If pnFrmTpo = 8 Then
                    rtfConsulta.Locked = False
                    'Y Consultas
                    Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                    If rs.RecordCount > 0 Then
                        Set fgeConsulta.Recordset = rs
                    End If
                End If
            End If
            If pnFrmTpo >= 9 Then
                'Absolución de Consulta
                'Muestra Consulta
                Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgeConsulta.Recordset = rs
                End If
                If pnFrmTpo = 9 Then
                    rtfAbsolucion.Locked = False
                    Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                    If rs.RecordCount > 0 Then
                        Set fgeAbsolucion.Recordset = rs
                    End If
                End If
            End If
            If pnFrmTpo >= 10 Then
                'OBSERVACIONES
                'Muestra Absolucion
                Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgeAbsolucion.Recordset = rs
                End If
                
                If pnFrmTpo = 10 Then
                    rtfObservacion.Locked = False
                    Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                    If rs.RecordCount > 0 Then
                        Set fgeObservacion.Recordset = rs
                    End If
                End If
            End If
            If pnFrmTpo >= 11 Then
                'OTROS XXX
                'Muestra Observación
                Set rs = clsDAdq.CargaSelPostor(clsDGnral.GetnMovNro(sSelNro))
                If rs.RecordCount > 0 Then
                    Set fgeObservacion.Recordset = rs
                End If
            End If
        End If
    End If
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 8, 4)
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

Private Sub fgeSel_Click()
Dim nCol As Integer
Dim nFil As Integer
Dim nPos As Integer
Dim i As Integer
Dim pColCot As Integer

Dim phCab As ColorConstants

pColCot = 5
phCab = vbBlue         '&H00008000&
nPos = fgeSel.Col - pColCot
nCol = fgeSel.Col
nFil = 1
If fgeSel.TextMatrix(1, 0) = "" Then
    Exit Sub
End If

If nCol >= 3 Then
    fgeSel.Col = nCol
    fgeSel.CellForeColor = vbBlue
   
    fgePro.Row = nPos
    fgeCot.Col = nCol
    If Len(fgeSel.TextMatrix(nFil, nCol)) = 0 Then
       fgeSel.TextMatrix(nFil, nCol) = "  X"
       fgePro.ForeColorRow (&HFF0000)
       For i = 1 To fgeCot.Rows - 1
           fgeCot.Row = i
           fgeCot.CellForeColor = &HFF0000
       Next
    Else
       fgeSel.TextMatrix(nFil, nCol) = ""
       fgePro.ForeColorRow (&H5)
       For i = 1 To fgeCot.Rows - 1
           fgeCot.Row = i
           fgeCot.CellForeColor = "&H00000005"
       Next
    End If
End If
End Sub

Private Sub TransObj(ByVal poRS As ADODB.Recordset)
Dim nCont As Integer, K As Integer, n As Integer
Dim rp As New ADODB.Recordset
Dim sObj As String

Dim clsDProv As DLogProveedor
Dim rs As ADODB.Recordset
Set clsDProv = New DLogProveedor
Set rs = New ADODB.Recordset

'nNumObj = poRS.RecordCount
Dim pColCot As Integer
pColCot = 5

'fgeSel.Cols = 5
'fgeSel.ColWidth(0) = 400
'fgeSel.ColWidth(1) = 0
'fgeSel.ColWidth(2) = 3000
'fgeSel.ColWidth(3) = 1
'fgeSel.ColWidth(4) = 1
'fgeSel.ColWidth(5) = 1
fgeSel.RowHeight(0) = 300
fgeSel.TextMatrix(1, 0) = "="
fgeSel.TextMatrix(1, 2) = "SELECCION DE PROVEEDORES"

'fgeCot.Rows = nNumObj + 1
K = 0
    
For nCont = 1 To fgeCot.Rows - 1
    sObj = fgeCot.TextMatrix(nCont, 1)
    'Carga Proveedores del producto X
    Set rs = clsDProv.CargaProveedorBS(ProBSProveedor, sObj)
    If Not rs.EOF Then
        Do While Not rs.EOF
            K = 0
            K = SeHalla(rs!cPersCod)
            If K = 0 Then
               'n = InsFlex(fgePro)
               fgePro.AdicionaFila
               n = fgePro.Rows - 1
               fgePro.TextMatrix(n, 0) = "P" & Format(n, "00")
               fgePro.TextMatrix(n, 1) = rs!cPersCod
               fgePro.TextMatrix(n, 2) = rs!cPersNombre
               fgePro.TextMatrix(n, 3) = IIf(IsNull(rs!cPersDireccDomicilio), "", rs!cPersDireccDomicilio)
               
               fgeCot.Cols = fgePro.Rows + pColCot
               fgeCot.ColWidth(fgePro.Rows + (pColCot - 1)) = 400
               
               fgeSel.Cols = fgePro.Rows + pColCot
               fgeSel.ColWidth(fgePro.Rows + (pColCot - 1)) = 400
               
               fgeCot.TextMatrix(0, fgePro.Rows + (pColCot - 1)) = "P" & Format(n, "00")
               fgeCot.TextMatrix(nCont, fgePro.Rows + (pColCot - 1)) = "  X"
               'fgeCot.ColAlignment(fgePro.Rows + (pColCot - 1)) = 4
            
               fgeSel.TextMatrix(0, fgePro.Rows + (pColCot - 1)) = "P" & Format(n, "00")
               'fgeSel.ColAlignment(fgePro.Rows + (pColCot - 1)) = 4
            
            Else
               fgeCot.TextMatrix(nCont, K + pColCot) = "  X"
            End If
            rs.MoveNext
        Loop
    End If
Next
End Sub

Function SeHalla(vCod As String) As Integer
Dim z As Integer
SeHalla = 0
For z = 1 To fgePro.Rows - 1
    If fgePro.TextMatrix(z, 1) = vCod Then
       SeHalla = z
       Exit For
    End If
Next
End Function

