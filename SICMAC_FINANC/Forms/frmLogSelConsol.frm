VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmLogSelConsol 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6135
   ClientLeft      =   360
   ClientTop       =   1665
   ClientWidth     =   10980
   Icon            =   "frmLogSelConsol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdq 
      Enabled         =   0   'False
      Height          =   390
      Left            =   4815
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin TabDlg.SSTab sstConsol 
      Height          =   4740
      Left            =   3930
      TabIndex        =   5
      Top             =   795
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmLogSelConsol.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCostoBase"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblResFec"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblResNro"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblEtiqueta(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblEtiqueta(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblEtiqueta(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblEtiqueta(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblEtiqueta(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSisAdj"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblEtiqueta(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fgeRespon"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fgeComite"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Adquisición"
      TabPicture(1)   =   "frmLogSelConsol.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgeBS"
      Tab(1).Control(1)=   "lblEtiqueta(8)"
      Tab(1).Control(2)=   "lblAdqNro"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Parámetros"
      TabPicture(2)   =   "frmLogSelConsol.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fgeParTec"
      Tab(2).Control(1)=   "fgeParEco"
      Tab(2).Control(2)=   "lblEtiqueta(11)"
      Tab(2).Control(3)=   "lblEtiqueta(10)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Otros"
      TabPicture(3)   =   "frmLogSelConsol.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgePostor"
      Tab(3).Control(1)=   "fgePublica"
      Tab(3).Control(2)=   "lblEtiqueta(12)"
      Tab(3).Control(3)=   "lblEtiqueta(1)"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "P.Técnica"
      TabPicture(4)   =   "frmLogSelConsol.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "P.Económica"
      TabPicture(5)   =   "frmLogSelConsol.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fgeBSEco"
      Tab(5).Control(1)=   "fgeCotEco"
      Tab(5).Control(2)=   "lblEtiqueta(15)"
      Tab(5).Control(3)=   "lblEtiqueta(14)"
      Tab(5).ControlCount=   4
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3690
         Left            =   -74745
         TabIndex        =   11
         Top             =   795
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   6509
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
      Begin Sicmact.FlexEdit fgePostor 
         Height          =   1770
         Left            =   -74715
         TabIndex        =   16
         Top             =   2730
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   3122
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Nombre-Observacion"
         EncabezadosAnchos=   "400-0-2800-2700"
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
         ColumnasAEditar =   "X-1-X-X"
         ListaControles  =   "0-1-0-0"
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgePublica 
         Height          =   1770
         Left            =   -74715
         TabIndex        =   18
         Top             =   645
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   3122
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cPersCod-Nombre-Fecha Inicial-Fecha Final"
         EncabezadosAnchos=   "400-0-3000-1200-1200"
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
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
         TipoBusPersona  =   1
      End
      Begin Sicmact.FlexEdit fgeComite 
         Height          =   1800
         Left            =   285
         TabIndex        =   20
         Top             =   2550
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   3175
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cAreaCod-Area-cPersCod-Nombre"
         EncabezadosAnchos=   "400-0-2500-0-3000"
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeParTec 
         Height          =   1665
         Left            =   -74715
         TabIndex        =   22
         Top             =   675
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   2937
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
         Height          =   1665
         Left            =   -74715
         TabIndex        =   24
         Top             =   2670
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   2937
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
      Begin Sicmact.FlexEdit fgeRespon 
         Height          =   930
         Left            =   285
         TabIndex        =   27
         Top             =   1275
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   1640
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Area-Nombre"
         EncabezadosAnchos=   "400-2500-3000"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeBSEco 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   31
         Top             =   630
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   4260
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Precio-Total-Cant.Prov-Prec.Prov-Total Prov"
         EncabezadosAnchos=   "400-0-2000-650-0-0-0-900-900-900"
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
      Begin Sicmact.FlexEdit fgeCotEco 
         Height          =   1260
         Left            =   -74760
         TabIndex        =   32
         Top             =   3375
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   2223
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Cotización-Codigo-Proveedor"
         EncabezadosAnchos=   "400-2200-0-3500"
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
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
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
         Index           =   15
         Left            =   -74685
         TabIndex        =   34
         Top             =   420
         Width           =   750
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
         Index           =   14
         Left            =   -74685
         TabIndex        =   33
         Top             =   3165
         Width           =   1185
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
         Index           =   4
         Left            =   315
         TabIndex        =   28
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label lblSisAdj 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4485
         TabIndex        =   26
         Top             =   690
         Width           =   2010
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
         Left            =   -74550
         TabIndex        =   25
         Top             =   2415
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
         Left            =   -74550
         TabIndex        =   23
         Top             =   420
         Width           =   1425
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Comité"
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
         Left            =   315
         TabIndex        =   21
         Top             =   2310
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Publicaciones"
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
         Left            =   -74550
         TabIndex        =   19
         Top             =   405
         Width           =   1530
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
         Index           =   1
         Left            =   -74550
         TabIndex        =   17
         Top             =   2490
         Width           =   855
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Sistema Adjudicación "
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
         Left            =   4515
         TabIndex        =   15
         Top             =   465
         Width           =   1950
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Costo base"
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
         Left            =   3345
         TabIndex        =   14
         Top             =   465
         Width           =   1095
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
         Left            =   -74715
         TabIndex        =   13
         Top             =   495
         Width           =   1125
      End
      Begin VB.Label lblAdqNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -73500
         TabIndex        =   12
         Top             =   450
         Width           =   2640
      End
      Begin VB.Label lblEtiqueta 
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
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   2
         Left            =   2235
         TabIndex        =   10
         Top             =   465
         Width           =   780
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Resolución"
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
         Left            =   315
         TabIndex        =   9
         Top             =   465
         Width           =   1110
      End
      Begin VB.Label lblResNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   285
         TabIndex        =   8
         Top             =   690
         Width           =   1875
      End
      Begin VB.Label lblResFec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2205
         TabIndex        =   7
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label lblCostoBase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3330
         TabIndex        =   6
         Top             =   690
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9015
      TabIndex        =   0
      Top             =   5625
      Width           =   1305
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   45
      Top             =   5640
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeProceso 
      Height          =   4755
      Left            =   120
      TabIndex        =   4
      Top             =   795
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   8387
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Seleccion-Resolución-Estado"
      EncabezadosAnchos=   "400-2000-0-1000"
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
      ColWidth0       =   405
      RowHeight0      =   285
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
      Height          =   210
      Index           =   13
      Left            =   4200
      TabIndex        =   29
      Top             =   510
      Width           =   795
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
      Left            =   315
      TabIndex        =   3
      Top             =   495
      Width           =   990
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1305
      TabIndex        =   2
      Top             =   90
      Width           =   3705
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
      Left            =   465
      TabIndex        =   1
      Top             =   135
      Width           =   555
   End
End
Attribute VB_Name = "frmLogSelConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String

Public Sub Inicio(ByVal psFormTpo As String)
psFrmTpo = psFormTpo
Me.Show 1
End Sub

Private Sub cmdAdq_Click()
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    Dim sSelNro As String, sSelTraNro As String, sAdqNro As String, sActualiza As String
    Dim nResult As Integer
    
    'Verifica que siempre este por lo menos UNO
    If fgeProceso.TextMatrix(1, 1) = "" Then Exit Sub
    sSelNro = fgeProceso.TextMatrix(fgeProceso.Row, 1)
    sAdqNro = Trim(lblAdqNro.Caption)
    If sSelNro = "" Then Exit Sub
    
    If psFrmTpo = "2" Then
        'RECHAZO
        If MsgBox("¿ Estás seguro de Rechazar este Proceso de Selección " & sSelNro & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoRechazado))
            clsDMov.InsertaMovRef sSelTraNro, sSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", _
                "", "", sActualiza, gLogSelEstadoRechazado
            
            'Libera la Adquisición
            If sAdqNro <> "" Then
                clsDMov.ActualizaAdquisicion sAdqNro, gLogAdqEstadoInicio, sActualiza
            End If
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
        If MsgBox("¿ Estás seguro de declarar Desierto este Proceso de Selección " & sSelNro & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoDesierto))
            clsDMov.InsertaMovRef sSelTraNro, sSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", _
                "", "", sActualiza, gLogSelEstadoDesierto
            
            'Libera la Adquisición
            If sAdqNro <> "" Then
                clsDMov.ActualizaAdquisicion sAdqNro, gLogAdqEstadoInicio, sActualiza
            End If
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
        If MsgBox("¿ Estás seguro de Terminar este Proceso de Selección " & sSelNro & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDGnral = New DLogGeneral
            sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDGnral = Nothing
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoAceptado))
            clsDMov.InsertaMovRef sSelTraNro, sSelNro
            
            'Actualiza LogSelección
            clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", _
                "", "", sActualiza, gLogSelEstadoAceptado
            
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

Private Sub fgeCotEco_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sSelCotNro As String
    
    'ECONOMICA
    sSelCotNro = fgeCotEco.TextMatrix(fgeCotEco.Row, 1)
    'Muestra datos
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegistro, sSelCotNro)
    If rs.RecordCount > 0 Then
        Set fgeBSEco.Recordset = rs
        fgeBSEco.AdicionaFila
        fgeBSEco.BackColorRow &HC0FFFF
        fgeBSEco.TextMatrix(fgeBSEco.Row, 2) = "T O T A L  R E F E R E N C I A L"
        fgeBSEco.TextMatrix(fgeBSEco.Row, 9) = Format(fgeBSEco.SumaRow(9), "#,##0.00")
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub fgeProceso_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDAdq As DLogAdquisi
    Dim rsPro As ADODB.Recordset, rs As ADODB.Recordset
    Dim sSelNro As String
    
    'Verifica que siempre este por lo menos UNO
    If fgeProceso.TextMatrix(1, 1) = "" Then
        cmdAdq.Enabled = False
        Exit Sub
    End If
    
    If psFrmTpo = "2" Or psFrmTpo = "3" Or psFrmTpo = "4" Then cmdAdq.Enabled = True
    sSelNro = fgeProceso.TextMatrix(fgeProceso.Row, 1)
    'Actualiza Detalle de Proceso
    Set clsDAdq = New DLogAdquisi
    Set rsPro = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rsPro = clsDAdq.CargaSeleccion(SelUnRegistro, sSelNro)
    If rsPro.RecordCount > 0 Then
        With rsPro
            Call Limpiar
            lblResNro.Caption = !cLogSelResNro
            lblResFec.Caption = Format(!dLogSelRes, "dd/mm/yyyy")
            lblCostoBase.Caption = Format(!nLogSelCostoBase, "#0.0")
            lblSisAdj.Caption = !cConsDescripcion
            lblAdqNro.Caption = !cLogAdqNro
            
            'Muestra Responsable
            'fgeRespon.AdicionaFila
            fgeRespon.TextMatrix(1, 1) = !cAreaDescripcion
            fgeRespon.TextMatrix(1, 2) = !cPersNombre
            
            'Muestra Comite
            Set rs = clsDAdq.CargaSelComite(sSelNro)
            If rs.RecordCount > 0 Then Set fgeComite.Recordset = rs
            
            'Muestra detalle de Bienes/Servicios
            Set rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegistro, !cLogAdqNro)
            If rs.RecordCount > 0 Then
                Set fgeBS.Recordset = rs
                fgeBS.AdicionaFila
                fgeBS.BackColorRow &HC0FFFF
                fgeBS.TextMatrix(fgeBS.Row, 2) = "T O T A L  R E F E R E N C I A L"
                fgeBS.TextMatrix(fgeBS.Row, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
            End If
            
            'Muestra Parametros Técnicos
            Set rs = clsDAdq.CargaSelParametro(sSelNro, "1")
            If rs.RecordCount > 0 Then Set fgeParTec.Recordset = rs
            'Muestra Parametros Económicos
            Set rs = clsDAdq.CargaSelParametro(sSelNro, "2")
            If rs.RecordCount > 0 Then Set fgeParEco.Recordset = rs
            
            'Muestra Publicaciones
            Set rs = clsDAdq.CargaSelPublica(sSelNro)
            If rs.RecordCount > 0 Then Set fgePublica.Recordset = rs
            
            'Muestra Postores
            Set rs = clsDAdq.CargaSelPostor(sSelNro)
            If rs.RecordCount > 0 Then Set fgePostor.Recordset = rs
            
            'Muestra Propuesta Técnica y Económica
            If psFrmTpo = "4" Then
                'Determina si es para la ACEPTACION (UNA COTIZACION)
                Set rs = clsDAdq.CargaSelCotiza(sSelNro, !cLogSelCotNro)
                If rs.RecordCount > 0 Then
                    'TECNICA
                    
                    'ECONOMICA
                    Set fgeCotEco.Recordset = rs
                    Call fgeCotEco_OnRowChange(fgeCotEco.Row, fgeCotEco.Col)
                End If
            Else
                'De lo contrario TODAS LAS COTIZACIONES
                Set rs = clsDAdq.CargaSelCotiza(sSelNro)
                If rs.RecordCount > 0 Then
                    'TECNICA
                    
                    'ECONOMICA
                    Set fgeCotEco.Recordset = rs
                    Call fgeCotEco_OnRowChange(fgeCotEco.Row, fgeCotEco.Col)
                End If
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
    lblAreaDes.Caption = Usuario.AreaNom
    
    If psFrmTpo = "1" Then
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
        Me.Caption = "Aceptación del Proceso de Selección"
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
    'fgeProceso.Clear
    'fgeProceso.FormaCabecera
    'fgeProceso.Rows = 2
    If psFrmTpo = "1" Then
        'CONSULTA
        Set rs = clsDAdq.CargaSeleccion(SelTodosGnral)
    ElseIf psFrmTpo = "2" Then
        'CANCELACION
        Set rs = clsDAdq.CargaSeleccion(SelTodosGnral, , gLogSelEstadoRechazado)
    ElseIf psFrmTpo = "3" Then
        'DESIERTO
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoPublicacion, gLogSelEstadoRegBase)
    ElseIf psFrmTpo = "4" Then
        'ACEPTACION
        Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoProcAdju)
    End If
    If rs.RecordCount > 0 Then
        Set fgeProceso.Recordset = rs
        Call fgeProceso_OnRowChange(fgeProceso.Row, fgeProceso.Col)
    Else
        fgeProceso.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub Limpiar()
    lblResNro.Caption = ""
    lblResFec.Caption = ""
    lblCostoBase.Caption = ""
    lblSisAdj.Caption = ""
    lblAdqNro.Caption = ""
    fgeComite.Clear
    fgeComite.FormaCabecera
    fgeComite.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeParTec.Clear
    fgeParTec.FormaCabecera
    fgeParTec.Rows = 2
    fgeParEco.Clear
    fgeParEco.FormaCabecera
    fgeParEco.Rows = 2
    fgePublica.Clear
    fgePublica.FormaCabecera
    fgePublica.Rows = 2
    fgePostor.Clear
    fgePostor.FormaCabecera
    fgePostor.Rows = 2
    fgeBSEco.Clear
    fgeBSEco.FormaCabecera
    fgeBSEco.Rows = 2
    fgeCotEco.Clear
    fgeCotEco.FormaCabecera
    fgeCotEco.Rows = 2
    
End Sub

