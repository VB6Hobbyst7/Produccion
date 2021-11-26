VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLogReqInicio 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requerimiento"
   ClientHeight    =   6075
   ClientLeft      =   465
   ClientTop       =   1740
   ClientWidth     =   10995
   Icon            =   "frmLogReqInicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin Sicmact.Usuario Usuario 
      Left            =   30
      Top             =   5580
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   3
      Left            =   6480
      TabIndex        =   18
      Top             =   5565
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   4980
      TabIndex        =   17
      Top             =   5565
      Width           =   1305
   End
   Begin Sicmact.TxtBuscar txtBuscar 
      Height          =   300
      Left            =   1185
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   529
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
      TipoBusqueda    =   2
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   1965
      TabIndex        =   15
      Top             =   5565
      Width           =   1305
   End
   Begin TabDlg.SSTab sstReq 
      Height          =   4800
      Left            =   105
      TabIndex        =   2
      Top             =   660
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   8467
      _Version        =   393216
      TabsPerRow      =   4
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
      TabCaption(0)   =   "S&ustentación"
      TabPicture(0)   =   "frmLogReqInicio.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEtiqueta(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEtiqueta(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "rtfDescri(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "rtfDescri(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Detalle"
      TabPicture(1)   =   "frmLogReqInicio.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboMoneda"
      Tab(1).Control(1)=   "cboAreas"
      Tab(1).Control(2)=   "fgeBSMes"
      Tab(1).Control(3)=   "fgeBS"
      Tab(1).Control(4)=   "cmdReqDet(1)"
      Tab(1).Control(5)=   "cmdReqDet(0)"
      Tab(1).Control(6)=   "fgeMes"
      Tab(1).Control(7)=   "lblMoneda"
      Tab(1).Control(8)=   "lblEtiqueta(8)"
      Tab(1).Control(9)=   "lblAreas"
      Tab(1).Control(10)=   "lblEtiqueta(6)"
      Tab(1).Control(11)=   "lblEtiqueta(5)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "&Flujo"
      TabPicture(2)   =   "frmLogReqInicio.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdIr(1)"
      Tab(2).Control(1)=   "cmdIr(0)"
      Tab(2).Control(2)=   "cmdReqFlu(1)"
      Tab(2).Control(3)=   "cmdReqFlu(0)"
      Tab(2).Control(4)=   "rtfObservacion"
      Tab(2).Control(5)=   "cmdReqFlu(2)"
      Tab(2).Control(6)=   "fgeFlujo"
      Tab(2).Control(7)=   "fgeDestino"
      Tab(2).Control(8)=   "lblDestino"
      Tab(2).Control(9)=   "lblFlujo"
      Tab(2).Control(10)=   "lblObservacion"
      Tab(2).ControlCount=   11
      Begin VB.ComboBox cboMoneda 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73815
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   390
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboAreas 
         Height          =   315
         Left            =   -70500
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   375
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.CommandButton cmdIr 
         Appearance      =   0  'Flat
         Height          =   450
         Index           =   1
         Left            =   -69870
         Picture         =   "frmLogReqInicio.frx":035E
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1605
         Width           =   555
      End
      Begin VB.CommandButton cmdIr 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   -69870
         Picture         =   "frmLogReqInicio.frx":0668
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   930
         Width           =   555
      End
      Begin VB.CommandButton cmdReqFlu 
         Caption         =   "Trámite"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -67545
         TabIndex        =   25
         Top             =   4155
         Width           =   1155
      End
      Begin VB.CommandButton cmdReqFlu 
         Caption         =   "Rechazar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -68760
         TabIndex        =   24
         Top             =   4155
         Width           =   1155
      End
      Begin Sicmact.FlexEdit fgeBSMes 
         Height          =   4035
         Left            =   -67350
         TabIndex        =   6
         Top             =   690
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   7117
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   1
         EncabezadosNombres=   "Mes-Código-Descripción-Cantidad"
         EncabezadosAnchos=   "400-0-1070-900"
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
         ColumnasAEditar =   "X-X-X-3"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "R-L-L-R"
         FormatosEdit    =   "0-0-0-2"
         CantEntero      =   6
         CantDecimales   =   1
         AvanceCeldas    =   1
         TextArray0      =   "Mes"
         lbFlexDuplicados=   0   'False
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   3255
         Left            =   -74895
         TabIndex        =   5
         Top             =   1005
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   5741
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "Item-Código-Descripción-Unidad-Precio Unit."
         EncabezadosAnchos=   "450-1200-3000-700-1000"
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
         ColumnasAEditar =   "X-1-X-X-4"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0"
         EncabezadosAlineacion=   "R-L-L-L-R"
         FormatosEdit    =   "0-0-0-0-2"
         CantEntero      =   6
         CantDecimales   =   1
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   285
      End
      Begin VB.CommandButton cmdReqDet 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -71010
         TabIndex        =   20
         Top             =   4365
         Width           =   1155
      End
      Begin VB.CommandButton cmdReqDet 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -72585
         TabIndex        =   19
         Top             =   4350
         Width           =   1155
      End
      Begin RichTextLib.RichTextBox rtfDescri 
         Height          =   4035
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   615
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7117
         _Version        =   393217
         ScrollBars      =   2
         MaxLength       =   8000
         TextRTF         =   $"frmLogReqInicio.frx":0972
      End
      Begin RichTextLib.RichTextBox rtfDescri 
         Height          =   4035
         Index           =   1
         Left            =   5370
         TabIndex        =   4
         Top             =   615
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7117
         _Version        =   393217
         ScrollBars      =   2
         MaxLength       =   8000
         TextRTF         =   $"frmLogReqInicio.frx":09E0
      End
      Begin Sicmact.FlexEdit fgeMes 
         Height          =   1425
         Left            =   -74895
         TabIndex        =   21
         Top             =   2850
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   2514
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Enero-Febrero-Marzo-Abril-Mayo-Junio-Julio-Agosto-Setiembre-Octubre-Noviembre-Diciembre"
         EncabezadosAnchos=   "400-550-550-550-550-550-550-550-550-550-550-550-550"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin RichTextLib.RichTextBox rtfObservacion 
         Height          =   1860
         Left            =   -74760
         TabIndex        =   23
         Top             =   2760
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   3281
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmLogReqInicio.frx":0A4E
      End
      Begin VB.CommandButton cmdReqFlu 
         Caption         =   "Aceptar"
         Height          =   330
         Index           =   2
         Left            =   -66330
         TabIndex        =   27
         Top             =   4155
         Visible         =   0   'False
         Width           =   1155
      End
      Begin Sicmact.FlexEdit fgeFlujo 
         Height          =   1860
         Left            =   -69225
         TabIndex        =   28
         Top             =   645
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   3281
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Nro.Trámite-Area-cEstadoCod-Estado-Comentario"
         EncabezadosAnchos=   "400-0-3200-0-800-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
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
         EncabezadosAlineacion=   "C-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         RowHeight0      =   225
      End
      Begin Sicmact.FlexEdit fgeDestino 
         Height          =   1860
         Left            =   -74730
         TabIndex        =   32
         Top             =   645
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   3281
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Código-Area"
         EncabezadosAnchos=   "400-0-4000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         RowHeight0      =   225
      End
      Begin VB.Label lblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -73815
         TabIndex        =   39
         Top             =   420
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Moneda :"
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
         Left            =   -74760
         TabIndex        =   37
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label lblAreas 
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
         Left            =   -71160
         TabIndex        =   34
         Top             =   435
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblDestino 
         Caption         =   "Para trámite"
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
         Left            =   -74670
         TabIndex        =   0
         Top             =   435
         Width           =   1500
      End
      Begin VB.Label lblFlujo 
         Caption         =   "Trámite"
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
         Left            =   -69180
         TabIndex        =   29
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label lblObservacion 
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
         Left            =   -74670
         TabIndex        =   26
         Top             =   2565
         Width           =   1500
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Mes"
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
         Left            =   -67170
         TabIndex        =   14
         Top             =   435
         Width           =   675
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
         Index           =   5
         Left            =   -74775
         TabIndex        =   13
         Top             =   735
         Width           =   1500
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Requerimiento"
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
         Left            =   5505
         TabIndex        =   12
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Necesidad"
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
         Left            =   240
         TabIndex        =   11
         Top             =   405
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8745
      TabIndex        =   7
      Top             =   5565
      Width           =   1305
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   3480
      TabIndex        =   22
      Top             =   5565
      Width           =   1305
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   8250
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblPeriodo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8250
      TabIndex        =   36
      Top             =   105
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1185
      TabIndex        =   16
      Top             =   60
      Width           =   4110
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
      Left            =   300
      TabIndex        =   10
      Top             =   390
      Width           =   825
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Año :"
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
      Left            =   7650
      TabIndex        =   9
      Top             =   135
      Width           =   660
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
      Left            =   315
      TabIndex        =   8
      Top             =   105
      Width           =   750
   End
End
Attribute VB_Name = "frmLogReqInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoReq As String, psFrmTpo As String, psReqNro As String
Dim clsDGnral As DLogGeneral
Dim clsDReq As DLogRequeri
Dim clsDBS As DLogBieSer
Dim b_Nuevo As Boolean
Dim pnRowAnt As Integer

Public Sub Inicio(ByVal psTipoReq As String, ByVal psFormTpo As String, Optional ByVal psRequeriNro As String = "")
psTpoReq = psTipoReq
psFrmTpo = psFormTpo
psReqNro = psRequeriNro
Me.Show 1
End Sub

Private Sub cboAreas_Click()
    Dim rs As ADODB.Recordset
    Dim sReqTraNro As String
    
    sReqTraNro = Right(Trim(cboAreas.Text), 25)
    If cboAreas.ListIndex + 1 = cboAreas.ListCount Then
        fgeBS.lbEditarFlex = True
        fgeBSMes.lbEditarFlex = True
        cmdReqDet(0).Enabled = True
        cmdReqDet(1).Enabled = True
        
        Set rs = New ADODB.Recordset
        'Cargar información del Detalle
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroTramiteUlt, clsDGnral.GetnMovNro(psReqNro))
        If rs.RecordCount > 0 Then Set fgeBS.Recordset = rs
        Set rs = Nothing

        'Cargar información del DetMes
        Set rs = clsDReq.CargaReqDetMes(ReqDetMesUltTraNro, clsDGnral.GetnMovNro(psReqNro))
        If rs.RecordCount > 0 Then Set fgeMes.Recordset = rs
        Set rs = Nothing

        'Actualiza fgeBSDetMes
        Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
        Set rs = Nothing
    Else
        fgeBS.lbEditarFlex = False
        fgeBSMes.lbEditarFlex = False
        cmdReqDet(0).Enabled = False
        cmdReqDet(1).Enabled = False
    
        Set rs = New ADODB.Recordset
        'Cargar información del Detalle
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroTramite, clsDGnral.GetnMovNro(psReqNro), clsDGnral.GetnMovNro(sReqTraNro))
        If rs.RecordCount > 0 Then Set fgeBS.Recordset = rs
        Set rs = Nothing

        'Cargar información del DetMes
        Set rs = clsDReq.CargaReqDetMes(ReqDetMesTraNro, clsDGnral.GetnMovNro(psReqNro), clsDGnral.GetnMovNro(sReqTraNro))
        If rs.RecordCount > 0 Then Set fgeMes.Recordset = rs
        Set rs = Nothing

        'Actualiza fgeBSDetMes
        Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
        Set rs = Nothing
    End If
End Sub

Private Sub cmdIr_Click(Index As Integer)
    Dim nDesRows As Integer, nFluRows As Integer
    Dim nDesRow As Integer, nFluRow As Integer
    nDesRows = fgeDestino.Rows:    nFluRows = fgeFlujo.Rows
    nDesRow = fgeDestino.Row:    nFluRow = fgeFlujo.Row
    
    If psFrmTpo = "1" Then
    ElseIf psFrmTpo = "2" Then
        'INICIO TRAMITE
        rtfObservacion.Text = ""
        If Index = 0 Then
            cmdIr(0).Enabled = False
            cmdIr(1).Enabled = True
            cmdReqFlu(0).Enabled = False
            cmdReqFlu(1).Enabled = True
            cmdReqFlu(2).Enabled = False
            'Ir a la derecha    -  And fgeFlujo.TextMatrix(nFluRows - 1, 3) <> ""
            If fgeDestino.TextMatrix(nDesRow, 1) <> "" And fgeFlujo.TextMatrix(nFluRow, 1) = "" Then
                fgeFlujo.AdicionaFila
                fgeFlujo.TextMatrix(nFluRow, 1) = fgeDestino.TextMatrix(nDesRow, 1)
                fgeFlujo.TextMatrix(nFluRow, 2) = fgeDestino.TextMatrix(nDesRow, 2)
                fgeFlujo.BackColorRow (&H80000018)
                fgeDestino.EliminaFila (nDesRow)
                rtfObservacion.Locked = False
            End If
        Else
            cmdIr(0).Enabled = True
            cmdIr(1).Enabled = False
            cmdReqFlu(0).Enabled = True
            cmdReqFlu(1).Enabled = False
            cmdReqFlu(2).Enabled = True
            'Ir a la izquierda
            If nFluRow = nFluRows - 1 And fgeFlujo.TextMatrix(nFluRow, 1) <> "" _
            And fgeFlujo.TextMatrix(nFluRow, 3) = "" Then
                fgeDestino.AdicionaFila
                nDesRows = fgeDestino.Rows
                fgeDestino.TextMatrix(nDesRows - 1, 1) = fgeFlujo.TextMatrix(nFluRows - 1, 1)
                fgeDestino.TextMatrix(nDesRows - 1, 2) = fgeFlujo.TextMatrix(nFluRows - 1, 2)
                fgeFlujo.BackColorRow (vbWhite)
                fgeFlujo.EliminaFila (nFluRows - 1)
                rtfObservacion.Locked = True
            End If
        End If
    ElseIf psFrmTpo = "3" Then
        'EN TRAMITE
        If Index = 0 Then
            'Ir a la derecha
            If fgeDestino.TextMatrix(nDesRow, 1) <> "" _
             And fgeFlujo.TextMatrix(nFluRows - 1, 3) <> "" Then
                cmdIr(0).Enabled = False
                cmdIr(1).Enabled = True
                cmdReqFlu(0).Enabled = False
                cmdReqFlu(1).Enabled = True
                cmdReqFlu(2).Enabled = False
                
                rtfObservacion.Text = ""
                fgeFlujo.AdicionaFila
                fgeFlujo.TextMatrix(nFluRows, 1) = fgeDestino.TextMatrix(nDesRow, 1)
                fgeFlujo.TextMatrix(nFluRows, 2) = fgeDestino.TextMatrix(nDesRow, 2)
                fgeFlujo.BackColorRow (&HC0FFC0)
                fgeDestino.EliminaFila (nDesRow)
                'rtfObservacion.Locked = False
            End If
        Else
            'Ir a la izquierda
            If nFluRow = nFluRows - 1 And fgeFlujo.TextMatrix(nFluRow, 1) <> "" _
            And fgeFlujo.TextMatrix(nFluRow, 3) = "" Then
                cmdIr(0).Enabled = True
                cmdIr(1).Enabled = False
                cmdReqFlu(0).Enabled = True
                cmdReqFlu(1).Enabled = False
                cmdReqFlu(2).Enabled = True
                
                fgeDestino.AdicionaFila
                nDesRows = fgeDestino.Rows
                fgeDestino.TextMatrix(nDesRows - 1, 1) = fgeFlujo.TextMatrix(nFluRows - 1, 1)
                fgeDestino.TextMatrix(nDesRows - 1, 2) = fgeFlujo.TextMatrix(nFluRows - 1, 2)
                fgeFlujo.BackColorRow (vbWhite)
                fgeFlujo.EliminaFila (nFluRows - 1)
                rtfObservacion.Locked = True
                rtfObservacion.Text = fgeFlujo.TextMatrix(nFluRows - 2, 5)
            End If
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset, rsDes As ADODB.Recordset
    Dim nCont As Integer
    Set clsDGnral = New DLogGeneral
    Set clsDReq = New DLogRequeri
    Set clsDBS = New DLogBieSer
    Set rs = New ADODB.Recordset
    Set rsDes = New ADODB.Recordset
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        txtBuscar.Enabled = False
        cmdReq(0).Enabled = False
        sstReq.Enabled = False
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    Else
        If psFrmTpo = "1" Then
            lblAreaDes.Caption = Usuario.AreaNom
        End If
    End If
    
    If psFrmTpo = "1" Then
        'Para cargar el PERIODO
        Set rs = clsDGnral.CargaPeriodo
        Call CargaCombo(rs, cboPeriodo)
        cboPeriodo.Visible = True
        
        'Carga el CBO con Monedas
        Set rs = clsDGnral.CargaConstante(gMoneda, False)
        Call CargaCombo(rs, cboMoneda)
        cboMoneda.Visible = True
    Else
        lblPeriodo.Visible = True
        lblMoneda.Visible = True
    End If
    
    Call Bloqueo
    
    If psTpoReq = "1" Then      'NORMAL
        If psFrmTpo = "1" Then
            Me.Caption = "Requerimiento Regular : Solicitud "
        ElseIf psFrmTpo = "2" Then
            Me.Caption = "Requerimiento Regular : Trámite : Nuevos"
        ElseIf psFrmTpo = "3" Then
            Me.Caption = ""
        Else
            Me.Caption = "Requerimiento Regular : Trámite : Tramitados"
        End If
    Else                        'EXTEMPORANEO
        If psFrmTpo = "1" Then
            Me.Caption = "Requerimiento Extemporaneo : Solicitud "
        ElseIf psFrmTpo = "2" Then
            Me.Caption = "Requerimiento Extemporaneo : Trámite : Nuevos"
        ElseIf psFrmTpo = "3" Then
            Me.Caption = ""
        Else
            Me.Caption = "Requerimiento Extemporaneo : Trámite : Tramitados"
        End If
    End If
    
    'Carga el TextBusca del Flex(fgeBS) con los bienes/servicios
    'fgeBS.rsTextBuscar = clsDBS.CargaBS(BsTodosArbol)
    fgeBS.rsTextBuscar = oAlmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "")
    
    'Carga Meses
    fgeBSMes.rsFlex = clsDGnral.CargaConstante(gMeses)
    
    If psFrmTpo = "1" Then
        cmdReq(0).Enabled = True
        'Generación y Mantenimiento
        sstReq.TabVisible(2) = False
        'Carga los requerimientos pendientes del area
        Call CargaTxtBuscar
    ElseIf psFrmTpo = "2" Then
        'Envio de Tramite Nuevo
        txtBuscar.Text = psReqNro
        txtBuscar.Enabled = False
        cmdReq(0).Visible = False
        cmdReq(1).Visible = False
        cmdReq(2).Visible = False
        cmdReq(3).Visible = False
        cmdReqDet(0).Visible = False
        cmdReqDet(1).Visible = False
        cmdReqFlu(0).Enabled = True
        cmdReqFlu(1).Enabled = False
        cmdReqFlu(2).Enabled = True
        cmdIr(0).Enabled = True
        cmdIr(1).Enabled = False
        
        'Carga el area a la que se enviará el requerimiento
        If Usuario.AreaTrami = gLogAreaTraEstadoAcepta Then
            Set rs = clsDGnral.CargaAreaSuperior(Usuario.AreaStru, True)
        Else
            Set rs = clsDGnral.CargaAreaSuperior(Usuario.AreaStru)
        End If
        If rs.RecordCount = 0 Then
            MsgBox "Problemas ar cargar información del Area a enviar", vbInformation, " Aviso"
            Exit Sub
        ElseIf rs.RecordCount = 1 Then
            Set fgeDestino.Recordset = rs
        ElseIf rs.RecordCount > 1 Then
            cmdReqFlu(2).Visible = True
            Set fgeDestino.Recordset = rs
        End If
        'Carga Información del requerimiento
        Call txtBuscar_EmiteDatos
    ElseIf psFrmTpo = "3" Then
        'Trámite Ingreso
        txtBuscar.Text = psReqNro
        txtBuscar.Enabled = False
        cmdReq(0).Visible = False
        cmdReq(1).Visible = False
        cmdReq(2).Visible = False
        cmdReq(3).Visible = False
        cmdReqDet(0).Enabled = True
        cmdReqDet(1).Enabled = True
        cmdReqFlu(0).Enabled = True
        cmdReqFlu(1).Enabled = False
        cmdReqFlu(2).Enabled = True
        cmdIr(0).Enabled = True
        cmdIr(1).Enabled = False
        lblAreas.Visible = True
        cboAreas.Visible = True
        cboAreas.Enabled = True
        fgeBS.lbEditarFlex = True
        fgeBSMes.lbEditarFlex = True
        'Carga el Flex que muestra los flujos del requerimiento
        Set rs = clsDReq.CargaReqTramite(ReqTraTodosArea, clsDGnral.GetnMovNro(psReqNro), 0)
        If rs.RecordCount > 0 Then
            Set fgeFlujo.Recordset = rs
            'Carga Cbo para cambios del detalle de las areas
            rs.MoveFirst
            For nCont = 1 To rs.RecordCount
                cboAreas.AddItem rs!cAreaDescripcion & Space(40) & clsDGnral.GetsMovNro(rs!nLogReqTraNro)
                rs.MoveNext
            Next
            'Ademas agrega area actual
            cboAreas.AddItem Usuario.AreaNom    'rs!cAreaDescripcion & Space(40) & rs!cLogReqTraNro
            cboAreas.ListIndex = nCont - 1
            'Carga el Flex que contestará el flujo
            Set rsDes = clsDReq.CargaReqTramite(ReqTraTodosAreaMasDes, clsDGnral.GetnMovNro(psReqNro), 0)
            If rsDes.RecordCount > 0 Then
                fgeFlujo.AdicionaFila
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 1) = rsDes(0)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 2) = rsDes(1)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 3) = rsDes(2)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 4) = rsDes(3)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 5) = rsDes(4)
                fgeFlujo.BackColorRow (&H80000018)
            End If
            'Variable Global de anterior row
            pnRowAnt = fgeFlujo.Row
            
            'Carga el area a la que se enviará el requerimiento
            If Usuario.AreaTrami = gLogAreaTraEstadoAcepta Then
                Set rs = clsDGnral.CargaAreaSuperior(Usuario.AreaStru, True)
            Else
                Set rs = clsDGnral.CargaAreaSuperior(Usuario.AreaStru)
            End If
            If rs.RecordCount = 0 Then
                MsgBox "Problemas ar cargar información del Area a enviar", vbInformation, " Aviso"
                Exit Sub
            ElseIf rs.RecordCount = 1 Then
                Set fgeDestino.Recordset = rs
            ElseIf rs.RecordCount > 1 Then
                cmdReqFlu(2).Visible = True
                Set fgeDestino.Recordset = rs
            End If
        'Else
            'fgeDestino.AdicionaFila
            'fgeDestino.TextMatrix(fgeFlujo.Rows - 1, 1) = Usuario.AreaCod
            'fgeDestino.TextMatrix(fgeFlujo.Rows - 1, 2) = Usuario.AreaNom
        End If
        'Carga Información del requerimiento
        Call txtBuscar_EmiteDatos
    ElseIf psFrmTpo = "4" Then
        'Trámite Egreso
        txtBuscar.Text = psReqNro
        txtBuscar.Enabled = False
        cmdReq(0).Visible = False
        cmdReq(1).Visible = False
        cmdReq(2).Visible = False
        cmdReq(3).Visible = False
        cmdReqDet(0).Visible = False
        cmdReqDet(1).Visible = False
        cmdReqFlu(0).Visible = False
        cmdReqFlu(1).Visible = False
        cmdReqFlu(2).Visible = False
        lblDestino.Visible = False
        fgeDestino.Visible = False
        cmdIr(0).Visible = False
        cmdIr(1).Visible = False
        lblObservacion.Visible = False
        rtfObservacion.Visible = False
        lblFlujo.Left = lblFlujo.Left - 5500
        fgeFlujo.Left = fgeFlujo.Left - 5500
        fgeFlujo.Width = fgeFlujo.Width + 5500
        fgeFlujo.Height = fgeFlujo.Height * 2
        fgeFlujo.EncabezadosAnchos = "400-0-2500-0-1400-5500"
        
        'Carga el Flex que muestra los flujos del requerimiento
        Set rs = clsDReq.CargaReqTramite(ReqTraTodosArea, clsDGnral.GetnMovNro(psReqNro), 0)
        If rs.RecordCount > 0 Then
            Set fgeFlujo.Recordset = rs
            Call fgeFlujo_OnRowChange(fgeFlujo.Row, fgeFlujo.Col)
        
            'Agrega el Flex que contestará el flujo como PENDIENTE
            Set rsDes = clsDReq.CargaReqTramite(ReqTraTodosAreaMasDes, clsDGnral.GetnMovNro(psReqNro), 0)
            If rsDes.RecordCount > 0 Then
                fgeFlujo.AdicionaFila
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 1) = rsDes(0)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 2) = rsDes(1)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 3) = rsDes(2)
                fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 4) = rsDes(3)
            End If
        End If
        'Carga Información del requerimiento
        Call txtBuscar_EmiteDatos
    Else
       MsgBox "Modo de Formulario no reconocido", vbInformation, "Aviso"
       Exit Sub
    End If
    Set rs = Nothing
End Sub


Private Sub cmdReq_Click(Index As Integer)
    Dim clsDMov As DLogMov
    Dim nReqNro As Long, nReqTraNro As Long
    Dim sReqNro As String, sReqTraNro As String, sBSCod As String, sActualiza As String
    Dim nMoneda As Integer
    Dim nRefPrecio As Currency, nCant As Currency
    Dim nBs As Integer, nBSMes As Integer, nResult As Integer
    'Botones de comandos principales
    If Index = 0 Then
        'Nuevo
        b_Nuevo = True
        Call Limpiar
        txtBuscar.Enabled = False
        txtBuscar.Text = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        cmdReq(0).Enabled = False
        cmdReq(1).Enabled = False
        cmdReq(2).Enabled = True
        cmdReq(3).Enabled = True
        cmdReqDet(0).Enabled = True
        cmdReqDet(1).Enabled = True
        rtfDescri(0).Locked = False
        rtfDescri(1).Locked = False
        cboPeriodo.Enabled = True
        cboMoneda.Enabled = True
        sstReq.Tab = 0
        rtfDescri(0).SetFocus
        
    ElseIf Index = 1 Then
        'Editar
        b_Nuevo = False
        cboPeriodo.Enabled = True
        cboMoneda.Enabled = True
        fgeBS.lbEditarFlex = True
        fgeBSMes.lbEditarFlex = True
        txtBuscar.Enabled = False
        cmdReq(0).Enabled = False
        cmdReq(1).Enabled = False
        cmdReq(2).Enabled = True
        cmdReq(3).Enabled = True
        cmdReqDet(0).Enabled = True
        cmdReqDet(1).Enabled = True
        If psFrmTpo = "1" Then
            rtfDescri(0).Locked = False
            rtfDescri(1).Locked = False
            sstReq.Tab = 0
        Else
            cmdReqFlu(0).Enabled = False
            cmdReqFlu(1).Enabled = False
            sstReq.Tab = 1
        End If
    ElseIf Index = 2 Then
        'Cancelar
        If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            If psFrmTpo = "1" Then
                txtBuscar.Enabled = True
                txtBuscar.Text = ""
                cboPeriodo.Enabled = False
                cboMoneda.Enabled = False
                cmdReq(0).Enabled = True
                cmdReq(1).Enabled = False
                cmdReq(2).Enabled = False
                cmdReq(3).Enabled = False
                cmdReqDet(0).Enabled = False
                cmdReqDet(1).Enabled = False
                Call Limpiar
                Call Bloqueo
                Call CargaTxtBuscar
            Else
                Call cmdSalir_Click
            End If
        End If
    ElseIf Index = 3 Then
        'Grabar
        'Valida Periodo
        If Trim(cboPeriodo.Text) = "" Then
            MsgBox "Falta determinar a que Periodo pertenece el requerimiento", vbInformation, " Aviso "
            Exit Sub
        End If
        'Valida Moneda
        If Trim(cboMoneda.Text) = "" Then
            MsgBox "Falta determinar la moneda del requerimiento", vbInformation, " Aviso "
            Exit Sub
        End If
        'Validación BIENSERVICIO
        If Not ValidaBS Then
            Exit Sub
        End If
        
        rtfDescri(0).Text = Replace(rtfDescri(0).Text, "'", "", , , vbTextCompare)
        rtfDescri(1).Text = Replace(rtfDescri(1).Text, "'", "", , , vbTextCompare)
        nMoneda = Val(Right(Trim(cboMoneda.Text), 1))
        
        If Not (nMoneda = gMonedaNacional Or nMoneda = gMonedaExtranjera) Then
            MsgBox "Moneda no reconocida", vbInformation, " Aviso "
            Exit Sub
        End If
        If MsgBox("¿ Estás seguro de Grabar el requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            
            sReqNro = txtBuscar.Text
            If psFrmTpo = "1" Then
                If b_Nuevo Then
                    sReqTraNro = sReqNro
                Else
                    sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                End If
            Else
                sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            End If
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            Set clsDMov = New DLogMov
            
            If b_Nuevo Then
                'Sólo Nuevo
                'Inserta MOV - MOVREF
                clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqRegistro)), "", gLogReqEstadoInicio
                nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                nReqNro = nReqTraNro
                clsDMov.InsertaMovRef nReqTraNro, nReqNro
                
                clsDMov.InsertaRequeri nReqNro, Val(cboPeriodo.Text), IIf(psTpoReq = "1", gLogReqTipoNormal, gLogReqTipoExtemporaneo), _
                     rtfDescri(0).Text, rtfDescri(1).Text
                
                clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                    "", gLogReqEstadoInicio, nMoneda, sActualiza
            Else
                'Sólo Edición
                'Inserta MOV - MOVREF
                clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqModifica)), "", gLogReqEstadoInicio
                nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                nReqNro = clsDMov.GetnMovNro(sReqNro)
                clsDMov.InsertaMovRef nReqTraNro, nReqNro
                
                clsDMov.ActualizaRequeri nReqNro, cboPeriodo.Text, IIf(psTpoReq = "1", gLogReqTipoNormal, gLogReqTipoExtemporaneo), _
                    rtfDescri(0).Text, rtfDescri(1).Text
                'Actualiza a valor con el se graban los detalles (mismo del requerimiento)
                nReqTraNro = nReqNro
                'Elimina solo cuando todavia no se ha iniciado requerimiento
                clsDMov.EliminaReqDetMes nReqNro, nReqTraNro
                clsDMov.EliminaReqDetalle nReqNro, nReqTraNro
                
                'Actualiza ReqTramite (Moneda)
                clsDMov.ActualizaReqTramite nReqNro, nReqTraNro, "", 0, "", nMoneda, sActualiza
            End If

            nBs = 0: nBSMes = 0
            For nBs = 1 To fgeBS.Rows - 1
                sBSCod = fgeBS.TextMatrix(nBs, 1)
                nRefPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 4) = "", 0, fgeBS.TextMatrix(nBs, 4)))
                clsDMov.InsertaReqDetalle nReqNro, nReqTraNro, sBSCod, _
                     nRefPrecio, sActualiza
                For nBSMes = 1 To fgeMes.Cols - 1
                    nCant = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                    If nCant > 0 Then
                        clsDMov.InsertaReqDetMes nReqNro, nReqTraNro, sBSCod, _
                             Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
                    End If
                Next
            Next
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            
            If nResult = 0 Then
                cboPeriodo.Enabled = False
                cboMoneda.Enabled = False
                cmdReq(0).Enabled = True
                cmdReq(1).Enabled = True
                cmdReq(2).Enabled = False
                cmdReq(3).Enabled = False
                cmdReqDet(0).Enabled = False
                cmdReqDet(1).Enabled = False
                txtBuscar.Enabled = True
                Call Bloqueo
                Call CargaTxtBuscar
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    Else
        MsgBox "Comand no reconocido", vbInformation, " Aviso"
    End If
End Sub

Private Sub cmdReqDet_Click(Index As Integer)
    Dim nBSRow As Integer
    'Botones de comandos del detalle de bienes/servicios
    If Index = 0 Then
        'Agregar en Flex
        fgeBS.AdicionaFila
        fgeBS.SetFocus
    ElseIf Index = 1 Then
        'Eliminar en Flex
        nBSRow = fgeBS.Row
        If MsgBox("¿ Estás seguro de eliminar " & fgeBS.TextMatrix(nBSRow, 2) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            fgeMes.EliminaFila nBSRow
            fgeBS.EliminaFila nBSRow
        End If
    End If
End Sub

Private Sub cmdReqFlu_Click(Index As Integer)
    Dim clsDMov As DLogMov
    Dim nReqNro As Integer, nReqTraNro As Integer
    Dim sReqNro As String, sReqTraNro As String, sReqTraNroAnt As String
    Dim sDestino As String, sActualiza As String
    Dim sBSCod As String, sObserva As String
    Dim nBs As Integer, nBSMes As Integer, nResult As Integer, nCont As Integer
    Dim nMoneda As Integer
    Dim nRefPrecio As Currency, nCant As Currency
    If psFrmTpo = "3" Then
        If Val(fgeFlujo.TextMatrix(fgeFlujo.Row, 3)) = gLogReqEstadoParaTramite Then
            'Solo si esta activo "para tramite"
            fgeFlujo.TextMatrix(fgeFlujo.Row, 5) = rtfObservacion.Text
        End If
    End If
    If Index = 1 Then
        If fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 3) <> "" Then
            MsgBox "No se ha seleccionado ningún área", vbInformation, " Aviso "
            Exit Sub
        End If
    End If
    sReqNro = txtBuscar.Text
    nMoneda = Val(Right(Trim(lblMoneda.Caption), 1))
    If Not (nMoneda = gMonedaNacional Or nMoneda = gMonedaExtranjera) Then
        MsgBox "Moneda no reconocida", vbInformation, " Aviso "
        Exit Sub
    End If
    If sReqNro <> "" Then
        Select Case Index
            Case 0:
                'RECHAZAR
                If psFrmTpo = "2" Then
                    sObserva = rtfObservacion.Text
                Else
                    For nCont = 1 To fgeFlujo.Rows - 1
                        If Val(fgeFlujo.TextMatrix(nCont, 3)) = gLogReqEstadoParaTramite Then
                            sObserva = fgeFlujo.TextMatrix(nCont, 5)
                            Exit For
                        End If
                    Next
                End If
                'Validación BIENSERVICIO
                If Not ValidaBS Then
                    Exit Sub
                End If
                
                If MsgBox("¿ Estás seguro de Rechazar el requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    If psFrmTpo = "2" Then
                        'INICIO TRAMITE
                        sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        
                        Set clsDMov = New DLogMov
                        'Inserta MOV - MOVREF
                        clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoRechazado
                        nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                        nReqNro = clsDMov.GetnMovNro(sReqNro)
                        clsDMov.InsertaMovRef nReqTraNro, nReqNro
                        
                        'Actualiza tramite anterior
                        clsDMov.ActualizaReqTramite nReqNro, nReqNro, "", gLogReqEstadoRechazado, _
                            sObserva, nMoneda, sActualiza
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            Call Bloqueo
                        Else
                            MsgBox "Error al rechazar el requerimiento", vbInformation, " Aviso "
                        End If
                    Else
                        'TRAMITE SUCESIVO
                        sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        
                        Set clsDMov = New DLogMov
                        'Inserta Mov - MovRef
                        clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoRechazado
                        nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                        nReqNro = clsDMov.GetnMovNro(sReqNro)
                        clsDMov.InsertaMovRef nReqTraNro, nReqNro
                        
                        'Inserta tramite
                        clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                            sObserva, gLogReqEstadoRechazado, nMoneda, sActualiza
                        
                        'Si no ha modificado detalle, lo agrega tal como está
                        nBs = 0: nBSMes = 0
                        For nBs = 1 To fgeBS.Rows - 1
                            sBSCod = fgeBS.TextMatrix(nBs, 1)
                            nRefPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 4) = "", 0, fgeBS.TextMatrix(nBs, 4)))
                            clsDMov.InsertaReqDetalle nReqNro, nReqTraNro, sBSCod, _
                                nRefPrecio, sActualiza
                            For nBSMes = 1 To fgeMes.Cols - 1
                                nCant = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                                If nCant > 0 Then
                                    clsDMov.InsertaReqDetMes nReqNro, nReqTraNro, sBSCod, _
                                         Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
                                End If
                            Next
                        Next
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            Call Bloqueo
                        Else
                            MsgBox "Error al rechazar el requerimiento", vbInformation, " Aviso "
                        End If
                    End If
                End If
            Case 1:
                'ENVIAR
                If psFrmTpo = "2" Then
                    'INICIO TRAMITE
                    sObserva = rtfObservacion.Text
                    sDestino = Trim(fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 1))
                    
                    If sDestino = "" Then
                        MsgBox "Determine el area a enviar requerimiento", vbInformation, " Aviso"
                        Exit Sub
                    End If
                    
                    If MsgBox("¿ Estás seguro de enviar el requerimiento " & vbCr _
                    & " al " & Trim(fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 2)) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                        sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        'Envio del INICIO
                        Set clsDMov = New DLogMov
                        'Inserta Mov - MovRef
                        clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoInicio
                        nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                        nReqNro = clsDMov.GetnMovNro(sReqNro)
                        clsDMov.InsertaMovRef nReqTraNro, nReqNro
                        
                        'Actualiza tramite anterior
                        clsDMov.ActualizaReqTramite nReqNro, nReqNro, sDestino, gLogReqEstadoInicio, _
                            sObserva, nMoneda, sActualiza
                        
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            Call Bloqueo
                        Else
                            MsgBox "Error al dar inicio al requerimiento", vbInformation, " Aviso "
                        End If
                    End If
                Else
                    'TRAMITES SUCESIVOS
                    For nCont = 1 To fgeFlujo.Rows - 1
                        If Val(fgeFlujo.TextMatrix(nCont, 3)) = gLogReqEstadoParaTramite Then
                            sObserva = fgeFlujo.TextMatrix(nCont, 5)
                            Exit For
                        End If
                    Next
                    sDestino = Trim(fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 1))
                    If sDestino = "" Then
                        MsgBox "Determine el area a enviar requerimiento", vbInformation, " Aviso"
                        Exit Sub
                    End If
                    'Validación BIENSERVICIO
                    If Not ValidaBS Then
                        Exit Sub
                    End If
                    If MsgBox("¿ Estás seguro de Enviar el requerimiento " & vbCr & " a " & Trim(fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 2)) & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                        sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        'Genera tramite para siguiente area
                        Set clsDMov = New DLogMov
                        'Inserta Mov - MovRef
                        clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoVB
                        nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                        nReqNro = clsDMov.GetnMovNro(sReqNro)
                        clsDMov.InsertaMovRef nReqTraNro, nReqNro
                        
                        clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, sDestino, _
                            sObserva, gLogReqEstadoVB, nMoneda, sActualiza
                        
                        'Si no ha modificado detalle, lo agrega tal como está
                        nBs = 0: nBSMes = 0
                        For nBs = 1 To fgeBS.Rows - 1
                            sBSCod = fgeBS.TextMatrix(nBs, 1)
                            nRefPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 4) = "", 0, fgeBS.TextMatrix(nBs, 4)))
                            clsDMov.InsertaReqDetalle nReqNro, nReqTraNro, sBSCod, _
                                nRefPrecio, sActualiza
                            For nBSMes = 1 To fgeMes.Cols - 1
                                nCant = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                                If nCant > 0 Then
                                    clsDMov.InsertaReqDetMes nReqNro, nReqTraNro, sBSCod, _
                                         Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
                                End If
                            Next
                        Next
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            Call Bloqueo
                        Else
                            MsgBox "Error al enviar el requerimiento", vbInformation, " Aviso "
                        End If
                    End If
                End If
            Case 2:
                'ACEPTAR
                'OJO VERIFICAR OBSERVACION
                If psFrmTpo = "2" Then
                    sObserva = rtfObservacion.Text
                Else
                    For nCont = 1 To fgeFlujo.Rows - 1
                        If Val(fgeFlujo.TextMatrix(nCont, 3)) = gLogReqEstadoParaTramite Then
                            sObserva = fgeFlujo.TextMatrix(nCont, 5)
                            Exit For
                        End If
                    Next
                End If
                'Validación BIENSERVICIO
                If Not ValidaBS Then
                    Exit Sub
                End If
                
                If MsgBox("¿ Estás seguro de Aceptar el requerimiento ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    'sReqTraNroAnt = fgeFlujo.TextMatrix(fgeFlujo.Rows - 2, 1)
                    sReqTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    
                    Set clsDMov = New DLogMov
                    'Inserta Mov - MovRef
                    clsDMov.InsertaMov sReqTraNro, Trim(Str(gLogOpeReqTramite)), "", gLogReqEstadoAcepPrevio
                    nReqTraNro = clsDMov.GetnMovNro(sReqTraNro)
                    nReqNro = clsDMov.GetnMovNro(sReqNro)
                    clsDMov.InsertaMovRef nReqTraNro, nReqNro
                    
                    clsDMov.InsertaReqTramite nReqNro, nReqTraNro, Usuario.AreaCod, "", _
                        sObserva, gLogReqEstadoAcepPrevio, nMoneda, sActualiza
                    
                    'Si no ha modificado detalle, lo agrega tal como está
                    nBs = 0: nBSMes = 0
                    For nBs = 1 To fgeBS.Rows - 1
                        sBSCod = fgeBS.TextMatrix(nBs, 1)
                        nRefPrecio = CCur(IIf(fgeBS.TextMatrix(nBs, 4) = "", 0, fgeBS.TextMatrix(nBs, 4)))
                        clsDMov.InsertaReqDetalle nReqNro, nReqTraNro, sBSCod, _
                            nRefPrecio, sActualiza
                        For nBSMes = 1 To fgeMes.Cols - 1
                            nCant = CCur(IIf(fgeMes.TextMatrix(nBs, nBSMes) = "", 0, fgeMes.TextMatrix(nBs, nBSMes)))
                            If nCant > 0 Then
                                clsDMov.InsertaReqDetMes nReqNro, nReqTraNro, sBSCod, _
                                     Val(fgeBSMes.TextMatrix(nBSMes, 1)), nCant
                            End If
                        Next
                    Next
                    
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        Call Bloqueo
                    Else
                        MsgBox "Error al aceptar el requerimiento", vbInformation, " Aviso "
                    End If
                End If
            Case Else
                MsgBox "Caso no reconocido", vbInformation, " Aviso "
        End Select
    End If
End Sub

Private Sub cmdSalir_Click()
    Set clsDGnral = Nothing
    Set clsDReq = Nothing
    Set clsDBS = Nothing
    Unload Me
End Sub

Private Sub fgeBS_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim rsBS As ADODB.Recordset
    'Agregar unidad al Flex
    If Not pbEsDuplicado Then
        Set rsBS = New ADODB.Recordset
        Set rsBS = clsDBS.CargaBS(BsUnRegistro, psDataCod)
        If rsBS.RecordCount > 0 Then fgeBS.TextMatrix(pnRow, 3) = rsBS!cConsUnidad     'cBSUnidad
        Set rsBS = Nothing
    End If
End Sub
Private Sub fgeBS_OnRowAdd(pnRow As Long)
    'Adiciona Fila
    fgeMes.Rows = fgeBS.Rows
    fgeMes.TextMatrix(fgeMes.Rows - 1, 0) = "."
    fgeBS.lbEditarFlex = True
    fgeBSMes.lbEditarFlex = True
    
    Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
End Sub
Private Sub fgeBS_OnRowChange(pnRow As Long, pnCol As Long)
    Dim nCont As Integer
    'Carga Meses del Item de acuerdo al Flex fgeMes
    For nCont = 1 To fgeBSMes.Rows - 1
        fgeBSMes.TextMatrix(nCont, 3) = fgeMes.TextMatrix(pnRow, nCont)
    Next
End Sub
Private Sub fgeBS_OnRowDelete()
    'Borra Fila
    If fgeBS.TextMatrix(fgeBS.Row, 0) = "" And fgeBS.Row = fgeBS.Rows - 1 Then
        fgeBSMes.lbEditarFlex = False
    End If
    
    Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
End Sub
Private Sub fgeBSMes_OnCellChange(pnRow As Long, pnCol As Long)
    fgeMes.TextMatrix(fgeBS.Row, pnRow) = fgeBSMes.TextMatrix(pnRow, pnCol)
End Sub

Private Sub fgeFlujo_OnRowChange(pnRow As Long, pnCol As Long)
    If Trim(fgeFlujo.TextMatrix(pnRow, 1)) <> "" Then
        If psFrmTpo = "3" Then
            'SOLO EN TRAMITE
            If fgeFlujo.TextMatrix(fgeFlujo.Rows - 1, 3) = "" Then
                If Val(fgeFlujo.TextMatrix(fgeFlujo.Row, 3)) = gLogReqEstadoParaTramite Then
                    rtfObservacion.Locked = False
                Else
                    rtfObservacion.Locked = True
                    If Val(fgeFlujo.TextMatrix(pnRowAnt, 3)) = gLogReqEstadoParaTramite Then
                        fgeFlujo.TextMatrix(pnRowAnt, 5) = rtfObservacion.Text
                    End If
                End If
            Else
                rtfObservacion.Locked = True
            End If
            rtfObservacion.Text = fgeFlujo.TextMatrix(pnRow, 5)
        End If
    Else
        rtfObservacion.Text = ""
    End If
    pnRowAnt = fgeFlujo.Row
End Sub

Private Sub Limpiar()
    Dim nCont As Integer
    'Carga los requerimientos pendientes del area
    Call CargaTxtBuscar
    'Otros
    'spinPeriodo.Valor = IIf(psTpoReq = "1", Year(gdFecSis) + 1, Year(gdFecSis))
    cboPeriodo.ListIndex = -1
    cboMoneda.ListIndex = -1
    rtfDescri(0).Text = ""
    rtfDescri(1).Text = ""
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeMes.Clear
    fgeMes.FormaCabecera
    fgeMes.Rows = 2
    For nCont = 1 To fgeBSMes.Rows - 1
        fgeBSMes.TextMatrix(nCont, 3) = ""
    Next
End Sub

Private Sub Bloqueo()
    rtfDescri(0).Locked = True
    rtfDescri(1).Locked = True
    cboPeriodo.Enabled = False
    fgeBS.lbEditarFlex = False
    fgeBSMes.lbEditarFlex = False
    If psFrmTpo <> "1" Then
        'Carga los requerimientos pendientes del area
        cmdReq(1).Enabled = False
        cmdReqDet(0).Enabled = False
        cmdReqDet(1).Enabled = False
        cmdReqFlu(0).Enabled = False
        cmdReqFlu(1).Enabled = False
        cmdReqFlu(2).Enabled = False
        rtfObservacion.Locked = True
        cmdIr(0).Enabled = False
        cmdIr(1).Enabled = False
        cboAreas.Enabled = False
    End If
End Sub


Private Sub rtfObservacion_GotFocus()
    If psFrmTpo = "2" Then
        If fgeFlujo.TextMatrix(fgeFlujo.Row, 1) = "" And cmdIr(0).Enabled = True Then
            rtfObservacion.Locked = False
        End If
    ElseIf psFrmTpo = "3" Then
        If Val(fgeFlujo.TextMatrix(fgeFlujo.Row, 3)) = gLogReqEstadoParaTramite And cmdIr(0).Enabled = True Then
            rtfObservacion.Locked = False
        End If
    End If
End Sub

Private Sub sstReq_Click(PreviousTab As Integer)
    If sstReq.Tab = 2 Then
        'Verificar que el detalle este en el ultimo nivel
        If cboAreas.ListIndex + 1 <> cboAreas.ListCount Then
            MsgBox "Se debe seleccionar el detalle de su area ", vbInformation, " Aviso "
            sstReq.Tab = PreviousTab
        End If
    End If
End Sub

Private Sub txtBuscar_EmiteDatos()
    
    Dim sReqNro As String, sBSCod As String
    Dim rs As ADODB.Recordset
    Dim nCont As Integer
    If psFrmTpo = "1" Then
        If txtBuscar.OK = False Then
            Exit Sub
        End If
    End If
    sReqNro = txtBuscar.Text
    If sReqNro <> "" Then
        If psFrmTpo = "1" Then
            cmdReq(1).Enabled = True
        End If
        Set rs = New ADODB.Recordset
        Set rs = clsDReq.CargaRequerimiento(Val(psTpoReq), ReqUnRegistro, "", clsDGnral.GetnMovNro(sReqNro))
        If rs.RecordCount = 1 Then
            With rs
                lblAreaDes.Caption = !cAreaDescripcion
                If psFrmTpo = "1" Then
                    cboPeriodo.Text = !nLogReqPeriodo
                    For nCont = 0 To cboMoneda.ListCount
                        cboMoneda.ListIndex = nCont
                        If Right(Trim(cboMoneda.Text), 1) = Right(Trim(!cLogReqMoneda), 1) Then
                            Exit For
                        End If
                    Next
                Else
                
                    lblPeriodo.Caption = !nLogReqPeriodo
                    lblMoneda.Caption = !cLogReqMoneda
                End If
                rtfDescri(0).Text = !cLogReqNecesidad
                rtfDescri(1).Text = !cLogReqRequerimiento
            End With
        Else
            cmdReqFlu(0).Enabled = False
            cmdReqFlu(1).Enabled = False
            Set rs = Nothing
            MsgBox "Problemas al cargar información del Requerimiento", vbInformation, " Aviso"
            Exit Sub
        End If
        Set rs = Nothing
        
        'Cargar información del Detalle
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroTramiteUlt, clsDGnral.GetnMovNro(sReqNro))
        If rs.RecordCount > 0 Then Set fgeBS.Recordset = rs
        Set rs = Nothing
        
        'Cargar información del DetMes
        Set rs = clsDReq.CargaReqDetMes(ReqDetMesUltTraNro, clsDGnral.GetnMovNro(sReqNro))
        If rs.RecordCount > 0 Then
            Set fgeMes.Recordset = rs
            For nCont = 1 To fgeMes.Rows - 1
                fgeMes.TextMatrix(nCont, 0) = nCont
            Next
        End If
        Set rs = Nothing
        
        'Actualiza fgeBSDetMes
        Call fgeBS_OnRowChange(fgeBS.Row, fgeBS.Col)
    End If
End Sub

Private Sub CargaTxtBuscar()
    Dim rsReqTree As ADODB.Recordset
    Set rsReqTree = New ADODB.Recordset
    'Carga los requerimientos pendientes del area
    Set rsReqTree = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosAreaFlex, Usuario.AreaCod)
    If rsReqTree.RecordCount > 0 Then
        'VER QUE SI ES UNO (1) HAY VARIACIONES
        txtBuscar.EditFlex = True
        txtBuscar.rs = rsReqTree
        txtBuscar.EditFlex = False
    Else
        txtBuscar.Enabled = False
    End If
    Set rsReqTree = Nothing
End Sub

Private Function ValidaBS() As Boolean
    Dim nBs As Integer, nBSMes As Integer, nCant As Integer
    'Validación de BienesServicios
    ValidaBS = True
    For nBs = 1 To fgeBS.Rows - 1
        If fgeBS.TextMatrix(nBs, 1) = "" Then
            MsgBox "Falta determinar el Bien/Servicio en el Item " & nBs, vbInformation, " Aviso "
            ValidaBS = False
            Exit Function
        End If
        nCant = 0
        For nBSMes = 1 To fgeMes.Cols - 1
            nCant = nCant + Val(fgeMes.TextMatrix(nBs, nBSMes))
        Next
        If nCant = 0 Then
            MsgBox "No se ha determinado las cantidades en los meses en el item " & nBs & " (" & fgeBS.TextMatrix(nBs, 2) & ")", vbInformation, " Aviso "
            ValidaBS = False
            Exit Function
        End If
    Next
End Function
