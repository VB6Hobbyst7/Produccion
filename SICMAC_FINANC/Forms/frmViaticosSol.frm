VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmViaticosSol 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de Arendir de Viáticos"
   ClientHeight    =   6870
   ClientLeft      =   1305
   ClientTop       =   1305
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViaticosSol.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrevio 
      Caption         =   "&Previo"
      Height          =   375
      Left            =   1680
      TabIndex        =   42
      Top             =   6375
      Width           =   1500
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eli&minar"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   6375
      Width           =   1500
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   6375
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   6375
      Width           =   1500
   End
   Begin Sicmact.Usuario user 
      Left            =   4380
      Top             =   6225
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Frame fraRecViaticos 
      Caption         =   "Recibo de Viáticos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   630
      Width           =   9495
      Begin VB.ComboBox cboCategoria 
         Height          =   330
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtNroViatico 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Left            =   6885
         TabIndex        =   3
         Top             =   165
         Width           =   2040
      End
      Begin Sicmact.TxtBuscar txtBuscaPers 
         Height          =   345
         Left            =   1020
         TabIndex        =   2
         Top             =   210
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   615
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "L.E./DNI. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6600
         TabIndex        =   31
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblpersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   30
         Top             =   570
         Width           =   5475
      End
      Begin VB.Label lblNrodoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7365
         TabIndex        =   29
         Top             =   570
         Width           =   1560
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Categoria :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1620
         TabIndex        =   25
         Top             =   1230
         Width           =   5295
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1620
         TabIndex        =   24
         Top             =   900
         Width           =   5280
      End
      Begin VB.Label lblAgecod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   23
         Top             =   1230
         Width           =   585
      End
      Begin VB.Label lblAreaCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   22
         Top             =   900
         Width           =   585
      End
      Begin VB.Label lblDesCargo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4710
         TabIndex        =   21
         Top             =   1560
         Width           =   2205
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   1290
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Area :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cargo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   18
         Top             =   1590
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Persona :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Left            =   6555
         TabIndex        =   16
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame fraDocRef 
      Caption         =   "Documento de Aprobación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   120
      TabIndex        =   12
      Top             =   -15
      Width           =   9450
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   315
         Left            =   2625
         TabIndex        =   1
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNroDocRef 
         Height          =   300
         Left            =   495
         MaxLength       =   15
         TabIndex        =   0
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   210
         Left            =   2040
         TabIndex        =   14
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   255
         Width           =   255
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3600
      Left            =   105
      TabIndex        =   27
      Top             =   2715
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   6350
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Detalle de &Viáticos"
      TabPicture(0)   =   "frmViaticosSol.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "De&talle de Costos"
      TabPicture(1)   =   "frmViaticosSol.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalleCostos"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDetalleCostos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   -74865
         TabIndex        =   33
         Top             =   390
         Width           =   8745
         Begin Sicmact.FlexEdit fgDetCostos 
            Height          =   2385
            Left            =   465
            TabIndex        =   34
            Top             =   210
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   4207
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Concepto-Descripcion-nivel-Importe"
            EncabezadosAnchos=   "450-1000-4000-0-1200"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-L-C-R"
            FormatosEdit    =   "0-0-0-0-2"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   450
            RowHeight0      =   285
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit fgAux 
            Height          =   1260
            Left            =   195
            TabIndex        =   37
            Top             =   1290
            Visible         =   0   'False
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   2223
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-COL1"
            EncabezadosAnchos=   "350-1000"
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
            ColumnasAEditar =   "X-X"
            ListaControles  =   "0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C"
            FormatosEdit    =   "0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   285
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblSubtotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5760
            TabIndex        =   39
            Top             =   2655
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total : "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5160
            TabIndex        =   38
            Top             =   2715
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   15
         TabIndex        =   28
         Top             =   390
         Width           =   9420
         Begin VB.TextBox txtMotivo 
            Height          =   675
            Left            =   105
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            ToolTipText     =   "Motivo de Viatico"
            Top             =   1815
            Width           =   5670
         End
         Begin Sicmact.FlexEdit fgDetViaticos 
            Height          =   1605
            Left            =   105
            TabIndex        =   4
            Top             =   180
            Width           =   9270
            _ExtentX        =   16351
            _ExtentY        =   2831
            Cols0           =   13
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Destino-Lugar-Ida-Vuelta-Partida-Dias-Retorno-Importe-Motivo-cMovNroDet-psOpeCod-nMovNroAtencion"
            EncabezadosAnchos=   "350-1500-1800-1000-1000-900-500-900-1200-0-0-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-5-6-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-3-0-3-3-2-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-L-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-3-0-2-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdeditar 
            Caption         =   "&Editar"
            Height          =   360
            Left            =   1440
            TabIndex        =   6
            Top             =   2565
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   1410
            TabIndex        =   8
            Top             =   2565
            Width           =   1215
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   360
            Left            =   195
            TabIndex        =   5
            Top             =   2565
            Width           =   1215
         End
         Begin VB.CommandButton cmdAceptaDet 
            Caption         =   "&Guardar"
            Height          =   360
            Left            =   195
            TabIndex        =   7
            Top             =   2565
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   6090
            TabIndex        =   36
            Top             =   1905
            Width           =   675
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   345
            Left            =   6795
            TabIndex        =   35
            Top             =   1845
            Width           =   1830
         End
      End
   End
End
Attribute VB_Name = "frmViaticosSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oGen As DGeneral
Dim oOpe As DOperacion
Dim oDoc As DDocumento
Dim oContFunc As NContFunciones
Dim oArendir As NARendir

Dim objPista As COMManejador.Pista

Dim lbSolicitud As Boolean
Dim lsDocTpo As String
Dim lbSalir As Boolean
Dim lbNuevo As Boolean
Dim lnFila As Long
Dim lsMovNroViat As String

'***Agregado por ELRO el 20120321, según OYP-RFC005-2012
Dim fsNroViatico As String
Public Property Get ATNroViatico() As String
ATNroViatico = fsNroViatico
End Property
Public Property Let ATNroViatico(ByVal vNewValue As String)
fsNroViatico = vNewValue
End Property
'***Fin Agregado por ELRO**************************************************

Public Sub Inicio(Optional ByVal pbSolicitud As Boolean = True)
lbSolicitud = pbSolicitud
Me.Show 1
End Sub

'***Agregado por ELRO el 20120321, según OYP-RFC005-2012
Public Sub iniciarEditar(ByVal psNroViatico As String, ByVal pcOpeCod As String)
lbSolicitud = False
ATNroViatico = psNroViatico
gsOpeCod = pcOpeCod
Me.Show 1
End Sub
'***Fin Agregado por ELRO**************************************************

Private Sub cboCategoria_Click()
    RecalculaTodosCostos
End Sub

Private Sub cboCategoria_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 0
    If cmdNuevo.Visible Then
        cmdNuevo.SetFocus
    Else
        fgDetViaticos.SetFocus
    End If
End If
End Sub

Private Sub cmdAceptaDet_Click()
If ValidaGrid = True Then
    'Modificado PASI TI-ERS060-2014
    'fgDetViaticos.TextMatrix(fgDetViaticos.row, 8) = txtMotivo
    If fgDetViaticos.Rows > 2 Then
        If fgDetViaticos.Row >= 2 Then
            If CDate(fgDetViaticos.TextMatrix(fgDetViaticos.Row, 5)) < CDate(fgDetViaticos.TextMatrix(fgDetViaticos.Row - 1, 7)) Then
                If (MsgBox("Las fechas de partida y de retorno  están sobreponiendose con las fecha de partida y retorno del tramo anterior, está seguro que desea continuar", vbInformation + vbYesNo) = vbNo) Then
                    fgDetViaticos.col = 5
                    fgDetViaticos.SetFocus
                    SendKeys "{ENTER}"
                    Exit Sub
                End If
            End If
        End If
     End If
    fgDetViaticos.TextMatrix(fgDetViaticos.Row, 9) = txtMotivo
    'end PASI
    HabilitaControles False
    lbNuevo = False
    fgDetViaticos.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
Dim oImp As NContImprimir
Dim lsImpresion As String

Set oArendir = New NARendir
Set oImp = New NContImprimir
If ValidaInterfaz = False Then Exit Sub
If MsgBox("Desea Grabar la solicitud de Arendir de viáticos", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If lbSolicitud Then
        lsMovNroViat = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    End If
'    If oArendir.GrabaSolicitudViaticos(lsMovNroViat, gdFecSis, nVal(lblTotal), gsOpeCod, _
'                                        txtMotivo, _
'                                        txtNroViatico, gdFecSis, txtNroDocRef, txtfecha, Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2), _
'                                        txtBuscaPers, Val(Right(Trim(cboCategoria), 2)), fgDetViaticos.GetRsNew, _
'                                        fgAux.GetRsNew, Not lbSolicitud) = 0 Then

'***Modificado por ELRO el 20120321, según OYP-RFC005-2012
'If oArendir.GrabaSolicitudViaticos(lsMovNroViat, gdFecSis, nVal(lblTotal), gsOpeCod, _
'                                        txtMotivo, _
'                                        txtNroViatico, txtfecha.Text, txtNroDocRef, txtfecha, Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2), _
'                                        txtBuscaPers, Val(Right(Trim(cboCategoria), 2)), fgDetViaticos.GetRsNew, _
'                                        fgAux.GetRsNew, Not lbSolicitud) = 0 Then
If oArendir.GrabaSolicitudViaticos(lsMovNroViat, gdFecSis, _
                                   nVal(lblTotal), gsOpeCod, _
                                   txtMotivo, txtNroViatico, _
                                   txtfecha.Text, txtNroDocRef, _
                                   txtfecha, "", "", _
                                   txtBuscaPers, Val(Right(Trim(cboCategoria), 2)), _
                                   fgDetViaticos.GetRsNew, fgAux.GetRsNew, Not lbSolicitud) = 0 Then
'***Fin Modificado por ELRO**************************************************
                                   
        lsImpresion = ImprimirReciboViaticoData(gnColPage, gdFecSis, gsOpeCod, _
                                                 gsInstCmac, gsSimbolo, _
                                                 gsNomCmac, gsNomCmacRUC, lsMovNroViat)
        
        EnviaPrevio lsImpresion, Me.Caption, gnLinPage
        objPista.InsertarPista gsOpeCod, lsMovNroViat, gsCodPersUser, GetMaquinaUsuario, "1", "Solicitud de Viaticos"
        If lbSolicitud Then
            If MsgBox("Desea Registrar otra Solicitud de Arendir de Viáticos", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                 'PASI20160210
                    txtBuscaPers = ""
                    lblAgecod = ""
                    lblAgeDesc = ""
                    lblAreaCod = ""
                    lblAreaDesc = ""
                    lblDesCargo = ""
                    lblNrodoc = ""
                    lblpersNombre = ""
                'end PASI****
                txtNroViatico = oContFunc.GeneraDocNro(Int(lsDocTpo))
                fgDetViaticos.Clear
                fgDetViaticos.FormaCabecera
                fgDetViaticos.Rows = 2
                LimpiaCostos
                txtNroDocRef = ""
                txtMotivo = ""
                lblTotal = "0.00"
                lblSubtotal = "0.00"
               
            Else
                '***Agregado por ELRO el 20120321, según OYP-RFC005-2012
                ATNroViatico = ""
                '***Fin Agregado por ELRO**************************************************
                Unload Me
            End If
        Else
            Unload Me
        End If
    End If
End If
Set oImp = Nothing
End Sub
Function ValidaInterfaz() As Boolean
ValidaInterfaz = False

'***Modificado por ELRO el 20120321, según OYP-RFC005-2012
'If Len(Trim(TxtBuscarArendir)) = 0 Then
'    MsgBox "Area que pertenece el Arendir no Ingresado", vbInformation, "Aviso"
'    TxtBuscarArendir.SetFocus
'    Exit Function
'End If
'***Fin Modificado por ELRO**************************************************

If Len(Trim(txtNroDocRef)) = 0 Then
    MsgBox "Documento de Referencia no Ingresado", vbInformation, "Aviso"
    txtNroDocRef.SetFocus
    Exit Function
End If
If ValFecha(txtfecha) = False Then
    Exit Function
End If
If Len(Trim(txtBuscaPers)) < 13 Or lblpersNombre = "" Then
    MsgBox "Código de Persona no Ingresada", vbInformation, "Aviso"
    txtBuscaPers.SetFocus
    Exit Function
End If
If lblAreaCod = "" Then
    MsgBox "Area no Encontrada o no Ingresada. Por favor Consulte con Sistemas ", vbInformation, "Aviso"
    Exit Function
End If
If Len(Trim(cboCategoria)) = 0 Then
    MsgBox "Categoria no Seleccionada", vbInformation, "Aviso"
    cboCategoria.SetFocus
    Exit Function
End If
If fgDetViaticos.TextMatrix(1, 0) = "" Then
    MsgBox "Detalles de viáticos no han sido ingresados", vbInformation, "Aviso"
    cmdNuevo.SetFocus
    Exit Function
End If
If nVal(lblTotal) = 0 Then
    MsgBox "Importe Total debe ser mayor a 0 ", vbInformation, "Aviso"
    cmdNuevo.SetFocus
    Exit Function
End If

ValidaInterfaz = True
End Function

Private Sub cmdCancelar_Click()
If lbNuevo Then
    fgAux.EliminaFila fgDetViaticos.Row
    fgDetViaticos.EliminaFila fgDetViaticos.Row
    If fgDetViaticos.TextMatrix(1, 0) = "" Then
        LimpiaCostos
        txtMotivo = ""
        lblTotal = "0.00"
    Else
        'Modificado PASI TI-ERS060-2014
'        If fgDetViaticos.TextMatrix(fgDetViaticos.row, 9) = "" Then
'            CargaCostos fgDetViaticos.row
'            txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.row, 8)
'            lblTotal = Format(fgDetViaticos.SumaRow(7), "#,#0.00")
'        End If
        
        If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10) = "" Then
            CargaCostos fgDetViaticos.Row
            txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.Row, 9)
            lblTotal = Format(fgDetViaticos.SumaRow(8), "#,#0.00")
        End If
        'end PASI
        
    End If
End If
HabilitaControles False
lbNuevo = False
End Sub

Private Sub cmdEditar_Click()
If fgDetViaticos.TextMatrix(1, 0) = "" Then Exit Sub

'Modificado PASI TI-ERS060-2014
'If fgDetViaticos.TextMatrix(fgDetViaticos.row, 9) <> "" Then
If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10) <> "" Then
'end PASI
    '***Modificado por ELRO el 20120503, según OYP-RFC005-2012
    'If nVal(fgDetViaticos.TextMatrix(fgDetViaticos.Row, 11)) = 0 Then
    '    MsgBox "Concepto de Viatico ya fue Grabado. Puede eliminarlo para agregar nuevo Concepto", vbInformation, "¡Aviso!"
    'Else
    '    MsgBox "Concepto de Viatico ya fue Atendido", vbInformation, "¡Aviso!"
    'End If
    'Exit Sub
    If (gsOpeCod <> CStr(gCGArendirViatSolcEditMN) And gsOpeCod <> CStr(gCGArendirViatSolcEditME)) Then
    
        If nVal(fgDetViaticos.TextMatrix(fgDetViaticos.Row, 12)) = 0 Then 'Modificado PASI TI-ERS060-2014 ; Cambiado Columna 11 por 12
            MsgBox "Concepto de Viatico ya fue Grabado. Puede eliminarlo para agregar nuevo Concepto", vbInformation, "¡Aviso!"
        Else
            MsgBox "Concepto de Viatico ya fue Atendido", vbInformation, "¡Aviso!"
        End If
       Exit Sub
    Else
        'fgDetViaticos.lbEditarFlex = True 'Comentado por TORE 13-04-2018 'Correción del error de las cabeceras del grid editables
    End If
    '***Fin Modificado por ELRO*******************************
End If

HabilitaControles True
lbNuevo = False
lnFila = fgDetViaticos.Row
End Sub

Private Sub cmdEliminar_Click()
Dim lsMsg As String
If fgDetViaticos.TextMatrix(1, 0) = "" Then Exit Sub
If nVal(fgDetViaticos.TextMatrix(fgDetViaticos.Row, 12)) <> 0 Then 'Modificado PASI TI-ERS060-2014; Cambiado Columna 11 por 12
   MsgBox "Detalle de Viaticos no se puede Eliminar. Ya fue atendido", vbInformation, "¡Aviso!"
   Exit Sub
End If
If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10) <> "" Then 'Modificado PASI TI-ERS060-2014; Cambiado Columna 9 por 10
    lsMsg = "¿ Seguro que desea eliminar Detalle de Viaticos Grabado anteriormente ? "
Else
    lsMsg = " ¿Desea eliminar el Detalle del viático Seleccionado? "
End If

If MsgBox(lsMsg, vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10) <> "" Then 'Modificado PASI TI-ERS060-2014; Cambiado Columna 9 por 10
        Dim oMov As New DMov
        oMov.EliminaMov fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10), GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge) 'Modificado PASI TI-ERS060-2014; Cambiado Columna 9 por 10
        Set oMov = Nothing
    Else
        fgAux.EliminaFila fgDetViaticos.Row
    End If
    fgDetViaticos.EliminaFila fgDetViaticos.Row
    If fgDetViaticos.TextMatrix(1, 0) = "" Then
        LimpiaCostos
        txtMotivo = ""
    Else
        CargaCostos fgDetViaticos.Row
        txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.Row, 9) 'Modificado PASI TI-ERS060-2014; Cambiado Columna 8 por 9
        lblTotal = Format(fgDetViaticos.SumaRow(8), gsFormatoNumeroView) 'Modificado PASI TI-ERS060-2014; Cambiado Columna 7 por 8
    End If
End If
End Sub
Private Sub cmdNuevo_Click()
If cboCategoria = "" Then
    MsgBox "Seleccione la categoria para el cálculo respectivo de los costos", vbInformation, "Aviso"
    If cboCategoria.Enabled Then cboCategoria.SetFocus
    Exit Sub
End If
'fgDetViaticos.lbEditarFlex = True 'Comentado por TORE 13-04-2018 'Correción del error de las cabeceras del grid editables
lbNuevo = True
fgDetViaticos.AdicionaFila
lnFila = fgDetViaticos.Row
fgAux.AdicionaFila
fgAux.TextMatrix(fgAux.Row, 1) = "X"
HabilitaControles True
fgDetViaticos.TextMatrix(fgDetViaticos.Row, 5) = gdFecSis
fgDetViaticos.col = 1

'*** PASI 20140301 TI-ERS050-2014
'fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosDestino)
 Dim oConstSist As NConstSistemas
 Set oConstSist = New NConstSistemas
fgDetViaticos.CargaCombo oConstSist.LeeRutasViaticos
'*** END PASI

txtMotivo = ""
fgDetViaticos.SetFocus
SendKeys "{ENTER}"

End Sub
Function ValidaGrid(Optional pbValDetCostos As Boolean = False) As Boolean
ValidaGrid = True
If fgDetViaticos.TextMatrix(1, 0) <> "" Then
    If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 1) = "" Then
        If pbValDetCostos = False Then
            MsgBox "Destino de viático no ingresado", vbInformation, "Aviso"
            fgDetViaticos.col = 1
        End If
        ValidaGrid = False
        '*** PASI 20140301 TI-ERS050-2014
        'fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosDestino)
        Dim oConstSist As NConstSistemas
        Set oConstSist = New NConstSistemas
        fgDetViaticos.CargaCombo oConstSist.LeeRutasViaticos
        '*** END PASI
        Exit Function
    End If
    If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 3) = "" Then
        If pbValDetCostos = False Then
            MsgBox "Transporte de Ida no ingresado", vbInformation, "Aviso"
            fgDetViaticos.col = 3
        End If
        ValidaGrid = False
        fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosTransporte)
        Exit Function
    End If
    If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 4) = "" Then
        If pbValDetCostos = False Then
            MsgBox "Transporte de vuelta no ingresado", vbInformation, "Aviso"
            fgDetViaticos.col = 4
        End If
        ValidaGrid = False
        fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosTransporte)
        Exit Function
    End If
    If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 5) = "" Then
        If pbValDetCostos = False Then
            MsgBox "Fecha de partida no Ingresada o no válida", vbInformation, "Aviso"
            fgDetViaticos.col = 5
        End If
        ValidaGrid = False
        Exit Function
    End If
    If Val(fgDetViaticos.TextMatrix(fgDetViaticos.Row, 6)) = 0 Then
        If pbValDetCostos = False Then
            MsgBox "Numeros de días no ingresados o es igual a cero", vbInformation, "Aviso"
            fgDetViaticos.col = 6
        End If
        ValidaGrid = False
        Exit Function
    End If
    
    If pbValDetCostos = False Then
        If Len(Trim(txtMotivo)) = 0 Then
            If pbValDetCostos = False Then
                MsgBox "Motivo no Ingresado", vbInformation, "Aviso"
                ValidaGrid = False
            End If
            txtMotivo.SetFocus
            Exit Function
        End If
    Else
        If Len(Trim(cboCategoria)) = 0 Then
            If pbValDetCostos = False Then
                MsgBox "Categoria no Ingresado para realizar el calculo respectivo", vbInformation, "Aviso"
                ValidaGrid = False
            End If
            cboCategoria.SetFocus
            Exit Function
        End If
    End If
End If
End Function

Private Sub cmdPrevio_Click()
Dim oImp As NContImprimir
Dim lsImpresion As String
Set oImp = New NContImprimir


If fgDetViaticos.TextMatrix(1, 1) = "" Or fgDetCostos.TextMatrix(1, 1) = "" Then
    MsgBox "No hay datos para imprimir", vbInformation, "Aviso"
    Exit Sub
End If
lsImpresion = ImprimirReciboViatico(fgDetViaticos.GetRsNew, fgDetCostos.GetRsNew, fgAux.GetRsNew, _
                gnColPage, gdFecSis, gsOpeCod, gsNomCmac, gsSimbolo, txtNroViatico, _
                txtNroDocRef, txtfecha, lblAreaCod, lblAreaDesc, lblAgecod, lblAgeDesc, _
                txtBuscaPers.Text, lblpersNombre, lblNrodoc, lblDesCargo, cboCategoria, True, _
                fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10), gsNomCmac, gsNomCmacRUC) 'Modificado PASI TI-ERS060-2014; Cambiado Columna 9 por 10 de fgDetViaticos
                
 
EnviaPrevio lsImpresion, Me.Caption, gnLinPage
Set oImp = Nothing

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgDetCostos_OnCellChange(pnRow As Long, pnCol As Long)
    fgAux.TextMatrix(fgDetViaticos.Row, fgDetCostos.Row + 1) = fgDetCostos.TextMatrix(fgDetCostos.Row, 4)
    fgDetViaticos.TextMatrix(fgDetViaticos.Row, 8) = Format(fgDetCostos.SumaRow(4), "#,#0.00") 'Modificado PASI TI-ERS060-2014; Cambiado Columna 7 por 8 de fgDetViaticos
    lblSubtotal = Format(fgDetCostos.SumaRow(4), gsFormatoNumeroView)
    lblTotal = Format(fgDetViaticos.SumaRow(8), gsFormatoNumeroView) 'Modificado PASI TI-ERS060-2014; Cambiado Columna 7 por 8
End Sub


Private Sub fgDetViaticos_GotFocus()
If fgDetViaticos.TextMatrix(1, 0) <> "" Then
   txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.Row, 9) 'Modificado PASI TI-ERS060-2014; Cambiado Columna 8 por 9
   RefrescaCostos fgDetViaticos.Row
End If
End Sub



Private Sub fgDetViaticos_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol <= 6 Then
    If Not fgDetViaticos.TextMatrix(fgDetViaticos.Row, 6) = "" Then 'TORE13042018 - Validación del numero de dias
        If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 6) > 99 Then
            fgDetViaticos.TextMatrix(fgDetViaticos.Row, 6) = ""
            fgDetViaticos.TextMatrix(fgDetViaticos.Row, 7) = ""
            fgDetViaticos.TextMatrix(fgDetViaticos.Row, 8) = ""
            MsgBox "N° de días no válido.", vbInformation, "Aviso" 'End TORE
        Else
            'fgDetViaticos.TextMatrix(fgDetViaticos.Row, 6) = fgDetViaticos.TextMatrix(fgDetViaticos.Row, 6)
            
            If fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10) = "" Then 'Modificado PASI TI-ERS060-2014; Cambiado Columna 9 por 10
                CargaCostos pnRow
            '***Agregado por ELRO el 20120503, según OYP-RFC005-2012
            ElseIf fgDetViaticos.TextMatrix(fgDetViaticos.Row, 10) <> "" And (gsOpeCod <> CStr(gCGArendirViatSolcEditMN) Or gsOpeCod <> CStr(gCGArendirViatSolcEditME)) Then 'Modificado PASI TI-ERS060-2014; Cambiado Columna 9 por 10
                CargaCostos pnRow
            '***Fin Agregado por ELRO*******************************
            End If
            'Agregado PASI TI-ERS060 2014
                ObtenerFechaRetorno pnRow
            'end PASI
            End If
            'Exit Sub
        End If
    End If
End Sub
'Agregado PASI TI-ERS060 2014
 Private Sub ObtenerFechaRetorno(pnRow As Long)
    If fgDetViaticos.TextMatrix(pnRow, 6) <> "" Then
        fgDetViaticos.TextMatrix(pnRow, 7) = DateAdd("d", fgDetViaticos.TextMatrix(pnRow, 6), CDate(fgDetViaticos.TextMatrix(pnRow, 5)))
    End If
 End Sub
'end PASI
Sub CargaCostos(pnRow As Long)
Dim rs As ADODB.Recordset
'***Modificado por ELRO el 20120925, según SATI INC1209250017 y INC1209260002
'Dim lnDestino As ViaticosDestino
Dim lnDestino As String
'***Fin Modificado por ELRO el 20120925**************************************
Dim lnIda As ViaticosTransporte
Dim lnVuelta As ViaticosTransporte
Dim lnCategCod As ViaticosCateg
Dim lnNumDias As Integer
Set rs = New ADODB.Recordset
Dim i As Integer

Set oArendir = New NARendir

If ValidaGrid(True) Then
        '***Modificado por ELRO el 20120925, según SATI INC1209250017 y INC1209260002
        'lnDestino = Val(Right(fgDetViaticos.TextMatrix(pnRow, 1), 2))
        lnDestino = Val(Trim(Right(fgDetViaticos.TextMatrix(pnRow, 1), 4)))
        '***Fin Modificado por ELRO el 20120925**************************************
        lnIda = Val(Right(fgDetViaticos.TextMatrix(pnRow, 3), 2))
        lnVuelta = Val(Right(fgDetViaticos.TextMatrix(pnRow, 4), 2))
        lnCategCod = Val(Right(cboCategoria, 2))
        lnNumDias = Val(fgDetViaticos.TextMatrix(pnRow, 6))
        Set rs = oArendir.GetConceptosViaticos(lnCategCod, lnDestino, lnIda, lnVuelta, lnNumDias)
        Set fgDetCostos.Recordset = rs
        
        'Modificado PASI TI-ERS060-2014
        'fgDetViaticos.TextMatrix(pnRow, 7) = Format(fgDetCostos.SumaRow(4), "#,#0.00")
        fgDetViaticos.TextMatrix(pnRow, 8) = Format(fgDetCostos.SumaRow(4), "#,#0.00")
        'end PASI
        
        lblSubtotal = Format(fgDetCostos.SumaRow(4), "#,#0.00")
        If Not rs.EOF Then
           i = 1
           Do While Not rs.EOF
                i = i + 1
                fgAux.TextMatrix(pnRow, i) = rs!ImporteConcepto
                rs.MoveNext
           Loop
        End If
        rs.Close
        Set rs = Nothing
        'Modificado PASI TI-ERS060-2014
        'lblTotal = Format(fgDetViaticos.SumaRow(7), "#,#0.00")
        lblTotal = Format(fgDetViaticos.SumaRow(8), "#,#0.00")
        'end PASI
        
Set oArendir = Nothing
        
End If
End Sub
'por si cambi el tipo de categoria del viatico
Sub RecalculaTodosCostos()
Dim rs As ADODB.Recordset
Dim oArendir As New NARendir
 '***Modificado por ELRO el 20120925, según SATI INC1209250017 y INC1209260002
'Dim lnDestino As ViaticosDestino
Dim lnDestino As String
'***Fin Modificado por ELRO el 20120925**************************************
Dim lnIda As ViaticosTransporte
Dim lnVuelta As ViaticosTransporte
Dim lnCategCod As ViaticosCateg
Dim lnNumDias As Integer
Set rs = New ADODB.Recordset
Dim i As Integer
Dim j As Integer

If fgDetViaticos.TextMatrix(1, 0) <> "" Then
    For j = 1 To fgDetViaticos.Rows - 1
        '***Modificado por ELRO el 20120925, según SATI INC1209250017 y INC1209260002
        'lnDestino = Val(Right(fgDetViaticos.TextMatrix(j, 1), 2))
        lnDestino = Val(Trim(Right(fgDetViaticos.TextMatrix(j, 1), 4)))
        '***Fin Modificado por ELRO el 20120925**************************************
        lnIda = Val(Right(fgDetViaticos.TextMatrix(j, 3), 2))
        lnVuelta = Val(Right(fgDetViaticos.TextMatrix(j, 4), 2))
        lnCategCod = Val(Right(cboCategoria, 2))
        lnNumDias = Val(fgDetViaticos.TextMatrix(j, 6))
        Set rs = oArendir.GetConceptosViaticos(lnCategCod, lnDestino, lnIda, lnVuelta, lnNumDias)
        Set fgDetCostos.Recordset = rs
        
        'Modificado PASI TI-ERS060-2014
        'fgDetViaticos.TextMatrix(j, 7) = Format(fgDetCostos.SumaRow(4), "#,#0.00")
        fgDetViaticos.TextMatrix(j, 8) = Format(fgDetCostos.SumaRow(4), "#,#0.00")
        'end PASI
        
        lblSubtotal = Format(fgDetCostos.SumaRow(4), "#,#0.00")
        If Not rs.EOF Then
           i = 1
           Do While Not rs.EOF
                i = i + 1
                fgAux.TextMatrix(j, i) = rs!ImporteConcepto
                rs.MoveNext
           Loop
        End If
        rs.Close
        Set rs = Nothing
        Set oArendir = Nothing
    Next
    
    'Modificado PASI TI-ERS060-2014
    'lblTotal = Format(fgDetViaticos.SumaRow(7), "#,#0.00")
    lblTotal = Format(fgDetViaticos.SumaRow(8), "#,#0.00")
    'end PASI
    
End If
End Sub

Private Sub fgDetViaticos_OnRowChange(pnRow As Long, pnCol As Long)
If fgDetViaticos.TextMatrix(1, 0) <> "" Then

   'Modificado PASI TI-ERS060-2014
   'txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.row, 8)
   txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.Row, 9)
   'end PASI
   
   RefrescaCostos fgDetViaticos.Row
End If
End Sub

Private Sub fgDetViaticos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
If pnCol = 5 Then
    If ValidaFecha(fgDetViaticos.TextMatrix(fgDetViaticos.Row, pnCol)) = "" Then
        If CDate(fgDetViaticos.TextMatrix(fgDetViaticos.Row, pnCol)) < gdFecSis Then
            MsgBox "Fecha de partida no puede ser menor que la del sistema", vbInformation, "Aviso"
            Cancel = False
        End If
    End If
End If
End Sub

Private Sub fgDetViaticos_RowColChange()
'Comentado por TORE 13-042018 ' Correción para cargar los controles desplegables correctos en las grillas correctas.
'If fgDetViaticos.Row <> lnFila Then fgDetViaticos.Row = lnFila
'Select Case fgDetViaticos.col
'Case 1
'    '*** PASI 20140301 TI-ERS050-2014
'    'fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosDestino)
'    Dim oConstSist As NConstSistemas
'    Set oConstSist = New NConstSistemas
'    fgDetViaticos.CargaCombo oConstSist.LeeRutasViaticos
'    '*** END PASI
'Case 3, 4
'    fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosTransporte)
'End Select 'End TORE
If fgDetViaticos.TextMatrix(1, 0) <> "" Then
   'Modificado PASI TI-ERS060-2014
   'txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.row, 8)
   txtMotivo = fgDetViaticos.TextMatrix(fgDetViaticos.Row, 9)
   'end PASI
   RefrescaCostos fgDetViaticos.Row
End If
End Sub

Private Sub fgDetViaticos_Click()
    If fgDetViaticos.col = 1 Then
        CargarDestino
    ElseIf fgDetViaticos.col = 3 Or fgDetViaticos.col = 4 Then
        CargarIdaVuelta
    End If
End Sub

Private Sub fgDetViaticos_KeyPress(KeyAscii As Integer)
    If fgDetViaticos.col = 1 Then
        CargarDestino
    ElseIf fgDetViaticos.col = 3 Or fgDetViaticos.col = 4 Then
        CargarIdaVuelta
    End If
End Sub


'TORE13042017 'Carga correcta de los controles desplegables en las grillas correctas
Private Sub CargarDestino()
    '*** PASI 20140301 TI-ERS050-2014
    'fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosDestino)
    Dim oConstSist As NConstSistemas
    Set oConstSist = New NConstSistemas
    fgDetViaticos.CargaCombo oConstSist.LeeRutasViaticos
    '*** END PASI
End Sub
Private Sub CargarIdaVuelta()
    fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosTransporte)
End Sub
'End TORE


Private Sub Form_Activate()
If lbSalir = True Then
    Unload Me
Else
    '***Agregado por ELRO el 20120321, según OYP-RFC005-2012
    If gsOpeCod = gCGArendirViatSolcEditMN Or gsOpeCod = gCGArendirViatSolcEditME Then
       txtNroViatico = ATNroViatico
       txtNroViatico_KeyPress (13)
       txtNroViatico.Enabled = False
       fraDocRef.Enabled = True
       txtfecha.Enabled = False
     End If
   '***Fin Agregado por ELRO**************************************************
End If
End Sub
Sub HabilitaControles(pbHab As Boolean)
cmdAceptaDet.Visible = pbHab
cmdCancelar.Visible = pbHab
'fgDetViaticos.lbEditarFlex = pbHab 'Comentado por TORE 13-04-2018 'Correción del error de las cabeceras del grid editables
fgDetCostos.lbEditarFlex = pbHab
fgDetViaticos.SoloFila = pbHab
txtMotivo.Locked = Not pbHab
cmdNuevo.Visible = Not pbHab
cmdeditar.Visible = Not pbHab
cmdEliminar.Enabled = Not pbHab
'cmdPrevio.Enabled = Not pbHab
cmdAceptar.Enabled = Not pbHab
End Sub

Private Sub Form_Load()
'***Modificado por ELRO el 20120321, según OYP-RFC005-2012
'Dim oAreas As DActualizaDatosArea
'***Fin Modificado por ELRO**************************************************
Set oGen = New DGeneral
Set oOpe = New DOperacion
Set oDoc = New DDocumento
Set oContFunc = New NContFunciones
Set oArendir = New NARendir
Set objPista = New COMManejador.Pista

CentraForm Me

SSTab1.Tab = 0
LimpiaCostos
lbSalir = False
txtfecha = gdFecSis
'***Modificado por ELRO el 20120321, según OYP-RFC005-2012
'Set oAreas = New DActualizaDatosArea
'TxtBuscarArendir.lbUltimaInstancia = True
'TxtBuscarArendir.psRaiz = "A Rendir de..."
'TxtBuscarArendir.rs = oAreas.GetAgenciasAreas(, 1)
'Set oAreas = Nothing
'***Fin Modificado por ELRO**************************************************


gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", gcMN, gcME)
CentraForm Me
If lbSolicitud Then
    Me.Caption = "Solicitud de Arendir Viaticos"
Else
    '***Modificado por ELRO el 20120321, según OYP-RFC005-2012
    'Me.Caption = "Ampliacion de Arendir Viaticos"
    If gsOpeCod = gCGArendirViatSolcEditMN Or gsOpeCod = gCGArendirViatSolcEditME Then
        Me.Caption = "Editar A Rendir Viaticos"
    Else
        Me.Caption = "Ampliación de A Rendir Viaticos"
    End If
    '***Fin Modificado por ELRO**************************************************
End If
CambiaTamañoCombo cboCategoria
CargaCombo cboCategoria, oGen.GetConstante(gViaticosCateg)

'*** PASI 20140301 TI-ERS050-2014
'fgDetViaticos.CargaCombo oGen.GetConstante(gViaticosDestino)
 Dim oConstSist As NConstSistemas
 Set oConstSist = New NConstSistemas
fgDetViaticos.CargaCombo oConstSist.LeeRutasViaticos
'*** END PASI

fgDetViaticos.TamañoCombo 170
lblTotal = "0.00"
lsDocTpo = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetAutogenerado)
If lsDocTpo = "" Then
    MsgBox "Tipo de documento no definido en esta operación." + vbCrLf + " Por favor consulte a sistemas", vbInformation, "Aviso"
    lbSalir = True
    Exit Sub
End If
If lbSolicitud = False Then
    fraDocRef.Enabled = False
    txtBuscaPers.Enabled = False
    cboCategoria.Enabled = False
    txtNroViatico.Locked = False
    '***Modificado por ELRO el 20120321, según OYP-RFC005-2012
    'fraarendirde.Enabled = False
    '***Fin Modificado por ELRO**************************************************
    cmdPrevio.Enabled = True
Else
    txtNroViatico.Locked = True
    cmdPrevio.Enabled = False
    txtNroViatico = oContFunc.GeneraDocNro(Int(lsDocTpo))
    '***Modificado por ELRO el 20120505, según OYP-RFC005-2012
    'CargaDatosPers gsCodUser
    CargaDatosPers "", gsCodPersUser
    '***Fin Modificado por ELRO*******************************
End If
End Sub
Sub LimpiaCostos()
Dim rs As ADODB.Recordset
Dim i As Integer
Set rs = New ADODB.Recordset
Set rs = oContFunc.GetObjetos(ObjConceptosARendir)
Set fgDetCostos.Recordset = rs
For i = 1 To fgDetCostos.Rows - 1
    fgDetCostos.TextMatrix(i, 4) = "0.00"
Next
fgAux.Clear
fgAux.FormaCabecera
fgAux.Rows = 2
fgAux.Cols = rs.RecordCount + 2  ' SE AGREGA UNA COLUMNA PARA LA INICIAL
If Not rs.EOF Then
   i = 1
   fgAux.ColWidth(1) = 0
   Do While Not rs.EOF  ' i = 1 To rs.RecordCount
        i = i + 1
        fgAux.TextMatrix(0, i) = rs.Fields(0).value
        fgAux.ColWidth(i) = 800
        rs.MoveNext
   Loop
End If
lblSubtotal = "0.00"
rs.Close
Set rs = Nothing

End Sub
Public Function CargaDatosPers(psUser As String, Optional psPersCod As String = "") As Boolean
'***Agregado por ELRO el 20120414, según OYP-RFC005-2012
Dim nSaldoPendienteMN  As Currency
Dim nSaldoPendienteME  As Currency
'***Fin Agregado por ELRO*******************************


CargaDatosPers = False

If psUser <> "" Then
    user.Inicio gsCodUser
    
Else
    user.DatosPers psPersCod
End If

If Trim(user.PersCod) <> "" Then
    '***Agregado por ELRO el 20120414, según OYP-RFC005-2012
    Set oArendir = New NARendir
    Dim RsRendicion  As ADODB.Recordset '********Agregado por PASI20131118 segun TI-ERS107-2013
    Set RsRendicion = New ADODB.Recordset '********Agregado por PASI20131118 segun TI-ERS107-2013
    Call oArendir.obtenerSaldoARendirViaticos(psPersCod, nSaldoPendienteMN, nSaldoPendienteME)
    If nSaldoPendienteMN > 0 Then
        MsgBox PstaNombre(user.UserNom) & " tiene un Saldo pendiente de " & nSaldoPendienteMN & " Nuevo Soles." & Chr(13) & "Primero sustente y/o rinda." & Chr(13) & "Consultar Reglamento de Entregas a Rendir en la Intranet.", vbInformation, "Aviso"
        CargaDatosPers = True
        Exit Function
    ElseIf nSaldoPendienteME > 0 Then
        MsgBox PstaNombre(user.UserNom) & " tiene un Saldo pendiente de " & nSaldoPendienteME & " Dólares." & Chr(13) & "Primero sustente y/o rinda." & Chr(13) & "Consultar Reglamento de Entregas a Rendir en la Intranet.", vbInformation, "Aviso"
        CargaDatosPers = True
        Exit Function
    ElseIf nSaldoPendienteMN = -1 Or nSaldoPendienteME = -1 Then
        MsgBox PstaNombre(user.UserNom) & " tiene una Solicitud pendiente por aprobar." & Chr(13) & "Primero que lo eliminen, para registrar una nueva Solicitud.", vbInformation, "Aviso"
        CargaDatosPers = True
        Exit Function
    End If
    '***Fin Agregado por ELRO*******************************
    
    '********Agregado por PASI20131118 segun TI-ERS107-2013
    Set RsRendicion = oArendir.ObtenerARendirViaticosParaRendirxPersona(psPersCod, "1")
    If RsRendicion.RecordCount > 0 Then
        MsgBox PstaNombre(user.UserNom) & "; tiene pendiente una rendición en Moneda Nacional, no se puede realizar una nueva solicitud hasta que se rinda cuenta de la anterior"
        CargaDatosPers = True
        Exit Function
    End If
    Set RsRendicion = Nothing
    Set RsRendicion = oArendir.ObtenerARendirViaticosParaRendirxPersona(psPersCod, "2")
    If RsRendicion.RecordCount > 0 Then
        MsgBox PstaNombre(user.UserNom) & "; tiene pendiente una rendición en Moneda Extranjera, no se puede realizar una nueva solicitud hasta que se rinda cuenta de la anterior"
        CargaDatosPers = True
        Exit Function
    End If
    Set RsRendicion = Nothing
    '********Fin Agregado por PASI20131118
    
    txtBuscaPers.Text = user.PersCod
    lblpersNombre = PstaNombre(user.UserNom)
    lblAgecod = user.CodAgeAct
    lblAgeDesc = user.DescAgeAct
    lblAreaCod = user.AreaCod
    lblAreaDesc = user.AreaNom
    lblDesCargo = user.PersCargo
    lblNrodoc = user.NroDNIUser
    If user.PersCategCod <> "" And lbSolicitud Then
        cboCategoria.Enabled = False
        cboCategoria = user.PersCategDesc & space(50) & user.PersCategCod
    Else
        If lbSolicitud Then
            cboCategoria.ListIndex = -1
            cboCategoria.Enabled = True
        End If
    End If
    Set oArendir = Nothing
    nSaldoPendienteMN = 0
    nSaldoPendienteME = 0
    CargaDatosPers = True
Else
    Exit Function
End If

'If User.PersCod = "" Then Exit Function 'Comentado por ELRO el 20120401, según OYP-RFC005-2012
'CargaDatosPers = True 'Comentado por ELRO el 20120401, según OYP-RFC005-2012
End Function


Private Sub txtBuscaPers_EmiteDatos()
If CargaDatosPers("", txtBuscaPers) = False Then
    If txtBuscaPers <> "" Then
        MsgBox "Persona Ingresada no se encuentra registrada como empleado de la Institucion", vbInformation, "Aviso"
        Exit Sub
    End If
Else
    If cboCategoria.Enabled Then
        cboCategoria.SetFocus
    Else
        fgDetViaticos.SetFocus
    End If
End If
End Sub
'***Modificado por ELRO el 20120321, según OYP-RFC005-2012
'Private Sub TxtBuscarArendir_EmiteDatos()
'lblDescArendir = Trim(TxtBuscarArendir.psDescripcion)
'If txtNroDocRef.Enabled And txtNroDocRef.Visible Then txtNroDocRef.SetFocus
'End Sub
'***Fin Modificado por ELRO**************************************************

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cboCategoria.Enabled Then
        cboCategoria.SetFocus
    Else
        cmdNuevo.SetFocus
    End If
End If
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdAceptaDet.Visible Then
        cmdAceptaDet.SetFocus
    Else
        cmdNuevo.SetFocus
    End If
End If
End Sub

Private Sub txtNroDocRef_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
    '***Modificado por ELRO el 20120404, según OYP-RFC005-2012
    'txtfecha.SetFocus
    If txtfecha.Enabled Then
        txtfecha.SetFocus
    Else
        cmdNuevo.SetFocus
    End If
    '***Fin Modificado por ELRO***********************************
End If
End Sub
Sub RefrescaCostos(pnRow As Long)
Dim i As Integer
fgAux.Row = pnRow
For i = 2 To fgAux.Cols - 1
    fgDetCostos.TextMatrix(i - 1, 4) = Format(fgAux.TextMatrix(pnRow, i), "#,#0.00")
Next
lblSubtotal = Format(fgDetCostos.SumaRow(4), "#,#0.00")
End Sub
Function CargaDatosViaticos(ByVal psNroViaticos As String) As Boolean
Dim rsDatos As ADODB.Recordset
Dim rsDet As ADODB.Recordset
Dim rsCostos As ADODB.Recordset
Dim lnFilaDet As Long
Dim lnFilaAux As Long
Dim i As Long

CargaDatosViaticos = False
Set rsDatos = New ADODB.Recordset
Set rsDet = New ADODB.Recordset
Set rsCostos = New ADODB.Recordset

txtBuscaPers = ""
lblAgecod = ""
lblAgeDesc = ""
lblAreaCod = ""
lblAreaDesc = ""
lblDesCargo = ""
lblNrodoc = ""
lblpersNombre = ""
txtMotivo = ""
lblSubtotal = "0.00"
lblTotal = "0.00"

fgDetViaticos.Clear
fgDetViaticos.FormaCabecera
fgDetViaticos.Rows = 2

LimpiaCostos
Set rsDatos = oArendir.GetDatosViaticos(psNroViaticos)
If Not rsDatos.EOF And Not rsDatos.EOF Then
    CargaDatosViaticos = True
    lsMovNroViat = rsDatos!cMovNro
    txtNroDocRef = rsDatos!NroDocRef
    txtfecha = rsDatos!dDocFecha
    txtBuscaPers = rsDatos!cPersCod
    lblpersNombre = PstaNombre(rsDatos!cpersnombre)
    lblAreaCod = rsDatos!cAreaCod
    lblAreaDesc = rsDatos!cAreaDescripcion
    lblAgecod = rsDatos!cAgeCod
    lblAgeDesc = rsDatos!cAgeDescripcion
    lblDesCargo = rsDatos!cRHCargoDescripcion
    lblNrodoc = IIf(rsDatos!DNI = "", rsDatos!RUC, rsDatos!DNI)
    '***Modificado por ELRO el 20120321, según OYP-RFC005-2012
    'TxtBuscarArendir = rsDatos!cCodAreaArendir + rsDatos!cAgeCodArendir
    'lblDescArendir = rsDatos!cAreaArendir + rsDatos!cAgeArendir
    '***Fin Modificado por ELRO**************************************************
    'Set rsDet = oArendir.GetDetalleViaticos(rsDatos!cMovNro)
    Set rsDet = oArendir.GetDetalleViaticos(rsDatos!nMovNro)
    If Not rsDet.EOF And Not rsDet.BOF Then
        cboCategoria = rsDet!Categoria & space(50) & rsDet!CodCat
        'FUNCIONARIO 1
        Do While Not rsDet.EOF
            fgDetViaticos.AdicionaFila
            lnFilaDet = fgDetViaticos.Row
            '&H00808000&
            fgDetViaticos.BackColorRow &HC0FFFF    ' &HC0C000
            fgDetViaticos.TextMatrix(lnFilaDet, 1) = rsDet!Destino & space(50) & rsDet!CodDestino
            fgDetViaticos.TextMatrix(lnFilaDet, 2) = rsDet!Lugar
            fgDetViaticos.TextMatrix(lnFilaDet, 3) = rsDet!TransIda & space(50) & rsDet!CodTrasnIda
            fgDetViaticos.TextMatrix(lnFilaDet, 4) = rsDet!TrasnVuelta & space(50) & rsDet!CodTransVuelta
            fgDetViaticos.TextMatrix(lnFilaDet, 5) = rsDet!dPartida
            fgDetViaticos.TextMatrix(lnFilaDet, 6) = rsDet!nMovViaticosDias
            'Modificado PASI TI-ERS060-2014
'            fgDetViaticos.TextMatrix(lnFilaDet, 7) = Format(IIf(IsNull(rsDet!nMovMonto), "0.00", rsDet!nMovMonto), "0.00")
'            fgDetViaticos.TextMatrix(lnFilaDet, 8) = rsDet!Motivo
'            fgDetViaticos.TextMatrix(lnFilaDet, 9) = rsDet!cMovNro
'            fgDetViaticos.TextMatrix(lnFilaDet, 10) = rsDet!cOpeCod
'            fgDetViaticos.TextMatrix(lnFilaDet, 11) = rsDet!nMovNroAtencion
            fgDetViaticos.TextMatrix(lnFilaDet, 7) = rsDet!dllegada
            fgDetViaticos.TextMatrix(lnFilaDet, 8) = Format(IIf(IsNull(rsDet!nMovMonto), "0.00", rsDet!nMovMonto), "0.00")
            fgDetViaticos.TextMatrix(lnFilaDet, 9) = rsDet!Motivo
            fgDetViaticos.TextMatrix(lnFilaDet, 10) = rsDet!cMovNro
            fgDetViaticos.TextMatrix(lnFilaDet, 11) = rsDet!cOpeCod
            fgDetViaticos.TextMatrix(lnFilaDet, 12) = rsDet!nMovNroAtencion
            'end PASI
            
            Set rsCostos = oArendir.GetCostosViaticos(rsDet!cMovNro)
            If Not rsCostos.EOF And Not rsCostos.BOF Then
                fgAux.AdicionaFila
                lnFilaAux = fgAux.Row
                fgAux.TextMatrix(lnFilaAux, 1) = "X"
                For i = 2 To fgAux.Cols - 1
                    fgAux.TextMatrix(lnFilaAux, i) = "0.00"
                Next
                Do While Not rsCostos.EOF
                    For i = 1 To fgAux.Cols - 1
                        If fgAux.TextMatrix(0, i) = Trim(rsCostos!cObjetoCod) Then
                            fgAux.TextMatrix(lnFilaAux, i) = rsCostos!Importe
                            Exit For
                        End If
                    Next
                    rsCostos.MoveNext
                Loop
            Else
                fgAux.AdicionaFila
            End If
            rsCostos.Close
            Set rsCostos = Nothing
            rsDet.MoveNext
        Loop
        RefrescaCostos fgAux.Rows - 1
    End If
    rsDet.Close
    Set rsDet = Nothing
    'Modificado PASI TI-ERS060-2014
    'lblTotal = Format(fgDetViaticos.SumaRow(7), "#,#0.00")
    lblTotal = Format(fgDetViaticos.SumaRow(8), "#,#0.00")
    'end PASI
    
End If
rsDatos.Close
Set rsDatos = Nothing
End Function

Private Sub txtNroViatico_GotFocus()
fEnfoque txtNroViatico
End Sub

Private Sub txtNroViatico_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If lbSolicitud = False Then
        If Len(Trim(txtNroViatico)) > 0 Then
            txtNroViatico = Format(txtNroViatico, String(8, "0"))
             If CargaDatosViaticos(txtNroViatico) = False Then
                MsgBox "Nro de Viático no encontrado", vbInformation, "Aviso"
                txtNroViatico = ""
             Else
                 fgDetViaticos.SetFocus
             End If
        End If
    Else
        If cboCategoria.Enabled Then
            cboCategoria.SetFocus
        End If
    End If
End If
End Sub


