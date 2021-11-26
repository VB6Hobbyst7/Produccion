VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GARANTÍA"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmGarantia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDarBajaGarantia 
      Caption         =   "Eliminar Garantía"
      Height          =   375
      Left            =   6600
      TabIndex        =   52
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdVerCreditos 
      Caption         =   "&Ver Créditos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   49
      ToolTipText     =   "Pulse para visualizar los créditos vinculados a la garantía"
      Top             =   7440
      Width           =   1810
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   36
      ToolTipText     =   "Grabar"
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   85
      TabIndex        =   18
      ToolTipText     =   "Nuevo"
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1120
      TabIndex        =   19
      ToolTipText     =   "Editar"
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9660
      TabIndex        =   21
      ToolTipText     =   "Salir"
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8620
      TabIndex        =   20
      ToolTipText     =   "Cancelar"
      Top             =   7440
      Width           =   1000
   End
   Begin TabDlg.SSTab TabGarantia 
      Height          =   5535
      Left            =   75
      TabIndex        =   23
      Top             =   1845
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pertenencia del Bien"
      TabPicture(0)   =   "frmGarantia.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraClasificacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPropiedad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraClasificacionSBS"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Valorización del Bien"
      TabPicture(1)   =   "frmGarantia.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraValorizacion"
      Tab(1).Control(1)=   "fraValorizacionHist"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Trámite Legal"
      TabPicture(2)   =   "frmGarantia.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraHistorialTramites"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraClasificacionSBS 
         Caption         =   "Clasificación"
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
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   10335
         Begin VB.ComboBox cmbTpoBien 
            Height          =   315
            Left            =   8640
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tip.contr.Bien:"
            Height          =   195
            Left            =   7320
            TabIndex        =   54
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label txtIdSupGarantAntDesemb 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4800
            TabIndex        =   48
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label txtIdSupGarant 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   47
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Clas. créditos nuevos:"
            Height          =   195
            Left            =   3120
            TabIndex        =   46
            Top             =   270
            Width           =   1560
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "SBS :"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   270
            Width           =   405
         End
      End
      Begin VB.Frame fraHistorialTramites 
         Caption         =   "Historial de Trámites ante Registros Públicos"
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
         Height          =   4935
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   10335
         Begin VB.CommandButton cmdAnulaLev 
            Caption         =   "Anula Lev"
            Height          =   350
            Left            =   9360
            TabIndex        =   53
            Top             =   1930
            Width           =   900
         End
         Begin VB.CommandButton cmdTramiteLegalInscripcion 
            Caption         =   "&Inscripc."
            Enabled         =   0   'False
            Height          =   345
            Left            =   9360
            TabIndex        =   51
            ToolTipText     =   "Actualizar Inscripción de Trámite Legal"
            Top             =   1560
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdTramiteLegalEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   9360
            TabIndex        =   42
            ToolTipText     =   "Eliminar Trámite Legal"
            Top             =   960
            Width           =   900
         End
         Begin VB.CommandButton cmdTramiteLegalDetalle 
            Caption         =   "&Detalle"
            Height          =   350
            Left            =   9360
            TabIndex        =   41
            ToolTipText     =   "Ver detalle de Trámite Legal"
            Top             =   240
            Width           =   900
         End
         Begin VB.CommandButton cmdTramiteLegalNuevo 
            Caption         =   "&Nuevo"
            Height          =   350
            Left            =   9360
            TabIndex        =   40
            ToolTipText     =   "Nuevo Trámite Legal"
            Top             =   600
            Width           =   900
         End
         Begin SICMACT.FlexEdit feTramiteLegal 
            Height          =   4575
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7223
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   1
            EncabezadosNombres=   "#-Fecha Ult. Act.-Usuario-Tipo Trámite-Oficina Registral-Estado-Gravamen-Index"
            EncabezadosAnchos=   "400-1900-800-2000-1600-1200-1200-0"
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-L-L-L-R-C"
            FormatosEdit    =   "0-0-0-0-0-0-2-0"
            CantEntero      =   15
            TextArray0      =   "#"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraValorizacionHist 
         Caption         =   "Variación Histórica del Valor de la Garantía"
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
         Height          =   3975
         Left            =   -74880
         TabIndex        =   30
         Top             =   1440
         Width           =   10335
         Begin VB.CommandButton cmdValorizacionEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   350
            Left            =   9350
            TabIndex        =   38
            ToolTipText     =   "Eliminar Valorización"
            Top             =   620
            Width           =   900
         End
         Begin VB.CommandButton cmdValorizacionDetalle 
            Caption         =   "&Detalle"
            Enabled         =   0   'False
            Height          =   350
            Left            =   9350
            TabIndex        =   37
            ToolTipText     =   "Ver detalle de Valorización"
            Top             =   240
            Width           =   900
         End
         Begin SICMACT.FlexEdit feValorizacion 
            Height          =   3615
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   6376
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   1
            EncabezadosNombres=   "#-Fecha Ult. Actualiz.-Usuario-Tipo de Valorización-Moneda-VRM/VRA-IndexMatriz"
            EncabezadosAnchos=   "400-1900-800-3200-1200-1200-0"
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-L-C-R-C"
            FormatosEdit    =   "0-5-0-0-0-2-0"
            CantEntero      =   15
            TextArray0      =   "#"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraValorizacion 
         Caption         =   "Nuevo Valor de Garantía"
         ClipControls    =   0   'False
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
         Height          =   855
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   10335
         Begin VB.CommandButton cmdValorizacionNuevo 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   350
            Left            =   3960
            TabIndex        =   15
            ToolTipText     =   "Nuevo"
            Top             =   350
            Width           =   900
         End
         Begin VB.ComboBox cmbValorizacion 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   380
            Width           =   360
         End
      End
      Begin VB.Frame fraPropiedad 
         Caption         =   "Propiedad del Bien"
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
         Height          =   3495
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   10335
         Begin VB.ComboBox cmbDocumentoTpo 
            Height          =   315
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtDocumentoNro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3735
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "txtPrincipal"
            Top             =   600
            Width           =   3090
         End
         Begin VB.TextBox txtEmisorNombre 
            Height          =   300
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Tag             =   "txtPrincipal"
            Top             =   240
            Width           =   6450
         End
         Begin VB.Frame fraPropietarios 
            Caption         =   "Propietarios"
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
            Height          =   2415
            Left            =   120
            TabIndex        =   32
            Top             =   980
            Width           =   10100
            Begin VB.CommandButton cmdPropietarioNuevo 
               Caption         =   "&Nuevo"
               Height          =   350
               Left            =   8110
               TabIndex        =   39
               ToolTipText     =   "Nuevo Propietario"
               Top             =   1975
               Width           =   900
            End
            Begin VB.CommandButton cmdPropietarioEliminar 
               Caption         =   "&Eliminar"
               Height          =   350
               Left            =   9045
               TabIndex        =   11
               ToolTipText     =   "Eliminar Propietario"
               Top             =   1975
               Width           =   900
            End
            Begin VB.CommandButton cmdPropietarioCancelar 
               Caption         =   "&Cancelar"
               Height          =   350
               Left            =   5265
               TabIndex        =   13
               ToolTipText     =   "Cancelar"
               Top             =   1975
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CommandButton cmdPropietarioAceptar 
               Caption         =   "&Aceptar"
               Height          =   350
               Left            =   4320
               TabIndex        =   12
               ToolTipText     =   "Aceptar"
               Top             =   1975
               Visible         =   0   'False
               Width           =   900
            End
            Begin SICMACT.FlexEdit fePropietarios 
               Height          =   1695
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   9840
               _ExtentX        =   17357
               _ExtentY        =   2990
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   1
               VisiblePopMenu  =   -1  'True
               EncabezadosNombres=   "#-Código-Nombre-Relación-Aux"
               EncabezadosAnchos=   "400-1500-5500-2200-0"
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
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-1-X-3-X"
               ListaControles  =   "0-1-0-3-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-L-L-C"
               FormatosEdit    =   "0-0-0-0-0"
               TextArray0      =   "#"
               lbFlexDuplicados=   0   'False
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   3
               ColWidth0       =   405
               RowHeight0      =   300
            End
         End
         Begin SICMACT.TxtBuscar txtEmisorCod 
            Height          =   285
            Left            =   1575
            TabIndex        =   5
            Top             =   240
            Width           =   2070
            _ExtentX        =   3651
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
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin MSMask.MaskEdBox txtConstataFecha 
            Height          =   300
            Left            =   8900
            TabIndex        =   9
            Top             =   600
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emisor:"
            Height          =   195
            Left            =   135
            TabIndex        =   35
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Constantación:"
            Height          =   195
            Left            =   7260
            TabIndex        =   33
            Top             =   645
            Width           =   1560
         End
      End
      Begin VB.Frame fraClasificacion 
         Caption         =   "Clasificación del Bien"
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
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   10335
         Begin VB.ComboBox cmbClasificacion 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   225
            Width           =   2895
         End
         Begin VB.ComboBox cmbBienGarantia 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   225
            Width           =   4215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Bien en Garantía:"
            Height          =   195
            Left            =   4560
            TabIndex        =   25
            Top             =   240
            Width           =   1260
         End
      End
   End
   Begin VB.Frame fraBusca 
      Height          =   1755
      Left            =   100
      TabIndex        =   22
      Top             =   0
      Width           =   10545
      Begin VB.CommandButton cmdBuscaPersona 
         Caption         =   "&Buscar Persona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8590
         TabIndex        =   43
         ToolTipText     =   "Pulse para buscar las garantías de una persona especifica"
         Top             =   250
         Width           =   1810
      End
      Begin VB.TextBox txtNroGarantBusca 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8590
         MaxLength       =   8
         TabIndex        =   1
         Tag             =   "txtPrincipal"
         Text            =   "0"
         ToolTipText     =   "N° de Garantía"
         Top             =   1230
         Width           =   960
      End
      Begin VB.CommandButton cmdBuscaGarantia 
         Caption         =   "&Aplicar"
         Height          =   315
         Left            =   9600
         TabIndex        =   2
         ToolTipText     =   "Busca por Número de Garantía"
         Top             =   1230
         Width           =   810
      End
      Begin MSComctlLib.ListView LstGarantia 
         Height          =   1335
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Garantía"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bien en Garantia"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Relación"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Condición DPF"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "N° Garantía:"
         Height          =   195
         Left            =   8640
         TabIndex        =   50
         Top             =   960
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmGarantia
'** Descripción : Para registro/edición/consulta de Garantías creado segun TI-ERS063-2014
'** Creación : EJVG, 20150105 10:00:00 AM
'****************************************************************************************

Option Explicit

Enum eAccionPropietario
    PropietarioNuevo = 1
    PropietarioEditar = 2
End Enum

Dim fnTipoInicio As eTipoInicioGarantia
Dim oGarantia As UGarantia

Dim fnPropietarioAccion As eAccionPropietario
Dim fnPropietarioNoMoverFila As Integer
Dim fRsPropietarioRelac As ADODB.Recordset
Dim fbFocoGrilla As Boolean

Dim fbAccion As Boolean
Dim fbPermisoTramiteLegal As Boolean
Dim fbNroGarantiaDigitada As Boolean
Dim fbPermisoDarBajaGarantia As Boolean 'JGPA TI-ERS026-2018
Dim fbAnulaLevGarantia As Boolean 'JGPA20190528 ACTA 093 - 2019
Dim fsNumGarant As String

Dim fbGrabar As Boolean
Dim fbDescobertura As Boolean
Dim fbEliminarRegCobCredProcesoXDescob As Boolean
Dim fbVerBienFuturo As Boolean 'CTI5 ERS012020
Dim fbTieneCredito As Boolean 'CTI5 ERS012020
Dim fnContadorMensajeBienFuturo As Integer 'CTI5 ERS012020
Dim fnTpoBienContrato As Integer 'CTI5 ERS012020
Dim fbCredDesembolsado As Boolean 'CTI5 ERS012020
Dim fdVigCredDesembolsado As Date 'CTI5 ERS012020
Dim fsCtaCod As String 'CTI5 ERS012020
Dim fsAgeCodAct As String 'CTI5 ERS012020



Public Sub Registrar()
    fnTipoInicio = RegistrarGarantia
    Label10.Visible = False 'CTI ERS0012020
    cmbTpoBien.Visible = False 'CTI ERS0012020
    Show 1
End Sub

Public Function Editar(Optional ByVal psNumGarant As String = "") As Boolean
    fsNumGarant = psNumGarant
    fnTipoInicio = EditarGarantia
    Label10.Visible = False 'CTI ERS0012020
    cmbTpoBien.Visible = False 'CTI ERS0012020
    Show 1
    Editar = fbGrabar
End Function

Public Sub Consultar(Optional ByVal psNumGarant As String = "")
    fsNumGarant = psNumGarant
    fnTipoInicio = ConsultarGarantia
    Label10.Visible = False 'CTI ERS0012020
    cmbTpoBien.Visible = False 'CTI ERS0012020
    Show 1
End Sub

'Private Sub cmbBienPrenda_Click()
'    Dim nCod As Long
'    If Trim(cmbBienPrenda.Text) = "" Then
'        nCod = 0
'    Else
'        nCod = CLng(Trim(Right(cmbBienPrenda.Text, 4)))
'    End If
'    cmbDocumentoTpo.Clear
'    cmbValorizacion.Clear
'    Call CargaBienPrendaDocsValorizacion(2, nCod)
'End Sub
Private Sub OcultarVerBienFuturo()
    'oGarantia.TpoBienContrato = 1
    If oGarantia.Clasificacion = 1 And oGarantia.BienGarantia = 6 Then
        Label10.Visible = True
        cmbTpoBien.Visible = True
        fbVerBienFuturo = True
        cmbTpoBien.Enabled = True

    Else
        Label10.Visible = False
        cmbTpoBien.Visible = False
        fbVerBienFuturo = False
        cmbTpoBien.Enabled = False
    End If
   
End Sub
Private Sub cmbBienGarantia_Click()
    'CargaTpoDocumento
    'CTI5 ERS0012020****************************
    oGarantia.BienGarantia = val(Trim(Right(cmbBienGarantia.Text, 3)))
    Call OcultarVerBienFuturo
    If fbVerBienFuturo = True Then
        If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
            Call CargaTpoDocumento
        ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
           Call CargaTpoDocumentoBienFuturo
        Else
            cmbDocumentoTpo.Clear
            txtDocumentoNro.Enabled = False
        End If
    Else
        Call CargaTpoDocumento
        cmbDocumentoTpo.Enabled = True
        txtDocumentoNro.Enabled = True
    End If
    cmbTpoBien.ListIndex = -1
    txtDocumentoNro.Text = ""
    '*******************************************
    CargaTpoValorizacion
    
    oGarantia.BienGarantia = val(Trim(Right(cmbBienGarantia.Text, 3)))
    
    If oGarantia.BienGarantia = 16 Then
        txtEmisorCod.Text = gsCodPersCMACT
        txtEmisorCod.psCodigoPersona = gsCodPersCMACT
        If EnfocaControl(txtEmisorCod) Then
            SendKeys "{Enter}"
        End If
    End If
End Sub

Private Sub cmbBienGarantia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not EnfocaControl(txtEmisorCod) Then
            EnfocaControl cmbDocumentoTpo
        End If
    End If
End Sub

Private Sub cmbClasificacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbBienGarantia
    End If
End Sub
'CTI5 ERS012020*******************************************
Private Sub ObtenerCorrelativoXXX()
    
    If val(Trim(Right(cmbDocumentoTpo, 3))) <= 10 Then
        txtDocumentoNro.Enabled = True
        txtDocumentoNro.Text = ""
    Else
        Dim oGarant As New COMNCredito.NCOMGarantia
        Dim rsTpoDocBienFuturo As New ADODB.Recordset
        Dim i As Integer
        txtDocumentoNro.Enabled = False
        Set rsTpoDocBienFuturo = oGarant.ObtenertTpoDocumentoBienFuturo
        Do While Not rsTpoDocBienFuturo.EOF
            If rsTpoDocBienFuturo!nConsValor = val(Trim(Right(cmbDocumentoTpo, 3))) Then
                txtDocumentoNro.Text = rsTpoDocBienFuturo!cCorrelativo
                '1
                 For i = 1 To fePropietarios.rows - 1
                    If val(Trim(Right(fePropietarios.TextMatrix(i, 3), 3))) = GarantiaPropietarioRelacionTitular Then
                       
                        Dim oRsDocTitular As New ADODB.Recordset
                        
                        Dim cNumeroDocTemporal As String
                        cNumeroDocTemporal = oGarant.ObtenertTpoDocumentoBienFuturoxGarantia(oGarantia.GarantID, val(Trim(Right(cmbDocumentoTpo.Text, 3))), "Correlativo")
                        If Right(Trim(cNumeroDocTemporal), 2) <> "-0" Then
                            txtDocumentoNro.Text = cNumeroDocTemporal
                        End If
                        Set oRsDocTitular = oGarant.ObtenerDocumentoIdentidadTitularGarantia(fePropietarios.TextMatrix(i, 1))
                        Do While Not oRsDocTitular.EOF
                            txtDocumentoNro.Text = Replace(txtDocumentoNro.Text, "XXXXXXXXX", Trim(oRsDocTitular!cPersIDnro))
                            
                            oRsDocTitular.MoveNext
                        Loop
                        oRsDocTitular.Close
                    End If
                 Next
                '2
            End If
            rsTpoDocBienFuturo.MoveNext
        Loop
        
        RSClose rsTpoDocBienFuturo
        Set oGarant = Nothing
    End If
    
End Sub
'********************************************************
Private Sub cmbDocumentoTpo_Click()
    'CTI5 ERS012020**************************************
    'oGarantia.DocumentoTpo = val(Trim(Right(cmbDocumentoTpo, 3)))
    If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
        'oGarantia.DocumentoTpo = val(Trim(Right(cmbDocumentoTpo, 3)))
    ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
           Call ObtenerCorrelativoXXX
        'oGarantia.DocumentoTpo = val(Trim(Right(cmbDocumentoTpo, 3)))
    End If
    oGarantia.DocumentoTpo = val(Trim(Right(cmbDocumentoTpo, 3)))
    '****************************************************
End Sub

Private Sub cmbDocumentoTpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtDocumentoNro
    End If
End Sub
'CTI5 ERS012020*******************************************
Private Sub CargaTpoDocumentoBienFuturo()
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim rsTpoDocBienFuturo As New ADODB.Recordset
   
    Set rsTpoDocBienFuturo = oGarant.ObtenertTpoDocumentoBienFuturo
    Call Llenar_Combo_con_Recordset(rsTpoDocBienFuturo, cmbDocumentoTpo)
    
    RSClose rsTpoDocBienFuturo
    Set oGarant = Nothing
End Sub
Private Sub cmbTpoBien_Click()
    
    If fbTieneCredito And oGarantia.TpoBienContrato = 1 Then
        fnContadorMensajeBienFuturo = fnContadorMensajeBienFuturo + 1
        
        If fnContadorMensajeBienFuturo = 1 Then
            MsgBox "Para ganrantias que coberturan créditos, no se puede volver a seleccionar como Bien Futuro", vbInformation, "Aviso"
            cmbTpoBien.ListIndex = IndiceListaCombo(cmbTpoBien, oGarantia.TpoBienContrato)
        End If
       
        If fnContadorMensajeBienFuturo = 2 Then
            fnContadorMensajeBienFuturo = 0
        End If
        Exit Sub
    End If
    
    oGarantia.TpoBienContrato = 0
    If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
        Call CargaTpoDocumento
        txtDocumentoNro.Enabled = True
        txtDocumentoNro.Text = ""
        oGarantia.TpoBienContrato = 1
        txtDocumentoNro.Enabled = True
        cmbDocumentoTpo.Enabled = True
        txtDocumentoNro.MaxLength = 25
        txtConstataFecha.Enabled = True

        If oGarantia.GarantID = "" Then
            txtConstataFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
            oGarantia.FechaConstata = Format(gdFecSis, "dd/mm/yyyy")
        Else
            If txtConstataFecha.Text = "__/__/____" Then
                txtConstataFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
                oGarantia.FechaConstata = Format(gdFecSis, "dd/mm/yyyy")
            End If
        End If
    ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
       Call CargaTpoDocumentoBienFuturo
       txtDocumentoNro.Enabled = False
       cmbDocumentoTpo.Enabled = True
       oGarantia.TpoBienContrato = 2
       txtDocumentoNro.Text = ""
       txtDocumentoNro.MaxLength = 200
       txtConstataFecha.Enabled = False
       If oGarantia.GarantID = "" Then
            txtConstataFecha.Text = "__/__/____"
       Else
            Dim oRsGarantiaDesem As ADODB.Recordset
            Dim oGarant As New COMNCredito.NCOMGarantia
            Set oRsGarantiaDesem = New ADODB.Recordset
            Set oRsGarantiaDesem = oGarant.ObtenerGarantiaCreditoDesemboldado(oGarantia.GarantID)
            If (oRsGarantiaDesem.EOF And oRsGarantiaDesem.BOF) Then
                txtConstataFecha.Text = "__/__/____"
            End If
            RSClose oRsGarantiaDesem
       End If
    End If
End Sub
'********************************************************

Private Sub cmbValorizacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdValorizacionNuevo
    End If
End Sub
'***JGPA20190528 ACTA 093 - 2019
Private Sub cmdAnulaLev_Click()
Dim Index As Integer
    On Error GoTo ErrAnulaLev
    If feTramiteLegal.TextMatrix(1, 0) = "" Then Exit Sub
    Index = feTramiteLegal.TextMatrix(feTramiteLegal.row, 7)
    
    fbAnulaLevGarantia = False
    
    If oGarantia.ObtenerTramiteLegal(oGarantia.IndexUltimoTramiteLegalActiva).nTipoTramite = LevantamientoGarantia Then
        If MsgBox("Eliminar el Levantamiento de la Garantía solo si el cliente no realizó dicho trámite ante los RR.PP.", vbYesNo + vbInformation, "Aviso") = vbYes Then
            Screen.MousePointer = 11
            oGarantia.EliminarTramiteLegal Index
            oGarantia.PermiteNuevoTramiteLegal = False
            SetFlexTramiteLegal
        
            MsgBox "Se procederá a Anular el Levantamiento de la Garantía con fecha " & feTramiteLegal.TextMatrix(feTramiteLegal.row, 1) & " al guardar los datos.", vbInformation, "Aviso"
            fbAnulaLevGarantia = True
            cmdAnulaLev.Enabled = False
            Screen.MousePointer = 0
        Else
            Exit Sub
        End If
    Else
        MsgBox "El último Trámite Legal no es un Levantamiento de Garantía.", vbInformation, "Aviso"
        Exit Sub
    End If
    Exit Sub
ErrAnulaLev:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'***End JGPA20190528

'Private Sub cmdAplicar_Click()
'    If LstGarantia.ListItems.Count = 0 Then
'        MsgBox "No existe Garantia que mostrar ", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    txtNroGarantBusca.Text = Trim(LstGarantia.SelectedItem)
'    cmdBuscaGarantia_Click
'End Sub

Private Sub cmdBuscaGarantia_Click()
    On Error GoTo ErrBusca
    Dim lsBusca As String
    Dim i As Integer
    
    lsBusca = Trim(txtNroGarantBusca.Text)
    
    If Len(lsBusca) <> 8 Then
        MsgBox "Ud debe ingresar el Nro. de Garantía a Buscar", vbInformation, "Aviso"
        EnfocaControl txtNroGarantBusca
        Exit Sub
    ElseIf lsBusca = "00000000" Then
        MsgBox "Ud debe ingresar el Nro. de Garantía a Buscar", vbInformation, "Aviso"
        EnfocaControl txtNroGarantBusca
        Exit Sub
    End If
    
    fbNroGarantiaDigitada = True
    For i = 1 To LstGarantia.ListItems.count
        If LstGarantia.ListItems(i).Text = lsBusca Then
            fbNroGarantiaDigitada = False
            Exit For
        End If
    Next
    Label10.Visible = False 'CTI ERS0012020
    cmbTpoBien.Visible = False 'CTI ERS0012020
    cmdCancelar_Click
    If Not cargarDatos(lsBusca) Then Exit Sub
    Exit Sub
ErrBusca:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdBuscaPersona_Click()
    Call cmdCancelar_Click
    ObtieneGarantiasPersona
    LstGarantia_Click
End Sub

Private Sub cmdCancelar_Click()
    Set oGarantia = New UGarantia
    HabilitaControles False
    LimpiarControles
    LimpiarVariables
    cmdPropietarioCancelar_Click
    
    If fnTipoInicio = RegistrarGarantia Then
        cmdNuevo.Enabled = True
        Caption = "GARANTÍA [ NUEVO ]"
    ElseIf fnTipoInicio = EditarGarantia Then
        cmdEditar.Enabled = True
        Caption = "GARANTÍA [ EDITAR ]"
    ElseIf fnTipoInicio = ConsultarGarantia Then
        Caption = "GARANTÍA [ CONSULTAR ]"
    End If
    TabGarantia.Tab = 0
    CmdGrabar.Enabled = False
    cmdVerCreditos.Enabled = False
    '**JGPA TI-ERS026-2018----------
    cmdDarBajaGarantia.Visible = False
    fbPermisoDarBajaGarantia = False
    '**END JGPA---------------------
    '**JGPA20190528 ACTA 093 - 2019----------------
    cmdAnulaLev.Visible = False
    fbAnulaLevGarantia = False
    '***End JGPA-------------------
    fbAccion = False
    fbVerBienFuturo = False
    fbTieneCredito = False
    fnContadorMensajeBienFuturo = 0
    fnTpoBienContrato = 0
    fbCredDesembolsado = False
    fsCtaCod = ""
    fsAgeCodAct = ""
End Sub
'**JGPA TI-ERS026-2018---------------
Private Sub cmdDarBajaGarantia_Click()
    Dim pro As New COMDCredito.DCOMGarantia
    Dim cob As New COMNCredito.NCOMCredito
    Dim R As New ADODB.Recordset
    Dim cNumCuenta As String
    Dim bDesv As Boolean
    Dim objPista As COMManejador.Pista
    
    If oGarantia.NroCredCancelados > 0 Then
        MsgBox "La garantía tiene un registro histórico de créditos cancelados. No se puede eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If oGarantia.NroCredVig2 > 0 Then
        MsgBox "La garantía está asociado a creditos vigentes o aprobados. No se puede eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If oGarantia.NroCredEnProceso2 > 0 Then
        Set R = pro.ObtenerCreditosEnProcesoxGarantia(oGarantia.GarantID)
        If R.RecordCount < 2 Then
            If MsgBox("La garantía esta asociada al crédito " & R!cCuenta & " con estado" & " " & R!cEstado & " .¿Desea desvincular la garantía al crédito?", vbYesNo, "Aviso") = vbNo Then Exit Sub
        
            cNumCuenta = R!cCuenta
            bDesv = cob.EliminarCoberturaGarantia(cNumCuenta)
        
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gCredRegistrarGravamen, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, , cNumCuenta, gCodigoCuenta
            Set objPista = Nothing

            If bDesv = True Then MsgBox "Se ha desvinculado correctamente el crédito de la garantía", vbInformation, "Aviso"
        Else
            If MsgBox("La garantía esta asociada a más de un crédito en proceso.¿Desea realizar el proceso de desvinculación?", vbYesNo, "Aviso") = vbNo Then Exit Sub
        
            Do While Not R.EOF
                cNumCuenta = R!cCuenta
                bDesv = cob.EliminarCoberturaGarantia(cNumCuenta)
                
                Set objPista = New COMManejador.Pista
                    objPista.InsertarPista gCredRegistrarGravamen, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, , cNumCuenta, gCodigoCuenta
                Set objPista = Nothing
                R.MoveNext
            Loop
        
            If bDesv = True Then
                    MsgBox "El proceso de desvinculación se he realizado correctamente", vbInformation, "Aviso"
                Else
                MsgBox "No se realizó la desvinculación", vbInformation, "Aviso"
                Exit Sub
            End If
          
        End If
        Set cob = Nothing
        R.Close
    
    End If

    If MsgBox("¿Está seguro de eliminar la garantía?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        fbPermisoDarBajaGarantia = False
        Exit Sub
    Else
        MsgBox "La garantía se eliminará al guardar los datos", vbInformation, "Aviso"
        fbPermisoDarBajaGarantia = True
        cmdDarBajaGarantia.Enabled = False
    End If
End Sub
'**END JGPA-------------------------

Private Sub cmdEditar_Click()
    Dim oGarant As New COMNCredito.NCOMGarantia 'CTI5 ERS0012020
    fsCtaCod = "" 'CTI5 ERS0012020
    fsAgeCodAct = "" 'CTI5 ERS0012020
    If oGarantia.GarantID = "" Then
        MsgBox "Ud. debe seleccionar primero a la Garantía", vbInformation, "Aviso"
        Exit Sub
    End If
    If oGarantia.Estado = 11 Then 'Eliminada
        MsgBox "Esta garantía no podrá ser modificada, por estar ELIMINADA.", vbInformation, "Aviso"
        Exit Sub
    End If
    '(Una DJ x Persona) -> Si en el anterior módulo han tenido varias garantías de Tipo Artefacto (Declaración Jurada)-> Todas estás están bloquedas, y deberán registrar una nueva para que trabajen solo esta nueva
    If oGarantia.DJBloqueada = True Then
        MsgBox "Esta garantía no podrá ser modificada por ser Declaración Jurada en el anterior módulo." & Chr(13) & Chr(13) & _
                "Con el nuevo módulo las Declaraciones Juradas serán registradas una vez por cada Titular," & Chr(13) & _
                "y como Nro. de Documento se debe especificar el DOI del titular de la Garantía." & Chr(13) & Chr(13) & _
                "La declaración Jurada debe tener la siguiente especificación:" & Chr(13) & Chr(13) & _
                " - Clasificación SBS    : BIENES MUEBLES" & Chr(13) & _
                " - Bien en Garantía    : ARTEFACTOS Y OTROS ARTÍCULOS DOMÉSTICOS" & Chr(13) & _
                " - Tipo de Documento: DECLARACIÓN JURADA" & Chr(13) & _
                " - Nro. Documento    : (DOI del titular de la Garantía)", vbInformation, "Aviso"
        Exit Sub
    End If
    
    
    If oGarantia.BloqueoLegal Then
        MsgBox "Esta garantía no podrá ser modificada, comuníquese con el Área Legal / Sup. Créditos en Agencias", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'CTI5 ERS0012020***************************************************************
    If val(Trim(Right(cmbTpoBien.Text, 3))) = gTpoBienFuturo Then
         Dim oRsGarantiaDesem As New ADODB.Recordset
         Set oRsGarantiaDesem = oGarant.ObtenerGarantiaCreditoDesemboldado(oGarantia.GarantID)
         
         Dim cNumGarantDesem As String
         txtConstataFecha.Enabled = False
         If Not (oRsGarantiaDesem.EOF And oRsGarantiaDesem.BOF) Then
            cNumGarantDesem = oRsGarantiaDesem!cNumGarant
            fdVigCredDesembolsado = oRsGarantiaDesem!dVigencia
            fsCtaCod = oRsGarantiaDesem!cCtaCod
            fsAgeCodAct = oRsGarantiaDesem!cAgeCodAct
         End If
        
         If Trim(cNumGarantDesem) <> "" Then
             Dim vle1 As New COMDCredito.DCOMGarantia
             Dim rsLe1 As ADODB.Recordset
        
             Set rsLe1 = vle1.VerificaCargosGarantiasDesembolsadosBienFuturo(gsCodUser)
             If Not (rsLe1.BOF And rsLe1.EOF) Then
                fbCredDesembolsado = True
                txtConstataFecha.Enabled = True
             Else
                MsgBox "Esta garantía no podrá ser modificada, comuníquese con el Área Legal / Sup. Créditos en Agencias", vbInformation, "Aviso"
                Exit Sub
             End If
             Set vle1 = Nothing
             RSClose rsLe1
         End If
         Set oGarant = Nothing
     End If
    '*****************************************************************************
    HabilitaControles True
    cmdEditar.Enabled = False
    cmdVerCreditos.Enabled = False
    CmdGrabar.Enabled = True
    
    'JGPA TI-ERS026-2018------------
    'SE agregó para verificar solo legal tenga esta opción
    Dim vle As New COMDCredito.DCOMGarantia
    Dim rsLe As ADODB.Recordset
    
    Set rsLe = vle.VerificaCargoLegal(gsCodUser)
    If Not (rsLe.BOF And rsLe.EOF) Then
        cmdDarBajaGarantia.Visible = True
        cmdDarBajaGarantia.Enabled = True
        '***JGPA20190528 ACTA 093 - 2019
        If oGarantia.ObtenerTramiteLegal(oGarantia.IndexUltimoTramiteLegalActiva).nTipoTramite = LevantamientoGarantia Then
            cmdAnulaLev.Visible = True
            cmdAnulaLev.Enabled = True
        End If
        '***End JGPA20190528
      Else
        cmdDarBajaGarantia.Visible = False
        cmdDarBajaGarantia.Enabled = False
        cmdAnulaLev.Visible = False 'JGPA20190528 ACTA 093 - 2019
    End If
    Set vle = Nothing
    RSClose rsLe
    'End JGPA TI-ERS026-2018------------
    
    fbAccion = True
    
    'feValorizacion_RowColChange
    feValorizacion_OnRowChange feValorizacion.row, feValorizacion.col
    'feTramiteLegal_RowColChange
    feTramiteLegal_OnRowChange feTramiteLegal.row, feTramiteLegal.col
    'CTI5 ERS0012020***********************************************
       
        Dim oRsGarantiaColocaciones As New ADODB.Recordset
        Set oRsGarantiaColocaciones = oGarant.ObtenerGarantiaColocaciones(oGarantia.GarantID)
        If Not (oRsGarantiaColocaciones.BOF And oRsGarantiaColocaciones.EOF) Then
            fbTieneCredito = True
        Else
            fbTieneCredito = False
        End If
        RSClose oRsGarantiaColocaciones
    
        If oGarantia.Clasificacion = 1 And oGarantia.BienGarantia = 6 Then
            If val(Trim(Right(cmbTpoBien.Text, 3))) = gTpoBienConstruido Then
                If fbTieneCredito Then
                    cmbTpoBien.Enabled = False
                    cmbBienGarantia.Enabled = False
                Else
                    cmbTpoBien.Enabled = True
                    cmbBienGarantia.Enabled = False
                End If
            Else
                cmbBienGarantia.Enabled = False
                cmbTpoBien.Enabled = True
                txtDocumentoNro.Enabled = False
                txtConstataFecha.Enabled = False
                If Not fbTieneCredito Then
                    cmbDocumentoTpo.Enabled = True
                    If val(Trim(Right(cmbDocumentoTpo, 3))) = 1 Then
                        txtDocumentoNro.Enabled = True
                    Else
                        txtDocumentoNro.Enabled = False
                    End If
                Else
                     cmbDocumentoTpo.Enabled = False
                     txtDocumentoNro.Enabled = False
                End If
            End If
        Else
            cmbTpoBien.Enabled = True
        End If
    fnTpoBienContrato = oGarantia.TpoBienContrato
    Call validaBienFuturo
    '**************************************************************
End Sub

Private Sub validaBienFuturo()
 If val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
  Dim bEncontrarDatosValorados As Boolean
'  bEncontrarDatosValorados = False
'  Dim Index As Integer
  bEncontrarDatosValorados = False
  Dim objValorizacion As tValorizacion
  objValorizacion = oGarantia.ObtenerValorizacion(oGarantia.IndexUltimaValorizacionActiva)
  If objValorizacion.nValorizacionTpo = TasacionInmobiliaria Then
    bEncontrarDatosValorados = True
  End If
  
'  For Index = 1 To feValorizacion.rows - 1
'    If feValorizacion.TextMatrix(Index, 3) = "TASACIÓN INMOBILIARIA" Then
'        bEncontrarDatosValorados = True
'    Else
'        bEncontrarDatosValorados = False
'    End If
'  Next
  
  If feTramiteLegal.row > 0 Then
    bEncontrarDatosValorados = True
  End If
  
 
'   If bEncontrarDatosValorados Then
'        cmbDocumentoTpo.Enabled = False
'   End If
End If
 
End Sub


Private Sub cmdGrabar_Click()
    Dim objPista As COMManejador.Pista
    Dim bExito As Boolean
    Dim bMostrarMsgGrabar As Boolean
    
    On Error GoTo ErrGrabar
    
    bMostrarMsgGrabar = True
    CmdGrabar.Enabled = False
    If Not ValidarGarantia Then
        CmdGrabar.Enabled = True
        Exit Sub
    End If
    
	'JOEP20210820
    If cmbTpoBien.Visible = True Then
        If Trim(Right(cmbClasificacion.Text, 3)) = "1" And Trim(Right(cmbBienGarantia.Text, 3)) = "6" Then
            MsgBox "La configuración de la garantía solo es aplicable a los productos:" & Chr(13) & _
            "* TECHO PROPIO " & Chr(13) & _
            "* FONDO MI VIVIENDA", vbInformation, "Aviso"
        End If
    End If
    'JOEP20210820			
    If fbEliminarRegCobCredProcesoXDescob Then
        bMostrarMsgGrabar = False
    End If
    
    If bMostrarMsgGrabar Then
        If MsgBox("¿Está seguro de guardar los datos de la Garantía?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            CmdGrabar.Enabled = True
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = 11
    
    oGarantia.CtaCodDesembolsado = ""
    oGarantia.CambiadoBienFuturo = False
    oGarantia.FechaSistema = gdFecSis
     If (fnTpoBienContrato = gTpoBienFuturo And val(Trim(Right(cmbTpoBien.Text, 3))) = gTpoBienConstruido) And fbCredDesembolsado = True Then
         oGarantia.CtaCodDesembolsado = fsCtaCod
         oGarantia.CambiadoBienFuturo = True
         oGarantia.AgeCtaDesembolsado = fsAgeCodAct
     End If
    
    bExito = oGarantia.GrabarDatosGarantia(fbEliminarRegCobCredProcesoXDescob, _
                                            fbPermisoDarBajaGarantia) 'JGPA added fbPermisoDarBajaGarantia TI-ERS026-2018
    Screen.MousePointer = 0
    CmdGrabar.Enabled = True
    
    fbGrabar = bExito
    
    If Not bExito Then
        MsgBox "Ha sucedido un error al grabar los datos de la Garantía, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    
    Set objPista = New COMManejador.Pista
    If fnTipoInicio = RegistrarGarantia Then
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, oGarantia.GarantID, gCodigoGarantia
    ElseIf fnTipoInicio = EditarGarantia Then
        'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , oGarantia.GarantID, gCodigoGarantia 'Comentado por JGPA20180716
        '*** JGPA TI-ERS026-2018
        If fbPermisoDarBajaGarantia = True Then
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, , oGarantia.GarantID, gCodigoGarantia
        ElseIf fbAnulaLevGarantia = True Then
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, "Anulación de Levantamiento de Garantía", oGarantia.GarantID, gCodigoGarantia 'JGPA20190528 ACTA 093 - 2019
        Else
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , oGarantia.GarantID, gCodigoGarantia
        End If
        '*** END JGPA
    End If
    'CTI5 ERS0012020***************************
    If fbTieneCredito = True Then
        If fnTpoBienContrato = gTpoBienFuturo And val(Trim(Right(cmbTpoBien.Text, 3))) = gTpoBienConstruido Then
             objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Cambio de bien futuro a bien construido", oGarantia.GarantID, gCodigoGarantia
        End If
    End If
    '******************************************
    Set objPista = Nothing
    

    
    MsgBox "Se ha grabado con éxito los datos de la Garantía N° " & oGarantia.GarantID, vbInformation, "Aviso"
    
    txtNroGarantBusca.Text = oGarantia.GarantID
    cmdBuscaGarantia_Click

    Exit Sub
ErrGrabar:
    CmdGrabar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdNuevo_Click()
    cmdCancelar_Click
    HabilitaControles True
    cmdNuevo.Enabled = False
    CmdGrabar.Enabled = True
    EnfocaControl cmbClasificacion
    
    fbAccion = True
    fbTieneCredito = False 'CTI ERS0012020
    txtDocumentoNro.Enabled = False 'CTI ERS0012020
    cmbDocumentoTpo.Enabled = False 'CTI ERS0012020
    fsCtaCod = "" 'CTI ERS0012020
    Label10.Visible = False 'CTI ERS0012020
    cmbTpoBien.Visible = False 'CTI ERS0012020
End Sub

Private Sub cmdPropietarioAceptar_Click()
    If Not validaPropietarios Then Exit Sub
    

    If fnPropietarioAccion = PropietarioNuevo Then
        oGarantia.InsertarPropietario fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 1), fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 2), fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 3), GarantFilaNueva
    ElseIf fnPropietarioAccion = PropietarioEditar Then
        oGarantia.ActualizarPropietario fnPropietarioNoMoverFila, fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 1), fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 2), fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 3)
    End If
    
    If val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
        If val(Trim(Right(cmbDocumentoTpo, 3))) <= 10 Then
        Else
            If val(Trim(Right(fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 3), 3))) = GarantiaPropietarioRelacionTitular Then
                Dim oGarant As New COMNCredito.NCOMGarantia
                Dim oRsDocTitular As New ADODB.Recordset
                txtDocumentoNro.Text = oGarant.ObtenertTpoDocumentoBienFuturoxGarantia(oGarantia.GarantID, val(Trim(Right(cmbDocumentoTpo.Text, 3))), "Correlativo")
                
                Set oRsDocTitular = oGarant.ObtenerDocumentoIdentidadTitularGarantia(fePropietarios.TextMatrix(fnPropietarioNoMoverFila, 1))
                Do While Not oRsDocTitular.EOF
                    txtDocumentoNro.Text = Replace(txtDocumentoNro.Text, "XXXXXXXXX", Trim(oRsDocTitular!cPersIDnro))
                    oRsDocTitular.MoveNext
                Loop
                oRsDocTitular.Close
            End If
        
        End If
    End If
    SetFlexPropietario
    EditarPropietario False
    fnPropietarioAccion = -1
    fnPropietarioNoMoverFila = -1
    

End Sub

Private Sub cmdPropietarioCancelar_Click()
    SetFlexPropietario
    EditarPropietario False
    fnPropietarioAccion = -1
    fnPropietarioNoMoverFila = -1
End Sub

Private Sub cmdPropietarioEditar_Click()
    If fePropietarios.TextMatrix(1, 0) = "" Then Exit Sub
       
    EditarPropietario True
    fnPropietarioAccion = PropietarioEditar
    fnPropietarioNoMoverFila = fePropietarios.row
End Sub

Private Sub cmdPropietarioEliminar_Click()
    On Error GoTo ErrEliminar
    If fePropietarios.TextMatrix(1, 0) = "" Then Exit Sub
    
    If fnTipoInicio = EditarGarantia Then
        If oGarantia.ObtenerPropietario(fePropietarios.row).nRelacionTpo = Titular Then
            If oGarantia.ObtenerValorizacion(oGarantia.IndexUltimaValorizacionActiva).nValorizacionTpo = GarantiaAutoliquidable Then
                MsgBox "No se puede eliminar al Titular de la Garantía, ya que tiene una Valorización AutoLiquidable.", vbInformation, "Aviso"
                Exit Sub
            End If
            If oGarantia.NroCredVinc > 0 Then
                MsgBox "No se puede eliminar al Titular de la Garantía, ya que actualmente tiene créditos vinculados", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
    
    If MsgBox("¿Desea quitar a " & fePropietarios.TextMatrix(fePropietarios.row, 2) & " de los Propietarios?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    oGarantia.EliminarPropietario fePropietarios.row
    SetFlexPropietario
    Screen.MousePointer = 0
    Exit Sub
ErrEliminar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdPropietarioNuevo_Click()
    If fePropietarios.TextMatrix(1, 0) <> "" Then
        If Not validaPropietarios() Then Exit Sub
    End If
    
    fePropietarios.AdicionaFila
    fePropietarios.SetFocus
    SendKeys "{ENTER}"
    
    EditarPropietario True
    fnPropietarioAccion = PropietarioNuevo
    fnPropietarioNoMoverFila = fePropietarios.row
End Sub

Private Sub EditarPropietario(ByVal pbEditar As Boolean)
    cmdPropietarioNuevo.Visible = Not pbEditar
    'cmdPropietarioEditar.Visible = Not pbEditar
    cmdPropietarioEliminar.Visible = Not pbEditar
    cmdPropietarioAceptar.Visible = pbEditar
    cmdPropietarioCancelar.Visible = pbEditar
    
    fePropietarios.lbEditarFlex = pbEditar
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmbClasificacion_Click()
    oGarantia.Clasificacion = val(Trim(Right(cmbClasificacion.Text, 3)))
    CargaBienGarantia
    CargarComboTpoBienContrato
    'CTI5 ERS0012020****************************
    Call OcultarVerBienFuturo
    If fbVerBienFuturo = True Then
        If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
            Call CargaTpoDocumento
        ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
           Call CargaTpoDocumentoBienFuturo
        Else
            cmbDocumentoTpo.Clear
        End If
    Else
         Call CargaTpoDocumento
    End If
    txtDocumentoNro.Text = ""
    cmbTpoBien.ListIndex = -1
    '*******************************************
    CargaTpoValorizacion
    
    'If fnTipoInicio = RegistrarGarantia Then
    '    txtEmisorCod.Text = ""
    '    txtEmisorCod.psCodigoPersona = ""
    '    txtEmisorNombre.Text = ""
    '    oGarantia.EmisorCod = ""
    '    oGarantia.EmisorNombre = ""
    '    txtEmisorCod.Enabled = True
    
    '    If oGarantia.Clasificacion = 4 Then 'Titulos Valores solo para Caja Maynas
    '        txtEmisorCod.Text = gsCodPersCMACT
    '        txtEmisorCod.psCodigoPersona = gsCodPersCMACT
    '        txtEmisorNombre.Text = gsNomCmac
    '        oGarantia.EmisorCod = txtEmisorCod.psCodigoPersona
    '        oGarantia.EmisorNombre = gsNomCmac
    '    End If
    'End If
End Sub

Private Sub ObtieneGarantiasPersona()
    Dim oGaran As COMDCredito.DCOMGarantia
    Dim R As ADODB.Recordset
    Dim oPers As COMDPersona.UCOMPersona
    Dim L As ListItem
    Dim i As Integer
        
    LstGarantia.ListItems.Clear
    Set oPers = frmBuscaPersona.Inicio
    If oPers Is Nothing Then Exit Sub
    
    Set oGaran = New COMDCredito.DCOMGarantia
    Set R = oGaran.RecuperaGarantiasxPersona(oPers.sPersCod)
    Set oGaran = Nothing
    
    If Not RSVacio(R) Then
        Caption = "Garantias de Cliente : " & UCase(oPers.sPersNombre)
        Do While Not R.EOF
            Set L = LstGarantia.ListItems.Add(, , R!cNumGarant)
            L.SubItems(1) = Trim(R!cDescripcion)
            L.SubItems(2) = Trim(R!cBienGarantia)
            L.SubItems(3) = Trim(R!cRelacion)
            L.SubItems(4) = Trim(R!cCondicionDPF)
            
            'L.ListSubItems(1).Bold = True
            If R!nmoneda = gMonedaExtranjera Then
                L.ListSubItems(1).ForeColor = RGB(0, 125, 0)
            Else
                L.ListSubItems(1).ForeColor = vbBlack
            End If
            
            R.MoveNext
        Loop
    Else
        MsgBox "La persona seleccionada no cuenta con Garantías", vbInformation, "Aviso"
    End If
    Set oPers = Nothing
End Sub
'CTI5 ERS0012020
Private Sub CargarComboTpoBienContrato()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rsTpoBienContrato As New ADODB.Recordset
    Set rsTpoBienContrato = oCons.RecuperaConstantes(gTipoBienContrato)
    Call Llenar_Combo_con_Recordset(rsTpoBienContrato, cmbTpoBien)
    RSClose rsTpoBienContrato
    Set oCons = Nothing
End Sub
'End
Private Sub CargaControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rsCons As New ADODB.Recordset
    
    'Cargar Clasificación
    Set rsCons = oCons.RecuperaConstantes(gGarantiaTpoClasif)
    Call Llenar_Combo_con_Recordset(rsCons, cmbClasificacion)
    RSClose rsCons
    
    Dim rsTpoBienContrato As New ADODB.Recordset 'CTI5 ERS0012020
    Set rsTpoBienContrato = oCons.RecuperaConstantes(gTipoBienContrato) 'CTI5 ERS0012020
    Call Llenar_Combo_con_Recordset(rsTpoBienContrato, cmbTpoBien) 'CTI5 ERS0012020
    RSClose rsTpoBienContrato
    
    Set oCons = Nothing
End Sub

Private Sub CargaVariables()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim oGen As New COMDConstSistema.DCOMGeneral
            
    Set fRsPropietarioRelac = oCons.RecuperaConstantes(gPersRelGarantia)
    
    'fbPermisoTramiteLegal = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gMantGarantiaTramiteLegal)
    
    Set oCons = Nothing
    Set oGen = Nothing
End Sub

Private Sub CargaBienGarantia()
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim rsGarant As New ADODB.Recordset
    Dim lnCod As Long
    
    If cmbClasificacion.ListIndex = -1 Then
        lnCod = 0
    Else
        lnCod = val(Trim(Right(cmbClasificacion.Text, 3)))
    End If

    Set rsGarant = oGarant.ObtenerConfigGarantClas(lnCod)
    Call Llenar_Combo_con_Recordset(rsGarant, cmbBienGarantia)
    
    RSClose rsGarant
    Set oGarant = Nothing
End Sub

Private Sub CargaTpoDocumento()
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim rsGarant As New ADODB.Recordset
    Dim lnCod As Long
    
    If cmbBienGarantia.ListIndex = -1 Then
        lnCod = 0
    Else
        lnCod = val(Trim(Right(cmbBienGarantia.Text, 3)))
    End If
    
    Set rsGarant = oGarant.ObtenerDocsGarantiasXID(lnCod)
    Call Llenar_Combo_con_Recordset(rsGarant, cmbDocumentoTpo)
    
    txtDocumentoNro.Enabled = True 'CTI5 ERS0032020
    
    RSClose rsGarant
    Set oGarant = Nothing
End Sub

Private Sub CargaTpoValorizacion()
    Dim oGarant As New COMNCredito.NCOMGarantia
    Dim rsGarant As New ADODB.Recordset
    Dim lnCod As Long
    
    If cmbBienGarantia.ListIndex = -1 Then
        lnCod = 0
    Else
        lnCod = val(Trim(Right(cmbBienGarantia.Text, 3)))
    End If

    Set rsGarant = oGarant.ObtenerConfigGarantTpoValor(lnCod)
    Call Llenar_Combo_con_Recordset(rsGarant, cmbValorizacion)
        
    RSClose rsGarant
    Set oGarant = Nothing
End Sub

'Private Sub CargaDocumentos()
'    Dim oGarant As New COMNCredito.NCOMGarantia
'    Dim rsGarant As New ADODB.Recordset
'
'    Set rsGarant = oGarant.ObtenerConfigGarantClas(pnCod)
'    Call Llenar_Combo_con_Recordset(rsGarant, cmbBienGarantia)
'
'    RSClose rsGarant
'    Set oGarant = Nothing
'End Sub

Private Sub cmdTramiteLegalDetalle_Click()
    Dim frm As frmGarantiaTramiteLegal
    Dim lvTramiteLegal As tTramiteLegal
    Dim lvTramiteLegal_ULT As tTramiteLegal
    Dim Index As Integer, i As Integer
    Dim lbVinculado As Boolean
    Dim lnCuenta As Integer
    Dim lbPrimero As Boolean
    Dim ldTramiteLegal As Date
    Dim lbConsultar As Boolean
        
    On Error GoTo ErrDetalle
    
    If feTramiteLegal.TextMatrix(1, 0) = "" Then Exit Sub
    
    lnCuenta = oGarantia.NroTramitesLegalActivas
    lbPrimero = IIf(lnCuenta = 1, True, False)
    Index = feTramiteLegal.TextMatrix(feTramiteLegal.row, 7) 'Indice de la Matriz que hace referencia
    lvTramiteLegal = oGarantia.ObtenerTramiteLegal(Index)
    lbVinculado = lvTramiteLegal.bVinculado
    ldTramiteLegal = lvTramiteLegal.dfecha
    
    Set frm = New frmGarantiaTramiteLegal
    
    'CTI5 ERS012020********************************************
    If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
          oGarantia.TpoBienContrato = 1
    ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
         oGarantia.TpoBienContrato = 2
    Else
         oGarantia.TpoBienContrato = 1
    End If
    '**********************************************************
    
    lbConsultar = (lbVinculado) Or (fnTipoInicio = ConsultarGarantia) Or (Not fbAccion) Or (IIf(Index <> oGarantia.IndexUltimoTramiteLegalActiva, True, False))
    
    If lbConsultar Then
        frm.Consultar oGarantia.GarantID, oGarantia.Moneda, lvTramiteLegal
        Set frm = Nothing
        Exit Sub
    End If
  
    
    lvTramiteLegal_ULT = oGarantia.ObtenerTramiteLegalUltima(ldTramiteLegal)
    If Not frm.Editar(oGarantia.GarantID, lbPrimero, oGarantia.Moneda, lvTramiteLegal, lvTramiteLegal_ULT) Then
        Set frm = Nothing
        Exit Sub
    End If
    
    oGarantia.ActualizarTramiteLegal Index, lvTramiteLegal
    SetFlexTramiteLegal
    
    Set frm = Nothing
    Exit Sub
ErrDetalle:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdTramiteLegalEliminar_Click()
    Dim Index As Integer
    
    On Error GoTo ErrEliminar
    If feTramiteLegal.TextMatrix(1, 0) = "" Then Exit Sub
    
    'EJVG20160226 ***
    If oGarantia.ObtenerValorizacion(oGarantia.IndexUltimaValorizacionActiva).nValorizacionTpo = GravamenFavorOtraIFi Then
        MsgBox "La última valorización es [GRAVAMEN A FAVOR DE OTRA IFI] no se podrá eliminar Trámite Legal.", vbInformation, "Aviso"
        Exit Sub
    End If
    'END EJVG *******
    
    Index = feTramiteLegal.TextMatrix(feTramiteLegal.row, 7)
    
    '*** JGPA TI-ERS026-2018
    If oGarantia.ObtenerTramiteLegal(Index).nEstado = Inscrita Then
        MsgBox "Solo se puede eliminar trámite legal pendiente de inscripción", vbInformation, "Aviso"
        Exit Sub
    End If
    '*** End JGPA
    
    If MsgBox("¿Desea eliminar el Trámite Legal con fecha " & feTramiteLegal.TextMatrix(feTramiteLegal.row, 1) & "?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    oGarantia.EliminarTramiteLegal Index
    oGarantia.PermiteNuevoTramiteLegal = True
    cmdTramiteLegalNuevo.Enabled = oGarantia.PermiteNuevoTramiteLegal()
    SetFlexTramiteLegal
    Screen.MousePointer = 0
    Call validaBienFuturo 'ERS0012020********************
    Exit Sub
ErrEliminar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdTramiteLegalInscripcion_Click()
    Dim frm As frmGarantiaActInscripcion
    Dim lvTramiteLegal As tTramiteLegal
    Dim Index As Integer
    Dim ldTramiteLegal As Date
    Dim bExito As Boolean
        
    On Error GoTo ErrInscripcion
    
    If feTramiteLegal.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = feTramiteLegal.TextMatrix(feTramiteLegal.row, 7) 'Indice de la Matriz que hace referencia
    lvTramiteLegal = oGarantia.ObtenerTramiteLegal(Index)
        
    Set frm = New frmGarantiaActInscripcion
    If Not frm.Inicio(oGarantia.GarantID, lvTramiteLegal) Then
        Set frm = Nothing
        Exit Sub
    End If
    
    oGarantia.ActualizarTramiteLegal Index, lvTramiteLegal
    SetFlexTramiteLegal
    
    Set frm = Nothing
    Exit Sub
ErrInscripcion:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdTramiteLegalNuevo_Click()
    Dim frm As frmGarantiaTramiteLegal
    Dim lvTramiteLegal As tTramiteLegal
    Dim lvTramiteLegal_ULT As tTramiteLegal
    Dim lbExisteValorizacion As Boolean
    Dim lbExisteTramite As Boolean
    Dim Index As Integer, IndexAct As Integer
    Dim lnUltValorizacion As eGarantiaTipoValorizacion
    
    On Error GoTo ErrNuevo
    
    lbExisteValorizacion = IIf(oGarantia.NroValorizacionesActivas > 0, True, False)
    lbExisteTramite = IIf(oGarantia.NroTramitesLegalActivas > 0, True, False)
    
    If gTpoBienFuturo = val(Trim(Right(cmbTpoBien.Text, 3))) Then
         MsgBox "Para tipos de garantias de bien futuro, no se puede crear TRÁMITE LEGAL.", vbInformation, "Aviso"
         TabGarantia.Tab = 0
         Exit Sub
    End If
    
    If Not lbExisteValorizacion Then
        MsgBox "Ud. debe primero ingresar la valorización de la Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 1
        EnfocaControl cmbValorizacion
        Exit Sub
    End If
    
    lvTramiteLegal_ULT = oGarantia.ObtenerTramiteLegalUltima()
    
    'Verificar que el Trámite Legal no esté pendiente
    If Not lvTramiteLegal_ULT.bMigrado Then
        If lvTramiteLegal_ULT.nEstado = Pendiente Then
            MsgBox "Para agregar una nuevo Trámite Legal primero debe de actualizar la Inscripción del último trámite Legal", vbInformation, "Aviso"
            EnfocaControl cmdTramiteLegalInscripcion
            Exit Sub
        End If
    End If
    
    'EJVG20160226 ***
    If oGarantia.ObtenerValorizacion(oGarantia.IndexUltimaValorizacionActiva).nValorizacionTpo = GravamenFavorOtraIFi Then
        MsgBox "La última valorización es [GRAVAMEN A FAVOR DE OTRA IFI] no se podrá realizar Trámite Legal.", vbInformation, "Aviso"
        Exit Sub
    End If
    'END EJVG *******
    
    'CTI5 ERS012020********************************************
    If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
          oGarantia.TpoBienContrato = 1
    ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
         oGarantia.TpoBienContrato = 2
    Else
         oGarantia.TpoBienContrato = 1
    End If
    '**********************************************************
    
    Set frm = New frmGarantiaTramiteLegal
    If Not frm.Registrar(oGarantia.GarantID, Not lbExisteTramite, oGarantia.Moneda, lvTramiteLegal, lvTramiteLegal_ULT) Then
        Set frm = Nothing
        Exit Sub
    End If
    
    oGarantia.InsertarTramiteLegal lvTramiteLegal, GarantFilaNueva, True
    oGarantia.PermiteNuevoTramiteLegal = False
    cmdTramiteLegalNuevo.Enabled = oGarantia.PermiteNuevoTramiteLegal
    SetFlexTramiteLegal
    
    Set frm = Nothing
    
    Call validaBienFuturo 'ERS0012020********************
    
    Exit Sub
ErrNuevo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdValorizacionNuevo_Click()
    Dim lnTpoValoriza As eGarantiaTipoValorizacion
    Dim oFrm1 As frmGarantiaValorDirectoTotal
    Dim oFrm2 As frmGarantiaValorDirectoDetallado
    Dim oFrm3 As frmGarantiaValorAutoLiquidable
    Dim oFrm4 As frmGarantiaValorTasacInmob
    Dim oFrm5 As frmGarantiaValorTasacVehicular
    Dim oFrm6 As frmGarantiaValorTasacMobilOtras
    Dim oFrm8 As frmGarantiaValorGravamenOtraIFI
    
    Dim lvValDirTot As tValorDirectoTotal
    Dim lvValDirTot_ULT_VAL As tValorDirectoTotal
    Dim lvValDirDet() As tValorDirectoDetallado
    Dim lvValDirDet_ULT_VAL() As tValorDirectoDetallado
    Dim lvValorAutoLiq As tValorAutoLiquidable
    Dim lvValorTasacionInmobiliaria As tValorTasacionInmobiliaria
    Dim lvValorTasacionInmobiliaria_ULT_VAL As tValorTasacionInmobiliaria
    Dim lvValorTasacionVehicular As tValorTasacionVehicular
    Dim lvValorTasacionVehicular_ULT_VAL As tValorTasacionVehicular
    Dim lvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
    Dim lvValorTasacionMobiliariaOtras_ULT_VAL As tValorTasacionMobiliariaOtras
    Dim lvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi
    Dim lvValorGravamenFavorOtraIFi_ULT_VAL As tValorGravamenFavorOtraIFi
    
    Dim ldFecha As Date
    Dim lsUsuario As String
    Dim lnMoneda As Moneda
    Dim lsGlosa As String
    Dim lbExisteValorizacion As Boolean
    Dim lsPersCodTit As String
    
    Dim lsUltimaActualizacion As String
       
    On Error GoTo ErrValorizaNuevo
    
    If cmbValorizacion.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Tipo de Valorizacion", vbInformation, "Aviso"
        EnfocaControl cmbValorizacion
        Exit Sub
    End If
    
    If Not validaPropietarios(True) Then Exit Sub
    
    lnTpoValoriza = CInt(Trim(Right(cmbValorizacion.Text, 3)))
    lbExisteValorizacion = IIf(oGarantia.NroValorizacionesActivas > 0, True, False)
    lsPersCodTit = oGarantia.ObtenerPropietarioEspecifico(Titular).sPersCod
        
    ReDim lvValDirDet(0)
    
    If lbExisteValorizacion Then
        lnMoneda = oGarantia.Moneda
    End If
    'CTI5 ERS0012020******************************************************************
    If eGarantiaTipoValorizacion.TasacionInmobiliaria <> lnTpoValoriza And gTpoBienFuturo = val(Trim(Right(cmbTpoBien.Text, 3))) Then
         MsgBox "Para tipos de garantias BIEN FUTURO, no se puede crear valorizaciones diferentes a inmobiliarias", vbInformation, "Aviso"
         TabGarantia.Tab = 1
         Exit Sub
    End If
    '*********************************************************************************
    Select Case lnTpoValoriza
        Case eGarantiaTipoValorizacion.ValorDirectoTotal
            lvValDirTot_ULT_VAL = oGarantia.ObtenerValorizacionUltima(ValorDirectoTotal)
            
            Set oFrm1 = New frmGarantiaValorDirectoTotal
            If Not oFrm1.Registrar(Not lbExisteValorizacion, lnMoneda, lsGlosa, lvValDirTot, lvValDirTot_ULT_VAL) Then
                Set oFrm1 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.ValorDirectoDetallado
            lvValDirDet_ULT_VAL = oGarantia.ObtenerValorizacionUltima(ValorDirectoDetallado)
            
            Set oFrm2 = New frmGarantiaValorDirectoDetallado
            If Not oFrm2.Registrar(Not lbExisteValorizacion, lnMoneda, lsGlosa, lvValDirDet, lvValDirDet_ULT_VAL) Then
                Set oFrm2 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
            If lsPersCodTit = "" Then
                MsgBox "No se ha indicado al Titular de la Garantía", vbInformation, "Aviso"
                TabGarantia.Tab = 0
                EnfocaControl fePropietarios
                Exit Sub
            End If
            If lbExisteValorizacion Then
                lvValorAutoLiq = oGarantia.ObtenerValorizacion(1).vValorAutoLiquidable
            End If
            Set oFrm3 = New frmGarantiaValorAutoLiquidable
            If Not oFrm3.Registrar(Not lbExisteValorizacion, lsPersCodTit, lnMoneda, lvValorAutoLiq, oGarantia.GarantID) Then
                Set oFrm3 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.TasacionInmobiliaria
            lvValorTasacionInmobiliaria_ULT_VAL = oGarantia.ObtenerValorizacionUltima(TasacionInmobiliaria)
            Dim lsEtapa As String
            Dim oGarant As New COMNCredito.NCOMGarantia
            
            If oGarantia.TpoBienContrato = 2 Then
                If Trim(Right(cmbDocumentoTpo, 3)) = "" Then
                    MsgBox "Debe seleccionar el tipo de documento de la garantía", vbExclamation, "Aviso"
                    Exit Sub
                End If
                Dim oRsEtapa As ADODB.Recordset
                Set oRsEtapa = New ADODB.Recordset
                Set oRsEtapa = oGarant.ObtenerGarantiaBienFuturo(val(Trim(Right(cmbDocumentoTpo, 3))))
                If Not (oRsEtapa.EOF And oRsEtapa.BOF) Then
                       lsEtapa = oRsEtapa!cEtapa
                End If
                 RSClose oRsEtapa
            Else
                lsEtapa = ""
                oGarantia.TpoBienContrato = 1
            End If
            
            Set oFrm4 = New frmGarantiaValorTasacInmob
            If Not oFrm4.Registrar(Not lbExisteValorizacion, lnMoneda, lvValorTasacionInmobiliaria, lvValorTasacionInmobiliaria_ULT_VAL, oGarantia.TpoBienContrato, lsEtapa, val(Trim(Right(cmbDocumentoTpo, 3)))) Then
                Set oFrm4 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.TasacionVehicular
            lvValorTasacionVehicular_ULT_VAL = oGarantia.ObtenerValorizacionUltima(TasacionVehicular)
            
            Set oFrm5 = New frmGarantiaValorTasacVehicular
            If Not oFrm5.Registrar(Not lbExisteValorizacion, lnMoneda, lvValorTasacionVehicular, lvValorTasacionVehicular_ULT_VAL) Then
                Set oFrm5 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
            lvValorTasacionMobiliariaOtras_ULT_VAL = oGarantia.ObtenerValorizacionUltima(OtrasTasacionesMobiliarias)
            
            Set oFrm6 = New frmGarantiaValorTasacMobilOtras
            If Not oFrm6.Registrar(Not lbExisteValorizacion, lnMoneda, lvValorTasacionMobiliariaOtras, lvValorTasacionMobiliariaOtras_ULT_VAL) Then
                Set oFrm6 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160224
            lvValorGravamenFavorOtraIFi_ULT_VAL = oGarantia.ObtenerValorizacionUltima(GravamenFavorOtraIFi)
            
            Set oFrm8 = frmGarantiaValorGravamenOtraIFI
            If Not oFrm8.Registrar(Not lbExisteValorizacion, lnMoneda, lvValorGravamenFavorOtraIFi, lvValorGravamenFavorOtraIFi_ULT_VAL) Then
                Set oFrm8 = Nothing
                Exit Sub
            End If
        Case Else
            MsgBox "Para este tipo de valorización no se ha establecido configuración alguna," & Chr(13) & "consulte con el Dpto. de TI", vbExclamation, "Aviso"
            Exit Sub
    End Select
    
    lsUltimaActualizacion = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), UCase(gsCodUser))
    ldFecha = fgFechaHoraMovDate(lsUltimaActualizacion)
    
    oGarantia.InsertarValorizacion ldFecha, lsUltimaActualizacion, lnTpoValoriza, Trim(Mid(cmbValorizacion.Text, 1, Len(cmbValorizacion.Text) - 3)), lsGlosa, GarantFilaNueva, False, lvValDirTot, lvValDirDet, lvValorAutoLiq, lvValorTasacionInmobiliaria, lvValorTasacionVehicular, lvValorTasacionMobiliariaOtras, True, lvValorGravamenFavorOtraIFi
    oGarantia.PermiteNuevaValorizacion = False
    
    If Not lbExisteValorizacion Then
        oGarantia.Moneda = lnMoneda
    End If
    
    cmdValorizacionNuevo.Enabled = False
    SetFlexValorizacion
    
    HabilitaClasificacion
    HabilitaBienGarantia
    HabilitarPertenencia
    
    If lnTpoValoriza = GarantiaAutoliquidable Then
        txtDocumentoNro.Text = lvValorAutoLiq.sCtaCod
    End If
    
    Set oFrm1 = Nothing
    Set oFrm2 = Nothing
    Set oFrm3 = Nothing
    Set oFrm4 = Nothing
    Set oFrm5 = Nothing
    Set oFrm6 = Nothing
    Set oFrm8 = Nothing
    Call validaBienFuturo 'ERS0012020********************
    Exit Sub
ErrValorizaNuevo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdValorizacionDetalle_Click()
    Dim oFrm1 As frmGarantiaValorDirectoTotal
    Dim oFrm2 As frmGarantiaValorDirectoDetallado
    Dim oFrm3 As frmGarantiaValorAutoLiquidable
    Dim oFrm4 As frmGarantiaValorTasacInmob
    Dim oFrm5 As frmGarantiaValorTasacVehicular
    Dim oFrm6 As frmGarantiaValorTasacMobilOtras
    Dim oFrm8 As frmGarantiaValorGravamenOtraIFI
    
    Dim lvValDirTot As tValorDirectoTotal
    Dim lvValDirTot_ULT_VAL As tValorDirectoTotal
    Dim lvValDirDet() As tValorDirectoDetallado
    Dim lvValDirDet_ULT_VAL() As tValorDirectoDetallado
    Dim lvValorAutoLiq As tValorAutoLiquidable
    Dim lvValorTasacionInmobiliaria As tValorTasacionInmobiliaria
    Dim lvValorTasacionInmobiliaria_ULT_VAL As tValorTasacionInmobiliaria
    Dim lvValorTasacionVehicular As tValorTasacionVehicular
    Dim lvValorTasacionVehicular_ULT_VAL As tValorTasacionVehicular
    Dim lvValorTasacionMobiliariaOtras As tValorTasacionMobiliariaOtras
    Dim lvValorTasacionMobiliariaOtras_ULT_VAL As tValorTasacionMobiliariaOtras
    Dim lvValorGravamenFavorOtraIFi As tValorGravamenFavorOtraIFi
    Dim lvValorGravamenFavorOtraIFi_ULT_VAL As tValorGravamenFavorOtraIFi
    
    Dim Index As Integer
    Dim ldValorizacion As Date
    Dim lnTpoValoriza As eGarantiaTipoValorizacion
    Dim lbVinculado As Boolean
    Dim lnCuenta As Integer
    Dim lnMoneda As Moneda
    Dim lnValorRealizacion As Currency
    Dim lsGlosa As String
    Dim lsPersCodTit As String
    Dim lbPrimero As Boolean
    Dim lbEditar As Boolean
    Dim lbConsultar As Boolean
    
    Dim lsUltimaActualizacion As String

    On Error GoTo ErrValorizaDetalle
    
    If feValorizacion.TextMatrix(1, 0) = "" Then Exit Sub
    If Not validaPropietarios(True) Then Exit Sub
    
    Index = feValorizacion.TextMatrix(feValorizacion.row, 6) 'Indice de la Matriz que hace referencia
    ldValorizacion = oGarantia.ObtenerValorizacion(Index).dfecha
    lnTpoValoriza = oGarantia.ObtenerValorizacion(Index).nValorizacionTpo
    lbVinculado = oGarantia.ObtenerValorizacion(Index).bVinculado
    lnCuenta = oGarantia.NroValorizacionesActivas
    lbPrimero = IIf(lnCuenta = 1, True, False)
    lnMoneda = oGarantia.Moneda
    lvValDirTot = oGarantia.ObtenerValorizacion(Index).vValorDirectoTotal
    lsGlosa = oGarantia.ObtenerValorizacion(Index).sGlosa
    lsPersCodTit = oGarantia.ObtenerPropietarioEspecifico(Titular).sPersCod
    
    ReDim lvValDirDet(0)
    
    lbConsultar = (lbVinculado) Or (fnTipoInicio = ConsultarGarantia) Or (Not fbAccion) Or (IIf(Index <> oGarantia.IndexUltimaValorizacionActiva, True, False))
    Select Case lnTpoValoriza
        Case eGarantiaTipoValorizacion.ValorDirectoTotal
            Set oFrm1 = New frmGarantiaValorDirectoTotal
            If lbConsultar Then
                oFrm1.Consultar lnMoneda, lsGlosa, lvValDirTot
                Set oFrm1 = Nothing 'EJVG20160225
                Exit Sub
            End If
            
            lvValDirTot_ULT_VAL = oGarantia.ObtenerValorizacionUltima(ValorDirectoTotal)
            If Not oFrm1.Editar(lbPrimero, lnMoneda, lsGlosa, lvValDirTot, lvValDirTot_ULT_VAL) Then
                Set oFrm1 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.ValorDirectoDetallado
            Set oFrm2 = New frmGarantiaValorDirectoDetallado
            lvValDirDet = oGarantia.ObtenerValorizacion(Index).vValorDirDet
            
            If lbConsultar Then
                oFrm2.Consultar lnMoneda, lnValorRealizacion, lsGlosa, lvValDirDet
                Set oFrm2 = Nothing 'EJVG20160225
                Exit Sub
            End If
            
            lvValDirDet_ULT_VAL = oGarantia.ObtenerValorizacionUltima(ValorDirectoDetallado, ldValorizacion)
            If Not oFrm2.Editar(lbPrimero, lnMoneda, lsGlosa, lvValDirDet, lvValDirDet_ULT_VAL) Then
                Set oFrm2 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.GarantiaAutoliquidable
            Set oFrm3 = New frmGarantiaValorAutoLiquidable
            lvValorAutoLiq = oGarantia.ObtenerValorizacion(Index).vValorAutoLiquidable
            
            If lbConsultar Then
                oFrm3.Consultar lsPersCodTit, lvValorAutoLiq, oGarantia.GarantID
                Set oFrm3 = Nothing 'EJVG20160225
                Exit Sub
            End If
            If Not oFrm3.Editar(lbPrimero, lsPersCodTit, lnMoneda, lvValorAutoLiq, oGarantia.GarantID) Then
                Set oFrm3 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.TasacionInmobiliaria
            Set oFrm4 = New frmGarantiaValorTasacInmob
            lvValorTasacionInmobiliaria = oGarantia.ObtenerValorizacion(Index).vValorTasacionInmobiliaria
            Dim lsEtapa As String
            Dim lsTpoDoc As String
            Dim oGarant As New COMNCredito.NCOMGarantia
            lsTpoDoc = val(Trim(Right(cmbDocumentoTpo, 3)))
            If oGarantia.TpoBienContrato = 2 Then
                If Trim(Right(cmbDocumentoTpo, 3)) = "" Then
                    MsgBox "Debe seleccionar el tipo de documento de la garantía", vbExclamation, "Aviso"
                    Exit Sub
                End If
                Dim oRsEtapa As ADODB.Recordset
                Set oRsEtapa = New ADODB.Recordset
                Set oRsEtapa = oGarant.ObtenerGarantiaBienFuturo(val(Trim(Right(cmbDocumentoTpo, 3))))
                If Not (oRsEtapa.EOF And oRsEtapa.BOF) Then
                       lsEtapa = oRsEtapa!cEtapa
                Else
                       lsEtapa = lvValorTasacionInmobiliaria.Set
                End If
                RSClose oRsEtapa
            Else
                lsEtapa = ""
                oGarantia.TpoBienContrato = 1
            End If
            If lbConsultar Then
                oFrm4.Consultar lnMoneda, lvValorTasacionInmobiliaria, oGarantia.TpoBienContrato, lsEtapa, lsTpoDoc
                Set oFrm4 = Nothing 'EJVG20160225
                Exit Sub
            End If
            
            lvValorTasacionInmobiliaria_ULT_VAL = oGarantia.ObtenerValorizacionUltima(TasacionInmobiliaria, ldValorizacion)
            If Not oFrm4.Editar(lbPrimero, lnMoneda, lvValorTasacionInmobiliaria, lvValorTasacionInmobiliaria_ULT_VAL, oGarantia.TpoBienContrato, lsEtapa, lsTpoDoc) Then
                Set oFrm4 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.TasacionVehicular
            Set oFrm5 = New frmGarantiaValorTasacVehicular
            lvValorTasacionVehicular = oGarantia.ObtenerValorizacion(Index).vValorTasacionVehicular
            
            If lbConsultar Then
                oFrm5.Consultar lnMoneda, lvValorTasacionVehicular
                Set oFrm5 = Nothing 'EJVG20160225
                Exit Sub
            End If
            
            lvValorTasacionVehicular_ULT_VAL = oGarantia.ObtenerValorizacionUltima(TasacionVehicular, ldValorizacion)
            If Not oFrm5.Editar(lbPrimero, lnMoneda, lvValorTasacionVehicular, lvValorTasacionVehicular_ULT_VAL) Then
                Set oFrm5 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.OtrasTasacionesMobiliarias
            Set oFrm6 = New frmGarantiaValorTasacMobilOtras
            lvValorTasacionMobiliariaOtras = oGarantia.ObtenerValorizacion(Index).vValorTasacionMobiliariaOtras
            
            If lbConsultar Then
                oFrm6.Consultar lnMoneda, lvValorTasacionMobiliariaOtras
                Set oFrm6 = Nothing 'EJVG20160225
                Exit Sub
            End If
            
            lvValorTasacionMobiliariaOtras_ULT_VAL = oGarantia.ObtenerValorizacionUltima(OtrasTasacionesMobiliarias, ldValorizacion)
            If Not oFrm6.Editar(lbPrimero, lnMoneda, lvValorTasacionMobiliariaOtras, lvValorTasacionMobiliariaOtras_ULT_VAL) Then
                Set oFrm6 = Nothing 'EJVG20160225
                Exit Sub
            End If
        Case eGarantiaTipoValorizacion.GravamenFavorOtraIFi 'EJVG20160224
            Set oFrm8 = New frmGarantiaValorGravamenOtraIFI
            lvValorGravamenFavorOtraIFi = oGarantia.ObtenerValorizacion(Index).vValorGravamenFavorOtraIFi
            
            If lbConsultar Then
                oFrm8.Consultar lnMoneda, lvValorGravamenFavorOtraIFi
                Set oFrm8 = Nothing 'EJVG20160225
                Exit Sub
            End If
            
            lvValorGravamenFavorOtraIFi_ULT_VAL = oGarantia.ObtenerValorizacionUltima(GravamenFavorOtraIFi)
            If Not oFrm8.Editar(lbPrimero, lnMoneda, lvValorGravamenFavorOtraIFi, lvValorGravamenFavorOtraIFi_ULT_VAL) Then
                Set oFrm8 = Nothing
                Exit Sub
            End If
        Case Else
            MsgBox "Para este tipo de valorización no se ha establecido configuración alguna," & Chr(13) & "consulte con el Dpto. de TI", vbExclamation, "Aviso"
            Exit Sub
    End Select
    
    lsUltimaActualizacion = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), UCase(gsCodUser))
    
    oGarantia.ActualizarValorizacion Index, lsUltimaActualizacion, lsGlosa, lvValDirTot, lvValDirDet, lvValorAutoLiq, lvValorTasacionInmobiliaria, lvValorTasacionVehicular, lvValorTasacionMobiliariaOtras, lvValorGravamenFavorOtraIFi
    If lbPrimero Then
        oGarantia.Moneda = lnMoneda
    End If
    SetFlexValorizacion
    
    If lnTpoValoriza = GarantiaAutoliquidable Then
        txtDocumentoNro.Text = lvValorAutoLiq.sCtaCod
    End If
    
    Set oFrm1 = Nothing
    Set oFrm2 = Nothing
    Set oFrm3 = Nothing
    Set oFrm4 = Nothing
    Set oFrm5 = Nothing
    Set oFrm6 = Nothing
    Set oFrm8 = Nothing
    Exit Sub
ErrValorizaDetalle:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdValorizacionEliminar_Click()
    Dim Index As Integer
    
    On Error GoTo ErrEliminar
    If feValorizacion.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = CInt(feValorizacion.TextMatrix(feValorizacion.row, 6))
    
    If MsgBox("¿Desea eliminar la valorización con fecha " & feValorizacion.TextMatrix(feValorizacion.row, 1) & "?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    oGarantia.EliminarValorizacion Index
    oGarantia.PermiteNuevaValorizacion = True
    
    cmdValorizacionNuevo.Enabled = True 'Siempre se va a eliminar la ultima valorización
    SetFlexValorizacion
    
    HabilitaClasificacion
    HabilitaBienGarantia
    HabilitarPertenencia
    
    Screen.MousePointer = 0
    Call validaBienFuturo 'ERS0012020********************
    Exit Sub
ErrEliminar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub feHistorialTra_GotFocus()
    fbFocoGrilla = True
End Sub

Private Sub feHistorialTra_LostFocus()
    fbFocoGrilla = False
End Sub

Private Sub cmdVerCreditos_Click()
    If oGarantia.GarantID = "" Then
        MsgBox "Ud. debe seleccionar primero a la Garantía", vbInformation, "Aviso"
        Exit Sub
    End If
    frmGarantiaCred.Inicio oGarantia.GarantID
End Sub

Private Sub fePropietarios_GotFocus()
    fbFocoGrilla = True
End Sub

Private Sub fePropietarios_LostFocus()
    fbFocoGrilla = False
End Sub

Private Sub fePropietarios_RowColChange()
    Dim rs As ADODB.Recordset
    If fePropietarios.lbEditarFlex Then
        If fePropietarios.col = 3 Then
            Set rs = fRsPropietarioRelac.Clone
            fePropietarios.CargaCombo rs
        End If
        fePropietarios.row = fnPropietarioNoMoverFila
    End If
    RSClose rs
End Sub

Private Sub feTramiteLegal_DblClick()
    If feTramiteLegal.TextMatrix(1, 0) = "" Then Exit Sub
    cmdTramiteLegalDetalle_Click
End Sub

Private Sub feTramiteLegal_OnRowChange(pnRow As Long, pnCol As Long)
    Dim Index As Integer
    If feTramiteLegal.TextMatrix(1, 0) = "" Then
        'Botón Detalle
        cmdTramiteLegalDetalle.Enabled = False
        'Botón Eliminar
        cmdTramiteLegalEliminar.Enabled = False
        'Botón Inscripción
        cmdTramiteLegalInscripcion.Enabled = False
    Else
        Index = feTramiteLegal.TextMatrix(feTramiteLegal.row, 7)
        
        'Botón Detalle
        cmdTramiteLegalDetalle.Enabled = True
        'Botón Eliminar
        If fbAccion And (oGarantia.ObtenerTramiteLegal(Index).bEliminar And fnTipoInicio <> ConsultarGarantia) Then
            cmdTramiteLegalEliminar.Enabled = True
         Else
            'JGPA TI-ERS026-2018------------
            'SE agregó para verificar solo legal tenga esta opción
            If fbAccion And (fnTipoInicio <> ConsultarGarantia) Then
                Dim vle As New COMDCredito.DCOMGarantia
                Dim rsLe As ADODB.Recordset
    
                Set rsLe = vle.VerificaCargoLegal(gsCodUser)
                If Not (rsLe.BOF And rsLe.EOF) Then
                    cmdTramiteLegalEliminar.Enabled = True
                 Else
                    cmdTramiteLegalEliminar.Enabled = False
                End If
                Set vle = Nothing
                RSClose rsLe
             Else
                cmdTramiteLegalEliminar.Enabled = False
            End If
            'End JGPA TI-ERS026-2018------------
            
            'cmdTramiteLegalEliminar.Enabled = False 'Comentado x JGPA20180801
        End If
        'Botón Inscripción
        If fbAccion _
                And fnTipoInicio = EditarGarantia _
                And Not oGarantia.ObtenerTramiteLegal(Index).bMigrado _
                And oGarantia.ObtenerTramiteLegal(Index).nEstado = Pendiente _
                And oGarantia.ObtenerTramiteLegal(Index).bVinculado = True _
                And Index = oGarantia.IndexUltimoTramiteLegalActiva Then
            cmdTramiteLegalInscripcion.Enabled = True
        Else
            cmdTramiteLegalInscripcion.Enabled = False
        End If
    End If
End Sub

Private Sub feTramiteLegal_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    
    Editar = Split(feTramiteLegal.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub

Private Sub feTramiteLegal_RowColChange()
'    Dim Index As Integer
'    If feTramiteLegal.TextMatrix(1, 0) = "" Then
'        'Botón Detalle
'        cmdTramiteLegalDetalle.Enabled = False
'        'Botón Eliminar
'        cmdTramiteLegalEliminar.Enabled = False
'        'Botón Inscripción
'        cmdTramiteLegalInscripcion.Enabled = False
'    Else
'        Index = feTramiteLegal.TextMatrix(feTramiteLegal.row, 7)
'
'        'Botón Detalle
'        cmdTramiteLegalDetalle.Enabled = True
'        'Botón Eliminar
'        If fbAccion And (oGarantia.ObtenerTramiteLegal(Index).bEliminar And fnTipoInicio <> ConsultarGarantia) Then
'            cmdTramiteLegalEliminar.Enabled = True
'        Else
'            cmdTramiteLegalEliminar.Enabled = False
'        End If
'        'Botón Inscripción
'        If fbAccion _
'                And fnTipoInicio = EditarGarantia _
'                And Not oGarantia.ObtenerTramiteLegal(Index).bMigrado _
'                And oGarantia.ObtenerTramiteLegal(Index).nEstado = Pendiente _
'                And oGarantia.ObtenerTramiteLegal(Index).bVinculado = True _
'                And Index = oGarantia.IndexUltimoTramiteLegalActiva Then
'            cmdTramiteLegalInscripcion.Enabled = True
'        Else
'            cmdTramiteLegalInscripcion.Enabled = False
'        End If
'    End If
End Sub

Private Sub feValorizacion_DblClick()
    If feValorizacion.TextMatrix(1, 0) = "" Then Exit Sub
    cmdValorizacionDetalle_Click
End Sub

Private Sub feValorizacion_GotFocus()
    fbFocoGrilla = True
End Sub

Private Sub feValorizacion_LostFocus()
    fbFocoGrilla = False
End Sub

Private Sub feValorizacion_OnRowChange(pnRow As Long, pnCol As Long)
    Dim Index As Integer
    If feValorizacion.TextMatrix(1, 0) = "" Then
        'Boton Detalle
        cmdValorizacionDetalle.Enabled = False
        'Boton Eliminar
        cmdValorizacionEliminar.Enabled = False
    Else
        Index = feValorizacion.TextMatrix(feValorizacion.row, 6)
        
        'Boton Detalle
        cmdValorizacionDetalle.Enabled = True
        'Boton Eliminar
        If fbAccion And (oGarantia.ObtenerValorizacion(Index).bEliminar And fnTipoInicio <> ConsultarGarantia) Then
            cmdValorizacionEliminar.Enabled = True
        Else
            cmdValorizacionEliminar.Enabled = False
        End If
    End If
End Sub

Private Sub feValorizacion_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    
    Editar = Split(feValorizacion.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub

Private Sub feValorizacion_RowColChange()
'    Dim Index As Integer
'    If feValorizacion.TextMatrix(1, 0) = "" Then
'        'Boton Detalle
'        cmdValorizacionDetalle.Enabled = False
'        'Boton Eliminar
'        cmdValorizacionEliminar.Enabled = False
'    Else
'        Index = feValorizacion.TextMatrix(feValorizacion.row, 6)
'
'        'Boton Detalle
'        cmdValorizacionDetalle.Enabled = True
'        'Boton Eliminar
'        If fbAccion And (oGarantia.ObtenerValorizacion(Index).bEliminar And fnTipoInicio <> ConsultarGarantia) Then
'            cmdValorizacionEliminar.Enabled = True
'        Else
'            cmdValorizacionEliminar.Enabled = False
'        End If
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If fbFocoGrilla Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    
    fbGrabar = False
    
    CargaControles
    CargaVariables
    Call CambiaTamañoCombo(cmbDocumentoTpo, 300)
    cmdCancelar_Click
    'If fnTipoInicio = ConsultarGarantia Then
    If fsNumGarant <> "" Then
        txtNroGarantBusca.Text = fsNumGarant
        cmdBuscaGarantia_Click
        fsNumGarant = ""
        fraBusca.Enabled = False
        cmdcancelar.Enabled = False
        CmdSalir.Cancel = True
    End If
    'End If
    Screen.MousePointer = 0
    
    gsOpeCod = gCredRegistrarGarantiaCli 'Log Grabar Garantías
End Sub

Private Sub HabilitaControles(ByVal pbValor As Boolean)
    HabilitaControlesPertenencia pbValor
    HabilitaControlesValuacion pbValor
    HabilitaControlesBienGarantia pbValor
    
    If pbValor Then
        HabilitaClasificacion
        HabilitaBienGarantia
        HabilitarPertenencia
    End If
    
    If pbValor And (oGarantia.VinculadoUltVAL And Not oGarantia.PermiteNuevaValorizacion) Then
        MsgBox "En estos momentos no es posible realizar modificaciones en las Valorizaciones ni Trámites Legales." & Chr(13) & Chr(13) & _
                "La Garantía está coberturando créditos que aún no se han desembolsado." & Chr(13) & Chr(13) & _
                "(Solo para Créditos Solicitados y Sugeridos)" & Chr(13) & _
                "En caso requiera editar estos datos de la Garantía siga las sgtes. instrucciones:" & Chr(13) & Chr(13) & _
                "1. Ingrese al Registro de Cobertura, seleccione el crédito y elimine la cobertura actual." & Chr(13) & _
                "2. Regrese a la opción de Garantía, ahora si podrá editarlo." & Chr(13) & _
                "3. Vuelva a ingresar al Registro de Cobertura para vincular con la garantía actualizada.", vbInformation, "Aviso"
    End If
End Sub

Private Sub HabilitaControlesPertenencia(ByVal pbValor As Boolean)
    cmbClasificacion.Enabled = pbValor
    cmbBienGarantia.Enabled = pbValor
    txtEmisorCod.Enabled = pbValor
    cmbDocumentoTpo.Enabled = pbValor
    txtDocumentoNro.Enabled = pbValor
    txtConstataFecha.Enabled = pbValor
    fePropietarios.lbEditarFlex = False
    cmdPropietarioAceptar.Enabled = pbValor
    cmdPropietarioCancelar.Enabled = pbValor
    cmdPropietarioNuevo.Enabled = pbValor
    'cmdPropietarioEditar.Enabled = pbValor
    cmdPropietarioEliminar.Enabled = pbValor
    cmbTpoBien.Enabled = pbValor 'CTI5 ERS001-2020
End Sub

Private Sub LimpiarControlesPertenencia()
    cmbClasificacion.ListIndex = -1
    cmbBienGarantia.ListIndex = -1
    txtIdSupGarant.Caption = ""
    txtIdSupGarantAntDesemb.Caption = ""
    txtEmisorCod.Text = ""
    txtEmisorNombre.Text = ""
    cmbDocumentoTpo.ListIndex = -1
    txtDocumentoNro.Text = ""
    txtConstataFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    FormateaFlex fePropietarios
End Sub

Private Sub HabilitaControlesValuacion(ByVal pbValor As Boolean)
    cmbValorizacion.Enabled = pbValor
    cmdValorizacionNuevo.Enabled = pbValor
    cmdValorizacionDetalle.Enabled = pbValor
    
    If pbValor Then
        cmdValorizacionNuevo.Enabled = oGarantia.PermiteNuevaValorizacion()
    End If
End Sub

Private Sub LimpiarControlesValuacion()
    cmbValorizacion.ListIndex = -1
    FormateaFlex feValorizacion
    cmdValorizacionDetalle.Enabled = False
    cmdValorizacionEliminar.Enabled = False
End Sub

Private Sub HabilitaControlesBienGarantia(ByVal pbValor As Boolean)
    cmdTramiteLegalDetalle.Enabled = pbValor
    cmdTramiteLegalNuevo.Enabled = pbValor
    
    If pbValor Then
        cmdTramiteLegalNuevo.Enabled = oGarantia.PermiteNuevoTramiteLegal
    End If
End Sub

Public Sub LimpiarVariables()

End Sub

Private Sub LimpiarControlesBienGarantia()
    FormateaFlex feTramiteLegal
End Sub

Public Sub LimpiarControlesBusca()
    LstGarantia.ListItems.Clear
    'txtNroGarantBusca.Text = "0"
End Sub

Private Sub LimpiarControles()
    If fbNroGarantiaDigitada Then
        LimpiarControlesBusca
        fbNroGarantiaDigitada = False
    End If
    LimpiarControlesPertenencia
    LimpiarControlesValuacion
    LimpiarControlesBienGarantia
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If fnTipoInicio = RegistrarGarantia Then
        If MsgBox("¿Esta seguro de salir del Registro de Garantía?", vbYesNo + vbInformation, "Aviso") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub LstGarantia_Click()
    If LstGarantia.ListItems.count > 0 Then
        txtNroGarantBusca.Text = LstGarantia.SelectedItem
    Else
        txtNroGarantBusca.Text = "0"
    End If
End Sub

Private Sub LstGarantia_DblClick()
    If LstGarantia.ListItems.count > 0 Then
        If Len(txtNroGarantBusca.Text) = 8 Then
            cmdBuscaGarantia_Click
        End If
    End If
End Sub

Private Sub LstGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
    'txtNroGarantBusca.Text = LstGarantia.SelectedItem
End Sub

Private Sub LstGarantia_KeyPress(KeyAscii As Integer)
    'txtNroGarantBusca.Text = LstGarantia.SelectedItem
    If KeyAscii = 13 Then
        EnfocaControl cmdBuscaGarantia
    End If
End Sub

Private Sub LstGarantia_KeyUp(KeyCode As Integer, Shift As Integer)
    LstGarantia_Click
End Sub

Private Sub TabGarantia_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If cmdPropietarioAceptar.Visible And cmdPropietarioCancelar.Enabled Then
            MsgBox "Pulse Aceptar para Registrar al Propietario", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl cmdPropietarioAceptar
        End If
    End If
End Sub

Private Sub txtConstataFecha_Change()
    If Len(Trim(ValidaFecha(txtConstataFecha.Text))) = 0 Then
        oGarantia.FechaConstata = CDate(txtConstataFecha.Text)
    End If
End Sub

Private Sub txtConstataFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdPropietarioNuevo
    End If
End Sub

Private Sub txtConstataFecha_LostFocus()
    Dim bMostrarMensaje As Boolean
    If IsDate(txtConstataFecha.Text) Then
        If CDate(txtConstataFecha.Text) > gdFecSis Then
            MsgBox "La fecha de constatación no puede ser mayor a la fecha de sistema", vbInformation, "Aviso"
            EnfocaControl txtConstataFecha
            Exit Sub
        End If
    Else
        bMostrarMensaje = True
        If (val(Trim(Right(cmbTpoBien.Text, 3))) = gTpoBienFuturo) Then
            If oGarantia.GarantID <> "" Then
                Dim oRsGarantiaDesem As ADODB.Recordset
                Dim oGarant As New COMNCredito.NCOMGarantia
                Set oRsGarantiaDesem = New ADODB.Recordset
                Set oRsGarantiaDesem = oGarant.ObtenerGarantiaCreditoDesemboldado(oGarantia.GarantID)
                If (oRsGarantiaDesem.EOF And oRsGarantiaDesem.BOF) Then
                    bMostrarMensaje = False
                End If
                RSClose oRsGarantiaDesem
            Else
                 bMostrarMensaje = False
            End If

        End If
        If bMostrarMensaje Then
            MsgBox "Ud. debe ingresar una fecha válida", vbInformation, "Aviso"
            EnfocaControl txtConstataFecha
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDocumentoNro_Change()
    oGarantia.DocumentoNro = txtDocumentoNro.Text
End Sub

'Private Sub txtDocumentoNro_KeyDown(KeyCode As Integer, Shift As Integer)
'    Clipboard.Clear
'    Clipboard.SetText ""
'End Sub

Private Sub txtDocumentoNro_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtConstataFecha
    End If
End Sub

Private Sub txtDocumentoNro_LostFocus()
    txtDocumentoNro.Text = Trim(UCase(txtDocumentoNro.Text))
End Sub

Private Sub txtDocumentoNro_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Clipboard.Clear
    'Clipboard.SetText ""
End Sub

Private Sub txtEmisorCod_EmiteDatos()
    oGarantia.EmisorCod = ""
    txtEmisorNombre.Text = ""
    If Trim(txtEmisorCod.psCodigoPersona) <> "" Then
        oGarantia.EmisorCod = txtEmisorCod.psCodigoPersona
        oGarantia.EmisorNombre = txtEmisorCod.psDescripcion
        txtEmisorNombre.Text = oGarantia.EmisorNombre
        EnfocaControl cmbDocumentoTpo
    End If
End Sub

Private Function validaPropietarios(Optional ByVal pbGrabar As Boolean = False) As Boolean
    Dim oPersona As COMDPersona.DCOMPersona
    Dim rsPersona As ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim lnNroPropietarios As Integer
    Dim lnNroTitular As Integer, lnNroRepresentante As Integer
    Dim lsPersCodTitular As String
    
    If fePropietarios.TextMatrix(1, 0) = "" Then
        MsgBox "No se ha ingresado ningún propietario", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    
    For i = 1 To fePropietarios.rows - 1
        For j = 1 To fePropietarios.cols - 1
            If fePropietarios.ColWidth(j) > 0 Then
                If Len(Trim(fePropietarios.TextMatrix(i, j))) = 0 Then
                    MsgBox "El campo " & UCase(fePropietarios.TextMatrix(0, j)) & " está vacio, verifique..", vbInformation, "Aviso"
                    TabGarantia.Tab = 0
                    EnfocaControl fePropietarios
                    fePropietarios.TopRow = i
                    fePropietarios.row = i
                    fePropietarios.col = j
                    Exit Function
                End If
            End If
        Next
    Next
    
    For i = 1 To fePropietarios.rows - 1
        lnNroPropietarios = lnNroPropietarios + 1
        If val(Trim(Right(fePropietarios.TextMatrix(i, 3), 3))) = GarantiaPropietarioRelacionTitular Then
            lnNroTitular = lnNroTitular + 1
            lsPersCodTitular = fePropietarios.TextMatrix(i, 1)
        End If
        If val(Trim(Right(fePropietarios.TextMatrix(i, 3), 3))) = GarantiaPropietarioRelacionRepresentante Then
            lnNroRepresentante = lnNroRepresentante + 1
        End If
    Next
    
    If lnNroTitular = 0 Then
        MsgBox "Ud. primero debe ingresar al Titular de la Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    ElseIf lnNroTitular = 1 Then
        If fePropietarios.rows - 1 > 1 Or pbGrabar Then 'Si hay más de un registro
            If Not ValidaPropietariosDatos(lsPersCodTitular) Then Exit Function
        End If
    ElseIf lnNroTitular > 1 Then
        MsgBox "Solo puede existir un Titular de la Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    
    validaPropietarios = True
End Function

Private Function ValidaPropietariosDatos(Optional ByVal psPersCodTitular As String = "") As Boolean
    Dim oPersona As COMDPersona.DCOMPersona
    Dim rsPersona As ADODB.Recordset
    Dim lnPersoneria As PersPersoneria, lnPersoneriaConyugue As PersPersoneria
    Dim lnNroTitular As Integer, lnNroRepresentante As Integer, lnNroConyugue As Integer
    Dim i As Integer
    Dim lsPersCodConyugue As String
    Dim lsSexoTit As String, lsSexoConyugue As String
    
    On Error GoTo errValidaPropietariosRepJur
    If psPersCodTitular = "" Then
        For i = 1 To fePropietarios.rows - 1
            If val(Trim(Right(fePropietarios.TextMatrix(i, 3), 3))) = GarantiaPropietarioRelacionTitular Then
                lnNroTitular = lnNroTitular + 1
                psPersCodTitular = fePropietarios.TextMatrix(i, 1)
            End If
        Next
    Else
        lnNroTitular = 1
    End If
    
    If psPersCodTitular = "" Then
        MsgBox "Ud. primero debe ingresar al Titular de la Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    
    If lnNroTitular > 1 Then
        MsgBox "Solo puede existir un Titular de la Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    
    Set oPersona = New COMDPersona.DCOMPersona
    Set rsPersona = New ADODB.Recordset
    Set rsPersona = oPersona.RecuperaDatosPersonaxGarantia(psPersCodTitular)
    If Not rsPersona.EOF Then
        lnPersoneria = rsPersona!nPersPersoneria
        lsSexoTit = rsPersona!cPersnatSexo
    End If
    RSClose rsPersona
    
    If lnPersoneria = 0 Then
        MsgBox "La personería del Titular no se ha podido determinar, verifique en el Módulo de Personas", vbExclamation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    
    For i = 1 To fePropietarios.rows - 1
        If val(Trim(Right(fePropietarios.TextMatrix(i, 3), 3))) = GarantiaPropietarioRelacionRepresentante Then
            lnNroRepresentante = lnNroRepresentante + 1
        End If
        If val(Trim(Right(fePropietarios.TextMatrix(i, 3), 3))) = GarantiaPropietarioRelacionConyugue Then
            lsPersCodConyugue = fePropietarios.TextMatrix(i, 1)
            lnNroConyugue = lnNroConyugue + 1
        End If
    Next
    
    If lnPersoneria <> gPersonaNat And lnNroRepresentante = 0 Then 'Si el Titular es Juridico debe tener al menos un representante
        MsgBox "Ud. primero debe ingresar al Representante de la Persona Juridica", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    If lnPersoneria <> gPersonaNat And lnNroConyugue > 0 Then 'Si el Titular es Juridico no debe tener conyugue
        MsgBox "El Titular es una Persona Juridica, no debe tener Conyugue", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    If lnPersoneria = gPersonaNat And lnNroConyugue > 1 Then 'Si el Titular es Natural solo puede tener un conyugue
        MsgBox "El Titular no puede tener más de un Conyugue", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    ElseIf lnPersoneria = gPersonaNat And lnNroConyugue = 1 Then 'Si el Titular es Natural debe tener un conyugue de sexo opuesto
        Set rsPersona = New ADODB.Recordset
        Set rsPersona = oPersona.RecuperaDatosPersonaxGarantia(lsPersCodConyugue)
        If Not rsPersona.EOF Then
            lsSexoConyugue = rsPersona!cPersnatSexo
        End If
        RSClose rsPersona
        If lsSexoConyugue = lsSexoTit Then
            MsgBox "El Titular y el Conyugue no pueden tener el mismo sexo", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl fePropietarios
            Exit Function
        End If
    End If
    If lnPersoneria = gPersonaNat And lnNroRepresentante > 0 Then 'Si el Titular es Natural solo puede tener un conyugue
        MsgBox "El Titular es una Persona Natural, no puede tener Representante", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl fePropietarios
        Exit Function
    End If
    
    ValidaPropietariosDatos = True
    Exit Function
errValidaPropietariosRepJur:
    ValidaPropietariosDatos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub SetFlexPropietario()
    Dim Index As Integer, IndexFlex As Integer
    
    FormateaFlex fePropietarios
    For Index = 1 To oGarantia.NroPropietarios
        If oGarantia.ObtenerPropietario(Index).nGarantiaCambiosFila <> GarantFilaEliminada And oGarantia.ObtenerPropietario(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            fePropietarios.AdicionaFila
            IndexFlex = fePropietarios.row
            fePropietarios.TextMatrix(IndexFlex, 1) = oGarantia.ObtenerPropietario(Index).sPersCod 'Cod Persona
            fePropietarios.TextMatrix(IndexFlex, 2) = oGarantia.ObtenerPropietario(Index).sPersNombre 'Nombre Persona
            fePropietarios.TextMatrix(IndexFlex, 3) = oGarantia.ObtenerPropietario(Index).sRelacionTpo 'Relacion
        End If
    Next
End Sub

Private Sub SetFlexValorizacion()
    Dim Index As Integer, IndexFlex As Integer
    
    FormateaFlex feValorizacion
    For Index = 1 To oGarantia.NroValorizaciones
        If oGarantia.ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaEliminada And oGarantia.ObtenerValorizacion(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            feValorizacion.AdicionaFila
            IndexFlex = feValorizacion.row
            feValorizacion.TextMatrix(IndexFlex, 1) = Format(fgFechaHoraMovDate(oGarantia.ObtenerValorizacion(Index).sUltimaActualizacion), "dd/mm/yyyy hh:mm:ss AMPM") 'Fecha
            feValorizacion.TextMatrix(IndexFlex, 2) = Right(oGarantia.ObtenerValorizacion(Index).sUltimaActualizacion, 4) 'Usuario
            feValorizacion.TextMatrix(IndexFlex, 3) = oGarantia.ObtenerValorizacion(Index).sValorizacionTpo 'Tipo Valorizacion
            feValorizacion.TextMatrix(IndexFlex, 4) = IIf(oGarantia.Moneda = gMonedaNacional, "SOLES", "DOLARES") 'Moneda
            feValorizacion.TextMatrix(IndexFlex, 5) = Format(oGarantia.ObtenerValorizacion(Index).nVRM, "#,##0.00") 'Monto
            feValorizacion.TextMatrix(IndexFlex, 6) = Index 'Indice de la Matriz que hace referencia
        End If
    Next
    'feValorizacion_RowColChange
    feValorizacion_OnRowChange feValorizacion.row, feValorizacion.col
End Sub

Private Sub SetFlexTramiteLegal()
    Dim Index As Integer, IndexFlex As Integer
    
    FormateaFlex feTramiteLegal
    For Index = 1 To oGarantia.NroTramitesLegal
        If oGarantia.ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaEliminada And oGarantia.ObtenerTramiteLegal(Index).nGarantiaCambiosFila <> GarantFilaOculta Then
            feTramiteLegal.AdicionaFila
            IndexFlex = feTramiteLegal.row
            feTramiteLegal.TextMatrix(IndexFlex, 1) = Format(fgFechaHoraMovDate(oGarantia.ObtenerTramiteLegal(Index).sUltimaActualizacion), "dd/mm/yyyy hh:mm:ss AMPM")
            feTramiteLegal.TextMatrix(IndexFlex, 2) = Right(oGarantia.ObtenerTramiteLegal(Index).sUltimaActualizacion, 4) 'Usuario
            feTramiteLegal.TextMatrix(IndexFlex, 3) = oGarantia.ObtenerTramiteLegal(Index).cTipoTramite 'Tipo Tramite
            feTramiteLegal.TextMatrix(IndexFlex, 4) = oGarantia.ObtenerTramiteLegal(Index).cOficinaRegistralNombre 'Tipo Tramite
            feTramiteLegal.TextMatrix(IndexFlex, 5) = oGarantia.ObtenerTramiteLegal(Index).cEstado 'Estado
            feTramiteLegal.TextMatrix(IndexFlex, 6) = Format(oGarantia.ObtenerTramiteLegal(Index).nGravamen, "#,##0.00") 'Estado
            feTramiteLegal.TextMatrix(IndexFlex, 7) = Index 'Indice de la Matriz que hace referencia
        End If
    Next
    'feTramiteLegal_RowColChange
    feTramiteLegal_OnRowChange feTramiteLegal.row, feTramiteLegal.col
End Sub

Private Sub txtEmisorCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmbDocumentoTpo
    End If
End Sub

Private Function ValidarGarantia() As Boolean
    Dim oDPer As DCOMPersona
    Dim lsGarantIDExiste As String
    Dim lsFecha As String
    Dim iVAL As Integer
    
    Dim lsCadenaCoberturaGarantia As String
    Dim lnVRM As Currency
    Dim lbTramiteLegal As Boolean
    Dim lnTipoTramiteLegal As Integer
    Dim lnGravamen As Currency
    Dim lnTpoDoc As Integer
    
    Dim lnClasificacion As eGarantiaClasificacionBien
    Dim bEsDOI As Boolean
    Dim bEmisorIgualTitular As Boolean
    
    If cmbClasificacion.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Clasificación del Bien", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl cmbClasificacion
        Exit Function
    Else
        lnClasificacion = CInt(Right(cmbClasificacion.Text, 3))
    End If
    If cmbBienGarantia.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Bien en Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl cmbBienGarantia
        Exit Function
    End If
    'CTI5 ERS001-2020**********************
    If oGarantia.Clasificacion = 1 And oGarantia.BienGarantia = 6 Then
        If cmbTpoBien.ListIndex = -1 Then
            MsgBox "Ud. debe seleccionar el Tip.contr.Bien de Garantía", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl cmbTpoBien
            Exit Function
        End If
    End If
    '**************************************
    If Len(Trim(txtEmisorCod.Text)) <> 13 Or Len(txtEmisorCod.psCodigoPersona) <> 13 Then
        MsgBox "Ud. debe seleccionar al Emisor del Bien", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl txtEmisorCod
        Exit Function
    End If
    If cmbDocumentoTpo.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Tipo de Documento", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl cmbDocumentoTpo
        Exit Function
    Else
        lnTpoDoc = CInt(Right(cmbDocumentoTpo, 5))
    End If
    If Len(Trim(txtDocumentoNro.Text)) = 0 Then
        MsgBox "Ud. debe especificar el Nro. de Documento", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl txtDocumentoNro
        Exit Function
    End If
    If Len(Trim(txtDocumentoNro.Text)) = 0 Then
        MsgBox "Ud. debe especificar el Nro. de Documento", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl txtDocumentoNro
        Exit Function
    End If
    
    If oGarantia.TpoBienContrato <> gTpoBienFuturo Or fbCredDesembolsado = True Then
        
        lsFecha = ValidaFecha(txtConstataFecha.Text)
        If Len(lsFecha) > 0 Then
            MsgBox lsFecha, vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl txtConstataFecha
            Exit Function
        Else
            If oGarantia.TpoBienContrato = gTpoBienConstruido Then
                If CDate(txtConstataFecha.Text) > gdFecSis Then
                    MsgBox "La fecha de Constatación no puede ser mayor a la fecha del Sistema", vbInformation, "Aviso"
                    TabGarantia.Tab = 0
                    EnfocaControl txtConstataFecha
                    Exit Function
                End If
            Else
                If CDate(txtConstataFecha.Text) < CDate(fdVigCredDesembolsado) Then
                    MsgBox "La fecha de Constatación no puede ser menor a la fecha del desembolso", vbInformation, "Aviso"
                    TabGarantia.Tab = 0
                    EnfocaControl txtConstataFecha
                    Exit Function
                 End If
            End If
        End If
    End If
    
    If Not validaPropietarios(True) Then Exit Function
    
    If cmdPropietarioAceptar.Visible And cmdPropietarioCancelar.Enabled Then
        MsgBox "Pulse Aceptar para Registrar al Propietario", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl cmdPropietarioAceptar
        Exit Function
    End If
    
    If feValorizacion.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de ingresar la Valorización de la Garantía", vbInformation, "Aviso"
        TabGarantia.Tab = 1
        EnfocaControl feValorizacion
        Exit Function
    Else
        iVAL = oGarantia.IndexUltimaValorizacionActiva
    End If

    If oGarantia.ObtenerValorizacion(iVAL).nValorizacionTpo = GarantiaAutoliquidable Then
        If txtEmisorCod.psCodigoPersona <> gsCodPersCMACT Then
            MsgBox "El Emisor del documento en una Valorización AutoLiquidable debe ser la entidad " & UCase(gsNomCmac), vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl txtEmisorCod
            Exit Function
        End If
        If lnTpoDoc <> 17 Then
            MsgBox "El único documento permitido para una Valorización AutoLiquidable es el Deposito a Plazo Fijo" & Chr(13) & "la que debe contener como Nro. de Documento el Nro. de Cuenta", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl cmbDocumentoTpo
            Exit Function
        End If
        If oGarantia.ObtenerValorizacion(iVAL).vValorAutoLiquidable.sCtaCod <> Trim(txtDocumentoNro.Text) Then
            MsgBox "El Nro. de Documento debe ser igual a la cuenta de la última Valorización, verifique!", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl txtDocumentoNro
            Exit Function
        End If
    ElseIf oGarantia.ObtenerValorizacion(iVAL).nValorizacionTpo = ValorDirectoDetallado Then  'Si es Detallado->Debe tener como documento declaración jurada
        If txtEmisorCod.psCodigoPersona <> oGarantia.ObtenerPropietarioEspecifico(Titular).sPersCod Then
            MsgBox "El Emisor del documento debe ser el mismo Titular de la Garantía", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl txtEmisorCod
            Exit Function
        End If
        If lnTpoDoc <> 93 Then
            MsgBox "El único documento permitido para una Valorización Directa Detallada es la Declaración Jurada," & Chr(13) & "la que debe contener como Nro. de Documento el DOI del titular de la Garantía", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl cmbDocumentoTpo
            Exit Function
        End If
        
        Set oDPer = New DCOMPersona
        bEsDOI = oDPer.EsDOIdeTitular(txtEmisorCod.psCodigoPersona, Trim(txtDocumentoNro.Text))
        Set oDPer = Nothing
        If Not bEsDOI Then
            MsgBox "Ud. debe ingresar como Nro. de Documento el DOI del titular de la Garantía", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl txtDocumentoNro
            Exit Function
        End If
        bEmisorIgualTitular = True
    Else 'If oGarantia.ObtenerValorizacion(iVAL).nValorizacionTpo = ValorDirectoTotal Then  'Si es Detallado->Debe tener como documento declaración jurada
        
        bEmisorIgualTitular = False
    End If
    
    If Not bEmisorIgualTitular Then
        If txtEmisorCod.psCodigoPersona = oGarantia.ObtenerPropietarioEspecifico(Titular).sPersCod Then
            MsgBox "El Emisor del documento no puede ser el mismo Titular de la Garantía", vbInformation, "Aviso"
            TabGarantia.Tab = 0
            EnfocaControl txtEmisorCod
            Exit Function
        End If
    End If
    
    If oGarantia.ExisteGarantia(lsGarantIDExiste) Then
        MsgBox "El documento de propiedad del bien ya existe para la Garantía N° " & lsGarantIDExiste & "." & Chr(13) & Chr(13) & "Por favor modifiquelo e inténtelo de nuevo.", vbInformation, "Aviso"
        TabGarantia.Tab = 0
        EnfocaControl txtDocumentoNro
        Exit Function
    End If
    
    If oGarantia.TpoBienContrato = gTpoBienFuturo Then
    Dim objValorizacion As tValorizacion
    objValorizacion = oGarantia.ObtenerValorizacion(oGarantia.IndexUltimaValorizacionActiva)
    
    If objValorizacion.nValorizacionTpo = TasacionInmobiliaria Then
        If Trim(objValorizacion.vValorTasacionInmobiliaria.sMz) = "" Then
            MsgBox "Debe ingresar la Manzana de la tasación inmobiliaria." & Chr(13) & Chr(13) & "Por favor modifiquelo e inténtelo de nuevo.", vbInformation, "Aviso"
            TabGarantia.Tab = 1
            Exit Function
        End If
        If Trim(objValorizacion.vValorTasacionInmobiliaria.sLt) = "" Then
            MsgBox "Debe ingresar el Lote de la tasación inmobiliaria." & Chr(13) & Chr(13) & "Por favor modifiquelo e inténtelo de nuevo.", vbInformation, "Aviso"
            TabGarantia.Tab = 1
            Exit Function
        End If
        If Trim(objValorizacion.vValorTasacionInmobiliaria.Set) = "" Then
            MsgBox "Debe ingresar la Etapa de la tasación inmobiliaria." & Chr(13) & Chr(13) & "Por favor modifiquelo e inténtelo de nuevo.", vbInformation, "Aviso"
            TabGarantia.Tab = 1
            Exit Function
        End If
    Else
        MsgBox "Para Garantías BIEN FUTURO se requiere contar con Tasación Inmobilitaria." & Chr(13) & Chr(13) & "Por favor modifiquelo e inténtelo de nuevo.", vbInformation, "Aviso"
        TabGarantia.Tab = 1
        Exit Function
    End If
    End If
    If oGarantia.Clasificacion = 1 And oGarantia.BienGarantia = 6 Then
    Dim objTramiteLegal As tTramiteLegal
    If oGarantia.TpoBienContrato = gTpoBienFuturo Then
       
        objTramiteLegal = oGarantia.ObtenerTramiteLegal(oGarantia.IndexUltimoTramiteLegalActiva)
        
        If Trim(objTramiteLegal.sNroPartidaRegistral) <> "" Then
                MsgBox "Para garantías BIEN FUTURO no se permite registrar TRAMITE LEGAL.", vbInformation, "Aviso"
                TabGarantia.Tab = 2
               
                Exit Function
        End If
    End If
    If oGarantia.TpoBienContrato = gTpoBienConstruido Then
        
        objTramiteLegal = oGarantia.ObtenerTramiteLegal(oGarantia.IndexUltimoTramiteLegalActiva)
        
        'If Trim(objTramiteLegal.sNroPartidaRegistral) = "" Then
                'MsgBox "Para garantías BIEN CONSTRUIDO es obligatorio registrar TRAMITE LEGAL.", vbInformation, "Aviso"
                'TabGarantia.Tab = 2
                'Exit Function
		'JOEP20210820 Mejora para garantias antes del pase Tip.Contr.Bien
        If oGarantia.SalaValida <> 1 Then
            If Trim(objTramiteLegal.sNroPartidaRegistral) = "" Then
                    MsgBox "Para garantías BIEN CONSTRUIDO es obligatorio registrar TRAMITE LEGAL.", vbInformation, "Aviso"
                    TabGarantia.Tab = 2
                    Exit Function
            End If
        End If
        'JOEP20210820 Mejora para garantias antes del pase Tip.Contr.Bien					  						   
        'End If
    End If
    End If
    'Lógica para validar monto de Garantía
    fbDescobertura = False
    fbEliminarRegCobCredProcesoXDescob = False
    If fnTipoInicio = EditarGarantia Then
        lnVRM = oGarantia.ObtenerValorizacion(oGarantia.IndexUltimaValorizacionActiva).nVRM
        If oGarantia.IndexUltimoTramiteLegalActiva > 0 Then
            lbTramiteLegal = True
            lnTipoTramiteLegal = oGarantia.ObtenerTramiteLegal(oGarantia.IndexUltimoTramiteLegalActiva).nTipoTramite
            lnGravamen = oGarantia.ObtenerTramiteLegal(oGarantia.IndexUltimoTramiteLegalActiva).nGravamen
        End If
        
        lsCadenaCoberturaGarantia = oGarantia.CadenaCoberturaGarantia(lbTramiteLegal, lnTipoTramiteLegal, lnVRM, lnGravamen)
        If Len(lsCadenaCoberturaGarantia) > 0 Then
            'EJVG20160420 ***
            fbDescobertura = True
            If oGarantia.ObviarDescobertura() Then
                If oGarantia.NroCredEnProceso > 0 Then 'Si la garantía está con créditos en proceso(Solicitado(s),Sugerido(s) y/o Aprobado(s))
                    If MsgBox(lsCadenaCoberturaGarantia & Chr(13) & Chr(13) & "La garantía tiene " & oGarantia.NroCredEnProceso & " crédito" & IIf(oGarantia.NroCredEnProceso > 1, "s", "") _
                        & " en proceso (Solicitado, Sugerido y/o Aprobado) para posterior desembolso." & Chr(13) & Chr(13) & "De continuar se va a eliminar los Registro de Coberturas con est" & IIf(oGarantia.NroCredEnProceso > 1, "os", "e") & " crédito" & IIf(oGarantia.NroCredEnProceso > 1, "s", "") & "." & Chr(13) & Chr(13) _
                        & "¿Desea continuar de todos modos con la actualización?" _
                        , vbInformation + vbYesNo, "Se va a Eliminar el Registro de Cobertura") = vbNo Then
                        Exit Function
                    Else
                        fbEliminarRegCobCredProcesoXDescob = True
                    End If
                End If
            Else
                MsgBox lsCadenaCoberturaGarantia, vbInformation, "No se podrá continuar"
                Exit Function
            End If
            'END EJVG *******
        End If
    End If
    
    VerificarFechaSistema Me, True 'validar fecha de sistema, en caso no apaguen sus PC los usuarios les sacará del sistema
    '>>>>  I M P O R T A N T E  <<<<  No agregar más validaciones de acá para abajo, ya que estás debería ser las últimas validaciones
    
    ValidarGarantia = True
End Function

Private Function cargarDatos(ByVal psGarantID As String) As Boolean
    On Error GoTo ErrCargarDatos
    Dim nDocumentoNroTemp As String 'CTI5 ERS0012020
    Screen.MousePointer = 11

    If Not oGarantia.cargarDatos(psGarantID) Then
        Screen.MousePointer = 0
        cmdCancelar_Click
        MsgBox "No se ha podido cargar los datos de la Garantía seleccionada", vbInformation, "Aviso"
        Exit Function
    End If
    fbTieneCredito = False
    cmbClasificacion.ListIndex = IndiceListaCombo(cmbClasificacion, oGarantia.Clasificacion)
    cmbBienGarantia.ListIndex = IndiceListaCombo(cmbBienGarantia, oGarantia.BienGarantia)
    txtIdSupGarant.Caption = oGarantia.ClasificacionSBS
    txtIdSupGarantAntDesemb.Caption = oGarantia.ClasificacionInterna
    txtEmisorCod.Text = oGarantia.EmisorCod
    txtEmisorCod.psCodigoPersona = oGarantia.EmisorCod
    txtEmisorNombre.Text = oGarantia.EmisorNombre
    cmbTpoBien.ListIndex = IndiceListaCombo(cmbTpoBien, oGarantia.TpoBienContrato) 'CTI5ERS001-2020
    
    'CTI5 ERS0012020****************************
    If oGarantia.DocumentoTpo = 161 Then
        oGarantia.DocumentoNro = ""
    End If
    txtDocumentoNro.MaxLength = 25
    nDocumentoNroTemp = oGarantia.DocumentoNro
    Call OcultarVerBienFuturo
    If fbVerBienFuturo = True Then
        If val(Trim(Right(cmbTpoBien.Text, 3))) = 1 Then
            Call CargaTpoDocumento
        ElseIf val(Trim(Right(cmbTpoBien.Text, 3))) = 2 Then
           Call CargaTpoDocumentoBienFuturo
           txtDocumentoNro.MaxLength = 200
        Else
            cmbDocumentoTpo.Clear
        End If
    Else
        Call CargaTpoDocumento
    End If
    txtDocumentoNro.Text = ""

    '*******************************************
    'CTI5 ERS01-2020**********************************************************************
    cmbDocumentoTpo.ListIndex = IndiceListaCombo(cmbDocumentoTpo, oGarantia.DocumentoTpo)
    txtDocumentoNro.Text = nDocumentoNroTemp

    txtConstataFecha.Text = Format(oGarantia.FechaConstata, gsFormatoFechaView)
    If oGarantia.FechaConstata = "01/01/1900" Then
        txtConstataFecha.Text = "__/__/____"
    End If
    '*************************************************************************************
    SetFlexPropietario
    SetFlexValorizacion
    SetFlexTramiteLegal
    
    cmdVerCreditos.Enabled = True
    'CTI5 ERS01-2020**********************************************************************
    cmbTpoBien.Enabled = False
    txtDocumentoNro.Enabled = False
    cmbDocumentoTpo.Enabled = False
    '*************************************************************************************
    cargarDatos = True
    Screen.MousePointer = 0
    Exit Function
ErrCargarDatos:
    Screen.MousePointer = 0
    cargarDatos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub txtNroGarantBusca_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdBuscaGarantia
    End If
End Sub

Private Sub txtNroGarantBusca_LostFocus()
    txtNroGarantBusca.Text = Format(Trim(txtNroGarantBusca.Text), "00000000")
End Sub

Private Sub HabilitaClasificacion()
    Dim bHabilita As Boolean
    'Bloqueamos control Clasificación si es que tiene valorizaciones(Importante para generar asientos contables) ***
    bHabilita = IIf(feValorizacion.TextMatrix(1, 0) <> "", False, True)
    '***************************************************************************************************************
    cmbClasificacion.Enabled = bHabilita
End Sub

Private Sub HabilitaBienGarantia()
    Dim bHabilita As Boolean
    'Bloqueamos control Bien en Garantía si es se puede agregar valorizaciones ***
    bHabilita = cmdValorizacionNuevo.Enabled
    '*****************************************************************************
    cmbBienGarantia.Enabled = bHabilita
End Sub
Private Sub HabilitarPertenencia()
    Dim bHabilita As Boolean
    'Si no tiene créditos vigentes y si puede agregar valorizaciones(Sin créditos pendientes de desembolso) ***
    bHabilita = IIf(oGarantia.NroCredVinc <= 0 And cmbBienGarantia.Enabled = True, True, False)
    '**********************************************************************************************************
    
    'txtEmisorCod.Enabled = bHabilita
    'cmbDocumentoTpo.Enabled = bHabilita
    'txtDocumentoNro.Enabled = bHabilita
    
    'If fnTipoInicio = RegistrarGarantia And bHabilita = False Then
    '    txtEmisorCod.Enabled = True
    '    cmbDocumentoTpo.Enabled = True
    '    txtDocumentoNro.Enabled = True
    'End If
    
    cmdPropietarioNuevo.Enabled = bHabilita
    cmdPropietarioEliminar.Enabled = bHabilita
    
    'If Not bHabilita Then
    '    If cmbDocumentoTpo.ListIndex = -1 Or Len(Trim(txtDocumentoNro.Text)) = 0 Then
    '        cmbDocumentoTpo.Enabled = True
    '        txtDocumentoNro.Enabled = True
    '    End If
    '    If Len(Trim(txtEmisorCod.psCodigoPersona)) = 0 Then
    '        txtEmisorCod.Enabled = True
    '    End If
    '    If oGarantia.ObtenerPropietarioEspecifico(Titular).sPersCod = txtEmisorCod.psCodigoPersona Then
    '        txtEmisorCod.Enabled = True
    '    End If
    'End If
End Sub
