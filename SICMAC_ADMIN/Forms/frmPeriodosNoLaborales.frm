VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPeriodosNoLaborales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SicmactAdmin.ctrRRHHGen ctrRRHHGen1 
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7650
      TabIndex        =   34
      Top             =   5565
      Width           =   1095
   End
   Begin TabDlg.SSTab Tab 
      Height          =   4275
      Left            =   30
      TabIndex        =   1
      Top             =   1275
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   7541
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
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
      TabCaption(0)   =   "Vacaciones"
      TabPicture(0)   =   "frmPeriodosNoLaborales.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdEliminar(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdImprimir(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraPerNoLab(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGrabar(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNuevo(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdEditar(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Permisos"
      TabPicture(1)   =   "frmPeriodosNoLaborales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdNuevo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdEditar(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraPerNoLab(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdImprimir(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdEliminar(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdGrabar(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdCancelar(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Descansos"
      TabPicture(2)   =   "frmPeriodosNoLaborales.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPerNoLab(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdImprimir(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdEliminar(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdEditar(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdNuevo(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdGrabar(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdCancelar(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Sanciones"
      TabPicture(3)   =   "frmPeriodosNoLaborales.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraPerNoLab(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdImprimir(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdEliminar(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdEditar(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdNuevo(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdGrabar(3)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdCancelar(3)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   -69960
         TabIndex        =   35
         Top             =   3900
         Width           =   3525
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   -73710
         TabIndex        =   33
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   3
         Left            =   -74910
         TabIndex        =   32
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Index           =   3
         Left            =   -74910
         TabIndex        =   31
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   3
         Left            =   -73710
         TabIndex        =   30
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   3
         Left            =   -72525
         TabIndex        =   29
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   3
         Left            =   -71310
         TabIndex        =   28
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Caption         =   "Sanciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   3405
         Index           =   3
         Left            =   -74925
         TabIndex        =   26
         Top             =   345
         Width           =   8565
         Begin SicmactAdmin.FlexEdit FlexPerNoLab 
            Height          =   3075
            Index           =   3
            Left            =   105
            TabIndex        =   27
            Top             =   240
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   5424
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
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
            BackColorControl=   -2147483643
            BackColorControl=   -2147483628
            BackColorControl=   -2147483643
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   -1
            RowHeight0      =   240
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   1290
         TabIndex        =   25
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   2
         Left            =   1290
         TabIndex        =   22
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   2
         Left            =   2475
         TabIndex        =   21
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   2
         Left            =   3690
         TabIndex        =   20
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Caption         =   "Descansos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3405
         Index           =   2
         Left            =   75
         TabIndex        =   18
         Top             =   360
         Width           =   8565
         Begin SicmactAdmin.FlexEdit FlexPerNoLab 
            Height          =   3075
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   5424
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cod Tipo-Tipo"
            EncabezadosAnchos=   "300-600-1500"
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483628
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   0
         Left            =   -73710
         TabIndex        =   6
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Index           =   0
         Left            =   -74910
         TabIndex        =   7
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   -73710
         TabIndex        =   9
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   0
         Left            =   -74910
         TabIndex        =   8
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Caption         =   "Vacaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3405
         Index           =   0
         Left            =   -74910
         TabIndex        =   2
         Top             =   345
         Width           =   8565
         Begin SicmactAdmin.FlexEdit FlexPerNoLab 
            Height          =   3075
            Index           =   0
            Left            =   105
            TabIndex        =   3
            Top             =   240
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   5424
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Prog Ini-Prog Fin-Ejec Ini-Ejec Fin-Comentario"
            EncabezadosAnchos=   "300-1200-1200-1200-1200-4000"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-5"
            ListaControles  =   "0-2-2-2-2-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483628
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R-R-R-L"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   0
         Left            =   -71310
         TabIndex        =   4
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   0
         Left            =   -72525
         TabIndex        =   5
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   -73710
         TabIndex        =   17
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   1
         Left            =   -74910
         TabIndex        =   16
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   1
         Left            =   -72525
         TabIndex        =   13
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   1
         Left            =   -71310
         TabIndex        =   12
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Frame fraPerNoLab 
         Caption         =   "Permisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3405
         Index           =   1
         Left            =   -74910
         TabIndex        =   10
         Top             =   360
         Width           =   8565
         Begin SicmactAdmin.FlexEdit FlexPerNoLab 
            Height          =   3075
            Index           =   1
            Left            =   105
            TabIndex        =   11
            Top             =   240
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   5424
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cod Tipo-Tipo-Fech Ini-Fech Fin-Sustento-Apro../Rech...-Estado-Observaciones"
            EncabezadosAnchos=   "300-600-1000-1500-1500-5000-800-3000-5000"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3-4-5-6-X-8"
            ListaControles  =   "0-1-0-2-2-0-1-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483628
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-R-L-L-R-R"
            FormatosEdit    =   "0-0-0-5-5-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   1
         Left            =   -73710
         TabIndex        =   14
         Top             =   3825
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Index           =   1
         Left            =   -74910
         TabIndex        =   15
         Top             =   3825
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPeriodosNoLaborales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub
