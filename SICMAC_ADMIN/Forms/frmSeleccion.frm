VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmRHSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "frmSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   7440
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8175
      TabIndex        =   61
      Top             =   5760
      Width           =   975
   End
   Begin VB.ComboBox cmbEval 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   105
      Width           =   7440
   End
   Begin TabDlg.SSTab Tab 
      Height          =   5205
      Left            =   30
      TabIndex        =   32
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9181
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   459
      WordWrap        =   0   'False
      BackColor       =   12632256
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmSeleccion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdImprimir"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGrabar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNuevo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdEliminar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEditar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraDatosEva"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Comite"
      TabPicture(1)   =   "frmSeleccion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraComite"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Examen"
      TabPicture(2)   =   "frmSeleccion.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fratexto"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Postula..."
      TabPicture(3)   =   "frmSeleccion.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdPostulantesCancelar"
      Tab(3).Control(1)=   "cmdPostulantesEditar"
      Tab(3).Control(2)=   "cmdPostulantesImprimir"
      Tab(3).Control(3)=   "fraPostulantesSeleccion"
      Tab(3).Control(4)=   "cmdPostulantesGrabar"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "&Curricula"
      TabPicture(4)   =   "frmSeleccion.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdExamenGrabar(0)"
      Tab(4).Control(1)=   "cmdExamenEditar(0)"
      Tab(4).Control(2)=   "cmdExamenCancelar(0)"
      Tab(4).Control(3)=   "fraExamenSeleccion(0)"
      Tab(4).Control(4)=   "cmdExamenImprimir(0)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Esc&rito"
      TabPicture(5)   =   "frmSeleccion.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdExamenCancelar(1)"
      Tab(5).Control(1)=   "cmdExamenImprimir(1)"
      Tab(5).Control(2)=   "fraExamenSeleccion(1)"
      Tab(5).Control(3)=   "cmdExamenGrabar(1)"
      Tab(5).Control(4)=   "cmdExamenEditar(1)"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Psicolo&gico"
      TabPicture(6)   =   "frmSeleccion.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdExamenGrabar(2)"
      Tab(6).Control(1)=   "fraExamenSeleccion(2)"
      Tab(6).Control(2)=   "cmdExamenImprimir(2)"
      Tab(6).Control(3)=   "cmdExamenEditar(2)"
      Tab(6).Control(4)=   "cmdExamenCancelar(2)"
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "Entrevis&ta"
      TabPicture(7)   =   "frmSeleccion.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdExamenCancelar(3)"
      Tab(7).Control(1)=   "cmdExamenEditar(3)"
      Tab(7).Control(2)=   "cmdExamenImprimir(3)"
      Tab(7).Control(3)=   "fraExamenSeleccion(3)"
      Tab(7).Control(4)=   "cmdExamenGrabar(3)"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   " Res&ultados"
      TabPicture(8)   =   "frmSeleccion.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "cmdExamenCancelar(4)"
      Tab(8).Control(1)=   "cmdExamenEditar(4)"
      Tab(8).Control(2)=   "cmdExamenImprimir(4)"
      Tab(8).Control(3)=   "fraExamenSeleccion(4)"
      Tab(8).Control(4)=   "cmdExamenGrabar(4)"
      Tab(8).Control(5)=   "cmdCerrar"
      Tab(8).ControlCount=   6
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   -71670
         TabIndex        =   84
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   4
         Left            =   -74880
         TabIndex        =   58
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraExamenSeleccion 
         Caption         =   "Resultados"
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
         Height          =   3810
         Index           =   4
         Left            =   -74895
         TabIndex        =   81
         Top             =   300
         Width           =   8940
         Begin VB.TextBox txtComentarioFinal 
            Appearance      =   0  'Flat
            Height          =   930
            Left            =   105
            MaxLength       =   200
            ScrollBars      =   3  'Both
            TabIndex        =   56
            Top             =   2760
            Width           =   8730
         End
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   2280
            Index           =   4
            Left            =   60
            TabIndex        =   55
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   4022
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-3000-1000"
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
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R"
            FormatosEdit    =   "0-0-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
         End
         Begin VB.Label lblComentarioFinal 
            Caption         =   "Comentario Final"
            Height          =   240
            Left            =   105
            TabIndex        =   83
            Top             =   2565
            Width           =   1800
         End
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   4
         Left            =   -72750
         TabIndex        =   59
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   3
         Left            =   -74900
         TabIndex        =   52
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraExamenSeleccion 
         Caption         =   "Notas Examen Entrevista"
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
         Height          =   3810
         Index           =   3
         Left            =   -74895
         TabIndex        =   80
         Top             =   300
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3450
            Index           =   3
            Left            =   45
            TabIndex        =   50
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   6085
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R"
            FormatosEdit    =   "0-0-0-2"
            AvanceCeldas    =   1
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   3
         Left            =   -72750
         TabIndex        =   53
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   2
         Left            =   -74900
         TabIndex        =   48
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraExamenSeleccion 
         Caption         =   "Notas Examen Psicologico"
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
         Height          =   3810
         Index           =   2
         Left            =   -74895
         TabIndex        =   79
         Top             =   300
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3450
            Index           =   2
            Left            =   60
            TabIndex        =   62
            Top             =   255
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   6085
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R"
            FormatosEdit    =   "0-0-0-2"
            AvanceCeldas    =   1
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   2
         Left            =   -72750
         TabIndex        =   5
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   1
         Left            =   -73830
         TabIndex        =   41
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   1
         Left            =   -74900
         TabIndex        =   44
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraExamenSeleccion 
         Caption         =   "Notas Examen Escrito"
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
         Height          =   3810
         Index           =   1
         Left            =   -74895
         TabIndex        =   77
         Top             =   300
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3450
            Index           =   1
            Left            =   60
            TabIndex        =   78
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   6085
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R"
            FormatosEdit    =   "0-0-0-2"
            AvanceCeldas    =   1
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   1
         Left            =   -72750
         TabIndex        =   45
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   0
         Left            =   -72780
         TabIndex        =   43
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraExamenSeleccion 
         Caption         =   "Notas Examen Curricular"
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
         Height          =   3810
         Index           =   0
         Left            =   -74895
         TabIndex        =   76
         Top             =   300
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3450
            Index           =   0
            Left            =   60
            TabIndex        =   38
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   6085
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R"
            FormatosEdit    =   "0-0-0-2"
            AvanceCeldas    =   1
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdPostulantesGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   -74895
         TabIndex        =   30
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraPostulantesSeleccion 
         Caption         =   "Personas Seleccion"
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
         Height          =   3795
         Left            =   -74895
         TabIndex        =   75
         Top             =   300
         Width           =   8895
         Begin VB.CommandButton cmdPostulantesNuevo 
            Caption         =   "&Nuevo"
            Height          =   375
            Left            =   6720
            TabIndex        =   33
            Top             =   3345
            Width           =   975
         End
         Begin VB.CommandButton cmdPostulantesEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   7815
            TabIndex        =   35
            Top             =   3345
            Width           =   975
         End
         Begin Sicmact.FlexEdit FlexPostulantes 
            Height          =   3015
            Left            =   60
            TabIndex        =   34
            Top             =   255
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   5318
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Codigo-Nombre"
            EncabezadosAnchos=   "500-1500-6000"
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
            ColumnasAEditar =   "X-1-X"
            ListaControles  =   "0-1-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbRsLoad        =   -1  'True
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   285
         End
      End
      Begin VB.CommandButton cmdPostulantesImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   -72810
         TabIndex        =   31
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdPostulantesEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   -73845
         TabIndex        =   27
         Top             =   4200
         Width           =   975
      End
      Begin VB.Frame fraDatosEva 
         Caption         =   "Datos"
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
         Height          =   4485
         Left            =   105
         TabIndex        =   37
         Top             =   300
         Width           =   8925
         Begin VB.TextBox txtNotaMax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   7
            Text            =   "100"
            Top             =   3120
            Width           =   1230
         End
         Begin Sicmact.TxtBuscar TxtBuscarAreaCargo 
            Height          =   270
            Left            =   1170
            TabIndex        =   4
            Top             =   1920
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   476
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
            sTitulo         =   ""
         End
         Begin VB.Frame fraExamenes 
            Caption         =   "Examenes"
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
            Height          =   1320
            Left            =   2550
            TabIndex        =   74
            Top             =   2595
            Width           =   6195
            Begin Spinner.uSpinner uscurricular 
               Height          =   255
               Left            =   300
               TabIndex        =   86
               Top             =   960
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               Max             =   5
               MaxLength       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin VB.CheckBox chkEnt 
               Caption         =   "Exa Entrevista"
               Height          =   195
               Left            =   4695
               TabIndex        =   11
               Top             =   300
               Width           =   1425
            End
            Begin VB.CheckBox chkPsi 
               Caption         =   "Exa Psciologico"
               Height          =   195
               Left            =   3255
               TabIndex        =   10
               Top             =   300
               Width           =   1425
            End
            Begin VB.CheckBox chkEsc 
               Caption         =   "Exa Escrito"
               Height          =   195
               Left            =   1785
               TabIndex        =   9
               Top             =   285
               Width           =   1425
            End
            Begin VB.CheckBox chkCur 
               Caption         =   "Exa Curricular"
               Height          =   195
               Left            =   300
               TabIndex        =   8
               Top             =   285
               Width           =   1425
            End
            Begin Spinner.uSpinner usescrito 
               Height          =   255
               Left            =   1785
               TabIndex        =   87
               Top             =   960
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               Max             =   5
               MaxLength       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner uspsicologico 
               Height          =   255
               Left            =   3240
               TabIndex        =   88
               Top             =   960
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               Max             =   5
               MaxLength       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin Spinner.uSpinner usentrevista 
               Height          =   255
               Left            =   4680
               TabIndex        =   89
               Top             =   960
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               Max             =   5
               MaxLength       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ponderados"
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
               Height          =   195
               Left            =   120
               TabIndex        =   85
               Top             =   600
               Visible         =   0   'False
               Width           =   1020
            End
         End
         Begin VB.ComboBox cmbTipoContrato 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2220
            Width           =   7575
         End
         Begin VB.Frame fraFechas 
            Caption         =   "Fechas"
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
            Height          =   600
            Left            =   105
            TabIndex        =   66
            Top             =   3840
            Width           =   8640
            Begin MSMask.MaskEdBox mskFF 
               Height          =   315
               Left            =   5040
               TabIndex        =   13
               Top             =   225
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFI 
               Height          =   315
               Left            =   1005
               TabIndex        =   12
               Top             =   225
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblFF 
               Caption         =   "Fin :"
               Height          =   255
               Left            =   4320
               TabIndex        =   68
               Top             =   255
               Width           =   735
            End
            Begin VB.Label lblFI 
               Caption         =   "Inicio :"
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   255
               Width           =   855
            End
         End
         Begin VB.TextBox txtDes 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   1170
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   960
            Width           =   7605
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1575
            Width           =   7605
         End
         Begin VB.Frame fraProp 
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
            Height          =   660
            Left            =   105
            TabIndex        =   70
            Top             =   240
            Width           =   8685
            Begin VB.ComboBox cmbEstado 
               Height          =   315
               Left            =   975
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   255
               Width           =   7590
            End
            Begin VB.Label lblEstado 
               Caption         =   "Estado :"
               Height          =   255
               Left            =   210
               TabIndex        =   71
               Top             =   270
               Width           =   735
            End
         End
         Begin VB.Label lblNotMaxima 
            Caption         =   "Nota Maxima :"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   3120
            Width           =   1050
         End
         Begin VB.Label lblTipoContrato 
            Caption         =   "Tpo.Contrato:"
            Height          =   255
            Left            =   105
            TabIndex        =   72
            Top             =   2250
            Width           =   1050
         End
         Begin VB.Label lblTipo 
            Caption         =   "Tipo"
            Height          =   255
            Left            =   90
            TabIndex        =   65
            Top             =   1605
            Width           =   615
         End
         Begin VB.Label lblDes 
            Caption         =   "Descripción :"
            Height          =   255
            Left            =   90
            TabIndex        =   64
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblCargo 
            Caption         =   "Cargo :"
            Height          =   255
            Left            =   90
            TabIndex        =   63
            Top             =   1928
            Width           =   735
         End
         Begin VB.Label lblDescripcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   2925
            TabIndex        =   60
            Top             =   1950
            Width           =   5820
         End
      End
      Begin VB.Frame fraComite 
         Caption         =   "Comite"
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
         Height          =   4260
         Left            =   -74880
         TabIndex        =   73
         Top             =   345
         Width           =   8880
         Begin VB.CommandButton cmdEliminarComite 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   7785
            TabIndex        =   22
            Top             =   3795
            Width           =   975
         End
         Begin VB.CommandButton cmdNuevoComite 
            Caption         =   "N&uevo"
            Height          =   375
            Left            =   6750
            TabIndex        =   20
            Top             =   3795
            Width           =   975
         End
         Begin Sicmact.FlexEdit FlexComite 
            Height          =   3480
            Left            =   90
            TabIndex        =   21
            Top             =   240
            Width           =   8670
            _ExtentX        =   15293
            _ExtentY        =   6138
            Cols0           =   4
            HighLight       =   1
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Codigo-Nombre-Cargo"
            EncabezadosAnchos=   "350-1800-4000-3000"
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
            ColumnasAEditar =   "X-1-X-3"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-3"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L"
            FormatosEdit    =   "0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   240
            CellBackColor   =   -2147483624
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   4785
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2100
         TabIndex        =   18
         Top             =   4785
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "N&uevo"
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   4800
         Width           =   975
      End
      Begin VB.Frame fratexto 
         Caption         =   "Texto"
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
         Height          =   4245
         Left            =   -74895
         TabIndex        =   69
         Top             =   300
         Width           =   8895
         Begin VB.CommandButton cmdCargarP 
            Caption         =   "&Cargar Examen Psicologico"
            Height          =   375
            Left            =   4665
            TabIndex        =   26
            Top             =   3810
            Width           =   4065
         End
         Begin VB.CommandButton cmdCargarE 
            Caption         =   "&Cargar Examen Escrito"
            Height          =   375
            Left            =   150
            TabIndex        =   24
            Top             =   3810
            Width           =   4035
         End
         Begin RichTextLib.RichTextBox REscrito 
            Height          =   3525
            Left            =   150
            TabIndex        =   23
            Top             =   225
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   6218
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmSeleccion.frx":0406
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox RPsicologico 
            Height          =   3525
            Left            =   4650
            TabIndex        =   25
            Top             =   225
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   6218
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmSeleccion.frx":0486
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdPostulantesCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   -73845
         TabIndex        =   28
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   -73845
         TabIndex        =   40
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   -73830
         TabIndex        =   42
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   2
         Left            =   -73830
         TabIndex        =   46
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   -73830
         TabIndex        =   47
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   3
         Left            =   -73830
         TabIndex        =   49
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   4
         Left            =   -73830
         TabIndex        =   54
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   -73830
         TabIndex        =   51
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   -73830
         TabIndex        =   57
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   0
         Left            =   -73845
         TabIndex        =   36
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   0
         Left            =   -74900
         TabIndex        =   39
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   105
         TabIndex        =   19
         Top             =   4800
         Width           =   975
      End
   End
   Begin VB.Label lblEval 
      Caption         =   "Evaluación :"
      Height          =   255
      Left            =   135
      TabIndex        =   29
      Top             =   135
      Width           =   975
   End
End
Attribute VB_Name = "frmRHSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnTipo As TipoOpe
Dim lsCodigo As String
Dim lbEditado As Boolean
Dim loVar As RHProcesoSeleccionTipo
Dim lnTipoSeleccionMnu As RHPSeleccionTpoMnu
Dim lbCerrar As Boolean
Dim lnModo As RHProcesoSeleccionModal
Dim lbGanador As Boolean
Dim lnTipEva As RHTipoOpeEvaluacion
Dim lnIndiceActual As Integer
'ALPA 20090122******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub CargaDatos(pbTodos As Boolean, Optional pbSeleccion As Boolean = True)
    Dim rsE As ADODB.Recordset
    Dim oDatoAreas As DActualizaDatosArea
    Set oDatoAreas = New DActualizaDatosArea
    Set rsE = New ADODB.Recordset
    
    Me.TxtBuscarAreaCargo.rs = oDatoAreas.GetCargosAreas
    If pbSeleccion Then
        Dim oEva As DActualizaProcesoSeleccion
        Set oEva = New DActualizaProcesoSeleccion
        If lnTipo = gTipoOpeConsulta And lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
            Set rsE = oEva.GetProcesosSeleccion(-1)
        Else
            Set rsE = oEva.GetProcesosSeleccion(RHProcesoSeleccionEstado.gRHProcSelEstIniCiado)
        End If
        CargaCombo rsE, Me.cmbEval, 5
        rsE.Close
        Set oEva = Nothing
    End If
    CargaComite
    If pbTodos Then
        Dim oCons As DConstantes
        Set oCons = New DConstantes
        Set rsE = New ADODB.Recordset
        Set rsE = oCons.GetConstante(gRHProcesoSeleccionTipo)
        CargaCombo rsE, Me.cmbTipo, 200
        rsE.Close
        Set rsE = oCons.GetConstante(gRHContratoTipo)
        CargaCombo rsE, Me.cmbTipoContrato, 200
        rsE.Close
        Set rsE = oCons.GetConstante(gRHEvaluacionComite)
        Me.FlexComite.CargaCombo rsE
        Set rsE = oCons.GetConstante(gRHProcesoSeleccionEstado)
        CargaCombo rsE, Me.cmbEstado, 200
        rsE.Close
        Set rsE = Nothing
        Set oCons = Nothing
    End If
End Sub

Private Sub chkCur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       chkEsc.SetFocus
    End If
End Sub

Private Sub chkEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.mskFI.SetFocus
    End If
End Sub

Private Sub chkEsc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       chkPsi.SetFocus
    End If
End Sub

Private Sub chkPsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       chkEnt.SetFocus
    End If
End Sub

Private Sub cmbEval_Click()
    Dim oEva As DActualizaProcesoSeleccion
    Dim rsE As New ADODB.Recordset
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rsE = oEva.GetProcesoSeleccion(Trim(Left(Me.cmbEval.Text, 6)))
    If Not (rsE.EOF And rsE.BOF) Then
        Me.txtDes.Text = rsE!Comentario
        Me.TxtBuscarAreaCargo = rsE!area & rsE!Cargo
        Me.lblDescripcion.Caption = Me.TxtBuscarAreaCargo.psDescripcion
        UbicaCombo Me.cmbTipo, rsE!Tipo
        UbicaCombo Me.cmbEstado, rsE!Estado
        UbicaCombo Me.cmbTipoContrato, IIf(IsNull(rsE!TipoContrato), "", rsE!TipoContrato)
        Me.mskFI.Text = Format(rsE!fi, gsFormatoFechaView)
        Me.mskFF.Text = Format(rsE!ff, gsFormatoFechaView)
        Me.REscrito.Text = rsE!escrito
        Me.RPsicologico.Text = rsE!Psico
        Me.chkCur.value = IIf(IIf(IsNull(rsE!bCur), False, rsE!bCur), 1, 0)
        Me.chkEsc.value = IIf(IIf(IsNull(rsE!bEsc), False, rsE!bEsc), 1, 0)
        Me.chkPsi.value = IIf(IIf(IsNull(rsE!bPsi), False, rsE!bPsi), 1, 0)
        Me.chkEnt.value = IIf(IIf(IsNull(rsE!bEnt), False, rsE!bEnt), 1, 0)
        Me.txtNotaMax.Text = IIf(IsNull(rsE!NotaMax), 0, rsE!NotaMax)
        Me.txtComentarioFinal.Text = IIf(IsNull(rsE!ComentarioFinal), "", rsE!ComentarioFinal)
        
        'agregado
        'nRHPesoExaCur,nRHPesoExaEsc,nRHPesoExaPsico,nRHPesoExaEnt
        'uscurricular.Valor = IIf(IsNull(rsE!nRHPesoExaCur), 1, rsE!nRHPesoExaCur)
        'usescrito.Valor = IIf(IsNull(rsE!nRHPesoExaEsc), 1, rsE!nRHPesoExaEsc)
        'uspsicologico.Valor = IIf(IsNull(rsE!nRHPesoExaPsico), 1, rsE!nRHPesoExaPsico)
        'usentrevista.Valor = IIf(IsNull(rsE!nRHPesoExaEnt), 1, rsE!nRHPesoExaEnt)
        
        
        CargaComite
    Else
        Me.FlexComite.Rows = 1
        Me.FlexComite.Rows = 2
        Me.FlexComite.FixedRows = 1
    End If
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuPost Then
        CargaDatosPostulantes
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaCur Or lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEsc Or lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaPsi Or lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEnt Then
        CargaDatosNotas
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
        CargaDatosPostulantes
        lnTipEva = RHTipoOpeEvaCurricular
        lnIndiceActual = 0
        CargaDatosNotas
        lnTipEva = RHTipoOpeEvaEscrito
        lnIndiceActual = 1
        CargaDatosNotas
        lnTipEva = RHTipoOpeEvaPsicologico
        lnIndiceActual = 2
        CargaDatosNotas
        lnTipEva = RHTipoOpeEvaEntrevista
        lnIndiceActual = 3
        CargaDatosNotas
        lnTipEva = RHTipoOpeEvaConsolidado
        lnIndiceActual = 4
        CargaDatosNotas
    End If
End Sub

Private Sub cmbEval_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lnTipoSeleccionMnu = RHPSeleccionTpoMnuSel Then
            If lnTipo = gTipoOpeMantenimiento Then
                Me.cmdEditar.SetFocus
            ElseIf lnTipo = gTipoOpeReporte Then
                Me.cmdImprimir.SetFocus
            ElseIf Me.cmdCerrar.Enabled Then
                Me.cmdCerrar.SetFocus
            End If
        ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuPost Then
            If Me.cmdPostulantesEditar.Enabled Then
                cmdPostulantesEditar.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtBuscarAreaCargo.SetFocus
    End If
End Sub

Private Sub cmbTipoContrato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.txtNotaMax.SetFocus
    End If
End Sub

Private Sub cmdCerrar_Click()
    Dim oEval As DActualizaProcesoSeleccion
    Set oEval = New DActualizaProcesoSeleccion
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    If Me.cmbEval.Text = "" Then Exit Sub
    If Not oEval.ValidaNotasNulas(Left(Me.cmbEval.Text, 6)) Then
        lbEditado = True
        If MsgBox("Desea Cerrar proceso Seleccion ?  - El Proceso no podra se modificado.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        UbicaCombo Me.cmbEstado, RHProcesoSeleccionEstado.gRHProcSelEstFinalizado
        oEva.ModificaEstadoProSelec Left(Me.cmbEval, 6), Right(Me.cmbEstado.Text, 1), GetMovNro(gsCodUser, gsCodAge)
        cmdCancelar_Click
        Me.txtComentarioFinal.Text = ""
    Else
        MsgBox "Debe de Ingresar todas la notas para poder cerrar un proceso de Seleccion.", vbInformation, "Aviso"
        Me.FlexExamen(4).Clear
        Me.FlexExamen(4).Rows = 2
        Me.FlexExamen(4).FormaCabecera
    End If
    
    Set oEva = Nothing
    Set oEval = Nothing
    
    Form_Load
End Sub

Private Sub cmdCancelar_Click()
    Limpia
    Activa False, lnTipo
    If lnTipo = gTipoOpeRegistro Then Me.cmdNuevo.SetFocus
End Sub

Private Sub cmdCargarE_Click()
    CDialog.CancelError = False
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos txt(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Me.REscrito.LoadFile CDialog.FileName, 1
    
    If REscrito.Text <> "" Then
        REscrito.Text = Replace(Me.REscrito.Text, Chr(38), "")
        REscrito.Text = Replace(Me.REscrito.Text, Chr(34), "")
        REscrito.Text = Replace(Me.REscrito.Text, oImpresora.gPrnSaltoLinea, "")
        REscrito.Text = Replace(Me.REscrito.Text, Chr(13), "")
        REscrito.Text = "-" & REscrito.Text
    End If
End Sub

Private Sub cmdCargarP_Click()
    CDialog.CancelError = False
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos txt(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Me.RPsicologico.LoadFile CDialog.FileName, 1
    
    If RPsicologico.Text <> "" Then
        RPsicologico.Text = Replace(Me.RPsicologico.Text, Chr(38), "")
        RPsicologico.Text = Replace(Me.RPsicologico.Text, Chr(34), "")
        RPsicologico.Text = Replace(Me.RPsicologico.Text, Chr(13), "")
        RPsicologico.Text = Replace(Me.RPsicologico.Text, oImpresora.gPrnSaltoLinea, "")
        RPsicologico.Text = "-" & RPsicologico.Text
    End If
End Sub

Private Sub cmdEditar_Click()
    If Me.cmbEval.Text = "" Then
        Me.cmbEval.SetFocus
        Exit Sub
    End If
    lbEditado = True
    Activa True, lnTipo
    Me.cmbEstado.SetFocus
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDes.SetFocus
    End If
End Sub

Private Sub CmdEliminar_Click()
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    If MsgBox("Desea Elimiar la Evaluación. " & Trim(Left(Me.cmbEval.Text, 50)) & Chr(13) & "Se eliminaran todas las personas relacionadas con esta Evaluacion.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    oEva.EliminaProSelec Trim(Left(Me.cmbEval.List(cmbEval.ListCount - 1), 6))
    CargaDatos False
    Limpia
End Sub

Private Sub cmdGrabar_Click()
    Dim oEval As NActualizaProcesoSeleccion
    Set oEval = New NActualizaProcesoSeleccion
    Dim lsUltActualizacion As String
    If Not Valida() Then Exit Sub
    
    If MsgBox("Desea Grabar los cambios", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lsUltActualizacion = GetMovNro(gsCodUser, gsCodAge)
    glsMovNro = lsUltActualizacion
    If lbEditado Then
        oEval.ModificaProSelec Left(Me.cmbEval, 6), Right(Me.cmbTipo, 1), Left(TxtBuscarAreaCargo, 3), Right(Me.TxtBuscarAreaCargo, 6), Format(Me.mskFI.Text, gsFormatoFecha), Format(Me.mskFF.Text, gsFormatoFecha), Right(cmbEstado, 1), Me.txtDes.Text, Me.RPsicologico.Text, Me.REscrito.Text, Me.chkCur.value, Me.chkEsc.value, Me.chkPsi.value, Me.chkEnt.value, Me.txtNotaMax.Text, Right(Me.cmbTipoContrato.Text, 1), lsUltActualizacion, uscurricular.Valor, usescrito.Valor, uspsicologico.Valor, usentrevista.Valor
        oEval.AgregaComiteProSelec Left(Me.cmbEval, 6), Me.FlexComite.GetRsNew, lsUltActualizacion
        'ALPA 20090122 **********************************************************
        gsOpeCod = LogPistaModificaProcesoSeleccion
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", , Left(Me.cmbEval, 6), gNumeroProcesoSeleccion
        '************************************************************************
    Else
        oEval.AgregaProSelec lsCodigo, Right(Me.cmbTipo, 1), Left(Me.TxtBuscarAreaCargo, 3), Right(Me.TxtBuscarAreaCargo, 6), Format(Me.mskFI.Text, gsFormatoFecha), Format(Me.mskFF.Text, gsFormatoFecha), Right(cmbEstado, 1), Me.txtDes.Text, Mid(Me.RPsicologico.Text, 8), Mid(Me.REscrito.Text, 8), Me.chkCur.value, Me.chkEsc.value, Me.chkPsi.value, Me.chkEnt.value, Me.txtNotaMax.Text, Right(Me.cmbTipoContrato.Text, 1), lsUltActualizacion, uscurricular.Valor, usescrito.Valor, uspsicologico.Valor, usentrevista.Valor
        oEval.AgregaComiteProSelec lsCodigo, Me.FlexComite.GetRsNew, lsUltActualizacion
         
        'ALPA 20090122 **********************************************************
        gsOpeCod = LogPistaRegistraProcesoSeleccion
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", , lsCodigo, gNumeroProcesoSeleccion
        '************************************************************************
         
    End If
        
    Limpia
    Activa False, lnTipo
    CargaDatos False
End Sub

Private Sub cmdImprimir_Click()
    Dim oEval As NActualizaProcesoSeleccion
    Dim lsCadena As String
    Dim lsCadenaTemp As String
    Dim lbRep(1) As Boolean
    Dim oPrevio As Previo.clsPrevio
    Set oEval = New NActualizaProcesoSeleccion
    Set oPrevio = New Previo.clsPrevio
    
    'frmImpreRRHH.Ini "Lista de Examenes;Acta de Ingresos;", "Evaluaicon", lbRep, gdFecSis, gdFecSis, False
    frmImpreRRHH.Ini "Lista de Examenes;", "Evaluacion", lbRep, gdFecSis, gdFecSis, False
    
    If lbRep(1) Then
        lsCadena = lsCadena & oEval.GetReporte(gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    If lbRep(1) And Me.cmbEval.Text <> "" Then
        If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        
        lsCadena = lsCadena & oEval.GetActa(Left(Me.cmbEval.Text, 6), gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    If lsCadena <> "" Then oPrevio.Show lsCadena, " Evaluaciones ", True, 66
    Set oEval = Nothing
    Set oPrevio = Nothing
End Sub

Private Sub CmdNuevo_Click()
    Me.cmbEval.ListIndex = -1
    lbEditado = False
    Limpia
    Activa True, lnTipo
    If lnTipo <> gTipoOpeRegistro Then
        Me.cmbEstado.SetFocus
    Else
        UbicaCombo cmbEstado, RHProcesoSeleccionEstado.gRHProcSelEstIniCiado
        Me.txtDes.SetFocus
    End If
End Sub

Private Sub cmdEliminarComite_Click()
    Me.FlexComite.EliminaFila Me.FlexComite.Row
End Sub

Private Sub cmdNuevoComite_Click()
    Dim oPersona As UPersona
    If Me.FlexComite.TextMatrix(Me.FlexComite.Rows - 1, 0) = "#" Then
        FlexComite.Rows = 2
    End If
    
    If Me.FlexComite.TextMatrix(Me.FlexComite.Rows - 1, 0) = "" Then
        FlexComite.AdicionaFila 1
    Else
        FlexComite.AdicionaFila CLng(Me.FlexComite.TextMatrix(FlexComite.Rows - 1, 0)) + 1
    End If
    FlexComite.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub




Private Sub FlexExamen_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If Not IsNumeric(Me.FlexExamen(lnIndiceActual).TextMatrix(pnRow, pnCol)) Then Exit Sub
    If CCur(Me.FlexExamen(lnIndiceActual).TextMatrix(pnRow, pnCol)) <= CCur(Me.txtNotaMax.Text) And CCur(Me.FlexExamen(lnIndiceActual).TextMatrix(pnRow, pnCol)) >= 0 Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub FlexPostulantes_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oEval As DActualizaProcesoSeleccion
    Set oEval = New DActualizaProcesoSeleccion
    
    'oEval.PersonaComite
End Sub

Private Sub FlexPostulantes_RowColChange()
    Dim oEval As DActualizaProcesoSeleccion
    Set oEval = New DActualizaProcesoSeleccion
    
    If oEval.PersonaComite(Left(Me.cmbEval.Text, 6), Me.FlexPostulantes.TextMatrix(FlexPostulantes.Row, 1)) Then
        Me.FlexPostulantes.TextMatrix(FlexPostulantes.Row, 1) = ""
        Me.FlexPostulantes.TextMatrix(FlexPostulantes.Row, 2) = ""
    End If

End Sub

Private Sub Form_Load()

    'ALPA 20090122 ***************************************************************************
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuSel Then
        If lnTipo = gTipoOpeRegistro Then
            CargaDatos True, False
        Else
            CargaDatos True
        End If
    Else
        CargaDatos True
    End If
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaCur Then
        lnTipEva = RHTipoOpeEvaCurricular
        lnIndiceActual = 0
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEnt Then
        lnTipEva = RHTipoOpeEvaEntrevista
        lnIndiceActual = 3
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEsc Then
        lnTipEva = RHTipoOpeEvaEscrito
        lnIndiceActual = 1
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaPsi Then
        lnTipEva = RHTipoOpeEvaPsicologico
        lnIndiceActual = 2
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
        lnTipEva = RHTipoOpeEvaConsolidado
        lnIndiceActual = 4
    End If
    
    Activa False, lnTipo
End Sub

Public Sub Ini(pnTipo As TipoOpe, pnTipoSeleccionMnu As RHPSeleccionTpoMnu, psCaption As String)
    lnTipoSeleccionMnu = pnTipoSeleccionMnu
    lnTipo = pnTipo
    lbCerrar = False
    Caption = psCaption
    IniTab
    Me.Show 1
End Sub

Private Sub mskFF_GotFocus()
    mskFF.SelStart = 0
    mskFF.SelLength = 10
End Sub

Private Sub mskFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

Private Sub mskFI_GotFocus()
    mskFI.SelStart = 0
    mskFI.SelLength = 10
End Sub

Private Sub Activa(pbvalor As Boolean, pnTipo As TipoOpe)
    'Objetos
    Me.cmbEval.Enabled = Not pbvalor
    Dim lnI As Integer
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuSel Then
        Me.cmdNuevo.Visible = Not pbvalor
        Me.cmdEditar.Visible = Not pbvalor
        Me.cmdGrabar.Visible = pbvalor
        Me.cmdSalir.Enabled = Not pbvalor
        Me.fraDatosEva.Enabled = pbvalor
        If pnTipo = gTipoOpeRegistro Then
            Me.cmdEditar.Visible = False
            Me.cmdCancelar.Enabled = pbvalor
            Me.fraProp.Enabled = False
            Me.fraComite.Enabled = pbvalor
            Me.cmbEval.Enabled = False
            Me.cmdEliminar.Visible = False
            Me.cmdImprimir.Visible = False
            Me.fratexto.Enabled = True
        ElseIf pnTipo = gTipoOpeMantenimiento Then
            Me.fraProp.Enabled = pbvalor
            Me.cmdCancelar.Visible = pbvalor
            Me.cmdEliminar.Enabled = Not pbvalor
            Me.cmdImprimir.Visible = False
            Me.cmdNuevo.Visible = False
            Me.cmdGrabar.Visible = True
            Me.cmdGrabar.Enabled = pbvalor
            Me.fraComite.Enabled = pbvalor
        ElseIf pnTipo = gTipoOpeConsulta Then
            Me.cmdEliminarComite.Visible = False
            Me.cmdNuevoComite.Visible = False
            Me.cmdPostulantesNuevo.Visible = False
            Me.cmdPostulantesEliminar.Visible = False
            Me.cmdEliminar.Enabled = False
            Me.fraProp.Enabled = False
            Me.cmdNuevo.Visible = False
            Me.cmdCancelar.Visible = False
            Me.cmdEditar.Visible = False
            Me.cmdImprimir.Visible = False
            Me.cmdEliminar.Visible = False
            Me.fraDatosEva.Enabled = pbvalor
            fraFechas.Enabled = pbvalor
            Me.cmdCargarE.Visible = pbvalor
            Me.cmdCargarP.Visible = pbvalor
            Me.cmdCerrar.Visible = lbCerrar
            Me.txtDes.Enabled = pbvalor
            Me.TxtBuscarAreaCargo.Enabled = pbvalor
            Me.cmbTipo.Enabled = pbvalor
            Me.fraComite.Enabled = False
        ElseIf pnTipo = gTipoOpeReporte Then
            Me.txtDes.Enabled = pbvalor
            Me.TxtBuscarAreaCargo.Enabled = pbvalor
            Me.cmbTipo.Enabled = pbvalor
            Me.cmdEliminar.Visible = pbvalor
            Me.cmdNuevo.Visible = pbvalor
            Me.cmdEditar.Visible = pbvalor
            Me.cmbEval.Enabled = True
            Me.fraProp.Enabled = pbvalor
            Me.fraComite.Enabled = False
            Me.fraExamenes.Enabled = False
            Me.cmdCancelar.Visible = False
            Me.cmdCargarE.Visible = False
            Me.cmdCargarP.Visible = False
            Me.cmdNuevoComite.Visible = False
            Me.cmdEliminarComite.Visible = False
        End If
    Else
        Me.cmdEliminar.Enabled = False
        Me.fraProp.Enabled = False
        Me.cmdNuevo.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdImprimir.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdCargarE.Visible = False
        Me.cmdCargarP.Visible = False
        Me.fraDatosEva.Enabled = pbvalor
        fraFechas.Enabled = pbvalor
        Me.cmdCerrar.Visible = False
        Me.txtDes.Enabled = pbvalor
        Me.TxtBuscarAreaCargo.Enabled = pbvalor
        Me.cmbTipo.Enabled = pbvalor
        Me.fraComite.Enabled = False
        Me.cmdEliminarComite.Visible = False
        Me.cmdNuevoComite.Visible = False
        Me.cmdSalir.Enabled = Not pbvalor
        
        If lnTipoSeleccionMnu = RHPSeleccionTpoMnuPost Then
            Me.cmdPostulantesEditar.Visible = Not pbvalor
            Me.cmdPostulantesGrabar.Enabled = pbvalor
            Me.cmdPostulantesCancelar.Visible = pbvalor
            Me.fraPostulantesSeleccion.Enabled = pbvalor
            Me.cmbEval.Enabled = Not pbvalor
            If lnTipo = gTipoOpeRegistro Then
                Me.cmdPostulantesEliminar.Visible = False
                Me.cmdPostulantesImprimir.Visible = False
            ElseIf lnTipo = gTipoOpeConsulta Then
                Me.cmdPostulantesEditar.Visible = False
                Me.cmdPostulantesImprimir.Visible = False
                Me.cmdPostulantesGrabar.Visible = False
                Me.cmdPostulantesCancelar.Visible = False
                Me.cmdPostulantesNuevo.Visible = False
                Me.cmdPostulantesEliminar.Visible = False
            ElseIf lnTipo = gTipoOpeReporte Then
                Me.cmdPostulantesEditar.Visible = False
                Me.cmdPostulantesNuevo.Visible = False
                Me.cmdPostulantesEliminar.Visible = False
                Me.cmdPostulantesGrabar.Visible = False
                Me.cmdPostulantesCancelar.Visible = False
            ElseIf lnTipo = gTipoOpeMantenimiento Then
                Me.cmdPostulantesNuevo.Enabled = False
                Me.cmdPostulantesImprimir.Visible = False
            End If
        
        ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
            Me.cmdPostulantesCancelar.Visible = False
            Me.cmdPostulantesEditar.Visible = False
            Me.cmdPostulantesGrabar.Visible = False
            Me.cmdPostulantesEliminar.Visible = False
            Me.cmdPostulantesNuevo.Visible = False
            
            For lnI = 0 To 3
                Me.cmdExamenEditar(lnI).Visible = False
                Me.cmdExamenGrabar(lnI).Visible = False
                Me.cmdExamenCancelar(lnI).Visible = False
                Me.fraExamenSeleccion(lnI).Enabled = False
                If lnTipo = gTipoOpeConsulta Then Me.cmdExamenImprimir(lnI).Visible = False
            Next
        
            If lnTipo = gTipoOpeConsulta Then
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
                Me.cmdExamenGrabar(lnIndiceActual).Visible = False
                Me.cmdExamenCancelar(lnIndiceActual).Visible = False
                If lnIndiceActual = 4 Then
                    Me.cmdExamenImprimir(lnIndiceActual).Visible = True
                Else
                    Me.cmdExamenImprimir(lnIndiceActual).Visible = False
                End If
                Me.cmdCerrar.Visible = False
                Me.fraExamenSeleccion(lnIndiceActual).Enabled = False
                Me.cmdPostulantesImprimir.Visible = False
            Else
                Me.cmdExamenEditar(lnIndiceActual).Visible = Not pbvalor
                Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbvalor
                Me.cmdExamenCancelar(lnIndiceActual).Visible = pbvalor
                Me.fraExamenSeleccion(lnIndiceActual).Enabled = pbvalor
                Me.cmdSalir.Enabled = Not pbvalor
                Me.cmdCerrar.Visible = True
                Me.cmdCerrar.Enabled = True
            End If

        Else
            Me.cmdExamenEditar(lnIndiceActual).Visible = Not pbvalor
            Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbvalor
            Me.cmdExamenCancelar(lnIndiceActual).Visible = pbvalor
            Me.fraExamenSeleccion(lnIndiceActual).Enabled = pbvalor
            Me.cmbEval.Enabled = Not pbvalor
        
            If lnTipo = gTipoOpeRegistro Then
                Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbvalor
                Me.cmdExamenImprimir(lnIndiceActual).Enabled = False
            ElseIf lnTipo = gTipoOpeMantenimiento Then
                Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbvalor
                Me.cmdExamenImprimir(lnIndiceActual).Enabled = pbvalor
                Me.cmdExamenImprimir(lnIndiceActual).Enabled = False
            ElseIf lnTipo = gTipoOpeConsulta Then
                Me.cmdExamenGrabar(lnIndiceActual).Visible = False
                Me.cmdExamenImprimir(lnIndiceActual).Visible = False
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
                Me.cmdExamenCancelar(lnIndiceActual).Visible = False
            ElseIf lnTipo = gTipoOpeReporte Then
                Me.cmdExamenGrabar(lnIndiceActual).Visible = False
                Me.cmdExamenCancelar(lnIndiceActual).Visible = False
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
                Me.cmdExamenCancelar(lnIndiceActual).Visible = False
                
            End If
        End If
    End If
End Sub

Private Sub mskFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFF.SetFocus
    End If
End Sub

Private Sub REscrito_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show REscrito.Text, "Examen Escrito", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub RPsicologico_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show RPsicologico.Text, "Examen Psicologico", True, 66
    Set oPrevio = Nothing
End Sub
Private Sub TxtBuscarAreaCargo_EmiteDatos()
    Me.lblDescripcion.Caption = TxtBuscarAreaCargo.psDescripcion
End Sub

Private Sub TxtBuscarAreaCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.mskFI.SetFocus
End Sub

Private Sub txtComentarioFinal_GotFocus()
    txtComentarioFinal.SelStart = 0
    txtComentarioFinal.SelLength = 200
End Sub

Private Sub TxtDes_GotFocus()
    txtDes.SelStart = 0
    txtDes.SelLength = 50
End Sub

Private Sub TxtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbTipo.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub Limpia()
    Me.cmbEval.ListIndex = -1
    Me.TxtBuscarAreaCargo.Text = ""
    Me.cmbTipo.ListIndex = -1
    Me.cmbEstado.ListIndex = -1
    Me.txtDes.Text = ""
    Me.mskFI.Text = "__/__/____"
    Me.mskFF.Text = "__/__/____"
    Me.REscrito.Text = ""
    Me.RPsicologico.Text = ""
    Me.lblDescripcion.Caption = ""
    Me.cmbTipoContrato.ListIndex = -1
    Me.chkCur.value = 0
    Me.chkPsi.value = 0
    Me.chkEnt.value = 0
    Me.chkEsc.value = 0
    Me.txtNotaMax.Text = "100"
    Me.FlexComite.Clear
    Me.FlexComite.FormaCabecera
End Sub

Private Function Valida() As Boolean
    Dim lnI As Integer
    Dim oPar As DLogGeneral
    Set oPar = New DLogGeneral
    Dim lnMesesRango As Integer
    
    lnMesesRango = oPar.CargaParametro(6000, 1010)
    
    
    Me.Tab.Tab = 1
    
    For lnI = 1 To Me.FlexComite.Rows - 1
        Me.FlexComite.Row = lnI
        If Me.FlexComite.TextMatrix(lnI, 1) = "" And Me.FlexComite.TextMatrix(lnI, 0) <> "" Then
            MsgBox "Debe ingresar una persona al comite.", vbInformation, "Aviso"
            Me.FlexComite.Col = 1
            Valida = False
            FlexComite.SetFocus
            Exit Function
        ElseIf Me.FlexComite.TextMatrix(lnI, 3) = "" And Me.FlexComite.TextMatrix(lnI, 0) <> "" Then
            MsgBox "Debe ingresar un cargo a la persona del comite.", vbInformation, "Aviso"
            Me.FlexComite.Col = 3
            Valida = False
            
            Exit Function
        Else
            Valida = True
        End If
    Next lnI
    
    Me.Tab.Tab = 0
    
    If Me.cmbEstado.Text = "" Then
        MsgBox "Debe Elegir un Estado.", vbInformation, "Aviso"
        cmbEstado.SetFocus
        Valida = False
    ElseIf Me.txtDes.Text = "" Then
        MsgBox "Debe Ingresar una descripción.", vbInformation, "Aviso"
        txtDes.SetFocus
        Valida = False
    ElseIf Me.cmbTipo.Text = "" Then
        MsgBox "Debe Elegir un Tipo.", vbInformation, "Aviso"
        cmbTipo.SetFocus
        Valida = False
    ElseIf Me.TxtBuscarAreaCargo.Text = "" Then
        MsgBox "Debe Elegir un Cargo.", vbInformation, "Aviso"
        TxtBuscarAreaCargo.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Me.txtNotaMax.Text) Then
        MsgBox "Debe Elegir un numero Valido.", vbInformation, "Aviso"
        txtNotaMax.SetFocus
        Valida = False
    ElseIf Me.cmbTipoContrato.Text = "" Then
        MsgBox "Debe Ingresar un tipo de Contrato.", vbInformation, "Aviso"
        cmbTipoContrato.SetFocus
        Valida = False
    ElseIf Not IsDate(Me.mskFI.Text) Then
        MsgBox "Debe Ingresar una Fecha Valida.", vbInformation, "Aviso"
        mskFI.SetFocus
        Valida = False
    ElseIf Not IsDate(Me.mskFF.Text) Then
        MsgBox "Debe Ingresar una Fecha Valida.", vbInformation, "Aviso"
        mskFF.SetFocus
        Valida = False
    ElseIf CDate(Me.mskFI.Text) < DateAdd("m", lnMesesRango * -1, gdFecSis) Or CDate(Me.mskFI.Text) > DateAdd("m", lnMesesRango, gdFecSis) Then
        MsgBox "Debe Ingresar una Fecha Valida. en el rango " & Format(DateAdd("m", lnMesesRango * -1, gdFecSis), gsFormatoFechaView) & " - " & Format(DateAdd("m", lnMesesRango, gdFecSis), gsFormatoFechaView), vbInformation, "Aviso"
        mskFI.SetFocus
        Valida = False
    ElseIf CDate(Me.mskFF.Text) < DateAdd("m", lnMesesRango * -1, gdFecSis) Or CDate(Me.mskFF.Text) > DateAdd("m", lnMesesRango, gdFecSis) Then
        MsgBox "Debe Ingresar una Fecha Valida. en el rango " & Format(DateAdd("m", lnMesesRango * -1, gdFecSis), gsFormatoFechaView) & " - " & Format(DateAdd("m", lnMesesRango, gdFecSis), gsFormatoFechaView), vbInformation, "Aviso"
        mskFF.SetFocus
        Valida = False
    ElseIf CDate(Me.mskFF.Text) <= CDate(Me.mskFI.Text) Then
        MsgBox "La fecha final no puede ser menor o igual que la fecha inicial.", vbInformation, "Aviso"
        mskFF.SetFocus
        Valida = False
    'ElseIf Me.REscrito.Text = "" And lnTipo = gTipoOpeMantenimiento Then
    '    MsgBox "Debe Ingresar el texto del examen escrito.", vbInformation, "Aviso"
    '    Me.Tab.Tab = 2
    '    Me.cmdCargarE.SetFocus
    '    Valida = False
    'ElseIf Me.RPsicologico.Text = "" And lnTipo = gTipoOpeMantenimiento Then
    '    MsgBox "Debe Ingresar el texto del examen Psocilogico.", vbInformation, "Aviso"
    '    Me.Tab.Tab = 2
    '    Me.cmdCargarP.SetFocus
    '    Valida = False
    Else
        Valida = True
    End If
End Function

Private Sub CargaComite()
    Dim oEva As DActualizaProcesoSeleccion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oEva = New DActualizaProcesoSeleccion
    Set rs = oEva.GetNomPersonasComite(Left(Me.cmbEval.Text, 6))
    If Not (rs.EOF And rs.BOF) Then
        Set Me.FlexComite.Recordset = rs
    Else
        FlexComite.Clear
        FlexComite.FormaCabecera
    End If
    
    Set oEva = Nothing
    rs.Close
    Set rs = Nothing
End Sub

Private Sub IniTab()
    Dim i As Integer
    
    For i = 0 To Me.Tab.Tabs - 1
        Me.Tab.TabVisible(i) = False
    Next i
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuSel Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.Tab = 0
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuPost Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(3) = True
        Me.Tab.Tab = 3
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaCur Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(4) = True
        Me.Tab.Tab = 4
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEsc Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(5) = True
        Me.Tab.Tab = 5
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaPsi Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(6) = True
        Me.Tab.Tab = 6
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEnt Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(7) = True
        Me.Tab.Tab = 7
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(3) = True
        Me.Tab.TabVisible(4) = True
        Me.Tab.TabVisible(5) = True
        Me.Tab.TabVisible(6) = True
        Me.Tab.TabVisible(7) = True
        Me.Tab.TabVisible(8) = True
        Me.Tab.Tab = 8
    End If
End Sub

Public Sub IniCerrar()
    lbCerrar = True
    lnTipo = gTipoOpeConsulta
    Me.Show 1
End Sub

'****************************************************
Private Sub cmdPostulantesCancelar_Click()
    CargaDatosPostulantes
    Activa False, lnTipo
End Sub

Private Sub cmdPostulantesEditar_Click()
    If Me.cmbEval.Text = "" Then Exit Sub
    Me.fraPostulantesSeleccion.Enabled = True
    Activa True, gTipoOpeRegistro
    If cmdPostulantesNuevo.Enabled Then
        Me.cmdPostulantesNuevo.SetFocus
    ElseIf Me.cmdPostulantesEliminar.Enabled Then
        Me.cmdPostulantesEliminar.SetFocus
    End If
End Sub

Private Sub cmdPostulantesEliminar_Click()
    Me.FlexPostulantes.EliminaFila CLng(FlexPostulantes.TextMatrix(FlexPostulantes.Row, 0))
End Sub

Private Sub cmdPostulantesGrabar_Click()
    Dim oCurDet As DActualizaProcesoSeleccion
    Dim i As Integer
    Set oCurDet = New DActualizaProcesoSeleccion

    If MsgBox("Desea Grabar la Información ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub

    oCurDet.EliminaPersonasProSelec Left(Me.cmbEval.Text, 6), Me.FlexPostulantes.GetRsNew
    For i = 1 To Me.FlexPostulantes.Rows - 1
        If Me.FlexPostulantes.TextMatrix(i, 1) <> "" Then
            oCurDet.AgregaPersonaProSelec Left(Me.cmbEval.Text, 6), Me.FlexPostulantes.TextMatrix(i, 1), GetMovNro(gsCodUser, gsCodAge)
        End If
    Next i
    'ALPA 20090122 **********************************************************
    If Right(frmRHSeleccion.Caption, 13) = "MANTENIMIENTO" Then
        gsOpeCod = LogPistaModificaPostulante
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, "2", , Left(Me.cmbEval.Text, 6), gNumeroProcesoSeleccion
    Else
        gsOpeCod = LogPistaRegistraPostulante
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, "1", , Left(Me.cmbEval.Text, 6), gNumeroProcesoSeleccion
    End If
    '************************************************************************
    Set oCurDet = Nothing
    cmdPostulantesCancelar_Click
End Sub

Private Sub cmdPostulantesNuevo_Click()
    If Me.cmbEval.Text = "" Then Exit Sub

    If Me.FlexPostulantes.TextMatrix(FlexPostulantes.Rows - 1, 0) = "" Then
        FlexPostulantes.AdicionaFila 1
    Else
        FlexPostulantes.AdicionaFila CLng(Me.FlexPostulantes.TextMatrix(FlexPostulantes.Rows - 1, 0)) + 1
    End If
    Me.FlexPostulantes.SetFocus
End Sub

Private Sub cmdPostulantesImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oPrevio = New Previo.clsPrevio
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion

    lsCadena = oEva.GetReporteEvaPersonas(Left(Me.cmbEval.Text, 6), gsNomAge, gsEmpresa, gdFecSis)
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66
    Set oPrevio = Nothing
    Set oEva = Nothing
End Sub
Private Sub CargaDatosPostulantes()
    Dim oCurDet As DActualizaProcesoSeleccion
    Set oCurDet = New DActualizaProcesoSeleccion
    Dim rsEva As ADODB.Recordset

    If Me.cmbEval.Text = "" Then
        Me.cmdPostulantesNuevo.Enabled = False
        Me.cmdPostulantesEliminar.Enabled = False
        Exit Sub
    End If

    Set rsEva = New ADODB.Recordset
    Set rsEva = oCurDet.GetProcesosSeleccionDet(Left(Me.cmbEval.Text, 6))

    If Not (rsEva.BOF And rsEva.EOF) Then
        Me.FlexPostulantes.rsFlex = rsEva
    Else
        Me.FlexPostulantes.Clear
        Me.FlexPostulantes.Rows = 2
        Me.FlexPostulantes.FormaCabecera
    End If

    Set oCurDet = Nothing
End Sub

Private Sub cmdExamenCancelar_Click(Index As Integer)
    Activa False, lnTipo
    CargaDatosNotas
    Form_Load
    
    Me.FlexExamen(lnIndiceActual).Clear
    Me.FlexExamen(lnIndiceActual).Rows = 2
    Me.FlexExamen(lnIndiceActual).FormaCabecera
    Me.txtComentarioFinal.Text = ""
End Sub

Private Sub cmdExamenGrabar_Click(Index As Integer)
    Dim oCurDet As NActualizaProcesoSeleccion
    Dim i As Integer
    Dim lsUltMov As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oCurDet = New NActualizaProcesoSeleccion

    If MsgBox("Desea Grabar la Información ??? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub

    Set rs = Me.FlexExamen(lnIndiceActual).GetRsNew
    glsMovNro = ""
    If lnTipEva = RHTipoOpeEvaConsolidado Then
        lsUltMov = GetMovNro(gsCodUser, gsCodAge)
        'ALPA 20090122*************************************
        glsMovNro = lsUltMov
        '**************************************************
        oCurDet.ModificaPersonaProSelec Left(Me.cmbEval.Text, 6), rs, CInt(lnTipEva), lsUltMov
        oCurDet.ModificaProSelecComentarioFinal Left(Me.cmbEval.Text, 6), Me.txtComentarioFinal.Text, lsUltMov
        gsOpeCod = LogPistaRegistraCierreProcesoSelección
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.txtComentarioFinal.Text, Left(Me.cmbEval.Text, 6), gNumeroProcesoSeleccion
    Else
        'ALPA 20090122*************************************
        glsMovNro = GetMovNro(gsCodUser, gsCodAge)
        '**************************************************
        oCurDet.ModificaPersonaProSelec Left(Me.cmbEval.Text, 6), rs, CInt(lnTipEva), glsMovNro
    'End If
        'ALPA 20090122 **********************************************************
        If gTipoOpeMantenimiento = lnTipo Then
            If lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaCur Then
                gsOpeCod = LogPistaModificaExamenCurricular
            ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEsc Then
                gsOpeCod = LogPistaModificaExamenEscrito
            ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaPsi Then
                gsOpeCod = LogPistaModificaPsicologico
            ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEnt Then
                gsOpeCod = LogPistaModificaExamenEntrevista
            End If
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", , Left(Me.cmbEval.Text, 6), gNumeroProcesoSeleccion
        ElseIf gTipoOpeRegistro = lnTipo Then
            If lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaCur Then
                gsOpeCod = LogPistaRegistraExamenCurricular
            ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEsc Then
                gsOpeCod = LogPistaRegistraExamenEscrito
            ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaPsi Then
                gsOpeCod = LogPistaRegistraPsicologico
            ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEnt Then
                gsOpeCod = LogPistaRegistraExamenEntrevista
            End If
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", , Left(Me.cmbEval.Text, 6), gNumeroProcesoSeleccion
        End If
    End If
    '*************************************************************************
    Set oCurDet = Nothing
    cmdExamenCancelar_Click 1
End Sub

Private Sub cmdExamenImprimir_Click(Index As Integer)
    If Me.cmbEval.Text = "" Then Exit Sub
    Dim oEva As NActualizaProcesoSeleccion
    Dim oPrevio As Previo.clsPrevio
    Dim lsCadena As String
    Set oPrevio = New Previo.clsPrevio
    Set oEva = New NActualizaProcesoSeleccion
    Dim lbRep(2) As Boolean
    
    If Me.cmbEval.Text = "" Then Exit Sub
    
    frmImpreRRHH.Ini "Resultados;Acta Proceso Selección;", Caption, lbRep, gdFecSis, gdFecSis, False
    
    If lbRep(1) Then
        lsCadena = lsCadena & oEva.GetReporteSelectPersonasNotas(Left(Me.cmbEval.Text, 6), gsNomAge, gsEmpresa, gdFecSis, CInt(lnTipEva))
    End If
    
    If lbRep(2) Then
        If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        lsCadena = lsCadena & oEva.GetActa(Left(Me.cmbEval.Text, 6), gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66

    Set oPrevio = Nothing
    Set oEva = Nothing
End Sub

Private Sub cmdExamenEditar_Click(Index As Integer)
    If Me.cmbEval.Text = "" Then Exit Sub
    If Me.FlexExamen(lnIndiceActual).TextMatrix(1, 0) = "" Then
        MsgBox "No Puede procesar, porque el proeceso de selecion no tiene postulantes.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If lnTipEva = RHTipoOpeEvaCurricular And Me.chkCur.value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnTipEva = RHTipoOpeEvaEntrevista And Me.chkEnt.value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnTipEva = RHTipoOpeEvaEscrito And Me.chkEsc.value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnTipEva = RHTipoOpeEvaPsicologico And Me.chkPsi.value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Activa True, lnTipo
    Me.FlexExamen(Index).SetFocus
End Sub

Private Sub FlexExamen_RowColChange(Index As Integer)
    If lnTipEva <> RHTipoOpeEvaConsolidado And lnTipo = gTipoOpeRegistro Then
        If Me.FlexExamen(Index).TextMatrix(FlexExamen(Index).Row, FlexExamen(Index).Cols - 1) = "0" Then
            Me.FlexExamen(Index).lbEditarFlex = False
        Else
            Me.FlexExamen(Index).lbEditarFlex = True
        End If
    End If
End Sub

Private Sub CargaDatosNotas()
    Dim oCurDet As DActualizaProcesoSeleccion
    Set oCurDet = New DActualizaProcesoSeleccion
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset

    Set rsEva = oCurDet.GetProcesosSeleccionDetExamen(Left(Me.cmbEval.Text, 6), CInt(lnTipEva))

    If Not (rsEva.BOF And rsEva.EOF) Then
        If lnTipEva = RHTipoOpeEvaConsolidado Then
            Me.FlexExamen(lnIndiceActual).EncabezadosAlineacion = "C-L-L-R-R-R-R-R-R"
            Me.FlexExamen(lnIndiceActual).EncabezadosAnchos = "300-0-3000-900-900-900-900-900-350"
            Me.FlexExamen(lnIndiceActual).EncabezadosNombres = "#-Codigo-Nombre-Curriculum-Escrito-Psicologico-Entrevista-Promedio-OK"
            Me.FlexExamen(lnIndiceActual).ColumnasAEditar = "X-X-X-X-X-X-X-X-8"
            Me.FlexExamen(lnIndiceActual).ListaControles = "0-0-0-0-0-0-0-0-4"
        End If
        
        Me.FlexExamen(lnIndiceActual).rsFlex = rsEva
        If lnTipEva = RHTipoOpeEvaConsolidado Then Exit Sub
        Me.FlexExamen(lnIndiceActual).EncabezadosAnchos = "300-1500-4800-1500-1"
    Else
        Me.FlexExamen(lnIndiceActual).Clear
        Me.FlexExamen(lnIndiceActual).Rows = 2
        Me.FlexExamen(lnIndiceActual).FormaCabecera
        If lnTipEva = RHTipoOpeEvaConsolidado Then Exit Sub
        Me.FlexExamen(lnIndiceActual).EncabezadosAnchos = "500-1500-4800-1500-1"
    End If

    Set oCurDet = Nothing
    Set rsEva = Nothing
End Sub

Private Sub txtNotaMax_GotFocus()
    txtNotaMax.SelStart = 0
    txtNotaMax.SelLength = 50
End Sub

Private Sub txtNotaMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       chkCur.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub
