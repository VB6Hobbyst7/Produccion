VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHEvaluacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmRHEvaluacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab 
      Height          =   4365
      Left            =   75
      TabIndex        =   3
      Top             =   420
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   7699
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmRHEvaluacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancelar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGrabar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatosEva"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdEliminar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdImprimir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdNuevo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Comite"
      TabPicture(1)   =   "frmRHEvaluacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraComite"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Area/Per.."
      TabPicture(2)   =   "frmRHEvaluacion.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraArea"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Texto"
      TabPicture(3)   =   "frmRHEvaluacion.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fratexto"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Curricular"
      TabPicture(4)   =   "frmRHEvaluacion.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdExamenEditar(0)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdExamenImprimir(0)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraExamenSeleccion(0)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdExamenGrabar(0)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdExamenCancelar(0)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Escrito"
      TabPicture(5)   =   "frmRHEvaluacion.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdExamenCancelar(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdExamenEditar(1)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdExamenGrabar(1)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "fraExamenSeleccion(1)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cmdExamenImprimir(1)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Psicologico"
      TabPicture(6)   =   "frmRHEvaluacion.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdExamenCancelar(2)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "cmdExamenEditar(2)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "cmdExamenGrabar(2)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "fraExamenSeleccion(2)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cmdExamenImprimir(2)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "Entrevista"
      TabPicture(7)   =   "frmRHEvaluacion.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdExamenCancelar(3)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "cmdExamenEditar(3)"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "cmdExamenGrabar(3)"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "fraExamenSeleccion(3)"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "cmdExamenImprimir(3)"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "Resultado"
      TabPicture(8)   =   "frmRHEvaluacion.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "cmdCerrar"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "fraExamenSeleccion(4)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "cmdExamenEditar(4)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "cmdExamenGrabar(4)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "cmdExamenImprimir(4)"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "cmdExamenCancelar(4)"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).ControlCount=   6
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   -71730
         TabIndex        =   75
         Top             =   3870
         Visible         =   0   'False
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
         Height          =   3465
         Index           =   4
         Left            =   -74895
         TabIndex        =   71
         Top             =   330
         Width           =   8940
         Begin VB.TextBox txtComentarioFinal 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   60
            MaxLength       =   200
            ScrollBars      =   3  'Both
            TabIndex        =   72
            Top             =   2940
            Width           =   8730
         End
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   2475
            Index           =   4
            Left            =   90
            TabIndex        =   73
            Top             =   255
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   4366
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
         Begin VB.Label lblComentarioFinal 
            Caption         =   "Comentario Final"
            Height          =   240
            Left            =   75
            TabIndex        =   74
            Top             =   2730
            Width           =   1800
         End
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   4
         Left            =   -73845
         TabIndex        =   69
         Top             =   3870
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   4
         Left            =   -74895
         TabIndex        =   68
         Top             =   3870
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   4
         Left            =   -72780
         TabIndex        =   67
         Top             =   3870
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   -73845
         TabIndex        =   66
         Top             =   3915
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   3
         Left            =   -73845
         TabIndex        =   65
         Top             =   3915
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   3
         Left            =   -74895
         TabIndex        =   64
         Top             =   3915
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
         Height          =   3465
         Index           =   3
         Left            =   -74895
         TabIndex        =   62
         Top             =   330
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3120
            Index           =   3
            Left            =   45
            TabIndex        =   63
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   5503
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
         Left            =   -72780
         TabIndex        =   61
         Top             =   3915
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   -73845
         TabIndex        =   60
         Top             =   3885
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   2
         Left            =   -73845
         TabIndex        =   59
         Top             =   3885
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   2
         Left            =   -74895
         TabIndex        =   58
         Top             =   3885
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
         Height          =   3465
         Index           =   2
         Left            =   -74895
         TabIndex        =   56
         Top             =   330
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3120
            Index           =   2
            Left            =   60
            TabIndex        =   57
            Top             =   270
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   5503
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
         Left            =   -72780
         TabIndex        =   55
         Top             =   3885
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   -73845
         TabIndex        =   54
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   1
         Left            =   -73845
         TabIndex        =   53
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   1
         Left            =   -74895
         TabIndex        =   52
         Top             =   3900
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
         Height          =   3465
         Index           =   1
         Left            =   -74895
         TabIndex        =   50
         Top             =   330
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3120
            Index           =   1
            Left            =   45
            TabIndex        =   51
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   5503
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
         Left            =   -72780
         TabIndex        =   49
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenEditar 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   0
         Left            =   -73860
         TabIndex        =   47
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "N&uevo"
         Height          =   375
         Left            =   105
         TabIndex        =   35
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   3360
         TabIndex        =   34
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2265
         TabIndex        =   33
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   1185
         TabIndex        =   32
         Top             =   3900
         Width           =   975
      End
      Begin VB.Frame fraArea 
         Caption         =   "Areas Personas"
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
         Height          =   3870
         Left            =   -74895
         TabIndex        =   24
         Top             =   330
         Width           =   8895
         Begin Sicmact.FlexEdit FlexAreasEmp 
            Height          =   3105
            Left            =   3825
            TabIndex        =   28
            Top             =   270
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   5477
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Codigo-Nombre-Exon...-Comentario-Bit-COD"
            EncabezadosAnchos=   "300-1200-2000-600-2000-0-0"
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
            ColumnasAEditar =   "X-X-X-3-4-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-4-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.CommandButton cmdEliminarA 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   1125
            TabIndex        =   27
            Top             =   3435
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarA 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   105
            TabIndex        =   26
            Top             =   3435
            Width           =   975
         End
         Begin Sicmact.FlexEdit FlexAreas 
            Height          =   3105
            Left            =   90
            TabIndex        =   25
            Top             =   270
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   5477
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Area-bit"
            EncabezadosAnchos=   "300-800-2500-0"
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
            ColumnasAEditar =   "X-1-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C"
            FormatosEdit    =   "0-0-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
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
         Height          =   3855
         Left            =   -74895
         TabIndex        =   20
         Top             =   330
         Width           =   8895
         Begin VB.CommandButton cmdNuevoComite 
            Caption         =   "N&uevo"
            Height          =   375
            Left            =   6735
            TabIndex        =   22
            Top             =   3420
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarComite 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   7770
            TabIndex        =   21
            Top             =   3405
            Width           =   975
         End
         Begin Sicmact.FlexEdit FlexComite 
            Height          =   3120
            Left            =   90
            TabIndex        =   23
            Top             =   240
            Width           =   8670
            _ExtentX        =   15293
            _ExtentY        =   5503
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
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
            Appearance      =   0
            RowHeight0      =   240
         End
      End
      Begin VB.Frame fraDatosEva 
         Caption         =   "Evaluación"
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
         Height          =   3510
         Left            =   105
         TabIndex        =   4
         Top             =   330
         Width           =   8880
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
            TabIndex        =   38
            Top             =   240
            Width           =   8625
            Begin VB.ComboBox cmbEstado 
               Height          =   315
               Left            =   1155
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   255
               Width           =   7350
            End
            Begin VB.Label lblEstado 
               Caption         =   "Estado :"
               Height          =   255
               Left            =   135
               TabIndex        =   40
               Top             =   285
               Width           =   735
            End
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
            Left            =   120
            TabIndex        =   16
            Top             =   2805
            Width           =   8640
            Begin MSMask.MaskEdBox mskFF 
               Height          =   315
               Left            =   5040
               TabIndex        =   14
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
            Begin VB.Label lblFI 
               Caption         =   "Inicio :"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   255
               Width           =   855
            End
            Begin VB.Label lblFF 
               Caption         =   "Fin :"
               Height          =   255
               Left            =   4320
               TabIndex        =   17
               Top             =   255
               Width           =   735
            End
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
            Height          =   600
            Left            =   2565
            TabIndex        =   15
            Top             =   1995
            Width           =   6195
            Begin VB.CheckBox chkCur 
               Caption         =   "Exa Curricular"
               Height          =   195
               Left            =   300
               TabIndex        =   9
               Top             =   285
               Width           =   1425
            End
            Begin VB.CheckBox chkEsc 
               Caption         =   "Exa Escrito"
               Height          =   195
               Left            =   1785
               TabIndex        =   10
               Top             =   285
               Width           =   1425
            End
            Begin VB.CheckBox chkPsi 
               Caption         =   "Exa Psiologico"
               Height          =   195
               Left            =   3255
               TabIndex        =   11
               Top             =   300
               Width           =   1425
            End
            Begin VB.CheckBox chkEnt 
               Caption         =   "Exa Entrevista"
               Height          =   195
               Left            =   4695
               TabIndex        =   12
               Top             =   300
               Width           =   1365
            End
         End
         Begin VB.TextBox txtNotaMax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1215
            MaxLength       =   8
            TabIndex        =   8
            Text            =   "100"
            Top             =   2160
            Width           =   1230
         End
         Begin VB.TextBox txtDes 
            Appearance      =   0  'Flat
            Height          =   765
            Left            =   1185
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1080
            Width           =   7500
         End
         Begin VB.Label lblNotMaxima 
            Caption         =   "Nota Maxima :"
            Height          =   255
            Left            =   135
            TabIndex        =   19
            Top             =   2175
            Width           =   1050
         End
         Begin VB.Label lblDescripcion 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   2610
            TabIndex        =   7
            Top             =   780
            Width           =   4875
         End
         Begin VB.Label lblDes 
            Caption         =   "Descripción :"
            Height          =   255
            Left            =   45
            TabIndex        =   6
            Top             =   1065
            Width           =   960
         End
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
         Height          =   3870
         Left            =   -74895
         TabIndex        =   29
         Top             =   330
         Width           =   8910
         Begin VB.CommandButton cmdCargarP 
            Caption         =   "&Cargar Examen Psicologico"
            Height          =   375
            Left            =   4695
            TabIndex        =   41
            Top             =   3420
            Width           =   4095
         End
         Begin VB.CommandButton cmdCargarE 
            Caption         =   "&Cargar Examen Escrito"
            Height          =   375
            Left            =   180
            TabIndex        =   30
            Top             =   3420
            Width           =   4095
         End
         Begin RichTextLib.RichTextBox REscrito 
            Height          =   3090
            Left            =   180
            TabIndex        =   31
            Top             =   255
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   5450
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmRHEvaluacion.frx":0406
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
            Height          =   3090
            Left            =   4665
            TabIndex        =   42
            Top             =   255
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   5450
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmRHEvaluacion.frx":047F
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
         Left            =   105
         TabIndex        =   36
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1185
         TabIndex        =   37
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Index           =   0
         Left            =   -72795
         TabIndex        =   43
         Top             =   3900
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
         Height          =   3465
         Index           =   0
         Left            =   -74895
         TabIndex        =   44
         Top             =   330
         Width           =   8940
         Begin Sicmact.FlexEdit FlexExamen 
            Height          =   3120
            Index           =   0
            Left            =   45
            TabIndex        =   45
            Top             =   255
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   5503
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Codigo-Nombre-Nota"
            EncabezadosAnchos=   "500-1500-5000-1500"
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
      Begin VB.CommandButton cmdExamenGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   0
         Left            =   -74910
         TabIndex        =   46
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   -73860
         TabIndex        =   48
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton cmdExamenCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   -73845
         TabIndex        =   70
         Top             =   3870
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbEval 
      Height          =   315
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   15
      Width           =   6720
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8190
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   5955
      Top             =   6420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEval 
      Caption         =   "Evaluación :"
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   45
      Width           =   975
   End
End
Attribute VB_Name = "frmRHEvaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnTipo As TipoOpe
Dim lsCodigo As String
Dim lbEditado As Boolean
Dim loVar As RHProcesoSeleccionTipo
Dim lnTipEva  As RHTipoOpeEvaluacion
Dim lbCerrar As Boolean
Dim lnTipoSeleccionMnu As RHPSeleccionTpoMnu
Dim lnIndiceActual As Integer

Private Sub CargaDatos(pbTodos As Boolean, Optional pbSeleccion As Boolean = True)
    Dim rsE As ADODB.Recordset
    Dim oDatoAreas As DActualizaDatosArea
    Set oDatoAreas = New DActualizaDatosArea
    
    Set rsE = New ADODB.Recordset
    
    Me.FlexAreas.rsTextBuscar = oDatoAreas.GetAreas
    
    If pbSeleccion Then
        Dim oEva As DActualizaProcesoSeleccion
        Set oEva = New DActualizaProcesoSeleccion
        
        If lnTipo = gTipoOpeConsulta And lnTipEva = RHTipoOpeEvaConsolidado Then
            Set rsE = oEva.GetProcesosEvaluacion()
        Else
            Set rsE = oEva.GetProcesosEvaluacion(lbCerrar)
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
        Set rsE = oCons.GetConstante(gRHProcesoSeleccionEstado)
        CargaCombo rsE, Me.cmbEstado, 200
        rsE.Close
        Set oCons = Nothing
    End If
End Sub

Private Sub chkCur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkEsc.SetFocus
    End If
End Sub

Private Sub chkEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFI.SetFocus
    End If
End Sub

Private Sub chkEsc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkPsi.SetFocus
    End If
End Sub

Private Sub chkPsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkEnt.SetFocus
    End If
End Sub

Private Sub cmbEval_Click()
    Dim oEva As DActualizaProcesoSeleccion
    Dim rsE As New ADODB.Recordset
    Dim i As Integer
    Dim J As Integer
    Set oEva = New DActualizaProcesoSeleccion
    
    Set rsE = oEva.GetProcesoEvaluacion(Trim(Left(Me.cmbEval.Text, 6)))
    
    If Not (rsE.EOF And rsE.BOF) Then
        Me.txtDes.Text = rsE!Comentario
        UbicaCombo Me.cmbEstado, rsE!Estado
        Me.mskFI.Text = Format(rsE!fi, gsFormatoFechaView)
        Me.mskFF.Text = Format(rsE!ff, gsFormatoFechaView)
        Me.REscrito.Text = rsE!escrito
        Me.RPsicologico.Text = rsE!Psico
        Me.txtNotaMax.Text = IIf(IsNull(rsE!nMax), "0", rsE!nMax)
        Me.txtComentarioFinal.Text = IIf(IsNull(rsE!Com), "", rsE!Com)
        
        If Not IsNull(rsE!bCur) Then
            Me.chkCur.Value = IIf(rsE!bCur, 1, 0)
        Else
            Me.chkCur.Value = 0
        End If
        
        If Not IsNull(rsE!bEnt) Then
            Me.chkEnt.Value = IIf(rsE!bEnt, 1, 0)
        Else
            Me.chkEnt.Value = 0
        End If
        
        If Not IsNull(rsE!bEsc) Then
            Me.chkEsc.Value = IIf(rsE!bEsc, 1, 0)
        Else
            Me.chkEsc.Value = 0
        End If
        
        If Not IsNull(rsE!bPsico) Then
            Me.chkPsi.Value = IIf(rsE!bPsico, 1, 0)
        Else
            Me.chkPsi.Value = 0
        End If
        
        CargaComite
        
        Set rsE = oEva.GetAreasEvaluacion(Trim(Left(Me.cmbEval.Text, 6)))
        If rsE.EOF And rsE.BOF Then
            FlexAreas.Clear
            Me.FlexAreas.Rows = 2
            FlexAreas.FormaCabecera
        Else
            Set Me.FlexAreas.Recordset = rsE
            Set rsE = oEva.GetAreasEvaluacionRRHH(Trim(Left(Me.cmbEval.Text, 6)))
            If rsE.EOF And rsE.BOF Then
                FlexAreasEmp.Clear
                Me.FlexAreasEmp.Rows = 2
                FlexAreasEmp.FormaCabecera
            Else
                Set Me.FlexAreasEmp.Recordset = rsE
                For J = 1 To FlexAreas.Rows - 1
                    For i = 1 To FlexAreasEmp.Rows - 1
                        If FlexAreasEmp.TextMatrix(i, 6) = FlexAreas.TextMatrix(J, 1) Then
                            FlexAreasEmp.TextMatrix(i, 5) = FlexAreas.TextMatrix(J, 0)
                        End If
                    Next i
                    FlexAreas.TextMatrix(J, 3) = FlexAreas.TextMatrix(J, 0)
                Next J
            End If
        End If
    Else
        Me.FlexComite.Rows = 1
        Me.FlexComite.Rows = 2
        Me.FlexComite.FixedRows = 1
    End If
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaCur Or lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEsc Or lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaPsi Or lnTipoSeleccionMnu = RHPSeleccionTpoMnuEvaEnt Then
        CargaDatosNotas
    ElseIf lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
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

Private Sub cmdAgregarA_Click()
    FlexAreas.ColumnasAEditar = "X-1-X"
    Me.FlexAreas.AdicionaFila
    Me.FlexAreas.SetFocus
End Sub

Private Sub cmdCerrar_Click()
    Dim oEval As DActualizaProcesoSeleccion
    Set oEval = New DActualizaProcesoSeleccion
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    If Me.cmbEval.Text = "" Then Exit Sub
    If Not oEval.ValidaNotasNulasEva(Left(Me.cmbEval.Text, 6)) Then
        lbEditado = True
        If MsgBox("Desea Cerrar proceso Seleccion ?,  - El Procesa no podra se modificado.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        UbicaCombo Me.cmbEstado, RHProcesoSeleccionEstado.gRHProcSelEstFinalizado
        oEva.ModificaEstadoEvaluacion Left(Me.cmbEval, 6), Right(Me.cmbEstado.Text, 1), GetMovNro(gsCodUser, gsCodAge)
        CmdCancelar_Click
    Else
        MsgBox "Debe de Ingresar todas la notas para poder cerrar un proceso de Evaluación Interna.", vbInformation, "Aviso"
    End If
    
    Set oEva = Nothing
    Set oEval = Nothing
End Sub

Private Sub CmdCancelar_Click()
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
End Sub

Private Sub cmdCargarP_Click()
    CDialog.CancelError = False
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos txt(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Me.RPsicologico.LoadFile CDialog.FileName, 1
End Sub

Private Sub cmdEditar_Click()
    If lnTipo = gTipoOpeRegistro Then
        Exit Sub
    End If
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

Private Sub cmdEliminar_Click()
    Dim oEva As NActualizaProcesoSeleccion
    Set oEva = New NActualizaProcesoSeleccion
    
    If MsgBox("Desea Elimiar la Evaluación. " & Trim(Left(Me.cmbEval.Text, 50)) & Chr(13) & "Se eliminaran todas las personas relacionadas con esta Evaluacion.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    oEva.EliminaProSelec Trim(Right(Me.cmbEval.List(cmbEval.ListCount - 1), 7))
    CargaDatos False
    Limpia
End Sub

Private Sub cmdEliminarA_Click()
    Dim i As Integer
    Dim lnPosAnt As Integer
    Dim lnTotal As Integer
    
    lnTotal = Me.FlexAreasEmp.Rows
    lnPosAnt = -1
    FlexAreasEmp.Row = 1
    For i = 1 To lnTotal
        If Me.FlexAreasEmp.TextMatrix(FlexAreasEmp.Row, 5) = FlexAreas.TextMatrix(FlexAreas.Row, 3) Then
            FlexAreasEmp.EliminaFila FlexAreasEmp.Row
        Else
            If FlexAreasEmp.Row < FlexAreasEmp.Rows - 1 Then FlexAreasEmp.Row = FlexAreasEmp.Row + 1
        End If
    Next i
    
    Me.FlexAreas.EliminaFila FlexAreas.Row
End Sub

Private Sub cmdEliminarComite_Click()
    Me.FlexComite.EliminaFila Me.FlexComite.Row
End Sub

Private Sub cmdGrabar_Click()
    Dim oEval As NActualizaProcesoSeleccion
    Set oEval = New NActualizaProcesoSeleccion
    Dim lsUltActualizacion As String
    
    If Not Valida() Then Exit Sub
        
    lsUltActualizacion = GetMovNro(gsCodUser, gsCodAge)
    If lbEditado Then
        oEval.ModificaProEval Left(Me.cmbEval, 6), Format(Me.mskFI.Text, gsFormatoFecha), Format(Me.mskFF.Text, gsFormatoFecha), Right(cmbEstado, 1), Me.txtDes.Text, Me.RPsicologico.Text, Me.REscrito.Text, lsUltActualizacion, FlexAreas.GetRsNew, FlexAreasEmp.GetRsNew, Me.txtNotaMax.Text, Str(Me.chkCur.Value), Str(Me.chkEsc.Value), Str(Me.chkPsi.Value), Str(Me.chkEnt.Value)
        oEval.AgregaComiteProEval Left(Me.cmbEval, 6), Me.FlexComite.GetRsNew, lsUltActualizacion
    Else
        oEval.AgregaProEval lsCodigo, Format(Me.mskFI.Text, gsFormatoFecha), Format(Me.mskFF.Text, gsFormatoFecha), Right(cmbEstado, 1), Me.txtDes.Text, Me.RPsicologico.Text, Me.REscrito.Text, lsUltActualizacion, Me.FlexAreas.GetRsNew, FlexAreasEmp.GetRsNew, Me.txtNotaMax.Text, Str(Me.chkCur.Value), Str(Me.chkEsc.Value), Str(Me.chkPsi.Value), Str(Me.chkEnt.Value)
        oEval.AgregaComiteProEval lsCodigo, Me.FlexComite.GetRsNew, lsUltActualizacion
    End If
        
    Limpia
    Activa False, lnTipo
    CargaDatos False
End Sub

Private Sub cmdImprimir_Click()
    Dim oEval As NActualizaProcesoSeleccion
    Dim lsCadena As String
    Dim lsCadenaTemp As String
    Dim lbRep(2) As Boolean
    Dim oPrevio As Previo.clsPrevio
    Set oEval = New NActualizaProcesoSeleccion
    Set oPrevio = New Previo.clsPrevio
    
    frmImpreRRHH.Ini "Lista de Examenes;Acta de Ingresos;", "Evaluaicon", lbRep, gdFecSis, gdFecSis, False
    
    If lbRep(1) Then
        lsCadena = lsCadena & oEval.GetReporte(gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    If lbRep(2) And Me.cmbEval.Text <> "" Then
        If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        
        lsCadena = lsCadena & oEval.GetActa(Left(Me.cmbEval.Text, 6), gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    If lsCadena <> "" Then oPrevio.Show lsCadena, " Evaluaciones ", True, 66
    Set oEval = Nothing
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
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


Private Sub FlexAreas_OnCellChange(pnRow As Long, pnCol As Long)
    Dim oEva As DActualizaDatosRRHH
    Set oEva = New DActualizaDatosRRHH
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If FlexAreas.TextMatrix(pnRow, pnCol) <> "" Then
        FlexAreas.TextMatrix(pnRow, 3) = FlexAreas.TextMatrix(pnRow, 0)
        Set rs = oEva.GetRRHHArea(FlexAreas.TextMatrix(pnRow, pnCol))
        If Not (rs.EOF And rs.BOF) Then
            While Not rs.EOF
                Me.FlexAreasEmp.AdicionaFila
                FlexAreasEmp.TextMatrix(FlexAreasEmp.Row, 1) = rs.Fields(0)
                FlexAreasEmp.TextMatrix(FlexAreasEmp.Row, 2) = rs.Fields(1)
                FlexAreasEmp.TextMatrix(FlexAreasEmp.Row, 5) = pnRow
                FlexAreasEmp.TextMatrix(FlexAreasEmp.Row, 6) = FlexAreas.TextMatrix(pnRow, pnCol)
                rs.MoveNext
            Wend
        End If
    End If
    Set oEva = Nothing
End Sub

Private Sub FlexAreas_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    pbEsDuplicado = True
End Sub

Private Sub FlexAreas_RowColChange()
    If FlexAreas.TextMatrix(FlexAreas.Row, 1) = "" Then
        FlexAreas.ColumnasAEditar = "X-1-X"
    Else
        FlexAreas.ColumnasAEditar = "X-X-X"
    End If
End Sub

Private Sub FlexExamen_RowColChange(Index As Integer)
    If lnTipEva <> RHTipoOpeEvaConsolidado And (lnTipo = gTipoOpeRegistro Or lnTipo = gTipoOpeMantenimiento) Then
        If lnTipo = gTipoOpeRegistro Then
            If Me.FlexExamen(Index).TextMatrix(FlexExamen(Index).Row, FlexExamen(Index).Cols - 1) = "1" Then
                Me.FlexExamen(Index).lbEditarFlex = True
            Else
                Me.FlexExamen(Index).lbEditarFlex = False
            End If
        Else
            If Me.FlexExamen(Index).TextMatrix(FlexExamen(Index).Row, FlexExamen(Index).Cols - 1) <> "0" Then
                Me.FlexExamen(Index).lbEditarFlex = True
            Else
                Me.FlexExamen(Index).lbEditarFlex = False
            End If
        End If
    End If
End Sub

Private Sub FlexExamen_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If Not IsNumeric(Me.FlexExamen(lnIndiceActual).TextMatrix(pnRow, pnCol)) Then Exit Sub
    If CCur(Me.FlexExamen(lnIndiceActual).TextMatrix(pnRow, pnCol)) <= CCur(Me.txtNotaMax.Text) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub Form_Load()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oCons As DConstantes
    Set oCons = New DConstantes
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuSel Then
        If lnTipo = gTipoOpeRegistro Then
            CargaDatos True, False
        Else
            CargaDatos True
        End If
    Else
        CargaDatos True
    End If
    
    Set rsE = oCons.GetConstante(gRHEvaluacionComite)
    Me.FlexComite.CargaCombo rsE
    
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

Private Sub Activa(pbValor As Boolean, pnTipo As TipoOpe)
    'Objetos
    Dim lnI As Integer
    Me.cmbEval.Enabled = Not pbValor
    
    If lnTipoSeleccionMnu = RHPSeleccionTpoMnuSel Then
        Me.cmdNuevo.Visible = Not pbValor
        Me.cmdEditar.Visible = Not pbValor
        Me.cmdGrabar.Visible = pbValor
        Me.CmdCancelar.Visible = pbValor
        Me.cmdSalir.Enabled = Not pbValor
        
        If pnTipo = gTipoOpeRegistro Then
            Me.CmdCancelar.Visible = pbValor
            Me.fraProp.Enabled = False
            Me.cmbEval.Enabled = False
            Me.cmdEliminar.Visible = False
            Me.cmdImprimir.Visible = False
            Me.fratexto.Enabled = False
            Me.cmdEditar.Visible = False
            
        ElseIf pnTipo = gTipoOpeMantenimiento Then
            Me.fraDatosEva.Enabled = pbValor
            Me.fraArea.Enabled = pbValor
            Me.fratexto.Enabled = pbValor
            Me.fraComite.Enabled = pbValor
            Me.fraProp.Enabled = pbValor
            Me.cmdEliminar.Enabled = Not pbValor
            Me.cmdImprimir.Visible = False
            Me.cmdNuevo.Visible = False
            Me.cmdGrabar.Enabled = pbValor
            Me.cmdGrabar.Visible = True
        ElseIf pnTipo = gTipoOpeConsulta Then
            Me.fraProp.Enabled = False
    
            Me.fraDatosEva.Enabled = pbValor
            fraFechas.Enabled = pbValor
    
            Me.cmdCargarE.Visible = pbValor
            Me.cmdCargarP.Visible = pbValor
            
            Me.cmdCerrar.Visible = lbCerrar
            Me.txtDes.Enabled = pbValor
            Me.cmdGrabar.Visible = False
            Me.cmdEditar.Visible = False
            Me.CmdCancelar.Visible = False
            Me.cmdNuevo.Visible = False
            Me.cmdImprimir.Visible = False
            Me.cmdEliminar.Visible = False
            Me.cmdNuevoComite.Visible = False
            Me.cmdEliminarComite.Visible = False
            Me.cmdAgregarA.Visible = False
            Me.cmdEliminarA.Visible = False
            Me.FlexAreas.lbEditarFlex = False
            Me.FlexAreasEmp.lbEditarFlex = False
            
        ElseIf pnTipo = gTipoOpeReporte Then
            Me.txtDes.Enabled = pbValor
            
            Me.cmdEliminar.Enabled = pbValor
            Me.cmdNuevo.Enabled = pbValor
            Me.cmdEditar.Enabled = pbValor
            Me.cmbEval.Enabled = True
            
            Me.fraProp.Enabled = pbValor
        End If
    Else
        Me.cmdGrabar.Visible = False
        Me.cmdEditar.Visible = False
        Me.CmdCancelar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdImprimir.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdNuevoComite.Visible = False
        Me.cmdEliminarComite.Visible = False
        Me.cmdAgregarA.Visible = False
        Me.cmdEliminarA.Visible = False

        Me.fraProp.Enabled = False
        Me.fraDatosEva.Enabled = pbValor
        fraFechas.Enabled = pbValor
        Me.cmdCargarE.Enabled = pbValor
        Me.cmdCargarP.Enabled = pbValor
        Me.cmdCerrar.Visible = False
        Me.txtDes.Enabled = pbValor
        Me.fraComite.Enabled = False
        Me.cmdSalir.Enabled = Not pbValor

        If lnTipoSeleccionMnu = RHPSeleccionTpoMnuResultado Then
            For lnI = 0 To 3
                Me.cmdExamenEditar(lnI).Visible = False
                Me.cmdExamenGrabar(lnI).Visible = False
                Me.cmdExamenCancelar(lnI).Visible = False
                Me.cmdExamenImprimir(lnI).Visible = False
                Me.fraExamenSeleccion(lnI).Visible = True
            Next lnI
        
            If lnTipo = gTipoOpeConsulta Then
                Me.cmdCerrar.Visible = False
                Me.cmdExamenImprimir(4).Visible = False
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
                Me.cmdExamenGrabar(lnIndiceActual).Visible = False
                Me.cmdExamenCancelar(lnIndiceActual).Visible = False
            Else
                Me.cmdExamenEditar(lnIndiceActual).Visible = Not pbValor
                Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbValor
                Me.cmdExamenCancelar(lnIndiceActual).Visible = pbValor
                Me.fraExamenSeleccion(lnIndiceActual).Enabled = pbValor
                
                Me.cmdSalir.Enabled = Not pbValor
                Me.cmdCerrar.Visible = True
                Me.cmdCerrar.Enabled = True
            End If
        Else
            Me.cmdExamenEditar(lnIndiceActual).Visible = Not pbValor
            Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbValor
            Me.cmdExamenCancelar(lnIndiceActual).Visible = pbValor
            Me.fraExamenSeleccion(lnIndiceActual).Enabled = pbValor
            Me.cmbEval.Enabled = Not pbValor
        
            If lnTipo = gTipoOpeRegistro Then
                Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbValor
                Me.cmdExamenImprimir(lnIndiceActual).Visible = False
            ElseIf lnTipo = gTipoOpeMantenimiento Then
                Me.cmdExamenGrabar(lnIndiceActual).Enabled = pbValor
                Me.cmdExamenImprimir(lnIndiceActual).Enabled = pbValor
                Me.cmdExamenImprimir(lnIndiceActual).Visible = False
            ElseIf lnTipo = gTipoOpeConsulta Then
                Me.cmdExamenGrabar(lnIndiceActual).Visible = False
                Me.cmdExamenImprimir(lnIndiceActual).Visible = False
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
                Me.fraExamenSeleccion(lnIndiceActual).Enabled = True
                Me.FlexExamen(lnIndiceActual).lbEditarFlex = False
            ElseIf lnTipo = gTipoOpeReporte Then
                Me.cmdExamenGrabar(lnIndiceActual).Visible = False
                Me.cmdExamenEditar(lnIndiceActual).Visible = False
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

Private Sub Tab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.Tab.Tab = 0 Then
            Me.cmdCargarE.SetFocus
        Else
            Me.cmdCargarP.SetFocus
        End If
    End If

End Sub

Private Sub TxtDes_GotFocus()
    txtDes.SelStart = 0
    txtDes.SelLength = 50
End Sub

Private Sub TxtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtNotaMax.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub Limpia()
    Me.Tab.Tab = 0
    Me.cmbEval.ListIndex = -1
    Me.cmbEstado.ListIndex = -1
    Me.txtDes.Text = ""
    Me.mskFI.Text = "__/__/____"
    Me.mskFF.Text = "__/__/____"
    Me.REscrito.Text = ""
    Me.RPsicologico.Text = ""
    Me.FlexComite.Clear
    Me.FlexComite.Rows = 2
    FlexComite.FormaCabecera
    Me.FlexAreas.Clear
    Me.FlexAreas.Rows = 2
    Me.FlexAreas.FormaCabecera
    Me.FlexAreasEmp.Clear
    Me.FlexAreasEmp.Rows = 2
    Me.FlexAreasEmp.FormaCabecera
    Me.chkCur.Value = 0
    Me.chkEsc.Value = 0
    Me.chkEnt.Value = 0
    Me.chkPsi.Value = 0
    Me.txtNotaMax = "100"
    Me.txtComentarioFinal.Text = ""
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
    ElseIf Me.REscrito.Text = "" And lnTipo = gTipoOpeMantenimiento Then
        MsgBox "Debe Ingresar el texto del examen escrito.", vbInformation, "Aviso"
        Me.Tab.Tab = 3
        Me.cmdCargarE.SetFocus
        Valida = False
    ElseIf Me.RPsicologico.Text = "" And lnTipo = gTipoOpeMantenimiento Then
        MsgBox "Debe Ingresar el texto del examen Psocilogico.", vbInformation, "Aviso"
        Me.Tab.Tab = 3
        Me.cmdCargarP.SetFocus
        Valida = False
    Else
        Valida = True
    End If
End Function

Private Sub CargaComite()
    Dim oEva As DActualizaProcesoSeleccion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oEva = New DActualizaProcesoSeleccion
    Set rs = oEva.GetNomPersonasComiteEval(Left(Me.cmbEval.Text, 6))
    If Not rs Is Nothing Then
        If Not (rs.EOF And rs.BOF) Then
            Set Me.FlexComite.Recordset = rs
        Else
            FlexComite.Clear
            FlexComite.Rows = 2
            FlexComite.FormaCabecera
        End If
        rs.Close
        Set rs = Nothing
    Else
        FlexComite.Clear
        FlexComite.Rows = 2
        FlexComite.FormaCabecera
    End If
    Set oEva = Nothing
End Sub

Public Sub IniCerrar()
    lbCerrar = True
    lnTipo = gTipoOpeConsulta
    Me.Show 1
End Sub

Private Sub txtNotaMax_GotFocus()
    txtNotaMax.SelStart = 0
    txtNotaMax.SelLength = 20
End Sub

Private Sub txtNotaMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkCur.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
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
        Me.Tab.TabVisible(3) = True
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

Private Sub cmdExamenCancelar_Click(Index As Integer)
    CargaDatosNotas
    Activa False, lnTipo
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
    
    If lnTipEva = RHTipoOpeEvaConsolidado Then
        lsUltMov = GetMovNro(gsCodUser, gsCodAge)
        'oCurDet.ModificaPersonaProEval Left(Me.cmbEval.Text, 6), rs, CInt(lnTipEva), lsUltMov
        oCurDet.ModificaProSelecComentarioFinalEval Left(Me.cmbEval.Text, 6), Me.txtComentarioFinal.Text, lsUltMov
    Else
        oCurDet.ModificaPersonaProEval Left(Me.cmbEval.Text, 6), rs, CInt(lnTipEva), GetMovNro(gsCodUser, gsCodAge)
    End If
    
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

    lsCadena = oEva.GetReporteEvaPersonasNotas(Left(Me.cmbEval.Text, 6), gsNomAge, gsEmpresa, gdFecSis, CInt(lnTipEva))
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
    
    If lnTipEva = RHTipoOpeEvaCurricular And Me.chkCur.Value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnTipEva = RHTipoOpeEvaEntrevista And Me.chkEnt.Value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnTipEva = RHTipoOpeEvaEscrito And Me.chkEsc.Value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnTipEva = RHTipoOpeEvaPsicologico And Me.chkPsi.Value = 0 Then
        MsgBox "No Puede procesar Agregar notas, porque el proceso de Seleccion no Cuenta con este tipo de evaluacion.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Activa True, lnTipo
    FlexExamen(Index).SetFocus
End Sub

Private Sub CargaDatosNotas()
    Dim oCurDet As DActualizaProcesoSeleccion
    Set oCurDet = New DActualizaProcesoSeleccion
    Dim rsEva As ADODB.Recordset
    Set rsEva = New ADODB.Recordset

    Set rsEva = oCurDet.GetProcesosEvaluacionDetExamen(Left(Me.cmbEval.Text, 6), CInt(lnTipEva))

    If Not (rsEva.BOF And rsEva.EOF) Then
        If lnTipEva = RHTipoOpeEvaConsolidado Then
            Me.FlexExamen(lnIndiceActual).EncabezadosAlineacion = "C-L-L-R-R-R-R-R-R"
            Me.FlexExamen(lnIndiceActual).EncabezadosAnchos = "300-1-3000-1000-1000-1000-1000-1000-1"
            Me.FlexExamen(lnIndiceActual).EncabezadosNombres = "#-Codigo-Nombre-Escrito-Psicologico-Entrevista-Curriculum-Promedio-ok"
            Me.FlexExamen(lnIndiceActual).ColumnasAEditar = "X-X-X-X-X-X-X-X-X"
            Me.FlexExamen(lnIndiceActual).ListaControles = "0-0-0-0-0-0-0-0-0"
        End If
        
        Me.FlexExamen(lnIndiceActual).rsFlex = rsEva
        If lnTipEva = RHTipoOpeEvaConsolidado Then Exit Sub
        Me.FlexExamen(lnIndiceActual).EncabezadosAnchos = "500-1500-4800-1500-1"
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

