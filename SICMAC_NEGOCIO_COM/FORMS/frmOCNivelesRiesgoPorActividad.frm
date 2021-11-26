VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOCNivelesRiesgoPorActividad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Niveles de Riesgo por Actividad"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "frmOCNivelesRiesgoPorActividad.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver Total "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   33
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   32
      Top             =   6120
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Nivel de Riesgo"
      TabPicture(0)   =   "frmOCNivelesRiesgoPorActividad.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraOcupacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraPeriodo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGenerar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGuardarPerfil"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCerrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9960
         TabIndex        =   13
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardarPerfil 
         Caption         =   "Guardar Perfil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9960
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame fraPeriodo 
         Caption         =   "Periodo"
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
         Left            =   7080
         TabIndex        =   7
         Top             =   480
         Width           =   2775
         Begin VB.TextBox txtAnio 
            Height          =   285
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame fraOcupacion 
         Caption         =   "Ocupación/CIIU"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   4815
         Begin VB.ComboBox cboCIIU 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame fraTipo 
         Caption         =   "Tipo"
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
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton optJuridica 
            Caption         =   "P. Juridica"
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optNatural 
            Caption         =   "P. Natural"
            Height          =   195
            Left            =   480
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   617
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Persona"
         TabPicture(0)   =   "frmOCNivelesRiesgoPorActividad.frx":0326
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Producto"
         TabPicture(1)   =   "frmOCNivelesRiesgoPorActividad.frx":0342
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Canales de Distribución"
         TabPicture(2)   =   "frmOCNivelesRiesgoPorActividad.frx":035E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame7"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Jurisdicción"
         TabPicture(3)   =   "frmOCNivelesRiesgoPorActividad.frx":037A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame8"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Comportamiento"
         TabPicture(4)   =   "frmOCNivelesRiesgoPorActividad.frx":0396
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame9"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Transacciones"
         TabPicture(5)   =   "frmOCNivelesRiesgoPorActividad.frx":03B2
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame10"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Resumen"
         TabPicture(6)   =   "frmOCNivelesRiesgoPorActividad.frx":03CE
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame4"
         Tab(6).ControlCount=   1
         Begin VB.Frame Frame10 
            Caption         =   "SubVariables"
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
            Height          =   3615
            Left            =   -73440
            TabIndex        =   29
            Top             =   600
            Width           =   7455
            Begin SICMACT.FlexEdit feTransacciones 
               Height          =   2775
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-SubVariables-Rango-Calificacion"
               EncabezadosAnchos=   "0-4000-1200-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "L-L-C-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "SubVariables"
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
            Height          =   3615
            Left            =   -73440
            TabIndex        =   27
            Top             =   600
            Width           =   7455
            Begin SICMACT.FlexEdit feComportamiento 
               Height          =   2775
               Left            =   240
               TabIndex        =   28
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-SubVariables-Rango-Calificacion"
               EncabezadosAnchos=   "0-4000-1200-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "SubVariables"
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
            Height          =   3615
            Left            =   -73440
            TabIndex        =   25
            Top             =   600
            Width           =   7455
            Begin SICMACT.FlexEdit feJurisdiccion 
               Height          =   2775
               Left            =   240
               TabIndex        =   26
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-SubVariables-Rango-Calificacion"
               EncabezadosAnchos=   "0-4000-1200-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "SubVariables"
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
            Height          =   3615
            Left            =   -73440
            TabIndex        =   23
            Top             =   600
            Width           =   7455
            Begin SICMACT.FlexEdit feCanales 
               Height          =   2775
               Left            =   240
               TabIndex        =   24
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-SubVariables-Rango-Calificacion"
               EncabezadosAnchos=   "0-4000-1200-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "SubVariables"
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
            Height          =   3615
            Left            =   -73440
            TabIndex        =   21
            Top             =   600
            Width           =   7455
            Begin SICMACT.FlexEdit feProducto 
               Height          =   2775
               Left            =   240
               TabIndex        =   22
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-SubVariables-Rango-Calificacion"
               EncabezadosAnchos=   "0-4000-1200-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "SubVariables"
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
            Height          =   3615
            Left            =   1560
            TabIndex        =   19
            Top             =   600
            Width           =   7455
            Begin SICMACT.FlexEdit fePersona 
               Height          =   2775
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4895
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-SubVariables-Rango-Calificacion"
               EncabezadosAnchos=   "0-4000-1200-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X-X"
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-C-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Variables"
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
            Height          =   3615
            Left            =   -73440
            TabIndex        =   14
            Top             =   600
            Width           =   7455
            Begin VB.TextBox txtCalificacion 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   2160
               TabIndex        =   17
               Text            =   "0.00"
               Top             =   3120
               Width           =   975
            End
            Begin SICMACT.FlexEdit feResumen 
               Height          =   2535
               Left            =   240
               TabIndex        =   15
               Top             =   360
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   4471
               Cols0           =   3
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Variables-Ponderado"
               EncabezadosAnchos=   "0-5000-1200"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-X-X"
               ListaControles  =   "0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-R"
               FormatosEdit    =   "0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.Label lblNivelRiesgo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   5760
               TabIndex        =   31
               Top             =   3120
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Nivel de Riesgo:"
               Height          =   255
               Left            =   4440
               TabIndex        =   18
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Resumen de Calificación:"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   3240
               Width           =   1935
            End
         End
      End
   End
End
Attribute VB_Name = "frmOCNivelesRiesgoPorActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'*** REQUERIMIENTO: TI-ERS106-2014 Y ANEXO-01
'*** USUARIO: FRHU
'*** FECHA CREACION: 17/09/2014
'********************************************
Option Explicit
Dim bFormatea As Boolean
Dim nTipoPersoneria As Integer, nMes As Integer
Dim cCodOcuOrCIIU As String, cAnio As String
Private Sub Form_Load()
    bFormatea = False
    Call CargaComboConstante(1010, cboMes)
    Call CargarOcupaciones
    Call CargarVariables
    Call CargarSubVariables
    cmdGuardarPerfil.Enabled = False
    cmdCancelar.Enabled = False
End Sub
Private Sub cmdGenerar_Click()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrGenerar
    'If MsgBox("Desea Generar con los parametros Ingresados", vbQuestion + vbYesNo, "AVISO") = vbNo Then Exit Sub 'FRHU 20141017 - Observacion
    If Not ValidaGenerar Then Exit Sub
    
    nTipoPersoneria = 1
    cCodOcuOrCIIU = Trim(Right(cboCIIU.Text, 5))
    If optJuridica.value = True Then nTipoPersoneria = 2
    nMes = CInt(Right(Trim(cboMes.Text), 2))
    cAnio = txtAnio.Text
    
    If oServ.ValidarNivelDeRiesgo(nTipoPersoneria, cCodOcuOrCIIU, nMes, cAnio) Then
        If MsgBox("El perfil con los parametros seleccionados ya fue generado. Desea generar nuevamente ?" & vbNewLine & _
                  "Si acepta, los datos anteriormente registrados con estos parametros se eliminaran.", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Call oServ.EliminarNivelDeRiesgo(nTipoPersoneria, cCodOcuOrCIIU, nMes, cAnio)
    Else
        If MsgBox("Desea Generar con los parametros Ingresados", vbQuestion + vbYesNo, "AVISO") = vbNo Then Exit Sub 'FRHU 20141017 - Observacion
    End If
    Screen.MousePointer = 11
    Set rs = oServ.obtenerPerfilDeRiesgoPorActividad(nTipoPersoneria, cCodOcuOrCIIU, nMes, cAnio)
    Screen.MousePointer = 0
    
    Call CargarValoresVariables(rs)
    fraTipo.Enabled = False
    fraPeriodo.Enabled = False
    fraOcupacion.Enabled = False
    cmdGenerar.Enabled = False
    cmdGuardarPerfil.Enabled = True
    cmdCancelar.Enabled = True
    Exit Sub
ErrGenerar:
    MsgBox "Ha ocurrido un error al procesar, vuelvo a intentarlo," & Chr(10) & " si persiste comuniquese con el Dpto de TI", vbCritical, "Aviso"
    Screen.MousePointer = 0
End Sub
Private Sub cmdCancelar_Click()
    If MsgBox("Desea Cancelar?. Los Datos no se guardarn.", vbQuestion + vbYesNo, "AVISO") = vbNo Then Exit Sub
    Call Limpiar
End Sub
Private Sub cmdVer_Click()
    frmOCOcupacionCIIU.Show 1
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdGuardarPerfil_Click()
    If MsgBox("Desea Guardar los Datos?", vbExclamation + vbYesNo, "AVISO") = vbNo Then Exit Sub
    Call GuardarPerfil
    MsgBox "Los datos se guardaron Correctamente", vbInformation, "AVISO"
    Call Limpiar
End Sub
Private Sub optNatural_Click()
    If optNatural.value = True Then
        cboCIIU.Clear
        Call CargarOcupaciones
    End If
End Sub
Private Sub optJuridica_Click()
    If optJuridica.value = True Then
        cboCIIU.Clear
        Call CargarCIIU
    End If
End Sub
Private Sub Limpiar()
    bFormatea = True
    Call CargarVariables
    Call CargarSubVariables
    txtAnio.Text = ""
    txtCalificacion.Text = "0.00"
    lblNivelRiesgo.Caption = ""
    fraTipo.Enabled = True
    fraPeriodo.Enabled = True
    fraOcupacion.Enabled = True
    cmdGenerar.Enabled = True
    cmdGuardarPerfil.Enabled = False
    cmdCancelar.Enabled = False
End Sub
Private Sub CargarVariables()
    Dim oVariable As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    Dim fila As Integer
    
    fila = 0
    If bFormatea Then Call FormateaFlex(feResumen)
    Set rs = oVariable.ObtenerNivelRiesgoVariable(1)
    Do While Not rs.EOF
        fila = fila + 1
        feResumen.AdicionaFila
        feResumen.TextMatrix(fila, 1) = rs!cVarDescripcion
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
Private Sub CargarSubVariables()
    Dim oSubVariable As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    Dim nFilaPersona As Integer, nFilaProducto As Integer, nFilaCanales As Integer
    Dim nFilaJuris As Integer, nFilaComport As Integer, nFilaTransa As Integer
    
    nFilaPersona = 0: nFilaProducto = 0: nFilaCanales = 0
    nFilaJuris = 0: nFilaComport = 0: nFilaTransa = 0
    Set rs = oSubVariable.ObtenerNivelRiesgoVariable(2)
    If bFormatea Then
        Call FormateaFlex(fePersona)
        Call FormateaFlex(feProducto)
        Call FormateaFlex(feCanales)
        Call FormateaFlex(feJurisdiccion)
        Call FormateaFlex(feComportamiento)
        Call FormateaFlex(feTransacciones)
    End If
    Do While Not rs.EOF
        If rs!nVarCod = 100 Then
            nFilaPersona = nFilaPersona + 1
            fePersona.AdicionaFila
            fePersona.TextMatrix(nFilaPersona, 1) = rs!cVarSubDescripcion
        End If
        If rs!nVarCod = 200 Then
            nFilaProducto = nFilaProducto + 1
            feProducto.AdicionaFila
            feProducto.TextMatrix(nFilaProducto, 1) = rs!cVarSubDescripcion
        End If
        If rs!nVarCod = 300 Then
            nFilaCanales = nFilaCanales + 1
            feCanales.AdicionaFila
            feCanales.TextMatrix(nFilaCanales, 1) = rs!cVarSubDescripcion
        End If
        If rs!nVarCod = 400 Then
            nFilaJuris = nFilaJuris + 1
            feJurisdiccion.AdicionaFila
            feJurisdiccion.TextMatrix(nFilaJuris, 1) = rs!cVarSubDescripcion
        End If
        If rs!nVarCod = 500 Then
            nFilaComport = nFilaComport + 1
            feComportamiento.AdicionaFila
            feComportamiento.TextMatrix(nFilaComport, 1) = rs!cVarSubDescripcion
        End If
        If rs!nVarCod = 600 Then
            nFilaTransa = nFilaTransa + 1
            feTransacciones.AdicionaFila
            feTransacciones.TextMatrix(nFilaTransa, 1) = rs!cVarSubDescripcion
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
Private Sub CargarCIIU()
    Dim oConstante As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
    
    Set rs = oConstante.ObtenerParaOficialCumplimientoCIIU()
    Do While Not rs.EOF
        cboCIIU.AddItem (rs!cCIIUdescripcion & Space(100) & rs!cCIIUcod)
        rs.MoveNext
    Loop
End Sub
Private Sub CargarOcupaciones()
    Dim oPersona As New COMDPersona.DCOMPersonas
    Dim rsOcupa As ADODB.Recordset
    
    Set rsOcupa = oPersona.CargarOcupaciones()
    Do While Not rsOcupa.EOF
        cboCIIU.AddItem (rsOcupa!cConsDescripcion) & Space(200) & Trim(rsOcupa!nConsValor)
        rsOcupa.MoveNext
    Loop
    Set rsOcupa = Nothing
    Set oPersona = Nothing
End Sub
Public Function ValidaGenerar() As Boolean
ValidaGenerar = False
If Trim(txtAnio.Text) = "" Then
    MsgBox "Ingrese correctamente el año", vbInformation, "Aviso"
    txtAnio.SetFocus
    Exit Function
End If
If val(txtAnio.Text) < 1900 Or val(txtAnio.Text) > 9972 Then
    MsgBox "Ingrese correctamente el año", vbInformation, "Aviso"
    txtAnio.SetFocus
    Exit Function
End If
If Trim(cboMes.Text) = "" Then
    MsgBox "Ingrese correctamente el mes", vbInformation, "Aviso"
    cboMes.SetFocus
    Exit Function
End If
If CInt(Trim(Right(cboMes.Text, 2))) >= CInt(Mid(gdFecSis, 4, 2)) Or val(txtAnio.Text) > Right(gdFecSis, 4) Then
    MsgBox "El periodo seleccionado debe ser anterior al mes actual", vbInformation, "Aviso"
    Exit Function
End If

ValidaGenerar = True
End Function
Private Sub GuardarPerfil()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim aValPers(1 To 10) As Integer
    Dim aValPersPond(1 To 10) As Double
    Dim aValProd(1 To 9) As Integer
    Dim aValProdPond(1 To 9) As Double
    Dim aValCan(1 To 6) As Integer
    Dim aValCanPond(1 To 6) As Double
    Dim aValJur(1 To 2) As Integer
    Dim aValJurPond(1 To 2) As Double
    Dim aValComp(1 To 4) As Integer
    Dim aValCompPond(1 To 4) As Double
    Dim aValTrans(1 To 4) As Integer
    Dim aValTransPond(1 To 4) As Double
    Dim nValPersona As Double, nValProductos As Double, nValCanales As Double
    Dim nValJurisdiccion As Double, nValComportamiento As Double, nValTransacciones As Double
    Dim nCalificacion As Double
    Dim cNivelRiesgo As String
    
    aValPers(1) = fePersona.TextMatrix(1, 2)
    aValPers(2) = fePersona.TextMatrix(2, 2)
    aValPers(3) = fePersona.TextMatrix(3, 2)
    aValPers(4) = fePersona.TextMatrix(4, 2)
    aValPers(5) = fePersona.TextMatrix(5, 2)
    aValPers(6) = fePersona.TextMatrix(6, 2)
    aValPers(7) = fePersona.TextMatrix(7, 2)
    aValPers(8) = fePersona.TextMatrix(8, 2)
    aValPers(9) = fePersona.TextMatrix(9, 2)
    aValPers(10) = fePersona.TextMatrix(10, 2)
    
    aValPersPond(1) = fePersona.TextMatrix(1, 3)
    aValPersPond(2) = fePersona.TextMatrix(2, 3)
    aValPersPond(3) = fePersona.TextMatrix(3, 3)
    aValPersPond(4) = fePersona.TextMatrix(4, 3)
    aValPersPond(5) = fePersona.TextMatrix(5, 3)
    aValPersPond(6) = fePersona.TextMatrix(6, 3)
    aValPersPond(7) = fePersona.TextMatrix(7, 3)
    aValPersPond(8) = fePersona.TextMatrix(8, 3)
    aValPersPond(9) = fePersona.TextMatrix(9, 3)
    aValPersPond(10) = fePersona.TextMatrix(10, 3)
    
    aValProd(1) = feProducto.TextMatrix(1, 2)
    aValProd(2) = feProducto.TextMatrix(2, 2)
    aValProd(3) = feProducto.TextMatrix(3, 2)
    aValProd(4) = feProducto.TextMatrix(4, 2)
    aValProd(5) = feProducto.TextMatrix(5, 2)
    aValProd(6) = feProducto.TextMatrix(6, 2)
    aValProd(7) = feProducto.TextMatrix(7, 2)
    aValProd(8) = feProducto.TextMatrix(8, 2)
    aValProd(9) = feProducto.TextMatrix(9, 2)
        
    aValProdPond(1) = feProducto.TextMatrix(1, 3)
    aValProdPond(2) = feProducto.TextMatrix(2, 3)
    aValProdPond(3) = feProducto.TextMatrix(3, 3)
    aValProdPond(4) = feProducto.TextMatrix(4, 3)
    aValProdPond(5) = feProducto.TextMatrix(5, 3)
    aValProdPond(6) = feProducto.TextMatrix(6, 3)
    aValProdPond(7) = feProducto.TextMatrix(7, 3)
    aValProdPond(8) = feProducto.TextMatrix(8, 3)
    aValProdPond(9) = feProducto.TextMatrix(9, 3)
    
    aValCan(1) = feCanales.TextMatrix(1, 2)
    aValCan(2) = feCanales.TextMatrix(2, 2)
    aValCan(3) = feCanales.TextMatrix(3, 2)
    aValCan(4) = feCanales.TextMatrix(4, 2)
    aValCan(5) = feCanales.TextMatrix(5, 2)
    aValCan(6) = feCanales.TextMatrix(6, 2)
        
    aValCanPond(1) = feCanales.TextMatrix(1, 3)
    aValCanPond(2) = feCanales.TextMatrix(2, 3)
    aValCanPond(3) = feCanales.TextMatrix(3, 3)
    aValCanPond(4) = feCanales.TextMatrix(4, 3)
    aValCanPond(5) = feCanales.TextMatrix(5, 3)
    aValCanPond(6) = feCanales.TextMatrix(6, 3)
    
    aValJur(1) = feJurisdiccion.TextMatrix(1, 2)
    aValJur(2) = feJurisdiccion.TextMatrix(2, 2)
        
    aValJurPond(1) = feJurisdiccion.TextMatrix(1, 3)
    aValJurPond(2) = feJurisdiccion.TextMatrix(2, 3)
    
    aValComp(1) = feComportamiento.TextMatrix(1, 2)
    aValComp(2) = feComportamiento.TextMatrix(2, 2)
    aValComp(3) = feComportamiento.TextMatrix(3, 2)
    aValComp(4) = feComportamiento.TextMatrix(4, 2)
        
    aValCompPond(1) = feComportamiento.TextMatrix(1, 3)
    aValCompPond(2) = feComportamiento.TextMatrix(2, 3)
    aValCompPond(3) = feComportamiento.TextMatrix(3, 3)
    aValCompPond(4) = feComportamiento.TextMatrix(4, 3)
    
    aValTrans(1) = feTransacciones.TextMatrix(1, 2)
    aValTrans(2) = feTransacciones.TextMatrix(2, 2)
    aValTrans(3) = feTransacciones.TextMatrix(3, 2)
    aValTrans(4) = feTransacciones.TextMatrix(4, 2)
        
    aValTransPond(1) = feTransacciones.TextMatrix(1, 3)
    aValTransPond(2) = feTransacciones.TextMatrix(2, 3)
    aValTransPond(3) = feTransacciones.TextMatrix(3, 3)
    aValTransPond(4) = feTransacciones.TextMatrix(4, 3)
    
    'RESUMEN
    nValPersona = feResumen.TextMatrix(1, 2)
    nValProductos = feResumen.TextMatrix(2, 2)
    nValCanales = feResumen.TextMatrix(3, 2)
    nValJurisdiccion = feResumen.TextMatrix(4, 2)
    nValComportamiento = feResumen.TextMatrix(5, 2)
    nValTransacciones = feResumen.TextMatrix(6, 2)
        
    nCalificacion = txtCalificacion.Text
    cNivelRiesgo = lblNivelRiesgo.Caption
    
    Call oServ.RegistrarPerfilPorActividad(nTipoPersoneria, cCodOcuOrCIIU, nMes, cAnio, _
                                           aValPers(), aValPersPond(), aValProd(), aValProdPond(), aValCan(), aValCanPond(), _
                                           aValJur(), aValJurPond(), aValComp(), aValCompPond(), aValTrans(), aValTransPond(), _
                                           nValPersona, nValProductos, nValCanales, nValJurisdiccion, nValComportamiento, nValTransacciones, _
                                           nCalificacion, cNivelRiesgo)
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub CargarValoresVariables(ByVal pRs As ADODB.Recordset)
    Dim fila As Integer
    
    If Not pRs.EOF And Not pRs.BOF Then
        fePersona.TextMatrix(1, 2) = pRs!ValPersEdad
        fePersona.TextMatrix(2, 2) = pRs!ValPersNac
        fePersona.TextMatrix(3, 2) = pRs!ValPersRO
        fePersona.TextMatrix(4, 2) = pRs!ValPersRD
        fePersona.TextMatrix(5, 2) = pRs!ValPersTA
        fePersona.TextMatrix(6, 2) = pRs!ValPersUA
        fePersona.TextMatrix(7, 2) = pRs!ValPersEA
        fePersona.TextMatrix(8, 2) = pRs!ValPersAC
        fePersona.TextMatrix(9, 2) = pRs!ValPersCPS
        fePersona.TextMatrix(10, 2) = pRs!ValPersRLN
        
        fePersona.TextMatrix(1, 3) = Format(pRs!ValPersEdadPond, "0.00")
        fePersona.TextMatrix(2, 3) = Format(pRs!ValPersNacPond, "0.00")
        fePersona.TextMatrix(3, 3) = Format(pRs!ValPersROPond, "0.00")
        fePersona.TextMatrix(4, 3) = Format(pRs!ValPersRDPond, "0.00")
        fePersona.TextMatrix(5, 3) = Format(pRs!ValPersTAPond, "0.00")
        fePersona.TextMatrix(6, 3) = Format(pRs!ValPersUAPond, "0.00")
        fePersona.TextMatrix(7, 3) = Format(pRs!ValPersEAPond, "0.00")
        fePersona.TextMatrix(8, 3) = Format(pRs!ValPersACPond, "0.00")
        fePersona.TextMatrix(9, 3) = Format(pRs!ValPersCPSPond, "0.00")
        fePersona.TextMatrix(10, 3) = Format(pRs!ValPersRLNPond, "0.00")
        
        feProducto.TextMatrix(1, 2) = pRs!ValProdCA
        feProducto.TextMatrix(2, 2) = pRs!ValProdCPF
        feProducto.TextMatrix(3, 2) = pRs!ValProdCTS
        feProducto.TextMatrix(4, 2) = pRs!ValProdCMiE
        feProducto.TextMatrix(5, 2) = pRs!ValProdCPE
        feProducto.TextMatrix(6, 2) = pRs!ValProdCMeE
        feProducto.TextMatrix(7, 2) = pRs!ValProdCCon
        feProducto.TextMatrix(8, 2) = pRs!ValProdCHV
        feProducto.TextMatrix(9, 2) = pRs!ValProdTWU
        
        feProducto.TextMatrix(1, 3) = Format(pRs!ValProdCAPond, "0.00")
        feProducto.TextMatrix(2, 3) = Format(pRs!ValProdCPFPond, "0.00")
        feProducto.TextMatrix(3, 3) = Format(pRs!ValProdCTSPond, "0.00")
        feProducto.TextMatrix(4, 3) = Format(pRs!ValProdCMiEPond, "0.00")
        feProducto.TextMatrix(5, 3) = Format(pRs!ValProdCPEPond, "0.00")
        feProducto.TextMatrix(6, 3) = Format(pRs!ValProdCMeEPond, "0.00")
        feProducto.TextMatrix(7, 3) = Format(pRs!ValProdCConPond, "0.00")
        feProducto.TextMatrix(8, 3) = Format(pRs!ValProdCHVPond, "0.00")
        feProducto.TextMatrix(9, 3) = Format(pRs!ValProdTWUPond, "0.00")
        
        feCanales.TextMatrix(1, 2) = pRs!ValCanVen
        feCanales.TextMatrix(2, 2) = pRs!ValCanCAN
        feCanales.TextMatrix(3, 2) = pRs!ValCanCAI
        feCanales.TextMatrix(4, 2) = pRs!ValCanPOSN
        feCanales.TextMatrix(5, 2) = pRs!ValCanPOSI
        feCanales.TextMatrix(6, 2) = pRs!ValCanTE
        
        feCanales.TextMatrix(1, 3) = Format(pRs!ValCanVenPond, "0.00")
        feCanales.TextMatrix(2, 3) = Format(pRs!ValCanCANPond, "0.00")
        feCanales.TextMatrix(3, 3) = Format(pRs!ValCanCAIPond, "0.00")
        feCanales.TextMatrix(4, 3) = Format(pRs!ValCanPOSNPond, "0.00")
        feCanales.TextMatrix(5, 3) = Format(pRs!ValCanPOSIPond, "0.00")
        feCanales.TextMatrix(6, 3) = Format(pRs!ValCanTEPond, "0.00")
        
        feJurisdiccion.TextMatrix(1, 2) = pRs!ValJurRea
        feJurisdiccion.TextMatrix(2, 2) = pRs!ValJurVinc
        
        feJurisdiccion.TextMatrix(1, 3) = Format(pRs!ValJurReaPond, "0.00")
        feJurisdiccion.TextMatrix(2, 3) = Format(pRs!ValJurVincPond, "0.00")
        
        feComportamiento.TextMatrix(1, 2) = pRs!ValCompOIFIS
        feComportamiento.TextMatrix(2, 2) = pRs!ValCompT1
        feComportamiento.TextMatrix(3, 2) = pRs!ValCompT2
        feComportamiento.TextMatrix(4, 2) = pRs!ValCompROA
        
        feComportamiento.TextMatrix(1, 3) = Format(pRs!ValCompOIFISPond, "0.00")
        feComportamiento.TextMatrix(2, 3) = Format(pRs!ValCompT1Pond, "0.00")
        feComportamiento.TextMatrix(3, 3) = Format(pRs!ValCompT2Pond, "0.00")
        feComportamiento.TextMatrix(4, 3) = Format(pRs!ValCompROAPond, "0.00")
        
        feTransacciones.TextMatrix(1, 2) = pRs!ValTransMON
        feTransacciones.TextMatrix(2, 2) = pRs!ValTransIDEP
        feTransacciones.TextMatrix(3, 2) = pRs!ValTransIRET
        feTransacciones.TextMatrix(4, 2) = pRs!ValTransMACUM
        
        feTransacciones.TextMatrix(1, 3) = Format(pRs!ValTransMONPond, "0.00")
        feTransacciones.TextMatrix(2, 3) = Format(pRs!ValTransIDEPPond, "0.00")
        feTransacciones.TextMatrix(3, 3) = Format(pRs!ValTransIRETPond, "0.00")
        feTransacciones.TextMatrix(4, 3) = Format(pRs!ValTransMACUMPond, "0.00")
        
        'RESUMEN
        feResumen.TextMatrix(1, 2) = Format(pRs!ValPersona, "0.00")
        feResumen.TextMatrix(2, 2) = Format(pRs!ValProductos, "0.00")
        feResumen.TextMatrix(3, 2) = Format(pRs!ValCanales, "0.00")
        feResumen.TextMatrix(4, 2) = Format(pRs!ValJurisdiccion, "0.00")
        feResumen.TextMatrix(5, 2) = Format(pRs!ValComportamiento, "0.00")
        feResumen.TextMatrix(6, 2) = Format(pRs!ValTransacciones, "0.00")
        
        txtCalificacion.Text = Format(pRs!ValFinal, "#,##0.00")
        lblNivelRiesgo.Caption = pRs!NivelRiesgoFinal
    End If
End Sub
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function
Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAnio.SetFocus
    End If
End Sub
