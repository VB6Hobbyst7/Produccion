VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormato4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos - Evaluación - Formato 4"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalFormato4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMNME 
      Caption         =   "MN - ME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6260
      TabIndex        =   147
      Top             =   10180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Hoja Evaluación"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2820
      TabIndex        =   37
      Top             =   10180
      Width           =   1530
   End
   Begin VB.CommandButton cmdCancelar4 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12240
      TabIndex        =   34
      Top             =   10180
      Width           =   1170
   End
   Begin VB.CommandButton cmdGuardar4 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11070
      TabIndex        =   33
      Top             =   10180
      Width           =   1170
   End
   Begin VB.CommandButton cmdFlujoCaja4 
      Caption         =   "Generar &Flujo Caja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4360
      TabIndex        =   78
      Top             =   10180
      Width           =   1890
   End
   Begin VB.CommandButton cmdVerCar 
      Caption         =   "&Ver CAR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1650
      TabIndex        =   36
      Top             =   10180
      Width           =   1170
   End
   Begin VB.CommandButton cmdInformeVisita4 
      Caption         =   "Infor&me de Visita"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   35
      Top             =   10180
      Width           =   1650
   End
   Begin TabDlg.SSTab SSTabIngresos 
      Height          =   7260
      Left            =   0
      TabIndex        =   1
      Top             =   2050
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   12806
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ingresos y Egresos"
      TabPicture(0)   =   "frmCredFormEvalFormato4.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Evaluación"
      TabPicture(1)   =   "frmCredFormEvalFormato4.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Propuesta del Crédito"
      TabPicture(2)   =   "frmCredFormEvalFormato4.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "framePropuesta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Comentarios y Referidos"
      TabPicture(3)   =   "frmCredFormEvalFormato4.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frameReferido"
      Tab(3).Control(1)=   "frameComentario"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Pasivos :"
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
         Height          =   6600
         Left            =   -68260
         TabIndex        =   124
         Top             =   310
         Width           =   6675
         Begin SICMACT.FlexEdit fePasivos 
            Height          =   6375
            Left            =   60
            TabIndex        =   17
            Top             =   200
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   11245
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "-Concepto-P. P.-P. E.-Total-nConsCod-nConsValor"
            EncabezadosAnchos=   "0-2630-1150-1150-1450-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-2-3-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-R-C-C"
            FormatosEdit    =   "0-0-2-2-2-2-2"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Otros Ingresos :"
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
         Height          =   2175
         Left            =   -67800
         TabIndex        =   123
         Top             =   3240
         Width           =   6135
         Begin SICMACT.FlexEdit feOtrosIngresos 
            Height          =   1815
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   3201
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-3500-1800-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Gastos Familiares : "
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
         Height          =   2775
         Left            =   -67800
         TabIndex        =   122
         Top             =   360
         Width           =   6135
         Begin SICMACT.FlexEdit feGastosFamiliares 
            Height          =   2415
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   4260
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-3500-1800-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame frameReferido 
         Caption         =   "Referidos :"
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
         Height          =   3255
         Left            =   -74640
         TabIndex        =   121
         Top             =   3360
         Width           =   12615
         Begin VB.CommandButton cmdQuitar4 
            Caption         =   "&Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10800
            TabIndex        =   32
            Top             =   2760
            Width           =   1170
         End
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   9480
            TabIndex        =   31
            Top             =   2760
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feReferidos 
            Height          =   2415
            Left            =   360
            TabIndex        =   30
            Top             =   240
            Width           =   11715
            _ExtentX        =   20664
            _ExtentY        =   4260
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "N-Nombres-DNI-Teléfono-Comentario-NroDNI-Aux"
            EncabezadosAnchos=   "350-4500-960-1260-4500-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "N"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Declaración PDT:"
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
         Height          =   1420
         Left            =   -69840
         TabIndex        =   101
         Top             =   5520
         Width           =   8235
         Begin SICMACT.FlexEdit feDeclaracionPDT 
            Height          =   1095
            Left            =   45
            TabIndex        =   21
            Top             =   240
            Width           =   8120
            _ExtentX        =   14314
            _ExtentY        =   1931
            Rows            =   3
            Cols0           =   9
            FixedCols       =   2
            HighLight       =   1
            EncabezadosNombres=   "-Mes/Detalle-nConsCod-nConsValor----Promedio-%Declarado"
            EncabezadosAnchos=   "0-1200-0-0-1350-1350-1350-1500-1230"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-X-4-5-6-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-R-R-R-C-R"
            FormatosEdit    =   "0-0-0-0-2-2-2-0-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Flujo de Caja Mensual :"
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
         Height          =   6880
         Left            =   -74880
         TabIndex        =   100
         Top             =   340
         Width           =   5000
         Begin SICMACT.FlexEdit feFlujoCajaMensual 
            Height          =   6660
            Left            =   60
            TabIndex        =   18
            Top             =   180
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   11748
            Cols0           =   6
            HighLight       =   1
            EncabezadosNombres=   "N-nConsCod-nConsValor-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-0-0-3150-1600-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-X-4-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-R-C"
            FormatosEdit    =   "0-0-0-0-2-0"
            TextArray0      =   "N"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame frameComentario 
         Caption         =   "Comentarios :"
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
         Height          =   2535
         Left            =   -74640
         TabIndex        =   86
         Top             =   360
         Width           =   12735
         Begin VB.TextBox txtComentario4 
            Height          =   2130
            IMEMode         =   3  'DISABLE
            Left            =   360
            MaxLength       =   3000
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   12015
         End
      End
      Begin VB.Frame framePropuesta 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   6135
         Left            =   120
         TabIndex        =   79
         Top             =   480
         Width           =   12975
         Begin VB.TextBox txtDestino4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   5400
            Width           =   12495
         End
         Begin VB.TextBox txtColaterales4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   4440
            Width           =   12615
         End
         Begin VB.TextBox txtFormalidadNegocio4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   3480
            Width           =   12615
         End
         Begin VB.TextBox txtExperiencia4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   2520
            Width           =   12615
         End
         Begin VB.TextBox txtGiroUbicacion4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1560
            Width           =   12615
         End
         Begin VB.TextBox txtEntornoFamiliar4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   600
            Width           =   12615
         End
         Begin MSMask.MaskEdBox txtFechaVisita 
            Height          =   300
            Left            =   11520
            TabIndex        =   22
            Top             =   240
            Width           =   1090
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Visita:"
            Height          =   195
            Left            =   10320
            TabIndex        =   120
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el destino y el impacto del mismo :"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   5160
            Width           =   2895
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   4200
            Width           =   2400
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   240
            TabIndex        =   83
            Top             =   3240
            Width           =   4770
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la experiencia crediticia:"
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   2280
            Width           =   2190
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   81
            Top             =   1320
            Width           =   2820
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Width           =   3795
         End
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73200
         TabIndex        =   77
         Top             =   6120
         Width           =   1170
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74640
         TabIndex        =   76
         Top             =   6120
         Width           =   1170
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   74
         Top             =   3360
         Width           =   9975
         Begin SICMACT.FlexEdit FlexEdit1 
            Height          =   1935
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   3413
            Cols0           =   6
            HighLight       =   1
            EncabezadosNombres=   "N°-Nombre-DNI-Telef.-Referido-DNI"
            EncabezadosAnchos=   "1000-2800-1000-1500-2300-1000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-C-C-C"
            FormatosEdit    =   "0-2-0-0-0-0"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   1005
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   72
         Top             =   360
         Width           =   9975
         Begin VB.TextBox Text1 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame frmCredEvalFormato1 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   6015
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtDestino 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   65
            Top             =   5280
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   480
            Width           =   9735
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   5040
            Width           =   2400
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   3795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Activos :"
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
         Height          =   5895
         Left            =   -74930
         TabIndex        =   58
         Top             =   310
         Width           =   6630
         Begin SICMACT.FlexEdit feActivos 
            Height          =   5535
            Left            =   40
            TabIndex        =   16
            Top             =   200
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   9763
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "-Concepto-P. P.-P. E.-Total-nConsCod-nConsValor"
            EncabezadosAnchos=   "0-2630-1150-1150-1450-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-2-3-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-R-R-C"
            FormatosEdit    =   "0-0-2-2-2-3-3"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   6015
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtDestino2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   5280
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtEntornoFamiliar2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   480
            Width           =   9735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   5040
            Width           =   2400
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   3795
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   43
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtComentario2 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Referidos :"
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   41
         Top             =   3240
         Width           =   9975
         Begin SICMACT.FlexEdit feReferidos2 
            Height          =   1935
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   3413
            Cols0           =   6
            HighLight       =   1
            EncabezadosNombres=   "N°-Nombre-DNI-Telef.-Referido-DNI"
            EncabezadosAnchos=   "1000-2800-1000-1500-2300-1000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-C-C-C"
            FormatosEdit    =   "0-2-0-0-0-0"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   1005
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73200
         TabIndex        =   40
         Top             =   5640
         Width           =   1170
      End
      Begin VB.CommandButton cmdAgregar2 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74640
         TabIndex        =   39
         Top             =   5640
         Width           =   1170
      End
      Begin VB.Line Line2 
         X1              =   -69000
         X2              =   -69000
         Y1              =   480
         Y2              =   5280
      End
   End
   Begin TabDlg.SSTab SSTabRatios 
      Height          =   860
      Left            =   0
      TabIndex        =   87
      Top             =   9310
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1508
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ratios e Indicadores"
      TabPicture(0)   =   "frmCredFormEvalFormato4.frx":037A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label33"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label32"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label19(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCapaAceptable"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblEndeAceptable"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtRentabilidadPat"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtLiquidezCte"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtExcedenteMensual"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtIngresoNeto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtEndeudamiento"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCapacidadNeta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Datos Flujo Caja Proyectado"
      TabPicture(1)   =   "frmCredFormEvalFormato4.frx":0396
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13(12)"
      Tab(1).Control(1)=   "Label13(11)"
      Tab(1).Control(2)=   "Label13(10)"
      Tab(1).Control(3)=   "Label13(9)"
      Tab(1).Control(4)=   "Label13(8)"
      Tab(1).Control(5)=   "Line4"
      Tab(1).Control(6)=   "Label13(7)"
      Tab(1).Control(7)=   "Label13(6)"
      Tab(1).Control(8)=   "Label13(5)"
      Tab(1).Control(9)=   "Label13(4)"
      Tab(1).Control(10)=   "Label13(3)"
      Tab(1).Control(11)=   "Label34"
      Tab(1).Control(12)=   "Label21"
      Tab(1).Control(13)=   "Label22"
      Tab(1).Control(14)=   "Label28"
      Tab(1).Control(15)=   "Label29"
      Tab(1).Control(16)=   "EditMoneyIncC4"
      Tab(1).Control(17)=   "EditMoneyIncGV4"
      Tab(1).Control(18)=   "EditMoneyIncPP4"
      Tab(1).Control(19)=   "EditMoneyIncCM4"
      Tab(1).Control(20)=   "EditMoneyIncVC4"
      Tab(1).ControlCount=   21
      Begin SICMACT.EditMoney txtCapacidadNeta 
         Height          =   300
         Left            =   1560
         TabIndex        =   88
         Top             =   330
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtEndeudamiento 
         Height          =   300
         Left            =   3700
         TabIndex        =   89
         Top             =   330
         Width           =   850
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIngresoNeto 
         Height          =   300
         Left            =   10200
         TabIndex        =   90
         Top             =   375
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtExcedenteMensual 
         Height          =   300
         Left            =   12240
         TabIndex        =   91
         Top             =   375
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtLiquidezCte 
         Height          =   300
         Left            =   7920
         TabIndex        =   92
         Top             =   375
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtRentabilidadPat 
         Height          =   300
         Left            =   5960
         TabIndex        =   93
         Top             =   330
         Width           =   850
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncVC4 
         Height          =   300
         Left            =   -74880
         TabIndex        =   127
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   500
         Width           =   1000
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "00.0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncCM4 
         Height          =   300
         Left            =   -72600
         TabIndex        =   128
         ToolTipText     =   "Incremento de Compras de Mercaderias - Anual"
         Top             =   500
         Width           =   1005
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "00.0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncPP4 
         Height          =   300
         Left            =   -70320
         TabIndex        =   129
         ToolTipText     =   "Incremento de Consumo - Anual"
         Top             =   500
         Width           =   1005
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "00.0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncGV4 
         Height          =   300
         Left            =   -68040
         TabIndex        =   130
         ToolTipText     =   "Incremento de Pago Personal -Anual"
         Top             =   500
         Width           =   1005
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "00.0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncC4 
         Height          =   300
         Left            =   -63480
         TabIndex        =   131
         ToolTipText     =   "Incremento de Gastos de Ventas - Anual"
         Top             =   510
         Width           =   1005
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "00.0"
      End
      Begin VB.Label Label29 
         Caption         =   "Anual"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   -62160
         TabIndex        =   146
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -66720
         TabIndex        =   145
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -69000
         TabIndex        =   144
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -71280
         TabIndex        =   143
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "Anual"
         Height          =   255
         Left            =   -73560
         TabIndex        =   142
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. ventas contado:"
         Height          =   195
         Index           =   3
         Left            =   -74880
         TabIndex        =   141
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. compra mercaderias:"
         Height          =   195
         Index           =   4
         Left            =   -72600
         TabIndex        =   140
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. de Consumo:"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   -64920
         TabIndex        =   139
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Pago Personal:"
         Height          =   195
         Index           =   6
         Left            =   -70320
         TabIndex        =   138
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Gasto Ventas:"
         Height          =   195
         Index           =   7
         Left            =   -68040
         TabIndex        =   137
         Top             =   300
         Width           =   1410
      End
      Begin VB.Line Line4 
         X1              =   -66000
         X2              =   -66000
         Y1              =   480
         Y2              =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   8
         Left            =   -73800
         TabIndex        =   136
         Top             =   550
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   9
         Left            =   -71520
         TabIndex        =   135
         Top             =   550
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   10
         Left            =   -69240
         TabIndex        =   134
         Top             =   555
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   11
         Left            =   -66960
         TabIndex        =   133
         Top             =   555
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   12
         Left            =   -62400
         TabIndex        =   132
         Top             =   550
         Width           =   165
      End
      Begin VB.Label lblEndeAceptable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3720
         TabIndex        =   126
         Top             =   645
         Width           =   750
      End
      Begin VB.Label lblCapaAceptable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1600
         TabIndex        =   125
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad de Pago:"
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   99
         Top             =   405
         Width           =   1440
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endeudamiento:"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   98
         Top             =   400
         Width           =   1170
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Neto:"
         Height          =   195
         Left            =   9180
         TabIndex        =   97
         Top             =   400
         Width           =   1005
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excedente:"
         Height          =   195
         Left            =   11440
         TabIndex        =   96
         Top             =   400
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidez Cte:"
         Height          =   195
         Index           =   3
         Left            =   6980
         TabIndex        =   95
         Top             =   400
         Width           =   930
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rentabilidad Pat.:"
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   94
         Top             =   400
         Width           =   1290
      End
      Begin VB.Line Line1 
         X1              =   6840
         X2              =   6840
         Y1              =   360
         Y2              =   720
      End
   End
   Begin TabDlg.SSTab SSTabInfoNego 
      Height          =   2050
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Información del Negocio"
      TabPicture(0)   =   "frmCredFormEvalFormato4.frx":03B2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ActXCodCta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frameLinea"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGiroNeg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtGiroNeg 
         Height          =   300
         Left            =   7680
         TabIndex        =   3
         Top             =   360
         Width           =   5475
      End
      Begin VB.Frame frameLinea 
         Height          =   255
         Left            =   4800
         TabIndex        =   116
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtNumLinea 
            Height          =   300
            Left            =   1800
            TabIndex        =   117
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Linea Automática :"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   118
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1380
         Left            =   120
         TabIndex        =   103
         Top             =   620
         Width           =   13215
         Begin VB.TextBox txtCondLocalOtros 
            Height          =   300
            Left            =   6765
            MaxLength       =   250
            TabIndex        =   15
            Top             =   1050
            Visible         =   0   'False
            Width           =   6315
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Otros"
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   14
            Top             =   1060
            Width           =   855
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Ambulante"
            Height          =   255
            Index           =   3
            Left            =   4680
            TabIndex        =   13
            Top             =   1060
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Alquilada"
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   12
            Top             =   1060
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Propia"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   11
            Top             =   1060
            Width           =   855
         End
         Begin VB.TextBox txtNombreCliente 
            Height          =   300
            Left            =   2400
            TabIndex        =   4
            Top             =   120
            Width           =   4155
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda 
            Height          =   300
            Left            =   12000
            TabIndex        =   9
            Top             =   420
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Spinner.uSpinner spnTiempoLocalAnio 
            Height          =   300
            Left            =   2400
            TabIndex        =   0
            Top             =   760
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   99
            MaxLength       =   2
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
         Begin Spinner.uSpinner spnTiempoLocalMes 
            Height          =   300
            Left            =   3720
            TabIndex        =   38
            Top             =   760
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   12
            MaxLength       =   2
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
         Begin SICMACT.EditMoney txtExposicionCredito 
            Height          =   300
            Left            =   11760
            TabIndex        =   5
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Spinner.uSpinner spnExpEmpAnio 
            Height          =   300
            Left            =   2400
            TabIndex        =   6
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   99
            MaxLength       =   2
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
            ForeColor       =   8421504
         End
         Begin Spinner.uSpinner spnExpEmpMes 
            Height          =   300
            Left            =   3720
            TabIndex        =   7
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            Max             =   12
            MaxLength       =   2
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
            ForeColor       =   8421504
         End
         Begin SICMACT.EditMoney txtUltEndeuda 
            Height          =   300
            Left            =   7680
            TabIndex        =   8
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox txtFechaEvaluacion 
            Height          =   300
            Left            =   12000
            TabIndex        =   10
            Top             =   740
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha último endeudamiento RCC :"
            Height          =   195
            Left            =   9225
            TabIndex        =   115
            Top             =   480
            Width           =   2520
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Último endeudamiento RCC :"
            Height          =   195
            Left            =   5640
            TabIndex        =   114
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4515
            TabIndex        =   113
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4515
            TabIndex        =   112
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   3195
            TabIndex        =   111
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   3195
            TabIndex        =   110
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exposición con este crédito :"
            Height          =   195
            Left            =   9615
            TabIndex        =   109
            Top             =   165
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición local :"
            Height          =   255
            Left            =   1200
            TabIndex        =   108
            Top             =   1060
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo en el mismo local :"
            Height          =   255
            Left            =   480
            TabIndex        =   107
            Top             =   780
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experiencia como empresario :"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   450
            Width           =   2295
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   1680
            TabIndex        =   105
            Top             =   160
            Width           =   555
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de evaluación al :"
            Height          =   195
            Left            =   10035
            TabIndex        =   104
            Top             =   780
            Width           =   1740
         End
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   280
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro del Negocio :"
         Height          =   255
         Left            =   6440
         TabIndex        =   119
         Top             =   390
         Width           =   1335
      End
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   13200
      Y1              =   9960
      Y2              =   9960
   End
End
Attribute VB_Name = "frmCredFormEvalFormato4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalFormato4
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit
    Dim sCtaCod As String
    Dim sPersCod As String
    Dim gsOpeCod As String
    Dim fnTipoRegMant As Integer
    Dim fnTipoPermiso As Integer
    Dim fbPermiteGrabar As Boolean
    Dim fbBloqueaTodo As Boolean
    Dim fnTotalRefGastoNego As Currency
    Dim fnTotalRefGastoFami As Currency
    Dim fsCliente As String
    Dim fsGiroNego As String
    Dim fsAnioExp As Integer
    Dim fsMesExp As Integer
    Dim fnEstado As Integer
    Dim fnMontoDeudaSbs As Currency
    Dim fnFechaDeudaSbs As Currency
    
    Dim fnCondLocal As Integer
    Dim MatIfiGastoNego As Variant
    Dim MatIfiGastoFami As Variant
    Dim MatReferidos As Variant
    
    Dim rsFeGastoNeg As ADODB.Recordset
    Dim rsFeDatGastoFam As ADODB.Recordset
    Dim rsFeDatOtrosIng As ADODB.Recordset
    Dim rsFeDatBalanGen As ADODB.Recordset
    Dim rsFeDatActivos As ADODB.Recordset
    Dim rsFeDatPasivos As ADODB.Recordset
    Dim rsFeDatPasivosNo As ADODB.Recordset
    Dim rsFeDatPatrimonio As ADODB.Recordset
    Dim rsFeDatRef As ADODB.Recordset
    Dim rsFeFlujoCaja As ADODB.Recordset
    Dim rsFeDatPDT As ADODB.Recordset
    
    Dim rsCredEval As ADODB.Recordset
    Dim rsDCredito As ADODB.Recordset
    Dim rsAceptableCritico As ADODB.Recordset
    Dim rsCapacPagoNeta As ADODB.Recordset
    Dim rsCuotaIFIs As ADODB.Recordset
    Dim rsPropuesta As ADODB.Recordset
        
    Dim rsDatPasivosNo As ADODB.Recordset
    Dim rsDatActivoPasivo As ADODB.Recordset
    Dim rsDatGastoNeg As ADODB.Recordset
    Dim rsDatGastoFam As ADODB.Recordset
    Dim rsDatOtrosIng As ADODB.Recordset
    Dim rsDatRef As ADODB.Recordset
    Dim rsDatRatioInd As ADODB.Recordset
    Dim rsDatIfiGastoNego As ADODB.Recordset
    Dim rsDatIfiGastoFami As ADODB.Recordset
    Dim rsDatPDT As ADODB.Recordset
    Dim rsDatPDTDet As ADODB.Recordset
    Dim rsDatFlujoCaja As ADODB.Recordset
    Dim rsDatActivos As ADODB.Recordset
    Dim rsDatPasivos As ADODB.Recordset
    
    Dim nMontoAct As Double
    Dim nMontoPas As Double
    Dim nMontoPat As Double
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim objPista As COMManejador.Pista
    Dim fnFormato As Integer
    Dim fnMontoIni As Double
    Dim lnMin As Double
    Dim lnMax As Double
    Dim lnMinDol As Double
    Dim lnMaxDol As Double
    Dim nTC As Double
    Dim i As Integer, j As Integer, K As Integer
    
    Dim sMes1 As String, sMes2 As String, sMes3 As String
    Dim nMes1 As Integer, nMes2 As Integer, nMes3 As Integer
    Dim nAnio1 As Integer, nAnio2 As Integer, nAnio3 As Integer
    Dim fbGrabar As Boolean
    Dim fnColocCondi As Integer
    Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agregó segun correo: RUSI
    
    'LUCV20160705 **********-> Trabajando con Matrices TYPE
    Dim lvPrincipalActivos() As tFormEvalPrincipalActivosFormato5 'Matriz Principal-> Activos
    Dim lvPrincipalPasivos() As tFormEvalPrincipalPasivosFormato5 'Matriz Principal-> Pasivos
    'Detalle de Activos
    Dim lvDetalleActivosCtasCobrar() As tFormEvalDetalleActivosCtasCobrarFormato5 'Ctas x Cobrar
    Dim oFrmCtaCobrarIfi As frmCredFormEvalCtasCobrarIfis                         'Formulario: Ctas x Cobrar

    'Fin LUCV20160705 <-**********
    'RECO20160916******************************
    Dim bActivoDetPP(2) As Boolean
    Dim bActivoDetPE(2) As Boolean

    'RECO FIN *********************************

    'JOEP20171102 Flujo de Caja
    Dim rsDatParamFlujoCajaForm4 As ADODB.Recordset
    Dim nMaximo As Integer
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim lcMovNro As String 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    
Private Sub cmdFlujoCaja4_Click()

On Error GoTo ErrorInicioExcel 'pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM

Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim bGeneraExcel As Boolean 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM


    'lsArchivo = App.Path & "\Spooler\FlujoCaja_Formato4" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls" 'comentado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
    lsArchivo = App.Path & "\Spooler\FlujoCaja_Formato4" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls" 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
    lbLibroOpen = ExcelInicio(lsArchivo, xlAplicacion, xlLibro)
    
    If lbLibroOpen Then
    bGeneraExcel = False 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        bGeneraExcel = generaExcelForm4 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        If bGeneraExcel Then 'modificado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
            ExcelFin lsArchivo, xlAplicacion, xlLibro, xlHoja1
            'AbrirArchivo "FlujoCaja_Formato4" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler" 'comentado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
            AbrirArchivo "FlujoCaja_Formato4" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler" 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        End If
    End If
    
Exit Sub 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
ErrorInicioExcel: 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
MsgBox Err.Description + "Error1: Error al iniciar la creación del excel, Comuníquese con el Area de TI", vbInformation, "Error" 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
'End If

End Sub

Public Function generaExcelForm4() As Boolean

    On Error GoTo ErrorInicioExcel 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM


    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim rsCabcera As New ADODB.Recordset
    Dim rsCuotas As New ADODB.Recordset
    Dim rsParFlujoCaja As New ADODB.Recordset
    Dim oCont As COMConecta.DCOMConecta
    Dim i As Integer
    Dim nCon As Integer
    Dim nFila As Integer
    Dim nCol As Integer
    Dim nColFin As Integer
    Dim A As Integer
    Dim nColInicio As Integer
    Dim Z As Integer
    
    Dim dFechaEval As Date

    generaExcelForm4 = True

    'proteger Libro
    'xlAplicacion.ActiveWorkbook.Protect (123) 'pti comentado
    
    'Adiciona una hoja
    ExcelAddHoja "Hoja1", xlLibro, xlHoja1, True
               
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    
    xlHoja1.Cells(2, 2) = "FLUJO DE CAJA MENSUAL PRESUPUESTADO"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).Font.Bold = True
    
    xlHoja1.Cells(4, 1) = "CLIENTE: "
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(5, 1) = "ANALISTA: "
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(6, 1) = "DNI: "
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(7, 1) = "RUC: "
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 1)).HorizontalAlignment = xlLeft
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCabecera  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsCabcera = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCuotas  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsCuotas = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosConceptos  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rs = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosParametros  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsParFlujoCaja = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing

'Cabecera
If Not (rsCabcera.EOF And rsCabcera.BOF) Then
    dFechaEval = rsCabcera!fechaEval
    
    xlHoja1.Cells(4, 2) = rsCabcera!NombreClie
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).Font.Bold = True

    xlHoja1.Cells(5, 2) = rsCabcera!NombreAnal
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).Font.Bold = True
    
    xlHoja1.Cells(6, 2) = rsCabcera!nDoc
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).Font.Bold = True
    
    xlHoja1.Cells(7, 2) = rsCabcera!nDocTrib
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).Font.Bold = True
    
Else
        MsgBox "Error, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm4 = False
        Exit Function
End If
    
    
    xlHoja1.Cells(9, 2) = "Conceptos / Meses"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 2)).Cells.Interior.Color = RGB(141, 180, 226)
        
    xlHoja1.Cells(9, 3) = "Flujo Mensual"
    xlHoja1.Cells(10, 3) = Format(dFechaEval, "mmm-yyyy")
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 3)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 3)).HorizontalAlignment = xlCenter
    
    CuadroExcel xlHoja1, 2, 9, 3, 10, True
    CuadroExcel xlHoja1, 2, 9, 3, 10, False
    
    nCon = 11
    
'Conceptos
    If Not (rs.EOF And rs.BOF) Then
        For i = 1 To rs.RecordCount
        
            CuadroExcel xlHoja1, 2, nCon, 3, nCon
        
            xlHoja1.Cells(nCon, 2) = rs!Descripcion
            xlHoja1.Cells(nCon, 3) = rs!Monto
            
            If rs!Descripcion = "INVERSION" Then
                nCon = nCon + 2
            Else
                nCon = nCon + 1
            End If
                        
            CuadroExcel xlHoja1, 2, nCon, 3, nCon - 1
            
            rs.MoveNext
        Next i
    Else
        MsgBox "Error, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm4 = False
        Exit Function
    End If
  
'Pie
    If Not (rsParFlujoCaja.EOF And rsParFlujoCaja.BOF) Then

        xlHoja1.Cells(nCon + 1, 2) = "DATOS ADICIONALES"
        xlHoja1.Range(xlHoja1.Cells(nCon + 1, 2), xlHoja1.Cells(nCon + 1, 2)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nCon + 1, 2), xlHoja1.Cells(nCon + 1, 2)).HorizontalAlignment = xlCenter
        CuadroExcel xlHoja1, 2, nCon + 1, 3, nCon + 1
        xlHoja1.Cells(nCon + 2, 2) = "Fecha de Pago"
        xlHoja1.Cells(nCon + 2, 3) = Format(rsParFlujoCaja!dFechaPago, "YYYY/mm/dd")
        CuadroExcel xlHoja1, 2, nCon + 2, 3, nCon + 2

        xlHoja1.Cells(nCon + 4, 3) = "Mes"
        xlHoja1.Cells(nCon + 4, 4) = "Anual"
        CuadroExcel xlHoja1, 3, nCon + 4, 4, nCon + 4
        xlHoja1.Range(xlHoja1.Cells(nCon + 4, 3), xlHoja1.Cells(nCon + 4, 4)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nCon + 4, 3), xlHoja1.Cells(nCon + 4, 4)).HorizontalAlignment = xlCenter

        xlHoja1.Cells(nCon + 5, 2) = "Incremento de ventas al contado "
        xlHoja1.Cells(nCon + 6, 2) = "Incremento de Compra de Mercaderias"
        xlHoja1.Cells(nCon + 7, 2) = "Incremento de Consumo"
        xlHoja1.Cells(nCon + 8, 2) = "Incremento de Pago Personal"
        xlHoja1.Cells(nCon + 9, 2) = "Ingremento de Gastos de Ventas"

        xlHoja1.Cells(nCon + 5, 3) = Format(((1 + rsParFlujoCaja!nIncVentCont / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 6, 3) = Format(((1 + rsParFlujoCaja!nIncCompMerc / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 7, 3) = Format(((1 + rsParFlujoCaja!nIncConsu / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 8, 3) = Format(((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 9, 3) = Format(((1 + rsParFlujoCaja!nIncGastvent / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"

        xlHoja1.Cells(nCon + 5, 4) = Format(rsParFlujoCaja!nIncVentCont, "#0.0") & "%"
        xlHoja1.Cells(nCon + 6, 4) = Format(rsParFlujoCaja!nIncCompMerc, "#0.0") & "%"
        xlHoja1.Cells(nCon + 7, 4) = Format(rsParFlujoCaja!nIncConsu, "#0.0") & "%"
        xlHoja1.Cells(nCon + 8, 4) = Format(rsParFlujoCaja!nIncPagPers, "#0.0") & "%"
        xlHoja1.Cells(nCon + 9, 4) = Format(rsParFlujoCaja!nIncGastvent, "#0.0") & "%"

        CuadroExcel xlHoja1, 2, nCon + 5, 4, nCon + 9, True
        CuadroExcel xlHoja1, 2, nCon + 5, 4, nCon + 9, False
        xlHoja1.Range(xlHoja1.Cells(nCon + 5, 3), xlHoja1.Cells(nCon + 9, 4)).HorizontalAlignment = xlCenter

    Else
        MsgBox "Registre los Datos de Flujo de Caja Proyectado, y dar click en Guardar", vbInformation, "!Aviso!"
        generaExcelForm4 = False
        Exit Function
    End If
    
'Obtener las Letras del Abecedario A-Z
    Dim MatAZ As Variant
    Dim P As Integer
    P = 1
    Set MatAZ = Nothing
    ReDim MatAZ(1, 140)
    For i = 65 To 90
        MatAZ(1, P) = ChrW(i)
        P = P + 1
    Next i
           
    Dim MatLetrasRep As Variant
    Dim Y As Integer
    Set MatLetrasRep = Nothing
    Y = 1
    ReDim MatLetrasRep(1, 131)
    For A = 1 To 130
        If A <= 26 Then
                MatLetrasRep(1, Y) = ChrW(65) & MatAZ(1, Y) 'AA,AB,AC......AZ
            Y = Y + 1
        ElseIf (A >= 27 And A <= 52) Then
            If A = 27 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(66) & MatAZ(1, P) 'BA,BB,BC......BZ
            Y = Y + 1
            P = P + 1
        ElseIf (A >= 53 And A <= 78) Then
            If A = 53 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(67) & MatAZ(1, P) 'CA,CB,CC......CZ
            Y = Y + 1
            P = P + 1
        ElseIf (A >= 79 And A <= 104) Then
            If A = 79 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(68) & MatAZ(1, P) 'DA,DB,DC......DZ
            Y = Y + 1
            P = P + 1
        End If
    Next A
    
''Cuotas
i = 0
Y = 0
Z = 0
nFila = 39
nCol = 4
nColInicio = 4
nColFin = 0
   If Not (rsCuotas.EOF And rsCuotas.BOF) Then
        For i = 1 To rsCuotas.RecordCount

            If i >= 24 Then
                Y = Y + 1
            End If

            xlHoja1.Cells(9, nCol) = rsCuotas!nCuota
            xlHoja1.Range(xlHoja1.Cells(9, 4), xlHoja1.Cells(9, nCol)).Cells.Interior.Color = RGB(141, 180, 226)
            xlHoja1.Range(xlHoja1.Cells(9, nCol), xlHoja1.Cells(9, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(9, nCol), xlHoja1.Cells(9, nCon)).HorizontalAlignment = xlCenter
            
            xlHoja1.Cells(10, nCol) = Format(rsCuotas!dFechaCuotas, "mmm-yyyy")
            xlHoja1.Range(xlHoja1.Cells(10, 4), xlHoja1.Cells(10, nCol)).Cells.Interior.Color = RGB(141, 180, 226)
            xlHoja1.Range(xlHoja1.Cells(10, nCol), xlHoja1.Cells(10, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(10, nCol), xlHoja1.Cells(10, nCon)).HorizontalAlignment = xlCenter

            'calculo Ingresos Operativos
            xlHoja1.Range(xlHoja1.Cells(11, 2), xlHoja1.Cells(11, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(11, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "12" & ":" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "15)"
            
            'Ventas al Contado
            xlHoja1.Cells(12, nCol) = Round((xlHoja1.Cells(12, nCol - 1) * ((1 + rsParFlujoCaja!nIncVentCont / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(12, nCol - 1)))
            
            'Ventas al Credito
            xlHoja1.Cells(13, nCol) = "=C13"
            
            'Ventas de Activo Fijo
            'xlHoja1.Cells(13, nCol) = "=C13"
            
            'Otros Ingresos
            xlHoja1.Cells(15, nCol) = "=C15"
            
            'calculo Engresos Operativos
            xlHoja1.Cells(16, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "17" & ":" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "31)"
            xlHoja1.Range(xlHoja1.Cells(16, 2), xlHoja1.Cells(16, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            
            'Egreso por Compras (Mercaderia)
            xlHoja1.Cells(17, nCol) = Round((xlHoja1.Cells(17, nCol - 1) * ((1 + rsParFlujoCaja!nIncCompMerc / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(17, nCol - 1)))
            
            'Personal
            xlHoja1.Cells(18, nCol) = Round((xlHoja1.Cells(18, nCol - 1) * ((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(18, nCol - 1)))
            
            'calculo Alquiler de Locales
            xlHoja1.Cells(19, nCol) = "=C19"

            'calculo Alquiler de Equipos
            xlHoja1.Cells(20, nCol) = "=C20"

            'calculo Servicios (luz....)
            xlHoja1.Cells(21, nCol) = "=C21"

            'calculo Utiles de oficinas
            xlHoja1.Cells(22, nCol) = "=C22"

            'calculo Rep y Mtto de Equipos
            xlHoja1.Cells(23, nCol) = "=C23"

            'calculo Rep y Mtto de Vehiculo
            xlHoja1.Cells(24, nCol) = "=C24"

            'calculo Seguro
            xlHoja1.Cells(25, nCol) = "=C25"

            'calculo Transporte/Combustible/ Gas
            xlHoja1.Cells(26, nCol) = "=C26"

            'calculo Contador
            xlHoja1.Cells(27, nCol) = "=C27"

            'calculo Sunat + Impuestos
            xlHoja1.Cells(28, nCol) = "=C28"

            'calculo Publicidad y otros gastos de ventas (**Nuevo)
            xlHoja1.Cells(29, nCol) = Round((xlHoja1.Cells(29, nCol - 1) * ((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(29, nCol - 1)))
            
            'calculo Otros
            xlHoja1.Cells(30, nCol) = "=C30"

            'calculo Consumo Per.Nat.
            xlHoja1.Cells(31, nCol) = Round((xlHoja1.Cells(31, nCol - 1) * ((1 + rsParFlujoCaja!nIncConsu / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(31, nCol - 1)))

            'calculo Flujo Operativo
            xlHoja1.Range(xlHoja1.Cells(32, 2), xlHoja1.Cells(32, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(32, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "11" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "16)"
            
            'Cobro de Prestamo y dividendos
            xlHoja1.Cells(33, nCol) = 0

            'Pago de cuota Prestamos vigentes
            xlHoja1.Cells(34, nCol) = "=C34"

            'Pago de cuotas de prestamos solicitado
            xlHoja1.Cells(35, nCol) = "=C35"

            'calculo Flujo Financiero
            xlHoja1.Range(xlHoja1.Cells(36, 2), xlHoja1.Cells(36, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(36, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "32" & "+" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "33" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "34" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "35)"
            
            'Inversio
            xlHoja1.Cells(37, nCol) = 0
            
            'calculo Saldo
            xlHoja1.Cells(39, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "36" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "37)"
             'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(39, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(39, nCol), xlHoja1.Cells(39, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            'calculo Saldo Disponible
            If i >= 25 Then
                Z = Z + 1
            End If
            xlHoja1.Cells(40, nCol) = "=(" & IIf(i >= 25, MatLetrasRep(1, Z), MatAZ(1, i + 2)) & "41)"
             'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(40, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(40, nCol), xlHoja1.Cells(40, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            'calculo Saldo Acumulado
            xlHoja1.Cells(41, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "39" & "+" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "40)"
             'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(41, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(41, nCol), xlHoja1.Cells(41, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            
            nCol = nCol + 1

            If (i Mod 12) = 0 Then
                nColFin = nCol - 1
                    xlHoja1.Cells(8, nColInicio) = "Año" & (i / 12)
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).HorizontalAlignment = xlCenter
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).MergeCells = True
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).Font.Bold = True
                nColInicio = nColFin + 1
            End If
            rsCuotas.MoveNext
        Next i
        
        If nColInicio <> nCol Then
            'Para la celda si no cumple un año
            xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nCol - 1)).MergeCells = True
        End If
        xlHoja1.Range(xlHoja1.Cells(8, 4), xlHoja1.Cells(8, nCol - 1)).Cells.Interior.Color = RGB(141, 180, 226)
        
        CuadroExcel xlHoja1, 4, 8, nCol - 1, 8

        For i = 0 To 33
            If i <= 28 Then
                CuadroExcel xlHoja1, 4, 9 + i, nCol - 1, 9 + i
            ElseIf i >= 31 Then
                CuadroExcel xlHoja1, 4, nFila, nCol - 1, nFila
                nFila = nFila + 1
            End If
        Next i
        
    Else
        MsgBox "Error al crear el Excel, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm4 = False
        Exit Function
    End If

xlHoja1.Cells.Select
xlHoja1.Cells.Font.Name = "Arial"
xlHoja1.Cells.Font.Size = 9
xlHoja1.Cells.EntireColumn.AutoFit

'xlAplicacion.Worksheets("Hoja1").Protect ("123")

MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "!Exito!"

rs.Close
rsCabcera.Close
rsParFlujoCaja.Close
rsCuotas.Close


Exit Function 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
ErrorInicioExcel: 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
MsgBox Err.Description + "Error 2: Error al iniciar la creación del excel comunicar a TI", vbInformation, "Error" 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM


End Function

'JOEP20180725 ERS034-2018
Private Sub cmdMNME_Click()
    Call frmCredFormEvalCredCel.Inicio(ActXCodCta.NroCuenta, 11)
End Sub
'JOEP20180725 ERS034-2018

Private Sub EditMoneyIncCM4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditMoneyIncPP4.SetFocus
        fEnfoque EditMoneyIncPP4
    End If
End Sub

Private Sub EditMoneyIncGV4_KeyPress(KeyAscii As Integer)
'    EditMoneyIncC4.SetFocus
'    fEnfoque EditMoneyIncC4
End Sub

Private Sub EditMoneyIncPP4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    EditMoneyIncGV4.SetFocus
    fEnfoque EditMoneyIncGV4
End If
End Sub

Private Sub EditMoneyIncVC4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    EditMoneyIncCM4.SetFocus
    fEnfoque EditMoneyIncCM4
End If
End Sub
'Flujo de Caja


'_____________________________________________________________________________________________________________
'******************************************LUCV20160525: EVENTOS Varios***************************************
Private Sub Form_Load()
    fbGrabar = False
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    EnfocaControl spnTiempoLocalAnio
    'RECO20160916 *********************
    bActivoDetPP(0) = False
    bActivoDetPP(1) = False
    bActivoDetPE(0) = False
    bActivoDetPE(1) = False
    'RECO FIN *************************

'JOEP20180725 ERS034-2018
    If fnTipoRegMant = 3 Then
        If Not ConsultaRiesgoCamCred(sCtaCod) Then
            cmdMNME.Visible = True
        End If
    End If
'JOEP20180725 ERS034-2018
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
End Sub

Private Sub fePasivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim pnMonto As Double
Dim index As Integer
Dim nTotal As Double
      
    If fePasivos.TextMatrix(1, 0) = "" Then Exit Sub
    index = CInt(fePasivos.TextMatrix(fePasivos.row, 0))

    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
        Case 7, 9 '*************************-> Ctas x Cobrar
            Set oFrmCtaCobrarIfi = New frmCredFormEvalCtasCobrarIfis
            If fePasivos.Col = 2 Then 'column P.P.
                If IsArray(lvPrincipalPasivos(index).vPPActivoCtaCobrar) Then
                    lvDetalleActivosCtasCobrar = lvPrincipalPasivos(index).vPPActivoCtaCobrar
                Else
                    ReDim lvDetalleActivosCtasCobrar(0)
                End If
            End If
            If fePasivos.Col = 3 Then 'column P.E.
                If IsArray(lvPrincipalPasivos(index).vPEActivoCtaCobrar) Then
                    lvDetalleActivosCtasCobrar = lvPrincipalPasivos(index).vPEActivoCtaCobrar
                Else
                    ReDim lvDetalleActivosCtasCobrar(0)
                End If
            End If
            
            If oFrmCtaCobrarIfi.Inicio(lvDetalleActivosCtasCobrar, nTotal, CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5)), CInt(fePasivos.TextMatrix(Me.fePasivos.row, 6)), fePasivos.TextMatrix(Me.fePasivos.row, 1), ActXCodCta.NroCuenta, IIf(fePasivos.Col = 2, 1, 2)) Then
                If fePasivos.Col = 2 Then 'column P.P.
                    lvPrincipalPasivos(index).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                    'RECO20160916***************
                    'If UBound(lvPrincipalPasivos(Index).vPPActivoCtaCobrar) > 0 Then
                        Select Case index
                        Case 7
                            bActivoDetPP(0) = True
                        Case 9
                            bActivoDetPP(1) = True
                        End Select
                    'End If
                    'RECO FIN*******************
                End If
                If fePasivos.Col = 3 Then ' columna P.E.
                    lvPrincipalPasivos(index).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                    'RECO20160916***************
                    'If UBound(lvPrincipalPasivos(Index).vPEActivoCtaCobrar) > 0 Then
                        Select Case index
                        Case 7
                            bActivoDetPE(0) = True
                        Case 9
                            bActivoDetPE(1) = True
                        End Select
                    'End If
                    'RECO FIN*******************
                End If
            End If
            If fePasivos.Col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
            End If

            If fePasivos.Col = 3 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
            End If
              Call CalculoTotal(2)
            'Fin - Ctas x Cobrar <-**********
        End Select
End Sub

Private Sub Cmdguardar4_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim GrabarDatos As Boolean
    Dim rsGastoFam As ADODB.Recordset
    Dim rsOtrosIng As ADODB.Recordset
    Dim rsBalGen As ADODB.Recordset
    Dim rsFlujoCaja As ADODB.Recordset
    Dim rsPDT As ADODB.Recordset
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsRatiosActual As ADODB.Recordset
    Dim rsRatiosAceptableCritico As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    'feDeclaracionPDT.TextMatrix(0, 3) = "MesDetalle"
    feDeclaracionPDT.TextMatrix(0, 4) = "Mes1"
    feDeclaracionPDT.TextMatrix(0, 5) = "Mes2"
    feDeclaracionPDT.TextMatrix(0, 6) = "Mes3"
    feDeclaracionPDT.TextMatrix(0, 8) = "VentasDeclaradas"
    Set rsPDT = IIf(feDeclaracionPDT.rows - 1 > 0, feDeclaracionPDT.GetRsNew(), Nothing)
    'feDeclaracionPDT.TextMatrix(0, 3) = "Mes/Detalle"
    feDeclaracionPDT.TextMatrix(0, 4) = sMes3
    feDeclaracionPDT.TextMatrix(0, 5) = sMes2
    feDeclaracionPDT.TextMatrix(0, 6) = sMes1
    feDeclaracionPDT.TextMatrix(0, 8) = "%Vent. Decl."

    Set rsFlujoCaja = IIf(feFlujoCajaMensual.rows - 1 > 0, feFlujoCajaMensual.GetRsNew(), Nothing)
    Set rsOtrosIng = IIf(feOtrosIngresos.rows - 1 > 0, feOtrosIngresos.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing)
    'Flex a Matriz Referidos **********->
        ReDim MatReferidos(feReferidos.rows - 1, 6)
        For i = 1 To feReferidos.rows - 1
            MatReferidos(i, 1) = feReferidos.TextMatrix(i, 0)
            MatReferidos(i, 2) = feReferidos.TextMatrix(i, 1)
            MatReferidos(i, 3) = feReferidos.TextMatrix(i, 2)
            MatReferidos(i, 4) = feReferidos.TextMatrix(i, 3)
            MatReferidos(i, 5) = feReferidos.TextMatrix(i, 4)
            MatReferidos(i, 6) = feReferidos.TextMatrix(i, 5)
         Next i
    'Fin Referidos
    
    If ValidaDatos Then
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        If txtUltEndeuda.Text = "__/__/____" Then
            txtUltEndeuda.Text = "01/01/1900"
        End If

        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        Set objPista = New COMManejador.Pista
        
    If fnTipoPermiso = 3 Then
    '***************************************************************************
    'LUCV20160709, PARA EL LLENADO DEL FORMATO  4 ******************************
    '***************************************************************************
            'RECO20160730****************************************************
            '->*****Pasivos
            If IsArray(lvPrincipalPasivos(7).vPPActivoCtaCobrar) Then '->PP: 109-Parte Cte de deuda de LP
                'If UBound(lvPrincipalPasivos(7).vPPActivoCtaCobrar) = 0 Then 'RECO20160916
                If UBound(lvPrincipalPasivos(7).vPPActivoCtaCobrar) <= 0 And bActivoDetPP(0) = False Then
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 109, 1)
                    lvPrincipalPasivos(7).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(7).vPEActivoCtaCobrar) Then '->PE
                'If UBound(lvPrincipalPasivos(7).vPEActivoCtaCobrar) <= 0 Then
                If UBound(lvPrincipalPasivos(7).vPEActivoCtaCobrar) <= 0 And bActivoDetPE(0) = False Then  'RECO20160916
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 109, 2)
                    lvPrincipalPasivos(7).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(9).vPPActivoCtaCobrar) Then '->PE: 201-Deuda financiera de LP
                'If UBound(lvPrincipalPasivos(9).vPPActivoCtaCobrar) = 0  Then
                If UBound(lvPrincipalPasivos(9).vPPActivoCtaCobrar) <= 0 And bActivoDetPP(1) = False Then 'RECO20160916
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 201, 1)
                    lvPrincipalPasivos(9).vPPActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            If IsArray(lvPrincipalPasivos(9).vPEActivoCtaCobrar) Then '->PP
                'If UBound(lvPrincipalPasivos(9).vPEActivoCtaCobrar) = 0 Then
                If UBound(lvPrincipalPasivos(9).vPEActivoCtaCobrar) <= 0 And bActivoDetPE(1) = False Then 'RECO20160916
                    Call CargaMatrizDatosMantenimientoCtaCobrar(lvDetalleActivosCtasCobrar, ActXCodCta.NroCuenta, 7026, 201, 2)
                    lvPrincipalPasivos(9).vPEActivoCtaCobrar = lvDetalleActivosCtasCobrar
                End If
                ReDim lvDetalleActivosCtasCobrar(0)
            End If
            'RECO FIN *******************************************************
    'Eliminamos Datos existentes
        If UBound(lvPrincipalActivos) > 0 Or UBound(lvPrincipalPasivos) > 0 Then
              Call oDCOMFormatosEval.EliminaCredFormEvalGrillaActiPasi(sCtaCod, fnFormato)
            Call oDCOMFormatosEval.EliminaCredFormEvalGrillaActiPasiDet(sCtaCod, fnFormato)
        End If

        '------ ACTIVOS(CredFormEvalActivoPasivo / CredFormEvalActivoPasivoDet)
            If UBound(lvPrincipalActivos) > 0 Then
                For i = 1 To UBound(lvPrincipalActivos)
                    'If CDbl(Me.feActivos.TextMatrix(i, 4)) > 0 Then
                        If i = 18 Then
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                            CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CCur(Me.feActivos.TextMatrix(i, 4)), _
                            CCur(Me.feActivos.TextMatrix(i, 2)), CCur(Me.feActivos.TextMatrix(i, 3)))
                        Else
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiDet(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                            CInt(Me.feActivos.TextMatrix(i, 5)), CInt(Me.feActivos.TextMatrix(i, 6)), CCur(Me.feActivos.TextMatrix(i, 4)), _
                            CCur(Me.feActivos.TextMatrix(i, 2)), CCur(Me.feActivos.TextMatrix(i, 3)))
                        End If
                    'End If
                Next i
            End If
        
        '--- PASIVOS (CredFormEvalActivoPasivo / CredFormEvalActivoPasivoDet)
            If UBound(lvPrincipalPasivos) > 0 Then
                For i = 1 To UBound(lvPrincipalPasivos)
                    'If CDbl(Me.fePasivos.TextMatrix(i, 4)) > 0 Then
                        If (i = 13) Then
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                            CInt(Me.fePasivos.TextMatrix(i, 5)), gCodTotalPatrimonio, CDbl(Me.fePasivos.TextMatrix(i, 4)), _
                            CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)))
                        End If
                    
                        If (i = 20) Or (i = 21) Then
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasi(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                            CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CDbl(Me.fePasivos.TextMatrix(i, 4)), _
                            CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)))
                        Else
                            Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiDet(sCtaCod, fnFormato, Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                            CInt(Me.fePasivos.TextMatrix(i, 5)), CInt(Me.fePasivos.TextMatrix(i, 6)), CDbl(Me.fePasivos.TextMatrix(i, 4)), _
                            CDbl(Me.fePasivos.TextMatrix(i, 2)), CDbl(Me.fePasivos.TextMatrix(i, 3)))
                        End If
                    'End If
                Next i
            End If
            
            '---------------------- PASIVOS -> Detalle Celdas (PP / PE)
            
        If UBound(lvPrincipalActivos) > 0 Then
           For i = 1 To UBound(lvPrincipalPasivos)
                   'Detalle de Celdas -> Cuentas por Cobrar ********************->
                       If IsArray(lvPrincipalPasivos(i).vPPActivoCtaCobrar) Then 'CtasxCobrar->PP
                           For j = 1 To UBound(lvPrincipalPasivos(i).vPPActivoCtaCobrar)
                               Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiCtaCobrar(sCtaCod, fnFormato, _
                                                                                       gCodPatrimonioPersonal, j, _
                                                                                       CInt(Me.fePasivos.TextMatrix(i, 5)), _
                                                                                       CInt(Me.fePasivos.TextMatrix(i, 6)), _
                                                                                       Format(lvPrincipalPasivos(i).vPPActivoCtaCobrar(j).dFecha, "yyyyMMdd"), _
                                                                                       lvPrincipalPasivos(i).vPPActivoCtaCobrar(j).cCtaporCobrar, _
                                                                                       lvPrincipalPasivos(i).vPPActivoCtaCobrar(j).nTotal)
                           Next j
                       End If
    
                       If IsArray(lvPrincipalPasivos(i).vPEActivoCtaCobrar) Then 'CtasxCobrar->PE
                           For j = 1 To UBound(lvPrincipalPasivos(i).vPEActivoCtaCobrar)
                               Call oDCOMFormatosEval.InsertaCredFormEvalActiPasiCtaCobrar(sCtaCod, fnFormato, _
                                                                                       gCodPatrimonioEmpresarial, j, _
                                                                                       CInt(Me.fePasivos.TextMatrix(i, 5)), _
                                                                                       CInt(Me.fePasivos.TextMatrix(i, 6)), _
                                                                                       Format(lvPrincipalPasivos(i).vPEActivoCtaCobrar(j).dFecha, "yyyyMMdd"), _
                                                                                       lvPrincipalPasivos(i).vPEActivoCtaCobrar(j).cCtaporCobrar, _
                                                                                       lvPrincipalPasivos(i).vPEActivoCtaCobrar(j).nTotal)
                           Next j
                       End If
            Next i
        End If
    'Fin <- ********** LUCV20160709
    
   'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Dim MatFlujoCaja As Variant
    Set MatFlujoCaja = Nothing
    ReDim MatFlujoCaja(1, 5)
        For i = 1 To 1
            MatFlujoCaja(i, 1) = EditMoneyIncVC4
            MatFlujoCaja(i, 2) = EditMoneyIncCM4
            MatFlujoCaja(i, 3) = EditMoneyIncPP4
            MatFlujoCaja(i, 4) = EditMoneyIncGV4
            MatFlujoCaja(i, 5) = EditMoneyIncC4
        Next i
   'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        GrabarDatos = oNCOMFormatosEval.GrabarCredFormEvalFormato1_5(sCtaCod, fnFormato, fnTipoRegMant, _
                                                                    Trim(txtGiroNeg.Text), CInt(spnExpEmpAnio.valor), CInt(spnExpEmpMes.valor), CInt(spnTiempoLocalAnio.valor), _
                                                                    CInt(spnTiempoLocalMes.valor), CDbl(txtUltEndeuda.Text), Format(txtFecUltEndeuda.Text, "yyyymmdd"), _
                                                                    fnCondLocal, IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text), CDbl(txtExposicionCredito.Text), _
                                                                    Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                                    Format(txtFechaVisita.Text, "yyyymmdd"), _
                                                                    txtEntornoFamiliar4.Text, txtGiroUbicacion4.Text, _
                                                                    txtExperiencia4.Text, txtFormalidadNegocio4.Text, _
                                                                    txtColaterales4, txtDestino4.Text, _
                                                                    txtComentario4.Text, MatReferidos, MatIfiGastoNego, MatIfiGastoFami, _
                                                                    rsGastoFam, rsOtrosIng, _
                                                                    , , , , , , _
                                                                    rsFlujoCaja, rsPDT, _
                                                                    gRatioCapacidadPago, _
                                                                    CDbl(Replace(txtCapacidadNeta.Text, "%", "")), _
                                                                    gRatioEndeudamiento, _
                                                                    CDbl(Replace(txtEndeudamiento.Text, "%", "")), _
                                                                    gRatioIngresoNetoNego, _
                                                                    CDbl(txtIngresoNeto.Text), _
                                                                    gRatioExcedenteMensual, _
                                                                    CDbl(txtExcedenteMensual.Text), _
                                                                    nMes1, nMes2, nMes3, nAnio1, nAnio2, nAnio3, fnColocCondi, MatFlujoCaja)
                                                                    
                                                                    'JOEP20171015 Flujo de Caja MatFlujoCaja
                                                                    
            Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod)
            Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
            Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod)
            'JOEP20180725 ERS034-2018
            Call EmiteFormRiesgoCamCred(sCtaCod)
'JOEP20180725 ERS034-2018
        ' Else
        'GrabarDatos = oNCOMFormatosEval.GrabarCredEvaluacionVerif(sCtaCod, Trim(txtVerif.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
            If GrabarDatos Then
                fbGrabar = True
                'RECO20161020 ERS060-2016 **********************************************************
                Dim oNCOMColocEval As New NCOMColocEval
                'Dim lcMovNro As String 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                
                If Not ValidaExisteRegProceso(sCtaCod, gTpoRegCtrlEvaluacion) Then
                   'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                   'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 4", sCtaCod, gCodigoCuenta 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                   Call oNCOMColocEval.insEstadosExpediente(sCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
                   Set oNCOMColocEval = Nothing
                End If
                'RECO FIN **************************************************************************
                If fnTipoRegMant = 1 Then
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 4", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
                Else
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 4", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                End If
                
                'Habilita / Deshabilita Botones - Text
                 If fnEstado = 2000 Then                '*****-> Si es Solicitado
                    If fnColocCondi <> 4 Then
                        Me.cmdInformeVisita4.Enabled = True
                        Me.cmdVerCar.Enabled = False
                    Else
                        Me.cmdInformeVisita4.Enabled = False
                        Me.cmdVerCar.Enabled = False
                    End If
                    Me.cmdGuardar4.Enabled = False
                    Me.cmdImprimir.Enabled = False
                Else                                    '*****-> Sugerido +
                    Me.cmdImprimir.Enabled = True
                    Me.cmdGuardar4.Enabled = False
                    If fnColocCondi <> 4 Then
                        Me.cmdVerCar.Enabled = True    'No refinanciado
                        Me.cmdInformeVisita4.Enabled = True
                    Else
                        Me.cmdVerCar.Enabled = False
                        Me.cmdInformeVisita4.Enabled = False
                    End If
                End If
                
                '*****->No Refinanciados (Propuesta Credito)
                If fnColocCondi <> 4 Then
                    txtFechaVisita.Enabled = True
                    txtEntornoFamiliar4.Enabled = True
                    txtGiroUbicacion4.Enabled = True
                    txtExperiencia4.Enabled = True
                    txtFormalidadNegocio4.Enabled = True
                    txtColaterales4.Enabled = True
                    txtDestino4.Enabled = True
                 Else
                    framePropuesta.Enabled = False
                    txtFechaVisita.Enabled = False
                    txtEntornoFamiliar4.Enabled = False
                    txtGiroUbicacion4.Enabled = False
                    txtExperiencia4.Enabled = False
                    txtFormalidadNegocio4.Enabled = False
                    txtColaterales4.Enabled = False
                    txtDestino4.Enabled = False
                End If
                '*****->Fin No Refinanciados
                
                'Actualizacion de los Ratios
                    txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"
                    txtEndeudamiento.Text = CStr(rsRatiosActual!nEndeuPat * 100) & "%"
                    txtLiquidezCte.Text = CStr(Format(rsRatiosActual!nLiquidezCte, "#0.00"))
                    txtRentabilidadPat.Text = CStr(rsRatiosActual!nRentaPatri * 100) & "%"
                    txtIngresoNeto.Text = Format(rsRatiosActual!nIngreNeto, "#,##0.00")
                    txtExcedenteMensual.Text = Format(rsRatiosActual!nExceMensual, "#,##0.00")
                    
                'Ratios: Aceptable / Critico ->*****
                    If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                    If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                        Me.lblCapaAceptable.Caption = "Aceptable"
                        Me.lblCapaAceptable.ForeColor = &H8000&
                    Else
                        Me.lblCapaAceptable.Caption = "Crítico"
                        Me.lblCapaAceptable.ForeColor = vbRed
                    End If
                    
                    If rsRatiosAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
                        Me.lblEndeAceptable.Caption = "Aceptable"
                        Me.lblEndeAceptable.ForeColor = &H8000&
                    Else
                        Me.lblEndeAceptable.Caption = "Crítico"
                        Me.lblEndeAceptable.ForeColor = vbRed
                    End If
                    Else
                        lblCapaAceptable.Visible = False
                        lblEndeAceptable.Visible = False
                    End If
                'Fin Ratios <-****
                    Set rsRatiosActual = Nothing
                    Set rsRatiosAceptableCritico = Nothing
            Else
                MsgBox "Hubo errores al grabar la información", vbError, "Error"
            End If
    'Else
    'MsgBox "Ha Ocurrido un Problema o Faltan Ingresar Datos", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdImprimir_Click()
    Call ImprimirFormatoEvaluacion
End Sub
Private Sub cmdVerCar_Click()
    Call GeneraVerCar
End Sub
Private Sub cmdCancelar4_Click()
    Unload frmCredFormEvalCuotasIfis
    Unload Me
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
End Sub
Private Sub cmdInformeVisita4_Click()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
       
    cmdInformeVisita4.Enabled = False
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atención"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes
    Set rsInfVisita = Nothing
    cmdInformeVisita4.Enabled = True
End Sub
Private Sub cmdAgregarRef_Click()
    If feReferidos.rows - 1 < 25 Then
        feReferidos.lbEditarFlex = True
        feReferidos.AdicionaFila
        feReferidos.SetFocus
        feReferidos.AvanceCeldas = Horizontal
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdQuitar4_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feReferidos.EliminaFila (feReferidos.row)
    End If
End Sub

'LUCV20160620, KeyPress / GotFocus / LostFocus ->**********
    'TAB0 -> Ingresos/Egresos
Private Sub spnTiempoLocalAnio_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        spnTiempoLocalMes.SetFocus
    End If
End Sub
Private Sub spnTiempoLocalMes_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        OptCondLocal(1).SetFocus
    End If
End Sub
Private Sub OptCondLocal_KeyPress(index As Integer, KeyAscii As Integer) 'CondicionLocal
    If KeyAscii = 13 Then
        feActivos.row = 2
        feActivos.Col = 2
        EnfocaControl feActivos
        SSTabIngresos.Tab = 0
        SendKeys "{Enter}"
        
    End If
End Sub

Private Sub txtCondLocalOtros_KeyPress(KeyAscii As Integer) 'OtroCondicionLocal
    If KeyAscii = 13 Then
        feActivos.row = 2
        feActivos.Col = 2
        EnfocaControl feActivos
        SSTabIngresos.Tab = 0
        SendKeys "{Enter}"
    End If
End Sub

'TAB1 ->PropuestaCredito
Private Sub txtFechaVisita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtEntornoFamiliar4.SetFocus
        If Not IsDate(txtFechaVisita) Then
            MsgBox "Verifique Dia,Mes,Año , Fecha Incorrecta", vbInformation, "Aviso"
            txtFechaVisita.SetFocus
        End If
    End If
End Sub

Private Sub txtEntornoFamiliar4_KeyPress(KeyAscii As Integer) 'Entornofamiliar
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
    txtGiroUbicacion4.SetFocus
    End If
End Sub
Private Sub txtGiroUbicacion4_KeyPress(KeyAscii As Integer) 'SobreGiro
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtExperiencia4.SetFocus
    End If
End Sub
Private Sub txtExperiencia4_KeyPress(KeyAscii As Integer) 'ExperienciaCrediticia
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtFormalidadNegocio4.SetFocus
    End If
End Sub
Private Sub txtFormalidadNegocio4_KeyPress(KeyAscii As Integer) 'ConsistenciaInformacion
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
    txtColaterales4.SetFocus
    End If
End Sub
Private Sub txtColaterales4_KeyPress(KeyAscii As Integer) 'Colaterales_Garantias
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtDestino4.SetFocus
    End If
End Sub
Private Sub txtDestino4_KeyPress(KeyAscii As Integer) 'Destino del crédito
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 3
        'If fnColocCondi = 1 Then 'LUCV20171115, Agregó segun correo: RUSI
        If Not fbTieneReferido6Meses Then
            txtComentario4.SetFocus
        Else
            cmdGuardar4.SetFocus
        End If
    End If
End Sub
    'TAB1 ->ComentarioReferido
Private Sub txtComentario4_KeyPress(KeyAscii As Integer) 'Referidos/ ComentariosReferidos
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 3
        If fnColocCondi = 1 Then
            cmdAgregarRef.SetFocus
        End If
    End If
End Sub
'LUCV20160620, KeyPress / GotFocus / LostFocus Fin <-**********

'Calcular Activos / Pasivos / FlujoCajaMensual
Private Sub feActivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    'CalculoTotal (1)
    Dim Editar() As String
    Editar = Split(Me.feActivos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub

Private Sub feActivos_RowColChange()
    If feActivos.Col = 2 Then
        feActivos.AvanceCeldas = Horizontal
    ElseIf feActivos.Col = 3 Then
        feActivos.AvanceCeldas = Vertical
    End If
    
    Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
     Case 1000, 100, 200, 300, 400, 500
         Me.feActivos.BackColorRow (&H80000000)
         Me.feActivos.ForeColorRow vbBlack, True
         Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
     Case Else
         Me.feActivos.BackColorRow (&HFFFFFF)
         Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
     End Select
     CalculoTotal (1)
End Sub
Private Sub feActivos_KeyPress(KeyAscii As Integer)
    If (Me.feActivos.Col = 4 And Me.feActivos.row = 18) Then
        Me.fePasivos.SetFocus
        fePasivos.row = 2
        fePasivos.Col = 2
        SendKeys "{f2}"
    End If
End Sub
'Private Sub feActivos_GotFocus()
'    If (Me.feActivos.Col = 3 And Me.feActivos.row = 18) Then
'        Me.fePasivos.SetFocus
'        fePasivos.row = 2
'        fePasivos.Col = 2
'        SendKeys "{TAB}"
'    End If
'End Sub

Private Sub feActivos_EnterCell()
'    If feActivos.Col = 2 Or feActivos.Col = 3 Then
'    Select Case CInt(feActivos.TextMatrix(feActivos.row, 6))
'    Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405
'      feActivos.ListaControles = "0-0-1-1-0-0-0"
'    Case Else
'      feActivos.ListaControles = "0-0-0-0-0-0-0"
'    End Select
'    End If
End Sub
Private Sub feActivos_Click() 'GastosFamiliares
'    If feActivos.Col = 2 Or feActivos.Col = 3 Then
'    Select Case CInt(feActivos.TextMatrix(feActivos.row, 6))
'    Case 102, 106, 301, 302, 303, 304, 401, 402, 403, 404, 405
'      feActivos.ListaControles = "0-0-1-1-0-0-0"
'    Case Else
'      feActivos.ListaControles = "0-0-0-0-0-0-0"
'    End Select
'    End If
End Sub

Private Sub feActivos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feActivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feActivos.TextMatrix(pnRow, pnCol) < 0 Then
            feActivos.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    Else
        feActivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    

    If (Me.feActivos.Col = 3 And Me.feActivos.row = 17) Then
        Me.fePasivos.SetFocus
        SSTabIngresos.Tab = 0
        fePasivos.row = 2
        fePasivos.Col = 2
        SendKeys "{TAB}"
    End If
    CalculoTotal (1)
End Sub

Private Sub fePasivos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.fePasivos.ColumnasAEditar, "-")
    If Me.fePasivos.row <> 2 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
    Call CalculoTotal(2)
End Sub
Private Sub fePasivos_RowColChange()
    If fePasivos.Col = 2 Or fePasivos.Col = 3 Then 'LUCV20160725
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
            Case 109, 201
              fePasivos.ListaControles = "0-0-1-1-0-0-0"
            Case Else
              fePasivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
    End If
    
    If fePasivos.Col = 2 Then
        fePasivos.AvanceCeldas = Horizontal
    ElseIf fePasivos.Col = 3 Then
        fePasivos.AvanceCeldas = Vertical
    End If
     
    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 1001, 1002, 100, 200, 300, 400, 500
            Me.fePasivos.BackColorRow (&H80000000)
            Me.fePasivos.ForeColorRow vbBlack, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 206
            Me.fePasivos.BackColorRow vbWhite, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 109, 201
            Me.fePasivos.BackColorRow &HC0FFFF, True
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case 301
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fePasivos.BackColorRow (&HFFFFFF)
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
    Call CalculoTotal(2)
End Sub
Private Sub fePasivos_EnterCell()
    
    If fePasivos.Col = 2 Or fePasivos.Col = 3 Then 'LUCV20160725
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
            Case 109, 201
              fePasivos.ListaControles = "0-0-1-1-0-0-0"
            Case Else
              fePasivos.ListaControles = "0-0-0-0-0-0-0"
        End Select
    End If
    
    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 1001, 1002, 100, 200, 300, 400, 500
            Me.fePasivos.BackColorRow (&H80000000)
            Me.fePasivos.ForeColorRow vbBlack, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 206
            Me.fePasivos.BackColorRow vbWhite, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 301
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 109, 201
            Me.fePasivos.BackColorRow &HC0FFFF, True
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case Else
            Me.fePasivos.BackColorRow (&HFFFFFF)
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
End Sub
Private Sub fePasivos_KeyPress(KeyAscii As Integer)
    If (Me.fePasivos.Col = 4 And Me.fePasivos.row = 21) Then
        SSTabIngresos.Tab = 1
        Me.feFlujoCajaMensual.SetFocus
        feFlujoCajaMensual.row = 1
        feFlujoCajaMensual.Col = 4
        SendKeys "{F2}"
    End If
End Sub
Private Sub fePasivos_Click()
    If fePasivos.Col = 2 Or fePasivos.Col = 3 Then 'LUCV20160725
    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
        Case 109, 201
          fePasivos.ListaControles = "0-0-1-1-0-0-0"
        Case Else
          fePasivos.ListaControles = "0-0-0-0-0-0-0"
    End Select
    End If

    Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
        Case 1000, 1001, 1002, 100, 200, 300, 400, 500
            Me.fePasivos.BackColorRow (&H80000000)
            Me.fePasivos.ForeColorRow vbBlack, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 206
            Me.fePasivos.BackColorRow vbWhite, True
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case 109, 201
            Me.fePasivos.BackColorRow &HC0FFFF, True
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        Case 301
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fePasivos.BackColorRow (&HFFFFFF)
            Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
End Sub

Private Sub fePasivos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(fePasivos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6))
            Case 305, 306 'Valores Negativos
                 fePasivos.TextMatrix(pnRow, pnCol) = Format((CCur(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")
            Case Else 'Valores Positivos
                If fePasivos.TextMatrix(pnRow, pnCol) < 0 Then
                  fePasivos.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(fePasivos.TextMatrix(pnRow, pnCol))), "#,#0.00")
                End If
         End Select
    Else
        fePasivos.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    
    If (Me.fePasivos.Col = 3 And Me.fePasivos.row = 19) Then
        SSTabIngresos.Tab = 1
        Me.feFlujoCajaMensual.SetFocus
        feFlujoCajaMensual.row = 1
        feFlujoCajaMensual.Col = 4
        SendKeys "{TAB}"
    End If
    
    Call CalculoTotal(2)
End Sub

'Para Buscar Cuotas IFIs (Flujo Caja)**********->
Private Sub feFlujoCajaMensual_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    'CalculoTotal (3)
    Dim Editar() As String
    Editar = Split(Me.feFlujoCajaMensual.ColumnasAEditar, "-")
    If Me.feFlujoCajaMensual.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
             Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub feFlujoCajaMensual_KeyPress(KeyAscii As Integer)
    'If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 20) Then'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 22) Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        SSTabIngresos.Tab = 1
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.Col = 3
        SendKeys "{Enter}"
    End If
End Sub
Private Sub feFlujoCajaMensual_Click() 'GastosNegocio
    If feFlujoCajaMensual.Col = 4 Then
        If CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiFlujoCaja Then
            feFlujoCajaMensual.ListaControles = "0-0-0-0-1-0"
        Else
            feFlujoCajaMensual.ListaControles = "0-0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
        'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.BackColorRow (&H80000000)
            Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
            Me.feFlujoCajaMensual.BackColorRow vbWhite, True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 19 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            'Me.feFlujoCajaMensual.CellBackColor = (&HC0FFFF)
            Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
        Case Else
            Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
    End Select
End Sub

Private Sub feFlujoCajaMensual_EnterCell() 'LUCV20160525 - Me permite Buscar OtrasCuotasIFIs (GastosNegocio)
    If feFlujoCajaMensual.Col = 4 Then
        If CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiFlujoCaja Then
            feFlujoCajaMensual.ListaControles = "0-0-0-0-1-0"
        Else
            feFlujoCajaMensual.ListaControles = "0-0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
    'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Me.feFlujoCajaMensual.BackColorRow (&H80000000)
        Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
    'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
        Me.feFlujoCajaMensual.BackColorRow vbWhite, True
        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
    'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Case 19 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        'Me.feFlujoCajaMensual.CellBackColor = (&HC0FFFF)
        Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
        Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
    Case Else
        Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
    End Select
End Sub

Private Sub feFlujoCajaMensual_RowColChange() 'PresionarEnter:Monto
    If feFlujoCajaMensual.Col = 4 Then
        feFlujoCajaMensual.AvanceCeldas = Vertical
    Else
        feFlujoCajaMensual.AvanceCeldas = Horizontal
    End If
    
    If feFlujoCajaMensual.Col = 4 Then
        If CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 0)) = gCodCuotaIfiFlujoCaja Then
            feFlujoCajaMensual.ListaControles = "0-0-0-0-1-0"
        Else
            feFlujoCajaMensual.ListaControles = "0-0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
        'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.BackColorRow (&H80000000)
            Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
            Me.feFlujoCajaMensual.BackColorRow vbWhite, True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
        'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 19 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
            Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
        Case Else
            Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
            Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
    End Select
End Sub
Private Sub feFlujoCajaMensual_OnClickTxtBuscar(psMontoIfiGastoNego As String, psDescripcion As String) 'Fujo Caja Mensual
    psMontoIfiGastoNego = 0
    psDescripcion = ""
    psDescripcion = feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3) 'Cuotas Otras IFIs
    psMontoIfiGastoNego = feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4) 'Monto
    
    If psMontoIfiGastoNego = 0 Then
        Set MatIfiGastoNego = Nothing
        fnTotalRefGastoNego = 0
        frmCredFormEvalCuotasIfis.Inicio (CCur(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4))), fnTotalRefGastoNego, MatIfiGastoNego, feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3)
        psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CCur(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 4))), fnTotalRefGastoNego, MatIfiGastoNego, feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 3)
        psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
    End If
End Sub
Private Sub feFlujoCajaMensual_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feFlujoCajaMensual.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feFlujoCajaMensual.TextMatrix(pnRow, pnCol) < 0 Then
            feFlujoCajaMensual.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    Else
        feFlujoCajaMensual.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    
    
    'If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 19) Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If (Me.feFlujoCajaMensual.Col = 4 And Me.feFlujoCajaMensual.row = 21) Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        SSTabIngresos.Tab = 1
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.Col = 3
        SendKeys "{TAB}"
    End If
    Call CalculoTotal(3)
    Call CalculoTotal(4)
End Sub
Private Sub feGastosFamiliares_KeyPress(KeyAscii As Integer)
        If (feGastosFamiliares.Col = 1 And feGastosFamiliares.row = 1) Or (feGastosFamiliares.Col = 3 And feGastosFamiliares.row = 7) Then
        If KeyAscii = 13 Then
            feOtrosIngresos.row = 1
            feOtrosIngresos.Col = 3
            EnfocaControl feOtrosIngresos
            SendKeys "{Enter}", True
        End If
    End If
End Sub

Private Sub feGastosFamiliares_Click() 'GastosFamiliares
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite, True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosFamiliares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosFamiliares.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub
Private Sub feGastosFamiliares_EnterCell() 'LUCV20160525 - Me permite Buscar CuotasIFIs(GastosFamiliares)
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite, True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosFamiliares_RowColChange() 'PresionarEnter:Monto
    If feGastosFamiliares.Col = 3 Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami Then
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If

    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
        Case gCodCuotaIfiGastoFami
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ForeColorRow (&H80000007), True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.BackColorRow vbWhite, True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Case Else
            Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String) 'GastosFamiliares
    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
    If psMontoIfiGastoFami = 0 Then
        fnTotalRefGastoFami = 0
        Set MatIfiGastoFami = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CCur(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2)
        psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CCur(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2)
        psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
    End If
End Sub
Private Sub feGastosFamiliares_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosFamiliares.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosFamiliares.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
    End If
End Sub

Private Sub OptCondLocal_Click(index As Integer)
    Select Case index
    Case 1, 2, 3
        Me.txtCondLocalOtros.Visible = False
        Me.txtCondLocalOtros.Text = ""
    Case 4
        Me.txtCondLocalOtros.Visible = True
        Me.txtCondLocalOtros.Text = ""
    End Select
    fnCondLocal = index
End Sub

'***** LUCV20160528 - OnCellChange / RowColChange
Private Sub feReferidos_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Or pnCol = 4 Then
        feReferidos.TextMatrix(pnRow, pnCol) = UCase(feReferidos.TextMatrix(pnRow, pnCol))
    End If
    
    Select Case pnCol
    Case 2
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                    Case Is > 0
                    Case Else
                        MsgBox "Por favor, verifique el DNI", vbInformation, "Alerta"
                        feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "El DNI, tiene que ser 8 dígitos.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
            
        Else
            MsgBox "El DNI, tiene que ser numérico.", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 3
        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 9 Then
                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Teléfono Mal Ingresado", vbInformation, "Alerta"
                    feReferidos.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "Faltan caracteres en el teléfono / celular.", vbInformation, "Alerta"
                feReferidos.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "El telefono, solo permite ingreso de datos tipo numérico." & Chr(10) & "Ejemplo: 065404040, 984047523 ", vbInformation, "Alerta"
            feReferidos.TextMatrix(pnRow, pnCol) = 0
        End If
'    Case 5
'        If IsNumeric(feReferidos.TextMatrix(pnRow, pnCol)) Then
'            If Len(feReferidos.TextMatrix(pnRow, pnCol)) = 8 Then
'                Select Case CCur(feReferidos.TextMatrix(pnRow, pnCol))
'                Case Is > 0
'                Case Else
'                    MsgBox "El DNI del referido, tiene que contener 8 dígitos", vbInformation, "Alerta"
'                    feReferidos.TextMatrix(pnRow, pnCol) = 0
'                    Exit Sub
'                End Select
'            Else
'                MsgBox "El DNI del referido, tiene que ser 8 dígitos", vbInformation, "Alerta"
'                feReferidos.TextMatrix(pnRow, pnCol) = 0
'            End If
'        Else
'            MsgBox "El DNI del referido, sólo permite ingreso de datos tipo numérico.", vbInformation, "Alerta"
'            feReferidos.TextMatrix(pnRow, pnCol) = 0
'        End If
    End Select
End Sub

Private Sub feReferidos_RowColChange()
    If feReferidos.Col = 1 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.Col = 2 Then
        feReferidos.MaxLength = "8"
    ElseIf feReferidos.Col = 3 Then
        feReferidos.MaxLength = "9"
    ElseIf feReferidos.Col = 4 Then
        feReferidos.MaxLength = "200"
    ElseIf feReferidos.Col = 5 Then
        feReferidos.MaxLength = "8"
    End If
End Sub

Private Sub feOtrosIngresos_RowColChange() 'PresionarEnter:Monto
    If feOtrosIngresos.Col = 3 Then
        feOtrosIngresos.AvanceCeldas = Vertical
    Else
        feOtrosIngresos.AvanceCeldas = Horizontal
    End If
End Sub
Private Sub feOtrosIngresos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feOtrosIngresos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feOtrosIngresos.TextMatrix(pnRow, pnCol) < 0 Then
            feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
    End If
    If Me.feOtrosIngresos.Col = 3 And Me.feOtrosIngresos.row = 5 Then
        Me.SSTabIngresos.Tab = 1
        Me.feDeclaracionPDT.SetFocus
        Me.feDeclaracionPDT.row = 1
        Me.feDeclaracionPDT.Col = 4
        SendKeys "{TAB}"
   End If
    
End Sub
Private Sub feDeclaracionPDT_KeyPress(KeyAscii As Integer)
 If (Me.feDeclaracionPDT.Col = 7 And Me.feDeclaracionPDT.row = 2) Or (Me.feDeclaracionPDT.Col = 8 And Me.feDeclaracionPDT.row = 2) Then
        Me.SSTabIngresos.Tab = 2
        If txtFechaVisita.Enabled Then  'ARLO20190330
            Me.txtFechaVisita.SetFocus
        End If
        SendKeys "{TAB}"
   End If
End Sub
Private Sub feDeclaracionPDT_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    'Call CalculoTotal(4)
End Sub
Private Sub feDeclaracionPDT_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feDeclaracionPDT.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feDeclaracionPDT.TextMatrix(pnRow, pnCol) < 0 Then
            feDeclaracionPDT.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    Else
        feDeclaracionPDT.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    
    If Me.feDeclaracionPDT.Col = 6 And Me.feDeclaracionPDT.row = 2 Then
        Me.SSTabIngresos.Tab = 2
        If txtFechaVisita.Enabled Then  'ARLO20190330
            Me.txtFechaVisita.SetFocus
        End If
        SendKeys "{TAB}"
   End If
    
    Call CalculoTotal(4)
End Sub
Private Sub feDeclaracionPDT_Click()
    Call CalculoTotal(4)
End Sub
'Fin <- LUCV20160528 - OnCellChange / RowColChange *****

'________________________________________________________________________________________________________________________
'*************************************************LUCV20160525: METODOS Varios **************************************************
Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                     ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer) As Boolean
 
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio

    If psCtaCod <> -1 Then 'CtaCod -> **********
        gsOpeCod = ""
        lcMovNro = "" 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
        sCtaCod = psCtaCod
        fnTipoRegMant = psTipoRegMant
        ActXCodCta.NroCuenta = sCtaCod
        
        '(3: Analista, 2: Coordinador, 1: JefeAgencia)
        fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo)  ' Obtener el tipo de Permiso, Segun Cargo
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
        
        If (rsDCredito!cActiGiro) = "" Then
            MsgBox "Por favor, actualizar los datos del cliente. " & Chr(13) & " (Actividad o Giro del negocio)", vbInformation, "Alerta"
            Exit Function
        End If
        
        '*****-> Datos básicos de cabecera de formato
        fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
        fnColocCondi = rsDCredito!nColocCondicion
        fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses   'Si tiene evaluacion registrada 6 meses (LUCV20171115, agregó según correo: RUSI)
        fsCliente = Trim(rsDCredito!cPersNombre)
        fsAnioExp = CInt(rsDCredito!nAnio)
        fsMesExp = CInt(rsDCredito!nMes)
        fnFechaDeudaSbs = IIf(rsDCredito!dFechaUltimaDeudaSBS = "", "__/__/____", rsDCredito!dFechaUltimaDeudaSBS)
        fnMontoDeudaSbs = CCur(rsDCredito!nMontoUltimaDeudaSBS)
        
        spnExpEmpAnio.valor = fsAnioExp
        spnExpEmpMes.valor = fsMesExp
        txtUltEndeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
        txtFecUltEndeuda.Text = Format(fnFechaDeudaSbs, "dd/mm/yyyy")
        txtExposicionCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
        txtFechaEvaluacion.Text = Format(gdFecSis, "dd/mm/yyyy")
        '<-***** Fin datos de cabecera
        
        Set rsDCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(sCtaCod) 'Ojo: Recuperar Credito Si ha sido Registrado el Form. Eval.
        Set rsAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod) 'Obtenemos Datos, Aceptable / Critico de los Ratios
        If fnTipoPermiso = 2 Then
           If rsDCredEval.RecordCount = 0 Then ' Si no hay credito registrado
                MsgBox "El analista no ha registrado la Evaluacion respectiva", vbInformation, "Aviso"
                fbPermiteGrabar = False
            Else
                fbPermiteGrabar = True
             End If
        End If
        
        Set rsDCredito = Nothing
        Set rsDCredEval = Nothing
        
        fnFormato = pnFormato
        fnEstado = pnEstado
        SSTabIngresos.Tab = 0
        fbPermiteGrabar = False
        fbBloqueaTodo = False
        frameLinea.Visible = False
    Else
        MsgBox "No se ha registrado el número de cuenta del crédito a evaluar ", vbInformation, "Aviso"
    End If 'Fin CtaCod <-*****
    
    Set oDCOMFormatosEval = Nothing
    Set oTipoCam = Nothing
    Call CargaControlesInicio
    
    If fnTipoRegMant = 3 Then
        fbBloqueaTodo = True
        'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        gsOpeCod = gCredConsultarEvaluacionCred
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 4", sCtaCod, gCodigoCuenta
        Set objPista = Nothing
        'Fin LUCV20181220
    End If
    
    'Carga de Datos Segun Evento: (Registrar / Mantenimiento) *****->
    If CargaDatos Then
        If CargaControlesTipoPermiso(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
            If fnTipoRegMant = 1 Then   'Para el Evento: "Registrar"
                If Not rsCredEval.EOF Then
                    Call Mantenimiento
                    fnTipoRegMant = 2
                Else
                    Call Registro
                    fnTipoRegMant = 1
                End If
            ElseIf fnTipoRegMant = 2 Then 'Para el Evento. "Mantenimiento"
                If rsCredEval.EOF Then
                    Call Registro
                    fnTipoRegMant = 1
                Else
                    Call Mantenimiento
                    fnTipoRegMant = 2
                End If
            ElseIf fnTipoRegMant = 3 Then  ' Para el Evento. "Consulta"
                    Call Mantenimiento
                    fnTipoRegMant = 3
            End If
        Else
            Unload Me
            Exit Function
        End If
    Else
        If CargaControlesTipoPermiso(1, False) Then
        End If
    End If
    'Fin Carga <-*****

     'Habilita / Deshabilita Botones - Text
        If fnEstado = 2000 Then             '*****-> Si es Solicitado
            'Me.cmdGuardar4.Enabled = True
            Me.cmdImprimir.Enabled = False
            Me.cmdInformeVisita4.Enabled = False
            cmdFlujoCaja4.Enabled = False 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = False
            Else
                Me.cmdVerCar.Enabled = False
            End If
        Else                                '*****-> Sugerido +
            'Me.cmdGuardar4.Enabled = True
            Me.cmdImprimir.Enabled = True
            cmdFlujoCaja4.Enabled = True 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = True 'No refinanciado
                Me.cmdInformeVisita4.Enabled = True
            Else
                Me.cmdVerCar.Enabled = False
                Me.cmdInformeVisita4.Enabled = False
            End If
        End If
    
    '*****->No Refinanciados (Propuesta Credito)
        If fnColocCondi <> 4 Then
            txtFechaVisita.Enabled = True
            txtEntornoFamiliar4.Enabled = True
            txtGiroUbicacion4.Enabled = True
            txtExperiencia4.Enabled = True
            txtFormalidadNegocio4.Enabled = True
            txtColaterales4.Enabled = True
            txtDestino4.Enabled = True
         Else
            framePropuesta.Enabled = False
            txtFechaVisita.Enabled = False
            txtEntornoFamiliar4.Enabled = False
            txtGiroUbicacion4.Enabled = False
            txtExperiencia4.Enabled = False
            txtFormalidadNegocio4.Enabled = False
            txtColaterales4.Enabled = False
            txtDestino4.Enabled = False
        End If
    '*****->Fin No Refinanciados
        
    Set rsAceptableCritico = Nothing
    fbGrabar = False
    Call CalculoTotal(2)
    If Not pbImprimir Then
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
End Function

Private Function DevolverMes(ByVal pnMes As Integer, ByRef pnAnio As Integer, ByRef pnMesN As Integer) As String 'Cargar Ultimo 3 Meses -> Registrar
    Dim nIndMes As Integer
    nIndMes = CInt(Mid(gdFecSis, 4, 2)) - pnMes
    pnAnio = CInt(Mid(gdFecSis, 7, 4))
        If nIndMes < 1 Then
            nIndMes = nIndMes + 12
            pnAnio = pnAnio - 1
        End If
    pnMesN = nIndMes
    DevolverMes = Choose(nIndMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

Private Function DevolverMesDatos(ByVal pnMes As Integer) As String 'Cargar 3 Ultimos Meses -> Para el Mantenimiento
    DevolverMesDatos = Choose(pnMes, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre")
End Function

'***** LUCV20160529 / feReferidos
Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosReferencia = False
        If feReferidos.rows - 1 < 2 Then
            MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
            cmdAgregarRef.SetFocus
            SSTabIngresos.Tab = 3
            ValidaDatosReferencia = False
            Exit Function
        End If
        For i = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del DNI
            If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
                For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 2)))
                    If (Mid(feReferidos.TextMatrix(i, 2), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 2), j, 1) > "9") Then
                       MsgBox "Uno de los Digitos del primer DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                       feReferidos.SetFocus
                       SSTabIngresos.Tab = 3
                       ValidaDatosReferencia = False
                       Exit Function
                    End If
                Next j
            End If
        Next i
        For i = 1 To feReferidos.rows - 1  'Verifica Longitud del DNI
            If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
                If Len(Trim(feReferidos.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                    MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    feReferidos.SetFocus
                    SSTabIngresos.Tab = 3
                    ValidaDatosReferencia = False
                    Exit Function
                End If
            End If
        Next i
        For i = 1 To feReferidos.rows - 1  'Verfica Tipo de Valores del Telefono
            If Trim(feReferidos.TextMatrix(i, 1)) <> "" Then
                For j = 1 To Len(Trim(feReferidos.TextMatrix(i, 3)))
                    If (Mid(feReferidos.TextMatrix(i, 3), j, 1) < "0" Or Mid(feReferidos.TextMatrix(i, 3), j, 1) > "9") Then
                       MsgBox "Uno de los Digitos del teléfono de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                       feReferidos.SetFocus
                       SSTabIngresos.Tab = 3
                       ValidaDatosReferencia = False
                       Exit Function
                    End If
                Next j
            End If
        Next i

        For i = 1 To feReferidos.rows - 1 'Verifica ambos DNI que no sean iguales
            For j = 1 To feReferidos.rows - 1
                If i <> j Then
                    If feReferidos.TextMatrix(i, 2) = feReferidos.TextMatrix(j, 2) Then
                        MsgBox "No se puede ingresar el mismo DNI mas de una vez en los referidos", vbInformation, "Alerta"
                        ValidaDatosReferencia = False
                        Exit Function
                    End If
                End If
            Next
        Next
    ValidaDatosReferencia = True
End Function

Public Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    ValidaGrillas = False
    For i = 1 To Flex.rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 3)) = "" Then
                ValidaGrillas = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrillas = True
End Function

Public Function ValidaDatos() As Boolean
    Dim nIndice As Integer
    Dim i As Integer
    ValidaDatos = False
    Dim lsMensajeIfi As String 'LUCV20161115
        If fnTipoPermiso = 3 Then
        '********** Para TAB:0 -> Ingresos y Egresos
            If spnTiempoLocalAnio.valor = "" Then
            MsgBox "Ingrese Tiempo en el mismo local: Años", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                spnTiempoLocalAnio.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If spnTiempoLocalMes.valor = "" Then
            MsgBox "Ingrese Tiempo en el mismo local: Meses", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                spnTiempoLocalMes.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If OptCondLocal(1).value = 0 And OptCondLocal(2).value = 0 And OptCondLocal(3).value = 0 And OptCondLocal(4).value = 0 Then
                MsgBox "Falta elegir la Condicion del local", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
            If txtCondLocalOtros.Visible = True Then
                If txtCondLocalOtros.Text = "" Then
                MsgBox "Ingrese la Descripcion de la Opcion: Otro Local", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 0
                    ValidaDatos = False
                    Exit Function
                End If
            End If
            If Trim(txtGiroNeg.Text) = "" Then
                MsgBox "Falta ingresar el Giro del Negocio, Favor Actualizar los Datos del Cliente", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                txtGiroNeg.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If Trim(txtFechaEvaluacion.Text) = "__/__/____" Then
                MsgBox "Falta Ingresar la Fecha de Evaluacion", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                txtFechaEvaluacion.SetFocus
                ValidaDatos = False
                Exit Function
            End If
    
        '********** Para TAB:1 -> Propuesta del Credito
            If fnColocCondi <> 4 Then 'Valida, si el credito no es refinanciado
                If Trim(txtFechaVisita.Text) = "__/__/____" Or Not IsDate(Trim(txtFechaVisita.Text)) Then
                    MsgBox "Falta ingresar la fecha de visita o el formato de la fecha no es el correcto." & Chr(10) & " Formato: DD/MM/YYY", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtFechaVisita.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                If txtEntornoFamiliar4.Text = "" Then
                    MsgBox "Por favor Ingrese, El Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtEntornoFamiliar4.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                If txtGiroUbicacion4.Text = "" Then
                    MsgBox "Por favor Ingrese, El Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtGiroUbicacion4.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                If txtExperiencia4.Text = "" Then
                    MsgBox "Por favor Ingrese, Sobre la Experiencia Crediticia", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtExperiencia4.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                If txtFormalidadNegocio4.Text = "" Then
                    MsgBox "Por favor Ingrese, La Formalidad del Negocio", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtFormalidadNegocio4.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                If txtColaterales4.Text = "" Then
                    MsgBox "Por favor Ingrese, Sobre las Garantias y Colaterales", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtColaterales4.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                If txtDestino4.Text = "" Then
                    MsgBox "Por favor Ingrese, El destino del Credito", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtDestino4.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
            
        '********** PARA TAB2 -> Comentarios y Referidos
            'LUCV25072016->*****, Si el cliente es Nuevo -> Referente es Obligatorio
            'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
            If Not fbTieneReferido6Meses Then
                frameReferido.Enabled = True
                frameComentario.Enabled = True
                    For i = 0 To feReferidos.rows - 1
                        If feReferidos.TextMatrix(i, 0) <> "" Then
                            If Trim(feReferidos.TextMatrix(i, 0)) = "" Or Trim(feReferidos.TextMatrix(i, 1)) = "" _
                                Or Trim(feReferidos.TextMatrix(i, 2)) = "" Or Trim(feReferidos.TextMatrix(i, 3)) = "" Or Trim(feReferidos.TextMatrix(i, 4)) = "" Then
                                MsgBox "Faltan datos en la lista de Referencias", vbInformation, "Aviso"
                                SSTabIngresos.Tab = 3
                                ValidaDatos = False
                                Exit Function
                            End If
                        End If
                    Next i
            
                    If ValidaDatosReferencia = False Then 'Contenido de feReferidos2: Referidos
                        SSTabIngresos.Tab = 3
                        ValidaDatos = False
                        Exit Function
                    End If
                    
                    If txtComentario4.Text = "" Then
                        MsgBox "Por favor Ingrese, Comentarios", vbInformation, "Aviso"
                        SSTabIngresos.Tab = 3
                        txtComentario4.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
            Else
                'si el cliente es nuevo-> referido obligatorio
                    frameReferido.Enabled = False
                    feReferidos.Enabled = False
                    cmdAgregarRef.Enabled = False
                    cmdQuitar4.Enabled = False
                    txtComentario4.Enabled = False 'Comentarios
                    frameComentario.Enabled = False
            End If
            'Fin LUCV25072016 <-*****
            
        '********** Para TAB:0 -> Validacion Grillas: GastosNegocio, OtrosIngresos, GastosFamiliares
            If ValidaGrillas(feOtrosIngresos) = False Then
                MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
            If ValidaGrillas(feGastosFamiliares) = False Then
                MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
        '********** Para TAB:1 -> Grilla Balance General
            For nIndice = 1 To feFlujoCajaMensual.rows - 1
                'Flujo de Caja Mensual
                'If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 20 Then'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 22 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    If val(Replace(feFlujoCajaMensual.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Saldo disponible = (Margen Bruto) - (Otros Egresos) " & Chr(10) & " - El saldo disponible no puede ser un valor menor que cero.", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 1
                        Me.feFlujoCajaMensual.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
                'If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 3 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                If feFlujoCajaMensual.TextMatrix(nIndice, 2) = 4 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    If val(Replace(feFlujoCajaMensual.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Flujo de Caja Mensual: Egresos por compras " & Chr(10) & " - El Valor ingresado tiene que ser un monto mayor a cero", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos.Tab = 1
                        Exit Function
                    End If
                End If
            Next
            For nIndice = 1 To feActivos.rows - 1
                'Activo Corriente
                If feActivos.TextMatrix(nIndice, 6) = 100 Then
                    If val(Replace(feActivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Activo Corriente: " & Chr(10) & " El total, no tiene que ser un valor menor o igual que cero. ", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
                'Total Activo
                If feActivos.TextMatrix(nIndice, 6) = 1000 Then
                    If val(Replace(feActivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Total Activo: " & Chr(10) & " El total, no tiene que ser un valor menor o igual que cero. ", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
                Next
                'Pasivo  Corriente
            For nIndice = 1 To fePasivos.rows - 1
                If fePasivos.TextMatrix(nIndice, 6) = 100 Then
                    If val(Replace(fePasivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Pasivo Corriente: " & Chr(10) & " - El total, no tiene que ser un valor menor o igual que cero ", vbInformation, "Alerta"
                        SSTabIngresos.Tab = 0
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            'Patrimonio
                If fePasivos.TextMatrix(nIndice, 6) = 300 Then
                    If val(Replace(fePasivos.TextMatrix(nIndice, 4), ",", "")) <= 0 Then
                        MsgBox "Patrimonio: " & Chr(10) & " - El total, no tiene que ser un valor menor o igual que cero ", vbInformation, "Alerta"
                         SSTabIngresos.Tab = 0
                         Me.fePasivos.SetFocus
                         ValidaDatos = False
                        Exit Function
                    End If
                End If
            Next
        'LUCV20161115, Agregó->Según ERS068-2016
        If Not ValidaIfiExisteCompraDeuda(sCtaCod, MatIfiGastoFami, MatIfiGastoNego, lsMensajeIfi) Or Len(Trim(lsMensajeIfi)) > 0 Then
            MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
            SSTabIngresos.Tab = 1
            Exit Function
        End If
            
    End If
    ValidaDatos = True
End Function

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, False, pPermiteGrabar)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
    CargaControlesTipoPermiso = False
    End If

    If pBloqueaTodo Then 'Para el Caso despues de dar Verificacion
        Call HabilitaControles(True, True, False)
        CargaControlesTipoPermiso = True
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaRatios As Boolean, ByVal pbHabilitaGuardar As Boolean)
    'HabilitacionControlesAnalistas:     pbHabilitaA = True
    'Tab0: Ingresos/Egresos
    spnTiempoLocalAnio.Enabled = pbHabilitaA
    spnTiempoLocalMes.Enabled = pbHabilitaA
    OptCondLocal(1).Enabled = pbHabilitaA
    OptCondLocal(2).Enabled = pbHabilitaA
    OptCondLocal(3).Enabled = pbHabilitaA
    OptCondLocal(4).Enabled = pbHabilitaA
    'txtFechaEvaluacion.Enabled = pbHabilitaA
    txtCondLocalOtros.Enabled = pbHabilitaA
    feActivos.Enabled = pbHabilitaA
    fePasivos.Enabled = pbHabilitaA
    
    'Tab1:  Flujo Caja Mensual
    feFlujoCajaMensual.Enabled = pbHabilitaA
    feGastosFamiliares.Enabled = pbHabilitaA
    feOtrosIngresos.Enabled = pbHabilitaA
    feDeclaracionPDT.Enabled = pbHabilitaA

    'Tab2: Propuesta/Credito
    txtFechaVisita.Enabled = pbHabilitaA
    txtEntornoFamiliar4.Enabled = pbHabilitaA
    txtGiroUbicacion4.Enabled = pbHabilitaA
    txtExperiencia4.Enabled = pbHabilitaA
    txtFormalidadNegocio4.Enabled = pbHabilitaA
    txtColaterales4.Enabled = pbHabilitaA
    txtDestino4.Enabled = pbHabilitaA

    'Tab3: Comentarios/Referidos
    txtComentario4.Enabled = pbHabilitaA
    feReferidos.Enabled = pbHabilitaA
    cmdAgregarRef.Enabled = pbHabilitaA
    cmdQuitar4.Enabled = pbHabilitaA
    frameReferido.Enabled = pbHabilitaA

   'txtVerif.Enabled = pbHabilitaB
    If fnEstado = 2000 Then
        SSTabRatios.Visible = False
    Else
        SSTabRatios.Visible = pbHabilitaRatios
    End If

    'cmdInformeVisita4.Enabled = pbHabilitaRatios
    'cmdVerCar.Enabled = pbHabilitaRatios
    'cmdImprimir.Enabled = pbHabilitaRatios
    cmdGuardar4.Enabled = pbHabilitaGuardar
End Function
Private Sub CargaControlesInicio()
    Call CargarFlexEdit
    
    'DesHabilita la CargaInicial de Controles
    ActXCodCta.Enabled = False
    txtNombreCliente.Enabled = False
    txtExposicionCredito.Enabled = False
    txtGiroNeg.Enabled = False
    txtUltEndeuda.Enabled = False
    txtFecUltEndeuda.Enabled = False
    spnExpEmpAnio.Enabled = False
    spnExpEmpMes.Enabled = False
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidadPat.Enabled = False
    txtLiquidezCte.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
End Sub
Private Sub CargarFlexEdit() 'Registrar New Formato Evaluacion
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMontoIni As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Dim nFila, NumRegRS As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    nMontoIni = Format(0, "00.00")
    
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(fnFormato, sCtaCod, _
                                                        rsFeGastoNeg, _
                                                        rsFeDatGastoFam, _
                                                        rsFeDatOtrosIng, _
                                                        rsFeDatBalanGen, _
                                                        rsFeDatActivos, _
                                                        rsFeDatPasivos, _
                                                        rsFeDatPasivosNo, _
                                                        rsFeDatPatrimonio, _
                                                        rsFeDatRef, _
                                                        rsFeFlujoCaja, _
                                                        , , , , rsFeDatPDT)
                                                                                                      
    'Flex Activos ->CargaInicial
    feActivos.Clear
    feActivos.FormaCabecera
    feActivos.rows = 2
    Call LimpiaFlex(feActivos)
        nFila = 0
        NumRegRS = 0
        NumRegRS = rsFeDatActivos.RecordCount
        ReDim lvPrincipalActivos(NumRegRS)
        Do While Not rsFeDatActivos.EOF
            feActivos.AdicionaFila
            lnFila = feActivos.row
            feActivos.TextMatrix(lnFila, 1) = rsFeDatActivos!cConsDescripcion
            feActivos.TextMatrix(lnFila, 2) = Format(rsFeDatActivos!nPP, "#,#0.00")
            feActivos.TextMatrix(lnFila, 3) = Format(rsFeDatActivos!nPE, "#,#0.00")
            feActivos.TextMatrix(lnFila, 4) = Format(rsFeDatActivos!nTotal, "#,#0.00")
            feActivos.TextMatrix(lnFila, 5) = rsFeDatActivos!nConsCod
            feActivos.TextMatrix(lnFila, 6) = rsFeDatActivos!nConsValor
            
            'Lena datos de Registro en Matrix "lvPrincipalActivosPasivos"
            lvPrincipalActivos(lnFila).cConcepto = rsFeDatActivos!cConsDescripcion
            lvPrincipalActivos(lnFila).nImportePP = rsFeDatActivos!nPP
            lvPrincipalActivos(lnFila).nImportePE = rsFeDatActivos!nPP
            lvPrincipalActivos(lnFila).nConsCod = rsFeDatActivos!nConsCod
            lvPrincipalActivos(lnFila).nConsValor = rsFeDatActivos!nConsValor
            
        Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
            Case 1000, 100, 200, 300, 400, 500
                Me.feActivos.BackColorRow (&H80000000)
                Me.feActivos.ForeColorRow vbBlack, True
            Case Else
                Me.feActivos.BackColorRow (&HFFFFFF)
                Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
        rsFeDatActivos.MoveNext
        Loop
    rsFeDatActivos.Close
    Set rsFeDatActivos = Nothing
                                                                                                                                                                                              
    'Flex Pasivos->CargaInicial
    fePasivos.Clear
    fePasivos.FormaCabecera
    fePasivos.rows = 2
    Call LimpiaFlex(fePasivos)
    nFila = 0
    NumRegRS = 0
    NumRegRS = rsFeDatPasivos.RecordCount
    ReDim lvPrincipalPasivos(NumRegRS)
        Do While Not rsFeDatPasivos.EOF
            fePasivos.AdicionaFila
            lnFila = fePasivos.row
            fePasivos.TextMatrix(lnFila, 1) = rsFeDatPasivos!cConsDescripcion
            fePasivos.TextMatrix(lnFila, 2) = Format(rsFeDatPasivos!nPP, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 3) = Format(rsFeDatPasivos!nPE, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 4) = Format(rsFeDatPasivos!nTotal, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 5) = rsFeDatPasivos!nConsCod
            fePasivos.TextMatrix(lnFila, 6) = rsFeDatPasivos!nConsValor
            
            lvPrincipalPasivos(lnFila).cConcepto = rsFeDatPasivos!cConsDescripcion
            lvPrincipalPasivos(lnFila).nImportePP = rsFeDatPasivos!nPP
            lvPrincipalPasivos(lnFila).nImportePE = rsFeDatPasivos!nPP
            lvPrincipalPasivos(lnFila).nConsCod = rsFeDatPasivos!nConsCod
            lvPrincipalPasivos(lnFila).nConsValor = rsFeDatPasivos!nConsValor
            
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
            Case 1000, 1001, 1002, 100, 200, 300, 400, 500
                Me.fePasivos.BackColorRow (&H80000000)
                Me.fePasivos.ForeColorRow vbBlack, True
                Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
            Case 109, 201
                Me.fePasivos.BackColorRow &HC0FFFF
                Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
            Case 206
                Me.fePasivos.BackColorRow vbWhite, True
                Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
            Case 301
                Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
            Case Else
                Me.fePasivos.BackColorRow (&HFFFFFF)
                Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
            rsFeDatPasivos.MoveNext
        Loop
    rsFeDatPasivos.Close
    Set rsFeDatPasivos = Nothing
                                                                                                                                               
    'Flex Flujo Caja Mensual
    feFlujoCajaMensual.Clear
    feFlujoCajaMensual.FormaCabecera
    feFlujoCajaMensual.rows = 2
    Call LimpiaFlex(feFlujoCajaMensual)
        Do While Not rsFeFlujoCaja.EOF
            feFlujoCajaMensual.AdicionaFila
            lnFila = feFlujoCajaMensual.row
            feFlujoCajaMensual.TextMatrix(lnFila, 1) = rsFeFlujoCaja!nConsCod
            feFlujoCajaMensual.TextMatrix(lnFila, 2) = rsFeFlujoCaja!nConsValor
            feFlujoCajaMensual.TextMatrix(lnFila, 3) = rsFeFlujoCaja!cConsDescripcion
            feFlujoCajaMensual.TextMatrix(lnFila, 4) = Format(rsFeFlujoCaja!nMonto, "#,#0.00")
            
            Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
                'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Me.feFlujoCajaMensual.BackColorRow (&H80000000)
                    Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
                    Me.feFlujoCajaMensual.BackColorRow vbWhite, True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Case 19 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
                    Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
                Case Else
                    Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
                    Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
            End Select
            rsFeFlujoCaja.MoveNext
        Loop
    rsFeFlujoCaja.Close
    Set rsFeFlujoCaja = Nothing
                                                                              
    'Flex otros Ingresos
    feOtrosIngresos.Clear
    feOtrosIngresos.FormaCabecera
    feOtrosIngresos.rows = 2
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsFeDatOtrosIng!nMonto, "#,##0.00")
            rsFeDatOtrosIng.MoveNext
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing

    'Gastos Familiares
    feGastosFamiliares.Clear
    feGastosFamiliares.FormaCabecera
    feGastosFamiliares.rows = 2
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsFeDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,##0.00")
                
            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
                Case gCodCuotaIfiGastoFami
                    Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Case gCodDeudaLCNUGastoFami
                    Me.feGastosFamiliares.BackColorRow vbWhite, True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Case Else
                    Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsFeDatGastoFam.MoveNext
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
    
    'Declaracion PDT
    sMes1 = DevolverMes(1, nAnio3, nMes3)
    sMes2 = DevolverMes(2, nAnio2, nMes2)
    sMes3 = DevolverMes(3, nAnio1, nMes1)
    
    feDeclaracionPDT.Clear
    feDeclaracionPDT.FormaCabecera
    feDeclaracionPDT.rows = 2
        
    feDeclaracionPDT.TextMatrix(0, 4) = sMes3
    feDeclaracionPDT.TextMatrix(0, 5) = sMes2
    feDeclaracionPDT.TextMatrix(0, 6) = sMes1
    
    feDeclaracionPDT.TextMatrix(0, 1) = "Mes/Detalle" '& Space(8)
    For i = 1 To 2
        feDeclaracionPDT.AdicionaFila
        'feDeclaracionPDT.TextMatrix(i, 1) = Choose(i, "Compras" & Space(8), "Ventas" & Space(8))
        feDeclaracionPDT.TextMatrix(i, 1) = rsFeDatPDT!cConsDescripcion
        feDeclaracionPDT.TextMatrix(i, 2) = rsFeDatPDT!nConsCod
        feDeclaracionPDT.TextMatrix(i, 3) = rsFeDatPDT!nConsValor
        feDeclaracionPDT.TextMatrix(i, 4) = Choose(i, "0.00", "0.00") 'Mes3
        feDeclaracionPDT.TextMatrix(i, 5) = Choose(i, "0.00", "0.00") 'Mes2
        feDeclaracionPDT.TextMatrix(i, 6) = Choose(i, "0.00", "0.00") 'Mes1
        feDeclaracionPDT.TextMatrix(i, 7) = Choose(i, "0.00", "0.00") 'Promedio
        feDeclaracionPDT.TextMatrix(i, 8) = Choose(i, "0.00", "0.00") '%Ventas
        rsFeDatPDT.MoveNext
    Next i
End Sub
Private Function CargaDatos() As Boolean 'Mantenimiento Formatos
On Error GoTo ErrorCargaDatos
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
 
    CargaDatos = oNCOMFormatosEval.CargaDatosCredEvaluacion2(sCtaCod, _
                                                            fnFormato, _
                                                            rsCredEval, _
                                                            rsDatGastoNeg, _
                                                            rsDatGastoFam, _
                                                            rsDatOtrosIng, _
                                                            rsDatRef, _
                                                            rsDatActivos, _
                                                            rsDatPasivos, _
                                                            rsCuotaIFIs, _
                                                            rsPropuesta, _
                                                            rsCapacPagoNeta, _
                                                            rsDatRatioInd, _
                                                            rsDatActivoPasivo, _
                                                            rsDatIfiGastoNego, _
                                                            rsDatIfiGastoFami, , _
                                                            rsDatFlujoCaja, _
                                                            rsDatPDT, _
                                                            rsDatPDTDet, _
                                                            gFormatoActivos, _
                                                            gFormatoPasivos, _
                                                            rsDatParamFlujoCajaForm4)
                                                            
                                                            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja rsDatParamFlujoCajaForm4
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbInformation, "Error"
End Function

Private Sub CalculoTotal(ByVal pnTipo As Integer)
    Dim nTotalActPP As Double 'Total Activos (PP:Patrimonio Empresarial | PE: Patrionio Personal)
    Dim nTotalActPE As Double
    Dim nTotalPasPP As Double 'Total Pasivos
    Dim nTotalPasPE As Double 'Total Pasivos
    
    Dim nActivoCtePP As Double 'Activo Cuenta Corriente PP
    Dim nActivoCtePE As Double 'Activo Cuenta Corriente PE
    Dim nInventarioPP As Double 'Inventario PP
    Dim nInventarioPE As Double 'Inventario PE
    Dim nActiFijoPP As Double 'Activo Fijo PP
    Dim nActiFijoPE As Double 'Activo Fijo PE
    
    Dim nPasiCtePP As Double 'Pasivo CtaCte
    Dim nPasiCtePE As Double 'Pasivo CtaCte
    Dim nPasiNoCtePP As Double 'Pasivo NO CtaCte
    Dim nPasiNoCtePE As Double 'Pasivo NO CtaCte
    Dim nPatriPP As Double 'Patrimonio
    Dim nPatriPE As Double 'Patrimonio
    Dim nExistePat As Double 'Saber si el monto del detalle es > 0
    Dim nPatriTotal As Double
    
    Dim nMargenCaja As Double 'FlujoCaja
    Dim nOtrosEgresos As Double 'FlujoCaja
    Dim nMontoDeclarado As Double
    Dim nTotalPP As Currency, nTotalPE As Currency
    
    Dim nTotalActiCte As Currency
    Dim nTotalActivo As Currency
    Dim nTotalPasivo As Currency
    Dim nCuotasIfisFlujo As Double
    
    Dim nCapitalPP As Currency 'Para Calculo Fila -> Capital(Patrimonio)
    Dim nCapitalPE As Currency
    Dim nCapitalTotal As Currency
    Dim nCapitalAdicPP As Currency 'Fila -> Capital Adicional
    Dim nCapitalAdicPE As Currency
    Dim nCapitalAdicTotal As Currency
    Dim nExcedenteRevalPP As Currency 'Fila ->Excedente de Revaluación
    Dim nExcedenteRevalPE As Currency
    Dim nExcedenteRevalTotal As Currency
    Dim nReservaLegalPP As Currency 'Fila ->Reserva Legal
    Dim nReservaLegalPE As Currency
    Dim nReservaLegalTotal As Currency
    Dim nResultadoEjercicioPP As Currency 'Fila -> Resultado del Ejercicio
    Dim nResultadoEjercicioPE As Currency
    Dim nResultadoEjercicioTotal As Currency
    Dim nResultadoAcumuladoPP As Currency 'Fila ->Resultado Acumulado
    Dim nResultadoAcumuladoPE As Currency
    Dim nResultadoAcumuladoTotal As Currency
    Dim nTotalActivoPP As Currency
    Dim nTotalActivoPE As Currency
         
    Dim nTotalPasivoPatrimonio As Currency
    Dim nTotalPatrimonio As Currency
           
    Dim nPorcentajeVentas As Double
    Dim nPorcentajeCompras As Double
    
    nTotalActPP = 0: nTotalActPE = 0: nTotalPasPP = 0: nTotalPasPE = 0
    nActivoCtePP = 0: nActivoCtePE = 0: nInventarioPP = 0: nInventarioPE = 0: nActiFijoPP = 0: nActiFijoPE = 0
    nPasiCtePP = 0: nPasiCtePE = 0: nPasiNoCtePP = 0: nPasiNoCtePE = 0: nPatriPP = 0: nPatriPE = 0
    nMargenCaja = 0: nOtrosEgresos = 0
    nMontoDeclarado = 0
    nTotalActiCte = 0: nTotalActivo = 0: nTotalPasivo = 0
    nCuotasIfisFlujo = 0: nPorcentajeVentas = 0: nPorcentajeCompras = 0
    
    nCapitalPP = 0: nCapitalPE = 0: nCapitalTotal = 0:
    nCapitalAdicPP = 0: nCapitalAdicPE = 0: nCapitalAdicTotal = 0
    nExcedenteRevalPP = 0: nExcedenteRevalPE = 0: nExcedenteRevalTotal = 0
    nReservaLegalPP = 0: nReservaLegalPE = 0: nReservaLegalTotal = 0
    nResultadoEjercicioPP = 0: nResultadoEjercicioPE = 0: nResultadoEjercicioTotal = 0:
    nResultadoAcumuladoPP = 0: nResultadoAcumuladoPE = 0: nResultadoAcumuladoTotal = 0:
    nTotalPasivoPatrimonio = 0: nTotalPatrimonio = 0
    
    
On Error GoTo ErrorCalculo
Select Case pnTipo
    Case 1:
            'ACTIVOS:**********->
            'Sumatoria: Activo Corriente
            For i = 2 To 6
                nActivoCtePP = nActivoCtePP + CDbl(Me.feActivos.TextMatrix(i, 2))
                nActivoCtePE = nActivoCtePE + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i
            Me.feActivos.TextMatrix(1, 2) = Format(nActivoCtePP, "#,#0.00") 'Resultado: Activo Cte PP
            Me.feActivos.TextMatrix(1, 3) = Format(nActivoCtePE, "#,#0.00") 'Resultado: Activo Cte PE

            'Sumatoria: Inventario
            For i = 8 To 11
                nInventarioPP = nInventarioPP + CDbl(Me.feActivos.TextMatrix(i, 2))
                nInventarioPE = nInventarioPE + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i
            Me.feActivos.TextMatrix(7, 2) = Format(nInventarioPP, "#,#0.00")
            Me.feActivos.TextMatrix(7, 3) = Format(nInventarioPE, "#,#0.00")
            
            'Sumatoria: Activo Fijo
            For i = 13 To 17
                nActiFijoPP = nActiFijoPP + CDbl(Me.feActivos.TextMatrix(i, 2))
                nActiFijoPE = nActiFijoPE + CDbl(Me.feActivos.TextMatrix(i, 3))
            Next i
            Me.feActivos.TextMatrix(12, 2) = Format(nActiFijoPP, "#,#0.00")
            Me.feActivos.TextMatrix(12, 3) = Format(nActiFijoPE, "#,#0.00")
    
            'Activo Total (PP | PE)
            Me.feActivos.TextMatrix(18, 2) = Format(nActivoCtePP + nActiFijoPP + nInventarioPP, "#,#0.00")
            Me.feActivos.TextMatrix(18, 3) = Format(nActivoCtePE + nActiFijoPE + nInventarioPE, "#,#0.00")
                
            'Columna Total
            For i = 2 To Me.feActivos.rows - 2 '(Sin Considerar al Activo Cte y al Total)
            nTotalActPP = CDbl(Me.feActivos.TextMatrix(i, 2))
            nTotalActPE = CDbl(Me.feActivos.TextMatrix(i, 3))
            Me.feActivos.TextMatrix(i, 4) = Format(nTotalActPP + nTotalActPE, "#,#0.00")
            Next i
            
            'Calculo del "TotalActivoCte":
            nTotalActiCte = Format(Me.feActivos.TextMatrix(7, 4) + nActivoCtePP + nActivoCtePE, "#,#0.00")
            Me.feActivos.TextMatrix(1, 4) = Format(CCur(nTotalActiCte), "#,#0.00")
            'Calculo del "TOTAL":
            nTotalActivo = Format(CCur(Me.feActivos.TextMatrix(1, 4)) + CCur(Me.feActivos.TextMatrix(12, 4)), "#,#0.00")
            Me.feActivos.TextMatrix(18, 4) = Format(nTotalActivo, "#,#0.00")
            
             Call CalculoTotal(2)
    Case 2:
            'PASIVOS:**********->
            'Sumatoria (PP/ PE): Pasivo Corriente
            For i = 2 To 7
                nPasiCtePP = nPasiCtePP + CDbl(Me.fePasivos.TextMatrix(i, 2))
                nPasiCtePE = nPasiCtePE + CDbl(Me.fePasivos.TextMatrix(i, 3))
            Next i
                Me.fePasivos.TextMatrix(1, 2) = Format(nPasiCtePP, "#,#0.00") 'Resultado: PAsivo Cte PP
                Me.fePasivos.TextMatrix(1, 3) = Format(nPasiCtePE, "#,#0.00") 'Resultado: Pasivo Cte PE
                nTotalPP = nPasiCtePP
                nTotalPE = nPasiCtePE
            
            'Sumatoria (PP/ PE): Pasivo No Corriente
            For i = 9 To 12
                nPasiNoCtePP = nPasiNoCtePP + CDbl(Me.fePasivos.TextMatrix(i, 2))
                nPasiNoCtePE = nPasiNoCtePE + CDbl(Me.fePasivos.TextMatrix(i, 3))
            Next i
                Me.fePasivos.TextMatrix(8, 2) = Format(nPasiNoCtePP, "#,#0.00")
                Me.fePasivos.TextMatrix(8, 3) = Format(nPasiNoCtePE, "#,#0.00")
                nTotalPP = nTotalPP + nPasiNoCtePP
                nTotalPE = nTotalPE + nPasiNoCtePE
            
            'Sumatoria Capital (Patrimonio)
            nTotalActivoPP = Me.feActivos.TextMatrix(18, 2)
            nTotalActivoPE = Me.feActivos.TextMatrix(18, 3)
            
            nCapitalAdicPP = Me.fePasivos.TextMatrix(15, 2)
            nCapitalAdicPE = Me.fePasivos.TextMatrix(15, 3)
            nCapitalAdicTotal = Me.fePasivos.TextMatrix(15, 4)
            nExcedenteRevalPP = Me.fePasivos.TextMatrix(16, 2)
            nExcedenteRevalPE = Me.fePasivos.TextMatrix(16, 3)
            nExcedenteRevalTotal = Me.fePasivos.TextMatrix(16, 4)
            nReservaLegalPP = Me.fePasivos.TextMatrix(17, 2)
            nReservaLegalPE = Me.fePasivos.TextMatrix(17, 3)
            nReservaLegalTotal = Me.fePasivos.TextMatrix(17, 4)
            nResultadoEjercicioPP = Me.fePasivos.TextMatrix(18, 2)
            nResultadoEjercicioPE = Me.fePasivos.TextMatrix(18, 3)
            nResultadoEjercicioTotal = Me.fePasivos.TextMatrix(18, 4)
            nResultadoAcumuladoPP = Me.fePasivos.TextMatrix(19, 2)
            nResultadoAcumuladoPE = Me.fePasivos.TextMatrix(19, 3)
            nResultadoAcumuladoTotal = Me.fePasivos.TextMatrix(19, 4)
            nTotalPasivo = Me.fePasivos.TextMatrix(20, 4)
                       
            nCapitalPP = nTotalActivoPP - nTotalPP - (nCapitalAdicPP + nExcedenteRevalPP + nReservaLegalPP + nResultadoEjercicioPP + nResultadoAcumuladoPP) 'Capital - Patrimonio PP
            nCapitalPE = nTotalActivoPE - nTotalPE - (nCapitalAdicPE + nExcedenteRevalPE + nReservaLegalPE + nResultadoEjercicioPE + nResultadoAcumuladoPE) 'Capital - Patrimonio PE
            nCapitalTotal = nTotalActivo - nTotalPasivo - (nCapitalAdicTotal + nExcedenteRevalTotal + nReservaLegalTotal + nResultadoEjercicioTotal + nResultadoAcumuladoTotal) 'Total Capital -Patrimonio
           
            Me.fePasivos.TextMatrix(14, 2) = Format(nCapitalPP, "#,#0.00")
            Me.fePasivos.TextMatrix(14, 3) = Format(nCapitalPE, "#,#0.00")
            Me.fePasivos.TextMatrix(14, 4) = Format(nCapitalTotal, "#,#0.00")
        
           'Verificar si Existe detalle Patrimonio
           For i = 14 To 19
                nExistePat = nExistePat + CDbl(Me.fePasivos.TextMatrix(i, 4))
           Next i
            
            'Sumatoria (PP/ PE): Patrimonio
            If nExistePat <> 0 Then
                For i = 14 To 19
                nPatriPP = nPatriPP + CDbl(Me.fePasivos.TextMatrix(i, 2))
                nPatriPE = nPatriPE + CDbl(Me.fePasivos.TextMatrix(i, 3))
                Next i
                Me.fePasivos.TextMatrix(13, 2) = Format(nPatriPP, "#,#0.00")
                Me.fePasivos.TextMatrix(13, 3) = Format(nPatriPE, "#,#0.00")
            Else
                nPatriTotal = Me.feActivos.TextMatrix(18, 4) - Me.fePasivos.TextMatrix(20, 4)
                Me.fePasivos.TextMatrix(13, 4) = Format(nPatriTotal, "#,#0.00")
            End If
    
            'Total Pasivo y Patrimonio (PP | PE)
            Me.fePasivos.TextMatrix(20, 2) = Format(nPasiCtePP + nPasiNoCtePP, "#,#0.00")
            Me.fePasivos.TextMatrix(20, 3) = Format(nPasiCtePE + nPasiNoCtePE, "#,#0.00")
            
            'Me.fePasivos.TextMatrix(21, 2) = Format(nPasiCtePP + nPasiNoCtePP + nPatriPP, "#,#0.00")
            'Me.fePasivos.TextMatrix(21, 3) = Format(nPasiCtePE + nPasiNoCtePE + nPatriPE, "#,#0.00")
                        
            'Columna Total= PP + PE
            For i = 1 To Me.fePasivos.rows - 1
                nTotalPasPP = CDbl(Me.fePasivos.TextMatrix(i, 2))
                nTotalPasPE = CDbl(Me.fePasivos.TextMatrix(i, 3))
                Me.fePasivos.TextMatrix(i, 4) = Format(nTotalPasPP + nTotalPasPE, "#,#0.00")
            Next i
            
            If nPatriTotal <> 0 Then ' Para el Caso que no Exista detalle de Patrimonio
               Me.fePasivos.TextMatrix(13, 4) = Format(nPatriTotal, "#,#0.00")
               Me.fePasivos.TextMatrix(21, 4) = Format(CDbl(Me.fePasivos.TextMatrix(1, 4)) + CDbl(Me.fePasivos.TextMatrix(8, 4)) + CDbl(Me.fePasivos.TextMatrix(13, 4)), "#,#0.00")
            End If
            
            'Calculo de Total Pasivo y Patrimonio
             nTotalPatrimonio = Me.fePasivos.TextMatrix(13, 4)
             nTotalPasivoPatrimonio = nTotalPasivo + nTotalPatrimonio
             Me.fePasivos.TextMatrix(21, 4) = Format(nTotalPasivoPatrimonio, "#,#0.00")
            
             Me.fePasivos.TextMatrix(21, 2) = Format(nPasiCtePP + nPasiNoCtePP + nPatriPP, "#,#0.00") 'Total Pasivo y Patrimonio
             Me.fePasivos.TextMatrix(21, 3) = Format(nPasiCtePE + nPasiNoCtePE + nPatriPE, "#,#0.00") 'Total Pasivo y Patrimonio
             
            
        Case 3:
            'Margen Bruto Caja
            'For i = 1 To 2 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            For i = 1 To 3 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                nMargenCaja = nMargenCaja + CDbl(Me.feFlujoCajaMensual.TextMatrix(i, 4))
            Next i
                nMargenCaja = nMargenCaja - CDbl(Me.feFlujoCajaMensual.TextMatrix(i, 4))
                'Me.feFlujoCajaMensual.TextMatrix(4, 4) = Format(nMargenCaja, "#,#0.00") 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                Me.feFlujoCajaMensual.TextMatrix(5, 4) = Format(nMargenCaja, "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
           
           'Otros Egresos
            'For i = 6 To 19 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            For i = 7 To 21 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                nOtrosEgresos = nOtrosEgresos + CDbl(Me.feFlujoCajaMensual.TextMatrix(i, 4))
            Next i
                'Me.feFlujoCajaMensual.TextMatrix(5, 4) = Format(nOtrosEgresos, "#,#0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                 Me.feFlujoCajaMensual.TextMatrix(6, 4) = Format(nOtrosEgresos, "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            
           'Saldo Disponible
             'nCuotasIfisFlujo = CDbl(feFlujoCajaMensual.TextMatrix(17, 4)) + CDbl(feFlujoCajaMensual.TextMatrix(18, 4))
             nCuotasIfisFlujo = CDbl(feFlujoCajaMensual.TextMatrix(18, 4)) + CDbl(feFlujoCajaMensual.TextMatrix(19, 4)) 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
             'Me.feFlujoCajaMensual.TextMatrix(20, 4) = Format(nMargenCaja - (CDbl(nOtrosEgresos) - CDbl(nCuotasIfisFlujo)), "#,#0.00")
             Me.feFlujoCajaMensual.TextMatrix(22, 4) = Format(nMargenCaja - (CDbl(nOtrosEgresos) - CDbl(nCuotasIfisFlujo)), "#,#0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            
        Case 4:
                'If CCur(feFlujoCajaMensual.TextMatrix(1, 4)) = 0 Or CCur(feFlujoCajaMensual.TextMatrix(3, 4)) = 0 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                If CCur(feFlujoCajaMensual.TextMatrix(1, 4)) = 0 Or CCur(feFlujoCajaMensual.TextMatrix(4, 4)) = 0 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                
                    MsgBox "Consideraciones en el flujo de caja mensual " & Chr(10) & " - El monto de ventas al contado, no tiene que ser cero." & Chr(10) & " - El monto de egresos por compras, no tiene que ser cero.", vbInformation, "Alerta"
                Exit Sub
                End If
                'Promedio Declaracion PDT
                For i = 1 To feDeclaracionPDT.rows - 1
                    nMontoDeclarado = CDbl(Me.feDeclaracionPDT.TextMatrix(i, 4)) + CDbl(Me.feDeclaracionPDT.TextMatrix(i, 5)) + CDbl(Me.feDeclaracionPDT.TextMatrix(i, 6))
                    nMontoDeclarado = nMontoDeclarado / 3
                    Me.feDeclaracionPDT.TextMatrix(i, 7) = Format(nMontoDeclarado, "#,#0.00")
                Next
            'Para el %Declarado
            nPorcentajeVentas = Round(CCur(feDeclaracionPDT.TextMatrix(1, 7)) / CCur(feFlujoCajaMensual.TextMatrix(1, 4)), 4)
            'nPorcentajeCompras = Round(CCur(feDeclaracionPDT.TextMatrix(2, 7)) / CCur(feFlujoCajaMensual.TextMatrix(3, 4)), 4) 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            nPorcentajeCompras = Round(CCur(feDeclaracionPDT.TextMatrix(2, 7)) / CCur(feFlujoCajaMensual.TextMatrix(4, 4)), 4) 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        
            Me.feDeclaracionPDT.TextMatrix(1, 8) = CStr(nPorcentajeVentas * 100) & "%"
            Me.feDeclaracionPDT.TextMatrix(2, 8) = CStr(nPorcentajeCompras * 100) & "%"
End Select

Exit Sub
ErrorCalculo:
    MsgBox "Aviso: Ingrese los datos Correctamente." & Chr(13) & "Detalles de error: " & Err.Description, vbInformation, "Aviso"
End Sub

Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtNombreCliente.Text = fsCliente
    txtGiroNeg.Text = fsGiroNego
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidadPat.Enabled = False
    txtLiquidezCte.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    
    'si el cliente es nuevo-> referido obligatorio
    'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
        frameReferido.Enabled = True
        feReferidos.Enabled = True
        cmdAgregarRef.Enabled = True
        cmdQuitar4.Enabled = True
        frameComentario.Enabled = True
        txtComentario4.Enabled = True
    Else
        frameReferido.Enabled = False
        feReferidos.Enabled = False
        cmdAgregarRef.Enabled = False
        cmdQuitar4.Enabled = False
        frameComentario.Enabled = False
        txtComentario4.Enabled = False
    End If
    
    'Ratios: Aceptable / Critico ->*****
    If Not (rsAceptableCritico.BOF Or rsAceptableCritico.EOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Crítico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
        
        If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
            Me.lblEndeAceptable.Caption = "Aceptable"
            Me.lblEndeAceptable.ForeColor = &H8000&
        Else
            Me.lblEndeAceptable.Caption = "Crítico"
            Me.lblEndeAceptable.ForeColor = vbRed
        End If
    Else
        lblCapaAceptable.Visible = False
        lblCapaAceptable.Visible = False
    End If
    'Fin Ratios <-****
    
    '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
        txtFechaVisita.Enabled = True
        txtEntornoFamiliar4.Enabled = True
        txtGiroUbicacion4.Enabled = True
        txtExperiencia4.Enabled = True
        txtFormalidadNegocio4.Enabled = True
        txtColaterales4.Enabled = True
        txtDestino4.Enabled = True
     Else
        framePropuesta.Enabled = False
        txtFechaVisita.Enabled = False
        txtEntornoFamiliar4.Enabled = False
        txtGiroUbicacion4.Enabled = False
        txtExperiencia4.Enabled = False
        txtFormalidadNegocio4.Enabled = False
        txtColaterales4.Enabled = False
        txtDestino4.Enabled = False
    End If
    '*****->Fin No Refinanciados
    
     'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        If Not (rsDatParamFlujoCajaForm4.BOF And rsDatParamFlujoCajaForm4.EOF) Then
            EditMoneyIncVC4.Text = rsDatParamFlujoCajaForm4!nIncVentCont
            EditMoneyIncCM4.Text = rsDatParamFlujoCajaForm4!nIncCompMerc
            EditMoneyIncPP4.Text = rsDatParamFlujoCajaForm4!nIncPagPers
            EditMoneyIncGV4.Text = rsDatParamFlujoCajaForm4!nIncGastvent
            EditMoneyIncC4.Text = rsDatParamFlujoCajaForm4!nIncConsu
        End If
        Set rsDatParamFlujoCajaForm4 = Nothing
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
End Function

Private Function Mantenimiento()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim lnFila As Integer
        If fnTipoPermiso = 3 Then
            gsOpeCod = gCredMantenimientoEvaluacionCred
        Else
            'gsOpeCod = gCredVerificacionEvaluacionCred
        End If
        
            'Para Botones *****
            If Not fbBloqueaTodo Then
                cmdInformeVisita4.Enabled = False
                cmdVerCar.Enabled = False
                cmdImprimir.Enabled = False
            End If
            
            'Ver Ratios *****
            If fnEstado > 2000 Then
                SSTabRatios.Visible = True
            Else
                SSTabRatios.Visible = False
                cmdInformeVisita4.Enabled = False
                cmdVerCar.Enabled = False
                cmdImprimir.Enabled = False
            End If
        
        
        'Ratios/ Indicadores
          txtCapacidadNeta.Enabled = False
          txtEndeudamiento.Enabled = False
          txtRentabilidadPat.Enabled = False
          txtLiquidezCte.Enabled = False
          txtIngresoNeto.Enabled = False
          txtExcedenteMensual.Enabled = False
        
        'Si el cliente es nuevo-> referido obligatorio
        'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
        If Not fbTieneReferido6Meses Then
            frameReferido.Enabled = True
            feReferidos.Enabled = True
            cmdAgregarRef.Enabled = True
            cmdQuitar4.Enabled = True
            frameComentario.Enabled = True
            txtComentario4.Enabled = True
        Else
            frameReferido.Enabled = False
            feReferidos.Enabled = False
            cmdAgregarRef.Enabled = False
            cmdQuitar4.Enabled = False
            frameComentario.Enabled = False
            txtComentario4.Enabled = False
        End If
        
        'Ratios: Aceptable / Critico ->*****
         If Not (rsAceptableCritico.EOF Or rsAceptableCritico.BOF) Then
            If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                Me.lblCapaAceptable.Caption = "Aceptable"
                Me.lblCapaAceptable.ForeColor = &H8000&
            Else
                Me.lblCapaAceptable.Caption = "Crítico"
                Me.lblCapaAceptable.ForeColor = vbRed
            End If
            
            If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
                Me.lblEndeAceptable.Caption = "Aceptable"
                Me.lblEndeAceptable.ForeColor = &H8000&
            Else
                Me.lblEndeAceptable.Caption = "Crítico"
                Me.lblEndeAceptable.ForeColor = vbRed
            End If
        Else
            Me.lblCapaAceptable.Visible = False
            Me.lblEndeAceptable.Visible = False
        End If
            'Fin Ratios <-****
            
        '*****->No Refinanciados (Propuesta Credito)
        If fnColocCondi <> 4 Then
            txtFechaVisita.Enabled = True
            txtEntornoFamiliar4.Enabled = True
            txtGiroUbicacion4.Enabled = True
            txtExperiencia4.Enabled = True
            txtFormalidadNegocio4.Enabled = True
            txtColaterales4.Enabled = True
            txtDestino4.Enabled = True
         Else
            framePropuesta.Enabled = False
            txtFechaVisita.Enabled = False
            txtEntornoFamiliar4.Enabled = False
            txtGiroUbicacion4.Enabled = False
            txtExperiencia4.Enabled = False
            txtFormalidadNegocio4.Enabled = False
            txtColaterales4.Enabled = False
            txtDestino4.Enabled = False
        End If
        '*****->Fin No Refinanciados
        
        'LUCV20160626, Para CARGAR CABECERA->**********
        Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
        ActXCodCta.NroCuenta = sCtaCod
        txtGiroNeg.Text = rsCredEval!cActividad
        txtNombreCliente.Text = fsCliente
        spnExpEmpAnio.valor = rsCredEval!nExpEmpAnio
        spnExpEmpMes.valor = rsCredEval!nExpEmpMes
        spnTiempoLocalAnio.valor = rsCredEval!nTmpoLocalAnio
        spnTiempoLocalMes.valor = rsCredEval!nTmpoLocalMes
        OptCondLocal(rsCredEval!nCondiLocal).value = 1
        txtCondLocalOtros.Text = rsCredEval!cCondiLocalOtro
        txtExposicionCredito.Text = Format(rsCredEval!nExposiCred, "#,##0.00")
        txtFechaEvaluacion.Text = Format(rsCredEval!dFecEval, "dd/mm/yyyy")
        txtUltEndeuda.Text = Format(rsCredEval!nUltEndeSBS, "#,##0.00")
        txtFecUltEndeuda.Text = Format(rsCredEval!dUltEndeuSBS, "dd/mm/yyyy")
         
        'LUCV20160626, Para CARGAR PROPUESTA->**********
        If fnColocCondi <> 4 Then
            txtFechaVisita.Text = Format(rsPropuesta!dFecVisita, "dd/mm/yyyy")
            txtEntornoFamiliar4.Text = Trim(rsPropuesta!cEntornoFami)
            txtGiroUbicacion4.Text = Trim(rsPropuesta!cGiroUbica)
            txtExperiencia4.Text = Trim(rsPropuesta!cExpeCrediticia)
            txtFormalidadNegocio4.Text = Trim(rsPropuesta!cFormalNegocio)
            txtColaterales4.Text = Trim(rsPropuesta!cColateGarantia)
            txtDestino4.Text = Trim(rsPropuesta!cDestino)
            txtComentario4.Text = Trim(rsCredEval!cComentario)
        End If
        'LUCV20160626, Para la CARGAR FLEX - Mantenimiento **********->
        
        'Call FormatearGrillas(feGastosFamiliares2)
        Call LimpiaFlex(feGastosFamiliares)
            Do While Not rsDatGastoFam.EOF
                feGastosFamiliares.AdicionaFila
                lnFila = feGastosFamiliares.row
                feGastosFamiliares.TextMatrix(lnFila, 1) = rsDatGastoFam!nConsValor
                feGastosFamiliares.TextMatrix(lnFila, 2) = rsDatGastoFam!cConsDescripcion
                feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsDatGastoFam!nMonto, "#,##0.00")
                     
            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) 'celda que  o se puede editar
                Case gCodCuotaIfiGastoFami
                    Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                    Me.feGastosFamiliares.ForeColorRow (&H80000007), True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Case gCodDeudaLCNUGastoFami
                    Me.feGastosFamiliares.BackColorRow vbWhite, True
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Case Else
                    Me.feGastosFamiliares.BackColorRow (&HFFFFFF)
                    Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsDatGastoFam.MoveNext
            Loop
        rsDatGastoFam.Close
        Set rsDatGastoFam = Nothing
        
        'Call FormatearGrillas(feOtrosIngresos2)
        Call LimpiaFlex(feOtrosIngresos)
            Do While Not rsDatOtrosIng.EOF
                feOtrosIngresos.AdicionaFila
                lnFila = feOtrosIngresos.row
                feOtrosIngresos.TextMatrix(lnFila, 1) = rsDatOtrosIng!nConsValor
                feOtrosIngresos.TextMatrix(lnFila, 2) = rsDatOtrosIng!cConsDescripcion
                feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
                rsDatOtrosIng.MoveNext
            Loop
        rsDatOtrosIng.Close
        Set rsDatOtrosIng = Nothing
        
        'Call FormatearGrillas(feCuotaIfis)
        Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
            Do While Not rsCuotaIFIs.EOF
                frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
                lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsCuotaIFIs!CDescripcion
                frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsCuotaIFIs!nMonto, "#,##0.00")
                rsCuotaIFIs.MoveNext
            Loop
        rsCuotaIFIs.Close
        Set rsCuotaIFIs = Nothing
        
        'Call FormatearGrillas(feReferidos2)
        Call LimpiaFlex(feReferidos)
            Do While Not rsDatRef.EOF
                feReferidos.AdicionaFila
                lnFila = feReferidos.row
                feReferidos.TextMatrix(lnFila, 0) = rsDatRef!nCodRef
                feReferidos.TextMatrix(lnFila, 1) = rsDatRef!cNombre
                feReferidos.TextMatrix(lnFila, 2) = rsDatRef!cDniNom
                feReferidos.TextMatrix(lnFila, 3) = rsDatRef!cTelf
                feReferidos.TextMatrix(lnFila, 4) = rsDatRef!cReferido
                feReferidos.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
                rsDatRef.MoveNext
            Loop
        rsDatRef.Close
        Set rsDatRef = Nothing
        
        'Call FormatearGrillas(feDeclaracionPDT)
        'Call LimpiaFlex(feDeclaracionPDT)
        lnFila = 1
        Do While Not rsDatPDTDet.EOF
            'feDeclaracionPDT.AdicionaFila
            feDeclaracionPDT.TextMatrix(lnFila, 2) = Format(rsDatPDTDet!nConsCod, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 3) = Format(rsDatPDTDet!nConsValor, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 4) = Format(rsDatPDTDet!nMontoMes1, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 5) = Format(rsDatPDTDet!nMontoMes2, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 6) = Format(rsDatPDTDet!nMontoMes3, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 7) = Format(rsDatPDTDet!nPromedio, "#,##0.00")
            feDeclaracionPDT.TextMatrix(lnFila, 8) = Format(rsDatPDTDet!nPorcentajeVent, "#,##0.00")
            rsDatPDTDet.MoveNext
            lnFila = lnFila + 1
        Loop
        rsDatPDTDet.Close
        Set rsDatPDTDet = Nothing
        feDeclaracionPDT.TextMatrix(0, 4) = DevolverMesDatos(CInt(rsDatPDT!nMes1))
        feDeclaracionPDT.TextMatrix(0, 5) = DevolverMesDatos(CInt(rsDatPDT!nMes2))
        feDeclaracionPDT.TextMatrix(0, 6) = DevolverMesDatos(CInt(rsDatPDT!nMes3))
        
        'Call FormatearGrillas(feFlujoCajaMensual)
        Call LimpiaFlex(feFlujoCajaMensual)
            Do While Not rsDatFlujoCaja.EOF
                feFlujoCajaMensual.AdicionaFila
                lnFila = feFlujoCajaMensual.row
                feFlujoCajaMensual.TextMatrix(lnFila, 1) = rsDatFlujoCaja!nConsCod
                feFlujoCajaMensual.TextMatrix(lnFila, 2) = rsDatFlujoCaja!nConsValor
                feFlujoCajaMensual.TextMatrix(lnFila, 3) = rsDatFlujoCaja!cConcepto
                feFlujoCajaMensual.TextMatrix(lnFila, 4) = Format(rsDatFlujoCaja!nMonto, "#,##0.00")
                
                Select Case CInt(feFlujoCajaMensual.TextMatrix(feFlujoCajaMensual.row, 2)) 'celda que  o se puede editar
                    'Case 4, 5, 20, 1000 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Case 5, 6, 22, 1000 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                        Me.feFlujoCajaMensual.BackColorRow (&H80000000)
                        Me.feFlujoCajaMensual.ForeColorRow vbBlack, True
                        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                    'Case 17 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Case 18 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                        Me.feFlujoCajaMensual.ForeColorRow (&H80000007)
                        Me.feFlujoCajaMensual.BackColorRow vbWhite, True
                        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-X-X"
                    'Case 18 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
                    Case 19 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                        Me.feFlujoCajaMensual.BackColorRow &HC0FFFF, True
                        Me.feFlujoCajaMensual.ForeColorRow (&H80000007), True
                        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
                    Case Else
                        Me.feFlujoCajaMensual.BackColorRow (&HFFFFFF)
                        Me.feFlujoCajaMensual.ColumnasAEditar = "X-X-X-X-4-X"
                End Select
                rsDatFlujoCaja.MoveNext
            Loop
        rsDatFlujoCaja.Close
        Set rsDatFlujoCaja = Nothing
    
      'Call FormatearGrillas(feActivo)
        Call LimpiaFlex(feActivos)
            Do While Not rsDatActivos.EOF
                feActivos.AdicionaFila
                lnFila = feActivos.row
                feActivos.TextMatrix(lnFila, 1) = rsDatActivos!cConsDescripcion
                feActivos.TextMatrix(lnFila, 2) = Format(rsDatActivos!PP, "#,##0.00")
                feActivos.TextMatrix(lnFila, 3) = Format(rsDatActivos!PE, "#,##0.00")
                feActivos.TextMatrix(lnFila, 4) = Format(rsDatActivos!nTotal, "#,##0.00")
                feActivos.TextMatrix(lnFila, 5) = rsDatActivos!nConsCod
                feActivos.TextMatrix(lnFila, 6) = rsDatActivos!nConsValor
                
                Select Case CInt(feActivos.TextMatrix(feActivos.row, 6)) 'celda que  o se puede editar
                    Case 1000, 100, 200, 300, 400, 500
                        Me.feActivos.BackColorRow (&H80000000)
                        Me.feActivos.ForeColorRow vbBlack, True
                    Case Else
                        Me.feActivos.BackColorRow (&HFFFFFF)
                        Me.feActivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                End Select
                rsDatActivos.MoveNext
            Loop
        rsDatActivos.Close
        Set rsDatActivos = Nothing
        
        
        'Call FormatearGrillas(fePasivo)
        Call LimpiaFlex(fePasivos)
            Do While Not rsDatPasivos.EOF
                fePasivos.AdicionaFila
                lnFila = fePasivos.row
                fePasivos.TextMatrix(lnFila, 1) = rsDatPasivos!cConsDescripcion
                fePasivos.TextMatrix(lnFila, 2) = Format(rsDatPasivos!PP, "#,##0.00")
                fePasivos.TextMatrix(lnFila, 3) = Format(rsDatPasivos!PE, "#,##0.00")
                fePasivos.TextMatrix(lnFila, 4) = Format(rsDatPasivos!nTotal, "#,##0.00")
                fePasivos.TextMatrix(lnFila, 5) = rsDatPasivos!nConsCod
                fePasivos.TextMatrix(lnFila, 6) = rsDatPasivos!nConsValor
                
                Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 6)) 'celda que  o se puede editar
                    Case 1000, 1001, 1002, 100, 200, 300, 400, 500
                        Me.fePasivos.BackColorRow (&H80000000)
                        Me.fePasivos.ForeColorRow vbBlack, True
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                    Case 206
                        Me.fePasivos.BackColorRow vbWhite, True
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                    Case 109, 201
                        Me.fePasivos.BackColorRow &HC0FFFF
                        Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                    Case 301
                        Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
                    Case Else
                        Me.fePasivos.BackColorRow (&HFFFFFF)
                        Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
                End Select
                rsDatPasivos.MoveNext
            Loop
        rsDatPasivos.Close
        Set rsDatPasivos = Nothing
        'LUCV20160626, Fin Carga Flex <-**********
        
            'Carga de rsDatIfiGastoNego -> Matrix
            ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
            i = 0
            Do While Not rsDatIfiGastoNego.EOF
                MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
                MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
                MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#0.00")
                rsDatIfiGastoNego.MoveNext
                  i = i + 1
            Loop
            rsDatIfiGastoNego.Close
            Set rsDatIfiGastoNego = Nothing
    
            'Carga de rsDatIfiGastoFami -> Matrix
            ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
            j = 0
            Do While Not rsDatIfiGastoFami.EOF
                MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
                MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
                MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#0.00")
                rsDatIfiGastoFami.MoveNext
                j = j + 1
            Loop
            rsDatIfiGastoFami.Close
            Set rsDatIfiGastoFami = Nothing
    
        'LUCV20160628, Para CARGA RATIOS/INDICADORES
        txtCapacidadNeta.Text = CStr(rsDatRatioInd!nCapPagNeta * 100) & "%"
        txtEndeudamiento.Text = CStr(rsDatRatioInd!nEndeuPat * 100) & "%"
        txtLiquidezCte.Text = CStr(Format(rsDatRatioInd!nLiquidezCte, "#0.00"))
        txtRentabilidadPat.Text = CStr(rsDatRatioInd!nRentaPatri * 100) & "%"
        txtIngresoNeto.Text = Format(rsDatRatioInd!nIngreNeto, "#,##0.00")
        txtExcedenteMensual.Text = Format(rsDatRatioInd!nExceMensual, "#,##0.00")
        
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        If Not (rsDatParamFlujoCajaForm4.BOF And rsDatParamFlujoCajaForm4.EOF) Then
            EditMoneyIncVC4.Text = rsDatParamFlujoCajaForm4!nIncVentCont
            EditMoneyIncCM4.Text = rsDatParamFlujoCajaForm4!nIncCompMerc
            EditMoneyIncPP4.Text = rsDatParamFlujoCajaForm4!nIncPagPers
            EditMoneyIncGV4.Text = rsDatParamFlujoCajaForm4!nIncGastvent
            EditMoneyIncC4.Text = rsDatParamFlujoCajaForm4!nIncConsu
        End If
        Set rsDatParamFlujoCajaForm4 = Nothing
       'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        
    Set rsDCredito = Nothing
End Function


Private Sub GeneraVerCar()
    Dim oCred As COMNCredito.NCOMFormatosEval
    Dim oDCredSbs As COMDCredito.DCOMFormatosEval
    Dim R As ADODB.Recordset
    Dim lcDNI, lcRUC As String
    Dim RSbs, RDatFin1, RCap As ADODB.Recordset
    Set oCred = New COMNCredito.NCOMFormatosEval
    Call oCred.RecuperaDatosInformeComercial(ActXCodCta.NroCuenta, R)
    Set oCred = Nothing
    
    If R.EOF And R.BOF Then
    MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
    Exit Sub
    End If
    
    lcDNI = Trim(R!dni_deudor)
    lcRUC = Trim(R!ruc_deudor)
    Set oDCredSbs = New COMDCredito.DCOMFormatosEval
    Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC)
    Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActXCodCta.NroCuenta, fnFormato)
    Set oDCredSbs = Nothing
    Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActXCodCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1)
End Sub


Private Sub ImprimirFormatoEvaluacion()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    
    Dim rsMostrarCuotasIfis As ADODB.Recordset
    Dim rsMostrarCuotasIfisGF As ADODB.Recordset
    Dim rsRatiosIndicadores As ADODB.Recordset
    
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    'Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
    Set rsInfVisita = oDCOMFormatosEval.MostrarFormatoSinConvenioInfVisCabecera(sCtaCod, fnFormato)
    
    Set rsMostrarCuotasIfis = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7022)
    Set rsMostrarCuotasIfisGF = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7023)
    Set rsRatiosIndicadores = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
    
    Dim A As Currency
    Dim nFila As Integer
    
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de Visita Nº " & sCtaCod
    oDoc.Title = "Informe de Visita Nº " & sCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoEvaluacion_" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
    If Not (rsInfVisita.BOF Or rsInfVisita.EOF) Then
        'Tamaño de hoja A4
        oDoc.NewPage A4_Vertical
    
        'Call CabeceraImpCuadros(rsInfVisita)
            '---------- cabecera
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox 120, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        oDoc.WTextBox 130, 55, 1, 160, "ACTIVOS", "F2", 7.5, hjustify
        oDoc.WTextBox 140, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = 140
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F2", 7.5, hRight
        
            For i = 1 To feActivos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feActivos.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feActivos.TextMatrix(i, 2), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(feActivos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(feActivos.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "PASIVOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "P.P", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240, 1, 160, "P.E.", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340, 1, 160, "TOTAL", "F2", 7.5, hRight
        
            For i = 1 To fePasivos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, fePasivos.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(fePasivos.TextMatrix(i, 2), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250, 15, 150, Format(fePasivos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350, 15, 150, Format(fePasivos.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        oDoc.NewPage A4_Vertical
            '---------- cabecera
        
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
    
        nFila = 140
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "FLUJO DE CAJA MENSUAL", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        A = 0
            For i = 1 To feFlujoCajaMensual.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feFlujoCajaMensual.TextMatrix(i, 3), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feFlujoCajaMensual.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
                A = A + feFlujoCajaMensual.TextMatrix(i, 4)
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "Total" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        
        oDoc.WTextBox nFila, 55, 1, 200, "FLUJO DE CAJA MENSUAL - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        A = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For i = 1 To rsMostrarCuotasIfis.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
                A = A + rsMostrarCuotasIfis!nMonto
                rsMostrarCuotasIfis.MoveNext
                nFila = nFila + 10
            Next i
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
    
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        A = 0
            For i = 1 To feGastosFamiliares.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feGastosFamiliares.TextMatrix(i, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosFamiliares.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                A = A + feGastosFamiliares.TextMatrix(i, 3)
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "Total" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        
        oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES  - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        A = 0
        If Not (rsMostrarCuotasIfisGF.BOF And rsMostrarCuotasIfisGF.EOF) Then
            For i = 1 To rsMostrarCuotasIfisGF.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfisGF!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfisGF!CDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfisGF!nMonto, "#,##0.00"), "F1", 7.5, hRight
                A = A + rsMostrarCuotasIfisGF!nMonto
                nFila = nFila + 10
                rsMostrarCuotasIfisGF.MoveNext
            Next i
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
        
            
        oDoc.WTextBox nFila, 55, 1, 160, "OTROS INGRESOS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
        A = 0
            For i = 1 To feOtrosIngresos.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feOtrosIngresos.TextMatrix(i, 2), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150, 15, 150, Format(feOtrosIngresos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
                A = A + feOtrosIngresos.TextMatrix(i, 3)
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 140, 1, 160, "Total" & Space(10) & Format(A, "#,##0.00"), "F2", 7.5, hRight
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        '----------------------------------------------------------------------------------------------------------------
        
        If nFila >= 700 Then
        
        'Tamaño de hoja A4
        oDoc.NewPage A4_Vertical
        
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
        
        
         nFila = 110
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
            oDoc.WTextBox nFila, 55, 1, 160, "DECLARACION PDT", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 4), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 5), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 6), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 440 - 80, 1, 160, "MONTO", "F2", 7.5, hRight
        
            For i = 1 To feDeclaracionPDT.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feDeclaracionPDT.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 5), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 6), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 450 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 7), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 550 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 8), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
    
        oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
        oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
        oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
    
        oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
        oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
        '----------------------------------------------------------------------------------------------------------------
        Else
        
        oDoc.WTextBox nFila, 55, 1, 160, "DECLARACION PDT", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
        oDoc.WTextBox nFila, 140 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 4), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 240 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 5), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 340 - 80, 1, 160, feDeclaracionPDT.TextMatrix(0, 6), "F2", 7.5, hRight
        oDoc.WTextBox nFila, 440 - 80, 1, 160, "MONTO", "F2", 7.5, hRight
        
            For i = 1 To feDeclaracionPDT.rows - 1
                nFila = nFila + 10
                oDoc.WTextBox nFila, 55, 15, 250, feDeclaracionPDT.TextMatrix(i, 1), "F1", 7.5, hLeft
                oDoc.WTextBox nFila, 150 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 4), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 250 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 5), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 350 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 6), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 450 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 7), "#,#0.00"), "F1", 7.5, hRight
                oDoc.WTextBox nFila, 550 - 80, 15, 150, Format(feDeclaracionPDT.TextMatrix(i, 8), "#,#0.00"), "F1", 7.5, hRight
            Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        '----------------------------------------------------------------------------------------------------------------
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de pago", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Patrimonial", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Cte.", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
    
        oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
        oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
        oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
    
        oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
        oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
        '----------------------------------------------------------------------------------------------------------------
        End If
        oDoc.PDFClose
        oDoc.Show
    Else
        MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If
    Set rsInfVisita = Nothing
End Sub



