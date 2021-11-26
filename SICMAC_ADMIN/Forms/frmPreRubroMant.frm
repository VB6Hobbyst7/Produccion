VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreRubroMant 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "frmPreRubroMant.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPresupuesto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Presupuesto"
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
      Height          =   795
      Left            =   15
      TabIndex        =   42
      Top             =   -15
      Width           =   8940
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   350
      Left            =   1065
      TabIndex        =   41
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   350
      Left            =   45
      TabIndex        =   40
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   350
      Left            =   3090
      TabIndex        =   39
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   350
      Left            =   2070
      TabIndex        =   38
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   1065
      TabIndex        =   37
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   45
      TabIndex        =   36
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   350
      Left            =   8055
      TabIndex        =   35
      Top             =   6840
      Width           =   975
   End
   Begin VB.Frame fraConceptos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Rubros"
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
      Height          =   5955
      Left            =   15
      TabIndex        =   0
      Top             =   825
      Width           =   9030
      Begin VB.Frame fraOperadores 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Operadores"
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
         Height          =   1935
         Left            =   6045
         TabIndex        =   24
         Top             =   2940
         Width           =   2910
         Begin VB.ComboBox cmbOpeInterna 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   225
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeAri 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   540
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeLog 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   870
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeCad 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1215
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeFec 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1545
            Width           =   2070
         End
         Begin VB.Label lblOpeAri 
            Caption         =   "Arimet."
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   60
            TabIndex        =   34
            Top             =   585
            Width           =   750
         End
         Begin VB.Label lblOpeLog 
            Caption         =   "Logicos"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   33
            Top             =   915
            Width           =   1215
         End
         Begin VB.Label lblOpeCad 
            Caption         =   "Cadena"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   75
            TabIndex        =   32
            Top             =   1275
            Width           =   1215
         End
         Begin VB.Label lblOpeFec 
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   90
            TabIndex        =   31
            Top             =   1590
            Width           =   1335
         End
         Begin VB.Label lblOpeInterna 
            Caption         =   "Internos"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   45
            TabIndex        =   30
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Frame fraCamposTabla 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Tabla/Campos"
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
         Height          =   975
         Left            =   6045
         TabIndex        =   19
         Top             =   4890
         Width           =   2910
         Begin VB.ComboBox cmbTablas 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   225
            Width           =   2070
         End
         Begin VB.ComboBox cmbCampos 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   2070
         End
         Begin VB.Label lblTablas 
            Caption         =   "Tablas"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   285
            Width           =   615
         End
         Begin VB.Label lblCampos 
            Caption         =   "Campos"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   75
            TabIndex        =   22
            Top             =   660
            Width           =   675
         End
      End
      Begin VB.Frame fraConcepto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Concepto"
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
         Height          =   2895
         Left            =   75
         TabIndex        =   17
         Top             =   2970
         Width           =   5895
         Begin MSComctlLib.ListView lvwCon 
            Height          =   2610
            Left            =   45
            TabIndex        =   18
            Top             =   225
            Width           =   5730
            _ExtentX        =   10107
            _ExtentY        =   4604
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Height          =   2700
         Left            =   75
         TabIndex        =   1
         Top             =   225
         Width           =   8895
         Begin VB.TextBox txtOrden 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7170
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "0"
            Top             =   253
            Width           =   630
         End
         Begin VB.ComboBox cmbConcep 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   870
            Width           =   5460
         End
         Begin VB.TextBox txtNomCon 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   750
            MaxLength       =   50
            TabIndex        =   6
            Top             =   255
            Width           =   5460
         End
         Begin VB.ComboBox cmbGrupo 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   555
            Width           =   5460
         End
         Begin VB.TextBox txtNemonico 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   7170
            MaxLength       =   50
            TabIndex        =   4
            Top             =   572
            Width           =   1620
         End
         Begin VB.TextBox txtNomImpre 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   7170
            MaxLength       =   50
            TabIndex        =   3
            Top             =   900
            Width           =   1620
         End
         Begin VB.TextBox txtForEdit 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   75
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   1440
            Width           =   8730
         End
         Begin VB.Label lblCodigoL 
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
            Left            =   8010
            TabIndex        =   16
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lblTipConcep 
            Caption         =   "Tipo:"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   15
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblNomConcep 
            Caption         =   "Nombre"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   14
            Top             =   290
            Width           =   735
         End
         Begin VB.Label lblGrupo 
            Caption         =   "Grupo"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   13
            Top             =   585
            Width           =   615
         End
         Begin VB.Label lblNemoConep 
            Caption         =   "Nemo:"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6330
            TabIndex        =   12
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblOrden 
            Caption         =   "Orden"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   6330
            TabIndex        =   11
            Top             =   290
            Width           =   615
         End
         Begin VB.Label lblNomImpre 
            Caption         =   "Impresión"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6330
            TabIndex        =   10
            Top             =   900
            Width           =   870
         End
         Begin VB.Label lblFormula 
            Caption         =   "Formula"
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
            Height          =   210
            Left            =   90
            TabIndex        =   9
            Top             =   1230
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frmPreRubroMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
