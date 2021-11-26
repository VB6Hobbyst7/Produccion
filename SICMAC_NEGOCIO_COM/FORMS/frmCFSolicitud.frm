VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCFSolicitud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Carta Fianza"
   ClientHeight    =   8535
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   7920
   Icon            =   "frmCFSolicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   7920
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Avalado "
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
      Height          =   675
      Left            =   60
      TabIndex        =   50
      Top             =   3120
      Width           =   7800
      Begin VB.TextBox txtConsorcio 
         Height          =   285
         Left            =   1080
         TabIndex        =   56
         Top             =   240
         Width           =   6135
      End
      Begin VB.CheckBox chkConsorcio 
         Caption         =   "Consorcio"
         Height          =   195
         Left            =   1080
         TabIndex        =   55
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarAvalado 
         BackColor       =   &H80000004&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7185
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Buscar al Acreedor"
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Avalado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   53
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblNomAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2220
         TabIndex        =   52
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame FraRenovacion 
      Enabled         =   0   'False
      Height          =   615
      Left            =   60
      TabIndex        =   45
      Top             =   7320
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdBuscar 
         Enabled         =   0   'False
         Height          =   345
         Left            =   5640
         Picture         =   "frmCFSolicitud.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Buscar ..."
         Top             =   200
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CheckBox ChkRenovacion 
         Caption         =   "Renovación"
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
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1455
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   1920
         TabIndex        =   48
         Top             =   170
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Credito"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   60
      TabIndex        =   44
      Top             =   75
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Frame fraCreditos 
      Height          =   3645
      Left            =   60
      TabIndex        =   28
      Top             =   3645
      Width           =   7800
      Begin VB.Frame frFinalidad 
         Height          =   1120
         Left            =   120
         TabIndex        =   60
         Top             =   2470
         Width           =   7575
         Begin VB.TextBox TxtFinalidad 
            Height          =   750
            Left            =   50
            MaxLength       =   700
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   320
            Width           =   7470
         End
         Begin VB.Label lblFinalidad 
            AutoSize        =   -1  'True
            Caption         =   "Finalidad "
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   5
            Left            =   50
            TabIndex        =   62
            Top             =   100
            Width           =   675
         End
      End
      Begin VB.Frame frModOtrs 
         Height          =   500
         Left            =   120
         TabIndex        =   57
         Top             =   1990
         Visible         =   0   'False
         Width           =   5940
         Begin VB.TextBox txtModOtrs 
            Height          =   300
            Left            =   1620
            TabIndex        =   59
            Top             =   150
            Width           =   4230
         End
         Begin VB.Label Label4 
            Caption         =   "Modalidad Otros"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   50
            TabIndex        =   58
            Top             =   195
            Width           =   1350
         End
      End
      Begin VB.TextBox TxtPeriodo 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6285
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1320
         Width           =   660
      End
      Begin VB.Frame fraLineaCred 
         Height          =   645
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   7590
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   6120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1380
         End
         Begin VB.ComboBox cboTipoCF 
            Height          =   315
            Left            =   1620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   5370
            TabIndex        =   31
            Top             =   285
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Sub Producto"
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   30
            Top             =   270
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.ComboBox cboAnalista 
         Height          =   315
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1665
         Width           =   3120
      End
      Begin VB.TextBox txtMontoSol 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1770
         MaxLength       =   15
         TabIndex        =   8
         Top             =   855
         Width           =   1260
      End
      Begin VB.ComboBox cboModalidad 
         Height          =   315
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1245
         Width           =   3120
      End
      Begin MSMask.MaskEdBox txtfechaAsig 
         Height          =   315
         Left            =   6285
         TabIndex        =   9
         Top             =   915
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaVencimiento 
         Height          =   315
         Left            =   6285
         TabIndex        =   12
         Top             =   1740
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   49
         Top             =   1365
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Analista Responsable "
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   36
         Top             =   1710
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto Solicitado "
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
         Index           =   1
         Left            =   165
         TabIndex        =   35
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asignación "
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
         Index           =   4
         Left            =   5175
         TabIndex        =   34
         Top             =   975
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad de Fianza"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   33
         Top             =   1305
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
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
         Index           =   0
         Left            =   5175
         TabIndex        =   32
         Top             =   1800
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "E&xaminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   105
      Width           =   1215
   End
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   60
      TabIndex        =   25
      Top             =   7920
      Width           =   7800
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   15
         Top             =   165
         Width           =   885
      End
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   27
         Top             =   165
         Width           =   885
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6800
         TabIndex        =   19
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Registrar Cobertura>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4645
         TabIndex        =   18
         Top             =   180
         Width           =   2105
      End
      Begin VB.CommandButton cmdGarantias 
         Caption         =   "Garan&tías..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3390
         TabIndex        =   17
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2460
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         TabIndex        =   14
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         TabIndex        =   26
         Top             =   165
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdRelaciones 
      Caption         =   "&Relaciones"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6495
      TabIndex        =   2
      Top             =   1380
      Width           =   1230
   End
   Begin VB.ComboBox cboCondicion 
      Height          =   315
      Left            =   4695
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1860
   End
   Begin VB.Frame fraAcreedor 
      Caption         =   "Acreedor"
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
      Height          =   675
      Left            =   60
      TabIndex        =   20
      Top             =   2490
      Width           =   7800
      Begin VB.CommandButton cmdBuscarAcreedor 
         BackColor       =   &H80000004&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7185
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Buscar al Acreedor"
         Top             =   225
         Width           =   420
      End
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2220
         TabIndex        =   23
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   22
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   270
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView ListaRelacion 
      Height          =   1020
      Left            =   300
      TabIndex        =   24
      Top             =   1005
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre de Cliente"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Relación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nº Cuenta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Valor. Rel."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nº D.I."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Nº D.T."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Tipo Persona"
         Object.Width           =   0
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   585
      Left            =   7380
      TabIndex        =   42
      Top             =   5670
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   1032
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCFSolicitud.frx":040C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fracliente 
      Caption         =   "Afianzado"
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
      Height          =   1950
      Left            =   60
      TabIndex        =   37
      Top             =   495
      Width           =   7770
      Begin VB.ComboBox cboFuentes 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   4380
      End
      Begin VB.CommandButton cmdFuentes 
         Caption         =   "&Fuentes Ingreso"
         Height          =   330
         Left            =   6000
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado "
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   40
         Tag             =   "txtcodigo"
         Top             =   195
         Width           =   1275
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   39
         Tag             =   "txtnombre"
         Top             =   195
         Width           =   5115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fuentes Ingreso "
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   38
         Top             =   1590
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Condición:"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3885
      TabIndex        =   43
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "frmCFSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oRelPersCred As UCredRelac_Cli
Private cmdEjecutar As Integer
Private oPersBuscada As COMDPersona.UCOMPersona
Dim objPista As COMManejador.Pista
Dim bConsorcio As Boolean 'WIOR 20120328
Dim nPeriodoMax As Long 'JOEP20181222 CP
Dim nPeriodoMin As Long 'JOEP20181222 CP

Private Sub LimpiaPantalla()
    Set oRelPersCred = New UCredRelac_Cli
    Call LimpiaControles(Me, True)
End Sub

Private Sub CargaDatos(ByVal psCta As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
Dim R As New ADODB.Recordset

    Set oCF = New COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
     Set R = oCF.RecuperaCartaFianzaSolicitud(psCta)
    Set oCF = Nothing
    If R Is Nothing Then
        MsgBox "No Se encuentra Informacion de la Carta Fianza"
        Exit Sub
    End If
    If Not (R.BOF And R.EOF) Then
        Call CP_CargaDatos 'JOEP20181227 CP
        cboCondicion.ListIndex = IndiceListaCombo(cboCondicion, R!nColocCondicion)
        lblCodigo.Caption = R!cPersCod
        lblNombre.Caption = PstaNombre(R!cPersNombre)
        
        Set oRelPersCred = New UCredRelac_Cli
        Call oRelPersCred.CargaRelacPersCred(psCta)
        Call ActualizarListaPersRelacCred
        Call CargaFuentesIngreso(oRelPersCred.TitularPersCod)
        'cboFuentes.ListIndex = IndiceListaCombo(cboFuentes, R!cNumFuente) 'LUCV20160919, Comentó Según ERS004-2016
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        lblCodAvalado.Caption = IIf(IsNull(R!cPersAvalado), "", R!cPersAvalado) 'MADM 20111020
        'WIOR 20120328-INICIO
        If R!cConsorcio <> "" Then
            Me.chkConsorcio.value = 1
            Me.txtConsorcio.Text = R!cConsorcio
        End If
       'WIOR -FIN
        If R!cAvalNombre <> "" Then
            Me.chkConsorcio.value = 0
            lblNomAvalado.Caption = IIf(IsNull(PstaNombre(R!cAvalNombre)), "", PstaNombre(R!cAvalNombre)) 'MADM 20111020
        End If
        'MAVM 20100615  BAS II
        'cboTipoCF.ListIndex = IndiceListaCombo(cboTipoCF, Mid(psCta, 6, 3))
        cboTipoCF.ListIndex = IndiceListaCombo(cboTipoCF, COMDConstantes.gColCFTpoProducto)
        cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Mid(psCta, 9, 1))
        cboModalidad.ListIndex = IndiceListaCombo(cboModalidad, R!nModalidad)
        cboAnalista.ListIndex = IndiceListaCombo(cboAnalista, R!cAnalista)
        txtfechaAsig.Text = Format(R!dAsignacion, "dd/mm/yyyy")
        txtFechaVencimiento.Text = Format(R!dVencimiento, "dd/mm/yyyy")
        TxtFinalidad.Text = Trim(R!cfinalidad)
        txtMontoSol.Text = Format(R!nMonto, "#0.00")
    'JOEP20181220 CP
        TxtPeriodo.Text = R!nPeriodo
        If R!OtrsModalidades <> "" Then
            frModOtrs.Visible = True
            Me.Width = 7965
            Me.Height = 8350
            frFinalidad.top = 2480
            fraCreditos.Height = 3700
            fracontrol.top = 7300
            txtModOtrs.Text = R!OtrsModalidades
        Else
            txtModOtrs.Text = ""
            frModOtrs.Visible = False
            Me.Width = 7965
            Me.Height = 7800
            frFinalidad.top = 2000
            fraCreditos.Height = 3200
            fracontrol.top = 6800
        End If
    'JOEP20181220 CP
    Else
        MsgBox "No Se encuentra Informacion de la Carta Fianza"
        Exit Sub
    End If
    R.Close
    Set R = Nothing
    'CMAC CUSCO
    Set R = New ADODB.Recordset
    Set oCF = New COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
     Set R = oCF.ObtenerCartaFianzaRenovacion(psCta)
    Set oCF = Nothing
    If Not (R.EOF And R.BOF) Then
        AXCodCta.NroCuenta = ""
        ChkRenovacion.value = 1
        AXCodCta.NroCuenta = Trim(R!cCtaCodAnterior)
    End If
    R.Close
    Set R = Nothing
End Sub

Private Function ValidaDatos() As Boolean
Dim dFteFecEval As Date
Dim dFteFecCaduc As Date
Dim sCad As String
Dim oPersona As UPersona_Cli 'COMDPersona.DCOMPersona
Dim oCF As COMDCartaFianza.DCOMCartaFianza
    ValidaDatos = True
    
    'Valida Condicion del Credito
    If cboCondicion.ListIndex = -1 Then
        MsgBox "Seleccione la Condicion del Credito", vbInformation, "Aviso"
        cboCondicion.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida Existencia del Titular
    If Not oRelPersCred.ExisteTitular Then
        MsgBox "La Carta Fianza Debe Poseer un Titular o Afianzado", vbInformation, "Aviso"
        Call cmdRelaciones_Click
        ValidaDatos = False
        Exit Function
    End If
    'Valida la Fuente de Ingreso
    'RECO20160914 *********************************
'    If cboFuentes.ListIndex = -1 Then
'        MsgBox "Seleccione una Fuente de Ingreso", vbInformation, "Aviso"
'        cmdFuentes.SetFocus
'        validaDatos = False
'        Exit Function
'    End If
    
    'Valida la Fecha de Caducidad del la Fuente de Ingreso
   
'    Dim rsFteIng As ADODB.Recordset
'    Dim rsFIDep As ADODB.Recordset
'    Dim rsFIInd As ADODB.Recordset
'    Dim oCred As COMNCredito.NCOMCredito
'
'    Set oCred = New COMNCredito.NCOMCredito
'
'    Call oCred.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, oRelPersCred.TitularPersCod, , cboFuentes.ListIndex)
'
'    Set oPersona = New UPersona_Cli
'    Call oPersona.RecuperaFtesdeIngreso(oRelPersCred.TitularPersCod, rsFteIng)
'    Call oPersona.RecuperaFtesIngresoDependiente(cboFuentes.ListIndex, rsFIDep)
'    Call oPersona.RecuperaFtesIngresoIndependiente(cboFuentes.ListIndex, rsFIInd)
'    dFteFecEval = oPersona.ObtenerFteIngFecEval(cboFuentes.ListIndex, IIf(oPersona.ObtenerFteIngIngresoTipo(Me.cboFuentes.ListIndex) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(cboFuentes.ListIndex) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(cboFuentes.ListIndex) - 1), oPersona.ObtenerFteIngIngresoTipo(Me.cboFuentes.ListIndex))
'    dFteFecCaduc = oPersona.ObtenerFteIngFecCaducac(cboFuentes.ListIndex, IIf(oPersona.ObtenerFteIngIngresoTipo(cboFuentes.ListIndex) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(cboFuentes.ListIndex) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(cboFuentes.ListIndex) - 1))
'
'    If gdFecSis > dFteFecCaduc Then
'        MsgBox "La Fuente de Ingreso es muy Antigua debe Ingresar otra Fuente de Ingreso", vbInformation, "Aviso"
'        cmdFuentes.SetFocus
'        validaDatos = False
'        Exit Function
'    End If
    'Valida si se ingreso acreedor
    If Trim(lblCodAcreedor.Caption) = "" Then
        MsgBox "Falta Ingresar el Acreedor de la Carta Fianza", vbInformation, "Aviso"
        cmdBuscarAcreedor.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida si se ingreso el Tipo de Carta Fianza
    If cboTipoCF.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Tipo de Carta Fianza", vbInformation, "Aviso"
        'cboTipoCF.SetFocus
        ValidaDatos = False
        cboTipoCF.ListIndex = IndiceListaCombo(cboTipoCF, COMDConstantes.gColCFTpoProducto)
        Exit Function
    End If
    'Valida Si se Selecciono la Moneda
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Moneda para la Carta Fianza", vbInformation, "Aviso"
        cboMoneda.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida que se halla Ingresado el Monto y sea Mayor que Cero
    If Trim(txtMontoSol.Text) = "" Then
        MsgBox "Falta Ingresar el Monto para la Carta Fianza", vbInformation, "Aviso"
        txtMontoSol.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida que se halla Ingresado el Monto y sea Mayor que Cero
    If CDbl(txtMontoSol.Text) <= 0 Then
        MsgBox "Monto de la Carta Fianza debe ser Mayor a Cero", vbInformation, "Aviso"
        txtMontoSol.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida si se selecciono la Modalidad de Carta Fianza
    If cboModalidad.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Modalidad de la Carta Fianza", vbInformation, "Aviso"
        cboModalidad.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida si se selecciono al Analista
    If cboAnalista.ListIndex = -1 Then
        MsgBox "Falta Seleccionar al Analista de la Carta Fianza", vbInformation, "Aviso"
        cboAnalista.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida si la Fecha Asignacion es Correcta
    sCad = ValidaFecha(txtfechaAsig.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtfechaAsig.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida si la Fecha de Vencimiento es Correcta
    sCad = ValidaFecha(txtFechaVencimiento.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtFechaVencimiento.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    'Valida que el Titular sea diferente del Acreedor
    If Trim(lblCodAcreedor.Caption) = Trim(lblCodigo.Caption) Then
        MsgBox "El Acreedor no puede ser el mismo que el Afianzado", vbInformation, "Aviso"
        cmdBuscarAcreedor.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If CDate(Format(gdFecSis, "dd/mm/yyyy")) > CDate(Format(txtFechaVencimiento.Text, "dd/mm/yyyy")) Then
        MsgBox "La Fecha de Vencimiento debe ser mayor a la fecha actual", vbInformation, "Aviso"
        txtFechaVencimiento.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    'By capi 26022009 para validar ingreso de finalidad de la carta fianza
    'WIOR 20120613 Aumento a 700 caracteres de la finalidad
    If Len(Trim(TxtFinalidad.Text)) >= 700 Then
        MsgBox "El texto de Finalidad no debe superar 700 caracteres, Resumir", vbInformation, "Aviso"
        'ALPA 20090417*********************
        TxtFinalidad.SetFocus
        'TxtFinalidad.Text.SetFocus
        '**********************************
        ValidaDatos = False
        Exit Function
    End If
        
    'End by
  
    
    'WIOR 20120328-INICIO-Valida que en caso de seleccionar consorcio se llene el campo consorcio
    If Me.chkConsorcio.value = 1 Then
        If Trim(Me.txtConsorcio.Text) = "" Then
            MsgBox "Debe Registrar el nombre del consorcio", vbInformation, "Aviso"
            Me.txtConsorcio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Len(Trim(txtConsorcio.Text)) > 500 Then
            MsgBox "El texto de Consorcio no debe superar 500 caracteres, Resumir", vbInformation, "Aviso"
            txtConsorcio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    'WIOR -FIN
    
    
    
    If CDate(Format(txtFechaVencimiento.Text, "dd/mm/yyyy")) < CDate(Format(txtfechaAsig, "dd/mm/yyyy")) Then
        MsgBox "Fecha de Asignación no puede ser mayor que Fecha de Vencimiento", vbInformation, "Aviso"
        txtfechaAsig.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    Set oPersona = Nothing
    
    If cmdEjecutar = 1 Then
        Set oCF = New COMDCartaFianza.DCOMCartaFianza
            If Len(Trim(Me.AXCodCta.NroCuenta)) = 18 Then
               If oCF.VerficarCartaFianzaRenovacion(Me.AXCodCta.NroCuenta) = False Then
                   MsgBox "Ya existe una Renovacion con este Nro de Carta Fianza", vbInformation, "Aviso"
                    Me.AXCodCta.Age = ""
                    Me.AXCodCta.Cuenta = ""
                    Me.AXCodCta.Prod = ""
                    Me.AXCodCta.CMAC = ""
                    Me.AXCodCta.Enabled = False
                    Me.cmdBuscar.Enabled = False
                    Me.ChkRenovacion.value = 0
                   ValidaDatos = False
                   Exit Function
               End If
            End If
        Set oCF = Nothing
    End If

'JOEP20181227 CP
If Not CP_ValidaMsg(3) Then
    ValidaDatos = False
    Exit Function
End If

If Not CP_ValidaMsg(1) Then
    ValidaDatos = False
    Exit Function
End If
'JOEP20181227 CP
End Function

Private Sub CargaFuentesIngreso(ByVal psPersCod As String)
Dim i As Integer

Dim oPersona As COMDPersona.DCOMPersona
    On Error GoTo ErrorCargaFuentesIngreso
  
    Set oPersona = New COMDPersona.DCOMPersona
    Call oPersona.RecuperaPersona(psPersCod)
    'Call oPersona.RecuperaFtesdeIngreso(psPersCod)
    oPersona.PersCodigo = oRelPersCred.TitularPersCod
    cboFuentes.Clear
    For i = 0 To oPersona.NumeroFtesIngreso - 1
        cboFuentes.AddItem oPersona.ObtenerFteIngRazonSocial(i) & Space(100 - Len(oPersona.ObtenerFteIngRazonSocial(i))) & oPersona.ObtenerFteIngFuente(i) & Space(50 - Len(oPersona.ObtenerFteIngFuente(i))) & oPersona.ObtenerFteIngcNumFuente(i)
    Next i
    If cboFuentes.ListCount > 0 Then
        cboFuentes.ListIndex = 0
    End If
    'Set oPersona = Nothing
   
    Exit Sub
   
ErrorCargaFuentesIngreso:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub


Private Sub ActualizarListaPersRelacCred()
    Dim s As ListItem
    Dim oPREDA As COMDPersona.DCOMPersonas 'JUEZ 20130717
    ListaRelacion.ListItems.Clear
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        'JUEZ 20130717 ********************************************************
        Set oPREDA = New COMDPersona.DCOMPersonas
        If oPREDA.VerificarPersonaPREDA(oRelPersCred.ObtenerCodigo, 1) Then
            MsgBox "El " & IIf(oRelPersCred.ObtenerValorRelac = CStr(gColRelPersTitular), "Titular", "cliente") & " " & oRelPersCred.ObtenerNombre & " es un cliente PREDA no sujeto de Crédito, consultar a Coordinador de Producto Agropecuario", vbInformation, "Aviso"
            If oRelPersCred.ObtenerValorRelac = CStr(gColRelPersTitular) Then
                ListaRelacion.ListItems.Clear
                cmdCancela_Click
                Exit Sub
            End If
            Call oRelPersCred.EliminarRelacion(oRelPersCred.ObtenerCodigo, oRelPersCred.ObtenerValorRelac)
        Else
        'END JUEZ *************************************************************
            Set s = ListaRelacion.ListItems.Add(, , oRelPersCred.ObtenerNombre)
            s.SubItems(1) = oRelPersCred.ObtenerRelac
            s.SubItems(2) = oRelPersCred.ObtenerCodigo
            s.SubItems(3) = oRelPersCred.ObtenerValorRelac
        End If
        oRelPersCred.siguiente
    Loop
   
End Sub

Private Sub HabilitaActualizacion(ByVal pbHabilita As Boolean)
    cmdExaminar.Enabled = Not pbHabilita
    ActXCodCta.Enabled = Not pbHabilita
    cboCondicion.Enabled = pbHabilita
    cmdRelaciones.Enabled = pbHabilita
    cboFuentes.Enabled = pbHabilita
    cmdFuentes.Enabled = pbHabilita
    cmdBuscarAcreedor.Enabled = pbHabilita
    cboTipoCF.Enabled = pbHabilita
    cboMoneda.Enabled = pbHabilita
    txtMontoSol.Enabled = pbHabilita
    txtfechaAsig.Enabled = pbHabilita
    cboModalidad.Enabled = pbHabilita
    cboAnalista.Enabled = pbHabilita
    TxtFinalidad.Enabled = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
    cmdGrabar.Enabled = pbHabilita
    cmdEditar.Visible = Not pbHabilita
    cmdCancela.Visible = pbHabilita
    cmdImprimir.Enabled = Not pbHabilita
    cmdImprimir.Visible = True
    cmdGarantias.Enabled = Not pbHabilita
    cmdGravar.Enabled = Not pbHabilita
    cmdsalir.Enabled = Not pbHabilita
    txtFechaVencimiento.Enabled = pbHabilita
    ListaRelacion.Enabled = pbHabilita
    FraRenovacion.Enabled = pbHabilita
    fraLineaCred.Enabled = pbHabilita
    Me.cmdBuscarAvalado.Enabled = pbHabilita 'WIOR 20120328
    Me.chkConsorcio.Enabled = pbHabilita 'WIOR 20120328
    txtConsorcio.Enabled = pbHabilita 'WIOR 20120328
    txtModOtrs.Enabled = pbHabilita 'JOEP20181220 CP
    TxtPeriodo.Enabled = pbHabilita 'JOEP20181220 CP
End Sub

Private Sub CargaAnalistas()
Dim R As ADODB.Recordset
Dim oDatos As COMDCredito.DCOMCredito
    
    Set oDatos = New COMDCredito.DCOMCredito
    Set R = oDatos.CargaAnalistas
    cboAnalista.Clear
    Do While Not R.EOF
        cboAnalista.AddItem PstaNombre(R!cPersNombre) & Space(100) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oDatos = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaControles()
Dim RTemp As ADODB.Recordset
Dim oCred As COMDCredito.DCOMCredito

    On Error GoTo ERRORCargaControles

    'Carga Condiciones de un Credito
    Call CargaComboConstante(gColocCredCondicion, cboCondicion)
    Call CambiaTamañoCombo(cboCondicion)

    'Carga Monedas
    Call CP_CargaCombox(5000) 'JOEP20181218 CP
    'Call CargaComboConstante(gMoneda, cboMoneda) 'Comento JOEP20181218 CP
    'Carga Analistas
    Call CargaAnalistas
    'Carga Tipos de Carta Fianza
    Set oCred = New COMDCredito.DCOMCredito
        Set RTemp = oCred.RecuperaProductosDeSolicitudDeCF
    Set oCred = Nothing
    cboTipoCF.Clear
    Do While Not RTemp.EOF
        cboTipoCF.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        cboTipoCF.ListIndex = 0
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    If cboTipoCF.ListCount > 0 Then
        cboTipoCF.ListIndex = 0
    End If
    Call CambiaTamañoCombo(cboTipoCF, 300)

    'Carga Modalidad de Carta Fianza
    Call CP_CargaCombox(49000) 'JOEP20181218 CP
    'Call CargaComboConstante(gColCFModalidad, cboModalidad) 'Comento JOEP20181218 CP
    'Call CambiaTamañoCombo(cboModalidad, 300) 'Comento JOEP20181218 CP
    Exit Sub

ERRORCargaControles:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CargaDatos(ActXCodCta.NroCuenta)
End If
End Sub


Private Sub cboAnalista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtfechaAsig.SetFocus
    End If
End Sub

Private Sub cboCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRelaciones.SetFocus
    End If
End Sub

Private Sub cboFuentes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscarAcreedor.SetFocus
    End If
End Sub

'JOEP20181218 CP
Private Sub cboModalidad_Click()
If Trim(Right(cboModalidad.Text, 9)) = "13" Then
    frModOtrs.Visible = True
    Me.Width = 7965
    Me.Height = 8350
    frFinalidad.top = 2480
    fraCreditos.Height = 3700
    fracontrol.top = 7300
    If txtModOtrs.Enabled = True Then
        txtModOtrs.SetFocus
    End If
Else
    txtModOtrs.Text = ""
    frModOtrs.Visible = False
    Me.Width = 7965
    Me.Height = 7800
    frFinalidad.top = 2000
    fraCreditos.Height = 3200
    fracontrol.top = 6800
End If
End Sub
'JOEP20181218 CP

Private Sub cboModalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboAnalista.SetFocus
    End If
End Sub

Private Sub cboMoneda_Click()
    If cboMoneda.Text <> "" Then
        If CInt(Right(cboMoneda.Text, 2)) = gMonedaNacional Then
            txtMontoSol.BackColor = vbWhite
        Else
            txtMontoSol.BackColor = RGB(200, 255, 200)
        End If
    End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoSol.SetFocus
    End If
End Sub

Private Sub cboTipoCF_Click()
Dim loCred As COMDCredito.DCOMCredito
Dim nValor As Integer
Dim nIndice As Integer
    Set loCred = New COMDCredito.DCOMCredito
    If Not oRelPersCred Is Nothing And Trim(cboTipoCF.Text) <> "" Then
        nValor = loCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, CInt(Mid(Right(cboTipoCF.Text, 3), 1, 3)))
        nIndice = IndiceListaCombo(cboTipoCF, Trim(str(nValor)))
        'If nIndice <> -1 Then
        '    LblCondProd.Caption = " " & cmbCondicion.List(nIndice)
        'End If
    End If
    Set loCred = Nothing
End Sub

Private Sub cboTipoCF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub
'WIOR 20120328-INICIO
Private Sub chkConsorcio_Click()
    If Me.chkConsorcio.value = 1 Then
        Me.txtConsorcio.Visible = True
        Me.cmdBuscarAvalado.Visible = False
        Me.lblCodAvalado.Visible = False
        Me.lblNomAvalado.Visible = False
        Me.lblCodAvalado.Caption = ""
        Me.lblNomAvalado.Caption = ""
        bConsorcio = True
      
    Else
        Me.txtConsorcio.Visible = False
        Me.txtConsorcio.Text = ""
        Me.cmdBuscarAvalado.Visible = True
        Me.lblCodAvalado.Visible = True
        Me.lblNomAvalado.Visible = True
        bConsorcio = False
    End If
End Sub
'WIOR - FIN
Private Sub ChkRenovacion_Click()
    If Me.ChkRenovacion.value = 1 Then
         Me.AXCodCta.Visible = True
         Me.cmdBuscar.Visible = True
         Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
         Me.AXCodCta.Age = ""
         Me.AXCodCta.Cuenta = ""
         Me.AXCodCta.Enabled = True
         Me.cmdBuscar.Enabled = True
    Else
        Me.AXCodCta.Age = ""
        Me.AXCodCta.Cuenta = ""
        Me.AXCodCta.Prod = ""
        Me.AXCodCta.CMAC = ""
        Me.AXCodCta.Enabled = False
        Me.cmdBuscar.Enabled = False
        Me.AXCodCta.Visible = False
        Me.cmdBuscar.Visible = False
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDCartaFianza.DCOMCartaFianza
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

If Me.ListaRelacion.ListItems.count > 0 Then
    lsPersNombre = Me.ListaRelacion.ListItems.Item(1).Text
    lsPersCod = Me.ListaRelacion.ListItems.Item(1).SubItems(2)
Else
    MsgBox "No existe Titular para realizar una Renovación", vbInformation, "Aviso"
    Exit Sub
End If

'Selecciona Estados

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDCartaFianza.DCOMCartaFianza
        Set lrContratos = loPersContrato.ObtieneCartaFianzaPersona(lsPersCod)
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        Me.AXCodCta.Enabled = True
        AXCodCta.SetFocusCuenta
    Else
        ChkRenovacion.value = 0
        AXCodCta.CMAC = ""
        AXCodCta.Prod = ""
        AXCodCta.Enabled = False
        cmdBuscar.Enabled = False
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdBuscarAcreedor_Click()
    Set oPersBuscada = New COMDPersona.UCOMPersona
    Set oPersBuscada = frmBuscaPersona.Inicio
    If oPersBuscada Is Nothing Then Exit Sub
    lblCodAcreedor.Caption = oPersBuscada.sPersCod
    lblNomAcreedor.Caption = oPersBuscada.sPersNombre
    'MAVM 20100612 BAS II
'    If cboTipoCF.Enabled = True Then
'        cboTipoCF.SetFocus
'    End If
    cboMoneda.SetFocus
    Set oPersBuscada = Nothing
End Sub

Private Sub cmdBuscarAvalado_Click()
  Set oPersBuscada = New COMDPersona.UCOMPersona
    Set oPersBuscada = frmBuscaPersona.Inicio
    If oPersBuscada Is Nothing Then Exit Sub
    lblCodAvalado.Caption = oPersBuscada.sPersCod
    lblNomAvalado.Caption = oPersBuscada.sPersNombre
  '  cboMoneda.SetFocus
    Set oPersBuscada = Nothing
End Sub

Private Sub cmdCancela_Click()
    Call LimpiaControles(Me, True)
    Call HabilitaActualizacion(False)
    ChkRenovacion.value = 0
    cboTipoCF.ListIndex = 0
End Sub

Private Sub cmdEditar_Click()
    If Len(Me.ActXCodCta.NroCuenta) < 18 Then Exit Sub
    cmdEjecutar = 2
    HabilitaActualizacion True
    cboMoneda.Enabled = False
    fraLineaCred.Enabled = False
    cboTipoCF.Enabled = False
    cboCondicion.Enabled = False 'JOEP20181221 CP
End Sub

Private Sub cmdExaminar_Click()
Dim sCta As String
    Call LimpiaPantalla
    'MAVM 20100604 Se agrego la var: gColCFTpoProducto BAS II***
    sCta = frmCFPersEstado.Inicio(Array(gColocEstSolic), "Solicitudes de Carta Fianza", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto))
    'MAVM ***
    If Len(Trim(sCta)) > 0 Then
        ActXCodCta.NroCuenta = sCta
        Call CargaDatos(sCta)
    End If
End Sub

Private Sub cmdFuentes_Click()
    Dim oPersona As COMDPersona.DCOMPersona
   
    If Not ExisteTitular Then
        MsgBox "Debe Ingresar el Titular del Credito", vbInformation, "Aviso"
        cmdRelaciones.SetFocus
        Exit Sub
    End If
    Set oPersona = New COMDPersona.DCOMPersona
    Call frmPersona.Inicio(oRelPersCred.TitularPersCod, PersonaActualiza)
    Call CargaFuentesIngreso(oRelPersCred.TitularPersCod)
    oPersona.PersCodigo = oRelPersCred.TitularPersCod
    Set oPersona = Nothing
    
End Sub

Private Sub cmdGarantias_Click()
    Dim frm As New frmGarantia 'ALPA20160330
    'If Len(AXCodCta.NroCuenta) = 18 Then
    If Len(ActXCodCta.NroCuenta) = 18 Then
        'EJVG20150712 ***
        'If gsProyectoActual = "H" Then
        '    frmPersGarantiasHC.Show 1
        'Else
        '    'frmPersGarantias.Show 1'WIOR 20140826 COMENTO
        '    frmPersGarantias.Inicio RegistroGarantia, , True 'WIOR 20140826 COMENTO
        'End If
        If MsgBox("Seleccione [SI] para Registrar Nuevas Garantías" & Chr(13) & "Seleccione [NO] para Editar Garantías", vbInformation + vbYesNo, "Aviso") = vbYes Then
            frm.Registrar
        Else
            frm.Editar
        End If
        Set frm = Nothing
        'END EJVG *******
    Else
       MsgBox "Ingrese los datos para la Carta Fianza", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim oCF As COMNCartaFianza.NCOMCartaFianza
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oPersona As COMDPersona.DCOMPersona
Dim psNuevaCta As String
Dim sMovAct As String
Dim sCtaCodA As String
Dim bRenovacion As String
    If Not ValidaDatos Then
        Exit Sub
    End If
    
    If ChkRenovacion.value = 1 Then
        sCtaCodA = AXCodCta.NroCuenta
        bRenovacion = True
    Else
        sCtaCodA = ""
        bRenovacion = False
    End If
    
    If MsgBox("Se va a Grabar los Datos, Desea Continuar?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbNo Then
        Exit Sub
    End If

    Call HabilitaActualizacion(False)

    Set oPersona = New COMDPersona.DCOMPersona
    If cmdEjecutar = 1 Then
        Set oGen = New COMDConstSistema.DCOMGeneral
            psNuevaCta = gsCodCMAC & oGen.GeneraNuevaCuenta(gsCodAge, CInt(Right(cboTipoCF.Text, 3)), CInt(Right(cboMoneda.Text, 2)))
        Set oGen = Nothing
        sMovAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        oPersona.RecuperaFtesdeIngreso (oRelPersCred.TitularPersCod)
        'oPersona.RecuperaFtesIngresoDependiente (cboFuentes.ListIndex) 'LUCV20160919
        'oPersona.RecuperaFtesIngresoIndependiente (cboFuentes.ListIndex) 'LUCV20160919
        Set oCF = New COMNCartaFianza.NCOMCartaFianza
        
        'MAVM 20100604 Se agregaron las var: gsCodAge, CInt(Right(cboTipoCF.Text, 3)) BAS II
        Call oCF.nCFRegistraSolicitud(psNuevaCta, gsCodAge, CInt(Right(cboTipoCF.Text, 2)), CInt(Right(cboMoneda.Text, 2)), CInt(Right(cboCondicion.Text, 2)), _
                                    CInt(Right(cboModalidad.Text, 2)), Trim(TxtFinalidad.Text), gdFecSis, oRelPersCred.ObtenerMatrizRelaciones, _
                                    Trim(Right(cboAnalista.Text, 20)), CDate(Me.txtfechaAsig.Text), CDate(txtFechaVencimiento.Text), CDbl(Me.txtMontoSol.Text), sMovAct, _
                                    -1, gdFecSis, _
                                    lblCodAcreedor.Caption, bRenovacion, sCtaCodA, gsCodAge, CInt(Right(cboTipoCF.Text, 3)), , lblCodAvalado.Caption, bConsorcio, Trim(Me.txtConsorcio.Text), Trim(txtModOtrs.Text)) 'WIOR 20120328
        'JOEP20181219 CP Trim(txtModOtrs.Text)
        '***
        'LUCV20160919, psNumFuente: oPersona.ObtenerFteIngFecEval(cboFuentes.ListIndex) = 1
        'LUCV20160919, pdPersEval:  oPersona.ObtenerFteIngFecEval(cboFuentes.ListIndex, IIf(oPersona.ObtenerFteIngIngresoTipo(cboFuentes.ListIndex) = gPersFteIngresoTipoDependiente, _
                                    oPersona.ObtenerFteIngIngresoNumeroFteDep(cboFuentes.ListIndex) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(cboFuentes.ListIndex) - 1), _
                                    oPersona.ObtenerFteIngIngresoTipo(cboFuentes.ListIndex)) -> Agregó: gdFecSis
        '*** PEAC 20090126
            objPista.InsertarPista gsOpeCod, sMovAct, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Registrar carta fianza", psNuevaCta, gCodigoCuenta
                        
        Set oCF = Nothing
        Call CargaDatos(psNuevaCta) 'WIOR 20120328
    Else
        psNuevaCta = ActXCodCta.NroCuenta
        sMovAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        oPersona.RecuperaFtesdeIngreso (oRelPersCred.TitularPersCod)
        'oPersona.RecuperaFtesIngresoDependiente (cboFuentes.ListIndex) 'LUCV20160919, Comentó según ERS004-2016
        'oPersona.RecuperaFtesIngresoIndependiente (cboFuentes.ListIndex)'LUCV20160919, Comentó según ERS004-2016
        Set oCF = New COMNCartaFianza.NCOMCartaFianza
        Call oCF.nCFActualizaSolicitud(psNuevaCta, gsCodAge, CInt(Right(cboTipoCF.Text, 2)), CInt(Right(cboMoneda.Text, 2)), _
                                      CInt(Right(cboCondicion.Text, 2)), CInt(Right(cboModalidad.Text, 2)), Trim(TxtFinalidad.Text), gdFecSis, _
                                      oRelPersCred.ObtenerMatrizRelaciones, Trim(Right(cboAnalista.Text, 20)), CDate(Me.txtfechaAsig.Text), _
                                      CDate(txtFechaVencimiento.Text), CDbl(Me.txtMontoSol.Text), sMovAct, _
                                      -1, gdFecSis, _
                                      lblCodAcreedor.Caption, bRenovacion, sCtaCodA, gsCodAge, CInt(Right(cboTipoCF.Text, 3)), , _
                                      lblCodAvalado.Caption, bConsorcio, Trim(Me.txtConsorcio.Text), Trim(txtModOtrs.Text)) 'WIOR 20120328
        
        'LUCV20160919, psNumFuente: oPersona.ObtenerFteIngcNumFuente(cboFuentes.ListIndex) = -1
        'LUCV20160919, pdPersEval:  oPersona.ObtenerFteIngFecEval(cboFuentes.ListIndex, IIf(oPersona.ObtenerFteIngIngresoTipo(cboFuentes.ListIndex) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(cboFuentes.ListIndex) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(cboFuentes.ListIndex) - 1), oPersona.ObtenerFteIngIngresoTipo(cboFuentes.ListIndex))
        
            ''*** PEAC 20090126
            objPista.InsertarPista gsOpeCod, sMovAct, gsCodPersUser, GetMaquinaUsuario, gModificar, "Modificar carta fianza", psNuevaCta, gCodigoCuenta
        
        Set oCF = Nothing
         Call CargaDatos(psNuevaCta) 'WIOR 20120328
    End If

    Set oPersona = Nothing
    ActXCodCta.NroCuenta = psNuevaCta
    ActXCodCta.Enabled = False
    
End Sub

Private Sub cmdGravar_Click()
  'If Len(AXCodCta.NroCuenta) = 18 Then
  If Len(ActXCodCta.NroCuenta) = 18 Then
    'frmCredGarantCred.Inicio PorSolicitud, ActXCodCta.NroCuenta, 1
    frmGarantiaCobertura.Inicio InicioGravamenxSolicitud, CartaFianza, ActXCodCta.NroCuenta 'EJVG20150712
    'JOEP20181220 CP trim(right(cboCondicion.Text,10))
  Else
    MsgBox "Ingrese los datos para la Carta Fianza", vbInformation, "Aviso"
  End If
End Sub

Private Sub cmdImprimir_Click()
Dim loImp As COMNCartaFianza.NCOMCartaFianzaReporte
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String
Dim lsmensaje As String
On Error GoTo ErrorImpresion
If Len(Me.ActXCodCta.NroCuenta) = 18 Then
    'frmImpresora.Show 1
    'If lbCancela = False Then
        Set loImp = New COMNCartaFianza.NCOMCartaFianzaReporte
            loImp.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsCadImprimir = loImp.nRepoDuplicado(Me.ActXCodCta.NroCuenta, 2, lsmensaje, gImpresora)
            If lsmensaje <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
        Set loImp = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.Show lsCadImprimir, "Carta Fianza - Solicitud ", True, , gImpresora
    'End If
End If
Exit Sub
ErrorImpresion:
    MsgBox "Error Nº [" & str(Err.Number) & "] " & Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdNuevo_Click()
    cmdEjecutar = 1
    Call LimpiaPantalla
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    Call HabilitaActualizacion(True)
    Call cmdRelaciones_Click
    cboTipoCF.ListIndex = 0
    cboCondicion.Enabled = False 'JOEP20181221 CP
End Sub

Private Sub cmdRelaciones_Click()
    Call frmCredRelaCta.Inicio(oRelPersCred, InicioSolicitud, , cmdEjecutar)
    Call ActualizarListaPersRelacCred
    'JUEZ 20130717 *****************************
    If ListaRelacion.ListItems.count = 0 Then
        cmdCancela_Click
        cmdNuevo.SetFocus
        Exit Sub
    End If
    'END JUEZ **********************************
    lblCodigo.Caption = oRelPersCred.TitularPersCod
    lblNombre.Caption = oRelPersCred.TitularNombre
    Call CargaFuentesIngreso(lblCodigo.Caption)
    Call DefineCondicionCF
    'cboFuentes.SetFocus 'LUCV20160919, Comentó. Según: ERS004-2016
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
 Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    'JOEP20181219 CP
    frModOtrs.Visible = False
    Me.Width = 7965
    Me.Height = 7800
    frFinalidad.top = 2000
    fraCreditos.Height = 3200
    fracontrol.top = 6800
    'JOEP20181219 CP
    
    Call CargaControles
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    Call HabilitaActualizacion(False)
     Me.txtConsorcio.Visible = False 'WIOR 20120330
     cboCondicion.Enabled = False 'JOEP20181219 CP
    txtMontoSol.Text = "0.00"
    txtfechaAsig.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtFechaVencimiento.Text = Format(gdFecSis, "dd/mm/yyyy")
    cmdEjecutar = -1
   bConsorcio = False 'WIOR 20120328
   Set objPista = New COMManejador.Pista
   gsOpeCod = gCredRegistrarCartaFianza
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRelPersCred = Nothing
    Set objPista = Nothing
    
     '***MARG ERS046-2016***AGREGADO 20161109***
    gsOpeCod = ""
    '***END MARG*******************************
End Sub


Private Sub txtfechaAsig_GotFocus()
    fEnfoque txtfechaAsig
End Sub
Private Sub txtfechaAsig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CP_CargaDatos 'JOEP20181218 CP
        TxtPeriodo.SetFocus
    End If
End Sub

Private Sub txtFechaVencimiento_GotFocus()
    fEnfoque txtFechaVencimiento
End Sub

Private Sub txtFechaVencimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtFinalidad.SetFocus
    End If
End Sub

Private Sub txtFechaVencimiento_LostFocus()
Dim sCad As String
    sCad = ValidaFecha(txtFechaVencimiento.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        Exit Sub
    End If
    If CDate(Format(txtFechaVencimiento.Text, "dd/mm/yyyy")) < CDate(Format(gdFecSis, "dd/mm/yyyy")) Then
        MsgBox "Fecha de Vencimiento no puede ser anterior a la fecha actual", vbInformation, "Aviso"
        txtFechaVencimiento.SetFocus
        Exit Sub
    Else
        If CDate(Format(txtFechaVencimiento.Text, "dd/mm/yyyy")) < CDate(Format(Me.txtfechaAsig, "dd/mm/yyyy")) Then
            MsgBox "Fecha de Vencimiento no puede ser menor que Fecha de Asignación", vbInformation, "Aviso"
            txtFechaVencimiento.SetFocus
            Exit Sub
        Else
            cmdGrabar.Enabled = True
            TxtFinalidad.SetFocus
        End If
    End If
    
End Sub

Private Sub TxtFinalidad_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtMontoSol_GotFocus()
    fEnfoque txtMontoSol
End Sub

Private Sub txtMontoSol_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMontoSol, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        cboModalidad.SetFocus
    End If
End Sub

Private Sub txtMontoSol_LostFocus()
    If Trim(txtMontoSol.Text) = "" Then
        txtMontoSol.Text = "0.00"
    End If
    txtMontoSol.Text = Format(txtMontoSol.Text, "#0.00")
End Sub

Private Sub DefineCondicionCF()
Dim loCred As COMDCredito.DCOMCredito
Dim nValor As Integer
Set loCred = New COMDCredito.DCOMCredito
    nValor = loCred.DefineCondicionCredito(oRelPersCred.TitularPersCod, gColCFTpoProducto)
Set loCred = Nothing
    cboCondicion.ListIndex = IndiceListaCombo(cboCondicion, Trim(str(nValor)))

End Sub

Private Function ExisteTitular() As Boolean
    ExisteTitular = False
    oRelPersCred.IniciarMatriz
    Do While Not oRelPersCred.EOF
        If oRelPersCred.ObtenerValorRelac = gColRelPersTitular Then
            ExisteTitular = True
            Exit Do
        End If
        oRelPersCred.siguiente
    Loop
End Function
    
Private Sub TxtPeriodo_Change()
    If IsNumeric(TxtPeriodo) Then
        If (txtfechaAsig.Text = "" Or txtfechaAsig.Text = "__/__/____") Then Exit Sub 'JOEP20181218 CP
        txtFechaVencimiento.Text = CDate(txtfechaAsig.Text) + CCur(TxtPeriodo.Text)
    End If
End Sub

Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CP_ValidaMsg(2) Then Exit Sub 'JOEP20181218 CP
        If IsNumeric(TxtPeriodo) Then
            If Not CP_ValidaMsg(1) Then Exit Sub 'JOEP20181218 CP
            txtFechaVencimiento.Text = CDate(txtfechaAsig.Text) + CCur(TxtPeriodo.Text)
            TxtFinalidad.SetFocus
        Else
            MsgBox "El periodo tiene que ser numerico", vbInformation, "Aviso"
            TxtPeriodo.Text = ""
        End If
    End If
End Sub

'JOEP20181218 CP
Private Sub CP_CargaDatos()
Dim oDCred As COMDCredito.DCOMCredito
Dim rsDefaut As ADODB.Recordset
Set oDCred = New COMDCredito.DCOMCredito

Set rsDefaut = oDCred.CatalogoProDefaut(514, 7000)

If Not (rsDefaut.BOF And rsDefaut.EOF) Then
    'TxtPeriodo.Text = rsDefaut!MinPlazo
    nPeriodoMin = rsDefaut!MinPlazo
    nPeriodoMax = rsDefaut!MaxPlazo
End If

End Sub

Private Function CP_ValidaMsg(ByVal nTpOp As Integer) As Boolean
CP_ValidaMsg = True
Select Case nTpOp
    Case 1
        If nPeriodoMin <> -1 Then
            If CInt(TxtPeriodo.Text) < nPeriodoMin Then
                MsgBox "El Periodo mínimo es " & nPeriodoMin & " días", vbInformation, "Aviso"
                TxtPeriodo.Text = nPeriodoMin
                CP_ValidaMsg = False
                Exit Function
            End If
            If CInt(TxtPeriodo.Text) > nPeriodoMax Then
                MsgBox "El Periodo máximo es " & nPeriodoMax & " días", vbInformation, "Aviso"
                TxtPeriodo.Text = nPeriodoMax
                CP_ValidaMsg = False
                Exit Function
            End If
        End If
    Case 2
        If (txtfechaAsig.Text = "" Or txtfechaAsig.Text = "__/__/____") Then
            MsgBox "Ingrese la Fecha de Asignacion", vbInformation, "Aviso"
            txtfechaAsig.SetFocus
            TxtPeriodo.Text = 0
            CP_ValidaMsg = False
            Exit Function
        End If
    Case 3
        If frModOtrs.Visible = True And txtModOtrs.Text = "" Then
            MsgBox "Registre Otras Modalidades", vbInformation, "Aviso"
            txtModOtrs.SetFocus
            CP_ValidaMsg = False
            Exit Function
        End If
End Select
End Function

Private Sub CP_CargaCombox(ByVal nParCod As Long)
Dim objCatalogoLlenaCombox As COMDCredito.DCOMCredito
Dim rsCatalogoCombox As ADODB.Recordset
Set objCatalogoLlenaCombox = New COMDCredito.DCOMCredito
Set rsCatalogoCombox = objCatalogoLlenaCombox.getCatalogoCombo("514", nParCod)

If Not (rsCatalogoCombox.BOF And rsCatalogoCombox.EOF) Then
    If nParCod = 5000 Then
        Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cboMoneda)
    ElseIf nParCod = 49000 Then
        Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cboModalidad)
        Call CambiaTamañoCombo(cboModalidad, 300)
    End If
End If

End Sub
'JOEP20181218 CP

