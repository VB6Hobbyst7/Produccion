VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogBienMnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Bienes"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   Icon            =   "frmLogBienMnt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   1050
   End
   Begin Sicmact.FlexEdit feBien 
      Height          =   2205
      Left            =   45
      TabIndex        =   4
      Top             =   1400
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   3889
      Cols0           =   26
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   $"frmLogBienMnt.frx":030A
      EncabezadosAnchos=   "350-0-1700-2200-1500-1400-1300-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-C-L-L-L-C-L-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      CantEntero      =   9
      TextArray0      =   "#"
      SelectionMode   =   1
      TipoBusqueda    =   0
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin TabDlg.SSTab TabBien 
      Height          =   3765
      Left            =   45
      TabIndex        =   30
      Top             =   3660
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   6641
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Activo Fijo"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdAFActualizar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAFCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   " &Bien No Depreciable"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "cmdBNDCancelar"
      Tab(1).Control(3)=   "cmdBNDActualizar"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Registro"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   8505
         Begin VB.TextBox txtAFInventarioCod 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtAFNombre 
            Height          =   285
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   7
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtAFAreaAgeNombre 
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1310
            Width           =   1740
         End
         Begin VB.TextBox txtAFPersonaNombre 
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1310
            Width           =   2055
         End
         Begin VB.TextBox txtAFDepreContAnio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            TabIndex        =   9
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAFDepreContMes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            TabIndex        =   10
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAFDepreTribMes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txtAFFechaConformidad 
            Height          =   315
            Left            =   7150
            TabIndex        =   6
            Top             =   240
            Width           =   1280
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   76873729
            CurrentDate     =   41414
         End
         Begin Sicmact.TxtBuscar txtAFPersonaCod 
            Height          =   255
            Left            =   4920
            TabIndex        =   12
            Top             =   1305
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            TipoBusqueda    =   3
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin Sicmact.TxtBuscar txtAFAreaAgeCod 
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   1310
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   450
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            EnabledText     =   0   'False
         End
         Begin VB.Label Label3 
            Caption         =   "Cód. Inventario:"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "F. Conformidad:"
            Height          =   255
            Left            =   4200
            TabIndex        =   58
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Área/Agencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Persona:"
            Height          =   255
            Left            =   4200
            TabIndex        =   55
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Tiempo Deprec. Cont."
            Height          =   255
            Left            =   4200
            TabIndex        =   54
            Top             =   975
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Año:"
            Height          =   255
            Left            =   5985
            TabIndex        =   53
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   7305
            TabIndex        =   52
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Tiempo Deprec. Trib."
            Height          =   255
            Left            =   4200
            TabIndex        =   51
            Top             =   615
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   7320
            TabIndex        =   50
            Top             =   615
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Caracteristicas del Bien"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   43
         Top             =   2200
         Width           =   8505
         Begin VB.TextBox txtAFMarca 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   13
            Top             =   235
            Width           =   2775
         End
         Begin VB.TextBox txtAFSerie 
            Height          =   285
            Left            =   5280
            MaxLength       =   50
            TabIndex        =   14
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtAFModelo 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   15
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label17 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label16 
            Caption         =   "Serie/caract:"
            Height          =   255
            Left            =   4200
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Modelo:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdAFCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   7560
         TabIndex        =   17
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdAFActualizar 
         Caption         =   "&Actualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6480
         TabIndex        =   16
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdBNDActualizar 
         Caption         =   "&Actualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -68520
         TabIndex        =   26
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdBNDCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -67440
         TabIndex        =   27
         Top             =   3330
         Width           =   1050
      End
      Begin VB.Frame Frame4 
         Caption         =   "Caracteristicas del Bien"
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   39
         Top             =   2205
         Width           =   8505
         Begin VB.TextBox txtBNDModelo 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   25
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtBNDSerie 
            Height          =   285
            Left            =   5280
            MaxLength       =   50
            TabIndex        =   24
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtBNDMarca 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   23
            Top             =   235
            Width           =   2775
         End
         Begin VB.Label Label19 
            Caption         =   "Modelo:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Serie/caract:"
            Height          =   255
            Left            =   4200
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos de Registro"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   8505
         Begin VB.TextBox txtBNDPersonaNombre 
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1310
            Width           =   2055
         End
         Begin VB.TextBox txtBNDAreaAgeNombre 
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1310
            Width           =   1740
         End
         Begin VB.TextBox txtBNDNombre 
            Height          =   285
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   20
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtBNDInventarioCod 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker txtBNDFechaConformidad 
            Height          =   315
            Left            =   7150
            TabIndex        =   19
            Top             =   240
            Width           =   1280
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   76873729
            CurrentDate     =   41414
         End
         Begin Sicmact.TxtBuscar txtBNDPersonaCod 
            Height          =   255
            Left            =   4920
            TabIndex        =   22
            Top             =   1305
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            TipoBusqueda    =   3
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin Sicmact.TxtBuscar txtBNDAreaAgeCod 
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   1310
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   450
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            EnabledText     =   0   'False
         End
         Begin VB.Label Label22 
            Caption         =   "Persona:"
            Height          =   255
            Left            =   4200
            TabIndex        =   38
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label23 
            Caption         =   "Área/Agencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label30 
            Caption         =   "F. Conformidad:"
            Height          =   255
            Left            =   4200
            TabIndex        =   35
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Cód. Inventario:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   270
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   60
         Top             =   495
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Producto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nro. Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nro. Cta Antigua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Participación"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "SaldoCont"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SaldoDisp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Motivo de Bloque"
            Object.Width           =   7231
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   65
         Top             =   3465
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73185
         TabIndex        =   64
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   63
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label lblSolesAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70815
         TabIndex        =   62
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67680
         TabIndex        =   61
         Top             =   3375
         Width           =   2145
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda Series"
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
      Height          =   1300
      Left            =   60
      TabIndex        =   28
      Top             =   40
      Width           =   7425
      Begin VB.TextBox txtSerieNombre 
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   880
         Width           =   3060
      End
      Begin VB.TextBox txtCategoriaNombre 
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   562
         Width           =   3060
      End
      Begin VB.TextBox txtAreaAgeNombre 
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   3060
      End
      Begin Sicmact.TxtBuscar txtAreaAgeCod 
         Height          =   255
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
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
      End
      Begin Sicmact.TxtBuscar txtCategoriaCod 
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   562
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
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
      End
      Begin Sicmact.TxtBuscar txtSerieCod 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   880
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   450
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
      End
      Begin VB.Label Label24 
         Caption         =   "Serie:"
         Height          =   255
         Left            =   360
         TabIndex        =   69
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Categoría:"
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Área/Agencia:"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLogBienMnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmLogBienMnt
'** Descripción : Mantenimiento de Bienes Activados creado segun ERS059-2013
'** Creación : EJVG, 20130615 09:00:00 AM
'***************************************************************************
Option Explicit
Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double
Dim fnMovNro As Long
Dim fnCategoriaBien As Integer
Dim fnBANCod As Integer
Dim fbActivaAF As Boolean, fnActivaAF_Id As Long
Dim fbActivaAF_AC As Boolean, fnActivaAF_AC_Id As Long, fnActivaAF_AC_Item As Long
Dim fbActivaBND As Boolean, fnActivaBND_Id As Long
Dim fbActivaBND_AC As Boolean, fnActivaBND_AC_Id As Long, fnActivaBND_AC_Item As Long
Dim fbDepreciado As Boolean
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub Form_Load()
    CentraForm Me
    fnFormTamanioIni = 4125
    fnFormTamanioActiva = 7965
    Height = fnFormTamanioIni
    Call CargaControles
End Sub
Private Sub CargaControles()
    Dim obj As New DActualizaDatosArea
    Dim oBien As New DBien
    txtAreaAgeCod.rs = obj.GetAgenciasAreas()
    txtCategoriaCod.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto("", "")
    txtAFAreaAgeCod.rs = obj.GetAgenciasAreas()
    txtBNDAreaAgeCod.rs = obj.GetAgenciasAreas()
    Set obj = Nothing
    Set oBien = Nothing
End Sub
Private Sub txtAreaAgeCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    txtAreaAgeNombre.Text = ""
    txtCategoriaCod.Text = ""
    txtCategoriaNombre.Text = ""
    If txtAreaAgeCod.Text <> "" Then
        txtAreaAgeNombre.Text = txtAreaAgeCod.psDescripcion
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
        txtCategoriaCod.rs = oBien.RecuperaCategoriasBienPaObjeto(False, lsAreaAgeCod)
    Else
        txtCategoriaCod.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    End If
    txtCategoriaCod_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
Private Sub txtAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCategoriaCod.SetFocus
    End If
End Sub
Private Sub txtAreaAgeCod_LostFocus()
    Dim oBien As New DBien
    If txtAreaAgeCod.Text = "" Then
        txtAreaAgeNombre.Text = ""
    End If
    Set oBien = Nothing
End Sub
Private Sub txtCategoriaCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    txtCategoriaNombre.Text = ""
    If txtCategoriaCod.Text <> "" Then
        txtCategoriaNombre.Text = txtCategoriaCod.psDescripcion
    End If
    If txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    txtSerieCod.Text = ""
    txtSerieNombre.Text = ""
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto(lsAreaAgeCod, txtCategoriaCod.Text)
    txtSerieCod_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
Private Sub txtCategoriaCod_LostFocus()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    If txtCategoriaCod.Text = "" Then
        txtCategoriaNombre.Text = ""
    End If
    Set oBien = Nothing
End Sub
Private Sub txtSerieCod_EmiteDatos()
    txtSerieNombre.Text = ""
    If txtSerieCod.Text <> "" Then
        txtSerieNombre.Text = txtSerieCod.psDescripcion
    End If
    Call LimpiaFlex(feBien)
    Height = fnFormTamanioIni
End Sub
Private Sub txtSerieCod_LostFocus()
    If txtSerieCod.Text = "" Then
        txtSerieNombre.Text = ""
    End If
End Sub
Private Sub cmdBuscar_Click()
    Dim oBien As New DBien
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    Dim lsAreaAgeCod As String
    
    If Not ValidaBuscar Then Exit Sub
    Screen.MousePointer = 11
    
    If txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    
    Set rs = oBien.RecuperaBienxMantenimiento2(lsAreaAgeCod, txtCategoriaCod.Text, txtSerieCod.Text)
    Call LimpiaFlex(feBien)
    Height = fnFormTamanioIni
    Do While Not rs.EOF
        feBien.AdicionaFila
        fila = feBien.row
        feBien.TextMatrix(fila, 1) = rs!nMovNro
        feBien.TextMatrix(fila, 2) = rs!cInventarioCod
        feBien.TextMatrix(fila, 3) = UCase(rs!cNombre)
        feBien.TextMatrix(fila, 4) = UCase(rs!cMarca)
        feBien.TextMatrix(fila, 5) = rs!cSerie
        feBien.TextMatrix(fila, 6) = Format(rs!dActivacion, gsFormatoFechaView)
        feBien.TextMatrix(fila, 7) = rs!nDepreciaTributMes
        feBien.TextMatrix(fila, 8) = rs!nDepreciaContabMes
        feBien.TextMatrix(fila, 9) = rs!cAreaAgeCod
        feBien.TextMatrix(fila, 10) = rs!cPersCod
        feBien.TextMatrix(fila, 11) = UCase(rs!cPersNombre)
        feBien.TextMatrix(fila, 12) = UCase(rs!cModelo)
        feBien.TextMatrix(fila, 13) = rs!nActivadoAF
        feBien.TextMatrix(fila, 14) = rs!nActivadoAF_Id
        feBien.TextMatrix(fila, 15) = rs!nActivadoAF_AC
        feBien.TextMatrix(fila, 16) = rs!nActivadoAF_AC_Id
        feBien.TextMatrix(fila, 17) = rs!nActivadoAF_AC_Item
        feBien.TextMatrix(fila, 18) = rs!nActivadoBND
        feBien.TextMatrix(fila, 19) = rs!nActivadoBND_Id
        feBien.TextMatrix(fila, 20) = rs!nActivadoBND_AC
        feBien.TextMatrix(fila, 21) = rs!nActivadoBND_AC_Id
        feBien.TextMatrix(fila, 22) = rs!nActivadoBND_AC_Item
        feBien.TextMatrix(fila, 23) = rs!cCategoBien
        feBien.TextMatrix(fila, 24) = rs!ban
        rs.MoveNext
    Loop
    feBien.TopRow = 1
    feBien.row = 1
    Screen.MousePointer = 0
    If FlexVacio(feBien) Then
        MsgBox "No se encontraron resultados de la búsqueda realizada", vbInformation, "Aviso"
    End If
End Sub
Private Function ValidaBuscar() As Boolean
    ValidaBuscar = True
End Function
Private Sub feBien_OnRowChange(pnRow As Long, pnCol As Long)
    If pnRow > 0 Then
        Call feBien_Click
    End If
End Sub
Private Sub feBien_Click()
    Dim xBien As New DBien
    Dim pnRow As Long, pnCol As Long
    
    pnRow = feBien.row
    pnCol = feBien.col
    Screen.MousePointer = 11
    If Not FlexVacio(feBien) Then
        If pnRow > 0 Then
            fnMovNro = CLng(Trim(feBien.TextMatrix(pnRow, 1)))
            fnCategoriaBien = Trim(feBien.TextMatrix(pnRow, 23))
            fnBANCod = CInt(Trim(feBien.TextMatrix(pnRow, 24)))
            fbActivaAF = IIf(CInt(Trim(feBien.TextMatrix(pnRow, 13))) = 1, True, False)
            fnActivaAF_Id = CInt(Trim(feBien.TextMatrix(pnRow, 14)))
            fbActivaAF_AC = IIf(CInt(Trim(feBien.TextMatrix(pnRow, 15))) = 1, True, False)
            fnActivaAF_AC_Id = CInt(Trim(feBien.TextMatrix(pnRow, 16)))
            fnActivaAF_AC_Item = CInt(Trim(feBien.TextMatrix(pnRow, 17)))
            fbActivaBND = IIf(CInt(Trim(feBien.TextMatrix(pnRow, 18))) = 1, True, False)
            fnActivaBND_Id = CInt(Trim(feBien.TextMatrix(pnRow, 19)))
            fbActivaBND_AC = IIf(CInt(Trim(feBien.TextMatrix(pnRow, 20))) = 1, True, False)
            fnActivaBND_AC_Id = CInt(Trim(feBien.TextMatrix(pnRow, 21)))
            fnActivaBND_AC_Item = CInt(Trim(feBien.TextMatrix(pnRow, 22)))
            'fbDepreciado = oALmacen.BuscaSiSerieFueDepre(Trim(feBien.TextMatrix(pnRow, 2)))
            'fbDepreciado = xBien.TieneDepreciacion(Trim(feBien.TextMatrix(pnRow, 2))) 'Comentado by NAGL 20191222
            'Para indicar si se puede modificar el tiempo de depreciación en el Día
            fbDepreciado = xBien.ValidaDepreciacionDia(fnMovNro, gdFecSis) ' NAGL 20191222 Según RFC1910190001
            
            If fnCategoriaBien = 1 Then 'Activo Fijo
                TabBien.TabVisible(0) = True
                TabBien.TabVisible(1) = False
                TabBien.Tab = 0
                Call llenarDatosAF(Trim(feBien.TextMatrix(pnRow, 2)), Trim(feBien.TextMatrix(pnRow, 3)), CDate(Trim(feBien.TextMatrix(pnRow, 6))), _
                                    Trim(feBien.TextMatrix(pnRow, 7)), Trim(feBien.TextMatrix(pnRow, 8)), Trim(feBien.TextMatrix(pnRow, 9)), _
                                    Trim(feBien.TextMatrix(pnRow, 10)), Trim(feBien.TextMatrix(pnRow, 11)), Trim(feBien.TextMatrix(pnRow, 4)), _
                                    Trim(feBien.TextMatrix(pnRow, 5)), Trim(feBien.TextMatrix(pnRow, 12)))
                
                If fbDepreciado = True Then
                    txtAFFechaConformidad.Enabled = True
                    txtAFDepreContAnio.Locked = False
                    txtAFDepreContMes.Locked = False
                Else
                    txtAFFechaConformidad.Enabled = False
                    txtAFDepreContAnio.Locked = True
                    txtAFDepreContMes.Locked = True
                End If 'NAGL 20191222
                
            Else
                TabBien.TabVisible(0) = False
                TabBien.TabVisible(1) = True
                TabBien.Tab = 1
                Call llenarDatosBND(Trim(feBien.TextMatrix(pnRow, 2)), Trim(feBien.TextMatrix(pnRow, 3)), CDate(Trim(feBien.TextMatrix(pnRow, 6))), _
                                    Trim(feBien.TextMatrix(pnRow, 9)), Trim(feBien.TextMatrix(pnRow, 10)), Trim(feBien.TextMatrix(pnRow, 11)), _
                                    Trim(feBien.TextMatrix(pnRow, 4)), Trim(feBien.TextMatrix(pnRow, 5)), Trim(feBien.TextMatrix(pnRow, 12)))
            End If
            Height = fnFormTamanioActiva
        End If
    End If
    Set xBien = Nothing
    Screen.MousePointer = 0
End Sub
Private Sub feBien_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feBien.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
'*************** ACTIVO FIJO ********************
Private Sub txtAFInventarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFFechaConformidad.SetFocus
    End If
End Sub
Private Sub txtAFFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFNombre.SetFocus
    End If
End Sub
Private Sub txtAFNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtAFDepreTribMes.SetFocus
    End If
End Sub
Private Sub txtAFDepreTribMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtAFDepreContAnio.SetFocus
    End If
End Sub
Private Sub txtAFDepreContAnio_GotFocus()
    txtAFDepreContAnio.Text = Round(txtAFDepreContAnio.Text)
End Sub
Private Sub txtAFDepreContAnio_LostFocus()
    txtAFDepreContAnio.Text = Val(txtAFDepreContAnio.Text)
End Sub
Private Sub txtAFDepreContAnio_Change()
    txtAFDepreContMes.Text = Val(txtAFDepreContAnio.Text) * 12
End Sub
Private Sub txtAFDepreContAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtAFDepreContMes.SetFocus
    End If
End Sub
Private Sub txtAFDepreContMes_LostFocus()
    txtAFDepreContMes.Text = Val(txtAFDepreContMes.Text)
End Sub
Private Sub txtAFDepreContMes_Change()
    txtAFDepreContAnio.Text = Val(txtAFDepreContMes.Text) / 12
End Sub
Private Sub txtAFDepreContMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        If txtAFAreaAgeCod.Enabled Then
            txtAFAreaAgeCod.SetFocus
        Else
            txtAFMarca.SetFocus
        End If
    End If
End Sub
Private Sub txtAFAreaAgeCod_EmiteDatos()
    txtAFAreaAgeNombre.Text = txtAFAreaAgeCod.psDescripcion
    txtAFDepreTribMes.Text = RecuperaMesesDepreciaTributariamente(fnBANCod, Mid(txtAFAreaAgeCod.Text, 4, 2))
End Sub
Private Sub txtAFPersonaCod_EmiteDatos()
    txtAFPersonaNombre.Text = txtAFPersonaCod.psDescripcion
End Sub
Private Sub txtAFAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtAFPersonaCod.Enabled Then
            txtAFPersonaCod.SetFocus
        Else
            txtAFMarca.SetFocus
        End If
    End If
End Sub
Private Sub txtAFPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFMarca.SetFocus
    End If
End Sub
Private Sub txtAFMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtAFSerie.SetFocus
    End If
End Sub
Private Sub txtAFSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtAFModelo.SetFocus
    End If
End Sub
Private Sub txtAFModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmdAFActualizar.SetFocus
    End If
End Sub
Private Sub cmdAFActualizar_Click()
    Dim oBien As DBien
    Dim lsInventarioCod As String, lsNombre As String
    Dim ldFechaIng As Date
    Dim lnDepreTributMes As Integer, lnDepreContabMes As Integer
    Dim lsAreaCod As String, lsAgeCod As String
    Dim lsMarca As String, lsSerie As String, lsModelo As String
    Dim bTrans As Boolean
    
    On Error GoTo ErrAFActualizar
    If Not validarActualizarAF Then Exit Sub
   
    If MsgBox("¿Esta seguro de actualizar los datos del Activo Fijo?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    lsInventarioCod = txtAFInventarioCod.Text
    lsNombre = Trim(txtAFNombre.Text)
    ldFechaIng = CDate(txtAFFechaConformidad.value)
    lnDepreTributMes = CInt(txtAFDepreTribMes.Text)
    lnDepreContabMes = CInt(txtAFDepreContMes.Text)
    lsAreaCod = Left(txtAFAreaAgeCod.Text, 3)
    lsAgeCod = IIf(Mid(txtAFAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAFAreaAgeCod.Text, 4, 2))
    lsMarca = Trim(txtAFMarca.Text)
    lsSerie = Trim(txtAFSerie.Text)
    lsModelo = Trim(txtAFModelo.Text)
    
    Set oBien = New DBien
    
    oBien.dBeginTrans
    bTrans = True
    'Agencia ni Persona se actualizará, para que en transferencia se haga
    Call oBien.ActualizarAF(fnMovNro, , ldFechaIng, lsNombre, lnDepreTributMes, lnDepreContabMes, , , , lsMarca, lsSerie, lsModelo)
    Call oBien.ActualizarVidaUtilAF(fnMovNro, fnMovNro, lnDepreContabMes)
    oBien.dCommitTrans
    bTrans = False
        'ARLO 20160126 ***
        gsopecod = LogPistaEntradaSalidaBien
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", "Actualizo el Bien N° Serie : " & txtAFInventarioCod.Text & " Nombre : " & Trim(txtAFNombre.Text)
        Set objPista = Nothing
        '**************
    Set oBien = Nothing
    MsgBox "Se ha actualizado los datos del Activo Fijo con éxito", vbInformation, "Aviso"
    Call llenarDatosAF("", "", gdFecSis, 0, 0, "", "", "", "", "", "")
    Call cmdBuscar_Click
    Exit Sub
ErrAFActualizar:
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdAFCancelar_Click()
    Height = fnFormTamanioIni
End Sub
Private Sub llenarDatosAF(ByVal psInventarioCod As String, ByVal psNombre As String, ByVal pdFechaAct As Date, ByVal pnDepreTrib As Long, ByVal pnDepreCont As Long, _
                            ByVal psAreaAgeCod As String, ByVal psPersCod As String, ByVal psPersNombre As String, ByVal psMarca As String, _
                            ByVal psSerie As String, ByVal psModelo As String)
    txtAFInventarioCod.Text = psInventarioCod
    txtAFNombre.Text = psNombre
    txtAFFechaConformidad.value = Format(pdFechaAct, gsFormatoFechaView)
    txtAFDepreTribMes.Text = pnDepreTrib
    txtAFDepreContMes.Text = pnDepreCont
    txtAFAreaAgeCod.Text = psAreaAgeCod
    Call txtAFAreaAgeCod_EmiteDatos
    If txtAFAreaAgeCod.Text = "" And Mid(psAreaAgeCod, 4, 2) = "01" Then
        txtAFAreaAgeCod.Text = Left(psAreaAgeCod, 3)
        Call txtAFAreaAgeCod_EmiteDatos
    End If
    txtAFPersonaCod.Text = psPersCod
    txtAFPersonaNombre.Text = psPersNombre
    txtAFMarca.Text = psMarca
    txtAFSerie.Text = psSerie
    txtAFModelo.Text = psModelo
End Sub
Private Function validarActualizarAF() As Boolean
    Dim valFecha As String
    validarActualizarAF = True
    If Len(Trim(txtAFInventarioCod.Text)) = 0 Then
        validarActualizarAF = False
        MsgBox "El Bien a Actualizar no tiene Código de Inventario", vbInformation, "Aviso"
        txtAFInventarioCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFNombre.Text)) = 0 Then
        validarActualizarAF = False
        MsgBox "Ud. debe de ingresar el Nombre del Bien", vbInformation, "Aviso"
        txtAFNombre.SetFocus
        Exit Function
    End If
    valFecha = ValidaFecha(txtAFFechaConformidad.value)
    If valFecha <> "" Then
        validarActualizarAF = False
        MsgBox valFecha, vbInformation, "Aviso"
        'txtAFFechaConformidad.SetFocus
        Exit Function
    End If
    If Val(txtAFDepreTribMes.Text) = 0 Then
        validarActualizarAF = False
        MsgBox "Ud. debe ingresar el tiempo en que se deprecia Tributariamente el Bien", vbInformation, "Aviso"
        txtAFDepreTribMes.SetFocus
        Exit Function
    End If
    If Val(txtAFDepreContMes.Text) = 0 Then
        validarActualizarAF = False
        MsgBox "Ud. debe ingresar el tiempo en que se deprecia Contablemente el Bien", vbInformation, "Aviso"
        txtAFDepreContMes.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFMarca.Text)) = 0 Then
        validarActualizarAF = False
        MsgBox "Ud. debe ingresar la Marca del Bien", vbInformation, "Aviso"
        txtAFMarca.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFSerie.Text)) = 0 Then
        validarActualizarAF = False
        MsgBox "Ud. debe ingresar la Serie del Bien", vbInformation, "Aviso"
        txtAFSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFModelo.Text)) = 0 Then
        validarActualizarAF = False
        MsgBox "Ud. debe ingresar el Modelo del Bien", vbInformation, "Aviso"
        txtAFModelo.SetFocus
        Exit Function
    End If
End Function
'*************** BND ********************
Private Sub txtBNDInventarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDFechaConformidad.SetFocus
    End If
End Sub
Private Sub txtBNDFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDNombre.SetFocus
    End If
End Sub
Private Sub txtBNDNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If txtBNDAreaAgeCod.Enabled Then
            txtBNDAreaAgeCod.SetFocus
        Else
            txtBNDMarca.SetFocus
        End If
    End If
End Sub
Private Sub txtBNDAreaAgeCod_EmiteDatos()
    txtBNDAreaAgeNombre.Text = txtBNDAreaAgeCod.psDescripcion
End Sub
Private Sub txtBNDPersonaCod_EmiteDatos()
    txtBNDPersonaNombre.Text = txtBNDPersonaCod.psDescripcion
End Sub
Private Sub txtBNDAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtBNDPersonaCod.Enabled Then
            txtBNDPersonaCod.SetFocus
        Else
            txtBNDMarca.SetFocus
        End If
    End If
End Sub
Private Sub txtBNDPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDMarca.SetFocus
    End If
End Sub
Private Sub txtBNDMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtBNDSerie.SetFocus
    End If
End Sub
Private Sub txtBNDSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtBNDModelo.SetFocus
    End If
End Sub
Private Sub txtBNDModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmdBNDActualizar.SetFocus
    End If
End Sub
Private Sub cmdBNDActualizar_Click()
    Dim oBien As DBien
    Dim lsInventarioCod As String, lsNombre As String
    Dim ldFechaIng As Date
    Dim lsAreaCod As String, lsAgeCod As String
    Dim lsMarca As String, lsSerie As String, lsModelo As String
    Dim bTrans As Boolean
    
    On Error GoTo ErrAFActualizar
    If Not validarActualizarBND Then Exit Sub
    
    If MsgBox("¿Esta seguro de actualizar los datos del Activo Fijo?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    lsInventarioCod = txtBNDInventarioCod.Text
    lsNombre = Trim(txtBNDNombre.Text)
    ldFechaIng = CDate(txtBNDFechaConformidad.value)
    lsAreaCod = Left(txtBNDAreaAgeCod.Text, 3)
    lsAgeCod = IIf(Mid(txtBNDAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtBNDAreaAgeCod.Text, 4, 2))
    lsMarca = Trim(txtBNDMarca.Text)
    lsSerie = Trim(txtBNDSerie.Text)
    lsModelo = Trim(txtBNDModelo.Text)
    
    Set oBien = New DBien
    
    oBien.dBeginTrans
    bTrans = True
    'Área/Agencia ni Persona se actualizará, para que en transferencia se haga
    Call oBien.ActualizarAF(fnMovNro, , ldFechaIng, lsNombre, , , , , , lsMarca, lsSerie, lsModelo)
    oBien.dCommitTrans
    bTrans = False
    Set oBien = Nothing
    MsgBox "Se ha actualizado los datos del Bien No Depreciable con éxito", vbInformation, "Aviso"
    Call llenarDatosBND("", "", gdFecSis, "", "", "", "", "", "")
    Call cmdBuscar_Click
    Exit Sub
ErrAFActualizar:
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdBNDCancelar_Click()
    Height = fnFormTamanioIni
End Sub
Private Sub llenarDatosBND(ByVal psInventarioCod As String, ByVal psNombre As String, ByVal pdFechaAct As Date, ByVal psAreaAgeCod As String, ByVal psPersCod As String, ByVal psPersNombre As String, ByVal psMarca As String, ByVal psSerie As String, ByVal psModelo As String)
    txtBNDInventarioCod.Text = psInventarioCod
    txtBNDNombre.Text = psNombre
    txtBNDFechaConformidad.value = Format(pdFechaAct, gsFormatoFechaView)
    txtBNDAreaAgeCod.Text = psAreaAgeCod
    Call txtBNDAreaAgeCod_EmiteDatos
    If txtBNDAreaAgeCod.Text = "" And Mid(psAreaAgeCod, 4, 2) = "01" Then
        txtBNDAreaAgeCod.Text = Left(psAreaAgeCod, 3)
        Call txtBNDAreaAgeCod_EmiteDatos
    End If
    txtBNDPersonaCod.Text = psPersCod
    txtBNDPersonaNombre.Text = psPersNombre
    txtBNDMarca.Text = psMarca
    txtBNDSerie.Text = psSerie
    txtBNDModelo.Text = psModelo
End Sub
Private Function validarActualizarBND() As Boolean
    Dim valFecha As String
    validarActualizarBND = True
    If Len(Trim(txtBNDInventarioCod.Text)) = 0 Then
        validarActualizarBND = False
        MsgBox "El Bien a Actualizar no tiene Código de Inventario", vbInformation, "Aviso"
        txtBNDInventarioCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDNombre.Text)) = 0 Then
        validarActualizarBND = False
        MsgBox "Ud. debe de ingresar el Nombre del Bien", vbInformation, "Aviso"
        txtBNDNombre.SetFocus
        Exit Function
    End If
    valFecha = ValidaFecha(txtBNDFechaConformidad.value)
    If valFecha <> "" Then
        validarActualizarBND = False
        MsgBox valFecha, vbInformation, "Aviso"
        txtBNDFechaConformidad.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDMarca.Text)) = 0 Then
        validarActualizarBND = False
        MsgBox "Ud. debe ingresar la Marca del Bien", vbInformation, "Aviso"
        txtBNDMarca.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDSerie.Text)) = 0 Then
        validarActualizarBND = False
        MsgBox "Ud. debe ingresar la Serie del Bien", vbInformation, "Aviso"
        txtBNDSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDModelo.Text)) = 0 Then
        validarActualizarBND = False
        MsgBox "Ud. debe ingresar el Modelo del Bien", vbInformation, "Aviso"
        txtBNDModelo.SetFocus
        Exit Function
    End If
End Function
'*****************************************************
Private Function RecuperaMesesDepreciaTributariamente(ByVal pnTpoActivo As Integer, ByRef psAgeCod As String) As Integer
    Dim dLog As New DLogDeprecia
    Dim rs As New ADODB.Recordset
    If psAgeCod = "" Then psAgeCod = "01"
    Set rs = dLog.ObtienePorcentVidaUtlAF(pnTpoActivo)
    RecuperaMesesDepreciaTributariamente = 0
    Do While Not rs.EOF
        If psAgeCod = rs!cAgeCod Then
            RecuperaMesesDepreciaTributariamente = rs!nDepreMesT
            Exit Do
        End If
        rs.MoveNext
    Loop
    Set dLog = Nothing
    Set rs = Nothing
End Function
