VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmColRecRConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pagos de Creditos Judiciales"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10935
   Icon            =   "frmColRecRConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3780
      Left            =   105
      TabIndex        =   2
      Top             =   3690
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   6668
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Gastos"
      TabPicture(0)   =   "frmColRecRConsulta.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FE1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblGastosPagados"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblGastosAcumulados"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Amortizaciones"
      TabPicture(1)   =   "frmColRecRConsulta.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTotalPendiente"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "lblTotalPagado"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "FE2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Datos &Expediente"
      TabPicture(2)   =   "frmColRecRConsulta.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Datos &Generales"
      TabPicture(3)   =   "frmColRecRConsulta.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Negociacion"
      TabPicture(4)   =   "frmColRecRConsulta.frx":037A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fraNegociacion"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame fraNegociacion 
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
         Height          =   3030
         Left            =   135
         TabIndex        =   94
         Top             =   495
         Width           =   10125
         Begin VB.TextBox txtNegComenta 
            Height          =   585
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   106
            Top             =   2295
            Width           =   7185
         End
         Begin VB.TextBox txtNegEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   8685
            TabIndex        =   98
            Top             =   810
            Width           =   1290
         End
         Begin VB.TextBox txtNegMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   8685
            TabIndex        =   97
            Top             =   2085
            Width           =   1290
         End
         Begin VB.TextBox txtNegCuotas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   8685
            TabIndex        =   96
            Top             =   1665
            Width           =   1290
         End
         Begin VB.TextBox txtNegNro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   8685
            TabIndex        =   95
            Top             =   375
            Width           =   1290
         End
         Begin MSMask.MaskEdBox TxtNegVigencia 
            Height          =   285
            Left            =   8685
            TabIndex        =   99
            Top             =   1260
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ListView lvwCalendario 
            Height          =   2025
            Left            =   120
            TabIndex        =   100
            Top             =   225
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   3572
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Monto"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Monto Pagado"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Estado"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label Label42 
            Caption         =   "Monto Neg"
            Height          =   255
            Left            =   7695
            TabIndex        =   105
            Top             =   2130
            Width           =   825
         End
         Begin VB.Label Label41 
            Caption         =   "Vigencia"
            Height          =   225
            Left            =   7695
            TabIndex        =   104
            Top             =   1305
            Width           =   705
         End
         Begin VB.Label Label40 
            Caption         =   "Estado"
            Height          =   225
            Left            =   7695
            TabIndex        =   103
            Top             =   855
            Width           =   675
         End
         Begin VB.Label Label39 
            Caption         =   "Cuotas"
            Height          =   255
            Left            =   7695
            TabIndex        =   102
            Top             =   1710
            Width           =   675
         End
         Begin VB.Label Label37 
            Caption         =   "Nro Negoc"
            Height          =   225
            Left            =   7695
            TabIndex        =   101
            Top             =   450
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   -74700
         TabIndex        =   78
         Top             =   600
         Width           =   9930
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Int. Morat."
            Height          =   195
            Left            =   6840
            TabIndex        =   92
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Int. Comp."
            Height          =   195
            Left            =   6840
            TabIndex        =   91
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label lblGenerales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   8160
            TabIndex        =   90
            Top             =   735
            Width           =   1170
         End
         Begin VB.Label lblGenerales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   8160
            TabIndex        =   89
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Línea de Crédito"
            Height          =   195
            Left            =   180
            TabIndex        =   88
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label lblGenerales 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   87
            Top             =   330
            Width           =   3900
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Analista"
            Height          =   195
            Left            =   180
            TabIndex        =   86
            Top             =   708
            Width           =   555
         End
         Begin VB.Label lblGenerales 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   85
            Top             =   705
            Width           =   3900
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Préstamo"
            Height          =   195
            Left            =   180
            TabIndex        =   84
            Top             =   1086
            Width           =   1155
         End
         Begin VB.Label lblGenerales 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   1500
            TabIndex        =   83
            Top             =   1080
            Width           =   1530
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Monto Préstamo"
            Height          =   195
            Left            =   180
            TabIndex        =   82
            Top             =   1464
            Width           =   1155
         End
         Begin VB.Label lblGenerales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   1500
            TabIndex        =   81
            Top             =   1470
            Width           =   1530
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Fec. Ing. Judicial"
            Height          =   195
            Left            =   180
            TabIndex        =   80
            Top             =   1845
            Width           =   1200
         End
         Begin VB.Label lblGenerales 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   1515
            TabIndex        =   79
            Top             =   1845
            Width           =   1530
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3015
         Left            =   -74700
         TabIndex        =   61
         Top             =   600
         Width           =   9930
         Begin VB.CommandButton cmdActuacionesProc 
            Caption         =   "Actuaciones &Procesales"
            Height          =   375
            Left            =   7560
            TabIndex        =   93
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label lblExpediente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   7
            Left            =   1500
            TabIndex        =   77
            Top             =   2235
            Width           =   5505
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Estado Procesal"
            Height          =   195
            Left            =   165
            TabIndex        =   76
            Top             =   2235
            Width           =   1155
         End
         Begin VB.Label lblExpediente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   1500
            TabIndex        =   75
            Top             =   1845
            Width           =   5505
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Vía Procesal"
            Height          =   195
            Left            =   165
            TabIndex        =   74
            Top             =   1854
            Width           =   915
         End
         Begin VB.Label lblExpediente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   1500
            TabIndex        =   73
            Top             =   1470
            Width           =   1530
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   4740
            TabIndex        =   72
            Top             =   1470
            Width           =   585
         End
         Begin VB.Label lblExpediente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   5475
            TabIndex        =   71
            Top             =   1470
            Width           =   1530
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Monto Petitorio"
            Height          =   195
            Left            =   165
            TabIndex        =   70
            Top             =   1473
            Width           =   1065
         End
         Begin VB.Label lblExpediente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   1500
            TabIndex        =   69
            Top             =   1086
            Width           =   5505
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio Proc."
            Height          =   195
            Left            =   165
            TabIndex        =   68
            Top             =   1092
            Width           =   1050
         End
         Begin VB.Label lblExpediente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   1500
            TabIndex        =   67
            Top             =   708
            Width           =   5505
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Estudio Jurídico"
            Height          =   195
            Left            =   165
            TabIndex        =   66
            Top             =   711
            Width           =   1140
         End
         Begin VB.Label lblExpediente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   5475
            TabIndex        =   65
            Top             =   330
            Width           =   1530
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Comisión"
            Height          =   195
            Left            =   4710
            TabIndex        =   64
            Top             =   330
            Width           =   630
         End
         Begin VB.Label lblExpediente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   63
            Top             =   330
            Width           =   1530
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Expediente"
            Height          =   195
            Left            =   165
            TabIndex        =   62
            Top             =   330
            Width           =   1140
         End
      End
      Begin SICMACT.FlexEdit FE1 
         Height          =   2550
         Left            =   -74715
         TabIndex        =   3
         Top             =   690
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   4498
         Cols0           =   5
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Fecha-Descripcion-Importe-Origen Gasto"
         EncabezadosAnchos=   "400-1000-3450-1200-3450"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-L"
         FormatosEdit    =   "0-0-0-2-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FE2 
         Height          =   2550
         Left            =   -74715
         TabIndex        =   4
         Top             =   690
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   4498
         Cols0           =   10
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Fecha-Operacion-Importe-Capital-Interés-Mora-Gastos-Saldo-Movimiento"
         EncabezadosAnchos=   "400-1000-2400-1000-1000-900-900-900-1000-1"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-2-2-2-2-2-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagado"
         Height          =   195
         Left            =   -74610
         TabIndex        =   60
         Top             =   3375
         Width           =   960
      End
      Begin VB.Label lblTotalPagado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -73095
         TabIndex        =   59
         Top             =   3375
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total Pendiente"
         Height          =   195
         Left            =   -70950
         TabIndex        =   58
         Top             =   3375
         Width           =   1125
      End
      Begin VB.Label lblTotalPendiente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -69735
         TabIndex        =   57
         Top             =   3375
         Width           =   1455
      End
      Begin VB.Label lblGastosPagados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -69735
         TabIndex        =   56
         Top             =   3375
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Pagados"
         Height          =   195
         Left            =   -70950
         TabIndex        =   55
         Top             =   3375
         Width           =   1170
      End
      Begin VB.Label lblGastosAcumulados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -73095
         TabIndex        =   54
         Top             =   3375
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Gastos Acumulados"
         Height          =   195
         Left            =   -74610
         TabIndex        =   53
         Top             =   3375
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   8325
      TabIndex        =   6
      Top             =   7590
      Width           =   1140
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
      Left            =   7065
      TabIndex        =   5
      Top             =   7590
      Width           =   1140
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Datos Generales"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   555
      Width           =   10665
      Begin VB.Frame fraActual 
         Caption         =   "Actual"
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
         Height          =   2175
         Left            =   5205
         TabIndex        =   15
         Top             =   720
         Width           =   2355
         Begin VB.Label Label33 
            Caption         =   "Gastos"
            Height          =   195
            Left            =   165
            TabIndex        =   52
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label32 
            Caption         =   "Mora"
            Height          =   195
            Left            =   165
            TabIndex        =   51
            Top             =   1020
            Width           =   645
         End
         Begin VB.Label Label31 
            Caption         =   "Interes"
            Height          =   195
            Left            =   165
            TabIndex        =   50
            Top             =   645
            Width           =   645
         End
         Begin VB.Label Label30 
            Caption         =   "Capital"
            Height          =   195
            Left            =   165
            TabIndex        =   49
            Top             =   285
            Width           =   645
         End
         Begin VB.Label lblActual 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   945
            TabIndex        =   48
            Top             =   285
            Width           =   1245
         End
         Begin VB.Label lblActual 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   945
            TabIndex        =   47
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label lblActual 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   945
            TabIndex        =   46
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label lblActual 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   945
            TabIndex        =   45
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label lblActual 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Index           =   4
            Left            =   615
            TabIndex        =   44
            Top             =   1755
            Width           =   1575
         End
      End
      Begin VB.Frame fraPagado 
         Caption         =   "Pagado"
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
         Height          =   2175
         Left            =   2685
         TabIndex        =   14
         Top             =   720
         Width           =   2355
         Begin VB.Label Label24 
            Caption         =   "Gastos"
            Height          =   195
            Left            =   180
            TabIndex        =   43
            Top             =   1380
            Width           =   645
         End
         Begin VB.Label Label23 
            Caption         =   "Mora"
            Height          =   195
            Left            =   180
            TabIndex        =   42
            Top             =   1020
            Width           =   645
         End
         Begin VB.Label Label22 
            Caption         =   "Interes"
            Height          =   195
            Left            =   180
            TabIndex        =   41
            Top             =   645
            Width           =   645
         End
         Begin VB.Label Label21 
            Caption         =   "Capital"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   285
            Width           =   645
         End
         Begin VB.Label lblPagado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   39
            Top             =   285
            Width           =   1245
         End
         Begin VB.Label lblPagado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   38
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label lblPagado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   960
            TabIndex        =   37
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label lblPagado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   960
            TabIndex        =   36
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label lblPagado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Index           =   4
            Left            =   630
            TabIndex        =   35
            Top             =   1755
            Width           =   1575
         End
      End
      Begin VB.Frame fraIngreso 
         Caption         =   "Ingreso a Judicial"
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
         Height          =   2190
         Left            =   210
         TabIndex        =   13
         Top             =   705
         Width           =   2355
         Begin VB.Label Label15 
            Caption         =   "Gastos"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   1380
            Width           =   645
         End
         Begin VB.Label Label14 
            Caption         =   "Mora"
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   1020
            Width           =   645
         End
         Begin VB.Label Label13 
            Caption         =   "Interes"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   645
            Width           =   645
         End
         Begin VB.Label Label12 
            Caption         =   "Capital"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   285
            Width           =   645
         End
         Begin VB.Label lblIngreso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   30
            Top             =   285
            Width           =   1245
         End
         Begin VB.Label lblIngreso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   29
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label lblIngreso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   960
            TabIndex        =   28
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label lblIngreso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   960
            TabIndex        =   27
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label lblIngreso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Index           =   4
            Left            =   630
            TabIndex        =   26
            Top             =   1755
            Width           =   1575
         End
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   8790
         TabIndex        =   25
         Top             =   2505
         Width           =   1665
      End
      Begin VB.Label lblTipo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8790
         TabIndex        =   24
         Top             =   2130
         Width           =   1665
      End
      Begin VB.Label lblCondicion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8790
         TabIndex        =   23
         Top             =   1770
         Width           =   1665
      End
      Begin VB.Label lblMetodoLiquidacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9255
         TabIndex        =   22
         Top             =   1395
         Width           =   1200
      End
      Begin VB.Label lblDocIdentidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8790
         TabIndex        =   21
         Top             =   1035
         Width           =   1665
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   7680
         TabIndex        =   20
         Top             =   2505
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   7680
         TabIndex        =   19
         Top             =   2130
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Condición:"
         Height          =   195
         Left            =   7680
         TabIndex        =   18
         Top             =   1770
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Método Liquidación:"
         Height          =   195
         Left            =   7680
         TabIndex        =   17
         Top             =   1395
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Natural:"
         Height          =   195
         Left            =   7680
         TabIndex        =   16
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label lblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8175
         TabIndex        =   12
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   315
         Width           =   645
      End
      Begin VB.Label lblCodPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2565
         TabIndex        =   9
         Top             =   285
         Width           =   5475
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar ..."
      CausesValidation=   0   'False
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
      Left            =   9780
      TabIndex        =   1
      Top             =   135
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
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
      Height          =   360
      Left            =   9540
      TabIndex        =   7
      Top             =   7590
      Width           =   1140
   End
   Begin VB.Label lbl_Reprogramado 
      Caption         =   "REPROGRAMADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6120
      TabIndex        =   107
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmColRecRConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnSaldoCap As Currency
Dim fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency
Dim fnTasaInt As Double, fnTasaIntMorat As Double
Dim fnTipoCalcIntComp As Integer, fnTipoCalcIntMora As Integer
'0 --> No Calcula
'1 --> Capital
'2 --> Capital + Int Comp
'3 --> Capital + Int comp + Int Morat

Dim fnFormaCalcIntComp As Integer, fnFormaCalcIntMora As Integer
'0 INTERES SIMPLE
'1 INTERES COMPUESTO

Dim fnIntCompGenerado As Currency, fnIntMoraGenerado As Currency

Private Sub cmdActuacionesProc_Click()
    Call frmColRecActuacionesProc.Inicia("C", AXCodCta.NroCuenta)
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

'On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then
        Exit Sub
    End If
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstRecCanJud & "," & gColocEstRecCanCast

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        BuscaCredito (AXCodCta.NroCuenta)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub


Private Sub cmdCancelar_Click()
    Limpiar
    Call HabilitaControles(False, True, True)
    AXCodCta.SetFocusAge
    lbl_Reprogramado.Visible = False 'FRHU 20141105 Observacion
End Sub

Private Sub cmdImprimir_Click()
Dim loRep As COMNColocRec.NCOMColRecRConsulta
Dim lscadimp As String
Dim loPrevio As previo.clsprevio
Set loRep = New COMNColocRec.NCOMColRecRConsulta
loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
lscadimp = loRep.nRepo150000_ConsultaCreditoenCobranzaJudicial(AXCodCta.NroCuenta, gImpresora, gdFecSis)
Set loRep = Nothing
    
    If Len(Trim(lscadimp)) > 0 Then
        Set loPrevio = New previo.clsprevio
        loPrevio.Show lscadimp, "Consulta de Créditos en Cobranza Judicial", True
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Public Sub Inicia(ByVal sCaption As String)
    Me.Caption = sCaption
    SSTab1.Tab = 0
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    CargaParametros
    Me.Show 1
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AXCodCta.NroCuenta)
End Sub

Private Sub BuscaCredito(ByVal psNroContrato As String)
Dim lbok As Boolean
Dim lrDatCredito As New ADODB.Recordset
Dim lrDatTotales As New ADODB.Recordset
Dim lrDatGastos As New ADODB.Recordset
Dim lrDatTotalesGastos As New ADODB.Recordset
Dim lrDatAmortizaciones As New ADODB.Recordset
Dim lrDatExpediente As New ADODB.Recordset
Dim lrDatDatosGenerales As New ADODB.Recordset
Dim loValCred As COMDColocRec.DCOMColRecRConsulta
Dim i As Integer
Dim Saldo As Currency
Dim lsFecUltPago As String
Dim lnDiasUltTrans As Integer

On Error GoTo ControlError
    
    fnIntCompGenerado = 0
    fnIntMoraGenerado = 0
    'Carga Datos
    Set loValCred = New COMDColocRec.DCOMColRecRConsulta
        Set lrDatCredito = loValCred.dObtieneDatosCabeceraRecuperacion(psNroContrato)
        Set lrDatTotales = loValCred.dObtieneDatosTotalesRecuperacion(psNroContrato)
        Set lrDatGastos = loValCred.dObtieneGastosRecuperacion(psNroContrato)
        Set lrDatTotalesGastos = loValCred.dObtieneTotalesGastosRecuperacion(psNroContrato)
        Set lrDatAmortizaciones = loValCred.dObtieneListaAmortizaciones(psNroContrato)
        Set lrDatExpediente = loValCred.dObtieneDatosExpediente(psNroContrato)
        Set lrDatDatosGenerales = loValCred.dObtieneDatosGenerales(psNroContrato)
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Or (lrDatCredito.BOF And lrDatCredito.EOF) Then  ' Hubo un Errora
        MsgBox "No se encontro el Credito ", vbInformation, "Aviso"
        Limpiar
        Set lrDatCredito = Nothing
        Set lrDatTotales = Nothing
        Set lrDatGastos = Nothing
        Set lrDatTotalesGastos = Nothing
        Set lrDatAmortizaciones = Nothing
        Set lrDatExpediente = Nothing
        Set lrDatDatosGenerales = Nothing
        Exit Sub
    End If
    
    
    'INICIO ORCR-20140913*********
     Dim oCred2 As COMDCredito.DCOMCreditos
     Set oCred2 = New COMDCredito.DCOMCreditos
    
    lbl_Reprogramado.Visible = oCred2.CreditoReprogramado(psNroContrato)
    'FIN ORCR-20140913************
    
    lblCodPers.Caption = lrDatCredito!cperscod
    lblNomPers.Caption = lrDatCredito!cPersNombre
    lblMoneda.Caption = lrDatCredito!cMoneda
    lblDocIdentidad.Caption = "" & lrDatCredito!cPersIDnro
    lblMetodoLiquidacion.Caption = lrDatCredito!cMetLiquid
    lblCondicion.Caption = lrDatCredito!sCondicion
    If lrDatCredito!cTipo = "EXTRAJUDICIAL" Then
        lblCondicion.Caption = "VENCIDO"
    End If
    
    lblEstado.Caption = lrDatCredito!ssEstado
    
    'lsFecUltPago = CDate(fgFechaHoraGrab(lrDatCredito!cUltimaActualizacion))   'RIRO 20210801 Comentado
    lsFecUltPago = CDate(lrDatCredito!dUltimoPago)                              'RIRO 20210801 add
    
    'Tasa de Interes
    
    fnSaldoCap = IIf(IsNull(lrDatTotales!CapitalActual), 0, lrDatTotales!CapitalActual)
    fnSaldoIntComp = IIf(IsNull(lrDatTotales!InteresActual), 0, lrDatTotales!InteresActual)
    fnSaldoIntMorat = IIf(IsNull(lrDatTotales!MoraActual), 0, lrDatTotales!MoraActual)
    fnTasaInt = IIf(IsNull(lrDatDatosGenerales!nTasaInt), 0, lrDatDatosGenerales!nTasaInt)
    fnTasaIntMorat = IIf(IsNull(lrDatDatosGenerales!nTasaIntMor), 0, lrDatDatosGenerales!nTasaIntMor)
    Set lrDatCredito = Nothing
    
    If Not lrDatTotalesGastos.EOF Then
        lblGastosAcumulados.Caption = IIf(IsNull(lrDatTotalesGastos!Importe), 0, Format(lrDatTotalesGastos!Importe, "0.00"))
        lblGastosPagados.Caption = IIf(IsNull(lrDatTotalesGastos!Pagado), 0, Format(lrDatTotalesGastos!Pagado, "0.00"))
    End If
    
    Set lrDatTotalesGastos = Nothing
     
    Set FE1.Recordset = lrDatGastos
    Set FE2.Recordset = lrDatAmortizaciones
    
    Set lrDatAmortizaciones = Nothing
    Set lrDatGastos = Nothing
      
      
    With lrDatTotales
        lblPagado(0).Caption = Format(!CapitalPagado, "#,##0.00")
        lblPagado(1).Caption = Format(!InteresPagado, "#,##0.00")
        lblPagado(2).Caption = Format(!MoraPagado, "#,##0.00")
        lblPagado(3).Caption = Format(!GastoPagado, "#,##0.00")
        lblPagado(4).Caption = Format(!CapitalPagado + !InteresPagado + !MoraPagado + !GastoPagado, "#,##0.00")
        lblTotalPagado.Caption = lblPagado(4).Caption
        
        lblActual(0).Caption = Format(!CapitalActual, "#,##0.00")
        lblActual(1).Caption = Format(!InteresActual, "#,##0.00")
        lblActual(2).Caption = Format(!MoraActual, "#,##0.00")
        lblActual(3).Caption = Format(!GastoActual, "#,##0.00")
        lblActual(4).Caption = Format(!CapitalActual + !InteresActual + !MoraActual + !GastoActual, "#,##0.00")
        lblTotalPendiente.Caption = lblActual(4).Caption
       
        lblIngreso(0).Caption = Format(!CapitalPagado + !CapitalActual, "#,##0.00")
        lblIngreso(1).Caption = Format(!InteresPagado + !InteresActual, "#,##0.00")
        lblIngreso(2).Caption = Format(!MoraPagado + !MoraActual, "#,##0.00")
        lblIngreso(3).Caption = Format(!GastoPagado + !GastoActual, "#,##0.00")
        lblIngreso(4).Caption = Format(!CapitalPagado + !InteresPagado + !MoraPagado + !GastoPagado + !CapitalActual + !InteresActual + !MoraActual + !GastoActual, "#,##0.00")
    End With
    
    Set lrDatTotales = Nothing
    
    '****** Calcula el interes a la fecha
    lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(lsFecUltPago, "dd/mm/yyyy"))
    Dim loCalcula As COMNColocRec.NCOMColRecCalculos
    'Calcula el Int Comp Generado
    Set loCalcula = New COMNColocRec.NCOMColRecCalculos
        If fnTipoCalcIntComp = 0 Then ' NoCalcula
            fnIntCompGenerado = 0
        ElseIf fnTipoCalcIntComp = 1 Then ' En base al capital
            If fnFormaCalcIntComp = 1 Then 'INTERES COMPUESTO
                fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
            Else
                'INTERES SIMPLE
                fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
            End If
        ElseIf fnTipoCalcIntComp = 2 Then ' En base a capit + int Comp
            If fnFormaCalcIntComp = 1 Then
                fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp)
            Else
                fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp)
            End If
        ElseIf fnTipoCalcIntComp = 3 Then ' En base a capit + int Comp + int Morat
            If fnFormaCalcIntComp = 1 Then
                fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
            Else
                fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
            End If
        End If
        If fnTipoCalcIntMora = 0 Then  ' NoCalcula
            fnIntMoraGenerado = 0
        ElseIf fnTipoCalcIntMora = 1 Then ' En base al capital
            If fnFormaCalcIntMora = 1 Then 'INTERES COMPUESTO
                fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
            Else
                'INTERES SIMPLE
                fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
            End If
        ElseIf fnTipoCalcIntMora = 2 Then ' En base a capit + int Comp
            If fnFormaCalcIntMora = 1 Then
                fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp)
            Else
                fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp)
            End If
        ElseIf fnTipoCalcIntMora = 3 Then ' En base a capit + int Comp + int Morat
            If fnFormaCalcIntMora = 1 Then
                fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
            Else
                fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
            End If
        End If
          
    Set loCalcula = Nothing
    'Agregamos el Int Calculado al Saldo Int Comp
    lblActual(1).Caption = CCur(lblActual(1).Caption) + fnIntCompGenerado
    lblActual(2).Caption = CCur(lblActual(2).Caption) + fnIntMoraGenerado
    lblActual(4).Caption = CCur(lblActual(4).Caption) + fnIntCompGenerado + fnIntMoraGenerado
    '******
    If lrDatExpediente Is Nothing Or (lrDatExpediente.BOF And lrDatExpediente.EOF) Then    ' No hay datos de Expediente
    Else
        With lrDatExpediente
            lblExpediente(0).Caption = !cNumExp
            lblExpediente(2).Caption = !cPersNombre
            lblExpediente(3).Caption = !cPersDireccDomicilio
            lblExpediente(5).Caption = !sMoneda
            lblExpediente(1).Caption = !sTipoComision & " " & Format(!nValor, "#,##0.00")
            lblExpediente(4).Caption = Format(!nMonPetit, "#,##0.00")
            lblExpediente(6).Caption = "" & !sViaProcesal
            lblExpediente(7).Caption = "" & !sEstadoProceso
        End With
    End If
    Set lrDatExpediente = Nothing
    
    With lrDatDatosGenerales
        lblGenerales(0).Caption = !cdescripcion
        lblGenerales(1).Caption = !Analista
        lblGenerales(2).Caption = Format(!FechaPrestamo, "dd/mm/yyyy")
        lblGenerales(3).Caption = Format(!nMontoCol, "#,##0.00")
        lblGenerales(4).Caption = Format(!FechaIngreso, "dd/mm/yyyy")
        lblGenerales(5).Caption = IIf(IsNull(!nTasaInt), 0, !nTasaInt)
        lblGenerales(6).Caption = IIf(IsNull(!nTasaIntMor), 0, !nTasaIntMor)
    End With
    
    Set lrDatDatosGenerales = Nothing
    
    With FE2
        Saldo = CDbl(lblIngreso(0).Caption)
        For i = 1 To FE2.Rows - 1
            .TextMatrix(i, 3) = Format(CDbl(.TextMatrix(i, 4)) + CDbl(.TextMatrix(i, 5)) + CDbl(.TextMatrix(i, 6)) + CDbl(.TextMatrix(i, 7)), "#,##0.00")
            If Right(.TextMatrix(i, 2), 6) <> "130100" Then
                .TextMatrix(i, 8) = Format(Saldo - CDbl(.TextMatrix(i, 4)), "#,##0.00")
            Else
                .TextMatrix(i, 8) = Format(Saldo, "#,##0.00")
            End If
            Saldo = CDbl(.TextMatrix(i, 8))
        Next
    End With
    ' Muestra Negociacion
    MuestraDatosNegocia (psNroContrato)
    cmdImprimir.Enabled = True
    If cmdImprimir.Visible Then
        cmdImprimir.SetFocus
    End If
        
    AXCodCta.Enabled = False
    Call HabilitaControles(True, True, True)
    'Me.AXMontoPago.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub HabilitaControles(ByVal pbcmdImprimir As Boolean, ByVal pbCmdCancelar As Boolean, _
            ByVal pbCmdSalir As Boolean)
    cmdImprimir.Enabled = pbcmdImprimir
    cmdCancelar.Enabled = pbCmdCancelar
    cmdSalir.Enabled = pbCmdSalir
End Sub

Private Sub Limpiar()
Dim i As Integer

    Me.AXCodCta.Enabled = True
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones

    lblCodPers.Caption = ""
    lblNomPers.Caption = ""
    lblCodPers.Caption = ""
    lblNomPers.Caption = ""
    lblMoneda.Caption = ""
    lblDocIdentidad.Caption = ""
    lblMetodoLiquidacion.Caption = ""
    lblCondicion.Caption = ""
    lblTipo.Caption = ""
    lblEstado.Caption = ""
    
    For i = 0 To 4
        lblIngreso(i).Caption = ""
        lblPagado(i).Caption = ""
        lblActual(i).Caption = ""
        lblExpediente(i).Caption = ""
        lblGenerales(i).Caption = ""
    Next
    
    For i = 5 To 7
        lblExpediente(i).Caption = ""
    Next
    
    lblGastosAcumulados.Caption = ""
    lblGastosPagados.Caption = ""
    lblTotalPagado.Caption = ""
    lblTotalPendiente.Caption = ""
    Set FE1.Recordset = Nothing
    Set FE2.Recordset = Nothing
    FE1.Rows = 2
    FE2.Rows = 2
End Sub
 
Public Sub MuestraPosicionCliente(ByVal psCodCta As String)
    If psCodCta <> "" Then
        AXCodCta.NroCuenta = psCodCta
        CargaParametros 'MADM 20120127
        BuscaCredito (psCodCta)
        cmdCancelar.Enabled = False
        AXCodCta.Enabled = False
        cmdBuscar.Enabled = False
        Me.Show 1
    End If

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

'TODO VERIFICA FUNCION MUSTRADATOSNEGOCIO *************************
'******************************************************************
Private Sub MuestraDatosNegocia(ByVal psCodCta As String)
On Error GoTo ControlError
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Dim reg As New ADODB.Recordset
Dim lsSQL As String
Dim L As ListItem
    ' Busca la Negociacion
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set reg = lcRec.ObtenerDatosCredNegociacion(psCodCta)
    Set lcRec = Nothing
    If reg.BOF And reg.EOF Then
        'MsgBox " No se Tiene Negociaciones Vigentes para Credito  " & psCodCta, vbInformation, " Aviso "
        'LimpiaDatos
        'AXCodCta.Enabled = True
        Exit Sub
    Else
        ' Mostrar los datos de Negociacion
        Me.txtNegNro.Text = reg!cNroNeg
        Me.TxtNegVigencia.Text = Format(reg!dFecVig, "dd/mm/yyyy")
        Me.txtNegEstado.Text = IIf(reg!cEstado = "V", "Vigente", "Cancelado")
        Me.txtNegMonto.Text = Format(reg!nMontoNeg, "#0.00")
        Me.txtNegCuotas.Text = Format(reg!nCuotasNeg, "#0.00")
        Me.txtNegComenta.Text = IIf(IsNull(reg!cComenta), "", reg!cComenta)
        reg.Close
        Set reg = Nothing
    End If
        ' Busca Plan de Pagos de Negociacion
        Set lcRec = New COMDColocRec.DCOMColRecNegociacion
           Set reg = lcRec.ObtenerPlanPagosNegocia(psCodCta, txtNegNro.Text)
        Set lcRec = Nothing
        
        If reg.BOF And reg.EOF Then
            MsgBox " Negociacion No Posee Plan Pagos " & psCodCta, vbInformation, " Aviso "
        Else
            ' Mostrar Plan de Pagos
            reg.MoveFirst
            Do While Not reg.EOF
                Set L = lvwCalendario.ListItems.Add(, , Trim(Str(reg!nNroCuota)))
                L.SubItems(1) = Format(reg!dFecVenc, "dd/mm/yyyy")
                L.SubItems(2) = Format(reg!nMonto, "#0.00")
                L.SubItems(3) = Format(reg!nMontoPag, "#0.00")
                L.SubItems(4) = reg!cEstado
                reg.MoveNext
            Loop
        End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

Private Sub CargaParametros()
Dim loParam As COMDConstSistema.NCOMConstSistema 'NConstSistemas
Set loParam = New COMDConstSistema.NCOMConstSistema
    fnTipoCalcIntComp = loParam.LeeConstSistema(151)
    fnTipoCalcIntMora = loParam.LeeConstSistema(152)
    fnFormaCalcIntComp = loParam.LeeConstSistema(202) ' CMACICA
    fnFormaCalcIntMora = loParam.LeeConstSistema(203) ' CMACICA
Set loParam = Nothing
End Sub
