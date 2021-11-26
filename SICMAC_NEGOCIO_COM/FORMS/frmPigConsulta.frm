VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPigConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Creditos Pignoraticios"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ForeColor       =   &H8000000D&
   Icon            =   "frmPigConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVer 
      Caption         =   "&Detalle"
      Height          =   405
      Left            =   5895
      TabIndex        =   74
      Top             =   7215
      Width           =   1020
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   8160
      TabIndex        =   1
      Top             =   7215
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   7020
      TabIndex        =   0
      Top             =   7215
      Width           =   1020
   End
   Begin VB.Frame fraContenedor 
      Height          =   7110
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   9075
      Begin VB.Frame Frame4 
         Caption         =   "Estado"
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
         Height          =   1515
         Left            =   7350
         TabIndex        =   8
         Top             =   1530
         Width           =   1620
         Begin VB.OptionButton obEstado 
            Caption         =   "Cancelados"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   55
            Top             =   948
            Width           =   1350
         End
         Begin VB.OptionButton obEstado 
            Caption         =   "Activos"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   54
            Top             =   195
            Value           =   -1  'True
            Width           =   1350
         End
         Begin VB.OptionButton obEstado 
            Caption         =   "Rematados"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   53
            Top             =   446
            Width           =   1350
         End
         Begin VB.OptionButton obEstado 
            Caption         =   "Adjudicados"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   52
            Top             =   697
            Width           =   1350
         End
         Begin VB.OptionButton obEstado 
            Caption         =   "Todos"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   51
            Top             =   1200
            Width           =   1350
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Relación de Creditos"
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
         Height          =   1515
         Left            =   135
         TabIndex        =   6
         Top             =   1545
         Width           =   7140
         Begin MSComctlLib.ListView lstCreditos 
            Height          =   1275
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2249
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cuenta"
               Object.Width           =   3140
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Situacion"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Agencia Origen"
               Object.Width           =   4410
            EndProperty
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Datos del Cliente"
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
         Height          =   825
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   660
         Width           =   8805
         Begin MSComctlLib.ListView lstClientes 
            Height          =   585
            Left            =   90
            TabIndex        =   5
            Top             =   180
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   1032
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   16777215
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
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
               Text            =   "Codigo"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cliente"
               Object.Width           =   5468
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Doc Ident."
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tipo de Cliente"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Calif.SBS"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   360
         Left            =   3930
         Picture         =   "frmPigConsulta.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar ..."
         Top             =   255
         Width           =   420
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3615
         Left            =   60
         TabIndex        =   9
         Top             =   3330
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   6376
         _Version        =   393216
         Style           =   1
         TabHeight       =   582
         TabMaxWidth     =   176
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmPigConsulta.frx":040C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fraContenedor(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Garantias"
         TabPicture(1)   =   "frmPigConsulta.frx":0428
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstGarantias"
         Tab(1).Control(1)=   "Tasaciones"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Movimientos"
         TabPicture(2)   =   "frmPigConsulta.frx":0444
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstDesembolsos"
         Tab(2).Control(1)=   "lstPagos"
         Tab(2).Control(2)=   "frDesembolsos"
         Tab(2).Control(3)=   "fmPagos"
         Tab(2).ControlCount=   4
         Begin VB.Frame fraContenedor 
            Height          =   1395
            Index           =   2
            Left            =   90
            TabIndex        =   28
            Top             =   375
            Width           =   8640
            Begin VB.Label lblMontoPrestamo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   1170
               TabIndex        =   57
               Top             =   375
               Width           =   945
            End
            Begin VB.Label lblVencimiento 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   3435
               TabIndex        =   56
               Top             =   405
               Width           =   945
            End
            Begin VB.Label Label5 
               Caption         =   "Vencimiento"
               Height          =   165
               Left            =   2430
               TabIndex        =   50
               Top             =   420
               Width           =   960
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Dias Atraso"
               Height          =   180
               Index           =   7
               Left            =   4965
               TabIndex        =   49
               Top             =   480
               Width           =   825
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Uso  Linea"
               Height          =   195
               Index           =   9
               Left            =   4590
               TabIndex        =   48
               Top             =   2085
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label2 
               Caption         =   "Emision"
               Height          =   195
               Left            =   2460
               TabIndex        =   47
               Top             =   870
               Width           =   750
            End
            Begin VB.Label Label3 
               Caption         =   "Ult.Transacc"
               Height          =   195
               Left            =   4740
               TabIndex        =   46
               Top             =   1980
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label Label4 
               Caption         =   "Fec.Rescate"
               Height          =   225
               Left            =   4680
               TabIndex        =   45
               Top             =   2175
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label Label9 
               Caption         =   "Tasa Interes"
               Height          =   270
               Left            =   6765
               TabIndex        =   44
               Top             =   465
               Width           =   1170
            End
            Begin VB.Label Label10 
               Caption         =   "Tasa Moratoria"
               Height          =   165
               Left            =   6765
               TabIndex        =   43
               Top             =   900
               Width           =   1170
            End
            Begin VB.Label Label7 
               Caption         =   "Aviso Remate"
               Height          =   195
               Left            =   6600
               TabIndex        =   42
               Top             =   2010
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Nro.Duplic"
               Height          =   180
               Left            =   4965
               TabIndex        =   41
               Top             =   885
               Width           =   915
            End
            Begin VB.Label lblEtiqueta 
               Caption         =   "Plazo "
               Height          =   210
               Index           =   11
               Left            =   165
               TabIndex        =   40
               Top             =   855
               Width           =   525
            End
            Begin VB.Label Label1 
               Caption         =   "Prestamo"
               Height          =   225
               Left            =   165
               TabIndex        =   39
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label lblPlazo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1170
               TabIndex        =   38
               Top             =   810
               Width           =   435
            End
            Begin VB.Label lblFecRescate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   5775
               TabIndex        =   37
               Top             =   2160
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblTasaMora 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   7965
               TabIndex        =   36
               Top             =   810
               Width           =   600
            End
            Begin VB.Label lblEmision 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   3420
               TabIndex        =   35
               Top             =   810
               Width           =   945
            End
            Begin VB.Label lblUsoLinea 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   5775
               TabIndex        =   34
               Top             =   2040
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.Label lblTasaInteres 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   7965
               TabIndex        =   33
               Top             =   420
               Width           =   600
            End
            Begin VB.Label lblUltPago 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   5835
               TabIndex        =   32
               Top             =   1920
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblAvisoRemate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   7800
               TabIndex        =   31
               Top             =   1995
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblDuplicados 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   5940
               TabIndex        =   30
               Top             =   810
               Width           =   630
            End
            Begin VB.Label lblDiasAtraso 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   5940
               TabIndex        =   29
               Top             =   435
               Width           =   630
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Proyección de Deuda"
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
            Height          =   2055
            Left            =   105
            TabIndex        =   25
            Top             =   3915
            Width           =   5805
            Begin VB.TextBox txtProyAmortizacion 
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
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4545
               TabIndex        =   91
               Top             =   435
               Width           =   945
            End
            Begin VB.TextBox txtProyNuevoPlazo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1425
               TabIndex        =   89
               Top             =   1299
               Width           =   480
            End
            Begin MSComCtl2.DTPicker dtpPosicion 
               Height          =   285
               Left            =   1410
               TabIndex        =   26
               Top             =   180
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               Format          =   63373313
               CurrentDate     =   37516
            End
            Begin MSComCtl2.DTPicker dtpProyProxVcto 
               Height          =   285
               Left            =   1425
               TabIndex        =   90
               Top             =   1695
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   503
               _Version        =   393216
               Format          =   63373313
               CurrentDate     =   37516
            End
            Begin VB.Line Line4 
               X1              =   3030
               X2              =   5655
               Y1              =   1875
               Y2              =   1875
            End
            Begin VB.Line Line3 
               X1              =   3030
               X2              =   3030
               Y1              =   255
               Y2              =   1890
            End
            Begin VB.Line Line2 
               X1              =   5640
               X2              =   5640
               Y1              =   255
               Y2              =   1875
            End
            Begin VB.Line Line1 
               X1              =   3030
               X2              =   5640
               Y1              =   255
               Y2              =   255
            End
            Begin VB.Label Label23 
               Caption         =   "Nuevo Saldo"
               Height          =   210
               Left            =   3225
               TabIndex        =   80
               Top             =   825
               Width           =   1215
            End
            Begin VB.Label lblProyNuevoPagMin 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   87
               Top             =   1485
               Width           =   945
            End
            Begin VB.Label Label33 
               Caption         =   "Nuevo Pago Min."
               Height          =   210
               Left            =   3225
               TabIndex        =   86
               Top             =   1530
               Width           =   1260
            End
            Begin VB.Label lblProyNuevoSaldo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   84
               Top             =   795
               Width           =   945
            End
            Begin VB.Label lblProyNuevaDeuda 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   83
               Top             =   1155
               Width           =   945
            End
            Begin VB.Label Label22 
               Caption         =   "Nueva Deuda"
               Height          =   210
               Left            =   3225
               TabIndex        =   79
               Top             =   1185
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "A Pagar"
               Height          =   195
               Left            =   3240
               TabIndex        =   77
               Top             =   465
               Width           =   1215
            End
            Begin VB.Label Label35 
               Caption         =   "Dias"
               Height          =   195
               Left            =   2010
               TabIndex        =   88
               Top             =   1380
               Width           =   420
            End
            Begin VB.Label Label31 
               Caption         =   "Nuevo Plazo"
               Height          =   195
               Left            =   135
               TabIndex        =   85
               Top             =   1383
               Width           =   990
            End
            Begin VB.Label lblProyDeuda 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1425
               TabIndex        =   82
               Top             =   573
               Width           =   945
            End
            Begin VB.Label lblProyPagMin 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1425
               TabIndex        =   81
               Top             =   936
               Width           =   945
            End
            Begin VB.Label Label21 
               Caption         =   "Proximo Vcto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   135
               TabIndex        =   78
               Top             =   1755
               Width           =   1260
            End
            Begin VB.Label Label11 
               Caption         =   "Pago Minimo"
               Height          =   210
               Left            =   135
               TabIndex        =   76
               Top             =   997
               Width           =   1020
            End
            Begin VB.Label Label6 
               Caption         =   "Deuda Total"
               Height          =   195
               Left            =   135
               TabIndex        =   75
               Top             =   626
               Width           =   1125
            End
            Begin VB.Label Label18 
               Caption         =   "Posicion al"
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
               Left            =   135
               TabIndex        =   27
               Top             =   255
               Width           =   975
            End
         End
         Begin VB.Frame frame1 
            Caption         =   "Deuda Actual"
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
            Height          =   1200
            Left            =   120
            TabIndex        =   12
            Top             =   1920
            Width           =   8640
            Begin VB.Label Label12 
               Caption         =   "Saldo Capital"
               Height          =   180
               Left            =   255
               TabIndex        =   24
               Top             =   435
               Width           =   1065
            End
            Begin VB.Label Label13 
               Caption         =   "Interes"
               Height          =   210
               Left            =   255
               TabIndex        =   23
               Top             =   690
               Width           =   1080
            End
            Begin VB.Label Label14 
               Caption         =   "Mora"
               Height          =   165
               Left            =   2910
               TabIndex        =   22
               Top             =   465
               Width           =   525
            End
            Begin VB.Label Label15 
               Caption         =   "Otros"
               Height          =   225
               Left            =   2910
               TabIndex        =   21
               Top             =   750
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Label Label16 
               Caption         =   "Pago Minimo"
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
               Left            =   195
               TabIndex        =   20
               Top             =   2280
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label Label17 
               Caption         =   "Total      S/."
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
               Left            =   5760
               TabIndex        =   19
               Top             =   795
               Width           =   1185
            End
            Begin VB.Label lblSaldoCapital 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   1545
               TabIndex        =   18
               Top             =   390
               Width           =   945
            End
            Begin VB.Label lblInteres 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   1545
               TabIndex        =   17
               Top             =   675
               Width           =   945
            End
            Begin VB.Label lblMora 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   4200
               TabIndex        =   16
               Top             =   435
               Width           =   945
            End
            Begin VB.Label lblOtros 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4200
               TabIndex        =   15
               Top             =   720
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblTotal 
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
               Height          =   255
               Index           =   19
               Left            =   7185
               TabIndex        =   14
               Top             =   765
               Width           =   1260
            End
            Begin VB.Label lblPagoMinimo 
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
               Height          =   255
               Index           =   20
               Left            =   1485
               TabIndex        =   13
               Top             =   2250
               Visible         =   0   'False
               Width           =   945
            End
         End
         Begin VB.Frame Tasaciones 
            Height          =   1080
            Left            =   -74895
            TabIndex        =   11
            Top             =   480
            Width           =   8715
            Begin VB.Label lblPiezas 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7485
               TabIndex        =   69
               Top             =   690
               Width           =   945
            End
            Begin VB.Label lblTasAdic 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1485
               TabIndex        =   68
               Top             =   1215
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblTasacion 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1215
               TabIndex        =   67
               Top             =   690
               Width           =   945
            End
            Begin VB.Label lblPNeto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4725
               TabIndex        =   66
               Top             =   690
               Width           =   945
            End
            Begin VB.Label lblPBruto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4710
               TabIndex        =   65
               Top             =   285
               Width           =   945
            End
            Begin VB.Label lblTasador 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1215
               TabIndex        =   64
               Top             =   285
               Width           =   915
            End
            Begin VB.Label lblTxtPiezas 
               Caption         =   "Nro.Piezas"
               Height          =   195
               Left            =   6540
               TabIndex        =   63
               Top             =   720
               Width           =   870
            End
            Begin VB.Label lblTxtTasacion 
               Caption         =   "Tot.Tasación"
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblTxtTasAdic 
               Caption         =   "Tot.Tasac Adic."
               Height          =   195
               Left            =   165
               TabIndex        =   61
               Top             =   1245
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lblTxtPBruto 
               Caption         =   "Tot.Peso Bruto"
               Height          =   195
               Left            =   3390
               TabIndex        =   60
               Top             =   330
               Width           =   1200
            End
            Begin VB.Label lblTxtPNeto 
               Caption         =   "Tot. PNeto"
               Height          =   195
               Left            =   3405
               TabIndex        =   59
               Top             =   705
               Width           =   915
            End
            Begin VB.Label lblTxtTasador 
               Caption         =   "Tasador"
               Height          =   195
               Left            =   135
               TabIndex        =   58
               Top             =   315
               Width           =   825
            End
         End
         Begin MSComctlLib.ListView lstGarantias 
            Height          =   1665
            Left            =   -74910
            TabIndex        =   10
            Top             =   1575
            Width           =   8685
            _ExtentX        =   15319
            _ExtentY        =   2937
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "It"
               Object.Width           =   441
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nro Piezas"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Material"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Peso Bruto"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Peso Neto"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Tasacion"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Descripción"
               Object.Width           =   2822
            EndProperty
         End
         Begin MSComctlLib.ListView lstDesembolsos 
            Height          =   1065
            Left            =   -74895
            TabIndex        =   72
            Top             =   660
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   1879
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Hora"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Nro.Mvto."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Agencia"
               Object.Width           =   4666
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Prestamo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Com.Tasación"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Usuario"
               Object.Width           =   1058
            EndProperty
         End
         Begin MSComctlLib.ListView lstPagos 
            Height          =   1335
            Left            =   -74865
            TabIndex        =   73
            Top             =   2040
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   2355
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
            NumItems        =   16
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Hora"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Nro.Mvto"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Operación"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Agencia"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Monto Pagado"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Usuario"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Capital"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "InteresComp"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "InteresMora"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Text            =   "ComVcda"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Text            =   "ComServ"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Text            =   "DerRemate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   14
               Text            =   "Penalidad"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   15
               Text            =   "CustDiferida"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Frame frDesembolsos 
            Caption         =   "Desembolsos"
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
            Height          =   1335
            Left            =   -74955
            TabIndex        =   70
            Top             =   465
            Width           =   8745
         End
         Begin VB.Frame fmPagos 
            Caption         =   "Pagos"
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
            Left            =   -74955
            TabIndex        =   71
            Top             =   1800
            Width           =   8700
         End
      End
      Begin SICMACT.ActXCodCta AxCodCta 
         Height          =   405
         Left            =   180
         TabIndex        =   92
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   714
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPigConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************'
'* CONSULTA DE CONTRATOS DE PIGNORATICIOS                                       *'
'* Archivo      :  frmPigConsulta.frm                                                                    *'
'* Resumen      :  Para un contrato todas sus condiciones y sus   *'
'*                  caracteristicas, asi como sus movimientos y garantias.                  *'
'***************************************************************************************'
Option Explicit

Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String

Dim fnVarPlzMin As Currency
Dim fnVarPlzMax As Currency
Dim fsEstados As String
Dim fsCliente As String
Dim lnEstadoCont As Integer
Dim fsCtaCod As String
Dim fnVarDiasVigencia As Integer
Dim fnVarDiasIntereses As Long
Dim fnVarDiasVencidos As Integer
Dim fnVarCapitalPagado As Currency    '*************** Se obtiene de fnVarPorcCapMin * Saldo del Prestamo.
Dim fnVarPorcCapMin As Currency       '*************** Porcentaje!! Minimo de capital a Amortizar.
Dim fnVarCapMinimo As Currency        '*************** Capital minimo que se puede amortizar a un credito.
Dim fnVarMontoCol As Currency
Dim fnVarSaldo As Currency
Dim fnVarIntCompensatorio As Currency
Dim fnVarIntMoratorio As Currency
Dim fnVarPenalidad As Currency
Dim fnVarDerRemate As Currency
Dim fnVarComServicio As Currency
Dim fnVarComVencida As Currency
Dim fnVarPagoMinimo As Currency
Dim fnTipTasacion As Integer

Dim fnVarTasaPreparacionRemate As Currency
Dim fnVarTasaImpuesto As Currency
Dim fnVarTasaCustodiaVencida As Currency
Dim fnVarTasaCustodia As Currency
Dim fnVarDiasCambCart As Currency
Dim fnVarTopRenovaciones As Currency
Dim fnVarTopRenovacionesNuevo As Currency


Public Sub Inicio()
Dim loFunct As DPigFunciones
   
    Set loFunct = New DPigFunciones
         fnVarPlzMin = loFunct.GetParamValor(8015)
         fnVarPlzMax = loFunct.GetParamValor(8016)
    Set loFunct = Nothing
    
    lstClientes.ListItems.Clear
    Limpiar
    SSTab1.Enabled = False
    Frame4.Enabled = False
    LstCreditos.Enabled = False
    cmdVer.Visible = False
    cmdCancelar.Enabled = False
    Me.Show 1
End Sub

Private Sub Limpiar()
   LstCreditos.ListItems.Clear
   LstGarantias.ListItems.Clear
   lstDesembolsos.ListItems.Clear
   lstPagos.ListItems.Clear
   dtpPosicion.value = Format(gdFecSis, "dd/mm/yyyy")
   lblProyDeuda.Caption = Format$(0, "###,##0.00 ")
   lblProyPagMin.Caption = Format$(0, "###,##0.00 ")
   dtpProyProxVcto.value = Format(gdFecSis, "dd/mm/yyyy")
   lblMontoPrestamo(0).Caption = Format$(0, "##,##0.00 ")
   lblPlazo(1).Caption = Format$(0, "#0")
   lblVencimiento(5).Caption = Format$("  /  /    ", "dd/mm/yyyy")
   lblEmision(4).Caption = Format$("  /  /    ", "dd/mm/yyyy")
   lblUltPago(6).Caption = Format$("  /  /    ", "dd/mm/yyyy")
   lblFecRescate(7).Caption = Format$("  /  /    ", "dd/mm/yyyy")
   lblUsoLinea(8).Caption = Format$(0, "#0")
   lblDiasAtraso(9).Caption = Format$(0, "#0")
   lblDuplicados(11).Caption = Format$(0, "#0")
   lblTasaInteres(12).Caption = Format$(0, "#0.00 ")
   lblTasaMora(13).Caption = Format$(0, "#0.00 ")
   lblAvisoRemate(14).Caption = ""
   lblSaldoCapital(15).Caption = Format$(0, "##,##0.00 ")
   LblMora(17).Caption = Format$(0, "##,##0.00 ")
   lblInteres(16).Caption = Format$(0, "##,##0.00 ")
   lblOtros.Caption = Format$(0, "##,##0.00 ")
   LblTotal(19).Caption = Format$(0, "##,##0.00 ")
   lblPagoMinimo(20).Caption = Format$(0, "##,##0.00 ")
   txtProyNuevoPlazo.Text = Format$(0, "#0")
   lblTasador.Caption = ""
   lblTasacion.Caption = Format$(0, "###,##0.00 ")
   lblTasAdic.Caption = Format$(0, "###,##0.00 ")
   LblPBruto.Caption = Format$(0, "###,##0.00 ")
   LblPNeto.Caption = Format$(0, "###,##0.00 ")
   lblPiezas.Caption = Format$(0, "#0 ")
   txtProyAmortizacion.Text = Format$(0, "###,##0.00 ")
   lblProyNuevoSaldo.Caption = Format$(0, "###,##0.00 ")
   lblProyNuevaDeuda.Caption = Format$(0, "###,##0.00 ")
   lblProyNuevoPagMin.Caption = Format$(0, "###,##0.00 ")
   AXCodCta.Cuenta = ""
   AXCodCta.Age = ""
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
Dim oDatos As dPigContrato
Dim rs As Recordset
Dim lsCodCte As String
Dim lsCodCta As String

If KeyAscii = 13 Then
    lsCodCta = AXCodCta.GetCuenta
    
    Set oDatos = New dPigContrato
    Set rs = oDatos.dObtieneDatosCreditoPignoraticioPersonas(lsCodCta)
    If Not (rs.EOF And rs.BOF) Then
        lsCodCte = rs!cPerscod
    Else
        MsgBox "Número de Contrato no Válido ó Contrato no Existe", vbInformation, "Aviso"
        Exit Sub
    End If
    MuestraDatosPersona (lsCodCte)
    Call MuestraContratos(lsCodCte, lsCodCta)
    MuestraDetalleContrato (lsCodCta)
End If

End Sub

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As DColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

BuscaContratos (lsPersCod)
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Busca contratos del cliente
Private Sub BuscaContratos(ByVal pscliente As String)
Dim lbOk As Boolean
Dim lstTmpCliente As ListItem
Dim lrCredPigPersonas As ADODB.Recordset
Dim loMuestraDatos As dPigContrato

     ' Mostra Datos sobre el Cliente
     Set lrCredPigPersonas = New ADODB.Recordset
     Set loMuestraDatos = New dPigContrato
         Set lrCredPigPersonas = loMuestraDatos.dObtieneDatosClienteCredPignoraticio(pscliente)
     Set loMuestraDatos = Nothing
    
    If (lrCredPigPersonas.EOF And lrCredPigPersonas.BOF) Or lrCredPigPersonas Is Nothing Then
        MsgBox "Cliente no posee Contratos en Pignoraticios ", vbInformation, "Aviso"
        Exit Sub
    Else
        lstClientes.ListItems.Clear
        Do While Not lrCredPigPersonas.EOF
            Set lstTmpCliente = lstClientes.ListItems.Add(, , Trim(lrCredPigPersonas!cPerscod))
                  lstTmpCliente.SubItems(1) = Trim(PstaNombre(lrCredPigPersonas!cPersNombre, False))
                  lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrCredPigPersonas!NroDNI), "", lrCredPigPersonas!NroDNI))
                  lstTmpCliente.SubItems(3) = Trim(IIf(IsNull(lrCredPigPersonas!DescCalif), "", lrCredPigPersonas!DescCalif))
                  lstTmpCliente.SubItems(4) = Trim(IIf(IsNull(lrCredPigPersonas!DescSBS), "", lrCredPigPersonas!DescSBS))
                  fsCliente = Trim(lrCredPigPersonas!cPerscod)
            lrCredPigPersonas.MoveNext
        Loop
    End If
    obEstado_Click (4)
    Set lrCredPigPersonas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Public Function MuestraContratos(ByVal pscliente As String, Optional ByVal psCtaCod As String) As Boolean
Dim lrDatos As ADODB.Recordset
Dim loValContrato As dPigContrato
Dim lstTmpCreditos As ListItem
    
    Set lrDatos = New ADODB.Recordset
    Set loValContrato = New dPigContrato
        If psCtaCod <> "" Then
            Set lrDatos = loValContrato.dBuscaContratosConsultaCredPignoraticio(pscliente, fsEstados, psCtaCod)
        Else
            Set lrDatos = loValContrato.dBuscaContratosConsultaCredPignoraticio(pscliente, fsEstados)
        End If
    Set loValContrato = Nothing
     
     If lrDatos Is Nothing Then
        MuestraContratos = False
        Exit Function
    Else
        fsCtaCod = lrDatos!cCtaCod
        LstCreditos.ListItems.Clear
        Do While Not lrDatos.EOF
            Set lstTmpCreditos = LstCreditos.ListItems.Add(, , Trim(lrDatos!cCtaCod))
                  lstTmpCreditos.SubItems(1) = mfgEstCredPigDesc(lrDatos!nPrdEstado)
                  lstTmpCreditos.SubItems(2) = Trim(lrDatos!DescAge)
             lrDatos.MoveNext
        Loop
        MuestraContratos = True
    End If

End Function

Private Sub cmdCancelar_Click()
   Limpiar
   lstClientes.ListItems.Clear
   LstCreditos.Enabled = False
   SSTab1.Enabled = False
   Frame4.Enabled = False
   cmdBuscar.Enabled = True
   cmdCancelar.Enabled = False
   AXCodCta.Enabled = True
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdVer_Click()
frmPigConsultaDetPagos.Inicio CCur(lstPagos.SelectedItem.SubItems(6)), CCur(lstPagos.SelectedItem.SubItems(8)), CCur(IIf(lstPagos.SelectedItem.SubItems(9) = "", 0, lstPagos.SelectedItem.SubItems(9))), _
                                             CCur(IIf(lstPagos.SelectedItem.SubItems(10) = "", 0, lstPagos.SelectedItem.SubItems(10))), CCur(IIf(lstPagos.SelectedItem.SubItems(11) = "", 0, lstPagos.SelectedItem.SubItems(11))), CCur(IIf(lstPagos.SelectedItem.SubItems(13) = "", 0, lstPagos.SelectedItem.SubItems(13))), _
                                             CCur(IIf(lstPagos.SelectedItem.SubItems(14) = "", 0, lstPagos.SelectedItem.SubItems(14))), CCur(IIf(lstPagos.SelectedItem.SubItems(12) = "", 0, lstPagos.SelectedItem.SubItems(12))), CCur(IIf(lstPagos.SelectedItem.SubItems(15) = "", 0, lstPagos.SelectedItem.SubItems(15)))
End Sub

Private Sub dtpPosicion_Change()

Dim lrDatos As ADODB.Recordset
Dim loValContrato As dPigContrato
Dim lrFunctVal As ADODB.Recordset
Dim loFunct As DPigFunciones
Dim loCalc As NPigCalculos

'    If lnEstadoCont = 2802 Or lnEstadoCont = 2803 Or lnEstadoCont = 2804 Then
'
'        If dtpPosicion.value < gdFecSis Then
'            dtpPosicion.value = gdFecSis
'            Exit Sub
'        End If
'
'        fnVarDiasVigencia = DateDiff("d", Format(lblEmision(4).Caption, "dd/mm/yyyy"), Format(dtpPosicion.value, "dd/mm/yyyy"))
'        fnVarDiasIntereses = DateDiff("d", Format(lblUltPago(6), "dd/mm/yyyy"), Format(dtpPosicion.value, "dd/mm/yyyy"))
'        fnVarDiasVencidos = DateDiff("d", Format(lblVencimiento(5), "dd/mm/yyyy"), Format(dtpPosicion.value, "dd/mm/yyyy"))
'
'        Set loCalc = New NPigCalculos
'             fnVarIntCompensatorio = Format(loCalc.nCalculaIntCompensatorio(fnVarSaldo, Val(lblTasaInteres(12)), fnVarDiasIntereses), "#0.00")
'
'        If fnVarDiasVencidos > 0 Then         '***** Si desea pagar pasado el vencimiento
'            fnVarIntMoratorio = Format(loCalc.nCalculaIntCompensatorio(fnVarSaldo, Val(lblTasaMora(13)), fnVarDiasVencidos), "#0.00")
'        Else
'            fnVarIntMoratorio = Format(0, "#0.00")
'        End If
'
'        Set lrFunctVal = New ADODB.Recordset
'             Set loFunct = New DPigFunciones
'
'        If fnVarDiasVigencia < loFunct.GetParamValor(8019) Then         '***** Si esta afecto a Penalidad
'           Set lrFunctVal = loFunct.GetConceptoValor(Val(gColPigConceptoCodPenalidad))
'           fnVarPenalidad = Format(loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!NVALOR, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
'                                                                IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldo), "#0.00")
'        Else
'            fnVarPenalidad = Format(0, "#0.00")
'        End If
'
'        If fnVarDiasVencidos > loFunct.GetParamValor(8021) Then        '***** Si se le cobrara Comision de Vencimiento
'           Set lrFunctVal = loFunct.GetConceptoValor(Val(gColPigConceptoCodComiVencida))
'           fnVarComVencida = Format(loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!NVALOR, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
'                                                                IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldo), "#0.00")
'        Else
'           fnVarComVencida = Format(0, "#0.00")
'        End If
'
'        If fnVarDiasVencidos > loFunct.GetParamValor(8024) Then        '***** Si se le cobrara Comision de Vencimiento
'           Set lrFunctVal = loFunct.GetConceptoValor(Val(8206))
'           fnVarDerRemate = Format(loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!NVALOR, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
'                                                                IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), fnVarSaldo), "#0.00")
'        Else
'           fnVarDerRemate = Format(0, "#0.00")
'        End If
'
'        lblProyDeuda.Caption = Format$(fnVarSaldo + fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComVencida + _
'                                                      fnVarPenalidad + fnVarComServicio + fnVarDerRemate, "###,##0.00 ")
'
'        lblProyPagMin.Caption = Format$(fnVarCapitalPagado + fnVarIntCompensatorio + fnVarIntMoratorio + _
'                                                       fnVarComVencida + fnVarComServicio + fnVarDerRemate, "###,##0.00 ")
'
'        dtpProyProxVcto.value = Format(DateAdd("d", Val(txtProyNuevoPlazo.Text), dtpPosicion.value), "dd/mm/yyyy")
'
'        Set loCalc = Nothing
'        Set loFunct = Nothing
'
'        If Val(Format(txtProyAmortizacion.Text, "#0.00")) >= Val(Format(lblProyPagMin.Caption, "#0.00")) Then
'            txtProyAmortizacion_KeyPress (13)
'        Else
'            txtProyAmortizacion.Text = Format$(0, "###,##0.00 ")
'            lblProyNuevoSaldo.Caption = Format$(0, "###,##0.00 ")
'            lblProyNuevaDeuda.Caption = Format$(0, "###,##0.00 ")
'            lblProyNuevoPagMin.Caption = Format$(0, "###,##0.00 ")
'        End If
'
'    End If
End Sub

Private Sub dtpProyProxVcto_Change()
    If DateDiff("d", dtpPosicion.value, dtpProyProxVcto.value) >= fnVarPlzMin And DateDiff("d", dtpPosicion.value, dtpProyProxVcto.value) <= fnVarPlzMax Then
        txtProyNuevoPlazo.Text = DateDiff("d", dtpPosicion.value, dtpProyProxVcto.value)
        txtProyAmortizacion.SetFocus
    Else
        dtpProyProxVcto.value = Format(DateAdd("d", Val(txtProyNuevoPlazo.Text), dtpPosicion.value), "dd/mm/yyyy")
        MsgBox "Plazo Fuera del Rango Permitido", vbInformation, " Aviso "
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
    AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXCodCta.Texto = "Credito"
    AXCodCta.Age = ""
    CargaParametros
End Sub

Private Sub lstCreditos_Click()
fsCtaCod = Trim(LstCreditos.SelectedItem)
AXCodCta.NroCuenta = fsCtaCod
MuestraDetalleContrato (fsCtaCod)
txtProyAmortizacion.Text = Format$(0, "###,##0.00 ")
lblProyNuevoSaldo.Caption = Format$(0, "###,##0.00 ")
lblProyNuevaDeuda.Caption = Format$(0, "###,##0.00 ")
lblProyNuevoPagMin.Caption = Format$(0, "###,##0.00 ")
End Sub

Private Sub lstDesembolsos_Click()
cmdVer.Visible = False
End Sub

Private Sub lstPagos_Click()
cmdVer.Visible = True
End Sub

Private Sub obEstado_Click(Index As Integer)
    Limpiar
    Select Case Index
          Case 0      '/****** ACTIVOS
                  fsEstados = gPigEstDesemb & "," & gPigEstReusoLin & "," & gPigEstRemat & "," & _
                                    gPigEstRematPRes & "," & gPigEstAmortiz
          Case 1      '/****** REMATADOS
                  fsEstados = gPigEstRematPFact & "," & gPigEstRematFact
          Case 2      '/****** ADJUDICADOS
                  fsEstados = gPigEstAdjud
          Case 3      '/****** CANCELADOS
                  fsEstados = gPigEstCancelPendRes & "," & gPigEstRescate
          Case 4      '/****** TODOS
                  fsEstados = gPigEstRegis & "," & gPigEstDesemb & "," & gPigEstReusoLin & "," & gPigEstRemat & "," & _
                                    gPigEstRematPRes & "," & gPigEstAmortiz & "," & gPigEstCancelPendRes & "," & _
                                    gPigEstRescate & "," & gPigEstRematPFact & "," & gPigEstRematFact & "," & _
                                    gPigEstAdjud
    End Select
    
    If MuestraContratos(fsCliente) Then MuestraDetalleContrato (fsCtaCod)

End Sub

Public Sub MuestraDetalleContrato(ByVal psCtaCod As String)
Dim lrDatos1 As ADODB.Recordset, lrDatos2 As ADODB.Recordset
Dim loValContrato As dPigContrato
    
Dim lrFunctVal As ADODB.Recordset
Dim loFunct As DPigFunciones
    
'Dim loCalc As NPigCalculos

Dim loMuestraContrato As DColPContrato

Set loMuestraContrato = New DColPContrato

'***********
    
    Set lrDatos1 = New ADODB.Recordset
    Set lrDatos2 = New ADODB.Recordset
    
    Set loValContrato = New dPigContrato
        'Set lrDatos = loValContrato.dObtieneDetalleCredPignoraticio(psCtaCod, fsEstados)
        
        Set lrDatos1 = loMuestraContrato.dObtieneDatosCreditoPignoraticio(psCtaCod)
     '   Set lrDatos2 = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyas(psCtaCod)
        Set lrDatos2 = loMuestraContrato.dObtieneTasasCreditoPignoraticio(psCtaCod)
    Set loValContrato = Nothing
     
     If lrDatos1 Is Nothing Or lrDatos2 Is Nothing Then
        Exit Sub
     Else
        lnEstadoCont = lrDatos1!nPrdEstado
        fnVarMontoCol = Format$(lrDatos1!nMontoCol, "##,##0.00 ")
        lblMontoPrestamo(0).Caption = Format$(lrDatos2!nMontoPrestamo, "##,##0.00 ")
        lblPlazo(1).Caption = Format$(lrDatos1!nPlazo, "#0")
        lblVencimiento(5).Caption = Format$(lrDatos1!dvenc, "dd/mm/yyyy")
        lblEmision(4).Caption = Format$(lrDatos1!dVigencia, "dd/mm/yyyy")
        'lblUltPago(6).Caption = Format$(lrDatos!dPrdEstado, "dd/mm/yyyy")
        'lblFecRescate(7).caption = ""
        'lblUsoLinea(8).Caption = Format$(lrDatos!nUsoLineaNro, "#0")
        lblDiasAtraso(9).Caption = DateDiff("d", Format(lrDatos1!dvenc, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
        lblDuplicados(11).Caption = Format$(lrDatos1!nNroDuplic, "#0")
        lblTasaInteres(12).Caption = Format$(lrDatos1!nTasaInteres, "#0.00 ")
        lblTasaMora(13).Caption = Format$(lrDatos2!nMora, "#0.00 ")
        'lblAvisoRemate(14).caption = ""
        fnVarSaldo = Format$(IIf(IsNull(lrDatos1!nSaldo), 0, lrDatos1!nSaldo), "##,##0.00 ")
        lblSaldoCapital(15).Caption = Format$(IIf(IsNull(lrDatos1!nSaldo), 0, lrDatos1!nSaldo), "##,##0.00 ")
        
        Dim loCalculos As NColPCalculos, fnVarInteresVencido As Double
        Set loCalculos = New NColPCalculos
      
        fnVarInteresVencido = loCalculos.nCalculaInteresMoratorio(fnVarSaldo, lrDatos2!nMora, Val(lblDiasAtraso(9).Caption))
        fnVarInteresVencido = Round(fnVarInteresVencido, 2)
        fnVarInteresVencido = IIf(fnVarInteresVencido < 0, 0, fnVarInteresVencido)
        LblMora(17).Caption = Format$(IIf(IsNull(fnVarInteresVencido), 0, fnVarInteresVencido), "##,##0.00 ")
        
   
        
      Dim fnVarFactor As Double, fnVarCostoCustodia As Double, fnVarInteres As Double, fnVarCostoCustodiaVencida As Double, fnVarImpuesto As Double
      Set loCalculos = New NColPCalculos
       
        fnVarFactor = loCalculos.nCalculaFactorRenovacion(Val(lblTasaInteres(12).Caption), lrDatos1!nPlazo)
    
        fnVarCostoCustodia = loCalculos.nCalculaCostoCustodia(lrDatos1!nTasacion, fnVarTasaCustodia, lrDatos1!nPlazo)
        fnVarCostoCustodiaVencida = loCalculos.nCalculaCostoCustodiaMoratorio(lrDatos1!nTasacion, fnVarTasaCustodiaVencida, Val(lblDiasAtraso(9).Caption))
        fnVarCostoCustodiaVencida = Round(fnVarCostoCustodiaVencida, 2)
    
    
        fnVarInteres = fnVarSaldo * fnVarFactor
    
        fnVarImpuesto = (fnVarInteresVencido + fnVarInteres + fnVarCostoCustodia + fnVarCostoCustodiaVencida) * fnVarTasaImpuesto
        
        'fnVarMontoMinimo = FNVARINTERESVENCIDO + fnVarCostoCustodiaVencida + fnVarInteres + fnVarCostoCustodia + fnVarImpuesto ' + fnVarCostoPreparacionRemate
    
       Set loCalculos = Nothing
        
        
        lblInteres(16).Caption = Format$(lrDatos1!nIntDIf, "##,##0.00 ")
        
'        Set lrFunctVal = loFunct.GetConceptoValor(Val(gColPigConceptoCodComiServ))
        
'        fnVarComServicio = Round(loCalc.nCalculaConcepto(lrFunctVal!nTpoValor, lrFunctVal!nValor, IIf(IsNull(lrFunctVal!nMontoMin), 0, lrFunctVal!nMontoMin), _
'                                                                    IIf(IsNull(lrFunctVal!nMontoMax), 0, lrFunctVal!nMontoMax), IIf(IsNull(lrDatos!nSaldo), 0, lrDatos!nSaldo)), 2) 'Comision para el nuevo Calendario

  '      fnVarComServicio = IIf(IsNull(lrDatos!nComisionServicio), 0, lrDatos!nComisionServicio)

  '      fnVarPorcCapMin = lrDatos!PorcCapitalMinimo
      '  fnVarCapitalPagado = Format((fnVarMontoCol * fnVarPorcCapMin / 100), "#,##0.00")
     '   If fnVarCapitalPagado > fnVarSaldo Then
    '       fnVarCapitalPagado = fnVarSaldo
     '   End If
    '    fnVarCapMinimo = Format(lrDatos!CapitalMinimo, "#0.00")
    '    lblOtros.Caption = Format$(IIf(IsNull(lrDatos!nDerechoRemate), 0, lrDatos!nDerechoRemate) + _
    '                                                IIf(IsNull(lrDatos!nComisionVenc), 0, lrDatos!nComisionVenc) + _
     '                                               fnVarPenalidad + fnVarComServicio, "##,##0.00 ")
        'lblTotal(19).Caption = Format$(IIf(IsNull(lrDatos!nSaldo), 0, lrDatos!nSaldo) + _
        '                                           IIf(IsNull(lrDatos!nDerechoRemate), 0, lrDatos!nDerechoRemate) + _
         '                                         IIf(IsNull(lrDatos!nComisionVenc), 0, lrDatos!nComisionVenc) + _
         '                                          IIf(IsNull(lrDatos!nInteresMora), 0, lrDatos!nInteresMora) + _
         '                                         fnVarPenalidad + fnVarComServicio + fnVarIntCompensatorio, "##,##0.00 ")
         
         
         LblTotal(19).Caption = Format$(IIf(IsNull(lrDatos1!nSaldo), 0, lrDatos1!nSaldo) + IIf(IsNull(fnVarInteresVencido), 0, fnVarInteresVencido) + fnVarIntCompensatorio, "##,##0.00 ")
         
         
         
         
         
         
         
         
      '  lblPagoMinimo(20).Caption = Format$(fnVarCapitalPagado + _
      '                                                      IIf(IsNull(lrDatos!nDerechoRemate), 0, lrDatos!nDerechoRemate) + _
       '                                                     IIf(IsNull(lrDatos!nComisionVenc), 0, lrDatos!nComisionVenc) + _
       '                                                     IIf(IsNull(lrDatos!nInteresMora), 0, lrDatos!nInteresMora) + _
        '                                                    fnVarComServicio + fnVarIntCompensatorio, "##,##0.00 ")
        fnVarPagoMinimo = Val(lblPagoMinimo(20).Caption)
        txtProyNuevoPlazo.Text = Format$(lrDatos1!nPlazo, "#0")
        
'        Set loCalc = Nothing
        Set loFunct = Nothing
        
        MuestraPiezasContrato (psCtaCod)
'        If lnEstadoCont = 2802 Or lnEstadoCont = 2803 Or lnEstadoCont = 2804 Then
'            dtpPosicion_Change
'        End If
        MuestraDesembolsosContrato (psCtaCod)
        MuestraPagosContrato (psCtaCod)
        
        SSTab1.Enabled = True
        Frame4.Enabled = True
        LstCreditos.Enabled = True
        cmdBuscar.Enabled = False
        cmdCancelar.Enabled = True
        
        If dtpPosicion.Enabled > True Then
            dtpPosicion.SetFocus
        End If
    End If

End Sub

Private Sub CargaParametros()
Dim loParam As DColPCalculos
Set loParam = New DColPCalculos

    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnVarTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
    fnVarDiasCambCart = loParam.dObtieneColocParametro(gConsColPDiasCambioCartera)
    fnVarTopRenovaciones = loParam.dObtieneColocParametro(gConsColPMaxNroRenovac)
    fnVarTopRenovacionesNuevo = loParam.dObtieneColocParametro(3057)
    
    
Set loParam = Nothing
End Sub




Public Sub MuestraPiezasContrato(ByVal psCtaCod As String)
Dim lrDatos As ADODB.Recordset
Dim loValContrato As dPigContrato
Dim lstTmpGarantias As ListItem
    
Dim lsTasador As String
Dim lnTotTas As Currency
Dim lnTotTasAdic As Currency
Dim lnPBruto As Currency
Dim lnPNeto As Currency
Dim lnTotPiezas As Integer

'***********
    If fnTipTasacion = 0 Then fnTipTasacion = 1
    Set lrDatos = New ADODB.Recordset
    Set loValContrato = New dPigContrato
        Set lrDatos = loValContrato.dObtieneDetalleJoyas(psCtaCod, fnTipTasacion)
    Set loValContrato = Nothing
     
     If lrDatos Is Nothing Then
        Exit Sub
    Else
        LstGarantias.ListItems.Clear
        
'        Do While Not lrDatos.EOF
'            Set lstTmpGarantias = lstGarantias.ListItems.Add(, , Trim(lrDatos!Item))
'            lstTmpGarantias.SubItems(1) = Trim(lrDatos!Material)
'            lstTmpGarantias.SubItems(2) = Trim(lrDatos!Tipo)
'            lstTmpGarantias.SubItems(3) = Trim(IIf(IsNull(lrDatos!SubTipo), "", lrDatos!SubTipo))
'            lstTmpGarantias.SubItems(4) = Trim(lrDatos!Estado)
'            lstTmpGarantias.SubItems(5) = Trim(IIf(IsNull(lrDatos!Observacion), "", lrDatos!Observacion))
'            lstTmpGarantias.SubItems(6) = Trim(IIf(IsNull(lrDatos!Situacion), "", lrDatos!Situacion))
'            lstTmpGarantias.SubItems(7) = Format$(lrDatos!PBruto, "#,##0.00 ")
'            lstTmpGarantias.SubItems(8) = Format$(lrDatos!pNeto, "#,##0.00 ")
'            lstTmpGarantias.SubItems(9) = Format$(lrDatos!Tasacion, "###,##0.00 ")
'            lstTmpGarantias.SubItems(10) = Format$(IIf(IsNull(lrDatos!TasAdicion), 0, lrDatos!TasAdicion), "###,##0.00 ")
'            lstTmpGarantias.SubItems(11) = Trim(IIf(IsNull(lrDatos!ObsAdicion), "", lrDatos!ObsAdicion))
'            lnTotTas = lnTotTas + lrDatos!Tasacion
'            lnTotTasAdic = lnTotTasAdic + IIf(IsNull(lrDatos!TasAdicion), 0, lrDatos!TasAdicion)
'            lnPBruto = lnPBruto + lrDatos!PBruto
'            lnPNeto = lnPNeto + lrDatos!pNeto
'            lnTotPiezas = lnTotPiezas + 1
'            lsTasador = lrDatos!Cuser
'            lrDatos.MoveNext
'        Loop

'Public Function fgMostrarJoyasDet(lstJoyasDet As ListView, ByVal prJoyas As ADODB.Recordset) As Boolean
Dim prJoyas As ADODB.Recordset
Dim loMuestraContrato As DColPContrato
Dim lstTmpCliente As ListItem

 lnTotTas = 0
 lnPBruto = 0
 lnPNeto = 0
 lnTotPiezas = 0


Set loMuestraContrato = New DColPContrato
    Set prJoyas = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(psCtaCod)

    If prJoyas.BOF And prJoyas.EOF Then
        MsgBox " Error al mostrar datos del Contrato ", vbCritical, " Aviso "
        Exit Sub
    Else
       LstGarantias.ListItems.Clear
        Do While Not prJoyas.EOF
            Set lstTmpCliente = LstGarantias.ListItems.Add(, , Trim(prJoyas!nItem))
                lstTmpCliente.SubItems(1) = ImpreFormat(prJoyas!npiezas, 3, 0)
                lstTmpCliente.SubItems(2) = Trim(prJoyas!ckilataje) & " k."
                lstTmpCliente.SubItems(3) = ImpreFormat(prJoyas!npesobruto, 4, 2)
                lstTmpCliente.SubItems(4) = ImpreFormat(prJoyas!npesoneto, 4, 2)
                lstTmpCliente.SubItems(5) = ImpreFormat(prJoyas!nvaltasac, 8, 2)
                lstTmpCliente.SubItems(6) = Trim(prJoyas!cdescrip)
                
                 lnTotTas = lnTotTas + prJoyas!nvaltasac
                 lnPBruto = lnPBruto + prJoyas!npesobruto
                 lnPNeto = lnPNeto + prJoyas!npesoneto
                 lnTotPiezas = lnTotPiezas + prJoyas!npiezas
                
            prJoyas.MoveNext
        Loop
        
        Set prJoyas = Nothing
        
    End If

        lblTasador.Caption = lsTasador
        lblTasacion.Caption = Format$(lnTotTas, "###,##0.00 ")
        lblTasAdic.Caption = Format$(lnTotTasAdic, "###,##0.00 ")
        LblPBruto.Caption = Format$(lnPBruto, "###,##0.00 ")
        LblPNeto.Caption = Format$(lnPNeto, "###,##0.00 ")
        lblPiezas.Caption = Format$(lnTotPiezas, "###,##0.00 ")
    End If

End Sub
'Private Sub optTasacion_Click(Index As Integer)
''    Select Case Index
''          Case 0
''                  fnTipTasacion = gPigTipoTasacNor
''          Case 1
''                  fnTipTasacion = gPigTipoTasacUsoLin
''          Case 2
''                  fnTipTasacion = gPigTipoTasacRetasac
''          Case 3
''                  fnTipTasacion = gPigTipoTasacRetasacVta
''    End Select
''
''    MuestraPiezasContrato (fsCtaCod)
'
'End Sub

Public Sub MuestraDesembolsosContrato(ByVal psCtaCod As String)
Dim lrDatos As ADODB.Recordset
Dim loValContrato As dPigContrato
Dim lstTmpDesembolsos As ListItem
    
Dim lnReg As Integer
'***********
    
    Set lrDatos = New ADODB.Recordset
    Set loValContrato = New dPigContrato
        Set lrDatos = loValContrato.dObtieneDesembolsosContrato(psCtaCod)
    Set loValContrato = Nothing
     
     If lrDatos Is Nothing Then
        Exit Sub
    Else
        lstDesembolsos.ListItems.Clear
        Do While Not lrDatos.EOF
            lnReg = lnReg + 1
            Set lstTmpDesembolsos = lstDesembolsos.ListItems.Add(, , lnReg)
                  lstTmpDesembolsos.SubItems(1) = Format$(Mid(Trim(lrDatos!cMovNro), 7, 2) & "/" & Mid(Trim(lrDatos!cMovNro), 5, 2) & "/" & Mid(Trim(lrDatos!cMovNro), 1, 4), "dd/mm/yyyy") 'Fecha
                  lstTmpDesembolsos.SubItems(2) = Mid(Trim(lrDatos!cMovNro), 9, 6)                                       'Hora
                  lstTmpDesembolsos.SubItems(3) = Format$(Trim(lrDatos!nMovNro), "#######0")
                  
                  lstTmpDesembolsos.SubItems(4) = Trim(lrDatos!DescAge)
                  lstTmpDesembolsos.SubItems(5) = Format$(lrDatos!capital, "###,##0.00")
                  lstTmpDesembolsos.SubItems(6) = Format$(lrDatos!ComTasacion, "#,##0.00 ")
                  lstTmpDesembolsos.SubItems(7) = Mid(Trim(lrDatos!cMovNro), 22, 4)                                    'Usuario
             lrDatos.MoveNext
        Loop
    End If

End Sub

Public Sub MuestraPagosContrato(ByVal psCtaCod As String)
Dim lrDatos As ADODB.Recordset
Dim loValContrato As dPigContrato
Dim lstTmpPagos As ListItem

Dim lnMontoPagado As Currency
Dim lnReg As Integer
    
'***********
    
    Set lrDatos = New ADODB.Recordset
    Set loValContrato = New dPigContrato
        Set lrDatos = loValContrato.dObtienePagosContrato(psCtaCod)
    Set loValContrato = Nothing
     
     If lrDatos Is Nothing Then
        Exit Sub
    Else
        lstPagos.ListItems.Clear
        Do While Not lrDatos.EOF
            lnReg = lnReg + 1
            Set lstTmpPagos = lstPagos.ListItems.Add(, , lnReg)
                  lstTmpPagos.SubItems(1) = Format$(Mid(Trim(lrDatos!cMovNro), 7, 2) & "/" & Mid(Trim(lrDatos!cMovNro), 5, 2) & "/" & Mid(Trim(lrDatos!cMovNro), 1, 4), "dd/mm/yyyy") 'Fecha
                  lstTmpPagos.SubItems(2) = Mid(Trim(lrDatos!cMovNro), 9, 2) & ":" & Mid(Trim(lrDatos!cMovNro), 11, 2) & ":" & Mid(Trim(lrDatos!cMovNro), 13, 2)   'Hora
                  lstTmpPagos.SubItems(3) = Format$(lrDatos!nMovNro, "#######0")
                  lstTmpPagos.SubItems(4) = Trim(lrDatos!cMovDesc)
                  lstTmpPagos.SubItems(5) = Trim(lrDatos!DescAge)
                  lnMontoPagado = Format$(lrDatos!capital + lrDatos!IntCompensatorio + lrDatos!IntMoratorio + _
                                                       lrDatos!ComServicio + lrDatos!ComVencida + lrDatos!DerRemate + _
                                                        lrDatos!CustDiferida, "###,##0.00")
                  lstTmpPagos.SubItems(6) = Format$(lnMontoPagado, "###,##0.00")
                  lstTmpPagos.SubItems(7) = Mid(Trim(lrDatos!cMovNro), 22, 4)                                      'Usuario
                  lstTmpPagos.SubItems(8) = Format$(lrDatos!capital, "###,##0.00")
                  lstTmpPagos.SubItems(9) = Format$(lrDatos!IntCompensatorio, "###,##0.00")
                  lstTmpPagos.SubItems(10) = Format$(lrDatos!IntMoratorio, "###,##0.00")
                  lstTmpPagos.SubItems(11) = Format$(lrDatos!ComVencida, "###,##0.00")
                  lstTmpPagos.SubItems(12) = Format$(lrDatos!ComServicio, "###,##0.00")
                  lstTmpPagos.SubItems(13) = Format$(lrDatos!DerRemate, "###,##0.00")
                  'lstTmpPagos.SubItems(14) = Format$(lrDatos!Penalidad, "###,##0.00")
                  lstTmpPagos.SubItems(14) = Format$(lrDatos!CustDiferida, "###,##0.00")
             lrDatos.MoveNext
        Loop
    End If
    cmdVer.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 cmdVer.Visible = False
End Sub

Private Sub txtProyAmortizacion_Click()
txtProyAmortizacion.SelLength = 12
End Sub

Private Sub txtProyAmortizacion_KeyPress(KeyAscii As Integer)
Dim loCalc As NPigCalculos
Dim lnNuevoSaldo As Currency
Dim lnProyAmort As Currency
Dim lnInteres As Currency
Dim lnMontoMinimo As Currency

   If KeyAscii = 13 Then
      lnProyAmort = Val(txtProyAmortizacion.Text)
      txtProyAmortizacion.Text = Format(txtProyAmortizacion.Text, "###,##0.00 ")
      lnMontoMinimo = fnVarCapitalPagado
      If lnProyAmort >= Val(Format(lblProyPagMin.Caption, "#0.00")) Then
         Set loCalc = New NPigCalculos
         lnNuevoSaldo = fnVarSaldo - (lnProyAmort - (fnVarIntCompensatorio + fnVarIntMoratorio + fnVarComServicio + _
                                                                         fnVarPenalidad + fnVarComVencida + fnVarDerRemate))
         lblProyNuevoSaldo.Caption = Format$(lnNuevoSaldo, "###,##0.00 ")
         lnInteres = Round(loCalc.nCalculaIntCompensatorio(lnNuevoSaldo, Val(lblTasaInteres(12).Caption), Val(txtProyNuevoPlazo.Text)), 2)
         If lnMontoMinimo > lnNuevoSaldo Then
            lnMontoMinimo = lnNuevoSaldo
         End If
         If lnMontoMinimo > 0 Then
             lblProyNuevaDeuda.Caption = Format$(lnNuevoSaldo + fnVarComServicio + lnInteres + fnVarPenalidad, "###,##0.00 ")
             lblProyNuevoPagMin.Caption = Format(lnMontoMinimo + fnVarComServicio + lnInteres, "###,##0.00 ")
         Else
             lblProyNuevaDeuda.Caption = Format$(0, "###,##0.00 ")
             lblProyNuevoPagMin.Caption = Format$(0, "###,##0.00 ")
         End If
      Else
         MsgBox "Monto a Pagar debe ser Mayor o Igual al Pago Minimo", vbInformation, " Aviso "
      End If
       txtProyAmortizacion.SetFocus
       txtProyAmortizacion.SelLength = 12
   End If
End Sub

Private Sub txtProyNuevoPlazo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If fnVarPlzMin <= txtProyNuevoPlazo.Text And txtProyNuevoPlazo.Text <= fnVarPlzMax Then
         dtpProyProxVcto.value = Format(DateAdd("d", Val(txtProyNuevoPlazo.Text), dtpPosicion.value), "dd/mm/yyyy")
         txtProyAmortizacion.SetFocus
      Else
         txtProyNuevoPlazo.Text = lblPlazo(1).Caption
         MsgBox "Plazo Fuera del Rango Permitido", vbInformation, " Aviso "
      End If
End If
End Sub

Public Sub MuestraContratoPosicion(ByVal psCtaCod As String, ByVal psCodCte As String)

    MuestraDatosPersona (psCodCte)
    Call MuestraContratos(psCodCte, psCtaCod)
    MuestraDetalleContrato (psCtaCod)
    dtpPosicion_Change
    fraContenedor(1).Enabled = False
    Frame4.Enabled = False
    cmdBuscar.Enabled = False
    LstCreditos.Enabled = False
    cmdCancelar.Enabled = False
    AXCodCta.Enabled = False
    AXCodCta.NroCuenta = psCtaCod
    Me.Show 1
    
End Sub

Public Sub MuestraDatosPersona(ByVal pscliente As String)
Dim lstTmpCliente As ListItem
Dim lrCredPigPersonas As ADODB.Recordset
Dim loMuestraDatos As dPigContrato

     ' Mostra Datos sobre el Cliente
     Set lrCredPigPersonas = New ADODB.Recordset
     Set loMuestraDatos = New dPigContrato
         Set lrCredPigPersonas = loMuestraDatos.dObtieneDatosClienteCredPignoraticio(pscliente)
     Set loMuestraDatos = Nothing
    
    If (lrCredPigPersonas.EOF And lrCredPigPersonas.BOF) Or lrCredPigPersonas Is Nothing Then
        MsgBox "Cliente no posee Contratos en Pignoraticios ", vbInformation, "Aviso"
        Exit Sub
    Else
        lstClientes.ListItems.Clear
        Do While Not lrCredPigPersonas.EOF
            Set lstTmpCliente = lstClientes.ListItems.Add(, , Trim(lrCredPigPersonas!cPerscod))
                  lstTmpCliente.SubItems(1) = Trim(PstaNombre(lrCredPigPersonas!cPersNombre, False))
                  lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrCredPigPersonas!NroDNI), "", lrCredPigPersonas!NroDNI))
                  lstTmpCliente.SubItems(3) = Trim(IIf(IsNull(lrCredPigPersonas!DescCalif), "", lrCredPigPersonas!DescCalif))
                  lstTmpCliente.SubItems(4) = Trim(IIf(IsNull(lrCredPigPersonas!DescSBS), "", lrCredPigPersonas!DescSBS))
                  fsCliente = Trim(lrCredPigPersonas!cPerscod)
            lrCredPigPersonas.MoveNext
        Loop
    End If
    
    Set lrCredPigPersonas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub


