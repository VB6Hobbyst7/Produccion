VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapTarifarioComision 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9375
   Icon            =   "frmCapTarifarioComision.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   300
      Left            =   8325
      TabIndex        =   196
      Top             =   7200
      Width           =   870
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   300
      Left            =   7380
      TabIndex        =   195
      Top             =   7200
      Width           =   870
   End
   Begin VB.CommandButton btnGuardarComo 
      Caption         =   "Guardar Como"
      Height          =   300
      Left            =   5085
      TabIndex        =   193
      Top             =   7200
      Width           =   1275
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "Editar"
      Height          =   300
      Left            =   1080
      TabIndex        =   192
      Top             =   7200
      Width           =   870
   End
   Begin VB.CommandButton btnNuevo 
      Caption         =   "Nuevo"
      Height          =   300
      Left            =   135
      TabIndex        =   191
      Top             =   7200
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Producto"
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
      Height          =   765
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   9180
      Begin VB.CommandButton btnExaminar 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   8010
         TabIndex        =   7
         Top             =   270
         Width           =   1050
      End
      Begin VB.CommandButton btnSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   300
         Left            =   6885
         TabIndex        =   6
         Top             =   270
         Width           =   1050
      End
      Begin VB.ComboBox cbPersoneria 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   2400
      End
      Begin VB.ComboBox cbProducto 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   2400
      End
      Begin VB.ComboBox cbGrupo 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Agencias:"
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   825
      End
   End
   Begin TabDlg.SSTab tabOperaciones 
      Height          =   5700
      Left            =   90
      TabIndex        =   0
      Top             =   945
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Operaciones en Cuenta"
      TabPicture(0)   =   "frmCapTarifarioComision.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Uso de ATM y Mon"
      TabPicture(1)   =   "frmCapTarifarioComision.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraVisa"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraExtr"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraOtro"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraMon"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraCaj"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Op en Ventanilla"
      TabPicture(2)   =   "frmCapTarifarioComision.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraVentanilla"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraExceso"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraRetiroOp"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraInterplaza"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraPlaza"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Serv. Asoc. Cta y Tarj"
      TabPicture(3)   =   "frmCapTarifarioComision.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraServicios"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraServicios 
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4920
         Left            =   -74775
         TabIndex        =   166
         Top             =   495
         Width           =   8700
         Begin VB.TextBox txtMantCuentaSolesSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3825
            TabIndex        =   181
            Text            =   "0"
            Top             =   540
            Width           =   780
         End
         Begin VB.ComboBox cbMantCuentaTipoSERV 
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   540
            Width           =   1365
         End
         Begin VB.TextBox txtMantCuentaDolaresSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6300
            TabIndex        =   179
            Text            =   "0"
            Top             =   540
            Width           =   780
         End
         Begin VB.TextBox txtDebitoRecaudoSolesSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3825
            TabIndex        =   178
            Text            =   "0"
            Top             =   1080
            Width           =   780
         End
         Begin VB.ComboBox cbDebitoRecaudoTipoSERV 
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   1080
            Width           =   1365
         End
         Begin VB.TextBox txtDebitoRecaudoDolaresSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6300
            TabIndex        =   176
            Text            =   "0"
            Top             =   1080
            Width           =   780
         End
         Begin VB.TextBox txtDebitoCreditoSolesSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3825
            TabIndex        =   175
            Text            =   "0"
            Top             =   1575
            Width           =   780
         End
         Begin VB.ComboBox cbDebitoCreditoTipoSERV 
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   1575
            Width           =   1365
         End
         Begin VB.TextBox txtDebitoCreditoDolaresSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6300
            TabIndex        =   173
            Text            =   "0"
            Top             =   1575
            Width           =   780
         End
         Begin VB.TextBox txtEnvioFisicoSolesSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3825
            TabIndex        =   172
            Text            =   "0"
            Top             =   2070
            Width           =   780
         End
         Begin VB.ComboBox cbEnvioFisicoTipoSERV 
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   2070
            Width           =   1365
         End
         Begin VB.TextBox txtEnvioFisicoDolaresSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6300
            TabIndex        =   170
            Text            =   "0"
            Top             =   2070
            Width           =   780
         End
         Begin VB.TextBox txtReposicionTarjetaSolesSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3825
            TabIndex        =   169
            Text            =   "0"
            Top             =   2565
            Width           =   780
         End
         Begin VB.ComboBox cbReposicionTarjetaTipoSERV 
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   2565
            Width           =   1365
         End
         Begin VB.TextBox txtReposicionTarjetaDolaresSERV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6300
            TabIndex        =   167
            Text            =   "0"
            Top             =   2565
            Width           =   780
         End
         Begin VB.Label Label77 
            Caption         =   "Comisión por Mantenimiento de cuenta"
            Height          =   285
            Left            =   270
            TabIndex        =   189
            Top             =   555
            Width           =   2760
         End
         Begin VB.Label Label78 
            Caption         =   "Soles:"
            Height          =   240
            Left            =   3825
            TabIndex        =   188
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label79 
            Caption         =   "Valor $:"
            Height          =   240
            Left            =   4770
            TabIndex        =   187
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label80 
            Caption         =   "Dolares"
            Height          =   240
            Left            =   6300
            TabIndex        =   186
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label81 
            Caption         =   "Débito para pagos de serv. de recaudo"
            Height          =   285
            Left            =   270
            TabIndex        =   185
            Top             =   1080
            Width           =   3030
         End
         Begin VB.Label Label82 
            Caption         =   "Débito para pagos de crédito"
            Height          =   285
            Left            =   270
            TabIndex        =   184
            Top             =   1620
            Width           =   2760
         End
         Begin VB.Label Label83 
            Caption         =   "Envío de estado de cuenta por medio físico:"
            Height          =   285
            Left            =   270
            TabIndex        =   183
            Top             =   2115
            Width           =   3345
         End
         Begin VB.Label Label84 
            Caption         =   "Comisión reposición de tarjeta débito Visa"
            Height          =   285
            Left            =   270
            TabIndex        =   182
            Top             =   2610
            Width           =   3345
         End
      End
      Begin VB.Frame fraVentanilla 
         Caption         =   "Otras Operaciones en Ventanilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   -74730
         TabIndex        =   144
         Top             =   3150
         Width           =   8655
         Begin VB.TextBox txtRetSinTarjDolaresVENT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6525
            TabIndex        =   159
            Text            =   "0"
            Top             =   1440
            Width           =   780
         End
         Begin VB.ComboBox cbRetSinTarjTipoVENT 
            Height          =   315
            Left            =   4995
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   1440
            Width           =   1365
         End
         Begin VB.TextBox txtRetSinTarjSolesVENT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4050
            TabIndex        =   157
            Text            =   "0"
            Top             =   1440
            Width           =   780
         End
         Begin VB.TextBox txtConsultaSaldoDolaresVENT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6525
            TabIndex        =   155
            Text            =   "0"
            Top             =   1035
            Width           =   780
         End
         Begin VB.ComboBox cbConsultaSaldoTipoVENT 
            Height          =   315
            Left            =   4995
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   1035
            Width           =   1365
         End
         Begin VB.TextBox txtConsultaSaldoSolesVENT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4050
            TabIndex        =   153
            Text            =   "0"
            Top             =   1035
            Width           =   780
         End
         Begin VB.TextBox txtExtractoCtasAhorroDolaresVENT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6525
            TabIndex        =   148
            Text            =   "0"
            Top             =   585
            Width           =   780
         End
         Begin VB.ComboBox cbExtractoCtasAhorroTipoVENT 
            Height          =   315
            Left            =   4995
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   585
            Width           =   1365
         End
         Begin VB.TextBox txtExtractoCtasAhorroSolesVENT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4050
            TabIndex        =   146
            Text            =   "0"
            Top             =   585
            Width           =   780
         End
         Begin VB.Label Label76 
            Caption         =   "Comisión por retiro sin tarjeta de Débito VISA"
            Height          =   285
            Left            =   405
            TabIndex        =   156
            Top             =   1440
            Width           =   3345
         End
         Begin VB.Label Label75 
            Caption         =   "Comisión por consulta de saldo en ventanilla"
            Height          =   285
            Left            =   405
            TabIndex        =   152
            Top             =   1035
            Width           =   3390
         End
         Begin VB.Label Label74 
            Caption         =   "Dolares"
            Height          =   240
            Left            =   6525
            TabIndex        =   151
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label73 
            Caption         =   "Valor $:"
            Height          =   240
            Left            =   4995
            TabIndex        =   150
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label72 
            Caption         =   "Soles:"
            Height          =   240
            Left            =   4050
            TabIndex        =   149
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label71 
            Caption         =   "Extracto de las cuentas de ahorros (por hoja)"
            Height          =   285
            Left            =   405
            TabIndex        =   145
            Top             =   585
            Width           =   3480
         End
      End
      Begin VB.Frame fraExceso 
         Caption         =   "Com. Por Exceso de Operaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74730
         TabIndex        =   128
         Top             =   2295
         Width           =   8655
         Begin VB.TextBox txtExcesoOpeTranDolaresVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7830
            TabIndex        =   141
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtExcesoOpeTranSolesVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6840
            TabIndex        =   140
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtExcesoOpeDepDolaresVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4950
            TabIndex        =   136
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtExcesoOpeDepSolesVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4005
            TabIndex        =   135
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtExcesoOpeRetDolaresVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2025
            TabIndex        =   131
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtExcesoOpeRetSolesVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1035
            TabIndex        =   130
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Label70 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7605
            TabIndex        =   143
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label69 
            Caption         =   "S/."
            Height          =   285
            Left            =   6525
            TabIndex        =   142
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label68 
            Caption         =   "Transf"
            Height          =   285
            Left            =   5895
            TabIndex        =   139
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label67 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4725
            TabIndex        =   138
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label66 
            Caption         =   "S/."
            Height          =   285
            Left            =   3645
            TabIndex        =   137
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label65 
            Caption         =   "Depósito:"
            Height          =   285
            Left            =   2880
            TabIndex        =   134
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label64 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   133
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label63 
            Caption         =   "S/."
            Height          =   285
            Left            =   675
            TabIndex        =   132
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label62 
            Caption         =   "Retiro:"
            Height          =   285
            Left            =   135
            TabIndex        =   129
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Frame fraRetiroOp 
         Caption         =   "Retiros con OP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -69510
         TabIndex        =   120
         Top             =   540
         Width           =   3435
         Begin VB.TextBox txtComExcesoOPdolaresVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   125
            Text            =   "0"
            Top             =   1170
            Width           =   645
         End
         Begin VB.TextBox txtComExcesoOPsolesVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   585
            TabIndex        =   124
            Text            =   "0"
            Top             =   1170
            Width           =   645
         End
         Begin VB.TextBox txtOpeLibresOPVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1395
            TabIndex        =   122
            Text            =   "0"
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label61 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   127
            Top             =   1170
            Width           =   195
         End
         Begin VB.Label Label60 
            Caption         =   "S/."
            Height          =   285
            Left            =   225
            TabIndex        =   126
            Top             =   1215
            Width           =   285
         End
         Begin VB.Label Label59 
            Caption         =   "Com. Exceso:"
            Height          =   285
            Left            =   225
            TabIndex        =   123
            Top             =   765
            Width           =   1140
         End
         Begin VB.Label Label58 
            Caption         =   "Nº Op. Libres:"
            Height          =   285
            Left            =   225
            TabIndex        =   121
            Top             =   360
            Width           =   1140
         End
      End
      Begin VB.Frame fraInterplaza 
         Caption         =   "Nº de Operaciones Libres Inter Plaza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74775
         TabIndex        =   119
         Top             =   1395
         Width           =   5100
         Begin VB.TextBox txtOpeLibresDepInterPlazaVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2385
            TabIndex        =   161
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtOpeLibresTransferInterPlazaVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4005
            TabIndex        =   160
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtOpeLibresRetInterPlazaVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   945
            TabIndex        =   162
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label57 
            Caption         =   "Retiros:"
            Height          =   285
            Left            =   315
            TabIndex        =   165
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label56 
            Caption         =   "Dep:"
            Height          =   285
            Left            =   1890
            TabIndex        =   164
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label55 
            Caption         =   "Transf:"
            Height          =   285
            Left            =   3375
            TabIndex        =   163
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame fraPlaza 
         Caption         =   "Nº de Operaciones Libres de la Misma Plaza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74730
         TabIndex        =   112
         Top             =   540
         Width           =   5100
         Begin VB.TextBox txtOpeLibresTransferVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4005
            TabIndex        =   117
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtOpeLibresDepVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2385
            TabIndex        =   115
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.TextBox txtOpeLibresRetirosVent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   945
            TabIndex        =   113
            Text            =   "0"
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label54 
            Caption         =   "Transf:"
            Height          =   285
            Left            =   3375
            TabIndex        =   118
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label52 
            Caption         =   "Dep:"
            Height          =   285
            Left            =   1890
            TabIndex        =   116
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label53 
            Caption         =   "Retiros:"
            Height          =   285
            Left            =   315
            TabIndex        =   114
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame fraVisa 
         Caption         =   "Compras con Visa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74820
         TabIndex        =   51
         Top             =   4725
         Width           =   8790
         Begin VB.TextBox txtComprasDolaresInterNacionalesVISA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7515
            TabIndex        =   109
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtComprasSolesInterNacionalesVISA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   108
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtComprasDolaresNacionalesVISA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   104
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox txtComprasSolesNacionalesVISA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   103
            Text            =   "0"
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Label51 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7290
            TabIndex        =   111
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label50 
            Caption         =   "S/."
            Height          =   285
            Left            =   6120
            TabIndex        =   110
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label49 
            Caption         =   "Nacionales:"
            Height          =   285
            Left            =   4545
            TabIndex        =   107
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label48 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   106
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label47 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   105
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label46 
            Caption         =   "Nacionales:"
            Height          =   285
            Left            =   225
            TabIndex        =   102
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame fraExtr 
         Caption         =   "Op. En ATM del Extranjero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -70455
         TabIndex        =   50
         Top             =   3060
         Width           =   4425
         Begin VB.TextBox txtCambioClaveDolaresExtranj 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   99
            Text            =   "0"
            Top             =   1080
            Width           =   645
         End
         Begin VB.TextBox txtCambioClaveSolesExtranj 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   98
            Text            =   "0"
            Top             =   1080
            Width           =   645
         End
         Begin VB.TextBox txtConsultaDolaresExtranj 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   94
            Text            =   "0"
            Top             =   675
            Width           =   645
         End
         Begin VB.TextBox txtConsultaSolesExtranj 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   93
            Text            =   "0"
            Top             =   675
            Width           =   645
         End
         Begin VB.TextBox txtRetDolaresExtranj 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   90
            Text            =   "0"
            Top             =   270
            Width           =   645
         End
         Begin VB.TextBox txtRetSolesExtranj 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   89
            Text            =   "0"
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label45 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   101
            Top             =   1080
            Width           =   195
         End
         Begin VB.Label Label44 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   100
            Top             =   1125
            Width           =   285
         End
         Begin VB.Label Label43 
            Caption         =   "Cambio de Clave:"
            Height          =   285
            Left            =   225
            TabIndex        =   97
            Top             =   1125
            Width           =   1320
         End
         Begin VB.Label Label42 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   96
            Top             =   675
            Width           =   195
         End
         Begin VB.Label Label41 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   95
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label40 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   92
            Top             =   270
            Width           =   195
         End
         Begin VB.Label Label39 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   91
            Top             =   315
            Width           =   285
         End
         Begin VB.Label Label38 
            Caption         =   "Consultas:"
            Height          =   285
            Left            =   225
            TabIndex        =   88
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label37 
            Caption         =   "Retiro:"
            Height          =   285
            Left            =   225
            TabIndex        =   87
            Top             =   300
            Width           =   870
         End
      End
      Begin VB.Frame fraOtro 
         Caption         =   "Op. En ATMs diferente a GlobalNet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -74820
         TabIndex        =   49
         Top             =   3060
         Width           =   4245
         Begin VB.TextBox txtConsultaDolaresOTRO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   84
            Text            =   "0"
            Top             =   765
            Width           =   645
         End
         Begin VB.TextBox txtConsultaSolesOTRO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   83
            Text            =   "0"
            Top             =   765
            Width           =   645
         End
         Begin VB.TextBox txtRetDolaresOTRO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   80
            Text            =   "0"
            Top             =   360
            Width           =   645
         End
         Begin VB.TextBox txtRetSolesOTRO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   79
            Text            =   "0"
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label36 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   86
            Top             =   765
            Width           =   195
         End
         Begin VB.Label Label35 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   85
            Top             =   810
            Width           =   285
         End
         Begin VB.Label Label34 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   82
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Label33 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   81
            Top             =   405
            Width           =   285
         End
         Begin VB.Label Label32 
            Caption         =   "Consultas:"
            Height          =   285
            Left            =   225
            TabIndex        =   78
            Top             =   810
            Width           =   870
         End
         Begin VB.Label Label31 
            Caption         =   "Retiro:"
            Height          =   285
            Left            =   225
            TabIndex        =   77
            Top             =   390
            Width           =   870
         End
      End
      Begin VB.Frame fraMon 
         Caption         =   "Uso de Monederos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -70455
         TabIndex        =   48
         Top             =   450
         Width           =   4425
         Begin VB.TextBox txtOpeLibresMON 
            Height          =   285
            Left            =   3015
            TabIndex        =   75
            Text            =   "0.9"
            Top             =   225
            Width           =   645
         End
         Begin SICMACT.FlexEdit grdExceoOperaciones 
            Height          =   1455
            Left            =   180
            TabIndex        =   198
            Top             =   900
            Width           =   4020
            _extentx        =   7091
            _extenty        =   2566
            cols0           =   4
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "col-nId-Concepto Moneda-Valor"
            encabezadosanchos=   "0-0-2300-1000"
            font            =   "frmCapTarifarioComision.frx":037A
            font            =   "frmCapTarifarioComision.frx":03A6
            font            =   "frmCapTarifarioComision.frx":03D2
            font            =   "frmCapTarifarioComision.frx":03FE
            font            =   "frmCapTarifarioComision.frx":042A
            fontfixed       =   "frmCapTarifarioComision.frx":0456
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-3"
            listacontroles  =   "0-0-0-0"
            encabezadosalineacion=   "C-C-C-R"
            formatosedit    =   "0-0-0-2"
            textarray0      =   "col"
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label Label30 
            Caption         =   "Com. Por exceso de Operaciones"
            Height          =   285
            Left            =   180
            TabIndex        =   76
            Top             =   585
            Width           =   2850
         End
         Begin VB.Label Label29 
            Caption         =   "Nº de Operaciones libres"
            Height          =   285
            Left            =   1170
            TabIndex        =   74
            Top             =   270
            Width           =   1905
         End
      End
      Begin VB.Frame fraCaj 
         Caption         =   "Uso Cajero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -74820
         TabIndex        =   47
         Top             =   450
         Width           =   4245
         Begin VB.TextBox txtCambioClaveDolaresExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   71
            Text            =   "0"
            Top             =   2025
            Width           =   645
         End
         Begin VB.TextBox txtCambioClaveSolesExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   70
            Text            =   "0"
            Top             =   2025
            Width           =   645
         End
         Begin VB.TextBox txtConsultaMovDolaresExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   67
            Text            =   "0"
            Top             =   1665
            Width           =   645
         End
         Begin VB.TextBox txtConsultaMovSolesExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   66
            Text            =   "0"
            Top             =   1665
            Width           =   645
         End
         Begin VB.TextBox txtConsultaSaldosDolaresExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   63
            Text            =   "0"
            Top             =   1305
            Width           =   645
         End
         Begin VB.TextBox txtConsultaSaldosSolesExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   62
            Text            =   "0"
            Top             =   1305
            Width           =   645
         End
         Begin VB.TextBox txtOpeLibresATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2925
            TabIndex        =   59
            Text            =   "0"
            Top             =   225
            Width           =   645
         End
         Begin VB.TextBox txtRetDolaresExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   58
            Text            =   "0"
            Top             =   945
            Width           =   645
         End
         Begin VB.TextBox txtRetSolesExcesoOpeATM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   57
            Text            =   "0"
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label85 
            Caption         =   "Com. Por exceso de Operaciones:"
            Height          =   240
            Left            =   225
            TabIndex        =   197
            Top             =   585
            Width           =   3075
         End
         Begin VB.Label Label28 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   73
            Top             =   2025
            Width           =   195
         End
         Begin VB.Label Label27 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   72
            Top             =   2070
            Width           =   285
         End
         Begin VB.Label Label26 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   69
            Top             =   1665
            Width           =   195
         End
         Begin VB.Label Label25 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   68
            Top             =   1710
            Width           =   285
         End
         Begin VB.Label Label24 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   65
            Top             =   1305
            Width           =   195
         End
         Begin VB.Label Label23 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   64
            Top             =   1350
            Width           =   285
         End
         Begin VB.Label Label22 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3150
            TabIndex        =   61
            Top             =   945
            Width           =   195
         End
         Begin VB.Label Label21 
            Caption         =   "S/."
            Height          =   285
            Left            =   1980
            TabIndex        =   60
            Top             =   990
            Width           =   285
         End
         Begin VB.Label Label20 
            Caption         =   "Cambio de clave:"
            Height          =   285
            Left            =   225
            TabIndex        =   56
            Top             =   2070
            Width           =   1500
         End
         Begin VB.Label Label19 
            Caption         =   "Consulta de Saldos:"
            Height          =   285
            Left            =   225
            TabIndex        =   55
            Top             =   1305
            Width           =   1500
         End
         Begin VB.Label Label18 
            Caption         =   "Consulta de Mov.:"
            Height          =   285
            Left            =   225
            TabIndex        =   54
            Top             =   1710
            Width           =   1500
         End
         Begin VB.Label Label17 
            Caption         =   "Retiro:"
            Height          =   285
            Left            =   225
            TabIndex        =   53
            Top             =   990
            Width           =   1500
         End
         Begin VB.Label Label16 
            Caption         =   "Nº de Operaciones libres"
            Height          =   285
            Left            =   990
            TabIndex        =   52
            Top             =   260
            Width           =   1905
         End
      End
      Begin VB.Frame fraCta 
         Caption         =   "Operaciones en otra localidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5010
         Left            =   190
         TabIndex        =   8
         Top             =   495
         Width           =   8745
         Begin VB.TextBox txtTranDolaresMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6255
            TabIndex        =   46
            Text            =   "0"
            Top             =   3600
            Width           =   645
         End
         Begin VB.TextBox txtTranSolesMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6255
            TabIndex        =   45
            Text            =   "0"
            Top             =   3195
            Width           =   645
         End
         Begin VB.TextBox txtTranDolaresMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5445
            TabIndex        =   44
            Text            =   "0"
            Top             =   3600
            Width           =   645
         End
         Begin VB.TextBox txtTranSolesMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5445
            TabIndex        =   43
            Text            =   "0"
            Top             =   3195
            Width           =   645
         End
         Begin VB.TextBox txtTranDolaresComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4635
            TabIndex        =   42
            Text            =   "0"
            Top             =   3600
            Width           =   645
         End
         Begin VB.TextBox txtTranSolesComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4635
            TabIndex        =   41
            Text            =   "0"
            Top             =   3195
            Width           =   645
         End
         Begin VB.ComboBox cbTranDolaresTipo 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   3600
            Width           =   1590
         End
         Begin VB.ComboBox cbTranSolesTipo 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   3195
            Width           =   1590
         End
         Begin VB.TextBox txtRetDolaresMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6255
            TabIndex        =   35
            Text            =   "0"
            Top             =   2385
            Width           =   645
         End
         Begin VB.TextBox txtRetSolesMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6255
            TabIndex        =   34
            Text            =   "0"
            Top             =   1980
            Width           =   645
         End
         Begin VB.TextBox txtRetDolaresMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5445
            TabIndex        =   33
            Text            =   "0"
            Top             =   2385
            Width           =   645
         End
         Begin VB.TextBox txtRetSolesMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5445
            TabIndex        =   32
            Text            =   "0"
            Top             =   1980
            Width           =   645
         End
         Begin VB.TextBox txtRetDolaresComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4635
            TabIndex        =   31
            Text            =   "0"
            Top             =   2385
            Width           =   645
         End
         Begin VB.TextBox txtRetSolesComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4635
            TabIndex        =   30
            Text            =   "0"
            Top             =   1980
            Width           =   645
         End
         Begin VB.ComboBox cbRetDolaresTipo 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2385
            Width           =   1590
         End
         Begin VB.ComboBox cbRetSolesTipo 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1980
            Width           =   1590
         End
         Begin VB.TextBox txtDepDolaresMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6255
            TabIndex        =   23
            Text            =   "0"
            Top             =   1125
            Width           =   645
         End
         Begin VB.TextBox txtDepSolesMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6255
            TabIndex        =   22
            Text            =   "0"
            Top             =   720
            Width           =   645
         End
         Begin VB.TextBox txtDepSolesMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5445
            TabIndex        =   20
            Text            =   "0"
            Top             =   720
            Width           =   645
         End
         Begin VB.TextBox txtDepDolaresComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4635
            TabIndex        =   18
            Text            =   "0"
            Top             =   1125
            Width           =   645
         End
         Begin VB.TextBox txtDepSolesComision 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4635
            TabIndex        =   17
            Text            =   "0"
            Top             =   720
            Width           =   645
         End
         Begin VB.ComboBox cbDepDolaresTipo 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1125
            Width           =   1590
         End
         Begin VB.ComboBox cbDepSolesTipo 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   720
            Width           =   1590
         End
         Begin VB.TextBox txtDepDolaresMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5445
            TabIndex        =   21
            Text            =   "0"
            Top             =   1125
            Width           =   645
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dolares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1755
            TabIndex        =   39
            Top             =   3600
            Width           =   960
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Soles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1755
            TabIndex        =   37
            Top             =   3195
            Width           =   960
         End
         Begin VB.Label Label13 
            Caption         =   "Transferencia entre cuentas:"
            Height          =   420
            Left            =   450
            TabIndex        =   36
            Top             =   3195
            Width           =   1230
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dolares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1755
            TabIndex        =   28
            Top             =   2385
            Width           =   960
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Soles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1755
            TabIndex        =   26
            Top             =   1980
            Width           =   960
         End
         Begin VB.Label Label10 
            Caption         =   "Retiros"
            Height          =   240
            Left            =   450
            TabIndex        =   25
            Top             =   1980
            Width           =   825
         End
         Begin VB.Label Label9 
            Caption         =   "Max"
            Height          =   240
            Left            =   6345
            TabIndex        =   24
            Top             =   405
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "Min:"
            Height          =   240
            Left            =   5580
            TabIndex        =   19
            Top             =   405
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Valor"
            Height          =   240
            Left            =   4725
            TabIndex        =   16
            Top             =   405
            Width           =   555
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo"
            Height          =   240
            Left            =   3375
            TabIndex        =   15
            Top             =   405
            Width           =   555
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dolares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1755
            TabIndex        =   13
            Top             =   1125
            Width           =   960
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Soles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1755
            TabIndex        =   11
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label3 
            Caption         =   "Moneda"
            Height          =   240
            Left            =   1890
            TabIndex        =   10
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Depositos:"
            Height          =   240
            Left            =   450
            TabIndex        =   9
            Top             =   720
            Width           =   825
         End
      End
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   300
      Left            =   6435
      TabIndex        =   194
      Top             =   7200
      Width           =   870
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version Seleccionada :"
      Height          =   240
      Left            =   180
      TabIndex        =   190
      Top             =   6795
      Width           =   2400
   End
End
Attribute VB_Name = "frmCapTarifarioComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioComision
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Creacion y Modificacion de Comisiones
'* CREACION       : RIRO, 20160420 10:00 AM
'************************************************************************************************************

Option Explicit

Private oCon As COMNCaptaGenerales.NCOMCaptaDefinicion
Private oGuardar As New frmCapTarifarioGuardar
Private nIdComision As Integer ' ID de la comision seleccionada
Private bFocoGrid As Boolean 'ver si el flexEdit tiene el foco
Private nTipoOperacion As Integer '0=Sin definir, 1=Nuevo, 2=Edicion

Private Sub CargarControles()
    
Dim oConstante As COMDConstSistema.DCOMGeneral
Dim oCon As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsTmp As ADODB.Recordset

Set oConstante = New COMDConstSistema.DCOMGeneral
Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion

'cargando los grupos
Set rsTmp = oCon.ObtenerGruposComision(2)
If Not rsTmp Is Nothing Then
    If Not rsTmp.EOF And Not rsTmp.BOF Then
        CargaCombo cbGrupo, rsTmp
        Set rsTmp = Nothing
    Else
        cbGrupo.Clear
    End If
Else
    cbGrupo.Clear
End If

'cargando los subproductos de ahorros
Set rsTmp = oConstante.GetConstante(2030, , "", "-")
CargaCombo cbProducto, rsTmp
Set rsTmp = Nothing

'cargando las personerias
Set rsTmp = oConstante.GetConstante(1002, , "'[12]'", "-")
CargaCombo cbPersoneria, rsTmp
Set rsTmp = Nothing

'cargando los combos de tipos.
cbDepSolesTipo.AddItem "Valor" & Space(50) & "1"
cbDepSolesTipo.AddItem "Porcentaje" & Space(50) & "2"

cbDepDolaresTipo.AddItem "Valor" & Space(50) & "1"
cbDepDolaresTipo.AddItem "Porcentaje" & Space(50) & "2"
cbDepDolaresTipo.AddItem "Equivalente" & Space(50) & "3"

'**

cbRetSolesTipo.AddItem "Valor" & Space(50) & "1"
cbRetSolesTipo.AddItem "Porcentaje" & Space(50) & "2"

cbRetDolaresTipo.AddItem "Valor" & Space(50) & "1"
cbRetDolaresTipo.AddItem "Porcentaje" & Space(50) & "2"
cbRetDolaresTipo.AddItem "Equivalente" & Space(50) & "3"

'***

cbTranSolesTipo.AddItem "Valor" & Space(50) & "1"
cbTranSolesTipo.AddItem "Porcentaje" & Space(50) & "2"

cbTranDolaresTipo.AddItem "Valor" & Space(50) & "1"
cbTranDolaresTipo.AddItem "Porcentaje" & Space(50) & "2"
cbTranDolaresTipo.AddItem "Equivalente" & Space(50) & "3"

'***

cbExtractoCtasAhorroTipoVENT.AddItem "Valor" & Space(50) & "1"
cbExtractoCtasAhorroTipoVENT.AddItem "Equivalente" & Space(50) & "2"
cbExtractoCtasAhorroTipoVENT.AddItem "N/A" & Space(50) & "3"

'***

cbConsultaSaldoTipoVENT.AddItem "Valor" & Space(50) & "1"
cbConsultaSaldoTipoVENT.AddItem "Equivalente" & Space(50) & "2"
cbConsultaSaldoTipoVENT.AddItem "N/A" & Space(50) & "3"

'***

cbRetSinTarjTipoVENT.AddItem "Valor" & Space(50) & "1"
cbRetSinTarjTipoVENT.AddItem "Equivalente" & Space(50) & "2"
cbRetSinTarjTipoVENT.AddItem "N/A" & Space(50) & "3"

'***

cbMantCuentaTipoSERV.AddItem "Valor" & Space(50) & "1"
cbMantCuentaTipoSERV.AddItem "Equivalente" & Space(50) & "2"
cbMantCuentaTipoSERV.AddItem "N/A" & Space(50) & "3"

'***

cbDebitoRecaudoTipoSERV.AddItem "Valor" & Space(50) & "1"
cbDebitoRecaudoTipoSERV.AddItem "Equivalente" & Space(50) & "2"
cbDebitoRecaudoTipoSERV.AddItem "N/A" & Space(50) & "3"

'***

cbDebitoCreditoTipoSERV.AddItem "Valor" & Space(50) & "1"
cbDebitoCreditoTipoSERV.AddItem "Equivalente" & Space(50) & "2"
cbDebitoCreditoTipoSERV.AddItem "N/A" & Space(50) & "3"

'***

cbEnvioFisicoTipoSERV.AddItem "Valor" & Space(50) & "1"
cbEnvioFisicoTipoSERV.AddItem "Equivalente" & Space(50) & "2"
cbEnvioFisicoTipoSERV.AddItem "N/A" & Space(50) & "3"

'***

cbReposicionTarjetaTipoSERV.AddItem "Valor" & Space(50) & "1"
cbReposicionTarjetaTipoSERV.AddItem "Equivalente" & Space(50) & "2"
cbReposicionTarjetaTipoSERV.AddItem "N/A" & Space(50) & "3"

'cargando Grid Monedero

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(1, 0) = 1
grdExceoOperaciones.TextMatrix(1, 1) = 5
grdExceoOperaciones.TextMatrix(1, 2) = "Retiro S/. 5.00"
grdExceoOperaciones.TextMatrix(1, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(2, 0) = 2
grdExceoOperaciones.TextMatrix(2, 1) = 10
grdExceoOperaciones.TextMatrix(2, 2) = "Retiro S/. 10.00"
grdExceoOperaciones.TextMatrix(2, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(3, 0) = 3
grdExceoOperaciones.TextMatrix(3, 1) = 15
grdExceoOperaciones.TextMatrix(3, 2) = "Retiro S/. 15.00"
grdExceoOperaciones.TextMatrix(3, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(4, 0) = 4
grdExceoOperaciones.TextMatrix(4, 1) = 20
grdExceoOperaciones.TextMatrix(4, 2) = "Retiro S/. 20.00"
grdExceoOperaciones.TextMatrix(4, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(5, 0) = 5
grdExceoOperaciones.TextMatrix(5, 1) = 25
grdExceoOperaciones.TextMatrix(5, 2) = "Retiro S/. 25.00"
grdExceoOperaciones.TextMatrix(5, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(6, 0) = 6
grdExceoOperaciones.TextMatrix(6, 1) = 30
grdExceoOperaciones.TextMatrix(6, 2) = "Retiro S/. 30.00"
grdExceoOperaciones.TextMatrix(6, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(7, 0) = 7
grdExceoOperaciones.TextMatrix(7, 1) = 35
grdExceoOperaciones.TextMatrix(7, 2) = "Retiro S/. 35.00"
grdExceoOperaciones.TextMatrix(7, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(8, 0) = 8
grdExceoOperaciones.TextMatrix(8, 1) = 40
grdExceoOperaciones.TextMatrix(8, 2) = "Retiro S/. 40.00"
grdExceoOperaciones.TextMatrix(8, 3) = "0.00"

grdExceoOperaciones.AdicionaFila
grdExceoOperaciones.TextMatrix(9, 0) = 8
grdExceoOperaciones.TextMatrix(9, 1) = 45
grdExceoOperaciones.TextMatrix(9, 2) = "Retiro S/. 45.00"
grdExceoOperaciones.TextMatrix(9, 3) = "0.00"

'seleccionando el primer registro de los combos
If cbGrupo.ListCount > 0 Then cbGrupo.ListIndex = 0
If cbProducto.ListCount > 0 Then cbProducto.ListIndex = 0
If cbPersoneria.ListCount > 0 Then cbPersoneria.ListIndex = 0
If cbDepSolesTipo.ListCount > 0 Then cbDepSolesTipo.ListIndex = 0
If cbDepDolaresTipo.ListCount > 0 Then cbDepDolaresTipo.ListIndex = 0
If cbRetSolesTipo.ListCount > 0 Then cbRetSolesTipo.ListIndex = 0
If cbRetDolaresTipo.ListCount > 0 Then cbRetDolaresTipo.ListIndex = 0
If cbTranSolesTipo.ListCount > 0 Then cbTranSolesTipo.ListIndex = 0
If cbTranDolaresTipo.ListCount > 0 Then cbTranDolaresTipo.ListIndex = 0
If cbExtractoCtasAhorroTipoVENT.ListCount > 0 Then cbExtractoCtasAhorroTipoVENT.ListIndex = 0
If cbConsultaSaldoTipoVENT.ListCount > 0 Then cbConsultaSaldoTipoVENT.ListIndex = 0
If cbRetSinTarjTipoVENT.ListCount > 0 Then cbRetSinTarjTipoVENT.ListIndex = 0
If cbMantCuentaTipoSERV.ListCount > 0 Then cbMantCuentaTipoSERV.ListIndex = 0
If cbDebitoRecaudoTipoSERV.ListCount > 0 Then cbDebitoRecaudoTipoSERV.ListIndex = 0
If cbDebitoCreditoTipoSERV.ListCount > 0 Then cbDebitoCreditoTipoSERV.ListIndex = 0
If cbEnvioFisicoTipoSERV.ListCount > 0 Then cbEnvioFisicoTipoSERV.ListIndex = 0
If cbTranDolaresTipo.ListCount > 0 Then cbReposicionTarjetaTipoSERV.ListIndex = 0

End Sub
Private Sub Limpiar(Optional nTipo As Integer = -1)
Dim Control As Control
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        Control.Text = "0"
    End If
    If TypeOf Control Is ComboBox Then
        If (nTipo = 1 Or nTipo = 2) And (Control.Name = "cbGrupo" Or Control.Name = "cbProducto" Or Control.Name = "cbPersoneria") Then
             
        Else
            Control.ListIndex = 0
        End If
    End If
Next
grdExceoOperaciones.TextMatrix(1, 3) = "0.00"
grdExceoOperaciones.TextMatrix(2, 3) = "0.00"
grdExceoOperaciones.TextMatrix(3, 3) = "0.00"
grdExceoOperaciones.TextMatrix(4, 3) = "0.00"
grdExceoOperaciones.TextMatrix(5, 3) = "0.00"
grdExceoOperaciones.TextMatrix(6, 3) = "0.00"
grdExceoOperaciones.TextMatrix(7, 3) = "0.00"
grdExceoOperaciones.TextMatrix(8, 3) = "0.00"
grdExceoOperaciones.TextMatrix(9, 3) = "0.00"
nTipoOperacion = IIf(nTipo < 0, 0, nTipoOperacion)
nIdComision = -1
tabOperaciones.Tab = 0
End Sub
Private Function Validar() As Boolean
Dim bValida As Boolean
Dim sMensaje As String
Dim Control As Control
bValida = True
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        If Not IsNumeric(Control.Text) Then
            bValida = False
        Else
           If CDbl(Control.Text) < 0 Then
            bValida = False
           End If
        End If
    End If
    If TypeOf Control Is ComboBox Then
        If Control.ListIndex < 0 Then
            bValida = False
        End If
    End If
Next
Validar = bValida
End Function
Private Sub BlqControles(ByVal nTipoBloqueo As Integer)

    'bloqueo al Inicio
    If nTipoBloqueo = 1 Then
        
        cbGrupo.Enabled = True
        cbProducto.Enabled = True
        cbPersoneria.Enabled = True
        btnSeleccionar.Enabled = True
        btnExaminar.Enabled = False
        btnNuevo.Enabled = False
        btnEditar.Enabled = False
        btnGuardarComo.Enabled = False
        btnGuardar.Enabled = False
        btnCancelar.Enabled = False
        btnSalir.Enabled = True
        
        fraCaj.Enabled = False
        fraCta.Enabled = False
        fraExceso.Enabled = False
        fraExtr.Enabled = False
        fraInterplaza.Enabled = False
        fraMon.Enabled = False
        fraOtro.Enabled = False
        fraPlaza.Enabled = False
        fraRetiroOp.Enabled = False
        fraServicios.Enabled = False
        fraVentanilla.Enabled = False
        fraVisa.Enabled = False
        
    'bloqueo al seleccionar
    ElseIf nTipoBloqueo = 2 Then
    
        cbGrupo.Enabled = False
        cbProducto.Enabled = False
        cbPersoneria.Enabled = False
        btnSeleccionar.Enabled = False
        btnExaminar.Enabled = True
        btnNuevo.Enabled = True
        btnEditar.Enabled = False
        btnGuardarComo.Enabled = False
        btnGuardar.Enabled = False
        btnCancelar.Enabled = True
        btnSalir.Enabled = True
        
        fraCaj.Enabled = False
        fraCta.Enabled = False
        fraExceso.Enabled = False
        fraExtr.Enabled = False
        fraInterplaza.Enabled = False
        fraMon.Enabled = False
        fraOtro.Enabled = False
        fraPlaza.Enabled = False
        fraRetiroOp.Enabled = False
        fraServicios.Enabled = False
        fraVentanilla.Enabled = False
        fraVisa.Enabled = False
    
    'bloqueo nuevo
    ElseIf nTipoBloqueo = 3 Then
    
        cbGrupo.Enabled = False
        cbProducto.Enabled = False
        cbPersoneria.Enabled = False
        btnSeleccionar.Enabled = False
        btnExaminar.Enabled = False
        btnNuevo.Enabled = False
        btnEditar.Enabled = False
        btnGuardarComo.Enabled = False
        btnGuardar.Enabled = True
        btnCancelar.Enabled = True
        btnSalir.Enabled = True
    
        fraCaj.Enabled = True
        fraCta.Enabled = True
        fraExceso.Enabled = True
        fraExtr.Enabled = True
        fraInterplaza.Enabled = True
        fraMon.Enabled = True
        fraOtro.Enabled = True
        fraPlaza.Enabled = True
        fraRetiroOp.Enabled = True
        fraServicios.Enabled = True
        fraVentanilla.Enabled = True
        fraVisa.Enabled = True
    
    'bloqueo examinar
    ElseIf nTipoBloqueo = 4 Then
    
        cbGrupo.Enabled = False
        cbProducto.Enabled = False
        cbPersoneria.Enabled = False
        btnSeleccionar.Enabled = False
        btnExaminar.Enabled = False
        btnNuevo.Enabled = True
        btnEditar.Enabled = True
        btnGuardarComo.Enabled = True
        btnGuardar.Enabled = False
        btnCancelar.Enabled = True
        btnSalir.Enabled = True
        
        fraCaj.Enabled = False
        fraCta.Enabled = False
        fraExceso.Enabled = False
        fraExtr.Enabled = False
        fraInterplaza.Enabled = False
        fraMon.Enabled = False
        fraOtro.Enabled = False
        fraPlaza.Enabled = False
        fraRetiroOp.Enabled = False
        fraServicios.Enabled = False
        fraVentanilla.Enabled = False
        fraVisa.Enabled = False
        
    'bloqueo editar
    ElseIf nTipoBloqueo = 5 Then
    
        cbGrupo.Enabled = False
        cbProducto.Enabled = False
        cbPersoneria.Enabled = False
        btnSeleccionar.Enabled = False
        btnExaminar.Enabled = False
        btnNuevo.Enabled = True
        btnEditar.Enabled = False
        btnGuardarComo.Enabled = True
        btnGuardar.Enabled = True
        btnCancelar.Enabled = True
        btnSalir.Enabled = True
        
        fraCaj.Enabled = True
        fraCta.Enabled = True
        fraExceso.Enabled = True
        fraExtr.Enabled = True
        fraInterplaza.Enabled = True
        fraMon.Enabled = True
        fraOtro.Enabled = True
        fraPlaza.Enabled = True
        fraRetiroOp.Enabled = True
        fraServicios.Enabled = True
        fraVentanilla.Enabled = True
        fraVisa.Enabled = True
    
    End If

End Sub
Private Sub Cancelar()
Limpiar
BlqControles (1)
End Sub
Private Sub btnCancelar_Click()
Cancelar
End Sub

Private Sub RecogerDatosComision(ByRef oComision As tComision)

Dim oCont As COMNContabilidad.NCOMContFunciones
Set oCont = New COMNContabilidad.NCOMContFunciones

oComision.MovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
oComision.Producto = 232 'Ahorros por defecto
oComision.SubProducto = CInt(Trim(Right(cbProducto.Text, 5)))
oComision.Personeria = CInt(Trim(Right(cbPersoneria.Text, 5)))
oComision.Grupo = Trim(Left(cbGrupo.Text, 5))
oComision.Estado = 1

' ******************************************** Pestaña: Operaciones en cuenta ********************************************

'Depositos soles
oComision.CtaDepositoSolesTipo = CInt(Trim(Right(cbDepSolesTipo.Text, 5)))
oComision.CtaDepositoSolesComision = CDbl(Trim(txtDepSolesComision.Text))
oComision.CtaDepositoSolesMin = CDbl(Trim(txtDepSolesMin.Text))
oComision.CtaDepositoSolesMax = CDbl(Trim(txtDepSolesMax.Text))

'Depositos dolares
oComision.CtaDepositoDolaresTipo = CInt(Trim(Right(cbDepDolaresTipo.Text, 5)))
oComision.CtaDepositoDolaresComision = CDbl(Trim(txtDepDolaresComision.Text))
oComision.CtaDepositoDolaresMin = CDbl(Trim(txtDepDolaresMin.Text))
oComision.CtaDepositoDolaresMax = CDbl(Trim(txtDepDolaresMax.Text))

'Retiros soles
oComision.CtaRetiroSolesTipo = CInt(Trim(Right(cbRetSolesTipo.Text, 5)))
oComision.CtaRetiroSolesComision = CDbl(Trim(txtRetSolesComision.Text))
oComision.CtaRetiroSolesMin = CDbl(Trim(txtRetSolesMin.Text))
oComision.CtaRetiroSolesMax = CDbl(Trim(txtRetSolesMax.Text))

'Retiros dolares
oComision.CtaRetiroDolaresTipo = CInt(Trim(Right(cbRetDolaresTipo.Text, 5)))
oComision.CtaRetiroDolaresComision = CDbl(Trim(txtRetDolaresComision.Text))
oComision.CtaRetiroDolaresMin = CDbl(Trim(txtRetDolaresMin.Text))
oComision.CtaRetiroDolaresMax = CDbl(Trim(txtRetDolaresMax.Text))

'Transferencia soles
oComision.CtaTransfSolesTipo = CInt(Trim(Right(cbTranSolesTipo.Text, 5)))
oComision.CtaTransfSolesComision = CDbl(Trim(txtTranSolesComision.Text))
oComision.CtaTransfSolesMin = CDbl(Trim(txtTranSolesMin.Text))
oComision.CtaTransfSolesMax = CDbl(Trim(txtTranSolesMax.Text))

'Transferencia dolares
oComision.CtaTransfDolaresTipo = CInt(Trim(Right(cbTranDolaresTipo.Text, 5)))
oComision.CtaTransfDolaresComision = CDbl(Trim(txtTranDolaresComision.Text))
oComision.CtaTransfDolaresMin = CDbl(Trim(txtTranDolaresMin.Text))
oComision.CtaTransfDolaresMax = CDbl(Trim(txtTranDolaresMax.Text))

' ******************************************** Pestaña: ATM y MON ***********************************************

'Operaciones libres
oComision.CajOperacionesLibresATM = CDbl(Trim(txtOpeLibresATM.Text))

'retiro
oComision.CajRetiroSolesExcesoOperacionesATM = CDbl(Trim(txtRetSolesExcesoOpeATM.Text))
oComision.CajRetiroDolaresExcesoOperacionesATM = CDbl(Trim(txtRetDolaresExcesoOpeATM.Text))

'consulta de saldos
oComision.CajConsultaSaldosSolesExcesoOperacionesATM = CDbl(Trim(txtConsultaSaldosSolesExcesoOpeATM.Text))
oComision.CajConsultaSaldosDolaresExcesoOperacionesATM = CDbl(Trim(txtConsultaSaldosDolaresExcesoOpeATM.Text))

'consulta de mov
oComision.CajConsultaMovSolesExcesoOperacionesATM = CDbl(Trim(txtConsultaMovSolesExcesoOpeATM.Text))
oComision.CajConsultaMovDolaresExcesoOperacionesATM = CDbl(Trim(txtConsultaMovDolaresExcesoOpeATM.Text))

'cambio de clave
oComision.CajCambioClaveSolesExcesoOperacionesATM = CDbl(Trim(txtCambioClaveSolesExcesoOpeATM.Text))
oComision.CajCambioClaveDolaresExcesoOperacionesATM = CDbl(Trim(txtCambioClaveDolaresExcesoOpeATM.Text))

'operaciones libres
oComision.CajOperacionesLibresMON = CInt(Trim(txtOpeLibresMON.Text))
oComision.CajRetiroSolesExcesoOperaciones05MON = CDbl(grdExceoOperaciones.TextMatrix(1, 3))
oComision.CajRetiroSolesExcesoOperaciones10MON = CDbl(grdExceoOperaciones.TextMatrix(2, 3))
oComision.CajRetiroSolesExcesoOperaciones15MON = CDbl(grdExceoOperaciones.TextMatrix(3, 3))
oComision.CajRetiroSolesExcesoOperaciones20MON = CDbl(grdExceoOperaciones.TextMatrix(4, 3))
oComision.CajRetiroSolesExcesoOperaciones25MON = CDbl(grdExceoOperaciones.TextMatrix(5, 3))
oComision.CajRetiroSolesExcesoOperaciones30MON = CDbl(grdExceoOperaciones.TextMatrix(6, 3))
oComision.CajRetiroSolesExcesoOperaciones35MON = CDbl(grdExceoOperaciones.TextMatrix(7, 3))
oComision.CajRetiroSolesExcesoOperaciones40MON = CDbl(grdExceoOperaciones.TextMatrix(8, 3))
oComision.CajRetiroSolesExcesoOperaciones45MON = CDbl(grdExceoOperaciones.TextMatrix(9, 3))

'atm diferente a globalnet
oComision.CajRetiroSolesOTRO = CDbl(Trim(txtRetSolesOTRO.Text))
oComision.CajRetiroDolaresOTRO = CDbl(Trim(txtRetDolaresOTRO.Text))

oComision.CajConsultaSaldosSolesOTRO = CDbl(Trim(txtConsultaSolesOTRO.Text))
oComision.CajConsultaSaldosDolaresOTRO = CDbl(Trim(txtConsultaDolaresOTRO.Text))

'atm del extranjero
oComision.CajRetiroSolesATMextranjero = CDbl(Trim(txtRetSolesExtranj.Text))
oComision.CajRetiroDolaresATMextranjero = CDbl(Trim(txtRetDolaresExtranj.Text))

oComision.CajConsultaSaldosSolesATMextranjero = CDbl(Trim(txtConsultaSolesExtranj.Text))
oComision.CajConsultaSaldosDolaresATMextranjero = CDbl(Trim(txtConsultaDolaresExtranj.Text))

oComision.CajCambioClaveSolesATMextranjero = CDbl(Trim(txtCambioClaveSolesExtranj.Text))
oComision.CajCambioClaveDolaresATMextranjero = CDbl(Trim(txtCambioClaveDolaresExtranj.Text))

'compras con visa
oComision.CajCompraSolesNacionalVISA = CDbl(Trim(txtComprasSolesNacionalesVISA.Text))
oComision.CajCompraDolaresNacionalVISA = CDbl(Trim(txtComprasDolaresNacionalesVISA.Text))

oComision.CajCompraSolesInternacionalVISA = CDbl(Trim(txtComprasSolesInterNacionalesVISA.Text))
oComision.CajCompraDolaresInternacionalVISA = CDbl(Trim(txtComprasDolaresInterNacionalesVISA.Text))

' ******************************************** Pestaña: Op. En Ventanilla ***********************************************

'nro de operaciones libres misma plaza
oComision.VntOperacionesLibresRetiro = CInt(txtOpeLibresRetirosVent.Text)
oComision.VntOperacionesLibresDeposito = CInt(txtOpeLibresDepVent.Text)
oComision.VntOperacionesLibresTransf = CInt(txtOpeLibresTransferVent.Text)

'nro de operaciones libres interplaza
oComision.VntOperacionesLibresInterplazaRetiro = CInt(txtOpeLibresRetInterPlazaVent.Text)
oComision.VntOperacionesLibresInterplazaDeposito = CInt(txtOpeLibresDepInterPlazaVent.Text)
oComision.VntOperacionesLibresInterplazaTransf = CInt(txtOpeLibresTransferInterPlazaVent.Text)

'Otras Operaciones.
oComision.VntOperacionesLibresOrdenesPago = CInt(txtOpeLibresOPVent.Text)
oComision.VntExcesoOperacionesOrdenesPagoSoles = CDbl(Trim(txtComExcesoOPsolesVent.Text))
oComision.VntExcesoOperacionesOrdenesPagoDolares = CDbl(Trim(txtComExcesoOPdolaresVent.Text))

'por exceso de operaciones
oComision.VntRetiroSolesExcesoOperaciones = CDbl(Trim(txtExcesoOpeRetSolesVent.Text))
oComision.VntRetiroDolaresExcesoOperaciones = CDbl(Trim(txtExcesoOpeRetDolaresVent.Text))

oComision.VntDepositoSolesExcesoOperaciones = CDbl(Trim(txtExcesoOpeDepSolesVent.Text))
oComision.VntDepositoDolaresExcesoOperaciones = CDbl(Trim(txtExcesoOpeDepDolaresVent.Text))

oComision.VntTranfSolesExcesoOperaciones = CDbl(Trim(txtExcesoOpeTranSolesVent.Text))
oComision.VntTranfDolaresExcesoOperaciones = CDbl(Trim(txtExcesoOpeTranDolaresVent.Text))

'otras operaciones en ventanilla
oComision.VntExtractoCtaAhorrosSoles = CDbl(Trim(txtExtractoCtasAhorroSolesVENT.Text))
oComision.VntExtractoCtaAhorrosDolares = CDbl(Trim(txtExtractoCtasAhorroDolaresVENT.Text))
oComision.VntExtractoCtaAhorrosValor = CInt(Trim(Right(cbExtractoCtasAhorroTipoVENT.Text, 5)))

oComision.VntConsultaSaldosVentanillaSoles = CDbl(Trim(txtConsultaSaldoSolesVENT.Text))
oComision.VntConsultaSaldosVentanillaDolares = CDbl(Trim(txtConsultaSaldoDolaresVENT.Text))
oComision.VntConsultaSaldosVentanillaValor = CInt(Trim(Right(cbConsultaSaldoTipoVENT.Text, 5)))

oComision.VntRetiroSinTarjetaDebitoSoles = CDbl(txtRetSinTarjSolesVENT.Text)
oComision.VntRetiroSinTarjetaDebitoDolares = CDbl(txtRetSinTarjDolaresVENT.Text)
oComision.VntRetiroSinTarjetaDebitoValor = CInt(Trim(Right(cbRetSinTarjTipoVENT.Text, 5)))

' ******************************************** Pestaña: Serv.  Asociados a Cta y Tarjetas ***********************

oComision.SrvMantenimientoCtaSoles = CDbl(Trim(txtMantCuentaSolesSERV.Text))
oComision.SrvMantenimientoCtaDolares = CDbl(Trim(txtMantCuentaDolaresSERV.Text))
oComision.SrvMantenimientoCtaValor = CInt(Trim(Right(cbMantCuentaTipoSERV.Text, 5)))

oComision.SrvDebitoServicioRecaudoSoles = CDbl(Trim(txtDebitoRecaudoSolesSERV.Text))
oComision.SrvDebitoServicioRecaudoDolares = CDbl(Trim(txtDebitoRecaudoDolaresSERV.Text))
oComision.SrvDebitoServicioRecaudoValor = CInt(Trim(Right(cbDebitoRecaudoTipoSERV.Text, 5)))

oComision.SrvDebitoPagoCreditoSoles = CDbl(Trim(txtDebitoCreditoSolesSERV.Text))
oComision.SrvDebitoPagoCreditoDolares = CDbl(Trim(txtDebitoCreditoDolaresSERV.Text))
oComision.SrvDebitoPagoCreditoValor = CInt(Trim(Right(cbDebitoCreditoTipoSERV.Text, 5)))

oComision.SrvEnvioEstadoCtaFisicoSoles = CDbl(Trim(txtEnvioFisicoSolesSERV.Text))
oComision.SrvEnvioEstadoCtaFisicoDolares = CDbl(Trim(txtEnvioFisicoDolaresSERV.Text))
oComision.SrvEnvioEstadoCtaFisicoValor = CInt(Trim(Right(cbEnvioFisicoTipoSERV.Text, 5)))

oComision.SrvReposicionTarjetaDebitoSoles = CDbl(Trim(txtReposicionTarjetaSolesSERV.Text))
oComision.SrvReposicionTarjetaDebitoDolares = CDbl(Trim(txtReposicionTarjetaDolaresSERV.Text))
oComision.SrvReposicionTarjetaDebitoValor = CInt(Trim(Right(cbReposicionTarjetaTipoSERV.Text, 5)))

' ******************************************** Fin de Pestañas **************************************************

End Sub

Private Sub Editar()
nTipoOperacion = 2 'Edicion
tabOperaciones.Tab = 0
BlqControles (5) ' Habilitar Controles Edicion
txtDepSolesComision.SetFocus
End Sub

Private Sub btnEditar_Click()
Editar
End Sub

Private Sub btnExaminar_Click()

Dim rsVersiones As ADODB.Recordset
Dim bVersiones As Boolean
Dim oExaminar As frmCapTarifarioExaminar
Dim oComision As tComision

'Buscando versiones que coincidan
Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
bVersiones = True
Set rsVersiones = oCon.ObtenerVersionesComision(CStr(Trim(Left(cbGrupo.Text, 5))), 232, CInt(Trim(Right(cbProducto.Text, 5))), CInt(Trim(Right(cbPersoneria.Text, 5))))
If Not rsVersiones Is Nothing Then
    If rsVersiones.RecordCount > 0 Then
        bVersiones = True
    Else
        bVersiones = False
    End If
Else
    bVersiones = False
End If

'mostrando lista de versiones encontradas
If bVersiones Then
    Set oExaminar = New frmCapTarifarioExaminar
    oExaminar.nTipo = 1 ' Comisiones
    oExaminar.rsExaminar = rsVersiones
    oExaminar.Show 1
    If oExaminar.bRespuesta Then
        nIdComision = oExaminar.Id
        If oCon.ObtenerComision(nIdComision, oComision) Then
        PintarComision oComision
        tabOperaciones.Tab = 0
        BlqControles (4)
        End If
    End If
Else
    MsgBox "No se encontraron versiones", vbInformation, "Aviso"
End If

End Sub

Private Sub PintarComision(poComision As tComision)

cbDepSolesTipo.ListIndex = IndiceListaCombo(cbDepSolesTipo, poComision.CtaDepositoSolesTipo)
txtDepSolesComision.Text = poComision.CtaDepositoSolesComision
txtDepSolesMin.Text = poComision.CtaDepositoSolesMin
txtDepSolesMax.Text = poComision.CtaDepositoSolesMax

cbDepDolaresTipo.ListIndex = IndiceListaCombo(cbDepDolaresTipo, poComision.CtaDepositoDolaresTipo)
txtDepDolaresComision.Text = poComision.CtaDepositoDolaresComision
txtDepDolaresMin.Text = poComision.CtaDepositoDolaresMin
txtDepDolaresMax.Text = poComision.CtaDepositoDolaresMax

cbRetSolesTipo.ListIndex = IndiceListaCombo(cbRetSolesTipo, poComision.CtaRetiroSolesTipo)
txtRetSolesComision.Text = poComision.CtaRetiroSolesComision
txtRetSolesMin.Text = poComision.CtaRetiroSolesMin
txtRetSolesMax.Text = poComision.CtaRetiroSolesMax

cbRetDolaresTipo.ListIndex = IndiceListaCombo(cbRetDolaresTipo, poComision.CtaRetiroDolaresTipo)
txtRetDolaresComision.Text = poComision.CtaRetiroDolaresComision
txtRetDolaresMin.Text = poComision.CtaRetiroDolaresMin
txtRetDolaresMax.Text = poComision.CtaRetiroDolaresMax

cbTranSolesTipo.ListIndex = IndiceListaCombo(cbTranSolesTipo, poComision.CtaTransfSolesTipo)
txtTranSolesComision.Text = poComision.CtaTransfSolesComision
txtTranSolesMin.Text = poComision.CtaTransfSolesMin
txtTranSolesMax.Text = poComision.CtaTransfSolesMax

cbTranDolaresTipo.ListIndex = IndiceListaCombo(cbTranDolaresTipo, poComision.CtaTransfDolaresTipo)
txtTranDolaresComision.Text = poComision.CtaTransfDolaresComision
txtTranDolaresMin.Text = poComision.CtaTransfDolaresMin
txtTranDolaresMax.Text = poComision.CtaTransfDolaresMax

txtOpeLibresATM.Text = poComision.CajOperacionesLibresATM
txtRetSolesExcesoOpeATM.Text = poComision.CajRetiroSolesExcesoOperacionesATM
txtRetDolaresExcesoOpeATM.Text = poComision.CajRetiroDolaresExcesoOperacionesATM
txtConsultaSaldosSolesExcesoOpeATM.Text = poComision.CajConsultaSaldosSolesExcesoOperacionesATM
txtConsultaSaldosDolaresExcesoOpeATM.Text = poComision.CajConsultaSaldosDolaresExcesoOperacionesATM
txtConsultaMovSolesExcesoOpeATM.Text = poComision.CajConsultaMovSolesExcesoOperacionesATM
txtConsultaMovDolaresExcesoOpeATM.Text = poComision.CajConsultaMovDolaresExcesoOperacionesATM
txtCambioClaveSolesExcesoOpeATM.Text = poComision.CajCambioClaveSolesExcesoOperacionesATM
txtCambioClaveDolaresExcesoOpeATM.Text = poComision.CajCambioClaveDolaresExcesoOperacionesATM
txtOpeLibresMON.Text = poComision.CajOperacionesLibresMON
txtRetSolesOTRO.Text = poComision.CajRetiroSolesOTRO
txtRetDolaresOTRO.Text = poComision.CajRetiroDolaresOTRO
txtConsultaSolesOTRO.Text = poComision.CajConsultaSaldosSolesOTRO
txtConsultaDolaresOTRO.Text = poComision.CajConsultaSaldosDolaresOTRO
txtRetSolesExtranj.Text = poComision.CajRetiroSolesATMextranjero
txtRetDolaresExtranj.Text = poComision.CajRetiroDolaresATMextranjero
txtConsultaSolesExtranj.Text = poComision.CajConsultaSaldosSolesATMextranjero
txtConsultaDolaresExtranj.Text = poComision.CajConsultaSaldosDolaresATMextranjero
txtCambioClaveSolesExtranj.Text = poComision.CajCambioClaveSolesATMextranjero
txtCambioClaveDolaresExtranj.Text = poComision.CajCambioClaveDolaresATMextranjero
txtComprasSolesNacionalesVISA.Text = poComision.CajCompraSolesNacionalVISA
txtComprasDolaresNacionalesVISA.Text = poComision.CajCompraDolaresNacionalVISA
txtComprasSolesInterNacionalesVISA.Text = poComision.CajCompraSolesInternacionalVISA
txtComprasDolaresInterNacionalesVISA.Text = poComision.CajCompraDolaresInternacionalVISA
txtOpeLibresRetirosVent.Text = poComision.VntOperacionesLibresRetiro
txtOpeLibresDepVent.Text = poComision.VntOperacionesLibresDeposito
txtOpeLibresTransferVent.Text = poComision.VntOperacionesLibresTransf
txtOpeLibresRetInterPlazaVent.Text = poComision.VntOperacionesLibresInterplazaRetiro
txtOpeLibresDepInterPlazaVent.Text = poComision.VntOperacionesLibresInterplazaDeposito
txtOpeLibresTransferInterPlazaVent.Text = poComision.VntOperacionesLibresInterplazaTransf
txtOpeLibresOPVent.Text = poComision.VntOperacionesLibresOrdenesPago
txtComExcesoOPsolesVent.Text = poComision.VntExcesoOperacionesOrdenesPagoSoles
txtComExcesoOPdolaresVent.Text = poComision.VntExcesoOperacionesOrdenesPagoDolares
txtExcesoOpeRetSolesVent.Text = poComision.VntRetiroSolesExcesoOperaciones
txtExcesoOpeRetDolaresVent.Text = poComision.VntRetiroDolaresExcesoOperaciones
txtExcesoOpeDepSolesVent.Text = poComision.VntDepositoSolesExcesoOperaciones
txtExcesoOpeDepDolaresVent.Text = poComision.VntDepositoDolaresExcesoOperaciones
txtExcesoOpeTranSolesVent.Text = poComision.VntTranfSolesExcesoOperaciones
txtExcesoOpeTranDolaresVent.Text = poComision.VntTranfDolaresExcesoOperaciones

txtExtractoCtasAhorroSolesVENT.Text = poComision.VntExtractoCtaAhorrosSoles
txtExtractoCtasAhorroDolaresVENT.Text = poComision.VntExtractoCtaAhorrosDolares
cbExtractoCtasAhorroTipoVENT.ListIndex = IndiceListaCombo(cbExtractoCtasAhorroTipoVENT, poComision.VntExtractoCtaAhorrosValor)

txtConsultaSaldoSolesVENT.Text = poComision.VntConsultaSaldosVentanillaSoles
txtConsultaSaldoDolaresVENT.Text = poComision.VntConsultaSaldosVentanillaDolares
cbConsultaSaldoTipoVENT.ListIndex = IndiceListaCombo(cbConsultaSaldoTipoVENT, poComision.VntConsultaSaldosVentanillaValor)

txtRetSinTarjSolesVENT.Text = poComision.VntRetiroSinTarjetaDebitoSoles
txtRetSinTarjDolaresVENT.Text = poComision.VntRetiroSinTarjetaDebitoDolares
cbRetSinTarjTipoVENT.ListIndex = IndiceListaCombo(cbRetSinTarjTipoVENT, poComision.VntRetiroSinTarjetaDebitoValor)

txtMantCuentaSolesSERV.Text = poComision.SrvMantenimientoCtaSoles
txtMantCuentaDolaresSERV.Text = poComision.SrvMantenimientoCtaDolares
cbMantCuentaTipoSERV.ListIndex = IndiceListaCombo(cbMantCuentaTipoSERV, poComision.SrvMantenimientoCtaValor)

txtDebitoRecaudoSolesSERV.Text = poComision.SrvDebitoServicioRecaudoSoles
txtDebitoRecaudoDolaresSERV.Text = poComision.SrvDebitoServicioRecaudoDolares
cbDebitoRecaudoTipoSERV.ListIndex = IndiceListaCombo(cbDebitoRecaudoTipoSERV, poComision.SrvDebitoServicioRecaudoValor)

txtDebitoCreditoSolesSERV.Text = poComision.SrvDebitoPagoCreditoSoles
txtDebitoCreditoDolaresSERV.Text = poComision.SrvDebitoPagoCreditoDolares
cbDebitoCreditoTipoSERV.ListIndex = IndiceListaCombo(cbDebitoCreditoTipoSERV, poComision.SrvDebitoPagoCreditoValor)

txtEnvioFisicoSolesSERV.Text = poComision.SrvEnvioEstadoCtaFisicoSoles
txtEnvioFisicoDolaresSERV.Text = poComision.SrvEnvioEstadoCtaFisicoDolares
cbEnvioFisicoTipoSERV.ListIndex = IndiceListaCombo(cbEnvioFisicoTipoSERV, poComision.SrvEnvioEstadoCtaFisicoValor)

txtReposicionTarjetaSolesSERV.Text = poComision.SrvReposicionTarjetaDebitoSoles
txtReposicionTarjetaDolaresSERV.Text = poComision.SrvReposicionTarjetaDebitoDolares
cbReposicionTarjetaTipoSERV.ListIndex = IndiceListaCombo(cbReposicionTarjetaTipoSERV, poComision.SrvReposicionTarjetaDebitoValor)

grdExceoOperaciones.TextMatrix(1, 3) = poComision.CajRetiroSolesExcesoOperaciones05MON
grdExceoOperaciones.TextMatrix(2, 3) = poComision.CajRetiroSolesExcesoOperaciones10MON
grdExceoOperaciones.TextMatrix(3, 3) = poComision.CajRetiroSolesExcesoOperaciones15MON
grdExceoOperaciones.TextMatrix(4, 3) = poComision.CajRetiroSolesExcesoOperaciones20MON
grdExceoOperaciones.TextMatrix(5, 3) = poComision.CajRetiroSolesExcesoOperaciones25MON
grdExceoOperaciones.TextMatrix(6, 3) = poComision.CajRetiroSolesExcesoOperaciones30MON
grdExceoOperaciones.TextMatrix(7, 3) = poComision.CajRetiroSolesExcesoOperaciones35MON
grdExceoOperaciones.TextMatrix(8, 3) = poComision.CajRetiroSolesExcesoOperaciones40MON
grdExceoOperaciones.TextMatrix(9, 3) = poComision.CajRetiroSolesExcesoOperaciones45MON

End Sub

Private Sub btnGuardar_Click()
    If Validar Then
                
        Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Dim oComision As tComision
        RecogerDatosComision oComision

        If nTipoOperacion = 1 Then 'Nueva version
            oComision.Version = oCon.ObtenerUltimaVersionComision(oComision.Grupo, oComision.Producto, oComision.SubProducto, oComision.Personeria)
            oComision.FechaRegistro = gdFecSis
            Set oGuardar = New frmCapTarifarioGuardar
            oGuardar.Comision = oComision
            oGuardar.nTipo = 1 'comisiones
            oGuardar.Caption = "Guardar..."
            oGuardar.Show 1
            If oGuardar.bRespuesta Then
                oComision.Glosa = oGuardar.sGosa
                oCon.AgregaComisionVersion oComision
                MsgBox "Se agregó la nueva version correctamente", vbInformation, "Aviso"
                Set oCon = Nothing
                Limpiar
                BlqControles (1)
            End If
            Set oGuardar = Nothing

        ElseIf nTipoOperacion = 2 Then ' Actualizacion de una version ya existente
            If MsgBox("Desea actualizar los datos de la versión seleccionada?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
                oComision.IdComision = nIdComision
                oCon.ActualizaComisionVersion oComision
                MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                Set oCon = Nothing
                Limpiar
                BlqControles (1)
            End If
        Else
            MsgBox "No hay operación asignada", vbInformation, "Aviso"
        End If
        Set oCon = Nothing

    Else
        MsgBox "Debe verificar los valores ingresados en la actual version", vbInformation, "Aviso"
    End If
End Sub

Private Sub Nuevo()
    Dim bRes As Boolean
    bRes = False
    If nTipoOperacion <> 0 Then
        If MsgBox("Si continúa se perderán los cambios, ¿desea continuar?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
            bRes = True
        End If
    Else
        bRes = True
    End If
    If bRes Then
        BlqControles (3)
        nTipoOperacion = 1 'Nuevo
        Limpiar nTipoOperacion
        txtDepSolesComision.SetFocus
    End If
End Sub

Private Sub btnGuardarComo_Click()
    If Validar Then
        Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set oGuardar = New frmCapComisionesGuardar
        Dim oComision As tComision
        RecogerDatosComision oComision
        nTipoOperacion = 1 ' Nueva Version
        oComision.Version = oCon.ObtenerUltimaVersionComision(oComision.Grupo, oComision.Producto, oComision.SubProducto, oComision.Personeria)
        oComision.FechaRegistro = gdFecSis
        oGuardar.Comision = oComision
        oGuardar.Caption = "Guardar Como..."
        oGuardar.Show 1
        If oGuardar.bRespuesta Then
            oComision.Glosa = oGuardar.sGosa
            oCon.AgregaComisionVersion oComision
            MsgBox "Se agregó la nueva version correctamente", vbInformation, "Aviso"
            Set oCon = Nothing
            Limpiar
            BlqControles (1)
        End If
        Set oCon = Nothing
        Set oGuardar = Nothing
    Else
        MsgBox "Debe verificar los valores ingresados en la actual version", vbInformation, "Aviso"
    End If
End Sub

Private Sub btnNuevo_Click()
Nuevo
End Sub

Private Sub btnSalir_Click()
    If MsgBox("Desea salir del formulario de comisiones?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub btnSeleccionar_Click()
    Dim bSeleccionar As Boolean
    bSeleccionar = True
    If cbGrupo.ListIndex < 0 Then
        bSeleccionar = False
    End If
    If cbProducto.ListIndex < 0 Then
        bSeleccionar = False
    End If
    If cbPersoneria.ListIndex < 0 Then
        bSeleccionar = False
    End If
    If Not bSeleccionar Then
        MsgBox "Debe seleccionar Agencia, Producto y Personería", vbInformation, "Aviso"
    Else
    BlqControles (2)
    End If
End Sub

Private Sub Form_Initialize()
bFocoGrid = False
nTipoOperacion = 0
nIdComision = -1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'controlando el Ctrl + V
    If KeyCode = 86 And Shift = 2 And bFocoGrid Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_Load()
CargarControles
BlqControles (1)
End Sub

'*************************** validacion de los campos numericos *****************************************

Private Sub GotFocus(ByRef txtValor As TextBox)
    If IsNumeric(txtValor.Text) Then
        txtValor.Text = txtValor.Text * 1
        txtValor.SelStart = 0
        txtValor.SelLength = Len(txtValor.Text)
        txtValor.SetFocus
    Else
        MsgBox "La casilla de texto debe contener valores numéricos", vbInformation, "Aviso"
        txtValor.Text = "0.00"
        txtValor.SelStart = 0
        txtValor.SelLength = Len(txtValor.Text)
        txtValor.SetFocus
    End If
End Sub
Private Sub LostFocus(ByRef txtValor As TextBox)
    If IsNumeric(txtValor.Text) Then
        txtValor.Text = txtValor.Text * 1
    Else
        MsgBox "El valor ingresado debe ser numérico", vbInformation, "Aviso"
        txtValor.Text = "0.00"
        txtValor.SelStart = 0
        txtValor.SelLength = Len(txtValor.Text)
        txtValor.SetFocus
    End If
End Sub

Private Sub grdExceoOperaciones_GotFocus()
bFocoGrid = True
End Sub

Private Sub grdExceoOperaciones_LostFocus()
bFocoGrid = False
End Sub

Private Sub grdExceoOperaciones_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdExceoOperaciones.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
    If pnCol = 3 Then
        If IsNumeric(grdExceoOperaciones.TextMatrix(pnRow, pnCol)) Then
            If CDbl(grdExceoOperaciones.TextMatrix(pnRow, pnCol)) < 0 Then
                Cancel = False
                MsgBox "El valor ingresado debe de ser mayor que cero", vbInformation, "Aviso"
                SendKeys "{Tab}", True
            Else
                grdExceoOperaciones.TextMatrix(pnRow, pnCol) = CDbl(grdExceoOperaciones.TextMatrix(pnRow, pnCol)) * 1
            End If
        Else
            Cancel = False
            MsgBox "El valor ingresado debe de ser numérico", vbInformation, "Aviso"
            SendKeys "{Tab}", True
        End If
    End If
End Sub

'txtDepSolesComision
Private Sub txtDepSolesComision_GotFocus()
GotFocus txtDepSolesComision
End Sub
Private Sub txtDepSolesComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtDepSolesComision, KeyAscii, 8, 2, False)
End Sub
Private Sub txtDepSolesComision_LostFocus()
LostFocus txtDepSolesComision
End Sub

'txtDepSolesMin
Private Sub txtDepSolesMin_GotFocus()
GotFocus txtDepSolesMin
End Sub
Private Sub txtDepSolesMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtDepSolesMin, KeyAscii, 8, 2, False)
End Sub
Private Sub txtDepSolesMin_LostFocus()
LostFocus txtDepSolesMin
End Sub

'txtDepSolesMax
Private Sub txtDepSolesMax_GotFocus()
GotFocus txtDepSolesMax
End Sub
Private Sub txtDepSolesMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtDepSolesMax, KeyAscii, 8, 2, False)
End Sub
Private Sub txtDepSolesMax_LostFocus()
LostFocus txtDepSolesMax
End Sub

'txtDepDolaresComision
Private Sub txtDepDolaresComision_GotFocus()
GotFocus txtDepDolaresComision
End Sub
Private Sub txtDepDolaresComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtDepDolaresComision, KeyAscii, 8, 2, False)
End Sub
Private Sub txtDepDolaresComision_LostFocus()
LostFocus txtDepDolaresComision
End Sub

'txtDepDolaresMin
Private Sub txtDepDolaresMin_GotFocus()
GotFocus txtDepDolaresMin
End Sub
Private Sub txtDepDolaresMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtDepDolaresMin, KeyAscii, 8, 2, False)
End Sub
Private Sub txtDepDolaresMin_LostFocus()
LostFocus txtDepDolaresMin
End Sub

'txtDepDolaresMax
Private Sub txtDepDolaresMax_GotFocus()
GotFocus txtDepDolaresMax
End Sub
Private Sub txtDepDolaresMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtDepDolaresMax, KeyAscii, 8, 2, False)
End Sub
Private Sub txtDepDolaresMax_LostFocus()
LostFocus txtDepDolaresMax
End Sub

'txtRetSolesComision
Private Sub txtRetSolesComision_GotFocus()
GotFocus txtRetSolesComision
End Sub
Private Sub txtRetSolesComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtRetSolesComision, KeyAscii, 8, 2, False)
End Sub
Private Sub txtRetSolesComision_LostFocus()
LostFocus txtRetSolesComision
End Sub

'txtRetSolesMin
Private Sub txtRetSolesMin_GotFocus()
GotFocus txtRetSolesMin
End Sub
Private Sub txtRetSolesMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtRetSolesMin, KeyAscii, 8, 2, False)
End Sub
Private Sub txtRetSolesMin_LostFocus()
LostFocus txtRetSolesMin
End Sub

'txtRetSolesMax
Private Sub txtRetSolesMax_GotFocus()
GotFocus txtRetSolesMax
End Sub
Private Sub txtRetSolesMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtRetSolesMax, KeyAscii, 8, 2, False)
End Sub
Private Sub txtRetSolesMax_LostFocus()
LostFocus txtRetSolesMax
End Sub

'txtRetDolaresComision
Private Sub txtRetDolaresComision_GotFocus()
GotFocus txtRetDolaresComision
End Sub
Private Sub txtRetDolaresComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtRetDolaresComision, KeyAscii, 8, 2, False)
End Sub
Private Sub txtRetDolaresComision_LostFocus()
LostFocus txtRetDolaresComision
End Sub

'txtRetDolaresMin
Private Sub txtRetDolaresMin_GotFocus()
GotFocus txtRetDolaresMin
End Sub
Private Sub txtRetDolaresMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtRetDolaresMin, KeyAscii, 8, 2, False)
End Sub
Private Sub txtRetDolaresMin_LostFocus()
LostFocus txtRetDolaresMin
End Sub

'txtRetDolaresMax
Private Sub txtRetDolaresMax_GotFocus()
GotFocus txtRetDolaresMax
End Sub
Private Sub txtRetDolaresMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtRetDolaresMax, KeyAscii, 8, 2, False)
End Sub
Private Sub txtRetDolaresMax_LostFocus()
LostFocus txtRetDolaresMax
End Sub

'txtTranSolesComision
Private Sub txtTranSolesComision_GotFocus()
GotFocus txtTranSolesComision
End Sub
Private Sub txtTranSolesComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTranSolesComision, KeyAscii, 8, 2, False)
End Sub
Private Sub txtTranSolesComision_LostFocus()
LostFocus txtTranSolesComision
End Sub

'txtTranSolesMin
Private Sub txtTranSolesMin_GotFocus()
GotFocus txtTranSolesMin
End Sub
Private Sub txtTranSolesMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTranSolesMin, KeyAscii, 8, 2, False)
End Sub
Private Sub txtTranSolesMin_LostFocus()
LostFocus txtTranSolesMin
End Sub

'txtTranSolesMax
Private Sub txtTranSolesMax_GotFocus()
GotFocus txtTranSolesMax
End Sub
Private Sub txtTranSolesMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTranSolesMax, KeyAscii, 8, 2, False)
End Sub
Private Sub txtTranSolesMax_LostFocus()
LostFocus txtTranSolesMax
End Sub

'txtTranDolaresComision
Private Sub txtTranDolaresComision_GotFocus()
GotFocus txtTranDolaresComision
End Sub
Private Sub txtTranDolaresComision_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTranDolaresComision, KeyAscii, 8, 2, False)
End Sub
Private Sub txtTranDolaresComision_LostFocus()
LostFocus txtTranDolaresComision
End Sub

'txtTranDolaresMin
Private Sub txtTranDolaresMin_GotFocus()
GotFocus txtTranDolaresMin
End Sub
Private Sub txtTranDolaresMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTranDolaresMin, KeyAscii, 8, 2, False)
End Sub
Private Sub txtTranDolaresMin_LostFocus()
LostFocus txtTranDolaresMin
End Sub

'txtTranDolaresMax
Private Sub txtTranDolaresMax_GotFocus()
GotFocus txtTranDolaresMax
End Sub
Private Sub txtTranDolaresMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTranDolaresMax, KeyAscii, 8, 2, False)
End Sub
Private Sub txtTranDolaresMax_LostFocus()
LostFocus txtTranDolaresMax
End Sub

'*************************** Fin de la validacion de los campos numericos **************************
