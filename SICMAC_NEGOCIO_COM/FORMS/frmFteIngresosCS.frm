VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFteIngresosCS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fuentes de Ingreso"
   ClientHeight    =   8175
   ClientLeft      =   4110
   ClientTop       =   1890
   ClientWidth     =   7635
   Icon            =   "frmFteIngresosCS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimir 
      Height          =   360
      Left            =   3180
      Picture         =   "frmFteIngresosCS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Imprimir Fuentes de Ingreso"
      Top             =   7820
      Width           =   450
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   345
      Left            =   2120
      TabIndex        =   79
      Top             =   7820
      Width           =   990
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   1125
      TabIndex        =   78
      Top             =   7820
      Width           =   990
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   345
      Left            =   120
      TabIndex        =   77
      Top             =   7820
      Width           =   990
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   5100
      TabIndex        =   26
      Top             =   7820
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame6 
      Caption         =   "Razon Social :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   135
      TabIndex        =   69
      Top             =   1160
      Width           =   7440
      Begin MSMask.MaskEdBox TxFecEval 
         Height          =   315
         Left            =   6060
         TabIndex        =   76
         Top             =   1260
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CmbFecha 
         Height          =   315
         Left            =   5655
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1245
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdUbigeo 
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
         Height          =   300
         Left            =   6765
         TabIndex        =   6
         Top             =   915
         Width           =   465
      End
      Begin VB.TextBox TxtRazSocTelef 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   1245
         Width           =   1755
      End
      Begin VB.TextBox TxtRazSocDirecc 
         Height          =   315
         Left            =   1815
         TabIndex        =   5
         Top             =   900
         Width           =   4920
      End
      Begin VB.TextBox TxtRazSocDescrip 
         Height          =   315
         Left            =   1815
         TabIndex        =   4
         Top             =   570
         Width           =   5415
      End
      Begin SICMACT.TxtBuscar TxtBRazonSoc 
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   225
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   503
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label30 
         Caption         =   "Fuentes Ingresos al :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         TabIndex        =   75
         Top             =   1290
         Width           =   1935
      End
      Begin VB.Label Label29 
         Caption         =   "Telefono :"
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
         Left            =   120
         TabIndex        =   73
         Top             =   1275
         Width           =   1050
      End
      Begin VB.Label Label27 
         Caption         =   "Direccion :"
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
         Left            =   135
         TabIndex        =   72
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Descrip. Actividad:"
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
         Left            =   135
         TabIndex        =   71
         Top             =   615
         Width           =   1770
      End
      Begin VB.Label LblRazonSoc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2250
         TabIndex        =   70
         Top             =   225
         Width           =   4980
      End
   End
   Begin VB.CommandButton CmdSalirCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6375
      TabIndex        =   27
      Top             =   7820
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   135
      TabIndex        =   29
      Top             =   -15
      Width           =   7440
      Begin VB.CheckBox ChkCostoProd 
         Caption         =   "Habilitar Costo de Produccion"
         Height          =   270
         Left            =   3540
         TabIndex        =   119
         Top             =   810
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         ItemData        =   "frmFteIngresosCS.frx":088C
         Left            =   1710
         List            =   "frmFteIngresosCS.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   1620
      End
      Begin VB.ComboBox CboTipoFte 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   5400
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "&Moneda :"
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
         Left            =   240
         TabIndex        =   64
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo de Fuente :"
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
         Left            =   240
         TabIndex        =   31
         Top             =   510
         Width           =   1425
      End
      Begin VB.Label LblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1710
         TabIndex        =   1
         Top             =   165
         Width           =   5400
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "&Cliente :"
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
         Left            =   255
         TabIndex        =   30
         Top             =   210
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTFuentes 
      Height          =   4965
      Left            =   105
      TabIndex        =   28
      Top             =   2820
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   8758
      _Version        =   393216
      Tab             =   1
      TabHeight       =   459
      TabCaption(0)   =   "In&gresos y Egresos"
      TabPicture(0)   =   "frmFteIngresosCS.frx":0890
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtCargo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DTPFecIni"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label26"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Balance"
      TabPicture(1)   =   "frmFteIngresosCS.frx":08AC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Costo de Produccion"
      TabPicture(2)   =   "frmFteIngresosCS.frx":08C8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CboTpoCul"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label35"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame7 
         Height          =   675
         Left            =   160
         TabIndex        =   89
         Top             =   4230
         Width           =   7185
         Begin RichTextLib.RichTextBox TxtComentariosBal 
            Height          =   405
            Left            =   120
            TabIndex        =   90
            Top             =   220
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   714
            _Version        =   393217
            MaxLength       =   300
            TextRTF         =   $"frmFteIngresosCS.frx":08E4
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   75
            TabIndex        =   91
            Top             =   15
            Width           =   840
         End
      End
      Begin VB.Frame Frame10 
         Height          =   615
         Left            =   -74730
         TabIndex        =   122
         Top             =   3990
         Width           =   6825
         Begin VB.CheckBox chkCosecha 
            Caption         =   "Cosecha"
            Height          =   225
            Left            =   4320
            TabIndex        =   127
            Top             =   270
            Width           =   1005
         End
         Begin VB.CheckBox chkOtros 
            Caption         =   "Otros"
            Height          =   225
            Left            =   5490
            TabIndex        =   126
            Top             =   270
            Width           =   1125
         End
         Begin VB.CheckBox chkDesAgricola 
            Caption         =   "Des.Agricola"
            Height          =   315
            Left            =   1260
            TabIndex        =   125
            Top             =   180
            Width           =   1215
         End
         Begin VB.CheckBox ChkMantenimiento 
            Caption         =   "Mantenimiento"
            Height          =   225
            Left            =   2610
            TabIndex        =   124
            Top             =   270
            Width           =   1335
         End
         Begin VB.CheckBox ChkSiembra 
            Caption         =   "Siembra"
            Height          =   285
            Left            =   180
            TabIndex        =   123
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Height          =   2880
         Left            =   -71565
         TabIndex        =   107
         Top             =   1125
         Width           =   3675
         Begin VB.ComboBox CboUnidad 
            Height          =   315
            Left            =   2685
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   705
            Width           =   870
         End
         Begin VB.TextBox TxtPreUni 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1515
            TabIndex        =   113
            Text            =   "0.00"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox TxtProd 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1515
            TabIndex        =   111
            Text            =   "0.00"
            Top             =   690
            Width           =   1125
         End
         Begin VB.TextBox TxtNumHec 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1500
            TabIndex        =   109
            Text            =   "0"
            Top             =   285
            Width           =   645
         End
         Begin VB.Label LblCostosIng 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1515
            TabIndex        =   120
            Top             =   1500
            Width           =   1125
         End
         Begin VB.Label LblCostosUtil 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1515
            TabIndex        =   118
            Top             =   2235
            Width           =   1125
         End
         Begin VB.Label Label48 
            Caption         =   "Utilidad               :"
            Height          =   330
            Left            =   180
            TabIndex        =   117
            Top             =   2250
            Width           =   1890
         End
         Begin VB.Label LblCostoEgr 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1500
            TabIndex        =   116
            Top             =   1875
            Width           =   1125
         End
         Begin VB.Label Label46 
            Caption         =   "Egresos              :"
            Height          =   330
            Left            =   165
            TabIndex        =   115
            Top             =   1890
            Width           =   1890
         End
         Begin VB.Label Label45 
            Caption         =   "Ingresos             :"
            Height          =   330
            Left            =   165
            TabIndex        =   114
            Top             =   1515
            Width           =   1890
         End
         Begin VB.Label Label44 
            Caption         =   "Precio Unitario   :"
            Height          =   330
            Left            =   150
            TabIndex        =   112
            Top             =   1125
            Width           =   1890
         End
         Begin VB.Label Label43 
            Caption         =   "Produccion        :"
            Height          =   330
            Left            =   135
            TabIndex        =   110
            Top             =   735
            Width           =   1890
         End
         Begin VB.Label Label42 
            Caption         =   "Hectareas          :"
            Height          =   330
            Left            =   135
            TabIndex        =   108
            Top             =   330
            Width           =   1890
         End
      End
      Begin VB.ComboBox CboTpoCul 
         Height          =   315
         Left            =   -73440
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   600
         Width           =   4335
      End
      Begin VB.Frame Frame8 
         Caption         =   "Rubro / Costo"
         Height          =   2865
         Left            =   -74730
         TabIndex        =   92
         Top             =   1125
         Width           =   3015
         Begin VB.TextBox TxtOtros 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   104
            Text            =   "0.00"
            Top             =   1800
            Width           =   1155
         End
         Begin VB.TextBox TxtPesticidas 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            TabIndex        =   102
            Text            =   "0.00"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox TxtInsumos 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   100
            Text            =   "0.00"
            Top             =   1065
            Width           =   1155
         End
         Begin VB.TextBox TxtJornal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1305
            TabIndex        =   98
            Text            =   "0.00"
            Top             =   690
            Width           =   1155
         End
         Begin VB.TextBox TxtMaq 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1350
            TabIndex        =   96
            Text            =   "0.00"
            Top             =   330
            Width           =   1155
         End
         Begin VB.Label LblCostoTotal 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1320
            TabIndex        =   106
            Top             =   2235
            Width           =   1155
         End
         Begin VB.Label Label41 
            Caption         =   "Costo Total     :"
            Height          =   330
            Left            =   150
            TabIndex        =   105
            Top             =   2250
            Width           =   1230
         End
         Begin VB.Label Label40 
            Caption         =   "Otros               :"
            Height          =   330
            Left            =   135
            TabIndex        =   103
            Top             =   1815
            Width           =   1230
         End
         Begin VB.Label Label39 
            Caption         =   "Pesticidas       :"
            Height          =   330
            Left            =   135
            TabIndex        =   101
            Top             =   1455
            Width           =   1230
         End
         Begin VB.Label Label38 
            Caption         =   "Insumos          :"
            Height          =   330
            Left            =   135
            TabIndex        =   99
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label37 
            Caption         =   "Jornales           :"
            Height          =   330
            Left            =   135
            TabIndex        =   97
            Top             =   735
            Width           =   1260
         End
         Begin VB.Label Label36 
            Caption         =   "Maquinaria      :"
            Height          =   330
            Left            =   135
            TabIndex        =   95
            Top             =   390
            Width           =   1260
         End
      End
      Begin VB.TextBox TxtCargo 
         Height          =   285
         Left            =   -72750
         TabIndex        =   13
         Top             =   2880
         Width           =   4560
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   315
         Left            =   -72735
         TabIndex        =   12
         Top             =   2445
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   68681729
         CurrentDate     =   37014
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   160
         TabIndex        =   41
         Top             =   315
         Width           =   7185
         Begin VB.TextBox txtactivofijo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2070
            MaxLength       =   13
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   1720
            Width           =   1320
         End
         Begin VB.TextBox txtDisponible 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1905
            MaxLength       =   13
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   760
            Width           =   1335
         End
         Begin VB.TextBox txtcuentas 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1905
            MaxLength       =   13
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1095
            Width           =   1335
         End
         Begin VB.TextBox txtInventario 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1905
            MaxLength       =   13
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   1410
            Width           =   1335
         End
         Begin VB.TextBox txtPrestCmact 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   5520
            MaxLength       =   13
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   1410
            Width           =   1335
         End
         Begin VB.TextBox txtOtrosPrest 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   5520
            MaxLength       =   13
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   1095
            Width           =   1335
         End
         Begin VB.TextBox txtProveedores 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   5520
            MaxLength       =   13
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   760
            Width           =   1335
         End
         Begin VB.Label lblPatrimonio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   5715
            TabIndex        =   42
            Top             =   1720
            Width           =   1320
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Activo :"
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
            Height          =   195
            Left            =   180
            TabIndex        =   58
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Activo Corriente :"
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
            Height          =   195
            Left            =   180
            TabIndex        =   57
            Top             =   495
            Width           =   1500
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Activo No Corriente :"
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
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   1740
            Width           =   1800
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Pasivo y Patrimonio :"
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
            Height          =   195
            Left            =   3810
            TabIndex        =   55
            Top             =   180
            Width           =   1800
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Pasivo :"
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
            Height          =   195
            Left            =   3810
            TabIndex        =   54
            Top             =   495
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Patrimonio :"
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
            Height          =   195
            Left            =   3810
            TabIndex        =   53
            Top             =   1740
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Disponible :"
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cuentas x Cobrar:"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   1155
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Inventario :"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label lblActCirc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   2070
            TabIndex        =   49
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lblActivo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   2070
            TabIndex        =   48
            Top             =   135
            Width           =   1320
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Prestamos CMAC-C"
            Height          =   195
            Left            =   3810
            TabIndex        =   47
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Otros Préstamos :"
            Height          =   195
            Left            =   3810
            TabIndex        =   46
            Top             =   1155
            Width           =   1245
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Proveedores :"
            Height          =   195
            Left            =   3810
            TabIndex        =   45
            Top             =   870
            Width           =   990
         End
         Begin VB.Label lblPasivo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   5715
            TabIndex        =   44
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lblPasPatrim 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   5715
            TabIndex        =   43
            Top             =   135
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1830
         Left            =   -74040
         TabIndex        =   34
         Top             =   540
         Width           =   5835
         Begin VB.TextBox TxtIngCon 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   495
            Width           =   1155
         End
         Begin VB.TextBox txtOtroIng 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   795
            Width           =   1155
         End
         Begin VB.TextBox txtIngFam 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4455
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   210
            Width           =   1155
         End
         Begin VB.TextBox txtEgreFam 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4455
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   525
            Width           =   1155
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso Cony:"
            Height          =   195
            Left            =   180
            TabIndex        =   87
            Top             =   570
            Width           =   975
         End
         Begin VB.Label LblIngresos 
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1440
            TabIndex        =   82
            Top             =   225
            Width           =   1155
         End
         Begin VB.Line Line1 
            X1              =   195
            X2              =   5640
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Left            =   4455
            TabIndex        =   40
            Top             =   1215
            Width           =   1200
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
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
            Height          =   195
            Left            =   2925
            TabIndex        =   39
            Top             =   1275
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Otros Ingresos:"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   870
            Width           =   1065
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos :"
            Height          =   195
            Left            =   210
            TabIndex        =   37
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblingreso 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso Cliente :"
            Height          =   195
            Left            =   2925
            TabIndex        =   36
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label lblEgreso 
            AutoSize        =   -1  'True
            Caption         =   "Egreso Familiar :"
            Height          =   195
            Left            =   2910
            TabIndex        =   35
            Top             =   585
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1200
         Left            =   -74100
         TabIndex        =   32
         Top             =   3180
         Width           =   6150
         Begin RichTextLib.RichTextBox Txtcomentarios 
            Height          =   915
            Left            =   90
            TabIndex        =   14
            Top             =   225
            Width           =   5910
            _ExtentX        =   10425
            _ExtentY        =   1614
            _Version        =   393217
            MaxLength       =   300
            TextRTF         =   $"frmFteIngresosCS.frx":0967
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   75
            TabIndex        =   33
            Top             =   15
            Width           =   840
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   160
         TabIndex        =   59
         Top             =   2310
         Width           =   7185
         Begin VB.TextBox TxtBalEgrFam 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   5460
            MaxLength       =   13
            TabIndex        =   85
            Text            =   "0.00"
            Top             =   790
            Width           =   1335
         End
         Begin VB.TextBox TxtBalIngFam 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1890
            MaxLength       =   13
            TabIndex        =   83
            Text            =   "0.00"
            Top             =   1110
            Width           =   1335
         End
         Begin VB.TextBox txtVentas 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1890
            MaxLength       =   13
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtrecuperacion 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0;(0)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   1890
            MaxLength       =   13
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   790
            Width           =   1335
         End
         Begin VB.TextBox txtcompras 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   5460
            MaxLength       =   13
            TabIndex        =   24
            Text            =   "0.00"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtOtrosEgresos 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   5460
            MaxLength       =   13
            TabIndex        =   25
            Text            =   "0.00"
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label lblEgresosB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   5685
            TabIndex        =   131
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label lblIngresosB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            Left            =   2100
            TabIndex        =   130
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Egresos:"
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
            Height          =   195
            Left            =   3810
            TabIndex        =   129
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos:"
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
            Height          =   195
            Left            =   180
            TabIndex        =   128
            Top             =   210
            Width           =   795
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Familiares        :"
            Height          =   195
            Left            =   3810
            TabIndex        =   86
            Top             =   1110
            Width           =   1635
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Otros Ingresos            :"
            Height          =   195
            Left            =   180
            TabIndex        =   84
            Top             =   1110
            Width           =   1605
         End
         Begin VB.Line Line2 
            X1              =   165
            X2              =   7020
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
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
            Height          =   195
            Left            =   3810
            TabIndex        =   68
            Top             =   1605
            Width           =   555
         End
         Begin VB.Label LblSaldoIngEgr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Left            =   5655
            TabIndex        =   67
            Top             =   1545
            Width           =   1320
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ventas :"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   525
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Rec. de Ctas x Cobrar :"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   810
            Width           =   1650
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Costos de Reposición :"
            Height          =   195
            Left            =   3810
            TabIndex        =   61
            Top             =   525
            Width           =   1635
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Administrativos:"
            Height          =   195
            Left            =   3810
            TabIndex        =   60
            Top             =   810
            Width           =   1635
         End
      End
      Begin VB.Label Label35 
         Caption         =   "Tipo de Cultivo  :"
         Height          =   330
         Left            =   -74730
         TabIndex        =   93
         Top             =   645
         Width           =   1260
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cargo : "
         Height          =   195
         Left            =   -74025
         TabIndex        =   66
         Top             =   2910
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio :"
         Height          =   195
         Left            =   -74025
         TabIndex        =   65
         Top             =   2490
         Width           =   1185
      End
   End
   Begin VB.CommandButton CmdFteAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   120
      TabIndex        =   80
      Top             =   7815
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton CmdFteCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   1125
      TabIndex        =   81
      Top             =   7815
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "frmFteIngresosCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPersona As UPersona_Cli ' COMDPersona.DCOMPersona   'DPersona
Dim nIndice As Integer
Dim nProcesoEjecutado As Integer '1 Nueva fte de Ingreso; 2 Editar fte de Ingreso ; 3 Consulta de Fte
Dim vsUbiGeo As String
Dim bEstadoCargando As Boolean
Dim nProcesoActual As Integer
Dim ldFecEval As Date

'Para calcular la Magnitud Empresarial
Public nPersMagnitudEmp As PersJurMagnitud

'Para el tema de impresiones
Dim sRUC As String
Dim sCiiu As String
Dim sCondDomicilio As String
Dim nNroEmpleados As Integer
Dim sDepartamento As String
Dim sProvincia As String
Dim sDistrito As String
Dim sZona As String
Dim sMagnitudEmp As String

'Revision de Calculo de Magnitud Empresarial
Private Sub CalculaMagnitudEmpresarial()
Dim nVentas As Double
'Dim nActivoFijo As Double

nVentas = CDbl(txtVentas.Text)
'nActivoFijo = CDbl(txtactivofijo.Text)

Select Case nVentas
    Case Is > 80000
        nPersMagnitudEmp = gPersJurMagnitudGrande
    Case Is > 54000
        nPersMagnitudEmp = gPersJurMagnitudMediana
    Case Is > 14000
        nPersMagnitudEmp = gPersJurMagnitudPequeña
    Case Else
        nPersMagnitudEmp = gPersJurMagnitudMicro
End Select

'Select Case nActivoFijo
'    Case Is > 180000
'        nPersMagnitudEmp = gPersJurMagnitudGrande
'    Case Is <= 180000
'        nPersMagnitudEmp = gPersJurMagnitudMediana
'    Case Is <= 80000
'        nPersMagnitudEmp = gPersJurMagnitudPequeña
'    Case Is <= 20000
'        nPersMagnitudEmp = gPersJurMagnitudMicro
'End Select

End Sub

Private Function ValidaDatosFuentesIngreso() As Boolean
Dim CadTemp As String
Dim i As Integer
Dim J As Integer
Dim nNumeFte As Integer

    ValidaDatosFuentesIngreso = True
    
    If TxFecEval.Visible Then
        CadTemp = ValidaFecha(TxFecEval.Text)
        If Len(CadTemp) > 0 Then
            MsgBox CadTemp, vbInformation, "Aviso"
            ValidaDatosFuentesIngreso = False
            Exit Function
        Else
            If CmbFecha.ListCount > 0 Then
                If CDate(ldFecEval) >= CDate(TxFecEval) Then
                    MsgBox "No Puede Ingresar una Fecha de Evaluacion Igual o Menor a la Ultima Fecha de Evaluacion de la Fuente Ingreso", vbInformation, "Aviso"
                    ValidaDatosFuentesIngreso = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    If CboTipoFte.ListIndex = -1 Then
        MsgBox "No se ha Seleccionado el Tipo de Fuente", vbInformation, "Aviso"
        CboTipoFte.SetFocus
        ValidaDatosFuentesIngreso = False
        Exit Function
    End If
    
    If CboMoneda.ListIndex = -1 Then
        MsgBox "No se ha Seleccionado la Moneda", vbInformation, "Aviso"
        CboMoneda.SetFocus
        ValidaDatosFuentesIngreso = False
        Exit Function
    End If
    
    If Len(Trim(TxtBRazonSoc.Text)) = 0 Then
        MsgBox "No Ingresado la Razon Social", vbInformation, "Aviso"
        TxtBRazonSoc.SetFocus
        ValidaDatosFuentesIngreso = False
        Exit Function
    End If
    
    CadTemp = ValidaFecha(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")))
    If Len(Trim(CadTemp)) <> 0 Then
        MsgBox CadTemp, vbInformation, "Aviso"
        DTPFecIni.SetFocus
        ValidaDatosFuentesIngreso = False
        Exit Function
    End If
    
    'Valida la Fecha de Evaluacion
    If nProcesoEjecutado = 1 Then
        CadTemp = ValidaFecha(TxFecEval.Text)
        If CadTemp <> "" Then
            MsgBox CadTemp, vbInformation, "Aviso"
            TxFecEval.SetFocus
            ValidaDatosFuentesIngreso = False
            Exit Function
        End If
    End If
    
    'Valida que Exista una Unica Fuente
    'If nProcesoEjecutado = 1 Then
    '    nNumeFte = 0
    '    For i = 0 To oPersona.NumeroFtesIngreso - 1
    '        If oPersona.ObtenerFteIngFecEval(i) = CDate(TxFecEval.Text) Then
    '            MsgBox "Ya existe una Fuente de Ingreso Con la Misma Fecha de Evaluacion", vbInformation, "Aviso"
    '            TxFecEval.SetFocus
    '        ValidaDatosFuentesIngreso = False
    '        Exit Function
    '        End If
    '    Next i
    'End If
    
    'Valida si se Ingreso el Balance en Caso de ser Fuente Independiente
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoIndependiente Then
        If CDbl(lblPatrimonio.Caption) <= 0 Then
            MsgBox "Falta Ingresar el Balance de la Fuente de Ingreso"
            SSTFuentes.Tab = 1
            txtDisponible.SetFocus
            ValidaDatosFuentesIngreso = False
            Exit Function
        End If
    End If
End Function
Private Sub HabilitaCabecera(ByVal pnHabilitar As Boolean)
    CboTipoFte.Enabled = pnHabilitar
    CboMoneda.Enabled = pnHabilitar
    TxtBRazonSoc.Enabled = pnHabilitar
    TxtRazSocDescrip.Enabled = pnHabilitar
    TxtRazSocDirecc.Enabled = pnHabilitar
    TxtRazSocTelef.Enabled = pnHabilitar
    CmdUbigeo.Enabled = pnHabilitar
End Sub
Private Sub HabilitaCostoProd(ByVal pnHabilitar As Boolean)
    
    CboTpoCul.Enabled = pnHabilitar
    TxtMaq.Enabled = pnHabilitar
    TxtJornal.Enabled = pnHabilitar
    TxtInsumos.Enabled = pnHabilitar
    TxtPesticidas.Enabled = pnHabilitar
    TxtOtros.Enabled = pnHabilitar
    TxtNumHec.Enabled = pnHabilitar
    TxtProd.Enabled = pnHabilitar
    CboUnidad.Enabled = pnHabilitar
    TxtPreUni.Enabled = pnHabilitar
    ChkSiembra.Enabled = pnHabilitar
    ChkMantenimiento.Enabled = pnHabilitar
    chkDesAgricola.Enabled = pnHabilitar
    chkOtros.Enabled = pnHabilitar
    chkCosecha.Enabled = pnHabilitar
End Sub

Private Sub HabilitaIngresosEgresos(ByVal pnHabilitar As Boolean)
    txtOtroIng.Enabled = pnHabilitar
    txtIngFam.Enabled = pnHabilitar
    txtEgreFam.Enabled = pnHabilitar
    DTPFecIni.Enabled = pnHabilitar
    TxtCargo.Enabled = pnHabilitar
    Txtcomentarios.Enabled = pnHabilitar
    TxtIngCon.Enabled = pnHabilitar
    ChkCostoProd.Enabled = pnHabilitar
End Sub


Private Sub HabilitaBalance(ByVal HabBalance As Boolean)
    TxtComentariosBal.Enabled = HabBalance
    Label6.Enabled = HabBalance
    lblActivo.Enabled = HabBalance
    Label9.Enabled = HabBalance
    lblPasPatrim.Enabled = HabBalance
    Label7.Enabled = HabBalance
    lblActCirc.Enabled = HabBalance
    Label10.Enabled = HabBalance
    lblPasivo.Enabled = HabBalance
    Label12.Enabled = HabBalance
    txtDisponible.Enabled = HabBalance
    Label19.Enabled = HabBalance
    txtProveedores.Enabled = HabBalance
    Label13.Enabled = HabBalance
    txtcuentas.Enabled = HabBalance
    Label18.Enabled = HabBalance
    txtOtrosPrest.Enabled = HabBalance
    Label14.Enabled = HabBalance
    txtInventario.Enabled = HabBalance
    Label17.Enabled = HabBalance
    txtPrestCmact.Enabled = HabBalance
    Label8.Enabled = HabBalance
    txtactivofijo.Enabled = HabBalance
    Label11.Enabled = HabBalance
    lblPatrimonio.Enabled = HabBalance
    Label15.Enabled = HabBalance
    txtVentas.Enabled = HabBalance
    Label5.Enabled = HabBalance
    txtcompras.Enabled = HabBalance
    Label20.Enabled = HabBalance
    txtrecuperacion.Enabled = HabBalance
    Label4.Enabled = HabBalance
    txtOtrosEgresos.Enabled = HabBalance
    Label47.Enabled = HabBalance
    Label49.Enabled = HabBalance
    Label31.Enabled = HabBalance
    Label32.Enabled = HabBalance
    lblIngresosB.Enabled = HabBalance
    lblEgresosB.Enabled = HabBalance
    TxtBalIngFam.Enabled = HabBalance
    TxtBalEgrFam.Enabled = HabBalance
    LblSaldoIngEgr.Enabled = HabBalance
    Frame4.Enabled = HabBalance
    Frame5.Enabled = HabBalance
    
    If Trim(Right(CboTipoFte.Text, 2)) = "2" Then
        Me.SSTFuentes.TabVisible(1) = True
        Me.SSTFuentes.TabVisible(0) = False
        Me.SSTFuentes.Tab = 1
    Else
        Me.SSTFuentes.TabVisible(1) = False
        Me.SSTFuentes.TabVisible(0) = True
        Me.SSTFuentes.Tab = 0
    End If
    
End Sub

Private Sub CargaControles()
'    Call CargaComboConstante(gPersFteIngresoTipo, CboTipoFte)
'    Call CargaComboConstante(gMoneda, CboMoneda)
'    Call CargaComboConstante(1046, CboTpoCul)
'    Call CargaComboConstante(1045, CboUnidad)

'Dim oPersona As COMDpersona.DCOMPersonas
'Set oPersona = New COMDpersona.DCOMPersonas
'Dim rsMoneda As ADODB.Recordset
'Dim rsTipoFte As ADODB.Recordset
'Dim rsTipoCul As ADODB.Recordset
'Dim rsUnidad As ADODB.Recordset
'
'Call oPersona.CargarDatosObjetosFteIngreso(rsTipoFte, rsMoneda, rsTipoCul, rsUnidad)
'
'Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
'Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
'Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
'Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
'
'Set rsMoneda = Nothing
'Set rsTipoFte = Nothing
'Set rsTipoCul = Nothing
'Set rsUnidad = Nothing
'Set oPersona = Nothing

End Sub

Private Sub CargaDatosFteIngreso(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli, _
                                ByVal rsFIDep As ADODB.Recordset, ByVal rsFIInd As ADODB.Recordset, ByVal rsFICos As ADODB.Recordset, _
                                Optional ByVal pnFteDetalle As Integer = -1)
 
Dim nUltFte As Integer

    LblCliente.Caption = PstaNombre(poPersona.NombreCompleto)
    ChkCostoProd.value = poPersona.ObtenerFteIngbCostoProd(pnIndice)
    TxtBRazonSoc.Text = poPersona.ObtenerFteIngFuente(pnIndice)
    
    '07-06-2006
    Call ObtenerDatosAdicionales(Trim(TxtBRazonSoc.Text))
    
    LblRazonSoc.Caption = poPersona.ObtenerFteIngRazonSocial(pnIndice)
    CboTipoFte.ListIndex = IndiceListaCombo(CboTipoFte, poPersona.ObtenerFteIngTipo(pnIndice))
    TxtRazSocDescrip.Text = poPersona.ObtenerFteIngRazSocDescrip(pnIndice)
    TxtRazSocDirecc.Text = poPersona.ObtenerFteIngRazSocDirecc(pnIndice)
    TxtRazSocTelef.Text = poPersona.ObtenerFteIngRazSocTelef(pnIndice)
    vsUbiGeo = poPersona.ObtenerFteIngRazSocUbiGeo(pnIndice)
    If CInt(poPersona.ObtenerFteIngTipo(pnIndice)) = gPersFteIngresoTipoDependiente Then
        If poPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) = 0 Then
            Call poPersona.RecuperaFtesIngresoDependiente(pnIndice, rsFIDep)
        End If
        Call HabilitaBalance(False)
    Else
        If poPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) = 0 Then
            Call poPersona.RecuperaFtesIngresoIndependiente(pnIndice, rsFIInd)
        End If
        Call HabilitaBalance(True)
    End If
    
    If ChkCostoProd.value = 1 Then
        Call poPersona.RecuperaFtesIngresoCostosProd(pnIndice, rsFICos)
        If pnFteDetalle = -1 Then
            nUltFte = poPersona.ObtenerFteIngIngresoNumeroCostoProd(pnIndice) - 1
        Else
            nUltFte = pnFteDetalle
        End If
        
        Call UbicaCombo(CboTpoCul, poPersona.ObtenerCostoProdnTpoCultivo(pnIndice, nUltFte))
        TxtMaq.Text = poPersona.ObtenerCostoProdnMaquinaria(pnIndice, nUltFte)
        TxtJornal.Text = poPersona.ObtenerCostoProdnJornales(pnIndice, nUltFte)
        TxtInsumos.Text = poPersona.ObtenerCostoProdnInsumos(pnIndice, nUltFte)
        TxtPesticidas.Text = poPersona.ObtenerCostoProdnPesticidas(pnIndice, nUltFte)
        TxtOtros.Text = poPersona.ObtenerCostoProdnOtros(pnIndice, nUltFte)
        
        'LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas.Text) + CDbl(TxtOtros.Text), "#0.00")
        
        TxtNumHec.Text = poPersona.ObtenerCostoProdnHectareas(pnIndice, nUltFte)
        TxtProd.Text = poPersona.ObtenerCostoProdnProduccion(pnIndice, nUltFte)
        TxtPreUni.Text = poPersona.ObtenerCostoProdnPreUni(pnIndice, nUltFte)
        Call PutOfChecked(ChkSiembra, poPersona.ObtenerCostoProdnSiembra(pnIndice, nUltFte))
        Call PutOfChecked(ChkMantenimiento, poPersona.ObtenerCostoProdnMantenimiento(pnIndice, nUltFte))
        Call PutOfChecked(chkDesAgricola, poPersona.ObtenerCostoProdnDesaAgricola(pnIndice, nUltFte))
        Call PutOfChecked(chkOtros, poPersona.ObtenerCostoProdnOtros(pnIndice, nUltFte))
        Call PutOfChecked(chkCosecha, poPersona.ObtenerCostoProdnCosecha(pnIndice, nUltFte))
        
        'LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
        
        'LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
        
        'LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) * CDbl(LblCostoEgr.Caption), "#0.00")
        
        Call UbicaCombo(CboUnidad, poPersona.ObtenerCostoProdnUniProd(pnIndice, nUltFte))
    End If
    
    CboMoneda.ListIndex = IndiceListaCombo(CboMoneda, poPersona.ObtenerFteIngMoneda(pnIndice))
    'Carga Ingresos y Egresos
    DTPFecIni.value = CDate(Format(poPersona.ObtenerFteIngInicioFuente(pnIndice), "dd/mm/yyyy"))
    TxtCargo.Text = poPersona.ObtenerFteIngCargo(pnIndice)
    Txtcomentarios.Text = poPersona.ObtenerFteIngComentarios(pnIndice)
    TxtComentariosBal.Text = poPersona.ObtenerFteIngComentarios(pnIndice)
        
    If poPersona.ObtenerFteIngIngresoTipo(pnIndice) = gPersFteIngresoTipoDependiente Then
        If pnFteDetalle = -1 Then
            nUltFte = poPersona.ObtenerFteIngIngresoNumeroFteDep(pnIndice) - 1
        Else
            nUltFte = pnFteDetalle
        End If
        TxtIngCon.Text = Format(poPersona.ObtenerFteIngIngresoConyugue(pnIndice, nUltFte), "#0.00")
        txtIngFam.Text = Format(poPersona.ObtenerFteIngIngresoFam(pnIndice, nUltFte), "#0.00")
        txtOtroIng.Text = Format(poPersona.ObtenerFteIngIngresoOtros(pnIndice, nUltFte), "#0.00")
        'LblIngresos.Caption = Format(poPersona.ObtenerFteIngIngresos(pnIndice, nUltFte), "#0.00")
        txtEgreFam.Text = Format(poPersona.ObtenerFteIngGastoFam(pnIndice, nUltFte), "#0.00")
        'lblSaldo.Caption = CDbl(txtIngFam.Text) + CDbl(txtOtroIng.Text) + CDbl(LblIngresos.Caption) - CDbl(txtEgreFam.Text)
        
    Else
        If pnFteDetalle = -1 Then
            nUltFte = poPersona.ObtenerFteIngIngresoNumeroFteIndep(pnIndice) - 1
        Else
            nUltFte = pnFteDetalle
        End If
        'Carga el Balance
        txtDisponible.Text = Format(poPersona.ObtenerFteIngActivoDisp(pnIndice, nUltFte), "#0.00")
        txtcuentas.Text = Format(poPersona.ObtenerFteIngCtasxCob(pnIndice, nUltFte), "#0.00")
        txtInventario.Text = Format(poPersona.ObtenerFteIngInventario(pnIndice, nUltFte), "#0.00")
        txtactivofijo.Text = Format(poPersona.ObtenerFteIngActivoFijo(pnIndice, nUltFte), "#0.00")
        
        txtProveedores.Text = Format(poPersona.ObtenerFteIngProveedores(pnIndice, nUltFte), "#0.00")
        txtOtrosPrest.Text = Format(poPersona.ObtenerFteIngOtrosCreditos(pnIndice, nUltFte), "#0.00")
        txtPrestCmact.Text = Format(poPersona.ObtenerFteIngCreditosCmact(pnIndice, nUltFte), "#0.00")
                
        txtVentas.Text = Format(poPersona.ObtenerFteIngVentas(pnIndice, nUltFte), "#0.00")
        txtrecuperacion.Text = Format(poPersona.ObtenerFteIngRecupCtasxCobrar(pnIndice, nUltFte), "#0.00")
        txtcompras.Text = Format(poPersona.ObtenerFteIngComprasMercad(pnIndice, nUltFte), "#0.00")
        txtOtrosEgresos.Text = Format(poPersona.ObtenerFteIngOtrosEgresos(pnIndice, nUltFte), "#0.00")
        TxtBalIngFam.Text = Format(poPersona.ObtenerFteIngBalIngFam(pnIndice, nUltFte), "#0.00")
        TxtBalEgrFam.Text = Format(poPersona.ObtenerFteIngBalEgrFam(pnIndice, nUltFte), "#0.00")
    End If
End Sub
Sub PutOfChecked(ByRef cChecked As CheckBox, ByVal pintValor)
    If pintValor = 1 Then
        cChecked.value = 1
    Else
        cChecked.value = 0
    End If
End Sub
Private Sub LimpiaFormulario()
    
    LblCliente.Caption = oPersona.NombreCompleto
    TxtBRazonSoc.Text = ""
    CboTipoFte.ListIndex = -1
    CboMoneda.ListIndex = -1
    LblIngresos.Caption = "0.00"
    txtIngFam.Text = "0.00"
    txtOtroIng.Text = "0.00"
    txtEgreFam.Text = "0.00"
    DTPFecIni.value = gdFecSis
    TxtCargo.Text = ""
    Txtcomentarios.Text = ""
    TxtComentariosBal.Text = ""
    lblActivo.Caption = "0.00"
    lblActCirc.Caption = "0.00"
    txtDisponible.Text = "0.00"
    txtcuentas.Text = "0.00"
    txtInventario.Text = "0.00"
    txtactivofijo.Text = "0.00"
    lblPasPatrim.Caption = "0.00"
    lblPasivo.Caption = "0.00"
    txtProveedores.Text = "0.00"
    txtOtrosPrest.Text = "0.00"
    txtPrestCmact.Text = "0.00"
    lblPatrimonio.Caption = "0.00"
    txtVentas.Text = "0.00"
    txtrecuperacion.Text = "0.00"
    txtcompras.Text = "0.00"
    txtOtrosEgresos.Text = "0.00"
End Sub

Private Sub LimpiaFuentesIngreso()
    LblIngresos.Caption = "0.00"
    txtIngFam.Text = "0.00"
    txtOtroIng.Text = "0.00"
    txtEgreFam.Text = "0.00"
    TxtIngCon.Text = "0.00"
    DTPFecIni.value = gdFecSis
    TxtCargo.Text = ""
    Txtcomentarios.Text = ""
    TxtComentariosBal.Text = ""
    lblActivo.Caption = "0.00"
    lblActCirc.Caption = "0.00"
    txtDisponible.Text = "0.00"
    txtcuentas.Text = "0.00"
    txtInventario.Text = "0.00"
    txtactivofijo.Text = "0.00"
    lblPasPatrim.Caption = "0.00"
    lblPasivo.Caption = "0.00"
    txtProveedores.Text = "0.00"
    txtOtrosPrest.Text = "0.00"
    txtPrestCmact.Text = "0.00"
    lblPatrimonio.Caption = "0.00"
    txtVentas.Text = "0.00"
    txtrecuperacion.Text = "0.00"
    txtcompras.Text = "0.00"
    txtOtrosEgresos.Text = "0.00"
    
    CboTpoCul.ListIndex = 0
    TxtMaq.Text = "0.00"
    TxtJornal.Text = "0.00"
    TxtInsumos.Text = "0.00"
    TxtPesticidas.Text = "0.00"
    TxtOtros.Text = "0.00"
    LblCostoTotal.Caption = "0.00"
    TxtProd.Text = "0.00"
    TxtPreUni.Text = "0.00"
    LblCostosIng.Caption = "0.00"
    LblCostoEgr.Caption = "0.00"
    LblCostosUtil.Caption = "0.00"
    TxtNumHec.Text = "0"
    ChkSiembra.value = 0
    ChkMantenimiento.value = 0
    chkDesAgricola.value = 0
    chkOtros.value = 0
    chkCosecha.value = 0
End Sub

Public Sub Editar(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli)
    Set oPersona = poPersona
    nIndice = pnIndice
    nProcesoEjecutado = 2
    bEstadoCargando = True
    'Call CargaControles
    'Call CargaDatosFteIngreso(pnIndice, poPersona)
    Call CargarDatos(pnIndice, poPersona)
    
    CmdAceptar.Visible = True
    CmdSalirCancelar.Caption = "&Cancelar"
    bEstadoCargando = False
    CmbFecha.Visible = True
    TxFecEval.Visible = False
    Call CargaComboFechaEval
    HabilitaCabecera False
    HabilitaBalance False
    HabilitaIngresosEgresos False
    HabilitaCostoProd False
    CmdAceptar.Visible = False
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        SSTFuentes.TabVisible(1) = False
        SSTFuentes.TabVisible(0) = True
        SSTFuentes.Tab = 0
    Else
        SSTFuentes.TabVisible(0) = False
        SSTFuentes.TabVisible(1) = True
        SSTFuentes.Tab = 1
    End If
    frmFteIngresosCS.Show 1
End Sub

Public Sub NuevaFteIngreso(ByRef poPersona As UPersona_Cli, Optional ByVal pnFteIndice As Integer = -1)
Dim oPersTemp As UPersona_Cli ' COMDPersona.DCOMPersona 'DPersona
    bEstadoCargando = True
    Set oPersona = poPersona
    'Call CargaControles
    Call CargarDatos(-1)
    Call LimpiaFormulario
    nProcesoEjecutado = 1
    CmdAceptar.Visible = True
    CmdSalirCancelar.Caption = "&Cancelar"
    bEstadoCargando = False
    CmbFecha.Visible = False
    TxFecEval.Visible = True
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    TxFecEval.Text = Format(gdFecSis, "dd/mm/yyyy")
    CmbFecha.Clear
    'If pnFteIndice <> -1 Then
    '    TxtBRazonSoc.Text = poPersona.ObtenerFteIngFuente(pnFteIndice)
    '    LblRazonSoc.Caption = poPersona.ObtenerFteIngRazonSocial(pnFteIndice)
    '    CboTipoFte.ListIndex = IndiceListaCombo(CboTipoFte, CInt(poPersona.ObtenerFteIngTipo(pnFteIndice)))
    '    CboMoneda.ListIndex = IndiceListaCombo(CboMoneda, CInt(poPersona.ObtenerFteIngMoneda(pnFteIndice)))
    '    TxtRazSocDescrip.Text = poPersona.ObtenerFteIngRazSocDescrip(pnFteIndice)
    '    TxtRazSocDirecc.Text = poPersona.ObtenerFteIngRazSocDirecc(pnFteIndice)
    '    TxtRazSocTelef.Text = poPersona.ObtenerFteIngRazSocTelef(pnFteIndice)
    '    Set oPersTemp = New DPersona
    '    Call oPersTemp.RecuperaPersona(TxtBRazonSoc.Text)
    '    vsUbiGeo = oPersTemp.UbicacionGeografica
    '    Set oPersTemp = Nothing
    'End If
    If ChkCostoProd.value = vbChecked Then
        ' se procede a ver lo de costos de producccion
        txtVentas.Enabled = False
        txtcompras.Enabled = False
    Else
        txtVentas.Enabled = True
        txtcompras.Enabled = True
    End If
     ldFecEval = 0  'ARCV 12-08-2006
    frmFteIngresosCS.Show 1
End Sub

Public Sub CargaComboFechaEval()
Dim i As Integer
    CmbFecha.Clear
    If oPersona.ObtenerFteIngIngresoTipo(nIndice) = gPersFteIngresoTipoDependiente Then
        For i = 0 To oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1
            CmbFecha.AddItem oPersona.ObtenerFteIngFecEval(nIndice, i, CInt(Right(CboTipoFte.Text, 2)))
        Next i
    Else
        For i = 0 To oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1
            CmbFecha.AddItem oPersona.ObtenerFteIngFecEval(nIndice, i, CInt(Right(CboTipoFte.Text, 2)))
        Next i
    End If
    bEstadoCargando = True
    If oPersona.ObtenerFteIngIngresoTipo(nIndice) = gPersFteIngresoTipoDependiente Then
        CmbFecha.ListIndex = IndiceListaCombo(CmbFecha, Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1, gPersFteIngresoTipoDependiente), "dd/mm/yyyy"))
        ldFecEval = Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1, gPersFteIngresoTipoDependiente), "dd/mm/yyyy")
    Else
        CmbFecha.ListIndex = IndiceListaCombo(CmbFecha, Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1, gPersFteIngresoTipoIndependiente), "dd/mm/yyyy"))
        ldFecEval = Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1, gPersFteIngresoTipoIndependiente), "dd/mm/yyyy")
    End If
    bEstadoCargando = False
End Sub

Public Sub ConsultarFuenteIngreso(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli)
    Set oPersona = poPersona
    nIndice = pnIndice
    nProcesoEjecutado = 3
    bEstadoCargando = True
    'Call CargaControles
    'Call CargaDatosFteIngreso(pnIndice, poPersona)
    Call CargarDatos(pnIndice, poPersona)
    CmdSalirCancelar.Caption = "&Salir"
    Call HabilitaCabecera(False)
    Call HabilitaBalance(False)
    Call HabilitaIngresosEgresos(False)
    CmbFecha.Visible = True
    TxFecEval.Visible = False
    Call CargaComboFechaEval
    bEstadoCargando = False
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        SSTFuentes.TabVisible(1) = False
        SSTFuentes.TabVisible(0) = True
        SSTFuentes.Tab = 0
    Else
        SSTFuentes.TabVisible(0) = False
        SSTFuentes.TabVisible(1) = True
        SSTFuentes.Tab = 1
    End If
    frmFteIngresosCS.Show 1
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtBRazonSoc.Enabled Then
            TxtBRazonSoc.SetFocus
        End If
    End If
End Sub

Private Sub CboTipoFte_Click()
    If Trim(Right(CboTipoFte.Text, 15)) = gPersFteIngresoTipoDependiente Then
        Call HabilitaBalance(False)
       ChkCostoProd.value = 0
       ChkCostoProd.Enabled = False
'        TxtBRazonSoc.Enabled = True
    Else
        Call HabilitaBalance(True)
        ChkCostoProd.Enabled = True
'        TxtBRazonSoc.Text = ""
'        TxtBRazonSoc.Enabled = False
'        LblRazonSoc.Caption = ""
    End If
End Sub

Private Sub CboTipoFte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboMoneda.SetFocus
    End If
End Sub

Private Sub CboUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPreUni.SetFocus
    End If
End Sub

Private Sub ChkCostoProd_Click()
    If ChkCostoProd.value = 1 Then
        SSTFuentes.TabVisible(2) = True
        SSTFuentes.Tab = 2
        txtVentas.Enabled = False
        txtcompras.Enabled = False
        CboTpoCul.Enabled = True
        TxtMaq.Enabled = True
        TxtJornal.Enabled = True
        TxtInsumos.Enabled = True
        TxtPesticidas.Enabled = True
        TxtOtros.Enabled = True
        TxtNumHec.Enabled = True
        TxtProd.Enabled = True
        CboUnidad.Enabled = True
        TxtPreUni.Enabled = True
        
    Else
        SSTFuentes.TabVisible(2) = False
        'Se Agrego
        txtVentas.Enabled = True
        txtcompras.Enabled = True
    End If
End Sub

Private Sub CmbFecha_Click()
Dim oPersonaD  As COMDPersona.DCOMPersona

    If bEstadoCargando Then
        Exit Sub
    End If
    If CmbFecha.ListCount <= 0 Then
        MsgBox "No Existe Fuente de Ingreso para Mostrar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(TxtBRazonSoc.Text)) <= 0 Then
        MsgBox "Falta Ingresar la Razon Social", vbInformation, "Aviso"
        Exit Sub
    End If
    If CmbFecha.ListIndex = -1 Then
        MsgBox "Seleccione una Fecha de Evaluacion del Credito", vbInformation, "Aviso"
        Exit Sub
    End If
    'Call CargaDatosFteIngreso(nIndice, oPersona, CmbFecha.ListIndex)
    Call CargarDatos(nIndice, oPersona, CmbFecha.ListIndex, False)
    'Verifica si ya esta asignado a un Credito
    HabilitaCabecera False
    HabilitaIngresosEgresos False
    
    HabilitaBalance False
    
    
    If Trim(Right(CboTipoFte.Text, 2)) = "2" Then
        Me.SSTFuentes.TabVisible(1) = True
        Me.SSTFuentes.TabVisible(0) = False
        Me.SSTFuentes.Tab = 1
    Else
        Me.SSTFuentes.TabVisible(1) = False
        Me.SSTFuentes.TabVisible(0) = True
        Me.SSTFuentes.Tab = 0
    End If
    
    Set oPersonaD = New COMDPersona.DCOMPersona
    Call oPersonaD.RecuperaFtesdeIngreso(oPersona.PersCodigo)
    If oPersonaD.FuenteIngresoAsignadaACredito(nIndice, CDate(CmbFecha.Text)) Then
        HabilitaIngresosEgresos False
        HabilitaBalance False
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
    Else
        CmdNuevo.Enabled = True
        CmdEditar.Enabled = True
        CmdEliminar.Enabled = True
    End If
    Set oPersonaD = Nothing
    If nProcesoEjecutado = 3 Then
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
    End If
End Sub

Private Sub CmbFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
            TxtIngCon.SetFocus
        Else
            txtDisponible.SetFocus
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()
'Dim oPersonaNeg As npersona
    
    If Not ValidaDatosFuentesIngreso Then
        Exit Sub
    End If
    
    If nProcesoEjecutado = 1 Then
        Call oPersona.AdicionaFteIngreso(CInt(Right(CboTipoFte.Text, 2)), IIf(ChkCostoProd.value = vbChecked, True, False))
        nIndice = oPersona.NumeroFtesIngreso - 1
    Else
        If oPersona.ObtenerFteIngTipoAct(nIndice) <> PersFilaNueva Then
            Call oPersona.ActualizarFteIngTipoAct(PersFilaModificada, nIndice)
            'If SSTFuentes.TabVisible(2) = True Then
             If ChkCostoProd.value = Checked Then
                Call oPersona.ActualizarCostoProdTipoAct(PersFilaModificada, nIndice, 0)
            End If
        End If
    End If
    
    If frmPersona.bNuevaPersona = False Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    
    Call oPersona.ActualizarFteIngTipoFuente(Trim(Right(CboTipoFte.Text, 20)), nIndice)
    Call oPersona.ActualizarFteIngMoneda(Trim(Right(CboMoneda.Text, 20)), nIndice)
    Call oPersona.ActualizarFteIngFuenteIngreso(Trim(TxtBRazonSoc.Text), LblRazonSoc.Caption, nIndice)
    Call oPersona.ActualizarFteIngComentarios(Txtcomentarios.Text, nIndice)
    Call oPersona.ActualizarFteIngComentarios(TxtComentariosBal.Text, nIndice)
    Call oPersona.ActualizarFteIngCargo(TxtCargo.Text, nIndice)
    Call oPersona.ActualizarFteIngRazSocDirecc(Trim(TxtRazSocDirecc.Text), nIndice)
    Call oPersona.ActualizarFteIngRazSocDescrip(Trim(TxtRazSocDescrip.Text), nIndice)
    Call oPersona.ActualizarFteIngRazSocTelef(Trim(TxtRazSocTelef.Text), nIndice)
    Call oPersona.ActualizarFteIngRazSocUbiGeo(Trim(vsUbiGeo), nIndice)
    Call oPersona.ActualizarFteIngFechaInicioFuente(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")), nIndice)
    If SSTFuentes.TabVisible(2) = True Then
         Call oPersona.ActualizarFteIngbCostoProd(ChkCostoProd.value, nIndice)
    End If
    
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        Call oPersona.ActualizarFteIngIngresos(CDbl(txtIngFam.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngIngOtros(CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
        Call oPersona.ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, 0)
        If TxFecEval.Visible Then
            Call oPersona.ActualizarFteDepIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
        End If
    Else
        Call oPersona.ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice, 0)
        If TxFecEval.Visible Then
            Call oPersona.ActualizarFteIndIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
        End If
        Call oPersona.ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice, 0)
        Call oPersona.ActualizarFteIngBalIngFam(CDbl(TxtBalIngFam.Text), nIndice, 0)
        Call oPersona.ActualizarFteIngBalEgrFam(CDbl(TxtBalEgrFam.Text), nIndice, 0)
        
    End If
    ' se verifica que el tab de produccion  este visible
    
    'If SSTFuentes.TabVisible(2) = True Then
    If ChkCostoProd.value = vbChecked Then
    'Actualiza Costos de Produccion
        If CmbFecha.Visible = True Then
            Call oPersona.ActualizarCostosdFecEval(CDate(IIf(IsDate(CmbFecha.Text), CmbFecha.Text, Date)), nIndice, 0)
        Else
            Call oPersona.ActualizarCostosdFecEval(CDate(IIf(IsDate(TxFecEval.Text), TxFecEval.Text, Date)), nIndice, 0)
        End If
        Call oPersona.ActualizarCostosnTpoCultivo(CInt(Right(CboTpoCul.Text, 2)), nIndice, 0)
        Call oPersona.ActualizarCostosnMaquinaria(CDbl(TxtMaq.Text), nIndice, 0)
        Call oPersona.ActualizarCostosnJornales(CDbl(TxtJornal), nIndice, 0)
        Call oPersona.ActualizarCostosnInsumos(CDbl(TxtInsumos), nIndice, 0)
        Call oPersona.ActualizarCostosnPesticidas(CDbl(TxtPesticidas), nIndice, 0)
        Call oPersona.ActualizarCostosnOtros(CDbl(TxtOtros), nIndice, 0)
        Call oPersona.ActualizarCostosnHectareas(CDbl(TxtNumHec), nIndice, 0)
        Call oPersona.ActualizarCostosnProduccion(CDbl(TxtProd), nIndice, 0)
        Call oPersona.ActualizarCostosnPreUni(CDbl(TxtPreUni), nIndice, 0)
        Call oPersona.ActualizarCostosnUniProd(CInt(Right(CboUnidad.Text, 2)), nIndice, 0)
        Call oPersona.ActualizarCostosTpoCultivo(Mid(CboTpoCul.Text, 1, Len(CboTpoCul.Text) - 2), nIndice, 0)
        Call oPersona.ActualizarCostossUniProd(Mid(CboUnidad.Text, 1, Len(CboUnidad.Text) - 2), nIndice, 0)
        
        Call oPersona.ActualizarCostonSiembra(GetValueOfChecked(ChkSiembra), nIndice, 0)
        Call oPersona.ActualizarCostonMantenimiento(GetValueOfChecked(ChkMantenimiento), nIndice, 0)
        Call oPersona.ActualizarCostonDesaAgricola(GetValueOfChecked(chkDesAgricola), nIndice, 0)
        Call oPersona.ActualizarCostonCosecha(GetValueOfChecked(chkCosecha), nIndice, 0)
        Call oPersona.ActualizarCostonPlanOtros(GetValueOfChecked(chkOtros), nIndice, 0)
        
   End If
        
    If nProcesoEjecutado = 1 Then
        'Set oPersonaNeg = New COMNPersona.NCOMPersona    ' COMDPersona.DCOMPersona 'npersona
        Call oPersona.ChequeoFuenteIngreso(nIndice, CInt(Right(CboTipoFte, 2)), 0)
        'Set oPersonaNeg = Nothing
    End If
    If nProcesoEjecutado <> 2 Then Call cmdImprimir_Click   'Al final
    Unload Me
End Sub

Private Sub cmdEditar_Click()
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        HabilitaBalance False
        HabilitaIngresosEgresos True
        SSTFuentes.Tab = 0
    Else
        HabilitaIngresosEgresos False
        HabilitaBalance True
        SSTFuentes.Tab = 1
    End If
    
    If Me.ChkCostoProd.value = 1 Then
        HabilitaCostoProd True
        txtVentas.Enabled = False
        txtcompras.Enabled = False
    Else
        txtVentas.Enabled = True
        txtcompras.Enabled = True
    End If
    
    HabilitaMantenimiento False
    CmbFecha.Enabled = False
    nProcesoActual = 2
    '***Modificacion LMMD******************
    Frame6.Enabled = True
    TxtBRazonSoc.Enabled = True
    TxtRazSocDescrip.Enabled = True
    TxtRazSocDirecc.Enabled = True
    TxtRazSocTelef.Enabled = True
    CmdUbigeo.Enabled = True
End Sub

Private Sub cmdeliminar_Click()
Dim oPersonaD As COMDPersona.DCOMPersona

    If MsgBox("Se va a Eliminar la Fuente de Ingreso de Fecha :" & Me.CmbFecha.Text & ", Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oPersonaD = New COMDPersona.DCOMPersona
    'If oPersonaD.FuenteIngresoAsignadaACredito(nIndice, CDate(CmbFecha.Text)) Then
    If oPersonaD.FuenteIngresoAsignadaACredito(oPersona.ObtenerFteIngcNumFuente(nIndice), CDate(CmbFecha.Text)) Then  'ARCV 14-08-2006
        MsgBox "La Fuente de Ingreso No se Puede Eliminar porque esta Asignada a un Credito", vbInformation, "Aviso"
        Set oPersonaD = Nothing
        Exit Sub
    End If
    Set oPersonaD = Nothing
    Call oPersona.ActualizarFteIngTipoActdetalle(PersFilaEliminda, nIndice, CmbFecha.ListIndex)
    Call CmbFecha.RemoveItem(CmbFecha.ListIndex)
End Sub

Private Sub CmdFteAceptar_Click()
Dim nIndiceAct As Integer
'Dim oPersonaNeg As UPersona_Cli
    If Not ValidaDatosFuentesIngreso Then
        Exit Sub
    End If

    'Si se va a adicionar una nueva fuente
    If nProcesoActual = 1 Then
        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
            Call oPersona.AdicionaFteIngresoDependiente(nIndice)
            nIndiceAct = oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1
        Else
            Call oPersona.AdicionaFteIngresoIndependiente(nIndice)
            nIndiceAct = oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1
            If ChkCostoProd.value = 1 Then
                Call oPersona.AdicionaFteIngresoCostoProd(nIndice)
            End If
        End If
                
        Call oPersona.ActualizarFteIngTipoActdetalle(PersFilaNueva, nIndice, nIndiceAct)
        CmbFecha.AddItem TxFecEval.Text
    Else
        nIndiceAct = CmbFecha.ListIndex
        Call oPersona.ActualizarFteIngTipoActdetalle(PersFilaModificada, nIndice, nIndiceAct)
    End If
    'Si se va a actualizar una fte de ingreso
    If nProcesoActual = 1 Or nProcesoActual = 2 Then
        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
            Call oPersona.ActualizarFteIngIngresos(CDbl(txtIngFam.Text) + CDbl(txtOtroIng.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngIngOtros(CDbl(txtOtroIng.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, nIndiceAct)
            If TxFecEval.Visible Then
                Call oPersona.ActualizarFteDepIngFecEval(CDate(TxFecEval.Text), nIndice, nIndiceAct)
            End If
        Else
            Call oPersona.ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice, nIndiceAct)
            If TxFecEval.Visible Then
                Call oPersona.ActualizarFteIndIngFecEval(CDate(TxFecEval.Text), nIndice, nIndiceAct)
            End If
            Call oPersona.ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngBalIngFam(CDbl(TxtBalIngFam.Text), nIndice, nIndiceAct)
            Call oPersona.ActualizarFteIngBalEgrFam(CDbl(TxtBalEgrFam.Text), nIndice, nIndiceAct)
        End If
    End If
    
    'Actualiza Costos de produccion
    If SSTFuentes.TabVisible(2) = True Then
        If TxFecEval.Visible Then
            Call oPersona.ActualizarCostosdFecEval(CDate(TxFecEval.Text), nIndice, nIndiceAct)
        Else
            Call oPersona.ActualizarCostosdFecEval(CDate(CmbFecha.Text), nIndice, nIndiceAct)
        End If
        Call oPersona.ActualizarCostosnTpoCultivo(CInt(Right(CboTpoCul.Text, 2)), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnMaquinaria(CDbl(TxtMaq.Text), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnJornales(CDbl(TxtJornal), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnInsumos(CDbl(TxtInsumos), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnPesticidas(CDbl(TxtPesticidas), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnOtros(CDbl(TxtOtros), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnHectareas(CDbl(TxtNumHec), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnProduccion(CDbl(TxtProd), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnPreUni(CDbl(TxtPreUni), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosnUniProd(CInt(Right(CboUnidad.Text, 2)), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostosTpoCultivo(Mid(CboTpoCul.Text, 1, Len(CboTpoCul.Text) - 2), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostossUniProd(Mid(CboUnidad.Text, 1, Len(CboUnidad.Text) - 2), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostonSiembra(GetValueOfChecked(ChkSiembra), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostonMantenimiento(GetValueOfChecked(ChkMantenimiento), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostonDesaAgricola(GetValueOfChecked(chkDesAgricola), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostonPlanOtros(GetValueOfChecked(chkOtros), nIndice, nIndiceAct)
        Call oPersona.ActualizarCostonCosecha(GetValueOfChecked(chkCosecha), nIndice, nIndiceAct)
  End If
 '   Set oPersonaNeg = New UPersona_Cli ' COMDPersona.DCOMPersona  'npersona
 '   Call oPersonaNeg.ChequeoFuenteIngreso(nIndice, CInt(Right(CboTipoFte, 2)), nIndiceAct)
 '   Set oPersonaNeg = Nothing
    Call oPersona.ChequeoFuenteIngreso(nIndice, CInt(Right(CboTipoFte, 2)), nIndiceAct)
    Call cmdImprimir_Click

    HabilitaBalance False
    HabilitaCostoProd False
    HabilitaMantenimiento True
    CmbFecha.Visible = True
    TxFecEval.Visible = False
    CmbFecha.Enabled = True
    CmdAceptar.Visible = True
        
End Sub

Function GetValueOfChecked(ByVal pCChecked As CheckBox) As Integer
        If pCChecked.value = vbChecked Then
           GetValueOfChecked = 1
        Else
            GetValueOfChecked = 0
        End If
End Function

Private Sub HabilitaMantenimiento(ByVal pbHabilita As Boolean)
    CmdNuevo.Visible = pbHabilita
    CmdEditar.Visible = pbHabilita
    CmdEliminar.Visible = pbHabilita
    CmdFteAceptar.Visible = Not pbHabilita
    CmdFteCancelar.Visible = Not pbHabilita
End Sub

Private Sub CmdFteCancelar_Click()
    HabilitaBalance False
    HabilitaIngresosEgresos False
    HabilitaMantenimiento True
    HabilitaCostoProd False
    CmbFecha.Visible = True
    CmbFecha.Enabled = True
    TxFecEval.Visible = False
    CmbFecha_Click
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        SSTFuentes.TabVisible(1) = False
        SSTFuentes.TabVisible(0) = True
        SSTFuentes.Tab = 0
    Else
        SSTFuentes.TabVisible(0) = False
        SSTFuentes.TabVisible(1) = True
        SSTFuentes.Tab = 1
    End If
End Sub

'Private Sub CmdImprimir_Click()
'Dim sCadImp As String
'Dim oPrev As previo.clsPrevio
'Dim oPersonaD As COMDPersona.DCOMPersona
'
'Dim bCostoProd As Boolean
'
'    If ChkCostoProd.value = vbChecked Then
'        bCostoProd = True
'    Else
'        bCostoProd = False
'    End If
'
'    Set oPrev = previo.clsPrevio
'    Set oPersonaD = New COMDPersona.DCOMPersona
'
'    Call LlenarDatosFteIngreso(oPersonaD)
'
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        'psPersCod, nIndice, gsNomAge, gdFecSis, bCostoProd, ""
'        sCadImp = oPersonaD.GenerarImpresionFteIngresoDependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
'    Else
'        sCadImp = oPersonaD.GenerarImpresionFteIngresoIndependiente_CS(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
'    End If
'    Set oPersonaD = Nothing
'    previo.Show sCadImp, "Evaluacion de Fuentes de Ingreso", False
'    Set oPrev = Nothing
'End Sub

'ARCV 12-08-2006
Private Sub cmdImprimir_Click()
Dim sCadImp As String
Dim oPrev As previo.clsPrevio
'Dim oPersonaD As COMDPersona.DCOMPersona

Dim bCostoProd As Boolean

    If ChkCostoProd.value = vbChecked Then
        bCostoProd = True
    Else
        bCostoProd = False
    End If
    
    Set oPrev = New previo.clsPrevio
'    Set oPersonaD = New COMDPersona.DCOMPersona
    
'    Call LlenarDatosFteIngreso(oPersonaD)
    
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        'psPersCod, nIndice, gsNomAge, gdFecSis, bCostoProd, ""
        'sCadImp = oPersonaD.GenerarImpresionFteIngresoDependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
        sCadImp = oPersona.GenerarImpresionFteIngresoDependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
    Else
        'sCadImp = oPersonaD.GenerarImpresionFteIngresoIndependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
        sCadImp = oPersona.GenerarImpresionFteIngresoIndependiente_CS(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
    End If
    'Set oPersonaD = Nothing
    previo.Show sCadImp, "Evaluacion de Fuentes de Ingreso", False
    Set oPrev = Nothing
End Sub
'-----------

Sub LlenarDatosFteIngreso(ByVal poPersona As COMDPersona.DCOMPersona)

Dim nIndex As Integer

With poPersona
     'If nProcesoEjecutado = 1 Then
    While nIndex <= nIndice
        Call .AdicionaFteIngreso(CInt(Right(CboTipoFte.Text, 2)), IIf(ChkCostoProd.value = vbChecked, True, False))
        nIndex = nIndex + 1
    Wend
    '   nIndice = oPersona.NumeroFtesIngreso - 1
    'Else
    '    If oPersona.ObtenerFteIngTipoAct(nIndice) <> PersFilaNueva Then
    '        Call .ActualizarFteIngTipoAct(PersFilaModificada, nIndice)
    '         If ChkCostoProd.value = Checked Then
    '            Call .ActualizarCostoProdTipoAct(PersFilaModificada, nIndice, 0)
    '        End If
    '    End If
        If ChkCostoProd.value = vbChecked Then
            Call .AdicionaFteIngresoCostoProd(nIndice)
        End If
    'End If

    'Datos Adicionales no incluidos para el Reporte
    .NombreCompleto = oPersona.NombreCompleto
    .PersCodigo = oPersona.PersCodigo

    If frmPersona.bNuevaPersona = False Then
        .TipoActualizacion = PersFilaModificada
    End If

    Call .ActualizarFteIngTipoFuente(Trim(Right(CboTipoFte.Text, 20)), nIndice)
    Call .ActualizarFteIngMoneda(Trim(Right(CboMoneda.Text, 20)), nIndice)
    Call .ActualizarFteIngFuenteIngreso(Trim(TxtBRazonSoc.Text), LblRazonSoc.Caption, nIndice)
    Call .ActualizarFteIngComentarios(Txtcomentarios.Text, nIndice)
    Call .ActualizarFteIngComentarios(TxtComentariosBal.Text, nIndice)
    Call .ActualizarFteIngCargo(TxtCargo.Text, nIndice)
    Call .ActualizarFteIngRazSocDirecc(Trim(TxtRazSocDirecc.Text), nIndice)
    Call .ActualizarFteIngRazSocDescrip(Trim(TxtRazSocDescrip.Text), nIndice)
    Call .ActualizarFteIngRazSocTelef(Trim(TxtRazSocTelef.Text), nIndice)
    Call .ActualizarFteIngRazSocUbiGeo(Trim(vsUbiGeo), nIndice)
    Call .ActualizarFteIngFechaInicioFuente(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")), nIndice)

'   Datos de la Fuente como Persona, no incluidos en el Reporte

    Call .ActualizarFteRuc(sRUC, nIndice) 'oPersona.ObtenerFteIngRuc(nIndice), nIndice)
    Call .ActualizarFteFecInicioAct(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")), nIndice) 'oPersona.ObtenerFteIngFecInicioAct(nIndice), nIndice)
    Call .ActualizarFteTipoPersJur(oPersona.ObtenerFteIngTipoPersJur(nIndice), nIndice)
    Call .ActualizarFteTelefono(TxtRazSocTelef.Text, nIndice) 'oPersona.ObtenerFteIngTelefono(nIndice), nIndice)
    Call .ActualizarFteCIUU(sCiiu, nIndice) 'oPersona.ObtenerFteIngCIUU(nIndice), nIndice)
    Call .ActualizarFteCondicionDomic(sCondDomicilio, nIndice)  'oPersona.ObtenerFteIngCondicionDomic(nIndice), nIndice)
    Call .ActualizarFteMagnitudEmp(sMagnitudEmp, nIndice) 'oPersona.ObtenerFteIngMagnitudEmp(nIndice), nIndice)
    Call .ActualizarFteNroEmpleados(nNroEmpleados, nIndice) 'oPersona.ObtenerFteIngNroEmpleados(nIndice), nIndice)
    Call .ActualizarFteDireccion(TxtRazSocDirecc.Text, nIndice) 'oPersona.ObtenerFteIngDireccion(nIndice), nIndice)
    Call .ActualizarFteDpto(sDepartamento, nIndice) 'oPersona.ObtenerFteIngDpto(nIndice), nIndice)
    Call .ActualizarFteProv(sProvincia, nIndice) 'oPersona.ObtenerFteIngProv(nIndice), nIndice)
    Call .ActualizarFteDist(sDistrito, nIndice)  '( oPersona.ObtenerFteIngDist(nIndice), nIndice)
    Call .ActualizarFteZona(sZona, nIndice) 'oPersona.ObtenerFteIngZona(nIndice), nIndice)

    If SSTFuentes.TabVisible(2) = True Then
         Call .ActualizarFteIngbCostoProd(ChkCostoProd.value, nIndice)
    End If

    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        Call .AdicionaFteIngresoDependiente(nIndice)
        Call .ActualizarFteIngIngresos(CDbl(txtIngFam.Text) + CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
        Call .ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice, 0)
        Call .ActualizarFteIngIngOtros(CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
        Call .ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice, 0)
        Call .ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, 0)
        If TxFecEval.Text <> "__/__/____" Then
            Call .ActualizarFteDepIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
        Else
            If CmbFecha.Text <> "" Then
                Call .ActualizarFteDepIngFecEval(CDate(CmbFecha.Text), nIndice, 0)
            End If
        End If
    Else
        Call .AdicionaFteIngresoIndependiente(nIndice)
        Call .ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice, 0)
        Call .ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice, 0)
        Call .ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice, 0)
        Call .ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice, 0)
        Call .ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice, 0)
        Call .ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice, 0)
        Call .ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice, 0)
        Call .ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice, 0)
        Call .ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice, 0)
        Call .ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice, 0)
        Call .ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice, 0)
        If TxFecEval.Visible Then
            Call .ActualizarFteIndIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
        End If
        Call .ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice, 0)
        Call .ActualizarFteIngBalIngFam(CDbl(TxtBalIngFam.Text), nIndice, 0)
        Call .ActualizarFteIngBalEgrFam(CDbl(TxtBalEgrFam.Text), nIndice, 0)

    End If
    ' se verifica que el tab de produccion  este visible

    'If SSTFuentes.TabVisible(2) = True Then
    If ChkCostoProd.value = vbChecked Then
    'Actualiza Costos de Produccion
        If CmbFecha.Visible = True Then
            Call .ActualizarCostosdFecEval(CDate(IIf(IsDate(CmbFecha.Text), CmbFecha.Text, Date)), nIndice, 0)
        Else
            Call .ActualizarCostosdFecEval(CDate(IIf(IsDate(TxFecEval.Text), TxFecEval.Text, Date)), nIndice, 0)
        End If
        Call .ActualizarCostosnTpoCultivo(CInt(Right(CboTpoCul.Text, 2)), nIndice, 0)
        Call .ActualizarCostosnMaquinaria(CDbl(TxtMaq.Text), nIndice, 0)
        Call .ActualizarCostosnJornales(CDbl(TxtJornal), nIndice, 0)
        Call .ActualizarCostosnInsumos(CDbl(TxtInsumos), nIndice, 0)
        Call .ActualizarCostosnPesticidas(CDbl(TxtPesticidas), nIndice, 0)
        Call .ActualizarCostosnOtros(CDbl(TxtOtros), nIndice, 0)
        Call .ActualizarCostosnHectareas(CDbl(TxtNumHec), nIndice, 0)
        Call .ActualizarCostosnProduccion(CDbl(TxtProd), nIndice, 0)
        Call .ActualizarCostosnPreUni(CDbl(TxtPreUni), nIndice, 0)
        Call .ActualizarCostosnUniProd(CInt(Right(CboUnidad.Text, 2)), nIndice, 0)
        Call .ActualizarCostosTpoCultivo(Mid(CboTpoCul.Text, 1, Len(CboTpoCul.Text) - 2), nIndice, 0)
        Call .ActualizarCostossUniProd(Mid(CboUnidad.Text, 1, Len(CboUnidad.Text) - 2), nIndice, 0)

        Call .ActualizarCostonSiembra(GetValueOfChecked(ChkSiembra), nIndice, 0)
        Call .ActualizarCostonMantenimiento(GetValueOfChecked(ChkMantenimiento), nIndice, 0)
        Call .ActualizarCostonDesaAgricola(GetValueOfChecked(chkDesAgricola), nIndice, 0)
        Call .ActualizarCostonCosecha(GetValueOfChecked(chkCosecha), nIndice, 0)
        Call .ActualizarCostonPlanOtros(GetValueOfChecked(chkOtros), nIndice, 0)

   End If

End With

End Sub

Sub CargarDatos(ByVal pnIndice As Integer, Optional ByVal poPersona As UPersona_Cli = Nothing, _
                Optional ByVal pnFteDetalle As Integer = -1, _
                Optional ByVal pbCargarControles As Boolean = True)
    
Dim oPersona As COMDPersona.DCOMPersonas
Set oPersona = New COMDPersona.DCOMPersonas
Dim rsMoneda As ADODB.Recordset
Dim rsTipoFte As ADODB.Recordset
Dim rsTipoCul As ADODB.Recordset
Dim rsUnidad As ADODB.Recordset
Dim rsFIDep As ADODB.Recordset
Dim rsFIInd As ADODB.Recordset
Dim rsFICos As ADODB.Recordset

If pnIndice = -1 Then
    Call oPersona.CargarDatosObjetosFteIngreso(rsTipoFte, rsMoneda, rsTipoCul, rsUnidad, rsFIDep, rsFIInd, rsFICos)
    If pbCargarControles Then
        Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
        Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
        Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
        Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
    End If
Else
    Call oPersona.CargarDatosObjetosFteIngreso(rsTipoFte, rsMoneda, rsTipoCul, rsUnidad, rsFIDep, rsFIInd, rsFICos, poPersona.ObtenerFteIngcNumFuente(pnIndice))
    If pbCargarControles Then
        Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
        Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
        Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
        Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
    End If
    Call CargaDatosFteIngreso(pnIndice, poPersona, rsFIDep, rsFIInd, rsFICos, pnFteDetalle)
End If

Set rsMoneda = Nothing
Set rsTipoFte = Nothing
Set rsTipoCul = Nothing
Set rsUnidad = Nothing
Set rsFIDep = Nothing
Set rsFIInd = Nothing
Set rsFICos = Nothing
Set oPersona = Nothing
End Sub


Private Sub cmdNuevo_Click()
    TxFecEval.Text = "__/__/____"
    CmbFecha.Visible = False
    TxFecEval.Visible = True
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        HabilitaBalance False
        HabilitaIngresosEgresos True
        SSTFuentes.Tab = 0
    Else
        HabilitaIngresosEgresos False
        HabilitaBalance True
        SSTFuentes.Tab = 1
    End If
    If Me.ChkCostoProd.value = 1 Then
        HabilitaCostoProd True
    End If
    Call LimpiaFuentesIngreso
    nProcesoActual = 1
    HabilitaMantenimiento False
    ChkCostoProd.Enabled = True
End Sub

Private Sub CmdSalirCancelar_Click()
    Unload Me
End Sub

Private Sub CmdUbigeo_Click()
    vsUbiGeo = Right(frmUbicacionGeo.Inicio(vsUbiGeo), 12)
End Sub

Private Sub DTPFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCargo.SetFocus
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    'me.Left = 600
    bEstadoCargando = False
    nProcesoActual = 0
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    SSTFuentes.TabVisible(2) = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call CalculaMagnitudEmpresarial
End Sub


Private Sub LblCostoEgr_Change()
    If ChkCostoProd.value = vbChecked Then
        txtcompras = LblCostoEgr
    End If
End Sub

Private Sub LblCostosIng_Change()
    If ChkCostoProd.value = vbChecked Then
        txtVentas = LblCostosIng
    End If
End Sub

Private Sub TxFecEval_GotFocus()
    fEnfoque TxFecEval
End Sub

Private Sub TxFecEval_KeyPress(KeyAscii As Integer)
'Dim oPersona As UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona

'Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona 'DPersona
    If KeyAscii = 13 Then
        If CDate(ldFecEval) >= CDate(TxFecEval) Then
            MsgBox "No Puede Ingresar una Fecha de Evaluacion Igual o Menor a la Ultima Fecha de Evaluacion de la Fuente Ingreso", vbInformation, "Aviso"
            Exit Sub
        End If
        If Trim(CboTipoFte.Text) <> "" Then
            If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
                TxtIngCon.SetFocus
            Else
                txtDisponible.SetFocus
            End If
        End If
    End If
'Set oPersona = Nothing

End Sub

Private Sub TxFecEval_LostFocus()
Dim sCad As String

    sCad = ValidaFecha(TxFecEval.Text)
    If Len(Trim(sCad)) > 0 Then
        MsgBox sCad, vbInformation, "Aviso"
        TxFecEval.SetFocus
    End If
    
End Sub

Private Sub txtactivofijo_Change()
   lblActivo.Caption = Format(CDbl(IIf(Trim(lblActCirc.Caption) = "", "0", lblActCirc.Caption)) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
   lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
   lblPasPatrim.Caption = Format(CDbl(lblPatrimonio.Caption) + CDbl(lblPasivo.Caption), "#0.00")
   
End Sub

Private Sub txtactivofijo_GotFocus()
    fEnfoque txtactivofijo
End Sub

Private Sub txtactivofijo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtactivofijo, KeyAscii, 12)
        If KeyAscii = 13 Then
            txtProveedores.SetFocus
        End If
End Sub

Private Sub txtactivofijo_LostFocus()
    txtactivofijo.Text = Format(IIf(Trim(txtactivofijo.Text) = "", 0, txtactivofijo.Text), "#0.00")
End Sub

Private Sub TxtBalEgrFam_Change()

LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

lblEgresosB.Caption = Format(CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) _
            + CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

End Sub

Private Sub TxtBalEgrFam_GotFocus()
    fEnfoque TxtBalEgrFam
End Sub

Private Sub TxtBalEgrFam_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtBalEgrFam, KeyAscii, 12)
    If KeyAscii = 13 Then
        If CmdFteAceptar.Visible Then
            CmdFteAceptar.SetFocus
        Else
            CmdAceptar.SetFocus
        End If
    End If
End Sub

Private Sub TxtBalEgrFam_LostFocus()
    If Len(Trim(TxtBalEgrFam.Text)) = 0 Then
        TxtBalEgrFam.Text = "0.00"
    End If
    TxtBalEgrFam.Text = Format(TxtBalEgrFam.Text, "#0.00")
End Sub

Private Sub TxtBalIngFam_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

lblIngresosB.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)), "#0.00")

End Sub

Private Sub TxtBalIngFam_GotFocus()
    fEnfoque TxtBalIngFam
End Sub

Private Sub TxtBalIngFam_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtBalIngFam, KeyAscii, 12)
    If KeyAscii = 13 Then
        TxtBalEgrFam.SetFocus
    End If
End Sub

Private Sub TxtBalIngFam_LostFocus()
    If Len(Trim(TxtBalIngFam.Text)) = 0 Then
        TxtBalIngFam.Text = "0.00"
    End If
    TxtBalIngFam.Text = Format(TxtBalIngFam.Text, "#0.00")
End Sub

Private Sub TxtBRazonSoc_EmiteDatos()
Dim oPersTemp As UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona

    LblRazonSoc.Caption = Trim(TxtBRazonSoc.psDescripcion)
    TxtRazSocDescrip = LblRazonSoc.Caption
    TxtRazSocDirecc.Text = TxtBRazonSoc.sPersDireccion
    Set oPersTemp = New UPersona_Cli ' COMDPersona.DCOMPersona 'DPersona
    Call oPersTemp.RecuperaPersona(TxtBRazonSoc.Text)
    vsUbiGeo = oPersTemp.UbicacionGeografica
    Set oPersTemp = Nothing
    TxtRazSocDescrip.SetFocus
        
    Call ObtenerDatosAdicionales(Trim(TxtBRazonSoc.Text))
End Sub

Private Function ObtenerDatosAdicionales(ByVal psPersCod As String)
    '03-06-2006
    Dim oPers As COMDPersona.DCOMPersonas
    Set oPers = New COMDPersona.DCOMPersonas
    Dim RS As ADODB.Recordset
    
    Set RS = oPers.ObtenerDatosReporteFteIngreso(psPersCod)
    
    Set oPers = Nothing
    
    If Not RS.EOF Then
        sRUC = RS!cPersIDnro
        sCiiu = RS!cCIIUdescripcion
        sCondDomicilio = RS!cConsDescripcion
        nNroEmpleados = RS!nPersJurEmpleados
        sMagnitudEmp = RS!Magnitud
        
        '05-06-2006
        sDepartamento = RS!cDep
        sProvincia = RS!cProv
        sDistrito = RS!cDist
        sZona = RS!cZon
        '---------------
    End If
End Function

Private Sub TxtCargo_GotFocus()
    fEnfoque TxtCargo
End Sub

Private Sub TxtCargo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        Txtcomentarios.SetFocus
    End If
End Sub

Private Sub txtcompras_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

lblEgresosB.Caption = Format(CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) _
            + CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

End Sub

Private Sub txtcompras_GotFocus()
    fEnfoque txtcompras
End Sub

Private Sub txtcompras_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtcompras, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtOtrosEgresos.SetFocus
    End If
End Sub

Private Sub txtcompras_LostFocus()
    txtcompras.Text = Format(IIf(Trim(txtcompras.Text) = "", "0.00", txtcompras.Text), "#0.00")
End Sub

Private Sub txtcuentas_Change()
    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) _
            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) _
            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
            
    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    
End Sub

Private Sub txtcuentas_GotFocus()
    fEnfoque txtcuentas
End Sub

Private Sub txtcuentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtcuentas, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtInventario.SetFocus
    End If
End Sub

Private Sub txtcuentas_LostFocus()
    txtcuentas.Text = Format(IIf(Trim(txtcuentas.Text) = "", 0, txtcuentas.Text), "#0.00")
End Sub

Private Sub txtDisponible_Change()
    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) _
            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) _
            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
            
    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    
End Sub

Private Sub txtDisponible_GotFocus()
    fEnfoque txtDisponible
End Sub

Private Sub txtDisponible_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtDisponible, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtcuentas.SetFocus
    End If
End Sub

Private Sub txtDisponible_LostFocus()
    txtDisponible.Text = Format(IIf(Trim(txtDisponible.Text) = "", 0, txtDisponible.Text), "#0.00")
End Sub

Private Sub txtEgreFam_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00"))
End Sub

Private Sub txtEgreFam_GotFocus()
    fEnfoque txtEgreFam
End Sub

Private Sub txtEgreFam_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEgreFam, KeyAscii)
    If KeyAscii = 13 Then
        DTPFecIni.SetFocus
    End If
End Sub

Private Sub txtEgreFam_LostFocus()
    If Len(Trim(txtEgreFam.Text)) > 0 Then
        txtEgreFam.Text = Format(txtEgreFam.Text, "#0.00")
    Else
        txtEgreFam.Text = "0.00"
    End If
End Sub

Private Sub TxtIngCon_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
    
End Sub

Private Sub TxtIngCon_GotFocus()
    fEnfoque TxtIngCon
End Sub

Private Sub TxtIngCon_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtIngCon, KeyAscii)
    If KeyAscii = 13 Then
        txtOtroIng.SetFocus
    End If
End Sub

Private Sub TxtIngCon_LostFocus()
    If Len(Trim(TxtIngCon.Text)) = 0 Then
        TxtIngCon.Text = "0.00"
    End If
    TxtIngCon.Text = Format(TxtIngCon.Text, "#0.00")
End Sub

Private Sub txtIngFam_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _

End Sub

Private Sub txtIngFam_GotFocus()
    fEnfoque txtIngFam
End Sub

Private Sub txtIngFam_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtIngFam, KeyAscii)
    If KeyAscii = 13 Then
        txtEgreFam.SetFocus
    End If
End Sub

Private Sub txtIngFam_LostFocus()
    If Len(Trim(txtIngFam.Text)) > 0 Then
        txtIngFam.Text = Format(txtIngFam.Text, "#0.00")
    Else
        txtIngFam.Text = "0.00"
    End If
End Sub

Private Sub TxtInsumos_Change()
    If Trim(TxtInsumos.Text) = "" Then
        TxtInsumos.Text = "0.00"
    End If
    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
End Sub

Private Sub TxtInsumos_GotFocus()
    fEnfoque TxtInsumos
End Sub

Private Sub TxtInsumos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtInsumos, KeyAscii)
    If KeyAscii = 13 Then
        TxtPesticidas.SetFocus
    End If
End Sub

Private Sub TxtInsumos_LostFocus()
    If Trim(TxtInsumos.Text) = "" Then
        TxtInsumos.Text = "0.00"
    End If
    TxtInsumos.Text = Format(TxtInsumos.Text, "#0.00")
End Sub

Private Sub txtInventario_Change()
    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) _
            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) _
            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
            
    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
        
End Sub

Private Sub txtInventario_GotFocus()
    fEnfoque txtInventario
End Sub

Private Sub txtInventario_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtInventario, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtactivofijo.SetFocus
    End If
End Sub

Private Sub txtInventario_LostFocus()
    txtInventario.Text = Format(IIf(Trim(txtInventario.Text) = "", 0, txtInventario.Text), "#0.00")
End Sub

Private Sub TxtJornal_Change()
    If Trim(TxtJornal.Text) = "" Then
        TxtJornal.Text = "0.00"
    End If
    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
End Sub

Private Sub TxtJornal_GotFocus()
    fEnfoque TxtJornal
End Sub

Private Sub TxtJornal_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtJornal, KeyAscii)
    If KeyAscii = 13 Then
        TxtInsumos.SetFocus
    End If
End Sub

Private Sub TxtJornal_LostFocus()
    If Trim(TxtJornal.Text) = "" Then
        TxtJornal.Text = "0.00"
    End If
    TxtJornal.Text = Format(TxtJornal.Text, "#0.00")
End Sub

Private Sub TxtMaq_Change()
    If Trim(TxtMaq.Text) = "" Then
        TxtMaq.Text = "0.00"
    End If
    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
End Sub

Private Sub TxtMaq_GotFocus()
    fEnfoque TxtMaq
End Sub

Private Sub TxtMaq_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMaq, KeyAscii)
    If KeyAscii = 13 Then
        TxtJornal.SetFocus
    End If
End Sub

Private Sub TxtMaq_LostFocus()
    If Trim(TxtMaq.Text) = "" Then
        TxtMaq.Text = "0.00"
    End If
    TxtMaq.Text = Format(TxtMaq.Text, "#0.00")
End Sub

Private Sub TxtNumHec_Change()
    If Trim(TxtNumHec.Text) = "" Then
        TxtNumHec.Text = "0"
    End If
    LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
    LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
    LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) - CDbl(LblCostoEgr.Caption), "#0.00")
    
End Sub

Private Sub TxtNumHec_GotFocus()
    fEnfoque TxtNumHec
End Sub

Private Sub TxtNumHec_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtNumHec, KeyAscii)
    If KeyAscii = 13 Then
        TxtProd.SetFocus
    End If
End Sub

Private Sub TxtNumHec_LostFocus()
    If Trim(TxtNumHec.Text) = "" Then
        TxtNumHec.Text = "0"
    End If
    TxtNumHec.Text = Format(TxtNumHec.Text, "#0.0")
End Sub

Private Sub txtOtroIng_Change()
    
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
    
    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00"))
            
End Sub

Private Sub txtOtroIng_GotFocus()
    fEnfoque txtOtroIng
End Sub

Private Sub txtOtroIng_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtOtroIng, KeyAscii)
    If KeyAscii = 13 Then
        txtIngFam.SetFocus
    End If
End Sub

Private Sub txtOtroIng_LostFocus()
    If Len(Trim(txtOtroIng.Text)) > 0 Then
        txtOtroIng.Text = Format(txtOtroIng.Text, "#0.00")
    Else
        txtOtroIng.Text = "0.00"
    End If
End Sub

Private Sub TxtOtros_Change()
    If Trim(TxtOtros.Text) = "" Then
        TxtOtros.Text = "0.00"
    End If
    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
End Sub

Private Sub TxtOtros_GotFocus()
    fEnfoque TxtOtros
End Sub

Private Sub TxtOtros_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtOtros, KeyAscii)
    If KeyAscii = 13 Then
        TxtNumHec.SetFocus
    End If
End Sub

Private Sub TxtOtros_LostFocus()
    If Trim(TxtOtros.Text) = "" Then
        TxtOtros.Text = "0.00"
    End If
    TxtOtros.Text = Format(TxtOtros.Text, "#0.00")
End Sub

Private Sub txtOtrosEgresos_Change()
    LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
    
    lblEgresosB.Caption = Format(CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) _
            + CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

End Sub

Private Sub txtOtrosEgresos_GotFocus()
    fEnfoque txtOtrosEgresos
End Sub

Private Sub txtOtrosEgresos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtOtrosEgresos, KeyAscii, 12)
    If KeyAscii = 13 Then
        TxtBalIngFam.SetFocus
    End If
End Sub

Private Sub txtOtrosEgresos_LostFocus()
    txtOtrosEgresos.Text = Format(IIf(Trim(txtOtrosEgresos.Text) = "", "0.00", txtOtrosEgresos.Text), "#0.00")
End Sub

Private Sub txtOtrosPrest_Change()
    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) _
            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) _
            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
            
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
    
End Sub

Private Sub txtOtrosPrest_GotFocus()
    fEnfoque txtOtrosPrest
End Sub

Private Sub txtOtrosPrest_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtOtrosPrest, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtPrestCmact.SetFocus
    End If
End Sub

Private Sub txtOtrosPrest_LostFocus()
    If Len(Trim(txtOtrosPrest.Text)) > 0 Then
        txtOtrosPrest.Text = Format(txtOtrosPrest.Text, "#0.00")
    Else
        txtOtrosPrest.Text = "0.00"
    End If
End Sub

Private Sub TxtPesticidas_Change()
    If Trim(TxtPesticidas.Text) = "" Then
        TxtPesticidas.Text = "0.00"
    End If
    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
End Sub

Private Sub TxtPesticidas_GotFocus()
    fEnfoque TxtPesticidas
End Sub

Private Sub TxtPesticidas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPesticidas, KeyAscii)
    If KeyAscii = 13 Then
        TxtOtros.SetFocus
    End If
End Sub

Private Sub TxtPesticidas_LostFocus()
    If Trim(TxtPesticidas.Text) = "" Then
        TxtPesticidas.Text = "0.00"
    End If
    TxtPesticidas.Text = Format(TxtPesticidas.Text, "#0.00")
End Sub

Private Sub txtPrestCmact_Change()
    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) _
            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) _
            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
            
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")

End Sub

Private Sub txtPrestCmact_GotFocus()
    fEnfoque txtPrestCmact
End Sub

Private Sub txtPrestCmact_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPrestCmact, KeyAscii, 12)
    If KeyAscii = 13 Then
     If txtVentas.Enabled = True Then
        txtVentas.SetFocus
    End If
    End If
End Sub

Private Sub txtPrestCmact_LostFocus()
    If Len(Trim(txtPrestCmact.Text)) > 0 Then
        txtPrestCmact.Text = Format(txtPrestCmact.Text, "#0.00")
    Else
        txtPrestCmact.Text = "0.00"
    End If
End Sub

Private Sub TxtPreUni_Change()
    
    If Trim(TxtPreUni.Text) = "" Then
        TxtPreUni.Text = "0.00"
    End If
    
    LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
    LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
    LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) - CDbl(LblCostoEgr.Caption), "#0.00")
End Sub

Private Sub TxtPreUni_GotFocus()
    fEnfoque TxtPreUni
End Sub

Private Sub TxtPreUni_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPreUni, KeyAscii)
    If KeyAscii = 13 Then
        CboTpoCul.SetFocus
    End If
End Sub

Private Sub TxtPreUni_LostFocus()
    If Trim(TxtPreUni.Text) = "" Then
        TxtPreUni.Text = "0.00"
    End If
    TxtPreUni.Text = Format(TxtPreUni.Text, "#0.00")
End Sub

Private Sub txtProd_Change()
    If Trim(TxtProd.Text) = "" Then
        TxtProd.Text = "0.00"
    End If
    LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
    LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
    LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) - CDbl(LblCostoEgr.Caption), "#0.00")
End Sub

Private Sub TxtProd_GotFocus()
    fEnfoque TxtProd
End Sub

Private Sub txtprod_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtProd, KeyAscii)
    If KeyAscii = 13 Then
        CboUnidad.SetFocus
    End If
End Sub

Private Sub TxtProd_LostFocus()
    If Trim(TxtProd.Text) = "" Then
        TxtProd.Text = "0.00"
    End If
    TxtProd.Text = Format(TxtProd.Text, "#0.00")
End Sub

Private Sub txtProveedores_Change()
    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) _
            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) _
            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
    
    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
End Sub

Private Sub txtProveedores_GotFocus()
    fEnfoque txtProveedores
End Sub

Private Sub txtProveedores_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtProveedores, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtOtrosPrest.SetFocus
    End If
End Sub


Private Sub txtProveedores_LostFocus()
    If Len(Trim(txtProveedores.Text)) > 0 Then
        txtProveedores.Text = Format(txtProveedores.Text, "#0.00")
    Else
        txtProveedores.Text = "0.00"
    End If
End Sub

Private Sub TxtRazSocDescrip_GotFocus()
    fEnfoque TxtRazSocDescrip
End Sub

Private Sub TxtRazSocDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtRazSocDirecc.SetFocus
    End If
End Sub

Private Sub TxtRazSocDirecc_GotFocus()
    fEnfoque TxtRazSocDirecc
End Sub

Private Sub TxtRazSocDirecc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdUbigeo.SetFocus
    End If
End Sub


Private Sub TxtRazSocTelef_GotFocus()
    fEnfoque TxtRazSocTelef
End Sub

Private Sub TxtRazSocTelef_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If TxFecEval.Enabled And TxFecEval.Visible Then
            TxFecEval.SetFocus
        End If
    End If
End Sub

Private Sub txtrecuperacion_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

lblIngresosB.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)), "#0.00")

End Sub

Private Sub txtrecuperacion_GotFocus()
    fEnfoque txtrecuperacion
End Sub

Private Sub txtrecuperacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtrecuperacion, KeyAscii, 12)
    If KeyAscii = 13 Then
        If txtcompras.Enabled = True Then
            txtcompras.SetFocus
        End If
    End If
End Sub

Private Sub txtrecuperacion_LostFocus()
    txtrecuperacion.Text = Format(IIf(Trim(txtrecuperacion.Text) = "", "0.00", txtrecuperacion.Text), "#0.00")
    
End Sub

Private Sub txtVentas_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

lblIngresosB.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)), "#0.00")

End Sub

Private Sub txtVentas_GotFocus()
    fEnfoque txtVentas
End Sub

Private Sub txtVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentas, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtrecuperacion.SetFocus
    End If
End Sub

Private Sub txtVentas_LostFocus()
    txtVentas.Text = Format(IIf(Trim(txtVentas.Text) = "", "0.00", txtVentas.Text), "#0.00")
End Sub
