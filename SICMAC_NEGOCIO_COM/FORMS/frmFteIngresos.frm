VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFteIngresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fuentes de Ingreso"
   ClientHeight    =   9645
   ClientLeft      =   4110
   ClientTop       =   1890
   ClientWidth     =   11040
   Icon            =   "frmFteIngresos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Height          =   420
      Left            =   3600
      Picture         =   "frmFteIngresos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   87
      ToolTipText     =   "Imprimir Fuentes de Ingreso"
      Top             =   9135
      Width           =   510
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   405
      Left            =   2430
      TabIndex        =   78
      Top             =   9135
      Width           =   1110
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   405
      Left            =   1305
      TabIndex        =   77
      Top             =   9135
      Width           =   1110
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   405
      Left            =   180
      TabIndex        =   76
      Top             =   9135
      Width           =   1110
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   4860
      TabIndex        =   25
      Top             =   9135
      Visible         =   0   'False
      Width           =   1365
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
      Height          =   1740
      Left            =   135
      TabIndex        =   68
      Top             =   840
      Width           =   10680
      Begin SICMACT.TxtBuscar TxtBRazonSoc 
         Height          =   255
         Left            =   240
         TabIndex        =   144
         Top             =   240
         Width           =   1935
         _extentx        =   1720
         _extenty        =   450
         appearance      =   1
         font            =   "frmFteIngresos.frx":088C
         tipobusqueda    =   3
      End
      Begin MSMask.MaskEdBox TxFecEval 
         Height          =   315
         Left            =   5940
         TabIndex        =   75
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
         Left            =   5775
         Style           =   2  'Dropdown List
         TabIndex        =   73
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
         TabIndex        =   5
         Top             =   915
         Width           =   465
      End
      Begin VB.TextBox TxtRazSocTelef 
         Height          =   315
         Left            =   8310
         TabIndex        =   7
         Top             =   285
         Width           =   1575
      End
      Begin VB.TextBox TxtRazSocDirecc 
         Height          =   315
         Left            =   1335
         TabIndex        =   4
         Top             =   900
         Width           =   5400
      End
      Begin VB.TextBox TxtRazSocDescrip 
         Height          =   315
         Left            =   1335
         MaxLength       =   80
         TabIndex        =   3
         Top             =   570
         Width           =   5895
      End
      Begin MSMask.MaskEdBox txtFecEEFF 
         Height          =   315
         Left            =   2280
         TabIndex        =   128
         Top             =   1320
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEEFF 
         AutoSize        =   -1  'True
         Caption         =   "Estados Financieros al :"
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
         Left            =   120
         TabIndex        =   127
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3960
         TabIndex        =   74
         Top             =   1290
         Width           =   1800
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
         Left            =   7320
         TabIndex        =   72
         Top             =   315
         Width           =   930
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
         TabIndex        =   71
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion :"
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
         TabIndex        =   70
         Top             =   615
         Width           =   1170
      End
      Begin VB.Label LblRazonSoc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2250
         TabIndex        =   69
         Top             =   225
         Width           =   4965
      End
   End
   Begin VB.CommandButton CmdSalirCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6255
      TabIndex        =   26
      Top             =   9135
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   135
      TabIndex        =   28
      Top             =   -15
      Width           =   10680
      Begin VB.CheckBox ChkCostoProd 
         Caption         =   "Habilitar Costo de Produccion"
         Height          =   195
         Left            =   7260
         TabIndex        =   118
         Top             =   570
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         ItemData        =   "frmFteIngresos.frx":08B8
         Left            =   8070
         List            =   "frmFteIngresos.frx":08BA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
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
         Left            =   7200
         TabIndex        =   63
         Top             =   240
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   210
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTFuentes 
      Height          =   6435
      Left            =   105
      TabIndex        =   27
      Top             =   2640
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   11351
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   459
      TabCaption(0)   =   "In&gresos y Egresos"
      TabPicture(0)   =   "frmFteIngresos.frx":08BC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtCargo"
      Tab(0).Control(1)=   "DTPFecIni"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Label26"
      Tab(0).Control(5)=   "Label3"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Balance"
      TabPicture(1)   =   "frmFteIngresos.frx":08D8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Costo de Produccion"
      TabPicture(2)   =   "frmFteIngresos.frx":08F4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label35"
      Tab(2).Control(1)=   "Frame8"
      Tab(2).Control(2)=   "CboTpoCul"
      Tab(2).Control(3)=   "Frame9"
      Tab(2).Control(4)=   "Frame10"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "HOJA EVALUACION"
      TabPicture(3)   =   "frmFteIngresos.frx":0910
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblTipoEval"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblTitulosEval"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblDescriEval"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblMonto1"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblMonto2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label47"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cboTipoEval"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cboConceptoEval"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtMonto"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "cmdGrabarEval"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "cboGrupoHojEval"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtMonto2"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cmdBorraLinHojaEval"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "cmdCargaRS"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "FEHojaEval"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtCodEval"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "cmdImprimeCodHojEval"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Estados Financieros"
      TabPicture(4)   =   "frmFteIngresos.frx":092C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label63"
      Tab(4).Control(1)=   "fraBalGeneral"
      Tab(4).Control(2)=   "fraEstResutado"
      Tab(4).Control(3)=   "fraFlujo"
      Tab(4).Control(4)=   "fraIndicadores"
      Tab(4).ControlCount=   5
      Begin VB.Frame fraIndicadores 
         Caption         =   "Indicadores de Riesgo Cambiario Créditicio"
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
         Height          =   1215
         Left            =   -70680
         TabIndex        =   207
         Top             =   5040
         Width           =   6375
         Begin VB.TextBox txtEFPosCambios 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   186
            Text            =   "0.00"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtEFIngresoME 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   187
            Text            =   "0.00"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtEFDeudaFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   188
            Text            =   "0.00"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtEFGastosME 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   189
            Text            =   "0.00"
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label73 
            Caption         =   "% Gastos en M.E."
            Height          =   315
            Left            =   3240
            TabIndex        =   211
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label79 
            Caption         =   "Posición de Cambios (S/. 000)"
            Height          =   435
            Left            =   120
            TabIndex        =   210
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label69 
            Caption         =   "Deuda Financiera en M.E."
            Height          =   375
            Left            =   3240
            TabIndex        =   209
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label65 
            Caption         =   "% Ingreso en M.E."
            Height          =   315
            Left            =   120
            TabIndex        =   208
            Top             =   840
            Width           =   1365
         End
      End
      Begin VB.Frame fraFlujo 
         Caption         =   "Estado de Flujo Efectivo"
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
         Left            =   -70680
         TabIndex        =   202
         Top             =   2880
         Width           =   4095
         Begin VB.TextBox txtEFFlujoEfec 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   185
            Text            =   "0.00"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtEFFlujoFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   184
            Text            =   "0.00"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtEFFujoInv 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   183
            Text            =   "0.00"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtEFFlujoOpe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   182
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label78 
            Caption         =   "Flujo de Efectivo por Act. de Inversión"
            Height          =   435
            Left            =   120
            TabIndex        =   206
            Top             =   720
            Width           =   1965
         End
         Begin VB.Label Label76 
            Caption         =   "Flujo de Efectivo por Act. de Financiamiento"
            Height          =   435
            Left            =   120
            TabIndex        =   205
            Top             =   1200
            Width           =   2085
         End
         Begin VB.Label Label75 
            Caption         =   "Flujo Efectivo Total"
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
            TabIndex        =   204
            Top             =   1680
            Width           =   2025
         End
         Begin VB.Label Label64 
            Caption         =   "Flujo de Efectivo por Act. de Operación"
            Height          =   435
            Left            =   120
            TabIndex        =   203
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraEstResutado 
         Caption         =   "Estado de Resultados"
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
         Height          =   3375
         Left            =   -74880
         TabIndex        =   193
         Top             =   2880
         Width           =   4095
         Begin VB.TextBox txtEFIngresoFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   179
            Text            =   "0.00"
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtEFUtNeta 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   181
            Text            =   "0.00"
            Top             =   3000
            Width           =   1695
         End
         Begin VB.TextBox txtEFVentas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   174
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEFCostVentas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   175
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFUtBruta 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   176
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtEFGastosOpe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   177
            Text            =   "0.00"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtEFUtOpe 
            Alignment       =   1  'Right Justify
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
            Left            =   2280
            TabIndex        =   178
            Text            =   "0.00"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtEFGastoFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   180
            Text            =   "0.00"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos Financieros"
            Height          =   195
            Left            =   120
            TabIndex        =   212
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "Ventas"
            Height          =   195
            Left            =   120
            TabIndex        =   200
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Utilidad/ (Pérdida)Neta"
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
            TabIndex        =   199
            Top             =   3000
            Width           =   1980
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Financieros"
            Height          =   195
            Left            =   120
            TabIndex        =   198
            Top             =   2640
            Width           =   1350
         End
         Begin VB.Label Label70 
            Caption         =   "Utilidad/ (Pérdida)Operativa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   197
            Top             =   1750
            Width           =   1620
         End
         Begin VB.Label Label68 
            Caption         =   "Utilidad/ (Pérdida)Bruta"
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
            TabIndex        =   196
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label Label67 
            Caption         =   "Gastos Operativos (Adm+Srv+Vtas)"
            Height          =   435
            Left            =   120
            TabIndex        =   195
            Top             =   1320
            Width           =   2085
         End
         Begin VB.Label Label66 
            Caption         =   "Costo de Ventas"
            Height          =   315
            Left            =   120
            TabIndex        =   194
            Top             =   600
            Width           =   1245
         End
      End
      Begin VB.Frame fraBalGeneral 
         Caption         =   "Balance General"
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
         Height          =   2415
         Left            =   -74880
         TabIndex        =   148
         Top             =   360
         Width           =   10575
         Begin VB.TextBox txtEFTotalPat 
            Alignment       =   1  'Right Justify
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
            Left            =   8760
            TabIndex        =   173
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtEFResulAcum 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8760
            TabIndex        =   172
            Text            =   "0.00"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtEFTotalPas 
            Alignment       =   1  'Right Justify
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
            Left            =   5280
            TabIndex        =   170
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtEFDeudaFinanL 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   169
            Text            =   "0.00"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtEFPasCorriente 
            Alignment       =   1  'Right Justify
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
            Left            =   5280
            TabIndex        =   168
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtEFDeudaFinanC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   167
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFTotalAct 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   165
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtEFActFijo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   164
            Text            =   "0.00"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtEFActCorriente 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   163
            Text            =   "0.00"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtEFExiste 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   162
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtEFCuentaCobrar 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   161
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFCapSocial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8760
            TabIndex        =   171
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEFProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5280
            TabIndex        =   166
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEFCajaBanco 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            TabIndex        =   160
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label62 
            Caption         =   "Capital Social + Adicional"
            Height          =   435
            Left            =   7440
            TabIndex        =   192
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label61 
            Caption         =   "Resultados Acumulados"
            Height          =   435
            Left            =   7440
            TabIndex        =   191
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label60 
            Caption         =   "Total Patrimonio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   7440
            TabIndex        =   190
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label Label59 
            Caption         =   "Cuentas x Cobrar Comerc."
            Height          =   435
            Left            =   120
            TabIndex        =   159
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Existencias"
            Height          =   195
            Left            =   120
            TabIndex        =   158
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Activo Corriente"
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
            Left            =   120
            TabIndex        =   157
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Activo Fijo"
            Height          =   195
            Left            =   120
            TabIndex        =   156
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Total Activo"
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
            Left            =   120
            TabIndex        =   155
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Proveedores"
            Height          =   195
            Left            =   3360
            TabIndex        =   154
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label53 
            Caption         =   "Deuda Financiera Cte."
            Height          =   315
            Left            =   3360
            TabIndex        =   153
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Pasivo Corriente"
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
            Left            =   3360
            TabIndex        =   152
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label Label51 
            Caption         =   "Deuda Financiera a L.P."
            Height          =   315
            Left            =   3360
            TabIndex        =   151
            Top             =   1320
            Width           =   1725
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Total Pasivo"
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
            Left            =   3360
            TabIndex        =   150
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Caja-Bancos"
            Height          =   195
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdImprimeCodHojEval 
         Caption         =   "Codigos Hoja Eval."
         Height          =   255
         Left            =   120
         TabIndex        =   147
         Top             =   5880
         Width           =   1695
      End
      Begin VB.TextBox txtCodEval 
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
         Left            =   1320
         TabIndex        =   133
         Top             =   1200
         Width           =   1575
      End
      Begin SICMACT.FlexEdit FEHojaEval 
         Height          =   4215
         Left            =   120
         TabIndex        =   145
         Top             =   1560
         Width           =   9615
         _extentx        =   16960
         _extenty        =   5953
         cols0           =   8
         highlight       =   1
         allowuserresizing=   3
         encabezadosnombres=   "#-Tipo Eval-Titulo Eval-Descripcion-Personal-Negocio-Unico-cCodHojEval"
         encabezadosanchos=   "400-1750-1400-2300-1200-1200-1200-0"
         font            =   "frmFteIngresos.frx":0948
         fontfixed       =   "frmFteIngresos.frx":0974
         columnasaeditar =   "X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-R-R-R-L"
         formatosedit    =   "0-0-0-0-2-2-2-0"
         textarray0      =   "#"
         lbultimainstancia=   -1  'True
         lbpuntero       =   -1  'True
         colwidth0       =   405
         rowheight0      =   300
      End
      Begin VB.CommandButton cmdCargaRS 
         Caption         =   "&Imprimir Hoja de Evaluación"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   143
         Top             =   5880
         Width           =   2775
      End
      Begin VB.CommandButton cmdBorraLinHojaEval 
         Caption         =   "Borrar linea"
         Height          =   255
         Left            =   8520
         TabIndex        =   142
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox txtMonto2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6840
         TabIndex        =   136
         Text            =   "0.00"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cboGrupoHojEval 
         Height          =   315
         ItemData        =   "frmFteIngresos.frx":09A2
         Left            =   1320
         List            =   "frmFteIngresos.frx":09A4
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdGrabarEval 
         Caption         =   "Añadir"
         Height          =   315
         Left            =   8640
         TabIndex        =   137
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4080
         TabIndex        =   135
         Text            =   "0.00"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cboConceptoEval 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   840
         Width           =   7215
      End
      Begin VB.ComboBox cboTipoEval 
         Height          =   315
         ItemData        =   "frmFteIngresos.frx":09A6
         Left            =   4080
         List            =   "frmFteIngresos.frx":09A8
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame10 
         Height          =   615
         Left            =   -74730
         TabIndex        =   121
         Top             =   4245
         Width           =   6825
         Begin VB.CheckBox chkCosecha 
            Caption         =   "Cosecha"
            Height          =   225
            Left            =   4320
            TabIndex        =   126
            Top             =   270
            Width           =   1005
         End
         Begin VB.CheckBox chkOtros 
            Caption         =   "Otros"
            Height          =   225
            Left            =   5490
            TabIndex        =   125
            Top             =   270
            Width           =   1125
         End
         Begin VB.CheckBox chkDesAgricola 
            Caption         =   "Des.Agricola"
            Height          =   315
            Left            =   1260
            TabIndex        =   124
            Top             =   180
            Width           =   1215
         End
         Begin VB.CheckBox ChkMantenimiento 
            Caption         =   "Mantenimiento"
            Height          =   225
            Left            =   2610
            TabIndex        =   123
            Top             =   270
            Width           =   1335
         End
         Begin VB.CheckBox ChkSiembra 
            Caption         =   "Siembra"
            Height          =   285
            Left            =   180
            TabIndex        =   122
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Height          =   2880
         Left            =   -71565
         TabIndex        =   106
         Top             =   1380
         Width           =   3675
         Begin VB.ComboBox CboUnidad 
            Height          =   315
            Left            =   2685
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   705
            Width           =   870
         End
         Begin VB.TextBox TxtPreUni 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1515
            TabIndex        =   112
            Text            =   "0.00"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox TxtProd 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1515
            TabIndex        =   110
            Text            =   "0.00"
            Top             =   690
            Width           =   1125
         End
         Begin VB.TextBox TxtNumHec 
            Alignment       =   2  'Center
            Height          =   330
            Left            =   1500
            TabIndex        =   108
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
            TabIndex        =   119
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
            TabIndex        =   117
            Top             =   2235
            Width           =   1125
         End
         Begin VB.Label Label48 
            Caption         =   "Utilidad               :"
            Height          =   330
            Left            =   180
            TabIndex        =   116
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
            TabIndex        =   115
            Top             =   1875
            Width           =   1125
         End
         Begin VB.Label Label46 
            Caption         =   "Egresos              :"
            Height          =   330
            Left            =   165
            TabIndex        =   114
            Top             =   1890
            Width           =   1890
         End
         Begin VB.Label Label45 
            Caption         =   "Ingresos             :"
            Height          =   330
            Left            =   165
            TabIndex        =   113
            Top             =   1515
            Width           =   1890
         End
         Begin VB.Label Label44 
            Caption         =   "Precio Unitario   :"
            Height          =   330
            Left            =   150
            TabIndex        =   111
            Top             =   1125
            Width           =   1890
         End
         Begin VB.Label Label43 
            Caption         =   "Produccion        :"
            Height          =   330
            Left            =   135
            TabIndex        =   109
            Top             =   735
            Width           =   1890
         End
         Begin VB.Label Label42 
            Caption         =   "Hectareas          :"
            Height          =   330
            Left            =   135
            TabIndex        =   107
            Top             =   330
            Width           =   1890
         End
      End
      Begin VB.ComboBox CboTpoCul 
         Height          =   315
         Left            =   -73440
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   855
         Width           =   4335
      End
      Begin VB.Frame Frame8 
         Caption         =   "Rubro / Costo"
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
         Height          =   2865
         Left            =   -74730
         TabIndex        =   91
         Top             =   1380
         Width           =   3015
         Begin VB.TextBox TxtOtros 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   103
            Text            =   "0.00"
            Top             =   1800
            Width           =   1155
         End
         Begin VB.TextBox TxtPesticidas 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            TabIndex        =   101
            Text            =   "0.00"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox TxtInsumos 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1320
            TabIndex        =   99
            Text            =   "0.00"
            Top             =   1065
            Width           =   1155
         End
         Begin VB.TextBox TxtJornal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1305
            TabIndex        =   97
            Text            =   "0.00"
            Top             =   690
            Width           =   1155
         End
         Begin VB.TextBox TxtMaq 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1350
            TabIndex        =   95
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
            TabIndex        =   105
            Top             =   2235
            Width           =   1155
         End
         Begin VB.Label Label41 
            Caption         =   "Costo Total     :"
            Height          =   330
            Left            =   150
            TabIndex        =   104
            Top             =   2250
            Width           =   1230
         End
         Begin VB.Label Label40 
            Caption         =   "Otros               :"
            Height          =   330
            Left            =   135
            TabIndex        =   102
            Top             =   1815
            Width           =   1230
         End
         Begin VB.Label Label39 
            Caption         =   "Pesticidas       :"
            Height          =   330
            Left            =   135
            TabIndex        =   100
            Top             =   1455
            Width           =   1230
         End
         Begin VB.Label Label38 
            Caption         =   "Insumos          :"
            Height          =   330
            Left            =   135
            TabIndex        =   98
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label37 
            Caption         =   "Jornales           :"
            Height          =   330
            Left            =   135
            TabIndex        =   96
            Top             =   735
            Width           =   1260
         End
         Begin VB.Label Label36 
            Caption         =   "Maquinaria      :"
            Height          =   330
            Left            =   135
            TabIndex        =   94
            Top             =   390
            Width           =   1260
         End
      End
      Begin VB.Frame Frame7 
         Height          =   780
         Left            =   -74820
         TabIndex        =   88
         Top             =   4185
         Width           =   7125
         Begin RichTextLib.RichTextBox TxtComentariosBal 
            Height          =   465
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   6810
            _ExtentX        =   12012
            _ExtentY        =   820
            _Version        =   393217
            MaxLength       =   300
            TextRTF         =   $"frmFteIngresos.frx":09AA
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   75
            TabIndex        =   90
            Top             =   15
            Width           =   840
         End
      End
      Begin VB.TextBox TxtCargo 
         Height          =   285
         Left            =   -72750
         TabIndex        =   12
         Top             =   3135
         Width           =   4560
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   315
         Left            =   -72735
         TabIndex        =   11
         Top             =   2700
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   173801473
         CurrentDate     =   37014
      End
      Begin VB.Frame Frame5 
         Caption         =   "Flujo de Caja Mensual"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1560
         Left            =   -74820
         TabIndex        =   58
         Top             =   2625
         Width           =   7140
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
            TabIndex        =   84
            Text            =   "0.00"
            Top             =   765
            Width           =   1530
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
            Left            =   2010
            MaxLength       =   13
            TabIndex        =   82
            Text            =   "0.00"
            Top             =   780
            Width           =   1350
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
            Left            =   2010
            MaxLength       =   13
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   195
            Width           =   1350
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
            Left            =   1995
            MaxLength       =   13
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   480
            Width           =   1350
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
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   180
            Width           =   1530
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
            TabIndex        =   24
            Text            =   "0.00"
            Top             =   465
            Width           =   1530
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Egresos Familiares"
            Height          =   195
            Left            =   3540
            TabIndex        =   85
            Top             =   810
            Width           =   1305
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos Familiares"
            Height          =   195
            Left            =   180
            TabIndex        =   83
            Top             =   810
            Width           =   1335
         End
         Begin VB.Line Line2 
            X1              =   105
            X2              =   6840
            Y1              =   1125
            Y2              =   1125
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
            Left            =   3885
            TabIndex        =   67
            Top             =   1245
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
            Left            =   5475
            TabIndex        =   66
            Top             =   1185
            Width           =   1515
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ventas :"
            Height          =   195
            Left            =   195
            TabIndex        =   62
            Top             =   225
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Rec. de Ctas x Cobrar :"
            Height          =   195
            Left            =   195
            TabIndex        =   61
            Top             =   510
            Width           =   1650
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Costo de Ventas :"
            Height          =   195
            Left            =   3525
            TabIndex        =   60
            Top             =   195
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Otros Egresos :"
            Height          =   195
            Left            =   3525
            TabIndex        =   59
            Top             =   495
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Balance General"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2220
         Left            =   -74805
         TabIndex        =   40
         Top             =   390
         Width           =   7125
         Begin VB.TextBox txtDisponible 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1725
            MaxLength       =   13
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   780
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
            Left            =   1725
            MaxLength       =   13
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   1065
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
            Left            =   1725
            MaxLength       =   13
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1350
            Width           =   1335
         End
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
            Left            =   1950
            MaxLength       =   13
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   1680
            Width           =   1455
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
            Left            =   5460
            MaxLength       =   13
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   1365
            Width           =   1485
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
            Left            =   5460
            MaxLength       =   13
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   1080
            Width           =   1485
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
            Left            =   5460
            MaxLength       =   13
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   780
            Width           =   1485
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
            Left            =   240
            TabIndex        =   57
            Top             =   225
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Activo Circulante :"
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
            Left            =   240
            TabIndex        =   56
            Top             =   495
            Width           =   1590
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Activo Fijo :"
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
            Left            =   240
            TabIndex        =   55
            Top             =   1740
            Width           =   1035
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
            Left            =   3750
            TabIndex        =   54
            Top             =   195
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
            Left            =   3750
            TabIndex        =   53
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
            Left            =   3750
            TabIndex        =   52
            Top             =   1785
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Disponible :"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cuentas x Cobrar:"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   1095
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Inventario :"
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   1380
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
            Left            =   1950
            TabIndex        =   48
            Top             =   480
            Width           =   1485
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
            Left            =   1950
            TabIndex        =   47
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Prestamos CMAC-M"
            Height          =   195
            Left            =   3750
            TabIndex        =   46
            Top             =   1365
            Width           =   1410
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Otros Préstamos :"
            Height          =   195
            Left            =   3750
            TabIndex        =   45
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Proveedores :"
            Height          =   195
            Left            =   3750
            TabIndex        =   44
            Top             =   825
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
            TabIndex        =   43
            Top             =   435
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
            TabIndex        =   42
            Top             =   135
            Width           =   1320
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
            TabIndex        =   41
            Top             =   1710
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1830
         Left            =   -74040
         TabIndex        =   33
         Top             =   795
         Width           =   5835
         Begin VB.TextBox TxtIngCon 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   495
            Width           =   1155
         End
         Begin VB.TextBox txtOtroIng 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   795
            Width           =   1155
         End
         Begin VB.TextBox txtIngFam 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4455
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   210
            Width           =   1155
         End
         Begin VB.TextBox txtEgreFam 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4455
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   525
            Width           =   1155
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso Cony:"
            Height          =   195
            Left            =   180
            TabIndex        =   86
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
            TabIndex        =   81
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   1275
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Otros Ingresos:"
            Height          =   195
            Left            =   180
            TabIndex        =   37
            Top             =   870
            Width           =   1065
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos :"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblingreso 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso Cliente :"
            Height          =   195
            Left            =   2925
            TabIndex        =   35
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label lblEgreso 
            AutoSize        =   -1  'True
            Caption         =   "Egreso Familiar :"
            Height          =   195
            Left            =   2910
            TabIndex        =   34
            Top             =   585
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1200
         Left            =   -74100
         TabIndex        =   31
         Top             =   3435
         Width           =   6150
         Begin RichTextLib.RichTextBox Txtcomentarios 
            Height          =   915
            Left            =   90
            TabIndex        =   13
            Top             =   225
            Width           =   5910
            _ExtentX        =   10425
            _ExtentY        =   1614
            _Version        =   393217
            MaxLength       =   300
            TextRTF         =   $"frmFteIngresos.frx":0A2C
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   75
            TabIndex        =   32
            Top             =   15
            Width           =   840
         End
      End
      Begin VB.Label Label63 
         Caption         =   $"frmFteIngresos.frx":0AAE
         Height          =   1395
         Left            =   -66480
         TabIndex        =   201
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   255
         Left            =   600
         TabIndex        =   146
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblMonto2 
         AutoSize        =   -1  'True
         Caption         =   "Empresarial:"
         Height          =   255
         Left            =   5880
         TabIndex        =   141
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblMonto1 
         AutoSize        =   -1  'True
         Caption         =   "Personal:"
         Height          =   195
         Left            =   3360
         TabIndex        =   140
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label lblDescriEval 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   360
         TabIndex        =   139
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblTitulosEval 
         AutoSize        =   -1  'True
         Caption         =   "Titulos:"
         Height          =   195
         Left            =   3480
         TabIndex        =   138
         Top             =   480
         Width           =   510
      End
      Begin VB.Label lblTipoEval 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Evaluación:"
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label35 
         Caption         =   "Tipo de Cultivo  :"
         Height          =   330
         Left            =   -74730
         TabIndex        =   92
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cargo : "
         Height          =   195
         Left            =   -74025
         TabIndex        =   65
         Top             =   3165
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio :"
         Height          =   195
         Left            =   -74025
         TabIndex        =   64
         Top             =   2745
         Width           =   1185
      End
   End
   Begin VB.CommandButton CmdFteAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   180
      TabIndex        =   79
      Top             =   9135
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton CmdFteCancelar 
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   1320
      TabIndex        =   80
      Top             =   9120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label71 
      AutoSize        =   -1  'True
      Caption         =   "FLUJO DE CAJA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      TabIndex        =   129
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmFteIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim oPersona As UPersona_Cli ' COMDPersona.DCOMPersona   'DPersona
'Dim nIndice As Integer
'Dim nProcesoEjecutado As Integer '1 Nueva fte de Ingreso; 2 Editar fte de Ingreso ; 3 Consulta de Fte
'Dim vsUbiGeo As String
'Dim bEstadoCargando As Boolean
'Dim nProcesoActual As Integer
'Dim ldFecEval As Date
'
'Private Function ValidaDatosFuentesIngreso() As Boolean
'Dim CadTemp As String
'Dim i As Integer
'Dim J As Integer
'Dim nNumeFte As Integer

    
    
'
'    ValidaDatosFuentesIngreso = True
'
'    If TxFecEval.Visible Then
'        CadTemp = ValidaFecha(TxFecEval.Text)
'        If Len(CadTemp) > 0 Then
'            MsgBox CadTemp, vbInformation, "Aviso"
'            ValidaDatosFuentesIngreso = False
'            Exit Function
'        Else
'            If CmbFecha.ListCount > 0 Then
'                If CDate(ldFecEval) >= CDate(TxFecEval) Then
'                    MsgBox "No Puede Ingresar una Fecha de Evaluacion Igual o Menor a la Ultima Fecha de Evaluacion de la Fuente Ingreso", vbInformation, "Aviso"
'                    ValidaDatosFuentesIngreso = False
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
'
'    If CboTipoFte.ListIndex = -1 Then
'        MsgBox "No se ha Seleccionado el Tipo de Fuente", vbInformation, "Aviso"
'        CboTipoFte.SetFocus
'        ValidaDatosFuentesIngreso = False
'        Exit Function
'    End If
'
'    If CboMoneda.ListIndex = -1 Then
'        MsgBox "No se ha Seleccionado la Moneda", vbInformation, "Aviso"
'        CboMoneda.SetFocus
'        ValidaDatosFuentesIngreso = False
'        Exit Function
'    End If
'
'    If Len(Trim(TxtBRazonSoc.Text)) = 0 Then
'        MsgBox "No Ingresado la Razon Social", vbInformation, "Aviso"
'        TxtBRazonSoc.SetFocus
'        ValidaDatosFuentesIngreso = False
'        Exit Function
'    End If
'
'    CadTemp = ValidaFecha(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")))
'    If Len(Trim(CadTemp)) <> 0 Then
'        MsgBox CadTemp, vbInformation, "Aviso"
'        DTPFecIni.SetFocus
'        ValidaDatosFuentesIngreso = False
'        Exit Function
'    End If
'
'    'Valida la Fecha de Evaluacion
'    If nProcesoEjecutado = 1 Then
'        CadTemp = ValidaFecha(TxFecEval.Text)
'        If CadTemp <> "" Then
'            MsgBox CadTemp, vbInformation, "Aviso"
'            TxFecEval.SetFocus
'            ValidaDatosFuentesIngreso = False
'            Exit Function
'        End If
'    End If
'
'    'Valida que Exista una Unica Fuente
'    'If nProcesoEjecutado = 1 Then
'    '    nNumeFte = 0
'    '    For i = 0 To oPersona.NumeroFtesIngreso - 1
'    '        If oPersona.ObtenerFteIngFecEval(i) = CDate(TxFecEval.Text) Then
'    '            MsgBox "Ya existe una Fuente de Ingreso Con la Misma Fecha de Evaluacion", vbInformation, "Aviso"
'    '            TxFecEval.SetFocus
'    '        ValidaDatosFuentesIngreso = False
'    '        Exit Function
'    '        End If
'    '    Next i
'    'End If
'
'    'Valida si se Ingreso el Balance en Caso de ser Fuente Independiente
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoIndependiente Then
'        If CDbl(lblPatrimonio.Caption) <= 0 Then
'            MsgBox "Falta Ingresar el Balance de la Fuente de Ingreso"
'            SSTFuentes.Tab = 1
'            txtDisponible.SetFocus
'            ValidaDatosFuentesIngreso = False
'            Exit Function
'        End If
'    End If
'End Function
'Private Sub HabilitaCabecera(ByVal pnHabilitar As Boolean)
'    CboTipoFte.Enabled = pnHabilitar
'    CboMoneda.Enabled = pnHabilitar
'    TxtBRazonSoc.Enabled = pnHabilitar
'    TxtRazSocDescrip.Enabled = pnHabilitar
'    TxtRazSocDirecc.Enabled = pnHabilitar
'    TxtRazSocTelef.Enabled = pnHabilitar
'    CmdUbigeo.Enabled = pnHabilitar
'End Sub
'Private Sub HabilitaCostoProd(ByVal pnHabilitar As Boolean)
'
'    CboTpoCul.Enabled = pnHabilitar
'    TxtMaq.Enabled = pnHabilitar
'    TxtJornal.Enabled = pnHabilitar
'    TxtInsumos.Enabled = pnHabilitar
'    TxtPesticidas.Enabled = pnHabilitar
'    TxtOtros.Enabled = pnHabilitar
'    TxtNumHec.Enabled = pnHabilitar
'    TxtProd.Enabled = pnHabilitar
'    CboUnidad.Enabled = pnHabilitar
'    TxtPreUni.Enabled = pnHabilitar
'    ChkSiembra.Enabled = pnHabilitar
'    ChkMantenimiento.Enabled = pnHabilitar
'    chkDesAgricola.Enabled = pnHabilitar
'    chkOtros.Enabled = pnHabilitar
'    chkCosecha.Enabled = pnHabilitar
'End Sub
'
'Private Sub HabilitaIngresosEgresos(ByVal pnHabilitar As Boolean)
'    txtOtroIng.Enabled = pnHabilitar
'    txtIngFam.Enabled = pnHabilitar
'    txtEgreFam.Enabled = pnHabilitar
'    DTPFecIni.Enabled = pnHabilitar
'    TxtCargo.Enabled = pnHabilitar
'    Txtcomentarios.Enabled = pnHabilitar
'    TxtIngCon.Enabled = pnHabilitar
'    ChkCostoProd.Enabled = pnHabilitar
'End Sub
'
'
'Private Sub HabilitaBalance(ByVal HabBalance As Boolean)
'    TxtComentariosBal.Enabled = HabBalance
'    Label6.Enabled = HabBalance
'    lblActivo.Enabled = HabBalance
'    Label9.Enabled = HabBalance
'    lblPasPatrim.Enabled = HabBalance
'    Label7.Enabled = HabBalance
'    lblActCirc.Enabled = HabBalance
'    Label10.Enabled = HabBalance
'    lblPasivo.Enabled = HabBalance
'    Label12.Enabled = HabBalance
'    txtDisponible.Enabled = HabBalance
'    Label19.Enabled = HabBalance
'    txtProveedores.Enabled = HabBalance
'    Label13.Enabled = HabBalance
'    txtcuentas.Enabled = HabBalance
'    Label18.Enabled = HabBalance
'    txtOtrosPrest.Enabled = HabBalance
'    Label14.Enabled = HabBalance
'    txtInventario.Enabled = HabBalance
'    Label17.Enabled = HabBalance
'    txtPrestCmact.Enabled = HabBalance
'    Label8.Enabled = HabBalance
'    txtactivofijo.Enabled = HabBalance
'    Label11.Enabled = HabBalance
'    lblPatrimonio.Enabled = HabBalance
'    Label15.Enabled = HabBalance
'    txtVentas.Enabled = HabBalance
'    Label5.Enabled = HabBalance
'    txtcompras.Enabled = HabBalance
'    Label20.Enabled = HabBalance
'    txtrecuperacion.Enabled = HabBalance
'    Label4.Enabled = HabBalance
'    txtOtrosEgresos.Enabled = HabBalance
'    'Label47.Enabled = HabBalance
'    'Label49.Enabled = HabBalance
'    Label31.Enabled = HabBalance
'    Label32.Enabled = HabBalance
'    'lblIngresosB.Enabled = HabBalance
'    'lblEgresosB.Enabled = HabBalance
'    TxtBalIngFam.Enabled = HabBalance
'    TxtBalEgrFam.Enabled = HabBalance
'    'txtOtrosEgresos.Enabled = HabBalance
'    LblSaldoIngEgr.Enabled = HabBalance
'    Frame4.Enabled = HabBalance
'    Frame5.Enabled = HabBalance
'
'    If Trim(Right(CboTipoFte.Text, 2)) = "2" Then
'        Me.SSTFuentes.TabVisible(1) = True
'        Me.SSTFuentes.TabVisible(0) = False
'        Me.SSTFuentes.Tab = 1
'    Else
'        Me.SSTFuentes.TabVisible(1) = False
'        Me.SSTFuentes.TabVisible(0) = True
'        Me.SSTFuentes.Tab = 0
'    End If
'
'End Sub
'
'Private Sub CargaControles()
''    Call CargaComboConstante(gPersFteIngresoTipo, CboTipoFte)
''    Call CargaComboConstante(gMoneda, CboMoneda)
''    Call CargaComboConstante(1046, CboTpoCul)
''    Call CargaComboConstante(1045, CboUnidad)
'
''Dim oPersona As COMDpersona.DCOMPersonas
''Set oPersona = New COMDpersona.DCOMPersonas
''Dim rsMoneda As ADODB.Recordset
''Dim rsTipoFte As ADODB.Recordset
''Dim rsTipoCul As ADODB.Recordset
''Dim rsUnidad As ADODB.Recordset
''
''Call oPersona.CargarDatosObjetosFteIngreso(rsTipoFte, rsMoneda, rsTipoCul, rsUnidad)
''
''Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
''Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
''Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
''Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
''
''Set rsMoneda = Nothing
''Set rsTipoFte = Nothing
''Set rsTipoCul = Nothing
''Set rsUnidad = Nothing
''Set oPersona = Nothing
'
'End Sub
'
'Private Sub CargaDatosFteIngreso(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli, '                                ByVal rsFIDep As ADODB.Recordset, ByVal rsFIInd As ADODB.Recordset, ByVal rsFICos As ADODB.Recordset, '                                Optional ByVal pnFteDetalle As Integer = -1)
'
'Dim nUltFte As Integer
'
'    LblCliente.Caption = PstaNombre(poPersona.NombreCompleto)
'    ChkCostoProd.value = poPersona.ObtenerFteIngbCostoProd(pnIndice)
'    TxtBRazonSoc.Text = poPersona.ObtenerFteIngFuente(pnIndice)
'    LblRazonSoc.Caption = poPersona.ObtenerFteIngRazonSocial(pnIndice)
'    CboTipoFte.ListIndex = IndiceListaCombo(CboTipoFte, poPersona.ObtenerFteIngTipo(pnIndice))
'    TxtRazSocDescrip.Text = poPersona.ObtenerFteIngRazSocDescrip(pnIndice)
'    TxtRazSocDirecc.Text = poPersona.ObtenerFteIngRazSocDirecc(pnIndice)
'    TxtRazSocTelef.Text = poPersona.ObtenerFteIngRazSocTelef(pnIndice)
'    vsUbiGeo = poPersona.ObtenerFteIngRazSocUbiGeo(pnIndice)
'    If CInt(poPersona.ObtenerFteIngTipo(pnIndice)) = gPersFteIngresoTipoDependiente Then
'        If poPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) = 0 Then
'            Call poPersona.RecuperaFtesIngresoDependiente(pnIndice, rsFIDep)
'        End If
'        Call HabilitaBalance(False)
'    Else
'        If poPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) = 0 Then
'            Call poPersona.RecuperaFtesIngresoIndependiente(pnIndice, rsFIInd)
'        End If
'        Call HabilitaBalance(True)
'    End If
'
'    If ChkCostoProd.value = 1 Then
'        Call poPersona.RecuperaFtesIngresoCostosProd(pnIndice, rsFICos)
'        If pnFteDetalle = -1 Then
'            nUltFte = poPersona.ObtenerFteIngIngresoNumeroCostoProd(pnIndice) - 1
'        Else
'            nUltFte = pnFteDetalle
'        End If
'
'        Call UbicaCombo(CboTpoCul, poPersona.ObtenerCostoProdnTpoCultivo(pnIndice, nUltFte))
'        TxtMaq.Text = poPersona.ObtenerCostoProdnMaquinaria(pnIndice, nUltFte)
'        TxtJornal.Text = poPersona.ObtenerCostoProdnJornales(pnIndice, nUltFte)
'        TxtInsumos.Text = poPersona.ObtenerCostoProdnInsumos(pnIndice, nUltFte)
'        TxtPesticidas.Text = poPersona.ObtenerCostoProdnPesticidas(pnIndice, nUltFte)
'        TxtOtros.Text = poPersona.ObtenerCostoProdnOtros(pnIndice, nUltFte)
'
'        'LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas.Text) + CDbl(TxtOtros.Text), "#0.00")
'
'        TxtNumHec.Text = poPersona.ObtenerCostoProdnHectareas(pnIndice, nUltFte)
'        TxtProd.Text = poPersona.ObtenerCostoProdnProduccion(pnIndice, nUltFte)
'        TxtPreUni.Text = poPersona.ObtenerCostoProdnPreUni(pnIndice, nUltFte)
'        Call PutOfChecked(ChkSiembra, poPersona.ObtenerCostoProdnSiembra(pnIndice, nUltFte))
'        Call PutOfChecked(ChkMantenimiento, poPersona.ObtenerCostoProdnMantenimiento(pnIndice, nUltFte))
'        Call PutOfChecked(chkDesAgricola, poPersona.ObtenerCostoProdnDesaAgricola(pnIndice, nUltFte))
'        Call PutOfChecked(chkOtros, poPersona.ObtenerCostoProdnOtros(pnIndice, nUltFte))
'        Call PutOfChecked(chkCosecha, poPersona.ObtenerCostoProdnCosecha(pnIndice, nUltFte))
'
'        'LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
'
'        'LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
'
'        'LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) * CDbl(LblCostoEgr.Caption), "#0.00")
'
'        Call UbicaCombo(CboUnidad, poPersona.ObtenerCostoProdnUniProd(pnIndice, nUltFte))
'    End If
'
'    CboMoneda.ListIndex = IndiceListaCombo(CboMoneda, poPersona.ObtenerFteIngMoneda(pnIndice))
'    'Carga Ingresos y Egresos
'    DTPFecIni.value = CDate(Format(poPersona.ObtenerFteIngInicioFuente(pnIndice), "dd/mm/yyyy"))
'    TxtCargo.Text = poPersona.ObtenerFteIngCargo(pnIndice)
'    Txtcomentarios.Text = poPersona.ObtenerFteIngComentarios(pnIndice)
'    TxtComentariosBal.Text = poPersona.ObtenerFteIngComentarios(pnIndice)
'
'    If poPersona.ObtenerFteIngIngresoTipo(pnIndice) = gPersFteIngresoTipoDependiente Then
'        If pnFteDetalle = -1 Then
'            nUltFte = poPersona.ObtenerFteIngIngresoNumeroFteDep(pnIndice) - 1
'        Else
'            nUltFte = pnFteDetalle
'        End If
'        TxtIngCon.Text = Format(poPersona.ObtenerFteIngIngresoConyugue(pnIndice, nUltFte), "#0.00")
'        txtIngFam.Text = Format(poPersona.ObtenerFteIngIngresoFam(pnIndice, nUltFte), "#0.00")
'        txtOtroIng.Text = Format(poPersona.ObtenerFteIngIngresoOtros(pnIndice, nUltFte), "#0.00")
'        'LblIngresos.Caption = Format(poPersona.ObtenerFteIngIngresos(pnIndice, nUltFte), "#0.00")
'        txtEgreFam.Text = Format(poPersona.ObtenerFteIngGastoFam(pnIndice, nUltFte), "#0.00")
'        'lblSaldo.Caption = CDbl(txtIngFam.Text) + CDbl(txtOtroIng.Text) + CDbl(LblIngresos.Caption) - CDbl(txtEgreFam.Text)
'
'    Else
'        If pnFteDetalle = -1 Then
'            nUltFte = poPersona.ObtenerFteIngIngresoNumeroFteIndep(pnIndice) - 1
'        Else
'            nUltFte = pnFteDetalle
'        End If
'        'Carga el Balance
'        txtDisponible.Text = Format(poPersona.ObtenerFteIngActivoDisp(pnIndice, nUltFte), "#0.00")
'        txtcuentas.Text = Format(poPersona.ObtenerFteIngCtasxCob(pnIndice, nUltFte), "#0.00")
'        txtInventario.Text = Format(poPersona.ObtenerFteIngInventario(pnIndice, nUltFte), "#0.00")
'        txtactivofijo.Text = Format(poPersona.ObtenerFteIngActivoFijo(pnIndice, nUltFte), "#0.00")
'
'        txtProveedores.Text = Format(poPersona.ObtenerFteIngProveedores(pnIndice, nUltFte), "#0.00")
'        txtOtrosPrest.Text = Format(poPersona.ObtenerFteIngOtrosCreditos(pnIndice, nUltFte), "#0.00")
'        txtPrestCmact.Text = Format(poPersona.ObtenerFteIngCreditosCmact(pnIndice, nUltFte), "#0.00")
'
'        txtVentas.Text = Format(poPersona.ObtenerFteIngVentas(pnIndice, nUltFte), "#0.00")
'        txtrecuperacion.Text = Format(poPersona.ObtenerFteIngRecupCtasxCobrar(pnIndice, nUltFte), "#0.00")
'        txtcompras.Text = Format(poPersona.ObtenerFteIngComprasMercad(pnIndice, nUltFte), "#0.00")
'        txtOtrosEgresos.Text = Format(poPersona.ObtenerFteIngOtrosEgresos(pnIndice, nUltFte), "#0.00")
'        TxtBalIngFam.Text = Format(poPersona.ObtenerFteIngBalIngFam(pnIndice, nUltFte), "#0.00")
'        TxtBalEgrFam.Text = Format(poPersona.ObtenerFteIngBalEgrFam(pnIndice, nUltFte), "#0.00")
'    End If
'End Sub
'Sub PutOfChecked(ByRef cChecked As CheckBox, ByVal pintValor)
'    If pintValor = 1 Then
'        cChecked.value = 1
'    Else
'        cChecked.value = 0
'    End If
'End Sub
'Private Sub LimpiaFormulario()
'
'    LblCliente.Caption = oPersona.NombreCompleto
'    TxtBRazonSoc.Text = ""
'    CboTipoFte.ListIndex = -1
'    CboMoneda.ListIndex = -1
'    LblIngresos.Caption = "0.00"
'    txtIngFam.Text = "0.00"
'    txtOtroIng.Text = "0.00"
'    txtEgreFam.Text = "0.00"
'    DTPFecIni.value = gdFecSis
'    TxtCargo.Text = ""
'    Txtcomentarios.Text = ""
'    TxtComentariosBal.Text = ""
'    lblActivo.Caption = "0.00"
'    lblActCirc.Caption = "0.00"
'    txtDisponible.Text = "0.00"
'    txtcuentas.Text = "0.00"
'    txtInventario.Text = "0.00"
'    txtactivofijo.Text = "0.00"
'    lblPasPatrim.Caption = "0.00"
'    lblPasivo.Caption = "0.00"
'    txtProveedores.Text = "0.00"
'    txtOtrosPrest.Text = "0.00"
'    txtPrestCmact.Text = "0.00"
'    lblPatrimonio.Caption = "0.00"
'    txtVentas.Text = "0.00"
'    txtrecuperacion.Text = "0.00"
'    txtcompras.Text = "0.00"
'    txtOtrosEgresos.Text = "0.00"
'End Sub
'
'Private Sub LimpiaFuentesIngreso()
'    LblIngresos.Caption = "0.00"
'    txtIngFam.Text = "0.00"
'    txtOtroIng.Text = "0.00"
'    txtEgreFam.Text = "0.00"
'    TxtIngCon.Text = "0.00"
'    DTPFecIni.value = gdFecSis
'    TxtCargo.Text = ""
'    Txtcomentarios.Text = ""
'    TxtComentariosBal.Text = ""
'    lblActivo.Caption = "0.00"
'    lblActCirc.Caption = "0.00"
'    txtDisponible.Text = "0.00"
'    txtcuentas.Text = "0.00"
'    txtInventario.Text = "0.00"
'    txtactivofijo.Text = "0.00"
'    lblPasPatrim.Caption = "0.00"
'    lblPasivo.Caption = "0.00"
'    txtProveedores.Text = "0.00"
'    txtOtrosPrest.Text = "0.00"
'    txtPrestCmact.Text = "0.00"
'    lblPatrimonio.Caption = "0.00"
'    txtVentas.Text = "0.00"
'    txtrecuperacion.Text = "0.00"
'    txtcompras.Text = "0.00"
'    txtOtrosEgresos.Text = "0.00"
'
'    CboTpoCul.ListIndex = 0
'    TxtMaq.Text = "0.00"
'    TxtJornal.Text = "0.00"
'    TxtInsumos.Text = "0.00"
'    TxtPesticidas.Text = "0.00"
'    TxtOtros.Text = "0.00"
'    LblCostoTotal.Caption = "0.00"
'    TxtProd.Text = "0.00"
'    TxtPreUni.Text = "0.00"
'    LblCostosIng.Caption = "0.00"
'    LblCostoEgr.Caption = "0.00"
'    LblCostosUtil.Caption = "0.00"
'    TxtNumHec.Text = "0"
'    ChkSiembra.value = 0
'    ChkMantenimiento.value = 0
'    chkDesAgricola.value = 0
'    chkOtros.value = 0
'    chkCosecha.value = 0
'End Sub
'
'Public Sub Editar(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli)
'    Set oPersona = poPersona
'    nIndice = pnIndice
'    nProcesoEjecutado = 2
'    bEstadoCargando = True
'    'Call CargaControles
'    'Call CargaDatosFteIngreso(pnIndice, poPersona)
'    Call CargarDatos(pnIndice, poPersona)
'
'    CmdAceptar.Visible = True
'    CmdSalirCancelar.Caption = "&Cancelar"
'    bEstadoCargando = False
'    CmbFecha.Visible = True
'    TxFecEval.Visible = False
'    Call CargaComboFechaEval
'    HabilitaCabecera False
'    HabilitaBalance False
'    HabilitaIngresosEgresos False
'    HabilitaCostoProd False
'    CmdAceptar.Visible = False
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        SSTFuentes.TabVisible(1) = False
'        SSTFuentes.TabVisible(0) = True
'        SSTFuentes.Tab = 0
'    Else
'        SSTFuentes.TabVisible(0) = False
'        SSTFuentes.TabVisible(1) = True
'        SSTFuentes.Tab = 1
'    End If
'    frmFteIngresos.Show 1
'End Sub
'
'Public Sub NuevaFteIngreso(ByRef poPersona As UPersona_Cli, Optional ByVal pnFteIndice As Integer = -1)
'Dim oPersTemp As UPersona_Cli ' COMDPersona.DCOMPersona 'DPersona
'    bEstadoCargando = True
'    Set oPersona = poPersona
'    'Call CargaControles
'    Call CargarDatos(-1)
'    Call LimpiaFormulario
'    nProcesoEjecutado = 1
'    CmdAceptar.Visible = True
'    CmdSalirCancelar.Caption = "&Cancelar"
'    bEstadoCargando = False
'    CmbFecha.Visible = False
'    TxFecEval.Visible = True
'    CmdNuevo.Enabled = False
'    CmdEditar.Enabled = False
'    CmdEliminar.Enabled = False
'    TxFecEval.Text = Format(gdFecSis, "dd/mm/yyyy")
'    CmbFecha.Clear
'    'If pnFteIndice <> -1 Then
'    '    TxtBRazonSoc.Text = poPersona.ObtenerFteIngFuente(pnFteIndice)
'    '    LblRazonSoc.Caption = poPersona.ObtenerFteIngRazonSocial(pnFteIndice)
'    '    CboTipoFte.ListIndex = IndiceListaCombo(CboTipoFte, CInt(poPersona.ObtenerFteIngTipo(pnFteIndice)))
'    '    CboMoneda.ListIndex = IndiceListaCombo(CboMoneda, CInt(poPersona.ObtenerFteIngMoneda(pnFteIndice)))
'    '    TxtRazSocDescrip.Text = poPersona.ObtenerFteIngRazSocDescrip(pnFteIndice)
'    '    TxtRazSocDirecc.Text = poPersona.ObtenerFteIngRazSocDirecc(pnFteIndice)
'    '    TxtRazSocTelef.Text = poPersona.ObtenerFteIngRazSocTelef(pnFteIndice)
'    '    Set oPersTemp = New DPersona
'    '    Call oPersTemp.RecuperaPersona(TxtBRazonSoc.Text)
'    '    vsUbiGeo = oPersTemp.UbicacionGeografica
'    '    Set oPersTemp = Nothing
'    'End If
'    If ChkCostoProd.value = vbChecked Then
'        ' se procede a ver lo de costos de producccion
'        txtVentas.Enabled = False
'        txtcompras.Enabled = False
'    Else
'        txtVentas.Enabled = True
'        txtcompras.Enabled = True
'    End If
'    frmFteIngresos.Show 1
'End Sub
'
'Public Sub CargaComboFechaEval()
'Dim i As Integer
'    CmbFecha.Clear
'    If oPersona.ObtenerFteIngIngresoTipo(nIndice) = gPersFteIngresoTipoDependiente Then
'        For i = 0 To oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1
'            CmbFecha.AddItem oPersona.ObtenerFteIngFecEval(nIndice, i, CInt(Right(CboTipoFte.Text, 2)))
'        Next i
'    Else
'        For i = 0 To oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1
'            CmbFecha.AddItem oPersona.ObtenerFteIngFecEval(nIndice, i, CInt(Right(CboTipoFte.Text, 2)))
'        Next i
'    End If
'    bEstadoCargando = True
'    If oPersona.ObtenerFteIngIngresoTipo(nIndice) = gPersFteIngresoTipoDependiente Then
'        CmbFecha.ListIndex = IndiceListaCombo(CmbFecha, Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1, gPersFteIngresoTipoDependiente), "dd/mm/yyyy"))
'        ldFecEval = Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1, gPersFteIngresoTipoDependiente), "dd/mm/yyyy")
'    Else
'        CmbFecha.ListIndex = IndiceListaCombo(CmbFecha, Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1, gPersFteIngresoTipoIndependiente), "dd/mm/yyyy"))
'        ldFecEval = Format(oPersona.ObtenerFteIngFecEval(nIndice, oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1, gPersFteIngresoTipoIndependiente), "dd/mm/yyyy")
'    End If
'    bEstadoCargando = False
'End Sub
'
'Public Sub ConsultarFuenteIngreso(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli)
'    Set oPersona = poPersona
'    nIndice = pnIndice
'    nProcesoEjecutado = 3
'    bEstadoCargando = True
'    'Call CargaControles
'    'Call CargaDatosFteIngreso(pnIndice, poPersona)
'    Call CargarDatos(pnIndice, poPersona)
'    CmdSalirCancelar.Caption = "&Salir"
'    Call HabilitaCabecera(False)
'    Call HabilitaBalance(False)
'    Call HabilitaIngresosEgresos(False)
'    CmbFecha.Visible = True
'    TxFecEval.Visible = False
'    Call CargaComboFechaEval
'    bEstadoCargando = False
'    CmdNuevo.Enabled = False
'    CmdEditar.Enabled = False
'    CmdEliminar.Enabled = False
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        SSTFuentes.TabVisible(1) = False
'        SSTFuentes.TabVisible(0) = True
'        SSTFuentes.Tab = 0
'    Else
'        SSTFuentes.TabVisible(0) = False
'        SSTFuentes.TabVisible(1) = True
'        SSTFuentes.Tab = 1
'    End If
'    frmFteIngresos.Show 1
'End Sub
'
'Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If TxtBRazonSoc.Enabled Then
'            TxtBRazonSoc.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub CboTipoFte_Click()
'    If Trim(Right(CboTipoFte.Text, 15)) = gPersFteIngresoTipoDependiente Then
'        Call HabilitaBalance(False)
'       ChkCostoProd.value = 0
'       ChkCostoProd.Enabled = False
''        TxtBRazonSoc.Enabled = True
'    Else
'        Call HabilitaBalance(True)
'        ChkCostoProd.Enabled = True
''        TxtBRazonSoc.Text = ""
''        TxtBRazonSoc.Enabled = False
''        LblRazonSoc.Caption = ""
'    End If
'End Sub
'
'Private Sub CboTipoFte_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        CboMoneda.SetFocus
'    End If
'End Sub
'
'Private Sub CboUnidad_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        TxtPreUni.SetFocus
'    End If
'End Sub
'
'Private Sub ChkCostoProd_Click()
'    If ChkCostoProd.value = 1 Then
'        SSTFuentes.TabVisible(2) = True
'        SSTFuentes.Tab = 2
'        txtVentas.Enabled = False
'        txtcompras.Enabled = False
'        CboTpoCul.Enabled = True
'        TxtMaq.Enabled = True
'        TxtJornal.Enabled = True
'        TxtInsumos.Enabled = True
'        TxtPesticidas.Enabled = True
'        TxtOtros.Enabled = True
'        TxtNumHec.Enabled = True
'        TxtProd.Enabled = True
'        CboUnidad.Enabled = True
'        TxtPreUni.Enabled = True
'
'    Else
'        SSTFuentes.TabVisible(2) = False
'        'Se Agrego
'        txtVentas.Enabled = True
'        txtcompras.Enabled = True
'    End If
'End Sub
'
'Private Sub CmbFecha_Click()
'Dim oPersonaD  As COMDPersona.DCOMPersona
'
'    If bEstadoCargando Then
'        Exit Sub
'    End If
'    If CmbFecha.ListCount <= 0 Then
'        MsgBox "No Existe Fuente de Ingreso para Mostrar", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If Len(Trim(TxtBRazonSoc.Text)) <= 0 Then
'        MsgBox "Falta Ingresar la Razon Social", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If CmbFecha.ListIndex = -1 Then
'        MsgBox "Seleccione una Fecha de Evaluacion del Credito", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    'Call CargaDatosFteIngreso(nIndice, oPersona, CmbFecha.ListIndex)
'    Call CargarDatos(nIndice, oPersona, CmbFecha.ListIndex, False)
'    'Verifica si ya esta asignado a un Credito
'    HabilitaCabecera False
'    HabilitaIngresosEgresos False
'
'    HabilitaBalance False
'
'
'    If Trim(Right(CboTipoFte.Text, 2)) = "2" Then
'        Me.SSTFuentes.TabVisible(1) = True
'        Me.SSTFuentes.TabVisible(0) = False
'        Me.SSTFuentes.Tab = 1
'    Else
'        Me.SSTFuentes.TabVisible(1) = False
'        Me.SSTFuentes.TabVisible(0) = True
'        Me.SSTFuentes.Tab = 0
'    End If
'
'    Set oPersonaD = New COMDPersona.DCOMPersona
'    Call oPersonaD.RecuperaFtesdeIngreso(oPersona.PersCodigo)
'    If oPersonaD.FuenteIngresoAsignadaACredito(nIndice, CDate(CmbFecha.Text)) Then
'        HabilitaIngresosEgresos False
'        HabilitaBalance False
'        CmdNuevo.Enabled = False
'        CmdEditar.Enabled = False
'        CmdEliminar.Enabled = False
'    Else
'        CmdNuevo.Enabled = True
'        CmdEditar.Enabled = True
'        CmdEliminar.Enabled = True
'    End If
'    Set oPersonaD = Nothing
'    If nProcesoEjecutado = 3 Then
'        CmdNuevo.Enabled = False
'        CmdEditar.Enabled = False
'        CmdEliminar.Enabled = False
'    End If
'End Sub
'
'Private Sub CmbFecha_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'            TxtIngCon.SetFocus
'        Else
'            txtDisponible.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub CmdAceptar_Click()
''Dim oPersonaNeg As npersona
'
'    If Not ValidaDatosFuentesIngreso Then
'        Exit Sub
'    End If
'
'    If nProcesoEjecutado = 1 Then
'        Call oPersona.AdicionaFteIngreso(CInt(Right(CboTipoFte.Text, 2)), IIf(ChkCostoProd.value = vbChecked, True, False))
'        nIndice = oPersona.NumeroFtesIngreso - 1
'    Else
'        If oPersona.ObtenerFteIngTipoAct(nIndice) <> PersFilaNueva Then
'            Call oPersona.ActualizarFteIngTipoAct(PersFilaModificada, nIndice)
'            'If SSTFuentes.TabVisible(2) = True Then
'             If ChkCostoProd.value = Checked Then
'                Call oPersona.ActualizarCostoProdTipoAct(PersFilaModificada, nIndice, 0)
'            End If
'        End If
'    End If
'
'    If frmPersona.bNuevaPersona = False Then
'        oPersona.TipoActualizacion = PersFilaModificada
'    End If
'
'    Call oPersona.ActualizarFteIngTipoFuente(Trim(Right(CboTipoFte.Text, 20)), nIndice)
'    Call oPersona.ActualizarFteIngMoneda(Trim(Right(CboMoneda.Text, 20)), nIndice)
'    Call oPersona.ActualizarFteIngFuenteIngreso(Trim(TxtBRazonSoc.Text), LblRazonSoc.Caption, nIndice)
'    Call oPersona.ActualizarFteIngComentarios(Txtcomentarios.Text, nIndice)
'    Call oPersona.ActualizarFteIngComentarios(TxtComentariosBal.Text, nIndice)
'    Call oPersona.ActualizarFteIngCargo(TxtCargo.Text, nIndice)
'    Call oPersona.ActualizarFteIngRazSocDirecc(Trim(TxtRazSocDirecc.Text), nIndice)
'    Call oPersona.ActualizarFteIngRazSocDescrip(Trim(TxtRazSocDescrip.Text), nIndice)
'    Call oPersona.ActualizarFteIngRazSocTelef(Trim(TxtRazSocTelef.Text), nIndice)
'    Call oPersona.ActualizarFteIngRazSocUbiGeo(Trim(vsUbiGeo), nIndice)
'    Call oPersona.ActualizarFteIngFechaInicioFuente(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")), nIndice)
'    If SSTFuentes.TabVisible(2) = True Then
'         Call oPersona.ActualizarFteIngbCostoProd(ChkCostoProd.value, nIndice)
'    End If
'
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        Call oPersona.ActualizarFteIngIngresos(CDbl(txtIngFam.Text) + CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
'        Call oPersona.ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngIngOtros(CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
'        Call oPersona.ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, 0)
'        If TxFecEval.Visible Then
'            Call oPersona.ActualizarFteDepIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
'        End If
'    Else
'        Call oPersona.ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice, 0)
'        If TxFecEval.Visible Then
'            Call oPersona.ActualizarFteIndIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
'        End If
'        Call oPersona.ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice, 0)
'        Call oPersona.ActualizarFteIngBalIngFam(CDbl(TxtBalIngFam.Text), nIndice, 0)
'        Call oPersona.ActualizarFteIngBalEgrFam(CDbl(TxtBalEgrFam.Text), nIndice, 0)
'
'    End If
'    ' se verifica que el tab de produccion  este visible
'
'    'If SSTFuentes.TabVisible(2) = True Then
'    If ChkCostoProd.value = vbChecked Then
'    'Actualiza Costos de Produccion
'        If CmbFecha.Visible = True Then
'            Call oPersona.ActualizarCostosdFecEval(CDate(IIf(IsDate(CmbFecha.Text), CmbFecha.Text, Date)), nIndice, 0)
'        Else
'            Call oPersona.ActualizarCostosdFecEval(CDate(IIf(IsDate(TxFecEval.Text), TxFecEval.Text, Date)), nIndice, 0)
'        End If
'        Call oPersona.ActualizarCostosnTpoCultivo(CInt(Right(CboTpoCul.Text, 2)), nIndice, 0)
'        Call oPersona.ActualizarCostosnMaquinaria(CDbl(TxtMaq.Text), nIndice, 0)
'        Call oPersona.ActualizarCostosnJornales(CDbl(TxtJornal), nIndice, 0)
'        Call oPersona.ActualizarCostosnInsumos(CDbl(TxtInsumos), nIndice, 0)
'        Call oPersona.ActualizarCostosnPesticidas(CDbl(TxtPesticidas), nIndice, 0)
'        Call oPersona.ActualizarCostosnOtros(CDbl(TxtOtros), nIndice, 0)
'        Call oPersona.ActualizarCostosnHectareas(CDbl(TxtNumHec), nIndice, 0)
'        Call oPersona.ActualizarCostosnProduccion(CDbl(TxtProd), nIndice, 0)
'        Call oPersona.ActualizarCostosnPreUni(CDbl(TxtPreUni), nIndice, 0)
'        Call oPersona.ActualizarCostosnUniProd(CInt(Right(CboUnidad.Text, 2)), nIndice, 0)
'        Call oPersona.ActualizarCostosTpoCultivo(Mid(CboTpoCul.Text, 1, Len(CboTpoCul.Text) - 2), nIndice, 0)
'        Call oPersona.ActualizarCostossUniProd(Mid(CboUnidad.Text, 1, Len(CboUnidad.Text) - 2), nIndice, 0)
'
'        Call oPersona.ActualizarCostonSiembra(GetValueOfChecked(ChkSiembra), nIndice, 0)
'        Call oPersona.ActualizarCostonMantenimiento(GetValueOfChecked(ChkMantenimiento), nIndice, 0)
'        Call oPersona.ActualizarCostonDesaAgricola(GetValueOfChecked(chkDesAgricola), nIndice, 0)
'        Call oPersona.ActualizarCostonCosecha(GetValueOfChecked(chkCosecha), nIndice, 0)
'        Call oPersona.ActualizarCostonPlanOtros(GetValueOfChecked(chkOtros), nIndice, 0)
'
'   End If
'
'    If nProcesoEjecutado = 1 Then
'        'Set oPersonaNeg = New COMNPersona.NCOMPersona    ' COMDPersona.DCOMPersona 'npersona
'        Call oPersona.ChequeoFuenteIngreso(nIndice, CInt(Right(CboTipoFte, 2)), 0)
'        'Set oPersonaNeg = Nothing
'    End If
'    Call cmdImprimir_Click
'    Unload Me
'End Sub
'
'Private Sub cmdEditar_Click()
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        HabilitaBalance False
'        HabilitaIngresosEgresos True
'        SSTFuentes.Tab = 0
'    Else
'        HabilitaIngresosEgresos False
'        HabilitaBalance True
'        SSTFuentes.Tab = 1
'    End If
'
'    If Me.ChkCostoProd.value = 1 Then
'        HabilitaCostoProd True
'        txtVentas.Enabled = False
'        txtcompras.Enabled = False
'    Else
'        txtVentas.Enabled = True
'        txtcompras.Enabled = True
'    End If
'
'    HabilitaMantenimiento False
'    CmbFecha.Enabled = False
'    nProcesoActual = 2
'    '***Modificacion LMMD******************
'    Frame6.Enabled = True
'    TxtBRazonSoc.Enabled = True
'    TxtRazSocDescrip.Enabled = True
'    TxtRazSocDirecc.Enabled = True
'    TxtRazSocTelef.Enabled = True
'    CmdUbigeo.Enabled = True
'End Sub
'
'Private Sub cmdeliminar_Click()
'Dim oPersonaD As COMDPersona.DCOMPersona
'
'    If MsgBox("Se va a Eliminar la Fuente de Ingreso de Fecha :" & Me.CmbFecha.Text & ", Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'    Set oPersonaD = New COMDPersona.DCOMPersona
'    If oPersonaD.FuenteIngresoAsignadaACredito(nIndice, CDate(CmbFecha.Text)) Then
'        MsgBox "La Fuente de Ingreso No se Puede Eliminar porque esta Asignada a un Credito", vbInformation, "Aviso"
'        Set oPersonaD = Nothing
'        Exit Sub
'    End If
'    Set oPersonaD = Nothing
'    Call oPersona.ActualizarFteIngTipoActdetalle(PersFilaEliminda, nIndice, CmbFecha.ListIndex)
'    Call CmbFecha.RemoveItem(CmbFecha.ListIndex)
'End Sub
'
'Private Sub CmdFteAceptar_Click()
'Dim nIndiceAct As Integer
''Dim oPersonaNeg As UPersona_Cli
'    If Not ValidaDatosFuentesIngreso Then
'        Exit Sub
'    End If
'
'    'Si se va a adicionar una nueva fuente
'    If nProcesoActual = 1 Then
'        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'            Call oPersona.AdicionaFteIngresoDependiente(nIndice)
'            nIndiceAct = oPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice) - 1
'        Else
'            Call oPersona.AdicionaFteIngresoIndependiente(nIndice)
'            nIndiceAct = oPersona.ObtenerFteIngIngresoNumeroFteIndep(nIndice) - 1
'            If ChkCostoProd.value = 1 Then
'                Call oPersona.AdicionaFteIngresoCostoProd(nIndice)
'            End If
'        End If
'
'        Call oPersona.ActualizarFteIngTipoActdetalle(PersFilaNueva, nIndice, nIndiceAct)
'        CmbFecha.AddItem TxFecEval.Text
'    Else
'        nIndiceAct = CmbFecha.ListIndex
'        Call oPersona.ActualizarFteIngTipoActdetalle(PersFilaModificada, nIndice, nIndiceAct)
'    End If
'    'Si se va a actualizar una fte de ingreso
'    If nProcesoActual = 1 Or nProcesoActual = 2 Then
'        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'            Call oPersona.ActualizarFteIngIngresos(CDbl(txtIngFam.Text) + CDbl(txtOtroIng.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngIngOtros(CDbl(txtOtroIng.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, nIndiceAct)
'            If TxFecEval.Visible Then
'                Call oPersona.ActualizarFteDepIngFecEval(CDate(TxFecEval.Text), nIndice, nIndiceAct)
'            End If
'        Else
'            Call oPersona.ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice, nIndiceAct)
'            If TxFecEval.Visible Then
'                Call oPersona.ActualizarFteIndIngFecEval(CDate(TxFecEval.Text), nIndice, nIndiceAct)
'            End If
'            Call oPersona.ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngBalIngFam(CDbl(TxtBalIngFam.Text), nIndice, nIndiceAct)
'            Call oPersona.ActualizarFteIngBalEgrFam(CDbl(TxtBalEgrFam.Text), nIndice, nIndiceAct)
'        End If
'    End If
'
'    'Actualiza Costos de produccion
'    If SSTFuentes.TabVisible(2) = True Then
'        If TxFecEval.Visible Then
'            Call oPersona.ActualizarCostosdFecEval(CDate(TxFecEval.Text), nIndice, nIndiceAct)
'        Else
'            Call oPersona.ActualizarCostosdFecEval(CDate(CmbFecha.Text), nIndice, nIndiceAct)
'        End If
'        Call oPersona.ActualizarCostosnTpoCultivo(CInt(Right(CboTpoCul.Text, 2)), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnMaquinaria(CDbl(TxtMaq.Text), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnJornales(CDbl(TxtJornal), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnInsumos(CDbl(TxtInsumos), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnPesticidas(CDbl(TxtPesticidas), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnOtros(CDbl(TxtOtros), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnHectareas(CDbl(TxtNumHec), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnProduccion(CDbl(TxtProd), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnPreUni(CDbl(TxtPreUni), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosnUniProd(CInt(Right(CboUnidad.Text, 2)), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostosTpoCultivo(Mid(CboTpoCul.Text, 1, Len(CboTpoCul.Text) - 2), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostossUniProd(Mid(CboUnidad.Text, 1, Len(CboUnidad.Text) - 2), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostonSiembra(GetValueOfChecked(ChkSiembra), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostonMantenimiento(GetValueOfChecked(ChkMantenimiento), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostonDesaAgricola(GetValueOfChecked(chkDesAgricola), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostonPlanOtros(GetValueOfChecked(chkOtros), nIndice, nIndiceAct)
'        Call oPersona.ActualizarCostonCosecha(GetValueOfChecked(chkCosecha), nIndice, nIndiceAct)
'  End If
' '   Set oPersonaNeg = New UPersona_Cli ' COMDPersona.DCOMPersona  'npersona
' '   Call oPersonaNeg.ChequeoFuenteIngreso(nIndice, CInt(Right(CboTipoFte, 2)), nIndiceAct)
' '   Set oPersonaNeg = Nothing
'    Call oPersona.ChequeoFuenteIngreso(nIndice, CInt(Right(CboTipoFte, 2)), nIndiceAct)
'    Call cmdImprimir_Click
'
'    HabilitaBalance False
'    HabilitaCostoProd False
'    HabilitaMantenimiento True
'    CmbFecha.Visible = True
'    TxFecEval.Visible = False
'    CmbFecha.Enabled = True
'    CmdAceptar.Visible = True
'
'End Sub
'
'Function GetValueOfChecked(ByVal pCChecked As CheckBox) As Integer
'        If pCChecked.value = vbChecked Then
'           GetValueOfChecked = 1
'        Else
'            GetValueOfChecked = 0
'        End If
'End Function
'
'Private Sub HabilitaMantenimiento(ByVal pbHabilita As Boolean)
'    CmdNuevo.Visible = pbHabilita
'    CmdEditar.Visible = pbHabilita
'    CmdEliminar.Visible = pbHabilita
'    CmdFteAceptar.Visible = Not pbHabilita
'    CmdFteCancelar.Visible = Not pbHabilita
'End Sub
'
'Private Sub CmdFteCancelar_Click()
'    HabilitaBalance False
'    HabilitaIngresosEgresos False
'    HabilitaMantenimiento True
'    HabilitaCostoProd False
'    CmbFecha.Visible = True
'    CmbFecha.Enabled = True
'    TxFecEval.Visible = False
'    CmbFecha_Click
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        SSTFuentes.TabVisible(1) = False
'        SSTFuentes.TabVisible(0) = True
'        SSTFuentes.Tab = 0
'    Else
'        SSTFuentes.TabVisible(0) = False
'        SSTFuentes.TabVisible(1) = True
'        SSTFuentes.Tab = 1
'    End If
'End Sub
'
'Private Sub cmdImprimir_Click()
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
'        sCadImp = oPersonaD.GenerarImpresionFteIngresoIndependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
'    End If
'    Set oPersonaD = Nothing
'    previo.Show sCadImp, "Evaluacion de Fuentes de Ingreso", False
'    Set oPrev = Nothing
'End Sub
'
'Sub LlenarDatosFteIngreso(ByVal poPersona As COMDPersona.DCOMPersona)
'
'Dim nIndex As Integer
'
'With poPersona
'     'If nProcesoEjecutado = 1 Then
'    While nIndex <= nIndice
'        Call .AdicionaFteIngreso(CInt(Right(CboTipoFte.Text, 2)), IIf(ChkCostoProd.value = vbChecked, True, False))
'        nIndex = nIndex + 1
'    Wend
'    '   nIndice = oPersona.NumeroFtesIngreso - 1
'    'Else
'    '    If oPersona.ObtenerFteIngTipoAct(nIndice) <> PersFilaNueva Then
'    '        Call .ActualizarFteIngTipoAct(PersFilaModificada, nIndice)
'    '         If ChkCostoProd.value = Checked Then
'    '            Call .ActualizarCostoProdTipoAct(PersFilaModificada, nIndice, 0)
'    '        End If
'    '    End If
'        If ChkCostoProd.value = vbChecked Then
'            Call .AdicionaFteIngresoCostoProd(nIndice)
'        End If
'    'End If
'
'    'Datos Adicionales no incluidos para el Reporte
'    .NombreCompleto = oPersona.NombreCompleto
'    .PersCodigo = oPersona.PersCodigo
'
'    If frmPersona.bNuevaPersona = False Then
'        .TipoActualizacion = PersFilaModificada
'    End If
'
'    Call .ActualizarFteIngTipoFuente(Trim(Right(CboTipoFte.Text, 20)), nIndice)
'    Call .ActualizarFteIngMoneda(Trim(Right(CboMoneda.Text, 20)), nIndice)
'    Call .ActualizarFteIngFuenteIngreso(Trim(TxtBRazonSoc.Text), LblRazonSoc.Caption, nIndice)
'    Call .ActualizarFteIngComentarios(Txtcomentarios.Text, nIndice)
'    Call .ActualizarFteIngComentarios(TxtComentariosBal.Text, nIndice)
'    Call .ActualizarFteIngCargo(TxtCargo.Text, nIndice)
'    Call .ActualizarFteIngRazSocDirecc(Trim(TxtRazSocDirecc.Text), nIndice)
'    Call .ActualizarFteIngRazSocDescrip(Trim(TxtRazSocDescrip.Text), nIndice)
'    Call .ActualizarFteIngRazSocTelef(Trim(TxtRazSocTelef.Text), nIndice)
'    Call .ActualizarFteIngRazSocUbiGeo(Trim(vsUbiGeo), nIndice)
'    Call .ActualizarFteIngFechaInicioFuente(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")), nIndice)
'
''   Datos de la Fuente como Persona, no incluidos en el Reporte
'    Call .ActualizarFteRuc(oPersona.ObtenerFteIngRuc(nIndice), nIndice)
'    Call .ActualizarFteFecInicioAct(oPersona.ObtenerFteIngFecInicioAct(nIndice), nIndice)
'    Call .ActualizarFteTipoPersJur(oPersona.ObtenerFteIngTipoPersJur(nIndice), nIndice)
'    Call .ActualizarFteTelefono(oPersona.ObtenerFteIngTelefono(nIndice), nIndice)
'    Call .ActualizarFteCIUU(oPersona.ObtenerFteIngCIUU(nIndice), nIndice)
'    Call .ActualizarFteCondicionDomic(oPersona.ObtenerFteIngCondicionDomic(nIndice), nIndice)
'    Call .ActualizarFteMagnitudEmp(oPersona.ObtenerFteIngMagnitudEmp(nIndice), nIndice)
'    Call .ActualizarFteNroEmpleados(oPersona.ObtenerFteIngNroEmpleados(nIndice), nIndice)
'    Call .ActualizarFteDireccion(oPersona.ObtenerFteIngDireccion(nIndice), nIndice)
'    Call .ActualizarFteDpto(oPersona.ObtenerFteIngDpto(nIndice), nIndice)
'    Call .ActualizarFteProv(oPersona.ObtenerFteIngProv(nIndice), nIndice)
'    Call .ActualizarFteDist(oPersona.ObtenerFteIngDist(nIndice), nIndice)
'    Call .ActualizarFteZona(oPersona.ObtenerFteIngZona(nIndice), nIndice)
'
'    If SSTFuentes.TabVisible(2) = True Then
'         Call .ActualizarFteIngbCostoProd(ChkCostoProd.value, nIndice)
'    End If
'
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        Call .AdicionaFteIngresoDependiente(nIndice)
'        Call .ActualizarFteIngIngresos(CDbl(txtIngFam.Text) + CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
'        Call .ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice, 0)
'        Call .ActualizarFteIngIngOtros(CDbl(IIf(txtOtroIng.Text = "", 0, txtOtroIng.Text)), nIndice, 0)
'        Call .ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice, 0)
'        Call .ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, 0)
'        If TxFecEval.Text <> "__/__/____" Then
'            Call .ActualizarFteDepIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
'        End If
'    Else
'        Call .AdicionaFteIngresoIndependiente(nIndice)
'        Call .ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice, 0)
'        Call .ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice, 0)
'        Call .ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice, 0)
'        Call .ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice, 0)
'        Call .ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice, 0)
'        Call .ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice, 0)
'        Call .ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice, 0)
'        Call .ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice, 0)
'        Call .ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice, 0)
'        Call .ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice, 0)
'        Call .ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice, 0)
'        If TxFecEval.Visible Then
'            Call .ActualizarFteIndIngFecEval(CDate(TxFecEval.Text), nIndice, 0)
'        End If
'        Call .ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice, 0)
'        Call .ActualizarFteIngBalIngFam(CDbl(TxtBalIngFam.Text), nIndice, 0)
'        Call .ActualizarFteIngBalEgrFam(CDbl(TxtBalEgrFam.Text), nIndice, 0)
'
'    End If
'    ' se verifica que el tab de produccion  este visible
'
'    'If SSTFuentes.TabVisible(2) = True Then
'    If ChkCostoProd.value = vbChecked Then
'    'Actualiza Costos de Produccion
'        If CmbFecha.Visible = True Then
'            Call .ActualizarCostosdFecEval(CDate(IIf(IsDate(CmbFecha.Text), CmbFecha.Text, Date)), nIndice, 0)
'        Else
'            Call .ActualizarCostosdFecEval(CDate(IIf(IsDate(TxFecEval.Text), TxFecEval.Text, Date)), nIndice, 0)
'        End If
'        Call .ActualizarCostosnTpoCultivo(CInt(Right(CboTpoCul.Text, 2)), nIndice, 0)
'        Call .ActualizarCostosnMaquinaria(CDbl(TxtMaq.Text), nIndice, 0)
'        Call .ActualizarCostosnJornales(CDbl(TxtJornal), nIndice, 0)
'        Call .ActualizarCostosnInsumos(CDbl(TxtInsumos), nIndice, 0)
'        Call .ActualizarCostosnPesticidas(CDbl(TxtPesticidas), nIndice, 0)
'        Call .ActualizarCostosnOtros(CDbl(TxtOtros), nIndice, 0)
'        Call .ActualizarCostosnHectareas(CDbl(TxtNumHec), nIndice, 0)
'        Call .ActualizarCostosnProduccion(CDbl(TxtProd), nIndice, 0)
'        Call .ActualizarCostosnPreUni(CDbl(TxtPreUni), nIndice, 0)
'        Call .ActualizarCostosnUniProd(CInt(Right(CboUnidad.Text, 2)), nIndice, 0)
'        Call .ActualizarCostosTpoCultivo(Mid(CboTpoCul.Text, 1, Len(CboTpoCul.Text) - 2), nIndice, 0)
'        Call .ActualizarCostossUniProd(Mid(CboUnidad.Text, 1, Len(CboUnidad.Text) - 2), nIndice, 0)
'
'        Call .ActualizarCostonSiembra(GetValueOfChecked(ChkSiembra), nIndice, 0)
'        Call .ActualizarCostonMantenimiento(GetValueOfChecked(ChkMantenimiento), nIndice, 0)
'        Call .ActualizarCostonDesaAgricola(GetValueOfChecked(chkDesAgricola), nIndice, 0)
'        Call .ActualizarCostonCosecha(GetValueOfChecked(chkCosecha), nIndice, 0)
'        Call .ActualizarCostonPlanOtros(GetValueOfChecked(chkOtros), nIndice, 0)
'
'   End If
'
'End With
'
'End Sub
'
'Sub CargarDatos(ByVal pnIndice As Integer, Optional ByVal poPersona As UPersona_Cli = Nothing, '                Optional ByVal pnFteDetalle As Integer = -1, '                Optional ByVal pbCargarControles As Boolean = True)
'
'Dim oPersona As COMDPersona.DCOMPersonas
'Set oPersona = New COMDPersona.DCOMPersonas
'Dim rsMoneda As ADODB.Recordset
'Dim rsTipoFte As ADODB.Recordset
'Dim rsTipoCul As ADODB.Recordset
'Dim rsUnidad As ADODB.Recordset
'Dim rsFIDep As ADODB.Recordset
'Dim rsFIInd As ADODB.Recordset
'Dim rsFICos As ADODB.Recordset
'
'If pnIndice = -1 Then
'    Call oPersona.CargarDatosObjetosFteIngreso(rsTipoFte, rsMoneda, rsTipoCul, rsUnidad, rsFIDep, rsFIInd, rsFICos)
'    If pbCargarControles Then
'        Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
'        Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
'        Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
'        Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
'    End If
'Else
'    Call oPersona.CargarDatosObjetosFteIngreso(rsTipoFte, rsMoneda, rsTipoCul, rsUnidad, rsFIDep, rsFIInd, rsFICos, poPersona.ObtenerFteIngcNumFuente(pnIndice))
'    If pbCargarControles Then
'        Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
'        Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
'        Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
'        Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
'    End If
'    Call CargaDatosFteIngreso(pnIndice, poPersona, rsFIDep, rsFIInd, rsFICos, pnFteDetalle)
'End If
'
'Set rsMoneda = Nothing
'Set rsTipoFte = Nothing
'Set rsTipoCul = Nothing
'Set rsUnidad = Nothing
'Set rsFIDep = Nothing
'Set rsFIInd = Nothing
'Set rsFICos = Nothing
'Set oPersona = Nothing
'End Sub
'
'
'Private Sub cmdNuevo_Click()
'    TxFecEval.Text = "__/__/____"
'    CmbFecha.Visible = False
'    TxFecEval.Visible = True
'    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'        HabilitaBalance False
'        HabilitaIngresosEgresos True
'        SSTFuentes.Tab = 0
'    Else
'        HabilitaIngresosEgresos False
'        HabilitaBalance True
'        SSTFuentes.Tab = 1
'    End If
'    If Me.ChkCostoProd.value = 1 Then
'        HabilitaCostoProd True
'    End If
'    Call LimpiaFuentesIngreso
'    nProcesoActual = 1
'    HabilitaMantenimiento False
'    ChkCostoProd.Enabled = True
'End Sub
'
'Private Sub CmdSalirCancelar_Click()
'    Unload Me
'End Sub
'
'Private Sub CmdUbigeo_Click()
'    vsUbiGeo = Right(frmUbicacionGeo.Inicio(vsUbiGeo), 12)
'End Sub
'
'Private Sub DTPFecIni_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        TxtCargo.SetFocus
'    End If
'End Sub
'
'Private Sub Form_Load()
'    CentraForm Me
'    Me.Top = 0
'    Me.Left = (Screen.Width - Me.Width) / 2
'    'me.Left = 600
'    bEstadoCargando = False
'    nProcesoActual = 0
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
'    SSTFuentes.TabVisible(2) = False
'End Sub
'
'
'Private Sub LblCostoEgr_Change()
'    If ChkCostoProd.value = vbChecked Then
'        txtcompras = LblCostoEgr
'    End If
'End Sub
'
'Private Sub LblCostosIng_Change()
'    If ChkCostoProd.value = vbChecked Then
'        txtVentas = LblCostosIng
'    End If
'End Sub
'
'Private Sub TxFecEval_GotFocus()
'    fEnfoque TxFecEval
'End Sub
'
'Private Sub TxFecEval_KeyPress(KeyAscii As Integer)
''Dim oPersona As UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona
'
''Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona 'DPersona
'    If KeyAscii = 13 Then
'        If CDate(ldFecEval) >= CDate(TxFecEval) Then
'            MsgBox "No Puede Ingresar una Fecha de Evaluacion Igual o Menor a la Ultima Fecha de Evaluacion de la Fuente Ingreso", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        If Trim(CboTipoFte.Text) <> "" Then
'            If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'                TxtIngCon.SetFocus
'            Else
'                txtDisponible.SetFocus
'            End If
'        End If
'    End If
''Set oPersona = Nothing
'
'End Sub
'
'Private Sub TxFecEval_LostFocus()
'Dim sCad As String
'
'    sCad = ValidaFecha(TxFecEval.Text)
'    If Len(Trim(sCad)) > 0 Then
'        MsgBox sCad, vbInformation, "Aviso"
'        TxFecEval.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtactivofijo_Change()
'   lblActivo.Caption = Format(CDbl(IIf(Trim(lblActCirc.Caption) = "", "0", lblActCirc.Caption)) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
'   lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'   lblPasPatrim.Caption = Format(CDbl(lblPatrimonio.Caption) + CDbl(lblPasivo.Caption), "#0.00")
'
'End Sub
'
'Private Sub txtactivofijo_GotFocus()
'    fEnfoque txtactivofijo
'End Sub
'
'Private Sub txtactivofijo_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtactivofijo, KeyAscii, 12)
'        If KeyAscii = 13 Then
'            txtProveedores.SetFocus
'        End If
'End Sub
'
'Private Sub txtactivofijo_LostFocus()
'    txtactivofijo.Text = Format(IIf(Trim(txtactivofijo.Text) = "", 0, txtactivofijo.Text), "#0.00")
'End Sub
'
'Private Sub TxtBalEgrFam_Change()
'LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) '            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) '            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) '            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) '            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
'End Sub
'
'Private Sub TxtBalEgrFam_GotFocus()
'    fEnfoque TxtBalEgrFam
'End Sub
'
'Private Sub TxtBalEgrFam_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtBalEgrFam, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        If CmdFteAceptar.Visible Then
'            CmdFteAceptar.SetFocus
'        Else
'            CmdAceptar.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub TxtBalEgrFam_LostFocus()
'    If Len(Trim(TxtBalEgrFam.Text)) = 0 Then
'        TxtBalEgrFam.Text = "0.00"
'    End If
'    TxtBalEgrFam.Text = Format(TxtBalEgrFam.Text, "#0.00")
'End Sub
'
'Private Sub TxtBalIngFam_Change()
'LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) '            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) '            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) '            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) '            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
'End Sub
'
'Private Sub TxtBalIngFam_GotFocus()
'    fEnfoque TxtBalIngFam
'End Sub
'
'Private Sub TxtBalIngFam_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtBalIngFam, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        TxtBalEgrFam.SetFocus
'    End If
'End Sub
'
'Private Sub TxtBalIngFam_LostFocus()
'    If Len(Trim(TxtBalIngFam.Text)) = 0 Then
'        TxtBalIngFam.Text = "0.00"
'    End If
'    TxtBalIngFam.Text = Format(TxtBalIngFam.Text, "#0.00")
'End Sub
'
'Private Sub TxtBRazonSoc_EmiteDatos()
'Dim oPersTemp As UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona
'
'    LblRazonSoc.Caption = Trim(TxtBRazonSoc.psDescripcion)
'    TxtRazSocDescrip = LblRazonSoc.Caption
'    TxtRazSocDirecc.Text = TxtBRazonSoc.sPersDireccion
'    Set oPersTemp = New UPersona_Cli ' COMDPersona.DCOMPersona 'DPersona
'    Call oPersTemp.RecuperaPersona(TxtBRazonSoc.Text)
'    vsUbiGeo = oPersTemp.UbicacionGeografica
'    Set oPersTemp = Nothing
'    TxtRazSocDescrip.SetFocus
'End Sub
'
'Private Sub TxtCargo_GotFocus()
'    fEnfoque TxtCargo
'End Sub
'
'Private Sub TxtCargo_KeyPress(KeyAscii As Integer)
'    KeyAscii = Letras(KeyAscii)
'    If KeyAscii = 13 Then
'        Txtcomentarios.SetFocus
'    End If
'End Sub
'
'Private Sub txtcompras_Change()
'LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) '            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) '            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) '            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) '            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
'End Sub
'
'Private Sub txtcompras_GotFocus()
'    fEnfoque txtcompras
'End Sub
'
'Private Sub txtcompras_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtcompras, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtOtrosEgresos.SetFocus
'    End If
'End Sub
'
'Private Sub txtcompras_LostFocus()
'    txtcompras.Text = Format(IIf(Trim(txtcompras.Text) = "", "0.00", txtcompras.Text), "#0.00")
'End Sub
'
'Private Sub txtcuentas_Change()
'    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) '            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) '            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
'
'    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
'    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
'    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'
'End Sub
'
'Private Sub txtcuentas_GotFocus()
'    fEnfoque txtcuentas
'End Sub
'
'Private Sub txtcuentas_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtcuentas, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtInventario.SetFocus
'    End If
'End Sub
'
'Private Sub txtcuentas_LostFocus()
'    txtcuentas.Text = Format(IIf(Trim(txtcuentas.Text) = "", 0, txtcuentas.Text), "#0.00")
'End Sub
'
'Private Sub txtDisponible_Change()
'    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) '            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) '            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
'
'    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
'    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
'    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'
'End Sub
'
'Private Sub txtDisponible_GotFocus()
'    fEnfoque txtDisponible
'End Sub
'
'Private Sub txtDisponible_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtDisponible, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtcuentas.SetFocus
'    End If
'End Sub
'
'Private Sub txtDisponible_LostFocus()
'    txtDisponible.Text = Format(IIf(Trim(txtDisponible.Text) = "", 0, txtDisponible.Text), "#0.00")
'End Sub
'
'Private Sub txtEgreFam_Change()
'    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) '            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
'    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
'
'    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00"))
'End Sub
'
'Private Sub txtEgreFam_GotFocus()
'    fEnfoque txtEgreFam
'End Sub
'
'Private Sub txtEgreFam_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtEgreFam, KeyAscii)
'    If KeyAscii = 13 Then
'        DTPFecIni.SetFocus
'    End If
'End Sub
'
'Private Sub txtEgreFam_LostFocus()
'    If Len(Trim(txtEgreFam.Text)) > 0 Then
'        txtEgreFam.Text = Format(txtEgreFam.Text, "#0.00")
'    Else
'        txtEgreFam.Text = "0.00"
'    End If
'End Sub
'
'Private Sub TxtIngCon_Change()
'    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) '            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
'    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
'
'    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) '
'End Sub
'
'Private Sub TxtIngCon_GotFocus()
'    fEnfoque TxtIngCon
'End Sub
'
'Private Sub TxtIngCon_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtIngCon, KeyAscii)
'    If KeyAscii = 13 Then
'        txtOtroIng.SetFocus
'    End If
'End Sub
'
'Private Sub TxtIngCon_LostFocus()
'    If Len(Trim(TxtIngCon.Text)) = 0 Then
'        TxtIngCon.Text = "0.00"
'    End If
'    TxtIngCon.Text = Format(TxtIngCon.Text, "#0.00")
'End Sub
'
'Private Sub txtIngFam_Change()
'    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) '            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
'    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
'
'    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) '
'End Sub
'
'Private Sub txtIngFam_GotFocus()
'    fEnfoque txtIngFam
'End Sub
'
'Private Sub txtIngFam_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtIngFam, KeyAscii)
'    If KeyAscii = 13 Then
'        txtEgreFam.SetFocus
'    End If
'End Sub
'
'Private Sub txtIngFam_LostFocus()
'    If Len(Trim(txtIngFam.Text)) > 0 Then
'        txtIngFam.Text = Format(txtIngFam.Text, "#0.00")
'    Else
'        txtIngFam.Text = "0.00"
'    End If
'End Sub
'
'Private Sub TxtInsumos_Change()
'    If Trim(TxtInsumos.Text) = "" Then
'        TxtInsumos.Text = "0.00"
'    End If
'    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
'End Sub
'
'Private Sub TxtInsumos_GotFocus()
'    fEnfoque TxtInsumos
'End Sub
'
'Private Sub TxtInsumos_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtInsumos, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtPesticidas.SetFocus
'    End If
'End Sub
'
'Private Sub TxtInsumos_LostFocus()
'    If Trim(TxtInsumos.Text) = "" Then
'        TxtInsumos.Text = "0.00"
'    End If
'    TxtInsumos.Text = Format(TxtInsumos.Text, "#0.00")
'End Sub
'
'Private Sub txtInventario_Change()
'    lblActCirc.Caption = Format(CDbl(IIf(Trim(txtDisponible.Text) = "", "0", txtDisponible.Text)) '            + CDbl(IIf(Trim(txtcuentas.Text) = "", "0", txtcuentas.Text)) '            + CDbl(IIf(Trim(txtInventario.Text) = "", "0", txtInventario.Text)), "#0.00")
'
'    lblActivo.Caption = Format(CDbl(lblActCirc.Caption) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text)), "#0.00")
'    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
'    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'
'End Sub
'
'Private Sub txtInventario_GotFocus()
'    fEnfoque txtInventario
'End Sub
'
'Private Sub txtInventario_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtInventario, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtactivofijo.SetFocus
'    End If
'End Sub
'
'Private Sub txtInventario_LostFocus()
'    txtInventario.Text = Format(IIf(Trim(txtInventario.Text) = "", 0, txtInventario.Text), "#0.00")
'End Sub
'
'Private Sub TxtJornal_Change()
'    If Trim(TxtJornal.Text) = "" Then
'        TxtJornal.Text = "0.00"
'    End If
'    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
'End Sub
'
'Private Sub TxtJornal_GotFocus()
'    fEnfoque TxtJornal
'End Sub
'
'Private Sub TxtJornal_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtJornal, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtInsumos.SetFocus
'    End If
'End Sub
'
'Private Sub TxtJornal_LostFocus()
'    If Trim(TxtJornal.Text) = "" Then
'        TxtJornal.Text = "0.00"
'    End If
'    TxtJornal.Text = Format(TxtJornal.Text, "#0.00")
'End Sub
'
'Private Sub TxtMaq_Change()
'    If Trim(TxtMaq.Text) = "" Then
'        TxtMaq.Text = "0.00"
'    End If
'    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
'End Sub
'
'Private Sub TxtMaq_GotFocus()
'    fEnfoque TxtMaq
'End Sub
'
'Private Sub TxtMaq_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtMaq, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtJornal.SetFocus
'    End If
'End Sub
'
'Private Sub TxtMaq_LostFocus()
'    If Trim(TxtMaq.Text) = "" Then
'        TxtMaq.Text = "0.00"
'    End If
'    TxtMaq.Text = Format(TxtMaq.Text, "#0.00")
'End Sub
'
'Private Sub TxtNumHec_Change()
'    If Trim(TxtNumHec.Text) = "" Then
'        TxtNumHec.Text = "0"
'    End If
'    LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
'    LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
'    LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) - CDbl(LblCostoEgr.Caption), "#0.00")
'
'End Sub
'
'Private Sub TxtNumHec_GotFocus()
'    fEnfoque TxtNumHec
'End Sub
'
'Private Sub TxtNumHec_KeyPress(KeyAscii As Integer)
''    KeyAscii = NumerosDecimales(TxtNumHec, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtProd.SetFocus
'    End If
'End Sub
'
'Private Sub TxtNumHec_LostFocus()
'    If Trim(TxtNumHec.Text) = "" Then
'        TxtNumHec.Text = "0"
'    End If
'    TxtNumHec.Text = Format(TxtNumHec.Text, "#0.0")
'End Sub
'
'Private Sub txtOtroIng_Change()
'
'    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00")) '            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
'    lblSaldo.Caption = Format(lblSaldo.Caption, "#0.00")
'
'    LblIngresos.Caption = CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) '            + CDbl(Format(IIf(Trim(TxtIngCon.Text) = "", "0", TxtIngCon.Text), "#0.00"))
'
'End Sub
'
'Private Sub txtOtroIng_GotFocus()
'    fEnfoque txtOtroIng
'End Sub
'
'Private Sub txtOtroIng_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtOtroIng, KeyAscii)
'    If KeyAscii = 13 Then
'        txtIngFam.SetFocus
'    End If
'End Sub
'
'Private Sub txtOtroIng_LostFocus()
'    If Len(Trim(txtOtroIng.Text)) > 0 Then
'        txtOtroIng.Text = Format(txtOtroIng.Text, "#0.00")
'    Else
'        txtOtroIng.Text = "0.00"
'    End If
'End Sub
'
'Private Sub TxtOtros_Change()
'    If Trim(TxtOtros.Text) = "" Then
'        TxtOtros.Text = "0.00"
'    End If
'    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
'End Sub
'
'Private Sub TxtOtros_GotFocus()
'    fEnfoque TxtOtros
'End Sub
'
'Private Sub TxtOtros_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtOtros, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtNumHec.SetFocus
'    End If
'End Sub
'
'Private Sub TxtOtros_LostFocus()
'    If Trim(TxtOtros.Text) = "" Then
'        TxtOtros.Text = "0.00"
'    End If
'    TxtOtros.Text = Format(TxtOtros.Text, "#0.00")
'End Sub
'
'Private Sub txtOtrosEgresos_Change()
'    LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) '            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) '            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) '            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) '            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
'End Sub
'
'Private Sub txtOtrosEgresos_GotFocus()
'    fEnfoque txtOtrosEgresos
'End Sub
'
'Private Sub txtOtrosEgresos_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtOtrosEgresos, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        TxtBalIngFam.SetFocus
'    End If
'End Sub
'
'Private Sub txtOtrosEgresos_LostFocus()
'    txtOtrosEgresos.Text = Format(IIf(Trim(txtOtrosEgresos.Text) = "", "0.00", txtOtrosEgresos.Text), "#0.00")
'End Sub
'
'Private Sub txtOtrosPrest_Change()
'    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) '            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) '            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
'
'    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
'
'End Sub
'
'Private Sub txtOtrosPrest_GotFocus()
'    fEnfoque txtOtrosPrest
'End Sub
'
'Private Sub txtOtrosPrest_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtOtrosPrest, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtPrestCmact.SetFocus
'    End If
'End Sub
'
'Private Sub txtOtrosPrest_LostFocus()
'    If Len(Trim(txtOtrosPrest.Text)) > 0 Then
'        txtOtrosPrest.Text = Format(txtOtrosPrest.Text, "#0.00")
'    Else
'        txtOtrosPrest.Text = "0.00"
'    End If
'End Sub
'
'Private Sub TxtPesticidas_Change()
'    If Trim(TxtPesticidas.Text) = "" Then
'        TxtPesticidas.Text = "0.00"
'    End If
'    LblCostoTotal.Caption = Format(CDbl(TxtMaq.Text) + CDbl(TxtJornal.Text) + CDbl(TxtInsumos.Text) + CDbl(TxtPesticidas) + CDbl(TxtOtros), "#0.00")
'End Sub
'
'Private Sub TxtPesticidas_GotFocus()
'    fEnfoque TxtPesticidas
'End Sub
'
'Private Sub TxtPesticidas_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtPesticidas, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtOtros.SetFocus
'    End If
'End Sub
'
'Private Sub TxtPesticidas_LostFocus()
'    If Trim(TxtPesticidas.Text) = "" Then
'        TxtPesticidas.Text = "0.00"
'    End If
'    TxtPesticidas.Text = Format(TxtPesticidas.Text, "#0.00")
'End Sub
'
'Private Sub txtPrestCmact_Change()
'    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) '            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) '            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
'
'    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
'
'End Sub
'
'Private Sub txtPrestCmact_GotFocus()
'    fEnfoque txtPrestCmact
'End Sub
'
'Private Sub txtPrestCmact_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtPrestCmact, KeyAscii, 12)
'    If KeyAscii = 13 Then
'     If txtVentas.Enabled = True Then
'        txtVentas.SetFocus
'    End If
'    End If
'End Sub
'
'Private Sub txtPrestCmact_LostFocus()
'    If Len(Trim(txtPrestCmact.Text)) > 0 Then
'        txtPrestCmact.Text = Format(txtPrestCmact.Text, "#0.00")
'    Else
'        txtPrestCmact.Text = "0.00"
'    End If
'End Sub
'
'Private Sub TxtPreUni_Change()
'
'    If Trim(TxtPreUni.Text) = "" Then
'        TxtPreUni.Text = "0.00"
'    End If
'
'    LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
'    LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
'    LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) - CDbl(LblCostoEgr.Caption), "#0.00")
'End Sub
'
'Private Sub TxtPreUni_GotFocus()
'    fEnfoque TxtPreUni
'End Sub
'
'Private Sub TxtPreUni_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtPreUni, KeyAscii)
'    If KeyAscii = 13 Then
'        CboTpoCul.SetFocus
'    End If
'End Sub
'
'Private Sub TxtPreUni_LostFocus()
'    If Trim(TxtPreUni.Text) = "" Then
'        TxtPreUni.Text = "0.00"
'    End If
'    TxtPreUni.Text = Format(TxtPreUni.Text, "#0.00")
'End Sub
'
'Private Sub txtProd_Change()
'    If Trim(TxtProd.Text) = "" Then
'        TxtProd.Text = "0.00"
'    End If
'    LblCostosIng.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(TxtProd.Text) * CDbl(TxtPreUni.Text), "#0.00")
'    LblCostoEgr.Caption = Format(CDbl(TxtNumHec.Text) * CDbl(LblCostoTotal.Caption), "#0.00")
'    LblCostosUtil.Caption = Format(CDbl(LblCostosIng.Caption) - CDbl(LblCostoEgr.Caption), "#0.00")
'End Sub
'
'Private Sub TxtProd_GotFocus()
'    fEnfoque TxtProd
'End Sub
'
'Private Sub txtprod_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtProd, KeyAscii)
'    If KeyAscii = 13 Then
'        CboUnidad.SetFocus
'    End If
'End Sub
'
'Private Sub TxtProd_LostFocus()
'    If Trim(TxtProd.Text) = "" Then
'        TxtProd.Text = "0.00"
'    End If
'    TxtProd.Text = Format(TxtProd.Text, "#0.00")
'End Sub
'
'Private Sub txtProveedores_Change()
'    lblPasivo.Caption = Format(CDbl(IIf(Trim(txtProveedores.Text) = "", "0", txtProveedores.Text)) '            + CDbl(IIf(Trim(txtOtrosPrest.Text) = "", "0", txtOtrosPrest.Text)) '            + CDbl(IIf(Trim(txtPrestCmact.Text) = "", "0", txtPrestCmact.Text)), "#0.00")
'
'    lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
'    lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblPatrimonio.Caption), "#0.00")
'End Sub
'
'Private Sub txtProveedores_GotFocus()
'    fEnfoque txtProveedores
'End Sub
'
'Private Sub txtProveedores_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtProveedores, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtOtrosPrest.SetFocus
'    End If
'End Sub
'
'
'Private Sub txtProveedores_LostFocus()
'    If Len(Trim(txtProveedores.Text)) > 0 Then
'        txtProveedores.Text = Format(txtProveedores.Text, "#0.00")
'    Else
'        txtProveedores.Text = "0.00"
'    End If
'End Sub
'
'Private Sub TxtRazSocDescrip_GotFocus()
'    fEnfoque TxtRazSocDescrip
'End Sub
'
'Private Sub TxtRazSocDescrip_KeyPress(KeyAscii As Integer)
'    KeyAscii = Letras(KeyAscii)
'    If KeyAscii = 13 Then
'        TxtRazSocDirecc.SetFocus
'    End If
'End Sub
'
'Private Sub TxtRazSocDirecc_GotFocus()
'    fEnfoque TxtRazSocDirecc
'End Sub
'
'Private Sub TxtRazSocDirecc_KeyPress(KeyAscii As Integer)
'    KeyAscii = Letras(KeyAscii)
'    If KeyAscii = 13 Then
'        CmdUbigeo.SetFocus
'    End If
'End Sub
'
'
'Private Sub TxtRazSocTelef_GotFocus()
'    fEnfoque TxtRazSocTelef
'End Sub
'
'Private Sub TxtRazSocTelef_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosEnteros(KeyAscii)
'    If KeyAscii = 13 Then
'        If TxFecEval.Enabled And TxFecEval.Visible Then
'            TxFecEval.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub txtrecuperacion_Change()
'LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) '            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) '            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) '            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) '            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
'
'End Sub
'
'Private Sub txtrecuperacion_GotFocus()
'    fEnfoque txtrecuperacion
'End Sub
'
'Private Sub txtrecuperacion_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtrecuperacion, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        If txtcompras.Enabled = True Then
'            txtcompras.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub txtrecuperacion_LostFocus()
'    txtrecuperacion.Text = Format(IIf(Trim(txtrecuperacion.Text) = "", "0.00", txtrecuperacion.Text), "#0.00")
'
'End Sub
'
'Private Sub txtVentas_Change()
'LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) '            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) '            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) '            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) '            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")
'End Sub
'
'Private Sub txtVentas_GotFocus()
'    fEnfoque txtVentas
'End Sub
'
'Private Sub txtVentas_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(txtVentas, KeyAscii, 12)
'    If KeyAscii = 13 Then
'        txtrecuperacion.SetFocus
'    End If
'End Sub
'
'Private Sub txtVentas_LostFocus()
'    txtVentas.Text = Format(IIf(Trim(txtVentas.Text) = "", "0.00", txtVentas.Text), "#0.00")
'End Sub

Option Explicit

Dim oPersona As UPersona_Cli ' COMDPersona.DCOMPersona   'DPersona
Dim nIndice As Integer
Dim nProcesoEjecutado As Integer '1 Nueva fte de Ingreso; 2 Editar fte de Ingreso ; 3 Consulta de Fte
Dim vsUbiGeo As String
Dim bEstadoCargando As Boolean
Dim nProcesoActual As Integer
Dim ldFecEval As Date
'**PEAC**2008/04/11**************************************************************
Dim MatrixHojaEval() As String
Dim nPos As Integer
Dim nDat As Integer

'Public rsHojEval As ADODB.Recordset
'**End**PEAC**2008/04/11**************************************************************
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
Dim sNumeroFtesIngreso As String
Dim dFecFteIng As Date
Dim sImpr As String
Dim rsHojEval As ADODB.Recordset
'Variables de Excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim lsHoja         As String
Dim lbLibroOpen As Boolean
'Matriz temporal
Dim MatrixHojaETemp() As String
Dim nPost As Integer

'*** PEAC 20080412
Dim oPers As COMDCredito.DCOMCredito
Dim rsOHE As ADODB.Recordset
Dim nPos1 As Integer
Dim nPos2 As Integer
Dim nPos3 As Integer
Dim nPos4 As Integer
Dim nTipEv As Integer
Dim lcNumFuente As String
'*** FIN PEAC
'WIOR 20140319 *******************************
Private fbEditarEF As Boolean
Private fMatEstFinan As Variant
Private rsEstFinan As Recordset
'WIOR FIN ************************************

'Fin Variables excel
'sNumeroFtesIngreso = "00000000"
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
    
    'peac 20071227
    If txtFecEEFF.Enabled Then
        
        CadTemp = ValidaFecha(txtFecEEFF.Text)
        If Len(CadTemp) > 0 Then
            MsgBox "Por favor ingrese la fecha del los Estados Financieros en forma correcta.", vbInformation, "Aviso"
            ValidaDatosFuentesIngreso = False
            Exit Function
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
    
    '*** PEAC 20080402 - AQUI VALIDA EL INGRESO DE LA HOJA DE EVALUACION



    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoIndependiente Then
        
'        If CDbl(lblPatrimonio.Caption) <= 0 Then
'            MsgBox "Falta Ingresar el Balance de la Fuente de Ingreso"
'
'            'SSTFuentes.Tab = 1
'            '***PEAC 20080402
'            SSTFuentes.Tab = 3
'
'            txtDisponible.SetFocus
'            ValidaDatosFuentesIngreso = False
'            Exit Function
'        End If
    
    End If

End Function

'*** PEAC 20080515
Private Sub ImprimeHistoriaCrediticia(ByVal psPersCod As String, ByVal psPersCodCli As String, ByVal psNumFuente As String)

Dim oDCred As COMDCredito.DCOMCredDoc
Dim oDCredD As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim RsHis As ADODB.Recordset
Dim rsRef As ADODB.Recordset
Dim RsDatFte As ADODB.Recordset
Dim prsRefComer As ADODB.Recordset

Dim oConecta  As COMConecta.DCOMConecta

Dim sCadImp As String, i As Integer, J As Integer
Dim lsNombreArchivo As String

Dim lnActivoTotal As Double
Dim lnPasivoTotal As Double
Dim lnPatrimonio As Double
Dim lnTotPasPatri As Double
Dim lnIngresoNeto As Double

    Screen.MousePointer = 11

    Set oDCredD = New COMDCredito.DCOMCredito
        Set RsDatFte = oDCredD.RecuperaDatosFteIng(psPersCod, psNumFuente)
        Set RsHis = oDCredD.RecuperaHisCred(psPersCodCli)
        Set rsRef = oDCredD.RecuperaReferenciaPersComer(psPersCodCli)
    Set oDCredD = Nothing

    If RsHis.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")

    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add

    ApExcel.Cells(1, 2).Formula = "CMAC MAYNAS"
    ApExcel.Cells(2, 2).Formula = UCase(gsNomAge)
    ApExcel.Range("B1", "B2").Font.Bold = True

    ApExcel.Cells(4, 2).Formula = "CLIENTE : " & Trim(LblCliente.Caption)
    ApExcel.Range("B1", "B4").Font.Bold = True
    ApExcel.Range("B4", "Q4").Borders(xlEdgeBottom).Weight = xlMedium
'----------------------------------------------------------------------
    ApExcel.Cells(6, 2).Formula = "RASON SOCIAL/NOMBRE COMERCIAL"
    ApExcel.Cells(6, 5).Formula = "RUC"
    ApExcel.Cells(6, 6).Formula = "INICIO ACT."
    ApExcel.Cells(6, 7).Formula = "PERS.JURIDICA"
    ApExcel.Cells(6, 10).Formula = "TELEFONO"
    ApExcel.Range("B6", "J6").Font.Bold = True

    ApExcel.Cells(8, 2).Formula = "ACTIVIDAD(CIIU)"
    ApExcel.Cells(8, 5).Formula = "COND.LOCAL"
    ApExcel.Cells(8, 6).Formula = "MAGNITUD EMP."
    ApExcel.Cells(8, 10).Formula = "Nº EMPLEADOS"
    ApExcel.Range("B8", "J8").Font.Bold = True
    
    ApExcel.Cells(10, 2).Formula = "DIRECCION(Avenida,Calle,Jiron)"
    ApExcel.Cells(10, 5).Formula = "DPTO."
    ApExcel.Cells(10, 6).Formula = "PROV."
    ApExcel.Cells(10, 7).Formula = "DISTRITO"
    ApExcel.Cells(10, 10).Formula = "ZONA"
    ApExcel.Range("B10", "J10").Font.Bold = True
    
    ApExcel.Range("B11", "Q11").Borders(xlEdgeBottom).Weight = xlMedium

'-----------------------------------------------------------------------
    ApExcel.Cells(7, 2).Formula = RsDatFte!cPersNombre
    ApExcel.Cells(7, 5).Formula = RsDatFte!Ruc
    ApExcel.Cells(7, 6).Formula = Format(RsDatFte!ini_activ, "dd/mm/yyyy")
    ApExcel.Cells(7, 7).Formula = RsDatFte!PersJur
    ApExcel.Cells(7, 10).Formula = RsDatFte!telf

    ApExcel.Cells(9, 2).Formula = RsDatFte!CIIU
    ApExcel.Cells(9, 5).Formula = RsDatFte!DireccCondicion
    ApExcel.Cells(9, 6).Formula = RsDatFte!magnitud
    ApExcel.Cells(9, 10).Formula = RsDatFte!numEmple

    ApExcel.Cells(11, 2).Formula = RsDatFte!direcc
    ApExcel.Cells(11, 5).Formula = RsDatFte!DPTO
    ApExcel.Cells(11, 6).Formula = RsDatFte!PROV
    ApExcel.Cells(11, 7).Formula = RsDatFte!Dist
    ApExcel.Cells(11, 10).Formula = RsDatFte!Zona

'-----------------------------------------------------------------------

    ApExcel.Cells(13, 2).Formula = "HISTORIA CREDITICIA"
    ApExcel.Cells(13, 2).Font.Bold = True
    'ApExcel.Cells(2, 2).HorizontalAlignment = 3

    ApExcel.Range("B15", "H15").Font.Bold = True
    ApExcel.Range("B15", "H15").Borders(xlEdgeBottom).Weight = xlMedium
    ApExcel.Cells(15, 2).Formula = "Nº"
    ApExcel.Cells(15, 3).Formula = "CREDITO"
    ApExcel.Cells(15, 4).Formula = "ESTADO"
    ApExcel.Cells(15, 5).Formula = "FEC.VIGENCIA"
    ApExcel.Cells(15, 6).Formula = "MONTO SOLICITADO"
    ApExcel.Cells(15, 7).Formula = "CUOTAS"
    ApExcel.Cells(15, 8).Formula = "ANALISTA"

i = 15
J = 0
Do While Not RsHis.EOF
    i = i + 1
    J = J + 1
    ApExcel.Cells(i, 2).Formula = J
    ApExcel.Cells(i, 3).Formula = "'" & RsHis!cCtaCod
    ApExcel.Cells(i, 4).Formula = "'" & RsHis!Estado
    ApExcel.Cells(i, 5).Formula = Format(RsHis!dVigencia, "mm/dd/yyyy")
    ApExcel.Cells(i, 6).Formula = RsHis!montosoli
    ApExcel.Cells(i, 7).Formula = RsHis!nCuotas
    ApExcel.Cells(i, 8).Formula = RsHis!cUser
    RsHis.MoveNext
    If RsHis.EOF Then
        Exit Do
    End If
Loop
RsHis.MoveFirst

i = i + 2

    ApExcel.Cells(i, 2).Formula = "EVOLUCION ECONOMICA FINANCIERA"
    ApExcel.Cells(i, 2).Font.Bold = True
    'ApExcel.Cells(i, 2).HorizontalAlignment = 3
i = i + 3

    ApExcel.Range("B" & i - 1, "Q" & i).Font.Bold = True
    ApExcel.Range("B" & i, "Q" & i).Borders(xlEdgeBottom).Weight = xlMedium
    ApExcel.Cells(i, 2).Formula = "Nº"
    ApExcel.Cells(i, 3).Formula = "CREDITO"
    ApExcel.Cells(i, 4).Formula = "FEC.EVAL."
    ApExcel.Cells(i - 1, 5).Formula = "ACTIVO"
    ApExcel.Cells(i, 5).Formula = "CTE."  '**
    ApExcel.Cells(i, 6).Formula = "INVENTARIO"
    ApExcel.Cells(i - 1, 7).Formula = "ACTIVO"
    ApExcel.Cells(i, 7).Formula = "FIJO"  '**
    ApExcel.Cells(i - 1, 8).Formula = "ACTIVO"
    ApExcel.Cells(i, 8).Formula = "TOTAL"  '**
    ApExcel.Cells(i - 1, 9).Formula = "PASIVO"
    ApExcel.Cells(i, 9).Formula = "CTE."  '**
    ApExcel.Cells(i - 1, 10).Formula = "PASIVO"
    ApExcel.Cells(i, 10).Formula = "NO CTE."  '**
    ApExcel.Cells(i - 1, 11).Formula = "PASIVO"
    ApExcel.Cells(i, 11).Formula = "TOTAL"  '**
    ApExcel.Cells(i, 12).Formula = "PATRIMONIO"
    ApExcel.Cells(i - 1, 13).Formula = "TOTAL PASIVO"
    ApExcel.Cells(i, 13).Formula = "Y PATRIMONIO"  '**
    ApExcel.Cells(i, 14).Formula = "INGRESOS"
    ApExcel.Cells(i - 1, 15).Formula = "COSTO"
    ApExcel.Cells(i, 15).Formula = "VENTAS"  '**
    ApExcel.Cells(i - 1, 16).Formula = "OTROS"
    ApExcel.Cells(i, 16).Formula = "EGRESOS"  '**
    ApExcel.Cells(i - 1, 17).Formula = "INGRESO"
    ApExcel.Cells(i, 17).Formula = "NETO"  '**

J = 0
Do While Not RsHis.EOF
    i = i + 1
    J = J + 1
    lnActivoTotal = RsHis!Act_cte + RsHis!act_fijo
    lnPasivoTotal = RsHis!pas_no_cte + RsHis!pas_cte
    lnPatrimonio = lnActivoTotal - lnPasivoTotal
    lnTotPasPatri = lnPasivoTotal + lnPatrimonio
    lnIngresoNeto = RsHis!ingresos - RsHis!cos_vtas - RsHis!otr_egre

    ApExcel.Cells(i, 2).Formula = J
    ApExcel.Cells(i, 3).Formula = "'" & RsHis!cCtaCod
    ApExcel.Cells(i, 4).Formula = Format(RsHis!dPersEval, "mm/dd/yyyy")
    ApExcel.Cells(i, 5).Formula = RsHis!Act_cte
    ApExcel.Cells(i, 6).Formula = RsHis!INVENTARIO
    ApExcel.Cells(i, 7).Formula = RsHis!act_fijo
    ApExcel.Cells(i, 8).Formula = lnActivoTotal
    ApExcel.Cells(i, 9).Formula = RsHis!pas_cte
    ApExcel.Cells(i, 10).Formula = RsHis!pas_no_cte
    ApExcel.Cells(i, 11).Formula = lnPasivoTotal
    ApExcel.Cells(i, 12).Formula = lnPatrimonio
    ApExcel.Cells(i, 13).Formula = lnTotPasPatri
    ApExcel.Cells(i, 14).Formula = RsHis!ingresos
    ApExcel.Cells(i, 15).Formula = RsHis!cos_vtas
    ApExcel.Cells(i, 16).Formula = RsHis!otr_egre
    ApExcel.Cells(i, 17).Formula = lnIngresoNeto
    
    RsHis.MoveNext
    If RsHis.EOF Then
        Exit Do
    End If
Loop

i = i + 2
ApExcel.Cells(i, 2).Formula = "VARIACION"
ApExcel.Cells(i, 2).Font.Bold = True
If J > 1 Then
    ApExcel.Cells(i, 5).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 6).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 7).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 8).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 9).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 10).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 11).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 12).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 13).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 14).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 15).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 16).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
    ApExcel.Cells(i, 17).Formula = "=+IF(R[-3]C>0,R[-2]C/R[-3]C,0)"
Else
    ApExcel.Cells(i, 5).Formula = 0
    ApExcel.Cells(i, 6).Formula = 0
    ApExcel.Cells(i, 7).Formula = 0
    ApExcel.Cells(i, 8).Formula = 0
    ApExcel.Cells(i, 9).Formula = 0
    ApExcel.Cells(i, 10).Formula = 0
    ApExcel.Cells(i, 11).Formula = 0
    ApExcel.Cells(i, 12).Formula = 0
    ApExcel.Cells(i, 13).Formula = 0
    ApExcel.Cells(i, 14).Formula = 0
    ApExcel.Cells(i, 15).Formula = 0
    ApExcel.Cells(i, 16).Formula = 0
    ApExcel.Cells(i, 17).Formula = 0
End If

    RsHis.Close
    Set RsHis = Nothing

i = i + 2

    ApExcel.Cells(i, 2).Formula = "REFERENCIAS PERSONALES / COMERCIALES"
    ApExcel.Cells(i, 2).Font.Bold = True
'rsRef

i = i + 2

    ApExcel.Range("B" & i - 1, "Q" & i).Font.Bold = True
    ApExcel.Range("B" & i, "Q" & i).Borders(xlEdgeBottom).Weight = xlMedium
    ApExcel.Cells(i, 2).Formula = "Nº"
    ApExcel.Cells(i, 3).Formula = "NOMBRE/REF.COMERCIAL"
    ApExcel.Cells(i, 4).Formula = "TIPO REFERENCIA"
    ApExcel.Cells(i, 5).Formula = "TELEFONO"
    ApExcel.Cells(i, 6).Formula = "COMENTARIO"
    ApExcel.Cells(i, 7).Formula = "DIRECCION"

J = 0
Do While Not rsRef.EOF
    i = i + 1
    J = J + 1
       
    ApExcel.Cells(i, 2).Formula = J
    ApExcel.Cells(i, 3).Formula = rsRef!cNomRefCom
    ApExcel.Cells(i, 4).Formula = rsRef!cConsDescripcion
    ApExcel.Cells(i, 5).Formula = rsRef!cFonoRefCom
    ApExcel.Cells(i, 6).Formula = rsRef!cComentarioRefCom
    ApExcel.Cells(i, 7).Formula = rsRef!cDireccionRefCom
    
    rsRef.MoveNext
    If rsRef.EOF Then
        Exit Do
    End If
Loop

    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Sub

Private Sub HabilitaCabecera(ByVal pnHabilitar As Boolean)
    CboTipoFte.Enabled = pnHabilitar
    CboMoneda.Enabled = pnHabilitar
    TxtBRazonSoc.Enabled = pnHabilitar
    TxtRazSocDescrip.Enabled = pnHabilitar
    TxtRazSocDirecc.Enabled = pnHabilitar
    TxtRazSocTelef.Enabled = pnHabilitar
    
    'peac 20071227
    txtFecEEFF.Enabled = pnHabilitar
    
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
    'Label47.Enabled = HabBalance
    'Label49.Enabled = HabBalance
    Label31.Enabled = HabBalance
    Label32.Enabled = HabBalance
    'lblIngresosB.Enabled = HabBalance
    'lblEgresosB.Enabled = HabBalance
    TxtBalIngFam.Enabled = HabBalance
    TxtBalEgrFam.Enabled = HabBalance
    LblSaldoIngEgr.Enabled = HabBalance
    Frame4.Enabled = HabBalance
    Frame5.Enabled = HabBalance
    
    
cboGrupoHojEval.Enabled = HabBalance
cboTipoEval.Enabled = HabBalance
cboConceptoEval.Enabled = HabBalance
cmdGrabarEval.Enabled = HabBalance
'cmdCargaRS.Enabled = HabBalance
cmdBorraLinHojaEval.Enabled = HabBalance
txtCodEval.Enabled = HabBalance
txtMonto.Enabled = HabBalance
txtMonto2.Enabled = HabBalance
FEHojaEval.Enabled = HabBalance

    
    If Trim(Right(CboTipoFte.Text, 2)) = "2" Then
    
        '*** PEAC 20080402
        'Me.SSTFuentes.TabVisible(1) = True
        Me.SSTFuentes.TabVisible(0) = False
        Me.SSTFuentes.TabVisible(1) = False
        Me.SSTFuentes.TabVisible(2) = False
        Me.SSTFuentes.TabVisible(3) = True
        
        Me.SSTFuentes.TabVisible(0) = False
        'Me.SSTFuentes.Tab = 1
        Me.SSTFuentes.Tab = 3
    Else
        '*** PEAC 20080402
        'Me.SSTFuentes.TabVisible(1) = False
        
        Me.SSTFuentes.TabVisible(0) = False
        Me.SSTFuentes.TabVisible(1) = False
        Me.SSTFuentes.TabVisible(2) = False
        Me.SSTFuentes.TabVisible(3) = True
        Me.SSTFuentes.Tab = 3
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

'*** PEAC 20080402
Dim oGarant As COMDCredito.DCOMGarantia
Dim oEvalua As COMDCredito.DCOMCredito

Dim rsClaseInmueble As ADODB.Recordset
Dim rsCategoria As ADODB.Recordset

'Dim rsTipoEvaluacion As ADODB.Recordset
Dim rsGrupoHojEval As ADODB.Recordset
Dim rsTipoImporte As ADODB.Recordset
Dim i As Integer
On Error GoTo ERRORCargaControles
    
    Set oEvalua = New COMDCredito.DCOMCredito
    
    Call oEvalua.CargarObjetosControles(rsGrupoHojEval, nTipEv)
    
    Set oEvalua = Nothing

    Call CargarGrupoHojEval(rsGrupoHojEval, nTipEv)
    fbEditarEF = False 'WIOR 20140319

    Exit Sub

ERRORCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
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
        
        'peac 20071227
        txtFecEEFF.Text = poPersona.ObtenerFteIngFecEEFF(pnIndice, nUltFte)
        
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
Public Sub ObtenerArreglo(sNumeroFtesIngreso As String, dFecFteIng As Date, poPersona As UPersona_Cli, pnIndice As Integer, Optional nFechaVal As Integer = 0)
    Dim result As ADODB.Recordset
    Dim rsConDPer As COMDpersona.DCOMPersonas
    Dim oCrDoc As New COMDCredito.DCOMCredDoc
    Dim i As Integer
    For i = 0 To nPos
        FEHojaEval.EliminaFila (1)
    Next i
    nPos = 0
    sNumeroFtesIngreso = poPersona.ObtenerFteIngCod(pnIndice)
    If nFechaVal = 0 Then
    If oCrDoc.ObtenerFechaProNumFuente(sNumeroFtesIngreso, gdFecSis).RecordCount > 0 Then
    dFecFteIng = Format(oCrDoc.ObtenerFechaProNumFuente(sNumeroFtesIngreso, gdFecSis)!valor, "dd/mm/yyyy")
    End If
    End If
    Set rsConDPer = New COMDpersona.DCOMPersonas
    Set result = rsConDPer.ObtenerDatosHojEvaluaci(sNumeroFtesIngreso, dFecFteIng)
    If result.RecordCount > 0 Then
            If Not result.EOF And Not result.BOF Then
                result.MoveFirst
            End If
    Do Until result.EOF
        FEHojaEval.AdicionaFila
        nPos = FEHojaEval.row - 1
       ' MatrixHojaEval(1, nPos) = FEHojaEval.Row
        ReDim Preserve MatrixHojaEval(1 To 7, 0 To nPos + 1)
        MatrixHojaEval(1, nPos) = result!cGrupo
        MatrixHojaEval(2, nPos) = result!CTipoEva
        MatrixHojaEval(3, nPos) = result!cDescripcion
        MatrixHojaEval(4, nPos) = result!nPersonal
        MatrixHojaEval(5, nPos) = result!nNegocio
        MatrixHojaEval(6, nPos) = result!nUnico
        MatrixHojaEval(7, nPos) = result!cCodHojEval
'
        FEHojaEval.TextMatrix(FEHojaEval.row, 1) = MatrixHojaEval(1, nPos)
        FEHojaEval.TextMatrix(FEHojaEval.row, 2) = MatrixHojaEval(2, nPos)
        FEHojaEval.TextMatrix(FEHojaEval.row, 3) = MatrixHojaEval(3, nPos)
        FEHojaEval.TextMatrix(FEHojaEval.row, 4) = MatrixHojaEval(4, nPos)
        FEHojaEval.TextMatrix(FEHojaEval.row, 5) = MatrixHojaEval(5, nPos)
        FEHojaEval.TextMatrix(FEHojaEval.row, 6) = MatrixHojaEval(6, nPos)
        FEHojaEval.TextMatrix(FEHojaEval.row, 7) = MatrixHojaEval(7, nPos)
        'FEHojaEval.TextMatrix(FEHojaEval.Row, 1) = MatrixHojaEval(1, nPos)
        result.MoveNext
    Loop
        'Exit Sub'WIOR 20140319 comento
    End If
    
    Call MostarEstadosFinancieros(sNumeroFtesIngreso, dFecFteIng) 'WIOR 20140319
End Sub

Public Sub Editar(ByVal pnIndice As Integer, ByRef poPersona As UPersona_Cli, ByRef Res As ADODB.Recordset)
    Set oPersona = poPersona
    Set rsHojEval = Res
   
    '**ALPA***13/04/2008*****************************************************************
    'Set Res = poPersona.ObtenerRsHojaEva(pnIndice).RecordCount
    'dFecFteIng = poPersona.ObtenerFteIngFecInicioAct(pnIndice)
     '**End*******************************************************************************
    'FtesIngreso(nNumFtes - 1).cNumFte = pRs!cNumFuente
    
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
    
        '*** PEAC 20080402
        'SSTFuentes.TabVisible(1) = False
        SSTFuentes.TabVisible(3) = True
        
        SSTFuentes.TabVisible(0) = False
        SSTFuentes.Tab = 3
        txtMonto2.Enabled = False
        dFecFteIng = Format(poPersona.ObtenerFteIngFecEval(pnIndice, poPersona.ObtenerFteIngIngresoNumeroFteDep(nIndice), gPersFteIngresoTipoDependiente), "dd/mm/yyyy")
    Else
        '*** PEAC 20080402
        SSTFuentes.TabVisible(0) = False
        
        'SSTFuentes.TabVisible(1) = True
        SSTFuentes.TabVisible(3) = True
        
        'SSTFuentes.Tab = 1
        SSTFuentes.Tab = 3
        txtMonto2.Enabled = True
        dFecFteIng = Format(poPersona.ObtenerFteIngInicioFuente(pnIndice), "dd/mm/yyyy")
    End If
    Call ObtenerArreglo(sNumeroFtesIngreso, dFecFteIng, poPersona, pnIndice)
    Call HabilitaEstFinancieros(False) 'WIOR 20140319
    Me.Show 1
End Sub

Public Sub NuevaFteIngreso(ByRef poPersona As UPersona_Cli, Optional ByVal pnFteIndice As Integer = -1, Optional ByVal nTipoEvalucion As Integer = 0)
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
    nTipEv = nTipoEvalucion
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
    Me.Show 1
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
    'dFecFteIng = oPersona.ObtenerFteIngFecInicioAct(pnIndice)
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
    Call HabilitaEstFinancieros(False) 'WIOR 20140313
    CmbFecha.Visible = True
    TxFecEval.Visible = False
    Call CargaComboFechaEval
    bEstadoCargando = False
    CmdNuevo.Enabled = False
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        SSTFuentes.TabVisible(1) = False
        '*** PEAC 20080402
        SSTFuentes.TabVisible(3) = True
        
        SSTFuentes.TabVisible(0) = False
        SSTFuentes.Tab = 3
        txtMonto2.Enabled = False
    Else
        SSTFuentes.TabVisible(0) = False
        'SSTFuentes.TabVisible(1) = True
        
        '*** PEAC 20080402
        SSTFuentes.TabVisible(3) = True
        'SSTFuentes.Tab = 1
        SSTFuentes.Tab = 3
        txtMonto2.Enabled = True
        'dFecFteIng = Format(oPersona.ObtenerFteIngFecEval(pnIndice, poPersona.ObtenerFteIngIngresoNumeroFteDep(pnIndice), gPersFteIngresoTipoDependiente), "dd/mm/yyyy")
    End If
        'dFecFteIng = Format(poPersona.ObtenerFteIngInicioFuente(pnIndice), "dd/mm/yyyy")
        dFecFteIng = Format(oPersona.ObtenerFteIngFecEval(pnIndice, poPersona.ObtenerFteIngIngresoNumeroFteDep(pnIndice), gPersFteIngresoTipoDependiente), "dd/mm/yyyy")
        Call ObtenerArreglo(sNumeroFtesIngreso, dFecFteIng, oPersona, pnIndice)
    Me.Show 1
End Sub

'*** PEAC 20080402
Sub CargarGrupoHojEval(ByVal prsGruHojEval As ADODB.Recordset, Optional nTipEv As Integer)
    Dim sDes As String
    Dim nCodigo As Integer
    On Error GoTo ErrHandler
    If nTipEv > 0 Then
        Do Until prsGruHojEval.EOF
            nCodigo = prsGruHojEval!cCodHojEval
            sDes = prsGruHojEval!cDescripcion

            cboGrupoHojEval.AddItem Trim(sDes) & Space(100) & Trim(str(nCodigo))
            cboGrupoHojEval.ItemData(cboGrupoHojEval.NewIndex) = nCodigo
            
            prsGruHojEval.MoveNext
        Loop
    End If
    Exit Sub
ErrHandler:
    MsgBox "Error al cargar datos 1", vbInformation, "AVISO"
End Sub



'*** PEAC 20080402
Sub CargarTipoEvaluacion(ByVal prsTipEva As ADODB.Recordset)
    Dim sDes As String
    Dim nCodigo As Integer
    
    On Error GoTo ErrHandler

        Do Until prsTipEva.EOF
            nCodigo = prsTipEva!nConsCod
            sDes = prsTipEva!cConsDescripcion

            cboTipoEval.AddItem Trim(sDes) & Space(100) & Trim(str(nCodigo))
            cboTipoEval.ItemData(cboTipoEval.NewIndex) = nCodigo
            
            prsTipEva.MoveNext
        Loop
    Exit Sub
ErrHandler:
    MsgBox "Error al cargar datos", vbInformation, "AVISO"
End Sub

'**DAOR 20080606, Comentado por errores de programación de PEAC **************
''***PEAC 20080402
'Sub CargarTipoImporte(ByVal prsTipImp As ADODB.Recordset)
'    Dim sDes As String
'    Dim nCodigo As Integer
'
'    On Error GoTo ErrHandler
'
'        Do Until prsTipImp.EOF
'            nCodigo = prsTipImp!nConsValor
'            sDes = prsTipImp!cConsDescripcion
'
'            cboTipoImporte.AddItem Trim(sDes) & Space(100) & Trim(Str(nCodigo))
'            cboTipoImporte.ItemData(cboTipoImporte.NewIndex) = nCodigo
'
'            prsTipImp.MoveNext
'        Loop
'    Exit Sub
'ErrHandler:
'    MsgBox "Error al cargar datos", vbInformation, "AVISO"
'End Sub
'******************************************************************************

'*** PEAC - OBTIENE BALANCE,FLUJO Y CONSUMO
Public Function ListaSuperGarantias() As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim sSQL As String
    On Error GoTo ErrHandler

    sSQL = "Select nConsValor,cConsDescripcion from constante where nconscod='1048' and nConsValor<>'1048'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaSuperGarantias = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    Exit Function
ErrHandler:
    If Not oConec Is Nothing Then Set oConec = Nothing
    Set ListaSuperGarantias = Null
End Function

Private Sub cboConceptoEval_Click()
    If cboConceptoEval.ListIndex = -1 Then
        MsgBox "Debe Escoger un Concepto de evaluación", vbInformation, "Aviso"
        Exit Sub
    Else
        txtCodEval.Text = Trim(Right(cboConceptoEval.Text, 10))
    End If
    
End Sub

Private Sub cboConceptoEval_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'SendKeys "{Tab}", True
    If txtMonto.Visible = True Then
        txtMonto.SetFocus
    Else
        txtMonto2.SetFocus
    End If
End If
End Sub

Private Sub cboGrupoHojEval_Click()
If nTipEv < 1 Then
    MsgBox "Debe Escoger un Tipo de Evaluación", vbInformation, "Aviso"
Else
    If cboGrupoHojEval.ListIndex <> -1 Then

        Call ReLoadcboGrupoHojEval(cboGrupoHojEval.ItemData(cboGrupoHojEval.ListIndex), IIf(nTipEv, nTipEv, 0))
        If CInt(Right(cboGrupoHojEval.Text, 2)) = "10" Then ' Flujo de Caja
            If nTipEv = 1 Then  'Dependdiente
                txtMonto.Visible = True
                lblMonto1.Visible = True
                txtMonto2.Visible = False
                lblMonto2.Visible = False
            Else                'Independiente
                txtMonto.Visible = True
                lblMonto1.Visible = True
                txtMonto2.Visible = True
                lblMonto2.Visible = True
                lblMonto2.Caption = "Empresarial:"
            End If
            'txtMonto.Visible = True
            'lblMonto2.Visible = True
            'lblMonto1.Visible = True
        Else ' Diferente de Flujo de Caja
            If nTipEv = 1 Then
                txtMonto2.Visible = True
                lblMonto2.Visible = True
                txtMonto2.Enabled = True
            Else
                txtMonto2.Visible = True
                lblMonto2.Visible = True
                txtMonto2.Enabled = True
            End If
            txtMonto.Visible = False
            lblMonto1.Visible = False
            lblMonto2.Caption = "Importe:"
        End If
     End If
End If
cboConceptoEval.Clear
End Sub

Private Sub cboGrupoHojEval_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtBRazonSoc.Enabled Then
            TxtBRazonSoc.SetFocus
        End If
    End If
End Sub

Private Sub cboTipoEval_Click()
    
    If cboTipoEval.ListIndex = -1 Then
            MsgBox "Debe Escoger un Concepto de evaluación", vbInformation, "Aviso"
            Exit Sub
    End If
        
    Dim rs As ADODB.Recordset
    Dim objDEval As COMDCredito.DCOMCredito
    
    'On Error GoTo ErrHandler
        Set objDEval = New COMDCredito.DCOMCredito

        Set rs = objDEval.CargarRelEval(Right(cboTipoEval.Text, 4), nTipEv)

        Set objDEval = Nothing
        If Not rs.EOF And Not rs.BOF Then
            cboConceptoEval.Clear
        End If

        Do Until rs.EOF
            cboConceptoEval.AddItem Trim(rs!cDescripcion) & Space(200) & Trim(str(rs!cCodHojEval))
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

        'Call ReLoadcboTipoEval(cboTipoEval.ItemData(cboTipoEval.ListIndex))

        'Call ReLoadcboTipoMoneda(cboTipoEval.ItemData(cboTipoEval.ListIndex))

End Sub


'*** PEAC 20080402
Sub ReLoadcboGrupoHojEval(ByVal pnIdGrupoHojEval As String, Optional nTipEval As Integer = 2)
    Dim rs As ADODB.Recordset

    Dim objDEval As COMDCredito.DCOMCredito
    
    On Error GoTo ErrHandler
        Set objDEval = New COMDCredito.DCOMCredito
        
        Set rs = objDEval.CargarGruEval(pnIdGrupoHojEval, nTipEval)
        
        Set objDEval = Nothing
        If Not rs.EOF And Not rs.BOF Then
            cboTipoEval.Clear
        End If
                
        Do Until rs.EOF
            cboTipoEval.AddItem Trim(rs!cDescripcion) & Space(100) & Trim(str(rs!cCodHojEval))
            rs.MoveNext
        Loop
        Set rs = Nothing
    Exit Sub
ErrHandler:
    If Not objDEval Is Nothing Then Set objDEval = Null
    
    If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error al cargar"
End Sub


'*** PEAC 20080402
Sub ReLoadcboTipoEval(ByVal pnIdSuperTipoEval As Integer)
    Dim rs As ADODB.Recordset
    
    Dim objDEval As COMDCredito.DCOMCredito
    
    On Error GoTo ErrHandler
        Set objDEval = New COMDCredito.DCOMCredito
        
        Set rs = objDEval.CargarRelEval(pnIdSuperTipoEval)
        
        Set objDEval = Nothing
        If Not rs.EOF And Not rs.BOF Then
            cboConceptoEval.Clear
        End If
                
        Do Until rs.EOF
            cboConceptoEval.AddItem Trim(rs!cDescripcion) & Space(100) & Trim(str(rs!nCodigo))
            rs.MoveNext
        Loop
        Set rs = Nothing
    Exit Sub
ErrHandler:
    If Not objDEval Is Nothing Then Set objDEval = Null
    
    If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error al cargar"
End Sub

'**DAOR 20080606, Comentado por errores de programacion de PEAC*******************
'*** PEAC 20080402
'Sub ReLoadcboTipoMoneda(ByVal pnIdSuperMone As Integer)
'    Dim rs As ADODB.Recordset
'
'    Dim objDMone As COMDCredito.DCOMCredito
'
'    On Error GoTo ErrHandler
'        Set objDMone = New COMDCredito.DCOMCredito
'
'        Set rs = objDMone.CargarTipoMone(pnIdSuperMone)
'
'        Set objDMone = Nothing
'        If Not rs.EOF And Not rs.BOF Then
'            cboTipoImporte.Clear
'        End If
'
'        Do Until rs.EOF
'            cboTipoImporte.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim(Str(rs!nConsValor))
'            rs.MoveNext
'        Loop
'
'        cboTipoImporte.ListIndex = 0
'
'        Set rs = Nothing
'    Exit Sub
'ErrHandler:
'    If Not objDMone Is Nothing Then Set objDMone = Null
'
'    If Not rs Is Nothing Then Set rs = Nothing
'    MsgBox "Error al cargar"
'End Sub
'********************************************************************************

Private Sub cboTipoEval_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub CboTipoFte_Click()
    If Trim(Right(CboTipoFte.Text, 15)) = gPersFteIngresoTipoDependiente Then
        Call HabilitaBalance(True)
       ChkCostoProd.value = 0
       ChkCostoProd.Enabled = False
       txtFecEEFF.Enabled = False
       nTipEv = gPersFteIngresoTipoDependiente
       'txtMonto2.Enabled = False
       txtMonto2.Visible = False
       lblMonto2.Visible = False
'        TxtBRazonSoc.Enabled = True
    Else
        Call HabilitaBalance(True)
        ChkCostoProd.Enabled = True
        txtFecEEFF.Enabled = True
        nTipEv = 2
        'txtMonto2.Enabled = True
        txtMonto2.Visible = True
        lblMonto2.Visible = True
        txtMonto.Visible = True
        lblMonto1.Visible = True
'        TxtBRazonSoc.Text = ""
'        TxtBRazonSoc.Enabled = False
'        LblRazonSoc.Caption = ""
    End If
    cboGrupoHojEval.Clear
    cboTipoEval.Clear
    cboConceptoEval.Clear
    Call CargaControles
    Set oPers = New COMDCredito.DCOMCredito
    Set rsOHE = oPers.ObtieneCodEvaluacion(nTipEv)
End Sub

Private Sub CboTipoFte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboMoneda.SetFocus
    End If
End Sub

Private Sub cboTipoImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
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
Dim oPersonaD  As COMDpersona.DCOMPersona
Dim sFecha  As String

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
    cmdCargaRS.Enabled = True 'PTI120170530 según ERS014-2017
    'Verifica si ya esta asignado a un Credito
    HabilitaCabecera False
    HabilitaIngresosEgresos False
    
    HabilitaBalance False
    
    
    If Trim(Right(CboTipoFte.Text, 2)) = "2" Then
        'Me.SSTFuentes.TabVisible(1) = True
        
        '*** PEAC 20080402
        Me.SSTFuentes.TabVisible(3) = True
        
        Me.SSTFuentes.TabVisible(0) = False
        'Me.SSTFuentes.Tab = 1
        Me.SSTFuentes.Tab = 3
       ' dFecFteIng = oPersona.ObtenerFteIngFecInicioAct(nIndice)
    Else
    
        'Me.SSTFuentes.TabVisible(1) = False
        
        '*** PEAC 20080402
        Me.SSTFuentes.TabVisible(3) = True
        Me.SSTFuentes.TabVisible(0) = False
        Me.SSTFuentes.Tab = 3
    End If
   
    Set oPersonaD = New COMDpersona.DCOMPersona
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
    sFecha = Left(CmbFecha.Text, 10)
    Call ObtenerArreglo(sNumeroFtesIngreso, CDate(sFecha), oPersona, nIndice, 1)
    Exit Sub
    'CmbFecha.Text = Left(sFecha, 10)
End Sub

Private Sub CmbFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbFecha_Click 'PTI120170530 según ERS014-2017
        cmdCargaRS.Enabled = True 'PTI120170530 según ERS014-2017
        'Me.cboGrupoHojEval.SetFocus 'Comentado PTI120170530 según ERS014-2017
        
'        If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'            TxtIngCon.SetFocus
'        Else
'            txtDisponible.SetFocus
'        End If
    End If
End Sub

Private Sub CmdAceptar_Click()
'Dim oPersonaNeg As npersona
Dim lnConta As Integer, i As Integer
    
    '*** PEAC 20080618 - verifica que se haya ingresado la expos.con este cred.
    lnConta = 0
    
    For i = 0 To nPos
        If MatrixHojaEval(7, i) = "400205" Then
            lnConta = lnConta + 1
        End If
    Next i
    
    If lnConta = 0 Then
        MsgBox "Falta ingresar ''EXPOSICION CON ESTE CREDITO''", vbInformation, "Aviso"
        Exit Sub
    End If
    
    '****************************************************************
    
    
    
    sImpr = ""
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
        Set rsHojEval = CargaRS()
        Call oPersona.ActualizarResFteIng(rsHojEval, nIndice)
        Call oPersona.ActualizarFteRSCInd(1, nIndice)
        'Call oPersona.ActualizarFteIngConyugue(CDbl(TxtIngCon.Text), nIndice, 0)
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
        
        '***PEAC 20080402
        'Public rsHojEval As ADODB.Recordset
        Set rsHojEval = CargaRS()
        Call oPersona.ActualizarResFteIng(rsHojEval, nIndice)
        'peac 20071227
        Call oPersona.ActualizarFteIngFecEEFF(CDate(txtFecEEFF.Text), nIndice, 0)
        Call oPersona.ActualizarFteRSCInd(1, nIndice)
    End If
    ' se verifica que el tab de produccion  este visible
    
    'WIOR 20140321 ***************************************
    Set rsEstFinan = CargaRSEstFinan()
    Call oPersona.ActualizarRsEstFinan(rsEstFinan, nIndice)
    'WIOR FIN ********************************************
    
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
    'Res = rsHojEval
    Unload Me
End Sub

Private Sub cmdBorraLinHojaEval_Click()
    Dim nXPos As Integer
    nXPos = FEHojaEval.row
    If nPos >= 1 Then
    If MsgBox("Esta Seguro de Eliminar este registro.", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        FEHojaEval.EliminaFila (FEHojaEval.row)
        If nPos >= 1 Then
      '  If FEHojaEval.Row < nPos - 1 Then
            Dim J As Integer
            For J = nXPos - 1 To nPos
                MatrixHojaEval(1, J) = MatrixHojaEval(1, J + 1)
                MatrixHojaEval(2, J) = MatrixHojaEval(2, J + 1)
                MatrixHojaEval(3, J) = MatrixHojaEval(3, J + 1)
                MatrixHojaEval(4, J) = MatrixHojaEval(4, J + 1)
                MatrixHojaEval(5, J) = MatrixHojaEval(5, J + 1)
                MatrixHojaEval(6, J) = MatrixHojaEval(6, J + 1)
                MatrixHojaEval(7, J) = MatrixHojaEval(7, J + 1)
            Next J
       '     End If
            nPos = nPos - 1
        Else
            nPos = nPos - 1
            nDat = 0
        End If
    End If
    Else
        If FEHojaEval.row >= 1 Then
        FEHojaEval.EliminaFila (1)
        End If
        nPos = -1
        nDat = 0
    End If
'    If nPos < 0 Then
'    nPos = 0
'    End If
    If nProcesoEjecutado = 1 Then
        If nPos < 0 Then
            CboTipoFte.Enabled = True
        End If
    End If
End Sub
Function ObtenerExistencia(ByVal rs As ADODB.Recordset, codigo As String) As String
Dim sretorno As String
sretorno = ""
If rs.RecordCount > 0 Then

With rs
    'If Not .EOF And Not .BOF Then
    If .EOF Or .BOF Then
        .MoveFirst
    End If
    Do Until .EOF
    '*********Inicio*************************************************************************
        If IIf(!cCodHojEval, !cCodHojEval, "00000") = codigo Then
        sretorno = codigo
        Exit Do
        End If
    '*******End******************************************************************************
    .MoveNext
    Loop
End With
End If
ObtenerExistencia = sretorno
End Function
Function ObtenerExistenciaRs(codigo As String, tamano As Integer, rs As ADODB.Recordset) As Integer
 Dim i As Integer
 Dim retorno As Integer
 retorno = -1
 For i = 0 To nPos
        If Left(MatrixHojaEval(7, i), tamano) = codigo Then
        '    If ObtenerExistencia(rs, codigo) <> "" Then
                    retorno = i
                'Exit For
         '   Else
         '       retorno = -1
                Exit For
         '   End If
        End If
Next i
ObtenerExistenciaRs = retorno
End Function
Function CargaRS(Optional rsMaq As ADODB.Recordset) As ADODB.Recordset
Dim rsHojEval As ADODB.Recordset
Dim i As Integer
Dim J As Integer
Dim contador As Integer
J = 0
If Len(FEHojaEval.TextMatrix(FEHojaEval.row, 1)) = 0 Then
    MsgBox "No existe datos para Grabar", vbInformation, "Aviso"
    Exit Function
End If

Set rsHojEval = New ADODB.Recordset

With rsHojEval
    'Crear RecordSet
    .Fields.Append "dPersEval", adDate              '1
    .Fields.Append "cNumFuente", adVarChar, 8       '2
    .Fields.Append "cCodHojEval", adVarChar, 8      '3
    .Fields.Append "cDescripcion", adVarChar, 250    '4
    .Fields.Append "nPersonal", adCurrency          '5
    .Fields.Append "nNegocio", adCurrency           '6
    .Fields.Append "nUnico", adCurrency             '7
    .Fields.Append "cTipoEval", adVarChar, 50       'So
    .Fields.Append "cTituloEval", adVarChar, 50     'So
    .Fields.Append "nTipoAct", adInteger            'So
    .Fields.Append "nValorEval", adCurrency, 50     'So
    .Open
    'Llenar Recordset
     Dim nEncontrado As Integer
     contador = 0
        If sImpr = "SI" Then
        '***************Inicio******************************************************
        
        If rsMaq.RecordCount > 0 Then
            If Not rsMaq.EOF And Not rsMaq.BOF Then
                rsMaq.MoveFirst
            End If
        Do Until rsMaq.EOF
                            J = 0
                            nEncontrado = 0
                            nEncontrado = ObtenerExistenciaRs(Trim(rsMaq!cCodHojEval), Len(Trim(rsMaq!cCodHojEval)), rsHojEval)
                         
                            If ObtenerExistencia(rsHojEval, Left(Trim(rsMaq!cCodHojEval), 2)) = "" And Len(Trim(rsMaq!cCodHojEval)) = 2 And nEncontrado > -1 Then
                                .AddNew
                                .Fields("dPersEval") = dFecFteIng
                                .Fields("cNumFuente") = sNumeroFtesIngreso
                                .Fields("cCodHojEval") = Left(MatrixHojaEval(7, nEncontrado), 2)
                                .Fields("cDescripcion") = MatrixHojaEval(1, nEncontrado)
                                .Fields("nPersonal") = 0#
                                .Fields("nNegocio") = 0#
                                .Fields("nUnico") = 0#
                                .Fields("cTipoEval") = MatrixHojaEval(1, nEncontrado)
                                .Fields("cTituloEval") = MatrixHojaEval(2, nEncontrado)
                                '.Fields("nTipoAct") = oPersona.ObtenerFteIngTipoAct(nIndice)
                                 J = 1
                                 contador = contador + 1
                            ElseIf ObtenerExistencia(rsHojEval, Left(Trim(rsMaq!cCodHojEval), 4)) = "" And Len(Trim(rsMaq!cCodHojEval)) = 4 And nEncontrado > -1 Then
                                .AddNew
                                .Fields("dPersEval") = dFecFteIng
                                .Fields("cNumFuente") = sNumeroFtesIngreso
                                .Fields("cCodHojEval") = Left(MatrixHojaEval(7, nEncontrado), 4)
                                .Fields("cDescripcion") = MatrixHojaEval(2, nEncontrado)
                                .Fields("nPersonal") = 0#
                                .Fields("nNegocio") = 0#
                                .Fields("nUnico") = 0#
                                .Fields("cTipoEval") = MatrixHojaEval(1, nEncontrado)
                                .Fields("cTituloEval") = MatrixHojaEval(2, nEncontrado)
                               ' .Fields("nTipoAct") = oPersona.ObtenerFteIngTipoAct(nIndice)
                                contador = contador + 1
                                 J = 1
                            ElseIf ObtenerExistencia(rsHojEval, Left(Trim(rsMaq!cCodHojEval), 6)) = "" And Len(Trim(rsMaq!cCodHojEval)) = 6 And nEncontrado > -1 Then
                                .AddNew
                                .Fields("dPersEval") = dFecFteIng
                                .Fields("cNumFuente") = sNumeroFtesIngreso
                                .Fields("cCodHojEval") = MatrixHojaEval(7, nEncontrado)
                                .Fields("cDescripcion") = MatrixHojaEval(3, nEncontrado)
                                .Fields("nPersonal") = IIf(Len(MatrixHojaEval(4, nEncontrado)) = 0, "0.00", MatrixHojaEval(4, nEncontrado))
                                .Fields("nNegocio") = IIf(Len(MatrixHojaEval(5, nEncontrado)) = 0, "0.00", MatrixHojaEval(5, nEncontrado))
                                .Fields("nUnico") = IIf(Len(MatrixHojaEval(6, nEncontrado)) = 0, "0.00", MatrixHojaEval(6, nEncontrado))
                                .Fields("cTipoEval") = MatrixHojaEval(1, nEncontrado)
                                .Fields("cTituloEval") = MatrixHojaEval(2, nEncontrado)
                                '.Fields("nTipoAct") = oPersona.ObtenerFteIngTipoAct(nIndice)
                                contador = contador + 1
                                 J = 1
                             ElseIf ObtenerExistencia(rsHojEval, Left(Trim(rsMaq!cCodHojEval), 8)) = "" And Len(Trim(rsMaq!cCodHojEval)) = 8 And nEncontrado > -1 Then
                                .AddNew
                                .Fields("dPersEval") = dFecFteIng
                                .Fields("cNumFuente") = sNumeroFtesIngreso
                                .Fields("cCodHojEval") = MatrixHojaEval(7, nEncontrado)
                                .Fields("cDescripcion") = MatrixHojaEval(3, nEncontrado)
                                .Fields("nPersonal") = IIf(Len(MatrixHojaEval(4, nEncontrado)) = 0, "0.00", MatrixHojaEval(4, nEncontrado))
                                .Fields("nNegocio") = IIf(Len(MatrixHojaEval(5, nEncontrado)) = 0, "0.00", MatrixHojaEval(5, nEncontrado))
                                .Fields("nUnico") = IIf(Len(MatrixHojaEval(6, nEncontrado)) = 0, "0.00", MatrixHojaEval(6, nEncontrado))
                                .Fields("cTipoEval") = MatrixHojaEval(1, nEncontrado)
                                .Fields("cTituloEval") = MatrixHojaEval(2, nEncontrado)
                                '.Fields("nTipoAct") = oPersona.ObtenerFteIngTipoAct(nIndice)
                                contador = contador + 1
                                 J = 1
                             Else
                                .AddNew
                                .Fields("dPersEval") = rsMaq!dPersEval
                                .Fields("cNumFuente") = rsMaq!cNumFuente
                                .Fields("cCodHojEval") = rsMaq!cCodHojEval
                                .Fields("cDescripcion") = rsMaq!cDescripcion
                                .Fields("nPersonal") = 0#
                                .Fields("nNegocio") = 0#
                                .Fields("nUnico") = 0#
                                .Fields("cTipoEval") = ""
                                .Fields("cTituloEval") = ""
                                .Fields("nTipoAct") = 1
                                contador = contador + 1
                               
                           End If
                         rsMaq.MoveNext
    Loop
   End If
       
        Else
'        If oPersona.ObtenerFteIngCod(nIndice) <> "" Then
'        sNumeroFtesIngreso = oPersona.ObtenerFteIngCod(nIndice)
'        End If
        'If TxFecEval.Text <> "__/__/____" Then
        If nProcesoActual = 1 Then ' Nuevo
            dFecFteIng = TxFecEval.Text
        ElseIf nProcesoActual = 2 Then ' Nuevo
            dFecFteIng = CDate(CmbFecha.Text) 'Editar
        Else
            dFecFteIng = TxFecEval.Text
        End If
        'If CStr(dFecFteIng) = "" Then
        '    dFecFteIng = TxFecEval.Text
       ' Else
        '    dFecFteIng = dFecFteIng
       ' End If
        For i = 0 To nPos
        J = 0
            If IIf(Len(MatrixHojaEval(4, i)) = 0, "0.00", MatrixHojaEval(4, i)) > 0# Then
                .AddNew
                .Fields("cNumFuente") = sNumeroFtesIngreso
                .Fields("cCodHojEval") = MatrixHojaEval(7, i)
                .Fields("dPersEval") = Format(dFecFteIng, "YYYY/mm/dd")
                .Fields("nTipoAct") = 1
                .Fields("nValorEval") = IIf(Len(MatrixHojaEval(4, i)) = 0, "0.00", MatrixHojaEval(4, i))
            End If
            If IIf(Len(MatrixHojaEval(5, i)) = 0, "0.00", MatrixHojaEval(5, i)) > 0# Then
                .AddNew
                .Fields("cNumFuente") = sNumeroFtesIngreso
                .Fields("cCodHojEval") = MatrixHojaEval(7, i)
                .Fields("dPersEval") = Format(dFecFteIng, "YYYY/mm/dd")
                .Fields("nTipoAct") = 2
                .Fields("nValorEval") = IIf(Len(MatrixHojaEval(5, i)) = 0, "0.00", MatrixHojaEval(5, i))
            End If
            If IIf(Len(MatrixHojaEval(6, i)) = 0, "0.00", MatrixHojaEval(6, i)) > 0# Then
                .AddNew
                .Fields("cNumFuente") = sNumeroFtesIngreso
                .Fields("cCodHojEval") = MatrixHojaEval(7, i)
                .Fields("dPersEval") = Format(dFecFteIng, "YYYY/mm/dd")
                .Fields("nTipoAct") = 3
                .Fields("nValorEval") = IIf(Len(MatrixHojaEval(6, i)) = 0, "0.00", MatrixHojaEval(6, i))
            End If
            Next i
        End If
    If Not .EOF Then .MoveFirst
End With
Set CargaRS = rsHojEval
'Set CargaRS = rsHojEval+rsHojEval
End Function
Private Sub cmdCargaRS_Click()
''    Dim rsHojEval As ADODB.Recordset
''    Dim rsHojMaq As ADODB.Recordset
''    Dim rsCabHojEval As ADODB.Recordset
''    Dim oDCom As New ComdCredito.DCOMCredito
''    Dim oDPer As New COMDPersona.DCOMPersonas
''    Set rsHojMaq = oDCom.CargarMaquetaHojaEval(nTipEv)
''    Set rsHojEval = CargaRS(rsHojMaq)
''    '**ALPA**20080412
''
''    Dim rsHojEva As New ADODB.Recordset
''    Dim sNFuFinan As String
''    Dim sTxtFec As Date
''
''    Set rsCabHojEval = oDPer.ObtenerDatosDocsPers(oPersona.PersCodigo)
''
''    If rsCabHojEval.RecordCount = 1 Then
''        If Not rsCabHojEval.BOF = True And Not rsCabHojEval.EOF = True Then
''            rsCabHojEval.MoveFirst
''        End If
''    Do Until rsCabHojEval.EOF
''        Call ImprimeHojaEvaluacionExcelCab(rsHojEval, rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis)
''        previo.Show imprime_HojaEvalucacion(rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis), "Hoja de Evaluación", True
''        rsCabHojEval.MoveNext
''    Loop
''    End If
Dim rsHojEval As ADODB.Recordset
Dim rsHojMaq As ADODB.Recordset
Dim rsCabHojEval As ADODB.Recordset

Dim oNCredito As COMNCredito.NCOMCredito
Dim oDCredito As COMDCredito.DCOMCredito
Dim oDPer As New COMDpersona.DCOMPersonas
Dim nCapaPa As Double

Set oDCredito = New COMDCredito.DCOMCredito
Set rsHojMaq = oDCredito.CargarMaquetaHojaEval(nTipEv)
Set rsHojEval = CargaRS(rsHojMaq)
Set rsCabHojEval = oDPer.ObtenerDatosDocsPers(oPersona.PersCodigo)
nCapaPa = 0

If rsCabHojEval.RecordCount = 1 Then
    If Not rsCabHojEval.BOF = True And Not rsCabHojEval.EOF = True Then
        rsCabHojEval.MoveFirst
    End If
Do Until rsCabHojEval.EOF
    Set oNCredito = New COMNCredito.NCOMCredito
    previo.Show oNCredito.GeneraMatrixEvaluacion(rsHojEval, rsCabHojEval!cPersona, rsCabHojEval!cPersCod, rsCabHojEval!cPerRUC, rsCabHojEval!cPerDNI, "CAJA MUNICIPAL DE MAYNAS", "OFICINA PRINCIPAL", gdFecSis, gsCodUser, , , , , nTipEv, 1), "Hoja de Evaluación", True
rsCabHojEval.MoveNext
Loop
End If
End Sub
'**ALPA**2008/04/11************************************************************
'******************************************************************************
Function buscar_refceldas(x, Y) As String
Dim celda As String
celda = ""
If Y = 1 Then
   celda = "A" & x
ElseIf Y = 2 Then
   celda = "B" & x
ElseIf Y = 3 Then
   celda = "C" & x
ElseIf Y = 4 Then
   celda = "D" & x
ElseIf Y = 5 Then
   celda = "E" & x
ElseIf Y = 6 Then
   celda = "F" & x
ElseIf Y = 7 Then
   celda = "G" & x
ElseIf Y = 8 Then
   celda = "H" & x
ElseIf Y = 9 Then
   celda = "I" & x
End If
buscar_refceldas = celda
End Function
''Private Sub ImprimeHojaEvaluacionExcelCab(ByRef rs As adodb.Recordset, ByRef cNomCliente As String, ByRef cCodCliente As String, ByRef cRucC As String, ByRef cDNIC As String, ByRef sEmpresa As String, ByRef sOficina As String, ByRef dFecha As Date)
''    xlHoja1.Range("A1:S1").EntireColumn.Font.FontStyle = "Arial"
''    xlHoja1.PageSetup.Orientation = xlLandscape
''    xlHoja1.PageSetup.CenterHorizontally = True
''    xlHoja1.PageSetup.Zoom = 75
''    xlHoja1.PageSetup.TopMargin = 2
''    xlHoja1.Range("A9:Z1").EntireColumn.Font.Size = 7
''    'xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignLeft
''    'gsCodUser
''    xlHoja1.Range("A1:A1").RowHeight = 17
''    xlHoja1.Range("A1:A1").ColumnWidth = 8
''    xlHoja1.Range("B1:B1").ColumnWidth = 40
''    xlHoja1.Range("C1:C1").ColumnWidth = 14
''    xlHoja1.Range("D1:D1").ColumnWidth = 12
''    xlHoja1.Range("E1:E1").ColumnWidth = 10
''    xlHoja1.Range("F1:F1").ColumnWidth = 14
''    xlHoja1.Range("G1:G1").ColumnWidth = 38
''    xlHoja1.Range("H1:H1").ColumnWidth = 20
''    xlHoja1.Range("I1:J1").ColumnWidth = 18
''    xlHoja1.Range("K1:K1").ColumnWidth = 20
''    xlHoja1.Range("L1:L1").ColumnWidth = 20
''    xlHoja1.Range("M1:M1").ColumnWidth = 16
''    xlHoja1.Range("N1:N1").ColumnWidth = 20
''    xlHoja1.Range("O1:O1").ColumnWidth = 18
''    xlHoja1.Range("P1:P1").ColumnWidth = 25
''    xlHoja1.Range("Q1:Q1").ColumnWidth = 16
''    xlHoja1.Range("R1:R1").ColumnWidth = 20
''    xlHoja1.Range("S1:S1").ColumnWidth = 16
''    xlHoja1.Range("T1:T1").ColumnWidth = 16
''    xlHoja1.Range("U1:U1").ColumnWidth = 16
''    xlHoja1.Range("V1:V1").ColumnWidth = 16
''    xlHoja1.Range("W1:W1").ColumnWidth = 16
''    xlHoja1.Range("X1:X1").ColumnWidth = 16
''
''    xlHoja1.Range("B1:B1").Font.Size = 12
''    xlHoja1.Range("A2:B4").Font.Size = 10
''    xlHoja1.Range("A1:I4").Font.Bold = True
''    'xlHoja1.Range("B1:B4").EntireColumn.HorizontalAlignment = xlHAlignLeft
''    xlHoja1.Cells(1, 2) = sEmpresa
''    xlHoja1.Cells(2, 2) = sOficina
''    xlHoja1.Cells(3, 2) = "Cliente :" & cNomCliente
''    xlHoja1.Cells(4, 2) = "Codigo  :" & cCodCliente
''    xlHoja1.Cells(1, 7) = "RUC     :" & cRucC
''    xlHoja1.Cells(2, 7) = "DNI     :" & cDNIC
''    xlHoja1.Cells(3, 7) = "Fecha   :" & Format(dFecha, "YYYY/mm/dd")
''    xlHoja1.Cells(4, 7) = "Usuario :" & gsCodUser
''    'xlHoja1.Range("B5").EntireColumn.HorizontalAlignment = xlHAlignLeft
''    'xlHoja1.Cells(2, 2) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
''   ' xlHoja1.Cells(4, 1) = "INSTITUCION : " '& psEmpresa
''    'xlHoja1.Range("B1:I2").Merge True
''    'xlHoja1.Range("A1:N2").HorizontalAlignment = xlHAlignCenter
''    'xlHoja1.Range("A" & nPosFin + 4 & ":U" & nPosFin + 4).Borders(xlEdgeTop).LineStyle = xlContinuous
''    'xlHoja1.Range("B7:E8").Borders(xlEdgeTop).LineStyle = xlContinuous
''    Dim PosicionX As Integer
''    Dim PosicionY As Integer
''    Dim PosicionX2 As Integer
''    Dim contador As Integer
''    Dim contador2 As Integer
''    Dim contador3 As Integer
''    Dim contador4  As Integer
''    Dim fSumaIte As Double
''
''    Dim fSumaSubnP As Double
''    Dim fSumaSubnN As Double
''    Dim fSumaSubnU As Double
''    Dim fSumaSubnT As Double
''
''    Dim fSumaSubnP2 As Double
''    Dim fSumaSubnN2 As Double
''    Dim fSumaSubnU2 As Double
''    Dim fSumaSubnT2 As Double
''
''    Dim fSumaSubnPG1 As Double
''    Dim fSumaSubnNG1 As Double
''    Dim fSumaSubnUG1 As Double
''    Dim fSumaSubnTG1 As Double
''
''    Dim fSumaSubnPG2 As Double
''    Dim fSumaSubnNG2 As Double
''    Dim fSumaSubnUG2 As Double
''    Dim fSumaSubnTG2 As Double
''    Dim fCostVe As Double
''    Dim PosY As Double
''    Dim PosY2 As Double
''    Dim PosYs As Double
''    Dim PosYs2 As Double
''    Dim PosicionXs As Integer
''    Dim PosicionX2s As Integer
''    contador = 6
''    contador2 = 6
''    nPost = 0
''     Do While Not rs.EOF
''        If Left(Trim(rs!cCodHojEval), 2) = "10" Then
''            contador = contador + 1
''            If contador = 7 Then
''                If Trim(rs!cCodHojEval) = "10" Then
''                    xlHoja1.Range("B5:E5").Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range("B5:E5").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range("B5:E5").Font.Bold = True
''                    xlHoja1.Cells(5, 2) = rs!cDescripcion
''                    '***Begin************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                    '***End**************************************************
''                    xlHoja1.Range("C7:C7").Font.Bold = True
''                    xlHoja1.Range("C7:C7").Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range("C7:C7").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    '**Begin**********************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(3, nPost) = "PERSONAL"
''                    MatrixHojaETemp(4, nPost) = "NEGOCIO"
''                    MatrixHojaETemp(5, nPost) = "TOTAL"
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                    '**End**********************
''                    xlHoja1.Cells(7, 3) = "PERSONAL"
''                    xlHoja1.Range("D7:D7").Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range("D7:D7").Font.Bold = True
''                    xlHoja1.Range("D7:D7").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(7, 4) = "NEGOCIO"
''                    xlHoja1.Range("E7:E7").Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range("E7:E7").Font.Bold = True
''                    xlHoja1.Range("E7:E7").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(7, 5) = "TOTAL"
''                End If
''            Else
''                fSumaIte = rs!nPersonal + rs!nNegocio
''                fSumaSubnP = fSumaSubnP + rs!nPersonal
''                fSumaSubnN = fSumaSubnN + rs!nNegocio
''                fSumaSubnT = fSumaSubnT + fSumaIte
''                fSumaSubnPG1 = fSumaSubnPG1 + rs!nPersonal
''                fSumaSubnNG1 = fSumaSubnNG1 + rs!nNegocio
''                fSumaSubnTG1 = fSumaSubnTG1 + fSumaIte
''
''                If Trim(rs!cCodHojEval) = "1004" Then
''                    '****Begin*********************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(3, nPost) = "PERSONAL"
''                    MatrixHojaETemp(4, nPost) = "NEGOCIO"
''                    MatrixHojaETemp(5, nPost) = "TOTAL"
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                    '*****End**********************************************************
''                    fSumaSubnPG2 = fSumaSubnPG1
''                    fSumaSubnNG2 = fSumaSubnNG1
''                    fSumaSubnTG2 = fSumaSubnTG1
''                    xlHoja1.Range(buscar_refceldas(contador, 2) & ":" & buscar_refceldas(contador, 2)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador, 2) & ":" & buscar_refceldas(contador, 2)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(contador, 2) & ":" & buscar_refceldas(contador, 2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(contador, 2) = "ACTIVO TOTAL"
''                    xlHoja1.Range(buscar_refceldas(contador, 3) & ":" & buscar_refceldas(contador, 3)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador, 3) & ":" & buscar_refceldas(contador, 3)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(contador, 3) & ":" & buscar_refceldas(contador, 3)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(contador, 3) = fSumaSubnPG1
''                    xlHoja1.Range(buscar_refceldas(contador, 4) & ":" & buscar_refceldas(contador, 4)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador, 4) & ":" & buscar_refceldas(contador, 4)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(contador, 4) & ":" & buscar_refceldas(contador, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(contador, 4) = fSumaSubnNG1
''                    xlHoja1.Range(buscar_refceldas(contador, 5) & ":" & buscar_refceldas(contador, 5)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador, 5) & ":" & buscar_refceldas(contador, 5)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(contador, 5) & ":" & buscar_refceldas(contador, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(contador, 5) = fSumaSubnTG1
''                    fSumaSubnPG1 = 0
''                    fSumaSubnNG1 = 0
''                    fSumaSubnTG1 = 0
''                    contador = contador + 1
''                End If
''                '****Begin**************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    MatrixHojaETemp(3, nPost) = CStr(rs!nPersonal)
''                    MatrixHojaETemp(4, nPost) = CStr(rs!nNegocio)
''                    MatrixHojaETemp(5, nPost) = CStr(fSumaIte)
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                '*****End****************************************************************
''                xlHoja1.Range(buscar_refceldas(contador, 2) & ":" & buscar_refceldas(contador, 2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                xlHoja1.Range(buscar_refceldas(contador, 3) & ":" & buscar_refceldas(contador, 3)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                xlHoja1.Range(buscar_refceldas(contador, 4) & ":" & buscar_refceldas(contador, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                xlHoja1.Range(buscar_refceldas(contador, 5) & ":" & buscar_refceldas(contador, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                xlHoja1.Cells(contador, 2) = rs!cDescripcion
''                xlHoja1.Cells(contador, 3) = rs!nPersonal
''                xlHoja1.Cells(contador, 4) = rs!nNegocio
''                xlHoja1.Cells(contador, 5) = fSumaIte
''                xlHoja1.Range(buscar_refceldas(contador, 3) & ":" & buscar_refceldas(contador, 5)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                xlHoja1.Range(buscar_refceldas(contador, 3) & ":" & buscar_refceldas(contador, 5)).NumberFormat = "#,##0.00"
''
''                If Len(rs!cCodHojEval) = 4 Then
''                If PosY <> 0 Then
''                    xlHoja1.Range(buscar_refceldas(PosY, 2) & ":" & buscar_refceldas(PosY, 5)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(PosY, 2) & ":" & buscar_refceldas(PosY, 5)).Font.Bold = True
''                    If rs!cCodHojEval <> "1002" Then
''                    '****Begin**************************************************************
''                    'nPost = nPost + 1
''                    'ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    'MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    MatrixHojaETemp(3, PosYs) = CStr(fSumaSubnP)
''                    MatrixHojaETemp(4, PosYs) = CStr(fSumaSubnN)
''                    MatrixHojaETemp(5, PosYs) = CStr(fSumaSubnT)
''                    MatrixHojaETemp(6, PosYs) = 1
''                    MatrixHojaETemp(7, PosYs) = 1
''                    '*****End****************************************************************
''                        xlHoja1.Cells(PosY, 3) = fSumaSubnP
''                        xlHoja1.Cells(PosY, 4) = fSumaSubnN
''                        xlHoja1.Cells(PosY, 5) = fSumaSubnT
''                        xlHoja1.Range(buscar_refceldas(PosY, 3) & ":" & buscar_refceldas(PosY, 5)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                        xlHoja1.Range(buscar_refceldas(PosY, 3) & ":" & buscar_refceldas(PosY, 5)).NumberFormat = "#,##0.00"
''
''                        If rs!cCodHojEval = "1004" Then
''                            '****Begin**************************************************************
''                            'nPost = nPost + 1
''                            'ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                            'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                            'MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                            MatrixHojaETemp(3, PosYs2) = CStr(fSumaSubnP2)
''                            MatrixHojaETemp(4, PosYs2) = CStr(fSumaSubnN2)
''                            MatrixHojaETemp(5, PosYs2) = CStr(fSumaSubnT2)
''                            MatrixHojaETemp(6, PosYs2) = 1
''                            MatrixHojaETemp(7, PosYs2) = 1
''                            '*****End****************************************************************
''                            xlHoja1.Cells(PosY2, 3) = fSumaSubnP2
''                            xlHoja1.Cells(PosY2, 4) = fSumaSubnN2
''                            xlHoja1.Cells(PosY2, 5) = fSumaSubnT2
''                            xlHoja1.Range(buscar_refceldas(PosY2, 3) & ":" & buscar_refceldas(PosY2, 5)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                            xlHoja1.Range(buscar_refceldas(PosY2, 3) & ":" & buscar_refceldas(PosY2, 5)).NumberFormat = "#,##0.00"
''                            fSumaSubnP2 = 0
''                            fSumaSubnN2 = 0
''                            fSumaSubnT2 = 0
''                        End If
''
''                    End If
''                End If
''                    If rs!cCodHojEval = "1002" Then
''                        PosY2 = PosY
'''                        xlHoja1.Cells(PosY, 3) = fSumaSubnP2
'''                        xlHoja1.Cells(PosY, 4) = fSumaSubnN2
'''                        xlHoja1.Cells(PosY, 5) = fSumaSubnT2
''                    End If
''                    PosY = 0
''                'If rs!cCodHojEval <> "1002" Then
''                    fSumaSubnP2 = fSumaSubnP2 + fSumaSubnP
''                    fSumaSubnN2 = fSumaSubnN2 + fSumaSubnN
''                    fSumaSubnT2 = fSumaSubnT2 + fSumaSubnT
''                    fSumaSubnP = 0
''                    fSumaSubnN = 0
''                    fSumaSubnT = 0
''                'End If
''                PosY = contador
''                PosYs = nPost
''                End If
''            End If
''        End If
''         If Left(Trim(rs!cCodHojEval), 2) = "20" Then
''          contador2 = contador2 + 1
''         If contador2 = 7 Then
''                If Trim(rs!cCodHojEval) = "20" Then
''                    PosY = 0
''                    PosYs = 0
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 2) & ":" & buscar_refceldas(contador + 3, 5)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 2) & ":" & buscar_refceldas(contador + 3, 5)).Font.Bold = True
''
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 3) & ":" & buscar_refceldas(contador + 3, 5)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 2) & ":" & buscar_refceldas(contador + 3, 5)).NumberFormat = "#,##0.00"
''
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 2) & ":" & buscar_refceldas(contador + 1, 2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 2, 2) & ":" & buscar_refceldas(contador + 2, 2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 3, 2) & ":" & buscar_refceldas(contador + 3, 2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 3) & ":" & buscar_refceldas(contador + 1, 3)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 4) & ":" & buscar_refceldas(contador + 1, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 1, 5) & ":" & buscar_refceldas(contador + 1, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''
''                    xlHoja1.Range(buscar_refceldas(contador + 2, 3) & ":" & buscar_refceldas(contador + 2, 3)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 2, 4) & ":" & buscar_refceldas(contador + 2, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 2, 5) & ":" & buscar_refceldas(contador + 2, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''
''                    xlHoja1.Range(buscar_refceldas(contador + 3, 3) & ":" & buscar_refceldas(contador + 3, 3)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 3, 4) & ":" & buscar_refceldas(contador + 3, 4)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador + 3, 5) & ":" & buscar_refceldas(contador + 3, 5)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    '****Begin**************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = "PASIVO TOTAL"
''                    MatrixHojaETemp(3, nPost) = CStr(fSumaSubnPG1)
''                    MatrixHojaETemp(4, nPost) = CStr(fSumaSubnPG2 - fSumaSubnPG1)
''                    MatrixHojaETemp(5, nPost) = CStr((fSumaSubnPG2 - fSumaSubnPG1) + fSumaSubnPG1)
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                    '*****End****************************************************************
''                    '****Begin**************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = "PATRIMONIO"
''                    MatrixHojaETemp(3, nPost) = CStr(fSumaSubnNG1)
''                    MatrixHojaETemp(4, nPost) = CStr(fSumaSubnNG2 - fSumaSubnNG1)
''                    MatrixHojaETemp(5, nPost) = CStr((fSumaSubnNG2 - fSumaSubnNG1) + fSumaSubnNG1)
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                    '*****End****************************************************************
''                    '****Begin**************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = "TOTAL PASIVO Y PATRIMONIO"
''                    MatrixHojaETemp(3, nPost) = CStr(fSumaSubnTG1)
''                    MatrixHojaETemp(4, nPost) = CStr(fSumaSubnTG2 - fSumaSubnTG1)
''                    MatrixHojaETemp(5, nPost) = CStr((fSumaSubnTG2 - fSumaSubnTG1) + fSumaSubnTG1)
''                    MatrixHojaETemp(6, nPost) = 1
''                    MatrixHojaETemp(7, nPost) = 1
''                    '*****End****************************************************************
''                    xlHoja1.Cells(contador + 1, 2) = "PASIVO TOTAL"
''                    xlHoja1.Cells(contador + 2, 2) = "PATRIMONIO"
''                    xlHoja1.Cells(contador + 3, 2) = "TOTAL PASIVO Y PATRIMONIO"
''
''                    xlHoja1.Cells(contador + 1, 3) = fSumaSubnPG1
''                    xlHoja1.Cells(contador + 1, 4) = fSumaSubnNG1
''                    xlHoja1.Cells(contador + 1, 5) = fSumaSubnTG1
''
''                    xlHoja1.Cells(contador + 2, 3) = fSumaSubnPG2 - fSumaSubnPG1
''                    xlHoja1.Cells(contador + 2, 4) = fSumaSubnNG2 - fSumaSubnNG1
''                    xlHoja1.Cells(contador + 2, 5) = fSumaSubnTG2 - fSumaSubnTG1
''
''                    xlHoja1.Cells(contador + 3, 3) = (fSumaSubnPG2 - fSumaSubnPG1) + fSumaSubnPG1
''                    xlHoja1.Cells(contador + 3, 4) = (fSumaSubnNG2 - fSumaSubnNG1) + fSumaSubnNG1
''                    xlHoja1.Cells(contador + 3, 5) = (fSumaSubnTG2 - fSumaSubnTG1) + fSumaSubnTG1
''                    '****Begin**************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    MatrixHojaETemp(6, nPost) = 2
''                    MatrixHojaETemp(7, nPost) = 2
''                    '*****End****************************************************************
''                    xlHoja1.Range(buscar_refceldas(5, 7) & ":" & buscar_refceldas(5, 8)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(5, 7) & ":" & buscar_refceldas(5, 8)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(5, 7) & ":" & buscar_refceldas(5, 8)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(5, 7) = rs!cDescripcion
''                  '  contador2 = contador2 + 1
''                End If
''            Else
''            If rs!cCodHojEval = "2002" Then
''                xlHoja1.Range(buscar_refceldas(contador2, 7) & ":" & buscar_refceldas(contador2, 8)).Cells.Interior.Color = RGB(220, 220, 220)
''                xlHoja1.Range(buscar_refceldas(contador2, 7) & ":" & buscar_refceldas(contador2, 8)).Font.Bold = True
''                xlHoja1.Range(buscar_refceldas(contador2, 7) & ":" & buscar_refceldas(contador2, 7)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''
''                xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).NumberFormat = "#,##0.00"
''                '****Begin**************************************************************
''                nPost = nPost + 1
''                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                MatrixHojaETemp(2, nPost) = "MARGEN BRUTO"
''                MatrixHojaETemp(3, nPost) = CStr(fCostVe + fSumaSubnU)
''                MatrixHojaETemp(6, nPost) = 2
''                MatrixHojaETemp(7, nPost) = 2
''                '*****End****************************************************************
''                xlHoja1.Cells(contador2, 7) = "MARGEN BRUTO"
''                xlHoja1.Cells(contador2, 8) = fCostVe + fSumaSubnU
''                contador2 = contador2 + 1
''            End If
''                xlHoja1.Range(buscar_refceldas(contador2, 7) & ":" & buscar_refceldas(contador2, 7)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''
''                xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).NumberFormat = "#,##0.00"
''                '****Begin***************************************************************
''                nPost = nPost + 1
''                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                MatrixHojaETemp(3, nPost) = CStr(rs!nUnico)
''                MatrixHojaETemp(6, nPost) = 2
''                MatrixHojaETemp(7, nPost) = 2
''                '*****End****************************************************************
''                xlHoja1.Cells(contador2, 7) = rs!cDescripcion
''                xlHoja1.Cells(contador2, 8) = rs!nUnico
''
''             If rs!cCodHojEval <> "200104" Then
''                fSumaSubnU = fSumaSubnU + rs!nUnico
''             End If
''             fSumaSubnUG1 = fSumaSubnUG1 + rs!nUnico
''             If Len(rs!cCodHojEval) = 4 Then
''                If PosY <> 0 Then
''                    xlHoja1.Range(buscar_refceldas(PosY, 7) & ":" & buscar_refceldas(PosY, 8)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(PosY, 7) & ":" & buscar_refceldas(PosY, 8)).Font.Bold = True
''
''                    xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                    xlHoja1.Range(buscar_refceldas(contador2, 8) & ":" & buscar_refceldas(contador2, 8)).NumberFormat = "#,##0.00"
''                    '****Begin***************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    'MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    MatrixHojaETemp(3, nPost) = CStr(fSumaSubnU)
''                    MatrixHojaETemp(6, nPost) = 2
''                    MatrixHojaETemp(7, nPost) = 2
''                    '*****End****************************************************************
''                    xlHoja1.Cells(PosY, 8) = fSumaSubnU
''                End If                '200104
''
''                 PosY = contador2
''                 PosYs = nPost
''             End If
''              If rs!cCodHojEval = "200104" Then
''                    fCostVe = rs!nUnico
''                End If
''            End If
''         End If
''         If Left(Trim(rs!cCodHojEval), 2) = "30" Then
''                contador4 = contador4 + 1
''                If contador4 = 1 Then
''                    PosicionX = 7
''                   contador3 = contador + 6
''                Else
''                   contador3 = contador3 + 1
''                End If
''
''               If contador4 = 1 Then
''                   If Trim(rs!cCodHojEval) = "30" Then
''                    fSumaSubnU = 0
''                    fSumaSubnUG1 = 0
''                    PosY = 0
''                    PosYs = 0
''                     '****Begin***************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = "INGRESO NETO"
''                    MatrixHojaETemp(3, nPost) = CStr(fSumaSubnU + fCostVe)
''                    MatrixHojaETemp(6, nPost) = 2
''                    MatrixHojaETemp(7, nPost) = 2
''                    '*****End****************************************************************
''                    xlHoja1.Range(buscar_refceldas(contador2 + 1, 7) & ":" & buscar_refceldas(contador2 + 1, 8)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador2 + 1, 7) & ":" & buscar_refceldas(contador2 + 1, 8)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(contador2 + 1, 7) & ":" & buscar_refceldas(contador2 + 1, 8)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(contador2 + 1, 7) = "INGRESO NETO"
''                    xlHoja1.Cells(contador2 + 1, 8) = fSumaSubnU + fCostVe
''                    xlHoja1.Range(buscar_refceldas(contador2 + 1, 8) & ":" & buscar_refceldas(contador2 + 1, 8)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                    xlHoja1.Range(buscar_refceldas(contador2 + 1, 8) & ":" & buscar_refceldas(contador2 + 1, 8)).NumberFormat = "#,##0.00"
''                    '****Begin***************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    'MatrixHojaETemp(3, nPost) = CStr(fSumaSubnU + fCostVe)
''                    MatrixHojaETemp(6, nPost) = 2
''                    MatrixHojaETemp(7, nPost) = 3
''                    '*****End****************************************************************
''                    xlHoja1.Range(buscar_refceldas(contador3 - 1, 2) & ":" & buscar_refceldas(contador3 - 1, 8)).Cells.Interior.Color = RGB(220, 220, 220)
''                    xlHoja1.Range(buscar_refceldas(contador3 - 1, 2) & ":" & buscar_refceldas(contador3 - 1, 8)).Font.Bold = True
''                    xlHoja1.Range(buscar_refceldas(contador3 - 1, 2) & ":" & buscar_refceldas(contador3 - 1, 8)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Cells(contador3 - 1, 2) = rs!cDescripcion
''                   End If
''                Else
''                If Trim(rs!cCodHojEval) = "3002" Then
''                 PosicionX2 = PosicionX
''                '****Begin***************************************************************
''                nPost = nPost + 1
''                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                'MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                MatrixHojaETemp(2, nPost) = "INGRESO TOTALES"
''                MatrixHojaETemp(3, nPost) = fSumaSubnU
''                MatrixHojaETemp(6, nPost) = 2
''                MatrixHojaETemp(7, nPost) = 3
''                nPost = nPost + 1
''                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                PosicionXs = nPost
''                '*****End****************************************************************
''                 xlHoja1.Range(buscar_refceldas(contador3, PosicionX) & ":" & buscar_refceldas(contador3, PosicionX + 1)).Cells.Interior.Color = RGB(220, 220, 220)
''                 xlHoja1.Range(buscar_refceldas(contador3, PosicionX) & ":" & buscar_refceldas(contador3, PosicionX + 1)).Font.Bold = True
''                 xlHoja1.Range(buscar_refceldas(contador3, PosicionX) & ":" & buscar_refceldas(contador3, PosicionX)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                 xlHoja1.Range(buscar_refceldas(contador3, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                 xlHoja1.Cells(contador3, PosicionX) = "INGRESO TOTALES"
''
''                 xlHoja1.Range(buscar_refceldas(contador3, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                 xlHoja1.Range(buscar_refceldas(contador3, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).NumberFormat = "#,##0.00"
''
''                 xlHoja1.Cells(contador3, PosicionX + 1) = fSumaSubnU
''                 contador3 = contador + 7
''                 PosicionX = 2
''                 fSumaSubnU = 0
''                End If
''                    xlHoja1.Range(buscar_refceldas(contador3, PosicionX) & ":" & buscar_refceldas(contador3, PosicionX)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador3, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''                    xlHoja1.Range(buscar_refceldas(contador3, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).EntireColumn.HorizontalAlignment = xlHAlignRight
''                    xlHoja1.Range(buscar_refceldas(contador3, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).NumberFormat = "#,##0.00"
''                    '****Begin***************************************************************
''                    nPost = nPost + 1
''                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''                    MatrixHojaETemp(1, nPost) = rs!cCodHojEval
''                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
''                    MatrixHojaETemp(3, nPost) = CStr(rs!nUnico)
''                    MatrixHojaETemp(6, nPost) = 2
''                    MatrixHojaETemp(7, nPost) = 4
''                    '*****End****************************************************************
''                    xlHoja1.Cells(contador3, PosicionX) = rs!cDescripcion
''                    xlHoja1.Cells(contador3, PosicionX + 1) = rs!nUnico
''                    fSumaSubnU = fSumaSubnU + rs!nUnico
''                    fSumaSubnUG1 = fSumaSubnUG1 + rs!nUnico
''                End If
''         End If
''        rs.MoveNext
''    Loop
''    xlHoja1.Range(buscar_refceldas(contador3 + 1, PosicionX) & ":" & buscar_refceldas(contador3 + 1, PosicionX)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''    xlHoja1.Range(buscar_refceldas(contador3 + 1, PosicionX + 1) & ":" & buscar_refceldas(contador3 + 1, PosicionX + 1)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''    xlHoja1.Range(buscar_refceldas(contador3 + 1, PosicionX) & ":" & buscar_refceldas(contador3 + 1, PosicionX + 1)).Cells.Interior.Color = RGB(220, 220, 220)
''    xlHoja1.Range(buscar_refceldas(contador3 + 1, PosicionX) & ":" & buscar_refceldas(contador3 + 1, PosicionX + 1)).Font.Bold = True
''
''    xlHoja1.Range(buscar_refceldas(contador3, PosicionX2) & ":" & buscar_refceldas(contador3, PosicionX2)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''    xlHoja1.Range(buscar_refceldas(contador3, PosicionX2 + 1) & ":" & buscar_refceldas(contador3, PosicionX2 + 1)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''
''    xlHoja1.Range(buscar_refceldas(contador3, PosicionX2) & ":" & buscar_refceldas(contador3, PosicionX2 + 1)).Cells.Interior.Color = RGB(220, 220, 220)
''    xlHoja1.Range(buscar_refceldas(contador3, PosicionX2) & ":" & buscar_refceldas(contador3, PosicionX2 + 1)).Font.Bold = True
''
''    xlHoja1.Range(buscar_refceldas(contador3 + 1, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).EntireColumn.HorizontalAlignment = xlHAlignRight
''    xlHoja1.Range(buscar_refceldas(contador3 + 1, PosicionX + 1) & ":" & buscar_refceldas(contador3, PosicionX + 1)).NumberFormat = "#,##0.00"
''
''    xlHoja1.Range(buscar_refceldas(contador3, PosicionX2 + 1) & ":" & buscar_refceldas(contador3, PosicionX2 + 1)).EntireColumn.HorizontalAlignment = xlHAlignRight
''    xlHoja1.Range(buscar_refceldas(contador3, PosicionX2 + 1) & ":" & buscar_refceldas(contador3, PosicionX2 + 1)).NumberFormat = "#,##0.00"
''    '****Begin***************************************************************
''    nPost = nPost + 1
''    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''    MatrixHojaETemp(2, nPost) = "EGRESO TOTALES"
''    MatrixHojaETemp(3, nPost) = fSumaSubnU
''    MatrixHojaETemp(6, nPost) = 2
''    MatrixHojaETemp(7, nPost) = 4
''    '*****End****************************************************************
''    '****Begin***************************************************************
''    nPost = nPost + 1
''    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
''    MatrixHojaETemp(2, PosicionXs) = "EXCEDENTE MENSUAL"
''    MatrixHojaETemp(3, PosicionXs) = fSumaSubnUG1
''    MatrixHojaETemp(6, PosicionXs) = 2
''    MatrixHojaETemp(7, PosicionXs) = 3
''    '*****End****************************************************************
''    xlHoja1.Cells(contador3 + 1, PosicionX) = "EGRESO TOTALES"
''    xlHoja1.Cells(contador3 + 1, PosicionX + 1) = fSumaSubnU
''    xlHoja1.Cells(contador3, PosicionX2) = "EXCEDENTE MENSUAL"
''    xlHoja1.Cells(contador3, PosicionX2 + 1) = fSumaSubnUG1
''    rs.Close
''
''    With xlHoja1.PageSetup
''        .LeftHeader = ""
''        .CenterHeader = ""
''        .RightHeader = ""
''        .LeftFooter = ""
''        .CenterFooter = ""
''        .RightFooter = ""
''
''        .PrintHeadings = False
''        .PrintGridlines = False
''        .PrintComments = xlPrintNoComments
''        .CenterHorizontally = True
''        .CenterVertically = False
''        .Orientation = xlLandscape
''        .Draft = False
''        .FirstPageNumber = xlAutomatic
''        .Order = xlDownThenOver
''        .BlackAndWhite = False
''        .Zoom = 55
''    End With
''End Sub
Private Sub ImprimeHojaEvaluacionExcelCab(ByRef rs As ADODB.Recordset, ByRef cNomCliente As String, ByRef cCodCliente As String, ByRef cRucC As String, ByRef cDNIC As String, ByRef sEmpresa As String, ByRef sOficina As String, ByRef dFecha As Date)
    
    Dim PosicionX As Integer
    Dim PosicionY As Integer
    Dim PosicionX2 As Integer
    Dim contador As Integer
    Dim contador2 As Integer
    Dim contador3 As Integer
    Dim contador4  As Integer
    Dim fSumaIte As Double
    
    Dim fSumaSubnP As Double
    Dim fSumaSubnN As Double
    Dim fSumaSubnU As Double
    Dim fSumaSubnT As Double
    
    Dim fSumaSubnP2 As Double
    Dim fSumaSubnN2 As Double
    Dim fSumaSubnU2 As Double
    Dim fSumaSubnT2 As Double
    '******* Sumatoria de PosYs
    Dim fSumaSubnPG1 As Double
    Dim fSumaSubnNG1 As Double
    Dim fSumaSubnUG1 As Double
    Dim fSumaSubnTG1 As Double
    '**************************
    '******* Sumatoria de PosYs2
    Dim fSumaSubnPG2 As Double
    Dim fSumaSubnNG2 As Double
    Dim fSumaSubnUG2 As Double
    Dim fSumaSubnTG2 As Double
    '**************************
    Dim fCostVe As Double
    Dim PosY As Double
    Dim PosY2 As Double
    Dim PosYs As Double 'Guarda Posicion de Codigo de 4 caracteres
    Dim PosYs2 As Double ' Guarda Posicion de 1001
    Dim PosicionXs As Integer
    Dim PosicionX2s As Integer
    Dim nTipoIm As Integer
    Dim nSATNe As Double
    Dim nSATEm As Double
    Dim nSATTo As Double
    Dim nMargenB As Double
    Dim nIngreT As Double
    Dim nEgresT As Double
    Dim nInNeto As Double
    
    nTipoIm = 3
    contador = 6
    contador2 = 6
    nPost = 0
     Do While Not rs.EOF
        If Left(Trim(rs!cCodHojEval), 2) = "10" Then
            contador = contador + 1
            If contador = 7 Then
                If Trim(rs!cCodHojEval) = "10" Then
                    '***Begin************************************************
                    nPos1 = nPost + 1
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = Trim(rs!cCodHojEval)
                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = "0"
                    MatrixHojaETemp(2, nPost) = ""
                    MatrixHojaETemp(3, nPost) = ""
                    MatrixHojaETemp(4, nPost) = ""
                    MatrixHojaETemp(5, nPost) = ""
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    '***End**************************************************
                    '**Begin**********************
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = "0"
                    MatrixHojaETemp(3, nPost) = "PERSONAL"
                    MatrixHojaETemp(4, nPost) = "NEGOCIO"
                    MatrixHojaETemp(5, nPost) = "TOTAL"
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    
                    '**End**********************
                End If
            Else
            '****Incluye datos al arreglo ---(Accion 01)
                
                fSumaIte = rs!nPersonal + rs!nNegocio
                fSumaSubnP = fSumaSubnP + rs!nPersonal
                fSumaSubnN = fSumaSubnN + rs!nNegocio
                fSumaSubnT = fSumaSubnT + fSumaIte
                fSumaSubnPG1 = fSumaSubnPG1 + rs!nPersonal
                fSumaSubnNG1 = fSumaSubnNG1 + rs!nNegocio
                fSumaSubnTG1 = fSumaSubnTG1 + fSumaIte
                
                If Len(Trim(rs!cCodHojEval)) = 6 And (Left(Trim(rs!cCodHojEval), 4) = "1001" Or Left(Trim(rs!cCodHojEval), 4) = "1002") Then
                    fSumaSubnP2 = fSumaSubnP2 + rs!nPersonal
                    fSumaSubnN2 = fSumaSubnN2 + rs!nNegocio
                    fSumaSubnT2 = fSumaSubnT2 + rs!nNegocio
                End If
                
                If Trim(rs!cCodHojEval) = "1004" Then
                    '****Begin*********************************************************
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(2, nPost) = "ACTIVO TOTAL"
                    MatrixHojaETemp(3, nPost) = fSumaSubnPG1
                    MatrixHojaETemp(4, nPost) = fSumaSubnNG1
                    MatrixHojaETemp(5, nPost) = fSumaSubnTG1
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    nSATNe = fSumaSubnPG1
                    nSATEm = fSumaSubnNG1
                    nSATTo = fSumaSubnTG1
                    '*****End**********************************************************
                    
                    fSumaSubnPG1 = 0
                    fSumaSubnNG1 = 0
                    fSumaSubnTG1 = 0
                    contador = contador + 1
                End If
                '****Begin**************************************************************
                If (rs!nNegocio <> 0 Or rs!nPersonal <> 0 Or rs!nNegocio <> 0) Or Len(Trim(rs!cCodHojEval)) = 4 Then
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = Trim(rs!cCodHojEval)
                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
                    MatrixHojaETemp(3, nPost) = CStr(rs!nPersonal)
                    MatrixHojaETemp(4, nPost) = CStr(rs!nNegocio)
                    MatrixHojaETemp(5, nPost) = CStr(fSumaIte)
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                End If
                '*****End****************************************************************
                '***Toma todos los codigos con 4 caracteres
                If Len(Trim(rs!cCodHojEval)) = 4 Then
                If PosY <> 0 Then
                    If Trim(rs!cCodHojEval) <> "1002" Then
                    '****Begin**************************************************************
                    MatrixHojaETemp(3, PosYs) = CStr(fSumaSubnP)
                    MatrixHojaETemp(4, PosYs) = CStr(fSumaSubnN)
                    MatrixHojaETemp(5, PosYs) = CStr(fSumaSubnT)
                    MatrixHojaETemp(6, PosYs) = 1
                    MatrixHojaETemp(7, PosYs) = 1
                    '*****End****************************************************************
                        If Trim(rs!cCodHojEval) = "1003" Then '1004
                            '****Begin**************************************************************
                            MatrixHojaETemp(3, PosYs2) = CStr(fSumaSubnP2)
                            MatrixHojaETemp(4, PosYs2) = CStr(fSumaSubnN2)
                            MatrixHojaETemp(5, PosYs2) = CStr(fSumaSubnT2)
                            MatrixHojaETemp(6, PosYs2) = 1
                            MatrixHojaETemp(7, PosYs2) = 1
                            '*****End****************************************************************
                        End If

                    End If
                End If
                    If Trim(rs!cCodHojEval) = "1001" Then
                        PosY2 = PosY
                        PosYs2 = nPost
                    End If
                    PosY = 0
                    fSumaSubnP2 = fSumaSubnP2 + fSumaSubnP
                    fSumaSubnN2 = fSumaSubnN2 + fSumaSubnN
                    fSumaSubnT2 = fSumaSubnT2 + fSumaSubnT
                    fSumaSubnP = 0
                    fSumaSubnN = 0
                    fSumaSubnT = 0
                PosY = contador
                PosYs = nPost
                End If
            End If
        End If
         If Left(Trim(rs!cCodHojEval), 2) = "20" Then
          contador2 = contador2 + 1
         If contador2 = 7 Then
                If Trim(rs!cCodHojEval) = "20" Then
                    PosY = 0
                    PosYs = 0
                    '****Begin**************************************************************
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(2, nPost) = "PASIVO TOTAL"
                    MatrixHojaETemp(3, nPost) = CStr(fSumaSubnPG1)
                    MatrixHojaETemp(4, nPost) = CStr(fSumaSubnNG1)
                    MatrixHojaETemp(5, nPost) = CStr(fSumaSubnTG1)
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    '*****End****************************************************************
                    '****Begin**************************************************************
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    'MatrixHojaETemp(1, nPost) = trim(RS!cCodHojEval)
                    MatrixHojaETemp(2, nPost) = "PATRIMONIO"
                    MatrixHojaETemp(3, nPost) = CStr(nSATNe - fSumaSubnPG1)
                    MatrixHojaETemp(4, nPost) = CStr(nSATEm - fSumaSubnNG1)
                    MatrixHojaETemp(5, nPost) = CStr(nSATTo - fSumaSubnTG1)
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    '*****End****************************************************************
                    '****Begin**************************************************************
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(2, nPost) = "TOTAL PASIVO Y PATRIMONIO"
                    MatrixHojaETemp(3, nPost) = CStr((nSATNe - fSumaSubnPG1) + fSumaSubnPG1)
                    MatrixHojaETemp(4, nPost) = CStr((nSATEm - fSumaSubnNG1) + fSumaSubnNG1)
                    MatrixHojaETemp(5, nPost) = CStr((nSATTo - fSumaSubnTG1) + fSumaSubnTG1)
                    MatrixHojaETemp(6, nPost) = 1
                    MatrixHojaETemp(7, nPost) = 1
                    '*****End****************************************************************
                    '****Begin**************************************************************
                    nPost = nPost + 1
                    nPos2 = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = Trim(rs!cCodHojEval)
                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
                    MatrixHojaETemp(6, nPost) = 2
                    MatrixHojaETemp(7, nPost) = 2
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = "0"
                    MatrixHojaETemp(2, nPost) = ""
                    MatrixHojaETemp(3, nPost) = ""
                    MatrixHojaETemp(4, nPost) = ""
                    MatrixHojaETemp(5, nPost) = ""
                    MatrixHojaETemp(6, nPost) = 2
                    MatrixHojaETemp(7, nPost) = 2
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = "0"
                    MatrixHojaETemp(2, nPost) = ""
                    MatrixHojaETemp(3, nPost) = ""
                    MatrixHojaETemp(4, nPost) = ""
                    MatrixHojaETemp(5, nPost) = ""
                    MatrixHojaETemp(6, nPost) = 2
                    MatrixHojaETemp(7, nPost) = 2
                    '*****End****************************************************************
                End If
            Else
            If Trim(rs!cCodHojEval) = "2002" Then
                '****Begin**************************************************************
                nPost = nPost + 1
                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                MatrixHojaETemp(2, nPost) = "MARGEN BRUTO"
                MatrixHojaETemp(3, nPost) = CStr(fSumaSubnU - fCostVe)
                MatrixHojaETemp(6, nPost) = 2
                MatrixHojaETemp(7, nPost) = 2
                nMargenB = fSumaSubnU - fCostVe
                '*****End****************************************************************
                contador2 = contador2 + 1
                nPost = nPost + 1
                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                MatrixHojaETemp(1, nPost) = "0"
                MatrixHojaETemp(2, nPost) = ""
                MatrixHojaETemp(3, nPost) = ""
                MatrixHojaETemp(4, nPost) = ""
                MatrixHojaETemp(5, nPost) = ""
                MatrixHojaETemp(6, nPost) = 2
                MatrixHojaETemp(7, nPost) = 2
            End If
                '****Begin***************************************************************
                If (rs!nUnico <> 0) Or Len(Trim(rs!cCodHojEval)) = 4 Then
                nPost = nPost + 1
                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                MatrixHojaETemp(1, nPost) = Trim(rs!cCodHojEval)
                MatrixHojaETemp(2, nPost) = rs!cDescripcion
                MatrixHojaETemp(3, nPost) = CStr(rs!nUnico)
                MatrixHojaETemp(6, nPost) = 2
                MatrixHojaETemp(7, nPost) = 2
                End If
                '*****End****************************************************************
             If Trim(rs!cCodHojEval) <> "200104" Then
                fSumaSubnU = fSumaSubnU + rs!nUnico
             End If
             fSumaSubnUG1 = fSumaSubnUG1 + rs!nUnico
             If Len(Trim(rs!cCodHojEval)) = 4 Then
                If PosYs <> 0 Then
                    '****Begin***************************************************************
                    'nPost = nPost + 1
                    'ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(3, PosYs) = CStr(fSumaSubnU)
                    MatrixHojaETemp(6, PosYs) = 2
                    MatrixHojaETemp(7, PosYs) = 2
                    fSumaSubnU = 0
                    '*****End****************************************************************
                    'xlHoja1.Cells(PosY, 8) = fSumaSubnU
                End If                '200104
               
                 PosY = contador2
                 PosYs = nPost
             End If
              If Trim(rs!cCodHojEval) = "200104" Then
                    fCostVe = rs!nUnico
                End If
            End If
         End If
         If Left(Trim(rs!cCodHojEval), 2) = "30" Then
                contador4 = contador4 + 1
                If contador4 = 1 Then
                    PosicionX = 7
                   contador3 = contador + 6
                Else
                   contador3 = contador3 + 1
                End If
            
               If contador4 = 1 Then
                   If Trim(rs!cCodHojEval) = "30" Then
                    'Otros Egresos
                    MatrixHojaETemp(3, PosYs) = CStr(fSumaSubnU)
                    MatrixHojaETemp(6, PosYs) = 2
                    MatrixHojaETemp(7, PosYs) = 2
                   
                     '****Begin***************************************************************
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(2, nPost) = "INGRESO NETO"
                    MatrixHojaETemp(3, nPost) = CStr(nMargenB - fSumaSubnU)
                    MatrixHojaETemp(6, nPost) = 2
                    MatrixHojaETemp(7, nPost) = 2
                    nInNeto = nMargenB - fSumaSubnU
                    fSumaSubnU = 0
                    fSumaSubnUG1 = 0
                    PosY = 0
                    PosYs = 0
                    '*****End****************************************************************
                    '****Begin***************************************************************
                    nPos3 = nPost + 1
                    nPost = nPost + 1
                    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                    MatrixHojaETemp(1, nPost) = Trim(rs!cCodHojEval)
                    MatrixHojaETemp(2, nPost) = rs!cDescripcion
                    MatrixHojaETemp(6, nPost) = 3
                    MatrixHojaETemp(7, nPost) = 3
                    '*****End****************************************************************
                   End If
                Else
                
                If Trim(rs!cCodHojEval) = "3002" Then
                nPost = nPost + 1
                PosicionX2 = PosicionX
                '****Begin***************************************************************
                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                MatrixHojaETemp(2, nPost) = "INGRESO TOTALES"
                MatrixHojaETemp(3, nPost) = CStr(nInNeto + fSumaSubnU)
                nIngreT = fSumaSubnU
                MatrixHojaETemp(6, nPost) = 3
                MatrixHojaETemp(7, nPost) = 3
                nPost = nPost + 1
                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                PosicionXs = nPost
                nPost = nPost + 1
                ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                nTipoIm = 4
                '*****End****************************************************************
                 contador3 = contador + 7
                 PosicionX = 2
                 fSumaSubnU = 0
                End If
                    '****Begin***************************************************************
                    If (rs!nUnico <> 0) Or Len(Trim(rs!cCodHojEval)) = 4 Or Trim(rs!cCodHojEval) = "300101" Then
                        nPost = nPost + 1
                        ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
                        MatrixHojaETemp(1, nPost) = Trim(rs!cCodHojEval)
                        MatrixHojaETemp(2, nPost) = rs!cDescripcion
                        MatrixHojaETemp(3, nPost) = CStr(rs!nUnico)
                    End If
                    If MatrixHojaETemp(1, nPost) = "300101" Then
                    MatrixHojaETemp(3, nPost) = CStr(nInNeto)
                    End If
                    
                    MatrixHojaETemp(6, nPost) = 3
                    MatrixHojaETemp(7, nPost) = nTipoIm
                    '*****End****************************************************************
                    fSumaSubnU = fSumaSubnU + rs!nUnico
                    fSumaSubnUG1 = fSumaSubnUG1 + rs!nUnico
                End If
         End If
        rs.MoveNext
    Loop
    '****Begin***************************************************************
    nPost = nPost + 1
    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
    MatrixHojaETemp(2, nPost) = "EGRESO TOTALES"
    MatrixHojaETemp(3, nPost) = fSumaSubnU
    MatrixHojaETemp(6, nPost) = 3
    MatrixHojaETemp(7, nPost) = 4
    nEgresT = fSumaSubnU
    '*****End****************************************************************
    '****Begin***************************************************************
    nPost = nPost + 1
    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
    MatrixHojaETemp(2, PosicionXs) = "EXCEDENTE MENSUAL"
    MatrixHojaETemp(3, PosicionXs) = (nIngreT + nInNeto) - nEgresT
    MatrixHojaETemp(6, PosicionXs) = 3
    MatrixHojaETemp(7, PosicionXs) = 3
     nPost = nPost + 1
     nPos4 = PosicionXs + 1
    ReDim Preserve MatrixHojaETemp(1 To 7, 0 To nPost + 1)
    MatrixHojaETemp(2, PosicionXs + 1) = ""
    MatrixHojaETemp(3, PosicionXs + 1) = ""
    MatrixHojaETemp(6, PosicionXs + 1) = 3
    MatrixHojaETemp(7, PosicionXs + 1) = 3
    '*****End****************************************************************
    rs.Close
End Sub
Private Function imprime_HojaEvalucacion(cPersona As String, cPersCod As String, cPerRUC As String, cPerDNI As String, sEmpresa As String, sOfic As String, gdFecSis As Date) As String
 Dim sCadImp As String
 Dim lsSaltoLin As String
 Dim lsNegritaOn As String
 Dim lsNegritaOff As String
 lsSaltoLin = Chr$(10)
 'lsNegritaOn = Chr$(27) + Chr$(71)
 'lsNegritaOff = Chr$(27) + Chr$(72)
 sCadImp = ""
 sCadImp = sCadImp & String(130, "-") & lsSaltoLin
sCadImp = sCadImp & "" & ImpreFormat(sEmpresa, 70) & ImpreFormat("FECHA:", 8) & gdFecSis & lsSaltoLin
sCadImp = sCadImp & "" & ImpreFormat(sOfic, 70) & ImpreFormat("USUARIO:", 8) & gsCodUser & lsSaltoLin
sCadImp = sCadImp & Space(30) & "HOJA DE EVALUACION" & lsSaltoLin
sCadImp = sCadImp & ImpreFormat("CODIDO CLIENTE: ", 30) & ImpreFormat(cPersCod, 40) & ImpreFormat("DNI: ", 5) & ImpreFormat(cPerDNI, 100) & lsSaltoLin
sCadImp = sCadImp & ImpreFormat("CLIENTE: ", 30) & ImpreFormat(cPersona, 40) & ImpreFormat("RUC: ", 5) & ImpreFormat(cPerRUC, 100) & lsSaltoLin
sCadImp = sCadImp & String(130, "-") & lsSaltoLin
sCadImp = sCadImp & lsSaltoLin
 Dim i As Integer
 For i = 1 To nPost - 1
  If MatrixHojaETemp(7, i) = "1" Then
  If Len(Trim(MatrixHojaETemp(1, i))) = 2 Or Len(Trim(MatrixHojaETemp(1, i))) = 1 Then
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & ImpreFormat(MatrixHojaETemp(2, i), 50) & ImpreFormat(MatrixHojaETemp(4, i), 8) & ImpreFormat(MatrixHojaETemp(3, i), 8) & ImpreFormat(MatrixHojaETemp(5, i), 8)
    If (i <= nPos2 - 2 And (i + nPos2 - 2) <= nPos3) And MatrixHojaETemp(7, i + (nPos2 - 2)) = "2" Then
    If Len(Trim(ImpreFormat(MatrixHojaETemp(2, i + (nPos2 - 2)), 50) & ImpreFormat(MatrixHojaETemp(3, i + (nPos2 - 2)), 8))) = 0 Then
        sCadImp = sCadImp
    Else
        sCadImp = sCadImp & Space(13) & ImpreFormat(MatrixHojaETemp(2, i + (nPos2 - 2)), 50) '& ImpreFormat(Round(MatrixHojaETemp(3, i + (nPos2 - 2)), 2), 8)
        
    End If
      
    End If
    sCadImp = sCadImp & lsSaltoLin
  Else
    sCadImp = sCadImp & lsNegritaOff
    sCadImp = sCadImp & ImpreFormat(MatrixHojaETemp(2, i), 50) & ImpreFormat(Round(MatrixHojaETemp(4, i), 2), 8) & ImpreFormat(Round(MatrixHojaETemp(3, i), 2), 8) & ImpreFormat(Round(MatrixHojaETemp(5, i), 2), 8)
    'If i <= 33 And (i + 35) <= 68 Then
    If i <= nPos2 And (i + (nPos2 - 2)) <= nPos3 Then
        If MatrixHojaETemp(2, i + (nPos2 - 2)) <> "" And MatrixHojaETemp(7, i + (nPos2 - 2)) = 2 Then
        sCadImp = sCadImp & Space(10) & ImpreFormat(MatrixHojaETemp(2, i + (nPos2 - 2)), 50) & ImpreFormat(Round(MatrixHojaETemp(3, i + (nPos2 - 2)), 2), 8)
        End If
    End If
    sCadImp = sCadImp & lsSaltoLin
  End If
End If
If MatrixHojaETemp(7, i) = "3" Then

 If Len(Trim(MatrixHojaETemp(1, i))) = 2 Or Len(Trim(MatrixHojaETemp(1, i))) = 1 Then
    sCadImp = sCadImp & lsNegritaOn
    If MatrixHojaETemp(1, i) = "30" Then
        sCadImp = sCadImp & String(130, "-") & lsSaltoLin
        sCadImp = sCadImp & lsSaltoLin
        sCadImp = sCadImp & ImpreFormat(MatrixHojaETemp(2, i), 50) & ImpreFormat(MatrixHojaETemp(3, i), 8)
        sCadImp = sCadImp & lsSaltoLin
    Else
        sCadImp = sCadImp & ImpreFormat(MatrixHojaETemp(2, i + (nPos4 - nPos3)), 50) & ImpreFormat(MatrixHojaETemp(3, i + (nPos4 - nPos3)), 8)
        sCadImp = sCadImp & Space(30) & ImpreFormat(MatrixHojaETemp(2, i), 50) & ImpreFormat(MatrixHojaETemp(3, i), 8)
    End If
    sCadImp = sCadImp & lsSaltoLin
  Else
        sCadImp = sCadImp & lsNegritaOff
        If MatrixHojaETemp(2, i + (nPos4 - nPos3)) <> "" Then
          If MatrixHojaETemp(3, i + (nPos4 - nPos3)) = "" Then
             MatrixHojaETemp(3, i + (nPos4 - nPos3)) = 0#
          End If
          If Len(MatrixHojaETemp(1, i)) = 4 Then
            sCadImp = sCadImp & ImpreFormat(MatrixHojaETemp(2, i + (nPos4 - nPos3)), 50) & Space(8)
          Else
          If Len(Trim(MatrixHojaETemp(2, i + (nPos4 - nPos3)))) > 0 Then
          sCadImp = sCadImp & ImpreFormat(MatrixHojaETemp(2, i + (nPos4 - nPos3)), 50) & ImpreFormat(Round(IIf(MatrixHojaETemp(3, i + (nPos4 - nPos3)), MatrixHojaETemp(3, i + (nPos4 - nPos3)), 0), 2), 8)
          End If
            
          End If
        Else
          sCadImp = sCadImp & Space(38)
        End If
        If MatrixHojaETemp(2, i) <> "" Then
            If Len(MatrixHojaETemp(1, i)) = 4 Then
                sCadImp = sCadImp & Space(30) & ImpreFormat(MatrixHojaETemp(2, i), 50)
            Else
                sCadImp = sCadImp & Space(30) & ImpreFormat(MatrixHojaETemp(2, i), 50) & ImpreFormat(Round(MatrixHojaETemp(3, i), 2), 8)
            End If

        Else
            sCadImp = sCadImp & Space(38)
        End If
        sCadImp = sCadImp & lsSaltoLin
  End If
End If
 Next i
imprime_HojaEvalucacion = sCadImp
End Function
'**End**ALPA******************************************************************
'*****************************************************************************

Private Sub CmdEditar_Click()
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        HabilitaBalance True
        HabilitaIngresosEgresos False
        SSTFuentes.Tab = 3
        
        'peac 20071227
        txtFecEEFF.Enabled = False
        
    Else
        HabilitaIngresosEgresos False
        HabilitaBalance True
        
        'SSTFuentes.Tab = 1
        '***PEAC 20080402
        SSTFuentes.Tab = 3
        
        'peac 20071227
        txtFecEEFF.Enabled = True
        
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
    
    'WIOR 20140319****************************
    If Not fbEditarEF Then
        Call HabilitaEstFinancieros(True)
    End If
    'WIOR FIN ********************************
End Sub

Private Sub CmdEliminar_Click()
Dim oPersonaD As COMDpersona.DCOMPersona

    If MsgBox("Se va a Eliminar la Fuente de Ingreso de Fecha :" & Me.CmbFecha.Text & ", Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oPersonaD = New COMDpersona.DCOMPersona
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
            
            'peac 20071227
            Call oPersona.ActualizarFteIndIngFecEEFF(CDate(txtFecEEFF.Text), nIndice, nIndiceAct)
            
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
   Call HabilitaEstFinancieros(False) 'WIOR 20140319
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
    
    'peac 20071227
    txtFecEEFF.Enabled = False
    
    CmbFecha_Click
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        'SSTFuentes.TabVisible(1) = False
        
        '*** PEAC 20080402
        SSTFuentes.TabVisible(3) = True
        
        SSTFuentes.TabVisible(0) = False
        SSTFuentes.Tab = 3
    Else
        SSTFuentes.TabVisible(0) = False
        
        'SSTFuentes.TabVisible(1) = True
        '*** PEAC 20080402
        SSTFuentes.TabVisible(3) = True
        
        'SSTFuentes.Tab = 1
        SSTFuentes.Tab = 3
    End If
    Call HabilitaEstFinancieros(False) 'WIOR 20140320
End Sub

Private Sub cmdGrabarEval_Click()
    Dim J As Integer
    Dim i As Integer
    Dim NCaDAr As Integer
    NCaDAr = 0
If val(txtMonto.Text) + val(txtMonto2.Text) = 0 Or Trim(Left(Me.cboGrupoHojEval.Text, 30)) = "" _
    Or Trim(Left(Me.cboTipoEval.Text, 30)) = "" Or Trim(Left(cboConceptoEval.Text, 30)) = "" Then
    MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
    Exit Sub
End If

If nDat = 1 Then
    For i = 0 To nPos
        If MatrixHojaEval(7, i) = val(Trim(Right(cboConceptoEval.Text, 8))) Then
            MsgBox "Este dato ya fue registrado...", vbInformation, "Aviso"
            txtMonto.Text = "0.00"
            txtMonto2.Text = "0.00"
            Exit Sub
        End If
    Next i
End If

    nDat = 1
    FEHojaEval.AdicionaFila
     'Dim MatrixHojaEval() As String
     If FEHojaEval.row = 1 Then
        ReDim MatrixHojaEval(1 To 7, 0 To 0)
     End If
     'Dim nPos As Integer
     nPos = FEHojaEval.row - 1
     MatrixHojaEval(1, nPos) = FEHojaEval.row
     ReDim Preserve MatrixHojaEval(1 To 7, 0 To UBound(MatrixHojaEval, 2) + 1)
     If nPos >= 1 Then
     For J = 0 To nPos - 1
        If Trim(Right(cboConceptoEval.Text, 8)) > MatrixHojaEval(7, nPos - 1) Then
            i = nPos
            Exit For
        End If
        If Trim(Right(cboConceptoEval.Text, 8)) < MatrixHojaEval(7, 0) Then
            i = 0
            Exit For
        End If
        If Trim(Right(cboConceptoEval.Text, 8)) > MatrixHojaEval(7, J) And Trim(Right(cboConceptoEval.Text, 8)) <= MatrixHojaEval(7, J + 1) Then
            i = J + 1
            Exit For
            '***********
        End If
    Next J
    For J = nPos - 1 To i Step -1
            MatrixHojaEval(1, J + 1) = MatrixHojaEval(1, J)
            MatrixHojaEval(2, J + 1) = MatrixHojaEval(2, J)
            MatrixHojaEval(3, J + 1) = MatrixHojaEval(3, J)
            MatrixHojaEval(4, J + 1) = MatrixHojaEval(4, J)
            MatrixHojaEval(5, J + 1) = MatrixHojaEval(5, J)
            MatrixHojaEval(6, J + 1) = MatrixHojaEval(6, J)
            MatrixHojaEval(7, J + 1) = MatrixHojaEval(7, J)
     Next J
            MatrixHojaEval(1, i) = Trim(Left(Me.cboGrupoHojEval.Text, 30))
            MatrixHojaEval(2, i) = Trim(Left(Me.cboTipoEval.Text, 30))
            MatrixHojaEval(3, i) = Trim(Left(cboConceptoEval.Text, 30))
            If CInt(Right(cboGrupoHojEval.Text, 2)) = 10 Then
               MatrixHojaEval(4, i) = Me.txtMonto.Text
               MatrixHojaEval(5, i) = Me.txtMonto2.Text
               MatrixHojaEval(6, i) = ""
            Else
               MatrixHojaEval(6, i) = Me.txtMonto2.Text
               MatrixHojaEval(4, i) = ""
               MatrixHojaEval(5, i) = ""
            End If
            MatrixHojaEval(7, i) = val(Trim(Right(cboConceptoEval.Text, 8)))
    Else
            MatrixHojaEval(1, nPos) = Trim(Left(Me.cboGrupoHojEval.Text, 30))
            MatrixHojaEval(2, nPos) = Trim(Left(Me.cboTipoEval.Text, 30))
            MatrixHojaEval(3, nPos) = Trim(Left(cboConceptoEval.Text, 30))
           If CInt(Right(cboGrupoHojEval.Text, 2)) = 10 Then
               MatrixHojaEval(4, nPos) = Me.txtMonto.Text
               MatrixHojaEval(5, nPos) = Me.txtMonto2.Text
            Else
               MatrixHojaEval(6, nPos) = Me.txtMonto2.Text
            End If
            MatrixHojaEval(7, nPos) = val(Trim(Right(cboConceptoEval.Text, 8)))
    End If
    'FEHojaEval.Clear
    For i = 0 To nPos
        FEHojaEval.EliminaFila (1)
    Next i
    For i = 0 To nPos
        FEHojaEval.AdicionaFila
        FEHojaEval.TextMatrix(FEHojaEval.row, 1) = MatrixHojaEval(1, i)
        FEHojaEval.TextMatrix(FEHojaEval.row, 2) = MatrixHojaEval(2, i)
        FEHojaEval.TextMatrix(FEHojaEval.row, 3) = MatrixHojaEval(3, i)
        FEHojaEval.TextMatrix(FEHojaEval.row, 4) = MatrixHojaEval(4, i)
        FEHojaEval.TextMatrix(FEHojaEval.row, 5) = MatrixHojaEval(5, i)
        FEHojaEval.TextMatrix(FEHojaEval.row, 6) = MatrixHojaEval(6, i)
        FEHojaEval.TextMatrix(FEHojaEval.row, 7) = MatrixHojaEval(7, i)
        NCaDAr = 1
    
    Next

    txtMonto.Text = "0.00"
    txtMonto2.Text = "0.00"
    
    'cboGrupoHojEval.SetFocus
    If NCaDAr = 1 Then
        CboTipoFte.Enabled = False
    End If
    Me.txtCodEval.SetFocus

End Sub

Private Sub cmdImprimeCodHojEval_Click()

    Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oNCredito As COMNCredito.NCOMCredDoc
   
    Set oPrev = New previo.clsprevio
     
    Set oNCredito = New COMNCredito.NCOMCredDoc
    sCadImp = oNCredito.ImprimeCodHojEval()
    Set oNCredito = Nothing
       
    previo.Show sCadImp, "Códigos de la Hoja de Evaluación.", False
    Set oPrev = Nothing

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
Dim oPrev As previo.clsprevio
'Dim oPersonaD As COMDPersona.DCOMPersona

Dim bCostoProd As Boolean

    If ChkCostoProd.value = vbChecked Then
        bCostoProd = True
    Else
        bCostoProd = False
    End If

    Set oPrev = New previo.clsprevio
    'Set oPersonaD = New COMDPersona.DCOMPersona

    'Call LlenarDatosFteIngreso(oPersonaD)

    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        'psPersCod, nIndice, gsNomAge, gdFecSis, bCostoProd, ""
        'sCadImp = oPersonaD.GenerarImpresionFteIngresoDependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
        sCadImp = oPersona.GenerarImpresionFteIngresoDependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
        
    Else
        'sCadImp = oPersona.GenerarImpresionFteIngresoIndependiente(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
        sCadImp = oPersona.GenerarImpresionFteIngresoIndependiente_CS(nIndice, gsNomAge, gdFecSis, bCostoProd, "")
    End If
    'Set oPersonaD = Nothing
    previo.Show sCadImp, "Evaluacion de Fuentes de Ingreso", False
    Set oPrev = Nothing
    
    'Call ImprimeHistoriaCrediticia(Trim(TxtBRazonSoc.Text), oPersona.PersCodigo, lcNumFuente)
    
End Sub
'-----------

Sub LlenarDatosFteIngreso(ByVal poPersona As COMDpersona.DCOMPersona)

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

Dim oPersona As COMDpersona.DCOMPersonas
Set oPersona = New COMDpersona.DCOMPersonas
Dim rsMoneda As ADODB.Recordset
Dim rsTipoFte As ADODB.Recordset
Dim rsTipoCul As ADODB.Recordset
Dim rsUnidad As ADODB.Recordset
Dim rsFIDep As ADODB.Recordset
Dim rsFIInd As ADODB.Recordset
Dim rsFICos As ADODB.Recordset

'*** PEAC 20080412
Set oPers = New COMDCredito.DCOMCredito
Set rsOHE = oPers.ObtieneCodEvaluacion(nTipEv)
Set oPers = Nothing
'*** FIN PEAC

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
    lcNumFuente = poPersona.ObtenerFteIngcNumFuente(pnIndice)
    If pbCargarControles Then
        Call Llenar_Combo_con_Recordset(rsTipoFte, CboTipoFte)
        Call Llenar_Combo_con_Recordset(rsMoneda, CboMoneda)
        Call Llenar_Combo_con_Recordset(rsTipoCul, CboTpoCul)
        Call Llenar_Combo_con_Recordset(rsUnidad, CboUnidad)
    End If
    Call CargaDatosFteIngreso(pnIndice, poPersona, rsFIDep, rsFIInd, rsFICos, pnFteDetalle)
End If

nDat = 0

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
    Dim i As Integer

    TxFecEval.Text = "__/__/____"
    
    'peac 20071227
    txtFecEEFF.Text = "__/__/____"
    
    CmbFecha.Visible = False
    TxFecEval.Visible = True
    If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
        HabilitaBalance True 'False
        HabilitaIngresosEgresos True
        SSTFuentes.Tab = 3
        
        txtFecEEFF.Enabled = False
        
    Else
        HabilitaIngresosEgresos False
        HabilitaBalance True
        
        '***PEAC 20080402
        'SSTFuentes.Tab = 1
        SSTFuentes.Tab = 3
        
        txtFecEEFF.Enabled = True
        
    End If
        
    If Me.ChkCostoProd.value = 1 Then
        HabilitaCostoProd True
    End If
    Call LimpiaFuentesIngreso
    nProcesoActual = 1
    HabilitaMantenimiento False
    ChkCostoProd.Enabled = True
    
    '*** PEAC 20080515
    For i = 0 To nPos
        FEHojaEval.EliminaFila (1)
    Next i
    'WIOR 20140320 ***************************************
    Call LimpiarEstFinan
    Call HabilitaEstFinancieros(True)
    'WIOR FIN ********************************************
End Sub

Private Sub CmdSalirCancelar_Click()
    nPos = 0
    nDat = 0
    Unload Me
End Sub

Private Sub CmdUbigeo_Click()
    vsUbiGeo = Right(frmUbicacionGeo.Inicio(vsUbiGeo), 12)
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()

rsOHE.Find rsOHE(2).Name & " = '" & txtCodEval.Text & "'", , , 1
If Not rsOHE.EOF Then
    cboGrupoHojEval.ListIndex = IndiceListaCombo(cboGrupoHojEval, Mid(txtCodEval.Text, 1, 2))
    cboTipoEval.ListIndex = IndiceListaCombo(cboTipoEval, Mid(txtCodEval.Text, 1, 4))
    cboConceptoEval.ListIndex = IndiceListaCombo(cboConceptoEval, Mid(txtCodEval.Text, 1, 6))
Else
    MsgBox "Código no Existe ...", vbInformation, "Aviso"
End If

End Sub

Private Sub Command2_Click()
    FEHojaEval.EliminaFila (FEHojaEval.row)
    If FEHojaEval.row > 1 Then
    Dim J As Integer
    For J = FEHojaEval.row - 1 To nPos - 1
        MatrixHojaEval(1, J) = MatrixHojaEval(1, J + 1)
        MatrixHojaEval(2, J) = MatrixHojaEval(2, J + 1)
        MatrixHojaEval(3, J) = MatrixHojaEval(3, J + 1)
        MatrixHojaEval(4, J) = MatrixHojaEval(4, J + 1)
        MatrixHojaEval(5, J) = MatrixHojaEval(5, J + 1)
        MatrixHojaEval(6, J) = MatrixHojaEval(6, J + 1)
        MatrixHojaEval(7, J) = MatrixHojaEval(7, J + 1)
    Next J
        nPos = nPos - 1
    Else
        nPos = 0
    End If
End Sub

Private Sub CmdUbigeo_LostFocus()
    If Me.txtFecEEFF.Enabled Then
        Me.txtFecEEFF.SetFocus
    Else
'        If TxFecEval.Visible Then
'            TxFecEval.SetFocus
'        Else
'            CmbFecha.SetFocus
'        End If
    End If
End Sub

Private Sub DTPFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCargo.SetFocus
    End If
End Sub

'Private Sub FEHojaEval_DblClick()
'    FEHojaEval.EliminaFila (FEHojaEval.Row)
'End Sub

Private Sub Form_Load()

    CentraForm Me
    Me.Top = 0
    sImpr = "SI"
    Me.Left = (Screen.Width - Me.Width) / 2
    'me.Left = 600
    bEstadoCargando = False
    nProcesoActual = 0
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    SSTFuentes.TabVisible(2) = False
    SSTFuentes.TabVisible(4) = False 'FRHU 20150311 ERS013-2015 'Los estados financieros ya no se mostraran en esta opcion y no tendran ninguna relacion con la fuente de ingresos.
    Call CargaControles
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
            
            Me.cboGrupoHojEval.SetFocus

'            If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'                TxtIngCon.SetFocus
'            Else
'                txtDisponible.SetFocus
'            End If
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

'lblEgresosB.Caption = Format(CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) _
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

'lblIngresosB.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
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
    Dim oPers As COMDpersona.DCOMPersonas
    Set oPers = New COMDpersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    
    Set rs = oPers.ObtenerDatosReporteFteIngreso(psPersCod)
    
    Set oPers = Nothing
    
    If Not rs.EOF Then
        sRUC = rs!cPersIDnro
        sCiiu = rs!cCIIUdescripcion
        sCondDomicilio = rs!cConsDescripcion
        nNroEmpleados = rs!nPersJurEmpleados
        sMagnitudEmp = rs!magnitud
        
        '05-06-2006
        sDepartamento = rs!cDep
        sProvincia = rs!cProv
        sDistrito = rs!cDist
        sZona = rs!cZon
        '---------------
    End If
End Function

Private Sub TxtCargo_GotFocus()
    fEnfoque TxtCargo
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        Txtcomentarios.SetFocus
    End If
End Sub

Private Sub txtCodEval_GotFocus()
    fEnfoque txtCodEval
End Sub

Private Sub txtCodEval_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If nTipEv > 0 Then
    If Len(txtCodEval.Text) >= 6 Then
        'rsOHE.Find rsOHE(2).Name & " = '" & txtCodEval.Text & "'", , , 1
        If Not rsOHE.EOF Then
        'If Not IndiceListaCombo(cboConceptoEval, Mid(txtCodEval.Text, 1, 8)) Then
            cboGrupoHojEval.ListIndex = IndiceListaCombo(cboGrupoHojEval, Mid(txtCodEval.Text, 1, 2))
            cboTipoEval.ListIndex = IndiceListaCombo(cboTipoEval, Mid(txtCodEval.Text, 1, 4))
            cboConceptoEval.ListIndex = IndiceListaCombo(cboConceptoEval, Mid(txtCodEval.Text, 1, 8))
'            If cboConceptoEval.DataSource < 1 Then
'                MsgBox "Código no Existe ...", vbInformation, "Aviso"
'                Me.txtCodEval.Text = ""
'            End If
            SendKeys "{Tab}", True
        Else
            MsgBox "Código no Existe ...", vbInformation, "Aviso"
            Me.txtCodEval.Text = ""
        End If
    Else
        MsgBox "Código incorrecto ...", vbInformation, "Aviso"
        Me.txtCodEval.Text = ""
    End If
Else
        MsgBox "Seleccionar fuente de Ingreso ...", vbInformation, "Aviso"
        Me.txtCodEval.Text = ""
End If

End If
End Sub


Private Sub txtcompras_Change()
LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)) _
            + CDbl(IIf(Trim(TxtBalIngFam.Text) = "", "0", TxtBalIngFam.Text)) _
            - CDbl(IIf(Trim(TxtBalEgrFam.Text) = "", "0", TxtBalEgrFam.Text)), "#0.00")

'lblEgresosB.Caption = Format(CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) _
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


'WIOR 20140313 *******************************************

Private Sub txtEFActFijo_GotFocus()
fEnfoque txtEFActFijo
End Sub

Private Sub txtEFActFijo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFActFijo, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFActFijo_LostFocus()
    If Trim(txtEFActFijo.Text) = "" Then
        txtEFActFijo.Text = "0.00"
    Else
        txtEFActFijo.Text = Format(txtEFActFijo.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(2)
End Sub

Private Sub txtEFCajaBanco_GotFocus()
fEnfoque txtEFCajaBanco
End Sub

Private Sub txtEFCajaBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFCajaBanco, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFCajaBanco_LostFocus()
    If Trim(txtEFCajaBanco.Text) = "" Then
        txtEFCajaBanco.Text = "0.00"
    Else
        txtEFCajaBanco.Text = Format(txtEFCajaBanco.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub


Private Sub txtEFCapSocial_GotFocus()
fEnfoque txtEFCapSocial
End Sub

Private Sub txtEFCapSocial_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFCapSocial, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFCapSocial_LostFocus()
    If Trim(txtEFCapSocial.Text) = "" Then
        txtEFCapSocial.Text = "0.00"
    Else
        txtEFCapSocial.Text = Format(txtEFCapSocial.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(5)
End Sub

Private Sub txtEFCostVentas_GotFocus()
fEnfoque txtEFCostVentas
End Sub

Private Sub txtEFCostVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFCostVentas, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFCostVentas_LostFocus()
    If Trim(txtEFCostVentas.Text) = "" Then
        txtEFCostVentas.Text = "0.00"
    Else
        txtEFCostVentas.Text = Format(txtEFCostVentas.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(6)
End Sub

Private Sub txtEFCuentaCobrar_GotFocus()
fEnfoque txtEFCuentaCobrar
End Sub

Private Sub txtEFCuentaCobrar_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFCuentaCobrar, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFCuentaCobrar_LostFocus()
    If Trim(txtEFCuentaCobrar.Text) = "" Then
        txtEFCuentaCobrar.Text = "0.00"
    Else
        txtEFCuentaCobrar.Text = Format(txtEFCuentaCobrar.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub

Private Sub txtEFDeudaFinan_GotFocus()
fEnfoque txtEFDeudaFinan
End Sub

Private Sub txtEFDeudaFinan_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFDeudaFinan, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFDeudaFinan_LostFocus()
    If Trim(txtEFDeudaFinan.Text) = "" Then
        txtEFDeudaFinan.Text = "0.00"
    Else
        txtEFDeudaFinan.Text = Format(txtEFDeudaFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub

Private Sub txtEFDeudaFinanC_GotFocus()
fEnfoque txtEFDeudaFinanC
End Sub

Private Sub txtEFDeudaFinanC_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFDeudaFinanC, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFDeudaFinanC_LostFocus()
    If Trim(txtEFDeudaFinanC.Text) = "" Then
        txtEFDeudaFinanC.Text = "0.00"
    Else
        txtEFDeudaFinanC.Text = Format(txtEFDeudaFinanC.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(3)
End Sub

Private Sub txtEFDeudaFinanL_GotFocus()
fEnfoque txtEFDeudaFinanL
End Sub

Private Sub txtEFDeudaFinanL_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFDeudaFinanL, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFDeudaFinanL_LostFocus()
    If Trim(txtEFDeudaFinanL.Text) = "" Then
        txtEFDeudaFinanL.Text = "0.00"
    Else
        txtEFDeudaFinanL.Text = Format(txtEFDeudaFinanL.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(3)
End Sub

Private Sub txtEFExiste_GotFocus()
fEnfoque txtEFExiste
End Sub

Private Sub txtEFExiste_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFExiste, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFExiste_LostFocus()
    If Trim(txtEFExiste.Text) = "" Then
        txtEFExiste.Text = "0.00"
    Else
        txtEFExiste.Text = Format(txtEFExiste.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub

Private Sub txtEFFlujoFinan_GotFocus()
fEnfoque txtEFFlujoFinan
End Sub

Private Sub txtEFFlujoFinan_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFFlujoFinan, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFFlujoFinan_LostFocus()
    If Trim(txtEFFlujoFinan.Text) = "" Then
        txtEFFlujoFinan.Text = "0.00"
    Else
        txtEFFlujoFinan.Text = Format(txtEFFlujoFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(9)
End Sub

Private Sub txtEFFlujoOpe_GotFocus()
fEnfoque txtEFFlujoOpe
End Sub

Private Sub txtEFFlujoOpe_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFFlujoOpe, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFFlujoOpe_LostFocus()
    If Trim(txtEFFlujoOpe.Text) = "" Then
        txtEFFlujoOpe.Text = "0.00"
    Else
        txtEFFlujoOpe.Text = Format(txtEFFlujoOpe.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(9)
End Sub

Private Sub txtEFFujoInv_GotFocus()
fEnfoque txtEFFujoInv
End Sub

Private Sub txtEFFujoInv_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFFujoInv, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFFujoInv_LostFocus()
    If Trim(txtEFFujoInv.Text) = "" Then
        txtEFFujoInv.Text = "0.00"
    Else
        txtEFFujoInv.Text = Format(txtEFFujoInv.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(9)
End Sub

Private Sub txtEFGastoFinan_GotFocus()
fEnfoque txtEFGastoFinan
End Sub

Private Sub txtEFGastoFinan_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFGastoFinan, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFGastoFinan_LostFocus()
    If Trim(txtEFGastoFinan.Text) = "" Then
        txtEFGastoFinan.Text = "0.00"
    Else
        txtEFGastoFinan.Text = Format(txtEFGastoFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(8)
End Sub

Private Sub txtEFGastosME_GotFocus()
fEnfoque txtEFGastosME
End Sub

Private Sub txtEFGastosME_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFGastosME, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFGastosME_LostFocus()
    If Trim(txtEFGastosME.Text) = "" Then
        txtEFGastosME.Text = "0.00"
    Else
        txtEFGastosME.Text = Format(txtEFGastosME.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub

Private Sub txtEFGastosOpe_GotFocus()
fEnfoque txtEFGastosOpe
End Sub

Private Sub txtEFGastosOpe_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFGastosOpe, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFGastosOpe_LostFocus()
    If Trim(txtEFGastosOpe.Text) = "" Then
        txtEFGastosOpe.Text = "0.00"
    Else
        txtEFGastosOpe.Text = Format(txtEFGastosOpe.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(7)
End Sub

Private Sub txtEFIngresoFinan_GotFocus()
fEnfoque txtEFIngresoFinan
End Sub

Private Sub txtEFIngresoFinan_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFIngresoFinan, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFIngresoFinan_LostFocus()
   If Trim(txtEFIngresoFinan.Text) = "" Then
        txtEFIngresoFinan.Text = "0.00"
    Else
        txtEFIngresoFinan.Text = Format(txtEFIngresoFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(8)
End Sub

Private Sub txtEFIngresoME_GotFocus()
fEnfoque txtEFIngresoME
End Sub

Private Sub txtEFIngresoME_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFIngresoME, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFIngresoME_LostFocus()
    If Trim(txtEFIngresoME.Text) = "" Then
        txtEFIngresoME.Text = "0.00"
    Else
        txtEFIngresoME.Text = Format(txtEFIngresoME.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub

Private Sub txtEFPosCambios_GotFocus()
fEnfoque txtEFPosCambios
End Sub

Private Sub txtEFPosCambios_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFPosCambios, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFPosCambios_LostFocus()
    If Trim(txtEFPosCambios.Text) = "" Then
        txtEFPosCambios.Text = "0.00"
    Else
        txtEFPosCambios.Text = Format(txtEFPosCambios.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub

Private Sub txtEFProveedor_GotFocus()
fEnfoque txtEFProveedor
End Sub

Private Sub txtEFProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFProveedor, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFProveedor_LostFocus()
    If Trim(txtEFProveedor.Text) = "" Then
        txtEFProveedor.Text = "0.00"
    Else
        txtEFProveedor.Text = Format(txtEFProveedor.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(3)
End Sub

Private Sub txtEFResulAcum_GotFocus()
fEnfoque txtEFResulAcum
End Sub

Private Sub txtEFResulAcum_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFResulAcum, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFResulAcum_LostFocus()
    If Trim(txtEFResulAcum.Text) = "" Then
        txtEFResulAcum.Text = "0.00"
    Else
        txtEFResulAcum.Text = Format(txtEFResulAcum.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(5)
End Sub

Private Sub txtEFVentas_GotFocus()
fEnfoque txtEFVentas
End Sub

Private Sub txtEFVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEFVentas, KeyAscii, , , True)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub

Private Sub txtEFVentas_LostFocus()
    If Trim(txtEFVentas.Text) = "" Then
        txtEFVentas.Text = "0.00"
    Else
        txtEFVentas.Text = Format(txtEFVentas.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(6)
End Sub

'WIOR FIN ************************************************

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

'peac 20071227
Private Sub txtFecEEFF_GotFocus()
    fEnfoque txtFecEEFF
End Sub

'peac 20071227
Private Sub txtFecEEFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If CDate(ldFecEval) >= CDate(txtFecEEFF) Then
'            MsgBox "No Puede Ingresar una Fecha Igual o Menor a la Ultima Fecha de EE.FF.", vbInformation, "Aviso"
'            Exit Sub
'        End If

'        If Trim(CboTipoFte.Text) <> "" Then
'            If CInt(Right(CboTipoFte.Text, 2)) = gPersFteIngresoTipoDependiente Then
'                TxtIngCon.SetFocus
'            Else
'                txtDisponible.SetFocus
'            End If
'        End If

        If Me.TxFecEval.Visible = True Then
            Me.TxFecEval.SetFocus
        End If
        
        If Me.CmbFecha.Visible = True Then
            Me.CmbFecha.SetFocus
        End If

    End If
End Sub

'peac 20071227
Private Sub txtFecEEFF_LostFocus()
'    Dim sCad As String
'    sCad = ValidaFecha(txtFecEEFF.Text)
'    If Len(Trim(sCad)) > 0 Then
'        MsgBox sCad, vbInformation, "Aviso"
'        txtFecEEFF.SetFocus
'    End If
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

Private Sub txtMonto_GotFocus()
fEnfoque txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii, , , True)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtMonto_LostFocus()
    If Trim(txtMonto.Text) = "" Then
        txtMonto.Text = "0.00"
    Else
        txtMonto.Text = Format(txtMonto.Text, "#0.00")
    End If

End Sub

Private Sub txtMonto2_GotFocus()
fEnfoque txtMonto2
End Sub

Private Sub txtMonto2_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto2, KeyAscii, , , True)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtMonto2_LostFocus()
    If Trim(txtMonto2.Text) = "" Then
        txtMonto2.Text = "0.00"
    Else
        txtMonto2.Text = Format(txtMonto2.Text, "#0.00")
    End If
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
    
    'lblEgresosB.Caption = Format(CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) _
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

'lblIngresosB.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
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

'lblIngresosB.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
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

'WIOR 20140313 *******************************************
Private Sub HabilitaEstFinancieros(ByVal HabEstFinan As Boolean)
Me.fraBalGeneral.Enabled = HabEstFinan
Me.fraEstResutado.Enabled = HabEstFinan
Me.fraFlujo.Enabled = HabEstFinan
Me.fraIndicadores.Enabled = HabEstFinan
End Sub

Private Sub CalculoTotal(ByVal pnTipo As Integer)
On Error GoTo ErrorCalculo

Select Case pnTipo
    Case 1:
            txtEFActCorriente.Text = Format(CDbl(txtEFCajaBanco.Text) + CDbl(txtEFCuentaCobrar.Text) + CDbl(txtEFExiste.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(2)
    Case 2:
            txtEFTotalAct.Text = Format(CDbl(txtEFActFijo.Text) + CDbl(txtEFActCorriente.Text), "###," & String(15, "#") & "#0.00")
    Case 3:
            txtEFPasCorriente.Text = Format(CDbl(txtEFProveedor.Text) + CDbl(txtEFDeudaFinanC.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(4)
    Case 4:
            txtEFTotalPas.Text = Format(CDbl(txtEFPasCorriente.Text) + CDbl(txtEFDeudaFinanL.Text), "###," & String(15, "#") & "#0.00")
    Case 5:
            txtEFTotalPat.Text = Format(CDbl(txtEFCapSocial.Text) + CDbl(txtEFResulAcum.Text), "###," & String(15, "#") & "#0.00")
    Case 6:
            txtEFUtBruta.Text = Format(CDbl(txtEFVentas.Text) - CDbl(txtEFCostVentas.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(7)
    Case 7:
            txtEFUtOpe.Text = Format(CDbl(txtEFUtBruta.Text) - CDbl(txtEFGastosOpe.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(8)
    Case 8:
            txtEFUtNeta.Text = Format(CDbl(txtEFUtOpe.Text) + CDbl(txtEFIngresoFinan.Text) - CDbl(txtEFGastoFinan.Text), "###," & String(15, "#") & "#0.00")
    Case 9:
            txtEFFlujoEfec.Text = Format(CDbl(txtEFFlujoOpe.Text) + CDbl(txtEFFujoInv.Text) + CDbl(txtEFFlujoFinan.Text), "###," & String(15, "#") & "#0.00")
End Select

Exit Sub

ErrorCalculo:
MsgBox "Error: Ingrese los datos Correctamente." & Chr(13) & "Detalles de error: " & Err.Description, vbCritical, "Error"

Select Case pnTipo
    Case 1:
            txtEFCajaBanco.Text = "0.00"
            txtEFCuentaCobrar.Text = "0.00"
            txtEFExiste.Text = "0.00"
    Case 2:
            txtEFActFijo.Text = "0.00"
    Case 3:
            txtEFProveedor.Text = "0.00"
            txtEFDeudaFinanC.Text = "0.00"
    Case 4:
            txtEFDeudaFinanL.Text = "0.00"
    Case 5:
            txtEFCapSocial.Text = "0.00"
            txtEFResulAcum.Text = "0.00"
    Case 6:
            txtEFVentas.Text = "0.00"
            txtEFCostVentas.Text = "0.00"
    Case 7:
            txtEFGastosOpe.Text = "0.00"
    Case 8:
            txtEFIngresoFinan.Text = "0.00"
            txtEFGastoFinan.Text = "0.00"
    Case 9:
            txtEFFlujoEfec.Text = "0.00"
            txtEFFlujoOpe.Text = "0.00"
            txtEFFujoInv.Text = "0.00"
            txtEFFlujoFinan.Text = "0.00"
End Select
Call CalculoTotal(pnTipo)
End Sub

Private Sub LimpiarEstFinan()
txtEFCajaBanco.Text = "0.00"
txtEFCuentaCobrar.Text = "0.00"
txtEFExiste.Text = "0.00"
txtEFActFijo.Text = "0.00"
txtEFActCorriente.Text = "0.00"
txtEFTotalAct.Text = "0.00"
txtEFProveedor.Text = "0.00"
txtEFDeudaFinanC.Text = "0.00"
txtEFPasCorriente.Text = "0.00"
txtEFDeudaFinanL.Text = "0.00"
txtEFTotalPas.Text = "0.00"
txtEFResulAcum.Text = "0.00"
txtEFCapSocial.Text = "0.00"
txtEFTotalPat.Text = "0.00"
txtEFCostVentas.Text = "0.00"
txtEFVentas.Text = "0.00"
txtEFUtBruta.Text = "0.00"
txtEFGastosOpe.Text = "0.00"
txtEFUtOpe.Text = "0.00"
txtEFIngresoFinan.Text = "0.00"
txtEFGastoFinan.Text = "0.00"
txtEFUtNeta.Text = "0.00"
txtEFFlujoOpe.Text = "0.00"
txtEFFujoInv.Text = "0.00"
txtEFFlujoFinan.Text = "0.00"
txtEFFlujoEfec.Text = "0.00"
txtEFPosCambios.Text = "0.00"
txtEFDeudaFinan.Text = "0.00"
txtEFIngresoME.Text = "0.00"
txtEFGastosME.Text = "0.00"
End Sub

Private Sub MostarEstadosFinancieros(ByVal psNumeroFtesIngreso As String, ByVal pdFecFteIng As Date)
Dim oDPersona  As COMDpersona.DCOMPersonas
Dim rsDatos  As ADODB.Recordset
Dim i As Integer
Dim sCadena As String
Set oDPersona = New COMDpersona.DCOMPersonas
Set rsDatos = Nothing
sCadena = "###," & String(15, "#") & "#0.00"
Call LimpiarEstFinan
Set rsDatos = oDPersona.RecuperaFuenteIngEstFinan(psNumeroFtesIngreso, pdFecFteIng)
 
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        fbEditarEF = True
        Call HabilitaEstFinancieros(False)
        For i = 1 To rsDatos.RecordCount
            Select Case CInt(rsDatos!nTpoEF)
                Case 1:
                    Select Case CInt(rsDatos!nSubTpoEF)
                        Case 1: txtEFCajaBanco.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 2: txtEFCuentaCobrar.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 3: txtEFExiste.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 4: txtEFActCorriente.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 5: txtEFActFijo.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 6: txtEFTotalAct.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 7: txtEFProveedor.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 8: txtEFDeudaFinanC.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 9: txtEFPasCorriente.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 10: txtEFDeudaFinanL.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 11: txtEFTotalPas.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 12: txtEFCapSocial.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 13: txtEFResulAcum.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 14: txtEFTotalPat.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                    End Select
                Case 2:
                    Select Case CInt(rsDatos!nSubTpoEF)
                        Case 1: txtEFVentas.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 2: txtEFCostVentas.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 3: txtEFUtBruta.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 4: txtEFGastosOpe.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 5: txtEFUtOpe.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 6: txtEFIngresoFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 7: txtEFGastoFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 8: txtEFUtNeta.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                    End Select
                Case 3:
                      Select Case CInt(rsDatos!nSubTpoEF)
                        Case 1: txtEFFlujoOpe.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 2: txtEFFujoInv.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 3: txtEFFlujoFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 4: txtEFFlujoEfec.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        End Select
                Case 4:
                     Select Case CInt(rsDatos!nSubTpoEF)
                        Case 1: txtEFPosCambios.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 2: txtEFDeudaFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 3: txtEFIngresoME.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        Case 4: txtEFGastosME.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                    End Select
            End Select
            rsDatos.MoveNext
        Next i
    Else
        fbEditarEF = False
    End If
End Sub

Private Function HayDatosEstFinan() As Boolean
HayDatosEstFinan = False

If Not (Trim(txtEFCajaBanco.Text) = "0.00" And _
    Trim(txtEFCuentaCobrar.Text) = "0.00" And _
    Trim(txtEFExiste.Text) = "0.00" And _
    Trim(txtEFActCorriente.Text) = "0.00" And _
    Trim(txtEFActFijo.Text) = "0.00" And _
    Trim(txtEFTotalAct.Text) = "0.00" And _
    Trim(txtEFProveedor.Text) = "0.00" And _
    Trim(txtEFDeudaFinanC.Text) = "0.00" And _
    Trim(txtEFPasCorriente.Text) = "0.00" And _
    Trim(txtEFDeudaFinanL.Text) = "0.00" And _
    Trim(txtEFTotalPas.Text) = "0.00" And _
    Trim(txtEFCapSocial.Text) = "0.00" And _
    Trim(txtEFResulAcum.Text) = "0.00" And _
    Trim(txtEFTotalPat.Text) = "0.00") Then

   HayDatosEstFinan = True

End If
   
If Not (Trim(txtEFVentas.Text) = "0.00" And _
   Trim(txtEFCostVentas.Text) = "0.00" And _
   Trim(txtEFUtBruta.Text) = "0.00" And _
   Trim(txtEFGastosOpe.Text) = "0.00" And _
   Trim(txtEFUtOpe.Text) = "0.00" And _
   Trim(txtEFIngresoFinan.Text) = "0.00" And _
   Trim(txtEFGastoFinan.Text) = "0.00" And _
   Trim(txtEFUtNeta.Text) = "0.00") Then
     HayDatosEstFinan = True
End If

If Not (Trim(txtEFFlujoOpe.Text) = "0.00" And _
   Trim(txtEFFujoInv.Text) = "0.00" And _
   Trim(txtEFFlujoFinan.Text) = "0.00" And _
   Trim(txtEFFlujoEfec.Text) = "0.00") Then
   
End If


If Not (Trim(txtEFPosCambios.Text) = "0.00" And _
   Trim(txtEFDeudaFinan.Text) = "0.00" And _
   Trim(txtEFIngresoME.Text) = "0.00" And _
   Trim(txtEFGastosME.Text) = "0.00") Then
     HayDatosEstFinan = True
End If

End Function

Private Function CargaRSEstFinan() As ADODB.Recordset
Dim rsEF As ADODB.Recordset
Dim i As Integer
Dim J As Integer

Set rsEF = New ADODB.Recordset

With rsEF
    'Crear RecordSet
    .Fields.Append "dPersEval", adDate          '1
    .Fields.Append "cNumFuente", adVarChar, 8   '2
    .Fields.Append "nTpoEF", adInteger          '3
    .Fields.Append "nSubTpoEF", adInteger       '4
    .Fields.Append "nMonto", adCurrency         '5
    .Open
    
    'Llenar Recordset
    If nProcesoActual = 1 Then ' Nuevo
        dFecFteIng = TxFecEval.Text
    ElseIf nProcesoActual = 2 Then ' Nuevo
        dFecFteIng = CDate(CmbFecha.Text) 'Editar
    Else
        dFecFteIng = TxFecEval.Text
    End If
    
    If HayDatosEstFinan Then
        For i = 1 To 4
            For J = 1 To 14
                .AddNew
                .Fields("dPersEval") = Format(dFecFteIng, "YYYY/mm/dd")
                .Fields("cNumFuente") = sNumeroFtesIngreso
                .Fields("nTpoEF") = i
                .Fields("nSubTpoEF") = J
                
                Select Case i
                    Case 1:
                        Select Case J
                            Case 1: .Fields("nMonto") = CDbl(txtEFCajaBanco.Text)
                            Case 2: .Fields("nMonto") = CDbl(txtEFCuentaCobrar.Text)
                            Case 3: .Fields("nMonto") = CDbl(txtEFExiste.Text)
                            Case 4: .Fields("nMonto") = CDbl(txtEFActCorriente.Text)
                            Case 5: .Fields("nMonto") = CDbl(txtEFActFijo.Text)
                            Case 6: .Fields("nMonto") = CDbl(txtEFTotalAct.Text)
                            Case 7: .Fields("nMonto") = CDbl(txtEFProveedor.Text)
                            Case 8: .Fields("nMonto") = CDbl(txtEFDeudaFinanC.Text)
                            Case 9: .Fields("nMonto") = CDbl(txtEFPasCorriente.Text)
                            Case 10: .Fields("nMonto") = CDbl(txtEFDeudaFinanL.Text)
                            Case 11: .Fields("nMonto") = CDbl(txtEFTotalPas.Text)
                            Case 12: .Fields("nMonto") = CDbl(txtEFCapSocial.Text)
                            Case 13: .Fields("nMonto") = CDbl(txtEFResulAcum.Text)
                            Case 14: .Fields("nMonto") = CDbl(txtEFTotalPat.Text)
                                    Exit For
                        End Select
                    Case 2:
                        Select Case J
                            Case 1: .Fields("nMonto") = CDbl(txtEFVentas.Text)
                            Case 2: .Fields("nMonto") = CDbl(txtEFCostVentas.Text)
                            Case 3: .Fields("nMonto") = CDbl(txtEFUtBruta.Text)
                            Case 4: .Fields("nMonto") = CDbl(txtEFGastosOpe.Text)
                            Case 5: .Fields("nMonto") = CDbl(txtEFUtOpe.Text)
                            Case 6: .Fields("nMonto") = CDbl(txtEFIngresoFinan.Text)
                            Case 7: .Fields("nMonto") = CDbl(txtEFGastoFinan.Text)
                            Case 8: .Fields("nMonto") = CDbl(txtEFUtNeta.Text)
                                    Exit For
                        End Select
                    Case 3:
                          Select Case J
                            Case 1: .Fields("nMonto") = CDbl(txtEFFlujoOpe.Text)
                            Case 2: .Fields("nMonto") = CDbl(txtEFFujoInv.Text)
                            Case 3: .Fields("nMonto") = CDbl(txtEFFlujoFinan.Text)
                            Case 4: .Fields("nMonto") = CDbl(txtEFFlujoEfec.Text)
                                    Exit For
                            End Select
                    Case 4:
                         Select Case J
                            Case 1: .Fields("nMonto") = CDbl(txtEFPosCambios.Text)
                            Case 2: .Fields("nMonto") = CDbl(txtEFDeudaFinan.Text)
                            Case 3: .Fields("nMonto") = CDbl(txtEFIngresoME.Text)
                            Case 4: .Fields("nMonto") = CDbl(txtEFGastosME.Text)
                                    Exit For
                        End Select
                End Select
                
                
            Next J
        Next i
    End If
    If Not .EOF Then .MoveFirst
End With

Set CargaRSEstFinan = rsEF
End Function
'WIOR FIN ************************************************
