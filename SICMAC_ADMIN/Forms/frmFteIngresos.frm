VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFteIngresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fuentes de Ingreso"
   ClientHeight    =   7560
   ClientLeft      =   2025
   ClientTop       =   705
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   4800
      TabIndex        =   26
      Top             =   7035
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
      Height          =   1665
      Left            =   135
      TabIndex        =   69
      Top             =   1200
      Width           =   7440
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
         Left            =   1335
         TabIndex        =   7
         Top             =   1245
         Width           =   1770
      End
      Begin VB.TextBox TxtRazSocDirecc 
         Height          =   315
         Left            =   1335
         TabIndex        =   5
         Top             =   900
         Width           =   5400
      End
      Begin VB.TextBox TxtRazSocDescrip 
         Height          =   315
         Left            =   1335
         TabIndex        =   4
         Top             =   570
         Width           =   5895
      End
      Begin Sicmact.TxtBuscar TxtBRazonSoc 
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
         TabIndex        =   71
         Top             =   615
         Width           =   1170
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
      Caption         =   "&Salir"
      Height          =   420
      Left            =   6225
      TabIndex        =   27
      Top             =   7035
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   135
      TabIndex        =   29
      Top             =   -15
      Width           =   7440
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         Left            =   1710
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
         Top             =   450
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
      Height          =   4110
      Left            =   105
      TabIndex        =   28
      Top             =   2880
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   7250
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "In&gresos y Egresos"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label26"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPFecIni"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtCargo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Balance"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.TextBox TxtCargo 
         Height          =   285
         Left            =   2235
         TabIndex        =   13
         Top             =   2220
         Width           =   4560
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         Top             =   1785
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55508993
         CurrentDate     =   37014
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ingresos y Egresos de la Empresa"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   -74820
         TabIndex        =   59
         Top             =   2610
         Width           =   7125
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
            Left            =   1995
            MaxLength       =   13
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   255
            Width           =   1080
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
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   540
            Width           =   1080
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
            Top             =   240
            Width           =   1080
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
            Top             =   525
            Width           =   1080
         End
         Begin VB.Line Line2 
            X1              =   105
            X2              =   6840
            Y1              =   900
            Y2              =   900
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
            Left            =   3870
            TabIndex        =   68
            Top             =   1035
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
            Left            =   5460
            TabIndex        =   67
            Top             =   975
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ventas :"
            Height          =   195
            Left            =   195
            TabIndex        =   63
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Rec. de Ctas x Cobrar :"
            Height          =   195
            Left            =   195
            TabIndex        =   62
            Top             =   570
            Width           =   1650
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Compra de Mercaderías :"
            Height          =   195
            Left            =   3525
            TabIndex        =   61
            Top             =   255
            Width           =   1800
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Otros Egresos :"
            Height          =   195
            Left            =   3525
            TabIndex        =   60
            Top             =   555
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Balance de Situacion"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   -74805
         TabIndex        =   41
         Top             =   375
         Width           =   7140
         Begin VB.TextBox txtDisponible 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1725
            MaxLength       =   13
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   900
            Width           =   990
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
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1185
            Width           =   990
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
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   1470
            Width           =   990
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
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   1800
            Width           =   1035
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
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   1485
            Width           =   1035
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
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   1200
            Width           =   1035
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
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   900
            Width           =   1035
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
            TabIndex        =   58
            Top             =   345
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
            TabIndex        =   57
            Top             =   615
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
            TabIndex        =   56
            Top             =   1860
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
            TabIndex        =   55
            Top             =   315
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
            TabIndex        =   54
            Top             =   615
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
            TabIndex        =   53
            Top             =   1905
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Disponible :"
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   930
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cuentas x Cobrar:"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   1215
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Inventario :"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   1500
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
            TabIndex        =   49
            Top             =   600
            Width           =   1035
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
            TabIndex        =   48
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Prestamos CMACT"
            Height          =   195
            Left            =   3750
            TabIndex        =   47
            Top             =   1485
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Otros Préstamos :"
            Height          =   195
            Left            =   3750
            TabIndex        =   46
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Proveedores :"
            Height          =   195
            Left            =   3750
            TabIndex        =   45
            Top             =   945
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
            Top             =   555
            Width           =   1020
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
            Top             =   255
            Width           =   1020
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
            Top             =   1830
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1320
         Left            =   908
         TabIndex        =   34
         Top             =   345
         Width           =   5835
         Begin VB.TextBox txtOtroIng 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1425
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   510
            Width           =   1155
         End
         Begin VB.TextBox txtIngresos 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1425
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   195
            Width           =   1155
         End
         Begin VB.TextBox txtIngFam 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4455
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   165
            Width           =   1155
         End
         Begin VB.TextBox txtEgreFam 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4455
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   480
            Width           =   1155
         End
         Begin VB.Line Line1 
            X1              =   180
            X2              =   5625
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Label lblSaldo 
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
            Left            =   4440
            TabIndex        =   40
            Top             =   930
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
            Left            =   2910
            TabIndex        =   39
            Top             =   990
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Otros Ingresos:"
            Height          =   195
            Left            =   165
            TabIndex        =   38
            Top             =   585
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
            Caption         =   "Ingreso Familiar :"
            Height          =   195
            Left            =   2925
            TabIndex        =   36
            Top             =   195
            Width           =   1200
         End
         Begin VB.Label lblEgreso 
            AutoSize        =   -1  'True
            Caption         =   "Egreso Familiar :"
            Height          =   195
            Left            =   2910
            TabIndex        =   35
            Top             =   540
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1110
         Left            =   660
         TabIndex        =   32
         Top             =   2550
         Width           =   6390
         Begin RichTextLib.RichTextBox Txtcomentarios 
            Height          =   825
            Left            =   90
            TabIndex        =   14
            Top             =   225
            Width           =   6210
            _ExtentX        =   10954
            _ExtentY        =   1455
            _Version        =   393217
            MaxLength       =   60
            TextRTF         =   $"frmFteIngresos.frx":0000
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
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Cargo : "
         Height          =   195
         Left            =   960
         TabIndex        =   66
         Top             =   2250
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio :"
         Height          =   195
         Left            =   960
         TabIndex        =   65
         Top             =   1830
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmFteIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPersona As DPersona
Dim nIndice As Integer
Dim nProcesoEjecutado As Integer '1 Nueva fte de Ingreso; 2 Editar fte de Ingreso ; 3 Consulta de Fte
Dim vsUbiGeo As String
Dim bEstadoCargando As Boolean
Private Function ValidaDatosFuentesIngreso() As Boolean
Dim CadTemp As String

    ValidaDatosFuentesIngreso = True
    
    If CboTipoFte.ListIndex = -1 Then
        MsgBox "No se ha Seleccionado el Tipo de Fuente", vbInformation, "Aviso"
        CboTipoFte.SetFocus
        ValidaDatosFuentesIngreso = False
        Exit Function
    End If
    
    If cboMoneda.ListIndex = -1 Then
        MsgBox "No se ha Seleccionado la Moneda", vbInformation, "Aviso"
        cboMoneda.SetFocus
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
    
End Function
Private Sub HabilitaCabecera(ByVal pnHabilitar As Boolean)
    CboTipoFte.Enabled = pnHabilitar
    cboMoneda.Enabled = pnHabilitar
    TxtBRazonSoc.Enabled = pnHabilitar
End Sub
Private Sub HabilitaIngresosEgresos(ByVal pnHabilitar As Boolean)
    txtIngresos.Enabled = pnHabilitar
    txtOtroIng.Enabled = pnHabilitar
    txtIngFam.Enabled = pnHabilitar
    txtEgreFam.Enabled = pnHabilitar
    DTPFecIni.Enabled = pnHabilitar
    TxtCargo.Enabled = pnHabilitar
    Txtcomentarios.Enabled = pnHabilitar
End Sub
Private Sub HabilitaBalance(ByVal HabBalance As Boolean)
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
    Frame4.Enabled = HabBalance
    Frame5.Enabled = HabBalance
End Sub

Private Sub CargaControles()
    Call CargaComboConstante(gPersFteIngresoTipo, CboTipoFte)
    Call CargaComboConstante(gMoneda, cboMoneda)
End Sub
Private Sub CargaDatosFteIngreso(ByVal pnIndice As Integer, ByRef poPersona As DPersona)
    
    'LblCliente.Caption = poPersona.NombreCompleto
    'TxtBRazonSoc.Text = poPersona.ObtenerFteIngFuente(pnIndice)
    'LblRazonSoc.Caption = poPersona.ObtenerFteIngRazonSocial(pnIndice)
    'CboTipoFte.ListIndex = IndiceListaCombo(CboTipoFte, poPersona.ObtenerFteIngTipo(pnIndice))
    'TxtRazSocDescrip.Text = poPersona.ObtenerFteIngRazSocDescrip(pnIndice)
    'TxtRazSocDirecc.Text = poPersona.ObtenerFteIngRazSocDirecc(pnIndice)
    'TxtRazSocTelef.Text = poPersona.ObtenerFteIngRazSocTelef(pnIndice)
    'vsUbiGeo = poPersona.ObtenerFteIngRazSocUbiGeo(pnIndice)
    'If CInt(poPersona.ObtenerFteIngTipo(pnIndice)) = gPersFteIngresoTipoDependiente Then
    '    Call HabilitaBalance(False)
    'Else
    '    Call HabilitaBalance(True)
    'End If
   '
    'cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, poPersona.ObtenerFteIngMoneda(pnIndice))
    
    'Carga Ingresos y Egresos
    'txtIngFam.Text = Format(poPersona.ObtenerFteIngIngresoFam(pnIndice), "#0.00")
    'txtOtroIng.Text = Format(poPersona.ObtenerFteIngIngresoOtros(pnIndice), "#0.00")
    'txtIngresos.Text = Format(poPersona.ObtenerFteIngIngresos(pnIndice), "#0.00")
    'txtEgreFam.Text = Format(poPersona.ObtenerFteIngGastoFam(pnIndice), "#0.00")
    'lblSaldo.Caption = CDbl(txtIngFam.Text) + CDbl(txtOtroIng.Text) + CDbl(txtIngresos.Text) - CDbl(txtEgreFam.Text)
    
    'DTPFecIni.value = CDate(Format(poPersona.ObtenerFteIngInicioFuente(pnIndice), "dd/mm/yyyy"))
    'TxtCargo.Text = poPersona.ObtenerFteIngCargo(pnIndice)
    'Txtcomentarios.Text = poPersona.ObtenerFteIngComentarios(pnIndice)
    
    'Carga el Balance
    'txtDisponible.Text = Format(poPersona.ObtenerFteIngActivoDisp(pnIndice), "#0.00")
    'txtcuentas.Text = Format(poPersona.ObtenerFteIngCtasxCob(pnIndice), "#0.00")
    'txtInventario.Text = Format(poPersona.ObtenerFteIngInventario(pnIndice), "#0.00")
    'txtactivofijo.Text = Format(poPersona.ObtenerFteIngActivoFijo(pnIndice), "#0.00")
    
    'txtProveedores.Text = Format(poPersona.ObtenerFteIngProveedores(pnIndice), "#0.00")
    'txtOtrosPrest.Text = Format(poPersona.ObtenerFteIngOtrosCreditos(pnIndice), "#0.00")
    'txtPrestCmact.Text = Format(poPersona.ObtenerFteIngCreditosCmact(pnIndice), "#0.00")
    
    'lblPatrimonio.Caption = Format(poPersona.ObtenerFteIngPasivoPatrim(pnIndice), "#0.00")
    'txtVentas.Text = Format(poPersona.ObtenerFteIngVentas(pnIndice), "#0.00")
    'txtrecuperacion.Text = Format(poPersona.ObtenerFteIngRecupCtasxCobrar(pnIndice), "#0.00")
    'txtcompras.Text = Format(poPersona.ObtenerFteIngComprasMercad(pnIndice), "#0.00")
    'txtOtrosEgresos.Text = Format(poPersona.ObtenerFteIngOtrosEgresos(pnIndice), "#0.00")
    
End Sub
Private Sub LimpiaFormulario()
    
    LblCliente.Caption = oPersona.NombreCompleto
    TxtBRazonSoc.Text = ""
    CboTipoFte.ListIndex = -1
    cboMoneda.ListIndex = -1
    txtIngresos.Text = "0.00"
    txtIngFam.Text = "0.00"
    txtOtroIng.Text = "0.00"
    txtEgreFam.Text = "0.00"
    DTPFecIni.value = gdFecSis
    TxtCargo.Text = ""
    Txtcomentarios.Text = ""
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

Public Sub Editar(ByVal pnIndice As Integer, ByRef poPersona As DPersona)
    Set oPersona = poPersona
    nIndice = pnIndice
    nProcesoEjecutado = 2
    bEstadoCargando = True
    Call CargaControles
    Call CargaDatosFteIngreso(pnIndice, poPersona)
    cmdAceptar.Visible = True
    CmdSalirCancelar.Caption = "&Cancelar"
    bEstadoCargando = False
    frmFteIngresos.Show 1
End Sub

Public Sub NuevaFteIngreso(ByRef poPersona As DPersona)
    bEstadoCargando = True
    Set oPersona = poPersona
    Call CargaControles
    Call LimpiaFormulario
    nProcesoEjecutado = 1
    cmdAceptar.Visible = True
    CmdSalirCancelar.Caption = "&Cancelar"
    bEstadoCargando = False
    frmFteIngresos.Show 1
End Sub

Public Sub ConsultarFuenteIngreso(ByVal pnIndice As Integer, ByRef poPersona As DPersona)
    Set oPersona = poPersona
    nIndice = pnIndice
    nProcesoEjecutado = 3
    bEstadoCargando = True
    Call CargaControles
    Call CargaDatosFteIngreso(pnIndice, poPersona)
    CmdSalirCancelar.Caption = "&Salir"
    Call HabilitaCabecera(False)
    Call HabilitaBalance(False)
    Call HabilitaIngresosEgresos(False)
    bEstadoCargando = False
    frmFteIngresos.Show 1
End Sub

Private Sub CboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtBRazonSoc.Enabled Then
            TxtBRazonSoc.SetFocus
        Else
            txtIngresos.SetFocus
        End If
    End If
End Sub

Private Sub CboTipoFte_Click()
    If Trim(Right(CboTipoFte.Text, 15)) = gPersFteIngresoTipoDependiente Then
        Call HabilitaBalance(False)
'        TxtBRazonSoc.Enabled = True
    Else
        Call HabilitaBalance(True)
'        TxtBRazonSoc.Text = ""
'        TxtBRazonSoc.Enabled = False
'        LblRazonSoc.Caption = ""
    End If
End Sub

Private Sub CboTipoFte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
'Dim oPersonaNeg As nPersona
    
    'If Not ValidaDatosFuentesIngreso Then
    '    Exit Sub
    'End If
    
    'If nProcesoEjecutado = 1 Then
    '    Call oPersona.AdicionaFteIngreso
    '    nIndice = oPersona.NumeroFtesIngreso - 1
    'Else
    '    If oPersona.ObtenerFteIngTipoAct(nIndice) <> PersFilaNueva Then
    '        Call oPersona.ActualizarFteIngTipoAct(PersFilaModificada, nIndice)
    '    End If
    'End If
    'Call oPersona.ActualizarFteIngTipoFuente(Trim(Right(CboTipoFte.Text, 20)), nIndice)
    'Call oPersona.ActualizarFteIngMoneda(Trim(Right(cboMoneda.Text, 20)), nIndice)
    'Call oPersona.ActualizarFteIngFuenteIngreso(Trim(TxtBRazonSoc.Text), LblRazonSoc.Caption, nIndice)
    'Call oPersona.ActualizarFteIngIngresos(CDbl(txtIngresos.Text), nIndice)
    'Call oPersona.ActualizarFteIngIngFam(CDbl(txtIngFam.Text), nIndice)
    'Call oPersona.ActualizarFteIngIngOtros(CDbl(txtOtroIng.Text), nIndice)
    'Call oPersona.ActualizarFteIngGastosFam(CDbl(txtEgreFam.Text), nIndice)
    'Call oPersona.ActualizarFteIngFechaInicioFuente(CDate(Format(DTPFecIni.value, "dd/mm/yyyy")), nIndice)
    'Call oPersona.ActualizarFteIngCargo(TxtCargo.Text, nIndice)
    'Call oPersona.ActualizarFteIngComentarios(Txtcomentarios.Text, nIndice)
    'Call oPersona.ActualizarFteIngActivDisp(CDbl(txtDisponible.Text), nIndice)
    'Call oPersona.ActualizarFteIngCtasxCob(CDbl(txtcuentas.Text), nIndice)
    'Call oPersona.ActualizarFteIngInventarios(CDbl(txtInventario.Text), nIndice)
    'Call oPersona.ActualizarFteIngActFijo(CDbl(txtactivofijo.Text), nIndice)
    'Call oPersona.ActualizarFteIngProveedores(CDbl(txtProveedores.Text), nIndice)
    'Call oPersona.ActualizarFteIngCreditosOtros(CDbl(txtOtrosPrest.Text), nIndice)
    'Call oPersona.ActualizarFteIngCreditosCmact(CDbl(txtPrestCmact.Text), nIndice)
    'Call oPersona.ActualizarFteIngVentas(CDbl(txtVentas.Text), nIndice)
    'Call oPersona.ActualizarFteIngRecupCtasxCob(CDbl(txtrecuperacion.Text), nIndice)
    'Call oPersona.ActualizarFteIngCompraMercad(CDbl(txtcompras.Text), nIndice)
    'Call oPersona.ActualizarFteIngOtrosEgresos(CDbl(txtOtrosEgresos.Text), nIndice)
    'Call oPersona.ActualizarFteIngFecEval(gdFecSis, nIndice)
    'Call oPersona.ActualizarFteIngPasivoPatrimonio(CDbl(lblPatrimonio.Caption), nIndice)
    'Call oPersona.ActualizarFteIngRazSocDirecc(Trim(TxtRazSocDirecc.Text), nIndice)
    'Call oPersona.ActualizarFteIngRazSocDescrip(Trim(TxtRazSocDescrip.Text), nIndice)
    'Call oPersona.ActualizarFteIngRazSocTelef(Trim(TxtRazSocTelef.Text), nIndice)
    'Call oPersona.ActualizarFteIngRazSocUbiGeo(Trim(vsUbiGeo), nIndice)
    
    'If nProcesoEjecutado = 1 Then
    '    Set oPersonaNeg = New nPersona
    '    Call oPersonaNeg.ChequeoFuenteIngreso(oPersona, nIndice)
    '    Set oPersonaNeg = Nothing
    'End If
    'Unload Me
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
    bEstadoCargando = False
End Sub

Private Sub txtactivofijo_Change()
   lblActivo.Caption = CDbl(IIf(Trim(lblActCirc.Caption) = "", "0", lblActCirc.Caption)) + CDbl(IIf(Trim(txtactivofijo.Text) = "", "0", txtactivofijo.Text))
   lblPasPatrim.Caption = Format(CDbl(lblPasivo.Caption) + CDbl(lblActivo.Caption), "#0.00")
   lblPatrimonio.Caption = Format(CDbl(lblActivo.Caption) - CDbl(lblPasivo.Caption), "#0.00")
    
End Sub

Private Sub txtactivofijo_GotFocus()
    fEnfoque txtactivofijo
End Sub

Private Sub txtactivofijo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtactivofijo, KeyAscii)
        If KeyAscii = 13 Then
            txtProveedores.SetFocus
        End If
End Sub

Private Sub txtactivofijo_LostFocus()
    txtactivofijo.Text = Format(IIf(Trim(txtactivofijo.Text) = "", 0, txtactivofijo.Text), "#0.00")
End Sub

Private Sub TxtBRazonSoc_EmiteDatos()
    LblRazonSoc.Caption = Trim(TxtBRazonSoc.psDescripcion)
    TxtRazSocDescrip.SetFocus
End Sub


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
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)), "#0.00")
End Sub

Private Sub txtcompras_GotFocus()
    fEnfoque txtcompras
End Sub

Private Sub txtcompras_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtcompras, KeyAscii)
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
    KeyAscii = NumerosDecimales(txtcuentas, KeyAscii)
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
    KeyAscii = NumerosDecimales(txtDisponible, KeyAscii)
    If KeyAscii = 13 Then
        txtcuentas.SetFocus
    End If
End Sub

Private Sub txtDisponible_LostFocus()
    txtDisponible.Text = Format(IIf(Trim(txtDisponible.Text) = "", 0, txtDisponible.Text), "#0.00")
End Sub

Private Sub txtEgreFam_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngresos.Text) = "", "0", txtIngresos.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
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
    If Len(Trim(txtEgreFam.Text)) = 0 Then
        txtEgreFam.Text = Format(0#, "#0.00")
    End If
End Sub

Private Sub txtIngFam_Change()
lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngresos.Text) = "", "0", txtIngresos.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
End Sub

Private Sub txtIngFam_GotFocus()
    fEnfoque txtIngFam
End Sub

Private Sub txtIngFam_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtIngFam, KeyAscii)
    If KeyAscii = 13 Then
        txtOtroIng.SetFocus
    End If
End Sub

Private Sub txtIngFam_LostFocus()
    If Len(Trim(txtIngFam.Text)) = 0 Then
        txtIngFam.Text = Format(0#, "#0.00")
    End If
End Sub

Private Sub txtIngresos_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngresos.Text) = "", "0", txtIngresos.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
            
End Sub

Private Sub txtIngresos_GotFocus()
    fEnfoque txtIngresos
End Sub

Private Sub txtIngresos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtIngresos, KeyAscii)
    If KeyAscii = 13 Then
        txtIngFam.SetFocus
    End If
End Sub

Private Sub txtIngresos_LostFocus()
    If Len(Trim(txtIngresos.Text)) = 0 Then
        txtIngresos.Text = Format(0#, "#0.00")
    End If
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
    KeyAscii = NumerosDecimales(txtInventario, KeyAscii)
    If KeyAscii = 13 Then
        txtactivofijo.SetFocus
    End If
End Sub

Private Sub txtInventario_LostFocus()
    txtInventario.Text = Format(IIf(Trim(txtInventario.Text) = "", 0, txtInventario.Text), "#0.00")
End Sub

Private Sub txtOtroIng_Change()
    lblSaldo.Caption = CDbl(Format(IIf(Trim(txtIngresos.Text) = "", "0", txtIngresos.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtIngFam.Text) = "", "0", txtIngFam.Text), "#0.00")) _
            + CDbl(Format(IIf(Trim(txtOtroIng.Text) = "", "0", txtOtroIng.Text), "#0.00")) _
            - CDbl(Format(IIf(Trim(txtEgreFam.Text) = "", "0", txtEgreFam.Text), "#0.00"))
            
End Sub

Private Sub txtOtroIng_GotFocus()
    fEnfoque txtOtroIng
End Sub

Private Sub txtOtroIng_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtOtroIng, KeyAscii)
    If KeyAscii = 13 Then
        txtEgreFam.SetFocus
    End If
End Sub

Private Sub txtOtroIng_LostFocus()
    If Len(Trim(txtOtroIng.Text)) = 0 Then
        txtOtroIng.Text = Format(0#, "#0.00")
    End If
End Sub

Private Sub txtOtrosEgresos_Change()
    LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)), "#0.00")
End Sub

Private Sub txtOtrosEgresos_GotFocus()
    fEnfoque txtOtrosEgresos
End Sub

Private Sub txtOtrosEgresos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtOtrosEgresos, KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
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
    KeyAscii = NumerosDecimales(txtOtrosPrest, KeyAscii)
    If KeyAscii = 13 Then
        txtPrestCmact.SetFocus
    End If
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
    KeyAscii = NumerosDecimales(txtPrestCmact, KeyAscii)
    If KeyAscii = 13 Then
        txtVentas.SetFocus
    End If
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
    KeyAscii = NumerosDecimales(txtProveedores, KeyAscii)
    If KeyAscii = 13 Then
        txtOtrosPrest.SetFocus
    End If
End Sub


Private Sub TxtRazSocDescrip_GotFocus()
    fEnfoque TxtRazSocDescrip
End Sub

Private Sub TxtRazSocDescrip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRazSocDirecc.SetFocus
    End If
End Sub

Private Sub TxtRazSocDirecc_GotFocus()
    fEnfoque TxtRazSocDirecc
End Sub

Private Sub TxtRazSocDirecc_KeyPress(KeyAscii As Integer)
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
        txtIngresos.SetFocus
    End If
End Sub

Private Sub txtrecuperacion_Change()
    LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)), "#0.00")

End Sub

Private Sub txtrecuperacion_GotFocus()
    fEnfoque txtrecuperacion
End Sub

Private Sub txtrecuperacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtrecuperacion, KeyAscii)
    If KeyAscii = 13 Then
        txtcompras.SetFocus
    End If
End Sub

Private Sub txtrecuperacion_LostFocus()
    txtrecuperacion.Text = Format(IIf(Trim(txtrecuperacion.Text) = "", "0.00", txtrecuperacion.Text), "#0.00")
    
End Sub

Private Sub txtVentas_Change()
    LblSaldoIngEgr.Caption = Format(CDbl(IIf(Trim(txtVentas.Text) = "", "0", txtVentas.Text)) _
            + CDbl(IIf(Trim(txtrecuperacion.Text) = "", "0", txtrecuperacion.Text)) _
            - CDbl(IIf(Trim(txtcompras.Text) = "", "0", txtcompras.Text)) - CDbl(IIf(Trim(txtOtrosEgresos.Text) = "", "0", txtOtrosEgresos.Text)), "#0.00")
End Sub

Private Sub txtVentas_GotFocus()
    fEnfoque txtVentas
End Sub

Private Sub txtVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentas, KeyAscii)
    If KeyAscii = 13 Then
        txtrecuperacion.SetFocus
    End If
End Sub

Private Sub txtVentas_LostFocus()
    txtVentas.Text = Format(IIf(Trim(txtVentas.Text) = "", "0.00", txtVentas.Text), "#0.00")
End Sub
