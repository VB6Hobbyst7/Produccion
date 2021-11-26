VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmVerRFA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Creditos RFA"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   30
      TabIndex        =   9
      Top             =   5850
      Width           =   10725
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9510
         TabIndex        =   11
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   1125
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   30
      TabIndex        =   8
      Top             =   1440
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   7805
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DIF- CMAC"
      TabPicture(0)   =   "FrmVerRFA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FlexDIF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "RFC - CMAC"
      TabPicture(1)   =   "FrmVerRFA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlexRFC"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "RFA - COFIDE"
      TabPicture(2)   =   "FrmVerRFA.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FlexRFA"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin SICMACT.FlexEdit FlexDIF 
         Height          =   3525
         Left            =   150
         TabIndex        =   45
         Top             =   360
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   6218
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cuota-Fec.Venc-Fec.Pag-Capital-Interes-Mora-Cofide-Gastos-Total-Credito"
         EncabezadosAnchos=   "500-1200-1200-1200-1200-1200-1200-1200-1200-2600"
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
         EncabezadosAlineacion=   "L-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Cuota"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame5 
         Height          =   525
         Left            =   -74940
         TabIndex        =   14
         Top             =   3870
         Width           =   10575
         Begin VB.Label LblGastosRFA 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   7710
            TabIndex        =   44
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Gastos"
            Height          =   195
            Left            =   6690
            TabIndex        =   43
            Top             =   240
            Width           =   945
         End
         Begin VB.Label LblCofideRFA 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   9810
            TabIndex        =   42
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Cofide"
            Height          =   195
            Left            =   8850
            TabIndex        =   41
            Top             =   210
            Width           =   900
         End
         Begin VB.Label LblInteresRFA 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3360
            TabIndex        =   40
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes"
            Height          =   195
            Left            =   2370
            TabIndex        =   39
            Top             =   240
            Width           =   930
         End
         Begin VB.Label LblMoraRFA 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   5340
            TabIndex        =   38
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Mora"
            Height          =   195
            Left            =   4530
            TabIndex        =   37
            Top             =   240
            Width           =   810
         End
         Begin VB.Label LblCapitalRFA 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   990
            TabIndex        =   36
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   0
            TabIndex        =   35
            Top             =   210
            Width           =   930
         End
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   90
         TabIndex        =   13
         Top             =   3810
         Width           =   10605
         Begin VB.Label LblGastosDif 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   5310
            TabIndex        =   24
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Gastos"
            Height          =   195
            Left            =   4320
            TabIndex        =   23
            Top             =   150
            Width           =   945
         End
         Begin VB.Label LblCofideDif 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   9480
            TabIndex        =   22
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Cofide"
            Height          =   195
            Left            =   8490
            TabIndex        =   21
            Top             =   210
            Width           =   900
         End
         Begin VB.Label LblInteresDIf 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3240
            TabIndex        =   20
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes"
            Height          =   195
            Left            =   2310
            TabIndex        =   19
            Top             =   180
            Width           =   930
         End
         Begin VB.Label LblMoraDif 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   7200
            TabIndex        =   18
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Mora"
            Height          =   195
            Left            =   6270
            TabIndex        =   17
            Top             =   180
            Width           =   810
         End
         Begin VB.Label LblSaldoCapitalDIF 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   1050
            TabIndex        =   16
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   60
            TabIndex        =   15
            Top             =   150
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Height          =   525
         Left            =   -74940
         TabIndex        =   12
         Top             =   3840
         Width           =   10635
         Begin VB.Label LblGastosRFC 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   7620
            TabIndex        =   34
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Gastos"
            Height          =   195
            Left            =   6630
            TabIndex        =   33
            Top             =   210
            Width           =   945
         End
         Begin VB.Label LblCofideRFC 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   9780
            TabIndex        =   32
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Cofide"
            Height          =   195
            Left            =   8760
            TabIndex        =   31
            Top             =   240
            Width           =   900
         End
         Begin VB.Label LblInteresRFC 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   3270
            TabIndex        =   30
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes"
            Height          =   195
            Left            =   2280
            TabIndex        =   29
            Top             =   240
            Width           =   930
         End
         Begin VB.Label LblMoraRFC 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   5340
            TabIndex        =   28
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Mora"
            Height          =   195
            Left            =   4410
            TabIndex        =   27
            Top             =   210
            Width           =   810
         End
         Begin VB.Label LblCapitalRFC 
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   990
            TabIndex        =   26
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   210
            Width           =   930
         End
      End
      Begin SICMACT.FlexEdit FlexRFC 
         Height          =   3525
         Left            =   -74940
         TabIndex        =   46
         Top             =   390
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   6218
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cuota-Fec.Venc-Fec.Pag-Capital-Interes-Mora-Cofide-Gastos-Total-Credito"
         EncabezadosAnchos=   "500-1200-1200-1200-1200-1200-1200-1200-1200-2600"
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
         EncabezadosAlineacion=   "L-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Cuota"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FlexRFA 
         Height          =   3525
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6218
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Cuota-Fec.Venc-Fec.Pag-Capital-Interes-Mora-Cofide-Gastos-Total-Credito"
         EncabezadosAnchos=   "500-1200-1200-1200-1200-1200-1200-1200-1200-2600"
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
         EncabezadosAlineacion=   "L-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Cuota"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10755
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8610
         TabIndex        =   7
         Top             =   180
         Width           =   1305
      End
      Begin VB.TextBox txtDireccion 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1560
         TabIndex        =   5
         Top             =   1020
         Width           =   3885
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   3885
      End
      Begin SICMACT.TxtBuscar TxtBuscar1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   180
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label LblDeuda 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   8790
         TabIndex        =   49
         Top             =   780
         Width           =   1845
      End
      Begin VB.Label Label6 
         Caption         =   "DEUDA TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7890
         TabIndex        =   48
         Top             =   870
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   5550
         Picture         =   "FrmVerRFA.frx":0054
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   1
         Top             =   210
         Width           =   660
      End
   End
End
Attribute VB_Name = "FrmVerRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdBuscar_Click()
    If TxtBuscar1.Text <> "" Then
        CargarDatos (TxtBuscar1.Text)
    Else
        MsgBox "Debe ingresar un cliente", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdCancelar_Click()
    txtNombre = ""
    txtDireccion = ""
    FlexRFA.Clear
    FlexRFC.Clear
    FlexDIF.Clear
    FlexRFA.FormaCabecera
    FlexRFC.FormaCabecera
    FlexDIF.FormaCabecera
    LblDeuda = ""
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 1
End Sub

Private Sub TxtBuscar1_EmiteDatos()
    Dim sCodigo As String
    Dim odRFa As COMDCredito.DCOMRFA
    Dim rs As ADODB.Recordset
    
    sCodigo = TxtBuscar1.Text
    
    Set odRFa = New COMDCredito.DCOMRFA
    Set rs = odRFa.BuscarPersona(sCodigo)
    Set odRFa = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        txtNombre.Text = rs!cPersNombre
        txtDireccion.Text = IIf(IsNull(rs!cPersDireccDomicilio), "", rs!cPersDireccDomicilio)
    End If
    Set rs = Nothing
End Sub

Sub CargarDatos(ByVal psPersCod As String)
    Dim rs As ADODB.Recordset
    Dim objDRFA As COMDCredito.DCOMRFA
    Dim nFila As Integer
    
    Set objDRFA = New COMDCredito.DCOMRFA
    Set rs = objDRFA.ListaCalendarioByPersona(psPersCod)
    Set objDRFA = Nothing
    
    Do Until rs.EOF
        If rs!Orden = "2" Then
            'rfc
          With FlexRFC
            nFila = CInt(rs!nCuota)
            .AdicionaFila
            .TextMatrix(nFila, 0) = rs!nCuota
            .TextMatrix(nFila, 1) = Format(rs!dvenc, "dd/mm/yyyy")
            .TextMatrix(nFila, 2) = Format(IIf(IsNull(rs!dPago) Or DatePart("yyyy", rs!dPago) = "1900", "", rs!dPago), "dd/mm/yyyy")
            .TextMatrix(nFila, 3) = Format(IIf(IsNull(rs!Capital), 0, rs!Capital), "#0.00")
            .TextMatrix(nFila, 4) = Format(IIf(IsNull(rs!Interes), 0, rs!Interes), "#0.00")
            .TextMatrix(nFila, 5) = Format(IIf(IsNull(rs!Mora), 0, rs!Mora), "#0.00")
            .TextMatrix(nFila, 6) = Format(IIf(IsNull(rs!Cofide), 0, rs!Cofide), "#0.00")
            .TextMatrix(nFila, 7) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos), "#0.00")
            .TextMatrix(nFila, 8) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos) + IIf(IsNull(rs!Capital), 0, rs!Capital) + IIf(IsNull(rs!Interes), 0, rs!Interes) + IIf(IsNull(rs!Mora), 0, rs!Mora) + IIf(IsNull(rs!Cofide), 0, rs!Cofide), "#0.00")
            .TextMatrix(nFila, 9) = Format(rs!cCtaCod)
          End With
            
        ElseIf rs!Orden = "1" Then
            'dif
            nFila = CInt(rs!nCuota)
            FlexDIF.AdicionaFila
            FlexDIF.TextMatrix(nFila, 0) = rs!nCuota
            FlexDIF.TextMatrix(nFila, 1) = Format(rs!dvenc, "dd/mm/yyyy")
            FlexDIF.TextMatrix(nFila, 2) = Format(IIf(IsNull(rs!dPago) Or DatePart("yyyy", rs!dPago) = "1900", "", rs!dPago), "dd/mm/yyyy")
            FlexDIF.TextMatrix(nFila, 3) = Format(IIf(IsNull(rs!Capital), 0, rs!Capital), "#0.00")
            FlexDIF.TextMatrix(nFila, 4) = Format(IIf(IsNull(rs!Interes), 0, rs!Interes), "#0.00")
            FlexDIF.TextMatrix(nFila, 5) = Format(IIf(IsNull(rs!Mora), 0, rs!Mora), "#0.00")
            FlexDIF.TextMatrix(nFila, 6) = Format(IIf(IsNull(rs!Cofide), 0, rs!Cofide), "#0.00")
            FlexDIF.TextMatrix(nFila, 7) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos), "#0.00")
            FlexDIF.TextMatrix(nFila, 8) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos) + IIf(IsNull(rs!Capital), 0, rs!Capital) + IIf(IsNull(rs!Interes), 0, rs!Interes) + IIf(IsNull(rs!Mora), 0, rs!Mora) + IIf(IsNull(rs!Cofide), 0, rs!Cofide), "#0.00")
            FlexDIF.TextMatrix(nFila, 9) = Format(rs!cCtaCod)
        Else
            'rfa
            With FlexRFA
                nFila = CInt(rs!nCuota)
            .AdicionaFila
            .TextMatrix(nFila, 0) = rs!nCuota
            .TextMatrix(nFila, 1) = Format(rs!dvenc, "dd/mm/yyyy")
            .TextMatrix(nFila, 2) = Format(IIf(IsNull(rs!dPago) Or DatePart("yyyy", rs!dPago) = "1900", "", rs!dPago), "dd/mm/yyyy")
            .TextMatrix(nFila, 3) = Format(IIf(IsNull(rs!Capital), 0, rs!Capital), "#0.00")
            .TextMatrix(nFila, 4) = Format(IIf(IsNull(rs!Interes), 0, rs!Interes), "#0.00")
            .TextMatrix(nFila, 5) = Format(IIf(IsNull(rs!Mora), 0, rs!Mora), "#0.00")
            .TextMatrix(nFila, 6) = Format(IIf(IsNull(rs!Cofide), 0, rs!Cofide), "#0.00")
            .TextMatrix(nFila, 7) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos), "#0.00")
            .TextMatrix(nFila, 8) = Format(IIf(IsNull(rs!Gastos), 0, rs!Gastos) + IIf(IsNull(rs!Capital), 0, rs!Capital) + IIf(IsNull(rs!Interes), 0, rs!Interes) + IIf(IsNull(rs!Mora), 0, rs!Mora) + IIf(IsNull(rs!Cofide), 0, rs!Cofide), "#0.00")
            .TextMatrix(nFila, 9) = Format(rs!cCtaCod)
            End With
        End If
        
        rs.MoveNext
    Loop
    Set rs = Nothing
    CarcularSaldos psPersCod
End Sub

Sub CarcularSaldos(ByVal psPersCod As String)
    Dim rs As ADODB.Recordset
    Dim objDRFA As COMDCredito.DCOMRFA
    Dim nFila As Integer
    
    Dim nDIFCapital As Double
    Dim nRFCCapital As Double
    Dim nRFACapital As Double
    
    Dim nDIFInteres As Double
    Dim nRFCInteres As Double
    Dim nRFAInteres As Double
    
    Dim nDIFMora As Double
    Dim nRFCMora As Double
    Dim nRFAMora As Double
    
    Dim nDIFGastos As Double
    Dim nRFCGastos As Double
    Dim nRFAGastos As Double
    
    Dim nDIFCofide As Double
    Dim nRFCCodife As Double
    Dim nRFACofide As Double
    
    
    nDIFCapital = 0
    nRFCCapital = 0
    nRFACapital = 0
    nDIFInteres = 0
    nRFCInteres = 0
    nRFAInteres = 0
    nDIFMora = 0
    nRFCMora = 0
    nRFAMora = 0
    nDIFGastos = 0
    nRFCGastos = 0
    nRFAGastos = 0
    nDIFCofide = 0
    nRFCCodife = 0
    nRFACofide = 0
    
    Set objDRFA = New COMDCredito.DCOMRFA
    Set rs = objDRFA.ListaCalendarioByPersona(psPersCod)
    Set objDRFA = Nothing
    
    Do Until rs.EOF
        If rs!Orden = "1" Then
            'Diferencial
          nDIFCapital = nDIFCapital + IIf(IsNull(rs!Capital), 0, rs!Capital)
          nDIFInteres = nDIFInteres + IIf(IsNull(rs!Interes), 0, rs!Interes)
          nDIFMora = nDIFMora + IIf(IsNull(rs!Mora), 0, rs!Mora)
          nDIFGastos = nDIFGastos + IIf(IsNull(rs!Gastos), 0, rs!Gastos)
          nDIFCofide = nDIFCofide + IIf(IsNull(rs!Cofide), 0, rs!Cofide)
        ElseIf rs!Orden = "2" Then
           nRFCCapital = nRFCCapital + IIf(IsNull(rs!Capital), 0, rs!Capital)
           nRFCInteres = nRFCInteres + IIf(IsNull(rs!Interes), 0, rs!Interes)
           nRFCMora = nRFCMora + IIf(IsNull(rs!Mora), 0, rs!Mora)
           nRFCGastos = nRFCGastos + IIf(IsNull(rs!Gastos), 0, rs!Gastos)
           nRFCCodife = nRFCCodife + IIf(IsNull(rs!Cofide), 0, rs!Cofide)
        ElseIf rs!Orden = "3" Then
            nRFACapital = nRFACapital + IIf(IsNull(rs!Capital), 0, rs!Capital)
            nRFAInteres = nRFAInteres + IIf(IsNull(rs!Interes), 0, rs!Interes)
            nRFAMora = nRFAMora + IIf(IsNull(rs!Mora), 0, rs!Mora)
            nRFAGastos = nRFAGastos + IIf(IsNull(rs!Gastos), 0, rs!Gastos)
            nRFACofide = nRFACofide + IIf(IsNull(rs!Cofide), 0, rs!Cofide)
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    LblSaldoCapitalDIF.Caption = Format(nDIFCapital, "#0.00")
    LblInteresDIf.Caption = Format(nDIFInteres, "#0.00")
    LblMoraDif.Caption = Format(nDIFMora, "#0.00")
    LblCofideDif.Caption = Format(nDIFCofide, "#0.00")
    LblGastosDif.Caption = Format(nDIFGastos, "#0.00")
    
    LblCapitalRFC.Caption = Format(nRFCCapital, "#0.00")
    LblInteresRFC.Caption = Format(nRFCInteres, "#0.00")
    LblMoraRFC.Caption = Format(nRFCMora, "#0.00")
    LblCofideRFC.Caption = Format(nRFCCodife, "#0.00")
    LblGastosRFC.Caption = Format(nRFCGastos, "#0.00")
    
    LblCapitalRFA.Caption = Format(nRFACapital, "#0.00")
    LblInteresRFA.Caption = Format(nRFAInteres, "#0.00")
    LblMoraRFA.Caption = Format(nRFAMora, "#0.00")
    LblCofideRFA.Caption = Format(nRFACofide, "#0.00")
    LblGastosRFA.Caption = Format(nRFAGastos, "#0.00")
    
    LblDeuda.Caption = Val(LblSaldoCapitalDIF) + Val(LblInteresDIf) + Val(LblMoraDif) + Val(LblCofideDif) + Val(LblGastosDif)
    LblDeuda = Val(LblDeuda) + Val(LblCapitalRFC) + Val(LblInteresRFC) + Val(LblMoraRFC) + Val(LblCofideRFC) + Val(LblGastosRFC)
    LblDeuda = Val(LblDeuda) + Val(LblCapitalRFA) + Val(LblInteresRFA) + Val(LblMoraRFA) + Val(LblCofideRFA) + Val(LblGastosRFA)
End Sub
