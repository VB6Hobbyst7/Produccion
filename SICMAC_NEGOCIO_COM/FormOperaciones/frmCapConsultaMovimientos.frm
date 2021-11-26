VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapConsultaMovimientos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7410
   ClientLeft      =   450
   ClientTop       =   1305
   ClientWidth     =   11565
   Icon            =   "frmCapConsultaMovimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7695
      TabIndex        =   9
      Top             =   6750
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8715
      TabIndex        =   10
      Top             =   6750
      Width           =   915
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   6750
      Width           =   915
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Datos Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1995
      Left            =   7860
      TabIndex        =   13
      Top             =   60
      Width           =   3600
      Begin VB.Frame fraRango 
         Height          =   975
         Left            =   180
         TabIndex        =   21
         Top             =   900
         Width           =   3135
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   2100
            TabIndex        =   7
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox txtNumMov 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   900
            MaxLength       =   6
            TabIndex        =   6
            Top             =   360
            Width           =   720
         End
         Begin MSMask.MaskEdBox txtFecFin 
            Height          =   315
            Left            =   540
            TabIndex        =   5
            Top             =   540
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin MSMask.MaskEdBox txtFecIni 
            Height          =   315
            Left            =   540
            TabIndex        =   4
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label lblNumMov 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   660
            TabIndex        =   28
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lblFin 
            AutoSize        =   -1  'True
            Caption         =   "Al :"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   225
         End
         Begin VB.Label lblIni 
            AutoSize        =   -1  'True
            Caption         =   "Del :"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.Frame fraTipoBus 
         Height          =   735
         Left            =   180
         TabIndex        =   20
         Top             =   180
         Width           =   3135
         Begin VB.OptionButton optTipoBus 
            Caption         =   "&Ultimos Movimientos"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   3
            Top             =   420
            Width           =   1815
         End
         Begin VB.OptionButton optTipoBus 
            Caption         =   "&Rango  Fechas"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   180
            Width           =   1515
         End
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1980
      Left            =   45
      TabIndex        =   11
      Top             =   60
      Width           =   7800
      Begin VB.Frame Frame2 
         Height          =   1305
         Left            =   5130
         TabIndex        =   14
         Top             =   570
         Width           =   2550
         Begin VB.Label lblTEA 
            Alignment       =   2  'Center
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
            Left            =   960
            TabIndex        =   41
            Top             =   555
            Width           =   1470
         End
         Begin VB.Label Label3 
            Caption         =   "TEA :"
            Height          =   255
            Left            =   180
            TabIndex        =   40
            Top             =   600
            Width           =   615
         End
         Begin VB.Label LblVence 
            AutoSize        =   -1  'True
            Caption         =   "Vence:"
            Height          =   195
            Left            =   165
            TabIndex        =   34
            Top             =   600
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblVencimiento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   960
            TabIndex        =   33
            Top             =   555
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   960
            TabIndex        =   19
            Top             =   870
            Width           =   1470
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   960
            TabIndex        =   18
            Top             =   225
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   930
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   285
            Width           =   690
         End
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1095
         Left            =   180
         TabIndex        =   1
         Top             =   780
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   1931
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cliente-RE-DIRECCION-cPersCod"
         EncabezadosAnchos=   "250-4000-500-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3900
         TabIndex        =   15
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Frame fraMov 
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4635
      Left            =   30
      TabIndex        =   12
      Top             =   2055
      Width           =   11460
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "&Ver Historico"
         Height          =   375
         Left            =   1440
         TabIndex        =   38
         Top             =   4155
         Visible         =   0   'False
         Width           =   1275
      End
      Begin TabDlg.SSTab SSMovs 
         Height          =   3135
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5530
         _Version        =   393216
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Movimientos"
         TabPicture(0)   =   "frmCapConsultaMovimientos.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblTotAbono"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblTotCargo"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "grdMov"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Movimientos ATM"
         TabPicture(1)   =   "frmCapConsultaMovimientos.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdMovP"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Débitos Automáticos"
         TabPicture(2)   =   "frmCapConsultaMovimientos.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdImpDebAuto"
         Tab(2).Control(1)=   "cmdAfilDebAuto"
         Tab(2).Control(2)=   "grdMovDA"
         Tab(2).ControlCount=   3
         Begin VB.CommandButton cmdImpDebAuto 
            Caption         =   "&Imprimir"
            Height          =   350
            Left            =   -64920
            TabIndex        =   45
            Top             =   2700
            Width           =   915
         End
         Begin VB.CommandButton cmdAfilDebAuto 
            Caption         =   "Ver Afiliaciones"
            Height          =   350
            Left            =   -66360
            TabIndex        =   44
            Top             =   2700
            Width           =   1395
         End
         Begin SICMACT.FlexEdit grdMov 
            Height          =   2295
            Left            =   60
            TabIndex        =   46
            Top             =   360
            Width           =   11050
            _ExtentX        =   19500
            _ExtentY        =   4048
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Fecha-Operacion-Documento-Abono-Cargo-Saldo Cnt.-Ag./Estab.-Usuario-Saldo Disp.-cOpeCod"
            EncabezadosAnchos=   "250-1450-2700-1000-1000-1000-1200-1300-800-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-2-2-2-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit grdMovP 
            Height          =   2175
            Left            =   -74940
            TabIndex        =   47
            Top             =   360
            Width           =   11050
            _ExtentX        =   19500
            _ExtentY        =   3836
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Fecha-Operacion-Documento-Abono-Cargo-Saldo Cnt.-Agencia-Usuario-Saldo Disp.-cOpeCod"
            EncabezadosAnchos=   "250-1450-3700-1000-1000-1000-0-1300-800-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-2-2-2-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit grdMovDA 
            Height          =   2295
            Left            =   -74940
            TabIndex        =   48
            Top             =   360
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4048
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Fecha-Operacion-Documento-Cargo-Saldo Cnt.-Agencia-Usuario-Saldo Disp.-cOpeCod"
            EncabezadosAnchos=   "250-1450-3700-1000-1000-1200-1300-800-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-2-2-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbOrdenaCol     =   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblTotCargo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   6400
            TabIndex        =   51
            Top             =   2760
            Width           =   1005
         End
         Begin VB.Label lblTotAbono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   5415
            TabIndex        =   50
            Top             =   2760
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3160
            TabIndex        =   49
            Top             =   2760
            Width           =   2295
         End
      End
      Begin VB.Label lblCampana 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   8880
         TabIndex        =   52
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblRemuneraciones 
         Caption         =   "Total 4 Últimas Remu. Brutas"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3900
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblTotalRemuneraciones 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   4200
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblCapRet 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO RET.  ACTUAL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7890
         TabIndex        =   36
         Top             =   4200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblSaldoRet 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   9690
         TabIndex        =   35
         Top             =   4200
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblFecSaldIni 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblSaldoInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO INICIAL AL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO CONTABLE :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4515
         TabIndex        =   27
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO DISPONIBLE :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4515
         TabIndex        =   26
         Top             =   3900
         Width           =   1815
      End
      Begin VB.Label lblSaldoContable 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6315
         TabIndex        =   25
         Top             =   4200
         Width           =   1200
      End
      Begin VB.Label lblSaldoDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6315
         TabIndex        =   24
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label lblAvisoCTS 
         Caption         =   "* Saldo total de retiro de todas las cuentas CTS del cliente con el mismo empleador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   42
         Top             =   3720
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.CheckBox ChkComision 
      Caption         =   "Cobrar comisión por página Impresa"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmCapConsultaMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As COMDConstantes.Producto
Dim nmoneda As COMDConstantes.Moneda
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim nSaldoDisponible As Double
Dim nSaldoContable As Double
Dim nInteres As Double
Dim nIntPag As Double
Dim nSaldoDispIni As Double
Dim nSaldoCntIni As Double
Dim nSaldoRetiroCTS As Double, nBloqueoParcial As Double
Dim nEstado As COMDConstantes.CaptacEstado
Dim nPersoneria As PersPersoneria
Dim cSiglas As String

Dim sNumTarj As String
Dim sCuenta As String
Dim oAho As COMDCaptaGenerales.DCOMCaptaGenerales
Dim nValCons As Integer
Dim fbPermisoLibreImp As Boolean 'JUEZ 20151229
Dim fsBoleta As String, sBoletaITF As String 'JUEZ 20151229

'Funcion de Impresion de Boletas
Private Sub ImprimeBoleta(ByVal psBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, psBoleta
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

'JUEZ 20151229 *****************************************************
Private Sub ImprimeExtractoCuenta(ByVal psImpCad As String)
Dim vPrevio As previo.clsprevio
Set vPrevio = New clsprevio
    vPrevio.PrintSpool sLpt, oImpresora.gPrnTpoLetraRoman1PDef & psImpCad & oImpresora.gPrnTamLetra10CPI
    Do While MsgBox("¿Desea Re-Imprimir el Extracto de Cuenta?", vbInformation + vbYesNo, "Aviso") = vbYes
        vPrevio.PrintSpool sLpt, oImpresora.gPrnTpoLetraRoman1PDef & psImpCad & oImpresora.gPrnTamLetra10CPI
    Loop
Set vPrevio = Nothing
End Sub
'END JUEZ **********************************************************

Private Function GetMontoDescuento(pnTipoDescuento As CaptacParametro, Optional pnCntPag As Integer = 0, Optional psCtaCod As String = "") As Double
'APRI20190109 ERS077-2018 Add psCtaCod
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As New ADODB.Recordset

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'Set rsPar = oParam.GetTarifaParametro(nOperacion, nmoneda, pnTipoDescuento)
    Set rsPar = oParam.GetTarifaParametro(nOperacion, nmoneda, pnTipoDescuento, psCtaCod) 'APRI20190109 ERS077-2018
Set oParam = Nothing


If rsPar.EOF And rsPar.BOF Then
    GetMontoDescuento = 0
Else
    Select Case pnTipoDescuento
        Case gDctoExtMNxPag, gDctoExtMExPag
            GetMontoDescuento = rsPar("nParValor") * pnCntPag
        Case gDctoExtMN, gDctoExtME
            GetMontoDescuento = rsPar("nParValor")
        Case Else
            GetMontoDescuento = rsPar("nParValor")
    End Select
End If
rsPar.Close
Set rsPar = Nothing
End Function

Private Function MuestraExtornos() As Boolean
    Dim rsPar As New ADODB.Recordset
    Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion

    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = oCap.GetTarifaParametro(nOperacion, nmoneda, gVisualizaExtornoExtracto)
    Set oCap = Nothing

    If rsPar.EOF And rsPar.BOF Then
        MuestraExtornos = True
    Else
        If rsPar("nParValor") = 0 Then
            MuestraExtornos = False
        Else
            MuestraExtornos = True
        End If
    End If
    rsPar.Close
    Set rsPar = Nothing
End Function

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsCta As ADODB.Recordset
    Dim rsRel  As New ADODB.Recordset
    Dim nRow As Long
    Dim sMoneda As String, sPersona As String
    Dim nTpoPrograma As Integer 'BRGO 20111226
    Dim lnRemBruCTS As Currency ''***Agregado por ELRO el 20121015, según OYP-RFC101-2012
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)

    If Not (rsCta.EOF And rsCta.BOF) Then
        nEstado = rsCta("nPrdEstado")
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy")
        lblEstado = rsCta("cEstado")
        lblEstado.ToolTipText = rsCta("cEstado")
        nPersoneria = rsCta("nPersoneria")
        nmoneda = CLng(Mid(sCuenta, 9, 1))
        lblTEA.Caption = Format(Round(ConvierteTNAaTEA(rsCta("nTasaInteres")), 2), "#,##0.00") & " %" 'Add By GITU 17-04-2013
        If nEstado = gCapEstAnulada Or nEstado = gCapEstCancelada Then
            nSaldoDisponible = 0
            nSaldoContable = 0
        Else
            nSaldoDisponible = rsCta("nSaldoDisp")
            nSaldoContable = rsCta("nSaldo")
        End If
        nInteres = rsCta("nIntAcum")
        nIntPag = 0
        If Mid(sCuenta, 6, 3) = Producto.gCapPlazoFijo Then
            nIntPag = rsCta("nIntpag")
        End If
    
        If nmoneda = gMonedaNacional Then
            sMoneda = "MN" '"MONEDA NACIONAL"
        Else
            sMoneda = "ME" '"MONEDA EXTRANJERA"
        End If
    
        Select Case nProducto
            Case gCapAhorros
                If rsCta("bOrdPag") Then
                    lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                Else
                    lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                End If
                nBloqueoParcial = rsCta("nBloqueoParcial")
            Case gCapPlazoFijo
                lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & sMoneda
            Case gCapCTS
                cSiglas = rsCta("EmpSiglas")
                lblMensaje = "DEPOSITO CTS" & Chr$(13) & sMoneda
                nSaldoRetiroCTS = Format$(rsCta("nSaldretiro"), "#,##0.00")
                '***Agregado por ELRO el 20121015, según OYP-RFC101-2012
                clsMant.obtenerHistorialCaptacSueldosCTS sCuenta, lnRemBruCTS
                lblTotalRemuneraciones = Format$(lnRemBruCTS, "#,##0.00")
                '***Fin Agregado por ELRO el 20121015
        End Select
            
        '*** BRGO 20111220 ***********************************
        nTpoPrograma = rsCta("nTpoPrograma")
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Set rsRel = clsGen.GetConstante(IIf(nProducto = gCapAhorros, gCaptacSubProdAhorros, IIf(nProducto = gCapPlazoFijo, gCaptacSubProdPlazoFijo, gCaptacSubProdCTS)), , CStr(nTpoPrograma))
        lblMensaje = lblMensaje & "-" & rsRel!CDescripcion
        Set clsGen = Nothing
        Set rsRel = Nothing
        '*** END BRGO ****************************************
        
        If nProducto = gCapCTS Then lblMensaje = lblMensaje & IIf(rsCta("bCeseLaboral"), " - CESE LABORAL", "") 'JUEZ 20140319
        
        If rsCta("cCampanaDesc") <> "" Then lblCampana.Caption = "TASA DE CAMPAÑA: " & rsCta("cCampanaDesc") 'JUEZ 20160425
        
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        Do While Not rsRel.EOF
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 2) = Left(UCase(rsRel("Relacion")), 2)
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Direccion"))
                grdCliente.TextMatrix(nRow, 4) = Trim(rsRel("cPersCod")) 'FRHU ERS077-2015 20151204
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        rsRel.Close
        Set rsRel = Nothing
        txtCuenta.Enabled = False
        cmdCancelar.Enabled = True
        fraBuscar.Enabled = True
'        optTipoBus(0).SetFocus
    Else
        MsgBox "Cuenta NO Existe", vbInformation, "Operacion"
        txtCuenta.SetFocus
    End If
    Set clsMant = Nothing
End Sub

Private Sub LimpiaControles()
cSiglas = ""
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera

'by capi 11112008
grdMovP.Clear
grdMovP.Rows = 2
grdMovP.FormaCabecera
grdMovP.Visible = True
'
grdMov.Clear
grdMov.Rows = 2
grdMov.FormaCabecera
grdMov.Visible = True

'JUEZ 20150224 **********
grdMovDA.Clear
grdMovDA.Rows = 2
grdMovDA.FormaCabecera
grdMovDA.Visible = True
'END JUEZ ***************

nPersoneria = -1
txtCuenta.Age = ""
txtCuenta.Cuenta = ""
lblApertura = ""
lblEstado = ""
lblTEA = ""
fraBuscar.Enabled = False
fraMov.Enabled = False
cmdImprimir.Enabled = False
cmdCancelar.Enabled = False
ChkComision.value = 0
ChkComision.Enabled = False
txtCuenta.Enabled = True
txtFecIni = "__/__/____"
txtFecFin = "__/__/____"
txtNumMov = ""
lblMensaje = ""
lblSaldoContable = ""
lblSaldoDisponible = ""
lblTotAbono = ""
lblTotCargo = ""
lblTotCargo.Visible = True
txtCuenta.SetFocus
lblFecSaldIni = ""
lblSaldoInicial = ""
lblSaldoRet.Caption = ""
'***Agregado por ELRO el 20121013, según OYP-RFC101-2012
lblTotalRemuneraciones.Caption = ""
'***Fin Agregado por ELRO el 20121013*******************
lblCampana.Caption = "" 'JUEZ 20160425
End Sub

Public Sub inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion)

Dim oGen As COMDConstSistema.DCOMGeneral 'JUEZ 20151229

    nProducto = nProd
    nOperacion = nOpe
    Select Case nProd
        Case gCapAhorros
            'Me.Caption = "Captaciones - Ultimos Movimientos - Ahorros"
            Me.Caption = "Captaciones - Extracto de Cuenta - Ahorros" 'JUEZ 20151229
            lblCapRet.Visible = True
            lblSaldoRet.Visible = True
            lblCapRet.Caption = "BLOQUEO PARCIAL:"
            lblAvisoCTS.Visible = False 'JUEZ 20130727
        Case gCapPlazoFijo
            'Me.Caption = "Captaciones - Ultimos Movimientos - Plazo Fijo"
            Me.Caption = "Captaciones - Extracto de Cuenta - Plazo Fijo" 'JUEZ 20151229
            lblCapRet.Visible = False
            lblSaldoRet.Visible = False
            lblAvisoCTS.Visible = False 'JUEZ 20130727
            SSMovs.TabVisible(2) = False 'JUEZ 20150224
        Case gCapCTS
            'Me.Caption = "Captaciones - Ultimos Movimientos - CTS"
            Me.Caption = "Captaciones - Extracto de Cuenta - CTS" 'JUEZ 20151229
            lblCapRet.Visible = True
            lblSaldoRet.Visible = True
            lblCapRet.Caption = "SALDO RETIRO:"
            '***Agregado por ELRO el 20121013, según OYP-RFC101-2012
            lblRemuneraciones.Visible = True
            lblTotalRemuneraciones.Visible = True
            cmdHistorico.Visible = True
            '***Fin Agregado por ELRO el 20121013*******************
            lblAvisoCTS.Visible = True 'JUEZ 20130727
    End Select
    txtCuenta.Prod = Trim(Str(nProducto))
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledCMAC = False
    txtCuenta.EnabledProd = False
    cmdCancelar.Enabled = False
    optTipoBus(0).value = True
    fraMov.Enabled = False
    fraBuscar.Enabled = False
    cmdImprimir.Enabled = False
    cmdCancelar.Enabled = False
    ChkComision.Enabled = False
    
    'JUEZ 20151229 *************************************************************
    Set oGen = New COMDConstSistema.DCOMGeneral
    fbPermisoLibreImp = oGen.VerificaExistePermisoCargo(gsCodCargo, gMantParamLibreImpConsMov)
    Set oGen = Nothing
    'JUEZ 20151229 *************************************************************
    
    'ADD By GITU para el uso de las operaciones con tarjeta
    Set oAho = New COMDCaptaGenerales.DCOMCaptaGenerales

    nValCons = oAho.GetConsultaSaldoSinTarjeta(gsCodCargo)
    
    If (gnCodOpeTarj = 1 And nValCons = 0) Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If sCuenta <> "123456789" Then
            If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
        
        
            If sCuenta <> "" Then
                txtCuenta.NroCuenta = sCuenta
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
            End If
            If sCuenta <> "" Then
                Me.Show 1
            End If
        End If
        If sCuenta <> "" Then
            Me.Show 1
        End If
    Else
        Me.Show 1
    End If
    'End GITU
End Sub

'JUEZ 20150224 *************************
Private Sub cmdAfilDebAuto_Click()
    frmServCobDebitoAuto.IniciaConsulta txtCuenta.NroCuenta
End Sub

Private Sub cmdImpDebAuto_Click()
Dim sCuenta As String
Dim sMoneda As String
Dim i As Integer, j As Integer
Dim nTotCargo As Double
Dim sTotCargo As String * 13
Dim sSaldoDisponible As String * 13, sSaldoRetiro As String * 13
Dim sSaldoContable As String * 13
Dim sInteresMes As String * 13, sInteresPagado As String * 13
Dim sTitRp1 As String, sTitRp2 As String, sUser As String
Dim rsDA As New ADODB.Recordset

Dim rsC As New ADODB.Recordset
Dim clsPrev As previo.clsprevio
Dim oCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
Dim lsCadImp As String
Dim nCntPag As Integer

Dim lsTEA As String

If Not grdMovDA.Rows - 1 > 1 Then
    MsgBox "No existen registros para realizar la impresión", vbInformation, "Aviso"
    Exit Sub
End If

sCuenta = txtCuenta.NroCuenta
sMoneda = IIf(nmoneda = gMonedaNacional, "SOLES", "DOLARES")
sTitRp1 = space(45) & "EXTRACTO DE CUENTA DEBITO AUTOMÁTICO"
If optTipoBus(1).value Then
    sTitRp2 = space(48) & "   ÚLTIMOS " & Trim(txtNumMov) & " MOVIMIENTOS"
Else
    sTitRp2 = space(47) & "( DEL " & txtFecIni & " AL " & txtFecFin & " )"
End If

lsTEA = lblTEA.Caption

RSet sTotCargo = lblTotCargo
RSet sSaldoDisponible = lblSaldoDisponible
RSet sSaldoContable = lblSaldoContable
RSet sInteresMes = Format$(nInteres, "#,##0.00")

If nProducto = gCapAhorros Or nProducto = gCapCTS Then
    RSet sSaldoRetiro = lblSaldoRet
Else
    RSet sInteresPagado = Format$(nIntPag, "#,##0.00")
End If
        
    With rsDA
        'Crear RecordSet
        .Fields.Append "dFecha", adDate
        .Fields.Append "sOperacion", adVarChar, 200
        .Fields.Append "sDocumento", adVarChar, 50
        .Fields.Append "sRetiro", adVarChar, 50
        .Fields.Append "sSaldoCnt", adVarChar, 50
        .Fields.Append "sAgencia", adVarChar, 200
        .Fields.Append "sUser", adVarChar, 50
        .Fields.Append "sOpeCod", adVarChar, 6
        .Open
        'Llenar Recordset
    
        For i = 1 To grdMovDA.Rows - 1
            .AddNew
            .Fields("dFecha") = Trim(grdMovDA.TextMatrix(i, 1))
            .Fields("sOperacion") = grdMovDA.TextMatrix(i, 2)
            .Fields("sDocumento") = grdMovDA.TextMatrix(i, 3)
            .Fields("sRetiro") = grdMovDA.TextMatrix(i, 4)
            .Fields("sSaldoCnt") = grdMovDA.TextMatrix(i, 5)
            .Fields("sAgencia") = Left(Mid(grdMovDA.TextMatrix(i, 6), 9), 10)
            .Fields("sUser") = grdMovDA.TextMatrix(i, 7)
            .Fields("sOpeCod") = grdMovDA.TextMatrix(i, 9)
            nTotCargo = nTotCargo + CDbl(grdMovDA.TextMatrix(i, 4))
        Next i
        rsDA.Sort = "dFecha,sOpeCod"
    End With
    
    sTotCargo = CStr(nTotCargo)
    
    With rsC
        'Crear RecordSet
        .Fields.Append "sNombre", adVarChar, 200
        .Fields.Append "sRelacion", adVarChar, 50
        .Fields.Append "sDireccion", adVarChar, 200
        .Open
        'Llenar Recordset
        For i = 1 To grdCliente.Rows - 1
            .AddNew
            .Fields("sNombre") = Trim(grdCliente.TextMatrix(i, 1))
            .Fields("sRelacion") = Trim(grdCliente.TextMatrix(i, 2))
            .Fields("sDireccion") = Trim(grdCliente.TextMatrix(i, 3))
        Next i
    End With
        
If nProducto = gCapCTS And cSiglas <> "" Then
   sCuenta = sCuenta & " - " & cSiglas
End If
Set oCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
oCapImp.IniciaImpresora gImpresora

lsCadImp = oCapImp.ImprimeConsultaMovimientoDebitoAuto(rsDA, rsC, gsCodUser, gsNomCmac, gsNomAge, sMoneda, gdFecSis, sTitRp1, sTitRp2, sCuenta, lblFecSaldIni.Caption, _
            lblSaldoInicial.Caption, sSaldoDisponible, sSaldoContable, sSaldoRetiro, sInteresMes, sInteresPagado, lblEstado.Caption, sTotCargo, nCntPag, lsTEA)
            
Set oCapImp = Nothing
Set clsPrev = New previo.clsprevio
clsPrev.Show lsCadImp, "Extracto de Cuenta", True, , gImpresora
Set clsPrev = Nothing

If Not CargoAutomatico(sCuenta, nCntPag, True) Then Exit Sub
   
End Sub
'END JUEZ ******************************

Private Sub cmdBuscar_Click()
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    'By Capi 11112008
    Dim rsCtaP As ADODB.Recordset
    '
    Dim rsCtaDA As ADODB.Recordset 'JUEZ 20150224
    Dim sCuenta As String
    Dim nAbonoTotal As Double, nCargoTotal As Double
    Dim nCapitalizacion As Double
    Dim sCtaAho As String
    
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Set clsGen = New COMDConstSistema.DCOMGeneral 'DGeneral
    
    On Error GoTo ErrorCmdBuscar_Click 'JUEZ 20150224
    
    If optTipoBus(0).value Then
        If Not IsDate(txtFecIni) Then
            MsgBox "Fecha No Válida", vbInformation, "SICMACM - Aviso"
            txtFecIni.SetFocus
            Exit Sub
        End If
        If Not IsDate(txtFecFin) Then
            MsgBox "Fecha No Válida", vbInformation, "SICMACM - Aviso"
            txtFecFin.SetFocus
            Exit Sub
        End If
        If DateDiff("d", CDate(txtFecFin), CDate(txtFecIni)) > 0 Then
            MsgBox "La Fecha Inicial no puede ser mayor que la Fecha Final", vbInformation, "SICMACM - Aviso"
            txtFecIni.SetFocus
            Exit Sub
        End If
    ElseIf optTipoBus(1).value Then
        If txtNumMov = "" Then
            MsgBox "Número de Movimientos NO Válidos", vbInformation, "SICMACM - Aviso"
            txtNumMov.SetFocus
            Exit Sub
        End If
        If CLng(Val(txtNumMov.Text)) = 0 Then
            MsgBox "Número de Movimientos NO Válidos", vbInformation, "SICMACM - Aviso"
            txtNumMov.SetFocus
            Exit Sub
        End If
    End If
    
    sCuenta = txtCuenta.NroCuenta
    
    '*** BRGO 20111220 **********************************************
    If Mid(sCuenta, 6, 3) = gCapPlazoFijo Then
        sCtaAho = clsGen.GetCuentaAntigua(sCuenta)
        If sCtaAho <> "" And Mid(sCtaAho, 6, 3) = gCapAhorros Then
            sCuenta = sCuenta & "," & sCtaAho
        End If
    End If
    '*** END BRGO ***************************************************
    
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    If optTipoBus(0).value Then
        Set rsCta = clsMant.GetMovimientosCuenta(sCuenta, CDate(txtFecIni), CDate(txtFecFin), , MuestraExtornos())
        'By Capi 11112008 caj aut
        Set rsCtaP = clsMant.GetMovimientosPendientesCuenta(sCuenta, CDate(txtFecIni), CDate(txtFecFin), , MuestraExtornos())
        '
        Set rsCtaDA = clsMant.GetMovimientosDebitosAutoCuenta(sCuenta, CDate(txtFecIni), CDate(txtFecFin), , MuestraExtornos()) 'JUEZ 20150224
    ElseIf optTipoBus(1).value Then
        Set rsCta = clsMant.GetMovimientosCuenta(sCuenta, , , CLng(Val(txtNumMov.Text)), MuestraExtornos())
        'By Capi 11112008 caj aut
        Set rsCtaP = clsMant.GetMovimientosPendientesCuenta(sCuenta, , , CLng(Val(txtNumMov.Text)), MuestraExtornos())
        '
        Set rsCtaDA = clsMant.GetMovimientosDebitosAutoCuenta(sCuenta, , , CLng(Val(txtNumMov.Text)), MuestraExtornos()) 'JUEZ 20150224
    End If

'If rsCta.EOF And rsCta.BOF Then
If rsCta.EOF And rsCta.BOF And rsCtaDA.EOF And rsCtaDA.BOF Then 'JUEZ 20150224
    MsgBox "No se encontraron movimientos para la cuenta " & sCuenta, vbInformation, "Aviso"
    If txtFecIni.Visible Then
        txtFecIni.SetFocus
    ElseIf txtNumMov.Visible Then
        txtNumMov.SetFocus
    End If
    Exit Sub
Else
    '*** BRGO 20111220 ***************************
    If sCtaAho <> "" Then
        sCuenta = Replace(sCuenta, "," & sCtaAho, "")
    End If
    '*** END BRGO ********************************
    If optTipoBus(0).value Then
        Set grdMov.Recordset = rsCta
        'By Capi 11112008 caj aut.
        Set grdMovP.Recordset = rsCtaP
        Set grdMovDA.Recordset = rsCtaDA 'JUEZ 20150224
        
        '***Agregado por ELRO el 20130212, según INC1301110013
        rsCta.MoveLast
        Do While Not rsCta.BOF
            If rsCta("cOpeCod") = 210401 Then
                nCapitalizacion = nCapitalizacion + Format$((-1) * rsCta("nAbono"), "#,##0.00")
            End If
            rsCta.MovePrevious
        Loop
        '***Fin Agregado por ELRO el 20130212*****************
        '
        'Set rsCta = clsMant.GetSaldoFecha(sCuenta, DateAdd("d", -1, CDate(txtFecIni))) 'RIRO20150511 ERS146-2014, comentado
        rsCta.MoveFirst 'RIRO20150511 ERS146-2014
                
        'ojo
        If rsCta.EOF And rsCta.BOF Then
            lblFecSaldIni = Format$(DateAdd("d", -1, CDate(txtFecIni)), "dd/mm/yyyy")
            lblSaldoInicial = "0.00"
        Else
            If IsDate(rsCta("Fecha")) Then
                lblFecSaldIni = Left(DateAdd("d", -1, rsCta("Fecha")), 12)
            End If
            'lblSaldoInicial = Format$(rsCta("nSaldoContable"), "#,##0.00")' RIRO20150511 ERS146-2014, comentado
            lblSaldoInicial = Format$(rsCta("nSaldoContable") - rsCta("nAbono") - rsCta("nCargo"), "#,##0.00") ' RIRO20150511 ERS146-2014, Add
        End If
    Else
        rsCta.MoveLast
        If rsCta.RecordCount <= CLng(Val(txtNumMov.Text)) Then
            'RIRO20150512 *****************
            If IsDate(rsCta("Fecha")) Then
                lblFecSaldIni = Left(DateAdd("d", -1, rsCta("Fecha")), 12)
            End If
            lblSaldoInicial = Format$(rsCta("nSaldoContable") - rsCta("nAbono") - rsCta("nCargo"), "#,##0.00")
            ' END RIRO ********************
            
            'lblFecSaldIni = Format$(DateAdd("d", -1, CDate(Left(rsCta("Fecha"), 10))), "dd/mm/yyyy") 'RIRO20150512, comentado
            'lblSaldoInicial = "0.00" 'RIRO20150512, comentado
                        
        Else
            lblFecSaldIni = Format$(CDate(Left(rsCta("Fecha"), 10)), "dd/mm/yyyy")
            lblSaldoInicial = Format$(rsCta("nSaldoContable"), "#,##0.00")
            rsCta.MovePrevious
        End If
        Do While Not rsCta.BOF
            grdMov.AdicionaFila
            grdMov.TextMatrix(grdMov.Rows - 1, 1) = rsCta("Fecha")
            grdMov.TextMatrix(grdMov.Rows - 1, 2) = rsCta("Operacion")
            grdMov.TextMatrix(grdMov.Rows - 1, 3) = rsCta("cDocumento")
           
           
            If rsCta("cOpeCod") = 210401 Then
                nCapitalizacion = nCapitalizacion + Format$((-1) * rsCta("nAbono"), "#,##0.00")
            End If
            grdMov.TextMatrix(grdMov.Rows - 1, 4) = Format$(rsCta("nAbono"), "#,##0.00")
            grdMov.TextMatrix(grdMov.Rows - 1, 5) = Format$(rsCta("nCargo"), "#,##0.00")
            grdMov.TextMatrix(grdMov.Rows - 1, 6) = Format$(rsCta("nSaldoContable"), "#,##0.00")
            grdMov.TextMatrix(grdMov.Rows - 1, 7) = rsCta("cAgencia") 'EAAS20180611 SE QUITO EAAS20190129 cAgenciaEstablecimientoc
            grdMov.TextMatrix(grdMov.Rows - 1, 8) = rsCta("cUsu")
            grdMov.TextMatrix(grdMov.Rows - 1, 10) = rsCta("cOpeCod") 'JUEZ 20140128
            rsCta.MovePrevious
        Loop
        'By capi 11112008 caj aut
        If rsCtaP.RecordCount > 0 Then
            rsCtaP.MoveLast
            Do While Not rsCtaP.BOF
                grdMovP.AdicionaFila
                grdMovP.TextMatrix(grdMovP.Rows - 1, 1) = rsCtaP("Fecha")
                grdMovP.TextMatrix(grdMovP.Rows - 1, 2) = rsCtaP("Operacion")
                grdMovP.TextMatrix(grdMovP.Rows - 1, 3) = rsCtaP("cDocumento")
                If rsCtaP("cOpeCod") = 210401 Then
                    nCapitalizacion = nCapitalizacion + Format$((-1) * rsCtaP("nAbono"), "#,##0.00")
                End If
                grdMovP.TextMatrix(grdMovP.Rows - 1, 4) = Format$(rsCtaP("nAbono"), "#,##0.00")
                grdMovP.TextMatrix(grdMovP.Rows - 1, 5) = Format$(rsCtaP("nCargo"), "#,##0.00")
                grdMovP.TextMatrix(grdMovP.Rows - 1, 6) = Format$(rsCtaP("nSaldoContable"), "#,##0.00") '999999999
                grdMovP.TextMatrix(grdMovP.Rows - 1, 7) = rsCtaP("cAgencia") 'EAAS20180611 SE QUITO EAAS20190129 cAgenciaEstablecimientoc
                grdMovP.TextMatrix(grdMovP.Rows - 1, 8) = rsCtaP("cUsu")
                grdMovP.TextMatrix(grdMovP.Rows - 1, 10) = rsCtaP("cOpeCod") 'JUEZ 20140128
                rsCtaP.MovePrevious
             Loop
           End If
            '
        'JUEZ 20150224 ***************************************
        If rsCtaDA.RecordCount > 0 Then
            rsCtaDA.MoveLast
            Do While Not rsCtaDA.BOF
                grdMovDA.AdicionaFila
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 1) = rsCtaDA("Fecha")
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 2) = rsCtaDA("Operacion")
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 3) = rsCtaDA("cDocumento")
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 4) = Format$(rsCtaDA("nCargo"), "#,##0.00")
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 5) = Format$(rsCtaDA("nSaldoContable"), "#,##0.00")
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 6) = rsCtaDA("cAgencia") 'EAAS20180611 SE QUITO EAAS20190129 cAgenciaEstablecimientoc
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 7) = rsCtaDA("cUsu")
                grdMovDA.TextMatrix(grdMovDA.Rows - 1, 9) = rsCtaDA("cOpeCod")
                rsCtaDA.MovePrevious
             Loop
        End If
        'END JUEZ ********************************************
    End If
    nAbonoTotal = grdMov.SumaRow(4) + nCapitalizacion
    nCargoTotal = grdMov.SumaRow(5)
    lblTotAbono = Format$(nAbonoTotal, "#,##0.00")
    lblTotCargo = Format$(nCargoTotal, "#,##0.00")
        
    If nEstado = gCapEstAnulada Or nEstado = gCapEstCancelada Then
        lblSaldoDisponible = Format$(0, "#,##0.00")
        lblSaldoContable = Format$(0, "#,##0.00")
        If nProducto = gCapAhorros Or nProducto = gCapCTS Then
            lblSaldoRet = "0.00"
        End If
    Else
        If nProducto = gCapAhorros Then
            lblSaldoDisponible = Format$(nSaldoDisponible - nBloqueoParcial, "#,##0.00")
        Else
            lblSaldoDisponible = Format$(nSaldoDisponible, "#,##0.00")
        End If
        lblSaldoContable = Format$(nSaldoContable, "#,##0.00")
        If nProducto = gCapAhorros Then
            lblSaldoRet = Format$(nBloqueoParcial, "#,##0.00")
        ElseIf nProducto = gCapCTS Then
            lblSaldoRet = Format$(nSaldoRetiroCTS, "#,##0.00")
        End If
    End If
    
    fraBuscar.Enabled = False
    fraMov.Enabled = True
    ChkComision.Enabled = True
    ChkComision.value = 1
    cmdImprimir.Enabled = True
    cmdImprimir.SetFocus
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim sMovNro As String
    Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        clsCap.CapConsultaSaldosMovimiento sCuenta, sMovNro, nOperacion, nSaldoDisponible, nSaldoContable, sNumTarj
    Set clsCap = Nothing
End If
Set clsMant = Nothing
rsCta.Close
Set rsCta = Nothing
Exit Sub
ErrorCmdBuscar_Click: 'JUEZ 20150224
    MsgBox err.Description, vbExclamation, "Aviso"
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub cmdHistorico_Click()
    Dim oForm As New frmCapHistoricoRemuneracionesCTS
    oForm.iniciarHistoricoRemBruCTS (txtCuenta.NroCuenta)
End Sub

Private Sub cmdImprimir_Click()
Dim sCuenta As String
Dim sMoneda As String
Dim i As Integer, j As Integer
Dim sTotAbono As String * 13
Dim sTotCargo As String * 13
Dim sSaldoDisponible As String * 13, sSaldoRetiro As String * 13
Dim sSaldoContable As String * 13
Dim sInteresMes As String * 13, sInteresPagado As String * 13
Dim sTitRp1 As String, sTitRp2 As String, sUser As String
Dim rsO As New ADODB.Recordset
'By capi 11112008
Dim rsOP As New ADODB.Recordset
'
Dim rsC As New ADODB.Recordset
Dim clsPrev As previo.clsprevio
Dim oCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
Dim lsCadImp As String
Dim nCntPag As Integer
'WIOR 20130122 *******************
Dim objPista As COMManejador.Pista
Dim sComentario As String
'WIOR FIN ************************

'--=========================================
'--INICIO EAAS20180511 SEGUN ERS TI-006-2018
'--=========================================
Dim rsCCI As New ADODB.Recordset
Dim oCCI As New COMNCaptaGenerales.NCOMCaptaGenerales
'--=========================================
'--FIN EAAS20180511 SEGUN ERS TI-006-2018
'--=========================================

Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento 'JUEZ 20151229

Dim lsTEA As String 'Add By GITU 18-04-2013
'FRHU ERS077-2015 20151204
Dim item As Integer
For item = 1 To grdCliente.Rows - 1
    Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(item, 4), nOperacion)
Next item
'FIN FRHU ERS077-2015

'WIOR 20121009 Clientes Observados ************************************
If nOperacion = gAhoConsMovimiento Or nOperacion = gPFConsMovimiento Or nOperacion = gCTSConsMovimiento Then
    Dim oDPersona As COMDPersona.DCOMPersona
    Dim rsPersonaCred As ADODB.Recordset
    Dim rsPersona As ADODB.Recordset
    Dim Cont As Integer
    Set oDPersona = New COMDPersona.DCOMPersona
            
    Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(Trim(txtCuenta.NroCuenta), gCapRelPersTitular)
        
    If rsPersonaCred.RecordCount > 0 Then
        If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
            For Cont = 0 To rsPersonaCred.RecordCount - 1
                Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cPersCod))
                If rsPersona.RecordCount > 0 Then
                    If Not (rsPersona.EOF And rsPersona.BOF) Then
                        If Trim(rsPersona!sUsual) = "3" Then
                            MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                            Call frmPersona.Inicio(Trim(rsPersonaCred!cPersCod), PersonaActualiza)
                        End If
                    End If
                End If
                Set rsPersona = Nothing
                rsPersonaCred.MoveNext
            Next Cont
        End If
    End If
End If
'WIOR FIN ***************************************************************

sCuenta = txtCuenta.NroCuenta
sMoneda = IIf(nmoneda = gMonedaNacional, "SOLES", "DOLARES")
sTitRp1 = space(55) & "EXTRACTO DE CUENTA"
If optTipoBus(1).value Then
    sTitRp2 = space(55) & "ÚLTIMOS " & Trim(txtNumMov) & " MOVIMIENTOS "
Else
    sTitRp2 = space(55) & "( DEL " & txtFecIni & " AL " & txtFecFin & " )"
End If

lsTEA = lblTEA.Caption 'Add By GITU 18-04-2013
Set rsCCI = oCCI.ObtieneCCI(sCuenta) '====================EAAS20180511 SEGUN ERS TI-006-2018
RSet sTotAbono = lblTotAbono
RSet sTotCargo = lblTotCargo
RSet sSaldoDisponible = lblSaldoDisponible
RSet sSaldoContable = lblSaldoContable
RSet sInteresMes = Format$(nInteres, "#,##0.00")

If nProducto = gCapAhorros Or nProducto = gCapCTS Then
    RSet sSaldoRetiro = lblSaldoRet
Else
    RSet sInteresPagado = Format$(nIntPag, "#,##0.00")
End If
        
        With rsO
            'Crear RecordSet
            .Fields.Append "dFecha", adDate
            .Fields.Append "sOperacion", adVarChar, 200
            .Fields.Append "sDocumento", adVarChar, 50
            .Fields.Append "sDeposito", adVarChar, 50
            .Fields.Append "sRetiro", adVarChar, 50
            .Fields.Append "sSaldoCnt", adVarChar, 50
            .Fields.Append "sAgencia", adVarChar, 200
            .Fields.Append "sUser", adVarChar, 50
            .Fields.Append "sOpeCod", adVarChar, 6 'JUEZ 20140128
            .Open
            'Llenar Recordset
        
            For i = 1 To grdMov.Rows - 1
                .AddNew
                .Fields("dFecha") = Trim(grdMov.TextMatrix(i, 1))
                
                .Fields("sOperacion") = grdMov.TextMatrix(i, 2)
                .Fields("sDocumento") = grdMov.TextMatrix(i, 3)
                .Fields("sDeposito") = grdMov.TextMatrix(i, 4)
                .Fields("sRetiro") = grdMov.TextMatrix(i, 5)
                .Fields("sSaldoCnt") = grdMov.TextMatrix(i, 6)
                .Fields("sAgencia") = Left(Mid(grdMov.TextMatrix(i, 7), 1, 20), 20) 'Comentado Por EAAS20180511
                '.Fields("sAgencia") = grdMov.TextMatrix(i, 7) '========= EAAS20180511 SEGUN ERS TI-006-2018
                .Fields("sUser") = grdMov.TextMatrix(i, 8)
                .Fields("sOpeCod") = grdMov.TextMatrix(i, 10)
            Next i
            'By Capi 11112008 para operaciones pendientes cajeros automaticos
            If grdMovP.row - 1 > 0 Then
                For i = 1 To grdMovP.Rows - 1
                    .AddNew
                    .Fields("dFecha") = Trim(grdMovP.TextMatrix(i, 1))
                    .Fields("sOperacion") = grdMovP.TextMatrix(i, 2)
                    .Fields("sDocumento") = grdMovP.TextMatrix(i, 3)
                    .Fields("sDeposito") = grdMovP.TextMatrix(i, 4)
                    .Fields("sRetiro") = grdMovP.TextMatrix(i, 5)
                    .Fields("sSaldoCnt") = grdMovP.TextMatrix(i, 6) '999999999
                    .Fields("sAgencia") = Left(Mid(grdMovP.TextMatrix(i, 7), 9), 10)
                    .Fields("sUser") = grdMovP.TextMatrix(i, 8)
                    '.Fields("sOpeCod") = grdMov.TextMatrix(i, 10) 'Comentado Por ARLO20170124
                    .Fields("sOpeCod") = grdMovP.TextMatrix(i, 10) 'Modificado Por ARLO20170124
             Next i
            End If
            'by capi 12112008
            'ordenando el recordset
            'rsO.Sort = "dFecha,sOperacion" 'Comentado por JUEZ 20131128
            'rsO.Sort = "dFecha,sOpeCod" 'JUEZ 20140128 'EAAS20180511 SEGUN ERS TI-006-2018 152
             
        End With
        
        
        With rsC
            'Crear RecordSet
            .Fields.Append "sNombre", adVarChar, 200
            .Fields.Append "sRelacion", adVarChar, 50
            .Fields.Append "sDireccion", adVarChar, 200
            .Fields.Append "sCCI", adVarChar, 200 '--========= EAAS20180511 SEGUN ERS TI-006-2018
            .Open
            'Llenar Recordset
            For i = 1 To grdCliente.Rows - 1
                .AddNew
                .Fields("sNombre") = Trim(grdCliente.TextMatrix(i, 1))
                .Fields("sRelacion") = Trim(grdCliente.TextMatrix(i, 2))
                .Fields("sDireccion") = Trim(grdCliente.TextMatrix(i, 3))
                .Fields("sCCI") = rsCCI!CCI '--========= EAAS20180511 SEGUN ERS TI-006-2018
            Next i
        End With
        
If nProducto = gCapCTS And cSiglas <> "" Then
   sCuenta = sCuenta & " - " & cSiglas
End If
Set oCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
oCapImp.IniciaImpresora gImpresora

lsCadImp = oCapImp.ImprimeConsultaMovimiento(rsO, rsC, gsCodUser, gsNomCmac, gsNomAge, sMoneda, gdFecSis, sTitRp1, sTitRp2, sCuenta, lblFecSaldIni.Caption, _
            lblSaldoInicial.Caption, sSaldoDisponible, sSaldoContable, sSaldoRetiro, sInteresMes, sInteresPagado, lblEstado.Caption, sTotAbono, sTotCargo, nCntPag, lsTEA)
            
Set oCapImp = Nothing
'WIOR 20130122 ************************************
'sComentario = "Consulta de Movimientos "
sComentario = "Extracto de Cuenta "
Select Case nOperacion
    Case gAhoConsMovimiento: sComentario = sComentario & "- AHORROS"
    Case gPFConsMovimiento: sComentario = sComentario & "- PLAZO FIJO"
    Case gCTSConsMovimiento: sComentario = sComentario & "- CTS"
End Select
Set objPista = New COMManejador.Pista
'objPista.InsertarPista nOperacion, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, sComentario, Trim(txtCuenta.NroCuenta), gCodigoCuenta
Set objPista = Nothing
'WIOR FIN *****************************************

'JUEZ 2015 ********************************************************
If Not fbPermisoLibreImp Then
    If Not CargoAutomatico(Left(sCuenta, 18), nCntPag) Then Exit Sub
    'ImprimeExtractoCuenta lsCadImp
End If
'END JUEZ *********************************************************

Set objPista = New COMManejador.Pista
objPista.InsertarPista nOperacion, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, sComentario, Trim(txtCuenta.NroCuenta), gCodigoCuenta 'JUEZ 20151229
Set objPista = Nothing

'JUEZ 20151229 ****************************************************
If (Not fbPermisoLibreImp) And (fsBoleta <> "") Then
    MsgBox "Se va a realizar la impresión del voucher de la comisión por extracto de cuenta", vbInformation, "Aviso"
    ImprimeBoleta fsBoleta, "Boleta de la Comisión"
End If
'END JUEZ *********************************************************

Set clsPrev = New previo.clsprevio
clsPrev.Show lsCadImp, "Extracto de Cuenta", True, , gImpresora
Set clsPrev = Nothing
'INICIO JHCU ENCUESTA 16-10-2019
Dim sOpecodEncuesta As String
sOpecodEncuesta = nOperacion
Encuestas gsCodUser, gsCodAge, "ERS0292019", sOpecodEncuesta
'FIN
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
        
    '**DAOR 20081125, Para tarjetas ***********************
    If KeyCode = vbKeyF10 And txtCuenta.Enabled Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, nProducto)
        If Val(Mid(sCuenta, 6, 3)) <> nProducto And sCuenta <> "" Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    '*******************************************************
    '-- =============================================
    '-- EAAS 20180405 SEGÚN ERS Nº006-2018
    '-- =============================================
    If txtCuenta.Enabled = False Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If

        If KeyCode = 17 And Shift = 2 Then
           KeyCode = 10
        End If

        If KeyCode = 67 And Shift = 2 Then
           KeyCode = 10
        End If

        If KeyCode = 113 And Shift = 0 Then
            KeyCode = 10
        End If
    End If
    '-- =============================================
    '-- EAAS 20180405 SEGÚN ERS Nº006-2018
    '-- =============================================
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub


Private Sub optTipoBus_Click(Index As Integer)
Select Case Index
    Case 0
        lblIni.Visible = True
        lblFin.Visible = True
        txtFecIni.Visible = True
        txtFecFin.Visible = True
        lblNumMov.Visible = False
        txtNumMov.Visible = False
    Case 1
        lblIni.Visible = False
        lblFin.Visible = False
        txtFecIni.Visible = False
        txtFecFin.Visible = False
        lblNumMov.Visible = True
        txtNumMov.Visible = True
End Select
End Sub

Private Sub optTipoBus_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFecIni.Visible Then
        txtFecIni.SetFocus
    ElseIf txtNumMov.Visible Then
        txtNumMov.SetFocus
    End If
End If
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        Dim sCuenta As String
        Dim clsGen As COMDConstSistema.DCOMGeneral
        Set clsGen = New COMDConstSistema.DCOMGeneral 'DGeneral
       
        sCta = txtCuenta.NroCuenta
        '*** BRGO 20111220 ******************************************
        If Mid(sCta, 6, 3) = gCapAhorros Then
            sCuenta = clsGen.GetCuentaNueva(sCta)
            If sCuenta <> "" And Mid(sCuenta, 6, 3) = gCapPlazoFijo Then
                MsgBox "La cuenta ingresada ha sido migrado a la cuenta de Plazo Fijo N°" & sCuenta & ". Favor ingrese a la opción de Consulta de Movimientos de Plazo Fijo", vbInformation, "Aviso"
                txtCuenta.Age = ""
                txtCuenta.Cuenta = ""
                txtCuenta.SetFocusAge
                Exit Sub
            End If
        End If
        '*** END BRGO ***********************************************
        ObtieneDatosCuenta sCta
        'frmSegSepelioAfiliacion.Inicio sCta
    End If
End Sub

Private Sub TxtFecFin_GotFocus()
With txtFecFin
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub

Private Sub txtFecIni_GotFocus()
With txtFecIni
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFecFin.SetFocus
End If
End Sub

Private Sub txtNumMov_GotFocus()
With txtNumMov
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtNumMov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Function CargoAutomatico(psCuenta As String, pnCntPag As Integer, Optional ByVal pbImpDebitoAuto As Boolean = False) As Boolean 'JUEZ 20150224 Se agregó pbImpDebitoAuto
'JUEZ 20151229 Cambios nueva forma de cobrar comisión

    Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento  'JUEZ 20151229
    Dim sMensajeCola As String
    Dim bExito As Boolean
    Dim nMonto As Currency
    'JUEZ 20151229 *********************
    Dim psCuentaImp As String
    Dim bValidaSaldo As Boolean
    Dim nTipoPago As Integer
    Dim RSClientes As ADODB.Recordset
    Dim sPersCod As String
    Dim sPersNombre As String
    
    bExito = False
    psCuentaImp = psCuenta
    fsBoleta = ""
    Set RSClientes = grdCliente.GetRsNew()
    'END JUEZ **************************
    
    'If chkComision.value = 1 Then
    'If chkComision.value = 1 Or pbImpDebitoAuto Then 'JUEZ 20150224
        'If nProducto = gCapAhorros And nPersoneria <> gPersonaJurCFLCMAC Then 'Ahorros y que no sean CMACs
        'If nProducto = gCapAhorros And nPersoneria <> gPersonaJurCFLCMAC Then 'Ahorros y que no sean CMACs
        If nPersoneria <> gPersonaJurCFLCMAC Then 'Que no sean CMACs
            If nEstado <> gCapEstAnulada Then
                sMensajeCola = Chr(13) + Str(pnCntPag) + " Pagina(s)"
                'MsgBox "Se va a realizar el cargo automático por el extracto de " + sMensajeCola, vbInformation, "Aviso"
                'If Not pbImpDebitoAuto Then MsgBox "Se va a realizar el cargo automático por el extracto de " + sMensajeCola, vbInformation, "Aviso" 'JUEZ 20150224
                
                If nmoneda = gMonedaNacional Then
                    'If (chkComision.value = 0) Then
                    '    nMonto = GetMontoDescuento(gDctoExtMN)
                    'Else
                        'nMonto = GetMontoDescuento(gDctoExtMNxPag, pnCntPag)
                        nMonto = GetMontoDescuento(gDctoExtMNxPag, pnCntPag, psCuenta) 'APRI20190109 ERS077-2018
                    'End If
                Else
                    'if (chkComision.value = 0) Then
                    '    nMonto = GetMontoDescuento(gDctoExtME)
                    'Else
                        'MAVM 20111110 ***
                        'nMonto = GetMontoDescuento(gDctoExtMExPag, pnCntPag)
                        'COMENTADO POR APRI20190109
                        'Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp
                        'Dim rs As ADODB.Recordset
                        'Set ObjTc = New COMDConstSistema.DCOMTipoCambioEsp
                        'Set rs = New ADODB.Recordset

                        'Set rs = ObjTc.GetTipoCambioCV(CCur(0))
                        'If Not (rs.EOF And rs.BOF) Then
                        '    Do Until rs.EOF
                        '        nMonto = Round(GetMontoDescuento(gDctoExtMNxPag, pnCntPag) / rs!nVenta, 2)
                        '        Exit Do
                        '    rs.MoveNext
                        '    Loop
                        'Set rs = Nothing
                        'Set ObjTc = Nothing
                        'End If
                        '***
                    'End If
                    nMonto = Round(GetMontoDescuento(gDctoExtMNxPag, pnCntPag, psCuenta), 2) 'APRI20190109 ERS077-2018
                End If
                If nMonto > 0 Then
                    Dim sMovNro As String
                    Dim oMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
                    Dim nFlag As Double, nITF As Currency
                    
                    Dim sMensaje As String
                    
                    nITF = fgITFCalculaImpuesto(nMonto)
                    
                    'JUEZ 20151229 *************************************************************
                    Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    If Not pbImpDebitoAuto Then
                        If Mid(psCuentaImp, 6, 3) = gCapPlazoFijo Then
                            bValidaSaldo = False
                        Else
                            bValidaSaldo = oCap.ValidaSaldoCuenta(psCuentaImp, nMonto + nITF)
                        End If
                        
                        If bValidaSaldo Then
                             If MsgBox("Esta opción generará un cobro de comisión de " & Format(nMonto / pnCntPag, "#,##0.00") & " por página, ¿Desea Continuar?", vbYesNo, "Aviso") = vbNo Then Exit Function
                        Else
                            MsgBox "Esta opción generará un cobro de comisión de " & Format(nMonto / pnCntPag, "#,##0.00") & " por página, por favor elija la forma de pago", vbInformation, "Aviso"
                            psCuenta = frmOpeComisionConsultaMovs.inicia(1, nMonto, nmoneda, RSClientes, nTipoPago, sPersCod, sPersNombre)
                            If nTipoPago = 0 Then Exit Function
                            If nTipoPago = 2 Then If psCuenta = "" Then Exit Function
                        End If
                    End If
                    'END JUEZ ******************************************************************
                    
                    Set oMov = New COMNContabilidad.NCOMContFunciones
                    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set oMov = Nothing
                    
                    oCap.IniciaImpresora gImpresora
                    'JUEZ 20150224 ********************************************
                    If pbImpDebitoAuto Then
                        If Not oCap.ValidaComisionExtractoServCobDebitoAuto(psCuentaImp, gdFecSis) Then
                            nMonto = 0
                            MsgBox "Este extracto de Débito Automático no realizará cargo de comisión por ser la primera vez en el mes", vbInformation, "Aviso"
                        Else
                            MsgBox "Se va a realizar el cargo automático por el extracto de Débito automático de " + sMensajeCola, vbInformation, "Aviso"
                        End If
                        nFlag = oCap.CapCargoCuentaAho(psCuenta, nMonto, gAhoDctoEmiExtDebitoAuto, sMovNro, "Descuento Emisión Extracto Debito Automático", , , , , , , , , gsNomAge, sLpt, , , , , gsCodCMAC, , gsCodAge, , False, nITF, , , , , sMensaje, fsBoleta, sBoletaITF, , , gbImpTMU)
                    Else
                        If nTipoPago = 1 Then
                            nFlag = oCap.OtrasOperaciones(sMovNro, "300439", nMonto, "", "Comisión Diversa : Descuento Emisión Extracto Cuenta " & psCuentaImp, nmoneda, sPersCod)  'JUEZ 20151019
                            Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
                            Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
                                fsBoleta = oBol.ImprimeBoleta("COMISIONES DIVERSAS", Left("Descuento Emisión Extracto", 36), "", Str(nMonto), sPersNombre, "________1", "", 0, "0", "", 0, 0, False, False, , , , False, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, 0)
                            Set oBol = Nothing
                        Else
                            If Mid(psCuenta, 6, 3) = gCapAhorros Then
                                nFlag = oCap.CapCargoCuentaAho(psCuenta, nMonto, gAhoDctoEmiExt, sMovNro, "Descuento Emisión Extracto Cuenta " & psCuentaImp, , , , , , , , , gsNomAge, sLpt, , , , , gsCodCMAC, , gsCodAge, , False, nITF, , , , , sMensaje, fsBoleta, sBoletaITF, , , gbImpTMU)
                            ElseIf Mid(psCuenta, 6, 3) = gCapCTS Then
                                nFlag = oCap.CapCargoCuentaCTS(psCuenta, nMonto, gCTSDctoEmiExt, sMovNro, "Descuento Emisión Extracto Cuenta " & psCuentaImp, , , , , gsCodCMAC, , gsNomAge, sLpt, , , , , Mid(psCuenta, 9, 1), , , , , , sMensaje, fsBoleta, gbImpTMU)
                            End If
                        End If
                    End If
                    'END JUEZ *************************************************
                    '***Se parametro false por ELRO el 20130322, según SATI INC1302130019
                    Set oCap = Nothing
                    
                    If nFlag = 0 Then
                        MsgBox "Cuenta No Posee saldo suficiente para el descuento.", vbInformation, "Aviso"
                        bExito = False
                    Else
                        bExito = True
                        'If nMonto > 0 Then
                        If nMonto > 0 And pbImpDebitoAuto Then
                            If sMensaje <> "" Then MsgBox sMensaje, vbInformation, "Aviso"
                            If fsBoleta <> "" Then ImprimeBoleta fsBoleta
                            If sBoletaITF <> "" Then ImprimeBoleta sBoletaITF, "Boleta ITF"
                        End If
                    End If
                Else
                    bExito = True
                End If
            Else
                bExito = True
            End If
        Else
            bExito = True
        End If
        
        CargoAutomatico = bExito
    'Else
    '    CargoAutomatico = bExito
    'End If
    psCuenta = psCuentaImp 'JUEZ 20151229
End Function






