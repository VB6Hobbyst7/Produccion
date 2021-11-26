VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaChicaRendicion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendicion de Caja Chica"
   ClientHeight    =   10125
   ClientLeft      =   1200
   ClientTop       =   1020
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "frmCajaChicaRendicion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRendicion 
      Caption         =   "&Confirmar Rend"
      Height          =   390
      Left            =   6735
      TabIndex        =   40
      Top             =   9600
      Width           =   1320
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   330
      Left            =   8265
      TabIndex        =   44
      Top             =   60
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdExtorn 
      Caption         =   "&Extorno Rend"
      Height          =   390
      Left            =   5325
      TabIndex        =   42
      Top             =   9600
      Width           =   1320
   End
   Begin VB.CommandButton cmdRecibo 
      Caption         =   "Duplicado &Recibo"
      Height          =   390
      Left            =   3525
      TabIndex        =   41
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8055
      TabIndex        =   39
      Top             =   9600
      Width           =   1320
   End
   Begin VB.CommandButton cmdAutorizacion 
      Caption         =   "&Autorización"
      Height          =   390
      Left            =   6720
      TabIndex        =   38
      Top             =   9600
      Width           =   1320
   End
   Begin VB.Frame fraAutorizacion 
      Caption         =   "Autorización"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3525
      TabIndex        =   29
      Top             =   1875
      Width           =   5850
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Monto Autorización :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   2430
         TabIndex        =   33
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label lblImporteAuto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   4275
         TabIndex        =   32
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label lblNroProcChAct 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1680
         TabIndex        =   31
         Top             =   195
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Caja Chica Actual :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   135
         TabIndex        =   30
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "Asiento &Contable"
      Height          =   390
      Left            =   1845
      TabIndex        =   14
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlanilla 
      Caption         =   "&Planilla Rendición"
      Height          =   390
      Left            =   165
      TabIndex        =   13
      Top             =   9600
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   12991
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   617
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Egresos Directos"
      TabPicture(0)   =   "frmCajaChicaRendicion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Entrega de Efectivo"
      TabPicture(1)   =   "frmCajaChicaRendicion.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FEEgresos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame5 
         Caption         =   "Comprobantes de Pago Detalle"
         Height          =   1455
         Left            =   120
         TabIndex        =   62
         Top             =   4440
         Width           =   9045
         Begin Sicmact.FlexEdit FEDetDocSust 
            Height          =   1095
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   1931
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-Código-Descripción-Monto"
            EncabezadosAnchos=   "300-1200-5000-1200"
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-R"
            FormatosEdit    =   "0-0-0-2"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Comprobantes de Pago - General"
         Height          =   1815
         Left            =   120
         TabIndex        =   60
         Top             =   2520
         Width           =   9045
         Begin Sicmact.FlexEdit fgSust 
            Height          =   1425
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   8760
            _ExtentX        =   15452
            _ExtentY        =   2514
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Nro-Tipo-Nro Doc-Fecha-Proveedor-Importe-cMovSust-cMovAtenc-Concepto"
            EncabezadosAnchos=   "400-500-1300-1000-3000-1200-0-0-0"
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-L-R-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-2-2-2-0"
            TextArray0      =   "Nro"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbOrdenaCol     =   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame FEEgresos 
         Caption         =   "Egresos"
         Height          =   1335
         Left            =   120
         TabIndex        =   46
         Top             =   5880
         Width           =   8775
         Begin VB.Frame FraSaldoArendir 
            Caption         =   "Saldo Rendido a Caja :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   630
            Left            =   120
            TabIndex        =   47
            Top             =   585
            Width           =   3495
            Begin VB.Label lblTotalArendRend 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2430
               TabIndex        =   51
               Top             =   210
               Width           =   990
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Importe :"
               Height          =   210
               Left            =   1770
               TabIndex        =   50
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lblFechaArendRend 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   660
               TabIndex        =   49
               Top             =   195
               Width           =   990
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Fecha :"
               Height          =   210
               Left            =   105
               TabIndex        =   48
               Top             =   255
               Width           =   540
            End
         End
         Begin VB.Label lblTotalPend 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   315
            Left            =   1650
            TabIndex        =   59
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Pendiente : "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   210
            Left            =   180
            TabIndex        =   58
            Top             =   270
            Width           =   1425
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Egresos Atendidos :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   4590
            TabIndex        =   57
            Top             =   315
            Width           =   2130
         End
         Begin VB.Label lblTotalEgresosArendir 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   7140
            TabIndex        =   56
            Top             =   270
            Width           =   1500
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Ingresos por Rendicion :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   4590
            TabIndex        =   55
            Top             =   630
            Width           =   2475
         End
         Begin VB.Label lblTotalIngPorArendir 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   7140
            TabIndex        =   54
            Top             =   585
            Width           =   1500
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Total Egresos por Rendicion :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   4590
            TabIndex        =   53
            Top             =   945
            Width           =   2415
         End
         Begin VB.Label lblTotalEgresPorArendir 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   300
            Left            =   7140
            TabIndex        =   52
            Top             =   900
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "A Rendir Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   150
         TabIndex        =   19
         Top             =   375
         Width           =   9045
         Begin Sicmact.FlexEdit fgArendir 
            Height          =   1680
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   2963
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Nro-Nro Doc-Solicitante-Fecha-Importe-Saldo-Concepto-cMovNroAtenc-cMovNroRend-nImporteRend-NroCajaCH"
            EncabezadosAnchos=   "400-1300-3000-1000-1200-1200-0-0-0-0-1000"
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-C-R-R-L-C-C-R-C"
            FormatosEdit    =   "0-0-0-0-2-2-0-0-0-2-0"
            TextArray0      =   "Nro"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbOrdenaCol     =   -1  'True
            ColWidth0       =   405
            RowHeight0      =   285
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6795
         Left            =   -74880
         TabIndex        =   15
         Top             =   405
         Width           =   9045
         Begin Sicmact.FlexEdit fgEgresos 
            Height          =   6045
            Left            =   150
            TabIndex        =   16
            Top             =   210
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   10663
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Nro-Tipo-Nro Doc-Fecha-Solicitante-Importe-Concepto-cMovNroAtenc-cDocTpo-cMovNrosol-nNroProc"
            EncabezadosAnchos=   "400-500-1300-1000-3500-1200-2500-0-0-0-0"
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-L-R-L-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0-0"
            TextArray0      =   "Nro"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbOrdenaCol     =   -1  'True
            ColWidth0       =   405
            RowHeight0      =   285
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblMontonoDesem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   2415
            TabIndex        =   35
            Top             =   6345
            Width           =   1500
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Monto no Desembolsado :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   225
            TabIndex        =   34
            Top             =   6390
            Width           =   2160
         End
         Begin VB.Label lblTotalEgresos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   7395
            TabIndex        =   18
            Top             =   6330
            Width           =   1500
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Total Egresos :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   6075
            TabIndex        =   17
            Top             =   6390
            Width           =   1230
         End
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1425
      Left            =   105
      TabIndex        =   0
      Top             =   330
      Width           =   9315
      Begin VB.ComboBox cboProceso 
         Height          =   330
         Left            =   8295
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   255
         Width           =   810
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   7275
         TabIndex        =   26
         Top             =   675
         Width           =   1530
         Begin VB.CheckBox chkRendida 
            Caption         =   "Rendida??"
            Height          =   240
            Left            =   150
            TabIndex        =   27
            Top             =   30
            Width           =   1275
         End
      End
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1155
         TabIndex        =   1
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbltotalRend 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   8040
         TabIndex        =   37
         Top             =   1035
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7260
         TabIndex        =   36
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsable :"
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   2565
         TabIndex        =   24
         Top             =   645
         Width           =   4545
      End
      Begin VB.Label lblPerscod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1155
         TabIndex        =   23
         Top             =   645
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Dif :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5715
         TabIndex        =   22
         Top             =   1110
         Width           =   300
      End
      Begin VB.Label lblDiferencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   6060
         TabIndex        =   21
         Top             =   1035
         Width           =   795
      End
      Begin VB.Label lblFechaApert 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1155
         TabIndex        =   11
         Top             =   1005
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Apertura/ Reembolso :"
         Height          =   405
         Left            =   90
         TabIndex        =   10
         Top             =   930
         Width           =   990
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4680
         TabIndex        =   9
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4110
         TabIndex        =   8
         Top             =   1095
         Width           =   540
      End
      Begin VB.Label lblImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   3105
         TabIndex        =   7
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Asignado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   2280
         TabIndex        =   6
         Top             =   1065
         Width           =   780
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7425
         TabIndex        =   5
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   210
         Left            =   7140
         TabIndex        =   4
         Top             =   330
         Width           =   255
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2580
         TabIndex        =   3
         Top             =   270
         Width           =   4305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   210
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   345
         Left            =   7080
         Top             =   1020
         Width           =   2190
      End
   End
   Begin VB.Label Label20 
      Caption         =   "Fecha :"
      Height          =   225
      Left            =   7230
      TabIndex        =   45
      Top             =   105
      Width           =   1005
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDFDFE&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   120
      TabIndex        =   43
      Top             =   1770
      Visible         =   0   'False
      Width           =   7785
   End
End
Attribute VB_Name = "frmCajaChicaRendicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipoProc As CHTipoProc
Dim oCH As nCajaChica
Dim lsCtaFondofijo As String
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsTituloFrame As String
Dim lsMovNroRend As String
Dim lbOk As Boolean
Dim lsCajaChicaNro As String
Dim fsOpeCodDes  As String '***ELro
Dim fsCtaFondofijoDes As String '***ELRO
Dim fsCtaFondoFijoDesF As String '***ELRO
Dim fsCtaDesembolso As String '****ELRO

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************


Public Sub Inicio(ByVal pnTipoProc As CHTipoProc, Optional psCajaChicaNro As String = "")
lsCajaChicaNro = psCajaChicaNro
lnTipoProc = pnTipoProc
Me.Show 1
End Sub
Private Sub cboProceso_Click()
If txtBuscarAreaCH <> "" Then
    lblNroProcCH = cboProceso
    CargaDatos
End If
End Sub


Private Sub cmdAsiento_Click()
Dim lsTexto As String
Dim oImp As NContImprimir
If ValidaInterfaz = False Then Exit Sub

'''gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", "S/.", "$.") 'marg ers044-2016
gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", gcPEN_SIMBOLO, "$.") 'marg ers044-2016
Set oImp = New NContImprimir
lsTexto = oImp.ImprimeAsientoRendicionCH(Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(Me.lblNroProcCH), _
            gsSimbolo, gnColPage, gnLinPage, gdFecSis, gsNomCmac, gsOpeCod, "", IIf(Me.chkRendida.value = 1, True, False), lsCtaFondofijo)
            
EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
Set oImp = Nothing
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo Asiento Contable de Chica de la  Agencia/Area : " & lblCajaChicaDesc & ""
            Set objPista = Nothing
            '*******
End Sub

Private Sub cmdAutorizacion_Click()
Dim lsMovNro As String
Dim lsTexto As String
Dim oConImp As NContImprimir
Dim oContFunc As NContFunciones
'***Agregado por ELRO el 20120618, según OYP-RFC047-2012
Dim oDOperacion As New DOperacion
Dim rsAut As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim lnMovNroAut As Long
Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim lsFechaDoc As String
Dim lsDocumento As String
Dim lsGlosa As String
Dim lsMovNroAut As String
Dim lsCadBol As String
Dim lnImporte As Currency
Dim lsCtaContDebeITF As String
Dim lsCtaContHaberITF As String
Dim lsMontoITF As Double
Dim lnMovNroApro As Long
Dim lsSubCta As String
'***Fin Agregado por ELRO*******************************
Dim lnTotalEgresos As Currency '***Agregado por ELRO el 20120828, según OYP-RFC104-2012

Set oContFunc = New NContFunciones
Set oConImp = New NContImprimir

If ValidaInterfaz(True) = False Then Exit Sub
If MsgBox("Desea Realizar la respectiva Autorización de Caja Chica ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    cmdAutorizacion.Enabled = False '***Agregado por ELRO, según SATI INC1301120002
    cmdSalir.Enabled = False '***Agregado por ELRO el 20130220, según SATI INC1302190017
    
    lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    '***Agregado por ELRO el 20120828, según OYP-RFC104-2012
    lnTotalEgresos = oCH.devolverTotalEgresos(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), CInt(lblNroProcChAct))
    '***Fin Agregado por ELRO el 20120828*******************
    
    If oCH.GrabaAutorizacionCh(lsMovNro, lsMovNroRend, gsFormatoFecha, gsOpeCod, Me.Caption, CCur(lblImporteAuto), CCur(lblSaldo), _
                            Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcChAct), lnMovNroAut) = 0 Then
                            '***Agregado por ELRO el parametro lnMovNroAut el 20120618
        
        '***Agregado por ELRO el 20120618, según OYP-RFC047-2012
        If lnMovNroAut > 0 Then
            lsMovNroAut = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            fsCtaDesembolso = oDOperacion.EmiteOpeCta(fsOpeCodDes, "H")
              
            If fsOpeCodDes = CStr(gCHApropacionApeMN) Or fsOpeCodDes = CStr(gCHApropacionApeME) Then
            '***Modificado por ELRO el 20130222, según SATI INC1301300007
            '    If CInt(lblNroProcCH + 1) > 1 Then
            '        fsCtaFondofijoDes = oDOperacion.EmiteOpeCta(fsOpeCodDes, "D", "1")
            '    Else
            '        fsCtaFondofijoDes = oDOperacion.EmiteOpeCta(fsOpeCodDes, "D")
            '    End If
            'Else
            '    fsCtaFondofijoDes = oDOperacion.EmiteOpeCta(fsOpeCodDes, "D")
            fsCtaFondofijoDes = oDOperacion.EmiteOpeCta(fsOpeCodDes, "D", "1")
            '***Fin Modificado por ELRO el 20130222**********************
            End If
            
            lsNroDoc = ""
            lsNroVoucher = ""
            lsFechaDoc = ""
            lsDocumento = ""
            lsGlosa = ""
            lnImporte = CCur(lblImporteAuto)
            Set rsAut = oCH.devolverDatosCajaChicaSinAprobar_2(lnMovNroAut, Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2))
            
           '***Cometado por ELRO el 20130222, según SATI INC1301300007
           'lsSubCta = oContFunc.GetFiltroObjetos(ObjCMACAgenciaArea, fsCtaFondofijoDes, Trim(txtBuscarAreaCH), False)
        
           'If lsSubCta <> "" Then
           '   If Mid(txtBuscarAreaCH, 1, 3) = "042" Then 'Recuperaciones
           '         fsCtaFondoFijoDesF = fsCtaFondofijoDes & lsSubCta
           '
           '   ElseIf Mid(txtBuscarAreaCH, 1, 3) = "043" Then  'Secretaria
           '         fsCtaFondoFijoDesF = fsCtaFondofijoDes & lsSubCta
                    
           '   ElseIf Mid(txtBuscarAreaCH, 1, 3) = "023" Then  'Logistica
           '         fsCtaFondoFijoDesF = fsCtaFondofijoDes & lsSubCta
           '   Else
           '         fsCtaFondoFijoDesF = fsCtaFondofijoDes & lsSubCta
           '   End If
           '
           'Else
           '   '***Modificado por ELRO el 20120927, según OYP-RFC111-2012
           '   'MsgBox "Sub Cuenta Contable " & lsSubCta & " no definida.", vbInformation, "Aviso"
           '   oCH.extornarAutorizacionCajaChica (lnMovNroAut)
           '   MsgBox "Sub Cuenta Contable " & lsSubCta & " de la Cuenta Contable " & fsCtaFondofijoDes & " no definida.", vbInformation, "¡Aviso!"
           '   '***Fin Modificado por ELRO el 20120927*******************
           '   Exit Sub
           'End If
           '***Cometado por ELRO el 20130222***************************
           lsCtaContDebeITF = oDOperacion.EmiteOpeCta(fsOpeCodDes, "D", 2)
           lsCtaContHaberITF = oDOperacion.EmiteOpeCta(fsOpeCodDes, "H", 2)
           
           lsMontoITF = fgTruncar(lnImporte * gnImpITF, 2)
           '***Agregado por ELRO el 20130220, según SATI INC1302190017
           lsGlosa = "Aprobación Reembolso Caja Chica " & lblCajaChicaDesc & " para el proceso " & lblNroProcChAct
           '***Fin Agregado por ELRO el 20130220**********************
           Call oCH.GrabaDesembolsoCH(lsMovNroAut, lnMovNroAut, _
                                      lblPerscod, _
                                      gsFormatoFecha, rs, fsOpeCodDes, _
                                      lsGlosa, _
                                      lnImporte, _
                                      Mid(Me.txtBuscarAreaCH, 1, 3), _
                                      Mid(Me.txtBuscarAreaCH, 4, 2), _
                                      Val(lblNroProcCH) + 1, "", _
                                      lsNroDoc, lsFechaDoc, _
                                      lsNroVoucher, rsAut, _
                                      "", 0, fsCtaFondofijoDes, _
                                      fsCtaDesembolso, CCur(lblSaldo), "", _
                                      gbBitCentral, , lsCtaContDebeITF, _
                                      lsCtaContHaberITF, lsMontoITF, lnMovNroApro, lnTotalEgresos)
            '***Parametro fsCtaFondofijoDes agregado por ELRO el 20130222, según SATI INC1301300007
            '***Agregado por ELRO el parametro lnSaldoActual, según OYP-RFC104-2012
            'lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, gsOpeDesc, gsNomCmac, gsOpeCod, _
            '                    lblPerscod, CCur(lblImporteAuto), gnColPage, , , False, True, txtBuscarAreaCH & "-" & Val(lblNroProcCH), _
            '                    lblCajaChicaDesc, , , , True)
            '    EnviaPrevio lsTexto, Me.Caption, gnLinPage
            If lnMovNroApro = 0 Then
                '***Agregado por ELRO el 20130220, según SATI INC1302190017
                oCH.extornarAutorizacionCajaChica (lnMovNroAut)
                '***Fin Agregado por ELRO el 20130220**********************
                MsgBox "No se realizó la operación", vbCritical, "Aviso"
                Unload Me
                Exit Sub
            End If
            '***Fin Agregado por ELRO*******************************
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Autorizo Reembolso de Chica del  Area : " & lblCajaChicaDesc & " |Encargado : " & lblPersNombre _
            & " |Monto : " & lblImporteAuto
            Set objPista = Nothing
            '*******
        End If
        
        If MsgBox("Desea Realizar otra Autorización de Caja Chica", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            Unload Me
        End If
        Me.txtBuscarAreaCH = ""
        Me.lblCajaChicaDesc = ""
        Me.lblNroProcCH = ""
        lblNroProcChAct = ""
        Limpia
        cboProceso.Clear
        '***Agregado por ELRO el 20121004, según OYP-RFC111-2012
        fsCtaFondoFijoDesF = ""
        fsCtaFondofijoDes = ""
        '***Fin Agregado por ELRO el 20121004*******************
        cmdAutorizacion.Enabled = True '***Agregado por ELRO, según SATI INC1301120002
        cmdSalir.Enabled = True '***Agregado por ELRO el 20130220, según SATI INC1302190017
    End If
End If
Set oConImp = Nothing

End Sub

Private Sub cmdExtorn_Click()
If ValidaInterfaz(True) = False Then Exit Sub

'***Agregado por ELRO el 20120808, según OYP-RFC047-2012
If gsCodCargo <> "006006" Then
    MsgBox "Solo el(la) Asistente de Contabilidad puede realizar el extorno.", vbInformation, "Aviso"
    Exit Sub
End If
'***Fin Agregado por ELRO*******************************

If MsgBox("Desea Realizar el Extorno de la Rendicion a contabilidad de la Caja chica N° [" & Val(lblNroProcCH) & "] ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If oCH.GrabaRendicionChExt(Mid(txtBuscarAreaCH, 4, 2), Left(txtBuscarAreaCH, 3), CInt(lblNroProcCH) + 1) = 0 Then
       MsgBox "Extorno grabado con exito.", vbInformation, "Aviso"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Se Realizo Extorno de Chica de la  Agencia/Area : " & lblCajaChicaDesc
            Set objPista = Nothing
            '*******
       Unload Me
    End If
End If
End Sub

Private Sub cmdPlanilla_Click()
Dim lsTexto As String
Dim oConImp As NContImprimir
Set oConImp = New NContImprimir
If ValidaInterfaz = False Then Exit Sub

'''gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", "S/.", "$.") 'marg ers044-2016
gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", gcPEN_SIMBOLO, "$.") 'marg ers044-2016
lsTexto = oConImp.GetImprimeDocRendicionCH(gnLinPage, gsNomCmac, gdFecSis, gnColPage, IIf(Me.chkRendida.value = 1, True, False), Trim(txtBuscarAreaCH), _
                                           Val(lblNroProcCH), lblCajaChicaDesc, lblFechaApert, _
                                           CCur(lblImporte), lblPerscod, lblPersNombre, lsCtaArendir, lsCtaPendiente, lsCtaFondofijo, _
                                           gsSimbolo, CCur(lblSaldo), fgArendir.GetRsNew, fgEgresos.GetRsNew)

EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
Set oConImp = Nothing
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Realizo Planilla Rendicion de Chica de la  Agencia/Area : " & lblCajaChicaDesc
            Set objPista = Nothing
            '*******
End Sub

Private Sub cmdRecibo_Click()
If Val(Me.fgEgresos.TextMatrix(fgEgresos.row, 8)) = TpoDocRecEgreso Then
    Dim lsRec As String
    Dim oImp As NContImprimir
    Dim oMov As New DMov
    Set oImp = New NContImprimir
    oImp.Inicio gsNomCmac, gsNomAge, Format(gdFecSis, gsFormatoFechaView)
    lsRec = oImp.ImprimeRecibo(oMov.GetcMovNro(fgEgresos.TextMatrix(fgEgresos.row, 7)))
    Set oImp = Nothing
    EnviaPrevio lsRec, "IMPRESION DE DOCUMENTO", gnLinPage, False
End If
End Sub

Private Sub cmdRendicion_Click()
Dim lsMovNro As String
Dim lsTexto As String
Dim oConImp As NContImprimir
Dim oContFunc As NContFunciones
Dim lnTotalRendicion As Currency
Dim lsGlosa As String '***Agregado por ELRO el 20130221, según INC1301300007

'*** PEAC 20100908
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date
'*** FIN PEAC

Set oContFunc = New NContFunciones
Set oConImp = New NContImprimir

If chkRendida Then Exit Sub '***Agregado por ELRO el 20130221, según INC1301300007

If ValidaInterfaz(True) = False Then Exit Sub

If Val(lblTotalPend) <> 0 Then
    MsgBox "Existe al menos una Entrega de Efectivo sin rendir.", vbInformation, "Aviso"
    Exit Sub
End If


'*** PEAC 20100908
If gsOpeCod = "401380" Then
    nMes = Month(CDate(Me.mskFecha.Text))
    nAnio = Year(CDate(Me.mskFecha.Text))
    dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(nAnio, "0000")) - 1
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible realizar la rendición a Contabilidad ya que la fecha ingresada pertenece a un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If
End If
'*** FIN PEAC


If MsgBox("Desea Realizar la rendicion a contabilida de la Caja chica N° [" & Val(lblNroProcCH) & "] ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    cmdRendicion.Enabled = False '***Agregado por ELRO, según SATI INC1301120002
    lsMovNro = oContFunc.GeneraMovNro(CDate(Me.mskFecha.Text), gsCodAge, gsCodUser)
    'lnTotalRendicion = CCur(lblImporte) - (CCur(lblSaldo) + CCur(lblMontonoDesem))
    lnTotalRendicion = CCur(lbltotalRend)
    lsGlosa = UCase("Confirmación de Rendición de la Caja Chica " & lblCajaChicaDesc & " para el proceso " & lblNroProcCH) '***Agregado por ELRO el 20130221, según INC1301300007
    
    If oCH.GrabaRendicionCh(lsMovNro, gsFormatoFecha, gsOpeCod, lsGlosa, lnTotalRendicion, CCur(lblSaldo), _
                            Mid(txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), CDate(Me.mskFecha.Text), gsCodUser, gsCodAge) = 0 Then
        '***Paranetro lsGlosa agregado por ELRO el 20130221, según INC1301300007
            CargaDatosCajaChica Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
            '''gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", "S/.", "$.") 'marg ers044-2016
            gsSimbolo = IIf(Mid(gsOpeCod, 3, 1) = "1", gcPEN_SIMBOLO, "$.") 'marg ers044-2016
            lsTexto = oConImp.ImprimeAsientoRendicionCH(Mid(Me.txtBuscarAreaCH, 1, 3), Mid(Me.txtBuscarAreaCH, 4, 2), Val(Me.lblNroProcCH), _
                        gsSimbolo, gnColPage, gnLinPage, CDate(Me.mskFecha.Text), gsNomCmac, gsOpeCod, "", True, lsCtaFondofijo)
                    
            lsTexto = lsTexto + oImpresora.gPrnSaltoPagina + oConImp.GetImprimeDocRendicionCH(gnLinPage, gsNomCmac, CDate(Me.mskFecha.Text), gnColPage, True, Trim(txtBuscarAreaCH), _
                                                       Val(lblNroProcCH), lblCajaChicaDesc, lblFechaApert, _
                                                       CCur(lblImporte), lblPerscod, lblPersNombre, lsCtaArendir, lsCtaPendiente, lsCtaFondofijo, _
                                                       gsSimbolo, CCur(lblSaldo), fgArendir.GetRsNew, fgEgresos.GetRsNew)
            
            EnviaPrevio lsTexto, Me.Caption, gnLinPage, False
            cmdRendicion.Enabled = False
            chkRendida.value = 1
            If lnTipoProc = gCHTipoProcArqueo Then
                lbOk = True
                Unload Me
            End If
        cmdRendicion.Enabled = True '***Agregado por ELRO, según SATI INC1301120002
    End If
End If
Set oConImp = Nothing
End Sub
Function ValidaInterfaz(Optional pbGrabar As Boolean = False) As Boolean
Dim lsFecha As String
Dim n As Integer

ValidaInterfaz = True
If Len(Trim(txtBuscarAreaCH.Text)) = 0 Then
    MsgBox "Ingrese la Caja Chica a Rendir", vbInformation, "Aviso"
    ValidaInterfaz = False
    Exit Function
End If
lsFecha = ValidaFecha(lblFechaApert)
If lsFecha <> "" Then
    MsgBox lsFecha & vbCrLf & "Fecha de Apertura o Reembolso no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    Exit Function
End If
If Me.fgEgresos.TextMatrix(1, 0) = "" Then
    If MsgBox("Caja Chica No Contiene Egresos directos. Desea Continuar ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        ValidaInterfaz = False
        Exit Function
    End If
End If
If fgArendir.TextMatrix(1, 0) = "" Then
    If MsgBox("Caja Chica No Contiene A Rendir Cuentas Solicitados. Desea Continuar ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        ValidaInterfaz = False
        Exit Function
    End If
End If

If mskFecha.Text = "__/__/____" Then
   MsgBox "Ingrese Fecha de Rendición", vbInformation, "AVISO"
   ValidaInterfaz = False
   Exit Function
End If
If pbGrabar Then
    If Val(lblTotalPend) <> 0 Then
        gnImporte = 0
        For n = 1 To fgArendir.Rows - 1
            If nVal(fgArendir.TextMatrix(n, 5)) <> 0 Then
               gnImporte = Val(fgArendir.TextMatrix(n, 4)) - (fgArendir.TextMatrix(n, 5))
            End If
        Next
        If gnImporte <> 0 Then
            MsgBox "Caja Chica posee Arendir en Proceso de Sustentación, aún no realiza la rendición correspondiente"
            ValidaInterfaz = False
            Exit Function
        End If
    End If
    If Val(lblDiferencia) <> 0 Then
        If MsgBox("Caja Chica posee diferencias. Desea Continuar? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
           ValidaInterfaz = False
           Exit Function
        End If
    End If
End If

'***Agregado por ELRO el 2012027, según TIC1209240007
If DateDiff("M", mskFecha, gdFecSis) <> 0 Then
    MsgBox "Debe ingresar una fecha en el mes y año vigente.", vbInformation, "¡Aviso!"
    mskFecha.SetFocus
End If
'***Fin Agregado por ELRO el 2012027*****************

End Function
Private Sub cmdSalir_Click()
lbOk = False
Unload Me
End Sub
Private Sub fgArendir_EnterCell()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
'Exacta Rend.MontoRendido = 0
'ConIngreso Rend.MontoRendido < 0
'ConEgreso and Rend.MontoRendido>0
If fgArendir.TextMatrix(1, 0) <> "" Then
    If fgArendir.TextMatrix(fgArendir.row, 8) <> "" Then
        lblFechaArendRend = Mid(fgArendir.TextMatrix(fgArendir.row, 8), 7, 2) & "/" & Mid(Me.fgArendir.TextMatrix(fgArendir.row, 8), 5, 2) & "/" & Mid(Me.fgArendir.TextMatrix(fgArendir.row, 8), 1, 4)
        lblTotalArendRend = Format(Abs(fgArendir.TextMatrix(fgArendir.row, 9)), "#,#0.00")
        Select Case Val(fgArendir.TextMatrix(fgArendir.row, 9))
            Case Is = 0
                FraSaldoArendir = lsTituloFrame & " - Exacta"
                Me.lblTotalArendRend.ForeColor = vbBlack
            Case Is < 0
                FraSaldoArendir = lsTituloFrame & " - Ingreso Caja"
                Me.lblTotalArendRend.ForeColor = vbBlue
            Case Is > 0
                FraSaldoArendir = lsTituloFrame & " - Egreso de Caja"
                Me.lblTotalArendRend.ForeColor = vbRed
        End Select
    Else
        FraSaldoArendir = lsTituloFrame & " - Sin Rendición"
        Me.lblTotalArendRend.ForeColor = vbBlack
        lblFechaArendRend = ""
        lblTotalArendRend = "0.00"
    End If
    Set rs = oCH.GetCHSustAtenArendir(lsCtaArendir, lsCtaPendiente, fgArendir.TextMatrix(fgArendir.row, 7))
    If Not rs.EOF And Not rs.BOF Then
        Set fgSust.Recordset = rs
        fgSust.FormatoPersNom 4
    Else
        fgSust.Clear
        fgSust.FormaCabecera
        fgSust.Rows = 2
    End If
    rs.Close
    Set rs = Nothing
    Call fgSust_EnterCell
End If
End Sub



Private Sub fgSust_EnterCell()
    Call LimpiaFlex(FEDetDocSust)
    
    If fgSust.TextMatrix(1, 0) <> "" Then
        Dim oNArendir As NARendir
        Set oNArendir = New NARendir
        Dim rsDetDocSus As ADODB.Recordset
        Set rsDetDocSus = New ADODB.Recordset
        Dim oDMov As New DMov
        Dim lnMovNroSus As Long
        Dim lnFila As Integer
   
        lnMovNroSus = oDMov.GetnMovNro(fgSust.TextMatrix(fgSust.row, 6))
        
        Set rsDetDocSus = oNArendir.devolverDetalleDocSustentariosArendir(lnMovNroSus, lsCtaArendir)
        
        Do While Not rsDetDocSus.EOF
            FEDetDocSust.AdicionaFila
            lnFila = FEDetDocSust.row
               
            FEDetDocSust.TextMatrix(lnFila, 1) = rsDetDocSus!cCtaContCod
            FEDetDocSust.TextMatrix(lnFila, 2) = rsDetDocSus!cMovDesc
            FEDetDocSust.TextMatrix(lnFila, 3) = Format(rsDetDocSus!nMovImporte, "#,#0.00")

            rsDetDocSus.MoveNext
        Loop
        rsDetDocSus.Close: Set rsDetDocSus = Nothing
        Set oNArendir = Nothing
    End If
End Sub



Private Sub Form_Load()
Dim oArendir As NARendir
Dim oOpe As DOperacion
Set oOpe = New DOperacion
Set oArendir = New NARendir
Set oCH = New nCajaChica
Me.Caption = gsOpeDesc

CentraForm Me
SSTab1.Tab = 0
cmdRendicion.Enabled = False
chkRendida.value = 0
cmdRendicion.Visible = False
cmdAutorizacion.Visible = False
fraAutorizacion.Visible = False
lsTituloFrame = "Saldo a Rendir a Caja  "
FraSaldoArendir = lsTituloFrame

'*** PEAC 20100129
If gsOpeCod = "401380" Or gsOpeCod = "401396" Or gsOpeCod = "401310" Then
    lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
    'lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
Else
    lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
End If
'*** FIN PEAC

'***Agregado por ELRO el 20120618, según OYP-RFC047-2012
fsOpeCodDes = IIf(Mid(gsOpeCod, 3, 1) = "1", CStr(gCHApropacionApeMN), CStr(gCHApropacionApeME))

'***Fin Agregado por ELRO*******************************

lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
lsCtaPendiente = oOpe.EmiteOpeCta(gsOpeCod, "D", "2")
txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
Select Case lnTipoProc
    Case gCHTipoProcRendicion
        cmdRendicion.Visible = True
        cmdExtorn.Visible = False
    Case gCHTipoProcHabilitacion
        cmdAutorizacion.Visible = True
        fraAutorizacion.Visible = True
        cmdExtorn.Visible = False
    Case gCHTipoProcArqueo
        txtBuscarAreaCH = lsCajaChicaNro
        txtBuscarAreaCH.Enabled = False
        cmdRendicion.Visible = True
        txtBuscarAreaCH_EmiteDatos
End Select

'***Agregado por ELRO el 20120625, según OYP-RFC-047-2012
If gsCodArea <> "021" And (gsOpeCod = CStr(gCHRendContabMN) Or gsOpeCod = CStr(gCHRendContabME)) Then
    cmdRendicion.Visible = False
End If
'***Fin Agregado por ELRO********************************
Set oArendir = Nothing
Set oOpe = Nothing
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 20
End Sub

Private Sub mskFecha_LostFocus()
    If Not IsDate(mskFecha) Then
        mskFecha.SetFocus
    '***Agregado por ELRO el 2012027, según TIC1209240007
    Else
        If DateDiff("M", mskFecha, gdFecSis) <> 0 Then
         MsgBox "Debe ingresar una fecha en el mes y año vigente.", vbInformation, "¡Aviso!"
         mskFecha.SetFocus
        End If
    '***Fin Agregado por ELRO el 2012027*****************
    End If
    
End Sub

Private Sub txtBuscarAreaCH_EmiteDatos()
If lnTipoProc = gCHTipoProcRendicion Or lnTipoProc = gCHTipoProcArqueo Then
    lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
    lblNroProcCH = oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
    CargaCombo cboProceso, oCH.GetCHRendidas(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
    cboProceso.AddItem lblNroProcCH
    CargaDatos
Else
    lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
    CargaCombo cboProceso, oCH.GetCHRendSinAutorizacion(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
    cboProceso.ListIndex = cboProceso.ListCount - 1
    lblNroProcChAct = oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
    If cboProceso.ListCount = 0 Then
        MsgBox "Caja chica no ha realizado ninguna Rendicion", vbInformation, "Aviso"
        Me.txtBuscarAreaCH = ""
        Me.lblCajaChicaDesc = ""
        Me.lblNroProcCH = ""
        lblNroProcChAct = ""
        Limpia
    End If
    
End If
End Sub
Public Sub CargaDatos()
Me.MousePointer = 11
Limpia
lsMovNroRend = oCH.GetMovCHProceso(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), gCHTipoProcRendicion)
If lsMovNroRend <> "" Then
    If lnTipoProc = gCHTipoProcRendicion Then
        If MsgBox("Caja Chica ya ha sido Rendida a Contabilidad." & Chr(13) & "Desea Continuar ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            txtBuscarAreaCH = ""
            lblNroProcCH = ""
            lblCajaChicaDesc = ""
            Me.MousePointer = 0
            Exit Sub
        Else
            cmdRendicion.Enabled = False
            chkRendida.value = 1
        End If
    Else
            cmdRendicion.Enabled = False
            chkRendida.value = 1
    End If
Else
    cmdRendicion.Enabled = True
    chkRendida.value = 0
End If
CargaDatosCajaChica Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
'***Modificado por ELRO el 20130222, según SATI INC1301300007
'SSTab1.Tab = 0
SSTab1.Tab = 1
'***Fin Modificado por ELRO el 20130222**********************
Me.MousePointer = 0
If fgSust.Visible And fgSust.Enabled Then fgSust.SetFocus
End Sub
Sub CargaDatosCajaChica(ByVal psAreaCh As String, ByVal psAgeCh As String, ByVal pnProcCH As Integer)
Dim rs As ADODB.Recordset
Dim lnNuevoMontoAsig As Currency
Set rs = New ADODB.Recordset
Dim lnPorcCH As Double
Dim oCon As New NConstSistemas
lnNuevoMontoAsig = oCH.GetDatosCajaChica(psAreaCh, psAgeCh, MontoAsig)

Set rs = oCH.GetRSDatosCajaChica(psAreaCh, psAgeCh, pnProcCH)
Limpia
If Not rs.EOF And Not rs.BOF Then
    lblFechaApert = IIf(IsNull(rs!dDesembolso), "", rs!dDesembolso)
    lblPerscod = rs!cPersCod
    lblPersNombre = PstaNombre(rs!cPersNombre)
    lblSaldo = Format(rs!nSaldo, "#,#0.00")
    
    If lnTipoProc = gCHTipoProcRendicion Then
        lnPorcCH = oCon.LeeConstSistema("213")
        If rs!nSaldo < rs!nMontoAsig * lnPorcCH Then
            lblMsg.Caption = "El Saldo es Menor al: " & lnPorcCH * 100 & "% del monto Total Asignado. Favor de realizar su rendición."
            lblMsg.Visible = True
        Else
            lblMsg.Visible = False
            lblMsg.Caption = ""
        End If
    End If
    If lnTipoProc = gCHTipoProcHabilitacion Then
        lblImporte = Format(lnNuevoMontoAsig, gsFormatoNumeroView)
    Else
        lblImporte = Format(rs!nMontoAsig, gsFormatoNumeroView)
    End If
    Set rs = oCH.GetSolEgresoDirectoAtend(psAreaCh, psAgeCh, pnProcCH, lsCtaFondofijo, False)
    If Not rs.EOF And Not rs.BOF Then
        Set fgEgresos.Recordset = rs
    End If
    fgEgresos.FormatoPersNom 4
    lblTotalEgresos = Format(fgEgresos.SumaRow(5), "#,#0.00")
    Set rs = oCH.GetCHTodosAtenArendir(psAreaCh, psAgeCh, pnProcCH, lsCtaArendir, lsCtaPendiente)
    If Not rs.EOF And Not rs.BOF Then
        Set fgArendir.Recordset = rs
    End If
    lblTotalPend = Format(fgArendir.SumaRow(5), "#,#0.00")
    fgArendir.FormatoPersNom 2
    lblMontonoDesem = Format(oCH.GetMontoNoDesembolsado(psAreaCh, psAgeCh, pnProcCH), "#,#0.00")
    lblMontonoDesem = Format(nVal(lblMontonoDesem) + oCH.GetMontoDesembolsadoNext(psAreaCh, psAgeCh, pnProcCH + 1), "#,#0.00")
    '***Se agrego (+1) para que tome el proximo proceso de caja chica por ELRO, según SATI INC1301120002
    Totales
    
    If lnTipoProc = gCHTipoProcHabilitacion Then
        lblImporteAuto = Format(CCur(lblImporte) - (CCur(lblSaldo) + CCur(lblMontonoDesem)), "#,#0.00")
    End If
    
End If
rs.Close
Set rs = Nothing
End Sub
Sub Limpia()
lblFechaApert = ""
lblPerscod = ""
lblPersNombre = ""
lblSaldo = "0.00"
lblImporte = "0.00"
lblFechaArendRend = ""
lblTotalArendRend = "0.00"
lblTotalEgresos = "0.00"
lblTotalPend = "0.00"
fgArendir.Clear
fgArendir.FormaCabecera
fgArendir.Rows = 2
lblTotalEgresosArendir = "0.00"
lblDiferencia = "0.00"


fgEgresos.Clear
fgEgresos.FormaCabecera
fgEgresos.Rows = 2

fgSust.Clear
fgSust.FormaCabecera
fgSust.Rows = 2
End Sub
Public Sub Totales()
Dim i As Integer
Dim lnTotalArendirEgresos As Currency
Dim lnDiferencia As Currency
lnTotalArendirEgresos = 0
lnDiferencia = 0
Dim lnTotalIngArendir As Currency
Dim lnTotalEgrArendir As Currency

lnTotalIngArendir = 0
lnTotalEgrArendir = 0
If fgArendir.TextMatrix(1, 0) <> "" Then
    For i = 1 To fgArendir.Rows - 1
        If Val(fgArendir.TextMatrix(i, 10)) = Val(lblNroProcCH) Then
            lnTotalArendirEgresos = lnTotalArendirEgresos + CCur(fgArendir.TextMatrix(i, 4))
        End If
        If Val(fgArendir.TextMatrix(i, 5)) = 0 Then
        Select Case Val(fgArendir.TextMatrix(i, 9))
            Case Is < 0
                lnTotalIngArendir = lnTotalIngArendir + Abs(CCur(fgArendir.TextMatrix(i, 9)))
            Case Is > 0
                lnTotalEgrArendir = lnTotalEgrArendir + Abs(CCur(fgArendir.TextMatrix(i, 9)))
        End Select
        End If
    Next
End If
lblTotalEgresosArendir = Format(lnTotalArendirEgresos, "0.00")
lblTotalEgresPorArendir = Format(lnTotalEgrArendir, "0.00")
lblTotalIngPorArendir = Format(lnTotalIngArendir, "0.00")

lnDiferencia = CCur(lblImporte) - (lnTotalArendirEgresos - CCur(lblTotalIngPorArendir) + CCur(lblTotalEgresPorArendir) + CCur(lblTotalEgresos))
lblDiferencia = Format(lnDiferencia - (CCur(lblSaldo) + CCur(lblMontonoDesem)), "#,#0.00")
If lblDiferencia < 0 Then
    lblDiferencia.ForeColor = vbRed
Else
    If lblDiferencia > 0 Then
        lblDiferencia.ForeColor = vbBlue
    Else
        lblDiferencia.ForeColor = vbBlack
    End If
End If
lbltotalRend = Format(CCur(lblImporte) - (CCur(lblSaldo) + CCur(lblMontonoDesem)), "#,#0.00")


End Sub
Public Property Get OK() As Boolean
OK = lbOk
End Property

Public Property Let OK(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property
