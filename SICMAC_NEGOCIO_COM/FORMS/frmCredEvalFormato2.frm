VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredEvalFormato2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos - Evaluación - Formato 2"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmCredEvalFormato2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9120
      TabIndex        =   40
      Top             =   9480
      Width           =   1170
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7920
      TabIndex        =   39
      Top             =   9480
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   6570
      Left            =   120
      TabIndex        =   42
      Top             =   2760
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11589
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Ingresos y Egresos"
      TabPicture(0)   =   "frmCredEvalFormato2.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Comentarios y Verificación"
      TabPicture(1)   =   "frmCredEvalFormato2.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   " Ventas y Costos "
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
         Height          =   660
         Left            =   120
         TabIndex        =   84
         Top             =   360
         Width           =   9975
         Begin VB.Label lblCostoVenta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8400
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblVentaProm 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3120
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9240
            TabIndex        =   87
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label16 
            Caption         =   "Costo de Venta :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   86
            Top             =   280
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Venta Promedio mes :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   85
            Top             =   280
            Width           =   1815
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " Ventas a la Semana "
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
         TabIndex        =   65
         Top             =   1030
         Width           =   9975
         Begin VB.CommandButton cmdQuitarVentSem 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   20
            Top             =   1260
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarVentSem 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   1260
            Width           =   1050
         End
         Begin SICMACT.FlexEdit fgVentasSem 
            Height          =   975
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   1720
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Nº-Producto-Venta alta-Dias-Venta baja-Dias-Total mes-% Costo-% Particip-% Real"
            EncabezadosAnchos=   "300-1800-1050-500-1050-500-1200-990-990-990"
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
            ColumnasAEditar =   "X-1-2-3-4-5-X-7-8-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-R-R-R-R-R-R"
            FormatosEdit    =   "0-0-2-3-2-3-2-2-2-2"
            TextArray0      =   "Nº"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Gastos del Negocio Mensual "
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
         TabIndex        =   64
         Top             =   2760
         Width           =   4935
         Begin VB.CommandButton cmdQuitarGastoNeg 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   23
            Top             =   1260
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarGastoNeg 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   1260
            Width           =   1050
         End
         Begin SICMACT.FlexEdit fgGastoNeg 
            Height          =   975
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   1720
            Cols0           =   4
            HighLight       =   1
            EncabezadosNombres=   "-Concepto-Monto-Aux"
            EncabezadosAnchos=   "300-2630-1400-0"
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
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-L"
            FormatosEdit    =   "0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin VB.Label lblTotalGastoNeg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   81
            Top             =   1260
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Gastos Familiares Mensual "
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
         Left            =   5160
         TabIndex        =   63
         Top             =   2760
         Width           =   4935
         Begin VB.CommandButton cmdQuitarGastoFam 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   1260
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarGastoFam 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   1260
            Width           =   1050
         End
         Begin SICMACT.FlexEdit fgGastoFam 
            Height          =   975
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   1720
            Cols0           =   4
            HighLight       =   1
            EncabezadosNombres=   "-Concepto-Monto-Aux"
            EncabezadosAnchos=   "300-2630-1400-0"
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
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-L"
            FormatosEdit    =   "0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin VB.Label lblTotalGastoFam 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3240
            TabIndex        =   82
            Top             =   1260
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Otros Ingresos Mensual "
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
         Height          =   1935
         Left            =   120
         TabIndex        =   62
         Top             =   4500
         Width           =   4935
         Begin VB.CommandButton cmdQuitarOtrosIng 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   29
            Top             =   1440
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarOtrosIng 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   1050
         End
         Begin SICMACT.FlexEdit fgOtrosIng 
            Height          =   1130
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   1984
            Cols0           =   4
            HighLight       =   1
            EncabezadosNombres=   "-Concepto-Monto-Aux"
            EncabezadosAnchos=   "300-2630-1400-0"
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
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-L"
            FormatosEdit    =   "0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin VB.Label lblTotalOtrosIng 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3120
            TabIndex        =   83
            Top             =   1455
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Cálculo referencial / Indicadores "
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
         Height          =   1935
         Left            =   5040
         TabIndex        =   46
         Top             =   4500
         Width           =   5055
         Begin VB.CommandButton cmdCalcular 
            Caption         =   "Calcular"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   33
            Top             =   840
            Width           =   1170
         End
         Begin SICMACT.EditMoney txtCalcMonto 
            Height          =   300
            Left            =   165
            TabIndex        =   30
            Top             =   855
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtCalcTEM 
            Height          =   300
            Left            =   1440
            TabIndex        =   31
            Top             =   855
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Spinner.uSpinner spnCalcCuotas 
            Height          =   315
            Left            =   2280
            TabIndex        =   32
            Top             =   855
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            Max             =   999
            Min             =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
         End
         Begin VB.Label Label20 
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   59
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   "TEM (%)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   58
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Nº Cuotas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   57
            Top             =   615
            Width           =   735
         End
         Begin VB.Label lblUtilNeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   52
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label lblExcedenteFam 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            TabIndex        =   51
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lblMontoMax 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   50
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label lblCuotaEstima 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   49
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblCuotaUNM 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3795
            TabIndex        =   48
            Top             =   1260
            Width           =   1140
         End
         Begin VB.Label lblCuotaExcedeFam 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3795
            TabIndex        =   47
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label Label18 
            Caption         =   "Util. Neta Mens."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   290
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Excedente Fam."
            Height          =   255
            Left            =   2520
            TabIndex        =   60
            Top             =   290
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Máx crédito : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label Label24 
            Caption         =   "Cta estimada mes: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   55
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Cuota / UNM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2480
            TabIndex        =   54
            Top             =   1310
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Cuota / Exd. Fam."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2480
            TabIndex        =   53
            Top             =   1570
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " Comentario "
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
         TabIndex        =   45
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtComent 
            Height          =   705
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " Referencias "
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
         Height          =   2775
         Left            =   -74880
         TabIndex        =   44
         Top             =   1680
         Width           =   9975
         Begin VB.CommandButton cmdQuitarRef 
            Caption         =   "Quitar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   37
            Top             =   2340
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   2340
            Width           =   1050
         End
         Begin SICMACT.FlexEdit fgRef 
            Height          =   2055
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   3625
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "Nº-Nombre-DNI-Telefono-Referido-DNl-Aux"
            EncabezadosAnchos=   "300-2730-1200-1200-2730-1200-0"
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
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "Nº"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " Verificación "
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
         TabIndex        =   43
         Top             =   4560
         Width           =   9975
         Begin VB.TextBox txtVerif 
            Height          =   705
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   9735
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   66
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Información del Negocio"
      TabPicture(0)   =   "frmCredEvalFormato2.frx":0342
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ActXCodCta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtGiroNeg"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame1 
         Height          =   1725
         Left            =   120
         TabIndex        =   67
         Top             =   800
         Width           =   9975
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Propia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   7
            Top             =   940
            Width           =   855
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Alquilada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   8
            Top             =   940
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Ambulante"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   4680
            TabIndex        =   9
            Top             =   940
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   10
            Top             =   940
            Width           =   855
         End
         Begin VB.ComboBox cboMontoSol 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCredEvalFormato2.frx":035E
            Left            =   7680
            List            =   "frmCredEvalFormato2.frx":0368
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1280
            Width           =   735
         End
         Begin VB.TextBox txtCondLocalOtros 
            Height          =   285
            Left            =   6840
            TabIndex        =   11
            Top             =   945
            Visible         =   0   'False
            Width           =   2955
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda 
            Height          =   300
            Left            =   8560
            TabIndex        =   6
            Top             =   560
            Width           =   1210
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Spinner.uSpinner spnTiempoLocalAnio 
            Height          =   315
            Left            =   2400
            TabIndex        =   4
            Top             =   540
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin Spinner.uSpinner spnTiempoLocalMes 
            Height          =   315
            Left            =   3720
            TabIndex        =   5
            Top             =   540
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            Min             =   1
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin Spinner.uSpinner spnCuotas 
            Height          =   315
            Left            =   4800
            TabIndex        =   13
            Top             =   1280
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   999
            Min             =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.EditMoney txtMontoSol 
            Height          =   300
            Left            =   8490
            TabIndex        =   15
            Top             =   1275
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtCuotaPagar 
            Height          =   295
            Left            =   2400
            TabIndex        =   12
            Top             =   1280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Spinner.uSpinner spnExpEmpAnio 
            Height          =   315
            Left            =   2400
            TabIndex        =   1
            Top             =   210
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin Spinner.uSpinner spnExpEmpMes 
            Height          =   315
            Left            =   3720
            TabIndex        =   2
            Top             =   210
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            Min             =   1
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.EditMoney txtUltEndeuda 
            Height          =   300
            Left            =   8560
            TabIndex        =   3
            Top             =   210
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Experiencia como empresario :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Tiempo en el mismo local :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Condición local :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Probable cuota a pagar (mes) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "años"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3200
            TabIndex        =   75
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "años"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3200
            TabIndex        =   74
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "meses"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4510
            TabIndex        =   73
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "meses"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4510
            TabIndex        =   72
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Último endeudamiento :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   71
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha último endeudamiento :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   70
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Nº Cuotas :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   69
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Monto solicitado :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   68
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.TextBox txtGiroNeg 
         Height          =   300
         Left            =   5760
         TabIndex        =   0
         Top             =   420
         Width           =   4155
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito"
      End
      Begin VB.Label Label1 
         Caption         =   "Giro del Negocio :"
         Height          =   255
         Left            =   4440
         TabIndex        =   80
         Top             =   445
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCredEvalFormato2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredEvalFormato2
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'**               creado segun RFC090-2012
'** Creación : JUEZ, 20120903 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim sCtaCod As String
Dim gsOpeCod As String
Dim fnTipoRegMant As Integer
Dim fnTipoPermiso As Integer
Dim fbPermiteGrabar As Boolean
Dim fbBloqueaTodo As Boolean
Dim lnIndMaximaCapPago As Double
Dim lnIndCuotaUNM As Double
Dim lnIndCuotaExcFam As Double
Dim lnCondLocal As Integer
Dim rsCredEval As ADODB.Recordset
Dim rsInd As ADODB.Recordset
Dim rsDatGastoNeg As ADODB.Recordset
Dim rsDatGastoFam As ADODB.Recordset
Dim rsDatOtrosIng As ADODB.Recordset
Dim rsDatVentaSem As ADODB.Recordset
Dim rsDatRef As ADODB.Recordset
Dim cSPrd As String, cPrd As String
Dim DCredito As COMDCredito.DCOMCredito
Dim objPista As COMManejador.Pista
Dim nFormato As Integer
Dim fnMontoIni As Double
Dim lnMin As Double, lnMax As Double
Dim lnMinDol As Double, lnMaxDol As Double
Dim nTC As Double
Dim fsCliente As String, fsUserAnalista As String

Public Sub Inicio(ByVal psCtaCod As String, ByVal psTipoRegMant As Integer)
    
    Dim oCred As COMNCredito.NCOMCredito
    Dim rsDCredito As ADODB.Recordset
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    
    Set oCred = New COMNCredito.NCOMCredito
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
    
    sCtaCod = psCtaCod
    fnTipoRegMant = psTipoRegMant
    '109017041002050005
    ActXCodCta.NroCuenta = sCtaCod
    fnTipoPermiso = oCred.ObtieneTipoPermisoCredEval(gsCodCargo)
    gsOpeCod = ""
    cSPrd = Mid(sCtaCod, 6, 3)
    cPrd = Mid(cSPrd, 1, 1) & "00"
    
    fbPermiteGrabar = False
    fbBloqueaTodo = False
    
    Set DCredito = New COMDCredito.DCOMCredito
    Set rsDCredito = DCredito.RecuperaSolicitudDatoBasicos(sCtaCod) 'JUEZ 20121216
    fnMontoIni = rsDCredito!nMonto
    fsCliente = rsDCredito!cPersNombre
    fsUserAnalista = rsDCredito!UserAnalista 'JUEZ 20121216
    
    Set rsDCredEval = DCredito.RecuperaColocacCredEval(sCtaCod)
    If fnTipoPermiso = 2 Then
        If rsDCredEval.RecordCount = 0 Then
            MsgBox "El analista no ha registrado la Evaluacion respectiva", vbExclamation, "Aviso"
            fbPermiteGrabar = False
        Else
            fbPermiteGrabar = True
        End If
    End If
    Set rsDCredito = Nothing
    Set rsDCredEval = Nothing
    
    Set rsDColCred = DCredito.RecuperaColocacCred(sCtaCod)
    If rsDColCred!nVerifCredEval = 1 Then
        MsgBox "Ud. no puede editar la evaluación, ya se realizó la verificacion del credito", vbExclamation, "Aviso"
        fbBloqueaTodo = True
    End If
    
    nFormato = DCredito.AsignarFormato(cPrd, cSPrd, fnMontoIni, lnMin, lnMax)
    lnMinDol = lnMin / nTC
    lnMaxDol = lnMax / nTC
    
    Set DCredito = Nothing
    Set oTipoCam = Nothing
    If CargaDatos Then
        If CargaControles(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
            If fnTipoRegMant = 1 Then
                If Not rsCredEval.EOF Then
                    'MsgBox "Ya se realizó el registro de la Evaluación", vbInformation, "Aviso"
                    Call Mantenimiento
                    fnTipoRegMant = 2
                Else
                    Call Registro
                    fnTipoRegMant = 1
                End If
            Else
                If rsCredEval.EOF Then
                    'MsgBox "Primero debe registrar los datos de la Evaluación", vbInformation, "Aviso"
                    Call Registro
                    fnTipoRegMant = 1
                Else
                    Call Mantenimiento
                    fnTipoRegMant = 2
                End If
            End If
        Else
            Unload Me
            Exit Sub
        End If
    Else
        If CargaControles(1, False) Then
        End If
    End If
    Me.Show 1
End Sub

Private Function CargaDatos() As Boolean

On Error GoTo ErrorCargaDatos

    Dim oCred As COMNCredito.NCOMCredito
    Dim i As Integer
    Set oCred = New COMNCredito.NCOMCredito
       
    CargaDatos = oCred.CargaDatosCredEvaluacion(sCtaCod, 1, rsCredEval, rsInd, rsDatGastoNeg, _
                                                rsDatGastoFam, rsDatOtrosIng, rsDatRef, rsDatVentaSem)
    
    If CargaDatos Then
        For i = 1 To rsInd.RecordCount
            If rsInd!cIndicadorID = "IND001" Or rsInd!cIndicadorID = "IND002" Then lnIndMaximaCapPago = rsInd!cIndicadorPorc / 100
            If rsInd!cIndicadorID = "IND003" Or rsInd!cIndicadorID = "IND004" Then lnIndCuotaUNM = rsInd!cIndicadorPorc / 100
            If rsInd!cIndicadorID = "IND005" Or rsInd!cIndicadorID = "IND006" Then lnIndCuotaExcFam = rsInd!cIndicadorPorc / 100
            rsInd.MoveNext
        Next
    End If
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbCritical, "Error"
End Function

Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtMontoSol.Text = Format(fnMontoIni, "#,##0.00")
    cboMontoSol.ListIndex = IndiceListaCombo(cboMontoSol, Mid(sCtaCod, 9, 1))
    lblVentaProm.Caption = Format(0, "#,##0.00")
    lblCostoVenta.Caption = Format(0, "#,##0.00")
    lblUtilNeta.Caption = Format(0, "#,##0.00")
    lblExcedenteFam.Caption = Format(0, "#,##0.00")
    
    lblTotalGastoNeg.Caption = Format(0, "#,##0.00")
    lblTotalGastoFam.Caption = Format(0, "#,##0.00")
    lblTotalOtrosIng.Caption = Format(0, "#,##0.00")
    
    txtCalcMonto.Text = Format(fnMontoIni, "#,##0.00")
    'txtCalcTEM.Text = Format(0, "#,##0.00")
    
    lblMontoMax.Caption = Format(0, "#,##0.00")
    lblCuotaEstima.Caption = Format(0, "#,##0.00")
    lblCuotaUNM.Caption = Format(0, "#,##0.00")
    lblCuotaExcedeFam.Caption = Format(0, "#,##0.00")
End Function

Private Function Mantenimiento()
    Dim lnFila As Integer
    If fnTipoPermiso = 3 Then
        gsOpeCod = gCredMantenimientoEvaluacionCred
    Else
        gsOpeCod = gCredVerificacionEvaluacionCred
    End If
    txtGiroNeg.Text = rsCredEval!cGiroNeg
    spnExpEmpAnio.valor = rsCredEval!nExpEmpAnio
    spnExpEmpMes.valor = rsCredEval!nExpEmpMes
    spnTiempoLocalAnio.valor = rsCredEval!nTiempoLocalAnio
    spnTiempoLocalMes.valor = rsCredEval!nTiempoLocalMes
    OptCondLocal(rsCredEval!nCondLocal).value = 1
    txtCondLocalOtros.Text = rsCredEval!cCondLocalOtros
    txtCuotaPagar.Text = Format(rsCredEval!cCuotaPagar, "#,##0.00")
    spnCuotas.valor = rsCredEval!nCuotas
    txtUltEndeuda.Text = Format(rsCredEval!cUltEndeuda, "#,##0.00")
    If rsCredEval!cUltEndeuda = 0 Then
        txtFecUltEndeuda.Enabled = False
    Else
        If fnTipoPermiso = 3 Then
            txtFecUltEndeuda.Enabled = True
        End If
    End If
    txtFecUltEndeuda.Text = Format(IIf(rsCredEval!cFecUltEndeuda = "01/01/1900", "__/__/____", rsCredEval!cFecUltEndeuda), "dd/mm/yyyy")
    cboMontoSol.ListIndex = IndiceListaCombo(cboMontoSol, rsCredEval!nmoneda)
    txtMontoSol.Text = Format(rsCredEval!nMontoSol, "#,##0.00")
    lblVentaProm.Caption = Format(rsCredEval!nVentaProm, "#,##0.00")
    lblCostoVenta.Caption = Format(rsCredEval!nCostoVenta, "#,##0.00")
    'Call FormatearGrillas(fgGastoNeg)
    Call LimpiaFlex(fgGastoNeg)
    Do While Not rsDatGastoNeg.EOF
        fgGastoNeg.AdicionaFila
        lnFila = fgGastoNeg.Row
        fgGastoNeg.TextMatrix(lnFila, 1) = rsDatGastoNeg!cConcepto
        fgGastoNeg.TextMatrix(lnFila, 2) = Format(rsDatGastoNeg!nMonto, "#,##0.00")
        lblTotalGastoNeg.Caption = Format(CDbl(IIf(lblTotalGastoNeg.Caption = "", 0, lblTotalGastoNeg.Caption)) + rsDatGastoNeg!nMonto, "#,##0.00")
        rsDatGastoNeg.MoveNext
    Loop
    rsDatGastoNeg.Close
    Set rsDatGastoNeg = Nothing
    'Call FormatearGrillas(fgGastoFam)
    Call LimpiaFlex(fgGastoFam)
    Do While Not rsDatGastoFam.EOF
        fgGastoFam.AdicionaFila
        lnFila = fgGastoFam.Row
        fgGastoFam.TextMatrix(lnFila, 1) = rsDatGastoFam!cConcepto
        fgGastoFam.TextMatrix(lnFila, 2) = Format(rsDatGastoFam!nMonto, "#,##0.00")
        lblTotalGastoFam.Caption = Format(CDbl(IIf(lblTotalGastoFam.Caption = "", 0, lblTotalGastoFam.Caption)) + rsDatGastoFam!nMonto, "#,##0.00")
        rsDatGastoFam.MoveNext
    Loop
    rsDatGastoFam.Close
    Set rsDatGastoFam = Nothing
    'Call FormatearGrillas(fgOtrosIng)
    Call LimpiaFlex(fgOtrosIng)
    Do While Not rsDatOtrosIng.EOF
        fgOtrosIng.AdicionaFila
        lnFila = fgOtrosIng.Row
        fgOtrosIng.TextMatrix(lnFila, 1) = rsDatOtrosIng!cConcepto
        fgOtrosIng.TextMatrix(lnFila, 2) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
        lblTotalOtrosIng.Caption = Format(CDbl(IIf(lblTotalOtrosIng.Caption = "", 0, lblTotalOtrosIng.Caption)) + rsDatOtrosIng!nMonto, "#,##0.00")
        rsDatOtrosIng.MoveNext
    Loop
    rsDatOtrosIng.Close
    Set rsDatOtrosIng = Nothing
    Call LimpiaFlex(fgVentasSem)
    Do While Not rsDatVentaSem.EOF
        fgVentasSem.AdicionaFila
        lnFila = fgVentasSem.Row
        fgVentasSem.TextMatrix(lnFila, 1) = rsDatVentaSem!cProducto
        fgVentasSem.TextMatrix(lnFila, 2) = Format(rsDatVentaSem!nVentaAlta, "#,##0.00")
        fgVentasSem.TextMatrix(lnFila, 3) = rsDatVentaSem!nDiasAlta
        fgVentasSem.TextMatrix(lnFila, 4) = Format(rsDatVentaSem!nVentaBaja, "#,##0.00")
        fgVentasSem.TextMatrix(lnFila, 5) = rsDatVentaSem!nDiasBaja
        fgVentasSem.TextMatrix(lnFila, 6) = Format(rsDatVentaSem!nTotalMes, "#,##0.00")
        fgVentasSem.TextMatrix(lnFila, 7) = Format(rsDatVentaSem!nCosto, "#,##0.00")
        fgVentasSem.TextMatrix(lnFila, 8) = Format(rsDatVentaSem!nParticip, "#,##0.00")
        fgVentasSem.TextMatrix(lnFila, 9) = Format(rsDatVentaSem!nReal, "#,##0.00")
        rsDatVentaSem.MoveNext
    Loop
    rsDatVentaSem.Close
    Set rsDatVentaSem = Nothing

    lblUtilNeta.Caption = Format(rsCredEval!nUtilNeta, "#,##0.00")
    lblExcedenteFam.Caption = Format(rsCredEval!nExcedenteFam, "#,##0.00")
    
    txtCalcMonto.Text = Format(rsCredEval!nMontoCalc, "#,##0.00")
    txtCalcTEM.Text = Format(rsCredEval!nTEMCalc, "#,##0.00")
    spnCalcCuotas.valor = rsCredEval!nCuotasCalc

    lblMontoMax.Caption = Format(rsCredEval!nMontoMax, "#,##0.00")
    lblCuotaEstima.Caption = Format(rsCredEval!nCuotaEstima, "#,##0.00")
    lblCuotaUNM.Caption = Format(rsCredEval!nCuotaUNM, "#,##0.00")
    lblCuotaExcedeFam.Caption = Format(rsCredEval!nCuotaExcedeFam, "#,##0.00")

    txtComent.Text = rsCredEval!cComent
    'Call FormatearGrillas(fgRef)
    Call LimpiaFlex(fgRef)
    Do While Not rsDatRef.EOF
        fgRef.AdicionaFila
        lnFila = fgRef.Row
        fgRef.TextMatrix(lnFila, 1) = rsDatRef!cNombre
        fgRef.TextMatrix(lnFila, 2) = rsDatRef!cDNI
        fgRef.TextMatrix(lnFila, 3) = rsDatRef!cTelef
        fgRef.TextMatrix(lnFila, 4) = rsDatRef!cReferido
        fgRef.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
        rsDatRef.MoveNext
    Loop
    rsDatRef.Close
    Set rsDatRef = Nothing
    
    txtVerif.Text = rsCredEval!cVerif
    
End Function

Private Function CargaControles(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, True, pPermiteGrabar)
        CargaControles = True
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControles = True
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        CargaControles = False
    End If
    If pBloqueaTodo Then
        Call HabilitaControles(False, False, False)
        CargaControles = True
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaB As Boolean, ByVal pbHabilitaGuardar As Boolean)
    txtGiroNeg.Enabled = pbHabilitaA
    spnExpEmpAnio.Enabled = pbHabilitaA
    spnExpEmpMes.Enabled = pbHabilitaA
    spnTiempoLocalAnio.Enabled = pbHabilitaA
    spnTiempoLocalMes.Enabled = pbHabilitaA
    OptCondLocal(1).Enabled = pbHabilitaA
    OptCondLocal(2).Enabled = pbHabilitaA
    OptCondLocal(3).Enabled = pbHabilitaA
    OptCondLocal(4).Enabled = pbHabilitaA
    txtCondLocalOtros.Enabled = pbHabilitaA
    txtCuotaPagar.Enabled = pbHabilitaA
    spnCuotas.Enabled = pbHabilitaA
    txtUltEndeuda.Enabled = pbHabilitaA
    txtFecUltEndeuda.Enabled = pbHabilitaA
    'cboMontoSol.Enabled = pbHabilitaA
    txtMontoSol.Enabled = pbHabilitaA
    fgVentasSem.Enabled = pbHabilitaA
    cmdAgregarVentSem.Enabled = pbHabilitaA
    cmdQuitarVentSem.Enabled = pbHabilitaA
    fgGastoNeg.Enabled = pbHabilitaA
    cmdAgregarGastoNeg.Enabled = pbHabilitaA
    cmdQuitarGastoNeg.Enabled = pbHabilitaA
    fgGastoFam.Enabled = pbHabilitaA
    cmdAgregarGastoFam.Enabled = pbHabilitaA
    cmdQuitarGastoFam.Enabled = pbHabilitaA
    fgOtrosIng.Enabled = pbHabilitaA
    cmdAgregarOtrosIng.Enabled = pbHabilitaA
    cmdQuitarOtrosIng.Enabled = pbHabilitaA
    txtCalcMonto.Enabled = pbHabilitaA
    txtCalcTEM.Enabled = pbHabilitaA
    spnCalcCuotas.Enabled = pbHabilitaA
    cmdCalcular.Enabled = pbHabilitaA
    txtComent.Enabled = pbHabilitaA
    fgRef.Enabled = pbHabilitaA
    cmdAgregarRef.Enabled = pbHabilitaA
    cmdQuitarRef.Enabled = pbHabilitaA
    
    txtVerif.Enabled = pbHabilitaB
    
    cmdGrabar.Enabled = pbHabilitaGuardar
    
    If Mid(sCtaCod, 9, 1) = "2" Then
        Me.txtMontoSol.BackColor = RGB(200, 255, 200)
        Me.txtCuotaPagar.BackColor = RGB(200, 255, 200)
        
        txtCalcMonto.BackColor = RGB(200, 255, 200)
        lblMontoMax.BackColor = RGB(200, 255, 200)
        lblCuotaEstima.BackColor = RGB(200, 255, 200)
        lblCuotaUNM.BackColor = RGB(200, 255, 200)
        lblCuotaExcedeFam.BackColor = RGB(200, 255, 200)
    Set DCredito = Nothing
    Else
        Me.txtMontoSol.BackColor = &HFFFFFF
        Me.txtCuotaPagar.BackColor = &HFFFFFF
        txtCalcMonto.BackColor = &HFFFFFF
    
        lblMontoMax.BackColor = &HFFFFFF
        lblCuotaEstima.BackColor = &HFFFFFF
        lblCuotaUNM.BackColor = &HFFFFFF
        lblCuotaExcedeFam.BackColor = &HFFFFFF
    End If
End Function

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGiroNeg.SetFocus
    End If
End Sub

Private Sub cboMontoSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoSol.SetFocus
    End If
End Sub

Private Sub cmdAgregarGastoFam_Click()
    If fgGastoFam.Rows - 1 < 25 Then
        fgGastoFam.AdicionaFila
        fgGastoFam.SetFocus
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdAgregarGastoNeg_Click()
    If fgGastoNeg.Rows - 1 < 25 Then
        fgGastoNeg.AdicionaFila
        fgGastoNeg.SetFocus
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdAgregarOtrosIng_Click()
    If fgOtrosIng.Rows - 1 < 25 Then
        fgOtrosIng.AdicionaFila
        fgOtrosIng.SetFocus
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdAgregarRef_Click()
    If fgRef.Rows - 1 < 25 Then
        fgRef.AdicionaFila
        fgRef.SetFocus
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdAgregarVentSem_Click()
    If Not ValidaPorcParticip Then
        MsgBox "El % de Participacion no debe sumar mas del 100%", vbInformation, "Aviso"
        Exit Sub
    End If
    fgVentasSem.AdicionaFila
    fgVentasSem.SetFocus
    SendKeys "{Enter}"
End Sub

Private Sub cmdCalcular_Click()

On Error GoTo ErrorCalcular
    Dim pnTEM As Double, pnMonto As Double
    Dim MontoMax As Double, CuotaEstima As Double
    Dim CuotaUNM As Double, CuotaExcedeFam As Double
    Dim pnCuotas As Integer
    Dim pnFormula As Double
    
    pnMonto = CDbl(txtCalcMonto.Text)
    pnTEM = CDbl(txtCalcTEM.Text) / 100
    pnCuotas = CInt(spnCalcCuotas.valor)
    
    pnFormula = (((pnTEM * ((1 + pnTEM) ^ pnCuotas))) / (((1 + pnTEM) ^ pnCuotas) - 1))
    
    MontoMax = (CDbl(lblExcedenteFam.Caption) * lnIndMaximaCapPago) / pnFormula
    CuotaEstima = pnMonto * pnFormula
    CuotaUNM = (CuotaEstima / CDbl(lblUtilNeta.Caption)) * 100
    CuotaExcedeFam = (CuotaEstima / CDbl(lblExcedenteFam.Caption)) * 100
    
    lblMontoMax.Caption = Format(MontoMax, "#,##0.00")
    lblCuotaEstima.Caption = Format(CuotaEstima, "#,##0.00")
    lblCuotaUNM.Caption = Format(CuotaUNM, "#,##0.00")
    lblCuotaExcedeFam = Format(CuotaExcedeFam, "#,##0.00")
    
    If Round(CDbl(lblMontoMax.Caption), 2) < Round(CDbl(txtCalcMonto.Text), 2) Then
        MsgBox "El Monto Máximo del Credito es menor al ingresado en el calculo", vbInformation, "Aviso"
        txtCalcMonto.SetFocus
        SSTab2.Tab = 0
        Exit Sub
    End If
    If Round(CDbl(lblCuotaEstima.Caption), 2) > Round(CDbl(txtCuotaPagar.Text), 2) Then
        MsgBox "La Couta Estimada a Pagar es mayor a la Probable Cuota por Pagar", vbInformation, "Aviso"
        txtCuotaPagar.SetFocus
        SSTab2.Tab = 0
        Exit Sub
    End If
    
    Exit Sub
ErrorCalcular:
    MsgBox Err.Description + ": Verifique que todos los datos esten ingresados", vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("Desea salir del Formato de Evaluación??", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub CmdGrabar_Click()
    
    Dim oCred As COMNCredito.NCOMCredito
    Dim GrabarDatos As Boolean
    Dim rsGastoNeg As ADODB.Recordset
    Dim rsGastoFam As ADODB.Recordset
    Dim rsOtrosIng As ADODB.Recordset
    Dim rsVentasSem As ADODB.Recordset
    Dim rsRef As ADODB.Recordset
    
    Set rsGastoNeg = IIf(fgGastoNeg.Rows - 1 > 0, fgGastoNeg.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(fgGastoFam.Rows - 1 > 0, fgGastoFam.GetRsNew(), Nothing)
    Set rsOtrosIng = IIf(fgOtrosIng.Rows - 1 > 0, fgOtrosIng.GetRsNew(), Nothing)
    
    Set rsRef = IIf(fgRef.Rows - 1 > 0, fgRef.GetRsNew(), Nothing)
    
    If validaDatos Then
        Set rsVentasSem = LlenaRecordSet_VentasSemana
         
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        If txtFecUltEndeuda.Text = "__/__/____" Then
            txtFecUltEndeuda.Text = "01/01/1900"
        End If
        
        Set oCred = New COMNCredito.NCOMCredito
        Set objPista = New COMManejador.Pista
        
        If fnTipoPermiso = 3 Then
            GrabarDatos = oCred.GrabarCredEvaluacionFormatoIyII(sCtaCod, fnTipoRegMant, Trim(Me.txtGiroNeg.Text), Me.spnExpEmpAnio.valor, Me.spnExpEmpMes.valor, _
                                                            Me.spnTiempoLocalAnio.valor, Me.spnTiempoLocalMes.valor, lnCondLocal, _
                                                            IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text), _
                                                            txtCuotaPagar.Text, spnCuotas.valor, txtUltEndeuda.Text, _
                                                            Format(txtFecUltEndeuda, "dd/mm/yyyy"), _
                                                            CInt(Trim(Right(cboMontoSol.Text, 2))), txtMontoSol.Text, lblVentaProm.Caption, _
                                                            lblCostoVenta.Caption, rsGastoNeg, rsGastoFam, rsOtrosIng, CDbl(lblUtilNeta.Caption), _
                                                            CDbl(lblExcedenteFam.Caption), CDbl(txtCalcMonto.value), CDbl(txtCalcTEM.value), _
                                                            spnCalcCuotas.valor, CDbl(lblMontoMax.Caption), CDbl(lblCuotaEstima.Caption), _
                                                            CDbl(lblCuotaUNM.Caption), CDbl(lblCuotaExcedeFam.Caption), Trim(txtComent.Text), rsRef, rsVentasSem, 2)
        Else
            GrabarDatos = oCred.GrabarCredEvaluacionVerif(sCtaCod, Trim(txtVerif.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
        
        If GrabarDatos Then
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 1", sCtaCod, gCodigoCuenta
            
            If txtFecUltEndeuda.Text = "01/01/1900" Then
                txtFecUltEndeuda.Text = "__/__/____"
            End If
            If fnTipoRegMant = 1 Then
                MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
            Else
                MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
            End If
            Call GeneraExcelFormato
            Unload Me
        Else
            MsgBox "Hubo errores al grabar la información", vbError, "Error"
        End If
    End If
End Sub

Private Sub cmdQuitarGastoFam_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(fgGastoFam.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        fgGastoFam.EliminaFila fgGastoFam.Row
        lblTotalGastoFam.Caption = Format(SumarCampo(fgGastoFam, 2), "#0.00")
        Call CalculaUtilidadNetaMensual
        Call CalculaExcedenteFam
    End If
End Sub

Private Sub cmdQuitarGastoNeg_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(fgGastoNeg.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        fgGastoNeg.EliminaFila fgGastoNeg.Row
        lblTotalGastoNeg.Caption = Format(SumarCampo(fgGastoNeg, 2), "#,##0.00")
        Call CalculaUtilidadNetaMensual
    End If
End Sub

Private Sub cmdQuitarOtrosIng_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(fgOtrosIng.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        fgOtrosIng.EliminaFila fgOtrosIng.Row
        lblTotalOtrosIng.Caption = Format(SumarCampo(fgOtrosIng, 2), "#,##0.00")
        Call CalculaUtilidadNetaMensual
        Call CalculaExcedenteFam
    End If
End Sub

Private Sub cmdQuitarRef_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(fgRef.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        fgRef.EliminaFila fgRef.Row
    End If
End Sub

Private Sub cmdQuitarVentSem_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(fgVentasSem.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        fgVentasSem.EliminaFila fgVentasSem.Row
        lblVentaProm.Caption = Format(SumarCampo(fgVentasSem, 6), "#,##0.00")
        lblCostoVenta.Caption = Format(SumarCampo(fgVentasSem, 9), "#,##0.00")
        Call CalculaUtilidadNetaMensual
        Call CalculaExcedenteFam
    End If
End Sub

Private Sub fgGastoFam_OnCellChange(pnRow As Long, pnCol As Long)
    If fgGastoFam.Col = 1 Then
        fgGastoFam.TextMatrix(fgGastoFam.Row, 1) = UCase(fgGastoFam.TextMatrix(fgGastoFam.Row, 1))
    ElseIf fgGastoFam.Col = 2 Then
        lblTotalGastoFam.Caption = Format(SumarCampo(fgGastoFam, 2), "#,##0.00")
        Call CalculaUtilidadNetaMensual
        Call CalculaExcedenteFam
        If cmdAgregarGastoFam.Enabled Then
            cmdAgregarGastoFam.SetFocus
        End If
    End If
End Sub

Private Sub fgGastoNeg_OnCellChange(pnRow As Long, pnCol As Long)
    If fgGastoNeg.Col = 1 Then
        fgGastoNeg.TextMatrix(fgGastoNeg.Row, 1) = UCase(fgGastoNeg.TextMatrix(fgGastoNeg.Row, 1))
    ElseIf fgGastoNeg.Col = 2 Then
        lblTotalGastoNeg.Caption = Format(SumarCampo(fgGastoNeg, 2), "#,##0.00")
        Call CalculaUtilidadNetaMensual
        If cmdAgregarGastoNeg.Enabled Then
            cmdAgregarGastoNeg.SetFocus
        End If
    End If
End Sub

Private Sub fgOtrosIng_OnCellChange(pnRow As Long, pnCol As Long)
    If fgOtrosIng.Col = 1 Then
        fgOtrosIng.TextMatrix(fgOtrosIng.Row, 1) = UCase(fgOtrosIng.TextMatrix(fgOtrosIng.Row, 1))
    ElseIf fgOtrosIng.Col = 2 Then
        lblTotalOtrosIng.Caption = Format(SumarCampo(fgOtrosIng, 2), "#,##0.00")
        Call CalculaUtilidadNetaMensual
        Call CalculaExcedenteFam
        If cmdAgregarOtrosIng.Enabled Then
            cmdAgregarOtrosIng.SetFocus
        End If
    End If
End Sub

Private Sub fgRef_OnCellChange(pnRow As Long, pnCol As Long)
    If fgRef.Col = 1 Or fgRef.Col = 4 Then
        fgRef.TextMatrix(fgRef.Row, fgRef.Col) = UCase(fgRef.TextMatrix(fgRef.Row, fgRef.Col))
    End If
End Sub

Private Sub fgVentasSem_OnCellChange(pnRow As Long, pnCol As Long)
    If fgVentasSem.Col = 1 Then
        fgVentasSem.TextMatrix(fgVentasSem.Row, 1) = UCase(fgVentasSem.TextMatrix(fgVentasSem.Row, 1))
    ElseIf fgVentasSem.Col = 2 Or fgVentasSem.Col = 3 Or fgVentasSem.Col = 4 Or fgVentasSem.Col = 5 Then
        fgVentasSem.TextMatrix(fgVentasSem.Row, 6) = Format(CalculaTotalMesVentasSem(CDbl(IIf(fgVentasSem.TextMatrix(fgVentasSem.Row, 2) = "", 0, fgVentasSem.TextMatrix(fgVentasSem.Row, 2))), _
                                                                                    CDbl(IIf(fgVentasSem.TextMatrix(fgVentasSem.Row, 3) = "", 0, fgVentasSem.TextMatrix(fgVentasSem.Row, 3))), _
                                                                                    CDbl(IIf(fgVentasSem.TextMatrix(fgVentasSem.Row, 4) = "", 0, fgVentasSem.TextMatrix(fgVentasSem.Row, 4))), _
                                                                                    CDbl(IIf(fgVentasSem.TextMatrix(fgVentasSem.Row, 5) = "", 0, fgVentasSem.TextMatrix(fgVentasSem.Row, 5)))), "#,##0.00")
    ElseIf fgVentasSem.Col = 7 Or fgVentasSem.Col = 8 Then
        fgVentasSem.TextMatrix(fgVentasSem.Row, 9) = Format(CalculaPorcRealVentasSem(CDbl(IIf(fgVentasSem.TextMatrix(fgVentasSem.Row, 7) = "", 0, fgVentasSem.TextMatrix(fgVentasSem.Row, 7))), _
                                                                                    CDbl(IIf(fgVentasSem.TextMatrix(fgVentasSem.Row, 8) = "", 0, fgVentasSem.TextMatrix(fgVentasSem.Row, 8)))), "#,##0.00")
        If Not ValidaPorcParticip Then
            MsgBox "El % de Participacion no debe sumar mas del 100%", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    lblVentaProm.Caption = Format(SumarCampo(fgVentasSem, 6), "#,##0.00")
    lblCostoVenta.Caption = Format(SumarCampo(fgVentasSem, 9), "#,##0.00")
    Call CalculaUtilidadNetaMensual
    Call CalculaExcedenteFam

End Sub

Public Function ValidaPorcParticip()
    ValidaPorcParticip = False
    If SumarCampo(fgVentasSem, 8) <= 100 Then
        ValidaPorcParticip = True
    End If
End Function

Private Sub OptCondLocal_Click(Index As Integer)
    Select Case Index
    Case 1, 2, 3
        Me.txtCondLocalOtros.Visible = False
        Me.txtCondLocalOtros.Text = ""
        'Me.txtCuotaPagar.SetFocus
    Case 4
        Me.txtCondLocalOtros.Visible = True
        Me.txtCondLocalOtros.Text = ""
        'Me.txtCondLocalOtros.SetFocus
    End Select
    lnCondLocal = Index
End Sub

Private Sub spnCalcCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdCalcular.SetFocus
    End If
End Sub

Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtUltEndeuda.SetFocus
    End If
End Sub

Private Sub spnExpEmpAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnExpEmpMes.SetFocus
    End If
End Sub

Private Sub spnExpEmpMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnTiempoLocalAnio.SetFocus
    End If
End Sub

Private Sub spnTiempoLocalAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnTiempoLocalMes.SetFocus
    End If
End Sub

Private Sub spnTiempoLocalMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OptCondLocal(1).SetFocus
    End If
End Sub

Private Sub txtCalcMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCalcTEM.SetFocus
    End If
End Sub

Private Sub txtCalcMonto_LostFocus()
    If txtCalcMonto.Text < lnMin Or txtCalcMonto.Text > lnMax Then
        MsgBox "El monto ingresado está fuera de los rangos de este formato.", vbInformation, "Aviso"
        MsgBox "El Monto Solicitado para este formato debe estar entre " & _
        IIf(Mid(sCtaCod, 9, 1) = "1", "S/. ", "$ ") & IIf(Mid(sCtaCod, 9, 1) = "1", Format(lnMin, "#,##0.00"), Format(lnMinDol, "#,##0.00")) & " y " & _
        IIf(Mid(sCtaCod, 9, 1) = "1", "S/. ", "$ ") & IIf(Mid(sCtaCod, 9, 1) = "1", Format(lnMax, "#,##0.00"), Format(lnMaxDol, "#,##0.00")), vbInformation, "Aviso"
        txtCalcMonto.SetFocus
    End If
End Sub

Private Sub txtCalcTEM_Change()
    If CDbl(txtCalcTEM.Text) > 100 Then
        txtCalcTEM.Text = Replace(Mid(lblCostoVenta.Caption, 1, Len(lblCostoVenta.Caption) - 1), ",", "")
    End If
End Sub

Private Sub txtCalcTEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnCalcCuotas.SetFocus
    End If
End Sub

Private Sub txtComent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAgregarRef.SetFocus
    End If
End Sub

Private Sub txtComent_LostFocus()
    txtComent.Text = UCase(txtComent.Text)
End Sub

'Private Sub txtCostoVenta_Change()
'    If Trim(txtCostoVenta.Text) <> "." Then
'        If CDbl(txtCostoVenta.Text) > 100 Then
'            txtCostoVenta.Text = Replace(Mid(txtCostoVenta.Text, 1, Len(txtCostoVenta.Text) - 1), ",", "")
'        End If
'    Else
'        txtCostoVenta.Text = ""
'    End If
'End Sub

'Private Sub txtCostoVenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        CalculaUtilidadNetaMensual
'        CalculaExcedenteFam
'        cmdAgregarGastoNeg.SetFocus
'    End If
'End Sub

'Private Sub txtCostoVenta_LostFocus()
'    CalculaUtilidadNetaMensual
'    CalculaExcedenteFam
'End Sub

Private Sub txtCuotaPagar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnCuotas.SetFocus
    End If
End Sub

Private Sub txtFecUltEndeuda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoSol.SetFocus
    End If
End Sub

Private Sub txtFecUltEndeuda_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFecUltEndeuda.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        If txtFecUltEndeuda.Enabled Then
            'txtFecUltEndeuda.SetFocus
            Exit Sub
        End If
    End If
    If CDate(IIf(txtFecUltEndeuda.Text = "__/__/____", "10/10/1900", txtFecUltEndeuda.Text)) > gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        'txtFecUltEndeuda.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtGiroNeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spnExpEmpAnio.SetFocus
    End If
End Sub

Private Sub txtGiroNeg_LostFocus()
    txtGiroNeg.Text = UCase(txtGiroNeg)
End Sub

Private Sub txtMontoSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarVentSem.SetFocus
    End If
End Sub

Private Sub txtMontoSol_LostFocus()
    If txtMontoSol.Text < lnMin Or txtMontoSol.Text > lnMax Then
        MsgBox "El monto ingresado está fuera de los rangos de este formato.", vbInformation, "Aviso"
        MsgBox "El Monto Solicitado para este formato debe estar entre " & _
        IIf(Mid(sCtaCod, 9, 1) = "1", "S/. ", "$ ") & IIf(Mid(sCtaCod, 9, 1) = "1", Format(lnMin, "#,##0.00"), Format(lnMinDol, "#,##0.00")) & " y " & _
        IIf(Mid(sCtaCod, 9, 1) = "1", "S/. ", "$ ") & IIf(Mid(sCtaCod, 9, 1) = "1", Format(lnMax, "#,##0.00"), Format(lnMaxDol, "#,##0.00")), vbInformation, "Aviso"
        txtMontoSol.SetFocus
    End If
End Sub

Private Sub txtUltEndeuda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUltEndeuda <> 0 Then
            txtFecUltEndeuda.Enabled = True
            txtFecUltEndeuda.SetFocus
        Else
            txtFecUltEndeuda.Enabled = False
            txtFecUltEndeuda.Text = "__/__/____"
            txtMontoSol.SetFocus
        End If
    End If
End Sub

Public Function CalculaTotalMesVentasSem(ByVal pnVentaAlta As Double, ByVal pnDiasAlta As Integer, ByVal pnVentaBaja As Double, ByVal pnDiasBaja As Double) As Double
On Error GoTo ErrorCalculaTotalMesVentasSem

    CalculaTotalMesVentasSem = ((pnVentaAlta * pnDiasAlta) + (pnVentaBaja * pnDiasBaja)) * 4
    Exit Function
ErrorCalculaTotalMesVentasSem:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Public Function CalculaPorcRealVentasSem(ByVal pnCosto As Double, ByVal pnParticipacion As Double) As Double
On Error GoTo ErrorCalculaPorcRealVentasSem

    CalculaPorcRealVentasSem = ((pnCosto * pnParticipacion) / 100)
    Exit Function
ErrorCalculaPorcRealVentasSem:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Public Sub CalculaUtilidadNetaMensual()
On Error GoTo ErrorCalculaUtilidadNetaMensual
    Dim pnTEM As Double, pnMonto As Double
    Dim MontoMax As Double, CuotaEstima As Double
    Dim CuotaUNM As Double, CuotaExcedeFam As Double
    Dim pnCuotas As Integer
    Dim pnFormula As Double
    
    lblUtilNeta.Caption = Format(CDbl(lblVentaProm.Caption) - (CDbl(lblVentaProm.Caption) * (CDbl(lblCostoVenta.Caption) / 100)) - CDbl(lblTotalGastoNeg.Caption), "#,##0.00")
    
    pnMonto = CDbl(txtCalcMonto.Text)
    pnTEM = CDbl(txtCalcTEM.Text) / 100
    pnCuotas = CInt(spnCalcCuotas.valor)
    
    If pnTEM <> 0 And CDbl(lblUtilNeta.Caption) <> 0 And CDbl(lblExcedenteFam.Caption) <> 0 Then
        pnFormula = (((pnTEM * ((1 + pnTEM) ^ pnCuotas))) / (((1 + pnTEM) ^ pnCuotas) - 1))
        
        MontoMax = (CDbl(lblExcedenteFam.Caption) * lnIndMaximaCapPago) / pnFormula
        CuotaEstima = pnMonto * pnFormula
        CuotaUNM = (CuotaEstima / CDbl(lblUtilNeta.Caption)) * 100
        CuotaExcedeFam = (CuotaEstima / CDbl(lblExcedenteFam.Caption)) * 100
        
        lblMontoMax.Caption = Format(MontoMax, "#,##0.00")
        lblCuotaEstima.Caption = Format(CuotaEstima, "#,##0.00")
        lblCuotaUNM.Caption = Format(CuotaUNM, "#,##0.00")
        lblCuotaExcedeFam = Format(CuotaExcedeFam, "#,##0.00")
    End If
    Exit Sub
ErrorCalculaUtilidadNetaMensual:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Sub CalculaExcedenteFam()
On Error GoTo ErrorCalculaExcedenteFam
    Dim pnTEM As Double, pnMonto As Double
    Dim MontoMax As Double, CuotaEstima As Double
    Dim CuotaUNM As Double, CuotaExcedeFam As Double
    Dim pnCuotas As Integer
    Dim pnFormula As Double
    
    lblExcedenteFam.Caption = Format(CDbl(lblUtilNeta.Caption) + CDbl(lblTotalOtrosIng.Caption) - CDbl(lblTotalGastoFam.Caption), "#,##0.00")
    
    pnMonto = CDbl(txtCalcMonto.Text)
    pnTEM = CDbl(txtCalcTEM.Text) / 100
    pnCuotas = CInt(spnCalcCuotas.valor)
    
    If pnTEM <> 0 And CDbl(lblUtilNeta.Caption) <> 0 And CDbl(lblExcedenteFam.Caption) <> 0 Then
        pnFormula = (((pnTEM * ((1 + pnTEM) ^ pnCuotas))) / (((1 + pnTEM) ^ pnCuotas) - 1))
        
        MontoMax = (CDbl(lblExcedenteFam.Caption) * lnIndMaximaCapPago) / pnFormula
        CuotaEstima = pnMonto * pnFormula
        CuotaUNM = (CuotaEstima / CDbl(lblUtilNeta.Caption)) * 100
        CuotaExcedeFam = (CuotaEstima / CDbl(lblExcedenteFam.Caption)) * 100
        
        lblMontoMax.Caption = Format(MontoMax, "#,##0.00")
        lblCuotaEstima.Caption = Format(CuotaEstima, "#,##0.00")
        lblCuotaUNM.Caption = Format(CuotaUNM, "#,##0.00")
        lblCuotaExcedeFam = Format(CuotaExcedeFam, "#,##0.00")
    End If
    Exit Sub
ErrorCalculaExcedenteFam:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, J As Integer
    ValidaDatosReferencia = False
        
    If fgRef.Rows - 1 < 2 Then
        MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
        cmdAgregarRef.SetFocus
        ValidaDatosReferencia = False
        Exit Function
    End If
    
    'Verfica Tipo de Valores del DNI
    For i = 1 To fgRef.Rows - 1
        If Trim(fgRef.TextMatrix(i, 1)) <> "" Then
            For J = 1 To Len(Trim(fgRef.TextMatrix(i, 2)))
                If (Mid(fgRef.TextMatrix(i, 2), J, 1) < "0" Or Mid(fgRef.TextMatrix(i, 2), J, 1) > "9") Then
                   MsgBox "Uno de los Digitos del primer DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   fgRef.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next J
        End If
    Next i
    
    'Verfica Longitud del DNI
    For i = 1 To fgRef.Rows - 1
        If Trim(fgRef.TextMatrix(i, 1)) <> "" Then
            If Len(Trim(fgRef.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                fgRef.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    
    'Verfica Tipo de Valores del Telefono
    For i = 1 To fgRef.Rows - 1
        If Trim(fgRef.TextMatrix(i, 1)) <> "" Then
            For J = 1 To Len(Trim(fgRef.TextMatrix(i, 3)))
                If (Mid(fgRef.TextMatrix(i, 3), J, 1) < "0" Or Mid(fgRef.TextMatrix(i, 3), J, 1) > "9") Then
                   MsgBox "Uno de los Digitos del teléfono de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   fgRef.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next J
        End If
    Next i
    
    'Verfica Tipo de Valores del DNI 2
    For i = 1 To fgRef.Rows - 1
        If Trim(fgRef.TextMatrix(i, 1)) <> "" Then
            For J = 1 To Len(Trim(fgRef.TextMatrix(i, 5)))
                If (Mid(fgRef.TextMatrix(i, 5), J, 1) < "0" Or Mid(fgRef.TextMatrix(i, 5), J, 1) > "9") Then
                   MsgBox "Uno de los Digitos del segundo DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   fgRef.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next J
        End If
    Next i
    
    'Verfica Longitud del DNI 2
    For i = 1 To fgRef.Rows - 1
        If Trim(fgRef.TextMatrix(i, 1)) <> "" Then
            If Len(Trim(fgRef.TextMatrix(i, 5))) <> gnNroDigitosDNI Then
                MsgBox "Segundo DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                fgRef.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    
    'Verfica ambos DNI que no sean iguales
    For i = 1 To fgRef.Rows - 1
        If Trim(fgRef.TextMatrix(i, 1)) <> "" Then
            If Trim(fgRef.TextMatrix(i, 2)) = Trim(fgRef.TextMatrix(i, 5)) Then
                MsgBox "Los DNI de la fila " & fgRef.Row & " son iguales", vbInformation, "Aviso"
                fgRef.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    ValidaDatosReferencia = True
End Function

Public Function validaDatos() As Boolean
    validaDatos = False
If fnTipoPermiso = 3 Then
    If Not ValidaPorcParticip Then
        MsgBox "El % de Participacion no debe sumar mas del 100%", vbInformation, "Aviso"
        validaDatos = False
        Exit Function
    End If
    If Round(CDbl(lblMontoMax.Caption), 2) < Round(CDbl(txtCalcMonto.Text), 2) Then
        MsgBox "El Monto Máximo del Credito es menor al ingresado en el calculo", vbInformation, "Aviso"
        txtCalcMonto.SetFocus
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Round(CDbl(lblCuotaEstima.Caption), 2) > Round(CDbl(txtCuotaPagar.Text), 2) Then
        MsgBox "La Couta Estimada a Pagar es mayor a la Probable Cuota por Pagar", vbInformation, "Aviso"
        txtCuotaPagar.SetFocus
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(txtGiroNeg.Text) = "" Then
        MsgBox "Falta ingresar el Giro del Negocio", vbInformation, "Aviso"
        txtGiroNeg.SetFocus
        validaDatos = False
        Exit Function
    End If
    If Trim(txtGiroNeg.Text) = "" Then
        MsgBox "Falta ingresar el Giro del Negocio", vbInformation, "Aviso"
        txtGiroNeg.SetFocus
        validaDatos = False
        Exit Function
    End If
    If OptCondLocal(1).value = 0 And OptCondLocal(2).value = 0 And OptCondLocal(3).value = 0 And OptCondLocal(4).value = 0 Then
        MsgBox "Falta elegir la Condicion del local", vbInformation, "Aviso"
        OptCondLocal(1).SetFocus
        validaDatos = False
        Exit Function
    End If
    If OptCondLocal(4).value = 1 Then
        If Trim(txtCondLocalOtros.Text) = "" Then
            MsgBox "Falta detallar la Condicion del local", vbInformation, "Aviso"
            txtCondLocalOtros.SetFocus
            validaDatos = False
            Exit Function
        End If
    End If
    If txtCuotaPagar.value = 0 Then
        MsgBox "Falta ingresar la Probable cuota a pagar", vbInformation, "Aviso"
        txtCuotaPagar.SetFocus
        validaDatos = False
        Exit Function
    End If
    If spnCuotas.valor = 0 Then
        MsgBox "Falta ingresar el nro de cuotas", vbInformation, "Aviso"
        spnCuotas.SetFocus
        validaDatos = False
        Exit Function
    End If
    If txtUltEndeuda.value <> 0 Then
        If Trim(txtFecUltEndeuda.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha del ultimo endeudamiento", vbInformation, "Aviso"
            txtFecUltEndeuda.SetFocus
            validaDatos = False
            Exit Function
        End If
    End If
    If cboMontoSol.ListIndex = -1 Then
        MsgBox "Falta seleccionar la moneda", vbInformation, "Aviso"
        cboMontoSol.SetFocus
        validaDatos = False
        Exit Function
    End If
    If txtMontoSol.value = 0 Then
        MsgBox "Falta ingresar el monto solicitado", vbInformation, "Aviso"
        txtMontoSol.SetFocus
        validaDatos = False
        Exit Function
    End If
'    If lblVentaProm.Caption = 0 Then
'        MsgBox "Falta ingresar la Venta Promedio en mes", vbInformation, "Aviso"
'        txtVentaProm.SetFocus
'        SSTab2.Tab = 0
'        ValidaDatos = False
'        Exit Function
'    End If
'    If txtCostoVenta.value = 0 Then
'        MsgBox "Falta ingresar el costo de Venta", vbInformation, "Aviso"
'        txtCostoVenta.SetFocus
'        SSTab2.Tab = 0
'        ValidaDatos = False
'        Exit Function
'    End If
    If Trim(lblUtilNeta.Caption) = "" Then
        MsgBox "Faltan datos para el calculo de la Utilidad Neta", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(lblExcedenteFam.Caption) = "" Then
        MsgBox "Faltan datos para el calculo del Excedente Familiar", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(lblMontoMax.Caption) = "" Then
        MsgBox "Faltan datos para el calculo del Monto maximo del credito", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(lblCuotaEstima.Caption) = "" Then
        MsgBox "Faltan datos para el calculo de la cuota estimada", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(lblCuotaUNM.Caption) = "" Then
        MsgBox "Faltan datos para el calculo de la Cuota / Utilidad Neta", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(lblCuotaExcedeFam.Caption) = "" Then
        MsgBox "Faltan datos para el calculo de la Cuota / Excedente Familiar", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    If Trim(txtComent.Text) = "" Then
        MsgBox "Faltan ingresar el comentario", vbInformation, "Aviso"
        txtComent.SetFocus
        SSTab2.Tab = 1
        validaDatos = False
        Exit Function
    End If
    
    Dim i As Integer
    
    For i = 1 To fgVentasSem.Rows - 1
        If fgVentasSem.TextMatrix(i, 0) <> "" Then
            If Trim(fgVentasSem.TextMatrix(i, 1)) = "" Or Trim(fgVentasSem.TextMatrix(i, 2)) = "" _
                Or Trim(fgVentasSem.TextMatrix(i, 3)) = "" Or Trim(fgVentasSem.TextMatrix(i, 4)) = "" _
                Or Trim(fgVentasSem.TextMatrix(i, 5)) = "" Or Trim(fgVentasSem.TextMatrix(i, 6)) = "" _
                Or Trim(fgVentasSem.TextMatrix(i, 7)) = "" Or Trim(fgVentasSem.TextMatrix(i, 8)) = "" _
                Or Trim(fgVentasSem.TextMatrix(i, 9)) = "" Then
                MsgBox "Faltan datos en la lista de las Ventas de la Semana", vbInformation, "Aviso"
                SSTab2.Tab = 0
                validaDatos = False
                Exit Function
            End If
        End If
    Next i
    
    If fgVentasSem.Rows - 1 < 2 Then
        
    End If
    
    If ValidaGrillas(fgGastoNeg) = False Then
        MsgBox "Faltan datos en la lista de Gastos del Negocio", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    
    If ValidaGrillas(fgGastoFam) = False Then
        MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If
    
    If ValidaGrillas(fgOtrosIng) = False Then
        MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
        SSTab2.Tab = 0
        validaDatos = False
        Exit Function
    End If

    For i = 1 To fgRef.Rows - 1
        If fgRef.TextMatrix(i, 0) <> "" Then
            If Trim(fgRef.TextMatrix(i, 1)) = "" Or Trim(fgRef.TextMatrix(i, 2)) = "" Or Trim(fgRef.TextMatrix(i, 3)) = "" Or Trim(fgRef.TextMatrix(i, 4)) = "" Or Trim(fgRef.TextMatrix(i, 5)) = "" Then
                MsgBox "Faltan datos en la lista de Referencias", vbInformation, "Aviso"
                SSTab2.Tab = 1
                validaDatos = False
                Exit Function
            End If
        End If
    Next i
    
    If ValidaDatosReferencia = False Then
        SSTab2.Tab = 1
        validaDatos = False
        Exit Function
    End If

    ElseIf fnTipoPermiso = 2 Then
        If Trim(txtVerif.Text) = "" Then
            MsgBox "Favor de ingresar la Validación respectiva", vbInformation, "Aviso"
            txtVerif.SetFocus
            SSTab2.Tab = 1
            validaDatos = False
            Exit Function
        End If
    End If

    validaDatos = True
End Function

Public Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    ValidaGrillas = False
    For i = 1 To Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 2)) = "" Then
                ValidaGrillas = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrillas = True
End Function

Private Sub txtUltEndeuda_LostFocus()
    If txtUltEndeuda <> 0 Then
        txtFecUltEndeuda.Enabled = True
        'txtFecUltEndeuda.SetFocus
    Else
        txtFecUltEndeuda.Enabled = False
        txtFecUltEndeuda.Text = "__/__/____"
        'txtMontoSol.SetFocus
    End If
End Sub

Private Sub txtVerif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdGrabar.SetFocus
    End If
End Sub

'Private Sub txtVentaProm_LostFocus()
'    CalculaUtilidadNetaMensual
'    CalculaExcedenteFam
'End Sub

Private Sub txtVerif_LostFocus()
    txtVerif.Text = UCase(txtVerif.Text)
End Sub

Private Function LlenaRecordSet_VentasSemana() As ADODB.Recordset
Dim rsVentaSem As ADODB.Recordset
Dim i As Integer
Set rsVentaSem = New ADODB.Recordset

With rsVentaSem
    .fields.Append "cProducto", adVarChar, 200
    .fields.Append "nVentaAlta", adDouble
    .fields.Append "nDiasAlta", adSmallInt
    .fields.Append "nVentaBaja", adDouble
    .fields.Append "nDiasBaja", adSmallInt
    .fields.Append "nTotalMes", adDouble
    .fields.Append "nCosto", adDouble
    .fields.Append "nParticip", adDouble
    .fields.Append "nReal", adDouble
    .Open
    
    For i = 1 To fgVentasSem.Rows - 1
        .AddNew
        .fields("cProducto") = fgVentasSem.TextMatrix(i, 1)
        .fields("nVentaAlta") = Format(fgVentasSem.TextMatrix(i, 2), "#,##0.00")
        .fields("nDiasAlta") = fgVentasSem.TextMatrix(i, 3)
        .fields("nVentaBaja") = Format(fgVentasSem.TextMatrix(i, 4), "#,##0.00")
        .fields("nDiasBaja") = fgVentasSem.TextMatrix(i, 5)
        .fields("nTotalMes") = Format(fgVentasSem.TextMatrix(i, 6), "#,##0.00")
        .fields("nCosto") = Format(fgVentasSem.TextMatrix(i, 7), "#,##0.00")
        .fields("nParticip") = Format(fgVentasSem.TextMatrix(i, 8), "#,##0.00")
        .fields("nReal") = Format(fgVentasSem.TextMatrix(i, 9), "#,##0.00")
    Next i
End With

Set LlenaRecordSet_VentasSemana = rsVentaSem
End Function

Private Sub GeneraExcelFormato()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim FilaPertenece As Integer
    Dim pnCondLocalCol As Integer
    Dim nTEA As Double
    Dim nCuotaUNM As Double, nCuotaExdFam As Double
    Dim i As Integer
    Dim IniTablas As Integer, IniTablaOtroIng As Integer, FinTablas As Integer
    Dim CeldaVacia1 As Integer, CeldaVacia2 As Integer
    Dim CeldaVacia3 As Integer, CeldaVacia4 As Integer
    Dim CeldaVacia5 As Integer
    Dim Celda As String
    
    On Error GoTo ErrorGeneraExcelFormato
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    lsNomHoja = "FORMATO2"
    lsFile = "CredEvalFormato2"
    
    lsArchivo = "\spooler\" & "Evaluacion_" & sCtaCod & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    fsCliente = PstaNombre(fsCliente, True)
    nTEA = ((1 + (CDbl(txtCalcTEM.Text) / 100)) ^ 12) - 1
    nCuotaUNM = CDbl(lblCuotaUNM.Caption) / 100
    nCuotaExdFam = CDbl(lblCuotaExcedeFam.Caption) / 100
    
    xlHoja1.Cells(2, 2) = "FORMATO 2. EVALUACIÓN DE CRÉDITOS HASTA " & IIf(Mid(sCtaCod, 9, 1) = "1", Format(lnMax, "#,##0.00"), Format(lnMaxDol, "#,##0.00"))
    xlHoja1.Cells(4, 3) = fsCliente
    xlHoja1.Cells(4, 17) = fsUserAnalista
    xlHoja1.Cells(8, 3) = sCtaCod
    xlHoja1.Cells(8, 10) = txtGiroNeg.Text
    xlHoja1.Cells(10, 5) = spnExpEmpAnio.valor
    xlHoja1.Cells(10, 8) = spnExpEmpMes.valor
    xlHoja1.Cells(12, 5) = spnTiempoLocalAnio.valor
    xlHoja1.Cells(12, 8) = spnTiempoLocalMes.valor
    xlHoja1.Cells(10, 17) = txtUltEndeuda.Text
    xlHoja1.Cells(12, 17) = IIf(txtFecUltEndeuda.Text = "__/__/____", "", txtFecUltEndeuda.Text)
    Select Case lnCondLocal
    Case 1
        pnCondLocalCol = 5
    Case 2
        pnCondLocalCol = 8
    Case 3
        pnCondLocalCol = 11
    Case 4
        pnCondLocalCol = 14
    End Select
    xlHoja1.Cells(14, pnCondLocalCol) = "X"
    xlHoja1.Cells(14, 16) = IIf(lnCondLocal = 4, txtCondLocalOtros, "")
    xlHoja1.Cells(16, 5) = txtCuotaPagar.Text
    xlHoja1.Cells(16, 11) = spnCuotas.valor
    xlHoja1.Cells(16, 17) = txtMontoSol.Text
    xlHoja1.Cells(20, 6) = lblVentaProm.Caption
    xlHoja1.Cells(20, 16) = lblCostoVenta.Caption
    
    'Ventas Semana 24 a 48
    IniTablas = 23
    For i = 1 To fgVentasSem.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = i 'Nº
        xlHoja1.Cells(IniTablas + i, 3) = fgVentasSem.TextMatrix(i, 1) 'Producto
        xlHoja1.Cells(IniTablas + i, 6) = fgVentasSem.TextMatrix(i, 2) 'Venta alta
        xlHoja1.Cells(IniTablas + i, 8) = fgVentasSem.TextMatrix(i, 3) 'dias alta
        xlHoja1.Cells(IniTablas + i, 9) = fgVentasSem.TextMatrix(i, 4) 'Venta baja
        xlHoja1.Cells(IniTablas + i, 11) = fgVentasSem.TextMatrix(i, 5) 'dias baja
        xlHoja1.Cells(IniTablas + i, 12) = fgVentasSem.TextMatrix(i, 6) 'Total mes
        xlHoja1.Cells(IniTablas + i, 14) = fgVentasSem.TextMatrix(i, 7) '% Costo
        xlHoja1.Cells(IniTablas + i, 15) = fgVentasSem.TextMatrix(i, 8) '% Participacion
        xlHoja1.Cells(IniTablas + i, 17) = fgVentasSem.TextMatrix(i, 9) '% Real
    Next i
    CeldaVacia5 = IniTablas + i
    FinTablas = 48
    If CeldaVacia5 < FinTablas Then
        For i = CeldaVacia5 To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    'Gasto Neg y Gasto Fam 52 a 76
    IniTablas = 51
    For i = 1 To fgGastoNeg.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = fgGastoNeg.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 6) = fgGastoNeg.TextMatrix(i, 2)
    Next i
    CeldaVacia1 = IniTablas + i
    FinTablas = 76
    xlHoja1.Cells(77, 6) = Format(lblTotalGastoNeg.Caption, "#,##0.00")
    If CeldaVacia1 < FinTablas Then
        For i = CeldaVacia1 To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    IniTablas = 51
    For i = 1 To fgGastoFam.Rows - 1
        xlHoja1.Cells(IniTablas + i, 10) = fgGastoFam.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 16) = fgGastoFam.TextMatrix(i, 2)
    Next i
    CeldaVacia2 = IniTablas + i
    FinTablas = 76
    xlHoja1.Cells(77, 16) = Format(lblTotalGastoFam.Caption, "#,##0.00")
    If IIf(CeldaVacia1 > CeldaVacia2, CeldaVacia1, CeldaVacia2) < FinTablas Then
        For i = IIf(CeldaVacia1 > CeldaVacia2, CeldaVacia1, CeldaVacia2) To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    'Otros Ingresos 81 a 130
    IniTablas = 80
    IniTablaOtroIng = 91
    For i = 1 To fgOtrosIng.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = fgOtrosIng.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 6) = fgOtrosIng.TextMatrix(i, 2)
        i = i + 1
    Next i
    CeldaVacia3 = IniTablas + i
    FinTablas = 130
    xlHoja1.Cells(131, 6) = Format(lblTotalOtrosIng.Caption, "#,##0.00")
    If IIf(CeldaVacia3 > IniTablaOtroIng, CeldaVacia3, IniTablaOtroIng) < FinTablas Then
        For i = IIf(CeldaVacia3 > IniTablaOtroIng, CeldaVacia3, IniTablaOtroIng) To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    xlHoja1.Cells(80, 16) = lnIndMaximaCapPago
    xlHoja1.Cells(82, 16) = Format(lblUtilNeta.Caption, "#,##0.00")
    xlHoja1.Cells(84, 16) = Format(lblExcedenteFam.Caption, "#,##0.00")
    xlHoja1.Cells(86, 16) = Format(lblMontoMax.Caption, "#,##0.00")
    xlHoja1.Cells(88, 16) = CStr(nCuotaExdFam)
    xlHoja1.Cells(90, 16) = CStr(nCuotaUNM)
    
    xlHoja1.Cells(137, 2) = gdFecSis
    xlHoja1.Cells(137, 4) = IIf(Mid(sCtaCod, 9, 1) = "1", "SOLES", "DOLARES")
    xlHoja1.Cells(137, 5) = txtCalcMonto.Text
    xlHoja1.Cells(137, 10) = spnCalcCuotas.valor
    xlHoja1.Cells(137, 12) = CStr(CDbl(txtCalcTEM.Text) / 100)
    xlHoja1.Cells(137, 14) = nTEA
    xlHoja1.Cells(137, 16) = Format(lblCuotaEstima.Caption, "#,##0.00")
    
    xlHoja1.Cells(140, 5) = CStr(nCuotaExdFam)
    xlHoja1.Cells(140, 14) = CStr(nCuotaUNM)
    
    xlHoja1.Cells(143, 2) = txtComent.Text
    
    'Referencia 147 a 171
    IniTablas = 146
    For i = 1 To fgRef.Rows - 1
        xlHoja1.Cells(IniTablas + i, 2) = i
        xlHoja1.Cells(IniTablas + i, 3) = fgRef.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 7) = fgRef.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 9) = fgRef.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 11) = fgRef.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 16) = fgRef.TextMatrix(i, 5)
    Next i
    CeldaVacia4 = IniTablas + i
    FinTablas = 171
    If CeldaVacia4 < FinTablas Then
        For i = CeldaVacia4 To FinTablas
            Celda = "A" & i & ":A" & i
            xlHoja1.Range(Celda).RowHeight = 0
        Next i
    End If
    
    xlHoja1.Cells(177, 2) = txtVerif.Text
    
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    MsgBox "Fromato Generado Satisfactoriamente en la ruta: " & psArchivoAGrabarC, vbInformation, "Aviso"
    
    Exit Sub
ErrorGeneraExcelFormato:
    MsgBox Err.Description, vbInformation, "Error!!"
End Sub
