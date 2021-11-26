VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredConfigCatalogoProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Catálogo de Productos"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17235
   Icon            =   "frmCredConfigCatalogoProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   17235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8880
      TabIndex        =   54
      Top             =   10080
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabMain 
      Height          =   9810
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16920
      _ExtentX        =   29845
      _ExtentY        =   17304
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Producto"
      TabPicture(0)   =   "frmCredConfigCatalogoProd.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameRelacProd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameFiltro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkEditar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGrabarRelProd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Características"
      TabPicture(1)   =   "frmCredConfigCatalogoProd.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdbajarCondRequ"
      Tab(1).Control(1)=   "cmdSubirCondRequ"
      Tab(1).Control(2)=   "cmdBajar"
      Tab(1).Control(3)=   "cmdSubir"
      Tab(1).Control(4)=   "Frame7"
      Tab(1).Control(5)=   "frmSecRang"
      Tab(1).Control(6)=   "chkRequisitos"
      Tab(1).Control(7)=   "chkCondiciones"
      Tab(1).Control(8)=   "cmdElimRowFlex"
      Tab(1).Control(9)=   "chkIndep"
      Tab(1).Control(10)=   "chkMatriz"
      Tab(1).Control(11)=   "cmdGrabarDatos"
      Tab(1).Control(12)=   "cmdQuitarParam"
      Tab(1).Control(13)=   "cmdAsigParam"
      Tab(1).Control(14)=   "lstParamMain"
      Tab(1).Control(15)=   "flxRelacParam"
      Tab(1).Control(16)=   "MSHFlexCond"
      Tab(1).Control(17)=   "frmCondRequ"
      Tab(1).Control(18)=   "cmdAgregar"
      Tab(1).Control(19)=   "frmParam"
      Tab(1).Control(20)=   "Label20"
      Tab(1).Control(21)=   "Label19"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Check List"
      TabPicture(2)   =   "frmCredConfigCatalogoProd.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdGrabarcheckList"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).Control(4)=   "fraCondicionesCheckList"
      Tab(2).Control(5)=   "fra"
      Tab(2).ControlCount=   6
      Begin VB.Frame fra 
         Caption         =   "Documento Niveles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   95
         Top             =   4200
         Width           =   7455
         Begin VB.CommandButton cmdQuitarDocNivelesCheckList 
            Caption         =   "-"
            Height          =   255
            Left            =   6960
            TabIndex        =   97
            ToolTipText     =   "Quitar Niveles"
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregarNivCheckList 
            Caption         =   "+"
            Height          =   255
            Left            =   6480
            TabIndex        =   96
            ToolTipText     =   "Agregar Niveles"
            Top             =   2280
            Width           =   375
         End
         Begin SICMACT.FlexEdit feNivelesCheckList 
            Height          =   1935
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3413
            Cols0           =   7
            ScrollBars      =   2
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "n-N°-Descripcion-Conf.-Nivel-Tipo Doc.-aux"
            EncabezadosAnchos=   "0-350-4400-520-650-920-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-4-5-X"
            ListaControles  =   "0-0-0-0-0-3-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R-R-L-C"
            FormatosEdit    =   "0-0-0-3-3-1-1"
            TextArray0      =   "n"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraCondicionesCheckList 
         Caption         =   "Condiciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -67320
         TabIndex        =   83
         Top             =   3720
         Width           =   8535
         Begin VB.CommandButton cmdAgregarParametroCheckList 
            Caption         =   "<--"
            Height          =   275
            Left            =   5640
            TabIndex        =   93
            ToolTipText     =   "Agregar Condiciones"
            Top             =   480
            Width           =   380
         End
         Begin VB.CommandButton cmdQuitarParametroCheckList 
            Caption         =   "-"
            Height          =   275
            Left            =   5640
            TabIndex        =   92
            ToolTipText     =   "Quitar Condiciones"
            Top             =   960
            Width           =   380
         End
         Begin VB.Frame fraMontosCheckList 
            Caption         =   "Montos"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            TabIndex        =   85
            Top             =   2400
            Width           =   5415
            Begin VB.CommandButton cmdAgregarMontoCheckList 
               Caption         =   "+"
               Height          =   255
               Left            =   4920
               TabIndex        =   91
               ToolTipText     =   "Agregar"
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtFinCheckList 
               Height          =   285
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   90
               Top             =   240
               Width           =   960
            End
            Begin VB.TextBox txtInicioCheckList 
               Height          =   285
               Left            =   840
               MaxLength       =   10
               TabIndex        =   89
               Top             =   240
               Width           =   960
            End
            Begin VB.ComboBox cmbIniCheckList 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox cmbFinCheckList 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   87
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox cmbUniMedCheckList 
               Height          =   315
               Left            =   3720
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.ListBox ListParametroCheckList 
            Height          =   2010
            Left            =   6120
            TabIndex        =   84
            Top             =   240
            Width           =   2295
         End
         Begin SICMACT.FlexEdit feCondicionesCheckList 
            Height          =   2055
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   3625
            Cols0           =   12
            ScrollBars      =   2
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Parámetro-Id_Param-Descripción-Id_Descrip-OpeIni-MontoIni-OpeFin-MontoFin-UnidaMedida-nCantConf-aux"
            EncabezadosAnchos=   "350-2000-0-2650-0-0-0-0-0-0-0-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-3-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-3-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
            CantEntero      =   20
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -67320
         TabIndex        =   79
         Top             =   1080
         Width           =   8535
         Begin VB.CommandButton cmdQuitarDet 
            Caption         =   "-"
            Height          =   255
            Left            =   8040
            TabIndex        =   81
            ToolTipText     =   "Quitar"
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregarDet 
            Caption         =   "+"
            Height          =   255
            Left            =   7560
            TabIndex        =   80
            ToolTipText     =   "Agregar"
            Top             =   2280
            Width           =   375
         End
         Begin SICMACT.FlexEdit feDetalle 
            Height          =   1935
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   3413
            Cols0           =   5
            ScrollBars      =   2
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "n-N°-Descripcion--aux"
            EncabezadosAnchos=   "0-350-6900-700-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-4-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0"
            TextArray0      =   "n"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Documentos Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74880
         TabIndex        =   72
         Top             =   1080
         Width           =   7455
         Begin VB.Frame Frame6 
            Caption         =   "Agregar Conf."
            Height          =   580
            Left            =   120
            TabIndex        =   76
            Top             =   2520
            Width           =   1215
            Begin VB.CommandButton cmdAgregarConfDoc 
               Caption         =   "+"
               Height          =   255
               Left            =   360
               TabIndex        =   77
               ToolTipText     =   "Agregar Cant. Conf."
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Doc. Principal"
            Height          =   580
            Left            =   5760
            TabIndex        =   73
            Top             =   2500
            Width           =   1575
            Begin VB.CommandButton cmdAgregarDoc 
               Caption         =   "+"
               Height          =   255
               Left            =   240
               TabIndex        =   75
               ToolTipText     =   "Agregar Documento"
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton cmdQuitarDoc 
               Caption         =   "-"
               Height          =   255
               Left            =   960
               TabIndex        =   74
               ToolTipText     =   "Quitar Documento"
               Top             =   240
               Width           =   375
            End
         End
         Begin SICMACT.FlexEdit feDocumentos 
            Height          =   2295
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   4048
            Cols0           =   5
            ScrollBars      =   2
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "n-N°-Descripcion-Conf.-aux"
            EncabezadosAnchos=   "0-350-5500-800-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-3-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-R-R"
            FormatosEdit    =   "0-0-0-3-3"
            TextArray0      =   "n"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   7455
         Begin VB.ComboBox cmbCatCheckLis 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   240
            Width           =   2350
         End
         Begin VB.ComboBox cmbProdCheckLis 
            Height          =   315
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   240
            Width           =   2350
         End
         Begin VB.Label Label7 
            Caption         =   "Producto: "
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
            Left            =   3600
            TabIndex        =   71
            Top             =   280
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Categoria:"
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
            TabIndex        =   70
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdbajarCondRequ 
         Caption         =   "Bajar"
         Height          =   375
         Left            =   -63800
         TabIndex        =   65
         Top             =   5280
         Width           =   825
      End
      Begin VB.CommandButton cmdSubirCondRequ 
         Caption         =   "Subir"
         Height          =   375
         Left            =   -63800
         TabIndex        =   64
         Top             =   4800
         Width           =   825
      End
      Begin VB.CommandButton cmdBajar 
         Caption         =   "Bajar"
         Height          =   335
         Left            =   -63742
         TabIndex        =   63
         Top             =   2280
         Width           =   773
      End
      Begin VB.CommandButton cmdSubir 
         Caption         =   "Subir"
         Height          =   335
         Left            =   -63742
         TabIndex        =   62
         Top             =   1800
         Width           =   773
      End
      Begin VB.Frame Frame7 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74760
         TabIndex        =   57
         Top             =   480
         Width           =   11945
         Begin VB.ComboBox cboTpoProdCar 
            Height          =   315
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
            Width           =   2350
         End
         Begin VB.ComboBox cboCatCar 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   240
            Width           =   2350
         End
         Begin VB.Label Label9 
            Caption         =   "Categoria:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   61
            Top             =   285
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Producto: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   60
            Top             =   285
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdGrabarcheckList 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   -59880
         TabIndex        =   55
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Frame frmSecRang 
         Caption         =   "Sección Rangos"
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
         Height          =   855
         Left            =   -74760
         TabIndex        =   25
         Top             =   8280
         Width           =   11960
         Begin VB.CheckBox chkHabRang 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1560
            TabIndex        =   53
            Top             =   0
            Width           =   255
         End
         Begin VB.ComboBox cboUndRang 
            Height          =   315
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   330
            Width           =   1095
         End
         Begin VB.ComboBox cboTpoDato 
            Height          =   315
            Left            =   10680
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdRang 
            Caption         =   "+"
            Height          =   255
            Left            =   7800
            TabIndex        =   37
            Top             =   360
            Width           =   375
         End
         Begin VB.ComboBox cboRang2 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   330
            Width           =   615
         End
         Begin VB.ComboBox cboRang1 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   330
            Width           =   615
         End
         Begin VB.TextBox txtRangMin 
            Height          =   285
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   31
            Top             =   360
            Width           =   960
         End
         Begin VB.TextBox txtRangMax 
            Height          =   285
            Left            =   5400
            MaxLength       =   10
            TabIndex        =   33
            Top             =   360
            Width           =   960
         End
         Begin VB.ComboBox cboTpoParam 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   315
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Tipo Dato"
            Height          =   255
            Left            =   9840
            TabIndex        =   27
            Top             =   35
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Parámetro"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.CheckBox chkRequisitos 
         Caption         =   "Requisitos"
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
         Left            =   -73200
         TabIndex        =   50
         Top             =   4005
         Width           =   1215
      End
      Begin VB.CheckBox chkCondiciones 
         Caption         =   "Condiciones"
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
         Left            =   -74640
         TabIndex        =   49
         Top             =   3975
         Width           =   1455
      End
      Begin VB.CommandButton cmdElimRowFlex 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   -63800
         TabIndex        =   48
         Top             =   5760
         Width           =   825
      End
      Begin VB.CheckBox chkIndep 
         Caption         =   "Independiente"
         Height          =   255
         Left            =   -66600
         TabIndex        =   46
         Top             =   1515
         Width           =   1335
      End
      Begin VB.CheckBox chkMatriz 
         Caption         =   "Matriz"
         Height          =   240
         Left            =   -68040
         TabIndex        =   45
         Top             =   1515
         Width           =   735
      End
      Begin VB.CommandButton cmdGrabarDatos 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   -69480
         TabIndex        =   44
         Top             =   9240
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitarParam 
         Caption         =   "--"
         Height          =   275
         Left            =   -71160
         TabIndex        =   38
         Top             =   2640
         Width           =   380
      End
      Begin VB.CommandButton cmdAsigParam 
         Caption         =   "-->"
         Height          =   275
         Left            =   -71160
         TabIndex        =   36
         Top             =   2280
         Width           =   380
      End
      Begin VB.ListBox lstParamMain 
         Height          =   2010
         Left            =   -74760
         TabIndex        =   35
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton cmdGrabarRelProd 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox chkEditar 
         Caption         =   "Editar"
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
         Left            =   600
         TabIndex        =   1
         Top             =   1560
         Width           =   855
      End
      Begin VB.Frame frameFiltro 
         Caption         =   "Filtrar Registros"
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
         Height          =   855
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   11895
         Begin VB.ComboBox cboProdCab 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   320
            Width           =   2350
         End
         Begin VB.ComboBox cboTpoProd 
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   320
            Width           =   2350
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "Nuevo"
            Height          =   375
            Left            =   9600
            TabIndex        =   19
            Top             =   280
            Width           =   1215
         End
         Begin VB.CommandButton cmdMostrar 
            Caption         =   "Mostrar"
            Height          =   375
            Left            =   8160
            TabIndex        =   18
            Top             =   280
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Categoría"
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
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   4200
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frameRelacProd 
         Height          =   3015
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   12675
         Begin VB.CommandButton cmdHabilitar 
            Caption         =   "Habilitar"
            Height          =   275
            Left            =   2985
            TabIndex        =   56
            Top             =   2665
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkNoVig 
            Caption         =   "No Vig."
            Height          =   255
            Left            =   3720
            TabIndex        =   52
            Top             =   355
            Width           =   855
         End
         Begin VB.ListBox LstProdCab 
            Height          =   2010
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton CmdQuitAsign 
            Caption         =   "--"
            Height          =   275
            Left            =   4755
            TabIndex        =   10
            Top             =   1650
            Width           =   380
         End
         Begin VB.CommandButton CmdAsigSubProd 
            Caption         =   "-->"
            Height          =   275
            Left            =   4755
            TabIndex        =   9
            Top             =   1200
            Width           =   380
         End
         Begin VB.ListBox LstTpoProd 
            Height          =   2010
            Left            =   2400
            TabIndex        =   8
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton cmdNuevoProd 
            Caption         =   "<--"
            Height          =   275
            Left            =   8640
            TabIndex        =   7
            Top             =   1250
            Width           =   380
         End
         Begin VB.CommandButton CmdAll 
            Caption         =   "<="
            Height          =   275
            Left            =   8640
            TabIndex        =   6
            Top             =   1700
            Width           =   380
         End
         Begin VB.CommandButton cboAdd 
            Caption         =   "+"
            Height          =   275
            Left            =   12105
            TabIndex        =   4
            Top             =   1250
            Width           =   380
         End
         Begin VB.CommandButton cboElim 
            Caption         =   "-"
            Height          =   275
            Left            =   12105
            TabIndex        =   3
            Top             =   1700
            Width           =   380
         End
         Begin SICMACT.FlexEdit flxNewProd 
            Height          =   2010
            Left            =   9195
            TabIndex        =   5
            Top             =   600
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   3545
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Nuevo Producto-Id_NuevoProd"
            EncabezadosAnchos=   "320-2000-800"
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
            ColumnasAEditar =   "X-1-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   315
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin MSComctlLib.ListView LstRelacion 
            Height          =   2010
            Left            =   5400
            TabIndex        =   12
            Top             =   600
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   3545
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Id_ProdAnt"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Producto Anterior"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Id_NewProd"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Nuevo Producto"
               Object.Width           =   2822
            EndProperty
         End
         Begin VB.Label lblCat 
            Caption         =   "Categoría"
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
            Left            =   250
            TabIndex        =   16
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label3 
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
            Height          =   255
            Left            =   2415
            TabIndex        =   15
            Top             =   345
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Relación"
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
            Left            =   5425
            TabIndex        =   14
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Nuevo Producto"
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
            Left            =   9240
            TabIndex        =   13
            Top             =   345
            Width           =   1695
         End
      End
      Begin SICMACT.FlexEdit flxRelacParam 
         Height          =   1695
         Left            =   -70440
         TabIndex        =   42
         Top             =   1800
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2990
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Parámetro-Id_Param-Descripción-Id_Descrip-Id_Rango-psOperIni-psRangIni-psOperFin-psRangFin-psUnidRang-Módulo-IdModulo"
         EncabezadosAnchos=   "350-2000-0-2950-0-0-0-0-0-0-0-1200-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X-X-X-X-X-X-X-11-X"
         ListaControles  =   "0-0-0-3-0-0-0-0-0-0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C-C-C-C-C-C-C-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         CantEntero      =   20
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexCond 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   41
         Top             =   4320
         Width           =   10670
         _ExtentX        =   18812
         _ExtentY        =   6376
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame frmCondRequ 
         Height          =   4200
         Left            =   -74760
         TabIndex        =   51
         Top             =   3960
         Width           =   11955
         Begin VB.CommandButton cmdEditar 
            Caption         =   "Editar"
            Height          =   375
            Left            =   10970
            TabIndex        =   66
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   335
         Left            =   -63742
         TabIndex        =   43
         Top             =   2760
         Width           =   773
      End
      Begin VB.Frame frmParam 
         Height          =   2055
         Left            =   -70560
         TabIndex        =   47
         Top             =   1560
         Width           =   7740
      End
      Begin VB.Label Label20 
         Caption         =   "Relación"
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
         Left            =   -70560
         TabIndex        =   40
         Top             =   1365
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Campos"
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   1300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCredConfigCatalogoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmCredConfigCatalogoProd
'*** Descripción : Formulario para realizar la configuración de Catálogo de Productos
'*** Creación : NAGL el 20180712
'********************************************************************************
Option Explicit 'JOEP

Dim psTpoParamIni As String
Dim pMatDocumentos As Variant 'JOEP
Dim pMatDetalle As Variant 'JOEP
Dim pMatCondicion As Variant 'JOEP
Dim nFilaSel As Integer
Dim psTipSelec As String
Dim pbEditarGeneral As Boolean

'******Para guardar las variable Iniciales, para la opción editar
Dim psIdParamAnt() As String
Dim psParametroAnt() As String
Dim psRangDescAnt() As String
Dim psOperIniAnt() As String
Dim psRangIniAnt() As String
Dim psOperFinAnt() As String
Dim psRangFinAnt() As String
Dim psUnidRangAnt() As String
Dim psValorUnidRAnt() As String
Dim psIdModuloAnt() As String
Dim psModuloDescAnt() As String
Dim psMovNroAnt() As String
Dim nOrd As Integer
'***************************************************************

Public Sub Inicio()
    Call CriterioProducto
    Call ControlesVisible("Prod")
    Call CargaListaParametrosMain
    Call CargaCabeceraMSHFlexCond
    chkCondiciones.value = 1
    pbEditarGeneral = False
    CentraForm Me
    Me.Show 1
End Sub

Public Sub CriterioProducto()
    
    Dim rsCat As New ADODB.Recordset
    Dim HeightIni As Integer
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
    Set rsCat = oCredCat.CargarTipoProd("Cab", "")
    Call LlenarComboRS(rsCat, cboProdCab) '->Producto
    Set rsCat = Nothing
    Set rsCat = oCredCat.CargarTipoProd("Cab", "")
    Call LlenarComboRS(rsCat, cboCatCar) '->Condiciones
    Call CambiaTamañoCombo(cboCatCar, 300)
    Set rsCat = Nothing
'JOEP20181204 ERS034-CPp
    Set rsCat = oCredCat.CargarTipoProd("Cab", "")
    Call LlenarComboRS(rsCat, cmbCatCheckLis)  '->Check List
    Call CambiaTamañoCombo(cmbCatCheckLis, 300)
'JOEP20181204 ERS034-CP
    
    cboProdCab.ListIndex = 0
    cboTpoProd.ListIndex = 0
    Set rsCat = Nothing
    FormateaFlex flxNewProd
    LstProdCab.Enabled = False
    LstTpoProd.Enabled = False
    LstRelacion.Enabled = False
    flxNewProd.Enabled = False
    CmdAsigSubProd.Enabled = False
    CmdQuitAsign.Enabled = False
    cmdNuevoProd.Enabled = False
    CmdAll.Enabled = False
    cboAdd.Enabled = False
    cboElim.Enabled = False
    cmdGrabarRelProd.Enabled = False
    chkEditar.value = 0
    chkEditar.Enabled = False
    chkNoVig.Enabled = False
'JOEP
Set oCredCat = Nothing
RSClose rsCat
'JOEP
End Sub

Public Sub ControlesVisible(SSTAB As String)
    If SSTAB = "Prod" Then
        SSTabMain.Height = 5200
        SSTabMain.Width = 13215
        Width = 13710
        Height = 6350
        cmdSalir.Left = 12120
        cmdSalir.Top = 5450
    ElseIf SSTAB = "Cond" Then
        chkIndep.value = 1
        SSTabMain.Height = 9750
        SSTabMain.Width = 12350 '11150
        Height = 10850 '9840
        Width = 12680 '11500
        cmdSalir.Left = 11250 '120
        cmdSalir.Top = 10000 '8760
    End If
End Sub

Private Sub cboProdCab_Click()
Dim rsProd As New ADODB.Recordset
Dim psRelac As String
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

If cboProdCab.Text <> "" Then
    psRelac = Mid(Trim(cboProdCab.Text), Len(Trim(cboProdCab.Text)) - 2, 1)
    Set rsProd = oCredCat.CargarTipoProd("Prod", psRelac)
    Call LlenarComboRS(rsProd, cboTpoProd)
    cboTpoProd.ListIndex = 0
End If
'JOEP
Set oCredCat = Nothing
RSClose rsProd
'JOEP
End Sub

Private Sub cboCatCar_Click()
Dim rsProd As New ADODB.Recordset
Dim psRelac As String
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
    psRelac = Mid(Trim(cboCatCar.Text), Len(Trim(cboCatCar.Text)) - 2, 1)
    Set rsProd = oCredCat.CargarTipoProd("Prod", psRelac)
    Call LlenarComboRS(rsProd, cboTpoProdCar)
    Call CambiaTamañoCombo(cboTpoProdCar, 300)
    cboTpoProdCar.ListIndex = 0
'JOEP
Set oCredCat = Nothing
RSClose rsProd
'JOEP
End Sub

Private Sub cboTpoProdCar_Click()
FormateaFlex flxRelacParam
If chkCondiciones = 1 Then
    Call CargarListadoCondicionesRequisitos("Cond")
Else
    Call CargarListadoCondicionesRequisitos("Req")
End If
End Sub

Public Sub LlenarComboRS(pRs As ADODB.Recordset, psObjeto As ComboBox)
psObjeto.Clear
Do While Not pRs.EOF
    psObjeto.AddItem Trim(pRs!Descrip) & Space(100) & Trim(pRs!Tipo)
    pRs.MoveNext
Loop
pRs.Close
End Sub

Private Sub chkEditar_Click()
If chkEditar.value = 1 Then
   cboProdCab.Enabled = False
   cboTpoProd.Enabled = False
   flxNewProd.lbEditarFlex = True
   CmdAsigSubProd.Enabled = True
   CmdQuitAsign.Enabled = True
   cmdNuevoProd.Enabled = True
   CmdAll.Enabled = True
   cboAdd.Enabled = True
   cboElim.Enabled = True
   chkNoVig.Enabled = True
   cmdGrabarRelProd.Enabled = True
   cmdMostrar.Caption = "Filtrar"
Else
   cboProdCab.Enabled = True
   cboTpoProd.Enabled = True
   flxNewProd.lbEditarFlex = False
   CmdAsigSubProd.Enabled = False
   CmdQuitAsign.Enabled = False
   cmdNuevoProd.Enabled = False
   CmdAll.Enabled = False
   cboAdd.Enabled = False
   cboElim.Enabled = False
   chkNoVig.Enabled = False
   cmdGrabarRelProd.Enabled = False
   cmdMostrar.Caption = "Mostrar"
End If
End Sub

Private Sub chkNoVig_Click()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim psRelacProd As String
Dim Item As ListItem
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

i = 1
If LstProdCab.Text <> "" Then
    psRelacProd = Mid(Trim(LstProdCab.Text), Len(Trim(LstProdCab.Text)) - 2, 1)
Else
    If psTipSelec = "Mostrar" Or LstTpoProd.ListCount > 1 Then
        psRelacProd = Mid(Trim(cboProdCab.Text), Len(Trim(cboProdCab.Text)) - 2, 1)
    Else
        MsgBox "Por favor seleccione alguna categoría", vbInformation, "Aviso"
        LstProdCab.SetFocus
        Exit Sub
    End If
End If
If chkNoVig.value = 1 Then
    LstTpoProd.Clear
    If LstTpoProd.ListIndex = -1 Then
        Set rs = oCredCat.CargarTipoProd("Prod", psRelacProd, "NoVig")
        Set oCredCat = Nothing
        Do While Not rs.EOF
            LstTpoProd.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
            rs.MoveNext
        Loop
    End If
    cmdHabilitar.Visible = True
Else
    LstTpoProd.Clear
    If LstTpoProd.ListIndex = -1 Then
        Set rs = oCredCat.CargarTipoProd("Prod", psRelacProd)
        Set oCredCat = Nothing
        Do While Not rs.EOF
            LstTpoProd.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
            rs.MoveNext
        Loop
    End If
    cmdHabilitar.Visible = False
End If

'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub cmdHabilitar_Click()
Dim psTpoProd
Dim psHab As String
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim psRelacProd As String
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

If LstTpoProd.Text <> "" Then
    psTpoProd = Mid(Trim(LstTpoProd.Text), Len(Trim(LstTpoProd.Text)) - 2, 3)
    If LstProdCab.Text <> "" Then
        psRelacProd = Mid(Trim(LstProdCab.Text), Len(Trim(LstProdCab.Text)) - 2, 1)
    Else
        psRelacProd = Mid(Trim(cboProdCab.Text), Len(Trim(cboProdCab.Text)) - 2, 1)
    End If
    psHab = oCredCat.HabilitarProducto(psTpoProd)
    If psHab = "SI" Then
        MsgBox "Se habilitó el Producto " & """" & Mid(Trim(LstTpoProd.Text), 1, Len(Trim(LstTpoProd.Text)) - 103) & """" & " satisfactoriamente", vbInformation, "Aviso"
        If MsgBox("¿Desea habilitar otro Producto..", vbYesNo + vbQuestion, "Atención") = vbYes Then
            chkNoVig.value = 1
            LstTpoProd.Clear
            If LstTpoProd.ListIndex = -1 Then
                Set rs = oCredCat.CargarTipoProd("Prod", psRelacProd, "NoVig")
                Set oCredCat = Nothing
                Do While Not rs.EOF
                    LstTpoProd.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
                    rs.MoveNext
                Loop
            End If
            cmdHabilitar.Visible = True
        Else
            chkNoVig.value = 0
        End If
            FormateaFlex flxNewProd
            i = 1
            Set oCredCat = New COMDCredito.DCOMCatalogoProd
            Set rs = oCredCat.ObtieneRelacProdEquiv(psRelacProd, "ProdNew")
            Do While Not rs.EOF
                flxNewProd.TextMatrix(i, 1) = rs!cConsDescripNew
                flxNewProd.TextMatrix(i, 2) = rs!cTpoProdNew
                i = i + 1
                flxNewProd.AdicionaFila
                rs.MoveNext
            Loop
            flxNewProd.EliminaFila (i)
            Set rs = Nothing
    Else
        MsgBox "El producto no puede ser habilitado, debido a que pertenece a una migración vigente.", vbInformation + vbOKOnly, "Aviso"
        Exit Sub
    End If
Else
    MsgBox "Por favor seleccione algún Producto", vbInformation, "Aviso"
    LstTpoProd.SetFocus
    Exit Sub
End If
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub cmdMostrar_Click()
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim psRelacProd As String
Dim Item As ListItem
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

If cmdMostrar.Caption = "Filtrar" Then
    cboProdCab.ListIndex = 0
    cboTpoProd.ListIndex = 0
    cboProdCab.Enabled = True
    cboTpoProd.Enabled = True
    FormateaFlex flxNewProd
    LstProdCab.Clear
    LstTpoProd.Clear
    LstRelacion.ListItems.Clear
    LstProdCab.Enabled = False
    LstTpoProd.Enabled = False
    LstRelacion.Enabled = False
    flxNewProd.Enabled = False
    CmdAsigSubProd.Enabled = False
    CmdQuitAsign.Enabled = False
    cmdNuevoProd.Enabled = False
    CmdAll.Enabled = False
    cboAdd.Enabled = False
    cboElim.Enabled = False
    cmdGrabarRelProd.Enabled = False
    chkEditar.value = 0
    chkEditar.Enabled = False
    chkNoVig.Enabled = False
    cmdMostrar.Caption = "Mostrar"
Else
    If cboProdCab.Text = "" Then
       MsgBox "Por favor seleccione alguna Categoría", vbInformation, "Aviso"
       Exit Sub
    Else
        i = 1
        psRelacProd = Mid(Trim(cboProdCab.Text), Len(Trim(cboProdCab.Text)) - 2, 1)
        Call CargarRelacionProductos
        FormateaFlex flxNewProd
        Set rs = oCredCat.ObtieneRelacProdEquiv(psRelacProd, "ProdNew")
        Do While Not rs.EOF
            flxNewProd.TextMatrix(i, 1) = rs!cConsDescripNew
            flxNewProd.TextMatrix(i, 2) = rs!cTpoProdNew
            i = i + 1
            flxNewProd.AdicionaFila
            rs.MoveNext
        Loop
        Set rs = Nothing
        chkCondiciones.value = 1
        cboCatCar.Text = cboProdCab.Text
        cboTpoProdCar.Text = cboTpoProd.Text
        flxNewProd.EliminaFila (i)
        chkEditar.value = 0
        chkEditar.Enabled = True
        psTipSelec = "Mostrar"
        chkNoVig.value = 0
        psTipSelec = ""
        LstProdCab.Enabled = True
        LstTpoProd.Enabled = True
        LstRelacion.Enabled = True
        flxNewProd.Enabled = True
    End If
End If

'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Public Sub CargarRelacionProductos(Optional psOpt As String)
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim psRelacProd As String
Dim Item As ListItem
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
'JOEP
i = 1
If psOpt = "" Then
    Call LimpiarControlesRelacProd(1)
    LstProdCab.AddItem cboProdCab.Text
End If

If cboProdCab.Text <> "" Then
    psRelacProd = Mid(Trim(cboProdCab.Text), Len(Trim(cboProdCab.Text)) - 2, 1)
Else
    psRelacProd = Mid(Trim(LstProdCab.Text), Len(Trim(LstProdCab.Text)) - 2, 1)
End If

Set oCredCat = New COMDCredito.DCOMCatalogoProd
LstTpoProd.Clear
If LstTpoProd.ListIndex = -1 Then
    Set rs = oCredCat.CargarTipoProd("Prod", psRelacProd)
    Set oCredCat = Nothing
    Do While Not rs.EOF
        LstTpoProd.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
        rs.MoveNext
    Loop
End If
Set rs = Nothing

Set oCredCat = New COMDCredito.DCOMCatalogoProd
LstRelacion.ListItems.Clear
Set rs = oCredCat.ObtieneRelacProdEquiv(psRelacProd)
Do While Not rs.EOF
    Set Item = LstRelacion.ListItems.Add(, , LstTpoProd.List(LstTpoProd.ListIndex))
    Item.SubItems(1) = rs!cTpoProdAnt
    Item.SubItems(2) = rs!cConsDescrip
    Item.SubItems(3) = rs!cTpoProdNew
    Item.SubItems(4) = rs!cConsDescripNew
    rs.MoveNext
Loop

'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Public Sub LimpiarControlesRelacProd(Optional OptLimpCab As Integer = 0)
    If OptLimpCab = 1 Then
        'LstProdCab.Clear 'No incluye la cabecera limpia
        LstTpoProd.Clear
        LstRelacion.ListItems.Clear
    Else
        LstTpoProd.Clear
        LstRelacion.ListItems.Clear
        FormateaFlex flxNewProd
    End If
End Sub

Private Sub cmdNuevo_Click()
Call CargarListProdCab
Call LimpiarControlesRelacProd
cboProdCab.ListIndex = -1
cboProdCab.Enabled = False
cboTpoProd.ListIndex = -1
cboTpoProd.Enabled = False
cmdMostrar.Caption = "Filtrar"

chkEditar.value = 0
chkEditar.Enabled = False
LstProdCab.Enabled = True
LstTpoProd.Enabled = True
LstRelacion.Enabled = True
flxNewProd.Enabled = True
flxNewProd.lbEditarFlex = True

'*******************************
CmdAsigSubProd.Enabled = False
CmdQuitAsign.Enabled = False
cmdNuevoProd.Enabled = False
CmdAll.Enabled = False
cboAdd.Enabled = False
cboElim.Enabled = False
chkNoVig.Enabled = False
cmdGrabarRelProd.Enabled = False
'** NAGL 20190926 Cambió de True a False***
End Sub

Private Sub cmdNuevoProd_Click()
Dim Item As ListItem
Dim nItemProd As Integer
Dim i As Integer
Dim nCant As Integer
Dim nPos As Integer

nItemProd = flxNewProd.row
i = 0

If flxNewProd.rows - 1 >= 1 Then
    For i = 1 To LstRelacion.ListItems.count
     If LstRelacion.ListItems(i).SubItems(3) = "" Then
        nPos = i
        Exit For
     End If
    Next i
    If nPos > 0 Then
       LstRelacion.ListItems(nPos).SubItems(3) = CStr(Trim(flxNewProd.TextMatrix(nItemProd, 2)))
       LstRelacion.ListItems(nPos).SubItems(4) = CStr(Trim(flxNewProd.TextMatrix(nItemProd, 1)))
    Else
       MsgBox "Debe agregar un Tipo de Producto Anterior", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub CmdAll_Click()
Dim Item As ListItem
Dim nItemProd As Integer
Dim i As Integer
Dim nCant As Integer
Dim nPos As Integer
i = 0
nItemProd = flxNewProd.row
If LstRelacion.ListItems.count <> 0 Then
    If LstRelacion.ListItems(1).SubItems(1) <> "" Then
        If flxNewProd.rows - 1 >= 1 Then
            For i = 1 To LstRelacion.ListItems.count
             If LstRelacion.ListItems(i).SubItems(3) = "" Then
                nPos = i
                LstRelacion.ListItems(nPos).SubItems(3) = CStr(Trim(flxNewProd.TextMatrix(nItemProd, 2)))
                LstRelacion.ListItems(nPos).SubItems(4) = CStr(Trim(flxNewProd.TextMatrix(nItemProd, 1)))
             End If
            Next i
        End If
    Else
        MsgBox "Debe agregar un Tipo de Producto Anterior para la migración", vbInformation, "Aviso"
    End If
Else
    MsgBox "Debe agregar un Tipo de Producto Anterior para la migración", vbInformation, "Aviso"
End If
End Sub

Private Sub CmdAsigSubProd_Click()
Dim Item As ListItem
  If LstTpoProd.ListIndex <> -1 Then
    Set Item = LstRelacion.ListItems.Add(, , LstTpoProd.List(LstTpoProd.ListIndex))
    Item.SubItems(1) = Trim(Mid(Trim(LstTpoProd.Text), Len(Trim(LstTpoProd.Text)) - 2, 3))
    Item.SubItems(2) = Trim(Mid(Trim(LstTpoProd.Text), 1, Len(Trim(LstTpoProd.Text)) - 3))
  Else
       MsgBox "Debe seleccionar una Tipo de Producto", vbInformation, "AVISO"
  End If
End Sub

Public Sub CargarListProdCab()
    Dim rs As New ADODB.Recordset
    'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
    LstProdCab.Clear
    Set rs = oCredCat.CargarTipoProd("Cab", "")
    Set oCredCat = Nothing
    Do While Not rs.EOF
        LstProdCab.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
        rs.MoveNext
    Loop
    'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Function DevuelvePosMatriz(ByVal iParam As Integer) As Integer
Dim i As Integer
Dim nItem As Integer
nItem = 0
    For i = 1 To CInt(flxNewProd.rows) - 1 'Para Controlar Productos Repetidos
            If Trim(flxNewProd.TextMatrix(i, 2)) = CDbl(Trim(flxNewProd.TextMatrix(iParam, 2))) + 1 Then
                If i <> flxNewProd.row Then
                    If Trim(flxNewProd.TextMatrix(i, 2)) = CDbl(Trim(flxNewProd.TextMatrix(flxNewProd.row, 2))) + 1 Then
                        nItem = flxNewProd.row
                        Exit Function
                    End If
                End If
            End If
    Next i
DevuelvePosMatriz = nItem
End Function

Private Sub cboAdd_Click()
Dim nCorr As Integer
Dim nIni As Integer
Dim sCorr As String
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
flxNewProd.lbEditarFlex = True
If LstProdCab.Text = "" Then
    nIni = CInt(Mid(Trim(cboProdCab.Text), Len(Trim(cboProdCab.Text)) - 2, 1))
Else
    nIni = CInt(Mid(Trim(LstProdCab.Text), Len(Trim(LstProdCab.Text)) - 2, 1))
End If
If CInt(flxNewProd.rows) - 1 > 1 And flxNewProd.TextMatrix(flxNewProd.rows - 1, 1) = "" Then
    flxNewProd.EliminaFila CInt(flxNewProd.rows) - 1
Else
    
    nCorr = oCredCat.ObtieneCorrelativoNewProd(CStr(nIni))
   
    If flxNewProd.TextMatrix(flxNewProd.rows - 1, 2) <> "" Then
        If CInt(flxNewProd.TextMatrix(flxNewProd.rows - 1, 2)) >= nCorr Then
            nCorr = CInt(flxNewProd.TextMatrix(flxNewProd.rows - 1, 2)) + 1
        End If
    End If
    sCorr = CStr(nCorr)
    flxNewProd.AdicionaFila
    flxNewProd.TextMatrix(flxNewProd.rows - 1, 1) = ""
    flxNewProd.TextMatrix(flxNewProd.rows - 1, 2) = sCorr
End If
'JOEP
Set oCredCat = Nothing
'JOEP
End Sub

Private Sub cboElim_Click()
Dim nItem As Integer
Dim psReg As String
Dim rs As New ADODB.Recordset
Dim i As Integer, nPos As Integer
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
nPos = 1
nItem = flxNewProd.row
psReg = oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(nItem, 2))))

If flxNewProd.TextMatrix(nItem, 1) <> "" Then
    If psReg = "SI" Then
        If MsgBox("¿Esta seguro de que desea eliminar el producto seleccionado?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
            psReg = oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(flxNewProd.row, 2))), "byItem")
            oCredCat.EliminaNewProducto (CStr(Trim(flxNewProd.TextMatrix(nItem, 2))))
            flxNewProd.EliminaFila nItem
           
            Call CargarRelacionProductos("ElimNewProd")
            For i = nItem To flxNewProd.rows - 1
                If flxNewProd.TextMatrix(nItem, 2) <> "" Then
                    If CStr(Trim(flxNewProd.TextMatrix(i - 1, 2))) = "Id_NuevoProd" And flxNewProd.rows - 1 > 1 Then
                        Exit For
                    ElseIf oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(i - 1, 2))), "byItem") = "NO" And oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(i, 2))), "byItem") = "NO" Then
                        Exit For
                    ElseIf oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(i - 1, 2))), "byItem") = "NO" And oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(i, 2))), "byItem") = "SI" And psReg = "NO" Then
                        Exit For
                    ElseIf psReg = "SI" And i <= flxNewProd.rows - 1 Then
'
                        If i <= flxNewProd.rows - 1 Or oCredCat.ObtieneValRegNewProd(CStr(Trim(flxNewProd.TextMatrix(i - 1, 2))), "byItem") = "NO" Then
                            flxNewProd.TextMatrix(i, 2) = CInt(flxNewProd.TextMatrix(i, 2)) - 1
                        End If
                    End If
                End If
            Next i
            chkNoVig.value = 0
            MsgBox "El Producto fue eliminado satisfactoriamente", vbInformation, "Aviso"
    Else
        MsgBox "El producto no puede ser eliminado, debido a que existen créditos registrados con dicho Producto, por lo que deberá recurrir a la migración", vbInformation, "Aviso"
    End If
Else
    flxNewProd.EliminaFila nItem
End If
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub CmdQuitAsign_Click()
Dim Item As ListItem
Dim psTpoProdAnt As String
Dim psTpoProdNew As String
Dim psDeshabRelac As String
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

If LstRelacion.ListItems.count <> 0 Then
    psTpoProdAnt = LstRelacion.ListItems(LstRelacion.SelectedItem.Index).SubItems(1)
    psTpoProdNew = LstRelacion.ListItems(LstRelacion.SelectedItem.Index).SubItems(3)
    If MsgBox("¿Desea quitar la relación seleccionada..", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
    psDeshabRelac = oCredCat.DeshabilitarMigracionRel(psTpoProdAnt, psTpoProdNew)
    If psDeshabRelac = "SI" Then
        LstRelacion.ListItems.Remove (LstRelacion.SelectedItem.Index)
        MsgBox "Se quitó la relación satisfactoriamente", vbInformation, "Aviso"
    Else
        MsgBox "No se puede quitar la relación seleccionada, debido a que existen créditos que dependen de dicha migración.", vbInformation, "Aviso"
        Exit Sub
    End If
End If
'JOEP
Set oCredCat = Nothing
'JOEP
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SSTabMain.Tab = 0
End Sub

Private Sub LstProdCab_DblClick()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim psRelacProd As String
    Dim Item As ListItem
    'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
    i = 1
    
    '***Agregado by NAGL 20190926***
    CmdAsigSubProd.Enabled = True
    CmdQuitAsign.Enabled = True
    cmdNuevoProd.Enabled = True
    CmdAll.Enabled = True
    cboAdd.Enabled = True
    cboElim.Enabled = True
    chkNoVig.Enabled = True
    cmdGrabarRelProd.Enabled = True
    '********************************
    
    chkNoVig.value = 0
    psRelacProd = Mid(Trim(LstProdCab.Text), Len(Trim(LstProdCab.Text)) - 2, 1)
    Call LimpiarControlesRelacProd
    
    Set rs = oCredCat.CargarTipoProd("Prod", psRelacProd)
    Set oCredCat = Nothing
    Do While Not rs.EOF
        LstTpoProd.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
    Set rs = oCredCat.ObtieneRelacProdEquiv(psRelacProd)
    Do While Not rs.EOF
        Set Item = LstRelacion.ListItems.Add(, , LstTpoProd.List(LstTpoProd.ListIndex))
        Item.SubItems(1) = rs!cTpoProdAnt
        Item.SubItems(2) = rs!cConsDescrip
        Item.SubItems(3) = rs!cTpoProdNew
        Item.SubItems(4) = rs!cConsDescripNew
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
    Set rs = oCredCat.ObtieneRelacProdEquiv(psRelacProd, "ProdNew")
    Do While Not rs.EOF
        flxNewProd.TextMatrix(i, 1) = rs!cConsDescripNew
        flxNewProd.TextMatrix(i, 2) = rs!cTpoProdNew
        i = i + 1
        flxNewProd.AdicionaFila
        rs.MoveNext
    Loop
    flxNewProd.EliminaFila (i)
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub cmdGrabarRelProd_Click()
Dim oCont As New NContFunciones
Dim i As Integer
Dim psMovNro As String
Dim MatListaRelNewProd As Variant
Dim MatNewProd As Variant
Dim pnCantRegListProd As Integer
Dim pnCantRegNewProd As Integer

'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

ReDim MatListaRelNewProd(LstRelacion.ListItems.count, 2)
ReDim MatNewProd(flxNewProd.rows - 1, 3)

chkNoVig.value = 0
pnCantRegListProd = LstRelacion.ListItems.count
pnCantRegNewProd = flxNewProd.rows - 1
    If pnCantRegNewProd <> 0 Then
        For i = 0 To pnCantRegListProd - 1
            MatListaRelNewProd(i, 0) = LstRelacion.ListItems(i + 1).SubItems(1)
            MatListaRelNewProd(i, 1) = LstRelacion.ListItems(i + 1).SubItems(3)
            MatListaRelNewProd(i, 2) = LstRelacion.ListItems(i + 1).SubItems(4)
        Next i
        
        For i = 1 To pnCantRegNewProd
            MatNewProd(i, 0) = CStr(Trim(flxNewProd.TextMatrix(i, 2)))
            MatNewProd(i, 1) = CStr(Trim(flxNewProd.TextMatrix(i, 1)))
        Next i
        
        psMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        If MsgBox("¿Desea guardar los cambios realizados..", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
           oCredCat.GrabarNewProducto MatListaRelNewProd, MatNewProd, psMovNro
           Call CargarRelacionProductos("AddReg") 'NAGL 20190926 Agregó Parámetro
           MsgBox "Los Datos se guardaron satisfactoriamente", vbInformation, "AVISO"
    Else
        MsgBox "No existe ningún registro en la Relación de Nuevos Productos", vbInformation, "AVISO"
    End If
    chkCondiciones.value = 0
'JOEP
Set oCredCat = Nothing
'JOEP
End Sub

'********PARTE DE CONDICIONES********'
Private Sub SSTabMain_Click(PreviousTab As Integer)
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
    'If PreviousTab = 0 Then
    If SSTabMain.Tab = 1 Then
         frmSecRang.Visible = True
         Call ControlesVisible("Cond")
         If cboCatCar.Text = "" Then
            cboCatCar.ListIndex = 0
         End If
         CentraForm Me
    'ElseIf PreviousTab = 1 Then
    ElseIf SSTabMain.Tab = 0 Then
        Call ControlesVisible("Prod")
        CentraForm Me
'JOEP20181204 ERS034-CP
    ElseIf SSTabMain.Tab = 2 Then
    Dim rsCmbCheckLis As ADODB.Recordset
        SSTabMain.Height = 7550
        SSTabMain.Width = 16800
        Width = 17160
        Height = 8145
        cmdSalir.Left = 13950
        cmdSalir.Top = 7200
        CentraForm Me
        Call habilitarControlesCheckList(1, False)
        
        Call CargaListaParametrosCheckList
        
        Set rsCmbCheckLis = oCredCat.CargarCboCondiciones(0, "UnidRangoIni")
        Call LlenarComboRS(rsCmbCheckLis, cmbIniCheckList)
        
        Set rsCmbCheckLis = Nothing
        Set rsCmbCheckLis = oCredCat.CargarCboCondiciones(0, "UnidRangoFin")
        Call LlenarComboRS(rsCmbCheckLis, cmbFinCheckList)
    
        Set rsCmbCheckLis = oCredCat.CargarCboCondiciones(4000, "UnidRango")
        Call LlenarComboRS(rsCmbCheckLis, cmbUniMedCheckList)
        cmbUniMedCheckList.RemoveItem (IndiceListaCombo(cmbUniMedCheckList, 1006))
        
        RSClose rsCmbCheckLis
'JOEP20181204 ERS034-CP
    End If
'JOEP
Set oCredCat = Nothing
'JOEP
End Sub

Private Sub cmdAsigParam_Click()
Dim psTpoDescrip As String
Dim psDescrip As String
Dim rs As New ADODB.Recordset
Dim nRelParam As Integer

'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
  If lstParamMain.ListIndex <> -1 Then
     For nRelParam = 1 To flxRelacParam.rows - 1
        If Trim(Right(lstParamMain.List(lstParamMain.ListIndex), 15)) = Trim(flxRelacParam.TextMatrix(nRelParam, 2)) Then
            Exit Sub
        End If
     Next nRelParam
     psTpoDescrip = Trim(Right(lstParamMain.List(lstParamMain.ListIndex), 10))
     psDescrip = Trim(Mid(lstParamMain.List(lstParamMain.ListIndex), 1, Len(lstParamMain.List(lstParamMain.ListIndex)) - 10))
     flxRelacParam.AdicionaFila
     flxRelacParam.lbEditarFlex = True
     flxRelacParam.TextMatrix(flxRelacParam.rows - 1, 1) = psDescrip
     flxRelacParam.TextMatrix(flxRelacParam.rows - 1, 2) = psTpoDescrip
     'Si es rango se habilitará la sección de Rangos
     Set rs = oCredCat.ObtieneListaParametro("ParamRang", CDbl(psTpoDescrip), "R")
     If rs.RecordCount <> 0 Then
        flxRelacParam.TextMatrix(flxRelacParam.rows - 1, 5) = "R"
     End If
     If flxRelacParam.rows - 1 > 1 And (chkMatriz.value = 1 And chkIndep.value = 0) Then
        flxRelacParam.TextMatrix(flxRelacParam.rows - 1, 12) = Trim(Right(flxRelacParam.TextMatrix(flxRelacParam.rows - 2, 11), 8))
        flxRelacParam.TextMatrix(flxRelacParam.rows - 1, 11) = flxRelacParam.TextMatrix(flxRelacParam.rows - 2, 11)
     End If
  Else
     MsgBox "Debe seleccionar una Tipo de Parámetro", vbInformation, "AVISO"
  End If
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub cmdQuitarParam_Click()
Dim pnPosicion As Integer
Dim psTpoDescrip As String
If (flxRelacParam.rows - 1 > 1 And flxRelacParam.TextMatrix(flxRelacParam.row, flxRelacParam.Col) <> "") Then
    'Elimina lo del Listado
    psTpoDescrip = Trim(Right(flxRelacParam.TextMatrix(flxRelacParam.row, 2), 15))
   
    'Elimina lo del FlexEdit
    pnPosicion = DevuelvePosicionGrid(psTpoDescrip)
    Call flxRelacParam.EliminaFila(pnPosicion)
Else
    Call flxRelacParam.EliminaFila(flxRelacParam.rows - 1)
End If
End Sub

Private Sub cmdSubir_Click()
Dim psIdParamA, psIdParamB As String
Dim psParametroA, psParametroB As String
Dim psIdDescripA, psIdDescripB As String
Dim psDescripcionA, psDescripcionB As String
Dim psRangoA, psRangoB As String
Dim nPosicion As Long
nPosicion = flxRelacParam.row
If nPosicion > 1 Then
    psParametroA = flxRelacParam.TextMatrix(nPosicion - 1, 1)
    psIdParamA = flxRelacParam.TextMatrix(nPosicion - 1, 2)
    psDescripcionA = flxRelacParam.TextMatrix(nPosicion - 1, 3)
    psIdDescripA = flxRelacParam.TextMatrix(nPosicion - 1, 4)
    psRangoA = flxRelacParam.TextMatrix(nPosicion - 1, 5)
    
    psParametroB = flxRelacParam.TextMatrix(nPosicion, 1)
    psIdParamB = flxRelacParam.TextMatrix(nPosicion, 2)
    psDescripcionB = flxRelacParam.TextMatrix(nPosicion, 3)
    psIdDescripB = flxRelacParam.TextMatrix(nPosicion, 4)
    psRangoB = flxRelacParam.TextMatrix(nPosicion, 5)
    
    flxRelacParam.TextMatrix(nPosicion - 1, 1) = psParametroB
    flxRelacParam.TextMatrix(nPosicion - 1, 2) = psIdParamB
    flxRelacParam.TextMatrix(nPosicion - 1, 3) = psDescripcionB
    flxRelacParam.TextMatrix(nPosicion - 1, 4) = psIdDescripB
    flxRelacParam.TextMatrix(nPosicion - 1, 5) = psRangoB
    
    flxRelacParam.TextMatrix(nPosicion, 1) = psParametroA
    flxRelacParam.TextMatrix(nPosicion, 2) = psIdParamA
    flxRelacParam.TextMatrix(nPosicion, 3) = psDescripcionA
    flxRelacParam.TextMatrix(nPosicion, 4) = psIdDescripA
    flxRelacParam.TextMatrix(nPosicion, 5) = psRangoA
    
    flxRelacParam.row = nPosicion - 1
    flxRelacParam.SetFocus
End If
End Sub

Private Sub cmdBajar_Click()
Dim psIdParamA, psIdParamB As String
Dim psParametroA, psParametroB As String
Dim psIdDescripA, psIdDescripB As String
Dim psDescripcionA, psDescripcionB As String
Dim psRangoA, psRangoB As String
Dim nPosicion As Long
nPosicion = flxRelacParam.row
If nPosicion >= 1 And flxRelacParam.rows - 1 > nPosicion Then
    psParametroA = flxRelacParam.TextMatrix(nPosicion + 1, 1)
    psIdParamA = flxRelacParam.TextMatrix(nPosicion + 1, 2)
    psDescripcionA = flxRelacParam.TextMatrix(nPosicion + 1, 3)
    psIdDescripA = flxRelacParam.TextMatrix(nPosicion + 1, 4)
    psRangoA = flxRelacParam.TextMatrix(nPosicion + 1, 5)
    
    psParametroB = flxRelacParam.TextMatrix(nPosicion, 1)
    psIdParamB = flxRelacParam.TextMatrix(nPosicion, 2)
    psDescripcionB = flxRelacParam.TextMatrix(nPosicion, 3)
    psIdDescripB = flxRelacParam.TextMatrix(nPosicion, 4)
    psRangoB = flxRelacParam.TextMatrix(nPosicion, 5)
    
    flxRelacParam.TextMatrix(nPosicion + 1, 1) = psParametroB
    flxRelacParam.TextMatrix(nPosicion + 1, 2) = psIdParamB
    flxRelacParam.TextMatrix(nPosicion + 1, 3) = psDescripcionB
    flxRelacParam.TextMatrix(nPosicion + 1, 4) = psIdDescripB
    flxRelacParam.TextMatrix(nPosicion + 1, 5) = psRangoB
    
    flxRelacParam.TextMatrix(nPosicion, 1) = psParametroA
    flxRelacParam.TextMatrix(nPosicion, 2) = psIdParamA
    flxRelacParam.TextMatrix(nPosicion, 3) = psDescripcionA
    flxRelacParam.TextMatrix(nPosicion, 4) = psIdDescripA
    flxRelacParam.TextMatrix(nPosicion, 5) = psRangoA
    flxRelacParam.row = nPosicion + 1
    flxRelacParam.SetFocus
End If
End Sub

Private Sub flxRelacParam_OnChangeCombo()
Dim iFila As Integer
    iFila = 1
    If flxRelacParam.Col = 3 Then
        flxRelacParam.TextMatrix(flxRelacParam.row, 4) = Trim(Right(flxRelacParam.TextMatrix(flxRelacParam.row, 3), 8))
        '****************************************************
        'flxRelacParam.TextMatrix(flxRelacParam.row, 5) = ""
        flxRelacParam.TextMatrix(flxRelacParam.row, 6) = ""
        flxRelacParam.TextMatrix(flxRelacParam.row, 7) = ""
        flxRelacParam.TextMatrix(flxRelacParam.row, 8) = ""
        flxRelacParam.TextMatrix(flxRelacParam.row, 9) = ""
        flxRelacParam.TextMatrix(flxRelacParam.row, 10) = ""
        '***************NAGL 20190923************************
    ElseIf flxRelacParam.Col = 11 Then
        flxRelacParam.TextMatrix(flxRelacParam.row, 12) = Trim(Right(flxRelacParam.TextMatrix(flxRelacParam.row, 11), 8))
    End If
    If (chkMatriz.value = 1 And chkIndep.value = 0) And flxRelacParam.Col = 11 Then
        Do While iFila <= flxRelacParam.rows - 1
            flxRelacParam.TextMatrix(iFila, 12) = Trim(Right(flxRelacParam.TextMatrix(flxRelacParam.row, 11), 8))
            flxRelacParam.TextMatrix(iFila, 11) = flxRelacParam.TextMatrix(flxRelacParam.row, 11)
            iFila = iFila + 1
        Loop
    End If
End Sub

Private Sub flxRelacParam_DblClick()
Dim psTpoDescrip As String
Dim RsCbo As New ADODB.Recordset
Dim RsMod As New ADODB.Recordset
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
flxRelacParam.TamañoCombo (0)
    If flxRelacParam.Col = 3 Then
        psTpoDescrip = flxRelacParam.TextMatrix(flxRelacParam.row, 2)
        If psTpoDescrip <> "" Then
          Set RsCbo = oCredCat.CargarCboCondiciones(CDbl(psTpoDescrip))
          flxRelacParam.CargaCombo RsCbo
        End If
        flxRelacParam.TamañoCombo (300) 'COMBO
    ElseIf flxRelacParam.Col = 11 Then
        Set RsMod = oCredCat.CargaTiposModulos()
        flxRelacParam.CargaCombo RsMod
        flxRelacParam.TamañoCombo (20)
    End If
'JOEP
Set oCredCat = Nothing
RSClose RsCbo
RSClose RsMod
'JOEP
End Sub

Public Sub CargaListaParametrosMain()
Dim rs As New ADODB.Recordset
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
lstParamMain.Clear
If lstParamMain.ListIndex = -1 Then
    Set rs = oCredCat.ObtieneListaParametro()
    Set oCredCat = Nothing
    Do While Not rs.EOF
        lstParamMain.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
        rs.MoveNext
    Loop
End If
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Public Function DevuelvePosicionGrid(psParParam As String) As Integer
Dim i As Integer
Dim nItem As Integer
nItem = 0
For i = 1 To CInt(flxRelacParam.rows) - 1 'Para Observar si ya se ha ingresado el Parámetro obtenido
    If CDbl(Trim(flxRelacParam.TextMatrix(i, 2))) = psParParam Then
       nItem = i
       Exit For
    End If
Next i
DevuelvePosicionGrid = nItem
End Function

'******Sección Check Matriz - Independiente******
Private Sub chkMatriz_Click()
    If chkMatriz.value = 1 Then
        chkIndep.value = 0
    ElseIf chkMatriz.value = 0 And chkIndep.value = 0 Then
        chkMatriz.value = 1
    End If
    LimpiaColumModulo
End Sub

Private Sub chkIndep_Click()
    If chkIndep.value = 1 Then
        chkMatriz.value = 0
    ElseIf chkIndep.value = 0 And chkMatriz.value = 0 Then
        chkIndep.value = 1
    End If
    LimpiaColumModulo
End Sub
'***********************************************
Public Sub LimpiaColumModulo()
Dim i As Integer
    For i = 1 To flxRelacParam.rows - 1
        flxRelacParam.TextMatrix(i, 12) = ""
        flxRelacParam.TextMatrix(i, 11) = ""
    Next i
End Sub

'***Sección Check Condiciones - Requisitos
Private Sub chkCondiciones_Click()
    chkHabRang.value = 0
    If chkCondiciones.value = 1 Then
        LimpiarTxtRango
        chkRequisitos.value = 0
        FormateaFlex flxRelacParam
        Call HabilitarControlesToCargarData(False) 'NAGL 20190923
        Call CargarListadoCondicionesRequisitos("Cond")
        Call HabilitarControlesToCargarData(True) 'NAGL 20190923
    ElseIf chkCondiciones.value = 0 And chkRequisitos.value = 0 Then
        chkCondiciones.value = 1
    End If
End Sub

Private Sub chkRequisitos_Click()
    chkHabRang.value = 0
    If chkRequisitos.value = 1 Then
        LimpiarTxtRango
        chkCondiciones.value = 0
        FormateaFlex flxRelacParam
        Call HabilitarControlesToCargarData(False) 'NAGL 20190923
        Call CargarListadoCondicionesRequisitos("Requ")
        Call HabilitarControlesToCargarData(True) 'NAGL 20190923
    ElseIf chkRequisitos.value = 0 And chkCondiciones.value = 0 Then
        chkRequisitos.value = 1
    End If
End Sub

Public Sub HabilitarControlesToCargarData(Optional pbHabilitar As Boolean = True, Optional psTipo As String = "")
If pbHabilitar = True Then
    cboCatCar.Enabled = True
    cboTpoProdCar.Enabled = True
    cmdAsigParam.Enabled = True
    cmdQuitarParam.Enabled = True
    cmdSubir.Enabled = True
    cmdBajar.Enabled = True
    cmdAgregar.Enabled = True
    chkMatriz.Enabled = True
    chkIndep.Enabled = True
    frmParam.Enabled = True
    cmdEditar.Enabled = True
    cmdSubirCondRequ.Enabled = True
    cmdbajarCondRequ.Enabled = True
    cmdElimRowFlex.Enabled = True
    chkHabRang.Enabled = True
    cmdGrabarDatos.Enabled = True
    If psTipo = "Grabar" Then
       chkCondiciones.Enabled = True
       chkRequisitos.Enabled = True
       cmdGrabarDatos.Enabled = True
       cmdGrabarDatos.Caption = "Grabar"
    End If
Else
    cboCatCar.Enabled = False
    cboTpoProdCar.Enabled = False
    cmdAsigParam.Enabled = False
    cmdQuitarParam.Enabled = False
    cmdSubir.Enabled = False
    cmdBajar.Enabled = False
    cmdAgregar.Enabled = False
    chkMatriz.Enabled = False
    chkIndep.Enabled = False
    frmParam.Enabled = False
    cmdEditar.Enabled = False
    cmdSubirCondRequ.Enabled = False
    cmdbajarCondRequ.Enabled = False
    cmdElimRowFlex.Enabled = False
    chkHabRang.Enabled = False
    cmdGrabarDatos.Enabled = False
    If psTipo = "Grabar" Then
       chkCondiciones.Enabled = False
       chkRequisitos.Enabled = False
       cmdGrabarDatos.Enabled = False
       cmdGrabarDatos.Caption = "Guardando"
    End If
End If
End Sub 'NAGL 20190923

'*****************************************

Private Sub cmdAgregar_Click()
chkHabRang.value = 0
If cmdAgregar.Caption = "Registrar" Then
    pbEditarGeneral = True
    If flxRelacParam.TextMatrix(1, 1) <> "" Then
        If ValidaPreRegistro("Matr") Then
            Call CargarMSHFlexCondRelacion(flxRelacParam, "Matriz", True) 'Actualización Matriz
            Call RowBackColor(MSHFlexCond.row)
        End If
    Else
        MsgBox "Falta ingresar parametros para editar el Registro", vbInformation, "Aviso"
        pbEditarGeneral = True
        cmdAgregar.SetFocus
        Exit Sub
    End If
Else
    If flxRelacParam.rows - 1 >= 1 And flxRelacParam.TextMatrix(1, 1) <> "" Then
        If chkCondiciones.value <> 0 Or chkRequisitos.value <> 0 Then
            If chkMatriz.value = 1 Then
                If ValidaPreRegistro("Matr") Then
                    Call CargarMSHFlexCondRelacion(flxRelacParam, "Matriz") 'Registro de Matriz
                    Call RowBackColor(MSHFlexCond.rows - 1)
                End If
            ElseIf chkIndep.value = 1 Then
                If ValidaPreRegistro("Ind") Then
                    Call CargarMSHFlexCondRelacion(flxRelacParam, "Ind") 'Registro de Independientes
                    Call RowBackColor(MSHFlexCond.rows - 1)
                End If
            Else
                MsgBox "Debe seleccionar un Tipo de Registro", vbInformation, "Aviso"
                chkIndep.SetFocus
            End If
        Else
            MsgBox "Debe seleccionar un Tipo de Característica: Condición o Requisito", vbInformation, "Aviso"
            chkCondiciones.SetFocus
        End If
    End If
End If
End Sub

Private Sub MSHFlexCond_EnterCell()
 Call RowBackColor(MSHFlexCond.row)
 Call CriterioRango("PorFila")
  chkHabRang.value = 0
End Sub

Private Sub CargarMSHFlexCondRelacion(flxRelacParam As FlexEdit, psReg As String, Optional pbEditar As Boolean = False)
Dim i As Integer
Dim psTpoParam As String, psParametro As String
Dim psTpoDescrip As String, psDescrip As String, psRango As String, psIdModulo As String, psModuloDesc As String
Dim nItemSec As Integer, pnTpoParam As Long
Dim pnCol As Integer, nCantRows As Integer, nCantRang As Integer, nId As Integer
Dim iFila As Integer

ReDim psIdParamAnt(100)
ReDim psParametroAnt(100)
ReDim psRangDescAnt(100)
ReDim psOperIniAnt(100)
ReDim psRangIniAnt(100)
ReDim psOperFinAnt(100)
ReDim psRangFinAnt(100)
ReDim psUnidRangAnt(100)
ReDim psValorUnidRAnt(100)
ReDim psIdModuloAnt(100)
ReDim psModuloDescAnt(100)
ReDim psMovNroAnt(100)

MSHFlexCond.SelectionMode = flexSelectionByRow
'MSHFlexCond.AllowUserResizing = flexResfizeColumns

nCantRows = 1
pnCol = 0
nCantRang = 0
nItemSec = 0
nId = 1
nOrd = 1

If flxRelacParam.rows - 1 >= 1 And flxRelacParam.TextMatrix(flxRelacParam.row, 2) <> "" Then
    If psReg = "Ind" Then
        pnCol = pnCol + 12
        If MSHFlexCond.cols < pnCol + 1 Then '3
            MSHFlexCond.cols = pnCol + 1 '3
        End If
        i = nCantRows
        
        psTpoParam = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 2))
        psParametro = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 1))
        psTpoDescrip = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 4))
        psDescrip = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 3))
        psRango = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 5))
        psIdModulo = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 12))
        psModuloDesc = Trim(Left(flxRelacParam.TextMatrix(flxRelacParam.row, 11), 50))
        
        If psDescrip <> "" Then
            psDescrip = Trim(Mid(psDescrip, 1, Len(psDescrip) - 15))
        End If
        
        nItemSec = LastPosicion(psTpoParam, 2)
        If nItemSec = 0 Then
            MSHFlexCond.AddItem ""
            If MSHFlexCond.TextMatrix(1, 1) = "" Then
                MSHFlexCond.RemoveItem 1
            End If
            nItemSec = MSHFlexCond.rows - 1
        Else
            MSHFlexCond.AddItem "", nItemSec
        End If
        
        MSHFlexCond.ColWidth(1) = 2000
        MSHFlexCond.ColWidth(2) = 0 '500
        MSHFlexCond.ColWidth(3) = 2500
        MSHFlexCond.ColWidth(4) = 0 '500
        MSHFlexCond.ColWidth(5) = 0 '500
        MSHFlexCond.ColWidth(6) = 0 '500
        MSHFlexCond.ColWidth(7) = 0 '500
        MSHFlexCond.ColWidth(8) = 0 '500
        MSHFlexCond.ColWidth(9) = 0 '500
        MSHFlexCond.ColWidth(10) = 0 '500
        MSHFlexCond.ColWidth(11) = 0 '500
        MSHFlexCond.ColWidth(12) = 0 '500
                
        MSHFlexCond.TextMatrix(0, 0) = "N°"
        MSHFlexCond.TextMatrix(0, 1) = "PARÁMETRO"
        MSHFlexCond.TextMatrix(0, 2) = "CodParam"
        MSHFlexCond.TextMatrix(0, 3) = "DESCRIPCIÓN"
        
        MSHFlexCond.TextMatrix(nItemSec, 1) = psParametro
        MSHFlexCond.TextMatrix(nItemSec, 2) = psTpoParam
        MSHFlexCond.TextMatrix(nItemSec, 3) = IIf(psRango <> "", "", psDescrip)
        MSHFlexCond.TextMatrix(nItemSec, 4) = ""
        MSHFlexCond.TextMatrix(nItemSec, 5) = ""
        MSHFlexCond.TextMatrix(nItemSec, 6) = ""
        MSHFlexCond.TextMatrix(nItemSec, 7) = ""
        MSHFlexCond.TextMatrix(nItemSec, 8) = ""
        MSHFlexCond.TextMatrix(nItemSec, 9) = psTpoDescrip
        MSHFlexCond.TextMatrix(nItemSec, 10) = psIdModulo
        MSHFlexCond.TextMatrix(nItemSec, 11) = psModuloDesc
        MSHFlexCond.TextMatrix(nItemSec, 12) = ""
        MSHFlexCond.MergeCol(1) = True
         
        MSHFlexCond.MergeCells = flexMergeRestrictColumns
        MSHFlexCond.MergeCol(1) = True
        
    Else
    
        For nCantRang = 1 To flxRelacParam.rows - 1
            pnCol = pnCol + 12
        Next nCantRang
        
        If MSHFlexCond.cols < pnCol + 1 Then '3
            MSHFlexCond.cols = pnCol + 1 '3
        End If
        
        If pbEditar = True Then
            nItemSec = MSHFlexCond.row 'Para la fila a Editar
        Else
            If nItemSec = 0 Then
                MSHFlexCond.AddItem ""
                If MSHFlexCond.TextMatrix(1, 1) = "" Then
                    MSHFlexCond.RemoveItem 1
                End If
                nItemSec = MSHFlexCond.rows - 1
            Else
                MSHFlexCond.AddItem "", nItemSec
            End If
        End If
        '*******Cuando se procede a Editar el Registro*********
        i = 1
        If pbEditar = True Then
            Do While i < MSHFlexCond.cols - 1
            
                psParametroAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i)
                psIdParamAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 1)
                psRangDescAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 2)
                psOperIniAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 3)
                psRangIniAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 4)
                psOperFinAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 5)
                psRangFinAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 6)
                psUnidRangAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 7)
                psValorUnidRAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 8)
                psIdModuloAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 9)
                psModuloDescAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 10)
                psMovNroAnt(nOrd) = MSHFlexCond.TextMatrix(nItemSec, i + 11)
        
                MSHFlexCond.TextMatrix(nItemSec, i) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 1) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 2) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 3) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 4) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 5) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 6) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 7) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 8) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 9) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 10) = ""
                MSHFlexCond.TextMatrix(nItemSec, i + 11) = ""
                
                i = i + 12
                nOrd = nOrd + 1
            Loop
        End If
        '******************************************************
        i = nCantRows
        Do While nCantRows <= flxRelacParam.rows - 1
               MSHFlexCond.MergeCells = flexMergeFree
               psTpoParam = Trim(flxRelacParam.TextMatrix(nCantRows, 2))
               psParametro = Trim(flxRelacParam.TextMatrix(nCantRows, 1))
               psTpoDescrip = Trim(flxRelacParam.TextMatrix(nCantRows, 4))
               psDescrip = Trim(flxRelacParam.TextMatrix(nCantRows, 3))
               psRango = Trim(flxRelacParam.TextMatrix(nCantRows, 5))
               psIdModulo = Trim(flxRelacParam.TextMatrix(flxRelacParam.row, 12))
               psModuloDesc = Trim(Left(flxRelacParam.TextMatrix(flxRelacParam.row, 11), 50))
                
               If psDescrip <> "" And pbEditar = False Then
                    psDescrip = Trim(Mid(psDescrip, 1, Len(psDescrip) - 15))
               End If
                
               If i = 1 Then
                   MSHFlexCond.ColWidth(i) = 2000
               Else
                    MSHFlexCond.ColWidth(i) = 0
               End If
               MSHFlexCond.ColWidth(i + 1) = 0 '500
               MSHFlexCond.ColWidth(i + 2) = 2500
               MSHFlexCond.ColWidth(i + 3) = 0 '500
               MSHFlexCond.ColWidth(i + 4) = 0 '500
               MSHFlexCond.ColWidth(i + 5) = 0 '500
               MSHFlexCond.ColWidth(i + 6) = 0 '500
               MSHFlexCond.ColWidth(i + 7) = 0 '500
               MSHFlexCond.ColWidth(i + 8) = 0 '500
               MSHFlexCond.ColWidth(i + 9) = 0 '500
               MSHFlexCond.ColWidth(i + 10) = 0 '500
               MSHFlexCond.ColWidth(i + 11) = 0 '500
               
               'MSHFlexCond.TextMatrix(nItemSec, 0) = CStr(MSHFlexCond.rows - 1)
               If pbEditar = False Then
                    MSHFlexCond.TextMatrix(nItemSec, i) = psParametro
                    MSHFlexCond.TextMatrix(nItemSec, i + 1) = psTpoParam
                    MSHFlexCond.TextMatrix(nItemSec, i + 2) = IIf(psRango <> "", "", psDescrip)
                    MSHFlexCond.TextMatrix(nItemSec, i + 3) = ""
                    MSHFlexCond.TextMatrix(nItemSec, i + 4) = ""
                    MSHFlexCond.TextMatrix(nItemSec, i + 5) = ""
                    MSHFlexCond.TextMatrix(nItemSec, i + 6) = ""
                    MSHFlexCond.TextMatrix(nItemSec, i + 7) = ""
                    MSHFlexCond.TextMatrix(nItemSec, i + 8) = psTpoDescrip
                    MSHFlexCond.TextMatrix(nItemSec, i + 9) = psIdModulo
                    MSHFlexCond.TextMatrix(nItemSec, i + 10) = psModuloDesc
                    MSHFlexCond.TextMatrix(nItemSec, i + 11) = ""
                Else
                    MSHFlexCond.TextMatrix(nItemSec, i) = psParametro
                    MSHFlexCond.TextMatrix(nItemSec, i + 1) = psTpoParam
                    MSHFlexCond.TextMatrix(nItemSec, i + 2) = psDescrip
                    MSHFlexCond.TextMatrix(nItemSec, i + 3) = Trim(flxRelacParam.TextMatrix(nCantRows, 6))
                    MSHFlexCond.TextMatrix(nItemSec, i + 4) = Trim(flxRelacParam.TextMatrix(nCantRows, 7))
                    MSHFlexCond.TextMatrix(nItemSec, i + 5) = Trim(flxRelacParam.TextMatrix(nCantRows, 8))
                    MSHFlexCond.TextMatrix(nItemSec, i + 6) = Trim(flxRelacParam.TextMatrix(nCantRows, 9))
                    MSHFlexCond.TextMatrix(nItemSec, i + 7) = Trim(flxRelacParam.TextMatrix(nCantRows, 10))
                    MSHFlexCond.TextMatrix(nItemSec, i + 8) = IIf(psTpoDescrip <> "", psTpoDescrip, psTpoParam)
                    MSHFlexCond.TextMatrix(nItemSec, i + 9) = psIdModulo
                    MSHFlexCond.TextMatrix(nItemSec, i + 10) = psModuloDesc
                    MSHFlexCond.TextMatrix(nItemSec, i + 11) = ""
                End If
            
               If i = 1 Then
                    MSHFlexCond.TextMatrix(0, i) = "PARÁMETRO"
               Else
                    MSHFlexCond.TextMatrix(0, i) = "DESCRIPCIÓN"
               End If
            
               MSHFlexCond.TextMatrix(0, i + 1) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 2) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 3) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 4) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 5) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 6) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 7) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 8) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 9) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 10) = "DESCRIPCIÓN"
               MSHFlexCond.TextMatrix(0, i + 11) = "DESCRIPCIÓN"
               
               MSHFlexCond.MergeCol(i) = True
               
               i = i + 12
            nCantRows = nCantRows + 1
        Loop
        MSHFlexCond.MergeCells = flexMergeFree
        MSHFlexCond.MergeCol(1) = True
        MSHFlexCond.MergeCol(2) = True
        MSHFlexCond.MergeRow(0) = True
    End If
  
    '*******ValidaControlRegDuplicado******
    If ValidaControlRegDuplicado(nItemSec) = True Then
        If pbEditar = False Then
            MSHFlexCond.RemoveItem nItemSec
            MsgBox "El registro ya ha sido agregado !!", vbInformation, "Atención"
        Else
            MsgBox "El registro ya ha sido agregado, por favor intente nuevamente!!", vbInformation, "Atención"
            i = 1
            nOrd = 1
            Do While i < MSHFlexCond.cols - 1
                MSHFlexCond.TextMatrix(nItemSec, i) = psParametroAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 1) = psIdParamAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 2) = psRangDescAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 3) = psOperIniAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 4) = psRangIniAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 5) = psOperFinAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 6) = psRangFinAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 7) = psUnidRangAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 8) = psValorUnidRAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 9) = psIdModuloAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 10) = psModuloDescAnt(nOrd)
                MSHFlexCond.TextMatrix(nItemSec, i + 11) = psMovNroAnt(nOrd)
                
                psParametroAnt(nOrd) = ""
                psIdParamAnt(nOrd) = ""
                psRangDescAnt(nOrd) = ""
                psOperIniAnt(nOrd) = ""
                psRangIniAnt(nOrd) = ""
                psOperFinAnt(nOrd) = ""
                psRangFinAnt(nOrd) = ""
                psUnidRangAnt(nOrd) = ""
                psValorUnidRAnt(nOrd) = ""
                psIdModuloAnt(nOrd) = ""
                psModuloDescAnt(nOrd) = ""
                psMovNroAnt(nOrd) = ""
                i = i + 12
                nOrd = nOrd + 1
           Loop
            pbEditarGeneral = True
        End If
        cmdAgregar.SetFocus
        Exit Sub
    End If
  
    '**********Correlativo**************
    For i = 1 To MSHFlexCond.rows - 1
        MSHFlexCond.TextMatrix(i, 0) = nId
        nId = nId + 1
    Next i
    
    If pbEditar = True Then
        cmdAgregar.Caption = "Agregar"
        'cmdAgregar.Width = 735
        cmdEditar.Caption = "Editar"
        MsgBox "Se actualizaron los datos correctamente", vbInformation, "Aviso"
        FormateaFlex flxRelacParam
        chkMatriz.value = 1
        chkIndep.value = 0
        chkMatriz.Enabled = True
        chkIndep.Enabled = True
        chkCondiciones.Enabled = True
        chkRequisitos.Enabled = True
        cmdSubirCondRequ.Enabled = True
        cmdbajarCondRequ.Enabled = True
        cmdElimRowFlex.Enabled = True
        cmdGrabarDatos.Enabled = True
        chkHabRang.Enabled = True
        MSHFlexCond.Enabled = True
        pbEditarGeneral = False
    End If
    '***********************************
    MSHFlexCond.SetFocus
End If
End Sub

Public Function ValidaControlRegDuplicado(nItemSec As Integer) As Boolean
Dim iCol As Integer, iRowMSH As Integer, pnDupl As Integer, pnRowDuplCant As Integer
Dim rsRang As New ADODB.Recordset
iRowMSH = 1
pnDupl = 0
pnRowDuplCant = 0
'*******ValidaControlRegDuplicado******
    iCol = 1
    Do While iRowMSH <= MSHFlexCond.rows - 1
        Do While iCol <= MSHFlexCond.cols - 1
            If nItemSec <> iRowMSH Then
                pnDupl = ControlRegistrosDuplicados(iRowMSH, iCol, MSHFlexCond.TextMatrix(nItemSec, iCol), MSHFlexCond.TextMatrix(nItemSec, iCol + 1), MSHFlexCond.TextMatrix(nItemSec, iCol + 2), MSHFlexCond.TextMatrix(nItemSec, iCol + 3), _
                                 MSHFlexCond.TextMatrix(nItemSec, iCol + 4), MSHFlexCond.TextMatrix(nItemSec, iCol + 5), MSHFlexCond.TextMatrix(nItemSec, iCol + 6), MSHFlexCond.TextMatrix(nItemSec, iCol + 7), MSHFlexCond.TextMatrix(nItemSec, iCol + 8), MSHFlexCond.TextMatrix(nItemSec, iCol + 9), MSHFlexCond.TextMatrix(nItemSec, iCol + 10))
            End If
            If pnDupl > 0 Then
               pnRowDuplCant = pnRowDuplCant + 1
            End If
            If pnRowDuplCant = (MSHFlexCond.cols - 1) / 12 Then
               ValidaControlRegDuplicado = True
               Exit Function
            End If
            iCol = iCol + 12
        Loop
        iCol = 1
        pnDupl = 0
        pnRowDuplCant = 0
        iRowMSH = iRowMSH + 1
    Loop
    ValidaControlRegDuplicado = False
'**************************************
End Function

Public Function ControlRegistrosDuplicados(pnRow As Integer, pnCol As Integer, psParametro As String, psTpoParam As String, psRangDesc As String, psOperIni As String, psRangIni As String, psOperFin As String, psRangFin As String, psUnidRang As String, psValorUnidR As String, psIdModulo As String, psModuloDesc As String) As Integer
Dim iFila As Long
Dim iCol As Integer
Dim pnRep As Integer
Dim pnIgual As Integer
iCol = pnCol
iFila = pnRow
pnIgual = 0
If MSHFlexCond.TextMatrix(iFila, iCol) = psParametro Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 1) = psTpoParam Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 2) = psRangDesc Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 3) = psOperIni Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 4) = psRangIni Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 5) = psOperFin Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 6) = psRangFin Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 7) = psUnidRang Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 8) = psValorUnidR Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 9) = psIdModulo Then pnIgual = pnIgual + 1
If MSHFlexCond.TextMatrix(iFila, iCol + 10) = psModuloDesc Then pnIgual = pnIgual + 1
If pnIgual = 11 Then
    pnRep = 1
Else
    pnRep = 0
End If
ControlRegistrosDuplicados = pnRep
End Function

Public Function ValidaPreRegistro(psTpoIndMat As String) As Boolean
Dim i As Integer
Dim rs As New ADODB.Recordset
If psTpoIndMat = "Matr" Then
    If flxRelacParam.TextMatrix(1, 1) <> "" Then
        For i = 1 To flxRelacParam.rows - 1
            If flxRelacParam.TextMatrix(i, 5) <> "R" And (flxRelacParam.TextMatrix(i, 4) = "" Or flxRelacParam.TextMatrix(i, 3) = "") Then 'NAGL 20190923 Agregó flxRelacParam.TextMatrix(i, 3)
               MsgBox "Falta registrar la descripción del " & CStr(flxRelacParam.TextMatrix(i, 1)) & ".", vbInformation, "Aviso"
               Exit Function
            ElseIf flxRelacParam.TextMatrix(i, 5) = "R" And (flxRelacParam.TextMatrix(i, 4) = "" Or flxRelacParam.TextMatrix(i, 3) = "") Then
               flxRelacParam.TextMatrix(i, 4) = ""
               flxRelacParam.TextMatrix(i, 6) = ""
               flxRelacParam.TextMatrix(i, 7) = ""
               flxRelacParam.TextMatrix(i, 8) = ""
               flxRelacParam.TextMatrix(i, 9) = ""
               flxRelacParam.TextMatrix(i, 10) = ""
            End If 'NAGL 20190923 Agregó esta condición
            If (flxRelacParam.TextMatrix(i, 12) = "" Or flxRelacParam.TextMatrix(i, 11) = "") Then 'NAGL 20190923 Agregó flxRelacParam.TextMatrix(i, 11)
                MsgBox "Falta registrar el módulo perteneciente a " & CStr(flxRelacParam.TextMatrix(i, 1)) & ".", vbInformation, "Aviso"
                Exit Function
            End If
        Next i
    End If
Else
    If flxRelacParam.TextMatrix(1, 1) <> "" Then
        If flxRelacParam.TextMatrix(flxRelacParam.row, 5) <> "R" And (flxRelacParam.TextMatrix(flxRelacParam.row, 4) = "" Or flxRelacParam.TextMatrix(flxRelacParam.row, 3) = "") Then 'NAGL 20190923 Agregó flxRelacParam.TextMatrix(flxRelacParam.row, 3)
           MsgBox "Falta registrar la descripción del " & CStr(flxRelacParam.TextMatrix(flxRelacParam.row, 1)) & ".", vbInformation, "Aviso"
           Exit Function
        ElseIf flxRelacParam.TextMatrix(flxRelacParam.row, 5) = "R" And (flxRelacParam.TextMatrix(flxRelacParam.row, 4) = "" Or flxRelacParam.TextMatrix(flxRelacParam.row, 3) = "") Then
               flxRelacParam.TextMatrix(flxRelacParam.row, 4) = ""
               flxRelacParam.TextMatrix(flxRelacParam.row, 6) = ""
               flxRelacParam.TextMatrix(flxRelacParam.row, 7) = ""
               flxRelacParam.TextMatrix(flxRelacParam.row, 8) = ""
               flxRelacParam.TextMatrix(flxRelacParam.row, 9) = ""
               flxRelacParam.TextMatrix(flxRelacParam.row, 10) = ""
        End If 'NAGL 20190923 Agregó esta condición
        If (flxRelacParam.TextMatrix(flxRelacParam.row, 12) = "" Or flxRelacParam.TextMatrix(flxRelacParam.row, 11) = "") Then 'NAGL 20190923 Agregó flxRelacParam.TextMatrix(flxRelacParam.row, 11)
            MsgBox "Falta registrar el módulo perteneciente a " & CStr(flxRelacParam.TextMatrix(flxRelacParam.row, 1)) & ".", vbInformation, "Aviso"
            Exit Function
        End If
    End If
End If
ValidaPreRegistro = True
End Function

Private Sub cmdElimRowFlex_Click()
Dim nRowFlx As Integer
Dim optChk As String
Dim iFila As Integer, nId As Integer
Dim nCantRows As Integer
nId = 1
nRowFlx = MSHFlexCond.row
If MSHFlexCond.rows - 1 >= 1 Then
    If nRowFlx = 1 And MSHFlexCond.rows - 1 < 2 Then
      Call CargaCabeceraMSHFlexCond
      MSHFlexCond.AddItem ""
      MSHFlexCond.RemoveItem 1
      Exit Sub
    ElseIf nRowFlx >= 1 Then
       MSHFlexCond.RemoveItem nRowFlx
    Else
        If chkCondiciones.value = 1 Then
            optChk = "Condiciones"
        Else
            optChk = "Requisitos"
        End If
        If MsgBox("¿Desea limpiar el listado completo de " & optChk & " ...", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
        LimpiaFlex MSHFlexCond
    End If
    For iFila = 1 To MSHFlexCond.rows - 1
        MSHFlexCond.TextMatrix(iFila, 0) = nId
        nId = nId + 1
    Next iFila
End If

LimpiarTxtRango ("byRow")
Call CriterioRango("PorFila")
Call RowBackColor(nRowFlx, "BackNext")
End Sub

Private Sub cmdEditar_Click()
If cmdEditar.Caption = "Cancelar" Then
    chkMatriz.value = 1
    chkIndep.value = 0
    chkMatriz.Enabled = True
    chkIndep.Enabled = True
    chkCondiciones.Enabled = True
    chkRequisitos.Enabled = True
    cmdSubirCondRequ.Enabled = True
    cmdbajarCondRequ.Enabled = True
    cmdElimRowFlex.Enabled = True
    chkHabRang.Enabled = True
    cmdGrabarDatos.Enabled = True
    MSHFlexCond.Enabled = True
    FormateaFlex flxRelacParam
    cmdAgregar.Caption = "Agregar"
    cmdEditar.Caption = "Editar"
    pbEditarGeneral = False
Else
If MSHFlexCond.rows - 1 >= 1 And MSHFlexCond.TextMatrix(1, 1) <> "" Then
    chkMatriz.value = 1
    chkIndep.value = 0
    chkMatriz.Enabled = False
    chkIndep.Enabled = False
    chkCondiciones.Enabled = False
    chkRequisitos.Enabled = False
    cmdSubirCondRequ.Enabled = False
    cmdbajarCondRequ.Enabled = False
    cmdElimRowFlex.Enabled = False
    chkHabRang.Enabled = False
    cmdGrabarDatos.Enabled = False
    MSHFlexCond.Enabled = False
    cmdAgregar.Caption = "Registrar"
    cmdEditar.Caption = "Cancelar"
    Call CargarDatosActualizar
End If
End If
End Sub

Public Sub CargarDatosActualizar()
Dim iColMSH As Integer
Dim iFilaFlx As Long
Dim nPosicion As Long
Dim rs As New ADODB.Recordset
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
nPosicion = MSHFlexCond.row
iFilaFlx = 1
iColMSH = 1
FormateaFlex flxRelacParam
Do While iColMSH <= MSHFlexCond.cols - 1
    If MSHFlexCond.TextMatrix(nPosicion, iColMSH + 1) <> "" Then
        flxRelacParam.AdicionaFila
        flxRelacParam.TextMatrix(iFilaFlx, 1) = MSHFlexCond.TextMatrix(nPosicion, iColMSH)
        flxRelacParam.TextMatrix(iFilaFlx, 2) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 1)
        flxRelacParam.TextMatrix(iFilaFlx, 3) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 2)
        Set rs = oCredCat.ObtieneListaParametro("ParamRang", CDbl(MSHFlexCond.TextMatrix(nPosicion, iColMSH + 1)), "R")
        If rs.RecordCount <> 0 Then
           flxRelacParam.TextMatrix(iFilaFlx, 5) = "R"
        Else
           flxRelacParam.TextMatrix(iFilaFlx, 5) = ""
        End If
        flxRelacParam.TextMatrix(iFilaFlx, 6) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 3)
        flxRelacParam.TextMatrix(iFilaFlx, 7) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 4)
        flxRelacParam.TextMatrix(iFilaFlx, 8) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 5)
        flxRelacParam.TextMatrix(iFilaFlx, 9) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 6)
        flxRelacParam.TextMatrix(iFilaFlx, 10) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 7)
        flxRelacParam.TextMatrix(iFilaFlx, 4) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 8)
        'flxRelacParam.TextMatrix(iFilaFlx, 11) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 10) 'Comentado by NAGL 20191210
        flxRelacParam.TextMatrix(iFilaFlx, 11) = MSHFlexCond.TextMatrix(nPosicion, iColMSH + 10) & Space(100) & IIf(CDbl(MSHFlexCond.TextMatrix(nPosicion, iColMSH + 9)) = 0, "", MSHFlexCond.TextMatrix(nPosicion, iColMSH + 9)) 'NAGL 20190923 Agregó la condicional del código del Módulo a afectar
        flxRelacParam.TextMatrix(iFilaFlx, 12) = IIf(CDbl(MSHFlexCond.TextMatrix(nPosicion, iColMSH + 9)) = 0, "", MSHFlexCond.TextMatrix(nPosicion, iColMSH + 9))
    End If
    iColMSH = iColMSH + 12
    iFilaFlx = iFilaFlx + 1
Loop
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub cmdSubirCondRequ_Click()
Dim psIdParamA, psIdParamB As String
Dim psParametroA, psParametroB As String
Dim psRangDescA, psRangDescB As String
Dim psOperIniA, psOperIniB As String
Dim psRangIniA, psRangIniB As String
Dim psOperFinA, psOperFinB As String
Dim psRangFinA, psRangFinB As String
Dim psUnidRangA, psUnidRangB As String
Dim psValorUnidRA, psValorUnidRB As String
Dim psIdModuloA, psIdModuloB As String
Dim psModuloDescA, psModuloDescB As String
Dim psMovNroA, psMovNroB As String
Dim nPosicion As Long
Dim iCol As Integer
nPosicion = MSHFlexCond.row
iCol = 1
If nPosicion > 1 Then
    Do While iCol <= MSHFlexCond.cols - 1
        psParametroA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol)
        psIdParamA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 1)
        psRangDescA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 2)
        psOperIniA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 3)
        psRangIniA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 4)
        psOperFinA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 5)
        psRangFinA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 6)
        psUnidRangA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 7)
        psValorUnidRA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 8)
        psIdModuloA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 9)
        psModuloDescA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 10)
        psMovNroA = MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 11)
        
        psParametroB = MSHFlexCond.TextMatrix(nPosicion, iCol)
        psIdParamB = MSHFlexCond.TextMatrix(nPosicion, iCol + 1)
        psRangDescB = MSHFlexCond.TextMatrix(nPosicion, iCol + 2)
        psOperIniB = MSHFlexCond.TextMatrix(nPosicion, iCol + 3)
        psRangIniB = MSHFlexCond.TextMatrix(nPosicion, iCol + 4)
        psOperFinB = MSHFlexCond.TextMatrix(nPosicion, iCol + 5)
        psRangFinB = MSHFlexCond.TextMatrix(nPosicion, iCol + 6)
        psUnidRangB = MSHFlexCond.TextMatrix(nPosicion, iCol + 7)
        psValorUnidRB = MSHFlexCond.TextMatrix(nPosicion, iCol + 8)
        psIdModuloB = MSHFlexCond.TextMatrix(nPosicion, iCol + 9)
        psModuloDescB = MSHFlexCond.TextMatrix(nPosicion, iCol + 10)
        psMovNroB = MSHFlexCond.TextMatrix(nPosicion, iCol + 11)
        
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol) = psParametroB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 1) = psIdParamB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 2) = psRangDescB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 3) = psOperIniB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 4) = psRangIniB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 5) = psOperFinB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 6) = psRangFinB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 7) = psUnidRangB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 8) = psValorUnidRB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 9) = psIdModuloB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 10) = psModuloDescB
        MSHFlexCond.TextMatrix(nPosicion - 1, iCol + 11) = psMovNroB
        
        MSHFlexCond.TextMatrix(nPosicion, iCol) = psParametroA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 1) = psIdParamA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 2) = psRangDescA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 3) = psOperIniA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 4) = psRangIniA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 5) = psOperFinA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 6) = psRangFinA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 7) = psUnidRangA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 8) = psValorUnidRA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 9) = psIdModuloA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 10) = psModuloDescA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 11) = psMovNroA
        
        iCol = iCol + 12
    Loop
    Call RowBackColor(nPosicion - 1)
    MSHFlexCond.SetFocus
End If
End Sub

Private Sub cmdbajarCondRequ_Click()
Dim psIdParamA, psIdParamB As String
Dim psParametroA, psParametroB As String
Dim psRangDescA, psRangDescB As String
Dim psOperIniA, psOperIniB As String
Dim psRangIniA, psRangIniB As String
Dim psOperFinA, psOperFinB As String
Dim psRangFinA, psRangFinB As String
Dim psUnidRangA, psUnidRangB As String
Dim psValorUnidRA, psValorUnidRB As String
Dim psIdModuloA, psIdModuloB As String
Dim psModuloDescA, psModuloDescB As String
Dim psMovNroA, psMovNroB As String

Dim nPosicion As Long
Dim iCol As Integer
nPosicion = MSHFlexCond.row
iCol = 1
If nPosicion >= 1 And MSHFlexCond.rows - 1 > nPosicion Then
    Do While iCol <= MSHFlexCond.cols - 1
        psParametroA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol)
        psIdParamA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 1)
        psRangDescA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 2)
        psOperIniA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 3)
        psRangIniA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 4)
        psOperFinA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 5)
        psRangFinA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 6)
        psUnidRangA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 7)
        psValorUnidRA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 8)
        psIdModuloA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 9)
        psModuloDescA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 10)
        psMovNroA = MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 11)
        
        psParametroB = MSHFlexCond.TextMatrix(nPosicion, iCol)
        psIdParamB = MSHFlexCond.TextMatrix(nPosicion, iCol + 1)
        psRangDescB = MSHFlexCond.TextMatrix(nPosicion, iCol + 2)
        psOperIniB = MSHFlexCond.TextMatrix(nPosicion, iCol + 3)
        psRangIniB = MSHFlexCond.TextMatrix(nPosicion, iCol + 4)
        psOperFinB = MSHFlexCond.TextMatrix(nPosicion, iCol + 5)
        psRangFinB = MSHFlexCond.TextMatrix(nPosicion, iCol + 6)
        psUnidRangB = MSHFlexCond.TextMatrix(nPosicion, iCol + 7)
        psValorUnidRB = MSHFlexCond.TextMatrix(nPosicion, iCol + 8)
        psIdModuloB = MSHFlexCond.TextMatrix(nPosicion, iCol + 9)
        psModuloDescB = MSHFlexCond.TextMatrix(nPosicion, iCol + 10)
        psMovNroB = MSHFlexCond.TextMatrix(nPosicion, iCol + 11)
        
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol) = psParametroB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 1) = psIdParamB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 2) = psRangDescB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 3) = psOperIniB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 4) = psRangIniB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 5) = psOperFinB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 6) = psRangFinB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 7) = psUnidRangB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 8) = psValorUnidRB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 9) = psIdModuloB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 10) = psModuloDescB
        MSHFlexCond.TextMatrix(nPosicion + 1, iCol + 11) = psMovNroB
        
        MSHFlexCond.TextMatrix(nPosicion, iCol) = psParametroA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 1) = psIdParamA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 2) = psRangDescA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 3) = psOperIniA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 4) = psRangIniA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 5) = psOperFinA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 6) = psRangFinA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 7) = psUnidRangA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 8) = psValorUnidRA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 9) = psIdModuloA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 10) = psModuloDescA
        MSHFlexCond.TextMatrix(nPosicion, iCol + 11) = psMovNroA
        
        iCol = iCol + 12
    Loop
    Call RowBackColor(nPosicion + 1)
    MSHFlexCond.SetFocus
End If
End Sub

Public Sub CargaCabeceraMSHFlexCond()
    'Const Cabecera As String = "" & "CodParam" & vbTab & "PARÁMETRO" & vbTab & "pTpo" & vbTab & "DESCRIPCIÓN"
    MSHFlexCond.ColWidth(0) = 350
    MSHFlexCond.ColWidth(1) = 2000
    MSHFlexCond.ColWidth(2) = 0 '500
    MSHFlexCond.ColWidth(3) = 2500
    MSHFlexCond.ColWidth(4) = 0 '500
    MSHFlexCond.ColWidth(5) = 0 '500
    MSHFlexCond.ColWidth(6) = 0 '500
    MSHFlexCond.ColWidth(7) = 0 '500
    MSHFlexCond.ColWidth(8) = 0 '500
    MSHFlexCond.ColWidth(9) = 0 '500
    MSHFlexCond.ColWidth(10) = 0 '500
    MSHFlexCond.ColWidth(11) = 0 '500
    MSHFlexCond.cols = 12
    MSHFlexCond.TextMatrix(0, 0) = "N°"
    MSHFlexCond.TextMatrix(0, 1) = "PARÁMETRO"
    MSHFlexCond.TextMatrix(0, 3) = "DESCRIPCIÓN"
End Sub

Private Sub cmdGrabarDatos_Click()
Dim optCondRequ As String, psTpoProd As String
Dim iFila As Integer, iCol As Integer
Dim psRang1Ant As String, psRang2Ant As String, psRang3Ant As String, psRang4Ant As String
Dim pnParCodAnt As Long, pnParValorAnt As Long, pnUnidRangoAnt As Long
Dim pnIdModAnt As Long
Dim psMovNroAnt As String
Dim psRang1 As String, psRang2 As String, psRang3 As String, psRang4 As String
Dim pnParCod As Long, pnParValor As Long, pnUnidRango As Long
Dim pnCodIdDepen As Long
Dim rsRang As New ADODB.Recordset
Dim psTpoRang As String, psTpoRangAnt As String
Dim pnIdModulo As Long
Dim psMovNro As String, lsMovNrOpt As String
Dim oCont As New NContFunciones
Dim psTipoCondReq As String 'NAGL 20190923
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

lsMovNrOpt = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
psTpoProd = Trim(Right(cboTpoProdCar.Text, 15))
If chkCondiciones.value = 1 Then
    optCondRequ = "Cond"
    psTipoCondReq = "Cond"
ElseIf chkRequisitos.value = 1 Then
    optCondRequ = "Req"
    psTipoCondReq = "Requ"
Else
    MsgBox "Debe seleccionar un Tipo de Carácteristica: Condición o Requisito", vbInformation, "Aviso"
    chkCondiciones.SetFocus
    Exit Sub
End If
On Error GoTo GrabarDatosErr 'NAGL 20190923

If (MSHFlexCond.rows) - 1 < 2 And MSHFlexCond.TextMatrix(1, 1) = "" Then Exit Sub
If MsgBox("¿Desea guardar los cambios realizados..", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
Call HabilitarControlesToCargarData(False, "Grabar") 'NAGL 20190923
Call oCredCat.LimpiarDataProducto(optCondRequ, psTpoProd, lsMovNrOpt) 'NAGL 20190923 Agregó lsMovNrOpt
For iFila = 1 To CInt(MSHFlexCond.rows) - 1
    For iCol = 1 To CInt(MSHFlexCond.cols) - 1
    If CInt(MSHFlexCond.cols) - 1 > iCol + 1 Then
        If IsNumeric(MSHFlexCond.TextMatrix(iFila, iCol + 1)) Then
            pnParCod = MSHFlexCond.TextMatrix(iFila, iCol + 1) 'CInt(IIf(MSHFlexCond.TextMatrix(iFila, iCol + 1) = "", 0, MSHFlexCond.TextMatrix(iFila, iCol + 1)))
        Else
            pnParCod = 0
        End If
        If CDbl(pnParCod) <> 0 Then
                psRang1 = MSHFlexCond.TextMatrix(iFila, iCol + 3)
                psRang2 = MSHFlexCond.TextMatrix(iFila, iCol + 4)
                psRang3 = MSHFlexCond.TextMatrix(iFila, iCol + 5)
                psRang4 = MSHFlexCond.TextMatrix(iFila, iCol + 6)
                pnUnidRango = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol + 7) = "", "0", MSHFlexCond.TextMatrix(iFila, iCol + 7)))
                pnParValor = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol + 8) = "", 0, MSHFlexCond.TextMatrix(iFila, iCol + 8)))
                pnIdModulo = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol + 9) = "", 0, MSHFlexCond.TextMatrix(iFila, iCol + 9)))
                psMovNro = lsMovNrOpt
                iCol = iCol + 11
                
            If iCol - 10 > 2 Then 'ante icol - 7
                'If psTpoRangAnt = "ConRang" And psTpoRang = "ConRang" Then
                    pnParCodAnt = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol - 22) = "", "0", MSHFlexCond.TextMatrix(iFila, iCol - 22)))
                    psRang1Ant = MSHFlexCond.TextMatrix(iFila, iCol - 20)
                    psRang2Ant = MSHFlexCond.TextMatrix(iFila, iCol - 19)
                    psRang3Ant = MSHFlexCond.TextMatrix(iFila, iCol - 18)
                    psRang4Ant = MSHFlexCond.TextMatrix(iFila, iCol - 17)
                    pnUnidRangoAnt = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol - 16) = "", "0", MSHFlexCond.TextMatrix(iFila, iCol - 16)))
                    pnParValorAnt = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol - 15) = "", "0", MSHFlexCond.TextMatrix(iFila, iCol - 15))) 'Desde el 9
                    pnIdModAnt = CDbl(IIf(MSHFlexCond.TextMatrix(iFila, iCol - 14) = "", "0", MSHFlexCond.TextMatrix(iFila, iCol - 14)))
                    psMovNroAnt = psMovNro
            End If
            pnCodIdDepen = oCredCat.ObtieneCodIdCondRequi(optCondRequ, psTpoProd, pnParCodAnt, pnParValorAnt, psRang1Ant, psRang2Ant, psRang3Ant, psRang4Ant, pnUnidRangoAnt, pnIdModAnt, psMovNroAnt)
            Call oCredCat.GrabarDatosCondicionRequisitos(optCondRequ, pnCodIdDepen, psTpoProd, pnParCod, pnParValor, psRang1, psRang2, psRang3, psRang4, pnUnidRango, pnIdModulo, psMovNro)
            pnCodIdDepen = 0
            pnParCodAnt = 0
            pnParValorAnt = 0
            psRang1Ant = ""
            psRang2Ant = ""
            psRang3Ant = ""
            psRang4Ant = ""
            pnUnidRangoAnt = 0
            pnIdModAnt = 0
            psMovNroAnt = ""
        End If
    End If
    Next iCol
Next iFila
MsgBox "Los Datos se guardaron satisfactoriamente", vbInformation, "Aviso"
Call HabilitarControlesToCargarData(True, "Grabar") 'NAGL 20190923
'JOEP
Set oCredCat = Nothing
RSClose rsRang
'JOEP
'*****************BEGIN NAGL*********************
Exit Sub
GrabarDatosErr:
    'Call RaiseError(MyUnhandledError, "Error De Grabación")
    Call oCredCat.RegresaDatosAnteriorProducto(optCondRequ, psTpoProd)
    Call CargarListadoCondicionesRequisitos(psTipoCondReq)
    MsgBox "Hubo un inconveniente en la red, Por favor intente realizar nuevamente la última configuración...!", vbExclamation, "Aviso"
    Call HabilitarControlesToCargarData(True, "Grabar")
'*****************NAGL 20190923******************
End Sub

Public Sub CargarListadoCondicionesRequisitos(psTpoCaract As String)
Dim psTpoProd As String
Dim iFila As Integer, iCol As Integer
Dim nCols As Integer
Dim nOrdenMain As Long, nOrdenMainAnt As Long
Dim rsCaract As New ADODB.Recordset
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
iFila = 1
nOrdenMain = 0
psTpoProd = Trim(Right(cboTpoProdCar.Text, 15))
LimpiaFlex MSHFlexCond
Set rsCaract = oCredCat.GetListadoCondicionesRequisitos(psTpoCaract, psTpoProd, "RC")
If rsCaract!nCols <> 0 Then
    nCols = rsCaract!nCols
    MSHFlexCond.cols = nCols
Else
    Call CargaCabeceraMSHFlexCond
    Call CriterioRango
    Exit Sub
End If
Set rsCaract = Nothing
Set rsCaract = oCredCat.GetListadoCondicionesRequisitos(psTpoCaract, psTpoProd)
Do While Not rsCaract.EOF
    nOrdenMain = rsCaract!nOrden
    If nOrdenMainAnt <> nOrdenMain Then
       iCol = 1
       MSHFlexCond.AddItem ""
       If MSHFlexCond.TextMatrix(1, 1) = "" Then
          MSHFlexCond.RemoveItem 1
       End If
    Else
        iFila = iFila - 1
    End If
    'If rsCaract!bEsRango = True Then 'Si es rango
           If iCol = 1 Then
              MSHFlexCond.ColWidth(iCol) = 2000
           Else
              MSHFlexCond.ColWidth(iCol) = 0
           End If
           MSHFlexCond.ColWidth(iCol + 1) = 0
           MSHFlexCond.ColWidth(iCol + 2) = 3000 '2500
           MSHFlexCond.ColWidth(iCol + 3) = 0
           MSHFlexCond.ColWidth(iCol + 4) = 0
           MSHFlexCond.ColWidth(iCol + 5) = 0
           MSHFlexCond.ColWidth(iCol + 6) = 0
           MSHFlexCond.ColWidth(iCol + 7) = 0
           MSHFlexCond.ColWidth(iCol + 8) = 0
           MSHFlexCond.ColWidth(iCol + 9) = 0
           MSHFlexCond.ColWidth(iCol + 10) = 0
           MSHFlexCond.ColWidth(iCol + 11) = 0
    
           MSHFlexCond.TextMatrix(iFila, 0) = CStr(MSHFlexCond.rows - 1)
           MSHFlexCond.TextMatrix(iFila, iCol) = rsCaract!cParCodDesc
           MSHFlexCond.TextMatrix(iFila, iCol + 1) = rsCaract!nParCod
           MSHFlexCond.TextMatrix(iFila, iCol + 2) = rsCaract!cRangoDesc
           MSHFlexCond.TextMatrix(iFila, iCol + 3) = rsCaract!cOperInicio
           MSHFlexCond.TextMatrix(iFila, iCol + 4) = rsCaract!cRangoInicio
           MSHFlexCond.TextMatrix(iFila, iCol + 5) = rsCaract!cOperFin
           MSHFlexCond.TextMatrix(iFila, iCol + 6) = rsCaract!cRangoFin
           MSHFlexCond.TextMatrix(iFila, iCol + 7) = rsCaract!nUndRango
           MSHFlexCond.TextMatrix(iFila, iCol + 8) = rsCaract!cValor
           MSHFlexCond.TextMatrix(iFila, iCol + 9) = rsCaract!nModulo
           MSHFlexCond.TextMatrix(iFila, iCol + 10) = rsCaract!cModuloDesc
           MSHFlexCond.TextMatrix(iFila, iCol + 11) = rsCaract!cMovNro
           If iCol = 1 Then
                MSHFlexCond.TextMatrix(0, iCol) = "PARÁMETRO"
           Else
                MSHFlexCond.TextMatrix(0, iCol) = "DESCRIPCIÓN"
           End If
    
           MSHFlexCond.TextMatrix(0, iCol + 1) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 2) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 3) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 4) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 5) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 6) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 7) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 8) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 9) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 10) = "DESCRIPCIÓN"
           MSHFlexCond.TextMatrix(0, iCol + 11) = "DESCRIPCIÓN"
           
           MSHFlexCond.MergeCol(iCol) = True
           iCol = iCol + 12
    nOrdenMainAnt = nOrdenMain
    iFila = iFila + 1
    rsCaract.MoveNext
Loop
MSHFlexCond.MergeCells = flexMergeFree
MSHFlexCond.MergeCol(1) = True
MSHFlexCond.MergeCol(2) = True
MSHFlexCond.MergeRow(0) = True
Call CriterioRango
Call RowBackColor(MSHFlexCond.rows - 1)
'JOEP
Set oCredCat = Nothing
RSClose rsCaract
'JOEP
End Sub

Public Function LastPosicion(psParParam As String, pnCol As Integer) As Integer
Dim i As Integer
Dim nItem As Integer
nItem = 0
For i = 1 To CInt(MSHFlexCond.rows) - 1 'Para Observar si ya se ha ingresado el Parámetro obtenido
        If Trim(MSHFlexCond.TextMatrix(i, pnCol)) = psParParam Then
           Do While Trim(MSHFlexCond.TextMatrix(i, pnCol)) = psParParam
              i = i + 1
              If i > (MSHFlexCond.rows) - 1 Then
                 Exit Do
              End If
           Loop
           nItem = i
        End If
Next i
LastPosicion = nItem
End Function

'***SECCION RANGO
Private Sub chkHabRang_Click()
    If chkHabRang.value = 1 Then
        Call CriterioRango("PorFila")
    Else
       cboTpoParam.Enabled = False
       cboTpoDato.Enabled = False
       cboUndRang.Enabled = False
       cmdRang.Enabled = False
       cboRang1.Enabled = False
       cboRang2.Enabled = False
       txtRangMin.Enabled = False
       txtRangMax.Enabled = False
       Call LimpiarTxtRango("byRow")
    End If
End Sub

Private Sub cmdRang_Click()
Dim iCol As Integer, iFila As Integer
Dim psTpoParam As Long, nCantCols As Integer
Dim psParamDescrip As String
Dim pnCompareIni As Long
Dim psParamDescripInt As String

'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

If ValidaRango Then
psTpoParam = CDbl(Trim(Right(cboTpoParam.Text, 15)))
psParamDescrip = Trim(Mid(cboTpoParam.Text, 1, Len(cboTpoParam.Text) - 15))
iFila = MSHFlexCond.row
For iCol = 1 To CInt(MSHFlexCond.cols) - 1
    pnCompareIni = iCol + 1
    If pnCompareIni > CInt(MSHFlexCond.cols) - 1 Then
       Exit For
    End If
    If MSHFlexCond.TextMatrix(iFila, pnCompareIni) = CStr(psTpoParam) Then
       psParamDescripInt = oCredCat.ObtieneParamDescripInt(CDbl(psTpoParam), CDbl(IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) = "", 0, MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7))))
       psParamDescrip = IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) <> MSHFlexCond.TextMatrix(iFila, pnCompareIni) And MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) <> "", psParamDescripInt, "")
       MSHFlexCond.TextMatrix(iFila, iCol + 2) = psParamDescrip & " " & Trim(Left(cboRang1.Text, 2)) & " " & txtRangMin.Text & " " & Trim(Left(cboRang2.Text, 2)) & " " & txtRangMax.Text & " " & Mid(cboUndRang.Text, 1, Len(cboUndRang.Text) - Len(Trim(Right(cboUndRang.Text, 10))))
       MSHFlexCond.TextMatrix(iFila, iCol + 3) = Trim(Left(cboRang1.Text, 2))
       MSHFlexCond.TextMatrix(iFila, iCol + 4) = txtRangMin.Text
       MSHFlexCond.TextMatrix(iFila, iCol + 5) = Trim(Left(cboRang2.Text, 2))
       MSHFlexCond.TextMatrix(iFila, iCol + 6) = txtRangMax.Text
       MSHFlexCond.TextMatrix(iFila, iCol + 7) = Trim(Right(cboUndRang.Text, 15))
       MSHFlexCond.TextMatrix(iFila, iCol + 8) = IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) = "", psTpoParam, MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7))
       iCol = iCol + 11
    End If
Next iCol
'*******ValidaControlRegDuplicado******
If ValidaControlRegDuplicado(iFila) = True Then
    MsgBox "Exite un registro con los mismos parámetros y rango ingresado, por favor cambie el rango ingresado!!", vbInformation, "Atención"
    For iCol = 1 To CInt(MSHFlexCond.cols) - 1
    pnCompareIni = iCol + 1
    If pnCompareIni > CInt(MSHFlexCond.cols) - 1 Then
       Exit For
    End If
    If MSHFlexCond.TextMatrix(iFila, pnCompareIni) = CStr(psTpoParam) Then
       psParamDescripInt = oCredCat.ObtieneParamDescripInt(CDbl(psTpoParam), CDbl(IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) = "", 0, MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7))))
       psParamDescrip = IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) <> MSHFlexCond.TextMatrix(iFila, pnCompareIni) And MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) <> "", psParamDescripInt, "")
       MSHFlexCond.TextMatrix(iFila, iCol + 2) = psParamDescrip
       MSHFlexCond.TextMatrix(iFila, iCol + 3) = ""
       MSHFlexCond.TextMatrix(iFila, iCol + 4) = ""
       MSHFlexCond.TextMatrix(iFila, iCol + 5) = ""
       MSHFlexCond.TextMatrix(iFila, iCol + 6) = ""
       MSHFlexCond.TextMatrix(iFila, iCol + 7) = ""
       MSHFlexCond.TextMatrix(iFila, iCol + 8) = IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7) = "", psTpoParam, MSHFlexCond.TextMatrix(iFila, pnCompareIni + 7))
       iCol = iCol + 8
    End If
Next iCol
    cmdRang.SetFocus
    Exit Sub
End If
'**************************************
End If
'JOEP
Set oCredCat = Nothing
'JOEP
End Sub

Public Function ValidaRango() As Boolean
If (Right(cboTpoDato.Text, 1) = 1 Or Right(cboTpoDato.Text, 1) = 2) Then
    If (cboTpoParam.Text = "") Then
        MsgBox "No se ha seleccionado ningún parámetro", vbInformation, "Aviso"
        Exit Function
    ElseIf (Trim(cboRang1.Text) <> "" And txtRangMin = "") Or (Trim(cboRang1.Text) = "" And txtRangMin <> "") Then
        MsgBox "Rango Inicial incorrecto", vbInformation, "Aviso"
        Exit Function
    ElseIf (Trim(cboRang2.Text) <> "" And txtRangMax = "" Or Trim(cboRang2.Text) = "" And txtRangMax <> "") Then
        MsgBox "Rango Final incorrecto", vbInformation, "Aviso"
        Exit Function
    ElseIf (Trim(cboRang1.Text) = "" And txtRangMin = "" And Trim(cboRang2.Text) <> "" And txtRangMax <> "") Then
        MsgBox "Falta establecer el Rango Inicial", vbInformation, "Aviso"
        Exit Function
    ElseIf (Trim(cboRang1.Text) <> "" And txtRangMin <> "") And (Trim(Right(cboUndRang.Text, 15)) = 0 Or cboUndRang.Text = "") Or ((Trim(cboRang1.Text) <> "" And txtRangMin <> "" And Trim(cboRang2.Text) <> "" And txtRangMax <> "") And (Trim(Right(cboUndRang.Text, 15)) = 0 Or cboUndRang.Text = "")) Then
        MsgBox "Falta establecer la Unidad del Rango", vbInformation, "Aviso"
        Exit Function
    End If

    If txtRangMin <> "" And txtRangMax <> "" Then
        If (CDbl(txtRangMin) > CDbl(txtRangMax)) Then
            MsgBox "El Rango Inicial no debe ser mayor al Rango Final", vbInformation, "Aviso"
            Exit Function
        End If
    End If
ElseIf Right(cboTpoDato.Text, 1) = 3 Then
    If (Trim(cboRang1.Text) <> "" And (txtRangMin = "" Or txtRangMin = "__/__/____")) Or (Trim(cboRang1.Text) = "" And (txtRangMin <> "" Or txtRangMin <> "__/__/____")) Then
       MsgBox "Rango Inicial incorrecto", vbInformation, "Aviso"
       Exit Function
    ElseIf (Trim(cboRang1.Text) <> "" And (txtRangMin = "" Or txtRangMin = "__/__/____")) Then
       If ValFecha(txtRangMin) = False Then
            txtRangMin = "__/__/___"
            txtRangMin.SetFocus
            Exit Function
       End If
    ElseIf (Trim(cboRang2.Text) <> "" And (txtRangMax = "" Or txtRangMax = "__/__/____")) Or (Trim(cboRang2.Text) = "" And (txtRangMax <> "" And txtRangMax <> "__/__/____")) Then
       MsgBox "Rango Final incorrecto", vbInformation, "Aviso"
       Exit Function
    ElseIf (Trim(cboRang2.Text) <> "" And (txtRangMax = "" Or txtRangMax = "__/__/____")) Then
        If ValFecha(txtRangMax) = False Then
           txtRangMax = "__/__/____"
           txtRangMax.SetFocus
           Exit Function
        End If
    ElseIf (Trim(cboRang1.Text) <> "" And (txtRangMin <> "" Or txtRangMin <> "__/__/____")) And (Trim(Right(cboUndRang.Text, 15)) = 0 Or cboUndRang.Text = "") Or ((Trim(cboRang1.Text) <> "" And (txtRangMin <> "" Or txtRangMin <> "__/__/____") And Trim(cboRang2.Text) <> "" And (txtRangMax <> "" Or txtRangMax <> "__/__/____")) And (Trim(Right(cboUndRang.Text, 15)) = 0 Or cboUndRang.Text = "")) Then
        MsgBox "Falta establecer la Unidad del Rango", vbInformation, "Aviso"
        Exit Function
    End If
    If (txtRangMin <> "" And txtRangMin <> "__/__/____") And (txtRangMax <> "" And txtRangMax <> "__/__/____") Then
        If (CDate(txtRangMin) > CDate(txtRangMax)) Then
            MsgBox "El Rango Inicial es mayor al Rango Final", vbInformation, "Aviso"
            Exit Function
        End If
    End If
    If txtRangMin = "__/__/____" Then
       txtRangMin = ""
    ElseIf txtRangMax = "__/__/____" Then
       txtRangMax = ""
    End If
End If
ValidaRango = True
End Function

Public Sub LimpiarTxtRango(Optional psClean As String)
If psClean = "byRow" Then
    cboTpoParam.Clear
    cboRang1.Clear
    cboRang2.Clear
    cboTpoDato.Clear
    cboUndRang.Clear
Else
    cboTpoParam.ListIndex = -1
    cboTpoDato.ListIndex = -1
    cboRang1.ListIndex = -1
    cboRang2.ListIndex = -1
    cboUndRang.ListIndex = -1
End If
    txtRangMin.Text = ""
    txtRangMax.Text = ""
End Sub

Public Sub CriterioRango(Optional psCritRang As String)
Dim RsCbo As New ADODB.Recordset
Dim i As Integer, iCol As Integer, iFila As Integer
Dim pnTpoParam As Long
Dim pnCompareIni As Long
Dim pnParCodInt As Long
Dim psHabRang As String
Dim nFilaRang As Long
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
psHabRang = ""
nFilaRang = MSHFlexCond.row
'SECCIÓN DE RANGOS
cboTpoParam.Clear
If nFilaRang <= CInt(MSHFlexCond.rows) - 1 Then
    If chkHabRang.value = 1 Then
        psHabRang = "HabRang"
    End If
    iFila = nFilaRang
    For iCol = 1 To CInt(MSHFlexCond.cols) - 1
        pnCompareIni = iCol + 1
        If pnCompareIni > CInt(MSHFlexCond.cols) - 1 Then
           Exit For
        End If
        If IsNumeric(MSHFlexCond.TextMatrix(iFila, pnCompareIni)) Then
         Set RsCbo = oCredCat.ObtieneListaParametro("ParamRang", CDbl(IIf(MSHFlexCond.TextMatrix(iFila, pnCompareIni) = "", 0, MSHFlexCond.TextMatrix(iFila, pnCompareIni)))) ',psHabRang
            If RsCbo.RecordCount <> 0 Then
                For i = 0 To cboTpoParam.ListCount - 1
                   If Trim(Right(cboTpoParam.List(i), 10)) = CStr(MSHFlexCond.TextMatrix(iFila, pnCompareIni)) Then
                       Exit For
                   End If
                Next i
                cboTpoParam.AddItem Trim(RsCbo!Descrip) & Space(100) & Trim(RsCbo!Tipo)
            End If
         End If
         iCol = iCol + 11
    Next iCol
    
    If psCritRang = "" And MSHFlexCond.rows - 1 <= 0 Then
        If flxRelacParam.TextMatrix(1, 2) <> "" Then
            If cboTpoParam.ListCount - 1 < 0 Then
                For i = 1 To CInt(flxRelacParam.rows) - 1 'Para Observar si ya se ha ingresado el Parámetro obtenido
                    pnTpoParam = CDbl(Trim(flxRelacParam.TextMatrix(i, 2)))
                    Set RsCbo = oCredCat.ObtieneListaParametro("ParamRang", pnTpoParam) ', psHabRang
                    If RsCbo.RecordCount <> 0 Then
                        cboTpoParam.AddItem Trim(RsCbo!Descrip) & Space(100) & Trim(RsCbo!Tipo)
                    End If
                Next i
            End If
        End If
    End If
    Set RsCbo = Nothing
    
    If cboTpoParam.ListCount - 1 >= 0 And psHabRang <> "" Then
       cboTpoParam.Enabled = True
    Else
       chkHabRang.value = 0
       cboTpoParam.Enabled = False
       cboTpoDato.Enabled = False
       cboUndRang.Enabled = False
       cmdRang.Enabled = False
       cboRang1.Enabled = False
       cboRang2.Enabled = False
       txtRangMin.Enabled = False
       txtRangMax.Enabled = False
       chkHabRang.Enabled = True
    End If
    
    If cboTpoParam.Enabled = True And psHabRang = "" Then
       chkHabRang.Enabled = False
    End If
    
        cboTpoDato.Clear
        cboUndRang.Clear
        For i = 0 To cboTpoParam.ListCount - 1
            pnParCodInt = CDbl(Trim(Right(cboTpoParam.List(i), 10)))
            Call ValidandoUnidadRango(pnParCodInt, "TpoDato")
            Call ValidandoUnidadRango(pnParCodInt, "UnidRango")
        Next i
        Set RsCbo = Nothing
        Set RsCbo = oCredCat.CargarCboCondiciones(0, "UnidRangoIni")
        Call LlenarComboRS(RsCbo, cboRang1)
        Set RsCbo = Nothing
        Set RsCbo = oCredCat.CargarCboCondiciones(0, "UnidRangoFin")
        Call LlenarComboRS(RsCbo, cboRang2)
        Set RsCbo = Nothing
End If
'JOEP
Set oCredCat = Nothing
RSClose RsCbo
'JOEP
End Sub

Public Sub ValidandoUnidadRango(pnParCodInt As Long, psDescrip As String)
    Dim i As Integer
    Dim RsCbo As New ADODB.Recordset
    Dim nExist As Integer
    'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
    Set RsCbo = oCredCat.CargarCboCondiciones(pnParCodInt, psDescrip)
    nExist = 0
    If psDescrip = "TpoDato" Then
        Do While Not RsCbo.EOF
           If cboTpoDato.ListCount = 0 Then
                cboTpoDato.AddItem Trim(RsCbo!Descrip) & Space(100) & Trim(RsCbo!Tipo)
           Else
                For i = 0 To cboTpoDato.ListCount - 1
                    If Trim(Right(cboTpoDato.List(i), 10)) = Trim(RsCbo!Tipo) Then
                       nExist = nExist + 1
                       Exit For
                    End If
                Next i
                If nExist = 0 Then
                    cboTpoDato.AddItem Trim(RsCbo!Descrip) & Space(100) & Trim(RsCbo!Tipo)
                End If
           End If
        RsCbo.MoveNext
       Loop
    Else
        Do While Not RsCbo.EOF
           If cboUndRang.ListCount = 0 Then
                cboUndRang.AddItem Trim(RsCbo!Descrip) & Space(100) & Trim(RsCbo!Tipo)
           Else
                For i = 0 To cboUndRang.ListCount - 1
                    If Trim(Right(cboUndRang.List(i), 10)) = Trim(RsCbo!Tipo) Then
                       nExist = nExist + 1
                       Exit For
                    End If
                Next i
                If nExist = 0 Then
                    cboUndRang.AddItem Trim(RsCbo!Descrip) & Space(100) & Trim(RsCbo!Tipo)
                End If
           End If
        RsCbo.MoveNext
       Loop
    End If
'JOEP
Set oCredCat = Nothing
RSClose RsCbo
'JOEP
End Sub

Private Sub cboTpoParam_Click()
Dim iCol As Integer, iFila As Integer, i As Integer
Dim psTpoParam As Long, nCantCols As Integer, pnParCodInt As Long
Dim pnCompareIni As Long
If cboTpoParam.Text = "" Then Exit Sub
psTpoParam = Trim(Right(cboTpoParam.Text, 15))
iFila = MSHFlexCond.row

If cboTpoParam.ListCount <> 0 Then
    cboTpoDato.Clear
    cboUndRang.Clear
    pnParCodInt = CDbl(Trim(Right(cboTpoParam.Text, 10)))
    Call ValidandoUnidadRango(pnParCodInt, "TpoDato")
    Call ValidandoUnidadRango(pnParCodInt, "UnidRango")
End If

For iCol = 1 To CInt(MSHFlexCond.cols) - 1
    pnCompareIni = iCol + 1
    If pnCompareIni > CInt(MSHFlexCond.cols) - 1 Then
       Exit For
    End If
    If MSHFlexCond.TextMatrix(iFila, pnCompareIni) = CStr(psTpoParam) Then
       For i = 0 To cboRang1.ListCount - 1
            If Trim(Left(cboRang1.List(i), 5)) = Trim(MSHFlexCond.TextMatrix(iFila, iCol + 3)) Then
               cboRang1.ListIndex = i
               Exit For
            End If
       Next i
       txtRangMin.Text = MSHFlexCond.TextMatrix(iFila, iCol + 4)
       For i = 0 To cboRang2.ListCount - 1
            If Trim(Left(cboRang2.List(i), 5)) = Trim(MSHFlexCond.TextMatrix(iFila, iCol + 5)) Then
               cboRang2.ListIndex = i
               Exit For
            End If
       Next i
       txtRangMax.Text = MSHFlexCond.TextMatrix(iFila, iCol + 6)
       For i = 0 To cboUndRang.ListCount - 1
            If Trim(Right(cboUndRang.List(i), 10)) = IIf(Trim(MSHFlexCond.TextMatrix(iFila, iCol + 7)) = "", "0", Trim(MSHFlexCond.TextMatrix(iFila, iCol + 7))) Then
               cboUndRang.ListIndex = i
               psTpoParamIni = "TpoIni"
               cboTpoDato.ListIndex = 0
               Exit For
            End If
       Next i
       iCol = iCol + 8
    End If
Next iCol

If cboTpoParam.ListIndex > -1 Then
    cboTpoDato.Enabled = True
    cboUndRang.Enabled = True
    cmdRang.Enabled = True
    cboRang1.Enabled = True
    cboRang2.Enabled = True
    txtRangMin.Enabled = True
    txtRangMax.Enabled = True
End If
End Sub

Private Sub cboTpoDato_Click()
Dim psTpoDato As String
psTpoDato = Trim(Right(cboTpoDato.Text, 3))
If psTpoParamIni <> "TpoIni" Then
    Select Case psTpoDato
        Case "1"
            txtRangMin = ""
            txtRangMax = ""
        Case "2"
            txtRangMin = ""
            txtRangMax = ""
        Case "3"
            txtRangMin = "__/__/____"
            txtRangMax = "__/__/____"
        Case "4"
            txtRangMin = ""
            txtRangMax = ""
    End Select
    cboRang1.ListIndex = -1
    cboRang2.ListIndex = -1
    cboUndRang.ListIndex = -1
End If
psTpoParamIni = ""
End Sub

'***END RANGO
Public Function ValExistParamDet(psParParamDet As String, psNameDet As String) As Boolean
Dim i As Integer
Dim nItem As Integer
nItem = 0
For i = 1 To CInt(MSHFlexCond.rows) - 1 'Para Observar si ya se ha ingresado el Parámetro obtenido
    If CInt(Trim(MSHFlexCond.TextMatrix(i, 2))) = psParParamDet Then
        MsgBox "El Parámetro " & """" & psNameDet & """" & " ya ha sido ingresado", vbInformation, "Aviso"
        Exit Function
    End If
Next i
ValExistParamDet = False
End Function

'Propiedades

Public Sub RowBackColor(nFila As Integer, Optional psBtn As String)
Dim nCol As Integer
Dim nRow As Integer
If (psBtn <> "BackNext") Then
    If nFila <= MSHFlexCond.rows - 1 And nFila > 0 Then
        For nRow = 1 To MSHFlexCond.rows - 1
            For nCol = 1 To MSHFlexCond.cols - 1
                MSHFlexCond.row = nRow
                MSHFlexCond.Col = nCol
                MSHFlexCond.CellBackColor = vbWhite
                MSHFlexCond.CellForeColor = vbBlack
            Next nCol
         Next nRow
        MSHFlexCond.row = nFila
        For nCol = 1 To MSHFlexCond.cols - 1
            MSHFlexCond.Col = nCol
            MSHFlexCond.CellBackColor = &H8000000D    'RGB(209, 222, 253)
            MSHFlexCond.CellForeColor = vbWhite
        Next nCol
    End If
Else
    For nRow = 1 To MSHFlexCond.rows - 1
        For nCol = 1 To MSHFlexCond.cols - 1
            MSHFlexCond.row = nRow
            MSHFlexCond.Col = nCol
            MSHFlexCond.CellBackColor = vbWhite
            MSHFlexCond.CellForeColor = vbBlack
        Next nCol
    Next nRow
    MSHFlexCond.row = IIf(nFila = 1, nFila + 1, nFila - 1)
    For nCol = 1 To MSHFlexCond.cols - 1
        MSHFlexCond.Col = nCol
        MSHFlexCond.CellBackColor = &H8000000D     'RGB(209, 222, 253)
        MSHFlexCond.CellForeColor = vbWhite
    Next nCol
End If
End Sub

Private Sub cboTpoParam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         cboTpoDato.SetFocus
    End If
End Sub

Private Sub cboTpoDato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         cboRang1.SetFocus
    End If
End Sub

Private Sub cboRang1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtRangMin.SetFocus
    End If
End Sub

Private Sub txtRangMin_GotFocus()
    fEnfoque txtRangMin
End Sub

Private Sub txtRangMin_KeyPress(KeyAscii As Integer)
Dim psTpoDato As String
psTpoDato = Trim(Right(cboTpoDato.Text, 3))
If psTpoDato = "1" Or psTpoDato = "2" Then
    KeyAscii = NumerosDecimales(txtRangMin, KeyAscii)
End If
    If KeyAscii = 13 Then
         If psTpoDato = "1" Or psTpoDato = "2" Then
            cboRang2.SetFocus
         ElseIf psTpoDato = "3" Then
            If ValFecha(txtRangMin) Then
                cboRang2.SetFocus
            End If
         Else
            txtRangMin.MaxLength = 25
            cboRang2.SetFocus
         End If
    End If
End Sub

Private Sub txtRangMin_LostFocus()
Dim psTpoDato As String
psTpoDato = Trim(Right(cboTpoDato.Text, 3))
    If psTpoDato <> "3" Then
       txtRangMin.MaxLength = 25
    End If
End Sub

Private Sub cboRang2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         txtRangMax.SetFocus
    End If
End Sub

Private Sub txtRangMax_GotFocus()
    fEnfoque txtRangMax
End Sub

Private Sub txtRangMax_KeyPress(KeyAscii As Integer)
Dim psTpoDato As String
psTpoDato = Trim(Right(cboTpoDato.Text, 3))
If psTpoDato = "1" Or psTpoDato = "2" Then
    KeyAscii = NumerosDecimales(txtRangMax, KeyAscii)
End If
If KeyAscii = 13 Then
    If psTpoDato = "1" Or psTpoDato = "2" Then
       cboUndRang.SetFocus
    ElseIf psTpoDato = "3" Then
       If ValFecha(txtRangMax) Then
           cboUndRang.SetFocus
       End If
    Else
       txtRangMax.MaxLength = 25
       cboUndRang.SetFocus
    End If
End If
End Sub

Private Sub txtRangMax_LostFocus()
Dim psTpoDato As String
psTpoDato = Trim(Right(cboTpoDato.Text, 3))
   
    If psTpoDato <> "3" Then
       txtRangMax.MaxLength = 25
    End If
End Sub

Private Sub cboUndRang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         cmdRang.SetFocus
    End If
End Sub


'JOEP20181204 ERS034-CP
Private Sub cmbProdCheckLis_Click()
Dim objDCred As COMDCredito.DCOMCatalogoProd
Dim rsDocDetCond As ADODB.Recordset
Dim i As Integer

Set objDCred = New COMDCredito.DCOMCatalogoProd
Set rsDocDetCond = objDCred.CargaFlexCheckList(1, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)))

LimpiaFlex feDocumentos
LimpiaFlex feDetalle
LimpiaFlex feNivelesCheckList
LimpiaFlex feCondicionesCheckList
If Not (rsDocDetCond.BOF And rsDocDetCond.EOF) Then
    For i = 1 To rsDocDetCond.RecordCount
        feDocumentos.AdicionaFila
        feDocumentos.TextMatrix(i, 1) = rsDocDetCond!cItem
        feDocumentos.TextMatrix(i, 2) = rsDocDetCond!cDescripcion
        rsDocDetCond.MoveNext
    Next i
Else
    Call habilitarControlesCheckList(1, False)
End If

Set objDCred = Nothing
RSClose rsDocDetCond
End Sub

Private Sub cmdAgregarConfDoc_Click()
Dim objDCred As COMDCredito.DCOMCatalogoProd
Dim rsDoc As ADODB.Recordset
Dim i As Integer

For i = 1 To feDocumentos.rows - 1
    If feDocumentos.TextMatrix(i, 1) <> feDocumentos.row Then
        feDocumentos.TextMatrix(i, 3) = ""
    End If
Next i

Set objDCred = New COMDCredito.DCOMCatalogoProd
Set rsDoc = objDCred.CargaFlexCheckList(6, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), , feDocumentos.TextMatrix(feDocumentos.row, 1))
If Not (rsDoc.BOF And rsDoc.EOF) Then
    feDocumentos.TextMatrix(feDocumentos.row, 3) = rsDoc!nCantConfMax + 1
Else
    feDocumentos.TextMatrix(feDocumentos.row, 3) = feDocumentos.TextMatrix(feDocumentos.row, 3) + 1
End If

    Call habilitarControlesCheckList(5, False)
    Call LimpiaFlex(feNivelesCheckList)
    Call LimpiaFlex(feDetalle)
    Call LimpiaFlex(feCondicionesCheckList)
RSClose rsDoc
End Sub

Private Sub cmbCatCheckLis_Click()
Dim rsProd As New ADODB.Recordset
Dim psRelac As String
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

    psRelac = Mid(Trim(cmbCatCheckLis.Text), Len(Trim(cmbCatCheckLis.Text)) - 2, 1)
    Set rsProd = oCredCat.CargarTipoProd("Prod", psRelac)
    Call LlenarComboRS(rsProd, cmbProdCheckLis)
    cboTpoProd.ListIndex = 0
    Call LimpiaFlex(feDocumentos)
    Call LimpiaFlex(feDetalle)
    Call LimpiaFlex(feNivelesCheckList)
    Call LimpiaFlex(feCondicionesCheckList)
'JOEP
Set oCredCat = Nothing
RSClose rsProd
'JOEP
End Sub

Public Sub CargaListaParametrosCheckList()
Dim rs As New ADODB.Recordset
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP
ListParametroCheckList.Clear
If ListParametroCheckList.ListIndex = -1 Then
    Set rs = oCredCat.ObtieneListaParametroCheckList(0)
    Set oCredCat = Nothing
    Do While Not rs.EOF
        ListParametroCheckList.AddItem Trim(rs!Descrip) & Space(100) & Trim(rs!Tipo)
        rs.MoveNext
    Loop
End If
'JOEP
Set oCredCat = Nothing
RSClose rs
'JOEP
End Sub

Private Sub cmbFinCheckList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFinCheckList.SetFocus
    End If
End Sub

Private Sub cmbUniMedCheckList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarMontoCheckList.SetFocus
    End If
End Sub

Private Sub cmdGrabarcheckList_Click()
Dim objGrabar As COMNCredito.NCOMCatalogoProd
Set objGrabar = New COMNCredito.NCOMCatalogoProd
Dim bGrabar As Boolean
Dim nItemPri As Integer
Dim i As Integer
Dim j As Integer
Dim nCantInte As Integer
Dim cNrMov As String
Dim cItemNivel As String
Dim cItemDet As String
Dim cItemDoc As String
cNrMov = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
j = 2
nCantInte = 0
If Not ValidaCheckList(4, 0) Then Exit Sub

For i = 1 To feDocumentos.rows - 1
    cItemDet = IIf(InStr(feDetalle.TextMatrix(feDetalle.row, 1), ".") = 0, feDetalle.TextMatrix(feDetalle.row, 1) & ".0", feDetalle.TextMatrix(feDetalle.row, 1))
    If feDocumentos.TextMatrix(i, 1) = Mid(cItemDet, InStr(cItemDet, ".") - (InStr(cItemDet, ".") - 1), InStr(cItemDet, ".") - 1) Then
        nCantInte = nCantInte + 1
    End If
Next i

If CInt((feNivelesCheckList.rows - 1)) >= 1 And feNivelesCheckList.TextMatrix(1, 1) <> "" Then
    For i = 1 To feNivelesCheckList.rows - 1
    cItemNivel = IIf(InStr(feNivelesCheckList.TextMatrix(i, 1), ".") = 0, feNivelesCheckList.TextMatrix(i, 1) & ".0", feNivelesCheckList.TextMatrix(i, 1))
    cItemDet = IIf(InStr(feDetalle.TextMatrix(feDetalle.row, 1), ".") = 0, feDetalle.TextMatrix(feDetalle.row, 1) & ".0", feDetalle.TextMatrix(feDetalle.row, 1))
    If Mid(cItemNivel, InStr(cItemNivel, ".") - (InStr(cItemNivel, ".") - 1), InStr(cItemNivel, ".") - 1) = Mid(cItemDet, InStr(cItemDet, ".") - (InStr(cItemDet, ".") - 1), InStr(cItemDet, ".") - 1) Then
            nCantInte = nCantInte + 1
        End If
    Next i
End If

ReDim pMatDocumentos(nCantInte, 5)
For i = 1 To (feDocumentos.rows - 1)
    cItemDet = IIf(InStr(feDetalle.TextMatrix(feDetalle.row, 1), ".") = 0, feDetalle.TextMatrix(feDetalle.row, 1) & ".0", feDetalle.TextMatrix(feDetalle.row, 1))
    cItemDoc = IIf(InStr(feDocumentos.TextMatrix(i, 1), ".") = 0, feDocumentos.TextMatrix(i, 1) & ".0", feDocumentos.TextMatrix(i, 1))
    If Mid(cItemDoc, InStr(cItemDoc, ".") - (InStr(cItemDoc, ".") - 1), InStr(cItemDoc, ".") - 1) = Mid(cItemDet, InStr(cItemDet, ".") - (InStr(cItemDet, ".") - 1), InStr(cItemDet, ".") - 1) Then
        If InStr(feDocumentos.TextMatrix(i, 1), ".") = 0 Then
            pMatDocumentos(1, 1) = feDocumentos.TextMatrix(i, 1)
            pMatDocumentos(1, 2) = feDocumentos.TextMatrix(i, 2)
            pMatDocumentos(1, 3) = feDocumentos.TextMatrix(i, 3)
            pMatDocumentos(1, 4) = feDocumentos.TextMatrix(i, 4)
        End If
    End If
Next i

If feNivelesCheckList.TextMatrix(1, 1) <> "" Then
For i = 1 To (feNivelesCheckList.rows - 1)
    cItemNivel = IIf(InStr(feNivelesCheckList.TextMatrix(i, 1), ".") = 0, feNivelesCheckList.TextMatrix(i, 1) & ".0", feNivelesCheckList.TextMatrix(i, 1))
    cItemDet = IIf(InStr(feDetalle.TextMatrix(feDetalle.row, 1), ".") = 0, feDetalle.TextMatrix(feDetalle.row, 1) & ".0", feDetalle.TextMatrix(feDetalle.row, 1))
    If Mid(cItemNivel, InStr(cItemNivel, ".") - (InStr(cItemNivel, ".") - 1), InStr(cItemNivel, ".") - 1) = Mid(cItemDet, InStr(cItemDet, ".") - (InStr(cItemDet, ".") - 1), InStr(cItemDet, ".") - 1) Then
        pMatDocumentos(j, 1) = feNivelesCheckList.TextMatrix(i, 1)
        pMatDocumentos(j, 2) = feNivelesCheckList.TextMatrix(i, 2)
        pMatDocumentos(j, 3) = feNivelesCheckList.TextMatrix(i, 3)
        pMatDocumentos(j, 4) = feNivelesCheckList.TextMatrix(i, 4)
        pMatDocumentos(j, 5) = Trim(Right(feNivelesCheckList.TextMatrix(i, 5), 5))
        j = j + 1
    End If
Next i
End If

i = 1
ReDim pMatDetalle(feDetalle.rows - 1, 3)
For i = 1 To feDetalle.rows - 1
    pMatDetalle(i, 1) = feDetalle.TextMatrix(i, 1)
    pMatDetalle(i, 2) = feDetalle.TextMatrix(i, 2)
    pMatDetalle(i, 3) = IIf(feDetalle.TextMatrix(i, 3) = ".", 1, 0)
Next i

i = 1
ReDim pMatCondicion(feCondicionesCheckList.rows - 1, 7)
For i = 1 To feCondicionesCheckList.rows - 1
    pMatCondicion(i, 1) = feCondicionesCheckList.TextMatrix(i, 2)
    pMatCondicion(i, 2) = feCondicionesCheckList.TextMatrix(i, 4)
    pMatCondicion(i, 3) = feCondicionesCheckList.TextMatrix(i, 5)
    pMatCondicion(i, 4) = feCondicionesCheckList.TextMatrix(i, 6)
    pMatCondicion(i, 5) = feCondicionesCheckList.TextMatrix(i, 7)
    pMatCondicion(i, 6) = feCondicionesCheckList.TextMatrix(i, 8)
    pMatCondicion(i, 7) = feCondicionesCheckList.TextMatrix(i, 9)
Next i

    bGrabar = objGrabar.GrabarCheckList(Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), pMatDocumentos, pMatDetalle, pMatCondicion, cNrMov, 1)
    If bGrabar = True Then
        MsgBox "Se registraron correctamente los datos", vbInformation, "Exito"
    Else
        MsgBox "Error al registrar los datos", vbInformation, "Error"
    End If
End Sub

Private Sub cmdQuitarDocNivelesCheckList_Click()
Dim cItemPriNiv As String
Dim i As Integer
Dim dato As String

If MsgBox("¿Esta seguro que desea eliminar el Nivel seleccionado?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub

If feNivelesCheckList.TextMatrix(1, 1) = "" Then Exit Sub

cItemPriNiv = feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1)

If Not ValidaCheckList(5, cItemPriNiv) Then Exit Sub

For i = 1 To feDetalle.rows - 1
    If cItemPriNiv = feDetalle.TextMatrix(i, 1) Then
      feDetalle.EliminaFila (IIf(feDetalle.TextMatrix(i, 0) = "", 0, feDetalle.TextMatrix(i, 0)))
    End If
Next i
    
feNivelesCheckList.EliminaFila (feNivelesCheckList.row)
    
End Sub

Private Sub txtFinCheckList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbUniMedCheckList.SetFocus
    End If
End Sub

Private Sub txtInicioCheckList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbFinCheckList.SetFocus
    End If
End Sub

Private Sub cmbIniCheckList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInicioCheckList.SetFocus
    End If
End Sub

Private Sub cmdAgregarDet_Click()
Dim Det As String
Dim Item As String
Dim iTemDet As String
'Verifica si se selecciono el N° Conf. del Documento Principal
If feDocumentos.TextMatrix(feDocumentos.row, 3) = "" Then
    Item = feDocumentos.TextMatrix(feDocumentos.row, 1)
    MsgBox "Seleccione el N° Conf. del Documento Principal" & Chr(13) & Item & ".- " & feDocumentos.TextMatrix(feDocumentos.row, 2) & "", vbInformation, "Aviso"
    feDocumentos.row = feDocumentos.row
    feDocumentos.Col = 3
    Exit Sub
End If

'Pregunta si se registrara los Niveles
If MsgBox("¿Marque [SI], si registrara detalle del Documento Principal" & Chr(13) & " Marque [NO], si registrara detalle del Documento Niveles?", vbYesNo + vbQuestion, "Atención") = vbYes Then
        
    Item = feDocumentos.TextMatrix(feDocumentos.row, 1)
    
    If feNivelesCheckList.TextMatrix(1, 1) <> "" Then
        If CInt(Item) <> CInt(IIf(feDetalle.TextMatrix(feDetalle.row, 1) = "", 0, feDetalle.TextMatrix(feDetalle.row, 1))) Then
            LimpiaFlex feDetalle
            LimpiaFlex feCondicionesCheckList
        End If
    End If
        
    MsgBox "Se ingresara el detalle del [Documento Principal seleccionado]: " & Chr(13) & Item & " - " & feDocumentos.TextMatrix(feDocumentos.row, 2) & ", Conf. N° " & feDocumentos.TextMatrix(feDocumentos.row, 3), vbInformation, "Aviso"
    Det = InputBox(Item & " - " & feDocumentos.TextMatrix(feDocumentos.row, 2) & ", Conf. N° " & feDocumentos.TextMatrix(feDocumentos.row, 3), "Detalle del Documento")
Else
    
    If feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) = "" Then
       MsgBox "Ingrese los Documentos Niveles", vbInformation, "Aviso"
       Exit Sub
    End If
    
    If Len(feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1)) > 1 Then
        If feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 4) = "" Then
            MsgBox "Ingrese el Nivel del panel Documento Niveles: " & Chr(13) & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) & " - " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2), vbInformation, "Aviso"
            feNivelesCheckList.SetFocus
            feNivelesCheckList.row = feNivelesCheckList.row
            feNivelesCheckList.Col = 4
            Exit Sub
        End If
        If feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 5) = "" Then
            MsgBox "Ingrese el Tipo de Documento del panel Documento Niveles" & Chr(13) & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) & " - " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2), vbInformation, "Aviso"
            feNivelesCheckList.SetFocus
            feNivelesCheckList.row = feNivelesCheckList.row
            feNivelesCheckList.Col = 5
            Exit Sub
        End If
    End If
    
    Item = IIf(feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) = "", "0.0", IIf(InStr(feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1), ".") = 0, feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) & ".0", feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1)))
    iTemDet = IIf(feDetalle.TextMatrix(feDetalle.row, 1) = "", "0.0", IIf(InStr(feDetalle.TextMatrix(feDetalle.row, 1), ".") = 0, feDetalle.TextMatrix(feDetalle.row, 1) & ".0", feDetalle.TextMatrix(feDetalle.row, 1)))
    
    If IIf(InStr(Item, ".") = 0, Item, Mid(Item, InStr(Item, ".") - (InStr(Item, ".") - 1), InStr(iTemDet, ".") - 1)) <> IIf(InStr(iTemDet, ".") = 0, iTemDet, Mid(iTemDet, InStr(iTemDet, ".") - (InStr(iTemDet, ".") - 1), InStr(iTemDet, ".") - 1)) Then
        LimpiaFlex feDetalle
        LimpiaFlex feCondicionesCheckList
    End If
        
    MsgBox "Se ingresara el detalle del [Nivel seleccionado]: " & Chr(13) & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) & " - " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2), vbInformation, "Aviso"
    Det = InputBox(feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) & " - " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2) & ", Conf. N° " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 3) & ", Nivel N° " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 4) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "Ingrese el Detalle", "Detalle de Documento Niveles")
End If

If Det <> "" Then
    feDetalle.AdicionaFila
    feDetalle.TextMatrix(feDetalle.row, 1) = Item
    feDetalle.TextMatrix(feDetalle.row, 2) = UCase(Det)
    Call habilitarControlesCheckList(3, True)
End If

End Sub

Private Sub cmdAgregarMontoCheckList_Click()
If Not MensajesCheckList(3) Then Exit Sub

    If Trim(Right(feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 1), 9)) = "4000" Then
        feCondicionesCheckList.ListaControles = "0-0-0-0-0-0"
    Else
        feCondicionesCheckList.ListaControles = "0-0-0-3-0-0"
    End If
    
    If Trim(Right(feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 1), 10)) = "4000" Then
        feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 5) = Left(cmbIniCheckList.Text, 2)
        feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 6) = txtInicioCheckList.Text
        feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 7) = Left(cmbFinCheckList.Text, 2)
        feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 8) = txtFinCheckList.Text
        feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 9) = Trim(Right(cmbUniMedCheckList.Text, 15))
    End If
    
    feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 3) = Left(cmbIniCheckList.Text, 2) & Space(2) & txtInicioCheckList.Text & Space(2) & Left(cmbFinCheckList.Text, 2) & Space(2) & txtFinCheckList.Text & Space(2) & Trim(Left(cmbUniMedCheckList.Text, 15))
    
    cmbIniCheckList.ListIndex = -1
    txtInicioCheckList.Text = ""
    cmbFinCheckList.ListIndex = -1
    txtFinCheckList.Text = ""
    cmbUniMedCheckList.ListIndex = -1
End Sub

Private Sub cmdAgregarNivCheckList_Click()
Dim Nivel As String
Dim DocPri As Integer
Dim DocSubIndice As Integer
Dim i As Integer
Dim Item As String
Dim objDCred As COMDCredito.DCOMCatalogoProd
Dim rsDoc As ADODB.Recordset

DocSubIndice = 0

If Not MensajesCheckList(1) Then Exit Sub
If Not MensajesCheckList(2) Then Exit Sub

If feDocumentos.TextMatrix(feDocumentos.row, 3) = "" Then
    MsgBox "Seleccione el [N° Conf.] del Documento Principal" & Chr(13) & feDocumentos.row & ".- [" & feDocumentos.TextMatrix(feDocumentos.row, 2) & "]", vbInformation, "Aviso"
    feDocumentos.row = feDocumentos.row
    feDocumentos.Col = 3
    Exit Sub
End If

If feDocumentos.TextMatrix(feDocumentos.row, 3) = "" Then
    Set objDCred = New COMDCredito.DCOMCatalogoProd
    Set rsDoc = objDCred.CargaFlexCheckList(4, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), , feDocumentos.TextMatrix(feDocumentos.row, 1))
    If Not (rsDoc.BOF And rsDoc.EOF) Then
        For i = 1 To rsDoc.RecordCount
            feDocumentos.TextMatrix(feDocumentos.row, 3) = rsDoc!nCantConf + 1
            rsDoc.MoveNext
        Next i
    End If
End If

DocPri = Trim(feDocumentos.TextMatrix(feDocumentos.row, 1))
For i = 1 To feNivelesCheckList.rows - 1
    Item = Trim(IIf(feNivelesCheckList.TextMatrix(i, 1) = "", 0, feNivelesCheckList.TextMatrix(i, 1)))
    If Item > 0 Then
        DocSubIndice = Mid(feNivelesCheckList.TextMatrix(i, 1), InStr(feNivelesCheckList.TextMatrix(i, 1), ".") + 1, 5)
    End If
Next i
    
Nivel = InputBox(DocPri & " - " & Trim(feDocumentos.TextMatrix(feDocumentos.row, 2)) & ", Conf N°." & Trim(feDocumentos.TextMatrix(feDocumentos.row, 3)) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "Ingrese el Documento del Nivel", "Documento Niveles")

If Nivel <> "" Then
    feNivelesCheckList.AdicionaFila
    DocSubIndice = DocSubIndice + 1
    feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) = DocPri & "." & DocSubIndice
    feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2) = UCase(Nivel)
    feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 3) = feDocumentos.TextMatrix(feDocumentos.row, 3)
End If

End Sub

Private Sub cmdAgregarParametroCheckList_Click()

feCondicionesCheckList.AdicionaFila
If Trim(Right(ListParametroCheckList.List(ListParametroCheckList.ListIndex), 9)) = "4000" Then
    feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 1) = ListParametroCheckList.List(ListParametroCheckList.ListIndex)
    feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 2) = Trim(Right(ListParametroCheckList.List(ListParametroCheckList.ListIndex), 9))
    Call habilitarControlesCheckList(4, True)
Else
    Call habilitarControlesCheckList(4, False)
    feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 1) = ListParametroCheckList.List(ListParametroCheckList.ListIndex)
    feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 2) = Trim(Right(ListParametroCheckList.List(ListParametroCheckList.ListIndex), 10))
End If
End Sub
Private Sub cmdQuitarParametroCheckList_Click()
If MsgBox("¿Esta seguro que desea eliminar la Condición seleccionada?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
    If Trim(Right(ListParametroCheckList.List(ListParametroCheckList.ListIndex), 9)) = "4000" Then
        Call habilitarControlesCheckList(4, False)
        feCondicionesCheckList.EliminaFila (feCondicionesCheckList.row)
    Else
        feCondicionesCheckList.EliminaFila (feCondicionesCheckList.row)
    End If
End Sub

Private Sub cmdQuitarDet_Click()
Dim i As Integer

If MsgBox("¿Esta seguro que desea eliminar el Detalle seleccionado?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub

    feDetalle.EliminaFila (feDetalle.row)
    
    For i = 1 To feDetalle.row
        If feDetalle.TextMatrix(i, 1) = "" Then
            Call habilitarControlesCheckList(2, True)
            Exit For
        End If
    Next i
End Sub

Private Sub cmdAgregarDoc_Click()
Dim Doc As String
Dim DocPri As Integer
Dim i As Integer
DocPri = 0

If Not MensajesCheckList(1) Then Exit Sub

For i = 1 To feDocumentos.rows - 1
    If Trim(feDocumentos.TextMatrix(i, 1)) <> "" Then
        DocPri = CInt(Trim(feDocumentos.TextMatrix(i, 1)))
    End If
Next i

Doc = InputBox("Ingrese el Documento Principal", "Documento Principal")

If Doc <> "" Then
    DocPri = IIf(DocPri = 0, 0, DocPri) + 1
    feDocumentos.AdicionaFila
    feDocumentos.TextMatrix(feDocumentos.row, 1) = DocPri
    feDocumentos.TextMatrix(feDocumentos.row, 2) = UCase(Doc)
    feDocumentos.TextMatrix(feDocumentos.row, 3) = 1
    Call habilitarControlesCheckList(2, True)
    
    For i = 1 To feDocumentos.rows - 1
        If feDocumentos.TextMatrix(i, 1) <> feDocumentos.row Then
            feDocumentos.TextMatrix(i, 3) = ""
        End If
    Next i
    
End If
LimpiaFlex feDetalle
LimpiaFlex feCondicionesCheckList
End Sub

Private Sub cmdQuitarDoc_Click()
Dim ItemDoc As String
Dim ConfDelete As String
Dim MaxConf As Integer
Dim CantConfTotal As Integer
Dim ItemCantDet As Integer
Dim ItemCantNiv As Integer
Dim i As Integer
Dim objDeleteDocDet As COMDCredito.DCOMCatalogoProd
Dim rsDeleteDocDet As ADODB.Recordset
Set objDeleteDocDet = New COMDCredito.DCOMCatalogoProd
ItemCantDet = 0
ItemCantNiv = 0

If MsgBox("¿Esta seguro que desea eliminar el Documento seleccionado?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub

ItemDoc = IIf(InStr(feDocumentos.TextMatrix(feDocumentos.row, 1), ".") = 0, Trim(feDocumentos.TextMatrix(feDocumentos.row, 1)) & ".0", Trim(feDocumentos.TextMatrix(feDocumentos.row, 1)))

Set objDeleteDocDet = New COMDCredito.DCOMCatalogoProd
Set rsDeleteDocDet = objDeleteDocDet.CargaFlexCheckList(6, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), IIf(ConfDelete = "", 0, ConfDelete), feDocumentos.TextMatrix(feDocumentos.row, 1))
    If (rsDeleteDocDet.BOF And rsDeleteDocDet.EOF) Then
        If Trim(feDocumentos.TextMatrix(feDocumentos.row, 3)) = "" And IIf(IsNull(rsDeleteDocDet!nCantConf), 0, rsDeleteDocDet!nCantConf) <> 0 Then
            MsgBox "Seleccione el N° Conf. a eliminar", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        If Trim(feDocumentos.TextMatrix(feDocumentos.row, 3)) = "" And IIf(IsNull(rsDeleteDocDet!nCantConf), 0, rsDeleteDocDet!nCantConf) <> 0 Then
            MsgBox "Seleccione el N° Conf. a eliminar", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

ConfDelete = Trim(feDocumentos.TextMatrix(feDocumentos.row, 3))

If Not ValidaCheckList(1, ItemDoc) Then Exit Sub
If Not ValidaCheckList(2, Mid(ItemDoc, Trim(InStr(ItemDoc, ".") - (InStr(ItemDoc, ".") - 1)), IIf(Len(ItemDoc) = 1, 1, Trim(InStr(ItemDoc, ".")) - 1))) Then Exit Sub
If Not ValidaCheckList(3, ItemDoc) Then Exit Sub

Set objDeleteDocDet = New COMDCredito.DCOMCatalogoProd
Set rsDeleteDocDet = objDeleteDocDet.CargaFlexCheckList(6, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), IIf(ConfDelete = "", 0, ConfDelete), feDocumentos.TextMatrix(feDocumentos.row, 1))
If Not (rsDeleteDocDet.BOF And rsDeleteDocDet.EOF) Then
    MaxConf = IIf(IsNull(rsDeleteDocDet!nCantConfMax), 0, rsDeleteDocDet!nCantConfMax)
    CantConfTotal = IIf(IsNull(rsDeleteDocDet!nCantConf), 0, rsDeleteDocDet!nCantConf)
Else
    MaxConf = 0
End If

Set rsDeleteDocDet = objDeleteDocDet.CargaFlexCheckList(2, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), IIf(ConfDelete = "", 0, ConfDelete), feDocumentos.TextMatrix(feDocumentos.row, 1))
If Not (rsDeleteDocDet.BOF And rsDeleteDocDet.EOF) Then

If feDocumentos.TextMatrix(feDocumentos.row, 3) < MaxConf Then
    MsgBox "No puede eliminar item menores", vbInformation, "Aviso"
    Exit Sub
End If
    
    If feDetalle.TextMatrix(1, 1) <> "" Then
        For i = 1 To feDetalle.rows - 1
            ItemCantDet = ItemCantDet + 1
            If CInt(feDetalle.TextMatrix(feDetalle.row, 1)) = ItemDoc Then
                feDetalle.EliminaFila (feDetalle.row)
            End If
        Next i
    End If
    
    If feNivelesCheckList.TextMatrix(1, 1) <> "" Then
        For i = 1 To feNivelesCheckList.rows - 1
            ItemCantNiv = ItemCantNiv + 1
            If CInt(feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1)) = ItemDoc Then
                feNivelesCheckList.EliminaFila (feNivelesCheckList.row)
            End If
        Next i
    End If
    
    If ItemCantNiv > 1 And ItemCantDet > 1 Then
        For i = 1 To feCondicionesCheckList.rows - 1
            feCondicionesCheckList.EliminaFila (feCondicionesCheckList.row)
        Next i
    End If
    
    Set objDeleteDocDet = New COMDCredito.DCOMCatalogoProd
    Call objDeleteDocDet.EliminaDatosCheckList(1, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), feDocumentos.TextMatrix(feDocumentos.row, 1), ConfDelete)
    Call objDeleteDocDet.EliminaDatosCheckList(2, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), feDocumentos.TextMatrix(feDocumentos.row, 1), ConfDelete)
    Call objDeleteDocDet.EliminaDatosCheckList(3, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), feDocumentos.TextMatrix(feDocumentos.row, 1), ConfDelete)
    
    If CantConfTotal = 1 Then
        feDocumentos.EliminaFila (feDocumentos.row)
    Else
        feDocumentos.TextMatrix(feDocumentos.row, 3) = ""
    End If
Else
    For i = 1 To feDetalle.rows - 1
        ItemCantDet = ItemCantDet + 1
        If feDetalle.TextMatrix(i, 1) = ItemDoc Then
            feDetalle.EliminaFila (i)
            Exit For
        End If
    Next i

    For i = 1 To feNivelesCheckList.rows - 1
        ItemCantNiv = ItemCantNiv + 1
        If feNivelesCheckList.TextMatrix(i, 1) = ItemDoc Then
            feNivelesCheckList.EliminaFila (i)
            Exit For
        End If
    Next i

    If ItemCantNiv = 1 And ItemCantDet = 1 Then
        For i = 1 To feCondicionesCheckList.rows - 1
            feCondicionesCheckList.EliminaFila (feCondicionesCheckList.row)
        Next i
    End If
    
    If MaxConf = 0 Then
        feDocumentos.EliminaFila (feDocumentos.row)
    Else
        feDocumentos.TextMatrix(feDocumentos.row, 3) = ""
    End If
End If
   
If feDocumentos.TextMatrix(feDocumentos.row, 1) = "" Then
    Call LimpiaFlex(feDetalle)
    Call LimpiaFlex(feCondicionesCheckList)
    Call habilitarControlesCheckList(1, False)
End If

Set objDeleteDocDet = Nothing
RSClose rsDeleteDocDet
End Sub

Private Sub feCondicionesCheckList_DblClick()
    'para editar
End Sub

Private Sub feCondicionesCheckList_OnChangeCombo()
    If Trim(Right(feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 2), 10)) <> "4000" Then
        feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 4) = Trim(Right(feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 3), 10))
    End If
End Sub

Private Sub feCondicionesCheckList_RowColChange()
Dim psTpoDescripCheckList As String
Dim RsCbo As New ADODB.Recordset
'JOEP
    Dim oCredCat As COMDCredito.DCOMCatalogoProd
    Set oCredCat = New COMDCredito.DCOMCatalogoProd
'JOEP

If Trim(Right((feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 2)), 10)) = "4000" Then
    Call habilitarControlesCheckList(4, True)
    feCondicionesCheckList.ListaControles = "0-0-0-0-0-0"
    feCondicionesCheckList.ColumnasAEditar = "X-X-X-X-X-X"
Else
    feCondicionesCheckList.ColumnasAEditar = "X-X-X-C-X-X"
    feCondicionesCheckList.ListaControles = "0-0-0-3-0-0"
    Call habilitarControlesCheckList(4, False)
    
    psTpoDescripCheckList = Trim((feCondicionesCheckList.TextMatrix(feCondicionesCheckList.row, 2)))
    
    If psTpoDescripCheckList <> "" Then
      Set RsCbo = oCredCat.CargarCboCondiciones(CDbl(psTpoDescripCheckList))
      feCondicionesCheckList.CargaCombo RsCbo
       feCondicionesCheckList.TamañoCombo (300)
    End If
End If
'JOEP
Set oCredCat = Nothing
RSClose RsCbo
'JOEP
End Sub

Private Sub habilitarControlesCheckList(ByVal nTp As Integer, ByVal TrueFalse As Boolean)
Select Case nTp
        Case 1
                'Documento
                cmdQuitarDoc.Enabled = TrueFalse
                cmdAgregarConfDoc.Enabled = TrueFalse
                'Niveles
                feNivelesCheckList.Enabled = TrueFalse
                cmdAgregarNivCheckList.Enabled = TrueFalse
                cmdQuitarDocNivelesCheckList.Enabled = TrueFalse
                'Detalle
                feDetalle.Enabled = TrueFalse
                cmdAgregarDet.Enabled = TrueFalse
                cmdQuitarDet.Enabled = TrueFalse
                'Condiciones
                feCondicionesCheckList.Enabled = TrueFalse
                cmdAgregarParametroCheckList.Enabled = TrueFalse
                cmdQuitarParametroCheckList.Enabled = TrueFalse
                ListParametroCheckList.Enabled = TrueFalse
                'Montos
                cmbIniCheckList.Enabled = TrueFalse
                txtInicioCheckList.Enabled = TrueFalse
                cmbFinCheckList.Enabled = TrueFalse
                txtFinCheckList.Enabled = TrueFalse
                cmbUniMedCheckList.Enabled = TrueFalse
                cmdAgregarMontoCheckList.Enabled = TrueFalse
                                               
                cmdGrabarcheckList.Enabled = TrueFalse
        Case 2
                'Documento
                    cmdQuitarDoc.Enabled = TrueFalse
                    cmdAgregarConfDoc.Enabled = TrueFalse
                'Niveles
                    feNivelesCheckList.Enabled = TrueFalse
                    cmdAgregarNivCheckList.Enabled = TrueFalse
                    cmdQuitarDocNivelesCheckList.Enabled = TrueFalse
                'Detalle
                    feDetalle.Enabled = TrueFalse
                    cmdAgregarDet.Enabled = TrueFalse
                    cmdQuitarDet.Enabled = TrueFalse
                'Grabar
                    cmdGrabarcheckList.Enabled = TrueFalse
        Case 3
            'Condiciones
                feCondicionesCheckList.Enabled = TrueFalse
                cmdAgregarParametroCheckList.Enabled = TrueFalse
                cmdQuitarParametroCheckList.Enabled = TrueFalse
                ListParametroCheckList.Enabled = TrueFalse
        Case 4
            'Montos
                fraMontosCheckList.Enabled = TrueFalse
                cmbIniCheckList.Enabled = TrueFalse
                txtInicioCheckList.Enabled = TrueFalse
                cmbFinCheckList.Enabled = TrueFalse
                txtFinCheckList.Enabled = TrueFalse
                cmbUniMedCheckList.Enabled = TrueFalse
                cmdAgregarMontoCheckList.Enabled = TrueFalse
        Case 5
                'Condiciones
                feCondicionesCheckList.Enabled = TrueFalse
                cmdAgregarParametroCheckList.Enabled = TrueFalse
                cmdQuitarParametroCheckList.Enabled = TrueFalse
                ListParametroCheckList.Enabled = TrueFalse
                'monto
                fraMontosCheckList.Enabled = TrueFalse
                cmbIniCheckList.Enabled = TrueFalse
                txtInicioCheckList.Enabled = TrueFalse
                cmbFinCheckList.Enabled = TrueFalse
                txtFinCheckList.Enabled = TrueFalse
                cmbUniMedCheckList.Enabled = TrueFalse
                cmdAgregarMontoCheckList.Enabled = TrueFalse
        Case 6
                'Detalle
                'cmdQuitarDet.Enabled = TrueFalse
        Case 7
                'Documento
                    cmdQuitarDoc.Enabled = TrueFalse
                    cmdAgregarConfDoc.Enabled = TrueFalse
                'Niveles
                    feNivelesCheckList.Enabled = TrueFalse
                    cmdAgregarNivCheckList.Enabled = TrueFalse
                    cmdQuitarDocNivelesCheckList.Enabled = TrueFalse
                'Detalle
                    feDetalle.Enabled = TrueFalse
                    cmdAgregarDet.Enabled = TrueFalse
                    cmdQuitarDet.Enabled = TrueFalse
                'condicion
                    feCondicionesCheckList.Enabled = TrueFalse
                    cmdAgregarParametroCheckList.Enabled = TrueFalse
                    cmdQuitarParametroCheckList.Enabled = TrueFalse
                    ListParametroCheckList.Enabled = TrueFalse
                'Grabar
                    cmdGrabarcheckList.Enabled = TrueFalse
End Select
End Sub

Private Function MensajesCheckList(ByVal nTp As Integer) As Boolean
MensajesCheckList = True
Select Case nTp
    Case 1
        If (cmbCatCheckLis.Text = "") Then
            MsgBox "Seleccione la Categoria", vbInformation, "Aviso"
            cmbCatCheckLis.SetFocus
            MensajesCheckList = False
            Exit Function
        ElseIf (cmbProdCheckLis.Text = "") Then
            MsgBox "Seleccione el Producto", vbInformation, "Aviso"
            cmbProdCheckLis.SetFocus
            MensajesCheckList = False
            Exit Function
        End If
    Case 2
        If Trim(Left(feDocumentos.TextMatrix(feDocumentos.row, 0), 1)) = "" Then
            MsgBox "Ingrese el Documento Principal", vbInformation, "Aviso"
            cmdAgregarDoc.SetFocus
            MensajesCheckList = False
            Exit Function
        End If
     Case 3
        If cmbIniCheckList.Text = "" Or txtInicioCheckList.Text = "" Or cmbFinCheckList.Text = "" Or txtFinCheckList.Text = "" Or cmbUniMedCheckList.Text = "" Then
            MsgBox "Ingrese los datos del parametro MONTO", vbInformation, "Aviso"
            cmbIniCheckList.SetFocus
            MensajesCheckList = False
            Exit Function
        End If
End Select
End Function

Private Function ValidaCheckList(ByVal pnTp As Integer, ByVal pnDato As String) As Boolean
Dim Item As String
Dim iTemMax As String
Dim iTemCab As Integer
Dim i As Integer
Dim j As Integer
Dim bExist As Integer
Dim dato As String
bExist = 0
    ValidaCheckList = True
Select Case pnTp
    Case 1
        For i = 1 To feDocumentos.rows - 1
        Item = IIf(InStr(Trim(feDocumentos.TextMatrix(i, 1)), ".") = 0, Trim(feDocumentos.TextMatrix(i, 1)) & ".0", Trim(feDocumentos.TextMatrix(i, 1)))
            If i = 1 And Len(pnDato) > 1 Then
                ValidaCheckList = True
                Exit For
            End If
            If Len(pnDato) = 1 Then
                If Len(Item) > 1 And Mid(Item, Trim(InStr(Item, ".") - (InStr(Item, ".") - 1)), IIf(Len(Item) = 1, 1, Trim(InStr(Item, ".")) - 1)) = pnDato And Trim(InStr(feDocumentos.TextMatrix(i, 1), ".")) > 0 Then
                    MsgBox "No puede eliminar el item " & pnDato & ", por que contiene sub items." & Chr(13) & " Si no hubiera sub item si se procede a eliminar.", vbInformation, "Aviso"
                    ValidaCheckList = False
                    Exit For
                End If
            End If
        Next i
    Case 2
        For i = 1 To feDocumentos.rows - 1
            Item = IIf(InStr(Trim(feDocumentos.TextMatrix(i, 1)), ".") = 0, Trim(feDocumentos.TextMatrix(i, 1)) & ".0", Trim(feDocumentos.TextMatrix(i, 1)))
            If Mid(Item, Trim(InStr(Item, ".") - (InStr(Item, ".") - 1)), IIf(Len(Item) = 1, 1, Trim(InStr(Item, ".")) - 1)) = pnDato Then
                iTemMax = Replace(Trim(feDocumentos.TextMatrix(i, 1)), ".", "")
            End If
        Next i
        
        For i = 1 To feDocumentos.rows - 1
            Item = IIf(InStr(Trim(feDocumentos.TextMatrix(i, 1)), ".") = 0, Trim(feDocumentos.TextMatrix(i, 1)) & ".0", Trim(feDocumentos.TextMatrix(i, 1)))
            If Len(pnDato) = 1 Then
                If Len(Item) > 1 And Mid(Item, Trim(InStr(Item, ".") - (InStr(Item, ".") - 1)), IIf(Len(Item) = 1, 1, Trim(InStr(Item, ".")) - 1)) = pnDato And CInt(Replace(Trim(feDocumentos.TextMatrix(feDocumentos.row, 1)), ".", "")) < CInt(iTemMax) Then
                    MsgBox "No puede eliminar Item menores.", vbInformation, "Aviso"
                    ValidaCheckList = False
                    Exit For
                End If
            End If
        Next i
    Case 3
        For i = 1 To feDocumentos.rows - 1
            If InStr(Trim(feDocumentos.TextMatrix(i, 1)), ".") = 0 Then
                iTemMax = Trim(feDocumentos.TextMatrix(i, 1))
                iTemCab = Trim(feDocumentos.TextMatrix(i, 1))
            Else
                iTemMax = Mid(feDocumentos.TextMatrix(i, 1), Trim(InStr(feDocumentos.TextMatrix(i, 1), ".") - (InStr(feDocumentos.TextMatrix(i, 1), ".") - 1)), IIf(Len(feDocumentos.TextMatrix(i, 1)) = 1, 1, Trim(InStr(feDocumentos.TextMatrix(i, 1), ".")) - 1))
                iTemCab = Mid(feDocumentos.TextMatrix(i, 1), Trim(InStr(feDocumentos.TextMatrix(i, 1), ".") - (InStr(feDocumentos.TextMatrix(i, 1), ".") - 1)), IIf(Len(feDocumentos.TextMatrix(i, 1)) = 1, 1, Trim(InStr(feDocumentos.TextMatrix(i, 1), ".")) - 1))
            End If
        Next i
        
        For i = 1 To feDocumentos.rows - 1
            If Len(Trim(feDocumentos.TextMatrix(feDocumentos.row, 1))) = 1 And CInt(Trim(feDocumentos.TextMatrix(feDocumentos.row, 1))) <= iTemCab And CInt(Replace(Trim(feDocumentos.TextMatrix(feDocumentos.row, 1)), ".", "")) < CInt(iTemMax) Then
                MsgBox "No puede eliminar Item menores del [Documento Principal].", vbInformation, "Aviso"
                ValidaCheckList = False
                Exit For
            End If
        Next i
    Case 4
        Dim valchek As Integer
        valchek = 0
        If cmbCatCheckLis.Text = "" Then
            MsgBox "Seleccione la Categoria", vbInformation, "Aviso"
            ValidaCheckList = False
            cmbCatCheckLis.SetFocus
            Exit Function
        ElseIf cmbProdCheckLis.Text = "" Then
            MsgBox "Seleccione el Producto", vbInformation, "Aviso"
            ValidaCheckList = False
            cmbProdCheckLis.SetFocus
            Exit Function
        ElseIf feDocumentos.TextMatrix(1, 1) = "" Then
            MsgBox "Ingrese los documentos", vbInformation, "Aviso"
            ValidaCheckList = False
            feDocumentos.SetFocus
            Exit Function
        ElseIf feDetalle.TextMatrix(1, 1) = "" Then
            MsgBox "Ingrese los detalles", vbInformation, "Aviso"
            ValidaCheckList = False
            feDetalle.SetFocus
            Exit Function
        ElseIf feCondicionesCheckList.TextMatrix(1, 1) = "" Then
            MsgBox "Ingrese las condiciones", vbInformation, "Aviso"
            ValidaCheckList = False
            feCondicionesCheckList.SetFocus
            Exit Function
        ElseIf feDocumentos.TextMatrix(feDocumentos.row, 3) = "" Then
            MsgBox "Ingrese N° Conf. a Guardar", vbInformation, "Aviso"
            ValidaCheckList = False
            feDocumentos.SetFocus
            Exit Function
        End If
        
        If feNivelesCheckList.TextMatrix(1, 4) <> "" Then
            For i = 1 To feNivelesCheckList.rows - 1
                If feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 4) = "" Then
                    MsgBox "Falta ingresar el Nivel del Documento " & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) & " - [" & feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2) & "]", vbInformation, "Aviso"
                    ValidaCheckList = False
                    feCondicionesCheckList.SetFocus
                    Exit Function
                End If
            Next i
            For i = 1 To feNivelesCheckList.rows - 1
                For j = 1 To feDetalle.rows - 1
                    If feNivelesCheckList.TextMatrix(i, 1) = feDetalle.TextMatrix(j, 1) Then
                        bExist = 1
                        Exit For
                    End If
                Next j
                If bExist = 0 Then
                    MsgBox "Falta ingresar el Detalle del Documento " & feNivelesCheckList.TextMatrix(j, 1) & " - [" & feNivelesCheckList.TextMatrix(j, 2) & "]", vbInformation, "Aviso"
                    ValidaCheckList = False
                    feCondicionesCheckList.SetFocus
                    Exit Function
                Else
                    bExist = 0
                End If
            Next i
        End If
        
        For i = 1 To feDetalle.rows - 1
            If feDetalle.TextMatrix(i, 3) = "." Then
                valchek = valchek + 1
            End If
        Next i
            
        If valchek = 0 Then
            If MsgBox("No esta seleccionando ningun Item del Detalle, ¿Desea continuar?", vbYesNo + vbQuestion, "Atención") = vbNo Then
                ValidaCheckList = False
                feDetalle.SetFocus
                Exit Function
            Else
                ValidaCheckList = True
            End If
        End If
    Case 5
        For i = 1 To feNivelesCheckList.rows - 1
            dato = IIf(feNivelesCheckList.TextMatrix(i, 1) = "", 0, feNivelesCheckList.TextMatrix(i, 1))
            iTemMax = Replace(Trim(dato), ".", "")
            iTemCab = Mid(dato, Trim(InStr(dato, ".") - (InStr(dato, ".") - 1)), IIf(Len(dato) = 1, 1, Trim(InStr(dato, ".")) - 1))
        Next i
        
        For i = 1 To feNivelesCheckList.rows - 1
            dato = IIf(feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1) = "", 0, feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 1))
            If Mid(dato, InStr(dato, ".") - (InStr(dato, ".") - 1), InStr(dato, ".") - 1) <= iTemCab And CInt(Replace(Trim(dato), ".", "")) < CInt(iTemMax) Then
                MsgBox "No puede eliminar Item menores del [Documento Niveles].", vbInformation, "Aviso"
                ValidaCheckList = False
                Exit For
            End If
        Next i
End Select

End Function

Private Sub feDocumentos_OnCellChange(pnRow As Long, pnCol As Long)
    If feDocumentos.Col = 4 Then
        feDocumentos.AvanceCeldas = Vertical
    Else
        feDocumentos.AvanceCeldas = Horizontal
    End If
    
    If feDocumentos.Col = 3 Or feDocumentos.Col = 1 Then
        If feDocumentos.row <> 1 Then
            feDocumentos.row = IIf(feDocumentos.Col = 1, feDocumentos.row - 1, feDocumentos.row)
        End If
        feDocumentos.Col = 3
    End If
    
End Sub

Private Sub feDocumentos_Click()
Dim Item As String
Item = Len(feDocumentos.TextMatrix(feDocumentos.row, 1))

If feDocumentos.TextMatrix(feDocumentos.row, 0) = "" Then Exit Sub

Select Case feDocumentos.Col
    Case 4
        If Item > 1 Then
            feDocumentos.ColumnasAEditar = "X-X-X-X-3"
        Else
             feDocumentos.ColumnasAEditar = "X-X-X-X-X"
        End If
    Case Else
        feDocumentos.ColumnasAEditar = "X-X-X-X-X"
End Select

End Sub

Private Sub feDocumentos_RowColChange()
Dim j As Integer
Dim i As Integer
Select Case feDocumentos.Col
    Case 3
        j = feDocumentos.rows - 1
        For i = 1 To feDocumentos.rows - 1
            If InStr(feDocumentos.TextMatrix(j, 1), ".") > 0 Then
                feDocumentos.EliminaFila (j)
                j = j - 1
            End If
        Next i
End Select
End Sub

Private Sub feDocumentos_DblClick()
Dim Item As String
Dim objDCred As COMDCredito.DCOMCatalogoProd
Dim rsDoc As ADODB.Recordset
Dim i As Integer
Dim j As Integer

If feDocumentos.TextMatrix(feDocumentos.row, 3) = "" Then
    feDocumentos.TextMatrix(feDocumentos.row, 3) = ""
End If
    
    If InStr(feDocumentos.TextMatrix(feDocumentos.row, 1), ".") = 0 Then
        Item = feDocumentos.TextMatrix(feDocumentos.row, 1)
    Else
        Item = Len(feDocumentos.TextMatrix(feDocumentos.row, 1))
    End If
        
    If feDocumentos.TextMatrix(feDocumentos.row, 0) = "" Then Exit Sub
    
    Select Case feDocumentos.Col
        Case 2
            Dim cItemDet As String
            If feDocumentos.Col = 2 Then
                If feDocumentos.TextMatrix(feDocumentos.row, 2) <> "" And feDocumentos.TextMatrix(feDocumentos.row, 3) <> "" Then
                    cItemDet = InputBox("", "Modificacion de Documento", feDocumentos.TextMatrix(feDocumentos.row, 2))
                    If cItemDet <> "" Then
                        feDocumentos.TextMatrix(feDocumentos.row, 2) = UCase(cItemDet)
                    End If
                    SendKeys "{Enter}"
                End If
            End If
        Case 3
                Set objDCred = New COMDCredito.DCOMCatalogoProd
                    LimpiaFlex feNivelesCheckList
                    LimpiaFlex feDetalle
                    LimpiaFlex feCondicionesCheckList
                    Set rsDoc = objDCred.CargaFlexCheckList(4, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), , feDocumentos.TextMatrix(feDocumentos.row, 1))
                    If Not (rsDoc.BOF And rsDoc.EOF) Then
                        feDocumentos.CargaCombo rsDoc
                        feDocumentos.ColumnasAEditar = "X-X-X-2-X"
                    Else
                        feDocumentos.ColumnasAEditar = "X-X-X-X-X"
                    End If
            For i = 1 To feDocumentos.rows - 1
                If i <> feDocumentos.row Then
                    feDocumentos.TextMatrix(i, 3) = ""
                End If
            Next i
            
        Case 4
            If Item > 1 Then
                feDocumentos.ColumnasAEditar = "X-X-X-X-3"
            End If
        Case Else
            feDocumentos.ColumnasAEditar = "X-X-X-X-X"
    End Select
End Sub

Private Sub feDocumentos_OnChangeCombo()
Dim objComb As COMDCredito.DCOMCatalogoProd
Dim rsCombCheckList As ADODB.Recordset
Dim i As Integer
Dim cItemPri As String
Dim nCantConf As Integer
Set objComb = New COMDCredito.DCOMCatalogoProd

If feDocumentos.TextMatrix(feDocumentos.row, 3) <> "" Then

    If feDocumentos.Col = 3 Then
        cItemPri = feDocumentos.TextMatrix(feDocumentos.row, 1)
        nCantConf = IIf(feDocumentos.TextMatrix(feDocumentos.row, 3) = "", 0, feDocumentos.TextMatrix(feDocumentos.row, 3))
    
        Set objComb = New COMDCredito.DCOMCatalogoProd
        Set rsCombCheckList = objComb.CargaFlexCheckList(5, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), nCantConf, cItemPri)
        If Not (rsCombCheckList.BOF And rsCombCheckList.EOF) Then
            LimpiaFlex feNivelesCheckList
            For i = 1 To rsCombCheckList.RecordCount
                feNivelesCheckList.AdicionaFila
                feNivelesCheckList.TextMatrix(i, 1) = rsCombCheckList!cItem
                feNivelesCheckList.TextMatrix(i, 2) = rsCombCheckList!cDescripcion
                feNivelesCheckList.TextMatrix(i, 3) = rsCombCheckList!nCantConf
                feNivelesCheckList.TextMatrix(i, 4) = rsCombCheckList!nNivel
                feNivelesCheckList.TextMatrix(i, 5) = rsCombCheckList!cConsDescripcion
            rsCombCheckList.MoveNext
            Next i
        Else
            LimpiaFlex feNivelesCheckList
        End If
     
        Set objComb = New COMDCredito.DCOMCatalogoProd
        Set rsCombCheckList = objComb.CargaFlexCheckList(2, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), nCantConf, cItemPri)
        If Not (rsCombCheckList.BOF And rsCombCheckList.EOF) Then
            LimpiaFlex feDetalle
            For i = 1 To rsCombCheckList.RecordCount
                feDetalle.AdicionaFila
                feDetalle.TextMatrix(i, 1) = rsCombCheckList!cItem
                feDetalle.TextMatrix(i, 2) = rsCombCheckList!cDescripcion
                feDetalle.TextMatrix(i, 3) = IIf(rsCombCheckList!nEstadoCheck, 1, 0)
                rsCombCheckList.MoveNext
            Next i
        Else
            LimpiaFlex feDetalle
        End If
    
        Set objComb = New COMDCredito.DCOMCatalogoProd
        Set rsCombCheckList = objComb.CargaFlexCheckList(3, Trim(Right(cmbCatCheckLis.Text, 10)), Trim(Right(cmbProdCheckLis.Text, 10)), nCantConf, cItemPri)
        If Not (rsCombCheckList.BOF And rsCombCheckList.EOF) Then
            LimpiaFlex feCondicionesCheckList
            For i = 1 To rsCombCheckList.RecordCount
                feCondicionesCheckList.AdicionaFila
                feCondicionesCheckList.TextMatrix(i, 1) = IIf(IsNull(rsCombCheckList!cParDesc), "", rsCombCheckList!cParDesc)
                feCondicionesCheckList.TextMatrix(i, 2) = IIf(IsNull(rsCombCheckList!nParCod), "", rsCombCheckList!nParCod)
                feCondicionesCheckList.TextMatrix(i, 3) = IIf(IsNull(rsCombCheckList!cParDescripcion), "", rsCombCheckList!cParDescripcion)
                feCondicionesCheckList.TextMatrix(i, 4) = IIf(IsNull(rsCombCheckList!nParValor), "", rsCombCheckList!nParValor)
                feCondicionesCheckList.TextMatrix(i, 5) = IIf(IsNull(rsCombCheckList!cOperadorIni), "", rsCombCheckList!cOperadorIni)
                feCondicionesCheckList.TextMatrix(i, 6) = IIf(IsNull(rsCombCheckList!nMontoIni), "", rsCombCheckList!nMontoIni)
                feCondicionesCheckList.TextMatrix(i, 7) = IIf(IsNull(rsCombCheckList!cOperadorFin), "", rsCombCheckList!cOperadorFin)
                feCondicionesCheckList.TextMatrix(i, 8) = IIf(IsNull(rsCombCheckList!nMontoFin), "", rsCombCheckList!nMontoFin)
                feCondicionesCheckList.TextMatrix(i, 9) = IIf(IsNull(rsCombCheckList!nUnidadMedida), "", rsCombCheckList!nUnidadMedida)
                rsCombCheckList.MoveNext
            Next i
        Else
            LimpiaFlex feCondicionesCheckList
        End If
    
        Call habilitarControlesCheckList(7, True)
    End If
End If
End Sub

Private Sub feNivelesCheckList_OnCellChange(pnRow As Long, pnCol As Long)
    If feNivelesCheckList.Col = 5 Or feNivelesCheckList.Col = 1 Then
        If feNivelesCheckList.row <> 1 Then
            feNivelesCheckList.row = IIf(feNivelesCheckList.Col = 1, feNivelesCheckList.row - 1, feNivelesCheckList.row)
        End If
        feNivelesCheckList.Col = 5
    End If
End Sub

Private Sub feNivelesCheckList_EnterCell()
    If feNivelesCheckList.Col = 4 Then
        feNivelesCheckList.AvanceCeldas = Vertical
    ElseIf feNivelesCheckList.Col = 5 Or feNivelesCheckList.Col = 1 Then
        If feNivelesCheckList.row <> 1 Then
            feNivelesCheckList.row = IIf(feNivelesCheckList.Col = 1, feNivelesCheckList.row - 1, feNivelesCheckList.row)
        End If
        feNivelesCheckList.Col = 5
    Else
        feNivelesCheckList.AvanceCeldas = Horizontal
    End If
End Sub

Private Sub feNivelesCheckList_DblClick()
Dim oCon As New COMDConstantes.DCOMConstantes
Dim cItemDet As String
    If feNivelesCheckList.Col = 5 Then
        feNivelesCheckList.CargaCombo oCon.RecuperaConstantes(90000)
        feNivelesCheckList.TamañoCombo (100)
    End If
    
    If feNivelesCheckList.Col = 2 Then
        If feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2) <> "" And (feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 4) <> "" And feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 5) <> "") Then
            cItemDet = InputBox("", "Modificacion de Niveles", feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2))
            If cItemDet <> "" Then
                feNivelesCheckList.TextMatrix(feNivelesCheckList.row, 2) = cItemDet
            End If
            SendKeys "{Enter}"
        End If
    End If
    
End Sub

Private Sub feNivelesCheckList_Click()
Dim oCon As New COMDConstantes.DCOMConstantes
    If feNivelesCheckList.Col = 5 Then
        feNivelesCheckList.CargaCombo oCon.RecuperaConstantes(90000)
        feNivelesCheckList.TamañoCombo (100)
    End If
End Sub

Private Sub feDetalle_DblClick()
Dim cItemDet As String
    If feDetalle.Col = 2 Then
        If feDetalle.TextMatrix(feDetalle.row, 2) <> "" Then
            cItemDet = InputBox("", "Modificacion de Detalle", feDetalle.TextMatrix(feDetalle.row, 2))
            If cItemDet <> "" Then
                feDetalle.TextMatrix(feDetalle.row, 2) = UCase(cItemDet)
            End If
            SendKeys "{Enter}"
        End If
    End If
End Sub
'JOEP20181204 ERS034-CP




