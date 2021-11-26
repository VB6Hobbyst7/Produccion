VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVacacionesGozadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vaciones Gozadas"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   Icon            =   "frmVacionesGozadas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8640
      TabIndex        =   18
      Top             =   6960
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6780
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   11959
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   706
      TabCaption(0)   =   "Consulta Vac"
      TabPicture(0)   =   "frmVacionesGozadas.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ctrRRHHGen1"
      Tab(0).Control(1)=   "FlexEdit1"
      Tab(0).Control(2)=   "CmdBuscar"
      Tab(0).Control(3)=   "txtAño"
      Tab(0).Control(4)=   "cboMes"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Registrar Vac"
      TabPicture(1)   =   "frmVacionesGozadas.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEstado"
      Tab(1).Control(1)=   "Flex"
      Tab(1).Control(2)=   "CmdAdicionar"
      Tab(1).Control(3)=   "CmdGrabar"
      Tab(1).Control(4)=   "CmdCancelar"
      Tab(1).Control(5)=   "CmdEliminar"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Abonar Vac. Mes"
      TabPicture(2)   =   "frmVacionesGozadas.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblProceso"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lnlTot"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "flex2"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdAbonar"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "CmdExcel3"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "mskFecha"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdAsiento"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtTotal"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cmdProvision"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtSaldoBalance"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtTotalSueldo"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtProvision"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "cmdConsol"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Provision Vac"
      TabPicture(3)   =   "frmVacionesGozadas.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LblProvision"
      Tab(3).Control(1)=   "FlexProv"
      Tab(3).Control(2)=   "CmdGenerarProv"
      Tab(3).Control(3)=   "CmdOcultar"
      Tab(3).Control(4)=   "CmdExpProvVac"
      Tab(3).Control(5)=   "cmdBasePDTprovicion"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Vacaciones Ejecutadas"
      TabPicture(4)   =   "frmVacionesGozadas.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label3"
      Tab(4).Control(1)=   "FlexVac"
      Tab(4).Control(2)=   "txtFechaVac"
      Tab(4).Control(3)=   "cmdExportaExcel5"
      Tab(4).Control(4)=   "cmdBasePDTvacacionE"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Base PDT"
      TabPicture(5)   =   "frmVacionesGozadas.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FlexBasePDT"
      Tab(5).Control(1)=   "ctrRRHHGen2"
      Tab(5).Control(2)=   "CmdEliminarPDT"
      Tab(5).Control(3)=   "CmdCancelarPDT"
      Tab(5).Control(4)=   "CmdGrabarPDT"
      Tab(5).Control(5)=   "CmdAdicionarPDT"
      Tab(5).Control(6)=   "cmdSubsidio"
      Tab(5).ControlCount=   7
      Begin VB.CommandButton cmdConsol 
         Caption         =   "&Consol Excel"
         Height          =   345
         Left            =   1560
         TabIndex        =   56
         Top             =   5880
         Width           =   1260
      End
      Begin VB.TextBox txtProvision 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8280
         TabIndex        =   54
         Text            =   "0.00"
         Top             =   5400
         Width           =   1100
      End
      Begin VB.TextBox txtTotalSueldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         TabIndex        =   52
         Text            =   "0.00"
         Top             =   5400
         Width           =   1100
      End
      Begin VB.TextBox txtSaldoBalance 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdProvision 
         Caption         =   "&P2: Provisionar"
         Height          =   345
         Left            =   6600
         TabIndex        =   49
         Top             =   5880
         Width           =   1300
      End
      Begin VB.CommandButton cmdSubsidio 
         Caption         =   "&Subsidio"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -70260
         TabIndex        =   4
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton cmdBasePDTvacacionE 
         Caption         =   "<< Base PDT >> "
         Height          =   345
         Left            =   -73185
         TabIndex        =   47
         Top             =   5400
         Width           =   1620
      End
      Begin VB.CommandButton cmdBasePDTprovicion 
         Caption         =   "<< Base PDT >> "
         Height          =   345
         Left            =   -70230
         TabIndex        =   46
         Top             =   5400
         Width           =   1620
      End
      Begin VB.CommandButton CmdAdicionarPDT 
         Caption         =   "Adicionar"
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
         Left            =   -71805
         TabIndex        =   44
         Top             =   960
         Width           =   1200
      End
      Begin VB.CommandButton CmdGrabarPDT 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   -67950
         TabIndex        =   43
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelarPDT 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -66795
         TabIndex        =   42
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminarPDT 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   -69120
         TabIndex        =   41
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton cmdExportaExcel5 
         Caption         =   "<< Exportar Excel >> "
         Height          =   345
         Left            =   -74880
         TabIndex        =   39
         Top             =   5400
         Width           =   1620
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   5400
         Width           =   1100
      End
      Begin VB.CommandButton cmdAsiento 
         Caption         =   "&P3: Asiento"
         Height          =   345
         Left            =   8040
         TabIndex        =   34
         Top             =   5865
         Width           =   1300
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   1050
         TabIndex        =   32
         Top             =   675
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   -69120
         TabIndex        =   31
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton CmdExpProvVac 
         Caption         =   "<< Exportar Excel >> "
         Height          =   345
         Left            =   -68520
         TabIndex        =   30
         Top             =   5400
         Width           =   1620
      End
      Begin VB.CommandButton CmdOcultar 
         Caption         =   "<< Ocultar >>"
         Height          =   315
         Left            =   -74880
         TabIndex        =   29
         Top             =   5415
         Width           =   1170
      End
      Begin VB.CommandButton CmdGenerarProv 
         Caption         =   "<< Generar >>"
         Height          =   345
         Left            =   -66810
         TabIndex        =   27
         Top             =   5400
         Width           =   1170
      End
      Begin VB.CommandButton CmdExcel3 
         Caption         =   "&Exportar Excel"
         Height          =   345
         Left            =   120
         TabIndex        =   23
         Top             =   5880
         Width           =   1260
      End
      Begin VB.Frame Frame2 
         Caption         =   "Busqueda"
         Height          =   1200
         Left            =   -74865
         TabIndex        =   20
         Top             =   75
         Width           =   1365
         Begin VB.OptionButton OptBusca 
            Caption         =   "Empleado"
            Height          =   315
            Index           =   1
            Left            =   135
            TabIndex        =   22
            Top             =   720
            Value           =   -1  'True
            Width           =   1050
         End
         Begin VB.OptionButton OptBusca 
            Caption         =   "Mes Año"
            Height          =   345
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -66795
         TabIndex        =   15
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   -67950
         TabIndex        =   14
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbonar 
         Caption         =   "&P1: Abonar Vac"
         Height          =   345
         Left            =   5160
         TabIndex        =   13
         Top             =   5880
         Width           =   1300
      End
      Begin VB.Frame Frame1 
         Height          =   585
         Left            =   5265
         TabIndex        =   9
         Top             =   450
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton OptEst 
            Caption         =   "Contratados"
            Height          =   210
            Index           =   2
            Left            =   2655
            TabIndex        =   12
            Top             =   240
            Width           =   1155
         End
         Begin VB.OptionButton OptEst 
            Caption         =   "Estables"
            Height          =   210
            Index           =   1
            Left            =   1335
            TabIndex        =   11
            Top             =   240
            Width           =   1005
         End
         Begin VB.OptionButton OptEst 
            Caption         =   "Todos"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.CommandButton CmdAdicionar 
         Caption         =   "Adicionar"
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
         Left            =   -71805
         TabIndex        =   6
         Top             =   960
         Width           =   1200
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   -70155
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   630
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtAño 
         Height          =   300
         Left            =   -68115
         MaxLength       =   4
         TabIndex        =   2
         Top             =   630
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CommandButton CmdBuscar 
         Height          =   600
         Left            =   -66405
         Picture         =   "frmVacionesGozadas.frx":0972
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   405
         Width           =   675
      End
      Begin Sicmact.FlexEdit Flex 
         Height          =   3915
         Left            =   -74835
         TabIndex        =   5
         Top             =   1350
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   6906
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "CodPersona-Codigo-Nombre-Fec Ini-Fec Fin-Dias-Comentario-CodPersona"
         EncabezadosAnchos=   "0-800-4000-1000-1000-800-4000-0"
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
         ColumnasAEditar =   "X-X-X-3-4-X-6-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-2-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-3-0-0"
         TextArray0      =   "CodPersona"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit FlexEdit1 
         Height          =   4380
         Left            =   -74850
         TabIndex        =   7
         Top             =   1380
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   7726
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Fec Ini Vac-Fec Ing Lab-Comentario"
         EncabezadosAnchos=   "300-800-4000-1200-1200-4000"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-L-L"
         FormatosEdit    =   "0-0-0-5-5-1"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.FlexEdit flex2 
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7355
         Cols0           =   15
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmVacionesGozadas.frx":0C7C
         EncabezadosAnchos=   "500-700-3000-900-0-600-900-1000-0-0-0-0-0-0-900"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C-R-R-R-L-C-L-L-L-L-R"
         FormatosEdit    =   "0-0-0-5-5-3-2-3-0-0-0-0-0-0-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.ctrRRHHGen ctrRRHHGen1 
         Height          =   1200
         Left            =   -73455
         TabIndex        =   19
         Top             =   60
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   2117
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Sicmact.FlexEdit FlexProv 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   26
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8281
         Cols0           =   15
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Fec Ini-Fec Fin-Dias-Sueldo-Import Vac-EsSalud-IES-Sueldo Anual-7 UIT-Base Imponible-15% Anual-Impuesto Mensual"
         EncabezadosAnchos=   "300-800-4000-1000-1000-1000-1200-1200-1200-1200-1200-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-R-R-R-R-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-5-5-0-0-0-0-0-0-0-0-2-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSMask.MaskEdBox txtFechaVac 
         Height          =   315
         Left            =   -66840
         TabIndex        =   37
         Top             =   120
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.FlexEdit FlexVac 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   38
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8281
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Nro Dias-Imp Quinta-Remuneracion"
         EncabezadosAnchos=   "300-800-4000-1000-1500-1500"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-R-R"
         FormatosEdit    =   "0-0-0-3-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.ctrRRHHGen ctrRRHHGen2 
         Height          =   1200
         Left            =   -74835
         TabIndex        =   45
         Top             =   105
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   2117
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Sicmact.FlexEdit FlexBasePDT 
         Height          =   3945
         Left            =   -74835
         TabIndex        =   48
         Top             =   1380
         Width           =   9165
         _ExtentX        =   16325
         _ExtentY        =   8281
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Nro Dias-Imp Quinta-Remuneracion-Periodo"
         EncabezadosAnchos=   "300-800-5000-0-0-0-1800"
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
         ColumnasAEditar =   "X-X-X-X-X-X-6"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-2"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-R-R-C"
         FormatosEdit    =   "0-0-0-3-2-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   360
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label6 
         Caption         =   "T. Prov:"
         Height          =   240
         Left            =   7560
         TabIndex        =   55
         Top             =   5400
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "T. Sueldo:"
         Height          =   240
         Left            =   3600
         TabIndex        =   53
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Balance:"
         Height          =   240
         Left            =   120
         TabIndex        =   51
         Top             =   5400
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "VACACIONES EJECUTADAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   -74880
         TabIndex        =   40
         Top             =   120
         Width           =   4110
      End
      Begin VB.Label lnlTot 
         Caption         =   "T. Vac:"
         Height          =   240
         Left            =   5760
         TabIndex        =   36
         Top             =   5400
         Width           =   525
      End
      Begin VB.Label lblProceso 
         Caption         =   "Proceso :"
         Height          =   240
         Left            =   285
         TabIndex        =   33
         Top             =   720
         Width           =   810
      End
      Begin VB.Label LblProvision 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "VACACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   -74880
         TabIndex        =   28
         Top             =   195
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Marzo 2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   5265
         TabIndex        =   25
         Top             =   105
         Width           =   1620
      End
      Begin VB.Label lblEstado 
         Caption         =   "Label5"
         Height          =   255
         Left            =   -72210
         TabIndex        =   17
         Top             =   570
         Width           =   3690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "ABONO DE VACACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   180
         TabIndex        =   16
         Top             =   135
         Width           =   3720
      End
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   240
      Left            =   135
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmVacacionesGozadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPersCod As String

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtAño.SetFocus
End If
End Sub

Private Sub cmdAbonar_Click()
Dim i As Integer
Dim opt  As Integer
Dim nValida As Integer
Dim RHV As DRHVacaciones
Dim nResult As Double
Dim nDias As Integer
Set RHV = New DRHVacaciones

If RHV.AbonoVacaciones(mskFecha.Text) > 0 Then
    MsgBox "Periodo " & Format(mskFecha.Text, "MMMM") & " " & Format(mskFecha.Text, "YYYY") & " Generado", vbInformation, "AVISO"
    Set RHV = Nothing
    Exit Sub
End If

opt = MsgBox("Esta Seguro de Abonar Vacaciones", vbInformation + vbYesNo, "AVISO")
If vbNo = opt Then Exit Sub

If flex2.TextMatrix(1, 1) = "" Then
    MsgBox "Cargue a los empleados de la CMAC-MAYNAS", vbInformation, "AVISO"
    Exit Sub
End If
PrgBar.Min = 1
PrgBar.Max = flex2.Rows
PrgBar.Visible = True
For i = 1 To flex2.Rows - 1
    'MAVM 20110715 ***
    'nValida = RHV.AbonaVacacionMes(flex2.TextMatrix(i, 8), Format(gdFecSis, "YYYY/MM/DD"), gsCodUser, flex2.TextMatrix(i, 3), flex2.TextMatrix(i, 4))
    nValida = RHV.AbonaVacacionMes(flex2.TextMatrix(i, 8), mskFecha, gsCodUser, flex2.TextMatrix(i, 3), FechaHora(gdFecSis))
    '***
    Select Case nValida
        Case -1
            flex2.TextMatrix(i, 9) = "Error Verificar Abono"
        Case 0
            flex2.TextMatrix(i, 9) = "Ya se Abono Vacaciones"
        Case 1
            flex2.TextMatrix(i, 9) = "Ok"
            If DateDiff("M", flex2.TextMatrix(i, 3), mskFecha) = 0 Then
                'Comentado Por MAVM 20110713 ***
                'If CInt(Mid(flex2.TextMatrix(i, 3), 1, 2)) = 1 Then
                '    nResult = 2.5
                'Else
                    'If Month(flex2.TextMatrix(i, 3)) = 2 Then
                    '    nDias = 30 - CInt(Mid(Me.flex2.TextMatrix(i, 3), 1, 2)) + 3
                    'Else
                    '    nDias = 30 - CInt(Mid(Me.flex2.TextMatrix(i, 3), 1, 2)) + 1
                    'End If
                '***
                    'Comentado Por MAVM 20110713 ***
                    'nDias = 30 - Mid(flex2.TextMatrix(i, 3), 1, 2) 'DateDiff("d", flex2.TextMatrix(i, 3), mskFecha)
                    'If nDias = 0 Then nDias = 1
                    'nResult = Round((nDias / 30) * 2.5, 2)
                    '***
                    nResult = 2.5 'MAVM 20110713
                'End If 'Comentado Por MAVM 20110713
            Else
                nResult = 2.5
            End If
            flex2.TextMatrix(i, 5) = flex2.TextMatrix(i, 5) + nResult
    End Select
    PrgBar.value = i
Next i
PrgBar.Visible = False
MsgBox "Vacaciones Abonadas", vbInformation, "AVISO"

Dim nOpt As Integer
If Me.OptEst(0).value Then
    nOpt = 0
End If
If Me.OptEst(1).value Then
    nOpt = 1
End If
If Me.OptEst(2).value Then
    nOpt = 2
End If

mskFecha_KeyPress (13) 'MAVM 20110713 ***
End Sub

'Private Sub CmdAdicionar_Click()
'    Dim i As Integer
'    For i = 1 To Flex.Rows - 1
'     If ctrRRHHGen.psCodigoEmpleado = Flex.TextMatrix(i, 1) Then
'        Exit Sub
'     End If
'    Next i
'    Me.Flex.AdicionaFila
'    Flex.TextMatrix(Flex.Rows - 1, 1) = ctrRRHHGen.psCodigoEmpleado
'    Flex.TextMatrix(Flex.Rows - 1, 2) = ctrRRHHGen.psNombreEmpledo
'    Flex.TextMatrix(Flex.Rows - 1, 7) = ctrRRHHGen.psCodigoPersona
'    'Me.cmdAbonar.Enabled = False
'End Sub

Private Sub CmdAdicionarPDT_Click()
    Dim i As Integer
    For i = 1 To FlexBasePDT.Rows - 1
     If ctrRRHHGen2.psCodigoEmpleado = FlexBasePDT.TextMatrix(i, 1) Then
        Exit Sub
     End If
    Next i
    With FlexBasePDT
        .AdicionaFila
        .TextMatrix(.Rows - 1, 1) = ctrRRHHGen2.psCodigoEmpleado
        .TextMatrix(.Rows - 1, 2) = ctrRRHHGen2.psNombreEmpledo
        .TextMatrix(.Rows - 1, 3) = 0
        .TextMatrix(.Rows - 1, 4) = 0
        .TextMatrix(.Rows - 1, 5) = 0
        .TextMatrix(.Rows - 1, 6) = DateAdd("m", -1, gdFecSis)
        If BuscarBasePDT(.TextMatrix(.Rows - 1, 1), .TextMatrix(.Rows - 1, 6)) Then cmdSubsidio.Enabled = True
    End With
    'Me.cmdAbonar.Enabled = False
End Sub

Private Function BuscarBasePDT(ByVal pcRHCod As String, ByVal pcPeriodo As String) As Boolean
On Error GoTo BuscarBasePDTErr
    Dim rs As ADODB.Recordset, sSQL As String, oCon  As DConecta
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select count(*) from RHBasePDT where cRHCod = '" & pcRHCod & "' and cPeriodo = '" & Format(CDate(pcPeriodo), "yyyymm") & "'"
        Set rs = oCon.CargaRecordSet(sSQL)
        oCon.CierraConexion
    End If
    If Not rs.EOF Then
        If rs(0) > 0 Then
            BuscarBasePDT = True
        Else
            BuscarBasePDT = False
        End If
    Else
        BuscarBasePDT = False
    End If
    Exit Function
BuscarBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub cmdAsiento_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim ldFechaAnt As Date
    Dim rsMontoAnt As New ADODB.Recordset
    Dim oCon As New DConecta
    Dim oMov As New DMov
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnItem As Long
    
    Dim oAsi As NContImprimir
    Set oAsi = New NContImprimir
    Dim lsCadena As String
    
    Dim lnAcum As Currency
    Dim lnAcumDiff As Currency
    Dim lnUltimo  As Currency
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    'ldFechaAnt = DateAdd("d", -1, CDate("01/" & Format(CDate(Me.mskFecha), "mm/yyyy")))
    sql = "Select cMovNro  from mov where cmovnro like '" & Format(mskFecha, gsFormatoMovFecha) & "%' And cOpeCod = '622501' And nMovflag = 0"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
     
    If Not rs.EOF And Not rs.BOF Then
        lsMovNro = rs!cMovNro
        MsgBox "El Asiento de Provision ya fue generado.", vbInformation
        
        lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
        oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
        
        oCon.CierraConexion
        Exit Sub
    Else
        rs.Close
    End If
    
    If MsgBox("Desea Generar Asiento Contable ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    'Comentado Por MAVM 20110715 ***
    'sql = "Select dbo.getsaldocta('" & Format(ldFechaAnt, gsFormatoFecha) & "','251503',1)"
    'Dim ldFechaAntBalance As Date
    'ldFechaAntBalance = DateAdd("d", -1, CDate("01/" & Format(CDate(Me.mskFecha), "mm/yyyy")))
    'ldFechaAntBalance = DateAdd("m", -1, Format(CDate(ldFechaAntBalance), "dd/mm/yyyy"))
    
    
    'Sql = "Select dbo.getsaldoctaAcumulado('" & Format(ldFechaAntBalance, gsFormatoFecha) & "','251503%',1)"
    'Set rsMontoAnt = oCon.CargaRecordSet(Sql)
    
    'Sql = " Select isnull(a.cage, b.cage) Age, IsNull(a.nmonto,0) MontoAct , IsNull(b.nmonto,0) MontoAnt, IsNull(a.nmonto,0) - IsNull(b.nmonto,0) Diff from dbo.RHGetProvisionVacacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "'," & CCur(Me.txtTotal.Text) & ") a" _
    '    & " full outer join dbo.RHGetProvisionVacacionesMesAnt('" & Format(CDate(Me.mskFecha.Text), gsFormatoFecha) & "'," & rsMontoAnt.Fields(0) & ") b on a.cage = b.cage"
    'Set rs = oCon.CargaRecordSet(Sql)
    '***
    
    'MAVM 20110715 ***
    sql = "Select cAgenciaActual  As Age, SUM (nProvision) as nRHVacacionesMonto"
    sql = sql & " From RRHH RH Inner Join MovVacaciones MV on RH.cPersCod = MV.cPersCod"
    sql = sql & " And MV.cPeriodo LIKE '" & Mid(Format(CDate(mskFecha), gsFormatoMovFecha), 1, 6) & "%'"
    sql = sql & " Group by cAgenciaActual"
    Set rs = oCon.CargaRecordSet(sql)
    'MAVM 20110715 ***
    
    'lsMovNro = oMov.GeneraMovNro(ldFechaAnt, Right(gsCodAge, 2), gsCodUser)
    lsMovNro = oMov.GeneraMovNro(mskFecha, Right(gsCodAge, 2), gsCodUser)
    lnAcum = 0
    lnItem = 0
    lnAcumDiff = 0
    
    oMov.BeginTrans
        'oMov.InsertaMov lsMovNro, "622501", "Provison de Planilla de Vacaciones. " & Format(ldFechaAnt, gsFormatoFechaView)
        oMov.InsertaMov lsMovNro, "622501", "Provison de Planilla de Vacaciones. " & Format(mskFecha, gsFormatoFechaView)
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        While Not rs.EOF
            'lnAcum = lnAcum + Round(rs!Diff, 2) + Round(rs!MontoAnt, 2)
            'lnAcumDiff = lnAcumDiff + Round(rs!Diff, 2)
            lnItem = lnItem + 1
            oMov.InsertaMovCta lnMovNro, lnItem, "451102" & rs!Age, Round(rs!nRHVacacionesMonto, 2)
            'JEOM ----------------------------------------------------------------------
            lnItem = lnItem + 1
            'oMov.InsertaMovCta lnMovNro, lnItem, "251503" & rs!Age, Round(rs!Diff, 2) * -1
            oMov.InsertaMovCta lnMovNro, lnItem, "251503" & rs!Age, Round(rs!nRHVacacionesMonto, 2) * -1
            'FIN----------------------------------------------------------------------
            'lnUltimo = rs!Diff
            rs.MoveNext
        Wend
        
        'Comentado Por MAVM 20110715 ***
'        If lnAcum <> CCur(Me.txtTotal.Text) Then
''            oMov.ActualizaMovCta lnMovNro, lnItem, , Format(lnUltimo + (CCur(Me.txtTotal.Text) - lnAcum), "0.00")
''            lnAcumDiff = lnAcumDiff + Format(lnUltimo + (CCur(Me.txtTotal.Text) - lnAcum), "0.00")
'            'JEOM----------------------------------------------------------------------------------------------
'            oMov.ActualizaMovCta lnMovNro, lnItem - 1, , lnUltimo + Format((CCur(Me.txtTotal.Text) - lnAcum), "0.00")
'            oMov.ActualizaMovCta lnMovNro, lnItem, , (lnUltimo + Format((CCur(Me.txtTotal.Text) - lnAcum), "0.00")) * -1
'            lnAcumDiff = lnAcumDiff + Format((CCur(Me.txtTotal.Text) - lnAcum), "0.00")
'            'FIN----------------------------------------------------------------------------------------------
'        End If
        '***
        
        'lnItem = lnItem + 1
        'oMov.InsertaMovCta lnMovNro, lnItem, "251503", lnAcumDiff * -1
        
    oMov.CommitTrans
        
    lsCadena = oAsi.ImprimeAsientoContable(lsMovNro, 66, 80, , , , , False)
    
    oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
    oCon.CierraConexion
End Sub

Private Sub cmdBasePDTprovicion_Click()
    GuardarBasePDTProviciones
End Sub

Private Sub cmdBasePDTvacacionE_Click()
    GuardarBasePDTVacacionesE
End Sub

Private Sub CmdBuscar_Click()
Dim i As Integer
Dim RHV As DRHVacaciones
Dim rs As ADODB.Recordset
Set RHV = New DRHVacaciones
If OptBusca(1).value = False Then
    Set rs = RHV.GetPersonalVacacionesMes(Me.txtAño & Trim(Right(Me.cboMes.Text, 3)))
Else
    Set rs = RHV.GetTrabajadorVacaciones(ctrRRHHGen1.psCodigoPersona)
End If
 FlexEdit1.Clear
 FlexEdit1.Rows = 2
 FlexEdit1.FormaCabecera
 
 While Not rs.EOF
    FlexEdit1.AdicionaFila
    FlexEdit1.TextMatrix(FlexEdit1.Rows - 1, 1) = rs!cRHCod
    FlexEdit1.TextMatrix(FlexEdit1.Rows - 1, 2) = rs!nombres
    FlexEdit1.TextMatrix(FlexEdit1.Rows - 1, 3) = rs!dFecIni
    FlexEdit1.TextMatrix(FlexEdit1.Rows - 1, 4) = rs!dFecFin
    FlexEdit1.TextMatrix(FlexEdit1.Rows - 1, 5) = rs!cComenta
    rs.MoveNext
 Wend

End Sub

Private Sub CmdCancelarPDT_Click()
Dim i As Integer
Me.CmdAdicionarPDT.Enabled = False
FlexBasePDT.Rows = 2
For i = 1 To FlexBasePDT.Cols - 1
    FlexBasePDT.TextMatrix(1, i) = ""
Next i
End Sub

Private Sub cmdConsol_Click()
Dim RHV As DRHVacaciones
Dim rs As New ADODB.Recordset
Dim RSVacGoz As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i, Y As Integer
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim dProv01Vert, dProv02Vert, dProv03Vert, dProv04Vert, dProv05Vert, dProv06Vert As Double
Dim dProv07Vert, dProv08Vert, dProv09Vert, dProv10Vert, dProv11Vert, dProv12Vert As Double
Dim dProvHoriz, dProvHorizSum, dProvHorizSumTot As Double
Dim dPago01Vert, dPago02Vert, dPago03Vert, dPago04Vert, dPago05Vert, dPago06Vert As Double
Dim dPago07Vert, dPago08Vert, dPago09Vert, dPago10Vert, dPago11Vert, dPago12Vert As Double
Dim dPagoHoriz, dPagoHorizSum, dPagoHorizSumTot As Double
Dim nDias As Double

Set RHV = New DRHVacaciones
Set rs = RHV.CargarProvVacacionesConsol(Mid(Format(mskFecha.Text, gsFormatoMovFecha), 1, 6))

Screen.MousePointer = 11
lsArchivo = "ProvVacConsol" & Format(Now, "yyyymm") & "_" & Format(Time(), "HHMMSS") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\Spooler\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\Spooler\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = Format(gdFecSis, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("B1") = "CAJA MAYNAS"
xlHoja1.Range("B1:C1").MergeCells = True
xlHoja1.Range("B1").Font.Bold = True
xlHoja1.Range("H1") = gdFecSis
xlHoja1.Range("H2") = gsCodUser
xlHoja1.Range("H2").HorizontalAlignment = xlRight
xlHoja1.Range("G6:G1000").HorizontalAlignment = xlCenter
xlHoja1.Range("F6:F1000").HorizontalAlignment = xlRight
xlHoja1.Range("H1:H2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Agencia"
xlHoja1.Range("E5") = "Area"
xlHoja1.Range("F5") = "Fecha Ingreso"
xlHoja1.Range("G5") = "Factor"
xlHoja1.Range("H5") = "Renumeracion"
xlHoja1.Range("I5") = "Dias Acum"
xlHoja1.Range("J5") = "Dias Prov Ejer"
xlHoja1.Range("K5") = "Dias Goza Ejer"
xlHoja1.Range("L5") = "Enero"
xlHoja1.Range("M5") = "Febrero"
xlHoja1.Range("N5") = "Marzo"
xlHoja1.Range("O5") = "Abril"
xlHoja1.Range("P5") = "Mayo"
xlHoja1.Range("Q5") = "Junio"
xlHoja1.Range("R5") = "Julio"
xlHoja1.Range("S5") = "Agosto"
xlHoja1.Range("T5") = "Setiembre"
xlHoja1.Range("U5") = "Octubre"
xlHoja1.Range("V5") = "Noviembre"
xlHoja1.Range("W5") = "Diciembre"
xlHoja1.Range("X5") = "Total Gasto"
xlHoja1.Range("Y5") = "Total Pagar Acum"
xlHoja1.Range("Z5") = "Pago Enero"
xlHoja1.Range("AA5") = "Pago Febrero"
xlHoja1.Range("AB5") = "Pago Marzo"
xlHoja1.Range("AC5") = "Pago Abril"
xlHoja1.Range("AD5") = "Pago Mayo"
xlHoja1.Range("AE5") = "Pago Junio"
xlHoja1.Range("AF5") = "Pago Julio"
xlHoja1.Range("AG5") = "Pago Agosto"
xlHoja1.Range("AH5") = "Pago Setiembre"
xlHoja1.Range("AI5") = "Pago Octubre"
xlHoja1.Range("AJ5") = "Pago Noviembre"
xlHoja1.Range("AK5") = "Pago Diciembre"
xlHoja1.Range("AL5") = "Total Pag"
xlHoja1.Range("AM5") = "Provision Año"

xlHoja1.Range("B4:H4").MergeCells = True
xlHoja1.Range("B4") = "PROVISION DE  VACACIONES " & " " & UCase(Format(DateAdd("m", -1, gdFecSis), "MMMM")) & " DEL " & Format(DateAdd("m", -1, gdFecSis), "YYYY")
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter

xlHoja1.Range("B5:AM5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:AM5").Interior.ColorIndex = 35
xlHoja1.Range("B5:AM5").Font.Bold = True

xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 45
xlHoja1.Range("D1").ColumnWidth = 25
xlHoja1.Range("E1").ColumnWidth = 25
xlHoja1.Range("F1").ColumnWidth = 12
xlHoja1.Range("G1").ColumnWidth = 9
xlHoja1.Range("H1").ColumnWidth = 13
xlHoja1.Range("I1").ColumnWidth = 13
xlHoja1.Range("J1").ColumnWidth = 13
xlHoja1.Range("K1").ColumnWidth = 13
xlHoja1.Range("L1").ColumnWidth = 12
xlHoja1.Range("M1").ColumnWidth = 12
xlHoja1.Range("N1").ColumnWidth = 12
xlHoja1.Range("O1").ColumnWidth = 12
xlHoja1.Range("P1").ColumnWidth = 12
xlHoja1.Range("Q1").ColumnWidth = 12
xlHoja1.Range("R1").ColumnWidth = 12
xlHoja1.Range("S1").ColumnWidth = 12
xlHoja1.Range("T1").ColumnWidth = 12
xlHoja1.Range("U1").ColumnWidth = 12
xlHoja1.Range("V1").ColumnWidth = 10
xlHoja1.Range("W1").ColumnWidth = 10
xlHoja1.Range("X1").ColumnWidth = 20
xlHoja1.Range("Y1").ColumnWidth = 20
xlHoja1.Range("Z1").ColumnWidth = 16
xlHoja1.Range("AA1").ColumnWidth = 16
xlHoja1.Range("AB1").ColumnWidth = 16
xlHoja1.Range("AC1").ColumnWidth = 16
xlHoja1.Range("AD1").ColumnWidth = 16
xlHoja1.Range("AE1").ColumnWidth = 16
xlHoja1.Range("AF1").ColumnWidth = 16
xlHoja1.Range("AG1").ColumnWidth = 16
xlHoja1.Range("AH1").ColumnWidth = 16
xlHoja1.Range("AI1").ColumnWidth = 16
xlHoja1.Range("AJ1").ColumnWidth = 16
xlHoja1.Range("AK1").ColumnWidth = 16
xlHoja1.Range("AL1").ColumnWidth = 20
xlHoja1.Range("AM1").ColumnWidth = 20

xlHoja1.Application.ActiveWindow.Zoom = 80
xlHoja1.Range("H6:H1000").Style = "Comma"
xlHoja1.Range("Z6:AK1000").Style = "Comma"
xlHoja1.Range("L6:W1000").Style = "Comma"
Y = 6

For i = 1 To rs.RecordCount
    xlHoja1.Range("B" & Y) = rs!cRHCod
    xlHoja1.Range("C" & Y) = rs!cPersNombre
    xlHoja1.Range("D" & Y) = rs!cAgeDescripcion
    xlHoja1.Range("E" & Y) = rs!cAreaDescripcion
    xlHoja1.Range("F" & Y) = "'" & rs!dIngreso
    xlHoja1.Range("G" & Y) = "30"
    xlHoja1.Range("H" & Y) = rs!nRHSueldoMonto
    xlHoja1.Range("I" & Y) = rs!nRHEmplVacacionesPend
    
    If Mid(rs!dIngreso, 7, 4) < "2011" Then
        xlHoja1.Range("J" & Y) = (DateDiff("M", "01/01/2011", Me.mskFecha.Text) + 1) * 2.5
    Else
        xlHoja1.Range("J" & Y) = (DateDiff("M", rs!dIngreso, Me.mskFecha.Text) + 1) * 2.5
    End If
   
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "01" Then
        xlHoja1.Range("L" & Y) = rs!Prov01
        dProv01Vert = dProv01Vert + rs!Prov01
        dProvHoriz = dProvHoriz + rs!Prov01
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "02" Then
        xlHoja1.Range("M" & Y) = rs!Prov02
        dProv02Vert = dProv02Vert + rs!Prov02
        dProvHoriz = dProvHoriz + rs!Prov02
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "03" Then
        xlHoja1.Range("N" & Y) = rs!Prov03
        dProv03Vert = dProv03Vert + rs!Prov03
        dProvHoriz = dProvHoriz + rs!Prov03
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "04" Then
        xlHoja1.Range("O" & Y) = rs!Prov04
        dProv04Vert = dProv04Vert + rs!Prov04
        dProvHoriz = dProvHoriz + rs!Prov04
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "05" Then
        xlHoja1.Range("P" & Y) = rs!Prov05
        dProv05Vert = dProv05Vert + rs!Prov05
        dProvHoriz = dProvHoriz + rs!Prov05
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "06" Then
        xlHoja1.Range("Q" & Y) = rs!Prov06
        dProv06Vert = dProv06Vert + rs!Prov06
        dProvHoriz = dProvHoriz + rs!Prov06
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "07" Then
        xlHoja1.Range("R" & Y) = rs!Prov07
        dProv07Vert = dProv07Vert + rs!Prov07
        dProvHoriz = dProvHoriz + rs!Prov07
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "08" Then
        xlHoja1.Range("S" & Y) = rs!Prov08
        dProv08Vert = dProv08Vert + rs!Prov08
        dProvHoriz = dProvHoriz + rs!Prov08
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "09" Then
        xlHoja1.Range("T" & Y) = rs!Prov09
        dProv09Vert = dProv09Vert + rs!Prov09
        dProvHoriz = dProvHoriz + rs!Prov09
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "10" Then
        xlHoja1.Range("U" & Y) = rs!Prov10
        dProv10Vert = dProv10Vert + rs!Prov10
        dProvHoriz = dProvHoriz + rs!Prov10
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "11" Then
        xlHoja1.Range("V" & Y) = rs!Prov11
        dProv11Vert = dProv11Vert + rs!Prov11
        dProvHoriz = dProvHoriz + rs!Prov11
    End If
    If Mid(Format(mskFecha.Text, gsFormatoMovFecha), 5, 2) >= "12" Then
        xlHoja1.Range("W" & Y) = rs!Prov12
        dProv12Vert = dProv12Vert + rs!Prov12
        dProvHoriz = dProvHoriz + rs!Prov12
    End If
    
    xlHoja1.Range("X" & Y) = Format(dProvHoriz, "#,##0.00")
    dProvHorizSum = dProvHorizSum + dProvHoriz
        
    Set RSVacGoz = RHV.CargarVacacionesGozadas(rs!cPersCod)
    If Mid(rs!dIngreso, 7, 4) < "2011" Then
        nDias = (DateDiff("M", rs!dIngreso, "31/12/2010") + 1) * 2.5
    End If
   
    If RSVacGoz.RecordCount <> "0" Then
        Dim x As Integer
        Dim dVacGoz As Double
        
        dVacGoz = 0
        For x = 0 To RSVacGoz.RecordCount - 1
            If Mid(RSVacGoz!cRRHHPeriodo, 1, 4) < "2011" Then
                nDias = nDias - RSVacGoz!Dias
            Else
                If Mid(RSVacGoz!cRRHHPeriodo, 1, 4) = "2011" Then

                    dVacGoz = dVacGoz + RSVacGoz!Dias
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "01" Then
                        xlHoja1.Range("Z" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago01Vert = dPago01Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "02" Then
                        xlHoja1.Range("AA" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago02Vert = dPago02Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "03" Then
                        xlHoja1.Range("AB" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago03Vert = dPago03Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "04" Then
                        xlHoja1.Range("AC" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago04Vert = dPago04Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "05" Then
                        xlHoja1.Range("AD" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago05Vert = dPago05Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "06" Then
                        xlHoja1.Range("AE" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago06Vert = dPago06Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "07" Then
                        xlHoja1.Range("AF" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago07Vert = dPago07Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "08" Then
                        xlHoja1.Range("AG" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago08Vert = dPago08Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "09" Then
                        xlHoja1.Range("AH" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago09Vert = dPago09Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "10" Then
                        xlHoja1.Range("AI" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago10Vert = dPago10Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "11" Then
                        xlHoja1.Range("AJ" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago11Vert = dPago11Vert + RSVacGoz!MontoVac
                    End If
                    
                    If Mid(RSVacGoz!cRRHHPeriodo, 5, 2) = "12" Then
                        xlHoja1.Range("AK" & Y) = Format(RSVacGoz!MontoVac, "#,##0.00")
                        dPago12Vert = dPago12Vert + RSVacGoz!MontoVac
                    End If
                    
                    dPagoHoriz = dPagoHoriz + RSVacGoz!MontoVac
                End If
            End If
        RSVacGoz.MoveNext
        Next x
        xlHoja1.Range("K" & Y) = IIf(dVacGoz = 0, "", dVacGoz)
    End If
    
    xlHoja1.Range("Y" & Y) = Format(((rs!nRHSueldoMonto / 30) * nDias) + xlHoja1.Range("X" & Y), "#,##0.00")
    
    xlHoja1.Range("AL" & Y) = Format(dPagoHoriz, "#,##0.00")
    xlHoja1.Range("AM" & Y) = Format(xlHoja1.Range("Y" & Y) - dPagoHoriz, "#,##0.00")
    dProvHorizSumTot = dProvHorizSumTot + xlHoja1.Range("Y" & Y)
    dPagoHorizSumTot = dPagoHorizSumTot + xlHoja1.Range("AM" & Y)
    dPagoHorizSum = dPagoHorizSum + dPagoHoriz
    dPagoHoriz = 0
    dProvHoriz = 0
    nDias = 0
    
    rs.MoveNext
    Y = Y + 1
Next i

xlHoja1.Cells(rs.RecordCount + 5, 2) = "TOTALES"
xlHoja1.Cells(rs.RecordCount + 5, 12) = Format(dProv01Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 13) = Format(dProv02Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 14) = Format(dProv03Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 15) = Format(dProv04Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 16) = Format(dProv05Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 17) = Format(dProv06Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 18) = Format(dProv07Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 19) = Format(dProv08Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 20) = Format(dProv09Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 21) = Format(dProv10Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 22) = Format(dProv11Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 23) = Format(dProv12Vert, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 5, 24) = Format(dProvHorizSum, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 25) = Format(dProvHorizSumTot, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 5, 26) = Format(dPago01Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 27) = Format(dPago02Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 28) = Format(dPago03Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 29) = Format(dPago04Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 30) = Format(dPago05Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 31) = Format(dPago06Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 32) = Format(dPago07Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 33) = Format(dPago08Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 34) = Format(dPago09Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 35) = Format(dPago10Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 36) = Format(dPago11Vert, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 37) = Format(dPago12Vert, "#,##0.00")

xlHoja1.Cells(rs.RecordCount + 5, 38) = Format(dPagoHorizSum, "#,##0.00")
xlHoja1.Cells(rs.RecordCount + 5, 39) = Format(dPagoHorizSumTot, "#,##0.00")
    
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente en la carpeta Spooler de SICMACT ADM", vbInformation, "Aviso"

CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
End Sub

Private Sub cmdEliminar_Click()
Dim opt As Integer
If Flex.Row = 0 Then Exit Sub
opt = MsgBox("Desea Eliminar ", vbInformation + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Flex.EliminaFila (Flex.Row)
End Sub

Private Sub CmdEliminarPDT_Click()
Dim opt As Integer
If FlexBasePDT.Row = 0 Then Exit Sub
opt = MsgBox("Desea Eliminar ", vbInformation + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
FlexBasePDT.EliminaFila (FlexBasePDT.Row)
End Sub

Private Sub CmdExcel3_Click()
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i As Integer
Dim Y As Integer
Dim sSuma As String
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim sCabecera As String

sCabecera = ""
If Me.OptEst(1).value = True Then sCabecera = "ESTABLES"
If Me.OptEst(2).value = True Then sCabecera = "CONTRATADOS"

If Me.flex2.TextMatrix(1, 1) = "" Then
    Exit Sub
End If

'On Error GoTo ErrorINFO4Excel
Screen.MousePointer = 11
lsArchivo = "ProvVac" & Format(DateAdd("m", -1, gdFecSis), "YYYYYMM") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = sCabecera & Format(gdFecSis, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("A1") = "CAJA MAYNAS"
xlHoja1.Range("A1").Font.Bold = True
xlHoja1.Range("H1") = gdFecSis
xlHoja1.Range("H2") = gsCodUser
xlHoja1.Range("H2").HorizontalAlignment = xlRight
xlHoja1.Range("H1:H2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Agencia"
xlHoja1.Range("E5") = "Fecha Ingreso"
xlHoja1.Range("F5") = "Remuneracion"
xlHoja1.Range("G5") = "Dias"
xlHoja1.Range("H5") = "Importe Vac"
xlHoja1.Range("I5") = "Importe Prov"
'xlHoja1.Range("J5") = "Area"
'xlHoja1.Range("k5") = "Ultimas Vacaciones"

xlHoja1.Range("B4:H4").MergeCells = True
xlHoja1.Range("B4") = "PROVISION DE  VACACIONES " & sCabecera & " " & UCase(Format(DateAdd("m", -1, gdFecSis), "MMMM")) & " DEL " & Format(DateAdd("m", -1, gdFecSis), "YYYY")
xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter
xlHoja1.Range("A5:I5").HorizontalAlignment = xlCenter
xlHoja1.Range("A5:I5").Interior.ColorIndex = 35
xlHoja1.Range("A5:I5").Font.Bold = True
xlHoja1.Range("A1").ColumnWidth = 6
xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 45
xlHoja1.Range("D1").ColumnWidth = 25
xlHoja1.Range("E1").ColumnWidth = 12
xlHoja1.Range("F1").ColumnWidth = 15
xlHoja1.Range("G1").ColumnWidth = 15
xlHoja1.Range("H1").ColumnWidth = 13
xlHoja1.Range("I1").ColumnWidth = 13
'xlHoja1.Range("J1").ColumnWidth = 20
'xlHoja1.Range("K1").ColumnWidth = 30
xlHoja1.Application.ActiveWindow.Zoom = 80
'xlHoja1.Range("D6:D1000").NumberFormat = "dd/mm/yyyy;@"
'xlHoja1.Range("E6:E1000").NumberFormat = "dd/mm/yyyy;@"
xlHoja1.Range("D6:E1000").Style = "Comma"
xlHoja1.Range("F6:F1000").Style = "Comma"
xlHoja1.Range("F6:H1000").Style = "Comma"
xlHoja1.Range("I6:I1000").Style = "Comma"
Y = 6
Me.PrgBar.Min = 1
Me.PrgBar.Max = flex2.Rows - 1

For i = 1 To flex2.Rows - 1
   xlHoja1.Range("A" & Y) = flex2.TextMatrix(i, 0)
   xlHoja1.Range("B" & Y) = flex2.TextMatrix(i, 1)
   xlHoja1.Range("C" & Y) = flex2.TextMatrix(i, 2)
   
   xlHoja1.Range("D" & Y) = "'" & flex2.TextMatrix(i, 10)
   
   xlHoja1.Range("E" & Y) = "'" & flex2.TextMatrix(i, 3)
   xlHoja1.Range("F" & Y) = flex2.TextMatrix(i, 6)
   xlHoja1.Range("G" & Y) = flex2.TextMatrix(i, 5)
   'xlHoja1.Range("F" & Y) = "'" & flex2.TextMatrix(i, 4)
   
   xlHoja1.Range("H" & Y) = flex2.TextMatrix(i, 7)
   'xlHoja1.Range("I" & Y) = "'" & flex2.TextMatrix(i, 10)
   xlHoja1.Range("I" & Y) = flex2.TextMatrix(i, 14)
   'xlHoja1.Range("K" & Y) = "'" & flex2.TextMatrix(i, 12) & "-" & flex2.TextMatrix(i, 13)
   Y = Y + 1
   Me.PrgBar.value = i
Next i
    
'xlHoja1.SaveAs App.path & "\SPOOLER\" & "VacPRov"
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente en la carpeta Spooler de SICMACT ADM", vbInformation, "Aviso"

CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
'ErrorINFO4Excel:
'    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'    xlLibro.Close
'    ' Cierra Microsoft Excel con el método Quit.
'    xlAplicacion.Quit
'    'Libera los objetos.
'    Set xlAplicacion = Nothing
'    Set xlLibro = Nothing
'    Set xlHoja1 = Nothing
End Sub

Private Sub cmdExportaExcel5_Click()
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i As Integer
Dim Y As Integer
Dim sSuma As String
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim sCabecera As String

sCabecera = "Vacaciones Ejecutadas " & Format(Me.txtFechaVac, "MMMM")

If Me.FlexVac.TextMatrix(1, 1) = "" Then
    Exit Sub
End If

'On Error GoTo ErrorINFO4Excel
Screen.MousePointer = 11
lsArchivo = "Vac_Eje_" & Format(Me.txtFechaVac, "YYYYMM") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = Format(Me.txtFechaVac, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("A1") = "CAJA TRUJILLO"
xlHoja1.Range("A1").Font.Bold = True
xlHoja1.Range("F1") = gdFecSis
xlHoja1.Range("F2") = gsCodUser
xlHoja1.Range("F2").HorizontalAlignment = xlRight
xlHoja1.Range("F1:F2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Dias"
xlHoja1.Range("E5") = "Imp Quinta"
xlHoja1.Range("F5") = "Rem Total"

xlHoja1.Range("B4:F4").MergeCells = True
xlHoja1.Range("B4") = "Vacaciones Ejecutadas de " & Format(Me.txtFechaVac, "MMMM")

xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:F5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:F5").Interior.ColorIndex = 35
xlHoja1.Range("B5:F5").Font.Bold = True
xlHoja1.Range("A1").ColumnWidth = 6
xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 45
xlHoja1.Range("D1").ColumnWidth = 12
xlHoja1.Range("E1").ColumnWidth = 15
xlHoja1.Range("F1").ColumnWidth = 15
xlHoja1.Application.ActiveWindow.Zoom = 80
xlHoja1.Range("E6:F1000").Style = "Comma"
Y = 6
Me.PrgBar.Min = 1
Me.PrgBar.Max = Me.FlexVac.Rows - 1

For i = 1 To FlexVac.Rows - 1
   xlHoja1.Range("A" & Y) = FlexVac.TextMatrix(i, 0)
   xlHoja1.Range("B" & Y) = FlexVac.TextMatrix(i, 1)
   xlHoja1.Range("C" & Y) = FlexVac.TextMatrix(i, 2)
   xlHoja1.Range("D" & Y) = FlexVac.TextMatrix(i, 3)
   xlHoja1.Range("E" & Y) = FlexVac.TextMatrix(i, 4)
   xlHoja1.Range("F" & Y) = FlexVac.TextMatrix(i, 5)
   Y = Y + 1
   Me.PrgBar.value = i
Next i

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente, en la carpeta Spooler del SICMACT ADM", vbInformation, "Aviso"
'CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
'ErrorINFO4Excel:
'    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'    xlLibro.Close
'    ' Cierra Microsoft Excel con el método Quit.
'    xlAplicacion.Quit
'    'Libera los objetos.
'    Set xlAplicacion = Nothing
'    Set xlLibro = Nothing
'    Set xlHoja1 = Nothing
End Sub

Private Sub CmdExpProvVac_Click()
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String

Dim lsNomHoja As String
Dim i As Integer
Dim Y As Integer
Dim sSuma As String
Dim lbExisteHoja As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim sCabecera As String

sCabecera = "Provision de Vacaciones" & Format(DateAdd("m", -1, gdFecSis), "MMMM")

If Me.FlexProv.TextMatrix(1, 1) = "" Then
    Exit Sub
End If

'On Error GoTo ErrorINFO4Excel
Screen.MousePointer = 11
lsArchivo = "P_VacPDT" & Format(DateAdd("m", -1, gdFecSis), "YYYYMM") & ".xls"

Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

lsNomHoja = Format(gdFecSis, "YYYYMM")
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next
If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

Me.PrgBar.Visible = True
xlHoja1.Range("A1") = "CAJA TRUJILLO"
xlHoja1.Range("A1").Font.Bold = True
xlHoja1.Range("F1") = gdFecSis
xlHoja1.Range("F2") = gsCodUser
xlHoja1.Range("F2").HorizontalAlignment = xlRight
xlHoja1.Range("F1:F2").Font.Bold = True

xlHoja1.Range("B5") = "Cod Emp"
xlHoja1.Range("C5") = "Nombre"
xlHoja1.Range("D5") = "Fec Ini"
xlHoja1.Range("E5") = "Fec Fin"
xlHoja1.Range("F5") = "Dias"
xlHoja1.Range("G5") = "Sueldo"
xlHoja1.Range("H5") = "Importe Vac"
xlHoja1.Range("I5") = "EsSalud"
xlHoja1.Range("J5") = "IES"
xlHoja1.Range("K5") = "Sueldo Anual"
xlHoja1.Range("L5") = "7 UIT"
xlHoja1.Range("M5") = "Base Disp"
xlHoja1.Range("N5") = "15% Anual"
xlHoja1.Range("O5") = "Imp. Anual"

xlHoja1.Range("B4:O4").MergeCells = True
'xlHoja1.Range("B4") = "Provision de Vacaciones " & Me.LblProvision

xlHoja1.Range("B4").Font.Bold = True
xlHoja1.Range("B4").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:O5").HorizontalAlignment = xlCenter
xlHoja1.Range("B5:O5").Interior.ColorIndex = 35
xlHoja1.Range("B5:O5").Font.Bold = True
xlHoja1.Range("A1").ColumnWidth = 6
xlHoja1.Range("B1").ColumnWidth = 9
xlHoja1.Range("C1").ColumnWidth = 45
xlHoja1.Range("D1").ColumnWidth = 12
xlHoja1.Range("E1").ColumnWidth = 12
xlHoja1.Range("F1").ColumnWidth = 11
xlHoja1.Range("G1").ColumnWidth = 13
xlHoja1.Range("H1").ColumnWidth = 13
xlHoja1.Range("I1").ColumnWidth = 13
xlHoja1.Range("J1").ColumnWidth = 13
xlHoja1.Range("K1").ColumnWidth = 13
xlHoja1.Range("L1").ColumnWidth = 13
xlHoja1.Range("M1").ColumnWidth = 13
xlHoja1.Range("N1").ColumnWidth = 13
xlHoja1.Range("O1").ColumnWidth = 13
xlHoja1.Application.ActiveWindow.Zoom = 80
'xlHoja1.Range("D6:D1000").NumberFormat = "dd/mm/yyyy;@"
'xlHoja1.Range("E6:E1000").NumberFormat = "dd/mm/yyyy;@"
xlHoja1.Range("D6:E1000").Style = "Comma"
xlHoja1.Range("F6:F1000").Style = "Comma"
xlHoja1.Range("F6:O1000").Style = "Comma"
Y = 6
Me.PrgBar.Min = 1
Me.PrgBar.Max = Me.FlexProv.Rows - 1

For i = 1 To FlexProv.Rows - 1
   xlHoja1.Range("A" & Y) = FlexProv.TextMatrix(i, 0)
   xlHoja1.Range("B" & Y) = FlexProv.TextMatrix(i, 1)
   xlHoja1.Range("C" & Y) = FlexProv.TextMatrix(i, 2)
   xlHoja1.Range("D" & Y) = "'" & FlexProv.TextMatrix(i, 3)
   xlHoja1.Range("E" & Y) = "'" & FlexProv.TextMatrix(i, 4)
   xlHoja1.Range("F" & Y) = FlexProv.TextMatrix(i, 5)
   xlHoja1.Range("G" & Y) = FlexProv.TextMatrix(i, 6)
   xlHoja1.Range("H" & Y) = FlexProv.TextMatrix(i, 7)
   xlHoja1.Range("I" & Y) = FlexProv.TextMatrix(i, 8)
   xlHoja1.Range("J" & Y) = FlexProv.TextMatrix(i, 9)
   xlHoja1.Range("K" & Y) = FlexProv.TextMatrix(i, 10)
   xlHoja1.Range("L" & Y) = FlexProv.TextMatrix(i, 11)
   xlHoja1.Range("M" & Y) = FlexProv.TextMatrix(i, 12)
   xlHoja1.Range("N" & Y) = FlexProv.TextMatrix(i, 13)
   xlHoja1.Range("O" & Y) = FlexProv.TextMatrix(i, 14)
   Y = Y + 1
   Me.PrgBar.value = i
Next i

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.

Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
Me.PrgBar.Visible = False
MsgBox "Se ha Generado el Archivo " & lsArchivo & " Satisfactoriamente, en la carpeta Spooler del SICMACT ADM", vbInformation, "Aviso"
'CargaArchivo lsArchivo, App.path & "\SPOOLER\"
Exit Sub
'ErrorINFO4Excel:
'    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
'    xlLibro.Close
'    ' Cierra Microsoft Excel con el método Quit.
'    xlAplicacion.Quit
'    'Libera los objetos.
'    Set xlAplicacion = Nothing
'    Set xlLibro = Nothing
'    Set xlHoja1 = Nothing

End Sub

Private Sub CmdGenerarProv_Click()
CargaProvisionMes
End Sub

Private Sub CmdGrabarPDT_Click()
Dim i As Integer
Dim j As Integer
Dim oCon  As DConecta, sSQL As String

If FlexBasePDT.TextMatrix(1, 1) = "" Then Exit Sub

For i = 1 To FlexBasePDT.Rows - 1
    For j = 3 To 4
        If FlexBasePDT.TextMatrix(i, j) = "" Then
            FlexBasePDT.Row = i
            FlexBasePDT.Col = j
            MsgBox "Datos Incompletos", vbInformation, "AVISO"
            Exit Sub
        End If
    Next j
Next i
Set oCon = New DConecta
If oCon.AbreConexion Then
    'Graba FlexBasePDT
    For i = 1 To FlexBasePDT.Rows - 1
        With FlexBasePDT
            sSQL = "insert into RHBasePDT(cRHCod,cPeriodo,nRemuneracion,nImpuesto,nDia) " & _
                   " values('" & .TextMatrix(i, 1) & "','" & Format(.TextMatrix(i, 6), "yyyymm") & "',0,0,0)"
            oCon.Ejecutar sSQL
        End With
    Next i
End If
MsgBox "Se grabo correctamente la información", vbInformation, "AVISO"
End Sub

Private Sub CmdOcultar_Click()
If CmdOcultar.Caption = "<< Ocultar >>" Then
    Me.FlexProv.ColWidth(3) = 0
    Me.FlexProv.ColWidth(4) = 0
    Me.FlexProv.ColWidth(5) = 0
    Me.FlexProv.ColWidth(6) = 0
    CmdOcultar.Caption = "<< Mostrar >>"
Else
    CmdOcultar.Caption = "<< Ocultar >>"
    Me.FlexProv.ColWidth(3) = 1000
    Me.FlexProv.ColWidth(4) = 1000
    Me.FlexProv.ColWidth(5) = 1000
    Me.FlexProv.ColWidth(6) = 1200
End If
End Sub

Private Sub cmdProvision_Click()
Dim dProvision As Double
Dim RHV As DRHVacaciones
Dim dMontoProv As Double
Dim i As Integer
Dim oCon As New DConecta 'MAVM 20110715
Dim sql As String 'MAVM 20110715
Dim rs As ADODB.Recordset

'Comentado Por MAVM 20110912
'If txtSaldoBalance.Text = "0.00" Or txtSaldoBalance.Text = "" Then
'    MsgBox "Debe Ingresar el Saldo del Balance", vbCritical, "AVISO"
'    Exit Sub
'End If

sql = "Select cMovNro  from mov where cmovnro like '" & Format(mskFecha, gsFormatoMovFecha) & "%' And cOpeCod = '622501' And nMovflag = 0"
oCon.AbreConexion
    
Set rs = oCon.CargaRecordSet(sql)
     
If Not rs.EOF And Not rs.BOF Then
    MsgBox "La Provision ya fue generado", vbInformation
    oCon.CierraConexion
    mskFecha_KeyPress (13)
    Exit Sub
Else
    rs.Close
End If

Set RHV = New DRHVacaciones
'Comentado Por MAVM 20110912
'dMontoProv = Round(((txtTotal.Text - txtSaldoBalance.Text) / Me.txtTotalSueldo), 8)

If flex2.Rows - 1 <> "0" Then
    PrgBar.Visible = True
    PrgBar.Min = 1
    PrgBar.Max = flex2.Rows - 1
End If
   
For i = 1 To flex2.Rows - 1
With flex2
    'Comentado Por MAVM 20110912
    '.TextMatrix(I, 14) = Format(.TextMatrix(I, 6) * dMontoProv, "#,##0.00")
    .TextMatrix(i, 14) = Format(.TextMatrix(i, 6) / 12, "#,##0.00")
    dProvision = dProvision + Format(IIf(IsNull(.TextMatrix(i, 14)), 0, .TextMatrix(i, 14)), "#,##0.00")
    RHV.ActualizaProvision Format(IIf(IsNull(.TextMatrix(i, 14)), 0, .TextMatrix(i, 14)), "#,##0.00"), .TextMatrix(i, 8), Format(mskFecha.Text, "YYYYMM")
    PrgBar.value = i
End With
Next i
    
txtProvision.Text = Format(dProvision, "#,##0.00")
PrgBar.Visible = False
MsgBox "Provision de Vacaciones realizada", vbInformation, "AVISO"
cmdAsiento.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CmdCancelar_Click()
Dim i As Integer
Me.CmdAdicionar.Enabled = False
Flex.Rows = 2
For i = 1 To Flex.Cols - 1
    Flex.TextMatrix(1, i) = ""
Next i
End Sub

Private Sub cmdGrabar_Click()
Dim i As Integer
Dim j As Integer
Dim RHV As DRHVacaciones

If Flex.TextMatrix(1, 1) = "" Then Exit Sub

For i = 1 To Flex.Rows - 1
    For j = 3 To 4
        If Flex.TextMatrix(i, j) = "" Then
            Flex.Row = i
            Flex.Col = j
            MsgBox "Datos Incompletos", vbInformation, "AVISO"
            Exit Sub
        End If
    Next j
Next i

Set RHV = New DRHVacaciones

'Graba Flex
For i = 1 To Flex.Rows - 1
    With Flex
         Call RHV.InsertaDiasVacaciones(.TextMatrix(i, 7), Format(.TextMatrix(i, 3), "YYYY/MM/DD"), Format(.TextMatrix(i, 4), "YYYY/MM/DD"), _
        .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 1), gsCodUser, Format(gdFecSis, "YYYY/MM/DD"))
    End With
Next i
MsgBox "Se grabo correctamente la información", vbInformation, "AVISO"
End Sub

Private Sub cmdSubsidio_Click()
On Error GoTo cmdSubsidioErr
    Dim oCon  As DConecta, sSQL As String, i As Integer
    Set oCon = New DConecta
    If MsgBox("Seguro que Desea Registrar Subsidios...?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If oCon.AbreConexion Then
        With FlexBasePDT
            i = 1
            Do While i < .Rows
                sSQL = "update RHBasePDTDetalle set bSubsidio=1 where cTipo='P' and cRHCod = '" & .TextMatrix(i, 1) & "' and cPeriodo = '" & Format(CDate(.TextMatrix(i, 6)), "yyyymm") & "'"
                oCon.Ejecutar sSQL
                i = i + 1
            Loop
        End With
        oCon.CierraConexion
    End If
    MsgBox "Se Registro los Subsidios Correctamente", vbInformation, "Aviso"
    Exit Sub
cmdSubsidioErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

'Private Sub ctrRRHHGen_EmiteDatos()
'    Dim oPersona As UPersona
'    Dim oRRHH As DActualizaDatosRRHH
'    Dim RHV As DRHVacaciones
'    'ClearScreen
'    Set oRRHH = New DActualizaDatosRRHH
'    Set oPersona = New UPersona
'    Set oPersona = frmBuscaPersona.Inicio(True)
'    If Not oPersona Is Nothing Then
'        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
'        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
'        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
'        Set RHV = New DRHVacaciones
'        If RHV.VerificaDespido(oPersona.sPersCod) Then
'            Me.CmdAdicionar.Enabled = True
'            MsgBox "El trabajador no se encuentra Labaorando en la CMACT", vbInformation, "AVISO"
'            Set oRRHH = Nothing
'            Exit Sub
'        Else
'            Me.CmdAdicionar.Enabled = True
'
'        End If
'        Set RHV = Nothing
'    End If
'    Set oRRHH = Nothing
'
'End Sub


'Private Sub ctrRRHHGen_KeyPress(KeyAscii As Integer)
'
'Dim oRRHH As DActualizaDatosRRHH
'Dim RHV As DRHVacaciones
'
'If KeyAscii = 13 Then
'        Dim rsR As ADODB.Recordset
'        Set oRRHH = New DActualizaDatosRRHH
'        ctrRRHHGen.psCodigoEmpleado = Left(ctrRRHHGen.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHHGen.psCodigoEmpleado, 2)), "00000")
'        Dim oCon As DActualizaDatosContrato
'        Set oCon = New DActualizaDatosContrato
'
'        Set rsR = oRRHH.GetRRHH(ctrRRHHGen.psCodigoEmpleado, gPersIdDNI)
'
'        If Not (rsR.EOF And rsR.BOF) Then
'            ctrRRHHGen.SpinnerValor = CInt(Right(ctrRRHHGen.psCodigoEmpleado, 5))
'            ctrRRHHGen.psCodigoPersona = rsR.Fields("Codigo")
'            ctrRRHHGen.psNombreEmpledo = rsR.Fields("Nombre")
'
'            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHHGen.psCodigoEmpleado)
'            rsR.Close
'            Set RHV = New DRHVacaciones
'            If RHV.VerificaDespido(ctrRRHHGen.psCodigoPersona) Then
'                MsgBox "El trabajador no se encuentra Labaorando en la CMACT", vbInformation, "AVISO"
'                Me.CmdAdicionar.Enabled = False
'                Set oRRHH = Nothing
'                Exit Sub
'            Else
'                Me.CmdAdicionar.Enabled = True
'            End If
'            Set RHV = Nothing
'            Me.CmdAdicionar.SetFocus
'            Exit Sub
'        Else
'            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
'            'ctrRRHHGen.SetFocus
'            Exit Sub
'        End If
'        rsR.Close
'        Set rsR = Nothing
'End If
'End Sub

Private Sub ctrRRHHGen1_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Me.ctrRRHHGen1.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen1.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen1.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen1.psCodigoPersona)
    End If
    Set oRRHH = Nothing
End Sub

Private Sub ctrRRHHGen1_KeyPress(KeyAscii As Integer)
Dim oRRHH As DActualizaDatosRRHH
If KeyAscii = 13 Then
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        ctrRRHHGen1.psCodigoEmpleado = Left(ctrRRHHGen1.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHHGen1.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        Set rsR = oRRHH.GetRRHH(ctrRRHHGen1.psCodigoEmpleado, gPersIdDNI)
        If Not (rsR.EOF And rsR.BOF) Then
            ctrRRHHGen1.SpinnerValor = CInt(Right(ctrRRHHGen1.psCodigoEmpleado, 5))
            ctrRRHHGen1.psCodigoPersona = rsR.Fields("Codigo")
            ctrRRHHGen1.psNombreEmpledo = rsR.Fields("Nombre")
            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHHGen1.psCodigoEmpleado)
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            ctrRRHHGen1.SetFocus
            Exit Sub
        End If
        rsR.Close
        Set rsR = Nothing
End If
End Sub

Private Sub ctrRRHHGen2_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Dim RHV As DRHVacaciones
    'ClearScreen
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Me.ctrRRHHGen2.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen2.psNombreEmpledo = oPersona.sPersNombre
        'Me.ctrRRHHGen2.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        Set RHV = New DRHVacaciones
        If RHV.VerificaDespido(oPersona.sPersCod) Then
            Me.CmdAdicionarPDT.Enabled = True
            MsgBox "El trabajador no se encuentra Labaorando en la CMACT", vbInformation, "AVISO"
            Set oRRHH = Nothing
            Exit Sub
        Else
            Me.CmdAdicionar.Enabled = True
        End If
        Set RHV = Nothing
    End If
    Set oRRHH = Nothing
    'Me.CmdAdicionar.SetFocus
End Sub

Private Sub ctrRRHHGen2_KeyPress(KeyAscii As Integer)
    Dim oRRHH As DActualizaDatosRRHH
Dim RHV As DRHVacaciones

If KeyAscii = 13 Then
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        ctrRRHHGen2.psCodigoEmpleado = Left(ctrRRHHGen2.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHHGen2.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(ctrRRHHGen2.psCodigoEmpleado, gPersIdDNI)
           
        If Not (rsR.EOF And rsR.BOF) Then
            ctrRRHHGen2.SpinnerValor = CInt(Right(ctrRRHHGen2.psCodigoEmpleado, 5))
            ctrRRHHGen2.psCodigoPersona = rsR.Fields("Codigo")
            ctrRRHHGen2.psNombreEmpledo = rsR.Fields("Nombre")
            
            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHHGen2.psCodigoEmpleado)
            rsR.Close
            Set RHV = New DRHVacaciones
            If RHV.VerificaDespido(ctrRRHHGen2.psCodigoPersona) Then
                MsgBox "El trabajador no se encuentra Labaorando en la CMACT", vbInformation, "AVISO"
                Me.CmdAdicionarPDT.Enabled = True
                Set oRRHH = Nothing
                Exit Sub
            Else
                Me.CmdAdicionarPDT.Enabled = True
            End If
            Set RHV = Nothing
            Me.CmdAdicionarPDT.SetFocus
            Exit Sub
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            'ctrRRHHGen.SetFocus
            Exit Sub
        End If
        rsR.Close
        Set rsR = Nothing
End If

End Sub

Private Sub Flex_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim nDias As Long
If pnCol = 4 Then
    If Flex.TextMatrix(pnRow, Flex.Col) = "" Then
    Else
        If Flex.TextMatrix(pnRow - 1, Flex.Col) = "" Then
            Flex.Row = pnRow - 1
            Flex.Col = pnCol
            Exit Sub
        End If
        
        nDias = DateDiff("d", Flex.TextMatrix(pnRow, pnCol - 1), Flex.TextMatrix(pnRow, pnCol))
        If nDias <= 0 Then
            MsgBox "Fecha Incorrecta", vbInformation, "AVISO"
            Flex.TextMatrix(pnRow, pnCol) = ""
            Exit Sub
        Else
            Flex.TextMatrix(pnRow, pnCol + 1) = nDias
        End If
    End If
End If
End Sub

Sub VisibleBusca(ByVal OptBusca As Integer, ByVal pbVisible As Boolean)
If OptBusca = 0 Then
    Me.ctrRRHHGen1.Visible = pbVisible
    Me.cboMes.Visible = Not pbVisible
    Me.txtAño.Visible = Not pbVisible
Else
    Me.ctrRRHHGen1.Visible = Not pbVisible
    Me.cboMes.Visible = pbVisible
    Me.txtAño.Visible = pbVisible
End If
End Sub

Private Sub Form_Load()
Dim FechaP As Date
SSTab1.Tab = 2 '0
'SSTab1.TabVisible(5) = False
CargaMeses
Me.txtAño = Format(gdFecSis, "YYYY")
Me.cboMes.ListIndex = (Month(gdFecSis) - 1)
Me.mskFecha.Text = Format(gdFecSis, gsFormatoFechaView)
Label1 = Format(DateAdd("m", -1, gdFecSis), "MMMM") & " " & Format(DateAdd("m", -1, gdFecSis), "YYYY")
FechaP = "01" & Mid(gdFecSis, 3, 10)
Me.mskFecha = DateAdd("d", -1, FechaP)
txtFechaVac = DateAdd("d", -1, FechaP)
End Sub

Sub CargaVaciones()
Dim sql As String
Dim RSS As New ADODB.Recordset
Dim Co  As DConecta
Set Co = New DConecta

Me.Flex.AdicionaFila

Co.AbreConexion
Set RSS = Co.Ejecutar(sql)
Co.CierraConexion
End Sub

Private Sub OptBusca_Click(Index As Integer)
FlexEdit1.Rows = 2
FlexEdit1.Clear
FlexEdit1.FormaCabecera
Call VisibleBusca(Index, False)
End Sub

'MAVM 20110713 ***
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
Dim RHV As DRHVacaciones
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim lnMonto As Currency
Dim lnTotalSueldo As Currency
Dim FechaProv As Date
Dim dProvision As Double 'MAVM 20110715

If KeyAscii = 13 Then
    If ValFecha(Me.mskFecha) = False Then
        Exit Sub
    End If
    FechaProv = mskFecha
    
    Set RHV = New DRHVacaciones
    Set rs = RHV.GetProvisionVacaciones(0, mskFecha.Text)
        
    'Borrar Flex2
    flex2.Rows = 2
    For i = 1 To flex2.Cols - 1
        flex2.TextMatrix(1, i) = ""
    Next i
    
    lnMonto = 0
        
    If Not (rs.EOF And rs.BOF) Then
        PrgBar.Visible = True
        PrgBar.Min = 1
        PrgBar.Max = rs.RecordCount
    End If
    While Not rs.EOF
        flex2.AdicionaFila
        flex2.TextMatrix(flex2.Rows - 1, 1) = rs!cRHCod
        flex2.TextMatrix(flex2.Rows - 1, 2) = rs!cPersNombre
        flex2.TextMatrix(flex2.Rows - 1, 3) = Format(rs!dIngreso, "DD/MM/YYYY")
        'flex2.TextMatrix(flex2.Rows - 1, 4) = DateAdd("D", -1, "01" & Mid(gdFecSis, 3, 8))
        flex2.TextMatrix(flex2.Rows - 1, 5) = rs!nRHEmplVacacionesPend
        flex2.TextMatrix(flex2.Rows - 1, 6) = Format(IIf(IsNull(rs!nBono), 0, rs!nBono), "#,##0.00")
        If flex2.TextMatrix(flex2.Rows - 1, 5) = 0 Then
            flex2.TextMatrix(flex2.Rows - 1, 7) = 0
        Else
            flex2.TextMatrix(flex2.Rows - 1, 7) = Format(IIf(IsNull(rs!nRHVacacionesMonto), 0, rs!nRHVacacionesMonto), "#,##0.00")
            'flex2.TextMatrix(flex2.Rows - 1, 14) = flex2.TextMatrix(flex2.Rows - 1, 7)
        End If
        lnMonto = lnMonto + Format(IIf(IsNull(rs!nRHVacacionesMonto), 0, rs!nRHVacacionesMonto), "#,##0.00")
        lnTotalSueldo = lnTotalSueldo + Format(IIf(IsNull(rs!nBono), 0, rs!nBono), "#,##0.00")
        flex2.TextMatrix(flex2.Rows - 1, 8) = rs!cPersCod
        
        flex2.TextMatrix(flex2.Rows - 1, 10) = rs!cAgeDescripcion
        'flex2.TextMatrix(flex2.Rows - 1, 11) = rs!cAreaDescripcion
        'flex2.TextMatrix(flex2.Rows - 1, 12) = IIf(IsNull(rs!feciniv), "", rs!feciniv)
        'flex2.TextMatrix(flex2.Rows - 1, 13) = IIf(IsNull(rs!fecfinv), "", rs!fecfinv)
        flex2.TextMatrix(flex2.Rows - 1, 14) = IIf(rs!nRHrovisionMonto <> 0, Format(rs!nRHrovisionMonto, "#,##0.00"), "")
        dProvision = dProvision + Format(IIf(flex2.TextMatrix(flex2.Rows - 1, 14) = "", 0, flex2.TextMatrix(flex2.Rows - 1, 14)), "#,##0.00")

        
        PrgBar.value = rs.Bookmark
        rs.MoveNext
    Wend
    PrgBar.Visible = False
    Me.txtTotal.Text = Format(lnMonto, "#,##0.00")
    Me.txtTotalSueldo.Text = Format(lnTotalSueldo, "#,##0.00")
    Me.txtProvision.Text = Format(dProvision, "#,##0.00")
End If
End Sub

'Comentado Por MAVM 20110713
'Private Sub OptEst_Click(Index As Integer)
'Dim RHV As DRHVacaciones
'Dim rs As New ADODB.Recordset
'Dim i As Integer
'Dim lnMonto As Currency
'
'Dim FechaProv As Date
'
'
'If ValFecha(Me.mskFecha) = False Then
'    Exit Sub
'End If
'FechaProv = mskFecha
'
'Set RHV = New DRHVacaciones
'Set rs = RHV.GetProvisionVacaciones(Index, gdFecSis) ', gdFecSis)
'
''Borrar Flex2
'flex2.Rows = 2
'For i = 1 To flex2.Cols - 1
'    flex2.TextMatrix(1, i) = ""
'Next i
'
'lnMonto = 0
'
'If Not (rs.EOF And rs.BOF) Then
'    PrgBar.Visible = True
'    PrgBar.Min = 1
'    PrgBar.Max = rs.RecordCount
'End If
'While Not rs.EOF
'    flex2.AdicionaFila
'    flex2.TextMatrix(flex2.Rows - 1, 1) = rs!cRHCod
'    flex2.TextMatrix(flex2.Rows - 1, 2) = rs!cPersNombre
'    flex2.TextMatrix(flex2.Rows - 1, 3) = Format(rs!dIngreso, "DD/MM/YYYY")
'    flex2.TextMatrix(flex2.Rows - 1, 4) = DateAdd("D", -1, "01" & Mid(gdFecSis, 3, 8))
'    flex2.TextMatrix(flex2.Rows - 1, 5) = rs!nRHEmplVacacionesPend
'    'flex2.TextMatrix(flex2.Rows - 1, 6) = rs!nRHSueldoMonto + rs!nBono
'    flex2.TextMatrix(flex2.Rows - 1, 6) = IIf(IsNull(rs!nBono), rs!nRHSueldoMonto, rs!nBono)
'    If flex2.TextMatrix(flex2.Rows - 1, 5) = 0 Then
'        flex2.TextMatrix(flex2.Rows - 1, 7) = 0
'    Else
'        'flex2.TextMatrix(flex2.Rows - 1, 7) = Format(Round((((rs!nRHSueldoMonto) / 30) * flex2.TextMatrix(flex2.Rows - 1, 5)), 2), "#0.00")
'        'lnMonto = lnMonto + Format((((rs!nRHSueldoMonto) / 30) * flex2.TextMatrix(flex2.Rows - 1, 5)), "#.00")
'
'        flex2.TextMatrix(flex2.Rows - 1, 7) = Format(Round((((IIf(IsNull(rs!nBono), rs!nRHSueldoMonto, rs!nBono)) / 30) * flex2.TextMatrix(flex2.Rows - 1, 5)), 2), "#0.00")
'        lnMonto = lnMonto + Format((((IIf(IsNull(rs!nBono), 0, rs!nBono)) / 30) * flex2.TextMatrix(flex2.Rows - 1, 5)), "#.00")
'    End If
'    flex2.TextMatrix(flex2.Rows - 1, 8) = rs!cPersCod
'    flex2.TextMatrix(flex2.Rows - 1, 10) = rs!cAgeDescripcion
'    flex2.TextMatrix(flex2.Rows - 1, 11) = rs!cAreaDescripcion
'    flex2.TextMatrix(flex2.Rows - 1, 12) = IIf(IsNull(rs!feciniv), "", rs!feciniv)
'    flex2.TextMatrix(flex2.Rows - 1, 13) = IIf(IsNull(rs!fecfinv), "", rs!fecfinv)
'
'
'    PrgBar.value = rs.Bookmark
'    rs.MoveNext
'Wend
'PrgBar.Visible = False
'
'Me.txtTotal.Text = Format(lnMonto, "#,##0.00")
'
'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case PreviousTab
    Case 0
    Case 1
    Case 2
    Case 3
End Select
End Sub

Private Sub CargaProvisionMes()
Dim RHVAC As DRHVacaciones
Dim sFecha As Date
Dim sFechaFin As Date
Dim nEsSalud As Double
Dim nIES As Double
Dim nCatV As Double
Dim nUIT As Double
Dim rs As ADODB.Recordset
Dim rsC As ADODB.Recordset

Dim RSS As ADODB.Recordset
Set RHVAC = New DRHVacaciones
Dim Monto As Double
sFecha = gdFecSis
sFecha = "01/" & Mid(gdFecSis, 4, 10)
sFecha = DateAdd("d", -1, sFecha)
sFechaFin = DateAdd("D", -1, "01" & Mid(DateAdd("M", 1, sFecha), 3, 10))
'MsgBox sFecha

'Set Rs = New ADODB.Recordset
Set rs = RHVAC.GetProvisionMes(CStr(sFecha), Month(sFecha))
Me.FlexProv.Clear
Me.FlexProv.FormaCabecera
Me.FlexProv.Rows = 2
Me.FlexProv.FixedCols = 3
nEsSalud = RHVAC.GetConceptoTablaImp("713")
nIES = RHVAC.GetConceptoTablaImp("714")
nCatV = RHVAC.GetConceptoTablaImp("712")
nUIT = RHVAC.GetConceptoTablaImp("724")
LblProvision = "PROVISION DE VACACIONES AL " & sFecha
If Not (rs.EOF And rs.BOF) Then
    Me.PrgBar.Min = 1
    Me.PrgBar.Max = rs.RecordCount
    Me.PrgBar.Visible = True
    While Not rs.EOF
        'If rs!CRHCod = "E00547" Then
        '    MsgBox "X"
        'End If
        If rs!nRHEmplVacacionesPend >= 30 Then
            Me.FlexProv.AdicionaFila
            With Me.FlexProv
                .TextMatrix(.Rows - 1, 1) = rs!cRHCod
                .TextMatrix(.Rows - 1, 2) = rs!cPersNombre
                .TextMatrix(.Rows - 1, 3) = Format(rs!dIngreso, "DD/MM/YYYY")
                .TextMatrix(.Rows - 1, 4) = sFechaFin
                .TextMatrix(.Rows - 1, 5) = 30 'Format(rs!nRHEmplVacacionesPend, "#0.00")
                .TextMatrix(.Rows - 1, 6) = Format(rs!nRHSueldoMonto, "#0.00")
                .TextMatrix(.Rows - 1, 7) = Format((rs!nRHSueldoMonto * 30) / 30, "#0.00")
                .TextMatrix(.Rows - 1, 8) = Format(.TextMatrix(.Rows - 1, 6) * nEsSalud, "#0.00")
                .TextMatrix(.Rows - 1, 9) = Format(.TextMatrix(.Rows - 1, 6) * nIES, "#0.00")
                .TextMatrix(.Rows - 1, 10) = Format(.TextMatrix(.Rows - 1, 6) * 14, "#0.00")
                .TextMatrix(.Rows - 1, 11) = Format(nUIT * 7, "#0.00")
                .TextMatrix(.Rows - 1, 12) = Format(.TextMatrix(.Rows - 1, 10) - (nUIT * 7), "#0.00")
                .TextMatrix(.Rows - 1, 13) = Format(.TextMatrix(.Rows - 1, 12) * nCatV, "#0.00")
                Monto = Format(.TextMatrix(.Rows - 1, 13) / 12, "#0.00")
                .TextMatrix(.Rows - 1, 14) = IIf(Monto > 0, Monto, 0)
             Me.PrgBar.value = rs.Bookmark
            End With
        End If
        rs.MoveNext
    Wend
    Me.PrgBar.Visible = False
End If

End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.CmdBuscar.SetFocus
End If
End Sub

Sub CargaMeses()
Me.cboMes.AddItem "ENERO" & Space(100) & "01"
Me.cboMes.AddItem "FEBRERO" & Space(100) & "02"
Me.cboMes.AddItem "MARZO" & Space(100) & "03"
Me.cboMes.AddItem "ABRIL" & Space(100) & "04"
Me.cboMes.AddItem "MAYO" & Space(100) & "05"
Me.cboMes.AddItem "JUNIO" & Space(100) & "06"
Me.cboMes.AddItem "JULIO" & Space(100) & "07"
Me.cboMes.AddItem "AGOSTO" & Space(100) & "08"
Me.cboMes.AddItem "SETIEMBRE" & Space(100) & "09"
Me.cboMes.AddItem "OCTUBRE" & Space(100) & "10"
Me.cboMes.AddItem "NOVIEMBRE" & Space(100) & "11"
Me.cboMes.AddItem "DICIEMBRE" & Space(100) & "12"

End Sub

Private Sub txtFechaVac_GotFocus()

    txtFechaVac.SelStart = 0
    txtFechaVac.SelLength = 50
End Sub


Private Sub txtFechaVac_KeyPress(KeyAscii As Integer)
Dim RHVAC As DRHVacaciones
Dim rs As ADODB.Recordset

 If KeyAscii = 13 Then
    If Not ValFecha(Me.txtFechaVac) Then
        Me.txtFechaVac.SetFocus
        Exit Sub
    End If
    Me.Label3 = "VACACIONES EJECUTADAS " & Format(Me.txtFechaVac, "MMMM")
    Me.FlexVac.Rows = 2
    Me.FlexVac.Clear
    Me.FlexVac.FormaCabecera
    Set RHVAC = New DRHVacaciones
    Set rs = RHVAC.GetVacacionesEjecutadas(Format(Me.txtFechaVac, "YYYYMM"))
    If Not (rs.EOF And rs.BOF) Then
        Me.PrgBar.Min = 0
        Me.PrgBar.Max = rs.RecordCount
        Me.PrgBar.Visible = True
    End If
    While Not rs.EOF
        Me.FlexVac.AdicionaFila
        Me.FlexVac.TextMatrix(FlexVac.Rows - 1, 1) = rs!codigo
        Me.FlexVac.TextMatrix(FlexVac.Rows - 1, 2) = rs!cPersNombre
        Me.FlexVac.TextMatrix(FlexVac.Rows - 1, 3) = rs!Dias_Vac
        Me.FlexVac.TextMatrix(FlexVac.Rows - 1, 4) = rs!Quinta
        Me.FlexVac.TextMatrix(FlexVac.Rows - 1, 5) = rs!Total
        Me.PrgBar.value = rs.Bookmark
        rs.MoveNext
    Wend
    If Not (rs.EOF And rs.BOF) Then Me.PrgBar.Visible = False
    Set rs = Nothing
    Set RHVAC = Nothing
End If

End Sub

'*****************************************************************************
'***********************Registro Base PDT
'******************************************************************************

Private Sub GuardarBasePDTProviciones()
On Error GoTo GuardarBasePDTErr
    Dim oCon  As DConecta, sSQL As String, i As Integer
    If MsgBox("¿Seguro que Desea Registrar los Datos para la Base del PDT?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        i = 1
        With FlexProv
            Do While i < .Rows
                
                sSQL = " declare @tmp int "
                sSQL = sSQL & " set @tmp=(select count(*) from RHBasePDTDetalle where cTipo='P' and cRHCod = '" & .TextMatrix(i, 1) & "' and cPeriodo='" & Format(txtFechaVac, "yyyymm") & "')"
                sSQL = sSQL & " if @tmp=0 "
                sSQL = sSQL & "insert into RHBasePDTDetalle(cRHCod,cPeriodo,nRemuneracion,nImpuesto,Dias,cTipo) " & _
                       " values('" & .TextMatrix(i, 1) & "','" & Format(.TextMatrix(i, 4), "yyyymm") & "'," & .TextMatrix(i, 7) & "," & .TextMatrix(i, 14) & "," & Format(.TextMatrix(i, 5), "###,###") & ",'P') "
                oCon.Ejecutar sSQL
                
                sSQL = " declare @tmp int, @Periodo varchar(6), @cod varchar(6) "
                sSQL = sSQL & " set @Periodo    = '" & Format(.TextMatrix(i, 4), "yyyymm") & "' "
                sSQL = sSQL & " set @cod        = '" & .TextMatrix(i, 1) & "' "
                sSQL = sSQL & " set @tmp=(select count(*) from RHBasePDT where cRHCod = @cod) " & _
                              " if @tmp=0 " & _
                              "     insert into RHBasePDT (cRHCod,cPeriodo,nRemuneracion,nImpuesto,nDia) " & _
                              "     values(@cod, @Periodo, 0, 0,0) "
                oCon.Ejecutar sSQL
                i = i + 1
            Loop
        End With
        oCon.CierraConexion
    End If
    MsgBox "Registro Completo", vbInformation, "Aviso"
    Exit Sub
GuardarBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub GuardarBasePDTVacacionesE()
On Error GoTo GuardarBasePDTErr
    Dim oCon  As DConecta, sSQL As String, i As Integer
    If MsgBox("¿Seguro que Desea Registrar los Datos para la Base del PDT?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        i = 1
        With FlexVac
            Do While i < .Rows
                sSQL = " declare @tmp int "
                sSQL = sSQL & " set @tmp=(select count(*) from RHBasePDTDetalle where cTipo='V' and cRHCod = '" & .TextMatrix(i, 1) & "' and cPeriodo='" & Format(txtFechaVac, "yyyymm") & "')"
                sSQL = sSQL & " if @tmp=0 "
                sSQL = sSQL & "insert into RHBasePDTDetalle(cRHCod,cPeriodo,nRemuneracion,nImpuesto,Dias,cTipo) " & _
                       " values('" & .TextMatrix(i, 1) & "','" & Format(txtFechaVac, "yyyymm") & "'," & Val(.TextMatrix(i, 5)) & "," & Val(.TextMatrix(i, 4)) & "," & Format(.TextMatrix(i, 3), "###,##0") & ",'V') "
                oCon.Ejecutar sSQL
                
                sSQL = " declare @tmp int, @Periodo varchar(6), @cod varchar(6) "
                sSQL = sSQL & " set @Periodo    = '" & Format(txtFechaVac, "yyyymm") & "' "
                sSQL = sSQL & " set @cod        = '" & .TextMatrix(i, 1) & "' "
                sSQL = sSQL & " set @tmp=(select count(*) from RHBasePDT where cRHCod = @cod) " & _
                              " if @tmp=0 " & _
                              "     insert into RHBasePDT (cRHCod,cPeriodo,nRemuneracion,nImpuesto,nDia) " & _
                              "     values(@cod, @Periodo, 0, 0 ,0) "
                oCon.Ejecutar sSQL
                i = i + 1
            Loop
        End With
        oCon.CierraConexion
    End If
    MsgBox "Registro Completo", vbInformation, "Aviso"
    Exit Sub
GuardarBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub
