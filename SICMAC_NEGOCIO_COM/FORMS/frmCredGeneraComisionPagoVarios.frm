VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredGeneraComisionPagoVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Convenio Pago Varios"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "frmCredGeneraComisionPagoVarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTDatosGen 
      Height          =   4095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   -2147483648
      TabCaption(0)   =   "Registro Convenio"
      TabPicture(0)   =   "frmCredGeneraComisionPagoVarios.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSimbolo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFecVig"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MontoSol"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdCancelar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Cmdgrabar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtdescripcion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboInstitucion"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboagencia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmdNuevo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboTipoConvenio"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fraCuenta"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Check2(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check2(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "&Mantenimiento Convenio"
      TabPicture(1)   =   "frmCredGeneraComisionPagoVarios.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(2)=   "Label19"
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(4)=   "Label21"
      Tab(1).Control(5)=   "Label22"
      Tab(1).Control(6)=   "Label26"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "txtFecVig1"
      Tab(1).Control(9)=   "MontoSol1"
      Tab(1).Control(10)=   "CmdCancelarMant"
      Tab(1).Control(11)=   "CmdEditar"
      Tab(1).Control(12)=   "cboInstitucion1"
      Tab(1).Control(13)=   "cboagencia1"
      Tab(1).Control(14)=   "txtdescripcion1"
      Tab(1).Control(15)=   "cboTipoConvenioMant"
      Tab(1).Control(16)=   "Frame1"
      Tab(1).Control(17)=   "Check2(2)"
      Tab(1).Control(18)=   "Check2(3)"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Listado Convenio"
      TabPicture(2)   =   "frmCredGeneraComisionPagoVarios.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label23"
      Tab(2).Control(1)=   "Label27"
      Tab(2).Control(2)=   "FEInstitucion"
      Tab(2).Control(3)=   "cmdImprimir"
      Tab(2).Control(4)=   "cboagencia2"
      Tab(2).Control(5)=   "ChKAgencia"
      Tab(2).Control(6)=   "CboTipoReporte"
      Tab(2).ControlCount=   7
      Begin VB.CheckBox Check2 
         Caption         =   "Identificador"
         Height          =   255
         Index           =   3
         Left            =   -72480
         TabIndex        =   76
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   2
         Left            =   -73800
         TabIndex        =   75
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Identificador"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   73
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   72
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cuenta"
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   68
         Top             =   1920
         Width           =   6720
         Begin SICMACT.ActXCodCta txtcuenta1 
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   661
            Texto           =   "Cuenta N°:"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblcuenta1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   555
            Left            =   3960
            TabIndex        =   70
            Top             =   150
            Width           =   2520
         End
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Cuenta"
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
         Height          =   735
         Left            =   120
         TabIndex        =   66
         Top             =   1920
         Width           =   6720
         Begin SICMACT.ActXCodCta txtCuenta 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   661
            Texto           =   "Cuenta N°:"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblcuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   555
            Left            =   3960
            TabIndex        =   67
            Top             =   150
            Width           =   2520
         End
      End
      Begin VB.ComboBox CboTipoReporte 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cboTipoConvenioMant 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox cboTipoConvenio 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCredGeneraComisionPagoVarios.frx":035E
         Left            =   1200
         List            =   "frmCredGeneraComisionPagoVarios.frx":0360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox ChKAgencia 
         Caption         =   "Todas las Agencias"
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
         TabIndex        =   61
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo"
         Height          =   300
         Left            =   2760
         TabIndex        =   9
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox cboagencia2 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -71520
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtdescripcion1 
         Enabled         =   0   'False
         Height          =   525
         Left            =   -73800
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   3120
         Width           =   5415
      End
      Begin VB.ComboBox cboagencia1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox cboInstitucion1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69600
         TabIndex        =   34
         ToolTipText     =   "Eliminar Comites"
         Top             =   3120
         Width           =   1020
      End
      Begin VB.ComboBox cboagencia 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -69480
         TabIndex        =   33
         Top             =   1020
         Width           =   1020
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.ListBox LstAnalista 
         Enabled         =   0   'False
         Height          =   2985
         Left            =   -74880
         Style           =   1  'Checkbox
         TabIndex        =   31
         Top             =   900
         Width           =   5265
      End
      Begin VB.ComboBox cboInstitucion 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox txtPersDireccDomicilio 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73740
         MaxLength       =   100
         TabIndex        =   30
         Top             =   2190
         Width           =   5200
      End
      Begin VB.ComboBox cmbPersDireccCondicion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73740
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2550
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ubicacion Geografica"
         Height          =   1665
         Left            =   -74865
         TabIndex        =   18
         Top             =   420
         Width           =   7365
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmCredGeneraComisionPagoVarios.frx":0362
            Left            =   2235
            List            =   "frmCredGeneraComisionPagoVarios.frx":0364
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1140
            Width           =   2190
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmCredGeneraComisionPagoVarios.frx":0366
            Left            =   4680
            List            =   "frmCredGeneraComisionPagoVarios.frx":0368
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "frmCredGeneraComisionPagoVarios.frx":036A
            Left            =   4680
            List            =   "frmCredGeneraComisionPagoVarios.frx":036C
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmCredGeneraComisionPagoVarios.frx":036E
            Left            =   2250
            List            =   "frmCredGeneraComisionPagoVarios.frx":0370
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmCredGeneraComisionPagoVarios.frx":0372
            Left            =   210
            List            =   "frmCredGeneraComisionPagoVarios.frx":0374
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2235
            TabIndex        =   28
            Top             =   900
            Width           =   600
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4680
            TabIndex        =   27
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2265
            TabIndex        =   25
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4695
            TabIndex        =   24
            Top             =   285
            Width           =   750
         End
      End
      Begin VB.TextBox txtdescripcion 
         Enabled         =   0   'False
         Height          =   525
         Left            =   1200
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3120
         Width           =   5415
      End
      Begin VB.CommandButton Cmdgrabar 
         Caption         =   "Grabar"
         Height          =   300
         Left            =   4080
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   5400
         TabIndex        =   10
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Grabar"
         Height          =   300
         Left            =   -70680
         TabIndex        =   17
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelarMant 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -69360
         TabIndex        =   16
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69360
         TabIndex        =   15
         ToolTipText     =   "Eliminar Comites"
         Top             =   3120
         Width           =   1020
      End
      Begin SICMACT.EditMoney MontoSol 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshComite 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   35
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSMask.MaskEdBox txtFecVig 
         Height          =   300
         Left            =   5400
         TabIndex        =   4
         Top             =   1200
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.EditMoney MontoSol1 
         Height          =   255
         Left            =   -73800
         TabIndex        =   49
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox txtFecVig1 
         Height          =   300
         Left            =   -69840
         TabIndex        =   50
         Top             =   1200
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.FlexEdit FEInstitucion 
         Height          =   2025
         Left            =   -74880
         TabIndex        =   60
         Top             =   960
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3572
         Cols0           =   4
         FixedCols       =   0
         HighLight       =   1
         EncabezadosNombres=   "#-Codigo-Descripcion-Vigencia"
         EncabezadosAnchos=   "400-1800-3000-1200"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-C"
         FormatosEdit    =   "0-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.Label Label14 
         Caption         =   "Buscar Por:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   74
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Buscar Por:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   65
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   63
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70560
         TabIndex        =   59
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "F. Vigencia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69840
         TabIndex        =   56
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label21 
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   55
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Institucion:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   54
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72360
         TabIndex        =   53
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Comis. Sol:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74880
         TabIndex        =   52
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label16 
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   51
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblmensaje 
         Caption         =   "Para eliminar escoja Agencia en la pestaña Datos Grales, y escoja de la lista un elemento..."
         Height          =   255
         Left            =   -74880
         TabIndex        =   48
         Top             =   3720
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
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
         Left            =   -74460
         TabIndex        =   47
         Top             =   2820
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
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
         Left            =   -74880
         TabIndex        =   46
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Institucion:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblPersDireccDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   43
         Top             =   2220
         Width           =   630
      End
      Begin VB.Label lblPersDireccCondicion 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Left            =   -74820
         TabIndex        =   42
         Top             =   2625
         Width           =   705
      End
      Begin VB.Label Label13 
         Caption         =   "Valor Comercial U$"
         Height          =   240
         Left            =   -71580
         TabIndex        =   41
         Top             =   2625
         Width           =   1440
      End
      Begin VB.Label lblRefDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Left            =   -74820
         TabIndex        =   40
         Top             =   3030
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "F. Vigencia:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4320
         TabIndex        =   39
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label lblSimbolo 
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         TabIndex        =   38
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Comision:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label15 
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCredGeneraComisionPagoVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bnovalidaCta As Boolean
Dim bnovalidaCta1 As Boolean
Dim objPista As COMManejador.Pista

Sub ConfigurarMShComite()
 MshComite.Clear
    MshComite.Cols = 3
    MshComite.Rows = 2

    With MshComite
        .TextMatrix(0, 0) = "Codigo"
        .TextMatrix(0, 1) = "Descripcion"
        .TextMatrix(0, 2) = "Vigencia"

        .ColWidth(1) = 2500
        .ColWidth(2) = 3000
    End With
End Sub

Private Sub CargaAgencia(Combo As ComboBox)
Dim loCargaAg As COMDColocPig.DCOMColPFunciones
Dim lrAgenc As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    Call llenar_cbo_agencia(lrAgenc, Combo)
    Exit Sub

ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaTipoConvenio(Combo As ComboBox)
Dim oCon As COMDConstantes.DCOMConstantes
Dim rsT As ADODB.Recordset

    Set oCon = New COMDConstantes.DCOMConstantes
    Set rsT = oCon.RecuperaConstantes(9090)
    Set oCon = Nothing
    Call llenar_cbo_tipoConvenio(rsT, Combo)
    Exit Sub

ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cboagencia1_Click()
    If cboagencia1.ListIndex <> -1 Then
        CargaCboConvenioporAgencia
    End If
End Sub

Private Sub cboagencia2_Click()
    Call CargaFEConvenioporAgencia
    cmdImprimir.Enabled = True
End Sub

Private Sub cboInstitucion1_Click()
   If cboInstitucion1.ListIndex <> -1 Then
        habilita_Controles1 True
        CargaDatosInstitucion2
   End If
End Sub

Private Sub cboTipoConvenioMant_Click()
    If cboTipoConvenioMant.ListIndex <> -1 Then
        CargaCboAgenciaporConvenio 1
    End If
End Sub

Private Sub CboTipoReporte_Click()
    If CboTipoReporte.ListIndex <> -1 Then
        CargaCboAgenciaporConvenio 2
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call limpia_Controles
    habilita_Controles False
    Me.cmdgrabar.Enabled = False
    Me.cmdNuevo.Enabled = True
    Me.cmdNuevo.SetFocus
End Sub

Private Sub CmdCancelarMant_Click()
    Call Limpia_Controles1
    habilita_Controles1 False
    Me.cboTipoConvenioMant.Enabled = True
    Me.cboTipoConvenioMant.SetFocus
End Sub

Private Sub cmdEditar_Click()
Dim oCredD As COMDCredito.DCOMCredito
Set oCredD = New COMDCredito.DCOMCredito

        If Me.lblcuenta1.Caption = "" Then
            MsgBox "Debe validar la cuenta ingresada para este convenio, Haga Enter sobre la cuenta", vbCritical, "Aviso"
            Exit Sub
        End If

        If bnovalidaCta1 = False Then
            MsgBox "La cuenta ingresada no es válida para este convenio", vbCritical, "Aviso"
            Exit Sub
        End If

        If Not valida_Controles1 Then
            Exit Sub
        End If

         If (Me.cboagencia1.Visible) Then
             nValAge = CInt(Right(Me.cboagencia1.Text, 2))
         Else
             nValAge = 0
         End If

        If CInt(Right(Me.cboTipoConvenioMant.Text, 2)) = 1 Then
             If Not oCredD.GetValConvenioInstitucionxAgenciaTipo(Trim(Right(Me.cboInstitucion1.Text, 13))) Then
                MsgBox "La Institución No Cuenta con un Convenio Total, Verifique!!", vbCritical
                Exit Sub
             End If
        Else
            If Not oCredD.GetValConvenioInstitucionxAgencia(Trim(Right(Me.cboInstitucion1.Text, 13)), nValAge) Then
                MsgBox "La Institución No Cuenta con un Convenio en esta Agencia, Verifique!!", vbCritical
                Exit Sub
            End If
        End If

        If MsgBox("Desea Grabar la Operacion??", vbQuestion + vbYesNo, "Aviso") = vbYes Then

            'MADM 20110601 - Parametros Busqueda
            If (Me.Check2(2).value = 1 And Me.Check2(3).value = 1) Then
                nValBusca = 2
            ElseIf (Me.Check2(3).value = 1) Then
                nValBusca = 1
            Else
                nValBusca = 0
            End If

                oCredD.ActualizarDatosPagoConvenioInstyAgencia Trim(Right(Me.cboInstitucion1.Text, 13)), nValAge, CDate(Me.txtFecVig1.Text), CInt(Right(Me.cboTipoConvenioMant, 2)), Trim(Me.txtdescripcion1.Text), CDbl(Me.MontoSol1.value), txtcuenta1.NroCuenta, nValBusca
                objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , txtCuenta.NroCuenta, gCodigoCuenta

                Set oCredD = Nothing
                Call Limpia_Controles1
                Call CargaAgencia(cboagencia1)
                Call CargaTipoConvenio(cboTipoConvenioMant)
                Call CargaDatosInstitucion2
                cmdeditar.Enabled = True
                cmdeditar.SetFocus
        End If
End Sub

Private Sub CmdGrabar_Click()
Dim oCredD As COMDCredito.DCOMCredito
Dim nValAge As Integer
Set oCredD = New COMDCredito.DCOMCredito

        If Me.lblcuenta.Caption = "" Then
            MsgBox "Debe validar la cuenta ingresada para este convenio", vbCritical, "Aviso"
            Exit Sub
        End If

        If bnovalidaCta = False Then
            MsgBox "La cuenta ingresada no es válida para este convenio", vbCritical, "Aviso"
            Exit Sub
        End If

        If Not valida_controles Then
            Exit Sub
        End If

        If oCredD.GetValConvenioInstitucionxAgenciaTipo(Trim(Right(Me.cboInstitucion.Text, 13))) Then
                MsgBox "La Institución Cuenta con un Convenio Total, seleccione Mantenimiento de Convenio !!", vbCritical
                Exit Sub
        End If

        If oCredD.GetValConvenioInstitucionxAgencia(Trim(Right(Me.cboInstitucion.Text, 13)), CInt(Right(Me.cboagencia.Text, 2))) Then
                MsgBox "La Institución Cuenta con un Convenio Tipo Agencia, seleccione Mantenimiento de Convenio !!", vbCritical
                Exit Sub
        End If

        If Me.Check2(0).value = False And Me.Check2(1).value = False Then
                MsgBox "Debe definir el Tipo de Búsqueda, seleccione Mantenimiento de Convenio !!", vbCritical
                Exit Sub
        End If

        If MsgBox("Desea Grabar la Operacion??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim nValBusca As Integer
            nValBusca = 0

            If (Me.cboagencia.Visible) Then
                nValAge = CInt(Right(Me.cboagencia.Text, 2))
            Else
                nValAge = 0
            End If

            'MADM 20110601 - Parametros Busqueda
            If (Me.Check2(0).value = 1 And Me.Check2(1).value = 1) Then
                nValBusca = 2
            ElseIf (Me.Check2(1).value = 1) Then
                nValBusca = 1
            Else
                nValBusca = 0
            End If

            oCredD.InsertarDatosPagoConvenioInstyAgencia Trim(Right(Me.cboInstitucion.Text, 13)), nValAge, gdFecSis, Me.txtFecVig.Text, CInt(Right(Me.cboTipoConvenio.Text, 2)), Me.txtdescripcion, Me.MontoSol.value, txtCuenta.NroCuenta, nValBusca
            Set oCredD = Nothing
    
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , txtCuenta.NroCuenta, gCodigoCuenta


            Call limpia_Controles
            habilita_Controles False
            Call CargaAgencia(cboagencia)
            Call CargaTipoConvenio(cboTipoConvenio)
            Call CargaInstitucion
            cmdNuevo.Enabled = True
            cmdNuevo.SetFocus
            Me.cmdgrabar.Enabled = False
        End If
End Sub
Public Function valida_controles() As Boolean
    Dim nMontoSol As Double
    valida_controles = True

    nMontoSol = MontoSol.value

    If Me.cboagencia.ListIndex = -1 Or Me.cboInstitucion.ListIndex = -1 Then
        MsgBox "Debe seleccionar Agencia / Institución Válida", vbInformation, "Aviso"
        valida_controles = False
        Exit Function
    End If

    If Len(Me.txtdescripcion.Text) = 0 Then
        MsgBox "Debe ingresar descripcion del Convenio", vbInformation, "Aviso"
        valida_controles = False
        Exit Function
    End If

    If Not (IsDate(Me.txtFecVig)) Then
        MsgBox "Debe ingresar una Fecha de Vigencia Correcta", vbInformation, "Aviso"
        valida_controles = False
        Exit Function
    End If

    If nMontoSol = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        valida_controles = False
        If MontoSol.Enabled Then MontoSol.SetFocus
        Exit Function
    Exit Function
End If

End Function
Public Function valida_Controles1() As Boolean
    Dim nMontoSol1 As Double

    valida_Controles1 = True
    nMontoSol1 = MontoSol1.value

    If Me.cboagencia1.ListIndex = -1 Or Me.cboInstitucion1.ListIndex = -1 Then
        MsgBox "Debe seleccionar Agencia / Institución Válida", vbInformation, "Aviso"
        valida_Controles1 = False
    End If

    If Len(Me.txtdescripcion1.Text) = 0 Then
        MsgBox "Debe ingresar descripcion del Convenio", vbInformation, "Aviso"
        valida_Controles1 = False
    End If

    If Not (IsDate(Me.txtFecVig1)) Then
        MsgBox "Debe ingresar una Fecha de Vigencia Correcta", vbInformation, "Aviso"
        valida_Controles1 = False
    End If

    If nMontoSol1 = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        If MontoSol1.Enabled Then MontoSol1.SetFocus
        valida_Controles1 = False
    Exit Function
End If

End Function

Private Sub cmdImprimir_Click()
Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oNCred As COMNCredito.NCOMCredito

    Set oPrev = New previo.clsprevio

    Set oNCred = New COMNCredito.NCOMCredito

    If Me.CboTipoReporte.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de convenio, Verifique !!", vbInformation, "Aviso"
        Exit Sub
    End If

    sCadImp = oNCred.ImprimeReporteConvenioPagoVarios(gsCodUser, gdFecSis, IIf(Trim(Right(Me.cboagencia2.Text, 2)) = "", 1, Trim(Right(Me.cboagencia2.Text, 2))), gsNomCmac, CInt(Right(Me.CboTipoReporte.Text, 2)), IIf(Me.ChKAgencia.value, 1, 0))

    previo.Show sCadImp, "Registro de Archivo Convenio de Servicios Registrados", False
    Set oPrev = Nothing
    Set oNCred = Nothing
End Sub

Private Sub cmdNuevo_Click()
    habilita_Controles True
    MontoSol.BackColor = &HC0FFFF
    lblSimbolo.Caption = "S/."
    Me.cmdgrabar.Enabled = True
    Me.cmdNuevo.Enabled = False
    bnovalidaCta = False
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 CentraForm Me
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    bnovalidaCta = False
    bnovalidaCta1 = False
    Me.txtCuenta.CMAC = "109"
    Me.txtcuenta1.CMAC = "109"
    habilita_Controles False
    Call CargaTipoConvenio(cboTipoConvenio)
    Call CargaTipoConvenio(cboTipoConvenioMant)
    Call CargaTipoConvenio(CboTipoReporte)
    Call CargaInstitucion
    Call CargaAgencia(cboagencia)
'    Call CargaAgencia(cboagencia1)
'    Call CargaAgencia(cboagencia2)
    ConfigurarMShComite
    Set objPista = New COMManejador.Pista
    gsOpeCod = 100917
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Sub habilita_Controles(pbloqueo As Boolean)
    Me.cboTipoConvenio.Enabled = pbloqueo
    Me.cboagencia.Enabled = pbloqueo
    Me.cboInstitucion.Enabled = pbloqueo
    Me.txtdescripcion.Enabled = pbloqueo
    Me.txtFecVig.Enabled = pbloqueo
    Me.MontoSol.Enabled = pbloqueo
    Me.txtCuenta.Enabled = pbloqueo
    Me.Check2(0).Enabled = pbloqueo
    Me.Check2(1).Enabled = pbloqueo
End Sub

Sub habilita_Controles1(pbloqueo As Boolean)
    Me.txtdescripcion1.Enabled = pbloqueo
    Me.txtFecVig1.Enabled = pbloqueo
    Me.MontoSol1.Enabled = pbloqueo
    Me.txtcuenta1.Enabled = pbloqueo
    Me.Check2(2).Enabled = pbloqueo
    Me.Check2(3).Enabled = pbloqueo
End Sub

Sub Limpia_Controles1()
    Me.cboTipoConvenioMant.ListIndex = -1
    Me.cboagencia1.ListIndex = -1
    Me.cboInstitucion1.ListIndex = -1
    Me.txtdescripcion1.Text = ""
    Me.txtFecVig1.Text = "__/__/____"
    Me.txtcuenta1.Age = ""
    Me.txtcuenta1.Cuenta = ""
    Me.txtcuenta1.Prod = ""
    bnovalidaCta1 = False
    lblcuenta1.Caption = ""
    MontoSol1.value = 0
    Me.Check2(2).value = 0
    Me.Check2(3).value = 0
End Sub

Sub limpia_Controles()
    Me.cboTipoConvenio.ListIndex = -1
    Me.cboagencia.ListIndex = -1
    Me.cboInstitucion.ListIndex = -1
    Me.txtdescripcion.Text = ""
    Me.txtFecVig.Text = "__/__/____"
    Me.MontoSol.Text = 0
    Me.txtCuenta.Age = ""
    Me.txtCuenta.Cuenta = ""
    Me.txtCuenta.Prod = ""
    bnovalidaCta = False
    lblcuenta.Caption = ""
    MontoSol.BackColor = &HC0FFFF
    lblSimbolo.Caption = "S/."
    Me.Check2(0).value = False
    Me.Check2(1).value = False
End Sub

Sub CargaFEConvenioporAgencia()
Dim rsCred As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito

    Set oCredD = New COMDCredito.DCOMCredito
    Set rsCred = New ADODB.Recordset
    Set rsCred = oCredD.GetDevuelveConvenioAgenciayTipo(CInt(Right(Me.cboagencia2.Text, 2)), CInt(Right(Me.CboTipoReporte.Text, 2)))

    Set oCredD = Nothing

    FEInstitucion.Clear
    FEInstitucion.Rows = 2
    FEInstitucion.FormaCabecera
    FEInstitucion.FormateaColumnas
    FEInstitucion.TextMatrix(1, 0) = "1"

    If Not (rsCred.EOF And rsCred.BOF) Then
        Call LimpiaFlex(FEInstitucion)
        Do While Not rsCred.EOF
            FEInstitucion.AdicionaFila
            FEInstitucion.TextMatrix(FEInstitucion.Row, 1) = rsCred!cPersCod
            FEInstitucion.TextMatrix(FEInstitucion.Row, 2) = rsCred!cPersNombre
            FEInstitucion.TextMatrix(FEInstitucion.Row, 3) = rsCred!fVigencia
            rsCred.MoveNext
        Loop

        rsCred.Close
        Set rsCred = Nothing

    Else
        MsgBox "No se encontraron Convenio de Pago de Servicios con esta Agencia", vbOKOnly + vbExclamation, "Atención"
        Set rsCred = Nothing
    End If
End Sub

Sub CargaCboConvenioporAgencia()
Dim rsCred As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito

    Set oCredD = New COMDCredito.DCOMCredito
    Set rsCred = New ADODB.Recordset
    Set rsCred = oCredD.GetDevuelveConvenioAgenciayTipo(IIf(Trim(Right(Me.cboagencia1.Text, 4)) <> "", Trim(Right(Me.cboagencia1.Text, 4)), gsCodAge), CInt(Right(Me.cboTipoConvenioMant.Text, 2)))

    Call llenar_cbo(rsCred, Me.cboInstitucion1)
    Me.txtdescripcion1.Text = ""
    Me.txtFecVig1.Text = "__/__/____"
    Me.MontoSol1.Text = 0
    Me.txtcuenta1.Texto = ""
    lblcuenta1.Caption = ""

    Set oCredD = Nothing
    Set rsCred = Nothing
    Exit Sub
End Sub

Sub CargaCboAgenciaporConvenio(ByVal pind As Integer)
Dim rsCred As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito
Dim pValor As Integer
Dim pCombo As ComboBox
    Set oCredD = New COMDCredito.DCOMCredito
    Set rsCred = New ADODB.Recordset

    '''''''''''''''''''''''''''''''''''''''''''''''''
    If pind = 1 Then
        If Me.cboTipoConvenioMant.ListIndex = -1 Then
           Exit Sub
        End If
        pValor = CInt(Right(Me.cboTipoConvenioMant.Text, 2))
    Else
        If Me.CboTipoReporte.ListIndex = -1 Then
           Exit Sub
        End If
        pValor = CInt(Right(Me.CboTipoReporte.Text, 2))
    End If
     Set rsCred = oCredD.GetDevuelveAgenciaxTipoConvenio(pValor)
    Call llenar_cbo_agencia(rsCred, IIf(pind = 1, Me.cboagencia1, Me.cboagencia2))

    Set oCredD = Nothing
    Set rsCred = Nothing
    Exit Sub
End Sub

Sub CargaCboAgenciaporConvenioReporte()
Dim rsCred As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito

    Set oCredD = New COMDCredito.DCOMCredito
    Set rsCred = New ADODB.Recordset
    Set rsCred = oCredD.GetDevuelveAgenciaxTipoConvenio(CInt(Right(Me.CboTipoReporte.Text, 2)))

    Call llenar_cbo_agencia(rsCred, Me.cboagencia2)

    Set oCredD = Nothing
    Set rsCred = Nothing
    Exit Sub
End Sub
Private Sub CargaInstitucion()
Dim rs As ADODB.Recordset
Dim oGen  As COMDConstSistema.DCOMGeneral
    On Error GoTo ERRORCargaInstitucion

    Set oGen = New COMDConstSistema.DCOMGeneral
    Set rs = New ADODB.Recordset
    Set rs = oGen.CargaInstituciones()

    Call llenar_cbo(rs, Me.cboInstitucion)

    Set oGen = Nothing
    Set rs = Nothing
    Exit Sub
ERRORCargaInstitucion:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaDatosInstitucion2()
Dim rs As ADODB.Recordset
Dim oCredD As COMDCredito.DCOMCredito
Dim sCuenta As String

    On Error GoTo ERRORCargaInstitucion
    sCuenta = ""
    If Me.cboTipoConvenioMant.ListIndex = -1 Or Me.cboInstitucion1.ListIndex = -1 Then
       Exit Sub
    End If

    Set oCredD = New COMDCredito.DCOMCredito
    Set rs = New ADODB.Recordset
    Set rs = oCredD.GetDevuelveConvenioInstitucionyAgenciaytipo(Trim(Right(Me.cboInstitucion1.Text, 13)), CInt(Right(Me.cboagencia1.Text, 2)), CInt(Right(Me.cboTipoConvenioMant.Text, 2)))

    If Not (rs.EOF And rs.BOF) Then
        Me.txtdescripcion1.Text = Trim(rs!cObsConv)
        Me.txtFecVig1.Text = CDate(rs!fVigencia)
        Me.MontoSol1.Text = CDbl(rs!nImporteComision)
        sCuenta = rs!cCuenta
        Me.txtcuenta1.CMAC = 109
        Me.txtcuenta1.Age = Mid(sCuenta, 4, 2)
        Me.txtcuenta1.Prod = Mid(sCuenta, 6, 3)
        Me.txtcuenta1.Cuenta = Mid(sCuenta, 9, 18)

        'MADM 20110601
         If CInt(rs!nTipoBus) = 2 Then
            Me.Check2(2).value = 1
            Me.Check2(3).value = 1
          ElseIf CInt(rs!nTipoBus) = 1 Then
            Me.Check2(2).value = 0
            Me.Check2(3).value = 1
          Else
              Me.Check2(2).value = 1
              Me.Check2(3).value = 0
          End If
        'END MADM

    rs.Close
    Set rs = Nothing
    End If

    Set oCredD = Nothing
    Exit Sub
ERRORCargaInstitucion:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub


Sub llenar_cbo(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(str(pRs!cPersCod))
    pRs.MoveNext
Loop
pRs.Close
End Sub

Sub llenar_cbo_agencia(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
Dim vage As String
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cAgeDescripcion) & Space(100) & Trim(str(pRs!cAgeCod))
    pRs.MoveNext
Loop
pRs.Close
End Sub

Sub llenar_cbo_tipoConvenio(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim(str(pRs!nConsValor))
    pRs.MoveNext
Loop
pRs.Close
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    Dim pnProducto As String
    sCta = txtCuenta.NroCuenta
    pnProducto = txtCuenta.Prod

    If EsHaberes(sCta) Then
       MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
       Exit Sub
    End If

   ObtieneDatosCuenta sCta, pnProducto
End If
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String, ByVal pnProducto As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim rsCta As New ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim lnTpoPrograma As Integer

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
sMsg = clsCap.ValidaCuentaOperacion(sCuenta, True)
Set clsCap = Nothing

If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsCta = New Recordset
        Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then

        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")

        If pnProducto = gCapAhorros Then
            lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        End If

        If lnTpoPrograma = 0 Then
            lblcuenta = ""
            bnovalidaCta = True
            MsgBox "La cuenta ingresada no se puede utilizar, para este convenio", vbCritical, "Aviso"
            Exit Sub
        End If

        nmoneda = CLng(Mid(sCuenta, 9, 1))

        If nmoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            MontoSol.BackColor = &HC0FFFF
            lblSimbolo.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            MontoSol.BackColor = &HC0FFC0
            lblSimbolo.Caption = "$"
        End If

        bnovalidaCta = True

        Select Case pnProducto
             Case gCapAhorros
                If rsCta("bOrdPag") Then
                    lblcuenta = lblmensaje & Chr$(13) & "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                    pbOrdPag = True
                Else
                    'AVMM 10-04-2007
                    If lnTpoPrograma = 1 Then
                        lblcuenta = lblcuenta & Chr$(13) & "AHORRO ÑAÑITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 2 Then
                        lblcuenta = lblcuenta & Chr$(13) & "AHORROS PANDERITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 3 Then
                        '*** PEAC 20090722
                        'lblMensaje = lblMensaje & Chr$(13) & "AHORROS PANDERO" & Chr$(13) & sMoneda
                        lblcuenta = lblcuenta & Chr$(13) & "AHORROS POCO A POCO AHORRO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 4 Then
                        lblcuenta = lblcuenta & Chr$(13) & "AHORROS DESTINO" & Chr$(13) & sMoneda
                    Else
                        lblcuenta = lblcuenta & Chr$(13) & "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                    End If
                    pbOrdPag = False
                End If

            Case gCapCTS
                lblcuenta = lblcuenta & Chr$(13) & "CTS" & Chr$(13) & sMoneda
        End Select
    End If
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
End Sub

Private Sub ObtieneDatosCuenta1(ByVal sCuenta As String, ByVal pnProducto As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Dim rsCta As New ADODB.Recordset, rsRel As New ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim lnTpoPrograma As Integer

Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
sMsg = clsCap.ValidaCuentaOperacion(sCuenta, True)
Set clsCap = Nothing

If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsCta = New Recordset
        Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then

        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")

        If pnProducto = gCapAhorros Then
            lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        End If

        'COMENTADO MADM 20120330
'        If lnTpoPrograma = 0 Then
'            lblcuenta1 = ""
'            MsgBox "La cuenta ingresada no se puede utilizar, para este convenio", vbCritical, "Aviso"
'            Exit Sub
'        End If

        nmoneda = CLng(Mid(sCuenta, 9, 1))

        If nmoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            MontoSol.BackColor = &HC0FFFF
            lblSimbolo.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            MontoSol.BackColor = &HC0FFC0
            lblSimbolo.Caption = "$"
        End If

        bnovalidaCta1 = True

        Select Case pnProducto
             Case gCapAhorros
                If rsCta("bOrdPag") Then
                    lblcuenta1 = lblcuenta1 & Chr$(13) & "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                    pbOrdPag = True
                Else
                    'AVMM 10-04-2007
                    If lnTpoPrograma = 1 Then
                        lblcuenta1 = lblcuenta1 & Chr$(13) & "AHORRO ÑAÑITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 2 Then
                        lblcuenta1 = lblcuenta1 & Chr$(13) & "AHORROS PANDERITO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 3 Then
                        '*** PEAC 20090722
                        'lblMensaje = lblMensaje & Chr$(13) & "AHORROS PANDERO" & Chr$(13) & sMoneda
                        lblcuenta1 = lblcuenta1 & Chr$(13) & "AHORROS POCO A POCO AHORRO" & Chr$(13) & sMoneda
                    ElseIf lnTpoPrograma = 4 Then
                        lblcuenta1 = lblcuenta1 & Chr$(13) & "AHORROS DESTINO" & Chr$(13) & sMoneda
                    Else
                        lblcuenta1 = lblcuenta1 & Chr$(13) & "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                    End If
                    pbOrdPag = False
                End If

            Case gCapCTS
                lblcuenta1 = lblcuenta1 & Chr$(13) & "CTS" & Chr$(13) & sMoneda
        End Select
    End If
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtcuenta1.SetFocus
End If
End Sub

Private Function EsHaberes(ByVal sCta As String) As Boolean
Dim cCap As COMDCaptaGenerales.COMDCaptAutorizacion
Set cCap = New COMDCaptaGenerales.COMDCaptAutorizacion
    EsHaberes = cCap.EsHaberes(sCta)
Set cCap = Nothing
End Function

Private Sub txtcuenta1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    Dim pnProducto As String
    sCta = txtcuenta1.NroCuenta
    pnProducto = txtcuenta1.Prod

    If EsHaberes(sCta) Then
       MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
       Exit Sub
    End If

   ObtieneDatosCuenta1 sCta, pnProducto
End If
End Sub
