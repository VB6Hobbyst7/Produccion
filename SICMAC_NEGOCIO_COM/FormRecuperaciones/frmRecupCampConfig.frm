VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecupCampConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Campañas de Recuperaciones"
   ClientHeight    =   5880
   ClientLeft      =   2970
   ClientTop       =   0
   ClientWidth     =   9375
   Icon            =   "frmRecupCampConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "Examinar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditar 
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
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin TabDlg.SSTab sstConfiguracion 
      Height          =   5175
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   741
      TabCaption(0)   =   "Datos de la Campaña"
      TabPicture(0)   =   "frmRecupCampConfig.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDatosCamp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAgencias"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraTipoCredito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Sub Campañas"
      TabPicture(1)   =   "frmRecupCampConfig.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraSubCampana"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraListado"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdQuitarSC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdEditarSC"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdNuevoSC"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdCancelarSC"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAceptarSC"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Niveles de Aprobación"
      TabPicture(2)   =   "frmRecupCampConfig.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "feNivelesApr"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Conf. Restricciòn"
      TabPicture(3)   =   "frmRecupCampConfig.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Exoneraciòn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74760
         TabIndex        =   52
         Top             =   2520
         Width           =   8655
         Begin VB.CommandButton cmdBuscarPer 
            Caption         =   "Buscar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6840
            TabIndex        =   65
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdCancExo 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6840
            TabIndex        =   64
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditExo 
            Caption         =   "Editar"
            Height          =   375
            Left            =   5040
            TabIndex        =   63
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtCantAdiDsctExo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   62
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtClienteExo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   61
            Top             =   570
            Width           =   5655
         End
         Begin VB.Label Label19 
            Caption         =   "Cant. Adicional Descuento:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   1125
            Width           =   2415
         End
         Begin VB.Label Label18 
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Asignaciòn Anual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74760
         TabIndex        =   51
         Top             =   600
         Width           =   8655
         Begin VB.CommandButton cmdCancAsigAnual 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6840
            TabIndex        =   58
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdEditAcept 
            Caption         =   "Editar"
            Height          =   375
            Left            =   6840
            TabIndex        =   57
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtCantMaxDesctAsigAnual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   56
            Top             =   930
            Width           =   2175
         End
         Begin VB.TextBox txtAnioAsigAnual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   55
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label17 
            Caption         =   "Cant. Max. Descuentos:"
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
            Left            =   240
            TabIndex        =   54
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label16 
            Caption         =   "Año:"
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
            Left            =   240
            TabIndex        =   53
            Top             =   520
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdAceptarSC 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7800
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelarSC 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7800
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevoSC 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   7800
         TabIndex        =   26
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditarSC 
         Caption         =   "Editar"
         Height          =   375
         Left            =   7800
         TabIndex        =   27
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuitarSC 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   7800
         TabIndex        =   28
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Frame fraListado 
         Caption         =   "Listado"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Width           =   7575
         Begin SICMACT.FlexEdit feSubCampanas 
            Height          =   2145
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   7290
            _ExtentX        =   12859
            _ExtentY        =   3784
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Días Atraso-Consideración-Cap-Int.-Mora-Gasto-Icv-Estado"
            EncabezadosAnchos=   "400-1000-1800-700-700-700-700-700-1000"
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
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0"
            BackColor       =   16777215
            EncabezadosAlineacion=   "C-C-L-R-C-C-C-R-C"
            FormatosEdit    =   "0-0-0-2-2-2-2-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            CellBackColor   =   16777215
         End
      End
      Begin VB.Frame fraSubCampana 
         Caption         =   "Sub Campaña"
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
         TabIndex        =   39
         Top             =   480
         Width           =   7575
         Begin VB.ComboBox cmbConsideracion 
            Height          =   315
            Left            =   5160
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cmbGarantias 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox txtConsideracion 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtDiasA 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtDiasB 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkTransferidos 
            Caption         =   "Transferidos"
            Height          =   255
            Left            =   2520
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkVencidos 
            Caption         =   "Vencidos"
            Height          =   255
            Left            =   1440
            TabIndex        =   20
            Top             =   1080
            Width           =   975
         End
         Begin SICMACT.EditMoney txtCapital 
            Height          =   300
            Left            =   2050
            TabIndex        =   16
            Top             =   720
            Width           =   600
            _ExtentX        =   1058
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
         Begin SICMACT.EditMoney txtInt 
            Height          =   300
            Left            =   3250
            TabIndex        =   17
            Top             =   720
            Width           =   600
            _ExtentX        =   1058
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
         Begin SICMACT.EditMoney txtMora 
            Height          =   300
            Left            =   4500
            TabIndex        =   18
            Top             =   720
            Width           =   600
            _ExtentX        =   1058
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
         Begin SICMACT.EditMoney txtGasto 
            Height          =   300
            Left            =   5750
            TabIndex        =   19
            Top             =   720
            Width           =   600
            _ExtentX        =   1058
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
         Begin SICMACT.EditMoney txtICV 
            Height          =   300
            Left            =   6840
            TabIndex        =   66
            Top             =   720
            Width           =   600
            _ExtentX        =   1058
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
         Begin VB.Label Label20 
            Caption         =   "ICV:"
            Height          =   255
            Left            =   6450
            TabIndex        =   67
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "-"
            Height          =   255
            Left            =   2160
            TabIndex        =   50
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label14 
            Caption         =   "Gasto:"
            Height          =   255
            Left            =   5200
            TabIndex        =   49
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Mora:"
            Height          =   255
            Left            =   4000
            TabIndex        =   48
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Int.:"
            Height          =   255
            Left            =   2850
            TabIndex        =   47
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Capital:"
            Height          =   255
            Left            =   1440
            TabIndex        =   46
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Consideración:"
            Height          =   255
            Left            =   3120
            TabIndex        =   45
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Garantías:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Estado Cred.:"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "% Descuentos:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Días de atraso:"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraTipoCredito 
         Caption         =   "Tipo de Creditos"
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
         Height          =   2535
         Left            =   -70320
         TabIndex        =   38
         Top             =   2400
         Width           =   4335
         Begin VB.ListBox lstTpoCred 
            Height          =   2085
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame fraAgencias 
         Caption         =   "Agencias"
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
         Height          =   2535
         Left            =   -74880
         TabIndex        =   37
         Top             =   2400
         Width           =   4455
         Begin VB.ListBox lstAgencias 
            Height          =   2085
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame fraDatosCamp 
         Caption         =   "Datos de la Campaña"
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
         Height          =   1815
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtAprobadoPor 
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   4455
         End
         Begin VB.TextBox txtNumMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5160
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "1"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtNombreCamp 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   360
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker dtpDesde 
            Height          =   300
            Left            =   1560
            TabIndex        =   7
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   119668737
            CurrentDate     =   37054
         End
         Begin MSComCtl2.DTPicker dtpHasta 
            Height          =   300
            Left            =   3120
            TabIndex        =   8
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Format          =   119668737
            CurrentDate     =   37054
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "-"
            Height          =   255
            Left            =   2880
            TabIndex        =   36
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label Label4 
            Caption         =   "N° Máximo de veces que el cliente puede acogerse a la campaña:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   5175
         End
         Begin VB.Label Label3 
            Caption         =   "Duración:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Aprobador por:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Campaña:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
      End
      Begin SICMACT.FlexEdit feNivelesApr 
         Height          =   3705
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   6535
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cod-Act-Nivel-Desde-Hasta-Aux"
         EncabezadosAnchos=   "400-0-400-3000-2000-2000-0"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-5-X"
         ListaControles  =   "0-0-4-0-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-R-L-R-R-C"
         FormatosEdit    =   "0-0-0-1-2-2-0"
         CantEntero      =   12
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   16777215
      End
   End
End
Attribute VB_Name = "frmRecupCampConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmRecupCampConfig
'** Descripción : Formulario para configurar las campaña de recuperaciones
'**               Creado segun TI-ERS035-2015
'** Creación    : WIOR, 20150522 09:00:00 AM
'**********************************************************************************************

Option Explicit
Private i As Integer
Private j As Integer
Private fnCodSub As Integer
Private pnTipoEjecucion As Integer
Private fbImpedirPegar As Boolean
Private MatSubCampana As Variant
Private MatAux As Variant
Private fnCod As Long
Private fnCant As Integer
Private fbDescCap, fbDescInt, fbDescMora, fbDescGasto  As Boolean
Private sPersCod As String 'CROB20180813 ERS055-2018

Private Sub cmdAceptarSC_Click()
Dim nId As Integer
Dim nCodSub As Long
fnCant = 0

If ValidaSubCamp Then
    If MsgBox("Estas seguro de grabar estos datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If Trim(feSubCampanas.TextMatrix(1, 0)) = "" Then
        fnCant = -1
    Else
        fnCant = feSubCampanas.Rows - 2
    End If
            
    If fnCodSub = -1 Then
        nId = fnCant + 1
        'ReDim Preserve MatSubCampana(12, 0 To nId)'JOEP
        ReDim Preserve MatSubCampana(13, 0 To nId) 'JOEP
        nCodSub = 0
    Else
        nId = fnCodSub
        nCodSub = MatSubCampana(0, nId)
    End If
    
    MatSubCampana(0, nId) = nCodSub
    MatSubCampana(1, nId) = Trim(txtDiasA.Text)
    MatSubCampana(2, nId) = Trim(txtDiasB.Text)
    MatSubCampana(3, nId) = Trim(txtConsideracion.Text)
    MatSubCampana(4, nId) = Trim(Right(cmbConsideracion.Text, 4))
    MatSubCampana(5, nId) = Trim(Left(cmbConsideracion.Text, 50))
    MatSubCampana(6, nId) = txtCapital.Text
    MatSubCampana(7, nId) = txtInt.Text
    MatSubCampana(8, nId) = TxtMora.Text
    MatSubCampana(9, nId) = txtGasto.Text
    MatSubCampana(10, nId) = chkVencidos.value 'JOEP
    MatSubCampana(11, nId) = chkTransferidos.value 'JOEP
    MatSubCampana(12, nId) = Trim(Right(cmbGarantias.Text, 4)) 'JOEP
    MatSubCampana(13, nId) = txtICV.Text 'JOEP
    
    LimpiarSubCampana
    Call HabilitaContrSubCamp(3, False)
    Call LlenarSubCamp
End If
End Sub

Private Sub LlenarSubCamp()
LimpiaFlex feSubCampanas
For i = 0 To UBound(MatSubCampana, 2)
    If Trim(MatSubCampana(0, 0)) <> "" Then
        feSubCampanas.AdicionaFila
        feSubCampanas.TextMatrix(i + 1, 1) = MatSubCampana(1, i) & " - " & MatSubCampana(2, i)
        feSubCampanas.TextMatrix(i + 1, 2) = Format(MatSubCampana(3, i), "00") & " " & MatSubCampana(5, i)
        feSubCampanas.TextMatrix(i + 1, 3) = Format(MatSubCampana(6, i), "#0.00")
        feSubCampanas.TextMatrix(i + 1, 4) = Format(MatSubCampana(7, i), "#0.00")
        feSubCampanas.TextMatrix(i + 1, 5) = Format(MatSubCampana(8, i), "#0.00")
        feSubCampanas.TextMatrix(i + 1, 6) = Format(MatSubCampana(9, i), "#0.00")
        
        feSubCampanas.TextMatrix(i + 1, 7) = Format(MatSubCampana(13, i), "#0.00") 'JOEP
        'feSubCampanas.TextMatrix(i + 1, 7) = "..."'JOEP
        feSubCampanas.TextMatrix(i + 1, 8) = "..." 'JOEP
    End If
Next i
End Sub
Private Function ValidaDatos() As Boolean
ValidaDatos = True

If Trim(txtNombreCamp.Text) = "" Then
    sstConfiguracion.Tab = 0
    MsgBox "Favor de ingresar el nombre de la campaña.", vbInformation, "Aviso"
    ValidaDatos = False
    txtNombreCamp.SetFocus
    Exit Function
End If

If Trim(txtAprobadoPor.Text) = "" Then
    sstConfiguracion.Tab = 0
    MsgBox "Favor de ingresar el porque fue aprobado la campaña.", vbInformation, "Aviso"
    ValidaDatos = False
    txtAprobadoPor.SetFocus
    Exit Function
End If

If CDate(dtpDesde.value) > CDate(dtpHasta.value) Then
    sstConfiguracion.Tab = 0
    MsgBox "La fecha de Inicio no puede ser mayor a la fecha final.", vbInformation, "Aviso"
    ValidaDatos = False
    dtpDesde.SetFocus
    Exit Function
End If

If Trim(txtNumMax.Text) = "" Or Not IsNumeric(txtNumMax.Text) Then
    sstConfiguracion.Tab = 0
    MsgBox "Favor de ingresar el número máximo de veces.", vbInformation, "Aviso"
    ValidaDatos = False
    txtAprobadoPor.SetFocus
    Exit Function
End If

If CDbl(txtNumMax.Text) < 1 Then
    sstConfiguracion.Tab = 0
    MsgBox "El número máximo de veces debe ser mayor a cero.", vbInformation, "Aviso"
    ValidaDatos = False
    txtAprobadoPor.SetFocus
    Exit Function
End If


ValidaDatos = False
For i = 0 To lstAgencias.ListCount - 1
    If lstAgencias.Selected(i) = True Then
        ValidaDatos = True
        Exit For
    End If
Next i

If Not ValidaDatos Then
    sstConfiguracion.Tab = 0
    MsgBox "Seleccione por lo menos una Agencia.", vbInformation, "Aviso"
    lstAgencias.SetFocus
    Exit Function
End If

ValidaDatos = False
For i = 0 To lstTpoCred.ListCount - 1
    If lstTpoCred.Selected(i) = True Then
        ValidaDatos = True
        Exit For
    End If
Next i

If Not ValidaDatos Then
    sstConfiguracion.Tab = 0
    MsgBox "Seleccione por lo menos 1 Tipo de Crédito.", vbInformation, "Aviso"
    lstTpoCred.SetFocus
    Exit Function
End If

If Trim(MatSubCampana(0, 0)) = "" Then
    sstConfiguracion.Tab = 1
    MsgBox "Ingrese por lo menos una Sub Campaña.", vbInformation, "Aviso"
    ValidaDatos = False
    cmdNuevoSC.SetFocus
    Exit Function
End If

ValidaDatos = False
For i = 1 To feNivelesApr.Rows - 1
    If Trim(feNivelesApr.TextMatrix(i, 2)) = "." Then
        ValidaDatos = True
        Exit For
    End If
Next i
    
If Not ValidaDatos Then
    sstConfiguracion.Tab = 2
    MsgBox "Seleccione por lo menos 1 Nivel de Aprobación.", vbInformation, "Aviso"
    feNivelesApr.SetFocus
    Exit Function
End If

For i = 1 To feNivelesApr.Rows - 1
    If Trim(feNivelesApr.TextMatrix(i, 2)) = "." Then
        If Not IsNumeric(feNivelesApr.TextMatrix(i, 5)) Then
            sstConfiguracion.Tab = 2
            MsgBox "Ingrese correctamente el valor ''Hasta'' del nivel de aprobación ''" & Trim(feNivelesApr.TextMatrix(i, 3)) & "''", vbInformation, "Aviso"
            ValidaDatos = False
            feNivelesApr.SetFocus
            Exit Function
        End If
        
        If CDbl(feNivelesApr.TextMatrix(i, 5)) = 0 Then
            sstConfiguracion.Tab = 2
            MsgBox "El valor ''Hasta'' del nivel de aprobación ''" & feNivelesApr.TextMatrix(i, 3) & "'' debe ser mayor a cero.", vbInformation, "Aviso"
            ValidaDatos = False
            feNivelesApr.SetFocus
            Exit Function
        End If
    
        If i > 1 Then
            'If Trim(feNivelesApr.TextMatrix(i, 2)) = "." Then
            For j = i - 1 To 0 Step -1
                If Trim(feNivelesApr.TextMatrix(j, 2)) = "." Then
                    Exit For
                End If
            Next j
            
            If j > 0 Then
                If (CDbl(feNivelesApr.TextMatrix(i, 5)) < CDbl(feNivelesApr.TextMatrix(j, 5))) Or (CDbl(feNivelesApr.TextMatrix(i, 4)) < CDbl(feNivelesApr.TextMatrix(j, 5))) Then
                    sstConfiguracion.Tab = 2
                    MsgBox "Lo valores del nivel de aprobación ''" & Trim(feNivelesApr.TextMatrix(i, 3)) & "'' deben ser mayores al de ''" & _
                            Trim(feNivelesApr.TextMatrix(j, 3)) & "''", vbInformation, "Aviso"
                    ValidaDatos = False
                    feNivelesApr.SetFocus
                    Exit Function
                End If
            End If
            'End If
            
    '        If CDbl(feNivelesApr.TextMatrix(i, 3)) < CDbl(feNivelesApr.TextMatrix(i - 1, 3)) Then
    '            sstConfiguracion.Tab = 2
    '            MsgBox "Lo valores del nivel de aprobación ''" & Trim(feNivelesApr.TextMatrix(i, 1)) & "'' deben ser mayores al de ''" & _
    '                    Trim(feNivelesApr.TextMatrix(i - 1, 1)) & "''", vbInformation, "Aviso"
    '            ValidaDatos = False
    '            feNivelesApr.SetFocus
    '            Exit Function
    '        End If
        End If
    End If
Next i

End Function

Private Function ValidaSubCamp() As Boolean
ValidaSubCamp = True

If Trim(txtDiasA.Text) = "" Or Not IsNumeric(txtDiasA.Text) Then
    MsgBox "Favor de ingresar el día de atraso inicial.", vbInformation, "Aviso"
    ValidaSubCamp = False
    txtDiasA.Text = "0"
    txtDiasA.SetFocus
    Exit Function
End If

If Trim(txtDiasB.Text) = "" Or Not IsNumeric(txtDiasB.Text) Then
    MsgBox "Favor de ingresar el día de atraso final.", vbInformation, "Aviso"
    ValidaSubCamp = False
    txtDiasB.Text = "0"
    txtDiasB.SetFocus
    Exit Function
End If

If CDbl(txtDiasA.Text) > CDbl(txtDiasB.Text) Then
    MsgBox "El día de atraso inicial no puede ser mayor al día de atraso final.", vbInformation, "Aviso"
    ValidaSubCamp = False
    txtDiasA.SetFocus
    Exit Function
End If

If Trim(txtConsideracion.Text) = "" Or Not IsNumeric(txtConsideracion.Text) Then
    MsgBox "Favor de ingresar la cantidad de consideración.", vbInformation, "Aviso"
    ValidaSubCamp = False
    txtConsideracion.SetFocus
    Exit Function
End If

If CDbl(txtConsideracion.Text) < 1 Then
    MsgBox "La cantidad de consideración debe ser mayor a cero.", vbInformation, "Aviso"
    ValidaSubCamp = False
    txtConsideracion.SetFocus
    Exit Function
End If

If Trim(cmbConsideracion.Text) = "" Then
    MsgBox "Favor de ingresar el Tipo de consideración", vbInformation, "Aviso"
    ValidaSubCamp = False
    cmbConsideracion.SetFocus
    Exit Function
End If

If fbDescCap Or fbDescInt Or fbDescMora Or fbDescGasto Then
    If CDbl(txtCapital.Text) = 0 And CDbl(txtInt.Text) = 0 And CDbl(TxtMora.Text) = 0 And CDbl(txtGasto.Text) = 0 Then
        MsgBox "Favor de ingresar por los menos uno de los Descuentos", vbInformation, "Aviso"
        ValidaSubCamp = False
        
        If TxtMora.Visible And TxtMora.Enabled Then
            TxtMora.SetFocus
        ElseIf txtInt.Visible And txtInt.Enabled Then
            txtInt.SetFocus
        ElseIf txtGasto.Visible And txtGasto.Enabled Then
            txtGasto.SetFocus
        ElseIf txtCapital.Visible And txtCapital.Enabled Then
            txtCapital.SetFocus
        End If
        
        Exit Function
    End If
End If

If chkVencidos.value = 0 And chkTransferidos.value = 0 Then
    MsgBox "Favor de checkear por lo menos uno de los estados de créditos", vbInformation, "Aviso"
    ValidaSubCamp = False
    chkVencidos.SetFocus
    Exit Function
End If

If Trim(cmbGarantias.Text) = "" Then
    MsgBox "Favor de ingresar el tipo garantía", vbInformation, "Aviso"
    ValidaSubCamp = False
    cmbGarantias.SetFocus
    Exit Function
End If
End Function

Private Sub cmdCancelar_Click()
If cmdGrabar.Enabled Then
    If MsgBox("Estas seguro de Cancelar la operación?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
End If
    
If pnTipoEjecucion = 1 Or pnTipoEjecucion = 0 Then
    Call Limpiar
ElseIf pnTipoEjecucion = 2 Then
    Call LimpiarSubCampana
    Call HabilitaContrSubCamp(1, False)
    Call CargarDatos(fnCod)
End If

Call LlenarSubCamp
pnTipoEjecucion = 0
End Sub

Private Sub cmdCancelarSC_Click()
LimpiarSubCampana
Call HabilitaContrSubCamp(3, False)
fnCodSub = -1
Call LlenarSubCamp
End Sub

Private Sub cmdEditar_Click()
pnTipoEjecucion = 2
Call HabilitarControles(4, False)
Call HabilitaContrSubCamp(2, False)
txtNombreCamp.SetFocus
sstConfiguracion.Tab = 0
End Sub

Private Sub cmdEditarSC_Click()
Call HabilitaContrSubCamp(3, True)
fnCodSub = feSubCampanas.row - 1

txtDiasA.Text = MatSubCampana(1, fnCodSub)
txtDiasB.Text = MatSubCampana(2, fnCodSub)
txtConsideracion.Text = MatSubCampana(3, fnCodSub)
cmbConsideracion.ListIndex = IndiceListaCombo(cmbConsideracion, MatSubCampana(4, fnCodSub))
txtCapital.Text = Format(MatSubCampana(6, fnCodSub), "#0.00")
txtInt.Text = Format(MatSubCampana(7, fnCodSub), "#0.00")
TxtMora.Text = Format(MatSubCampana(8, fnCodSub), "#0.00")
txtGasto.Text = Format(MatSubCampana(9, fnCodSub), "#0.00")
chkVencidos.value = MatSubCampana(10, fnCodSub)
chkTransferidos.value = MatSubCampana(11, fnCodSub)
cmbGarantias.ListIndex = IndiceListaCombo(cmbGarantias, MatSubCampana(12, fnCodSub))
txtICV.Text = Format(MatSubCampana(13, fnCodSub), "#0.00") 'JOEP

If txtDiasA.Visible And txtDiasA.Enabled Then
    txtDiasA.SetFocus
End If
End Sub

Private Sub cmdExaminar_Click()
Call Limpiar
fnCod = frmRecupCampLista.Inicio
If fnCod = 0 Then
    MsgBox "No ha seleccionado alguna campaña", vbInformation, "Aviso"
    Call Limpiar
Else
    Call CargarDatos(fnCod)
End If
End Sub

Private Sub cmdGrabar_Click()
If ValidaDatos Then
    If MsgBox("Estas seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    GrabarDatos
End If
End Sub

Private Sub cmdNuevo_Click()
pnTipoEjecucion = 1
Call HabilitarControles(2, False)
Call HabilitaContrSubCamp(2, False)

If txtNombreCamp.Visible And txtNombreCamp.Enabled Then
    txtNombreCamp.SetFocus
End If
sstConfiguracion.Tab = 0
End Sub

Private Sub cmdNuevoSC_Click()
Call HabilitaContrSubCamp(3, True)
txtDiasA.SetFocus
fnCodSub = -1
End Sub

Private Sub cmdQuitarSC_Click()
Dim nContador As Integer
nContador = 0
If MsgBox("Estas seguro de quitar la Sub Campaña", vbInformation + vbYesNo, "Aviso") = vbYes Then
    fnCodSub = feSubCampanas.row - 1
    
    ReDim MatAux(12, 0)
    For i = 0 To UBound(MatSubCampana, 2)
        If i <> fnCodSub Then
            ReDim Preserve MatAux(12, 0 To nContador)
            MatAux(0, nContador) = MatSubCampana(0, i)
            MatAux(1, nContador) = MatSubCampana(1, i)
            MatAux(2, nContador) = MatSubCampana(2, i)
            MatAux(3, nContador) = MatSubCampana(3, i)
            MatAux(4, nContador) = MatSubCampana(4, i)
            MatAux(5, nContador) = MatSubCampana(5, i)
            MatAux(6, nContador) = MatSubCampana(6, i)
            MatAux(7, nContador) = MatSubCampana(7, i)
            MatAux(8, nContador) = MatSubCampana(8, i)
            MatAux(9, nContador) = MatSubCampana(9, i)
            MatAux(10, nContador) = MatSubCampana(10, i)
            MatAux(11, nContador) = MatSubCampana(11, i)
            MatAux(12, nContador) = MatSubCampana(12, i)
            MatAux(13, nContador) = MatSubCampana(13, i) 'JOEP
            nContador = nContador + 1
        End If
    Next i
    

    MatSubCampana = MatAux
    Call LlenarSubCamp
End If
End Sub

Private Sub feNivelesApr_GotFocus()
fbImpedirPegar = True
End Sub

Private Sub feNivelesApr_LostFocus()
fbImpedirPegar = False
End Sub

Private Sub feNivelesApr_OnCellChange(pnRow As Long, pnCol As Long)
If Not IsNumeric(feNivelesApr.TextMatrix(pnRow, pnCol)) Then
    feNivelesApr.TextMatrix(pnRow, pnCol) = "0.00"
End If

For i = pnRow + 1 To feNivelesApr.Rows - 1
    'If feNivelesApr.Rows > i Then
        If Trim(feNivelesApr.TextMatrix(i, 2)) = "." Then
            feNivelesApr.TextMatrix(i, pnCol - 1) = Format(CDbl(feNivelesApr.TextMatrix(pnRow, pnCol)) + 0.01, "#0.00")
            Exit For
        End If
    'End If
Next i

If pnRow > 1 Then
    For j = pnRow - 1 To 0 Step -1
        If Trim(feNivelesApr.TextMatrix(j, 2)) = "." Then
            Exit For
        End If
    Next j
    
    If j > 0 Then
        feNivelesApr.TextMatrix(pnRow, pnCol - 1) = Format(CDbl(feNivelesApr.TextMatrix(j, pnCol)) + 0.01, "#0.00")
    End If

End If


'If feNivelesApr.Rows > pnRow + 1 Then
'    feNivelesApr.TextMatrix(pnRow + 1, pnCol - 1) = CDbl(feNivelesApr.TextMatrix(pnRow, pnCol)) + 0.01
'End If
End Sub

Private Sub feNivelesApr_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If pnCol = 2 Then
    If Trim(feNivelesApr.TextMatrix(pnRow, pnCol)) = "" Then
        feNivelesApr.TextMatrix(pnRow, 4) = "0.00"
        feNivelesApr.TextMatrix(pnRow, 5) = "0.00"
    End If
End If
End Sub

Private Sub feNivelesApr_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim sColumnas() As String
sColumnas = Split(feNivelesApr.ColumnasAEditar, "-")
If sColumnas(pnCol) = "X" Or (pnCol <> 2 And Trim(feNivelesApr.TextMatrix(pnRow, 2)) = "") Then
   Cancel = False
   SendKeys "{Tab}", True
   Exit Sub
End If
End Sub


Private Sub feSubCampanas_DblClick()
Dim sEstados As String

'If feSubCampanas.Col = 7 Then'JOEP
If feSubCampanas.Col = 8 Then 'JOEP
    sEstados = IIf(MatSubCampana(10, feSubCampanas.row - 1) = 0, "", "- Vencidos")
    sEstados = sEstados & IIf(MatSubCampana(11, feSubCampanas.row - 1) = 0, "", IIf(Trim(sEstados) = "", "", Chr(10)) & "- Transferidos")

    MsgBox sEstados, vbDefaultButton1, "Estados"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If fbImpedirPegar Then
    'KeyCode = 86 y Shift = 2 es Ctrl V
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End If
End Sub

Private Sub Form_Load()
pnTipoEjecucion = 0
Call HabilitarControles(1, True)
Call HabilitaContrSubCamp(1, False)
Call cargarControles
sstConfiguracion.Tab = 0
fnCod = 0
'ReDim MatSubCampana(12, 0)'joep
ReDim MatSubCampana(13, 0) 'joep
ReDim MatAux(12, 0)
End Sub
Private Sub GrabarDatos()
Dim bgrabar As Boolean
Dim oNCredito As COMNCredito.NCOMCredito
Dim RsNiveles As ADODB.Recordset
Dim ArrAgencias As Variant
Dim ArrTpoCred As Variant
Dim nContador As Integer

On Error GoTo ErrorGrabarDatos

Set RsNiveles = IIf(feNivelesApr.Rows - 1 > 0, feNivelesApr.GetRsNew(), Nothing)


nContador = 0
ReDim ArrAgencias(0)
For i = 0 To lstAgencias.ListCount - 1
    If lstAgencias.Selected(i) = True Then
        ReDim Preserve ArrAgencias(nContador)
        ArrAgencias(nContador) = Trim(Right(lstAgencias.List(i), 4))
        nContador = nContador + 1
    End If
Next i

nContador = 0
ReDim ArrTpoCred(0)
For i = 0 To lstTpoCred.ListCount - 1
    If lstTpoCred.Selected(i) = True Then
        ReDim Preserve ArrTpoCred(nContador)
        ArrTpoCred(nContador) = Trim(Right(lstTpoCred.List(i), 4))
        nContador = nContador + 1
    End If
Next i

If pnTipoEjecucion = 1 Then
    fnCod = 0
End If

Set oNCredito = New COMNCredito.NCOMCredito
bgrabar = oNCredito.GrabarCampanaRecuperaciones(fnCod, Trim(txtNombreCamp.Text), Trim(txtAprobadoPor.Text), CDate(dtpDesde.value), CDate(dtpHasta.value), _
                                                CInt(txtNumMax.Text), ArrAgencias, ArrTpoCred, MatSubCampana, RsNiveles)

If bgrabar Then
    If pnTipoEjecucion = 1 Then
        MsgBox "Los datos se grabaron correctamente.", vbInformation, "Aviso"
    ElseIf pnTipoEjecucion = 2 Then
        MsgBox "Los datos se actualizaron correctamente.", vbInformation, "Aviso"
    End If
                
    Call HabilitarControles(3, True)
    Call HabilitaContrSubCamp(1, False)
    pnTipoEjecucion = 2
Else
     MsgBox "Hubo errores al grabar la información", vbError, "Error"
End If

Exit Sub
ErrorGrabarDatos:
MsgBox Err.Number & " - " & Err.Description, vbError, "Error En Proceso"
End Sub

Private Sub cargarControles()
Dim oDAge As COMDConstantes.DCOMAgencias
Dim oDCredito As COMDCredito.DCOMCredito
Dim oDConstante As COMDConstantes.DCOMConstantes
Dim rsDatos As ADODB.Recordset

fbImpedirPegar = False
Me.dtpDesde.value = gdFecSis
Me.dtpHasta.value = gdFecSis

'Carga Lista Agencia
Set oDAge = New COMDConstantes.DCOMAgencias
Set rsDatos = oDAge.ObtieneAgencias()
Call CargarList(Me.lstAgencias, rsDatos)
Set oDAge = Nothing

'Carga Lista Tipo Credito
Set oDCredito = New COMDCredito.DCOMCredito
Set rsDatos = oDCredito.RecuperaTipoCredCabecera()
Call CargarList(Me.lstTpoCred, rsDatos)
Set oDCredito = Nothing

'Llenar Combos
Set oDConstante = New COMDConstantes.DCOMConstantes
Call Llenar_Combo_con_Recordset(oDConstante.RecuperaConstantes(7550), Me.cmbConsideracion)
Call Llenar_Combo_con_Recordset(oDConstante.RecuperaConstantes(7551), Me.cmbGarantias)
Set oDConstante = Nothing

'Cargar Niveles de Aprobación
Call CargarNiveles

'RecuperarCampanaRecupDescActivos
Set oDCredito = New COMDCredito.DCOMCredito
Set rsDatos = oDCredito.RecuperarCampanaRecupDescActivos()
Set oDCredito = Nothing
fbDescCap = False
fbDescInt = False
fbDescMora = False
fbDescGasto = False
If Not (rsDatos.EOF And rsDatos.BOF) Then
    For i = 1 To rsDatos.RecordCount
        Select Case rsDatos!nTpoDesc
            Case 1: 'Desc. Capital
                    fbDescCap = CBool(rsDatos!bActivo)
            Case 2: 'Desc. Interes
                    fbDescInt = CBool(rsDatos!bActivo)
            Case 3: 'Desc. Mora
                    fbDescMora = CBool(rsDatos!bActivo)
            Case 4: 'Desc. Gasto
                    fbDescGasto = CBool(rsDatos!bActivo)
        End Select
        rsDatos.MoveNext
    Next i
End If
Set rsDatos = Nothing

'CROB20180813 ERS055-2018
txtAnioAsigAnual.Text = Year(gdFecSis)
txtCantMaxDesctAsigAnual.Text = ObtenerValorCampanaMaxAnual(Year(gdFecSis))
'CROB20180813 ERS055-2018
End Sub
Private Sub CargarNiveles()
Dim oDConstante As COMDConstantes.DCOMConstantes
Dim rsDatos As ADODB.Recordset

LimpiaFlex Me.feNivelesApr
Set oDConstante = New COMDConstantes.DCOMConstantes
Set rsDatos = oDConstante.RecuperaConstantes(7552)
For i = 1 To rsDatos.RecordCount
    feNivelesApr.AdicionaFila
    feNivelesApr.TextMatrix(i, 1) = Trim(rsDatos!nConsValor)
    feNivelesApr.TextMatrix(i, 2) = ""
    feNivelesApr.TextMatrix(i, 3) = Trim(rsDatos!cConsDescripcion)
    feNivelesApr.TextMatrix(i, 4) = "0.00"
    feNivelesApr.TextMatrix(i, 5) = "0.00"
    rsDatos.MoveNext
Next i
Set rsDatos = Nothing
Set oDConstante = Nothing
End Sub
Private Sub Limpiar()
Call CheckLista(False, lstAgencias)
Call CheckLista(False, lstTpoCred)
Call HabilitarControles(1, True)
Call HabilitaContrSubCamp(1, False)

'Pestaña Datos
Me.txtNombreCamp.Text = ""
Me.txtAprobadoPor.Text = ""
Me.dtpDesde.value = gdFecSis
Me.dtpHasta.value = gdFecSis
Me.txtNumMax.Text = "1"
fnCod = 0
LimpiaFlex feSubCampanas
'ReDim MatSubCampana(12, 0)'joep
ReDim MatSubCampana(13, 0) 'joep
ReDim MatAux(12, 0)
Call CargarNiveles
End Sub

Private Sub LimpiarSubCampana()
'Sub Campaña
txtDiasA.Text = ""
txtDiasB.Text = ""
txtConsideracion.Text = ""
cmbConsideracion.ListIndex = -1
txtCapital.Text = "0"
txtInt.Text = "0"
TxtMora.Text = "0"
txtGasto.Text = "0"
txtICV.Text = "0" 'JOEP
chkVencidos.value = 0
chkTransferidos.value = 0
cmbGarantias.ListIndex = -1
End Sub

Private Sub CargarDatos(ByVal pnId As Long)
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsDatos As ADODB.Recordset
Dim nContador As Integer

Set oDCredito = New COMDCredito.DCOMCredito

Call CheckLista(False, lstAgencias)
Call CheckLista(False, lstTpoCred)

Set rsDatos = oDCredito.RecuperarCampanaRecup(pnId)
If Not (rsDatos.EOF And rsDatos.BOF) Then
    txtNombreCamp.Text = Trim(rsDatos!cNombre)
    txtAprobadoPor.Text = Trim(rsDatos!cAprobado)
    dtpDesde.value = Format(rsDatos!dfechaini, "dd/mm/yyyy")
    dtpHasta.value = Format(rsDatos!dfechafin, "dd/mm/yyyy")
    txtNumMax.Text = CInt(rsDatos!nNumMax)
    
    Set rsDatos = Nothing
    Set rsDatos = oDCredito.RecuperarCampanaRecupAgencias(pnId)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        nContador = 0
        For i = 1 To rsDatos.RecordCount
            For j = 0 To lstAgencias.ListCount - 1
                If Trim(Right(lstAgencias.List(j), 4)) = Trim(rsDatos!cAgeCod) Then
                    lstAgencias.Selected(j) = True
                    nContador = nContador + 1
                    Exit For
                End If
            Next j
            
            If rsDatos.RecordCount = nContador Then
                Exit For
            End If
            rsDatos.MoveNext
        Next i
    End If
    
    Set rsDatos = Nothing
    Set rsDatos = oDCredito.RecuperarCampanaRecupTpoCred(pnId)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        nContador = 0
        For i = 1 To rsDatos.RecordCount
            For j = 0 To lstTpoCred.ListCount - 1
                If Trim(Right(lstTpoCred.List(j), 4)) = Trim(rsDatos!cTpoCredCod) Then
                    lstTpoCred.Selected(j) = True
                    nContador = nContador + 1
                    Exit For
                End If
            Next j
            
            If rsDatos.RecordCount = nContador Then
                Exit For
            End If
            rsDatos.MoveNext
        Next i
    End If
    
    Set rsDatos = Nothing
    Set rsDatos = oDCredito.RecuperarCampanaRecupSubCamp(pnId)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        ReDim Preserve MatSubCampana(13, 0 To rsDatos.RecordCount - 1)
        For i = 0 To rsDatos.RecordCount - 1
            MatSubCampana(0, i) = rsDatos!nIdSubCamp
            MatSubCampana(1, i) = rsDatos!nDiasAtrasoIni
            MatSubCampana(2, i) = rsDatos!nDiasAtrasoFin
            MatSubCampana(3, i) = rsDatos!nConsidera
            MatSubCampana(4, i) = rsDatos!nTpoConsidera
            MatSubCampana(5, i) = rsDatos!cTpoConsidera
            MatSubCampana(6, i) = rsDatos!nDescCap
            MatSubCampana(7, i) = rsDatos!nDescInt
            MatSubCampana(8, i) = rsDatos!nDescMora
            MatSubCampana(9, i) = rsDatos!nDescGasto
            MatSubCampana(10, i) = IIf(CBool(rsDatos!bVencidos), 1, 0)
            MatSubCampana(11, i) = IIf(CBool(rsDatos!bTransferidos), 1, 0)
            MatSubCampana(12, i) = rsDatos!nTpoGarantia
            
            MatSubCampana(13, i) = rsDatos!nDescIcv 'JOEP
            rsDatos.MoveNext
        Next i
    End If
    
    Set rsDatos = Nothing
    Set rsDatos = oDCredito.RecuperarCampanaRecupNivelesApr(pnId)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        If feNivelesApr.TextMatrix(1, 0) <> "" Then
            nContador = 0
            For i = 1 To rsDatos.RecordCount
                For j = 1 To feNivelesApr.Rows - 1
                    If Trim(feNivelesApr.TextMatrix(j, 1)) = Trim(rsDatos!nNivel) Then
                        feNivelesApr.TextMatrix(j, 2) = "1"
                        feNivelesApr.TextMatrix(j, 4) = Format(rsDatos!nMontoDesde, "#0.00")
                        feNivelesApr.TextMatrix(j, 5) = Format(rsDatos!nMontoHasta, "#0.00")
                        nContador = nContador + 1
                        Exit For
                    End If
                Next j
                
                If rsDatos.RecordCount = nContador Then
                    Exit For
                End If
                rsDatos.MoveNext
            Next i
            feNivelesApr.TextMatrix(1, 4) = "0.00"
        End If
    End If
    
    Call HabilitarControles(3, True)
    Call LlenarSubCamp
End If
Set rsDatos = Nothing
Set oDCredito = Nothing
End Sub

Private Sub HabilitarControles(ByVal pnTipo As Integer, ByVal pbHabilita As Boolean)

'Botones Generales
Select Case pnTipo
    Case 1: 'Inicio
            cmdexaminar.Enabled = pbHabilita
            cmdNuevo.Enabled = pbHabilita
            cmdEditar.Enabled = Not pbHabilita
            cmdGrabar.Enabled = Not pbHabilita
    Case 2: 'Boton Nuevo
            cmdexaminar.Enabled = pbHabilita
            cmdNuevo.Enabled = pbHabilita
            cmdEditar.Enabled = pbHabilita
            cmdGrabar.Enabled = Not pbHabilita
    Case 3: 'Boton Examinar(Despues de Cargar Datos)
            cmdexaminar.Enabled = pbHabilita
            cmdNuevo.Enabled = Not pbHabilita
            cmdEditar.Enabled = pbHabilita
            cmdGrabar.Enabled = Not pbHabilita
    Case 4: 'Boton Editar
            cmdexaminar.Enabled = pbHabilita
            cmdNuevo.Enabled = pbHabilita
            cmdEditar.Enabled = pbHabilita
            cmdGrabar.Enabled = Not pbHabilita
End Select

'Pestaña Datos
'Datos
txtNombreCamp.Enabled = Not pbHabilita
txtAprobadoPor.Enabled = Not pbHabilita
dtpDesde.Enabled = Not pbHabilita
dtpHasta.Enabled = Not pbHabilita
txtNumMax.Enabled = Not pbHabilita

'Agencias
lstAgencias.Enabled = Not pbHabilita

'Tipo
lstTpoCred.Enabled = Not pbHabilita

'Niveles de Aprobacion
feNivelesApr.Enabled = Not pbHabilita
End Sub

Private Sub HabilitaContrSubCamp(ByVal pnTipo As Integer, ByVal pbHabilita As Boolean)
'Botones Generales
Select Case pnTipo
    Case 1: 'Inicio
            cmdAceptarSC.Enabled = pbHabilita
            cmdCancelarSC.Enabled = pbHabilita
            cmdNuevoSC.Enabled = pbHabilita
            cmdEditarSC.Enabled = pbHabilita
            cmdQuitarSC.Enabled = pbHabilita
            feSubCampanas.Enabled = pbHabilita
    Case 2: 'Boton Nuevo General
            cmdAceptarSC.Enabled = pbHabilita
            cmdCancelarSC.Enabled = pbHabilita
            cmdNuevoSC.Enabled = Not pbHabilita
            cmdEditarSC.Enabled = Not pbHabilita
            cmdQuitarSC.Enabled = Not pbHabilita
            feSubCampanas.Enabled = Not pbHabilita
    Case 3: 'Boton NuevoSC, cmdAceptarSC,cmdCancelarSC
            cmdAceptarSC.Enabled = pbHabilita
            cmdCancelarSC.Enabled = pbHabilita
            cmdNuevoSC.Enabled = Not pbHabilita
            cmdEditarSC.Enabled = Not pbHabilita
            cmdQuitarSC.Enabled = Not pbHabilita
            feSubCampanas.Enabled = Not pbHabilita
            
            If pnTipoEjecucion = 1 Then
                'cmdNuevo.Enabled = IIf(pbHabilita, False, True)
                cmdEditar.Enabled = IIf(pbHabilita, False, False)
                cmdGrabar.Enabled = IIf(pbHabilita, False, True)
                cmdcancelar.Enabled = IIf(pbHabilita, False, True)
            ElseIf pnTipoEjecucion = 2 Then
                cmdNuevo.Enabled = IIf(pbHabilita, False, False)
                'cmdEditar.Enabled = IIf(pbHabilita, False, True)
                cmdGrabar.Enabled = IIf(pbHabilita, False, True)
                cmdcancelar.Enabled = IIf(pbHabilita, False, True)
            End If
End Select

'Registro Sub Campaña
txtDiasA.Enabled = pbHabilita
txtDiasB.Enabled = pbHabilita
txtConsideracion.Enabled = pbHabilita
cmbConsideracion.Enabled = pbHabilita
txtCapital.Enabled = IIf(Not fbDescCap, False, pbHabilita)
txtInt.Enabled = IIf(Not fbDescInt, False, pbHabilita)
TxtMora.Enabled = IIf(Not fbDescMora, False, pbHabilita)
txtGasto.Enabled = IIf(Not fbDescGasto, False, pbHabilita)
chkVencidos.Enabled = pbHabilita
chkTransferidos.Enabled = pbHabilita
cmbGarantias.Enabled = pbHabilita
End Sub

Private Sub CargarList(ByRef lista As ListBox, ByVal prsDatos As ADODB.Recordset)
lista.Clear

For i = 1 To prsDatos.RecordCount
    lista.AddItem Trim(prsDatos!cConsDescripcion) & Space(100) & Trim(prsDatos!nConsValor)
    prsDatos.MoveNext
Next i
End Sub

Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
For i = 0 To lstLista.ListCount - 1
    lstLista.Selected(i) = bCheck
Next i
End Sub

Private Sub txtCapital_Change()
If Trim(txtCapital.Text) <> "." Then
    If CDbl(txtCapital.Text) > 100 Then
        txtCapital.Text = Replace(Mid(txtCapital.Text, 1, Len(txtCapital.Text) - 1), ",", "")
    End If
Else
    txtCapital.Text = "0.00"
End If
End Sub

Private Sub txtCapital_LostFocus()
txtCapital.Text = Format(txtCapital.Text, "#0.00")
End Sub

Private Sub txtConsideracion_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtConsideracion_LostFocus()
If IsNumeric(txtConsideracion.Text) Then
    If CInt(txtConsideracion.Text) < 1 Then
        txtConsideracion.Text = "1"
    End If
Else
    txtConsideracion.Text = "1"
End If
End Sub

Private Sub txtDiasA_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtDiasA_LostFocus()
If IsNumeric(txtDiasA.Text) Then
    If CInt(txtDiasA.Text) < 0 Then
        txtDiasA.Text = "0"
    End If
Else
    txtDiasA.Text = "0"
End If
End Sub

Private Sub txtDiasB_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtDiasB_LostFocus()
If IsNumeric(txtDiasB.Text) Then
    If CInt(txtDiasB.Text) < 0 Then
        txtDiasB.Text = "0"
    End If
Else
    txtDiasB.Text = "0"
End If
End Sub

Private Sub txtGasto_Change()
If Trim(txtGasto.Text) <> "." Then
    If CDbl(txtGasto.Text) > 100 Then
        txtGasto.Text = Replace(Mid(txtGasto.Text, 1, Len(txtGasto.Text) - 1), ",", "")
    End If
Else
    txtGasto.Text = "0.00"
End If
End Sub

Private Sub txtGasto_LostFocus()
txtGasto.Text = Format(txtGasto.Text, "#0.00")
End Sub

Private Sub txtInt_Change()
If Trim(txtInt.Text) <> "." Then
    If CDbl(txtInt.Text) > 100 Then
        txtInt.Text = Replace(Mid(txtInt.Text, 1, Len(txtInt.Text) - 1), ",", "")
    End If
Else
    txtInt.Text = "0.00"
End If
End Sub

Private Sub txtInt_LostFocus()
txtInt.Text = Format(txtInt.Text, "#0.00")
End Sub

Private Sub txtMora_Change()
If Trim(TxtMora.Text) <> "." Then
    If CDbl(TxtMora.Text) > 100 Then
        TxtMora.Text = Replace(Mid(TxtMora.Text, 1, Len(TxtMora.Text) - 1), ",", "")
    End If
Else
    TxtMora.Text = "0.00"
End If
End Sub

Private Sub txtMora_LostFocus()
TxtMora.Text = Format(TxtMora.Text, "#0.00")
End Sub

'JOEP
Private Sub txticv_change()
If Trim(txtICV.Text) <> "." Then
    If CDbl(txtICV.Text) > 100 Then
        txtICV.Text = Replace(Mid(txtICV.Text, 1, Len(txtICV.Text) - 1), ",", "")
    End If
Else
    txtICV.Text = "0.00"
End If
End Sub

Private Sub txtIcv_LostFocus()
txtICV.Text = Format(txtICV.Text, "#0.00")
End Sub
'JOEP

Private Sub txtNumMax_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If IsNumeric(txtNumMax.Text) Then
    If CInt(txtNumMax.Text) < 1 Then
        txtNumMax.Text = "1"
    End If
Else
    txtNumMax.Text = "1"
End If
End Sub

Private Sub txtNumMax_LostFocus()
If IsNumeric(txtNumMax.Text) Then
    If CInt(txtNumMax.Text) < 1 Then
        txtNumMax.Text = "1"
    End If
Else
    txtNumMax.Text = "1"
End If
End Sub

'CROB20180813 ERS055-2018
Private Sub cmdEditAcept_Click()
Dim nCantMax As Byte
    If cmdEditAcept.Caption = "Editar" Then
        Call habilitarBtnsAsigAnual(True, "Aceptar")
        txtCantMaxDesctAsigAnual.SetFocus
    Else
        'Ejecuta Actualizacion y Registro
        nCantMax = CByte(txtCantMaxDesctAsigAnual.Text)
        If nCantMax > 0 Then
            If MsgBox("¿Está seguro que desea actualizar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                Call RegistrarCampanasMaximaAnual(nCantMax, gsCodUser)
                Call habilitarBtnsAsigAnual(False, "Editar")
                MsgBox "Los datos se actualizaron correctamente.", vbInformation, "Aviso"
            End If
        Else
            MsgBox "El valor debe ser mayor a 0.", vbInformation, "Aviso"
            txtCantMaxDesctAsigAnual.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancAsigAnual_Click()
    'Obtengo el ultimo valor desde la BD y asigno el valor a txtCantMaxDesctAsigAnual
    txtCantMaxDesctAsigAnual.Text = ObtenerValorCampanaMaxAnual(Year(gdFecSis))
    'Bloquea botones
    Call habilitarBtnsAsigAnual(False, "Editar")
End Sub

Private Sub habilitarBtnsAsigAnual(ByVal bValor As Boolean, ByVal sCaption As String)
    txtCantMaxDesctAsigAnual.Enabled = bValor
    cmdCancAsigAnual.Enabled = bValor
    cmdEditAcept.Caption = sCaption
End Sub

Private Sub txtCantMaxDesctAsigAnual_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdEditAcept.SetFocus
    End If
End Sub

Private Sub RegistrarCampanasMaximaAnual(ByVal nCantMaxAnual As Byte, ByVal sUsuarioReg As String)
    Dim oNCredito As COMNCredito.NCOMCredito
    Set oNCredito = New COMNCredito.NCOMCredito

    Call oNCredito.RegistrarCampanaMaximaAnual(nCantMaxAnual, sUsuarioReg)
End Sub

Private Function ObtenerValorCampanaMaxAnual(ByVal cAnio As String) As Byte
    Dim oDCredito As COMDCredito.DCOMCredito
    Set oDCredito = New COMDCredito.DCOMCredito
    ObtenerValorCampanaMaxAnual = oDCredito.ObtenerValorMaxCampanasAnual(cAnio)!nCantMaxAnual
End Function

Private Sub habilitarBtnsExoneracion(ByVal bValor As Boolean, ByVal sCaption As String)
    txtCantAdiDsctExo.Enabled = bValor
    cmdBuscarPer.Enabled = bValor
    cmdCancExo.Enabled = bValor
    cmdEditExo.Caption = sCaption
End Sub

Private Function ObtenerValorCantAdicionalDesctosCampana(ByVal sPersCod As String) As ADODB.Recordset
    Dim oDCredito As COMDCredito.DCOMCredito
    Set oDCredito = New COMDCredito.DCOMCredito
    Set ObtenerValorCantAdicionalDesctosCampana = oDCredito.ObtenerValorCantAdicionalDesctosCampana(sPersCod)
End Function

Private Sub cmdEditExo_Click()
Dim nCantDescMax As Byte
    If cmdEditExo.Caption = "Editar" Then
        Call habilitarBtnsExoneracion(True, "Aceptar")
        cmdBuscarPer.SetFocus
    Else
        'Ejecuta Actualizacion y Registro
        nCantDescMax = CByte(txtCantAdiDsctExo.Text)
        If nCantDescMax > 0 Then
            If MsgBox("¿Está seguro que desea actualizar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                Call RegistrarDescuentoCampanaAdicionalAnual(sPersCod, nCantDescMax, gsCodUser)
                Call habilitarBtnsExoneracion(False, "Editar")
                MsgBox "Los datos se actualizaron correctamente.", vbInformation, "Aviso"
            End If
        Else
            MsgBox "El valor debe ser mayor a 0.", vbInformation, "Aviso"
            txtCantAdiDsctExo.SetFocus
        End If
    End If
End Sub

Private Sub cmdBuscarPer_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim rsCred As ADODB.Recordset
    
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        'LblPersCod.Caption = oPersona.sPersCod
        txtClienteExo.Text = oPersona.sPersNombre
    'Request query
    sPersCod = oPersona.sPersCod
    Set rsCred = ObtenerValorCantAdicionalDesctosCampana(sPersCod)
    txtCantAdiDsctExo.Text = rsCred!nCantAdicional
    txtCantAdiDsctExo.SetFocus
    Else
        Exit Sub
    End If
End Sub

Private Sub RegistrarDescuentoCampanaAdicionalAnual(ByVal psPersCod As String, ByVal pnCantDescMax As Byte, ByVal psUsuarioReg As String)
    Dim oNCredito As COMNCredito.NCOMCredito
    Set oNCredito = New COMNCredito.NCOMCredito

    Call oNCredito.RegistrarDescuentoCampanaAdicionalAnual(psPersCod, pnCantDescMax, psUsuarioReg)
End Sub

Private Sub cmdCancExo_Click()
    'Bloquea botones
    Call habilitarBtnsExoneracion(False, "Editar")
    txtClienteExo.Text = ""
    txtCantAdiDsctExo.Text = ""
    sPersCod = ""
    cmdEditExo.SetFocus
    'Obtengo el ultimo valor desde la BD y asigno el valor a txtCantAdiDsctExo
End Sub

Private Sub txtCantAdiDsctExo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdEditExo.SetFocus
    End If
End Sub
'CROB20180813 ERS055-2018

