VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersClienteSensible 
   Caption         =   "Cliente Procedimiento Reforzado"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   Icon            =   "frmPersClienteSensible.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   6600
      Width           =   1695
   End
   Begin TabDlg.SSTab tabClienteSensible 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Cliente/Conyuge"
      TabPicture(0)   =   "frmPersClienteSensible.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Parientes Cliente/Conyuge"
      TabPicture(1)   =   "frmPersClienteSensible.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Referencias Econónica"
      TabPicture(2)   =   "frmPersClienteSensible.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Referencia Financiera"
      TabPicture(3)   =   "frmPersClienteSensible.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Referencias Patrimonial"
      TabPicture(4)   =   "frmPersClienteSensible.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Persona Jurídica"
      TabPicture(5)   =   "frmPersClienteSensible.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdCancelarPersonaJuridica"
      Tab(5).Control(1)=   "cmdAceptarPersonaJuridica"
      Tab(5).Control(2)=   "Frame10"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Proveedores y/o Clientes"
      TabPicture(6)   =   "frmPersClienteSensible.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame9"
      Tab(6).Control(1)=   "Frame8"
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Conyuge"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   77
         Top             =   2040
         Width           =   9255
         Begin VB.TextBox TxtApellidosConyuge 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   86
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox TxtCentroLaborConyuge 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            TabIndex        =   85
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox TxtNombresConyuge 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   84
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cboNacionalidadConyuge 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox txtIngresoPromedioConyuge 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   7800
            MaxLength       =   9
            TabIndex        =   81
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox cboOcupacionConyuge 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   1680
            Width           =   3975
         End
         Begin VB.ComboBox cboDOIConyuge 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox TxtNumeroDoiConyuge 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   78
            Top             =   600
            Width           =   1575
         End
         Begin MSMask.MaskEdBox TxtFechaNacimientoConyuge 
            Height          =   315
            Left            =   7800
            TabIndex        =   82
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            Caption         =   "Apellidos"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Centro Laboral"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Nombres"
            Height          =   255
            Left            =   5160
            TabIndex        =   93
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Fecha de Nacimiento"
            Height          =   255
            Left            =   6120
            TabIndex        =   92
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Nacionalidad"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso Promedio (S/.) "
            Height          =   195
            Left            =   6000
            TabIndex        =   90
            Top             =   1380
            Width           =   1635
         End
         Begin VB.Label Label15 
            Caption         =   "Profesion u Ocupación"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Documento de Identidad"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Nº"
            Height          =   255
            Left            =   4680
            TabIndex        =   87
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parientes del Cliente"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   70
         Top             =   480
         Width           =   9255
         Begin VB.CommandButton cmdAgregarRelacionCliente 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   76
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarRelacionCliente 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   75
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarRelacionCliente 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   74
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelarRelacionCliente 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   73
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarRelacionCliente 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   72
            Top             =   1920
            Width           =   975
         End
         Begin SICMACT.FlexEdit FERelacionCliente 
            Height          =   2535
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   4471
            Cols0           =   5
            FixedCols       =   0
            HighLight       =   2
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Código-Apellidos y Nombres-Nº de DNI-salto"
            EncabezadosAnchos=   "400-1600-4200-1300-1"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-3-0-0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-C"
            FormatosEdit    =   "0-0-1-0-0"
            CantEntero      =   3
            TextArray0      =   "Nº"
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Principales Clientes"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   58
         Top             =   3480
         Width           =   6615
         Begin VB.CommandButton cmdCancelarCliente 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   69
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarCliente 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   68
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarCliente 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   67
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarCliente 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   66
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarCliente 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   65
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FECliente 
            Height          =   2535
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4471
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Nombre de Cliente"
            EncabezadosAnchos=   "400-4500"
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
            ColumnasAEditar =   "X-1"
            ListaControles  =   "0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L"
            FormatosEdit    =   "0-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdCancelarPersonaJuridica 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65520
         TabIndex        =   57
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPersonaJuridica 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65520
         TabIndex        =   56
         Top             =   2400
         Width           =   975
      End
      Begin VB.Frame Frame11 
         Caption         =   "Otros bienes"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   46
         Top             =   3480
         Width           =   6615
         Begin VB.CommandButton cmdCancelarOtroBien 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   52
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarOtroBien 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   51
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarOtroBien 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   50
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarOtroBien 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   49
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarOtroBien 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FEOtroBien 
            Height          =   2535
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4471
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Bienes"
            EncabezadosAnchos=   "400-4500"
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
            ColumnasAEditar =   "X-1"
            ListaControles  =   "0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L"
            FormatosEdit    =   "0-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Patrimonal"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   7815
         Begin VB.CommandButton cmdCancelarReferenciaPatrimonial 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            TabIndex        =   45
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarReferenciaPatrimonial 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            TabIndex        =   44
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarReferenciaPatrimonial 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            TabIndex        =   43
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarReferenciaPatrimonial 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            TabIndex        =   42
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarReferenciaPatrimonial 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FEReferenciaPatrimonial 
            Height          =   2535
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   4471
            Cols0           =   4
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Bienes Inmuebles/Muebles-Valor(US$)-salto"
            EncabezadosAnchos=   "400-4500-1200-1"
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
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-R-C"
            FormatosEdit    =   "0-0-2-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Financieras"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   9735
         Begin VB.CommandButton cmdAgregarReferenciaFinanciera 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8640
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarReferenciaFinanciera 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8640
            TabIndex        =   36
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarReferenciaFinanciera 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8640
            TabIndex        =   35
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarReferenciaFinanciera 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8640
            TabIndex        =   34
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelarReferenciaFinanciera 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8640
            TabIndex        =   33
            Top             =   2400
            Width           =   975
         End
         Begin SICMACT.FlexEdit FEReferenciaFinanciera 
            Height          =   2535
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4471
            Cols0           =   5
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Banco donde tiene cuenta-Tipo de Producto-Analista/Funcionario-salto"
            EncabezadosAnchos=   "400-3500-2000-2000-1"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-3-0-0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-C"
            FormatosEdit    =   "0-0-0-1-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Persona Jurídica en las que fuera Accionaista o Representante Legal"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   10455
         Begin VB.CommandButton cmdEliminarPersonaJuridica 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   55
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarPersonaJuridica 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   54
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarPersonaJuridica 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9360
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FEPersonaJuridica 
            Height          =   2535
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   4471
            Cols0           =   6
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Nombre de la Empresa-RUC-% Participación-Ingreso(S/.)-salto"
            EncabezadosAnchos=   "400-4500-1400-1200-1200-1"
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
            ColumnasAEditar =   "X-1-2-3-4-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-R-R-R-C"
            FormatosEdit    =   "0-0-0-3-2-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Principales Proveedores"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton cmdCancelarProveedor 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   64
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarProveedor 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   63
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarProveedor 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   62
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarProveedor 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   61
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarProveedor 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FEProveedor 
            Height          =   2535
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4471
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Nombre de Proveedor"
            EncabezadosAnchos=   "400-4500"
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
            ColumnasAEditar =   "X-1"
            ListaControles  =   "0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L"
            FormatosEdit    =   "0-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Economicas"
         Height          =   2895
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   12735
         Begin VB.CommandButton cmdAceptarReferenciaEconomica 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11640
            TabIndex        =   27
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelarReferenciaEconomica 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11640
            TabIndex        =   26
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarReferenciaEconomica 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11640
            TabIndex        =   25
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarReferenciaEconomica 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11640
            TabIndex        =   24
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarReferenciaEconomica 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11640
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FEReferenciaEconomica 
            Height          =   2535
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   4471
            Cols0           =   7
            FixedCols       =   0
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Cargo Público 5años ulti-Fecha de Inicio-Fecha de Cese-Entidad Labora(ó) en-Ingreso(S/.)-salto"
            EncabezadosAnchos=   "400-3500-1200-1200-3500-1200-1"
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
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-2-2-0-0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-L-R-C"
            FormatosEdit    =   "0-0-0-0-0-2-0"
            TextArray0      =   "Nº"
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Parientes del Conyuge"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   2
         Top             =   3480
         Width           =   9255
         Begin VB.CommandButton cmdAceptarRelacionConyuge 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   22
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelarRelacionConyuge 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   21
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarRelacionConyuge 
            Caption         =   "E&liminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   20
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdEditarRelacionConyuge 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarRelacionConyuge 
            Caption         =   "&Nuevo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8160
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin SICMACT.FlexEdit FERelacionConyuge 
            Height          =   2535
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   4471
            Cols0           =   5
            FixedCols       =   0
            HighLight       =   2
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Código-Apellidos y Nombres-Nº de DNI-salto"
            EncabezadosAnchos=   "400-1600-4200-1300-1"
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
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-3-0-0-0"
            BackColor       =   12648447
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-C"
            FormatosEdit    =   "0-0-1-0-0"
            CantEntero      =   3
            TextArray0      =   "Nº"
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   12648447
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Cliente"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   9255
         Begin VB.TextBox TxtCentroLaboralCliente 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   15
            Top             =   960
            Width           =   4335
         End
         Begin SICMACT.TxtBuscar TxtBuscarCliente 
            Height          =   255
            Left            =   1320
            TabIndex        =   13
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
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
         Begin VB.TextBox TxtPasaporteCliente 
            Height          =   285
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TxtCanetExtranjeriaCliente 
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TxtDniCliente 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TxtNombreCliente 
            Height          =   285
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label9 
            Caption         =   "Centro Laboral"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "C.E."
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Pasaporte"
            Height          =   255
            Left            =   5760
            TabIndex        =   10
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "D.N.I."
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmPersClienteSensible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************************
'** Nombre : frmPersClienteSensible
'** Descripción : Formulario para generar Anexo Nº 16 Ficha de Conocimiento
'                 del Cliente Sensible y Anexo Nº 5 Declaración Jurada de
'                 Origen de Fondos.
'** Creación : ELRO, 20110718 07:40:21 PM
'********************************************************************

Dim oCliente As COMDPersona.UCOMPersona
Dim oConyuge As COMDPersona.UCOMPersona
Dim oParentescoCliente As COMNPersona.NCOMPersona
Dim oParentescoConyuge As COMNPersona.NCOMPersona
Dim iAccion As Integer '1:Nuevo, 2:Editar, 3:Eliminar
Dim FilaNoEditar As Integer
Dim fsCodigoPersona As String
Dim fnCodigoClienteProcesoReforzado As Integer
Dim fcMovPersonaAtendido As String
Dim fcMovPersonaAutorizado As String
Dim fsMotivoRegistro As String
Dim fsEstadoCivil As String
Dim fsRelacionEntidad As String

Private Type TParienteCliente
nRelacionId As Integer
vRelacion As String
vNombreCompleto As String
vDni As String
End Type

Private ParienteCliente() As TParienteCliente
Private nNumeroParienteCliente As Integer

Private Type TParienteConyuge
nConyugeId As Integer
nRelacionId As Integer
vRelacion As String
vNombreCompleto As String
vDni As String
End Type

Private ParienteConyuge() As TParienteConyuge
Private nNumeroParienteConyuge As Integer

Private Type TReferenciaEconomica
vCargoPublico As String
vFechaInicio As String
vFechaCese As String
vEntidadLabora As String
dIngreso As Double
End Type

Private ReferenciaEconomica() As TReferenciaEconomica
Private nNumeroReferenciaEconomica As Integer

Private Type TReferenciaFinanciera
vCodigoBanco As String
vNombreBanco As String
vTipoProducto As String
vFuncionarioNegocio As String
End Type

Private ReferenciaFinanciera() As TReferenciaFinanciera
Private nNumeroReferenciaFinanciera As Integer

Private Type TReferenciaPatrimonial
vBien As String
dValor As Double
End Type

Private ReferenciaPatrimonial() As TReferenciaPatrimonial
Private nNumeroReferenciaPatrimonial As Integer

Private OtroBien() As TReferenciaPatrimonial
Private nNumeroOtroBien As Integer

Private Type TPersonaJuridica
vNombreEmpresa As String
vRuc As String
nParticipacion As Integer
dIngreso As Double
End Type

Private PersonaJuridica() As TPersonaJuridica
Private nNumeroPersonaJuridica As Integer

Private Type TProveedor
vNombreEmpresa As String
End Type

Private Proveedor() As TProveedor
Private nNumeroProveedor As Integer

Private Cliente() As TProveedor
Private nNumeroCliente As Integer



Private Sub actualizarNumeroParienteCliente(ByVal pvNuevoValor As Integer)
 nNumeroParienteCliente = pvNuevoValor
End Sub

Private Function recuperarNumeroParienteCliente() As Integer
 recuperarNumeroParienteCliente = nNumeroParienteCliente
End Function

Private Sub actualizarRelacionIdTitular(ByVal pnRelacionId As Integer, ByVal pnFila As Integer)
ParienteCliente(pnFila).nRelacionId = pnRelacionId
End Sub

Private Function recuperarRelacionIdTitular(ByVal pnFila) As Integer
recuperarRelacionIdTitular = ParienteCliente(pnFila).nRelacionId
End Function

Private Sub actualizarRelacionTitular(ByVal pvRelacion As String, ByVal pnFila As Integer)
ParienteCliente(pnFila).vRelacion = pvRelacion
End Sub

Private Function recuperarRelacionTitular(ByVal pnFila) As String
recuperarRelacionTitular = ParienteCliente(pnFila).vRelacion
End Function

Private Sub actualizarNombreCompletoTitular(ByVal pvNombreCompleto As String, ByVal pnFila As Integer)
ParienteCliente(pnFila).vNombreCompleto = pvNombreCompleto
End Sub

Private Function recuperarNombreCompletoTitular(ByVal pnFila) As String
recuperarNombreCompletoTitular = ParienteCliente(pnFila).vNombreCompleto
End Function

Private Sub actualizarDniTitular(ByVal pvDni As String, ByVal pnFila As Integer)
ParienteCliente(pnFila).vDni = pvDni
End Sub

Private Function recuperarDniTitular(ByVal pnFila) As String
recuperarDniTitular = ParienteCliente(pnFila).vDni
End Function

Public Sub adicionarParienteCliente()
ReDim Preserve ParienteCliente(recuperarNumeroParienteCliente)
End Sub

Private Sub actualizarNumeroParienteConyuge(ByVal pvNuevoValor As Integer)
 nNumeroParienteConyuge = pvNuevoValor
End Sub

Private Function recuperarNumeroParienteConyuge() As Integer
 recuperarNumeroParienteConyuge = nNumeroParienteConyuge
End Function

Private Sub actualizarConyuge(ByVal pnConyugeId As Integer, ByVal pnFila As Integer)
ParienteConyuge(pnFila).nConyugeId = pnConyugeId
End Sub

Private Function recuperarConyuge(ByVal pnFila) As String
recuperarConyuge = ParienteConyuge(pnFila).nConyugeId
End Function

Private Sub actualizarRelacionIdConyuge(ByVal pnRelacionId As Integer, ByVal pnFila As Integer)
ParienteConyuge(pnFila).nRelacionId = pnRelacionId
End Sub

Private Function recuperarRelacionIdConyuge(ByVal pnFila) As Integer
recuperarRelacionIdConyuge = ParienteConyuge(pnFila).nRelacionId
End Function

Private Sub actualizarRelacionConyuge(ByVal pvRelacion As String, ByVal pnFila As Integer)
ParienteConyuge(pnFila).vRelacion = pvRelacion
End Sub

Private Function recuperarRelacionConyuge(ByVal pnFila) As String
recuperarRelacionConyuge = ParienteConyuge(pnFila).vRelacion
End Function

Private Sub actualizarNombreCompletoConyuge(ByVal pvNombreCompleto As String, ByVal pnFila As Integer)
ParienteConyuge(pnFila).vNombreCompleto = pvNombreCompleto
End Sub

Private Function recuperarNombreCompletoConyuge(ByVal pnFila) As String
recuperarNombreCompletoConyuge = ParienteConyuge(pnFila).vNombreCompleto
End Function

Private Sub actualizarDniConyuge(ByVal pvDni As String, ByVal pnFila As Integer)
ParienteConyuge(pnFila).vDni = pvDni
End Sub

Private Function recuperarDniConyuge(ByVal pnFila) As String
recuperarDniConyuge = ParienteConyuge(pnFila).vDni
End Function

Public Sub adicionarParienteConyuge()
ReDim Preserve ParienteConyuge(recuperarNumeroParienteConyuge)
End Sub

Private Sub actualizarNumeroReferenciaEconomica(ByVal pvNuevoValor As Integer)
 nNumeroReferenciaEconomica = pvNuevoValor
End Sub

Private Function recuperarNumeroReferenciaEconomica() As Integer
 recuperarNumeroReferenciaEconomica = nNumeroReferenciaEconomica
End Function

Private Sub actualizarCargoPublicoReferenciaEconomica(ByVal pvCargoPublico As String, ByVal pnFila As Integer)
ReferenciaEconomica(pnFila).vCargoPublico = pvCargoPublico
End Sub

Private Function recuperarCargoPublicoReferenciaEconomica(ByVal pnFila) As String
recuperarCargoPublicoReferenciaEconomica = ReferenciaEconomica(pnFila).vCargoPublico
End Function

Private Sub actualizarFechaInicioReferenciaEconomica(ByVal pvFechaInicio As String, ByVal pnFila As Integer)
ReferenciaEconomica(pnFila).vFechaInicio = pvFechaInicio
End Sub

Private Function recuperarFechaInicioReferenciaEconomica(ByVal pnFila) As String
recuperarFechaInicioReferenciaEconomica = ReferenciaEconomica(pnFila).vFechaInicio
End Function

Private Sub actualizarFechaCeseReferenciaEconomica(ByVal pvFechaCese As String, ByVal pnFila As Integer)
ReferenciaEconomica(pnFila).vFechaCese = pvFechaCese
End Sub

Private Function recuperarFechaCeseReferenciaEconomica(ByVal pnFila) As String
recuperarFechaCeseReferenciaEconomica = ReferenciaEconomica(pnFila).vFechaCese
End Function

Private Sub actualizarEntidadLaboraReferenciaEconomica(ByVal pvEntidadLabora As String, ByVal pnFila As Integer)
ReferenciaEconomica(pnFila).vEntidadLabora = pvEntidadLabora
End Sub

Private Function recuperarEntidadLaboraReferenciaEconomica(ByVal pnFila) As String
recuperarEntidadLaboraReferenciaEconomica = ReferenciaEconomica(pnFila).vEntidadLabora
End Function

Private Sub actualizarIngresoReferenciaEconomica(ByVal pdIngreso As Double, ByVal pnFila As Integer)
ReferenciaEconomica(pnFila).dIngreso = pdIngreso
End Sub

Private Function recuperarIngresoReferenciaEconomica(ByVal pnFila) As String
recuperarIngresoReferenciaEconomica = ReferenciaEconomica(pnFila).dIngreso
End Function

Public Sub adicionarReferenciaEconomica()
ReDim Preserve ReferenciaEconomica(recuperarNumeroReferenciaEconomica)
End Sub

Private Sub actualizarNumeroReferenciaFinanciera(ByVal pvNuevoValor As Integer)
 nNumeroReferenciaFinanciera = pvNuevoValor
End Sub

Private Function recuperarNumeroReferenciaFinanciera() As Integer
 recuperarNumeroReferenciaFinanciera = nNumeroReferenciaFinanciera
End Function

Private Sub actualizarCodigoBancoReferenciaFinanciera(ByVal pvCodigoBanco As String, ByVal pnFila As Integer)
ReferenciaFinanciera(pnFila).vCodigoBanco = pvCodigoBanco
End Sub

Private Function recuperarCodigoBancoReferenciaFinanciera(ByVal pnFila) As String
recuperarCodigoBancoReferenciaFinanciera = ReferenciaFinanciera(pnFila).vCodigoBanco
End Function

Private Sub actualizarNombreBancoReferenciaFinanciera(ByVal pvNombreBanco As String, ByVal pnFila As Integer)
ReferenciaFinanciera(pnFila).vNombreBanco = pvNombreBanco
End Sub

Private Function recuperarNombreBancoReferenciaFinanciera(ByVal pnFila) As String
recuperarNombreBancoReferenciaFinanciera = ReferenciaFinanciera(pnFila).vNombreBanco
End Function

Private Sub actualizarTipoProductoReferenciaFinanciera(ByVal pvTipoProducto As String, ByVal pnFila As Integer)
ReferenciaFinanciera(pnFila).vTipoProducto = pvTipoProducto
End Sub

Private Function recuperarTipoProductoReferenciaFinanciera(ByVal pnFila) As String
recuperarTipoProductoReferenciaFinanciera = ReferenciaFinanciera(pnFila).vTipoProducto
End Function

Private Sub actualizarFuncionarioNegocioReferenciaFinanciera(ByVal pvFuncionarioNegocio As String, ByVal pnFila As Integer)
ReferenciaFinanciera(pnFila).vFuncionarioNegocio = pvFuncionarioNegocio
End Sub

Private Function recuperarFuncionarioNegocioReferenciaFinanciera(ByVal pnFila) As String
recuperarFuncionarioNegocioReferenciaFinanciera = ReferenciaFinanciera(pnFila).vFuncionarioNegocio
End Function

Public Sub adicionarReferenciaFinanciera()
ReDim Preserve ReferenciaFinanciera(recuperarNumeroReferenciaFinanciera)
End Sub

Private Sub actualizarNumeroReferenciaPatrimonial(ByVal pvNuevoValor As Integer)
 nNumeroReferenciaPatrimonial = pvNuevoValor
End Sub

Private Function recuperarNumeroReferenciaPatrimonial() As Integer
 recuperarNumeroReferenciaPatrimonial = nNumeroReferenciaPatrimonial
End Function

Private Sub actualizarBienReferenciaPatrimonial(ByVal pvBien As String, ByVal pnFila As Integer)
ReferenciaPatrimonial(pnFila).vBien = pvBien
End Sub

Private Function recuperarBienReferenciaPatrimonial(ByVal pnFila) As String
recuperarBienReferenciaPatrimonial = ReferenciaPatrimonial(pnFila).vBien
End Function

Private Sub actualizarValorReferenciaPatrimonial(ByVal pdValor As Double, ByVal pnFila As Integer)
ReferenciaPatrimonial(pnFila).dValor = pdValor
End Sub

Private Function recuperarValorReferenciaPatrimonial(ByVal pnFila) As String
recuperarValorReferenciaPatrimonial = ReferenciaPatrimonial(pnFila).dValor
End Function

Public Sub adicionarReferenciaPatrimonial()
ReDim Preserve ReferenciaPatrimonial(recuperarNumeroReferenciaPatrimonial)
End Sub

Private Sub actualizarNumeroOtroBien(ByVal pvNuevoValor As Integer)
 nNumeroOtroBien = pvNuevoValor
End Sub

Private Function recuperarNumeroOtroBien() As Integer
 recuperarNumeroOtroBien = nNumeroOtroBien
End Function

Private Sub actualizarOtroBien(ByVal pvBien As String, ByVal pnFila As Integer)
OtroBien(pnFila).vBien = pvBien
End Sub

Private Function recuperarOtroBien(ByVal pnFila) As String
recuperarOtroBien = OtroBien(pnFila).vBien
End Function

Public Sub adicionarOtroBien()
ReDim Preserve OtroBien(recuperarNumeroOtroBien)
End Sub

Private Sub actualizarNumeroPersonaJuridica(ByVal pvNuevoValor As Integer)
 nNumeroPersonaJuridica = pvNuevoValor
End Sub

Private Function recuperarNumeroPersonaJuridica() As Integer
 recuperarNumeroPersonaJuridica = nNumeroPersonaJuridica
End Function

Private Sub actualizarNombreEmpresaPersonaJuridica(ByVal pvNombreEmpresa As String, ByVal pnFila As Integer)
PersonaJuridica(pnFila).vNombreEmpresa = pvNombreEmpresa
End Sub

Private Function recuperarNombreEmpresaPersonaJuridica(ByVal pnFila) As String
recuperarNombreEmpresaPersonaJuridica = PersonaJuridica(pnFila).vNombreEmpresa
End Function

Private Sub actualizarRucPersonaJuridica(ByVal pvRuc As String, ByVal pnFila As Integer)
PersonaJuridica(pnFila).vRuc = pvRuc
End Sub

Private Function recuperarRucPersonaJuridica(ByVal pnFila) As String
recuperarRucPersonaJuridica = PersonaJuridica(pnFila).vRuc
End Function

Private Sub actualizarParticipacionPersonaJuridica(ByVal pnParticipacion As Integer, ByVal pnFila As Integer)
PersonaJuridica(pnFila).nParticipacion = pnParticipacion
End Sub

Private Function recuperarParticipacionPersonaJuridica(ByVal pnFila) As String
recuperarParticipacionPersonaJuridica = PersonaJuridica(pnFila).nParticipacion
End Function

Private Sub actualizarIngresoPersonaJuridica(ByVal pdIngreso As Double, ByVal pnFila As Integer)
PersonaJuridica(pnFila).dIngreso = pdIngreso
End Sub

Private Function recuperarIngresoPersonaJuridica(ByVal pnFila) As String
recuperarIngresoPersonaJuridica = PersonaJuridica(pnFila).dIngreso
End Function

Public Sub adicionarPersonaJuridica()
ReDim Preserve PersonaJuridica(recuperarNumeroPersonaJuridica)
End Sub

Private Sub actualizarNumeroProveedor(ByVal pvNuevoValor As Integer)
 nNumeroProveedor = pvNuevoValor
End Sub

Private Function recuperarNumeroProveedor() As Integer
 recuperarNumeroProveedor = nNumeroProveedor
End Function

Private Sub actualizarNombreEmpresaProveedor(ByVal pvNombreEmpresa As String, ByVal pnFila As Integer)
Proveedor(pnFila).vNombreEmpresa = pvNombreEmpresa
End Sub

Private Function recuperarNombreEmpresaProveedor(ByVal pnFila) As String
recuperarNombreEmpresaProveedor = Proveedor(pnFila).vNombreEmpresa
End Function

Public Sub adicionarProveedor()
ReDim Preserve Proveedor(recuperarNumeroProveedor)
End Sub

Private Sub actualizarNumeroCliente(ByVal pvNuevoValor As Integer)
 nNumeroCliente = pvNuevoValor
End Sub

Private Function recuperarNumeroCliente() As Integer
 recuperarNumeroCliente = nNumeroCliente
End Function

Private Sub actualizarNombreEmpresaCliente(ByVal pvNombreEmpresa As String, ByVal pnFila As Integer)
Cliente(pnFila).vNombreEmpresa = pvNombreEmpresa
End Sub

Private Function recuperarNombreEmpresaCliente(ByVal pnFila) As String
recuperarNombreEmpresaCliente = Cliente(pnFila).vNombreEmpresa
End Function

Public Sub adicionarCliente()
ReDim Preserve Cliente(recuperarNumeroCliente)
End Sub


Private Sub cargarControles()

Dim lrsParentescoCliente As ADODB.Recordset
Dim lrsParentescoConyuge As ADODB.Recordset
Dim lrsTipoDocumento As ADODB.Recordset
Dim lrsPaises As ADODB.Recordset
Dim lrsOcupaciones As ADODB.Recordset
Dim lrsBancos As ADODB.Recordset

FilaNoEditar = -1

Set oParentescoCliente = New COMNPersona.NCOMPersona
Set oParentescoConyuge = New COMNPersona.NCOMPersona

Set lrsParentescoCliente = New ADODB.Recordset
Set lrsParentescoConyuge = New ADODB.Recordset
Set lrsTipoDocumento = New ADODB.Recordset
Set lrsPaises = New ADODB.Recordset
Set lrsOcupaciones = New ADODB.Recordset
Set lrsBancos = New ADODB.Recordset

Set lrsParentescoCliente = oParentescoCliente.listarParentescosClienteReforzado("T")
Set lrsParentescoConyuge = oParentescoConyuge.listarParentescosClienteReforzado("N")
Set lrsTipoDocumento = oParentescoConyuge.listarTiposDocumentosClienteReforzado
Set lrsPaises = oParentescoConyuge.listarPaisesClienteReforzado("P")
Set lrsOcupaciones = oParentescoConyuge.listarOcupacionesClienteReforzado
Set lrsBancos = oParentescoConyuge.listarBancosClienteReforzado


'Cargar en el combo los datos de los parentescos del cliente
Me.FERelacionCliente.CargaCombo lrsParentescoCliente
'Cargar en el combo bancos del cliente
Me.FEReferenciaFinanciera.CargaCombo lrsBancos
'Cargar en el combo los datos de los parentescos del conyuge
Me.FERelacionConyuge.CargaCombo lrsParentescoConyuge
'Cargar en el combo tipo de documento del conyuge
Do While Not lrsTipoDocumento.EOF
Me.cboDOIConyuge.AddItem (lrsTipoDocumento!vTipoDocumento)
lrsTipoDocumento.MoveNext
Loop
lrsTipoDocumento.Close
cboDOIConyuge.ListIndex = 0
'Cargar en el combo nacionalidad del conyuge
Do While Not lrsPaises.EOF
Me.cboNacionalidadConyuge.AddItem (lrsPaises!vNacionalidad & Space(130) & lrsPaises!nNacionalidad)
lrsPaises.MoveNext
Loop
lrsPaises.Close
cboNacionalidadConyuge.ListIndex = IndiceListaCombo(cboNacionalidadConyuge, "04028")
'Cargar en el combo ocupacion del conyuge
Do While Not lrsOcupaciones.EOF
Me.cboOcupacionConyuge.AddItem (lrsOcupaciones!vOcupacion)
lrsOcupaciones.MoveNext
Loop
lrsOcupaciones.Close


'lrsParentescoCliente.Close
'lrsParentescoConyuge.Clone

Set lrsParentescoCliente = Nothing
Set lrsParentescoConyuge = Nothing
Set lrsTipoDocumento = Nothing
Set lrsPaises = Nothing
Set lrsOcupaciones = Nothing
Set lrsBancos = Nothing


Set oParentescoCliente = Nothing
Set oParentescoConyuge = Nothing

End Sub
Private Function validarDatosCliente() As Boolean
Dim J As Integer
validarDatosCliente = True

If Trim(Me.TxtCentroLaboralCliente.Text) = "" Then
    MsgBox "Debe ingresar el Centro Laboral del Cliente", vbInformation, "Aviso"
    validarDatosCliente = False
    TxtCentroLaboralCliente.SetFocus
    Exit Function
End If

End Function

Private Function validarParentescoCliente() As Boolean
Dim i, J As Integer

validarParentescoCliente = True

'Valida el parentesco
 If Trim(FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 1)) = "" Then
    MsgBox "Debe elegir el parentesco de la persona", vbInformation, "Aviso"
    validarParentescoCliente = False
    FERelacionCliente.SetFocus
    Exit Function
 End If
    
'Valida el nombre del Titular no esté como relación
 If FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 2) = Me.TxtNombreCliente.Text Then
    MsgBox "No se puede agregar el nombre del Titular en la lista de pariente", vbInformation, "Aviso"
    validarParentescoCliente = False
    FERelacionCliente.SetFocus
    Exit Function
 End If
 'Valida el DNI del Titular no esté como relación
  If FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3) = Me.TxtDniCliente.Text Then
    MsgBox "No se puede agregar el DNI del Titular en la lista de pariente", vbInformation, "Aviso"
    validarParentescoCliente = False
    FERelacionCliente.SetFocus
    Exit Function
 End If
    
'Valida el nombre del Conyuge no esté en la relación
 If Trim(FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 2)) = Trim(Me.TxtApellidosConyuge.Text & Me.TxtNombresConyuge.Text) Then
    MsgBox "No se puede agregar el nombre del Conyuge en la lista de pariente", vbInformation, "Aviso"
    validarParentescoCliente = False
    FERelacionCliente.SetFocus
    Exit Function
 End If
 
 'Valida el DNI del Conyuge no esté en la relación
 If FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3) = Me.TxtNumeroDoiConyuge.Text And _
    Trim(Me.TxtNumeroDoiConyuge.Text) <> "" And Trim(Right(Me.cboDOIConyuge.Text, 3)) = "1" Then
    MsgBox "No se puede agregar el documento de identidad del Conyuge en la lista de pariente", vbInformation, "Aviso"
    validarParentescoCliente = False
    FERelacionCliente.SetFocus
    Exit Function
 End If
    
 'Valida si se ingreso los apellidos y nombres del pariente del cliente
  If Len(Trim(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 2))) = 0 Then
     MsgBox "Falta Ingresar los apellidos y nombres del pariente", vbInformation, "Aviso"
     validarParentescoCliente = False
     FERelacionCliente.SetFocus
     Exit Function
  End If
    
 'Valida la longitud del DNI y sus valores
  If Trim(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3)) <> "" Then
     If Len(Trim(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3))) <> gnNroDigitosDNI Then
        MsgBox "El número de DNI no es de " & gnNroDigitosDNI & " dígitos", vbInformation, "Aviso"
        validarParentescoCliente = False
        FERelacionCliente.SetFocus
        Exit Function
      End If
           
      For J = 1 To Len(Trim(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3)))
          If (Mid(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3), J, 1) < "0" Or _
             Mid(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.row, 3), J, 1) > "9") Then
             MsgBox "Uno de los dígitos del DNI no es un número", vbInformation, "Aviso"
             validarParentescoCliente = False
             FERelacionCliente.SetFocus
             Exit Function
            End If
         Next J
   End If
   
'Valida duplicidad de documento
If Me.FERelacionCliente.Rows >= 3 Then
   For i = 1 To Me.FERelacionCliente.Rows - 2
       If Trim(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.Rows - 1, 3)) <> "" And _
          Trim(Me.FERelacionCliente.TextMatrix(i, 3)) <> "" Then
          If Trim(Me.FERelacionCliente.TextMatrix(Me.FERelacionCliente.Rows - 1, 3)) = Trim(Me.FERelacionCliente.TextMatrix(i, 3)) Then
             MsgBox "Existe DNI duplicado", vbInformation, "Aviso"
             validarParentescoCliente = False
             FERelacionCliente.SetFocus
             Exit Function
          End If
       End If
   Next i
End If
End Function

Private Sub cargarParientesClientes(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FERelacionCliente.lbEditarFlex = True
    Call LimpiaFlex(Me.FERelacionCliente)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroParienteCliente - 1
    If i <> pnFila Then
       Me.FERelacionCliente.AdicionaFila
       Me.FERelacionCliente.TextMatrix(J, 1) = recuperarRelacionTitular(i)
       Me.FERelacionCliente.TextMatrix(J, 2) = recuperarNombreCompletoTitular(i)
       Me.FERelacionCliente.TextMatrix(J, 3) = recuperarDniTitular(i)
       
       Call actualizarRelacionIdTitular(CInt(Trim(Right(Me.FERelacionCliente.TextMatrix(J, 1), 10))), J)
       Call actualizarRelacionTitular(Me.FERelacionCliente.TextMatrix(J, 1), J)
       Call actualizarNombreCompletoTitular(Me.FERelacionCliente.TextMatrix(J, 2), J)
       Call actualizarDniTitular(Me.FERelacionCliente.TextMatrix(J, 3), J)
       J = J + 1
    End If
    Next i
    Call actualizarNumeroParienteCliente(Me.FERelacionCliente.Rows)
    Else
        For i = 1 To recuperarNumeroParienteCliente - 1
        Me.FERelacionCliente.AdicionaFila
        Me.FERelacionCliente.TextMatrix(i, 1) = recuperarRelacionTitular(i)
        Me.FERelacionCliente.TextMatrix(i, 2) = recuperarNombreCompletoTitular(i)
        Me.FERelacionCliente.TextMatrix(i, 3) = recuperarDniTitular(i)
    Next i
    End If
    
    Me.FERelacionCliente.lbEditarFlex = False
End Sub

Private Function validarDatosConyuge() As Boolean
Dim J As Integer
validarDatosConyuge = True
'valida apellidos del conyuge
If Trim(Me.TxtApellidosConyuge.Text) = "" Then
 MsgBox "Debe ingresar los apellidos del conyuge", vbInformation, "Aviso"
      validarDatosConyuge = False
      TxtApellidosConyuge.SetFocus
      Exit Function
End If

'valida nombres del conyuge
If Trim(Me.TxtNombresConyuge.Text) = "" Then
 MsgBox "Debe ingresar el(los) nombre(s) del conyuge", vbInformation, "Aviso"
      validarDatosConyuge = False
      TxtNombresConyuge.SetFocus
      Exit Function
End If

'valida tipo de documento de conyuge
If Trim(Me.cboDOIConyuge.Text) = "" Then
 MsgBox "Debe seleccionar el tipo de documento del conyuge", vbInformation, "Aviso"
      validarDatosConyuge = False
      cboDOIConyuge.SetFocus
      Exit Function
End If

'Valida la longitud del DNI y sus valores
If Trim(Me.TxtNumeroDoiConyuge.Text) = "" Then
 MsgBox "Debe ingresar el número del documento del conyuge", vbInformation, "Aviso"
      validarDatosConyuge = False
      TxtNumeroDoiConyuge.SetFocus
      Exit Function
End If

If Trim(Me.TxtNumeroDoiConyuge.Text) <> "" And Trim(Right((Me.cboDOIConyuge.Text), 3)) = "1" Then
   If Len(Trim(Me.TxtNumeroDoiConyuge.Text)) <> gnNroDigitosDNI Then
      MsgBox "El número de DNI no es de " & gnNroDigitosDNI & " dígitos", vbInformation, "Aviso"
      validarDatosConyuge = False
      TxtNumeroDoiConyuge.SetFocus
      Exit Function
    End If
      
    For J = 1 To Len(Trim(Me.TxtNumeroDoiConyuge.Text))
        If (Mid(Trim(Me.TxtNumeroDoiConyuge.Text), J, 1) < "0" Or _
           Mid(Trim(Me.TxtNumeroDoiConyuge.Text), J, 1) > "9") Then
           MsgBox "Uno de los dígitos del DNI no es un número", vbInformation, "Aviso"
           validarDatosConyuge = False
           TxtNumeroDoiConyuge.SetFocus
           Exit Function
        End If
     Next J
  End If
      
 'valida nacionalidad de conyuge
 If Trim(Me.cboNacionalidadConyuge.Text) = "" Then
    MsgBox "Debe seleccionar la nacionalidad del conyuge", vbInformation, "Aviso"
    validarDatosConyuge = False
    cboNacionalidadConyuge.SetFocus
    Exit Function
 End If
'valida fecha de nacimento del conyuge
 If Len(Trim(ValidaFecha(Me.TxtFechaNacimientoConyuge.Text))) <> 0 Then
    MsgBox ValidaFecha(Me.TxtFechaNacimientoConyuge.Text), vbInformation, "Aviso"
    validarDatosConyuge = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If

'validad centro laboral del conyuge
If Trim(Me.TxtCentroLaborConyuge.Text) = "" Then
    MsgBox "Debe ingresar el centro laboral del conyuge", vbInformation, "Aviso"
    validarDatosConyuge = False
    cboNacionalidadConyuge.SetFocus
    Exit Function
 End If

'valida el ingreso promedio del conyuge
If Trim(Me.txtIngresoPromedioConyuge.Text) = "" Or Trim(Me.txtIngresoPromedioConyuge.Text) = "0" Then
    MsgBox "Debe ingresar el centro laboral del conyuge", vbInformation, "Aviso"
    validarDatosConyuge = False
    txtIngresoPromedioConyuge.SetFocus
    Exit Function
 End If
      
'valida la ocupacion del conyuge
If Trim(Me.cboOcupacionConyuge.Text) = "" Then
    MsgBox "Debe seleccionar la ocupacion del conyuge", vbInformation, "Aviso"
    validarDatosConyuge = False
    cboOcupacionConyuge.SetFocus
    Exit Function
 End If
      
End Function

Private Function validarParentescoConyuge() As Boolean
Dim i, J As Integer

validarParentescoConyuge = True

'Valida el parentesco
 If Trim(FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 1)) = "" Then
    MsgBox "Debe elegir el parentesco de la persona", vbInformation, "Aviso"
    validarParentescoConyuge = False
    FERelacionConyuge.SetFocus
    Exit Function
 End If
    
'Valida el nombre del Titular no esté como relación
 If FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 2) = Me.TxtNombreCliente.Text Then
    MsgBox "No se puede agregar el nombre del Titular en la lista de pariente", vbInformation, "Aviso"
    validarParentescoConyuge = False
    FERelacionConyuge.SetFocus
    Exit Function
 End If
 'Valida el DNI del Titular no esté como relación
  If FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3) = Me.TxtDniCliente.Text Then
    MsgBox "No se puede agregar el DNI del Titular en la lista de pariente", vbInformation, "Aviso"
    validarParentescoConyuge = False
    FERelacionConyuge.SetFocus
    Exit Function
 End If
    
'Valida el nombre del Conyuge no esté en la relación
 If Trim(FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 2)) = Trim(Me.TxtApellidosConyuge.Text & Me.TxtNombresConyuge.Text) Then
    MsgBox "No se puede agregar el nombre del Conyuge en la lista de pariente", vbInformation, "Aviso"
    validarParentescoConyuge = False
    FERelacionConyuge.SetFocus
    Exit Function
 End If
 
 'Valida el DNI del Conyuge no esté en la relación
 If FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3) = Me.TxtNumeroDoiConyuge.Text And _
    Trim(Me.TxtNumeroDoiConyuge.Text) <> "" And Trim(Right((Me.cboDOIConyuge.Text), 3)) = "1" Then
    MsgBox "No se puede agregar el documento de identidad del Conyuge en la lista de pariente", vbInformation, "Aviso"
    validarParentescoConyuge = False
    FERelacionConyuge.SetFocus
    Exit Function
 End If
    
 'Valida si se ingreso los apellidos y nombres del pariente del Conyuge
  If Len(Trim(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 2))) = 0 Then
     MsgBox "Falta Ingresar los apellidos y nombres del pariente", vbInformation, "Aviso"
     validarParentescoConyuge = False
     FERelacionConyuge.SetFocus
     Exit Function
  End If
    
 'Valida la longitud de DNI y sus valores
  If Trim(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3)) <> "" Then
     If Len(Trim(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3))) <> gnNroDigitosDNI Then
        MsgBox "El número de DNI no es de " & gnNroDigitosDNI & " dígitos", vbInformation, "Aviso"
        validarParentescoConyuge = False
        FERelacionConyuge.SetFocus
        Exit Function
      End If
           
      For J = 1 To Len(Trim(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3)))
          If (Mid(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3), J, 1) < "0" Or _
             Mid(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.row, 3), J, 1) > "9") Then
             MsgBox "Uno de los dígitos del DNI no es un número", vbInformation, "Aviso"
             validarParentescoConyuge = False
             FERelacionConyuge.SetFocus
             Exit Function
            End If
         Next J
   End If
   
'Valida duplicidad de documento
If Me.FERelacionConyuge.Rows >= 3 Then
   For i = 1 To Me.FERelacionConyuge.Rows - 2
       If Trim(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.Rows - 1, 3)) <> "" And _
          Trim(Me.FERelacionConyuge.TextMatrix(i, 3)) <> "" Then
          If Trim(Me.FERelacionConyuge.TextMatrix(Me.FERelacionConyuge.Rows - 1, 3)) = Trim(Me.FERelacionConyuge.TextMatrix(i, 3)) Then
             MsgBox "Existe DNI duplicado", vbInformation, "Aviso"
             validarParentescoConyuge = False
             FERelacionConyuge.SetFocus
             Exit Function
          End If
       End If
   Next i
End If

End Function

Private Sub cargarParientesConyuges(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FERelacionConyuge.lbEditarFlex = True
    Call LimpiaFlex(Me.FERelacionConyuge)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroParienteConyuge - 1
    If i <> pnFila Then
       Me.FERelacionConyuge.AdicionaFila
       Me.FERelacionConyuge.TextMatrix(J, 1) = recuperarRelacionConyuge(i)
       Me.FERelacionConyuge.TextMatrix(J, 2) = recuperarNombreCompletoConyuge(i)
       Me.FERelacionConyuge.TextMatrix(J, 3) = recuperarDniConyuge(i)
       
       Call actualizarRelacionIdConyuge(CInt(Trim(Right(Me.FERelacionConyuge.TextMatrix(J, 1), 10))), J)
       Call actualizarRelacionConyuge(Me.FERelacionConyuge.TextMatrix(J, 1), J)
       Call actualizarNombreCompletoConyuge(Me.FERelacionConyuge.TextMatrix(J, 2), J)
       Call actualizarDniConyuge(Me.FERelacionConyuge.TextMatrix(J, 3), J)
       J = J + 1
    End If
    Next i
    Call actualizarNumeroParienteConyuge(Me.FERelacionConyuge.Rows)
    Else
        For i = 1 To recuperarNumeroParienteConyuge - 1
        Me.FERelacionConyuge.AdicionaFila
        Me.FERelacionConyuge.TextMatrix(i, 1) = recuperarRelacionConyuge(i)
        Me.FERelacionConyuge.TextMatrix(i, 2) = recuperarNombreCompletoConyuge(i)
        Me.FERelacionConyuge.TextMatrix(i, 3) = recuperarDniConyuge(i)
    Next i
    End If
    
    Me.FERelacionConyuge.lbEditarFlex = False
End Sub

Private Function validarReferenciaEconomica() As Boolean

validarReferenciaEconomica = True

'Valida el cargo público
 If Trim(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 1)) = "" Then
    MsgBox "Debe ingresar el cargo público", vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
 
 'Valida la fecha de inicio
 If Len(Trim(ValidaFecha(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 2)))) <> 0 Then
    MsgBox ValidaFecha(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 2)), vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
 
  'Valida la fecha de cese
 If Len(Trim(ValidaFecha(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 3)))) <> 0 Then
    MsgBox ValidaFecha(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 3)), vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
 
 'Compara las fechas inicio y cese que no sean iguales
 If CDate(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 2)) >= CDate(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 3)) Then
    MsgBox "La fecha de ingreso debe ser menor y diferente que la fecha de cese", vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
  
 'Valida el entidad
 If Trim(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 4)) = "" Then
    MsgBox "Debe ingresar la entidad donde trabajo", vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
 
 'Valida el importe
 If Trim(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 5)) = "" Then
    MsgBox "Debe ingresar el importe del cargo", vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
 
 If Trim(Me.FEReferenciaEconomica.TextMatrix(Me.FEReferenciaEconomica.row, 5)) = "0" Then
    MsgBox "El importe del cargo debe ser mayor que cero", vbInformation, "Aviso"
    validarReferenciaEconomica = False
    FEReferenciaEconomica.SetFocus
    Exit Function
 End If
  
End Function

Private Sub cargarReferenciasEconomicas(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FEReferenciaEconomica.lbEditarFlex = True
    Call LimpiaFlex(Me.FEReferenciaEconomica)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroReferenciaEconomica - 1
    If i <> pnFila Then
       Me.FEReferenciaEconomica.AdicionaFila
       Me.FEReferenciaEconomica.TextMatrix(J, 1) = recuperarCargoPublicoReferenciaEconomica(i)
       Me.FEReferenciaEconomica.TextMatrix(J, 2) = recuperarFechaInicioReferenciaEconomica(i)
       Me.FEReferenciaEconomica.TextMatrix(J, 3) = recuperarFechaCeseReferenciaEconomica(i)
       Me.FEReferenciaEconomica.TextMatrix(J, 4) = recuperarEntidadLaboraReferenciaEconomica(i)
       Me.FEReferenciaEconomica.TextMatrix(J, 5) = recuperarIngresoReferenciaEconomica(i)
       
       Call actualizarCargoPublicoReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(J, 1), J)
       Call actualizarFechaInicioReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(J, 2), J)
       Call actualizarFechaCeseReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(J, 3), J)
       Call actualizarEntidadLaboraReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(J, 4), J)
        Call actualizarIngresoReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(J, 5), J)
       J = J + 1
    End If
    Next i
    Call actualizarNumeroReferenciaEconomica(Me.FEReferenciaEconomica.Rows)
    Else
        For i = 1 To recuperarNumeroReferenciaEconomica - 1
        Me.FEReferenciaEconomica.AdicionaFila
        Me.FEReferenciaEconomica.TextMatrix(i, 1) = recuperarCargoPublicoReferenciaEconomica(i)
        Me.FEReferenciaEconomica.TextMatrix(i, 2) = recuperarFechaInicioReferenciaEconomica(i)
        Me.FEReferenciaEconomica.TextMatrix(i, 3) = recuperarFechaCeseReferenciaEconomica(i)
        Me.FEReferenciaEconomica.TextMatrix(i, 4) = recuperarEntidadLaboraReferenciaEconomica(i)
        Me.FEReferenciaEconomica.TextMatrix(i, 5) = recuperarIngresoReferenciaEconomica(i)
    Next i
    End If
    
    Me.FEReferenciaEconomica.lbEditarFlex = False
End Sub

Private Function validarReferenciaFinanciera() As Boolean

validarReferenciaFinanciera = True

'Valida el banco
 If Trim(Me.FEReferenciaFinanciera.TextMatrix(Me.FEReferenciaFinanciera.row, 1)) = "" Then
    MsgBox "Debe seleccionar un Banco", vbInformation, "Aviso"
    validarReferenciaFinanciera = False
    FEReferenciaFinanciera.SetFocus
    Exit Function
 End If
 
'Valida el tipo de producto
 If Trim(Me.FEReferenciaFinanciera.TextMatrix(Me.FEReferenciaFinanciera.row, 2)) = "" Then
    MsgBox "Debe ingresar el tipo de producto", vbInformation, "Aviso"
    validarReferenciaFinanciera = False
    FEReferenciaFinanciera.SetFocus
    Exit Function
 End If
    
End Function

Private Sub cargarReferenciasFinancieras(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FEReferenciaFinanciera.lbEditarFlex = True
    Call LimpiaFlex(Me.FEReferenciaFinanciera)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroReferenciaFinanciera - 1
    If i <> pnFila Then
       Me.FEReferenciaFinanciera.AdicionaFila
       Me.FEReferenciaFinanciera.TextMatrix(J, 1) = recuperarNombreBancoReferenciaFinanciera(i)
       Me.FEReferenciaFinanciera.TextMatrix(J, 2) = recuperarTipoProductoReferenciaFinanciera(i)
       Me.FEReferenciaFinanciera.TextMatrix(J, 3) = recuperarFuncionarioNegocioReferenciaFinanciera(i)
              
       Call actualizarNombreBancoReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(J, 1), J)
       Call actualizarTipoProductoReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(J, 2), J)
       Call actualizarFuncionarioNegocioReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(J, 3), J)
       Call actualizarCodigoBancoReferenciaFinanciera(Trim(Right(Me.FEReferenciaFinanciera.TextMatrix(J, 1), 15)), J)
       J = J + 1
    End If
    Next i
    Call actualizarNumeroReferenciaFinanciera(Me.FEReferenciaFinanciera.Rows)
    Else
        For i = 1 To recuperarNumeroReferenciaFinanciera - 1
        Me.FEReferenciaFinanciera.AdicionaFila
        Me.FEReferenciaFinanciera.TextMatrix(i, 1) = recuperarNombreBancoReferenciaFinanciera(i)
        Me.FEReferenciaFinanciera.TextMatrix(i, 2) = recuperarTipoProductoReferenciaFinanciera(i)
        Me.FEReferenciaFinanciera.TextMatrix(i, 3) = recuperarFuncionarioNegocioReferenciaFinanciera(i)
    
    Next i
    End If
    
    Me.FEReferenciaFinanciera.lbEditarFlex = False
End Sub

Private Function validarReferenciaPatrimonial() As Boolean

validarReferenciaPatrimonial = True

'Valida el bienes inmuebles/muebles
 If Trim(Me.FEReferenciaPatrimonial.TextMatrix(Me.FEReferenciaPatrimonial.row, 1)) = "" Then
    MsgBox "Debe ingresar un bien inmueble/mueble", vbInformation, "Aviso"
    validarReferenciaPatrimonial = False
    FEReferenciaPatrimonial.SetFocus
    Exit Function
 End If
 
'Valida el valor
 If Trim(Me.FEReferenciaPatrimonial.TextMatrix(Me.FEReferenciaPatrimonial.row, 2)) = "" Then
    MsgBox "Debe ingresar el valor del bien inmueble/mueble", vbInformation, "Aviso"
    validarReferenciaPatrimonial = False
    FEReferenciaPatrimonial.SetFocus
    Exit Function
 End If
 
 If Trim(Me.FEReferenciaPatrimonial.TextMatrix(Me.FEReferenciaPatrimonial.row, 2)) = "0" Then
    MsgBox "El valor del bien inmueble/mueble debe ser mayor que cero", vbInformation, "Aviso"
    validarReferenciaPatrimonial = False
    FEReferenciaPatrimonial.SetFocus
    Exit Function
 End If
    
End Function

Private Sub cargarReferenciasPatrimoniales(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FEReferenciaPatrimonial.lbEditarFlex = True
    Call LimpiaFlex(Me.FEReferenciaPatrimonial)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroReferenciaPatrimonial - 1
    If i <> pnFila Then
       Me.FEReferenciaPatrimonial.AdicionaFila
       Me.FEReferenciaPatrimonial.TextMatrix(J, 1) = recuperarBienReferenciaPatrimonial(i)
       Me.FEReferenciaPatrimonial.TextMatrix(J, 2) = recuperarValorReferenciaPatrimonial(i)
              
       Call actualizarBienReferenciaPatrimonial(Me.FEReferenciaPatrimonial.TextMatrix(J, 1), J)
       Call actualizarValorReferenciaPatrimonial(Me.FEReferenciaPatrimonial.TextMatrix(J, 2), J)
       J = J + 1
    End If
    Next i
    Call actualizarNumeroReferenciaPatrimonial(Me.FEReferenciaPatrimonial.Rows)
    Else
        For i = 1 To recuperarNumeroReferenciaPatrimonial - 1
        Me.FEReferenciaPatrimonial.AdicionaFila
        Me.FEReferenciaPatrimonial.TextMatrix(i, 1) = recuperarBienReferenciaPatrimonial(i)
        Me.FEReferenciaPatrimonial.TextMatrix(i, 2) = recuperarValorReferenciaPatrimonial(i)
    
    Next i
    End If
    
    Me.FEReferenciaPatrimonial.lbEditarFlex = False
End Sub

Private Function validarOtroBien() As Boolean

validarOtroBien = True

'Valida otro bien
 If Trim(Me.FEOtroBien.TextMatrix(Me.FEOtroBien.row, 1)) = "" Then
    MsgBox "Debe ingresar un bien", vbInformation, "Aviso"
    validarOtroBien = False
    FEOtroBien.SetFocus
    Exit Function
 End If
 
End Function

Private Sub cargarOtrosBienes(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FEOtroBien.lbEditarFlex = True
    Call LimpiaFlex(Me.FEOtroBien)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroOtroBien - 1
    If i <> pnFila Then
       Me.FEOtroBien.AdicionaFila
       Me.FEOtroBien.TextMatrix(J, 1) = recuperarOtroBien(i)
                   
       Call actualizarOtroBien(Me.FEOtroBien.TextMatrix(J, 1), J)

       J = J + 1
    End If
    Next i
    Call actualizarNumeroOtroBien(Me.FEOtroBien.Rows)
    Else
        For i = 1 To recuperarNumeroOtroBien - 1
        Me.FEOtroBien.AdicionaFila
        Me.FEOtroBien.TextMatrix(i, 1) = recuperarOtroBien(i)
       
    Next i
    End If
    
    Me.FEOtroBien.lbEditarFlex = False
End Sub

Private Function validarPersonaJuridica() As Boolean
Dim i, J As Integer

validarPersonaJuridica = True

'Valida el nombre de empresa
 If Trim(FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 1)) = "" Then
    MsgBox "Debe ingresar el nombre de la empresa", vbInformation, "Aviso"
    validarPersonaJuridica = False
    FEPersonaJuridica.SetFocus
    Exit Function
 End If

 'Valida la longitud del RUC
 If Len(Trim(Me.FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 2))) <> gnNroDigitosRUC Then
    MsgBox "El número del RUC no es de " & gnNroDigitosRUC & " dígitos", vbInformation, "Aviso"
    validarPersonaJuridica = False
    FEPersonaJuridica.SetFocus
    Exit Function
 End If
           
 For J = 1 To Len(Trim(Me.FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 2)))
    If (Mid(Me.FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 2), J, 1) < "0" Or _
        Mid(Me.FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 2), J, 1) > "9") Then
        MsgBox "Uno de los dígitos del RUC no es un número", vbInformation, "Aviso"
        validarPersonaJuridica = False
        FEPersonaJuridica.SetFocus
        Exit Function
     End If
  Next J
   
'Valida duplicidad de documento y nombre de empresa
If Me.FEPersonaJuridica.Rows >= 3 Then
   For i = 1 To Me.FEPersonaJuridica.Rows - 2
       If Trim(Me.FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.Rows - 1, 1)) = Trim(Me.FEPersonaJuridica.TextMatrix(i, 1)) Or _
          Trim(Me.FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.Rows - 1, 2)) = Trim(Me.FEPersonaJuridica.TextMatrix(i, 2)) Then
          MsgBox "Existe RUC duplicado", vbInformation, "Aviso"
          validarPersonaJuridica = False
          FEPersonaJuridica.SetFocus
          Exit Function
       End If
     
   Next i
End If

'Valida la participación
 If Trim(FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 3)) = "" Then
    MsgBox "Debe ingresar el % de la participación", vbInformation, "Aviso"
    validarPersonaJuridica = False
    FEPersonaJuridica.SetFocus
    Exit Function
 End If
 
 'Valida el ingreso
 If Trim(FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 4)) = "" Then
    MsgBox "Debe ingresar el ingreso", vbInformation, "Aviso"
    validarPersonaJuridica = False
    FEPersonaJuridica.SetFocus
    Exit Function
 End If
 
 If (Trim(FEPersonaJuridica.TextMatrix(Me.FEPersonaJuridica.row, 4))) = 0 Then
    MsgBox "El ingreso  debe ser mayor que cero", vbInformation, "Aviso"
    validarPersonaJuridica = False
    FEPersonaJuridica.SetFocus
    Exit Function
 End If
 
End Function

Private Sub cargarPersonasJuridicas(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FEPersonaJuridica.lbEditarFlex = True
    Call LimpiaFlex(Me.FEPersonaJuridica)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroPersonaJuridica - 1
    If i <> pnFila Then
       Me.FEPersonaJuridica.AdicionaFila
       Me.FEPersonaJuridica.TextMatrix(J, 1) = recuperarNombreEmpresaPersonaJuridica(i)
       Me.FEPersonaJuridica.TextMatrix(J, 2) = recuperarRucPersonaJuridica(i)
       Me.FEPersonaJuridica.TextMatrix(J, 3) = recuperarParticipacionPersonaJuridica(i)
       Me.FEPersonaJuridica.TextMatrix(J, 4) = recuperarIngresoPersonaJuridica(i)
              
       Call actualizarNombreEmpresaPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(J, 1), J)
       Call actualizarRucPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(J, 2), J)
       Call actualizarParticipacionPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(J, 3), J)
       Call actualizarIngresoPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(J, 4), J)
       J = J + 1
    End If
    Next i
    Call actualizarNumeroPersonaJuridica(Me.FEPersonaJuridica.Rows)
    Else
        For i = 1 To recuperarNumeroPersonaJuridica - 1
        Me.FEPersonaJuridica.AdicionaFila
        Me.FEPersonaJuridica.TextMatrix(i, 1) = recuperarNombreEmpresaPersonaJuridica(i)
        Me.FEPersonaJuridica.TextMatrix(i, 2) = recuperarRucPersonaJuridica(i)
        Me.FEPersonaJuridica.TextMatrix(i, 3) = recuperarParticipacionPersonaJuridica(i)
        Me.FEPersonaJuridica.TextMatrix(i, 4) = recuperarIngresoPersonaJuridica(i)
    
    Next i
    End If
    
    Me.FEPersonaJuridica.lbEditarFlex = False
End Sub

Private Function validarProveedor() As Boolean

validarProveedor = True

'Valida nombre proveedor
 If Trim(Me.FEProveedor.TextMatrix(Me.FEProveedor.row, 1)) = "" Then
    MsgBox "Debe ingresar el nombre del proveedor", vbInformation, "Aviso"
    validarProveedor = False
    FEProveedor.SetFocus
    Exit Function
 End If
 
End Function


Private Sub cargarProveedores(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FEProveedor.lbEditarFlex = True
    Call LimpiaFlex(Me.FEProveedor)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroProveedor - 1
    If i <> pnFila Then
       Me.FEProveedor.AdicionaFila
       Me.FEProveedor.TextMatrix(J, 1) = recuperarNombreEmpresaProveedor(i)
                   
       Call actualizarNombreEmpresaProveedor(Me.FEProveedor.TextMatrix(J, 1), J)

       J = J + 1
    End If
    Next i
    Call actualizarNumeroProveedor(Me.FEProveedor.Rows)
    Else
        For i = 1 To recuperarNumeroProveedor - 1
        Me.FEProveedor.AdicionaFila
        Me.FEProveedor.TextMatrix(i, 1) = recuperarNombreEmpresaProveedor(i)
       
    Next i
    End If
    
    Me.FEProveedor.lbEditarFlex = False
End Sub

Private Function validarCliente() As Boolean

validarCliente = True

'Valida nombre Cliente
 If Trim(Me.FECliente.TextMatrix(Me.FECliente.row, 1)) = "" Then
    MsgBox "Debe ingresar el nombre del Cliente", vbInformation, "Aviso"
    validarCliente = False
    FECliente.SetFocus
    Exit Function
 End If
 
End Function

Private Sub cargarClientes(Optional ByVal pnFila As Integer = 0)
Dim i, J As Integer
    Me.FECliente.lbEditarFlex = True
    Call LimpiaFlex(Me.FECliente)
    If iAccion = 3 And pnFila > 0 Then
    J = 1
    For i = 1 To recuperarNumeroCliente - 1
    If i <> pnFila Then
       Me.FECliente.AdicionaFila
       Me.FECliente.TextMatrix(J, 1) = recuperarNombreEmpresaCliente(i)
                   
       Call actualizarNombreEmpresaCliente(Me.FECliente.TextMatrix(J, 1), J)

       J = J + 1
    End If
    Next i
    Call actualizarNumeroCliente(Me.FECliente.Rows)
    Else
        For i = 1 To recuperarNumeroCliente - 1
        Me.FECliente.AdicionaFila
        Me.FECliente.TextMatrix(i, 1) = recuperarNombreEmpresaCliente(i)
       
    Next i
    End If
    
    Me.FECliente.lbEditarFlex = False
End Sub

Private Sub limpiarDatosCliente()
Me.TxtBuscarCliente.Text = ""
Me.TxtNombreCliente.Text = ""
Me.TxtDniCliente.Text = ""
Me.TxtCanetExtranjeriaCliente.Text = ""
Me.TxtPasaporteCliente.Text = ""
Me.TxtCentroLaboralCliente.Text = ""
End Sub

Private Sub limpiarDatosConyuge()
Me.TxtApellidosConyuge.Text = ""
Me.TxtNombresConyuge.Text = ""
Me.TxtNumeroDoiConyuge.Text = ""
Me.TxtCentroLaborConyuge.Text = ""
Me.txtIngresoPromedioConyuge.Text = 0
End Sub

Private Sub habiliraDatosCliente()
Me.TxtCentroLaboralCliente.Enabled = True
End Sub

Private Sub deshabilitarDatosCliente()
Me.TxtCentroLaboralCliente.Enabled = False
End Sub

Private Sub habilitarDatosConyuge()
Me.TxtApellidosConyuge.Enabled = True
Me.TxtNombresConyuge.Enabled = True
Me.cboDOIConyuge.Enabled = True
Me.TxtNumeroDoiConyuge.Enabled = True
Me.cboNacionalidadConyuge.Enabled = True
Me.TxtFechaNacimientoConyuge.Enabled = True
Me.TxtCentroLaborConyuge.Enabled = True
Me.txtIngresoPromedioConyuge.Enabled = True
Me.cboOcupacionConyuge.Enabled = True
End Sub

Private Sub deshabilitarDatosConyuge()
Me.TxtApellidosConyuge.Enabled = False
Me.TxtNombresConyuge.Enabled = False
Me.cboDOIConyuge.Enabled = False
Me.TxtNumeroDoiConyuge.Enabled = False
Me.cboNacionalidadConyuge.Enabled = False
Me.TxtFechaNacimientoConyuge.Enabled = False
Me.TxtCentroLaborConyuge.Enabled = False
Me.txtIngresoPromedioConyuge.Enabled = False
Me.cboOcupacionConyuge.Enabled = False
End Sub

Private Sub habilitarBotonesParentescoCliente()
Me.cmdAgregarRelacionCliente.Enabled = True
Me.cmdEditarRelacionCliente.Enabled = True
Me.cmdEliminarRelacionCliente.Enabled = True
End Sub

Private Sub deshabilitarBotonesParentescoCliente()
Me.cmdAgregarRelacionCliente.Enabled = False
Me.cmdEditarRelacionCliente.Enabled = False
Me.cmdEliminarRelacionCliente.Enabled = False
End Sub

Private Sub habilitarBotonesParentescoConyuge()
Me.cmdAgregarRelacionConyuge.Enabled = True
Me.cmdEditarRelacionConyuge.Enabled = True
Me.cmdEliminarRelacionConyuge.Enabled = True
End Sub

Private Sub deshabilitarBotonesParentescoConyuge()
Me.cmdAgregarRelacionConyuge.Enabled = False
Me.cmdEditarRelacionConyuge.Enabled = False
Me.cmdEliminarRelacionConyuge.Enabled = False
End Sub

Private Sub habilitarBotonesReferenciaEconomica()
Me.cmdAgregarReferenciaEconomica.Enabled = True
Me.cmdEditarReferenciaEconomica.Enabled = True
Me.cmdEliminarReferenciaEconomica.Enabled = True
End Sub

Private Sub deshabilitarBotonesReferenciaEconomica()
Me.cmdAgregarReferenciaEconomica.Enabled = False
Me.cmdEditarReferenciaEconomica.Enabled = False
Me.cmdEliminarReferenciaEconomica.Enabled = False
End Sub

Private Sub habilitarBotonesReferenciaFinanciera()
Me.cmdAgregarReferenciaFinanciera.Enabled = True
Me.cmdEditarReferenciaFinanciera.Enabled = True
Me.cmdEliminarReferenciaFinanciera.Enabled = True
End Sub

Private Sub deshabilitarBotonesReferenciaFinanciera()
Me.cmdAgregarReferenciaFinanciera.Enabled = False
Me.cmdEditarReferenciaFinanciera.Enabled = False
Me.cmdEliminarReferenciaFinanciera.Enabled = False
End Sub

Private Sub habilitarBotonesReferenciaPatrimonial()
Me.cmdAgregarReferenciaPatrimonial.Enabled = True
Me.cmdEditarReferenciaPatrimonial.Enabled = True
Me.cmdEliminarReferenciaPatrimonial.Enabled = True
End Sub

Private Sub deshabilitarBotonesReferenciaPatrimonial()
Me.cmdAgregarReferenciaPatrimonial.Enabled = False
Me.cmdEditarReferenciaPatrimonial.Enabled = False
Me.cmdEliminarReferenciaPatrimonial.Enabled = False
End Sub

Private Sub habilitarBotonesOtroBien()
Me.cmdAgregarOtroBien.Enabled = True
Me.cmdEditarOtroBien.Enabled = True
Me.cmdEliminarOtroBien.Enabled = True
End Sub

Private Sub deshabilitarBotonesOtroBien()
Me.cmdAgregarOtroBien.Enabled = False
Me.cmdEditarOtroBien.Enabled = False
Me.cmdEliminarOtroBien.Enabled = False
End Sub

Private Sub habilitarBotonesPersonaJuridica()
Me.cmdAgregarPersonaJuridica.Enabled = True
Me.cmdEditarPersonaJuridica.Enabled = True
Me.cmdEliminarPersonaJuridica.Enabled = True
End Sub

Private Sub deshabilitarBotonesPersonaJuridica()
Me.cmdAgregarPersonaJuridica.Enabled = False
Me.cmdEditarPersonaJuridica.Enabled = False
Me.cmdEliminarPersonaJuridica.Enabled = False
End Sub

Private Sub habilitarBotonesProveedor()
Me.cmdAgregarProveedor.Enabled = True
Me.cmdEditarProveedor.Enabled = True
Me.cmdEliminarProveedor.Enabled = True
End Sub

Private Sub deshabilitarBotonesProveedor()
Me.cmdAgregarProveedor.Enabled = False
Me.cmdEditarProveedor.Enabled = False
Me.cmdEliminarProveedor.Enabled = False
End Sub

Private Sub habilitarBotonesCliente()
Me.cmdAgregarCliente.Enabled = True
Me.cmdEditarCliente.Enabled = True
Me.cmdEliminarCliente.Enabled = True
End Sub

Private Sub deshabilitarBotonesCliente()
Me.cmdAgregarCliente.Enabled = False
Me.cmdEditarCliente.Enabled = False
Me.cmdEliminarCliente.Enabled = False
End Sub






Private Sub Form_Activate()
'Call Restingir
If fsMotivoRegistro = "PEPS" Then
    If fsEstadoCivil = "2" Then
        Me.tabClienteSensible.Tab = 0
        If Me.TxtApellidosConyuge.Enabled And Me.TxtApellidosConyuge.Visible Then
            Me.TxtApellidosConyuge.SetFocus
        End If
    Else
        Me.tabClienteSensible.Tab = 1
        If Me.FERelacionCliente.Visible Then
            Me.FERelacionCliente.SetFocus
        End If
    End If
Else
    Me.tabClienteSensible.Tab = 0

    If Me.TxtCentroLaboralCliente.Enabled And Me.TxtCentroLaboralCliente.Visible Then
        TxtCentroLaboralCliente.SetFocus
    End If
End If

End Sub


Private Sub Form_Load()
CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Call cargarControles
Call limpiarDatosCliente
Call limpiarDatosConyuge
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If fnCodigoClienteProcesoReforzado = 0 Then
    MsgBox ("Aún no registro el Cliente de Procedimiento Reforzado")
    Cancel = 1
    End If
End Sub

Private Sub TxtBuscarCliente_EmiteDatos()
    If Me.TxtBuscarCliente.Text = "" Then
       MsgBox ("Aún no a buscado a una Cliente")
       Exit Sub
    Else
       Dim lrsDatosPrincipales As New ADODB.Recordset
       Set oParentescoCliente = New COMNPersona.NCOMPersona
       Set lrsDatosPrincipales = oParentescoCliente.mostrarDatosPrincipalesClienteProcesoReforzado(TxtBuscarCliente.Text)
             
       If lrsDatosPrincipales.RecordCount > 0 Then
           Me.TxtNombreCliente.Text = lrsDatosPrincipales.Fields(0)
           Me.TxtDniCliente.Text = lrsDatosPrincipales.Fields(1)
           Me.TxtCanetExtranjeriaCliente.Text = lrsDatosPrincipales.Fields(2)
           Me.TxtPasaporteCliente.Text = lrsDatosPrincipales.Fields(3)
           fsCodigoPersona = Me.TxtBuscarCliente.Text
        Else
        MsgBox "No se encontro la persona", vbInformation, "Aviso"
        TxtBuscarCliente.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub cmdAgregarRelacionCliente_Click()
Me.cmdGuardar.Enabled = False
'Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.TxtBuscarCliente.Enabled = False
Me.cmdAgregarRelacionCliente.Enabled = False
Me.cmdEditarRelacionCliente.Enabled = False
Me.cmdEliminarRelacionCliente.Enabled = False
Me.FERelacionCliente.lbEditarFlex = True
Me.FERelacionCliente.AdicionaFila
FilaNoEditar = Me.FERelacionCliente.Rows - 1
iAccion = 1
Me.cmdAceptarRelacionCliente.Enabled = True
Me.cmdCancelarRelacionCliente.Enabled = True
FERelacionCliente.SetFocus
End Sub

Private Sub cmdEditarRelacionCliente_Click()
Me.cmdGuardar.Enabled = False
'Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.TxtBuscarCliente.Enabled = False
Me.cmdAgregarRelacionCliente.Enabled = False
Me.cmdEditarRelacionCliente.Enabled = False
Me.cmdEliminarRelacionCliente.Enabled = False
Me.FERelacionCliente.lbEditarFlex = True
FilaNoEditar = Me.FERelacionCliente.row
iAccion = 2
Me.cmdAceptarRelacionCliente.Enabled = True
Me.cmdCancelarRelacionCliente.Enabled = True
FERelacionCliente.SetFocus
End Sub

Private Sub cmdEliminarRelacionCliente_Click()
  If MsgBox("Esta Seguro que desea eliminar la persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarParientesClientes(Me.FERelacionCliente.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarRelacionCliente_Click()
Dim i As Integer
If validarParentescoCliente = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FERelacionCliente.row
Call actualizarNumeroParienteCliente(Me.FERelacionCliente.Rows)
Call adicionarParienteCliente
Call actualizarRelacionIdTitular(CInt(Trim(Right(Me.FERelacionCliente.TextMatrix(i, 1), 10))), i)
Call actualizarRelacionTitular(Me.FERelacionCliente.TextMatrix(i, 1), i)
Call actualizarNombreCompletoTitular(Me.FERelacionCliente.TextMatrix(i, 2), i)
Call actualizarDniTitular(Me.FERelacionCliente.TextMatrix(i, 3), i)
End If

If iAccion = 2 Then
i = Me.FERelacionCliente.row
Call actualizarRelacionIdTitular(CInt(Trim(Right(Me.FERelacionCliente.TextMatrix(i, 1), 10))), i)
Call actualizarRelacionTitular(Me.FERelacionCliente.TextMatrix(i, 1), i)
Call actualizarNombreCompletoTitular(Me.FERelacionCliente.TextMatrix(i, 2), i)
Call actualizarDniTitular(Me.FERelacionCliente.TextMatrix(i, 3), i)
End If

Me.cmdAgregarRelacionCliente.Enabled = True
Me.cmdEditarRelacionCliente.Enabled = True
Me.cmdEliminarRelacionCliente.Enabled = True
Me.FERelacionCliente.lbEditarFlex = False
Me.cmdAceptarRelacionCliente.Enabled = False
Me.cmdCancelarRelacionCliente.Enabled = False
Me.TxtBuscarCliente.Enabled = True
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarRelacionCliente_Click()
Call cargarParientesClientes
Me.cmdAgregarRelacionCliente.Enabled = True
Me.cmdEditarRelacionCliente.Enabled = True
Me.cmdEliminarRelacionCliente.Enabled = True
Me.FERelacionCliente.lbEditarFlex = False
Me.cmdAceptarRelacionCliente.Enabled = False
Me.cmdCancelarRelacionCliente.Enabled = False
Me.TxtBuscarCliente.Enabled = True
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub


Private Sub cmdAgregarRelacionConyuge_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
'Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarRelacionConyuge.Enabled = False
Me.cmdEditarRelacionConyuge.Enabled = False
Me.cmdEliminarRelacionConyuge.Enabled = False
Me.FERelacionConyuge.lbEditarFlex = True
Me.FERelacionConyuge.AdicionaFila
FilaNoEditar = Me.FERelacionConyuge.Rows - 1
iAccion = 1
Me.cmdAceptarRelacionConyuge.Enabled = True
Me.cmdCancelarRelacionConyuge.Enabled = True
FERelacionConyuge.SetFocus
End Sub

Private Sub cmdEditarRelacionConyuge_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
'Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarRelacionConyuge.Enabled = False
Me.cmdEditarRelacionConyuge.Enabled = False
Me.cmdEliminarRelacionConyuge.Enabled = False
Me.FERelacionConyuge.lbEditarFlex = True
FilaNoEditar = Me.FERelacionConyuge.row
iAccion = 2
Me.cmdAceptarRelacionConyuge.Enabled = True
Me.cmdCancelarRelacionConyuge.Enabled = True
FERelacionConyuge.SetFocus
End Sub

Private Sub cmdEliminarRelacionConyuge_Click()
  If MsgBox("Esta Seguro que desea eliminar la persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarParientesConyuges(Me.FERelacionConyuge.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarRelacionConyuge_Click()
Dim i As Integer
If validarParentescoConyuge = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FERelacionConyuge.row
Call actualizarNumeroParienteConyuge(Me.FERelacionConyuge.Rows)
Call adicionarParienteConyuge
Call actualizarRelacionIdConyuge(CInt(Trim(Right(Me.FERelacionConyuge.TextMatrix(i, 1), 10))), i)
Call actualizarRelacionConyuge(Me.FERelacionConyuge.TextMatrix(i, 1), i)
Call actualizarNombreCompletoConyuge(Me.FERelacionConyuge.TextMatrix(i, 2), i)
Call actualizarDniConyuge(Me.FERelacionConyuge.TextMatrix(i, 3), i)
End If

If iAccion = 2 Then
i = Me.FERelacionConyuge.row
Call actualizarRelacionIdConyuge(CInt(Trim(Right(Me.FERelacionConyuge.TextMatrix(i, 1), 10))), i)
Call actualizarRelacionConyuge(Me.FERelacionConyuge.TextMatrix(i, 1), i)
Call actualizarNombreCompletoConyuge(Me.FERelacionConyuge.TextMatrix(i, 2), i)
Call actualizarDniConyuge(Me.FERelacionConyuge.TextMatrix(i, 3), i)
End If

Me.cmdAgregarRelacionConyuge.Enabled = True
Me.cmdEditarRelacionConyuge.Enabled = True
Me.cmdEliminarRelacionConyuge.Enabled = True
Me.FERelacionConyuge.lbEditarFlex = False
Me.cmdAceptarRelacionConyuge.Enabled = False
Me.cmdCancelarRelacionConyuge.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarRelacionConyuge_Click()
Call cargarParientesConyuges
Me.cmdAgregarRelacionConyuge.Enabled = True
Me.cmdEditarRelacionConyuge.Enabled = True
Me.cmdEliminarRelacionConyuge.Enabled = True
Me.FERelacionConyuge.lbEditarFlex = False
Me.cmdAceptarRelacionConyuge.Enabled = False
Me.cmdCancelarRelacionConyuge.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdAgregarReferenciaEconomica_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
'Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarReferenciaEconomica.Enabled = False
Me.cmdEditarReferenciaEconomica.Enabled = False
Me.cmdEliminarReferenciaEconomica.Enabled = False
Me.FEReferenciaEconomica.lbEditarFlex = True
Me.FEReferenciaEconomica.AdicionaFila
FilaNoEditar = Me.FEReferenciaEconomica.Rows - 1
iAccion = 1
Me.cmdAceptarReferenciaEconomica.Enabled = True
Me.cmdCancelarReferenciaEconomica.Enabled = True
FEReferenciaEconomica.SetFocus
End Sub

Private Sub cmdEditarReferenciaEconomica_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
'Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarReferenciaEconomica.Enabled = False
Me.cmdEditarReferenciaEconomica.Enabled = False
Me.cmdEliminarReferenciaEconomica.Enabled = False
Me.FEReferenciaEconomica.lbEditarFlex = True
FilaNoEditar = Me.FEReferenciaEconomica.row
iAccion = 2
Me.cmdAceptarReferenciaEconomica.Enabled = True
Me.cmdCancelarReferenciaEconomica.Enabled = True
FEReferenciaEconomica.SetFocus
End Sub

Private Sub cmdEliminarReferenciaEconomica_Click()
  If MsgBox("Esta Seguro que desea eliminar la referencia económica", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarReferenciasEconomicas(Me.FEReferenciaEconomica.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarReferenciaEconomica_Click()
Dim i As Integer
If validarReferenciaEconomica = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FEReferenciaEconomica.row
Call actualizarNumeroReferenciaEconomica(Me.FEReferenciaEconomica.Rows)
Call adicionarReferenciaEconomica
Call actualizarCargoPublicoReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 1), i)
Call actualizarFechaInicioReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 2), i)
Call actualizarFechaCeseReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 3), i)
Call actualizarEntidadLaboraReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 4), i)
Call actualizarIngresoReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 5), i)
End If

If iAccion = 2 Then
i = Me.FEReferenciaEconomica.row
Call actualizarCargoPublicoReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 1), i)
Call actualizarFechaInicioReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 2), i)
Call actualizarFechaCeseReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 3), i)
Call actualizarEntidadLaboraReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 4), i)
Call actualizarIngresoReferenciaEconomica(Me.FEReferenciaEconomica.TextMatrix(i, 5), i)
End If

Me.cmdAgregarReferenciaEconomica.Enabled = True
Me.cmdEditarReferenciaEconomica.Enabled = True
Me.cmdEliminarReferenciaEconomica.Enabled = True
Me.FEReferenciaEconomica.lbEditarFlex = False
Me.cmdAceptarReferenciaEconomica.Enabled = False
Me.cmdCancelarReferenciaEconomica.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarReferenciaEconomica_Click()
Call cargarReferenciasEconomicas
Me.cmdAgregarReferenciaEconomica.Enabled = True
Me.cmdEditarReferenciaEconomica.Enabled = True
Me.cmdEliminarReferenciaEconomica.Enabled = True
Me.FEReferenciaEconomica.lbEditarFlex = False
Me.cmdAceptarReferenciaEconomica.Enabled = False
Me.cmdCancelarReferenciaEconomica.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub


Private Sub cmdAgregarReferenciaFinanciera_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
'Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarReferenciaFinanciera.Enabled = False
Me.cmdEditarReferenciaFinanciera.Enabled = False
Me.cmdEliminarReferenciaFinanciera.Enabled = False
Me.FEReferenciaFinanciera.lbEditarFlex = True
Me.FEReferenciaFinanciera.AdicionaFila
FilaNoEditar = Me.FEReferenciaFinanciera.Rows - 1
iAccion = 1
Me.cmdAceptarReferenciaFinanciera.Enabled = True
Me.cmdCancelarReferenciaFinanciera.Enabled = True
FEReferenciaFinanciera.SetFocus
End Sub

Private Sub cmdEditarReferenciaFinanciera_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
'Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarReferenciaFinanciera.Enabled = False
Me.cmdEditarReferenciaFinanciera.Enabled = False
Me.cmdEliminarReferenciaFinanciera.Enabled = False
Me.FEReferenciaFinanciera.lbEditarFlex = True
FilaNoEditar = Me.FEReferenciaFinanciera.row
iAccion = 2
Me.cmdAceptarReferenciaFinanciera.Enabled = True
Me.cmdCancelarReferenciaFinanciera.Enabled = True
FEReferenciaFinanciera.SetFocus
End Sub

Private Sub cmdEliminarReferenciaFinanciera_Click()
  If MsgBox("Esta Seguro que desea eliminar la referencia finaciera", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarReferenciasFinancieras(Me.FEReferenciaFinanciera.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarReferenciaFinanciera_Click()
Dim i As Integer
If validarReferenciaFinanciera = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FEReferenciaFinanciera.row
Call actualizarNumeroReferenciaFinanciera(Me.FEReferenciaFinanciera.Rows)
Call adicionarReferenciaFinanciera
Call actualizarCodigoBancoReferenciaFinanciera(Trim(Right(Me.FEReferenciaFinanciera.TextMatrix(i, 1), 15)), i)
Call actualizarNombreBancoReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(i, 1), i)
Call actualizarTipoProductoReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(i, 2), i)
Call actualizarFuncionarioNegocioReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(i, 3), i)
End If

If iAccion = 2 Then
i = Me.FEReferenciaFinanciera.row
Call actualizarCodigoBancoReferenciaFinanciera(Trim(Right(Me.FEReferenciaFinanciera.TextMatrix(i, 1), 15)), i)
Call actualizarNombreBancoReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(i, 1), i)
Call actualizarTipoProductoReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(i, 2), i)
Call actualizarFuncionarioNegocioReferenciaFinanciera(Me.FEReferenciaFinanciera.TextMatrix(i, 3), i)
End If

Me.cmdAgregarReferenciaFinanciera.Enabled = True
Me.cmdEditarReferenciaFinanciera.Enabled = True
Me.cmdEliminarReferenciaFinanciera.Enabled = True
Me.FEReferenciaFinanciera.lbEditarFlex = False
Me.cmdAceptarReferenciaFinanciera.Enabled = False
Me.cmdCancelarReferenciaFinanciera.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarReferenciaFinanciera_Click()
Call cargarReferenciasFinancieras
Me.cmdAgregarReferenciaFinanciera.Enabled = True
Me.cmdEditarReferenciaFinanciera.Enabled = True
Me.cmdEliminarReferenciaFinanciera.Enabled = True
Me.FEReferenciaFinanciera.lbEditarFlex = False
Me.cmdAceptarReferenciaFinanciera.Enabled = False
Me.cmdCancelarReferenciaFinanciera.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdAgregarReferenciaPatrimonial_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
'Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarReferenciaPatrimonial.Enabled = False
Me.cmdEditarReferenciaPatrimonial.Enabled = False
Me.cmdEliminarReferenciaPatrimonial.Enabled = False
Me.FEReferenciaPatrimonial.lbEditarFlex = True
Me.FEReferenciaPatrimonial.AdicionaFila
FilaNoEditar = Me.FEReferenciaPatrimonial.Rows - 1
iAccion = 1
Me.cmdAceptarReferenciaPatrimonial.Enabled = True
Me.cmdCancelarReferenciaPatrimonial.Enabled = True
FEReferenciaPatrimonial.SetFocus
End Sub

Private Sub cmdEditarReferenciaPatrimonial_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
'Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarReferenciaPatrimonial.Enabled = False
Me.cmdEditarReferenciaPatrimonial.Enabled = False
Me.cmdEliminarReferenciaPatrimonial.Enabled = False
Me.FEReferenciaPatrimonial.lbEditarFlex = True
FilaNoEditar = Me.FEReferenciaPatrimonial.row
iAccion = 2
Me.cmdAceptarReferenciaPatrimonial.Enabled = True
Me.cmdCancelarReferenciaPatrimonial.Enabled = True
FEReferenciaPatrimonial.SetFocus
End Sub

Private Sub cmdEliminarReferenciaPatrimonial_Click()
  If MsgBox("Esta Seguro que desea eliminar la referencia patrimonial", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarReferenciasPatrimoniales(Me.FEReferenciaPatrimonial.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarReferenciaPatrimonial_Click()
Dim i As Integer
If validarReferenciaPatrimonial = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FEReferenciaPatrimonial.row
Call actualizarNumeroReferenciaPatrimonial(Me.FEReferenciaPatrimonial.Rows)
Call adicionarReferenciaPatrimonial
Call actualizarBienReferenciaPatrimonial(Me.FEReferenciaPatrimonial.TextMatrix(i, 1), i)
Call actualizarValorReferenciaPatrimonial(Me.FEReferenciaPatrimonial.TextMatrix(i, 2), i)
End If

If iAccion = 2 Then
i = Me.FEReferenciaPatrimonial.row
Call actualizarBienReferenciaPatrimonial(Me.FEReferenciaPatrimonial.TextMatrix(i, 1), i)
Call actualizarValorReferenciaPatrimonial(Me.FEReferenciaPatrimonial.TextMatrix(i, 2), i)
End If

Me.cmdAgregarReferenciaPatrimonial.Enabled = True
Me.cmdEditarReferenciaPatrimonial.Enabled = True
Me.cmdEliminarReferenciaPatrimonial.Enabled = True
Me.FEReferenciaPatrimonial.lbEditarFlex = False
Me.cmdAceptarReferenciaPatrimonial.Enabled = False
Me.cmdCancelarReferenciaPatrimonial.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarReferenciaPatrimonial_Click()
Call cargarReferenciasPatrimoniales
Me.cmdAgregarReferenciaPatrimonial.Enabled = True
Me.cmdEditarReferenciaPatrimonial.Enabled = True
Me.cmdEliminarReferenciaPatrimonial.Enabled = True
Me.FEReferenciaPatrimonial.lbEditarFlex = False
Me.cmdAceptarReferenciaPatrimonial.Enabled = False
Me.cmdCancelarReferenciaPatrimonial.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdAgregarOtroBien_Click()
Me.cmdAgregarOtroBien.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
'Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdEditarOtroBien.Enabled = False
Me.cmdEliminarOtroBien.Enabled = False
Me.FEOtroBien.lbEditarFlex = True
Me.FEOtroBien.AdicionaFila
FilaNoEditar = Me.FEOtroBien.Rows - 1
iAccion = 1
Me.cmdAceptarOtroBien.Enabled = True
Me.cmdCancelarOtroBien.Enabled = True
FEOtroBien.SetFocus
End Sub

Private Sub cmdEditarOtroBien_Click()
Me.cmdAgregarOtroBien.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
'Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdEditarOtroBien.Enabled = False
Me.cmdEliminarOtroBien.Enabled = False
Me.FEOtroBien.lbEditarFlex = True
FilaNoEditar = Me.FEOtroBien.row
iAccion = 2
Me.cmdAceptarOtroBien.Enabled = True
Me.cmdCancelarOtroBien.Enabled = True
FEOtroBien.SetFocus
End Sub

Private Sub cmdEliminarOtroBien_Click()
  If MsgBox("Esta Seguro que desea eliminar el bien", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarOtrosBienes(Me.FEOtroBien.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarOtroBien_Click()
Dim i As Integer
If validarOtroBien = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FEOtroBien.row
Call actualizarNumeroOtroBien(Me.FEOtroBien.Rows)
Call adicionarOtroBien
Call actualizarOtroBien(Me.FEOtroBien.TextMatrix(i, 1), i)

End If

If iAccion = 2 Then
i = Me.FEOtroBien.row
Call actualizarOtroBien(Me.FEOtroBien.TextMatrix(i, 1), i)
End If

Me.cmdAgregarOtroBien.Enabled = True
Me.cmdEditarOtroBien.Enabled = True
Me.cmdEliminarOtroBien.Enabled = True
Me.FEOtroBien.lbEditarFlex = False
Me.cmdAceptarOtroBien.Enabled = False
Me.cmdCancelarOtroBien.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarOtroBien_Click()
Call cargarOtrosBienes
Me.cmdAgregarOtroBien.Enabled = True
Me.cmdEditarOtroBien.Enabled = True
Me.cmdEliminarOtroBien.Enabled = True
Me.FEOtroBien.lbEditarFlex = False
Me.cmdAceptarOtroBien.Enabled = False
Me.cmdCancelarOtroBien.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdAgregarPersonaJuridica_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
'Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarPersonaJuridica.Enabled = False
Me.cmdEditarPersonaJuridica.Enabled = False
Me.cmdEliminarPersonaJuridica.Enabled = False
Me.FEPersonaJuridica.lbEditarFlex = True
Me.FEPersonaJuridica.AdicionaFila
FilaNoEditar = Me.FEPersonaJuridica.Rows - 1
iAccion = 1
Me.cmdAceptarPersonaJuridica.Enabled = True
Me.cmdCancelarPersonaJuridica.Enabled = True
FEPersonaJuridica.SetFocus
End Sub

Private Sub cmdEditarPersonaJuridica_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
'Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarPersonaJuridica.Enabled = False
Me.cmdEditarPersonaJuridica.Enabled = False
Me.cmdEliminarPersonaJuridica.Enabled = False
Me.FEPersonaJuridica.lbEditarFlex = True
FilaNoEditar = Me.FEPersonaJuridica.row
iAccion = 2
Me.cmdAceptarPersonaJuridica.Enabled = True
Me.cmdCancelarPersonaJuridica.Enabled = True
FEPersonaJuridica.SetFocus
End Sub

Private Sub cmdEliminarPersonaJuridica_Click()
  If MsgBox("Esta Seguro que desea eliminar la persona jurídica", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarPersonasJuridicas(Me.FEPersonaJuridica.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarPersonaJuridica_Click()
Dim i As Integer
If validarPersonaJuridica = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FEPersonaJuridica.row
Call actualizarNumeroPersonaJuridica(Me.FEPersonaJuridica.Rows)
Call adicionarPersonaJuridica
Call actualizarNombreEmpresaPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 1), i)
Call actualizarRucPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 2), i)
Call actualizarParticipacionPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 3), i)
Call actualizarIngresoPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 4), i)
End If

If iAccion = 2 Then
i = Me.FEPersonaJuridica.row
Call actualizarNombreEmpresaPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 1), i)
Call actualizarRucPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 2), i)
Call actualizarParticipacionPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 3), i)
Call actualizarIngresoPersonaJuridica(Me.FEPersonaJuridica.TextMatrix(i, 4), i)
End If

Me.cmdAgregarPersonaJuridica.Enabled = True
Me.cmdEditarPersonaJuridica.Enabled = True
Me.cmdEliminarPersonaJuridica.Enabled = True
Me.FEPersonaJuridica.lbEditarFlex = False
Me.cmdAceptarPersonaJuridica.Enabled = False
Me.cmdCancelarPersonaJuridica.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarPersonaJuridica_Click()
Call cargarPersonasJuridicas
Me.cmdAgregarPersonaJuridica.Enabled = True
Me.cmdEditarPersonaJuridica.Enabled = True
Me.cmdEliminarPersonaJuridica.Enabled = True
Me.FEPersonaJuridica.lbEditarFlex = False
Me.cmdAceptarPersonaJuridica.Enabled = False
Me.cmdCancelarPersonaJuridica.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdAgregarProveedor_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
'Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarProveedor.Enabled = False
Me.cmdEditarProveedor.Enabled = False
Me.cmdEliminarProveedor.Enabled = False
Me.FEProveedor.lbEditarFlex = True
Me.FEProveedor.AdicionaFila
FilaNoEditar = Me.FEProveedor.Rows - 1
iAccion = 1
Me.cmdAceptarProveedor.Enabled = True
Me.cmdCancelarProveedor.Enabled = True
FEProveedor.SetFocus
End Sub

Private Sub cmdEditarProveedor_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
'Call deshabilitarBotonesProveedor
Call deshabilitarBotonesCliente
Me.cmdAgregarProveedor.Enabled = False
Me.cmdEditarProveedor.Enabled = False
Me.cmdEliminarProveedor.Enabled = False
Me.FEProveedor.lbEditarFlex = True
FilaNoEditar = Me.FEProveedor.row
iAccion = 2
Me.cmdAceptarProveedor.Enabled = True
Me.cmdCancelarProveedor.Enabled = True
FEProveedor.SetFocus
End Sub

Private Sub cmdEliminarProveedor_Click()
  If MsgBox("Esta Seguro que desea eliminar el proveedor", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarProveedores(Me.FEProveedor.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarProveedor_Click()
Dim i As Integer
If validarProveedor = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FEProveedor.row
Call actualizarNumeroProveedor(Me.FEProveedor.Rows)
Call adicionarProveedor
Call actualizarNombreEmpresaProveedor(Me.FEProveedor.TextMatrix(i, 1), i)

End If

If iAccion = 2 Then
i = Me.FEProveedor.row
Call actualizarNombreEmpresaProveedor(Me.FEProveedor.TextMatrix(i, 1), i)
End If

Me.cmdAgregarProveedor.Enabled = True
Me.cmdEditarProveedor.Enabled = True
Me.cmdEliminarProveedor.Enabled = True
Me.FEProveedor.lbEditarFlex = False
Me.cmdAceptarProveedor.Enabled = False
Me.cmdCancelarProveedor.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarProveedor_Click()
Call cargarProveedores
Me.cmdAgregarProveedor.Enabled = True
Me.cmdEditarProveedor.Enabled = True
Me.cmdEliminarProveedor.Enabled = True
Me.FEProveedor.lbEditarFlex = False
Me.cmdAceptarProveedor.Enabled = False
Me.cmdCancelarProveedor.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdAgregarCliente_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
'Call deshabilitarBotonesCliente
Me.cmdAgregarCliente.Enabled = False
Me.cmdEditarCliente.Enabled = False
Me.cmdEliminarCliente.Enabled = False
Me.FECliente.lbEditarFlex = True
Me.FECliente.AdicionaFila
FilaNoEditar = Me.FECliente.Rows - 1
iAccion = 1
Me.cmdAceptarCliente.Enabled = True
Me.cmdCancelarCliente.Enabled = True
FECliente.SetFocus
End Sub

Private Sub cmdEditarCliente_Click()
Me.cmdGuardar.Enabled = False
Call deshabilitarBotonesParentescoCliente
Call deshabilitarBotonesParentescoConyuge
Call deshabilitarBotonesReferenciaEconomica
Call deshabilitarBotonesReferenciaFinanciera
Call deshabilitarBotonesReferenciaPatrimonial
Call deshabilitarBotonesOtroBien
Call deshabilitarBotonesPersonaJuridica
Call deshabilitarBotonesProveedor
'Call deshabilitarBotonesCliente
Me.cmdAgregarCliente.Enabled = False
Me.cmdEditarCliente.Enabled = False
Me.cmdEliminarCliente.Enabled = False
Me.FECliente.lbEditarFlex = True
FilaNoEditar = Me.FECliente.row
iAccion = 2
Me.cmdAceptarCliente.Enabled = True
Me.cmdCancelarCliente.Enabled = True
FECliente.SetFocus
End Sub

Private Sub cmdEliminarCliente_Click()
  If MsgBox("Esta Seguro que desea eliminar el cliente", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        iAccion = 3
        Call cargarClientes(Me.FECliente.row)
        FilaNoEditar = -1
    End If
End Sub

Private Sub cmdAceptarCliente_Click()
Dim i As Integer
If validarCliente = False Then
   Exit Sub
End If
If iAccion = 1 Then
i = Me.FECliente.row
Call actualizarNumeroCliente(Me.FECliente.Rows)
Call adicionarCliente
Call actualizarNombreEmpresaCliente(Me.FECliente.TextMatrix(i, 1), i)

End If

If iAccion = 2 Then
i = Me.FECliente.row
Call actualizarNombreEmpresaCliente(Me.FECliente.TextMatrix(i, 1), i)
End If

Me.cmdAgregarCliente.Enabled = True
Me.cmdEditarCliente.Enabled = True
Me.cmdEliminarCliente.Enabled = True
Me.FECliente.lbEditarFlex = False
Me.cmdAceptarCliente.Enabled = False
Me.cmdCancelarCliente.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub

Private Sub cmdCancelarCliente_Click()
Call cargarClientes
Me.cmdAgregarCliente.Enabled = True
Me.cmdEditarCliente.Enabled = True
Me.cmdEliminarCliente.Enabled = True
Me.FECliente.lbEditarFlex = False
Me.cmdAceptarCliente.Enabled = False
Me.cmdCancelarCliente.Enabled = False
FilaNoEditar = -1
Me.cmdGuardar.Enabled = True
Call Restingir
End Sub


Private Sub FERelacionCliente_Click()
Call FERelacionCliente_RowColChange
End Sub

Private Sub FERelacionCliente_EnterCell()
FERelacionCliente_RowColChange
End Sub

Private Sub FERelacionCliente_RowColChange()


 Dim lrsParentescoCliente As ADODB.Recordset

 Set oParentescoCliente = New COMNPersona.NCOMPersona
 Set lrsParentescoCliente = New ADODB.Recordset

 If Me.FERelacionCliente.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FERelacionCliente.row = FilaNoEditar
    End If

    Set lrsParentescoCliente = oParentescoCliente.listarParentescosClienteReforzado("T")

    Select Case Me.FERelacionCliente.Col
           Case 1
                Me.FERelacionCliente.CargaCombo lrsParentescoCliente
    End Select
    Set lrsParentescoCliente = Nothing
    Set oParentescoCliente = Nothing
End If

End Sub

Private Sub FERelacionConyuge_Click()
Call FERelacionConyuge_RowColChange
End Sub

Private Sub FERelacionConyuge_EnterCell()
FERelacionConyuge_RowColChange
End Sub

Private Sub FERelacionConyuge_RowColChange()


 Dim lrsParentescoConyuge As ADODB.Recordset

 Set oParentescoConyuge = New COMNPersona.NCOMPersona
 Set lrsParentescoConyuge = New ADODB.Recordset

 If Me.FERelacionConyuge.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FERelacionConyuge.row = FilaNoEditar
    End If

    Set lrsParentescoConyuge = oParentescoConyuge.listarParentescosClienteReforzado("C")

    Select Case Me.FERelacionConyuge.Col
           Case 1
                Me.FERelacionConyuge.CargaCombo lrsParentescoConyuge
    End Select
    Set lrsParentescoConyuge = Nothing
    Set oParentescoConyuge = Nothing
End If

End Sub

Private Sub FEReferenciaEconomica_Click()
Call FEReferenciaEconomica_RowColChange
End Sub

Private Sub FEReferenciaEconomica_EnterCell()
FEReferenciaEconomica_RowColChange
End Sub

Private Sub FEReferenciaEconomica_RowColChange()

 If Me.FEReferenciaEconomica.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FEReferenciaEconomica.row = FilaNoEditar
    End If

End If

End Sub

Private Sub FEReferenciaFinanciera_Click()
Call FEReferenciaFinanciera_RowColChange
End Sub

Private Sub FEReferenciaFinanciera_EnterCell()
FEReferenciaFinanciera_RowColChange
End Sub

Private Sub FEReferenciaFinanciera_RowColChange()


 Dim lrsBancos As ADODB.Recordset

 Set oParentescoCliente = New COMNPersona.NCOMPersona
 Set lrsBancos = New ADODB.Recordset

 If Me.FEReferenciaFinanciera.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FEReferenciaFinanciera.row = FilaNoEditar
    End If

    Set lrsBancos = oParentescoCliente.listarBancosClienteReforzado

    Select Case Me.FEReferenciaFinanciera.Col
           Case 1
                Me.FEReferenciaFinanciera.CargaCombo lrsBancos
    End Select
    Set lrsBancos = Nothing
    Set oParentescoCliente = Nothing
End If

End Sub

Private Sub FEReferenciaPatrimonial_Click()
Call FEReferenciaPatrimonial_RowColChange
End Sub

Private Sub FEReferenciaPatrimonial_EnterCell()
FEReferenciaPatrimonial_RowColChange
End Sub

Private Sub FEReferenciaPatrimonial_RowColChange()

 If Me.FEReferenciaPatrimonial.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FEReferenciaPatrimonial.row = FilaNoEditar
    End If

End If
End Sub

Private Sub FEOtroBien_Click()
Call FEOtroBien_RowColChange
End Sub

Private Sub FEOtroBien_EnterCell()
FEOtroBien_RowColChange
End Sub

Private Sub FEOtroBien_RowColChange()

 If Me.FEOtroBien.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FEOtroBien.row = FilaNoEditar
    End If

End If
End Sub

Private Sub FEPersonaJuridica_Click()
Call FEPersonaJuridica_RowColChange
End Sub

Private Sub FEPersonaJuridica_EnterCell()
FEPersonaJuridica_RowColChange
End Sub

Private Sub FEPersonaJuridica_RowColChange()

 If Me.FEPersonaJuridica.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FEPersonaJuridica.row = FilaNoEditar
    End If

End If
End Sub

Private Sub FEProveedor_Click()
Call FEProveedor_RowColChange
End Sub

Private Sub FEProveedor_EnterCell()
FEProveedor_RowColChange
End Sub

Private Sub FEProveedor_RowColChange()

 If Me.FEProveedor.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FEProveedor.row = FilaNoEditar
    End If

End If
End Sub

Private Sub FECliente_Click()
Call FECliente_RowColChange
End Sub

Private Sub FECliente_EnterCell()
FECliente_RowColChange
End Sub

Private Sub FECliente_RowColChange()

 If Me.FECliente.lbEditarFlex Then
    If FilaNoEditar <> -1 Then
       Me.FECliente.row = FilaNoEditar
    End If

End If
End Sub



Private Sub cmdGuardar_Click()
    Dim i As Integer
    Dim lrsConyugeId As ADODB.Recordset
    Dim lrsClienteProcesoReforzadoId As ADODB.Recordset
    
On Error GoTo ErrorcmdGuardar
Screen.MousePointer = 11

    If Trim(fsCodigoPersona) = "" Then
        Exit Sub
    End If

    If validarDatosCliente = False Then
        Exit Sub
    End If

    If fsEstadoCivil = "2" Then
        If validarDatosConyuge = False Then
           Exit Sub
        End If
    End If

    If fsMotivoRegistro = "PEPS" And recuperarNumeroParienteCliente < 2 Then
        MsgBox "Debe ingresar los parientes del cliente", vbInformation, "Aviso"
        Me.tabClienteSensible.Tab = 1
        FERelacionCliente.SetFocus
        Exit Sub
    End If

    If fsMotivoRegistro = "PEPS" And fsEstadoCivil = "2" And recuperarNumeroParienteConyuge < 2 Then
        MsgBox "Debe ingresar los parientes del conyuge", vbInformation, "Aviso"
        Me.tabClienteSensible.Tab = 1
        FERelacionConyuge.SetFocus
        Exit Sub
    End If

    If fsMotivoRegistro = "PEPS" And recuperarNumeroReferenciaEconomica < 0 Then 'WIOR 20130121 SE CAMBIO DE LA REFERENCIA ECONOMICA DE 2 A 0
        MsgBox "Debe ingresar las referencias economicas del cliente", vbInformation, "Aviso"
        Me.tabClienteSensible.Tab = 2
        FEReferenciaEconomica.SetFocus
        Exit Sub
    End If

    If recuperarNumeroReferenciaFinanciera < 0 Then 'WIOR 20130121SE CAMBIO DE LA REFERENCIA DE FINANCIERA 2 A 0
        MsgBox "Debe ingresar las referencias financieras del cliente", vbInformation, "Aviso"
        Me.tabClienteSensible.Tab = 3
        FEReferenciaFinanciera.SetFocus
         Exit Sub
    End If

    If recuperarNumeroReferenciaPatrimonial < 2 Then
        MsgBox "Debe ingresar las referencias patrimoniales del cliente", vbInformation, "Aviso"
        Me.tabClienteSensible.Tab = 4
        FEReferenciaPatrimonial.SetFocus
         Exit Sub
    End If

    Set oParentescoCliente = New COMNPersona.NCOMPersona
    Set oParentescoConyuge = New COMNPersona.NCOMPersona
    Set lrsConyugeId = New ADODB.Recordset
    Set lrsClienteProcesoReforzadoId = New ADODB.Recordset



    Call oParentescoCliente.ingresarClienteProcesoReforzado(fsCodigoPersona, _
                                                            Me.TxtCentroLaboralCliente, _
                                                            CStr(gdFecSis), _
                                                            fcMovPersonaAtendido, _
                                                            fcMovPersonaAutorizado, _
                                                            fsMotivoRegistro)

    Set lrsClienteProcesoReforzadoId = oParentescoCliente.mostarLlaveClienteProcesoReforzado(fsCodigoPersona)

    fnCodigoClienteProcesoReforzado = lrsClienteProcesoReforzadoId.Fields(0)
    
    lrsClienteProcesoReforzadoId.Close: Set lrsClienteProcesoReforzadoId = Nothing
  
    If fnCodigoClienteProcesoReforzado = 0 Then
        Exit Sub
    End If

    If fsEstadoCivil = "2" Then
        Call oParentescoConyuge.ingresarConyugeClienteReforzado(fnCodigoClienteProcesoReforzado, _
                                                                Me.TxtNombresConyuge.Text, _
                                                                Me.TxtApellidosConyuge.Text, _
                                                                CInt(Trim(Left(Right(Me.cboDOIConyuge.Text, 15), 10))), _
                                                                CInt(Trim(Right(Me.cboDOIConyuge.Text, 3))), _
                                                                Me.TxtNumeroDoiConyuge.Text, _
                                                                Trim(Right(Me.cboNacionalidadConyuge.Text, 10)), _
                                                                Me.TxtFechaNacimientoConyuge.Text, _
                                                                Me.TxtCentroLaborConyuge.Text, _
                                                                CInt(Trim(Right(Me.cboOcupacionConyuge.Text, 10))), _
                                                                CDbl(Me.txtIngresoPromedioConyuge.Text))
                                                                
        Set lrsConyugeId = oParentescoConyuge.mostarLlaveConyuge(fnCodigoClienteProcesoReforzado)
    End If

    If fsMotivoRegistro = "PEPS" Then
        If recuperarNumeroParienteCliente >= 2 And _
            Trim(Me.FERelacionCliente.TextMatrix(1, 0)) <> "" Then
            For i = 1 To recuperarNumeroParienteCliente - 1
                Call oParentescoCliente.ingresarParentescoClienteReforzado(fnCodigoClienteProcesoReforzado, _
                                                                           recuperarRelacionIdTitular(i), _
                                                                           recuperarNombreCompletoTitular(i), _
                                                                           recuperarDniTitular(i))
            Next i
        End If
    
       If fsEstadoCivil = "2" Then
            If recuperarNumeroParienteConyuge >= 2 And _
                Trim(Me.FERelacionConyuge.TextMatrix(1, 0)) <> "" Then
                For i = 1 To recuperarNumeroParienteConyuge - 1
                    Call oParentescoConyuge.ingresarParentescoConyuge(lrsConyugeId.Fields(0), _
                                                                      recuperarRelacionIdConyuge(i), _
                                                                      recuperarNombreCompletoConyuge(i), _
                                                                      recuperarDniConyuge(i))
                Next i
            End If
            lrsConyugeId.Close: Set lrsConyugeId = Nothing
        End If
    
        If recuperarNumeroReferenciaEconomica >= 2 And _
            Trim(Me.FEReferenciaEconomica.TextMatrix(1, 0)) <> "" Then
            For i = 1 To recuperarNumeroReferenciaEconomica - 1
                Call oParentescoCliente.ingresarReferenciaEconomica(fnCodigoClienteProcesoReforzado, _
                                                                    recuperarCargoPublicoReferenciaEconomica(i), _
                                                                    recuperarFechaInicioReferenciaEconomica(i), _
                                                                    recuperarFechaCeseReferenciaEconomica(i), _
                                                                    recuperarEntidadLaboraReferenciaEconomica(i), _
                                                                    recuperarIngresoReferenciaEconomica(i))
                                                         
                                                  
             Next i
        End If
      
    End If

If recuperarNumeroReferenciaFinanciera >= 2 And _
   Trim(Me.FEReferenciaFinanciera.TextMatrix(1, 0)) <> "" Then
   For i = 1 To recuperarNumeroReferenciaFinanciera - 1
       Call oParentescoCliente.ingresarReferenciaFinanciera(fnCodigoClienteProcesoReforzado, _
                                                            recuperarCodigoBancoReferenciaFinanciera(i), _
                                                            recuperarTipoProductoReferenciaFinanciera(i), _
                                                            recuperarFuncionarioNegocioReferenciaFinanciera(i))
                                                  
                                           
   Next i
End If

If recuperarNumeroReferenciaPatrimonial >= 2 And _
   Trim(Me.FEReferenciaPatrimonial.TextMatrix(1, 0)) <> "" Then
   For i = 1 To recuperarNumeroReferenciaPatrimonial - 1
       Call oParentescoCliente.ingresarReferenciaPatrimonial(fnCodigoClienteProcesoReforzado, _
                                                             recuperarBienReferenciaPatrimonial(i), _
                                                             recuperarValorReferenciaPatrimonial(i))
                                                  
                                       
   Next i
End If

If recuperarNumeroOtroBien >= 2 And Trim(Me.FEOtroBien.TextMatrix(1, 0)) <> "" Then
   For i = 1 To recuperarNumeroOtroBien - 1
       Call oParentescoCliente.ingresarOtroBien(fnCodigoClienteProcesoReforzado, _
                                                recuperarOtroBien(i))
                                                  
          
   Next i
End If

If recuperarNumeroPersonaJuridica >= 2 And Trim(Me.FEPersonaJuridica.TextMatrix(1, 0)) <> "" Then
   For i = 1 To recuperarNumeroPersonaJuridica - 1
       Call oParentescoCliente.ingresarPersonaJuridica(fnCodigoClienteProcesoReforzado, _
                                                       recuperarNombreEmpresaPersonaJuridica(i), _
                                                       recuperarRucPersonaJuridica(i), _
                                                       recuperarParticipacionPersonaJuridica(i), _
                                                       recuperarIngresoPersonaJuridica(i))
                                                  
                                           
   Next i
End If

If recuperarNumeroProveedor >= 2 And Trim(Me.FEProveedor.TextMatrix(1, 0)) <> "" Then
   For i = 1 To recuperarNumeroProveedor - 1
       Call oParentescoCliente.ingresarProveedorCliente(fnCodigoClienteProcesoReforzado, _
                                                        "P", _
                                                        recuperarNombreEmpresaProveedor(i))
                                                  
                                           
    Next i
End If

If recuperarNumeroCliente >= 2 And Trim(Me.FECliente.TextMatrix(1, 0)) <> "" Then
   For i = 1 To recuperarNumeroCliente - 1
       Call oParentescoCliente.ingresarProveedorCliente(fnCodigoClienteProcesoReforzado, _
                                                        "C", _
                                                        recuperarNombreEmpresaCliente(i))
                                                  
                                           
    Next i
End If
Set oParentescoCliente = Nothing
Set oParentescoConyuge = Nothing
'Call ImprimeAnexo1 'marg 20160719
'Call ImprimeAnexo2 'marg 20160719
MsgBox "Se guardaron correctamente los datos del Cliente de Procedimiento Reforzado", vbInformation, "Aviso"
Unload Me
Exit Sub

ErrorcmdGuardar:
Screen.MousePointer = 0
MsgBox Err.Description & "frmPersClienteSensivle: Ocurrio un error intentar guardar el Cliente de Procedimiento Reforzado, por favor comuniquese con el Área de TI", vbExclamation, "Aviso"

End Sub

Private Sub TxtCentroLaboralCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        If fsEstadoCivil = "2" Then
            TxtApellidosConyuge.SetFocus
        Else
            If fsMotivoRegistro = "PEPS" Then
                Me.tabClienteSensible.Tab = 1
                Me.FERelacionCliente.SetFocus
            Else
                Me.tabClienteSensible.Tab = 3
                Me.FEReferenciaFinanciera.SetFocus
            End If
        End If
        
    End If
End Sub
Private Sub TxtApellidosConyuge_KeyPress(KeyAscii As Integer)
  KeyAscii = Letras(KeyAscii, True)
  KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtNombresConyuge.SetFocus
    End If
End Sub

Private Sub TxtNombresConyuge_KeyPress(KeyAscii As Integer)
  KeyAscii = Letras(KeyAscii, True)
  KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    If KeyAscii = 13 Then
        cboDOIConyuge.SetFocus
    End If
End Sub

Private Sub cboDOIConyuge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        TxtNumeroDoiConyuge.SetFocus
    End If
End Sub

Private Sub TxtNumeroDoiConyuge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        cboNacionalidadConyuge.SetFocus
    End If
End Sub

Private Sub cboNacionalidadConyuge_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        TxtFechaNacimientoConyuge.SetFocus
    End If
End Sub

Private Sub TxtFechaNacimientoConyuge_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        TxtCentroLaborConyuge.SetFocus
    End If
End Sub

Private Sub TxtCentroLaborConyuge_KeyPress(KeyAscii As Integer)
  KeyAscii = Letras(KeyAscii, True)
 If KeyAscii = 13 Then
        txtIngresoPromedioConyuge.SetFocus
    End If
End Sub

Private Sub txtIngresoPromedioConyuge_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtIngresoPromedioConyuge, KeyAscii)

 If KeyAscii = 13 Then
        cboOcupacionConyuge.SetFocus
 End If
End Sub

Private Sub cboOcupacionConyuge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fsMotivoRegistro = "PEPS" Then
             Me.tabClienteSensible.Tab = 1
             Me.FERelacionCliente.SetFocus
         Else
             Me.tabClienteSensible.Tab = 3
             Me.FEReferenciaFinanciera.SetFocus
         End If
    End If
End Sub

Private Sub ImprimeAnexo1()
Dim lrsDatosCliente As New ADODB.Recordset
Dim lrsDatosConyuge As New ADODB.Recordset
Dim lrsRefernciaEconomica As New ADODB.Recordset
Dim lrsRefernciaFinanciera As New ADODB.Recordset
Dim lrsRefernciaPatrimonial As New ADODB.Recordset
Dim lrsPersonaJuridica As New ADODB.Recordset
Dim lrsProveedor As New ADODB.Recordset
Dim lrsCliente As New ADODB.Recordset
Dim lrsParienteCliente As New ADODB.Recordset
Dim lrsConyugeKey  As New ADODB.Recordset
Dim lrsParienteConyuge As New ADODB.Recordset
Dim lrsRelacionEntidad As New ADODB.Recordset
Dim lsPlantilla, lsTipoDoiConyuge, lsTipoDoiConyuge2, lsTipoDoiConyuge3 As String

Set oParentescoCliente = New COMNPersona.NCOMPersona
Set oParentescoConyuge = New COMNPersona.NCOMPersona

On Error GoTo ErrGeneraRepo
Screen.MousePointer = 11


Set lrsDatosCliente = oParentescoCliente.mostrarDatosClienteProcesoReforzado(fnCodigoClienteProcesoReforzado)
Set lrsDatosConyuge = oParentescoConyuge.mostrarDatosConyugeClienteProcesoReforzado(fnCodigoClienteProcesoReforzado)
Set lrsRefernciaEconomica = oParentescoCliente.mostrarReferenciaEconomicaClienteProcesoReforzado(fnCodigoClienteProcesoReforzado)
Set lrsRefernciaFinanciera = oParentescoCliente.mostrarReferenciaFinancieraClienteProcesoReforzado(fnCodigoClienteProcesoReforzado)
Set lrsRefernciaPatrimonial = oParentescoCliente.mostrarReferenciaPatrimonialClienteProcesoReforzado(fnCodigoClienteProcesoReforzado, "P")
Set lrsPersonaJuridica = oParentescoCliente.mostrarPersonaJuridicaClienteProcesoReforzado(fnCodigoClienteProcesoReforzado)
Set lrsProveedor = oParentescoCliente.mostrarProveedorClienteProcesoReforzado(fnCodigoClienteProcesoReforzado, "P")
Set lrsCliente = oParentescoCliente.mostrarProveedorClienteProcesoReforzado(fnCodigoClienteProcesoReforzado, "C")


If fsMotivoRegistro = "PEPS" Then
   Set lrsParienteCliente = oParentescoCliente.mostrarParienteClienteProcesoReforzado(fnCodigoClienteProcesoReforzado)
   If fsEstadoCivil = "2" Then
      Set lrsConyugeKey = oParentescoConyuge.mostarLlaveConyuge(fnCodigoClienteProcesoReforzado)
      Set lrsParienteConyuge = oParentescoCliente.mostrarParienteConyugeClienteProcesoReforzado(lrsConyugeKey.Fields(0))
      lrsConyugeKey.Close: Set lrsConyugeKey = Nothing
   End If

End If

Set lrsRelacionEntidad = oParentescoCliente.mostrarRelacionEntidadClienteProcesoReforzado(fsCodigoPersona)
Dim wApp As Word.Application
Dim wAppSource As Word.Application
Set wApp = New Word.Application
Set wAppSource = New Word.Application

lsPlantilla = App.Path & "\FormatoCarta\PlantillaClienteReforzado.doc"

wAppSource.Documents.Open FileName:=lsPlantilla
wAppSource.ActiveDocument.Content.Copy
wApp.Documents.Add


wApp.Application.Selection.TypeParagraph
wApp.Application.Selection.Paste
wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
wApp.Selection.MoveEnd

With wApp.Selection.Find
         .Text = "<<tipocliente>>"
         .Replacement.Text = lrsDatosCliente.Fields("vTipoCliente")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<apellidoscl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vApellidos")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<nombrescl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vNombres")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<dnicl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vDni")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<cecl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vCe")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<pasaportecl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vPasaporte")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<fechanacimientocl>>"
         .Replacement.Text = CStr(lrsDatosCliente.Fields("dFechaNacimiento"))
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<profesioncl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vProfesion")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<nacionalidadcl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vNacionalidad")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<centrolaborcl>>"
         .Replacement.Text = lrsDatosCliente.Fields("vCentroLabor")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<ingresomensualcl>>"
         .Replacement.Text = Format(lrsDatosCliente.Fields("mIngresoPromedio"), "#,##0.00")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<direccion>>"
         .Replacement.Text = lrsDatosCliente.Fields("vDireccion")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<distrito>>"
         .Replacement.Text = lrsDatosCliente.Fields("vDistrito")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<provincia>>"
         .Replacement.Text = lrsDatosCliente.Fields("vProvincia")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<departamento>>"
         .Replacement.Text = lrsDatosCliente.Fields("vDepartamento")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<telefonos>>"
         .Replacement.Text = lrsDatosCliente.Fields("vTelefonos")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<correo>>"
         .Replacement.Text = lrsDatosCliente.Fields("vCorreo")
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<fecharegistro>>"
         .Replacement.Text = CStr(lrsDatosCliente.Fields("dFechaRegistro"))
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<autorizado>>"
         .Replacement.Text = CStr(lrsDatosCliente.Fields("vPersonaAutoriza"))
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<atendido>>"
         .Replacement.Text = CStr(lrsDatosCliente.Fields("vPersonaAtencion"))
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<cargo>>"
         .Replacement.Text = CStr(lrsDatosCliente.Fields("vCargo"))
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

With wApp.Selection.Find
         .Text = "<<agencia>>"
         .Replacement.Text = CStr(lrsDatosCliente.Fields("vAgencia"))
         .Forward = True
         .Wrap = wdFindContinue
         .Format = False
         .Execute Replace:=wdReplaceAll
End With

lrsDatosCliente.Close: Set lrsDatosCliente = Nothing

If lrsDatosConyuge.RecordCount > 0 Then
   With wApp.Selection.Find
            .Text = "<<apellidosco>>"
            .Replacement.Text = lrsDatosConyuge.Fields("vApellidos")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

   With wApp.Selection.Find
            .Text = "<<nombresco>>"
            .Replacement.Text = lrsDatosConyuge.Fields("vNombres")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With

    Select Case lrsDatosConyuge.Fields("nTipoDocumento")
           Case 1
                lsTipoDoiConyuge = "<<dnico>>"
                lsTipoDoiConyuge2 = "<<ceco>>"
                lsTipoDoiConyuge3 = "<<pasaporteco>>"
            Case 4
                 lsTipoDoiConyuge = "<<ceco>>"
                 lsTipoDoiConyuge2 = "<<dnico>>"
                 lsTipoDoiConyuge3 = "<<pasaporteco>>"
            Case 11
                 lsTipoDoiConyuge = "<<pasaporteco>>"
                 lsTipoDoiConyuge2 = "<<ceco>>"
                 lsTipoDoiConyuge3 = "<<dnico>>"
     End Select

     With wApp.Selection.Find
              .Text = lsTipoDoiConyuge
              .Replacement.Text = lrsDatosConyuge.Fields("vNumeroDocumento")
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
    End With
 
    With wApp.Selection.Find
             .Text = lsTipoDoiConyuge2
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
     End With

     With wApp.Selection.Find
              .Text = lsTipoDoiConyuge3
              .Replacement.Text = ""
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
     End With

     With wApp.Selection.Find
              .Text = "<<fechanacimientoco>>"
              .Replacement.Text = CStr(lrsDatosConyuge.Fields("dFechaNacimento"))
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
     End With

     With wApp.Selection.Find
              .Text = "<<profesionco>>"
              .Replacement.Text = lrsDatosConyuge.Fields("vProfesion")
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<nacionalidadco>>"
             .Replacement.Text = lrsDatosConyuge.Fields("vNacionalidad")
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<centrolaborco>>"
             .Replacement.Text = lrsDatosConyuge.Fields("vCentroLaboral")
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<ingresomensualco>>"
             .Replacement.Text = Format(lrsDatosConyuge.Fields("mIngresoPromedio"), "#,##0.00")
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

Else

   With wApp.Selection.Find
            .Text = "<<apellidosco>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

   With wApp.Selection.Find
            .Text = "<<nombresco>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<dnico>>"
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With
 
    With wApp.Selection.Find
             .Text = "<<ceco>>"
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
     End With

     With wApp.Selection.Find
              .Text = "<<pasaporteco>>"
              .Replacement.Text = ""
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
     End With

     With wApp.Selection.Find
              .Text = "<<fechanacimientoco>>"
              .Replacement.Text = ""
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
     End With

     With wApp.Selection.Find
              .Text = "<<profesionco>>"
              .Replacement.Text = ""
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<nacionalidadco>>"
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<centrolaborco>>"
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<ingresomensualco>>"
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

End If

lrsDatosConyuge.Close: Set lrsDatosConyuge = Nothing

If lrsRefernciaEconomica.RecordCount > 0 Then

 Dim tblNew1 As Table
 Dim rowNew1 As row
 Dim celTable1 As Cell

 Set tblNew1 = wApp.ActiveDocument.Tables(3)
  
 Do While Not lrsRefernciaEconomica.EOF
 DoEvents
 Set rowNew1 = tblNew1.Rows.Add(BeforeRow:=tblNew1.Rows(tblNew1.Rows.count))
 
 For Each celTable1 In rowNew1.Cells
  Select Case celTable1.ColumnIndex
   Case 1
    celTable1.Range.InsertAfter Text:=lrsRefernciaEconomica!vCargoPublico
   Case 2
    celTable1.Range.InsertAfter Text:=CStr(lrsRefernciaEconomica!dFechaInicio)
   Case 3
    celTable1.Range.InsertAfter Text:=CStr(lrsRefernciaEconomica!dFechaCese)
   Case 4
    celTable1.Range.InsertAfter Text:=lrsRefernciaEconomica!vEntidadLabora
   Case 5
    celTable1.Range.InsertAfter Text:=Format(lrsRefernciaEconomica!mIngreso, "#,##0.00")
   End Select
 Next celTable1

lrsRefernciaEconomica.MoveNext
If lrsRefernciaEconomica.EOF Then
   Exit Do
End If
Loop

tblNew1.Rows(tblNew1.Rows.count).Delete

End If

lrsRefernciaEconomica.Close: Set lrsRefernciaEconomica = Nothing

If lrsRefernciaFinanciera.RecordCount > 0 Then

 Dim tblNew2 As Table
 Dim rowNew2 As row
 Dim celTable2 As Cell

 Set tblNew2 = wApp.ActiveDocument.Tables(4)
 
 Do While Not lrsRefernciaFinanciera.EOF
 DoEvents
 Set rowNew2 = tblNew2.Rows.Add(BeforeRow:=tblNew2.Rows(tblNew2.Rows.count))
 
 For Each celTable2 In rowNew2.Cells
  Select Case celTable2.ColumnIndex
   Case 1
    celTable2.Range.InsertAfter Text:=lrsRefernciaFinanciera!vBanco
   Case 2
    celTable2.Range.InsertAfter Text:=lrsRefernciaFinanciera!vTipoProducto
   Case 3
    celTable2.Range.InsertAfter Text:=lrsRefernciaFinanciera!vFuncionarioNegocio
   End Select
 Next celTable2

lrsRefernciaFinanciera.MoveNext
If lrsRefernciaFinanciera.EOF Then
   Exit Do
End If
Loop

tblNew2.Rows(tblNew2.Rows.count).Delete

End If

lrsRefernciaFinanciera.Close: Set lrsRefernciaFinanciera = Nothing

If lrsRefernciaPatrimonial.RecordCount > 0 Then

Dim tblNew3 As Table
Dim rowNew3 As row
Dim celTable3 As Cell

 Set tblNew3 = wApp.ActiveDocument.Tables(5)
 
 Do While Not lrsRefernciaPatrimonial.EOF
 DoEvents
 Set rowNew3 = tblNew3.Rows.Add(BeforeRow:=tblNew3.Rows(tblNew3.Rows.count))
 
 For Each celTable3 In rowNew3.Cells
  Select Case celTable3.ColumnIndex
   Case 1
    celTable3.Range.InsertAfter Text:=lrsRefernciaPatrimonial!vBien
   Case 2
    celTable3.Range.InsertAfter Text:=Format(lrsRefernciaPatrimonial!mValor, "#,##0.00")
   End Select
 Next celTable3

lrsRefernciaPatrimonial.MoveNext
If lrsRefernciaPatrimonial.EOF Then
   Exit Do
End If
Loop

tblNew3.Rows(tblNew3.Rows.count).Delete

End If

lrsRefernciaPatrimonial.Close: Set lrsRefernciaPatrimonial = Nothing

If lrsPersonaJuridica.RecordCount > 0 Then
Dim tblNew4 As Table
Dim rowNew4 As row
Dim celTable4 As Cell

 Set tblNew4 = wApp.ActiveDocument.Tables(6)
 
 Do While Not lrsPersonaJuridica.EOF
 DoEvents
 Set rowNew4 = tblNew4.Rows.Add(BeforeRow:=tblNew4.Rows(tblNew4.Rows.count))
 
 For Each celTable4 In rowNew4.Cells
  Select Case celTable4.ColumnIndex
   Case 1
    celTable4.Range.InsertAfter Text:=lrsPersonaJuridica!vNombreEmpresa
   Case 2
    celTable4.Range.InsertAfter Text:=lrsPersonaJuridica!vRuc
    Case 3
    celTable4.Range.InsertAfter Text:=lrsPersonaJuridica!nParticipacion
    Case 4
    celTable4.Range.InsertAfter Text:=Format(lrsPersonaJuridica!mIngreso, "#,##0.00")
   End Select
 Next celTable4

lrsPersonaJuridica.MoveNext
If lrsPersonaJuridica.EOF Then
   Exit Do
End If
Loop

tblNew4.Rows(tblNew4.Rows.count).Delete

End If

lrsPersonaJuridica.Close: Set lrsPersonaJuridica = Nothing

If lrsProveedor.RecordCount > 0 And lrsProveedor.RecordCount >= lrsCliente.RecordCount Then
Dim tblNew5 As Table
Dim rowNew5 As row
Dim celTable5 As Cell

 Set tblNew5 = wApp.ActiveDocument.Tables(7)
 
 Do While Not lrsProveedor.EOF
 DoEvents
 Set rowNew5 = tblNew5.Rows.Add(BeforeRow:=tblNew5.Rows(tblNew5.Rows.count))
   
 For Each celTable5 In rowNew5.Cells
  Select Case celTable5.ColumnIndex
   Case 1
     celTable5.Range.InsertAfter Text:=lrsProveedor!vNombreEmpresa
    End Select
 Next celTable5

lrsProveedor.MoveNext
If lrsProveedor.EOF Then
   Exit Do
End If
Loop

End If

If lrsCliente.RecordCount > 0 And lrsCliente.RecordCount > lrsProveedor.RecordCount Then
Dim tblNew8 As Table
Dim rowNew8 As row
Dim celTable8 As Cell

 Set tblNew8 = wApp.ActiveDocument.Tables(7)
 
 Do While Not lrsCliente.EOF
 DoEvents
 Set rowNew8 = tblNew8.Rows.Add(BeforeRow:=tblNew8.Rows(tblNew8.Rows.count))
   
 For Each celTable8 In rowNew8.Cells
  Select Case celTable8.ColumnIndex
   Case 2
     celTable8.Range.InsertAfter Text:=lrsCliente!vNombreEmpresa
    End Select
 Next celTable8

lrsCliente.MoveNext
If lrsCliente.EOF Then
   Exit Do
End If
Loop

End If

If lrsProveedor.RecordCount > 0 And lrsProveedor.RecordCount < lrsCliente.RecordCount Then
Dim tblNew9 As Table
Dim i As Integer
Set tblNew9 = wApp.ActiveDocument.Tables(7)
 i = 0
 Do While Not lrsProveedor.EOF
 DoEvents
   
tblNew9.Cell(2 + i, 1).Range.InsertAfter Text:=lrsProveedor!vNombreEmpresa
i = i + 1
lrsProveedor.MoveNext
If lrsProveedor.EOF Then
   Exit Do
End If
Loop

End If

If lrsCliente.RecordCount > 0 And lrsCliente.RecordCount <= lrsProveedor.RecordCount Then
Dim tblNew10 As Table
Dim J As Integer
 Set tblNew10 = wApp.ActiveDocument.Tables(7)
 J = 0
 Do While Not lrsCliente.EOF
 DoEvents
 tblNew10.Cell(2 + J, 2).Range.InsertAfter Text:=lrsCliente!vNombreEmpresa
J = J + 1
lrsCliente.MoveNext
If lrsCliente.EOF Then
   Exit Do
End If
Loop

End If

Dim tblNew11 As Table
Set tblNew11 = wApp.ActiveDocument.Tables(7)

If tblNew11.Rows.count > 2 Then
tblNew11.Rows(tblNew11.Rows.count).Delete
End If

lrsProveedor.Close: Set lrsProveedor = Nothing
lrsCliente.Close: Set lrsCliente = Nothing

If fsMotivoRegistro = "PEPS" Then
    If lrsParienteCliente.RecordCount > 0 Then
        Dim tblNew6 As Table
        Dim rowNew6 As row
        Dim celTable6 As Cell

        Set tblNew6 = wApp.ActiveDocument.Tables(9)
  
        Do While Not lrsParienteCliente.EOF
            DoEvents
            Set rowNew6 = tblNew6.Rows.Add(BeforeRow:=tblNew6.Rows(tblNew6.Rows.count))
 
            For Each celTable6 In rowNew6.Cells
                Select Case celTable6.ColumnIndex
                    Case 1
                        celTable6.Range.InsertAfter Text:=lrsParienteCliente!vCodigo
                    Case 2
                        celTable6.Range.InsertAfter Text:=lrsParienteCliente!vNombreCompleto
                    Case 3
                        celTable6.Range.InsertAfter Text:=lrsParienteCliente!vDni
                End Select
            Next celTable6

            lrsParienteCliente.MoveNext
            If lrsParienteCliente.EOF Then
                Exit Do
            End If
        Loop

        tblNew6.Rows(tblNew6.Rows.count).Delete

    End If

    lrsParienteCliente.Close: Set lrsParienteCliente = Nothing

    If fsEstadoCivil = "2" Then
        If lrsParienteConyuge.RecordCount > 0 Then
            Dim tblNew7 As Table
            Dim rowNew7 As row
            Dim celTable7 As Cell

            Set tblNew7 = wApp.ActiveDocument.Tables(10)
 
            Do While Not lrsParienteConyuge.EOF
                DoEvents
                Set rowNew7 = tblNew7.Rows.Add(BeforeRow:=tblNew7.Rows(tblNew7.Rows.count))
 
                For Each celTable7 In rowNew7.Cells
                    Select Case celTable7.ColumnIndex
                        Case 1
                            celTable7.Range.InsertAfter Text:=lrsParienteConyuge!vCodigo
                        Case 2
                            celTable7.Range.InsertAfter Text:=lrsParienteConyuge!vNombreCompleto
                        Case 3
                            celTable7.Range.InsertAfter Text:=lrsParienteConyuge!vDni
                    End Select
                Next celTable7

                lrsParienteConyuge.MoveNext
                If lrsParienteConyuge.EOF Then
                    Exit Do
                End If
            Loop

            tblNew7.Rows(tblNew7.Rows.count).Delete
        End If
        lrsParienteConyuge.Close: Set lrsParienteConyuge = Nothing
    End If


End If

If lrsRelacionEntidad.RecordCount > 0 Then

   With wApp.Selection.Find
            .Text = "<<a>>"
            .Replacement.Text = lrsRelacionEntidad.Fields("vAhorro")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With
 
   With wApp.Selection.Find
            .Text = "<<c>>"
            .Replacement.Text = lrsRelacionEntidad.Fields("vCredito")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

   With wApp.Selection.Find
            .Text = "<<s>>"
            .Replacement.Text = lrsRelacionEntidad.Fields("vServicio")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

   With wApp.Selection.Find
            .Text = "<<o>>"
            .Replacement.Text = lrsRelacionEntidad.Fields("vOtro")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With
   
 Else

   With wApp.Selection.Find
            .Text = "<<a>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With
 
   With wApp.Selection.Find
            .Text = "<<c>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

   With wApp.Selection.Find
            .Text = "<<s>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

   With wApp.Selection.Find
            .Text = "<<o>>"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
   End With

End If

lrsRelacionEntidad.Close: Set lrsRelacionEntidad = Nothing

Screen.MousePointer = 0
wAppSource.ActiveDocument.Close
wApp.ActiveDocument.CopyStylesFromTemplate (lsPlantilla)
wApp.ActiveDocument.SaveAs (App.Path & "\Spooler\ClienteProcesoReforzado" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".doc")
wApp.Visible = True
Set wAppSource = Nothing
Set wApp = Nothing
Set oParentescoCliente = Nothing
Set oParentescoConyuge = Nothing

Exit Sub

ErrGeneraRepo:
        Screen.MousePointer = 0
        MsgBox "frmPersClienteSensivle: No tiene el formato PlantillaClienteReforzado, comuniquese con el Área de TI", vbInformation, "Aviso"
End Sub

Private Sub ImprimeAnexo2()
    Dim lrsDatosPrincipalesCliente As New ADODB.Recordset
    Dim lrsBienesInmuebles As New ADODB.Recordset
    Dim lrsOtrosBienes As New ADODB.Recordset
    Dim lsPlantillaAnexo2 As String
    Dim lstipodocumento, lsdocumento As String
    
    Set oParentescoCliente = New COMNPersona.NCOMPersona
    Set oParentescoConyuge = New COMNPersona.NCOMPersona
    
    On Error GoTo ErrGeneraRepo
    Screen.MousePointer = 11


    Set lrsDatosPrincipalesCliente = oParentescoCliente.mostrarDatosPrincipalesClienteProcesoReforzado(fsCodigoPersona)
    Set lrsBienesInmuebles = oParentescoCliente.mostrarReferenciaPatrimonialClienteProcesoReforzado(fnCodigoClienteProcesoReforzado, "P")
    Set lrsOtrosBienes = oParentescoCliente.mostrarReferenciaPatrimonialClienteProcesoReforzado(fnCodigoClienteProcesoReforzado, "O")
    
    
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application

    lsPlantillaAnexo2 = App.Path & "\FormatoCarta\PlantillaClienteReforzado2.doc"
        
    wAppSource.Documents.Open FileName:=lsPlantillaAnexo2
    wAppSource.ActiveDocument.Content.Copy
    
    wApp.Documents.Add
    wApp.Application.Selection.TypeParagraph
    wApp.Application.Selection.Paste
    wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
    wApp.Selection.MoveEnd

    With wApp.Selection.Find
             .Text = "<<nombrecliente>>"
             .Replacement.Text = lrsDatosPrincipalesCliente.Fields(0)
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With
    
    If lrsDatosPrincipalesCliente.Fields(1) <> "" Then
        lstipodocumento = "DNI"
        lsdocumento = lrsDatosPrincipalesCliente.Fields(1)
    Else
        If lrsDatosPrincipalesCliente.Fields(2) <> "" Then
        lstipodocumento = "CE"
        lsdocumento = lrsDatosPrincipalesCliente.Fields(2)
        Else
            If lrsDatosPrincipalesCliente.Fields(3) <> "" Then
            lstipodocumento = "Pasaporte"
            lsdocumento = lrsDatosPrincipalesCliente.Fields(3)
            End If
        End If
    End If
    
    

    With wApp.Selection.Find
             .Text = "<<tipodocumento>>"
             .Replacement.Text = lstipodocumento
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<documento>>"
             .Replacement.Text = lsdocumento
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    With wApp.Selection.Find
             .Text = "<<ciudad>>"
             .Replacement.Text = ""
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With
    
    With wApp.Selection.Find
             .Text = "<<fecha>>"
             .Replacement.Text = Format(gdFecSis, "Long Date")
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .Execute Replace:=wdReplaceAll
    End With

    lrsDatosPrincipalesCliente.Close: Set lrsDatosPrincipalesCliente = Nothing
    
    
    If lrsBienesInmuebles.RecordCount > 0 Then

        Dim tblNew1 As Table
        Dim rowNew1 As row
        Dim celTable1 As Cell
    
        Set tblNew1 = wApp.ActiveDocument.Tables(1)
        
        Do While Not lrsBienesInmuebles.EOF
            DoEvents
            Set rowNew1 = tblNew1.Rows.Add(BeforeRow:=tblNew1.Rows(tblNew1.Rows.count))
 
            For Each celTable1 In rowNew1.Cells
                Select Case celTable1.ColumnIndex
                    Case 1
                        celTable1.Range.InsertAfter Text:=lrsBienesInmuebles!vBien
                    Case 2
                        celTable1.Range.InsertAfter Text:=Format(lrsBienesInmuebles!mValor, "#,##0.00")
                End Select
            Next celTable1

            lrsBienesInmuebles.MoveNext
            If lrsBienesInmuebles.EOF Then
                Exit Do
            End If
        Loop

        tblNew1.Rows(tblNew1.Rows.count).Delete

    End If

    lrsBienesInmuebles.Close: Set lrsBienesInmuebles = Nothing

    
    If lrsOtrosBienes.RecordCount > 0 Then

        Dim tblNew2 As Table
        Dim rowNew2 As row
        Dim celTable2 As Cell

        Set tblNew2 = wApp.ActiveDocument.Tables(2)
        
        Do While Not lrsOtrosBienes.EOF
            DoEvents
            Set rowNew2 = tblNew2.Rows.Add(BeforeRow:=tblNew2.Rows(tblNew2.Rows.count))
 
            For Each celTable2 In rowNew2.Cells
                celTable2.Range.InsertAfter Text:=lrsOtrosBienes!vBien
            Next celTable2

            lrsOtrosBienes.MoveNext
    
            If lrsOtrosBienes.EOF Then
                Exit Do
            End If
        Loop

        tblNew2.Rows(tblNew2.Rows.count).Delete

    End If
    
    lrsOtrosBienes.Close: Set lrsOtrosBienes = Nothing
    
    Screen.MousePointer = 0
    wAppSource.ActiveDocument.Close
    wApp.ActiveDocument.CopyStylesFromTemplate (lsPlantillaAnexo2)
    wApp.ActiveDocument.SaveAs (App.Path & "\Spooler\ClienteProcesoReforzado2" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".doc")
    wApp.Visible = True
    Set wAppSource = Nothing
    Set wApp = Nothing
    Set oParentescoCliente = Nothing
    Set oParentescoConyuge = Nothing
    
    Exit Sub
       
ErrGeneraRepo:
            Screen.MousePointer = 0
            MsgBox "frmPersClienteSensivle: No tiene el formato PlantillaClienteReforzado2, comuniquese con el Área de TI", vbInformation, "Aviso"
    
End Sub



Public Sub Inicio(ByVal pcPerCodCliente As String, _
                  ByVal pcMovPersonaAtendido As String, _
                  ByVal pcMovPersonaAutorizado As String, _
                  ByVal psMotivoRegistro As String, _
                  ByVal psEstadoCivil As String, _
                  ByVal psCentroLaboral)
             

Me.TxtBuscarCliente.Text = pcPerCodCliente
Me.TxtBuscarCliente.Enabled = False
Me.TxtCentroLaboralCliente.Text = psCentroLaboral
fsEstadoCivil = psEstadoCivil
fsMotivoRegistro = psMotivoRegistro
fcMovPersonaAutorizado = pcMovPersonaAutorizado
fcMovPersonaAtendido = pcMovPersonaAtendido

Call Restingir
Call TxtBuscarCliente_EmiteDatos
Me.Show 1
End Sub

Private Sub Restingir()

    If fsMotivoRegistro = "PEPS" Then
        deshabilitarDatosCliente
        
        If fsEstadoCivil = "2" Then
            Call habilitarDatosConyuge
            Call habilitarBotonesParentescoConyuge
        Else
            Call deshabilitarDatosConyuge
            Call deshabilitarBotonesParentescoConyuge
        End If
        
        Call habilitarBotonesReferenciaEconomica
        Call habilitarBotonesReferenciaFinanciera
        Call habilitarBotonesReferenciaPatrimonial
        Call habilitarBotonesOtroBien
        Call habilitarBotonesPersonaJuridica
        Call habilitarBotonesProveedor
        Call habilitarBotonesCliente
        Call habilitarBotonesParentescoCliente
 
    Else
        habiliraDatosCliente
        
    If fsEstadoCivil = "2" Then
        Call habilitarDatosConyuge
    Else
        Call deshabilitarDatosConyuge
    End If
    
    Call deshabilitarBotonesReferenciaEconomica
    Call habilitarBotonesReferenciaFinanciera
    Call habilitarBotonesReferenciaPatrimonial
    Call habilitarBotonesOtroBien
    Call habilitarBotonesPersonaJuridica
    Call habilitarBotonesProveedor
    Call habilitarBotonesCliente
    Call deshabilitarBotonesParentescoCliente
    Call deshabilitarBotonesParentescoConyuge
End If
End Sub





