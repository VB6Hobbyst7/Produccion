VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{C70F746A-6712-11D5-BF5D-000103686C75}#1.0#0"; "PryImage.ocx"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personas: Mantenimiento"
   ClientHeight    =   6480
   ClientLeft      =   705
   ClientTop       =   1440
   ClientWidth     =   10680
   Icon            =   "PersonasMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdPersCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9360
      TabIndex        =   59
      ToolTipText     =   "Cancelar Todos los cambios Realizados"
      Top             =   6105
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdPersAceptar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8100
      TabIndex        =   58
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   6105
      Visible         =   0   'False
      Width           =   1230
   End
   Begin Sicmact.TxtBuscar TxtBCodPers 
      Height          =   330
      Left            =   675
      TabIndex        =   0
      Top             =   97
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgImg 
      Height          =   480
      Left            =   10110
      TabIndex        =   48
      Top             =   2550
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   847
      Filtro          =   "Archivos bmp|*.bmp|Archivos jpg|*.jpg|Archivos gif|*.gif|Todos los Archivos|*.*"
      Altura          =   280
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   300
      Left            =   9600
      TabIndex        =   27
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   300
      Left            =   9600
      TabIndex        =   28
      Top             =   465
      Width           =   1000
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   300
      Left            =   9600
      TabIndex        =   29
      Top             =   825
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTDatosGen 
      Height          =   3045
      Left            =   105
      TabIndex        =   38
      Top             =   2985
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   5371
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "PersonasMantenimiento.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPersNac"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPersTelefono"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPersCIIU"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPersEstado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblRelPers"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPersTelefono"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbPersEstado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CboPersCiiu"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPersNacCreac"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPersTelefono2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtEmail"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCodSbs"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmbRelPers"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "&Domicilio"
      TabPicture(1)   =   "PersonasMantenimiento.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPersDireccDomicilio"
      Tab(1).Control(1)=   "lblPersDireccCondicion"
      Tab(1).Control(2)=   "txtPersDireccDomicilio"
      Tab(1).Control(3)=   "cmbPersDireccCondicion"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox CmbRelPers 
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0342
         Left            =   5490
         List            =   "PersonasMantenimiento.frx":0344
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   615
         Width           =   1710
      End
      Begin VB.TextBox txtCodSbs 
         Height          =   300
         Left            =   4725
         TabIndex        =   91
         Top             =   1530
         Width           =   1785
      End
      Begin VB.TextBox TxtEmail 
         Height          =   300
         Left            =   1200
         TabIndex        =   72
         Top             =   1530
         Width           =   2985
      End
      Begin VB.TextBox txtPersTelefono2 
         Height          =   300
         Left            =   3165
         TabIndex        =   70
         Top             =   1140
         Width           =   1785
      End
      Begin MSMask.MaskEdBox txtPersNacCreac 
         Height          =   330
         Left            =   2640
         TabIndex        =   13
         Top             =   615
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPersCiiu 
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0346
         Left            =   1215
         List            =   "PersonasMantenimiento.frx":0348
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   5520
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ubicacion Geografica"
         Height          =   1665
         Left            =   -74865
         TabIndex        =   51
         Top             =   420
         Width           =   7155
         Begin VB.ComboBox cmbPersUbiGeo 
            Height          =   315
            Index           =   0
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   525
            Width           =   1815
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Height          =   315
            Index           =   1
            Left            =   2145
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Height          =   315
            Index           =   4
            Left            =   4500
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Height          =   315
            Index           =   2
            Left            =   4500
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Height          =   315
            Index           =   3
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1140
            Width           =   2190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4515
            TabIndex        =   56
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2160
            TabIndex        =   55
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   210
            TabIndex        =   54
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4500
            TabIndex        =   53
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2130
            TabIndex        =   52
            Top             =   900
            Width           =   600
         End
      End
      Begin VB.ComboBox cmbPersDireccCondicion 
         Height          =   315
         Left            =   -73740
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2550
         Width           =   1695
      End
      Begin VB.TextBox txtPersDireccDomicilio 
         Height          =   285
         Left            =   -73740
         TabIndex        =   22
         Top             =   2190
         Width           =   5200
      End
      Begin VB.ComboBox cmbPersEstado 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2370
         Width           =   3690
      End
      Begin VB.TextBox txtPersTelefono 
         Height          =   300
         Left            =   1200
         TabIndex        =   14
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Label LblRelPers 
         Caption         =   "Relac:  Institucion :"
         Height          =   255
         Left            =   4125
         TabIndex        =   94
         Top             =   645
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SBS:"
         Height          =   195
         Left            =   4290
         TabIndex        =   92
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         Height          =   195
         Left            =   285
         TabIndex        =   73
         Top             =   1590
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   ","
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
         Left            =   3060
         TabIndex        =   71
         Top             =   1260
         Width           =   75
      End
      Begin VB.Label lblPersDireccCondicion 
         AutoSize        =   -1  'True
         Caption         =   "Condicion"
         Height          =   195
         Left            =   -74820
         TabIndex        =   50
         Top             =   2580
         Width           =   705
      End
      Begin VB.Label lblPersDireccDomicilio 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   -74820
         TabIndex        =   49
         Top             =   2220
         Width           =   630
      End
      Begin VB.Label lblPersEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   285
         TabIndex        =   46
         Top             =   2370
         Width           =   495
      End
      Begin VB.Label lblPersCIIU 
         AutoSize        =   -1  'True
         Caption         =   "CIIU"
         Height          =   195
         Left            =   285
         TabIndex        =   41
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label lblPersTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Telefono(s)"
         Height          =   195
         Left            =   285
         TabIndex        =   40
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblPersNac 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Nacimiento/Creación"
         Height          =   195
         Left            =   285
         TabIndex        =   39
         Top             =   630
         Width           =   2220
      End
   End
   Begin VB.ComboBox cmbPersPersoneria 
      Height          =   315
      Left            =   3615
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   2835
   End
   Begin TabDlg.SSTab SSTabs 
      Height          =   2400
      Left            =   120
      TabIndex        =   32
      Top             =   555
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4233
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Persona &Natural"
      TabPicture(0)   =   "PersonasMantenimiento.frx":034A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPersNombreAM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPersNatEstCiv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPersNatHijos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblApCasada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPersNombreN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPeso"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblTpoSangre"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPersNatSexo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPersNombreAP"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblTalla"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbPersNatEstCiv"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPersNatHijos"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtPeso"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtTipoSangre"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbPersNatSexo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtTalla"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtPersNombreAP"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtPersNombreAM"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtApellidoCasada"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtPersNombreN"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Personal &Juridica"
      TabPicture(1)   =   "PersonasMantenimiento.frx":0366
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPersNombre"
      Tab(1).Control(1)=   "lblPersJurSiglas"
      Tab(1).Control(2)=   "lblPersJurTpo"
      Tab(1).Control(3)=   "lblPersJurMagnitud"
      Tab(1).Control(4)=   "lblPersJurEmpleados"
      Tab(1).Control(5)=   "txtPersNombreRS"
      Tab(1).Control(6)=   "TxtSiglas"
      Tab(1).Control(7)=   "cmbPersJurTpo"
      Tab(1).Control(8)=   "cmbPersJurMagnitud"
      Tab(1).Control(9)=   "txtPersJurEmpleados"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "&Relaciones con otras Personas"
      TabPicture(2)   =   "PersonasMantenimiento.frx":0382
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPersRelacNew"
      Tab(2).Control(1)=   "cmdPersRelacDel"
      Tab(2).Control(2)=   "FERelPers"
      Tab(2).Control(3)=   "cmdPersRelacEditar"
      Tab(2).Control(4)=   "cmdPersRelacCancelar"
      Tab(2).Control(5)=   "cmdPersRelacAceptar"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&Fuentes de Ingreso"
      TabPicture(3)   =   "PersonasMantenimiento.frx":039E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FEFteIng"
      Tab(3).Control(1)=   "CmdFteIngEditar"
      Tab(3).Control(2)=   "CmdFteIngEliminar"
      Tab(3).Control(3)=   "CmdFteIngNuevo"
      Tab(3).Control(4)=   "CmdPersFteConsultar"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Fir&ma"
      TabPicture(4)   =   "PersonasMantenimiento.frx":03BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CmdActFirma"
      Tab(4).Control(1)=   "IDBFirma"
      Tab(4).ControlCount=   2
      Begin VB.TextBox txtPersNombreN 
         Height          =   300
         Left            =   1545
         TabIndex        =   90
         Top             =   1890
         Width           =   2895
      End
      Begin VB.TextBox TxtApellidoCasada 
         Height          =   300
         Left            =   1545
         TabIndex        =   89
         Top             =   1515
         Width           =   2895
      End
      Begin VB.TextBox txtPersNombreAM 
         Height          =   300
         Left            =   1545
         TabIndex        =   88
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtPersNombreAP 
         Height          =   300
         Left            =   1545
         TabIndex        =   87
         Top             =   675
         Width           =   2895
      End
      Begin VB.CommandButton CmdPersFteConsultar 
         Caption         =   "Consultar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   84
         Top             =   1860
         Width           =   1140
      End
      Begin PryImage.ImageDB IDBFirma 
         Height          =   2025
         Left            =   -73335
         TabIndex        =   80
         Top             =   315
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   3572
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdPersRelacAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   79
         Top             =   1620
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   78
         Top             =   1935
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngNuevo 
         Caption         =   "Nuevo"
         Height          =   300
         Left            =   -67260
         TabIndex        =   77
         Top             =   465
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngEliminar 
         Caption         =   "Eliminar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   76
         Top             =   1095
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngEditar 
         Caption         =   "Editar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   75
         Top             =   780
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacEditar 
         Caption         =   "Editar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   74
         Top             =   780
         Width           =   1140
      End
      Begin VB.TextBox TxtTalla 
         Height          =   300
         Left            =   8205
         TabIndex        =   68
         Top             =   615
         Width           =   870
      End
      Begin VB.ComboBox cmbPersNatSexo 
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":03D6
         Left            =   5580
         List            =   "PersonasMantenimiento.frx":03D8
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   608
         Width           =   1710
      End
      Begin VB.TextBox TxtTipoSangre 
         Height          =   300
         Left            =   8205
         TabIndex        =   63
         Top             =   1065
         Width           =   870
      End
      Begin VB.TextBox TxtPeso 
         Height          =   300
         Left            =   5580
         TabIndex        =   61
         Top             =   1860
         Width           =   870
      End
      Begin Sicmact.FlexEdit FEFteIng 
         Height          =   1785
         Left            =   -74775
         TabIndex        =   11
         Top             =   435
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   3149
         Cols0           =   6
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Tipo-Razon Social-F. Caducac.-F. Evaluacion-Moneda"
         EncabezadosAnchos=   "400-600-4000-1000-1200-1200"
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
         ColumnasAEditar =   "X-1-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-L-C-L"
         FormatosEdit    =   "3-1-1-0-0-1"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit FERelPers 
         Height          =   1785
         Left            =   -74775
         TabIndex        =   8
         Top             =   435
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   3149
         Cols0           =   7
         FixedCols       =   0
         HighLight       =   2
         EncabezadosNombres=   "Item-Codigo-Nombres-Relacion-Beneficiario-Porcentaje-Asist.Med.Priv."
         EncabezadosAnchos=   "400-1300-4000-1400-2000-1200-2500"
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
         ColumnasAEditar =   "X-1-X-3-4-5-6"
         ListaControles  =   "0-1-0-3-3-0-3"
         EncabezadosAlineacion=   "C-L-L-L-L-R-L"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin VB.CommandButton CmdActFirma 
         Caption         =   "&Actualizar Firma"
         Height          =   375
         Left            =   -67740
         TabIndex        =   12
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmdPersRelacDel 
         Caption         =   "Eliminar"
         Height          =   300
         Left            =   -67260
         TabIndex        =   10
         Top             =   1095
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacNew 
         Caption         =   "Nuevo"
         Height          =   300
         Left            =   -67260
         TabIndex        =   9
         Top             =   465
         Width           =   1140
      End
      Begin VB.TextBox txtPersJurEmpleados 
         Height          =   285
         Left            =   -73680
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox cmbPersJurMagnitud 
         Height          =   315
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1305
         Width           =   3000
      End
      Begin VB.ComboBox cmbPersJurTpo 
         Height          =   315
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   3000
      End
      Begin VB.TextBox TxtSiglas 
         Height          =   285
         Left            =   -73680
         TabIndex        =   30
         Top             =   960
         Width           =   1500
      End
      Begin VB.TextBox txtPersNombreRS 
         Height          =   300
         Left            =   -73680
         TabIndex        =   4
         Top             =   600
         Width           =   7815
      End
      Begin VB.TextBox txtPersNatHijos 
         Height          =   300
         Left            =   5580
         TabIndex        =   3
         Top             =   1485
         Width           =   300
      End
      Begin VB.ComboBox cmbPersNatEstCiv 
         Height          =   315
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1065
         Width           =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "m."
         Height          =   195
         Left            =   9105
         TabIndex        =   86
         Top             =   675
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kg."
         Height          =   195
         Left            =   6510
         TabIndex        =   85
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label LblTalla 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
         Height          =   195
         Left            =   7740
         TabIndex        =   69
         Top             =   675
         Width           =   345
      End
      Begin VB.Label lblPersNombreAP 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   668
         Width           =   1155
      End
      Begin VB.Label lblPersNatSexo 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Left            =   5145
         TabIndex        =   66
         Top             =   668
         Width           =   360
      End
      Begin VB.Label LblTpoSangre 
         AutoSize        =   -1  'True
         Caption         =   "T. Sangre"
         Height          =   195
         Left            =   7410
         TabIndex        =   64
         Top             =   1125
         Width           =   705
      End
      Begin VB.Label lblPeso 
         AutoSize        =   -1  'True
         Caption         =   "Peso "
         Height          =   195
         Left            =   5100
         TabIndex        =   62
         Top             =   1913
         Width           =   405
      End
      Begin VB.Label lblPersNombreN 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Left            =   135
         TabIndex        =   60
         Top             =   1913
         Width           =   630
      End
      Begin VB.Label lblApCasada 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Casada"
         Height          =   195
         Left            =   135
         TabIndex        =   57
         Top             =   1538
         Width           =   1140
      End
      Begin VB.Label lblPersJurEmpleados 
         AutoSize        =   -1  'True
         Caption         =   "N° Empleados"
         Height          =   195
         Left            =   -74880
         TabIndex        =   47
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblPersJurMagnitud 
         AutoSize        =   -1  'True
         Caption         =   "Magnitud Empresarial"
         Height          =   195
         Left            =   -70560
         TabIndex        =   45
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label lblPersJurTpo 
         AutoSize        =   -1  'True
         Caption         =   "Tpo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   44
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label lblPersJurSiglas 
         AutoSize        =   -1  'True
         Caption         =   "Siglas"
         Height          =   195
         Left            =   -74880
         TabIndex        =   43
         Top             =   960
         Width           =   420
      End
      Begin VB.Label lblPersNombre 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social"
         Height          =   195
         Left            =   -74880
         TabIndex        =   42
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblPersNatHijos 
         AutoSize        =   -1  'True
         Caption         =   "N°Hijos"
         Height          =   195
         Left            =   4980
         TabIndex        =   37
         Top             =   1538
         Width           =   525
      End
      Begin VB.Label lblPersNatEstCiv 
         AutoSize        =   -1  'True
         Caption         =   "Estado Civil"
         Height          =   195
         Left            =   4680
         TabIndex        =   36
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label lblPersNombreAM 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1125
         Width           =   1185
      End
   End
   Begin TabDlg.SSTab SSTIdent 
      Height          =   3015
      Left            =   7575
      TabIndex        =   34
      Top             =   3000
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Identificación"
      TabPicture(0)   =   "PersonasMantenimiento.frx":03DA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdPersIDCancelar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPersIDAceptar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPersIDedit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPersIDDel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FEDocs"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPersIDnew"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdPersIDnew 
         Caption         =   "Nuevo"
         Height          =   300
         Left            =   180
         TabIndex        =   24
         Top             =   2460
         Width           =   885
      End
      Begin Sicmact.FlexEdit FEDocs 
         Height          =   1920
         Left            =   135
         TabIndex        =   81
         Top             =   435
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   3387
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Tipo-Numero"
         EncabezadosAnchos=   "350-1200-1200"
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
         ColumnasAEditar =   "X-1-2"
         ListaControles  =   "0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   285
      End
      Begin VB.CommandButton cmdPersIDDel 
         Caption         =   "Eliminar"
         Height          =   300
         Left            =   2070
         TabIndex        =   26
         Top             =   2460
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDedit 
         Caption         =   "Editar"
         Height          =   300
         Left            =   1125
         TabIndex        =   25
         Top             =   2460
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   180
         TabIndex        =   82
         Top             =   2460
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   1125
         TabIndex        =   83
         Top             =   2460
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Label lblPersPersoneria 
      AutoSize        =   -1  'True
      Caption         =   "Personería:"
      Height          =   195
      Left            =   2745
      TabIndex        =   33
      Top             =   165
      Width           =   825
   End
   Begin VB.Label lblPersCod 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   60
      TabIndex        =   31
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "frmPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum

Enum TPersonaTipoInicio
    PersonaConsulta = 1
    PersonaActualiza = 2
End Enum


Dim NomMoverSSTabs As Integer


Dim oBuscaPersona As UPersona
Dim oPersona As DPersona
Dim Nivel1() As String
Dim ContNiv1 As Integer
Dim Nivel2() As String
Dim ContNiv2 As Integer
Dim Nivel3() As String
Dim ContNiv3 As Integer
Dim Nivel4() As String
Dim ContNiv4 As Integer
Dim Nivel5() As String
Dim ContNiv5 As Integer
Dim bEstadoCargando As Boolean

'Para flexEdit de Relacion con Personas
Dim cmdPersRelaEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FERelPersNoMoverdeFila As Integer

'Para flexEdit de Documentos
Dim cmdPersDocEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FEDocsPersNoMoverdeFila As Integer

'Para Fuentes de Ingreso
Dim cmdPersFteIngresoEjecutado As Integer
Dim FEFtePersNoMoverdeFila As Integer

Dim BotonEditar As Boolean
Dim BotonNuevo As Boolean

Public Sub Registrar()
    cmdEditar.Enabled = True
    cmdNuevo.Enabled = True
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Show 1
End Sub

Public Sub Mantenimeinto()
    cmdEditar.Enabled = True
    cmdNuevo.Enabled = False
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Show 1
End Sub

Public Sub Consultar()
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = False
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Show 1
End Sub
Public Sub Inicio(ByVal cPersCod As String, ByVal pTipoInicio As TPersonaTipoInicio)
    
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New DPersona
    TxtBCodPers.Text = cPersCod
    Call oPersona.RecuperaPersona(Trim(TxtBCodPers.Text))
    If oPersona.PersCodigo = "" Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    End If
    Call CargaDatos
    TxtBCodPers.Enabled = False
    cmdNuevo.Enabled = False
    
    If pTipoInicio = PersonaActualiza Then
        cmdEditar.Enabled = True
    Else
        cmdEditar.Enabled = False
    End If
    BotonNuevo = cmdNuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Show 1
End Sub


Private Function ValidaDatosDocumentos() As Boolean
    ValidaDatosDocumentos = True
    If Len(Trim(FEDocs.TextMatrix(FEDocs.Row, 1))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Documento", vbInformation, "Aviso"
        FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    If Len(Trim(FEDocs.TextMatrix(FEDocs.Row, 2))) = 0 Then
        MsgBox "Falta Ingresar el Numero de Documento", vbInformation, "Aviso"
        FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
End Function
Private Function ValidaDatosPersRelacion() As Boolean
    ValidaDatosPersRelacion = True
    
    'Falta Persona a Relacionar
    If Len(Trim(FERelPers.TextMatrix(FERelPers.Row, 1))) = 0 Then
        MsgBox "Falta Ingresar la Persona con la que se va a Relacionar", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta Ingresar el tipo de Relacion
    If Len(Trim(FERelPers.TextMatrix(FERelPers.Row, 3))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Relacion con la Persona", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta el Tipo de Beneficio
    If Len(Trim(FERelPers.TextMatrix(FERelPers.Row, 4))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Beneficio con la Persona", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta ingresar el tipo de Asistenacia Medica Privada AMP
    If Len(Trim(FERelPers.TextMatrix(FERelPers.Row, 6))) = 0 Then
        MsgBox "Falta Ingresar el Tipo Asist. Med. Privada", vbInformation, "Aviso"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
End Function
Private Function ValidaControles() As Boolean
Dim CadTmp As String
    
    ValidaControles = True
    
    If cmbPersPersoneria.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Personeria", vbInformation, "Aviso"
        cmbPersPersoneria.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    If oPersona.Personeria = gPersonaNat Then
    'Valida Controles de Persona Natural
        If Len(Trim(txtPersNombreAP.Text)) = 0 Then
            MsgBox "Falta Ingresar el Apellido Paterno", vbInformation, "Aviso"
            txtPersNombreAP.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If Len(Trim(txtPersNombreAM.Text)) = 0 Then
            MsgBox "Falta Ingresar el Apellido Materno", vbInformation, "Aviso"
            txtPersNombreAM.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If Len(Trim(txtPersNombreN.Text)) = 0 Then
            MsgBox "Falta Ingresar Nombres de la Persona", vbInformation, "Aviso"
            txtPersNombreN.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If cmbPersNatSexo.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Sexo de la Persona", vbInformation, "Aviso"
            cmbPersNatSexo.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If cmbPersNatEstCiv.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Estado Civil de la Persona", vbInformation, "Aviso"
            cmbPersNatEstCiv.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If Len(Trim(txtPersNatHijos.Text)) = 0 Then
            MsgBox "Falta Ingresar el Numero de Hijos", vbInformation, "Aviso"
            txtPersNatHijos.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
'        If Len(Trim(TxtPeso.Text)) = 0 Then
'            MsgBox "Falta Ingresar el Peso de la Persona", vbInformation, "Aviso"
'            ValidaControles = False
'            Exit Function
'        End If
'        If Len(Trim(TxtTalla.Text)) = 0 Then
'            MsgBox "Falta Ingresar la Talla de la Persona", vbInformation, "Aviso"
'            ValidaControles = False
'            Exit Function
'        End If
'        If Len(Trim(TxtTipoSangre.Text)) = 0 Then
'            MsgBox "Falta Ingresar el Tipo de Sangre de la Persona de la Persona", vbInformation, "Aviso"
'            ValidaControles = False
'            Exit Function
'        End If
    Else
    'Valida Controles de Persona Juridica
        If Len(Trim(txtPersNombreRS.Text)) = 0 Then
            MsgBox "Falta Ingresar la razon Social de la Persona", vbInformation, "Aviso"
            txtPersNombreRS.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
    
        If Len(Trim(TxtSiglas.Text)) = 0 Then
            MsgBox "Falta Ingresar la Siglas de la Persona", vbInformation, "Aviso"
            TxtSiglas.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If cmbPersJurTpo.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Tipo de Persona Juridica", vbInformation, "Aviso"
            cmbPersJurTpo.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If cmbPersJurMagnitud.ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Magnitud Empresarial de la Persona", vbInformation, "Aviso"
            cmbPersJurMagnitud.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If Len(Trim(txtPersJurEmpleados.Text)) = 0 Then
            MsgBox "Falta Ingresar el Numero de Empleados ", vbInformation, "Aviso"
            txtPersJurEmpleados.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
    End If
    
    'Valida Datos Generales
    CadTmp = ValidaFecha(CDate(Format(txtPersNacCreac.Text, "dd/mm/yyyy")))
    If Len(CadTmp) > 0 Then
        MsgBox CadTmp, vbInformation, "Aviso"
        txtPersNacCreac.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    If CboPersCiiu.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Ciiu de la Persona", vbInformation, "Aviso"
        CboPersCiiu.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    If cmbPersEstado.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Estado de la Persona", vbInformation, "Aviso"
        cmbPersEstado.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    If cmbPersUbiGeo(4).ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Ubicacion Geografica de la Persona", vbInformation, "Aviso"
        cmbPersUbiGeo(4).SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 1
        Exit Function
    End If
    
    If Len(Trim(txtPersDireccDomicilio.Text)) = 0 Then
        MsgBox "Falta Ingresar el Domicilio de la Persona", vbInformation, "Aviso"
        txtPersDireccDomicilio.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 1
        Exit Function
    End If
    
    If cmbPersDireccCondicion.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Condicion del Domicilio de la Persona", vbInformation, "Aviso"
        cmbPersDireccCondicion.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 1
        Exit Function
    End If
    
End Function

Private Sub HabilitaControlesPersona(ByVal pbBloqueo As Boolean)
    
    cmbPersPersoneria.Enabled = pbBloqueo
    
    If Not oPersona Is Nothing Then
        If oPersona.Personeria = 0 Then
            'Ficha de Persona natural
            Call HabilitaFichaPersonaNat(pbBloqueo)
        Else
            'Ficha de Persona Juridica
            Call HabilitaFichaPersonaJur(pbBloqueo)
        End If
    Else
        'Ficha de Persona natural
        Call HabilitaFichaPersonaNat(pbBloqueo)
        'Ficha de Persona Juridica
        Call HabilitaFichaPersonaJur(pbBloqueo)
    End If
        
    'Ficha de Relaciones de Persona
    FERelPers.lbEditarFlex = pbBloqueo
    cmdPersRelacNew.Enabled = pbBloqueo
    cmdPersRelacEditar.Enabled = pbBloqueo
    cmdPersRelacDel.Enabled = pbBloqueo
    cmdPersRelacAceptar.Enabled = pbBloqueo
    cmdPersRelacCancelar.Enabled = pbBloqueo
    
    'Ficha de Fuentes de Ingreso
    HabilitaControlesPersonaFtesIngreso (pbBloqueo)
    'Firma
    CmdActFirma.Enabled = pbBloqueo
        
    'Ficha de Datos Generales
    txtPersNacCreac.Enabled = pbBloqueo
    txtPersTelefono.Enabled = pbBloqueo
    txtPersTelefono2.Enabled = pbBloqueo
    TxtEmail.Enabled = pbBloqueo
    CboPersCiiu.Enabled = pbBloqueo
    cmbPersEstado.Enabled = pbBloqueo
    
    'Ficha de Ubicacion Geografica
    cmbPersUbiGeo(0).Enabled = pbBloqueo
    cmbPersUbiGeo(1).Enabled = pbBloqueo
    cmbPersUbiGeo(2).Enabled = pbBloqueo
    cmbPersUbiGeo(3).Enabled = pbBloqueo
    cmbPersUbiGeo(4).Enabled = pbBloqueo
    txtPersDireccDomicilio.Enabled = pbBloqueo
    cmbPersDireccCondicion.Enabled = pbBloqueo
    
    'Ficha de Identificacion
    SSTIdent.Enabled = pbBloqueo
    FEDocs.lbEditarFlex = IIf(pbBloqueo, False, True)
End Sub
Private Sub HabilitaControlesPersonaFtesIngreso(ByVal pbBloqueo As Boolean)
    FEFteIng.lbEditarFlex = pbBloqueo
    CmdFteIngNuevo.Enabled = pbBloqueo
    CmdFteIngEditar.Enabled = pbBloqueo
    CmdFteIngEliminar.Enabled = pbBloqueo
    'CmdPersFteConsultar.Enabled = pbBloqueo
End Sub
Private Sub HabilitaFichaPersonaNat(ByVal pbFicActiva As Boolean)
    lblPersNombreAP.Enabled = pbFicActiva
    txtPersNombreAP.Enabled = pbFicActiva
    lblPersNombreAM.Enabled = pbFicActiva
    txtPersNombreAM.Enabled = pbFicActiva
    lblApCasada.Enabled = pbFicActiva
    TxtApellidoCasada.Enabled = pbFicActiva
    lblPersNombreN.Enabled = pbFicActiva
    txtPersNombreN.Enabled = pbFicActiva
    lblPersNatSexo.Enabled = pbFicActiva
    cmbPersNatSexo.Enabled = pbFicActiva
    lblPersNatEstCiv.Enabled = pbFicActiva
    cmbPersNatEstCiv.Enabled = pbFicActiva
    lblPersNatHijos.Enabled = pbFicActiva
    txtPersNatHijos.Enabled = pbFicActiva
    lblPeso.Enabled = pbFicActiva
    TxtPeso.Enabled = pbFicActiva
    LblTalla.Enabled = pbFicActiva
    TxtTalla.Enabled = pbFicActiva
    LblTpoSangre.Enabled = pbFicActiva
    TxtTipoSangre.Enabled = pbFicActiva
    
End Sub
Private Sub HabilitaFichaPersonaJur(ByVal pbFicActiva As Boolean)
    txtPersNombreRS.Enabled = pbFicActiva
    TxtSiglas.Enabled = pbFicActiva
    cmbPersJurTpo.Enabled = pbFicActiva
    txtPersJurEmpleados.Enabled = pbFicActiva
    cmbPersJurMagnitud.Enabled = pbFicActiva
    lblPersNombre.Enabled = pbFicActiva
    lblPersJurSiglas.Enabled = pbFicActiva
    lblPersJurTpo.Enabled = pbFicActiva
    lblPersJurEmpleados.Enabled = pbFicActiva
    lblPersJurMagnitud.Enabled = pbFicActiva
End Sub

Private Sub CargaControles()
Dim sSql As String
Dim Conn As DConecta
Dim R As ADODB.Recordset
Dim i As Integer
Dim oConstante As DConstante
On Error GoTo ERRORCargaControles
    bEstadoCargando = True
    cmdPersDocEjecutado = 0
    cmdPersRelaEjecutado = 0
    cmdPersFteIngresoEjecutado = 0
    NomMoverSSTabs = -1
    FERelPersNoMoverdeFila = -1
    FEDocsPersNoMoverdeFila = -1
    FEFtePersNoMoverdeFila = -1
    Set oConstante = New DConstante
    Set Conn = New DConecta
    Conn.AbreConexion
    
    txtPersNacCreac.Text = Format(gdFecSis, "dd/mm/yyyy")
    'Carga Combo de Flex de Relaciones de Persona
    FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
    
    'Carga Combo de Documentos
    FEDocs.CargaCombo oConstante.RecuperaConstantes(gPersIdTipo)
    
    Set oConstante = Nothing
    'Carga Condiciones de Domicilio
    cmbPersDireccCondicion.Clear
    Call CargaComboConstante(gPersCondDomic, cmbPersDireccCondicion)
    
    'Carga Personeria
    cmbPersPersoneria.Clear
    Call CargaComboConstante(gPersPersoneria, cmbPersPersoneria)
    
    'Carga Tipos de Sexo de Personas
    cmbPersNatSexo.AddItem "FEMENINO" & Space(50) & "F"
    cmbPersNatSexo.AddItem "MASCULINO" & Space(50) & "M"
    
    'Carga Magnitud Empresarial
    cmbPersJurMagnitud.Clear
    Call CargaComboConstante(gPersJurMagnitud, cmbPersJurMagnitud)
        
    'Carga Relacion de Persona con Institucion
    CmbRelPers.Clear
    Call CargaComboConstante(gPersRelacionInst, CmbRelPers)
    
    'Carga Condicion de Domicilio
    cmbPersDireccCondicion.Clear
    Call CargaComboConstante(gPersCondDomic, cmbPersDireccCondicion)
    
    'Carga Estado Civil
    cmbPersNatEstCiv.Clear
    Call CargaComboConstante(gPersEstadoCivil, cmbPersNatEstCiv)
    
    'Carga Ubicaciones Geograficas
    Call CargaUbicacionesGeograficas
    
    
    'Carga Ciiu
    sSql = "Select cCIIUcod,cCIIUdescripcion from CIIU Order by cCIIUdescripcion"
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        CboPersCiiu.AddItem Trim(R!cCIIUdescripcion) & Space(100) & Trim(R!cCIIUcod)
        R.MoveNext
    Loop
    R.Close
    
    'CARGA TIPOS DE PERSONA JURIDICA
    cmbPersJurTpo.Clear
    sSql = "Select cPersJurTpoCod,cPersJurTpoDesc  from persjurtpo Order by cPersJurTpoDesc"
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        cmbPersJurTpo.AddItem Trim(R!cPersJurTpoDesc) & Space(100) & Trim(R!cPersJurTpoCod)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    bEstadoCargando = True
    Exit Sub
    
    
ERRORCargaControles:
    MsgBox Err.Description, vbExclamation, "Aviso"
    
End Sub
Private Sub CargaControlEstadoPersona(ByVal pnTipoPers As Integer)
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset

On Error GoTo ERRORCargaControles
    Set Conn = New DConecta
    Conn.AbreConexion
    cmbPersEstado.Clear
    
    'Carga Estados de la Persona
    sSql = "Select nConsValor,cConsDescripcion From Constante Where nConsCod = " & Trim(Str(gPersEstado)) & " and nConsValor <> " & Trim(Str(gPersEstado))
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        If pnTipoPers = 0 Then
            If Len(Trim(R!nConsValor)) = 1 Then
                cmbPersEstado.AddItem Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(Str(R!nConsValor)), 2)
            End If
        Else
            If Len(Trim(R!nConsValor)) > 1 Then
                cmbPersEstado.AddItem Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(R!nConsValor), 2)
            End If
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    Conn.CierraConexion
    Set Conn = Nothing
    Exit Sub
    
ERRORCargaControles:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
Private Sub DistribuyeApellidos(ByVal bApellCasada As Boolean)
    If bApellCasada = True Then
        lblApCasada.Visible = True
        TxtApellidoCasada.Visible = True
        lblPersNombreAP.Top = 660
        lblPersNombreAP.Left = 120
        txtPersNombreAP.Top = 660
        txtPersNombreAP.Left = 1680
        lblPersNombreAM.Top = 1065
        lblPersNombreAM.Left = 120
        txtPersNombreAM.Top = 1065
        txtPersNombreAM.Left = 1680
        lblApCasada.Top = 1485
        lblApCasada.Left = 135
        TxtApellidoCasada.Top = 1485
        TxtApellidoCasada.Left = 1680
        lblPersNombreN.Top = 1860
        lblPersNombreN.Left = 135
        txtPersNombreN.Top = 1875
        txtPersNombreN.Left = 1680
    Else
        lblPersNombreAP.Top = 660
        lblPersNombreAP.Left = 120
        txtPersNombreAP.Top = 660
        txtPersNombreAP.Left = 1680
        lblPersNombreAM.Top = 1140
        lblPersNombreAM.Left = 120
        txtPersNombreAM.Top = 1140
        txtPersNombreAM.Left = 1680
        lblApCasada.Visible = False
        TxtApellidoCasada.Visible = False
        lblPersNombreN.Top = 1620
        lblPersNombreN.Left = 120
        txtPersNombreN.Top = 1620
        txtPersNombreN.Left = 1680
    End If
End Sub
Private Sub CargaDocumentos()
Dim i As Integer
    Call LimpiaFlex(FEDocs)
    For i = 0 To oPersona.NumeroDocumentos - 1
        FEDocs.AdicionaFila
        'Columna de Tipo de Documento
        FEDocs.TextMatrix(i + 1, 1) = Trim(oPersona.ObtenerTipoDoc(i))
        'Columna de Numero de Documento
        FEDocs.TextMatrix(i + 1, 2) = oPersona.ObtenerNumeroDoc(i)
    Next i
End Sub
Private Sub CargaRelacionesPersonas()
Dim i As Integer
    FERelPers.lbEditarFlex = True
    Call LimpiaFlex(FERelPers)
    For i = 0 To oPersona.NumeroRelacPers - 1
        FERelPers.AdicionaFila
        'Codigo
        FERelPers.TextMatrix(i + 1, 1) = oPersona.ObtenerRelaPersCodigo(i)
        'Apellidos y Nombres
        FERelPers.TextMatrix(i + 1, 2) = oPersona.ObtenerRelaPersNombres(i)
        'Relacion
        FERelPers.TextMatrix(i + 1, 3) = oPersona.ObtenerRelaPersRelacion(i)
        'Beneficiario
        FERelPers.TextMatrix(i + 1, 4) = oPersona.ObtenerRelaPersBenef(i)
        'Beneficiario Porcentaje
        FERelPers.TextMatrix(i + 1, 5) = Format(oPersona.ObtenerRelaPersBenefPorc(i), "#0.00")
        'Asistencia medica Privada
        FERelPers.TextMatrix(i + 1, 6) = oPersona.ObtenerRelaPersAMP(i)
    Next i
    FERelPers.lbEditarFlex = False
End Sub
Private Sub CargaFuentesIngreso()
Dim i As Integer
    Call LimpiaFlex(FEFteIng)
    For i = 0 To oPersona.NumeroFtesIngreso - 1
        FEFteIng.AdicionaFila
        FEFteIng.TextMatrix(i + 1, 0) = i + 1
        FEFteIng.TextMatrix(i + 1, 1) = IIf(oPersona.ObtenerFteIngTipo(i) = "1", "D", "I") 'Tipo de Fte de Ingreso
        FEFteIng.TextMatrix(i + 1, 2) = oPersona.ObtenerFteIngRazonSocial(i) 'Razon Social de Fte de Ingreso
        FEFteIng.TextMatrix(i + 1, 3) = Format(oPersona.ObtenerFteIngFecCaducac(i), "dd/mm/yyyy") 'Fecha de Caducacion de la Fte de Ingreso
        FEFteIng.TextMatrix(i + 1, 4) = Format(oPersona.ObtenerFteIngFecEval(i), "dd/mm/yyyy") 'Fecha de Evaluacion de la Fte de Ingreso
        FEFteIng.TextMatrix(i + 1, 5) = IIf(oPersona.ObtenerFteIngMoneda(i) = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
    Next i
End Sub
Private Sub CargaDatos()
Dim i As Integer
     
    bEstadoCargando = True
    SSTDatosGen.Tab = 0
    SSTabs.Tab = 0
    'Carga Personeria
    cmbPersPersoneria.ListIndex = IndiceListaCombo(cmbPersPersoneria, Trim(Str(oPersona.Personeria)))
    
    'Habilita o Deshabilita Ficha de Persona Juridica
    
    Call HabilitaFichaPersonaJur(False)
    Call HabilitaFichaPersonaNat(False)
    
    
    'Carga Ubicacion Georgrafica
    If Len(Trim(oPersona.UbicacionGeografica)) = 12 Then
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "PER")
        cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "1" & Mid(oPersona.UbicacionGeografica, 2, 2) & String(9, "0"))
        cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "2" & Mid(oPersona.UbicacionGeografica, 2, 4) & String(7, "0"))
        cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & "3" & Mid(oPersona.UbicacionGeografica, 2, 6) & String(5, "0"))
        cmbPersUbiGeo(4).ListIndex = IndiceListaCombo(cmbPersUbiGeo(4), Space(30) & oPersona.UbicacionGeografica)
    Else
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & oPersona.UbicacionGeografica)
        cmbPersUbiGeo(1).Clear
        cmbPersUbiGeo(1).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(1).ListIndex = 0
        cmbPersUbiGeo(2).Clear
        cmbPersUbiGeo(2).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(2).ListIndex = 0
        cmbPersUbiGeo(3).Clear
        cmbPersUbiGeo(3).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(3).ListIndex = 0
        cmbPersUbiGeo(4).Clear
        cmbPersUbiGeo(4).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(4).ListIndex = 0
    End If
    
    'Carga Direccion
    txtPersDireccDomicilio.Text = oPersona.Domicilio
    
    'Selecciona la Condicion del Doicilio
    cmbPersDireccCondicion.ListIndex = IndiceListaCombo(cmbPersDireccCondicion, oPersona.CondicionDomicilio)
    
    'Carga Ficha 1
    If oPersona.Personeria = gPersonaNat Then
        txtPersNombreAP.Text = oPersona.ApellidoPaterno
        txtPersNombreAM.Text = oPersona.ApellidoMaterno
        txtPersNombreN.Text = oPersona.Nombres
    Else
        txtPersNombreRS.Text = oPersona.NombreCompleto
    End If
    TxtTalla.Text = Format(oPersona.Talla, "#0.00")
    TxtPeso.Text = Format(oPersona.Peso, "#0.00")
    TxtEmail.Text = oPersona.Email
    TxtTipoSangre.Text = oPersona.TipoSangre
    txtPersTelefono2.Text = oPersona.Telefonos2
    If oPersona.Sexo = "F" Then
        TxtApellidoCasada.Text = oPersona.ApellidoCasada
        cmbPersNatSexo.ListIndex = 0
        Call DistribuyeApellidos(True)
    Else
        cmbPersNatSexo.ListIndex = 1
        Call DistribuyeApellidos(False)
    End If
    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, oPersona.EstadoCivil)
    txtPersNatHijos.Text = Trim(Str(oPersona.Hijos))
    
    'Carga Datos Generales
    txtPersNacCreac.Text = Format(oPersona.FechaNacimiento, "dd/mm/yyyy")
    txtPersTelefono.Text = oPersona.Telefonos
    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, oPersona.CIIU)
    
    Call CargaControlEstadoPersona(oPersona.Personeria)
    cmbPersEstado.ListIndex = IndiceListaCombo(cmbPersEstado, oPersona.Estado)
    
    'Carga Razon Social
    TxtSiglas.Text = oPersona.Siglas
    
    'Selecciona el Tipo de Persona Juridica
    cmbPersJurTpo.ListIndex = IndiceListaCombo(cmbPersJurTpo, Trim(Str(IIf(oPersona.TipoPersonaJur = "", -1, oPersona.TipoPersonaJur))))

    'Selecciona la magnitud Empresarial
    cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, Trim(oPersona.MagnitudEmpresarial))
    
    'Carga Numero de Empleados
    txtPersJurEmpleados.Text = Trim(Str(oPersona.NumerosEmpleados))
    
    'Carga Los Documentos de la Persona
    Call CargaDocumentos
    
    'Carga las Relaciones de las Personas
    Call CargaRelacionesPersonas
    
    'Carga las Fuentes de Ingresos de las Personas
    Call CargaFuentesIngreso
    
    'Carga Firma
    Call IDBFirma.CargarFirma(oPersona.RFirma)
    
    If oPersona.Personeria = 0 Then
        SSTabs.Tab = 0
    Else
        SSTabs.Tab = 1
    End If
    bEstadoCargando = False
End Sub

Private Sub CargaUbicacionesGeograficas()
Dim Conn As DConecta
Dim sSql As String
Dim R As ADODB.Recordset
Dim i As Integer
Dim nPos As Integer

On Error GoTo ErrCargaUbicacionesGeograficas
    Set Conn = New DConecta
    'Carga Niveles
    sSql = "select *,1 p from UbicacionGeografica where len(cUbiGeoCod)=3 "
    sSql = sSql & " Union "
    sSql = sSql & " Select *, 2 p from UbicacionGeografica where cUbiGeoCod like '1%'"
    sSql = sSql & " Union "
    sSql = sSql & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '2%' "
    sSql = sSql & " Union "
    sSql = sSql & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '3%' "
    sSql = sSql & " Union "
    sSql = sSql & " select *, 5 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
    ContNiv1 = 0
    ContNiv2 = 0
    ContNiv3 = 0
    ContNiv4 = 0
    ContNiv5 = 0
    
    Conn.AbreConexion
    Set R = Conn.CargaRecordSet(sSql)
    Do While Not R.EOF
        Select Case R!P
            Case 1 'Pais
                ContNiv1 = ContNiv1 + 1
                ReDim Preserve Nivel1(ContNiv1)
                Nivel1(ContNiv1 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 2 ' Departamento
                ContNiv2 = ContNiv2 + 1
                ReDim Preserve Nivel2(ContNiv2)
                Nivel2(ContNiv2 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 3 'Provincia
                ContNiv3 = ContNiv3 + 1
                ReDim Preserve Nivel3(ContNiv3)
                Nivel3(ContNiv3 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 4 'Distrito
                ContNiv4 = ContNiv4 + 1
                ReDim Preserve Nivel4(ContNiv4)
                Nivel4(ContNiv4 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
            Case 5 'Zona
                ContNiv5 = ContNiv5 + 1
                ReDim Preserve Nivel5(ContNiv5)
                Nivel5(ContNiv5 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
        End Select
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Conn.CierraConexion
    Set Conn = Nothing
    
    'Carga el Nivel1 en el Control
    cmbPersUbiGeo(0).Clear
    For i = 0 To ContNiv1 - 1
        cmbPersUbiGeo(0).AddItem Nivel1(i)
        If Trim(Right(Nivel1(i), 10)) = "PER" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbPersUbiGeo(2).Clear
    cmbPersUbiGeo(3).Clear
    cmbPersUbiGeo(4).Clear
    Exit Sub
    
ErrCargaUbicacionesGeograficas:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub
Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TTipoCombo)
Dim i As Integer
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboDpto
            cmbPersUbiGeo(1).Clear
            If sCodigo = "PER" Then
                For i = 0 To ContNiv2 - 1
                    cmbPersUbiGeo(1).AddItem Nivel2(i)
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboProv
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(2).Clear
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel3(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel4(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(4).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv5 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel5(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(4).AddItem Nivel5(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(4).AddItem psValor
            End If
    End Select
End Sub

Private Sub CboPersCiiu_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 20))
    End If
End Sub

Private Sub CboPersCiiu_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 20))
    End If
End Sub

Private Sub CboPersCiiu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPersEstado.SetFocus
    End If
End Sub

Private Sub cmbPersDireccCondicion_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CondicionDomicilio = Trim(Right(cmbPersDireccCondicion.Text, 10))
    End If
End Sub

Private Sub cmbPersDireccCondicion_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CondicionDomicilio = Trim(Right(cmbPersDireccCondicion.Text, 10))
    End If
End Sub

Private Sub cmbPersEstado_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Estado = Trim(Right(cmbPersEstado.Text, 10))
      End If
End Sub

Private Sub cmbPersEstado_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Estado = Trim(Right(cmbPersEstado.Text, 10))
    End If
End Sub

Private Sub cmbPersEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SSTDatosGen.Tab = 1
        cmbPersUbiGeo(0).SetFocus
    End If
End Sub

Private Sub cmbPersJurMagnitud_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
        oPersona.MagnitudEmpresarial = Trim(Right(cmbPersJurMagnitud.Text, 15))
    End If
End Sub

Private Sub cmbPersJurMagnitud_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.MagnitudEmpresarial = Trim(Right(cmbPersJurMagnitud.Text, 15))
    End If
End Sub

Private Sub cmbPersJurMagnitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNacCreac.SetFocus
    End If
End Sub

Private Sub cmbPersJurTpo_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersJurTpo.Text, 10))
        End If
    End If
End Sub

Private Sub cmbPersJurTpo_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersJurTpo.Text, 10))
    End If
End Sub

Private Sub cmbPersJurTpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersJurEmpleados.SetFocus
    End If
End Sub

Private Sub cmbPersNatEstCiv_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
  End If
End Sub

Private Sub cmbPersNatEstCiv_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
    End If
End Sub

Private Sub cmbPersNatEstCiv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNatHijos.SetFocus
    End If
End Sub

Private Sub cmbPersNatSexo_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.Sexo = Trim(RTrim(cmbPersNatSexo.Text, 10))
  End If
End Sub

Private Sub cmbPersNatSexo_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
        If oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
        End If
    End If
End Sub

Private Sub cmbPersNatSexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPersNatEstCiv.SetFocus
    End If
End Sub

Private Sub cmbPersPersoneria_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersPersoneria.Text, 15))
    End If
End Sub

Private Sub cmbPersPersoneria_Click()
    If Not bEstadoCargando Then
        If oPersona.Personeria <> CInt(Trim(Right(cmbPersPersoneria.Text, 15))) Then
            oPersona.Personeria = Trim(Right(IIf(cmbPersPersoneria.Text = "", "0", cmbPersPersoneria.Text), 15))
            If oPersona.Personeria <> 0 Then
                Call HabilitaFichaPersonaJur(True)
                Call HabilitaFichaPersonaNat(False)
            Else
                Call HabilitaFichaPersonaJur(False)
                Call HabilitaFichaPersonaNat(True)
            End If
            cmbPersEstado.ListIndex = -1
        End If
    End If
    Call CargaControlEstadoPersona(oPersona.Personeria)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Personeria = Trim(Right(IIf(cmbPersPersoneria.Text = "", "0", cmbPersPersoneria.Text), 15))
    End If
End Sub

Private Sub cmbPersPersoneria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPersNombreAP.Enabled Then
            txtPersNombreAP.SetFocus
        Else
            txtPersNombreRS.SetFocus
        End If
    End If
End Sub

Private Sub cmbPersUbiGeo_Change(Index As Integer)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
    End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
        Select Case Index
            Case 0 'Combo Pais
                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboDpto)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(2).Clear
                    cmbPersUbiGeo(3).Clear
                    cmbPersUbiGeo(4).Clear
                End If
            Case 1 'Combo Dpto
                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboProv)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(3).Clear
                    cmbPersUbiGeo(4).Clear
                End If
            Case 2 'Combo Provincia
                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboDist)
                If Not bEstadoCargando Then
                    cmbPersUbiGeo(4).Clear
                End If
            Case 3 'Combo Distrito
                Call ActualizaCombo(cmbPersUbiGeo(3).Text, ComboZona)
        End Select
        
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
        End If

End Sub

Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            txtPersDireccDomicilio.SetFocus
        End If
    End If
End Sub


Private Sub CmbRelPers_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.PersonaRelaInst = Trim(RTrim(CmbRelPers.Text, 10))
  End If
End Sub

Private Sub CmbRelPers_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PersonaRelaInst = Trim(Right(CmbRelPers.Text, 10))
    End If
End Sub

Private Sub CmbRelPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersTelefono.SetFocus
    End If
End Sub

Private Sub CmdActFirma_Click()
Dim sRuta As String
    CdlgImg.nHwd = Me.hwnd
    CdlgImg.Show
    sRuta = CdlgImg.Ruta
    If Len(Trim(sRuta)) > 0 Then
        IDBFirma.RutaImagen = sRuta
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
            Call IDBFirma.GrabarFirma(oPersona.RFirma, oPersona.PersCodigo, "")
        Else
            Call IDBFirma.GrabarFirma(oPersona.RFirma, oPersona.PersCodigo, "")
        End If
    End If
End Sub
    
Private Sub CmdEditar_Click()
    If oPersona Is Nothing Then
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPersona.PersCodigo = "" Then
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    CmdPersAceptar.Enabled = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Enabled = True
    CmdPersCancelar.Visible = True
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = False
    Call HabilitaControlesPersona(True)
    
    oPersona.TipoActualizacion = PersFilaModificada
    
End Sub


Private Sub CmdFteIngEditar_Click()
    If oPersona.NumeroFtesIngreso > 0 Then
        Call frmFteIngresos.Editar(FEFteIng.Row - 1, oPersona)
    Else
        MsgBox "No Existe Fuentes de Ingreso para Editar", vbInformation, "Aviso"
    End If
End Sub


Private Sub CmdFteIngEliminar_Click()
    If MsgBox("Esta Seguro que Desea Eliminar esta Fuente de Ingreso", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarFteIngTipoAct(PersFilaEliminda, FEFteIng.Row - 1)
        Call CmdPersAceptar_Click
        Call CargaFuentesIngreso
    End If
End Sub

Private Sub CmdFteIngNuevo_Click()

'    CmdFteIngNuevo.Enabled = False
'    CmdFteIngEliminar.Enabled = False
'    CmdFteIngEditar.Enabled = False
'    CmdPersFteConsultar.Enabled = False
'    NomMoverSSTabs = SSTabs.Tab
'    FEFteIng.lbEditarFlex = True
'    cmdPersFteIngresoEjecutado = 1
'    FEFtePersNoMoverdeFila = FEFteIng.Rows - 1
    Call frmFteIngresos.NuevaFteIngreso(oPersona)
    Call CargaFuentesIngreso
End Sub

Private Sub CmdNuevo_Click()
    Call HabilitaControlesPersona(True)
    Call LimpiarPantalla
    
    'Inicializa Controles
    TxtBCodPers.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTabs.Enabled = True
    SSTabs.Tab = 0
    SSTDatosGen.Enabled = True
    SSTDatosGen.Tab = 0
    SSTIdent.Enabled = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    If oPersona Is Nothing Then
       Set oPersona = New DPersona
    End If
    Call oPersona.NuevaPersona
    oPersona.TipoActualizacion = PersFilaNueva
    
    Call HabilitaControlesPersonaFtesIngreso(False)
    cmbPersNatSexo.ListIndex = 0
    cmbPersPersoneria.ListIndex = 0
    cmbPersPersoneria.SetFocus
    
End Sub

Private Sub CmdPersAceptar_Click()
Dim oPersonaNeg As NPersona
Dim R As ADODB.Recordset

    If Not ValidaControles Then
        Exit Sub
    End If
    
    'Verifica Homonimia
    If oPersona.TipoActualizacion = PersFilaNueva Then
        Set oPersonaNeg = New NPersona
        Set R = oPersonaNeg.VerificandoHomonimia(oPersona)
        If Not R.BOF And Not R.EOF Then
            Call frmMuestraHomonimia.Inicio(R)
        End If
        R.Close
        Set R = Nothing
        Set oPersonaNeg = Nothing
    End If

    If MsgBox("Se va a proceder a guardar los datos de la Persona, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
        
    
    On Error GoTo ErrorCmdPersAceptar
    
    'Verificando Duplicidad de Documento
    Set oPersonaNeg = New NPersona
    If oPersonaNeg.VerificandoDuplicidadDoc(oPersona) <> -1 Then
        MsgBox "Documento " & Trim(Left(FEDocs.TextMatrix(FEDocs.Row, 1), 30)) & " se Encuentra Duplicado", vbInformation, "Aviso"
        cmdPersIDedit.SetFocus
        Set oPersonaNeg = Nothing
        Exit Sub
    End If
    Set oPersonaNeg = Nothing
    
    'Verificando Tamaño de Documento
    Set oPersonaNeg = New NPersona
    If oPersonaNeg.VerificandoTamanioDoc(oPersona) <> -1 Then
        MsgBox "Documento " & Trim(Left(FEDocs.TextMatrix(FEDocs.Row, 1), 30)) & " numero de digitos Incorrecto", vbInformation, "Aviso"
        cmdPersIDedit.SetFocus
        Set oPersonaNeg = Nothing
        Exit Sub
    End If
    Set oPersonaNeg = Nothing
    Screen.MousePointer = 11
    
    
    'Genera el Nuevo codigo de la Persona
    If oPersona.TipoActualizacion = PersFilaNueva Then
        Set oPersonaNeg = New NPersona
        oPersona.PersCodigo = oPersonaNeg.GenerarNuevoCodigoPersona(oPersona, gsCodCMAC & gsCodAge)
        Set oPersonaNeg = Nothing
    End If
    oPersona.CampoActualizacion = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, Right(gsCodAge, 2))
    oPersona.GrabarCambiosPersona gsCodAge
    Call HabilitaControlesPersona(False)
    TxtBCodPers.Text = oPersona.PersCodigo
    
    Screen.MousePointer = 0
    
    MsgBox "Datos Grabados", vbInformation, "Aviso"
    Call HabilitaControlesPersona(False)
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Value = False
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    NomMoverSSTabs = -1
    
    'Habilita Todos los Controles
    SSTabs.Enabled = True
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdSalir.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Visible = False
    cmdNuevo.Enabled = BotonNuevo
    cmdEditar.Enabled = BotonEditar
    
        
    
    Exit Sub
    
ErrorCmdPersAceptar:
    MsgBox Err.Description, vbExclamation, "Aviso"
    Call CmdPersCancelar_Click
End Sub

Private Sub CmdPersCancelar_Click()
    'Habilita Todos los Controles
    NomMoverSSTabs = -1
    SSTabs.Enabled = True
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdSalir.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Visible = False
    
    If Trim(TxtBCodPers.Text) <> "" Then
        Call TxtBCodPers_EmiteDatos
    Else
        Call LimpiarPantalla
        Call HabilitaControlesPersona(False)
        Call HabilitaControlesPersonaFtesIngreso(False)
    End If
    cmdNuevo.Enabled = BotonNuevo
    cmdEditar.Enabled = BotonEditar
End Sub

Private Sub CmdPersFteConsultar_Click()
    If oPersona.NumeroFtesIngreso > 0 Then
        Call frmFteIngresos.ConsultarFuenteIngreso(FEFteIng.Row - 1, oPersona)
    Else
        MsgBox "No se puede Consultar la Fuente de Ingreso", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdPersIDAceptar_Click()

    If Not ValidaDatosDocumentos Then
        Exit Sub
    End If
        
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    
    If cmdPersDocEjecutado = 1 Then
        Call oPersona.AdicionaDocumento(PersFilaNueva, FEDocs.TextMatrix(FEDocs.Row, 2), FEDocs.TextMatrix(FEDocs.Row, 1))
        Call oPersona.ActualizarDocsTipoAct(PersFilaNueva, FEDocs.Row - 1)
    Else
        If cmdPersDocEjecutado = 2 Then
            If oPersona.ObtenerDocTipoAct(FEDocs.Row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarDocsTipoAct(PersFilaModificada, FEDocs.Row - 1)
            End If
        End If
    End If
    
    'Tipo de Docmumento
    Call oPersona.ActualizaDocsTipo(FEDocs.TextMatrix(FEDocs.Row, 1), FEDocs.Row - 1)
    'Tipo de Numero
    Call oPersona.ActualizaDocsNumero(FEDocs.TextMatrix(FEDocs.Row, 2), FEDocs.Row - 1)
    
    'Habilitar Controles
    cmdPersDocEjecutado = 0
    FEDocsPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEDocs.lbEditarFlex = False
    cmdPersIDnew.Enabled = True
    cmdPersIDnew.Visible = True
    cmdPersIDedit.Enabled = True
    cmdPersIDedit.Visible = True
    cmdPersIDDel.Enabled = True
    cmdPersIDAceptar.Enabled = False
    cmdPersIDAceptar.Visible = False
    cmdPersIDCancelar.Enabled = False
    cmdPersIDCancelar.Visible = False
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTabs.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Enabled = True
        CmdPersCancelar.Visible = True
    End If
    FEDocs.SetFocus
End Sub

Private Sub cmdPersIDCancelar_Click()

    Call CargaDocumentos
    
    'Habilitar Controles
    FEDocsPersNoMoverdeFila = -1
    cmdPersDocEjecutado = -1
    NomMoverSSTabs = -1
    FEDocs.lbEditarFlex = False
    cmdPersIDnew.Enabled = True
    cmdPersIDedit.Enabled = True
    cmdPersIDDel.Enabled = True
    cmdPersIDnew.Visible = True
    cmdPersIDedit.Visible = True
    cmdPersIDDel.Visible = True
    cmdPersIDAceptar.Enabled = False
    cmdPersIDCancelar.Enabled = False
    cmdPersIDAceptar.Visible = False
    cmdPersIDCancelar.Visible = False
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTabs.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FEDocs.SetFocus

End Sub

Private Sub cmdPersIDDel_Click()
    If MsgBox("Esta Seguro que Desea Eliminar este Documento", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarDocsTipoAct(PersFilaEliminda, FEDocs.Row - 1)
        Call FEDocs.EliminaFila(FEDocs.Row)
        Call CmdPersAceptar_Click
        Call CargaDocumentos
    End If
End Sub

Private Sub cmdPersIDedit_Click()
    cmdPersDocEjecutado = 2
    FEDocsPersNoMoverdeFila = FEDocs.Row
    FEDocs.lbEditarFlex = True
    FEDocs.SetFocus
    cmdPersIDnew.Enabled = False
    cmdPersIDnew.Visible = False
    cmdPersIDedit.Enabled = False
    cmdPersIDedit.Visible = False
    cmdPersIDDel.Enabled = False
    cmdPersIDAceptar.Enabled = True
    cmdPersIDAceptar.Visible = True
    cmdPersIDCancelar.Enabled = True
    cmdPersIDCancelar.Visible = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = False
    cmbPersPersoneria.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTDatosGen.Enabled = False
    SSTabs.Enabled = False
    
End Sub


Private Sub cmdPersIDnew_Click()
    FEDocs.AdicionaFila
    cmdPersDocEjecutado = 1
    FEDocsPersNoMoverdeFila = FEDocs.Rows - 1
    FEDocs.lbEditarFlex = True
    FEDocs.SetFocus
    cmdPersIDnew.Enabled = False
    cmdPersIDnew.Visible = False
    cmdPersIDedit.Enabled = False
    cmdPersIDedit.Visible = False
    cmdPersIDDel.Enabled = False
    cmdPersIDAceptar.Enabled = True
    cmdPersIDAceptar.Visible = True
    cmdPersIDCancelar.Enabled = True
    cmdPersIDCancelar.Visible = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = False
    cmbPersPersoneria.Enabled = False
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTDatosGen.Enabled = False
    SSTabs.Enabled = False
End Sub

Private Sub cmdPersRelacAceptar_Click()
    
    If Len(Trim(FERelPers.TextMatrix(FERelPers.Row, 5))) = 0 Then
        FERelPers.TextMatrix(FERelPers.Row, 5) = "0.00"
    End If
    
    If Not ValidaDatosPersRelacion Then
        Exit Sub
    End If
    
   If cmdPersRelaEjecutado = 1 Then
        Call oPersona.AdicionaPersonaRelacion
        Call oPersona.ActualizarRelacPersTipoAct(PersFilaNueva, FERelPers.Row - 1)
    Else
        If cmdPersRelaEjecutado = 2 Then
            If oPersona.ObtenerRelaPersTipoAct(FERelPers.Row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRelacPersTipoAct(PersFilaModificada, FERelPers.Row - 1)
            End If
        End If
    End If
    
    'Apellidos y Nombres
    Call oPersona.ActualizaPersRelaPersona(FERelPers.TextMatrix(FERelPers.Row, 2), FERelPers.TextMatrix(FERelPers.Row, 1), FERelPers.Row - 1)
    ' Relacion
    Call oPersona.ActualizaPersRelaRelacion(FERelPers.TextMatrix(FERelPers.Row, 3), FERelPers.Row - 1)
    'Beneficiario
    Call oPersona.ActualizarRelaPersBenef(FERelPers.TextMatrix(FERelPers.Row, 4), FERelPers.Row - 1)
    'Beneficiario Porcentaje
    Call oPersona.ActualizarRelaPersBenefPorc(CDbl(FERelPers.TextMatrix(FERelPers.Row, 5)), FERelPers.Row - 1)
    'AMP
    Call oPersona.ActualizarRelaPersAMP(FERelPers.TextMatrix(FERelPers.Row, 6), FERelPers.Row - 1)
    
    'Habilitar Controles
    cmdPersRelaEjecutado = 0
    FERelPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FERelPers.lbEditarFlex = False
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FERelPers.SetFocus
End Sub

Private Sub cmdPersRelacCancelar_Click()
    CargaRelacionesPersonas
    'Habilitar Controles
    FERelPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FERelPers.lbEditarFlex = False
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = True
    cmdPersRelacDel.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FERelPers.SetFocus

End Sub

Private Sub cmdPersRelacDel_Click()
    If MsgBox("Esta Seguro que Desea Eliminar La Relacion con esta Persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRelacPersTipoAct(PersFilaEliminda, FERelPers.Row - 1)
        Call CmdPersAceptar_Click
        Call CargaRelacionesPersonas
    End If
End Sub

Private Sub cmdPersRelacEditar_Click()
    If oPersona.NumeroRelacPers > 0 Then
        cmdPersRelaEjecutado = 2
        FERelPersNoMoverdeFila = FERelPers.Row
        NomMoverSSTabs = SSTabs.Tab
        FERelPers.lbEditarFlex = True
        FERelPers.SetFocus
        cmdPersRelacNew.Enabled = False
        cmdPersRelacEditar.Enabled = False
        cmdPersRelacDel.Enabled = False
        cmdPersRelacAceptar.Visible = True
        cmdPersRelacCancelar.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        FERelPers.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdPersRelacNew_Click()
    cmdPersRelacAceptar.Visible = True
    cmdPersRelacCancelar.Visible = True
    cmdPersRelacNew.Enabled = False
    cmdPersRelacDel.Enabled = False
    cmdPersRelacEditar.Enabled = False
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    cmdPersRelaEjecutado = 1
    FERelPersNoMoverdeFila = FERelPers.Rows - 1
    FERelPers.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set oPersona = Nothing
    Unload Me
End Sub


Private Sub FEDocs_RowColChange()
    If cmdPersDocEjecutado = 1 Or cmdPersDocEjecutado = 2 Then
        FEDocs.Row = FEDocsPersNoMoverdeFila
    End If
End Sub

Private Sub FERelPers_RowColChange()
Dim oConstante As DConstante
    If FERelPers.lbEditarFlex Then
        If FERelPersNoMoverdeFila <> -1 Then
            FERelPers.Row = FERelPersNoMoverdeFila
        End If
        Set oConstante = New DConstante
        Select Case FERelPers.Col
            Case 3 'Relacion de Persona
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
            Case 4 'Beneficiario
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacBenef)
            Case 6 'AMP
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacAmp)
        End Select
        Set oConstante = Nothing
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CargaControles
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    TxtBCodPers.Enabled = True
    Screen.MousePointer = 0
End Sub


Private Sub SSTabs_Click(PreviousTab As Integer)
    If NomMoverSSTabs > -1 Then
        SSTabs.Tab = NomMoverSSTabs
    End If
End Sub

Private Sub TxtApellidoCasada_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ApellidoCasada = Trim(TxtApellidoCasada.Text)
    End If
End Sub

Private Sub TxtApellidoCasada_GotFocus()
    fEnfoque TxtApellidoCasada
End Sub

Private Sub TxtApellidoCasada_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtPersNombreN.SetFocus
    End If
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New DPersona
    Call oPersona.RecuperaPersona(Trim(TxtBCodPers.Text))
    If oPersona.PersCodigo = "" Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    End If
    Call CargaDatos
End Sub

Private Sub LimpiarPantalla()
Dim i As Integer
    bEstadoCargando = True
    'TxtBCodPers.Text = ""
    cmbPersPersoneria.ListIndex = -1
    txtPersNombreAP.Text = ""
    txtPersNombreAM.Text = ""
    txtPersNombreN.Text = ""
    cmbPersNatSexo.ListIndex = -1
    cmbPersNatEstCiv.ListIndex = -1
    txtPersNatHijos.Text = ""
    txtPersNacCreac.Text = "__/__/____"
    txtPersTelefono.Text = ""
    CboPersCiiu.ListIndex = -1
    cmbPersEstado.ListIndex = -1
    TxtTalla.Text = ""
    TxtPeso.Text = ""
    TxtTipoSangre.Text = ""
    txtPersTelefono2.Text = ""
    TxtEmail.Text = ""
    Call LimpiaFlex(FEDocs)
    
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), "PER")
    cmbPersUbiGeo(1).ListIndex = -1
    cmbPersUbiGeo(2).ListIndex = -1
    cmbPersUbiGeo(3).ListIndex = -1
    cmbPersUbiGeo(4).ListIndex = -1
    txtPersDireccDomicilio.Text = ""
    cmbPersDireccCondicion.ListIndex = -1
    txtPersNombreRS.Text = ""
    TxtSiglas.Text = ""
    cmbPersJurTpo.ListIndex = -1
    cmbPersJurMagnitud.ListIndex = -1
    txtPersJurEmpleados.Text = ""
    
    Call LimpiaFlex(FERelPers)
    Call LimpiaFlex(FEFteIng)
    IDBFirma.RutaImagen = ""
    bEstadoCargando = False
End Sub

Private Sub TxtBCodPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbPersPersoneria.Enabled Then
            cmbPersPersoneria.SetFocus
        Else
            SSTabs.Tab = 0
            If txtPersNombreAP.Enabled Then
                txtPersNombreAP.SetFocus
            End If
        End If
    End If
End Sub


Private Sub txtCodSbs_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CodigoSbs = Trim(txtCodSbs.Text)
    End If
End Sub

Private Sub txtCodSbs_GotFocus()
    fEnfoque txtCodSbs
End Sub

Private Sub txtCodSbs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboPersCiiu.SetFocus
    End If
End Sub

Private Sub TxtEmail_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Email = Trim(TxtEmail.Text)
    End If
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCodSbs.SetFocus
    End If
End Sub

Private Sub txtPersDireccDomicilio_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Domicilio = Trim(txtPersDireccDomicilio.Text)
    End If
End Sub

Private Sub txtPersDireccDomicilio_GotFocus()
    fEnfoque txtPersDireccDomicilio
End Sub

Private Sub txtPersDireccDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersDireccCondicion.SetFocus
    End If
End Sub

Private Sub txtPersJurEmpleados_Change()

    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NumerosEmpleados = CInt(IIf(Trim(txtPersJurEmpleados.Text) = "", "0", Trim(txtPersJurEmpleados.Text)))
    End If
End Sub

Private Sub txtPersJurEmpleados_GotFocus()
    fEnfoque txtPersJurEmpleados
End Sub

Private Sub txtPersJurEmpleados_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersJurMagnitud.SetFocus
    End If
End Sub

Private Sub txtPersJurEmpleados_LostFocus()
    txtPersJurEmpleados.Text = IIf(Trim(txtPersJurEmpleados.Text) = "", "0", Trim(txtPersJurEmpleados.Text))
End Sub

Private Sub txtPersNacCreac_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
        If Len(Trim(ValidaFecha(txtPersNacCreac.Text))) = 0 Then
            oPersona.FechaNacimiento = CDate(txtPersNacCreac.Text)
        End If
    End If
    
End Sub

Private Sub txtPersNacCreac_GotFocus()
    fEnfoque txtPersNacCreac
End Sub

Private Sub txtPersNacCreac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbRelPers.SetFocus
    End If
End Sub

Private Sub txtPersNatHijos_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Hijos = CInt(Trim(IIf(Trim(txtPersNatHijos.Text) = "", "0", txtPersNatHijos.Text)))
      End If
End Sub

Private Sub txtPersNatHijos_GotFocus()
    fEnfoque txtPersNatHijos
End Sub

Private Sub txtPersNatHijos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtPeso.SetFocus
    End If
End Sub

Private Sub txtPersNatHijos_LostFocus()
    txtPersNatHijos.Text = IIf(Trim(txtPersNatHijos.Text) = "", "0", txtPersNatHijos.Text)
End Sub

Private Sub txtPersNombreAM_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.ApellidoMaterno = Trim(txtPersNombreAM.Text)
   End If
End Sub

Private Sub txtPersNombreAM_GotFocus()
    fEnfoque txtPersNombreAM
End Sub

Private Sub txtPersNombreAM_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        If TxtApellidoCasada.Visible Then
            TxtApellidoCasada.SetFocus
        Else
            txtPersNombreN.SetFocus
        End If
    End If
End Sub

Private Sub txtPersNombreAP_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ApellidoPaterno = Trim(txtPersNombreAP.Text)
    End If
End Sub

Private Sub txtPersNombreAP_GotFocus()
    fEnfoque txtPersNombreAP
End Sub

Private Sub txtPersNombreAP_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtPersNombreAM.SetFocus
    End If
End Sub

Private Sub txtPersNombreN_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
      oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.Nombres = Trim(txtPersNombreN.Text)
  End If
End Sub

Private Sub txtPersNombreN_GotFocus()
    fEnfoque txtPersNombreN
End Sub

Private Sub txtPersNombreN_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersNatSexo.SetFocus
    End If
End Sub

Private Sub txtPersNombreRS_Change()

    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NombreCompleto = Trim(txtPersNombreRS.Text)
    End If
End Sub

Private Sub txtPersNombreRS_GotFocus()
    fEnfoque txtPersNombreRS
End Sub

Private Sub txtPersNombreRS_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtSiglas.SetFocus
    End If
End Sub

Private Sub txtPersTelefono_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Telefonos = Trim(txtPersTelefono.Text)
    End If
End Sub

Private Sub txtPersTelefono_GotFocus()
    fEnfoque txtPersTelefono
End Sub

Private Sub txtPersTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtPersTelefono2.SetFocus
    End If
End Sub

Private Sub txtPersTelefono2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Telefonos2 = Trim(txtPersTelefono2.Text)
    End If
End Sub

Private Sub txtPersTelefono2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtEmail.SetFocus
    End If
End Sub

Private Sub TxtPeso_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Peso = CDbl(Format(IIf(Trim(TxtPeso.Text) = "", "0", TxtPeso.Text), "#0.00"))
      End If
End Sub

Private Sub TxtPeso_GotFocus()
    fEnfoque TxtPeso
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPeso, KeyAscii)
    If KeyAscii = 13 Then
        TxtTalla.SetFocus
    End If
End Sub

Private Sub TxtPeso_LostFocus()
    TxtPeso.Text = IIf(Trim(TxtPeso.Text) = "", "0", TxtPeso.Text)
End Sub

Private Sub TxtSiglas_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Siglas = Trim(TxtSiglas.Text)
    End If
End Sub

Private Sub TxtSiglas_GotFocus()
    fEnfoque TxtSiglas
End Sub

Private Sub TxtSiglas_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersJurTpo.SetFocus
    End If
End Sub

Private Sub TxtTalla_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Talla = CDbl(IIf(Trim(TxtTalla.Text) = "", "0.00", Trim(TxtTalla.Text)))
    End If
End Sub

Private Sub TxtTalla_GotFocus()
    fEnfoque TxtTalla
End Sub

Private Sub TxtTalla_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtTalla, KeyAscii)
    If KeyAscii = 13 Then
        TxtTipoSangre.SetFocus
    End If
End Sub

Private Sub TxtTalla_LostFocus()
    TxtTalla.Text = IIf(Trim(TxtTalla.Text) = "", "0.00", Trim(TxtTalla.Text))
End Sub

Private Sub TxtTipoSangre_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoSangre = Trim(TxtTipoSangre.Text)
    End If
End Sub

Private Sub TxtTipoSangre_GotFocus()
    fEnfoque TxtTipoSangre
End Sub

Private Sub TxtTipoSangre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNacCreac.SetFocus
    End If
End Sub
