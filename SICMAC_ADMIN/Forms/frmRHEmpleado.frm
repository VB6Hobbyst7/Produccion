VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHEmpleado 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   Icon            =   "frmRHEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.ctrRRHHGen ctrRRHHGen 
      Height          =   1200
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
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
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   30
      Top             =   6825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9765
      TabIndex        =   0
      Top             =   6375
      Width           =   1095
   End
   Begin TabDlg.SSTab Tab 
      Height          =   5085
      Left            =   45
      TabIndex        =   2
      Top             =   1245
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   8969
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      WordWrap        =   0   'False
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dato!P.."
      TabPicture(0)   =   "frmRHEmpleado.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCancelar"
      Tab(0).Control(1)=   "fraGeneralidades"
      Tab(0).Control(2)=   "cmdGrabar"
      Tab(0).Control(3)=   "cmdEditar"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Familia"
      TabPicture(1)   =   "frmRHEmpleado.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraFamiliares"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "F&oto"
      TabPicture(2)   =   "frmRHEmpleado.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraFoto"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Historia"
      TabPicture(3)   =   "frmRHEmpleado.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraContrato(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Contrato"
      TabPicture(4)   =   "frmRHEmpleado.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraContrato(1)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Empleado"
      TabPicture(5)   =   "frmRHEmpleado.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraManArea_"
      Tab(5).ControlCount=   1
      Begin VB.Frame fraContrato 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Contrato"
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
         Height          =   4575
         Index           =   1
         Left            =   -74850
         TabIndex        =   22
         Top             =   360
         Width           =   10605
         Begin VB.CommandButton cmdGrabarContrato 
            Caption         =   "&Grabar"
            Height          =   375
            Left            =   5415
            TabIndex        =   24
            Top             =   4095
            Width           =   1095
         End
         Begin VB.CommandButton cmdAsigCont 
            Caption         =   "&Asigna Cont."
            Height          =   375
            Left            =   4170
            TabIndex        =   23
            Top             =   4080
            Width           =   1095
         End
         Begin RichTextLib.RichTextBox richContrato 
            Height          =   3780
            Left            =   150
            TabIndex        =   25
            Top             =   240
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   6668
            _Version        =   393217
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmRHEmpleado.frx":03B2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   -73680
         TabIndex        =   21
         Top             =   4635
         Width           =   1095
      End
      Begin VB.Frame fraGeneralidades 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Datos Personales"
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
         Height          =   4215
         Left            =   -74850
         TabIndex        =   10
         Top             =   360
         Width           =   10590
         Begin VB.CheckBox chkAgregarAPlanillas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Agrega a Planilla"
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   8835
            TabIndex        =   92
            Top             =   255
            Width           =   1500
         End
         Begin Sicmact.TxtBuscar TxtBuscarUsuario 
            Height          =   330
            Left            =   1230
            TabIndex        =   11
            Top             =   240
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
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
         End
         Begin Sicmact.FlexEdit FlexDoc 
            Height          =   1935
            Left            =   120
            TabIndex        =   32
            Top             =   2175
            Width           =   10230
            _ExtentX        =   18045
            _ExtentY        =   3413
            Cols0           =   3
            HighLight       =   1
            EncabezadosNombres=   "#-Documento-Numero"
            EncabezadosAnchos=   "300-3000-4000"
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
            ColumnasAEditar =   "X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblEdadD 
            Caption         =   "Edad :"
            Height          =   180
            Left            =   8010
            TabIndex        =   94
            Top             =   705
            Width           =   540
         End
         Begin VB.Label lblEdad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NO ESPECIFICADO"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8550
            TabIndex        =   93
            Top             =   630
            Width           =   1815
         End
         Begin VB.Label lblRHEstado 
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3000
            TabIndex        =   87
            Top             =   270
            Width           =   5610
         End
         Begin VB.Label lblEstadoRes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1230
            TabIndex        =   13
            Top             =   630
            Width           =   6555
         End
         Begin VB.Label lblUsuario 
            Caption         =   "Usuario"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblEstado 
            Caption         =   "Estado :"
            Height          =   255
            Left            =   165
            TabIndex        =   19
            Top             =   660
            Width           =   840
         End
         Begin VB.Label lblAgenciaAsignada 
            Caption         =   "Asignado a :"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1035
            Width           =   990
         End
         Begin VB.Label lblAgenciaActual 
            Caption         =   "Ubic. Actual :"
            Height          =   255
            Left            =   105
            TabIndex        =   17
            Top             =   1395
            Width           =   1020
         End
         Begin VB.Label lblCargo 
            Caption         =   "Cargo :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1785
            Width           =   1005
         End
         Begin VB.Label lblAgenciaRes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1230
            TabIndex        =   15
            Top             =   1350
            Width           =   9165
         End
         Begin VB.Label lblCargoRes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1230
            TabIndex        =   14
            Top             =   1725
            Width           =   9165
         End
         Begin VB.Label lblAgenciaAsignadaRes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1230
            TabIndex        =   12
            Top             =   990
            Width           =   9165
         End
      End
      Begin VB.Frame fraFoto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Foto"
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
         Height          =   4530
         Left            =   -74895
         TabIndex        =   7
         Top             =   375
         Width           =   10650
         Begin VB.CommandButton Command1 
            Caption         =   "Exportar"
            Height          =   375
            Left            =   6600
            TabIndex        =   96
            Top             =   4020
            Width           =   1215
         End
         Begin VB.CommandButton cmdFoto 
            Caption         =   "&Asigna Foto"
            Height          =   375
            Left            =   3990
            TabIndex        =   9
            Top             =   4020
            Width           =   1095
         End
         Begin VB.CommandButton cmdGrabarFoto 
            Caption         =   "&Grabar"
            Height          =   375
            Left            =   5220
            TabIndex        =   8
            Top             =   4020
            Width           =   1095
         End
         Begin VB.Image picFoto 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3675
            Left            =   3465
            Stretch         =   -1  'True
            Top             =   270
            Width           =   3315
         End
      End
      Begin VB.Frame fraFamiliares 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Familiares"
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
         Height          =   4590
         Left            =   -74865
         TabIndex        =   6
         Top             =   375
         Width           =   10635
         Begin Sicmact.FlexEdit FlexFamiliares 
            Height          =   4260
            Left            =   150
            TabIndex        =   33
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7514
            Cols0           =   3
            HighLight       =   1
            EncabezadosNombres=   "#-Nombre-Parentesco"
            EncabezadosAnchos=   "400-6000-3500"
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
            ColumnasAEditar =   "X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraContrato 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "File del RRHH"
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
         Height          =   4590
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   10500
         Begin Sicmact.FlexEdit FlexContrato 
            Height          =   1335
            Left            =   105
            TabIndex        =   31
            ToolTipText     =   "Haga doble Click Sobre el contrato que dese visualizar"
            Top             =   240
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   2355
            Cols0           =   13
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-NumContrato-Tipo-Fecha Fin-Contrato Ini-Contrato Fin-Area Asig-Age Asig-Area Actual-Age Actual-Cargo-Sueldo-Comentario"
            EncabezadosAnchos=   "300-1200-2000-1200-1200-1200-2500-2500-2500-2500-2500-1500-4000"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-R-R-L-L-L-L-L-R-L"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin TabDlg.SSTab TabDetalle 
            Height          =   2865
            Left            =   105
            TabIndex        =   34
            Top             =   1650
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   5054
            _Version        =   393216
            Tabs            =   16
            TabsPerRow      =   8
            TabHeight       =   520
            WordWrap        =   0   'False
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Area!Cargos"
            TabPicture(0)   =   "frmRHEmpleado.frx":0433
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "cmdAgregarArea(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdCancelarArea(0)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "cmdGrabarArea(0)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "fraManArea(0)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Asist.Med"
            TabPicture(1)   =   "frmRHEmpleado.frx":044F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraManArea(5)"
            Tab(1).Control(1)=   "cmdGrabarArea(1)"
            Tab(1).Control(2)=   "cmdCancelarArea(1)"
            Tab(1).Control(3)=   "cmdAgregarArea(1)"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Estados"
            TabPicture(2)   =   "frmRHEmpleado.frx":046B
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FlexEstado"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Sueldos"
            TabPicture(3)   =   "frmRHEmpleado.frx":0487
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "fraManArea(1)"
            Tab(3).Control(1)=   "cmdGrabarArea(2)"
            Tab(3).Control(2)=   "cmdCancelarArea(2)"
            Tab(3).Control(3)=   "cmdAgregarArea(2)"
            Tab(3).ControlCount=   4
            TabCaption(4)   =   "Sist.Pension"
            TabPicture(4)   =   "frmRHEmpleado.frx":04A3
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "fraManArea(4)"
            Tab(4).Control(1)=   "cmdGrabarArea(3)"
            Tab(4).Control(2)=   "cmdCancelarArea(3)"
            Tab(4).Control(3)=   "cmdAgregarArea(3)"
            Tab(4).ControlCount=   4
            TabCaption(5)   =   "Comentario"
            TabPicture(5)   =   "frmRHEmpleado.frx":04BF
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "cmdAgregarArea(4)"
            Tab(5).Control(1)=   "cmdCancelarArea(4)"
            Tab(5).Control(2)=   "cmdGrabarArea(4)"
            Tab(5).Control(3)=   "fraManArea(6)"
            Tab(5).ControlCount=   4
            TabCaption(6)   =   "Infor-Social"
            TabPicture(6)   =   "frmRHEmpleado.frx":04DB
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "flexInfSoc"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "Vacaciones"
            TabPicture(7)   =   "frmRHEmpleado.frx":04F7
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "FlexPerNoLab(0)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Lic!Permisos"
            TabPicture(8)   =   "frmRHEmpleado.frx":0513
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "FlexPerNoLab(1)"
            Tab(8).ControlCount=   1
            TabCaption(9)   =   "Sanciones"
            TabPicture(9)   =   "frmRHEmpleado.frx":052F
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "FlexPerNoLab(3)"
            Tab(9).ControlCount=   1
            TabCaption(10)  =   "Horario"
            TabPicture(10)  =   "frmRHEmpleado.frx":054B
            Tab(10).ControlEnabled=   0   'False
            Tab(10).Control(0)=   "FlexDia"
            Tab(10).Control(1)=   "FlexHor"
            Tab(10).ControlCount=   2
            TabCaption(11)  =   "Con!Adendas"
            TabPicture(11)  =   "frmRHEmpleado.frx":0567
            Tab(11).ControlEnabled=   0   'False
            Tab(11).Control(0)=   "fraManArea(7)"
            Tab(11).Control(1)=   "cmdGrabarArea(5)"
            Tab(11).Control(2)=   "cmdCancelarArea(5)"
            Tab(11).Control(3)=   "cmdAgregarArea(5)"
            Tab(11).Control(4)=   "richCont"
            Tab(11).ControlCount=   5
            TabCaption(12)  =   "Descansos"
            TabPicture(12)  =   "frmRHEmpleado.frx":0583
            Tab(12).ControlEnabled=   0   'False
            Tab(12).Control(0)=   "FlexPerNoLab(2)"
            Tab(12).ControlCount=   1
            TabCaption(13)  =   "Merit/Demer."
            TabPicture(13)  =   "frmRHEmpleado.frx":059F
            Tab(13).ControlEnabled=   0   'False
            Tab(13).Control(0)=   "FlexPer"
            Tab(13).ControlCount=   1
            TabCaption(14)  =   "Curriculum"
            TabPicture(14)  =   "frmRHEmpleado.frx":05BB
            Tab(14).ControlEnabled=   0   'False
            Tab(14).Control(0)=   "Flex"
            Tab(14).ControlCount=   1
            TabCaption(15)  =   "Act.Ext.Cur"
            TabPicture(15)  =   "frmRHEmpleado.frx":05D7
            Tab(15).ControlEnabled=   0   'False
            Tab(15).Control(0)=   "FlexExtra"
            Tab(15).ControlCount=   1
            Begin VB.Frame fraManArea 
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
               Height          =   2145
               Index           =   0
               Left            =   75
               TabIndex        =   67
               Top             =   630
               Width           =   8745
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1920
                  Index           =   0
                  Left            =   75
                  TabIndex        =   68
                  Top             =   195
                  Width           =   8580
                  _ExtentX        =   15134
                  _ExtentY        =   3387
                  Cols0           =   13
                  HighLight       =   1
                  AllowUserResizing=   3
                  EncabezadosNombres=   $"frmRHEmpleado.frx":05F3
                  EncabezadosAnchos=   "300-800-2500-1800-800-2500-800-2500-1800-800-2500-1800-1800"
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
                  ColumnasAEditar =   "X-1-X-X-4-X-6-X-X-9-X-11-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0-0-1-0-1-0-0-1-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-L-L-L-L-L"
                  FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   300
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2145
               Index           =   5
               Left            =   -74925
               TabIndex        =   65
               Top             =   645
               Width           =   8745
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1920
                  Index           =   1
                  Left            =   90
                  TabIndex        =   66
                  Top             =   165
                  Width           =   8580
                  _ExtentX        =   15134
                  _ExtentY        =   3387
                  Cols0           =   5
                  HighLight       =   1
                  AllowUserResizing=   3
                  EncabezadosNombres=   "#-Cod AMP-Asistencia Med Priv.-Comentario-Movimiento"
                  EncabezadosAnchos=   "300-800-2500-3000-1600"
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
                  ColumnasAEditar =   "X-1-X-3-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-L-L"
                  FormatosEdit    =   "0-0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   300
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2130
               Index           =   2
               Left            =   -74925
               TabIndex        =   63
               Top             =   675
               Width           =   8715
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   50
                  Index           =   22
                  Left            =   0
                  TabIndex        =   64
                  Top             =   0
                  Width           =   50
                  _ExtentX        =   79
                  _ExtentY        =   79
                  Cols0           =   7
                  HighLight       =   1
                  AllowUserResizing=   3
                  EncabezadosNombres=   "#-Cod Fondo-Fondo-Cod AFP-AFP-Comentario-Movimiento"
                  EncabezadosAnchos=   "400-1000-1200-1200-1200-5000-1500"
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
                  ColumnasAEditar =   "X-1-X-3-X-5-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0-1-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-L-L-L-C"
                  FormatosEdit    =   "0-0-0-0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   405
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2160
               Index           =   3
               Left            =   -74910
               TabIndex        =   61
               Top             =   645
               Width           =   8715
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1815
                  Index           =   23
                  Left            =   75
                  TabIndex        =   62
                  Top             =   225
                  Width           =   8520
                  _ExtentX        =   15028
                  _ExtentY        =   3201
                  Cols0           =   5
                  HighLight       =   1
                  AllowUserResizing=   3
                  EncabezadosNombres=   "#-Cod AMP-Asist Med Priv-Comentario-Movimiento"
                  EncabezadosAnchos=   "400-800-2500-3500-1200"
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
                  ColumnasAEditar =   "X-1-X-3-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-L-C"
                  FormatosEdit    =   "0-0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   405
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2160
               Index           =   4
               Left            =   -74910
               TabIndex        =   59
               Top             =   645
               Width           =   8715
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1920
                  Index           =   3
                  Left            =   105
                  TabIndex        =   60
                  Top             =   150
                  Width           =   8520
                  _ExtentX        =   15028
                  _ExtentY        =   3387
                  Cols0           =   7
                  HighLight       =   1
                  AllowUserResizing=   1
                  EncabezadosNombres=   "#-Sis.Pen-Sist. Pensiones-Cod AFP-AFP-Comentario-Movimiento"
                  EncabezadosAnchos=   "300-800-1200-1200-1200-3000-1800"
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
                  ColumnasAEditar =   "X-1-X-3-X-5-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0-1-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-L-L-L-L-L-L"
                  FormatosEdit    =   "0-0-0-0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   300
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2160
               Index           =   1
               Left            =   -74880
               TabIndex        =   57
               Top             =   630
               Width           =   8595
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1890
                  Index           =   2
                  Left            =   60
                  TabIndex        =   58
                  Top             =   195
                  Width           =   8460
                  _ExtentX        =   14923
                  _ExtentY        =   3334
                  Cols0           =   4
                  HighLight       =   1
                  EncabezadosNombres=   "#-Sueldo-Comentario-Movimiento"
                  EncabezadosAnchos=   "300-1000-4000-2500"
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
                  ColumnasAEditar =   "X-1-2-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-0-0-0"
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  BackColorControl=   -2147483643
                  EncabezadosAlineacion=   "C-R-L-L"
                  FormatosEdit    =   "0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   300
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2160
               Index           =   6
               Left            =   -74910
               TabIndex        =   55
               Top             =   660
               Width           =   8760
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1815
                  Index           =   4
                  Left            =   60
                  TabIndex        =   56
                  Top             =   225
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   3201
                  Cols0           =   3
                  HighLight       =   1
                  AllowUserResizing=   3
                  EncabezadosNombres=   "#-Comentario-Actualizacion"
                  EncabezadosAnchos=   "400-5500-2200"
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
                  ColumnasAEditar =   "X-1-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-0-0"
                  BackColor       =   -2147483639
                  BackColorControl=   -2147483639
                  BackColorControl=   -2147483639
                  BackColorControl=   -2147483639
                  EncabezadosAlineacion=   "C-L-L"
                  FormatosEdit    =   "0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   405
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
                  CellBackColor   =   -2147483639
               End
            End
            Begin VB.Frame fraManArea 
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
               Height          =   2100
               Index           =   7
               Left            =   -74940
               TabIndex        =   53
               Top             =   660
               Width           =   8760
               Begin Sicmact.FlexEdit FlexVista 
                  Height          =   1815
                  Index           =   5
                  Left            =   75
                  TabIndex        =   54
                  Top             =   195
                  Width           =   8595
                  _ExtentX        =   15161
                  _ExtentY        =   3201
                  Cols0           =   7
                  HighLight       =   1
                  AllowUserResizing=   3
                  EncabezadosNombres=   "#-Tpo-Nom.Tpo-Inicio-Fin-Texto-Movimiento"
                  EncabezadosAnchos=   "300-500-2500-1000-1000-800-1800"
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
                  ColumnasAEditar =   "X-1-X-3-4-5-X"
                  TextStyleFixed  =   3
                  ListaControles  =   "0-1-0-2-2-1-0"
                  BackColor       =   -2147483639
                  BackColorControl=   -2147483639
                  BackColorControl=   -2147483639
                  BackColorControl=   -2147483639
                  EncabezadosAlineacion=   "C-R-R-R-R-C-L"
                  FormatosEdit    =   "0-0-0-0-0"
                  TextArray0      =   "#"
                  lbEditarFlex    =   -1  'True
                  lbUltimaInstancia=   -1  'True
                  Appearance      =   0
                  ColWidth0       =   300
                  RowHeight0      =   300
                  ForeColorFixed  =   -2147483630
                  CellBackColor   =   -2147483639
               End
            End
            Begin VB.CommandButton cmdGrabarArea 
               Caption         =   "&Grabar"
               Height          =   375
               Index           =   0
               Left            =   8910
               TabIndex        =   52
               Top             =   1335
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancelarArea 
               Caption         =   "&Cancelar"
               Height          =   375
               Index           =   0
               Left            =   8910
               TabIndex        =   51
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdAgregarArea 
               Caption         =   "&Agregar"
               Height          =   375
               Index           =   0
               Left            =   8910
               TabIndex        =   50
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdGrabarArea 
               Caption         =   "&Grabar"
               Height          =   375
               Index           =   1
               Left            =   -66090
               TabIndex        =   49
               Top             =   1335
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancelarArea 
               Caption         =   "&Cancelar"
               Height          =   375
               Index           =   1
               Left            =   -66090
               TabIndex        =   48
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdAgregarArea 
               Caption         =   "&Agregar"
               Height          =   375
               Index           =   1
               Left            =   -66090
               TabIndex        =   47
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdGrabarArea 
               Caption         =   "&Grabar"
               Height          =   375
               Index           =   2
               Left            =   -66090
               TabIndex        =   46
               Top             =   1335
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancelarArea 
               Caption         =   "&Cancelar"
               Height          =   375
               Index           =   2
               Left            =   -66090
               TabIndex        =   45
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdAgregarArea 
               Caption         =   "&Agregar"
               Height          =   375
               Index           =   2
               Left            =   -66090
               TabIndex        =   44
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdGrabarArea 
               Caption         =   "&Grabar"
               Height          =   375
               Index           =   3
               Left            =   -66090
               TabIndex        =   43
               Top             =   1335
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancelarArea 
               Caption         =   "&Cancelar"
               Height          =   375
               Index           =   3
               Left            =   -66090
               TabIndex        =   42
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdAgregarArea 
               Caption         =   "&Agregar"
               Height          =   375
               Index           =   3
               Left            =   -66090
               TabIndex        =   41
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdGrabarArea 
               Caption         =   "&Grabar"
               Height          =   375
               Index           =   4
               Left            =   -66075
               TabIndex        =   40
               Top             =   1335
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancelarArea 
               Caption         =   "&Cancelar"
               Height          =   375
               Index           =   4
               Left            =   -66090
               TabIndex        =   39
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdAgregarArea 
               Caption         =   "&Agregar"
               Height          =   375
               Index           =   4
               Left            =   -66090
               TabIndex        =   38
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdGrabarArea 
               Caption         =   "&Grabar"
               Height          =   375
               Index           =   5
               Left            =   -66090
               TabIndex        =   37
               Top             =   1335
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancelarArea 
               Caption         =   "&Cancelar"
               Height          =   375
               Index           =   5
               Left            =   -66090
               TabIndex        =   36
               Top             =   870
               Width           =   1095
            End
            Begin VB.CommandButton cmdAgregarArea 
               Caption         =   "&Agregar"
               Height          =   375
               Index           =   5
               Left            =   -66090
               TabIndex        =   35
               Top             =   870
               Width           =   1095
            End
            Begin Sicmact.FlexEdit flexInfSoc 
               Height          =   2100
               Left            =   -74910
               TabIndex        =   69
               Top             =   705
               Width           =   9870
               _ExtentX        =   17410
               _ExtentY        =   3704
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Fecha-Comentario-bit"
               EncabezadosAnchos=   "300-2400-5500-0"
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
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-R-L-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin Sicmact.FlexEdit FlexEstado 
               Height          =   2070
               Left            =   -74910
               TabIndex        =   70
               Top             =   690
               Width           =   9900
               _ExtentX        =   17463
               _ExtentY        =   3651
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   3
               EncabezadosNombres=   "#-Estado-Comentario-Movimiento"
               EncabezadosAnchos=   "400-2500-4000-2500"
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
               ColumnasAEditar =   "X-X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-R-L-L"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               Appearance      =   0
               ColWidth0       =   405
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin Sicmact.FlexEdit FlexPerNoLab 
               Height          =   2085
               Index           =   0
               Left            =   -74895
               TabIndex        =   71
               Top             =   705
               Width           =   9900
               _ExtentX        =   17463
               _ExtentY        =   3678
               Cols0           =   13
               EncabezadosNombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
               EncabezadosAnchos=   "300-800-2500-2500-2500-2500-2500-5000-0-0-0-0-0"
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
               ColumnasAEditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
               EncabezadosAlineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
               FormatosEdit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
               TextArray0      =   "#"
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
            End
            Begin Sicmact.FlexEdit FlexPerNoLab 
               Height          =   2085
               Index           =   1
               Left            =   -74895
               TabIndex        =   72
               Top             =   705
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   3678
               Cols0           =   13
               EncabezadosNombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
               EncabezadosAnchos=   "300-800-2500-2500-2500-2500-2500-5000-800-2500-5000-0-0"
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
               ColumnasAEditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
               EncabezadosAlineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
               FormatosEdit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
            End
            Begin Sicmact.FlexEdit FlexPerNoLab 
               Height          =   2085
               Index           =   2
               Left            =   -74895
               TabIndex        =   73
               Top             =   705
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   3678
               Cols0           =   13
               EncabezadosNombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
               EncabezadosAnchos=   "300-800-2500-2500-2500-0-0-5000-0-0-0-0-0"
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
               ColumnasAEditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
               EncabezadosAlineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
               FormatosEdit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
               TextArray0      =   "#"
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
            End
            Begin Sicmact.FlexEdit FlexPerNoLab 
               Height          =   2085
               Index           =   3
               Left            =   -74895
               TabIndex        =   74
               Top             =   705
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   3678
               Cols0           =   13
               EncabezadosNombres=   "#-CodTipo-Tipo-Sol Ini-Sol Fin-Ejec Ini-Ejec Fin-Comentario-Cod Estado-Estado-Observacion-bit1-bit2"
               EncabezadosAnchos=   "300-800-2500-2500-2500-0-0-5000-0-0-0-0-0"
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
               ColumnasAEditar =   "X-1-X-3-4-5-6-7-8-X-10-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-2-2-2-2-0-1-0-0-0-0"
               EncabezadosAlineacion=   "C-L-L-R-R-R-R-L-L-L-L-C-C"
               FormatosEdit    =   "0-0-0-5-5-5-5-0-0-0-0-0-0"
               TextArray0      =   "#"
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
            End
            Begin Sicmact.FlexEdit FlexDia 
               Height          =   2085
               Left            =   -70155
               TabIndex        =   75
               Top             =   705
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   3678
               Cols0           =   7
               HighLight       =   1
               AllowUserResizing=   3
               EncabezadosNombres=   "#-Cod Dia-Dia-Turno-Hora Ini-Hora Fin-Existe"
               EncabezadosAnchos=   "300-800-800-1200-800-800-0"
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
               ColumnasAEditar =   "X-1-X-3-4-5-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-3-2-2-0"
               EncabezadosAlineacion=   "C-L-L-L-R-R-C"
               FormatosEdit    =   "0-0-0-6-6-6-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin Sicmact.FlexEdit FlexHor 
               Height          =   2085
               Left            =   -74895
               TabIndex        =   76
               Top             =   705
               Width           =   4740
               _ExtentX        =   8361
               _ExtentY        =   3678
               Cols0           =   4
               HighLight       =   1
               AllowUserResizing=   1
               EncabezadosNombres=   "#-Fecha-Comentario-bit"
               EncabezadosAnchos=   "300-1200-4000-0"
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
               ColumnasAEditar =   "X-1-2-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-2-0-0"
               EncabezadosAlineacion=   "C-R-L-C"
               FormatosEdit    =   "0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin RichTextLib.RichTextBox richCont 
               Height          =   420
               Left            =   -65970
               TabIndex        =   77
               Top             =   2070
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
               _Version        =   393217
               ReadOnly        =   -1  'True
               Appearance      =   0
               TextRTF         =   $"frmRHEmpleado.frx":0689
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Sicmact.FlexEdit FlexPer 
               Height          =   2040
               Left            =   -74895
               TabIndex        =   78
               Top             =   705
               Width           =   9900
               _ExtentX        =   17463
               _ExtentY        =   3598
               Cols0           =   8
               HighLight       =   1
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Cod.Tpo-Tipo-Fecha-Observaciones-bit-bit1-Movimiento"
               EncabezadosAnchos=   "300-800-1800-900-3000-0-0-1800"
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
               ColumnasAEditar =   "X-1-X-3-4-X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-2-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-L-C-C-L"
               FormatosEdit    =   "0-0-0-0-0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin Sicmact.FlexEdit Flex 
               Height          =   2025
               Left            =   -74895
               TabIndex        =   79
               Top             =   705
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   3572
               Cols0           =   15
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   $"frmRHEmpleado.frx":070A
               EncabezadosAnchos=   "300-1000-2500-3000-1200-2000-1200-1200-1200-2000-1200-5000-2500-0-0"
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
               ColumnasAEditar =   "X-1-X-3-4-X-6-7-8-X-10-11-X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-0-1-0-2-2-1-0-0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-L-L-R-L-R-C-C"
               FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-3-0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin Sicmact.FlexEdit FlexExtra 
               Height          =   2025
               Left            =   -74895
               TabIndex        =   95
               Top             =   705
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   3572
               Cols0           =   14
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Cod.Tpo-Tipo-Cod Activ-Actividad-Años.Pract-Costo-Cod.Niv-Nivel-Otorgado CMACT-Comentario-UltimaActualizacion-BitCod-BitItem"
               EncabezadosAnchos=   "300-750-3500-1000-3500-1000-1000-800-2500-1500-5000-2500-0-0"
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
               ColumnasAEditar =   "X-1-X-3-X-5-6-7-X-9-10-X-X-X"
               TextStyleFixed  =   3
               ListaControles  =   "0-1-0-1-0-0-0-1-0-4-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-L-L-L-L-R-R-L-L-L-L-L-C-C"
               FormatosEdit    =   "0-0-0-0-0-3-2-0-0-0-1-0-0-0"
               TextArray0      =   "#"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               lbBuscaDuplicadoText=   -1  'True
               Appearance      =   0
               ColWidth0       =   300
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   -74850
         TabIndex        =   4
         Top             =   4635
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   -73680
         TabIndex        =   3
         Top             =   4635
         Width           =   1095
      End
      Begin VB.Frame fraManArea_ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Empleados"
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
         Height          =   4575
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   10620
         Begin VB.TextBox txtUbicacion 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1680
            TabIndex        =   99
            Top             =   3120
            Width           =   3195
         End
         Begin MSMask.MaskEdBox mskFecIng 
            Height          =   270
            Left            =   1680
            TabIndex        =   89
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Sicmact.TxtBuscar txtAgencia 
            Height          =   345
            Left            =   5580
            TabIndex        =   83
            Top             =   645
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   609
            Appearance      =   0
            BackColor       =   12648447
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
            sTitulo         =   ""
         End
         Begin VB.CommandButton cmdEliminaCta 
            Caption         =   "&Elimina"
            Height          =   315
            Left            =   8340
            TabIndex        =   82
            Top             =   2475
            Width           =   1035
         End
         Begin VB.CommandButton cmdAgregaCta 
            Caption         =   "&Agregar"
            Height          =   315
            Left            =   9420
            TabIndex        =   81
            Top             =   2475
            Width           =   1035
         End
         Begin Sicmact.FlexEdit flexCuenta 
            Height          =   1395
            Left            =   5595
            TabIndex        =   80
            Top             =   1035
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   2461
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cuenta-tarj"
            EncabezadosAnchos=   "300-2000-1"
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
            ColumnasAEditar =   "X-1-X"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.TextBox txtNumSeg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   28
            Top             =   2520
            Width           =   3195
         End
         Begin VB.TextBox txtCUSPP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   27
            Top             =   1320
            Width           =   3195
         End
         Begin MSMask.MaskEdBox mskTar 
            Height          =   420
            Left            =   6600
            TabIndex        =   85
            Top             =   3480
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   741
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-####-####-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFecCese 
            Height          =   270
            Left            =   1680
            TabIndex        =   90
            Top             =   960
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFecAfilia 
            Height          =   270
            Left            =   1680
            TabIndex        =   98
            Top             =   1920
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000000&
            Caption         =   "Ubicación"
            Height          =   255
            Left            =   420
            TabIndex        =   100
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Afiliacion"
            Height          =   195
            Left            =   420
            TabIndex        =   97
            Top             =   1920
            Width           =   1125
         End
         Begin VB.Label lblFechaCese 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Cese"
            Height          =   255
            Left            =   405
            TabIndex        =   91
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label lblFechaIng 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ing."
            Height          =   255
            Left            =   405
            TabIndex        =   88
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label lblTar 
            Caption         =   "Tarjeta :"
            Height          =   255
            Left            =   5880
            TabIndex        =   86
            Top             =   3570
            Width           =   645
         End
         Begin VB.Label lblAgencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6420
            TabIndex        =   84
            Top             =   660
            Width           =   4020
         End
         Begin VB.Line Line1 
            X1              =   5445
            X2              =   5445
            Y1              =   120
            Y2              =   4530
         End
         Begin VB.Label lblNumSeg 
            BackStyle       =   0  'Transparent
            Caption         =   "Num Seg :"
            Height          =   195
            Left            =   420
            TabIndex        =   30
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label lblCUSPP 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSSP :"
            Height          =   240
            Left            =   420
            TabIndex        =   29
            Top             =   1440
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "frmRHEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lbModEmp As Boolean
Dim lbContrato As Boolean
Dim lnSueldo As Currency
Dim lsCarCod As String
Dim lsEstado As String
Dim lnTipo As TipoOpe
Dim lbEscapa As Boolean
Dim lsCodPers As String
Dim lsCodEva As String
Dim lnContratoMantTpo As RHContratoMantTpo
Dim lnRowAnt As Integer
Dim lbRegistro As Boolean
Dim lsRutaFoto As String

'->***** LUCV20181220, Anexo01 de Acta 199-2018
Dim objPista As COMManejador.Pista
Dim lsMovNro As String
'<-***** Fin LUCV20181220

Private Sub cmdAgregaCta_Click()
    Me.flexCuenta.AdicionaFila
End Sub

Private Sub cmdAgregarArea_Click(Index As Integer)
'Agregado by NAGL ERS074-2017 20171209
Dim ActRH As New DActualizaDatosRRHH
Dim rs As New ADODB.Recordset
Dim oArea As New DActualizaDatosArea
Dim psPersCod As String
'*****END NAGL ERS074-2017 20171209******
    If Me.ctrRRHHGen.psCodigoPersona = "" Then
        Me.ctrRRHHGen.SetFocus
        Exit Sub
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoCargo Then
        If (Me.FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 12) = "" And FlexVista(Index).Rows <> 2) Then
            MsgBox "Solo se debe ingresar un registro.", vbInformation, "Aviso"
            Exit Sub
        End If
        '***Agregado por ELRO el 20111219, según Acta N° 346-2011/TI-D
        If Me.ctrRRHHGen.psCodigoPersona = gsCodPersUser Then
            MsgBox "No puede agregar un cargo a su persona, solicite que lo haga otro usuario de RRHH", vbInformation, "Aviso"
            Exit Sub
        End If
        '***Fin Agregado por ELRO*************************************
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoSueldo Then
        If (Me.FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3) = "" And FlexVista(Index).Rows <> 2) Then
            MsgBox "Solo se debe ingresar un registro.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoSisPens Then
        If (Me.FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6) = "" And FlexVista(Index).Rows <> 2) Then
            MsgBox "Solo se debe ingresar un registro.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoComentario Then
        If (Me.FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 2) = "" And FlexVista(Index).Rows <> 2) Then
            MsgBox "Solo se debe ingresar un registro.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    Me.FlexVista(Index).AdicionaFila
     '************BEGIN NAGL ERS074-2017 20171212**************
    If Index = RHContratoMantTpo.RHContratoMantTpoCargo Then
        psPersCod = Me.ctrRRHHGen.psCodigoPersona
        Set rs = ActRH.ObtenerUltimoCargoOficial(psPersCod)
        If Not (rs.BOF And rs.EOF) Then
             FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6) = rs!CarOfi
             FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 7) = rs!CargoOficial
             FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 8) = rs!AreaOficial
             FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 9) = rs!AgeOfi
             FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 10) = rs!AgenciaOficial
        End If
    End If '*** NAGL ERS074-2017 20171209
    ActivaArea True, Index
    FlexVista_RowColChange Index
    FlexVista(Index).SetFocus
End Sub

Private Sub cmdAsigCont_Click()
    CDialog.CancelError = False
    On Error GoTo ErrHandler
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos Bmp(*.txt)|*.txt"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Me.richContrato.LoadFile CDialog.FileName, 1
    richContrato.Text = Replace(richContrato.Text, "'", "")
    Exit Sub
ErrHandler:
End Sub

Private Sub cmdCancelar_Click()
    ClearScreen
    Activa False, CInt(lnContratoMantTpo)
End Sub

Private Sub cmdCancelarArea_Click(Index As Integer)
    ActivaArea False, Index
    CargaData Me.ctrRRHHGen.psCodigoPersona
End Sub

Private Sub cmdEditar_Click()
    If Me.ctrRRHHGen.psCodigoPersona = "" Then
        Me.ctrRRHHGen.SetFocus
        Exit Sub
    End If

    Activa True, CInt(lnContratoMantTpo)
End Sub


Private Sub ActivaArea(pbvalor As Boolean, pnIndex As Integer)
    Dim i As Integer
    
    If pbvalor Then
        For i = 0 To cmdAgregarArea.Count - 1
            If pnIndex = i Or pnIndex = -1 Then
                Me.fraManArea(i).Enabled = True
                Me.cmdCancelarArea(i).Visible = pbvalor
                Me.cmdAgregarArea(i).Visible = Not pbvalor
                Me.cmdGrabarArea(i).Enabled = pbvalor
                Me.cmdGrabarArea(i).Enabled = pbvalor
                Me.FlexVista(i).lbEditarFlex = pbvalor
            Else
                Me.FlexVista(i).lbEditarFlex = pbvalor
                Me.cmdAgregarArea(i).Enabled = False
                Me.cmdCancelarArea(i).Visible = False
                
                Me.cmdGrabarArea(i).Visible = pbvalor
                Me.cmdGrabarArea(i).Enabled = False
            End If
        Next i
    Else
        For i = 0 To cmdAgregarArea.Count - 1
            If i = pnIndex Then
                Me.cmdCancelarArea(i).Visible = pbvalor
                Me.cmdAgregarArea(i).Visible = Not pbvalor
                Me.cmdGrabarArea(i).Enabled = pbvalor
                Me.cmdGrabarArea(i).Visible = True
                Me.FlexVista(i).lbEditarFlex = pbvalor
            Else
                Me.cmdCancelarArea(i).Visible = False
                Me.cmdGrabarArea(i).Visible = False
            End If
        Next i
    End If
    Me.cmdSalir.Enabled = Not pbvalor
    Me.ctrRRHHGen.Enabled = Not pbvalor
End Sub

Private Sub cmdEliminaCta_Click()
    If MsgBox("Se eliminara la cuenta " & Me.flexCuenta.TextMatrix(flexCuenta.row, 1), vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Me.flexCuenta.EliminaFila flexCuenta.row
    
End Sub

Private Sub cmdFoto_Click()
    CDialog.CancelError = False
    On Error GoTo ErrHandler
    CDialog.Flags = cdlOFNHideReadOnly
    CDialog.Filter = "Archivos Bmp(*.bmp)|*.bmp"
    CDialog.FilterIndex = 2
    CDialog.ShowOpen
    Set Me.picFoto = LoadPicture(CDialog.FileName)
    
    lsRutaFoto = CDialog.FileName
    
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdGrabar_Click()
    Dim oRh As NActualizaDatosRRHH
    Dim oCon As NActualizaDatosContrato
    Dim lnSueldoG As Currency
    Dim lsCargoG As String
    Dim lsEstadoG As String
    Set oCon = New NActualizaDatosContrato
    Set oRh = New NActualizaDatosRRHH
    
    
    If Not IsDate(Me.mskFecIng.Text) Then
        MsgBox "Debe ingresar una fecha de ingreso valida.", vbInformation, "Aviso"
        Me.Tab.Tab = 5
        Me.mskFecIng.SetFocus
        Exit Sub
    End If
    
    If Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "E" Or Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "F" Or Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "P" Then
        If Not oRh.ValidaCtasEmpleado(Me.flexCuenta.GetRsNew, gbBitCentral) Then
            MsgBox "Debe Ingresar Cuentas Validas para los empleados.", vbInformation, "Aviso"
            Me.Tab.Tab = 5
            Me.flexCuenta.SetFocus
            Exit Sub
        End If
    End If
    
    If MsgBox("Desea Grabar ??? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    oRh.ModificaUsuario Me.ctrRRHHGen.psCodigoPersona, Me.TxtBuscarUsuario.Text, GetMovNro(gsCodUser, gsCodAge), Me.chkAgregarAPlanillas.value, txtUbicacion.Text
    
    If Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "E" Then
        oRh.ModificaCUSSPSeguro Me.ctrRRHHGen.psCodigoPersona, txtCUSPP.Text, IIf(txtNumSeg.Text = "", "''", txtNumSeg.Text), CDate(Me.mskFecIng.Text), GetMovNro(gsCodUser, gsCodAge), Me.mskFecCese.Text, Me.mskFecAfilia.Text
    End If
        
    oRh.ModificaCuentaTarj Me.flexCuenta.GetRsNew, Me.mskTar.Text, Me.ctrRRHHGen.psCodigoPersona, GetMovNro(gsCodUser, gsCodAge)
    
    'LUCV20181220, Anexo01 de Acta 199-2018 (sólo para consulta)
        Set objPista = New COMManejador.Pista
        lsMovNro = GetMovNro(gsCodUser, gsCodAge)
        If lnTipo = gTipoOpeRegistro Then
            objPista.InsertarPista LogPistaRegistraProcesoSeleccion, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, Me.Caption, Me.ctrRRHHGen.psCodigoPersona, gCodigoPersona
            Set objPista = Nothing
        ElseIf lnTipo = gTipoOpeMantenimiento Then
            objPista.InsertarPista LogPistaModificaProcesoSeleccion, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, Me.Caption, Me.ctrRRHHGen.psCodigoPersona, gCodigoPersona
            Set objPista = Nothing
        End If
    'Fin LUCV20181220
        
    ClearScreen
    Activa False, CInt(lnContratoMantTpo)
End Sub

Private Function ValidaDatosArea(Index As Integer) As Boolean
Dim nItem As Integer
If FlexVista(Index).TextMatrix(CInt(FlexVista(Index).Rows) - 1, 1) = "" Or FlexVista(Index).TextMatrix(CInt(FlexVista(Index).Rows) - 1, 4) = "" Or FlexVista(Index).TextMatrix(CInt(FlexVista(Index).Rows) - 1, 6) = "" Or FlexVista(Index).TextMatrix(CInt(FlexVista(Index).Rows) - 1, 9) = "" Then
   'nItem = CInt(FlexVista(Index).Rows) - 1
    MsgBox "El Registro no puede ser grabado, por la existencia de campos vacíos..!!", vbInformation, "Aviso"
   'FlexVista(Index).EliminaFila (nItem)
   FlexVista(Index).SetFocus
   Exit Function
End If
   ValidaDatosArea = True
End Function 'NAGL ERS0742017 20171205

Private Sub cmdGrabarArea_Click(Index As Integer)
    Dim oRh As DActualizaDatosRRHH
    Dim oCont As DActualizaDatosContrato
    Set oCont = New DActualizaDatosContrato
    Dim i As Integer
    Set oRh = New DActualizaDatosRRHH
    Dim lsContrato As String
    
    If Me.ctrRRHHGen.psCodigoPersona = "" Then
        Exit Sub
    End If
    
    '***Agregado por ELRO el 20111219, según Acta N° 346-2011/TI-D
    If Me.ctrRRHHGen.psCodigoPersona = gsCodPersUser Then
        MsgBox "No puede agregar un cargo a su persona, solicite que lo haga otro usuario de RRHH", vbInformation, "Aviso"
        Exit Sub
    End If
    '***Fin Agregado por ELRO*************************************
    
    If MsgBox("Desea Grabar los cambios ???", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    If Index = 0 Then
        Set oRh = New DActualizaDatosRRHH
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 12) = "" Then
           If ValidaDatosArea(Index) Then 'NAGL ERS0742017 20171205
                oRh.AgregaRHCargo Me.ctrRRHHGen.psCodigoPersona, FechaHora(gdFecSis), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 4), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 9), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 11), GetMovNro(gsCodUser, gsCodAge)
           End If
        Else
            MsgBox "Solo se puede agregar, no se puede actualizar en esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = 1 Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = "" Then
            MsgBox "Deb ingresar un valor valido.", vbInformation, "Aviso"
            FlexVista(Index).row = FlexVista(Index).Rows - 1
            FlexVista(Index).col = 1
            Exit Sub
        End If
        
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 4) = "" Then
            oRh.AgregaRRHHAMP Me.ctrRRHHGen.psCodigoPersona, FechaHora(gdFecSis), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3), GetMovNro(gsCodUser, gsCodAge)
        Else
            MsgBox "Solo se puede agregar, no se puede actualizar en esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = 3 Then
        If Not IsNumeric(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1)) Then
            MsgBox "Debe ingresar una un tipo de sistema de pensiones valido.", vbInformation, "Aviso"
            FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = ""
            FlexVista(Index).col = 1
            Exit Sub
        ElseIf FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP And FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3) = "" Then
            MsgBox "Debe ingresar una AFP valido.", vbInformation, "Aviso"
            FlexVista(Index).col = 3
            Exit Sub
        ElseIf FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 5) = "" Then
            MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
            FlexVista(Index).col = 5
            Exit Sub
        End If
        
        
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6) = "" Then
            oRh.AgregaRRHHFondo Me.ctrRRHHGen.psCodigoPersona, FechaHora(gdFecSis), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), Right(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 4), 13), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 5), GetMovNro(gsCodUser, gsCodAge)
        Else
            MsgBox "Solo se puede agregar, no se puede actualizar en esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = 2 Then 'RHSueldos
        If Not IsNumeric(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1)) Then
            MsgBox "Debe ingresar una un sueldo valido.", vbInformation, "Aviso"
            FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = ""
            FlexVista(Index).col = 1
            Exit Sub
        ElseIf FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 2) = "" Then
            MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
            FlexVista(Index).col = 2
            Exit Sub
        End If
         
        
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3) = "" Then
            oCont.AgregaSueldo Me.ctrRRHHGen.psCodigoPersona, FechaHora(gdFecSis), Format(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), "#.00"), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 2), GetMovNro(gsCodUser, gsCodAge)
        Else
            MsgBox "Solo se puede agregar, no se puede actualizar en esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = 4 Then 'RHComentario
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 2) = "" Then
            oRh.AgregaRRHHComentario Me.ctrRRHHGen.psCodigoPersona, FechaHora(gdFecSis), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), GetMovNro(gsCodUser, gsCodAge)
        Else
            MsgBox "Solo se puede agregar, no se puede actualizar en esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
    ElseIf Index = 5 Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6) = "" Then
            lsContrato = oCont.GeCodContrato(Me.ctrRRHHGen.psCodigoPersona)
            oCont.AgregaContrato Me.ctrRRHHGen.psCodigoPersona, lsContrato, FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 5), GetMovNro(gsCodUser, gsCodAge), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1)
            oCont.AgregaContratoDet Me.ctrRRHHGen.psCodigoPersona, lsContrato, FechaHora(gdFecSis), Format(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3), gsFormatoFecha), Format(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 4), gsFormatoFecha), FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 5), GetMovNro(gsCodUser, gsCodAge)
        Else
            MsgBox "Solo se puede agregar, no se puede actualizar en esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If Index <> 8 Then
        Me.FlexVista(Index).Clear
        Me.FlexVista(Index).Rows = 2
        Me.FlexVista(Index).FormaCabecera
    End If
    
    ActivaArea False, Index
    CargaData Me.ctrRRHHGen.psCodigoPersona
End Sub

Private Sub cmdGrabarContrato_Click()
    If Me.ctrRRHHGen.psCodigoPersona = "" Then Exit Sub
    
    If MsgBox("Desea grabar el Contrato ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    Dim oRh As DActualizaDatosRRHH
    Set oRh = New DActualizaDatosRRHH

    oRh.ModificaContratoTextoUlt Me.ctrRRHHGen.psCodigoPersona, Me.richContrato.Text, GetMovNro(gsCodUser, gsCodAge)
    MsgBox "Grabación Satisfactoria.", vbInformation, "Aviso"
    
    Set oRh = Nothing
End Sub

Private Sub cmdGrabarFoto_Click()
    Dim oRh As DActualizaDatosRRHH
    Set oRh = New DActualizaDatosRRHH
    
    If ctrRRHHGen.psCodigoPersona = "" Then Exit Sub
    
    If Me.picFoto.Picture <> 0 Then
        oRh.ModificaFoto Me.ctrRRHHGen.psCodigoPersona, Me.picFoto, GetMovNro(gsCodUser, gsCodAge), lsRutaFoto
        MsgBox "Grabación Satisfactoria.", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
SavePicture picFoto, "c:\" + ctrRRHHGen.psCodigoPersona + ".jpg"
MsgBox "Grabado en :" + "c:\" + ctrRRHHGen.psCodigoPersona + ".jpg", vbInformation
End Sub

Private Sub ctrRRHHGen_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        ClearScreen
        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        CargaData Me.ctrRRHHGen.psCodigoPersona
    End If
End Sub

Private Sub ctrRRHHGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRRHH As DActualizaDatosRRHH
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        ctrRRHHGen.psCodigoEmpleado = Left(ctrRRHHGen.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHHGen.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(ctrRRHHGen.psCodigoEmpleado, gPersIdDNI)
           
        If Not (rsR.EOF And rsR.BOF) Then
            ctrRRHHGen.SpinnerValor = CInt(Right(ctrRRHHGen.psCodigoEmpleado, 5))
            ctrRRHHGen.psCodigoPersona = rsR.Fields("Codigo")
            ctrRRHHGen.psNombreEmpledo = rsR.Fields("Nombre")
            rsR.Close
            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHHGen.psCodigoEmpleado)
            CargaData Me.ctrRRHHGen.psCodigoPersona
            If cmdEditar.Enabled And cmdEditar.Visible Then
                Me.cmdEditar.SetFocus
            End If
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            ClearScreen
            ctrRRHHGen.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub FlexContrato_DblClick()
    Dim oRh As DActualizaDatosRRHH
    Set oRh = New DActualizaDatosRRHH
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    oPrevio.Show oRh.GetRRHHContratoTexto(Me.ctrRRHHGen.psCodigoPersona, Me.FlexContrato.TextMatrix(FlexContrato.row, 1)), "CONTRATO " & Me.FlexContrato.TextMatrix(FlexContrato.row, 2), True, 66

    Set oRh = Nothing
    Set oPrevio = Nothing
End Sub

Private Sub FlexContrato_OnRowChange(pnRow As Long, pnCol As Long)
    Dim oRh As DActualizaDatosRRHH
    Dim rsRH As ADODB.Recordset
    Set rsRH = New ADODB.Recordset
    Set oRh = New DActualizaDatosRRHH
            
    Set rsRH = oRh.GetRRHHFileDetalleContrato(Me.ctrRRHHGen.psCodigoPersona, Me.FlexContrato.TextMatrix(pnRow, 1))
    '0 SetFlexEdit Me.flexAdenda, rsRH
    
    '0 Me.richCont.Text = Me.flexAdenda.TextMatrix(1, flexAdenda.Cols - 1)
     
    Set oRh = Nothing
End Sub


Private Sub FlexHor_RowColChange()
    Dim rsD As ADODB.Recordset
    Dim oHor As DActualizaDatosHorarios
    
    If FlexHor.row Then
        If Not IsDate(Me.FlexHor.TextMatrix(FlexHor.row, 1)) Then
            lnRowAnt = FlexHor.row
            Exit Sub
        End If
        
        Set oHor = New DActualizaDatosHorarios
        Set rsD = New ADODB.Recordset
        
        
        Set rsD = oHor.GetHorariosDetalle(Me.ctrRRHHGen.psCodigoPersona, Format(CDate(Me.FlexHor.TextMatrix(FlexHor.row, 1)), gsFormatoFecha))
        
        If Not (rsD.EOF And rsD.BOF) Then
            Set Me.FlexDia.Recordset = rsD
        Else
            FlexDia.Clear
            FlexDia.Rows = 2
            FlexDia.FormaCabecera
        End If
        lnRowAnt = FlexHor.row
    End If
    
    If Me.FlexHor.TextMatrix(FlexHor.row, 3) = "1" Then
        FlexHor.ColumnasAEditar = "X-X-X-X"
    Else
        FlexHor.ColumnasAEditar = "X-1-2-X"
    End If
End Sub

Private Sub FlexVista_DblClick(Index As Integer)
    If Index = RHContratoMantTpo.RHContratoMantTpoAdenda Then
        If Me.FlexVista(Index).col = 5 Then
            CDialog.CancelError = False
            'On Error GoTo ErrHandler
            CDialog.Flags = cdlOFNHideReadOnly
            CDialog.Filter = "Archivos Bmp(*.txt)|*.txt"
            CDialog.FilterIndex = 2
            CDialog.ShowOpen
            Me.richContrato.LoadFile CDialog.FileName, 1
            FlexVista(Index).TextMatrix(Me.FlexVista(Index).row, 5) = richContrato.Text
        ElseIf Me.FlexVista(Index).col = 6 Then
            Dim oPrevio As Previo.clsPrevio
            Set oPrevio = New Previo.clsPrevio
            If FlexVista(Index).TextMatrix(FlexVista(Index).row, 5) <> "" Then
                oPrevio.Show FlexVista(Index).TextMatrix(FlexVista(Index).row, 5), Caption, True
            End If
            Set oPrevio = Nothing
        End If
    End If
End Sub

Private Sub FlexVista_EnterCell(Index As Integer)
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    
    If Index = RHContratoMantTpo.RHContratoMantTpoCargo Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) <> "" And FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3) = "" Then 'NAGL 20171209
           FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3) = oArea.GetNomArea(Left(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), 3))
        End If
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6) <> "" And FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 8) = "" Then 'NAGL 20171209
           FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 8) = oArea.GetNomArea(Left(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6), 3))
        End If
    End If
End Sub

Private Sub FlexVista_OnEnterTextBuscar(Index As Integer, psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    If Index = RHContratoMantTpo.RHContratoMantTpoAdenda Then
        If pnCol = 5 Then
            CDialog.CancelError = False
            On Error GoTo ErrHandler
            CDialog.Flags = cdlOFNHideReadOnly
            CDialog.Filter = "Archivos Bmp(*.txt)|*.txt"
            CDialog.FilterIndex = 2
            CDialog.ShowOpen
            Me.richContrato.LoadFile CDialog.FileName, 1
            FlexVista(Index).TextMatrix(pnRow, pnCol) = richContrato.Text
            psDataCod = richContrato.Text
        Else
        
        End If
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoSisPens Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) <> "0" Then
           FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 3) = ""
           FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 4) = ""
        End If
    End If
    
    
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub FlexVista_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    
    If Index = 2 And pnCol = 1 Then
        If Not IsNumeric(FlexVista(Index).TextMatrix(pnRow, pnCol)) Then
            FlexVista(Index).TextMatrix(pnRow, pnCol) = "0.00"
        End If
        If CCur(FlexVista(Index).TextMatrix(pnRow, pnCol)) < 0 Then
            Cancel = False
            MsgBox "Debe ingresar un sueldo mayor a cero.", vbInformation, "Aviso"
        End If
    End If
    
End Sub

Private Sub FlexVista_RowColChange(Index As Integer)
    Dim oCon As DConstantes
    Dim oAFP As DActualizaDatosAFP
    Dim oAMP As DActualizaAsistMedicaPrivada
    Set oAMP = New DActualizaAsistMedicaPrivada
    Set oCon = New DConstantes
    Set oAFP = New DActualizaDatosAFP
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsE As ADODB.Recordset
    Dim oDatoAreas As DActualizaDatosArea
    Set oDatoAreas = New DActualizaDatosArea
    Set rsE = New ADODB.Recordset
    'If Index = 1 Or Index = 2 Then
    '    If FlexManArea(Index).Col = 1 Then
    '        Set rs = oCon.GetAgencias(, , True, , Me.ctrRRHHGen.psCodigoPersona)
    '        Me.FlexManArea(Index).rsTextBuscar = rs
    '    End If
    'ElseIf Index = 4 Then
    '    If FlexManArea(Index).Col = 1 Then
    '        Set rs = oCon.GetConstante(gRHEmpleadoFonfoTipo, , , True)
    '        Me.FlexManArea(Index).rsTextBuscar = rs
    '    ElseIf FlexManArea(Index).TextMatrix(FlexManArea(Index).Row, 1) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP And FlexManArea(Index).Col = 3 Then
    '        FlexManArea(Index).ColumnasAEditar = "X-1-X-3-X-5"
    '        Set rs = oAFP.GetAFP(True)
    '        Me.FlexManArea(Index).rsTextBuscar = rs
    '    Else
    '        FlexManArea(Index).ColumnasAEditar = "X-1-X-X-X-5"
    '    End If
    'Else
    If Index = RHContratoMantTpo.RHContratoMantTpoCargo Then
        If FlexVista(Index).col = 1 Or FlexVista(Index).col = 6 Then
            Set rs = oDatoAreas.GetCargosAreas
            Me.FlexVista(Index).rsTextBuscar = rs
        ElseIf FlexVista(Index).col = 4 Then
            Set rs = oCon.GetAgencias(, , True, Left(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1), 3))
            Me.FlexVista(Index).rsTextBuscar = rs
        ElseIf FlexVista(Index).col = 9 Then
            Set rs = oCon.GetAgencias(, , True, Left(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 6), 3))
            Me.FlexVista(Index).rsTextBuscar = rs
        End If
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoSisPens Then
        
        If FlexVista(Index).col = 1 Then
            Set rs = oCon.GetConstante(gRHEmpleadoFonfoTipo, , , True)
            FlexVista(Index).rsTextBuscar = rs
        ElseIf FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = "" Then
            FlexVista(Index).ColumnasAEditar = "X-1-X-X-X-5-X"
            Exit Sub
        ElseIf Not IsNumeric(FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1)) Then
            FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = ""
            Exit Sub
        ElseIf FlexVista(Index).TextMatrix(FlexVista(Index).Rows - 1, 1) = RHEmpleadoFonfoTipo.RHEmpleadoFonfoTipoAFP And FlexVista(Index).col = 3 Then
            FlexVista(Index).ColumnasAEditar = "X-1-X-3-X-5-X"
            Set rs = oAFP.GetAFP(True)
            Me.FlexVista(Index).rsTextBuscar = rs
        Else
            FlexVista(Index).ColumnasAEditar = "X-1-X-X-X-5-X"
        End If
    ElseIf Index = RHContratoMantTpo.RHContratoMantTpoAMP Then
        If FlexVista(Index).col = 1 Then
            Set rs = oAMP.GetAsisMedPriv(True)
            Me.FlexVista(Index).rsTextBuscar = rs
        End If
    End If
    
    If Index = 0 Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).row, 12) = "" Then
            FlexVista(Index).lbEditarFlex = True
        Else
            FlexVista(Index).lbEditarFlex = False
        End If
    ElseIf Index = 1 Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).row, 4) = "" Then
            FlexVista(Index).lbEditarFlex = True
        Else
            FlexVista(Index).lbEditarFlex = False
        End If
    ElseIf Index = 3 Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).row, 6) = "" Then
            FlexVista(Index).lbEditarFlex = True
        Else
            FlexVista(Index).lbEditarFlex = False
        End If
    ElseIf Index = 2 Then 'RHSueldos
        If FlexVista(Index).TextMatrix(FlexVista(Index).row, 3) = "" Then
            FlexVista(Index).lbEditarFlex = True
        Else
            FlexVista(Index).lbEditarFlex = False
        End If
    ElseIf Index = 4 Then 'RHComentario
        If FlexVista(Index).TextMatrix(FlexVista(Index).row, 2) = "" Then
            FlexVista(Index).lbEditarFlex = True
        Else
            FlexVista(Index).lbEditarFlex = False
        End If
    ElseIf Index = 5 Then
        If FlexVista(Index).TextMatrix(FlexVista(Index).row, 4) = "" Then
            FlexVista(Index).lbEditarFlex = True
        Else
            FlexVista(Index).lbEditarFlex = False
        End If

        If Me.FlexVista(Index).col = 1 Then
            FlexVista(Index).rsTextBuscar = oCon.GetConstante(gRHContratoTipo, , , True)
        Else
            FlexVista(Index).rsTextBuscar = rsE
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim lsCad As String
    Dim oCon As DConstantes
    Dim rs As ADODB.Recordset
    Dim oDatoAreas As DActualizaDatosArea
    Set oDatoAreas = New DActualizaDatosArea
    Set rs = New ADODB.Recordset
    Set oCon = New DConstantes
    ActivaArea False, -1
    Activa False, CInt(lnContratoMantTpo)
    Set oCon = New DConstantes
    Set rs = oDatoAreas.GetCargosAreas
    Me.FlexVista(0).rsTextBuscar = rs
    Me.txtAgencia.rs = oCon.GetAgencias(, , True)
    
    Set oCon = Nothing
End Sub



Private Sub mskFecIng_GotFocus()
    mskFecIng.SelStart = 0
    mskFecIng.SelLength = 50
End Sub

Private Sub mskFecIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCUSPP.SetFocus
    End If
End Sub

Private Sub picFoto_DblClick()
    frmRHEmpFotoMayor.Ini ctrRRHHGen.psNombreEmpledo, Me.picFoto
End Sub


Private Sub RichContrato_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    oPrevio.Show Me.richContrato.Text, "Contrato", True, 66
    Set oPrevio = Nothing
End Sub

'Private Sub RRHH_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.Tab.Tab = 0
'        Dim oRRHH As DActualizaDatosRRHH
'        Dim rsR As ADODB.Recordset
'        Set oRRHH = New DActualizaDatosRRHH
'        RRHH.psCodigoEmpleado = Left(RRHH.psCodigoEmpleado, 1) & Format(Trim(Mid(RRHH.psCodigoEmpleado, 2)), "00000")
'        Dim oCon As DActualizaDatosContrato
'        Set oCon = New DActualizaDatosContrato
'
'        Set rsR = oRRHH.GetRRHH(RRHH.psCodigoEmpleado, gPersIdDNI)
'
'        If Not (rsR.EOF And rsR.BOF) Then
'            RRHH.SpinnerValor = CInt(Right(RRHH.psCodigoEmpleado, 5))
'
'            RRHH.psCodigoPersona = rsR.Fields("Codigo")
'            RRHH.psNombreEmpledo = rsR.Fields("Nombre")
'            RRHH.psDireccionPersona = rsR.Fields("Direccion")
'            RRHH.psDNIPersona = IIf(IsNull(rsR.Fields("ID")), "", rsR.Fields("ID"))
'            RRHH.psSueldoContrato = Format(rsR.Fields("Sueldo"), "#,##0.00")
'            RRHH.psFechaNacimiento = Format(rsR.Fields("Fecha"), gsFormatoFechaView)
'
'            rsR.Close
'            Set rsR = oRRHH.GetRRHHGeneralidades(RRHH.psCodigoEmpleado)
'
'            Me.MskFecIngreso.Text = Format(IIf(IsNull(rsR.Fields("FI")), "  /  /    ", rsR.Fields("FI")), gsFormatoFechaView)
'            Me.mskFecCese.Text = Format(IIf(IsNull(rsR.Fields("FF")), "  /  /    ", rsR.Fields("FF")), gsFormatoFechaView)
'            Me.TxtBuscaUsuario.Text = rsR.Fields("Usuario")
'            Me.txtCodigo.Text = RRHH.psCodigoEmpleado
'            UbicaCombo Me.cmbAgenciaActual, rsR.Fields("AgeAct")
'            Me.TxtBuscarAgeArea = rsR.Fields("AreCod") & rsR.Fields("AgeAsig")
'            Me.lblArea = Me.TxtBuscarAgeArea.psDescripcion
'            UbicaCombo Me.cmbEstado, rsR.Fields("Estado")
'
'            UbicaCombo Me.cmbTipo, Left(RRHH.psCodigoEmpleado, 1)
'            UbicaCombo Me.cmbCargo, IIf(IsNull(rsR.Fields("Descripcion")), "", rsR.Fields("Descripcion"))
'
'            If Left(Me.RRHH.psCodigoEmpleado, 1) = "E" Then
'                rsR.Close
'                Set rsR = oRRHH.GetRRHHEmpleado(RRHH.psCodigoPersona)
'
'                If Not (rsR.EOF And rsR.BOF) Then
'                    RRHH.SpinnerValor = CInt(Right(RRHH.psCodigoEmpleado, 5))
'
'                    UbicaCombo Me.cmbNivel, IIf(IsNull(rsR!Nivel), "", rsR!Nivel)
'                    UbicaCombo Me.cmbEmpAMPCod, IIf(IsNull(rsR!AMP), "", rsR!AMP)
'                    UbicaCombo Me.cmbAFP, rsR.Fields("AFP")
'                    Me.txtDiasvacaciones.Text = rsR.Fields("VacPen")
'                    If rsR.Fields("Condi") Then
'                        Me.chkEstable.Value = 1
'                    Else
'                        Me.chkEstable.Value = 0
'                    End If
'                    Me.mskFEstable.Text = Format(IIf(IsNull(rsR.Fields("dEstab")), "  /  /    ", rsR.Fields("dEstab")), gsFormatoFechaView)
'                    Me.lblCUSPPID.Caption = IIf(IsNull(rsR.Fields("CUSPP")), "", rsR.Fields("CUSPP"))
'                End If
'            End If
'
'            rsR.Close
'            Set rsR = oCon.GetContratos(RRHH.psCodigoPersona)
'            CargaCombo rsR, Me.cmbContrato, 2
'            cmbContrato.AddItem "<<NUEVO>>"
'            If Me.cmbContrato.ListCount > 0 And lnTipo <> gTipoOpeRegistro Then
'                cmbContrato.ListIndex = 0
'            Else
'                cmbContrato.ListIndex = Me.cmbContrato.ListCount - 1
'            End If
'
'
'            rsR.Close
'            Set rsR = oCon.GetSueldoContrato(RRHH.psCodigoPersona)
'            If Not (rsR.EOF And rsR.BOF) Then
'                CargaCombo rsR, Me.cmbSueldo, , 1, 0
'                If Me.cmbSueldo.ListCount > 0 Then cmbSueldo.ListIndex = 0
'            End If
'
'            If cmdNuevo.Enabled Then Me.cmdNuevo.SetFocus
'
'            Set oCon = Nothing
'        Else
'            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
'            ClearScreen
'            RRHH.SetFocus
'        End If
'
'        rsR.Close
'        Set rsR = Nothing
'    Else
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    End If
'End Sub

Private Sub ClearScreen()
    Set Me.picFoto = LoadPicture
End Sub

Private Sub Activa(pbvalor As Boolean, pnIndice As Integer)
    Me.cmdEditar.Visible = Not pbvalor
    Me.cmdGrabar.Enabled = pbvalor
    Me.cmdCancelar.Visible = pbvalor
    Me.cmdSalir.Enabled = Not pbvalor
    Me.fraGeneralidades.Enabled = pbvalor
    Dim i As Integer
    
    If lnTipo = gTipoOpeMantenimiento Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(3) = True
        Me.Tab.TabVisible(4) = True
        Me.cmdAsigCont.Visible = True
        Me.cmdGrabarContrato.Enabled = True
        
        
        If Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "E" Or Me.ctrRRHHGen.psCodigoEmpleado = "" Or Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "P" Or Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "F" Or Left(Me.ctrRRHHGen.psCodigoEmpleado, 1) = "D" Then
            fraManArea_.Enabled = pbvalor
        End If
        
        Me.cmdEditar.Visible = Not pbvalor
        Me.cmdGrabar.Enabled = pbvalor
        Me.cmdCancelar.Visible = pbvalor
        Me.cmdSalir.Enabled = Not pbvalor
        Me.fraGeneralidades.Enabled = pbvalor
        Me.cmdFoto.Visible = True
        Me.cmdGrabarFoto.Visible = True
        
        For i = 0 To cmdAgregarArea.Count - 1
            Me.cmdAgregarArea(i).Visible = False
            Me.cmdCancelarArea(i).Visible = False
            Me.cmdGrabarArea(i).Enabled = False
        Next i
    ElseIf lnTipo = gTipoOpeConsulta Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(3) = True
        Me.Tab.TabVisible(4) = True
        Me.cmdAsigCont.Visible = False
        Me.cmdGrabarContrato.Visible = False
        fraManArea_.Enabled = False
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdSalir.Enabled = Not pbvalor
        Me.fraGeneralidades.Enabled = pbvalor
        Me.cmdFoto.Visible = False
        Me.cmdGrabarFoto.Visible = False
        
        For i = 0 To cmdAgregarArea.Count - 1
            Me.cmdAgregarArea(i).Visible = False
            Me.cmdCancelarArea(i).Visible = False
            Me.cmdGrabarArea(i).Enabled = False
        Next i
        
        
    ElseIf lnTipo <> gTipoOpeRegistro Then
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(3) = True
        Me.Tab.TabVisible(4) = True
        Me.cmdAsigCont.Visible = False
        Me.cmdGrabarContrato.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdFoto.Visible = False
        Me.cmdGrabarFoto.Visible = False
        Me.fraManArea_.Enabled = False
        Me.cmdGrabarContrato.Visible = False
        Me.cmdAsigCont.Visible = False
        
    Else
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
        Me.Tab.TabVisible(2) = True
        Me.Tab.TabVisible(3) = True
        Me.Tab.TabVisible(4) = True
        Me.Tab.TabVisible(pnIndice) = True
        Me.cmdCancelar.Visible = False
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdGrabarFoto.Visible = False
        Me.cmdFoto.Visible = False
        Me.cmdGrabarContrato.Visible = False
        Me.cmdAsigCont.Visible = False
        Me.fraManArea_.Enabled = False
        
        
        For i = 0 To cmdAgregarArea.Count - 1
            If i = pnIndice Then
                'Me.fraManArea(I).Enabled = True
                Me.cmdAgregarArea(i).Visible = True
                Me.cmdCancelarArea(i).Visible = False
                Me.cmdGrabarArea(i).Enabled = True
            Else
                Me.cmdAgregarArea(i).Visible = False
                Me.cmdCancelarArea(i).Visible = False
                Me.cmdGrabarArea(i).Enabled = False
            End If
        Next i
        Me.Tab.Tab = 0
    End If
End Sub

Private Function Valida() As Boolean
    If TxtBuscarUsuario.Text = "" Then
        MsgBox "Debe Ingresar un Usuario del Sistema Operativo.", vbInformation, "Aviso"
        Valida = False
        Me.Tab.Tab = 0
        TxtBuscarUsuario.SetFocus
        Exit Function
    Else
        Valida = True
    End If
End Function

Public Sub Ini(pnTipo As TipoOpe, pnContratoMantTpo As RHContratoMantTpo, psCaption As String)
    lnTipo = pnTipo
    IniTab True
    Caption = psCaption
    
    If pnTipo <> gTipoOpeMantenimiento Then
        Activa False, CInt(pnContratoMantTpo)
    Else
        Activa False, CInt(pnContratoMantTpo)
    End If
    ActivaArea False, CInt(pnContratoMantTpo)
    Me.Tab.Tab = 0
    Me.TabDetalle.Tab = 0
    Me.Show 1
End Sub

Private Sub CargaData(psPersCod As String)
    Dim oRh As DActualizaDatosRRHH
    Dim rsRH As ADODB.Recordset
    Set rsRH = New ADODB.Recordset
    Set oRh = New DActualizaDatosRRHH
    Dim oPNL As DPeriodoNoLaborado
    Set oPNL = New DPeriodoNoLaborado
    Dim oInf As DActualizaDatosInformeSocial
    Set oInf = New DActualizaDatosInformeSocial
    Dim oHor As DActualizaDatosHorarios
    Set oHor = New DActualizaDatosHorarios
    Dim oCurT As DActualizaDatosCurriculum
    Set oCurT = New DActualizaDatosCurriculum
    
    Dim oDem As DMeritosDemeritos
    Set oDem = New DMeritosDemeritos
    
    Dim lsCad As String
    
    Me.lblRHEstado.Caption = oRh.GetRRHHEstado(psPersCod)
    
    Set rsRH = oRh.GetRRHHGeneralidades(psPersCod)
        
    If rsRH.EOF And rsRH.BOF Then
        Me.TxtBuscarUsuario.Text = ""
        Me.lblEstadoRes.Caption = "" 'rsRH!Estado
        Me.lblAgenciaAsignadaRes.Caption = "" 'IIf(IsNull(rsRH!AreaAsig), "", rsRH!AreaAsig) & "  -  " & IIf(IsNull(rsRH!Asig), "", rsRH!Asig)
        Me.lblAgenciaRes.Caption = "" ' IIf(IsNull(rsRH!AreaAct), "", rsRH!AreaAct) & "  -  " & IIf(IsNull(rsRH!Act), "", rsRH!Act)
        Me.lblCargoRes.Caption = "" ' rsRH!cRHCargoDescripcion
        Me.txtUbicacion.Text = ""
    Else
        Me.TxtBuscarUsuario.Text = IIf(IsNull(rsRH!usu), "", rsRH!usu)
        Me.lblEstadoRes.Caption = rsRH!Estado
        Me.lblAgenciaAsignadaRes.Caption = IIf(IsNull(rsRH!AreaAsig), "", rsRH!AreaAsig) & "  -  " & IIf(IsNull(rsRH!Asig), "", rsRH!Asig)
        Me.lblAgenciaRes.Caption = IIf(IsNull(rsRH!AreaAct), "", rsRH!AreaAct) & "  -  " & IIf(IsNull(rsRH!Act), "", rsRH!Act)
        Me.lblCargoRes.Caption = rsRH!cRHCargoDescripcion
        Me.chkAgregarAPlanillas.value = IIf(rsRH!AgregaPlanilla, 1, 0)
        If IsNull(rsRH!FecNac) Then
            Me.lblEdad.Caption = "NO ESPECIFICADO"
        Else
            Me.lblEdad.Caption = Str(DateDiff("yyyy", rsRH!FecNac, gdFecSis)) & " AÑOS "
        End If
        'ALPA 20110115*************************
        Me.txtUbicacion.Text = rsRH!cUbicacion
        '**************************************
    End If
    rsRH.Close
    
    Set rsRH = GetRRHHCuentas(psPersCod)
    If rsRH.EOF And rsRH.BOF Then
        Me.flexCuenta.Clear
        Me.flexCuenta.Rows = 2
        Me.flexCuenta.FormaCabecera
    Else
        Me.flexCuenta.rsFlex = rsRH
    End If
    If rsRH.State = 1 Then rsRH.Close
    
    Set rsRH = GetRRHHTarjeta(psPersCod)
    If rsRH.EOF And rsRH.BOF Then
        Me.mskTar.Mask = ""
        Me.mskTar.Text = ""
        Me.mskTar.Mask = "####-####-####-####"
    Else
        If IsNull(rsRH.Fields(0)) Or Not IsNumeric(rsRH.Fields(0)) Then
            Me.mskTar.Mask = ""
            Me.mskTar.Text = ""
            Me.mskTar.Mask = "####-####-####-####"
        Else
            Me.mskTar.Text = Format(rsRH.Fields(0), "0###-####-####-####")
        End If
    End If
    rsRH.Close
    
    Set rsRH = oRh.GetRRHHGeneralidadesDoc(psPersCod)
    If rsRH.EOF And rsRH.BOF Then
        Me.FlexDoc.Clear
        Me.FlexDoc.Rows = 2
        Me.FlexDoc.FormaCabecera
    Else
        Set Me.FlexDoc.Recordset = rsRH
    End If
    
    rsRH.Close
    Set rsRH = oRh.GetRRHHFamiliares(psPersCod)
    If rsRH.EOF And rsRH.BOF Then
        Me.FlexFamiliares.Clear
        Me.FlexFamiliares.Rows = 2
        Me.FlexFamiliares.FormaCabecera
    Else
        Set Me.FlexFamiliares.Recordset = rsRH
    End If
    
    rsRH.Close
    Set rsRH = oRh.GetRRHHContratos(psPersCod)
    If rsRH.EOF And rsRH.BOF Then
        Me.FlexContrato.Clear
        Me.FlexContrato.Rows = 2
        Me.FlexContrato.FormaCabecera
    Else
        Set Me.FlexContrato.Recordset = rsRH
    End If
    
        rsRH.Close
        Set rsRH = oRh.GetRRHHFileAsig(psPersCod)
        SetFlexEdit Me.FlexVista(0), rsRH
        
        rsRH.Close
        Set rsRH = oRh.GetRRHHFileEstado(psPersCod)
        SetFlexEdit Me.FlexEstado, rsRH
        
        rsRH.Close
        Set rsRH = oRh.GetRRHHFileSueldo(psPersCod)
        SetFlexEdit Me.FlexVista(2), rsRH
        
        rsRH.Close
        Set rsRH = oRh.GetRRHHFilePensiones(psPersCod)
        SetFlexEdit Me.FlexVista(3), rsRH

        rsRH.Close
        Set rsRH = oRh.GetRRHHFileAMP(psPersCod)
        SetFlexEdit Me.FlexVista(1), rsRH

        rsRH.Close
        Set rsRH = oRh.GetRRHHFileComentario(psPersCod)
        SetFlexEdit FlexVista(4), rsRH

        rsRH.Close
        Set rsRH = oRh.GetRRHHFileDetalleContrato(psPersCod)
        SetFlexEdit FlexVista(5), rsRH
        
        Me.richContrato.Text = oRh.GetRRHHContratoTexto(psPersCod)
        
        rsRH.Close 'vacaciones
        Set rsRH = oPNL.GetRHPeriodoNoLabPersona(psPersCod, CInt(RHEstadosTpo.RHEstadosTpoVacaciones))
        SetFlexEdit Me.FlexPerNoLab(0), rsRH
        
        rsRH.Close 'Permisos
        Set rsRH = oPNL.GetRHPeriodoNoLabPersona(psPersCod, CInt(RHEstadosTpo.RHEstadosTpoPermisosLicencias))
        SetFlexEdit Me.FlexPerNoLab(1), rsRH
        
        rsRH.Close 'Descansos
        Set rsRH = oPNL.GetRHPeriodoNoLabPersona(psPersCod, CInt(RHEstadosTpo.RHEstadosTpoSubsidiado))
        SetFlexEdit Me.FlexPerNoLab(2), rsRH
        
        rsRH.Close 'Descansos
        Set rsRH = oPNL.GetRHPeriodoNoLabPersona(psPersCod, CInt(RHEstadosTpo.RHEstadosTpoSuspendido))
        SetFlexEdit Me.FlexPerNoLab(3), rsRH
        
        rsRH.Close 'InfSoc
        Set rsRH = oInf.GetInformesSociales(psPersCod, 1)
        SetFlexEdit Me.flexInfSoc, rsRH
        
        rsRH.Close 'CUSSP
        Set rsRH = oRh.GetCUSSPSeguro(psPersCod)
        If Not (rsRH.EOF And rsRH.BOF) Then
            Me.txtCUSPP.Text = rsRH.Fields(0) & ""
            Me.txtNumSeg.Text = rsRH.Fields(1) & ""
            If IsDate(rsRH!dFecAfiliacion) Then
            mskFecAfilia.Text = Format(rsRH!dFecAfiliacion, gsFormatoFechaView)
            Else
            mskFecAfilia.Text = "__/__/____"
            End If
            Else
            mskFecAfilia.Text = "__/__/____"
        End If
        
        rsRH.Close 'CUSSP
        Set rsRH = oHor.GetHorarios(psPersCod)
        SetFlexEdit Me.FlexHor, rsRH
        
        rsRH.Close 'MERITOS DEMERITOS
        Set rsRH = oDem.GetMerDems(psPersCod)
        SetFlexEdit Me.FlexPer, rsRH
        
        rsRH.Close 'CURRICULUM
        Set rsRH = oCurT.GetCurriculums(psPersCod)
        Flex.rsFlex = rsRH
        
        'CURRICULUM ACTIVIDADES EXTRACURRICULARES
        Set rsRH = oCurT.GetCurriculumsExtra(psPersCod)
        Me.FlexExtra.rsFlex = rsRH
            
        Me.mskFecIng.Text = Format(oRh.GetFecIngEmp(psPersCod), gsFormatoFechaView)
        
        lsCad = oRh.GetFecCeseEmp(psPersCod)
        If IsDate(lsCad) Then
            mskFecCese.Text = Format(oRh.GetFecCeseEmp(psPersCod), gsFormatoFechaView)
        Else
            mskFecCese.Text = "__/__/____"
        End If
        
        Set rsRH = oRh.GetFoto(psPersCod)
        
        If Not (rsRH.EOF And rsRH.BOF) Then
            GetPicture rsRH.Fields(0), picFoto
        End If
        
        'LUCV20181220, Anexo01 de Acta 199-2018 (sólo para consulta)
        Set objPista = New COMManejador.Pista
        If lnTipo = gTipoOpeConsulta Then
            lsMovNro = GetMovNro(gsCodUser, gsCodAge)
            objPista.InsertarPista LogPistaConsultaProcesoSeleccion, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, Me.Caption, psPersCod, gCodigoPersona
            Set objPista = Nothing
        End If
        'Fin LUCV20181220
        rsRH.Close
End Sub

Public Sub IniRegistroEva(psPersCod As String, psCodEva As String)
    lnTipo = gTipoOpeRegistro
    lsCodPers = psPersCod
    lsCodEva = psCodEva
    lbRegistro = True
    Me.Show 1
End Sub

Private Sub richCont_DblClick()
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    oPrevio.Show Me.richCont.Text, "Contrato/Adenda", True, 66
End Sub

Private Sub txtAgencia_EmiteDatos()
    Me.lblAgencia.Caption = txtAgencia.psDescripcion
End Sub

Private Sub TxtBuscarUsuario_Click(psCodigo As String, psDescripcion As String)
    Dim lsUsuario As String
    TxtBuscarUsuario.Text = frmRHUsuarios.Ini
    If lsUsuario <> "" Then Me.TxtBuscarUsuario.Text = lsUsuario
End Sub

Private Sub TxtBuscarUsuario_EmiteDatos()
    Dim lsUsuario As String
    lsUsuario = frmRHUsuarios.Ini
    If lsUsuario <> "" Then TxtBuscarUsuario = lsUsuario
End Sub

Private Sub IniTab(Optional pbvalor As Boolean = False)
    Dim i As Integer
    For i = 0 To Me.Tab.Tabs - 1
        Me.Tab.TabVisible(i) = pbvalor
    Next i
End Sub

Private Sub txtCUSPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNumSeg.SetFocus
    End If
End Sub


