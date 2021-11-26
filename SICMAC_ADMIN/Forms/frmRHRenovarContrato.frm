VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHRenovarContrato 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmRHRenovarContrato.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   2910
      Top             =   3435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   7650
      TabIndex        =   12
      Top             =   3435
      Width           =   1110
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   360
      Left            =   1260
      TabIndex        =   11
      Top             =   3435
      Width           =   1110
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   90
      TabIndex        =   10
      Top             =   3435
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   2565
      Left            =   75
      TabIndex        =   6
      Top             =   810
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4524
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmRHRenovarContrato.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contrato"
      TabPicture(1)   =   "frmRHRenovarContrato.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContrato"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   8535
         Begin VB.TextBox txtComentario 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   105
            TabIndex        =   17
            Top             =   915
            Width           =   8370
         End
         Begin MSMask.MaskEdBox mskIni 
            Height          =   300
            Left            =   1935
            TabIndex        =   15
            Top             =   405
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskFin 
            Height          =   300
            Left            =   4530
            TabIndex        =   16
            Top             =   405
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame fraContrato 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   -74940
         TabIndex        =   8
         Top             =   360
         Width           =   8535
         Begin VB.CommandButton cmdCagar 
            Caption         =   "&Cargar"
            Height          =   360
            Left            =   3750
            TabIndex        =   13
            Top             =   1635
            Width           =   1110
         End
         Begin RichTextLib.RichTextBox R 
            Height          =   1320
            Left            =   120
            TabIndex        =   9
            Top             =   255
            Width           =   8280
            _ExtentX        =   14605
            _ExtentY        =   2328
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmRHRenovarContrato.frx":0342
         End
      End
   End
   Begin Sicmact.TxtBuscar txtEmpleado 
      Height          =   285
      Left            =   1125
      TabIndex        =   0
      Top             =   465
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
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
      Enabled         =   0   'False
      Enabled         =   0   'False
      Appearance      =   0
      TipoBusqueda    =   7
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
   Begin Sicmact.TxtBuscar txtTpoCon 
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   105
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
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
      Enabled         =   0   'False
      Enabled         =   0   'False
      Appearance      =   0
      sTitulo         =   ""
      EnabledText     =   0   'False
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "C&ancelar"
      Height          =   360
      Left            =   1260
      TabIndex        =   14
      Top             =   3435
      Width           =   1110
   End
   Begin VB.Label lblTpoConRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   2940
      TabIndex        =   5
      Top             =   90
      Width           =   5820
   End
   Begin VB.Label lblPersG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   2970
      TabIndex        =   4
      Top             =   435
      Width           =   5790
   End
   Begin VB.Label lblTpoCon 
      Caption         =   "Tpo.Contrato :"
      Height          =   255
      Left            =   45
      TabIndex        =   3
      Top             =   135
      Width           =   1065
   End
   Begin VB.Label lblPers 
      Caption         =   "Empleado"
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   480
      Width           =   1050
   End
End
Attribute VB_Name = "frmRHRenovarContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
