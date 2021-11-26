VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHGeneralidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   6210
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3645
      TabIndex        =   4
      Top             =   6210
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2445
      TabIndex        =   3
      Top             =   6210
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1245
      TabIndex        =   2
      Top             =   6210
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   6210
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   6255
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SicmactAdmin.ctrRRHH RRHH 
      Height          =   1905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3360
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
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   45
      TabIndex        =   6
      Top             =   6210
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1245
      TabIndex        =   7
      Top             =   6210
      Width           =   1095
   End
   Begin TabDlg.SSTab Tab 
      Height          =   4200
      Left            =   0
      TabIndex        =   8
      Top             =   1950
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7408
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Generalidades"
      TabPicture(0)   =   "frmRHGeneralidades.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FlexEdit1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Horario"
      TabPicture(1)   =   "frmRHGeneralidades.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlexEdit2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Comentarios"
      TabPicture(2)   =   "frmRHGeneralidades.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "List1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.TextBox Text1 
         Height          =   1560
         Left            =   270
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2445
         Width           =   7560
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   180
         TabIndex        =   11
         Top             =   420
         Width           =   7680
      End
      Begin SicmactAdmin.FlexEdit FlexEdit2 
         Height          =   3645
         Left            =   -74895
         TabIndex        =   9
         Top             =   420
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   6429
         HighLight       =   1
         AllowUserResizing=   3
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   240
      End
      Begin SicmactAdmin.FlexEdit FlexEdit1 
         Height          =   3630
         Left            =   -74895
         TabIndex        =   10
         Top             =   420
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   6403
         HighLight       =   1
         AllowUserResizing=   3
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
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   240
      End
   End
End
Attribute VB_Name = "frmRHGeneralidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
