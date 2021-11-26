VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmLogisticaMantenimientoV 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Incidencias & Carga"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "FrmLogisticaMantenimientoV.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   7845
      TabIndex        =   16
      Top             =   6645
      Width           =   1455
   End
   Begin VB.Frame FrmLogisticaMantenimientoV 
      Height          =   945
      Left            =   45
      TabIndex        =   1
      Top             =   -45
      Width           =   7455
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5685
         TabIndex        =   26
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Placa"
         Height          =   195
         Left            =   5130
         TabIndex        =   25
         Top             =   615
         Width           =   405
      End
      Begin VB.Label LblModelo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1080
         TabIndex        =   5
         Top             =   570
         Width           =   3795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   585
         Width           =   525
      End
      Begin VB.Label LblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1095
         TabIndex        =   3
         Top             =   240
         Width           =   6165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Chofer"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   465
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5655
      Left            =   60
      TabIndex        =   0
      Top             =   945
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   9975
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   512
      TabMaxWidth     =   18
      TabCaption(0)   =   "Incidencias"
      TabPicture(0)   =   "FrmLogisticaMantenimientoV.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FlexEdit1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Carga"
      TabPicture(1)   =   "FrmLogisticaMantenimientoV.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FlexEdit2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin Sicmact.FlexEdit FlexEdit2 
         Height          =   4125
         Left            =   225
         TabIndex        =   24
         Top             =   1065
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   7276
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
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
         ColWidth0       =   -1
         RowHeight0      =   240
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame2 
         Height          =   945
         Left            =   210
         TabIndex        =   17
         Top             =   45
         Width           =   8895
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1185
            TabIndex        =   20
            Top             =   570
            Width           =   7485
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1170
            TabIndex        =   19
            Text            =   "2004/01/31"
            Top             =   180
            Width           =   1035
         End
         Begin VB.ComboBox CboDestino 
            Height          =   315
            Left            =   5790
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   165
            Width           =   2850
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   225
            TabIndex        =   23
            Top             =   615
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   255
            TabIndex        =   22
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Destino"
            Height          =   195
            Left            =   5055
            TabIndex        =   21
            Top             =   240
            Width           =   540
         End
      End
      Begin Sicmact.FlexEdit FlexEdit1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   15
         Top             =   1455
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6588
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
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
         ColWidth0       =   -1
         RowHeight0      =   240
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame Frame1 
         Height          =   1350
         Left            =   -74820
         TabIndex        =   6
         Top             =   45
         Width           =   8895
         Begin VB.ComboBox CboTipo 
            Height          =   315
            Left            =   5790
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   165
            Width           =   2850
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1170
            TabIndex        =   12
            Text            =   "2004/01/31"
            Top             =   180
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1170
            TabIndex        =   10
            Top             =   540
            Width           =   7485
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1185
            TabIndex        =   8
            Top             =   885
            Width           =   7485
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   5295
            TabIndex        =   14
            Top             =   255
            Width           =   315
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   255
            TabIndex        =   11
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Lugar"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   225
            TabIndex        =   7
            Top             =   930
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "FrmLogisticaMantenimientoV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub
