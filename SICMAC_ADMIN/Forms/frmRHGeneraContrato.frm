VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHGeneraContrato 
   Caption         =   "Generacion Contrato"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   Icon            =   "frmRHGeneraContrato.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7905
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   49
      Text            =   "Combo1"
      Top             =   7320
      Width           =   2775
   End
   Begin Sicmact.ctrRRHHGen ctrRRHHGen1 
      Height          =   1200
      Left            =   360
      TabIndex        =   48
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
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
   Begin VB.TextBox txtSueldo 
      Height          =   285
      Left            =   5400
      TabIndex        =   44
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtCargoAsignado 
      Height          =   285
      Left            =   1800
      TabIndex        =   43
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtAreaActual 
      Height          =   285
      Left            =   5400
      TabIndex        =   42
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtAreaAsignada 
      Height          =   285
      Left            =   1800
      TabIndex        =   41
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtAgeActual 
      Height          =   285
      Left            =   5400
      TabIndex        =   40
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtAgeAsignada 
      Height          =   285
      Left            =   1800
      TabIndex        =   39
      Top             =   2400
      Width           =   2175
   End
   Begin MSMask.MaskEdBox MEFechaFin 
      Height          =   285
      Left            =   5400
      TabIndex        =   38
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MEFechaInicio 
      Height          =   285
      Left            =   1800
      TabIndex        =   37
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtnumContrato 
      Height          =   285
      Left            =   1800
      TabIndex        =   36
      Top             =   1440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Direccion Procedencia"
      TabPicture(0)   =   "frmRHGeneraContrato.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Direccion Actual"
      TabPicture(1)   =   "frmRHGeneraContrato.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdEditar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdGrabar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdCancelar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6360
         TabIndex        =   47
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   5040
         TabIndex        =   46
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   3600
         TabIndex        =   45
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ubicacion Geografica"
         Height          =   2145
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   7605
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   9
            Left            =   2235
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1140
            Width           =   2190
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   8
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   7
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   6
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   525
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   24
            Top             =   1680
            Width           =   5200
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2235
            TabIndex        =   35
            Top             =   900
            Width           =   600
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4680
            TabIndex        =   34
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   210
            TabIndex        =   33
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2265
            TabIndex        =   32
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4695
            TabIndex        =   31
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   1710
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ubicacion Geografica"
         Height          =   2145
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   7605
         Begin VB.TextBox txtPersDireccDomicilio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   21
            Top             =   1680
            Width           =   5200
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   525
            Width           =   1815
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   2235
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1140
            Width           =   2190
         End
         Begin VB.Label lblPersDireccDomicilio 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio Procedencia"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   1710
            Width           =   1575
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4695
            TabIndex        =   20
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2265
            TabIndex        =   19
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4680
            TabIndex        =   17
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2235
            TabIndex        =   16
            Top             =   900
            Width           =   600
         End
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Sueldo"
      Height          =   195
      Left            =   4200
      TabIndex        =   8
      Top             =   3405
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cargo Asignado"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3405
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Area Actual"
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   2925
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Area Asignada"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   2925
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "AG Actual"
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   2445
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "AG Asignada"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2445
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fin Contrato"
      Height          =   195
      Left            =   4200
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Incio Contrato"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1965
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numero Contrato"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1200
   End
End
Attribute VB_Name = "frmRHGeneraContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
