VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInvActivarBien_NEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activación de Bien"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "frmInvActivarBien_NEW.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   5520
      TabIndex        =   7
      Top             =   3210
      Width           =   1050
   End
   Begin VB.ComboBox cboTpoBien 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3210
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selección del Documento Origen"
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
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   40
      Width           =   8740
      Begin VB.TextBox txtOrdenNombre 
         Height          =   285
         Left            =   4155
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   440
         Width           =   3255
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "&Cargar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   7440
         TabIndex        =   4
         Top             =   440
         Width           =   1050
      End
      Begin VB.ComboBox cboDocOrigen 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   440
         Width           =   2055
      End
      Begin Sicmact.TxtBuscar txtOrdenCod 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   435
         Width           =   1740
         _extentx        =   3069
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmInvActivarBien_NEW.frx":030A
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Documento Origen:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   210
         Width           =   1455
      End
   End
   Begin Sicmact.FlexEdit feOrden 
      Height          =   2205
      Left            =   40
      TabIndex        =   5
      Top             =   960
      Width           =   8760
      _extentx        =   15452
      _extenty        =   3889
      cols0           =   12
      highlight       =   1
      allowuserresizing=   1
      encabezadosnombres=   "#-nMovItem-Ag. Cod-Ag. Destino-Objeto-Descripción-Cant.-B. Activa-P.U.-Total-CodAgeCodInv-Aux"
      encabezadosanchos=   "350-0-0-1400-1400-3000-800-850-800-900-0-0"
      font            =   "frmInvActivarBien_NEW.frx":0336
      font            =   "frmInvActivarBien_NEW.frx":035E
      font            =   "frmInvActivarBien_NEW.frx":0386
      font            =   "frmInvActivarBien_NEW.frx":03AE
      font            =   "frmInvActivarBien_NEW.frx":03D6
      fontfixed       =   "frmInvActivarBien_NEW.frx":03FE
      tipobusqueda    =   0
      columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-L-L-L-C-C-R-R-L-C"
      formatosedit    =   "0-0-0-0-0-0-1-0-2-2-0-0"
      cantentero      =   9
      textarray0      =   "#"
      selectionmode   =   1
      lbpuntero       =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   345
      rowheight0      =   300
   End
   Begin TabDlg.SSTab TabActivaBien 
      Height          =   3765
      Left            =   75
      TabIndex        =   51
      Top             =   3600
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   6641
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Activo Fijo"
      TabPicture(0)   =   "frmInvActivarBien_NEW.frx":0424
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAFCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAFActivar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   " &Bien No Depreciable"
      TabPicture(1)   =   "frmInvActivarBien_NEW.frx":0440
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBNDActivar"
      Tab(1).Control(1)=   "cmdBNDCancelar"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Activo Compuesto"
      TabPicture(2)   =   "frmInvActivarBien_NEW.frx":045C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "cmdACCancelar"
      Tab(2).Control(2)=   "cmdACActivar"
      Tab(2).Control(3)=   "Frame6"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Mejora Componente"
      TabPicture(3)   =   "frmInvActivarBien_NEW.frx":0478
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "cmdMCCancelar"
      Tab(3).Control(3)=   "cmdMCActivar"
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame9 
         Caption         =   "Datos de Registro"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   108
         Top             =   480
         Width           =   8505
         Begin VB.TextBox txtMCActCompMejoraNombre 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   1275
            Width           =   2775
         End
         Begin VB.TextBox txtMCDepreTribMes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtMCDepreContMes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            TabIndex        =   44
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtMCDepreContAnio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            TabIndex        =   43
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtMCNombre 
            Height          =   285
            Left            =   1320
            MaxLength       =   300
            TabIndex        =   41
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtMCInventarioCod 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker txtMCFechaConformidad 
            Height          =   315
            Left            =   7150
            TabIndex        =   40
            Top             =   240
            Width           =   1280
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   130088961
            CurrentDate     =   41414
         End
         Begin Sicmact.TxtBuscar txtMCActCompMejoraCod 
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1275
            Width           =   1125
            _extentx        =   1984
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":0494
            appearance      =   1
         End
         Begin VB.Label Label24 
            Caption         =   "Activo Componente a mejorar:"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   1035
            Width           =   3255
         End
         Begin VB.Label Label52 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   7320
            TabIndex        =   116
            Top             =   615
            Width           =   375
         End
         Begin VB.Label Label51 
            Caption         =   "Tiempo Deprec. Trib."
            Height          =   255
            Left            =   4200
            TabIndex        =   115
            Top             =   615
            Width           =   1575
         End
         Begin VB.Label Label50 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   7305
            TabIndex        =   114
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label49 
            Caption         =   "Año:"
            Height          =   255
            Left            =   5985
            TabIndex        =   113
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label48 
            Caption         =   "Tiempo Deprec. Cont."
            Height          =   255
            Left            =   4200
            TabIndex        =   112
            Top             =   975
            Width           =   1575
         End
         Begin VB.Label Label40 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label39 
            Caption         =   "F. Conformidad:"
            Height          =   255
            Left            =   4200
            TabIndex        =   110
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label38 
            Caption         =   "Cód. Inventario:"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datos de Registro"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   100
         Top             =   480
         Width           =   8505
         Begin VB.TextBox txtACInventarioCod 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtACNombre 
            Height          =   285
            Left            =   1320
            MaxLength       =   300
            TabIndex        =   30
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtACAreaAgeNombre 
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   1310
            Width           =   1740
         End
         Begin VB.TextBox txtACPersonaNombre 
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   101
            Top             =   1310
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker txtACFechaConformidad 
            Height          =   315
            Left            =   7150
            TabIndex        =   29
            Top             =   240
            Width           =   1280
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   130088961
            CurrentDate     =   41414
         End
         Begin Sicmact.TxtBuscar txtACPersonaCod 
            Height          =   255
            Left            =   4920
            TabIndex        =   32
            Top             =   1305
            Width           =   1455
            _extentx        =   2566
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":04C0
            appearance      =   1
            tipobusqueda    =   3
            stitulo         =   ""
         End
         Begin Sicmact.TxtBuscar txtACAreaAgeCod 
            Height          =   255
            Left            =   1320
            TabIndex        =   31
            Top             =   1310
            Width           =   1000
            _extentx        =   1773
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":04EC
            appearance      =   1
         End
         Begin VB.Label Label47 
            Caption         =   "Cód. Inventario:"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label46 
            Caption         =   "F. Conformidad:"
            Height          =   255
            Left            =   4200
            TabIndex        =   106
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label45 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label44 
            Caption         =   "Área/Agencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label43 
            Caption         =   "Persona:"
            Height          =   255
            Left            =   4200
            TabIndex        =   103
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Caracteristicas del Bien"
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   96
         Top             =   2205
         Width           =   8505
         Begin VB.TextBox txtMCMarca 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   46
            Top             =   235
            Width           =   2775
         End
         Begin VB.TextBox txtMCSerie 
            Height          =   285
            Left            =   5280
            MaxLength       =   499
            TabIndex        =   47
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtMCModelo 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   48
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label37 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label36 
            Caption         =   "Serie/caract:"
            Height          =   255
            Left            =   4200
            TabIndex        =   98
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label35 
            Caption         =   "Modelo:"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdMCCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -67440
         TabIndex        =   50
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdMCActivar 
         Caption         =   "&Activar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -68520
         TabIndex        =   49
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdACActivar 
         Caption         =   "&Activar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -68520
         TabIndex        =   37
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdACCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -67440
         TabIndex        =   38
         Top             =   3330
         Width           =   1050
      End
      Begin VB.Frame Frame7 
         Caption         =   "Caracteristicas del Bien"
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   92
         Top             =   2205
         Width           =   8505
         Begin VB.CommandButton cmdACComponentes 
            Caption         =   "&Registrar Componentes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Left            =   4200
            TabIndex        =   36
            Top             =   600
            Width           =   1890
         End
         Begin VB.TextBox txtACModelo 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   35
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtACSerie 
            Height          =   285
            Left            =   5280
            MaxLength       =   499
            TabIndex        =   34
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtACMarca 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   33
            Top             =   235
            Width           =   2775
         End
         Begin VB.Label Label34 
            Caption         =   "Modelo:"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "Serie/caract:"
            Height          =   255
            Left            =   4200
            TabIndex        =   94
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label32 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos de Registro"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   76
         Top             =   480
         Width           =   8505
         Begin VB.TextBox txtBNDInventarioCod 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtBNDNombre 
            Height          =   285
            Left            =   1320
            MaxLength       =   300
            TabIndex        =   20
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtBNDAreaAgeNombre 
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   1310
            Width           =   1740
         End
         Begin VB.TextBox txtBNDPersonaNombre 
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   1310
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker txtBNDFechaConformidad 
            Height          =   315
            Left            =   7150
            TabIndex        =   19
            Top             =   240
            Width           =   1280
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   127598593
            CurrentDate     =   41414
         End
         Begin Sicmact.TxtBuscar txtBNDPersonaCod 
            Height          =   255
            Left            =   4920
            TabIndex        =   22
            Top             =   1305
            Width           =   1455
            _extentx        =   2566
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":0518
            appearance      =   1
            tipobusqueda    =   3
            stitulo         =   ""
         End
         Begin Sicmact.TxtBuscar txtBNDAreaAgeCod 
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   1310
            Width           =   1000
            _extentx        =   1773
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":0544
            appearance      =   1
         End
         Begin VB.Label Label31 
            Caption         =   "Cód. Inventario:"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "F. Conformidad:"
            Height          =   255
            Left            =   4200
            TabIndex        =   82
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Área/Agencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Persona:"
            Height          =   255
            Left            =   4200
            TabIndex        =   79
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Caracteristicas del Bien"
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   72
         Top             =   2205
         Width           =   8505
         Begin VB.TextBox txtBNDMarca 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   23
            Top             =   235
            Width           =   2775
         End
         Begin VB.TextBox txtBNDSerie 
            Height          =   285
            Left            =   5280
            MaxLength       =   499
            TabIndex        =   24
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtBNDModelo 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   25
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label21 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "Serie/caract:"
            Height          =   255
            Left            =   4200
            TabIndex        =   74
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Modelo:"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdBNDCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -67440
         TabIndex        =   27
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdBNDActivar 
         Caption         =   "&Activar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   -68520
         TabIndex        =   26
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdAFActivar 
         Caption         =   "&Activar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6480
         TabIndex        =   16
         Top             =   3330
         Width           =   1050
      End
      Begin VB.CommandButton cmdAFCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   7560
         TabIndex        =   17
         Top             =   3330
         Width           =   1050
      End
      Begin VB.Frame Frame3 
         Caption         =   "Caracteristicas del Bien"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   67
         Top             =   2200
         Width           =   8505
         Begin VB.TextBox txtAFModelo 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   15
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtAFSerie 
            Height          =   285
            Left            =   5280
            MaxLength       =   499
            TabIndex        =   14
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtAFMarca 
            Height          =   285
            Left            =   1320
            MaxLength       =   499
            TabIndex        =   13
            Top             =   235
            Width           =   2775
         End
         Begin VB.Label Label18 
            Caption         =   "Modelo:"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Serie/caract:"
            Height          =   255
            Left            =   4200
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Marca:"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Registro"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   8505
         Begin VB.TextBox txtAFDepreTribMes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtAFDepreContMes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            TabIndex        =   86
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAFDepreContAnio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6360
            TabIndex        =   85
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAFPersonaNombre 
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1310
            Width           =   2055
         End
         Begin VB.TextBox txtAFAreaAgeNombre 
            Height          =   285
            Left            =   2320
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   1310
            Width           =   1740
         End
         Begin VB.TextBox txtAFNombre 
            Height          =   285
            Left            =   1320
            MaxLength       =   300
            TabIndex        =   10
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtAFInventarioCod 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker txtAFFechaConformidad 
            Height          =   315
            Left            =   7150
            TabIndex        =   9
            Top             =   240
            Width           =   1280
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   127598593
            CurrentDate     =   41414
         End
         Begin Sicmact.TxtBuscar txtAFPersonaCod 
            Height          =   255
            Left            =   4920
            TabIndex        =   12
            Top             =   1305
            Width           =   1455
            _extentx        =   2566
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":0570
            appearance      =   1
            tipobusqueda    =   3
            stitulo         =   ""
         End
         Begin Sicmact.TxtBuscar txtAFAreaAgeCod 
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   1310
            Width           =   1000
            _extentx        =   1773
            _extenty        =   450
            appearance      =   1
            appearance      =   1
            font            =   "frmInvActivarBien_NEW.frx":059C
            appearance      =   1
         End
         Begin VB.Label Label12 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   7320
            TabIndex        =   91
            Top             =   615
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Tiempo Deprec. Trib."
            Height          =   255
            Left            =   4200
            TabIndex        =   90
            Top             =   615
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   7305
            TabIndex        =   89
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Año:"
            Height          =   255
            Left            =   5985
            TabIndex        =   88
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "Tiempo Deprec. Cont."
            Height          =   255
            Left            =   4200
            TabIndex        =   84
            Top             =   975
            Width           =   1575
         End
         Begin VB.Label Label15 
            Caption         =   "Persona:"
            Height          =   255
            Left            =   4200
            TabIndex        =   65
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Área/Agencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "F. Conformidad:"
            Height          =   255
            Left            =   4200
            TabIndex        =   61
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Cód. Inventario:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   270
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   52
         Top             =   495
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Producto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nro. Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nro. Cta Antigua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Participación"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "SaldoCont"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SaldoDisp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Motivo de Bloque"
            Object.Width           =   7231
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67680
         TabIndex        =   57
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblSolesAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70815
         TabIndex        =   56
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   55
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
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
         Left            =   -73185
         TabIndex        =   54
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   53
         Top             =   3465
         Width           =   525
      End
   End
   Begin Sicmact.Usuario oUser 
      Left            =   8520
      Top             =   -120
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Bien:"
      Height          =   255
      Left            =   2400
      TabIndex        =   58
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmInvActivarBien_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmInvActivarBien_NEW
'** Descripción : Nueva Activación de Bienes creado segun ERS059-2013
'** Creación : EJVG, 20130518 09:00:00 AM
'********************************************************************
Option Explicit
Dim fnMoneda As Moneda
Dim fnFormTamanioIni As Double, fnFormTamanioActiva As Double

Dim fnDocOrigen As Integer
Dim fnMovNro As Long, fnMovItem As Long
Dim fnBANCod As Integer
Dim fsObjetoCod As String
Dim fnMonto As Currency
Dim fsAreaAgeCodInv As String
Dim fMatComponente As Variant
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Public Sub Inicio(ByVal pnMoneda As Moneda)
    fnFormTamanioIni = 4050
    fnFormTamanioActiva = 7875
    fnMoneda = pnMoneda
    Show 1
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call InicializaControlesGeneral
End Sub
Private Sub InicializaControlesGeneral()
    Height = fnFormTamanioIni
    Call ListarDocOrigen
    Call ListarTipoBien
    txtOrdenCod.Text = ""
    txtOrdenNombre.Text = ""
End Sub
Private Sub cboDocOrigen_Click()
    Dim obj As New DBien
    Dim lnTpoDocOrigen As Integer
    lnTpoDocOrigen = CInt(Right(cboDocOrigen.Text, 2))
    
    On Error GoTo ErrCboDocOrigen_Click
    Screen.MousePointer = 11
    
    txtOrdenCod.Text = ""
    txtOrdenNombre.Text = ""
    cancela_busqueda_actual
    
    fnDocOrigen = lnTpoDocOrigen
    Select Case lnTpoDocOrigen
        Case LogTipoDocOrigenActivaBien.OrdenCompra
            txtOrdenCod.Enabled = True
            cmdCargar.Enabled = True
            txtOrdenCod.rs = obj.ListarOrdenCompra(fnMoneda)
        Case LogTipoDocOrigenActivaBien.OrdenServicio
            txtOrdenCod.Enabled = True
            cmdCargar.Enabled = True
            txtOrdenCod.rs = obj.ListarOrdenCompra(fnMoneda)
    End Select
    Set obj = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCboDocOrigen_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cboDocOrigen_KeyPress(KeyAscii As Integer)
    Dim lnTpoDocOrigen As Integer
    If KeyAscii = 13 And cboDocOrigen.ListIndex <> -1 Then
        lnTpoDocOrigen = CInt(Right(cboDocOrigen.Text, 2))
        Select Case lnTpoDocOrigen
            Case LogTipoDocOrigenActivaBien.OrdenCompra, LogTipoDocOrigenActivaBien.OrdenServicio
                txtOrdenCod.SetFocus
        End Select
    End If
End Sub
Private Sub txtOrdenCod_EmiteDatos()
    txtOrdenNombre.Text = txtOrdenCod.psDescripcion
    cancela_busqueda_actual
End Sub
Private Sub cmdCargar_Click()
    Dim obj As New DBien
    Dim rs As New ADODB.Recordset
    Dim lnMovNro As Long
    Dim lnTpoDocOrigen As Integer
    Dim i As Long
    
    If txtOrdenCod.Text = "" Then
        MsgBox "Ud. debe seleccionar primero la Orden para Activar los Bienes", vbInformation, "Aviso"
        txtOrdenCod.SetFocus
        Exit Sub
    End If
    
    lnTpoDocOrigen = CInt(Right(cboDocOrigen.Text, 2))
    lnMovNro = CLng(Right(Me.txtOrdenNombre.Text, 12))
    
    cancela_busqueda_actual
    
    If txtOrdenCod.Text <> "" Then
        Select Case lnTpoDocOrigen
            Case LogTipoDocOrigenActivaBien.OrdenCompra
                Set rs = obj.ListarOrdenCompraDet(lnMovNro)
                For i = 1 To rs.RecordCount
                    feOrden.AdicionaFila
                    feOrden.TextMatrix(feOrden.row, 1) = rs!nMovItem
                    feOrden.TextMatrix(feOrden.row, 2) = rs!cAgeCod
                    feOrden.TextMatrix(feOrden.row, 3) = rs!cAgeDescripcion
                    feOrden.TextMatrix(feOrden.row, 4) = rs!cObjetoCod
                    feOrden.TextMatrix(feOrden.row, 5) = rs!cBSDescripcion
                    feOrden.TextMatrix(feOrden.row, 6) = rs!nMovCant
                    feOrden.TextMatrix(feOrden.row, 7) = rs!nMovCantAct
                    feOrden.TextMatrix(feOrden.row, 8) = Format(rs!nPU, gsFormatoNumeroView)
                    feOrden.TextMatrix(feOrden.row, 9) = Format(rs!nMovImporte, gsFormatoNumeroView)
                    feOrden.TextMatrix(feOrden.row, 10) = rs!cAreaAgeCodInv
                    rs.MoveNext
                Next
            Case LogTipoDocOrigenActivaBien.OrdenServicio
                Set rs = obj.ListarOrdenCompraDet(lnMovNro)
        End Select
    Else
        MsgBox "Ud. primero debe seleccionar la Orden", vbInformation, "Aviso"
        Exit Sub
    End If
    feOrden.TopRow = 1
    feOrden.row = 1
    feOrden.SetFocus
    Set obj = Nothing
End Sub
Private Sub feOrden_OnRowChange(pnRow As Long, pnCol As Long)
    Height = fnFormTamanioIni
End Sub
Private Sub feOrden_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(feOrden.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub cboTpoBien_Click()
    Height = fnFormTamanioIni
End Sub
Private Sub cmdProcesar_Click()
    Dim oBien As New DBien
    Dim lnTpoBienActivar As Integer
    Dim lsBSCod As String, lsAgeCod As String
    Dim lsDescripcion As String
    Dim lnTab As Integer
    Dim N As Integer
    Dim bTrans As Boolean
    Dim lnCant As Integer, lnCantActivada As Integer
    
    On Error GoTo ErrProcesar
    If FlexVacio(feOrden) Then
        MsgBox "Ud. debe seleccionar el Bien a activar", vbInformation, "Aviso"
        feOrden.SetFocus
        Exit Sub
    End If
    If cboTpoBien.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el Tipo de Bien a activar", vbInformation, "Aviso"
        cboTpoBien.SetFocus
        Exit Sub
    End If
    
    fnMovNro = CLng(Right(Me.txtOrdenNombre.Text, 12))
    fnMovItem = CInt(feOrden.TextMatrix(feOrden.row, 1))
    fsObjetoCod = feOrden.TextMatrix(feOrden.row, 4)
    fnBANCod = oBien.RecuperaCodigoBAN(fsObjetoCod)
    fnMonto = CCur(feOrden.TextMatrix(feOrden.row, 8))
    lsAgeCod = Trim(feOrden.TextMatrix(feOrden.row, 2))
    lnTpoBienActivar = CInt(Right(cboTpoBien.Text, 2))
    lsDescripcion = UCase(feOrden.TextMatrix(feOrden.row, 5))
    fsAreaAgeCodInv = feOrden.TextMatrix(feOrden.row, 10)
    lnCant = feOrden.TextMatrix(feOrden.row, 6)
    lnCantActivada = feOrden.TextMatrix(feOrden.row, 7)
    
    If lnCantActivada >= lnCant Then
        MsgBox "Se ha llegado al Máximo de Activaciones, no se puede continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Select Case lnTpoBienActivar
        Case LogTipoBienesNew.ActivoFijo
            If fnDocOrigen = 1 And Mid(fsObjetoCod, 2, 2) <> "12" Then 'Valida sea AF
                MsgBox "Para activar Bienes como ACTIVOS FIJOS, necesariamente debe ser un Activo Fijo", vbInformation, "Aviso"
                feOrden.SetFocus
                Exit Sub
            End If
            limpiarDatosAF
            'txtAFInventarioCod.Text = DevolverCorrelativoAF(fsObjetoCod, IIf(Mid(fsAreaAgeCodInv, 4, 2) = "", "01", Mid(fsAreaAgeCodInv, 4, 2)))
             txtAFInventarioCod.Text = DevolverCorrelativoAF(fsObjetoCod, IIf(lsAgeCod = "", "01", lsAgeCod)) ' *** PEAC 20140527
            
            txtAFNombre.Text = lsDescripcion
            txtAFAreaAgeCod.Text = fsAreaAgeCodInv
            txtAFAreaAgeCod_EmiteDatos
            If fnBANCod = 0 Then MsgBox "El presente Bien No tiene configurado el Tipo de Activo Fijo", vbInformation, "Aviso"
        Case LogTipoBienesNew.BienNoDepreciable
            If fnDocOrigen = 1 And Mid(fsObjetoCod, 2, 2) <> "13" Then 'Valida sea BND
                MsgBox "Para activar Bienes como BIEN NO DEPRECIABLE, necesariamente debe ser un Bien No Depreciable", vbInformation, "Aviso"
                feOrden.SetFocus
                Exit Sub
            End If
            limpiarDatosBND
            'txtBNDInventarioCod.Text = DevolverCorrelativoBND(IIf(Mid(fsAreaAgeCodInv, 4, 2) = "", "01", Mid(fsAreaAgeCodInv, 4, 2)))
             txtBNDInventarioCod.Text = DevolverCorrelativoBND(IIf(lsAgeCod = "", "01", lsAgeCod)) '*** PEAC 20140527

            txtBNDNombre.Text = lsDescripcion
            txtBNDAreaAgeCod.Text = fsAreaAgeCodInv
            txtBNDAreaAgeCod_EmiteDatos
        Case LogTipoBienesNew.ActivoCompuesto
            If fnDocOrigen = 1 And Mid(fsObjetoCod, 2, 2) <> "12" Then 'Valida sea AF
                MsgBox "Para activar Bienes como ACTIVOS COMPUESTO, necesariamente debe ser un Activo Fijo", vbInformation, "Aviso"
                feOrden.SetFocus
                Exit Sub
            End If
            LimpiarDatosAC
            '*** PEAC 20140527
            'txtACInventarioCod.Text = DevolverCorrelativoAF(fsObjetoCod, IIf(Mid(fsAreaAgeCodInv, 4, 2) = "", "01", Mid(fsAreaAgeCodInv, 4, 2)))
             txtACInventarioCod.Text = DevolverCorrelativoAF(fsObjetoCod, IIf(lsAgeCod = "", "01", lsAgeCod))
            '*** FIN PEAC
            txtACNombre.Text = lsDescripcion
            txtACAreaAgeCod.Text = fsAreaAgeCodInv
            txtACAreaAgeCod_EmiteDatos
        Case LogTipoBienesNew.MejoraComponente
            'Valida sea MC
            If fnDocOrigen = 1 And Not (Mid(fsObjetoCod, 2, 2) = "12" Or Mid(fsObjetoCod, 2, 2) = "13") Then
                MsgBox "Para activar Bienes como MEJORA DE COMPONENTE, necesariamente debe ser un ACTIVO FIJO o un BIEN NO DEPRECIABLE", vbInformation, "Aviso"
                feOrden.SetFocus
                Exit Sub
            End If
            LimpiarDatosMC
            'Generamos Código Inventario
            If fnDocOrigen = 1 Then
                Select Case Mid(fsObjetoCod, 2, 2)
                    Case "12"
                        'txtMCInventarioCod.Text = DevolverCorrelativoAF(fsObjetoCod, IIf(Mid(fsAreaAgeCodInv, 4, 2) = "", "01", Mid(fsAreaAgeCodInv, 4, 2)))
                         txtMCInventarioCod.Text = DevolverCorrelativoAF(fsObjetoCod, IIf(lsAgeCod = "", "01", lsAgeCod)) '*** PEAC 20140527
                    Case "13"
                        'txtMCInventarioCod.Text = DevolverCorrelativoBND(IIf(Mid(fsAreaAgeCodInv, 4, 2) = "", "01", Mid(fsAreaAgeCodInv, 4, 2)))
                         txtMCInventarioCod.Text = DevolverCorrelativoBND(IIf(lsAgeCod = "", "01", lsAgeCod)) '*** PEAC 20140527
                End Select
            End If
            txtMCNombre.Text = lsDescripcion
            txtMCDepreTribMes.Text = RecuperaMesesDepreciaTributariamente(fnBANCod, Mid(fsAreaAgeCodInv, 4, 2))
        Case LogTipoBienesNew.BienNoActivable
            If MsgBox("Esta seguro de activar el Bien como NO Activable." & Chr(10) & "Tenga en cuenta que no se puede deshacer este proceso" & Chr(10) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
                Exit Sub
            End If
            Set oBien = New DBien
            oBien.dBeginTrans
            bTrans = True
            Call oBien.InsertaBienNoActivable(fnMovNro, fnMovItem, fnDocOrigen, fsObjetoCod)
            oBien.dCommitTrans
            bTrans = False
            Set oBien = Nothing
            Call cmdCargar_Click
            Exit Sub
    End Select
    
    For lnTab = 0 To TabActivaBien.Tabs - 1
        TabActivaBien.TabVisible(lnTab) = False
    Next
    TabActivaBien.TabVisible(lnTpoBienActivar - 1) = True
    TabActivaBien.Tab = lnTpoBienActivar - 1
        
    Height = fnFormTamanioActiva
    CentraForm Me
    Set oBien = Nothing
    Exit Sub
ErrProcesar:
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub ListarDocOrigen()
    cboDocOrigen.Clear
    cboDocOrigen.AddItem "Orden de Compra" & Space(200) & "1"
    'cboDocOrigen.AddItem "Orden de Servicio" & Space(200) & "2"
    'cboDocOrigen.AddItem "Compra Directa" & Space(200) & "3"
End Sub
Private Sub ListarTipoBien()
    cboTpoBien.Clear
    cboTpoBien.AddItem "Activo Fijo" & Space(100) & "1"
    cboTpoBien.AddItem "Bienes No Depreciables" & Space(200) & "2"
    cboTpoBien.AddItem "Activo Compuesto" & Space(200) & "3"
    cboTpoBien.AddItem "Mejora Componente" & Space(200) & "4"
    cboTpoBien.AddItem "Bien No Activable" & Space(200) & "5"
End Sub
'****************************       Activo Fijo         *************************************
Private Sub InicializaControlesAF()
    Dim obj As New DActualizaDatosArea
    Me.txtAFAreaAgeCod.rs = obj.GetAgenciasAreas
    Set obj = Nothing
End Sub
Private Sub txtAFInventarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFFechaConformidad.SetFocus
    End If
End Sub
Private Sub txtAFFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFNombre.SetFocus
    End If
End Sub
Private Sub txtAFNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtAFDepreTribMes.SetFocus
    End If
End Sub
Private Sub txtAFDepreTribMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtAFDepreContAnio.SetFocus
    End If
End Sub
Private Sub txtAFDepreContAnio_GotFocus()
    txtAFDepreContAnio.Text = Round(txtAFDepreContAnio.Text)
End Sub
Private Sub txtAFDepreContAnio_LostFocus()
    txtAFDepreContAnio.Text = Val(txtAFDepreContAnio.Text)
End Sub
Private Sub txtAFDepreContAnio_Change()
    txtAFDepreContMes.Text = Val(txtAFDepreContAnio.Text) * 12
End Sub
Private Sub txtAFDepreContAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtAFDepreContMes.SetFocus
    End If
End Sub
Private Sub txtAFDepreContMes_LostFocus()
    txtAFDepreContMes.Text = Val(txtAFDepreContMes.Text)
End Sub
Private Sub txtAFDepreContMes_Change()
    txtAFDepreContAnio.Text = Val(txtAFDepreContMes.Text) / 12
End Sub
Private Sub txtAFDepreContMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtAFAreaAgeCod.SetFocus
    End If
End Sub
Private Sub txtAFAreaAgeCod_EmiteDatos()
    txtAFAreaAgeNombre.Text = txtAFAreaAgeCod.psDescripcion
    txtAFDepreTribMes.Text = RecuperaMesesDepreciaTributariamente(fnBANCod, Mid(txtAFAreaAgeCod.Text, 4, 2))
End Sub
Private Sub txtAFPersonaCod_EmiteDatos()
    txtAFPersonaNombre.Text = txtAFPersonaCod.psDescripcion
End Sub
Private Sub txtAFAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFPersonaCod.SetFocus
    End If
End Sub
Private Sub txtAFPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAFMarca.SetFocus
    End If
End Sub
Private Sub txtAFMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtAFSerie.SetFocus
    End If
End Sub
Private Sub txtAFSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtAFModelo.SetFocus
    End If
End Sub
Private Sub txtAFModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmdAFActivar.SetFocus
    End If
End Sub
'Private Sub cmdAFActivar_Click()
'    Dim oAF As New DMov
'    Dim oOpe As New DOperacion
'    Dim lnMovNro As Long, lnMovNroAjus As Long
'    Dim lsMovNro As String, lsMovNroAjus As String
'    Dim lsCtaCont As String, lsAgencia As String, lsCtaOpeBS As String, lsTipo As String
'    Dim bTrans As Boolean
'
'    Dim lsInventarioCod As String, lsNombre As String
'    Dim ldFechaIng As Date
'    Dim lnTiempoDepreTributMes As Integer, lnTiempoDepreContabMes As Integer
'    Dim lsAreaCod As String, lsAgeCod As String
'    Dim lsPersonaCod As String
'    Dim lsMarca As String, lsSerie As String, lsModelo As String
'
'    If Not validarGrabarAF Then Exit Sub
'    If fnBANCod = 0 Then
'        MsgBox "El presente Bien No tiene configurado el Tipo de Activo Fijo", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If MsgBox("¿Esta seguro de activar el Bien como ACTIVO FIJO?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'
'    On Error GoTo ErrCmdAFActivar
'    lsInventarioCod = Trim(txtAFInventarioCod.Text)
'    lsNombre = Trim(txtAFNombre.Text)
'    ldFechaIng = CDate(Me.txtAFFechaConformidad.value)
'    lnTiempoDepreTributMes = txtAFDepreTribMes.Text
'    lnTiempoDepreContabMes = txtAFDepreContMes.Text
'    lsAreaCod = Left(txtAFAreaAgeCod.Text, 3)
'    lsAgeCod = IIf(Mid(txtAFAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAFAreaAgeCod.Text, 4, 2))
'    lsPersonaCod = Trim(txtAFPersonaCod.Text)
'    lsMarca = Trim(txtAFMarca.Text)
'    lsSerie = Trim(txtAFSerie.Text)
'    lsModelo = Trim(txtAFModelo.Text)
'
'    oAF.BeginTrans
'    bTrans = True
'
'    Call oAF.InsertaInventarioAF(lsInventarioCod, lsNombre, txtAFAreaAgeNombre.Text, lsMarca, lsModelo, lsInventarioCod, Format(ldFechaIng, "dd/mm/yyyy"), fnMovNro, fsObjetoCod, fnMovItem, fnDocOrigen)
'    lsMovNro = oAF.GeneraMovNro(ldFechaIng, Right(gsCodAge, 2), gsCodUser)
'    oAF.InsertaMov lsMovNro, gnDepAF, "Depre. de ActivoFijo " & Left(txtAFNombre.Text, 275)
'    lnMovNro = oAF.GetnMovNro(lsMovNro)
'    lsMovNroAjus = oAF.GeneraMovNro(ldFechaIng, Right(gsCodAge, 2), gsCodUser, lsMovNro)
'    oAF.InsertaMov lsMovNroAjus, gnDepAjusteAF, "Deprr. Ajustada de ActivoFijo " & Left(txtAFNombre.Text, 272)
'    lnMovNroAjus = oAF.GetnMovNro(lsMovNroAjus)
'
'    'If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaOpeBS = oAF.GetCtaDep(fsObjetoCod)
'    oAF.InsertaMovBSActivoFijoUnico Year(ldFechaIng), lnMovNro, fsObjetoCod, lsInventarioCod, fnMonto, 0, "0", ldFechaIng, lsAreaCod, lsAgeCod, lnTiempoDepreContabMes, 0, lsNombre, "1", "0", "1", lsInventarioCod, lsInventarioCod, ldFechaIng, ldFechaIng, fnBANCod, "0", lsPersonaCod, "1", 0, lsMarca, lsModelo, lsSerie, lnTiempoDepreTributMes
'
'    'Depre. Historica
'    oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, fsObjetoCod, lsInventarioCod, lnMovNro, CStr(fnBANCod)
'    'lsCtaCont = oAF.GetOpeCtaCtaOtro(gnDepAF, lsCtaOpeBS, "", False)
'    'If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNro, 1, lsCtaCont, CCur(txtDepHist.Text) * -1
'
'    'If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
'    'If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNro, 2, lsCtaCont, CCur(Me.txtDepHist.Text)
'
'    'Depre. Ajsutada
'    'If Val(Me.txtDepreAcum.Text) <> 0 Then lsCtaCont = oOpe.EmiteOpeCta(gnDepAF, "D")
'    oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, fsObjetoCod, lsInventarioCod, lnMovNroAjus, CStr(fnBANCod)
'    'If Val(Me.txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNroAjus, 1, lsCtaCont, CCur(txtDepreAcum.Text) * -1
'
'    'lsCtaCont = oAF.GetOpeCtaCta(gnDepAF, lsCtaOpeBS, "")
'    'lsCtaCont = Left(lsCtaCont, 2) & "6" & Mid(lsCtaCont, 4, 100)
'    'If Val(txtDepreAcum.Text) <> 0 Then oAF.InsertaMovCta lnMovNroAjus, 2, lsCtaCont, CCur(txtDepreAcum.Text)
'
'    If lsPersonaCod <> "" Then
'        oAF.InsertaMovGasto lnMovNro, lsPersonaCod, ""
'        oAF.InsertaMovGasto lnMovNroAjus, lsPersonaCod, ""
'    End If
'
'    oAF.CommitTrans
'    bTrans = False
'
'    MsgBox "Se ha activado el Bien como ACTIVO FIJO con éxito", vbInformation, "Aviso"
'    LimpiarDatosAF
'    cmdCargar_Click
'    Set oAF = Nothing
'    Exit Sub
'ErrCmdAFActivar:
'    MsgBox Err.Description, vbCritical, "Aviso"
'    If bTrans Then
'        oAF.RollbackTrans
'        Set oAF = Nothing
'    End If
'End Sub
Private Sub cmdAFActivar_Click()
    Dim oAF As New DMov
    Dim oBien As New DBien
    Dim lsInventarioCod As String, lsNombre As String
    Dim ldFechaIng As Date
    Dim lnTiempoDepreTributMes As Integer, lnTiempoDepreContabMes As Integer
    Dim lsAreaCod As String, lsAgeCod As String
    Dim lsPersonaCod As String
    Dim lsMarca As String, lsSerie As String, lsModelo As String
    Dim lbRegistraAF As Boolean
    Dim bTrans As Boolean, bTransBien As Boolean
    Dim lnMovNroAF As Long
    
    If Not validarGrabarAF Then Exit Sub
    If fnBANCod = 0 Then
        MsgBox "El presente Bien No tiene configurado el Tipo de Activo Fijo", vbCritical, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Esta seguro de activar el Bien como ACTIVO FIJO?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrCmdAFActivar
    lsInventarioCod = Trim(txtAFInventarioCod.Text)
    lsNombre = Trim(txtAFNombre.Text)
    ldFechaIng = CDate(Me.txtAFFechaConformidad.value)
    lnTiempoDepreTributMes = txtAFDepreTribMes.Text
    lnTiempoDepreContabMes = txtAFDepreContMes.Text
    lsAreaCod = Left(txtAFAreaAgeCod.Text, 3)
    lsAgeCod = IIf(Mid(txtAFAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAFAreaAgeCod.Text, 4, 2))
    lsPersonaCod = Trim(txtAFPersonaCod.Text)
    lsMarca = Trim(txtAFMarca.Text)
    lsSerie = Trim(txtAFSerie.Text)
    lsModelo = Trim(txtAFModelo.Text)

    oAF.BeginTrans
    oBien.dBeginTrans
    bTrans = True
    bTransBien = True
    Call RegistraActivoFijo(fnMovNro, fnMovItem, fnDocOrigen, lsInventarioCod, fsObjetoCod, lsNombre, ldFechaIng, lnTiempoDepreTributMes, _
                                    lnTiempoDepreContabMes, lsAreaCod, lsAgeCod, txtAFAreaAgeNombre.Text, lsPersonaCod, lsMarca, lsSerie, lsModelo, _
                                    fnMonto, fnBANCod, oAF, lnMovNroAF, , fnMoneda)
    Call oBien.InsertaAF(fnMovNro, fnMovItem, lnMovNroAF)
    oAF.CommitTrans
    oBien.dCommitTrans
    bTrans = False
    bTransBien = False
    MsgBox "Se ha activado el Bien como ACTIVO FIJO con éxito", vbInformation, "Aviso"
        'ARLO 20160126 ***
        gsOpeCod = LogPistaEntradaSalidaBien
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", "Se ha activado el Bien como ACTIVO FIJO con éxito N° Serie : " & lsInventarioCod & " Nombre : " & lsNombre
        Set objPista = Nothing
        '**************
    limpiarDatosAF
    cmdCargar_Click
    Set oAF = Nothing
    Exit Sub
ErrCmdAFActivar:
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTrans Then
        oAF.RollbackTrans
        Set oAF = Nothing
    End If
    If bTransBien Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
End Sub
Private Sub cmdAFCancelar_Click()
    Height = fnFormTamanioIni
End Sub
Private Function validarGrabarAF() As Boolean
    Dim valFecha As String
    validarGrabarAF = True
    If Len(Trim(txtAFInventarioCod.Text)) = 0 Then
        MsgBox "El Bien a Activar no tiene Código de Inventario", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFInventarioCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFNombre.Text)) = 0 Then
        MsgBox "Ud. debe de ingresar el Nombre del Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFNombre.SetFocus
        Exit Function
    End If
    valFecha = ValidaFecha(txtAFFechaConformidad.value)
    If valFecha <> "" Then
        MsgBox valFecha, vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFFechaConformidad.SetFocus
        Exit Function
    End If
    If Val(txtAFDepreTribMes.Text) = 0 Then
        MsgBox "Ud. debe ingresar el tiempo en que se deprecia Tributariamente el Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFDepreTribMes.SetFocus
        Exit Function
    End If
    If Val(txtAFDepreContMes.Text) = 0 Then
        MsgBox "Ud. debe ingresar el tiempo en que se deprecia Contablemente el Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFDepreContMes.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFAreaAgeCod.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Área que está asignado el Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFAreaAgeCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFPersonaCod.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Persona al que está asignado el Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFPersonaCod.SetFocus
        Exit Function
    Else
        oUser.DatosPers (txtAFPersonaCod.Text)
        If oUser.AreaCod <> Left(txtAFAreaAgeCod.Text, 3) Then
            MsgBox "La Persona ingresada debe pertenecer al Área seleccionada", vbInformation, "Aviso"
            validarGrabarAF = False
            txtAFPersonaCod.SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(txtAFMarca.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Marca del Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFMarca.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFSerie.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Serie del Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtAFModelo.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Modelo del Bien", vbInformation, "Aviso"
        validarGrabarAF = False
        txtAFModelo.SetFocus
        Exit Function
    End If
End Function
Private Sub limpiarDatosAF()
    txtAFInventarioCod.Text = ""
    txtAFNombre.Text = ""
    txtAFFechaConformidad.value = Format(gdFecSis, "dd/mm/yyyy")
    txtAFDepreTribMes.Text = 0
    txtAFDepreContMes.Text = 0
    txtAFDepreContAnio.Text = 0
    txtAFAreaAgeCod.Text = ""
    txtAFAreaAgeNombre.Text = ""
    txtAFPersonaCod.Text = ""
    txtAFPersonaNombre.Text = ""
    txtAFMarca.Text = ""
    txtAFSerie.Text = ""
    txtAFModelo.Text = ""
    Call InicializaControlesAF
End Sub
Private Function DevolverCorrelativoAF(ByVal psBSCod As String, ByVal psAgeCod As String) As String
    Dim oBien As New DBien
    DevolverCorrelativoAF = oBien.RecuperaCorrelativoAF(psBSCod, psAgeCod)
    Set oBien = Nothing
End Function
'******************************     End Activo Fijo     *************************************
'******************************             BND         *************************************
Private Sub InicializaControlesBND()
    Dim obj As New DActualizaDatosArea
    Me.txtBNDAreaAgeCod.rs = obj.GetAgenciasAreas
    Set obj = Nothing
End Sub
Private Sub txtBNDInventarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDFechaConformidad.SetFocus
    End If
End Sub
Private Sub txtBNDFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDNombre.SetFocus
    End If
End Sub
Private Sub txtBNDNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtBNDAreaAgeCod.SetFocus
    End If
End Sub
Private Sub txtBNDAreaAgeCod_EmiteDatos()
    txtBNDAreaAgeNombre.Text = txtBNDAreaAgeCod.psDescripcion
End Sub
Private Sub txtBNDPersonaCod_EmiteDatos()
    txtBNDPersonaNombre.Text = txtBNDPersonaCod.psDescripcion
End Sub
Private Sub txtBNDAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDPersonaCod.SetFocus
    End If
End Sub
Private Sub txtBNDPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBNDMarca.SetFocus
    End If
End Sub
Private Sub txtBNDMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtBNDSerie.SetFocus
    End If
End Sub
Private Sub txtBNDSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtBNDModelo.SetFocus
    End If
End Sub
Private Sub txtBNDModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmdBNDActivar.SetFocus
    End If
End Sub
Private Sub cmdBNDActivar_Click()
    Dim oAF As New DMov
    Dim oBien As New DBien
    Dim bTrans As Boolean, bTransMov As Boolean
    Dim lsInventarioCod As String, lsNombre As String
    Dim ldFechaIng As Date
    Dim lsAreaCod As String, lsAgeCod As String
    Dim lsPersonaCod As String
    Dim lsMarca As String, lsSerie As String, lsModelo As String
    Dim lnMovNroAF As Long
    
    If Not validarGrabarBND Then Exit Sub
    If MsgBox("¿Esta seguro de activar el Bien como BIEN NO DEPRECIABLE?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrCmdBNDActivar
    lsInventarioCod = Trim(txtBNDInventarioCod.Text)
    lsNombre = Trim(txtBNDNombre.Text)
    ldFechaIng = CDate(txtBNDFechaConformidad.value)
    lsAreaCod = Left(txtBNDAreaAgeCod.Text, 3)
    lsAgeCod = IIf(Mid(txtBNDAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtBNDAreaAgeCod.Text, 4, 2))
    lsPersonaCod = Trim(txtBNDPersonaCod.Text)
    lsMarca = Trim(txtBNDMarca.Text)
    lsSerie = Trim(txtBNDSerie.Text)
    lsModelo = Trim(txtBNDModelo.Text)

    oAF.BeginTrans
    oBien.dBeginTrans
    bTrans = True
    bTransMov = True
    Call RegistraActivoFijo(fnMovNro, fnMovItem, fnDocOrigen, lsInventarioCod, fsObjetoCod, lsNombre, ldFechaIng, 0, _
                                0, lsAreaCod, lsAgeCod, txtBNDAreaAgeNombre.Text, lsPersonaCod, lsMarca, lsSerie, lsModelo, _
                                fnMonto, fnBANCod, oAF, lnMovNroAF, "0", fnMoneda)
    Call oBien.InsertaBienNoDepreciable(fnMovNro, fnMovItem, fnDocOrigen, lsInventarioCod, fsObjetoCod, lsNombre, fnMonto, ldFechaIng, lsAreaCod, lsAgeCod, lsMarca, lsSerie, lsModelo, lnMovNroAF)
    oBien.dCommitTrans
    oAF.CommitTrans
    bTrans = False
    bTransMov = False
    
    MsgBox "Se ha activado el Bien como BIEN NO DEPRECIABLE con éxito", vbInformation, "Aviso"
    limpiarDatosBND
    cmdCargar_Click
    Set oBien = Nothing
    Exit Sub
ErrCmdBNDActivar:
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    If bTransMov Then
        oAF.RollbackTrans
        Set oAF = Nothing
    End If
End Sub
Private Sub cmdBNDCancelar_Click()
    Height = fnFormTamanioIni
End Sub
Private Function validarGrabarBND() As Boolean
    Dim valFecha As String
    validarGrabarBND = True
    If Len(Trim(txtBNDInventarioCod.Text)) = 0 Then
        MsgBox "El Bien a Activar no tiene Código de Inventario", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDInventarioCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDNombre.Text)) = 0 Then
        MsgBox "Ud. debe de ingresar el Nombre del Bien", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDNombre.SetFocus
        Exit Function
    End If
    valFecha = ValidaFecha(txtBNDFechaConformidad.value)
    If valFecha <> "" Then
        MsgBox valFecha, vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDFechaConformidad.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDAreaAgeCod.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Área que está asignado el Bien", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDAreaAgeCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDPersonaCod.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Persona al que está asignado el Bien", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDPersonaCod.SetFocus
        Exit Function
    Else
        oUser.DatosPers (txtBNDPersonaCod.Text)
        If oUser.AreaCod <> Left(txtBNDAreaAgeCod.Text, 3) Then
            MsgBox "La Persona ingresada debe pertenecer al Área seleccionada", vbInformation, "Aviso"
            validarGrabarBND = False
            txtBNDPersonaCod.SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(txtBNDMarca.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Marca del Bien", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDMarca.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDSerie.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Serie del Bien", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtBNDModelo.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Modelo del Bien", vbInformation, "Aviso"
        validarGrabarBND = False
        txtBNDModelo.SetFocus
        Exit Function
    End If
End Function
Private Sub limpiarDatosBND()
    txtBNDInventarioCod.Text = ""
    txtBNDNombre.Text = ""
    txtBNDFechaConformidad.value = Format(gdFecSis, "dd/mm/yyyy")
    txtBNDAreaAgeCod.Text = ""
    txtBNDAreaAgeNombre.Text = ""
    txtBNDPersonaCod.Text = ""
    txtBNDPersonaNombre.Text = ""
    txtBNDMarca.Text = ""
    txtBNDSerie.Text = ""
    txtBNDModelo.Text = ""
    Call InicializaControlesBND
End Sub
Private Function DevolverCorrelativoBND(ByVal psAgeCod As String) As String
    Dim oBien As New DBien
    DevolverCorrelativoBND = oBien.RecuperaCorrelativoBND(psAgeCod)
    Set oBien = Nothing
End Function
'************************************      End BND    ***************************************
'******************************             Activo Compuesto         *************************************
Private Sub InicializaControlesAC()
    Dim obj As New DActualizaDatosArea
    Me.txtACAreaAgeCod.rs = obj.GetAgenciasAreas
    Set obj = Nothing
End Sub
Private Sub txtACInventarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtACFechaConformidad.SetFocus
    End If
End Sub
Private Sub txtACFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtACNombre.SetFocus
    End If
End Sub
Private Sub txtACNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtACAreaAgeCod.SetFocus
    End If
End Sub
Private Sub txtACAreaAgeCod_EmiteDatos()
    txtACAreaAgeNombre.Text = txtACAreaAgeCod.psDescripcion
    If txtACAreaAgeCod.Text <> fsAreaAgeCodInv Then
        Set fMatComponente = Nothing
        ReDim fMatComponente(10, 0)
        MsgBox "Ahora ya puede registrar los componentes respectivos", vbInformation, "Aviso"
    End If
End Sub
Private Sub txtACPersonaCod_EmiteDatos()
    txtACPersonaNombre.Text = txtACPersonaCod.psDescripcion
End Sub
Private Sub txtACAreaAgeCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtACPersonaCod.SetFocus
    End If
End Sub
Private Sub txtACPersonaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtACMarca.SetFocus
    End If
End Sub
Private Sub txtACMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtACSerie.SetFocus
    End If
End Sub
Private Sub txtACSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtACModelo.SetFocus
    End If
End Sub
Private Sub txtACModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmdACComponentes.SetFocus
    End If
End Sub
Private Sub cmdACComponentes_Click()
    Dim lsCorreBND As String
    Dim lsAreaAgeCod As String
    lsAreaAgeCod = txtACAreaAgeCod.Text
    If lsAreaAgeCod = "" Then
        MsgBox "Ud.debe primero de seleccionar la Área/Agencia", vbInformation, "Aviso"
        txtACAreaAgeCod.SetFocus
        Exit Sub
    End If
    lsCorreBND = DevolverCorrelativoBND(IIf(Mid(fsAreaAgeCodInv, 4, 2) = "", "01", Mid(fsAreaAgeCodInv, 4, 2)))
    lsCorreBND = Left(lsCorreBND, Len(lsCorreBND) - 5) & Format(CLng(Right(lsCorreBND, 5)) - 1, "00000")
    fMatComponente = frmInvActivarBienComponente.Inicio(fMatComponente, txtACInventarioCod.Text, lsCorreBND, lsAreaAgeCod)
End Sub
Private Sub cmdACActivar_Click()
    Dim oLog As DLogBieSer
    Dim oBien As New DBien
    Dim oAF As New DMov
    Dim bTrans As Boolean, bTransMov As Boolean
    Dim lsInventarioCod As String, lsNombre As String
    Dim ldFechaIng As Date
    Dim lsAreaCod As String, lsAgeCod As String
    Dim lsPersonaCod As String
    Dim lsMarca As String, lsSerie As String, lsModelo As String
    Dim i As Integer
    Dim lnIDCabecera As Long
    Dim lnMontoDetalle As Currency
    Dim lnBANCodDet As Integer
    Dim MatDetalle As Variant
    Dim lnMovNroAF As Long
    
    If Not validarGrabarAC Then Exit Sub
    'Valida Monto de Activación
    For i = 1 To UBound(fMatComponente, 2)
        lnMontoDetalle = lnMontoDetalle + fMatComponente(10, i)
    Next
    If lnMontoDetalle <> fnMonto Then
        MsgBox "El monto del Objeto que se esta activando " & Format(fnMonto, gsFormatoNumeroView) & Chr(10) & "es diferente al total de los componentes registrados " & Format(lnMontoDetalle, gsFormatoNumeroView), vbInformation, "Aviso"
        Exit Sub
    End If
    ReDim MatDetalle(1, 0)
    For i = 1 To UBound(fMatComponente, 2)
        ReDim Preserve MatDetalle(1, i)
        MatDetalle(1, i) = 0
        If fMatComponente(1, i) = 1 Then 'Si es AF
            lnBANCodDet = oBien.RecuperaCodigoBAN(fMatComponente(2, i))
            If lnBANCodDet = 0 Then
                MsgBox "El componente registrado " & UCase(CStr(fMatComponente(3, i))) & " No tiene configurado el Tipo de Activo Fijo", vbCritical, "Aviso"
                Exit Sub
            End If
            MatDetalle(1, i) = lnBANCodDet
        End If
    Next
    
    If MsgBox("¿Esta seguro de activar el Bien como ACTIVO COMPUESTO?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrCmdACActivar
    lsInventarioCod = Trim(txtACInventarioCod.Text)
    lsNombre = Trim(txtACNombre.Text)
    ldFechaIng = CDate(txtACFechaConformidad.value)
    lsAreaCod = Left(txtACAreaAgeCod.Text, 3)
    lsAgeCod = IIf(Mid(txtACAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtACAreaAgeCod.Text, 4, 2))
    lsPersonaCod = Trim(txtACPersonaCod.Text)
    lsMarca = Trim(txtACMarca.Text)
    lsSerie = Trim(txtACSerie.Text)
    lsModelo = Trim(txtACModelo.Text)
    
    Set oBien = New DBien
    Set oAF = New DMov

    oBien.dBeginTrans
    oAF.BeginTrans
    bTrans = True
    bTransMov = True

    lnIDCabecera = oBien.InsertaBienActivoCompuesto(fnMovNro, fnMovItem, fnDocOrigen, lsInventarioCod, fsObjetoCod, lsNombre, fnMonto, _
                                                    ldFechaIng, lsAreaCod, lsAgeCod, lsPersonaCod, lsMarca, lsSerie, lsModelo)
    
    For i = 1 To UBound(fMatComponente, 2)
        lnMovNroAF = 0
        '1. Tipo,2. Objeto,3. Nombre Objeto,4. Cod. Inventario,5. Depreciacion Contable (meses),6. Depreciacion Tributaria (meses),'7. Marca,8. Modelo,9. Serie,10. Precio
        If fMatComponente(1, i) = 1 Then 'Si es Activo Fijo
            Call RegistraActivoFijo(fnMovNro, fnMovItem, fnDocOrigen, fMatComponente(4, i), fMatComponente(2, i), fMatComponente(3, i), ldFechaIng, fMatComponente(6, i), fMatComponente(5, i), lsAreaCod, lsAgeCod, "", lsPersonaCod, _
                                    fMatComponente(7, i), fMatComponente(9, i), fMatComponente(8, i), fMatComponente(10, i), MatDetalle(1, i), oAF, lnMovNroAF, , fnMoneda)
        Else 'BND
            Call RegistraActivoFijo(fnMovNro, fnMovItem, fnDocOrigen, fMatComponente(4, i), fMatComponente(2, i), fMatComponente(3, i), ldFechaIng, 0, 0, lsAreaCod, lsAgeCod, "", lsPersonaCod, _
                                    fMatComponente(7, i), fMatComponente(9, i), fMatComponente(8, i), fMatComponente(10, i), MatDetalle(1, i), oAF, lnMovNroAF, "0", fnMoneda)
        End If
        Call oBien.InsertaBienActivoCompuestoDet(lnIDCabecera, i, fnDocOrigen, fMatComponente(1, i), fMatComponente(4, i), _
            fMatComponente(2, i), fMatComponente(3, i), fMatComponente(10, i), ldFechaIng, fMatComponente(6, i), fMatComponente(5, i), _
            lsAreaCod, lsAgeCod, lsPersonaCod, fMatComponente(7, i), fMatComponente(9, i), fMatComponente(8, i), lnMovNroAF)
    Next

    oBien.dCommitTrans
    oAF.CommitTrans
    bTrans = False
    bTransMov = False
    
    MsgBox "Se ha activado el Bien como ACTIVO COMPUESTO con éxito", vbInformation, "Aviso"
    LimpiarDatosAC
    cmdCargar_Click
    Set oBien = Nothing
    Exit Sub
ErrCmdACActivar:
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    If bTransMov Then
        oAF.RollbackTrans
        Set oAF = Nothing
    End If
End Sub
Private Sub cmdACCancelar_Click()
    Height = fnFormTamanioIni
End Sub
Private Function validarGrabarAC() As Boolean
    Dim valFecha As String
    validarGrabarAC = True
    If Len(Trim(txtACInventarioCod.Text)) = 0 Then
        MsgBox "El Bien a Activar no tiene Código de Inventario", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACInventarioCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtACNombre.Text)) = 0 Then
        MsgBox "Ud. debe de ingresar el Nombre del Bien", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACNombre.SetFocus
        Exit Function
    End If
    valFecha = ValidaFecha(txtACFechaConformidad.value)
    If valFecha <> "" Then
        MsgBox valFecha, vbInformation, "Aviso"
        validarGrabarAC = False
        txtACFechaConformidad.SetFocus
        Exit Function
    End If
    If Len(Trim(txtACAreaAgeCod.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Área que está asignado el Bien", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACAreaAgeCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtACPersonaCod.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Persona al que está asignado el Bien", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACPersonaCod.SetFocus
        Exit Function
    Else
        oUser.DatosPers (txtACPersonaCod.Text)
        If oUser.AreaCod <> Left(txtACAreaAgeCod.Text, 3) Then
            MsgBox "La Persona ingresada debe pertenecer al Área seleccionada", vbInformation, "Aviso"
            validarGrabarAC = False
            txtACPersonaCod.SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(txtACMarca.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Marca del Bien", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACMarca.SetFocus
        Exit Function
    End If
    If Len(Trim(txtACSerie.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Serie del Bien", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtACModelo.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Modelo del Bien", vbInformation, "Aviso"
        validarGrabarAC = False
        txtACModelo.SetFocus
        Exit Function
    End If
    If UBound(fMatComponente, 2) = 0 Then
        MsgBox "Ud. aun no ha realizado el Registro de Componentes", vbInformation, "Aviso"
        validarGrabarAC = False
        cmdACComponentes.SetFocus
        Exit Function
    End If
End Function
Private Sub LimpiarDatosAC()
    txtACInventarioCod.Text = ""
    txtACNombre.Text = ""
    txtACFechaConformidad.value = Format(gdFecSis, "dd/mm/yyyy")
    txtACAreaAgeCod.Text = ""
    txtACAreaAgeNombre.Text = ""
    txtACPersonaCod.Text = ""
    txtACPersonaNombre.Text = ""
    txtACMarca.Text = ""
    txtACSerie.Text = ""
    txtACModelo.Text = ""
    Set fMatComponente = Nothing
    ReDim fMatComponente(10, 0)
    Call InicializaControlesAC
End Sub
'Private Function DevolverCorrelativoAC(ByVal psAgeCod As String) As String
'    Dim oBien As New DBien
'    DevolverCorrelativoAC = oBien.RecuperaCorrelativoAC(psAgeCod)
'    Set oBien = Nothing
'End Function
'************************************      End Activo Compuesto    ***************************************
'****************************       Mejora Componente         *************************************
Private Sub InicializaControlesMC()
    Dim obj As New DBien
    txtMCActCompMejoraCod.rs = obj.RecuperaMejoraComponentePaArbol()
    Set obj = Nothing
End Sub
Private Sub txtMCInventarioCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMCFechaConformidad.SetFocus
    End If
End Sub
Private Sub txtMCFechaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMCNombre.SetFocus
    End If
End Sub
Private Sub txtMCNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtMCDepreTribMes.SetFocus
    End If
End Sub
Private Sub txtMCDepreTribMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtMCDepreContAnio.SetFocus
    End If
End Sub
Private Sub txtMCDepreContAnio_GotFocus()
    txtMCDepreContAnio.Text = Round(txtMCDepreContAnio.Text)
End Sub
Private Sub txtMCDepreContAnio_LostFocus()
    txtMCDepreContAnio.Text = Val(txtMCDepreContAnio.Text)
End Sub
Private Sub txtMCDepreContAnio_Change()
    txtMCDepreContMes.Text = Val(txtMCDepreContAnio.Text) * 12
End Sub
Private Sub txtMCDepreContAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        txtMCDepreContMes.SetFocus
    End If
End Sub
Private Sub txtMCDepreContMes_LostFocus()
    txtMCDepreContMes.Text = Val(txtMCDepreContMes.Text)
End Sub
Private Sub txtMCDepreContMes_Change()
    txtMCDepreContAnio.Text = Val(txtMCDepreContMes.Text) / 12
End Sub
Private Sub txtMCDepreContMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        If txtMCActCompMejoraCod.Enabled Then
            txtMCActCompMejoraCod.SetFocus
        Else
            txtMCMarca.SetFocus
        End If
    End If
End Sub
Private Sub txtMCActCompMejoraCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMCMarca.SetFocus
    End If
End Sub
Private Sub txtMCActCompMejoraCod_EmiteDatos()
    txtMCActCompMejoraNombre.Text = txtMCActCompMejoraCod.psDescripcion
End Sub
Private Sub txtMCMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtMCSerie.SetFocus
    End If
End Sub
Private Sub txtMCSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtMCModelo.SetFocus
    End If
End Sub
Private Sub txtMCModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmdMCActivar.SetFocus
    End If
End Sub
Private Sub cmdMCActivar_Click()
    Dim oAF As New DMov
    Dim oBien As New DBien
    Dim rs As New ADODB.Recordset
    Dim bTrans As Boolean, bTransMov As Boolean
    Dim lsInventarioCod As String, lsNombre As String
    Dim ldFechaIng As Date
    Dim lsAreaCod As String, lsAgeCod As String
    Dim lsPersonaCod As String
    Dim lsMarca As String, lsSerie As String, lsModelo As String
    Dim lnDepreciaTribMes As Integer, lnDepreciaContMes As Integer
    Dim lnId As Long, lnItem As Integer
    Dim i As Integer
    Dim lnMovNroAF As Long
    
    If Not validarGrabarMC Then Exit Sub
    If fnDocOrigen = 1 Then
        If Mid(fsObjetoCod, 2, 2) = "12" And fnBANCod = 0 Then
            MsgBox "El presente Bien No tiene configurado el Tipo de Activo Fijo", vbCritical, "Aviso"
            Exit Sub
        End If
    End If
    
    If MsgBox("¿Esta seguro de activar el Bien como MEJORA DE COMPONENTE?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrCmdMCActivar
    lsInventarioCod = Trim(txtMCInventarioCod.Text)
    lsNombre = Trim(txtMCNombre.Text)
    ldFechaIng = CDate(txtMCFechaConformidad.value)
    lnDepreciaTribMes = CInt(txtMCDepreTribMes.Text)
    lnDepreciaContMes = CInt(txtMCDepreContMes.Text)
    lsMarca = Trim(txtMCMarca.Text)
    lsSerie = Trim(txtMCSerie.Text)
    lsModelo = Trim(txtMCModelo.Text)
    lnId = CLng(txtMCActCompMejoraCod.Text)
    lnItem = 0
    Set rs = oBien.RecuperaBienActivoCompuesto(lnId)
    If Not rs.EOF Then
        lsAreaCod = rs!cAreaCod
        lsAgeCod = rs!cAgeCod
        lsPersonaCod = rs!cPersCod
    End If

    oBien.dBeginTrans
    oAF.BeginTrans
    bTrans = True
    bTransMov = True

    If fnDocOrigen = 1 Then
        If Mid(fsObjetoCod, 2, 2) = "12" Then 'Si es Activo Fijo
            Call RegistraActivoFijo(fnMovNro, fnMovItem, fnDocOrigen, lsInventarioCod, fsObjetoCod, lsNombre, _
                                    ldFechaIng, lnDepreciaTribMes, lnDepreciaContMes, lsAreaCod, lsAgeCod, "", lsPersonaCod, lsMarca, _
                                    lsSerie, lsModelo, fnMonto, fnBANCod, oAF, lnMovNroAF, , fnMoneda)
        End If
    End If
    Call oBien.InsertaBienActivoCompuestoDet(lnId, lnItem, fnDocOrigen, IIf(Mid(fsObjetoCod, 2, 2) = "12", 1, 2), _
                                            lsInventarioCod, fsObjetoCod, lsNombre, fnMonto, ldFechaIng, lnDepreciaTribMes, lnDepreciaContMes, _
                                            lsAreaCod, lsAgeCod, lsPersonaCod, lsMarca, lsSerie, lsModelo, lnMovNroAF)
    Call oBien.InsertaBienMejoraComponente(fnMovNro, fnMovItem, lnId, lnItem, CDate(gdFecSis & " " & Time))
    
    oBien.dCommitTrans
    oAF.CommitTrans
    bTrans = False
    bTransMov = False
    MsgBox "Se ha activado el Bien como MEJORA DE COMPONENTE con éxito", vbInformation, "Aviso"
    LimpiarDatosMC
    cmdCargar_Click
    Set oBien = Nothing
    Exit Sub
ErrCmdMCActivar:
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
    If bTransMov Then
        oAF.RollbackTrans
        Set oAF = Nothing
    End If
End Sub
Private Sub cmdMCCancelar_Click()
    Height = fnFormTamanioIni
End Sub
Private Function validarGrabarMC() As Boolean
    Dim valFecha As String
    validarGrabarMC = True
    If Len(Trim(txtMCInventarioCod.Text)) = 0 Then
        MsgBox "El Bien a Activar no tiene Código de Inventario", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCInventarioCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtMCNombre.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nombre del Bien", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCNombre.SetFocus
        Exit Function
    End If
    valFecha = ValidaFecha(txtMCFechaConformidad.value)
    If valFecha <> "" Then
        MsgBox valFecha, vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCFechaConformidad.SetFocus
        Exit Function
    End If
    If Val(txtMCDepreTribMes.Text) = 0 Then
        MsgBox "Ud. debe ingresar el tiempo en que se deprecia Tributariamente el Bien", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCDepreTribMes.SetFocus
        Exit Function
    End If
    If Val(txtMCDepreContMes.Text) = 0 Then
        MsgBox "Ud. debe ingresar el tiempo en que se deprecia Contablemente el Bien", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCDepreContMes.SetFocus
        Exit Function
    End If
    If Len(txtMCActCompMejoraCod.Text) = 0 Then
        MsgBox "Ud. debe ingresar el Activo Compuesto a Mejorar", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCActCompMejoraCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtMCMarca.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Marca del Bien", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCMarca.SetFocus
        Exit Function
    End If
    If Len(Trim(txtMCSerie.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Serie del Bien", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(txtMCModelo.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Modelo del Bien", vbInformation, "Aviso"
        validarGrabarMC = False
        txtMCModelo.SetFocus
        Exit Function
    End If
End Function
Private Sub LimpiarDatosMC()
    txtMCInventarioCod.Text = ""
    txtMCNombre.Text = ""
    txtMCFechaConformidad.value = Format(gdFecSis, "dd/mm/yyyy")
    txtMCDepreTribMes.Text = 0
    txtMCDepreContMes.Text = 0
    txtMCDepreContAnio.Text = 0
    txtMCActCompMejoraCod.Text = ""
    txtMCActCompMejoraNombre.Text = ""
    txtMCMarca.Text = ""
    txtMCSerie.Text = ""
    txtMCModelo.Text = ""
    Call InicializaControlesMC
End Sub
'******************************     End Mejora Componente     *************************************
Private Sub cancela_busqueda_actual()
    Call FormateaFlex(feOrden)
    Height = fnFormTamanioIni
    cboTpoBien.ListIndex = -1
End Sub
Private Function RecuperaMesesDepreciaTributariamente(ByVal pnTpoActivo As Integer, ByRef psAgeCod As String) As Integer
    Dim dLog As New DLogDeprecia
    Dim rs As New ADODB.Recordset
    If psAgeCod = "" Then psAgeCod = "01"
    Set rs = dLog.ObtienePorcentVidaUtlAF(pnTpoActivo)
    RecuperaMesesDepreciaTributariamente = 0
    Do While Not rs.EOF
        If psAgeCod = rs!cAgeCod Then
            RecuperaMesesDepreciaTributariamente = rs!nDepreMesT
            Exit Do
        End If
        rs.MoveNext
    Loop
    Set dLog = Nothing
    Set rs = Nothing
End Function
Public Sub RegistraActivoFijo(ByVal pnMovNroOrden As Long, ByVal pnMovItemOrden As Integer, ByVal pnDocOrigen As Integer, _
                                    ByVal psInventarioCod As String, ByVal psObjetoCod As String, ByVal psNombre As String, ByVal pdFechaIng As Date, _
                                    ByVal pnTiempoDepreTribMes As Integer, ByVal pnTiempoDepreContMes As Integer, ByVal psAreaCod As String, _
                                    ByVal psAgeCod As String, ByVal psAreaAgeNombre As String, ByVal psPersonaCod As String, ByVal psMarca As String, _
                                    ByVal psSerie As String, ByVal psModelo As String, ByVal pnMonto As Currency, ByVal pnBANCod As Integer, ByVal oAF As DMov, _
                                    ByRef pnMovNroAF As Long, Optional ByVal psCategoriaBien As String = "1", Optional ByVal pnMoneda As Moneda = gMonedaNacional)
    Dim lnMovNro As Long, lnMovNroAjus As Long
    Dim lsMovNro As String, lsMovNroAjus As String
    Dim lsCtaCont As String, lsCtaOpeBS As String

    Call oAF.InsertaInventarioAF(psInventarioCod, psNombre, psAreaAgeNombre, psMarca, psModelo, psInventarioCod, Format(pdFechaIng, "dd/mm/yyyy"), pnMovNroOrden, psObjetoCod, pnMovItemOrden, pnDocOrigen)
    lsMovNro = oAF.GeneraMovNro(pdFechaIng, Right(gsCodAge, 2), gsCodUser)
    oAF.InsertaMov lsMovNro, gnDepAF, "Depre. de ActivoFijo " & Left(psNombre, 275)
    lnMovNro = oAF.GetnMovNro(lsMovNro)
    lsMovNroAjus = oAF.GeneraMovNro(pdFechaIng, Right(gsCodAge, 2), gsCodUser, lsMovNro)
    oAF.InsertaMov lsMovNroAjus, gnDepAjusteAF, "Deprr. Ajustada de ActivoFijo " & Left(psNombre, 272)
    lnMovNroAjus = oAF.GetnMovNro(lsMovNroAjus)

    oAF.InsertaMovBSActivoFijoUnico Year(pdFechaIng), lnMovNro, psObjetoCod, psInventarioCod, pnMonto, 0, "0", pdFechaIng, psAreaCod, psAgeCod, pnTiempoDepreContMes, 0, psNombre, "1", "0", "1", psInventarioCod, psInventarioCod, pdFechaIng, pdFechaIng, pnBANCod, "0", psPersonaCod, psCategoriaBien, 0, psMarca, psModelo, psSerie, pnTiempoDepreTribMes, pnMoneda
    'Depre. Historica
    oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, psObjetoCod, psInventarioCod, lnMovNro, CStr(pnBANCod)
    'Depre. Ajsutada
    oAF.InsertaMovBSAF Year(gdFecSis), lnMovNro, 1, psObjetoCod, psInventarioCod, lnMovNroAjus, CStr(pnBANCod)
    If psPersonaCod <> "" Then
        oAF.InsertaMovGasto lnMovNro, psPersonaCod, ""
        oAF.InsertaMovGasto lnMovNroAjus, psPersonaCod, ""
    End If
    pnMovNroAF = lnMovNro
End Sub
