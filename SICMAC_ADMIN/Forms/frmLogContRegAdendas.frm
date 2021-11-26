VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmLogContRegAdendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Registro de Adendas"
   ClientHeight    =   7800
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   8280
   Icon            =   "frmLogContRegAdendas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkArchivo 
      Height          =   255
      Left            =   1680
      TabIndex        =   69
      Top             =   6720
      Width           =   255
   End
   Begin VB.Frame fraArchivoAdenda 
      Caption         =   "Archivo Adenda"
      Enabled         =   0   'False
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
      Height          =   1005
      Left            =   120
      TabIndex        =   65
      Top             =   6720
      Width           =   8040
      Begin VB.CommandButton cmdBuscarArchivo 
         Caption         =   "E&xaminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   66
         ToolTipText     =   "Buscar Credito"
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Contrato Digital"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblNombreArchivo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1320
         TabIndex        =   67
         Tag             =   "txtnombre"
         Top             =   360
         Width           =   4935
      End
   End
   Begin TabDlg.SSTab SSTContratos 
      Height          =   6420
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11324
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Complementarias"
      TabPicture(0)   =   "frmLogContRegAdendas.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraContratoCOMP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAdendaCOMP"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraCronograma"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGlosa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdRegistrarCOMP"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancelarCOMP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Adicionales"
      TabPicture(1)   =   "frmLogContRegAdendas.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContratoAD"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraAdendaAD"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRegistrarAD"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdCancelarAD"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Reducciones"
      TabPicture(2)   =   "frmLogContRegAdendas.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraContratoRED"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraAdendaRED"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdRegistrarRED"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdCancelarRED"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Items del Contrato"
      TabPicture(3)   =   "frmLogContRegAdendas.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraItemContrato"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraItemContrato 
         Caption         =   "Items Relacionados al contrato"
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
         Height          =   5775
         Left            =   -74880
         TabIndex        =   70
         Top             =   480
         Width           =   7935
         Begin VB.CommandButton cmdAgregarItemCont 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   5280
            Width           =   855
         End
         Begin VB.CommandButton cmdQuitarItemCont 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   960
            TabIndex        =   71
            Top             =   5280
            Width           =   975
         End
         Begin Sicmact.FlexEdit feOrden 
            Height          =   3255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   5741
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Ag.Des.-Objeto-Descripcion-Solic.-P.Unitario-SubTotal-CtaContCod"
            EncabezadosAnchos=   "0-800-900-3000-700-1100-1100-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-1-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-R-R-R-L"
            FormatosEdit    =   "0-0-0-0-3-2-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit feObj 
            Height          =   615
            Left            =   9240
            TabIndex        =   74
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   1085
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Id-Objeto Orden-CtaContCod-CtaContDesc-Filtro-CodObjeto"
            EncabezadosAnchos=   "0-400-800-800-800-800-800"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.CommandButton cmdCancelarRED 
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
         Height          =   360
         Left            =   -68160
         TabIndex        =   60
         Top             =   5880
         Width           =   1110
      End
      Begin VB.CommandButton cmdRegistrarRED 
         Caption         =   "&Registrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69280
         TabIndex        =   59
         Top             =   5880
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancelarAD 
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
         Height          =   360
         Left            =   -68160
         TabIndex        =   58
         Top             =   5880
         Width           =   1110
      End
      Begin VB.CommandButton cmdRegistrarAD 
         Caption         =   "&Registrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69280
         TabIndex        =   57
         Top             =   5880
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancelarCOMP 
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
         Height          =   360
         Left            =   6910
         TabIndex        =   56
         Top             =   5760
         Width           =   1110
      End
      Begin VB.CommandButton cmdRegistrarCOMP 
         Caption         =   "&Registrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         TabIndex        =   55
         Top             =   5760
         Width           =   1110
      End
      Begin VB.Frame fraAdendaRED 
         Caption         =   "Adenda"
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
         Height          =   3900
         Left            =   -74880
         TabIndex        =   46
         Top             =   1560
         Width           =   7800
         Begin VB.ComboBox cboCuotaHastaRED 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1680
            Width           =   1020
         End
         Begin VB.ComboBox cboCuotaDesdeRED 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1200
            Width           =   1020
         End
         Begin VB.TextBox txtMontoRED 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   48
            Top             =   720
            Width           =   1500
         End
         Begin VB.TextBox txtRazonRED 
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
            Height          =   1530
            Left            =   1200
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   2160
            Width           =   6060
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Hasta Cuota:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1680
            Width           =   930
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Desde Cuota:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblNAdendaRED 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1200
            TabIndex        =   52
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nº Adenda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Reducción:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Razon:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   49
            Top             =   2160
            Width           =   510
         End
      End
      Begin VB.Frame fraAdendaAD 
         Caption         =   "Adenda"
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
         Height          =   3900
         Left            =   -74880
         TabIndex        =   37
         Top             =   1560
         Width           =   7800
         Begin VB.ComboBox cboCuotaHastaAD 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1680
            Width           =   1020
         End
         Begin VB.ComboBox cboCuotaDesdeAD 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   1200
            Width           =   1020
         End
         Begin VB.TextBox txtRazonAD 
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
            Height          =   1530
            Left            =   1200
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   2160
            Width           =   6060
         End
         Begin VB.TextBox txtMontoExtra 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   43
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Razon:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   44
            Top             =   2160
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extra:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nº Adenda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   270
            Width           =   825
         End
         Begin VB.Label lblNAdendaAD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1200
            TabIndex        =   40
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desde Cuota:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta Cuota:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   930
         End
      End
      Begin VB.Frame fraContratoRED 
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
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   7800
         Begin VB.Label lblProveedorRED 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   36
            Tag             =   "txtnombre"
            Top             =   600
            Width           =   6495
         End
         Begin VB.Label lblNContratoRED 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   35
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   870
         End
      End
      Begin VB.Frame fraContratoAD 
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
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   7800
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   780
         End
         Begin VB.Label lblNContratoAD 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   29
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblProveedorAD 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   28
            Tag             =   "txtnombre"
            Top             =   600
            Width           =   6495
         End
      End
      Begin VB.TextBox txtGlosa 
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
         Height          =   450
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   5760
         Width           =   5460
      End
      Begin VB.Frame fraCronograma 
         Caption         =   "Cronograma"
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
         Height          =   3165
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   7800
         Begin VB.TextBox txtMontoCro 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Left            =   5040
            MaxLength       =   15
            TabIndex        =   17
            Top             =   840
            Width           =   1380
         End
         Begin VB.ComboBox cboMonedaCro 
            Height          =   315
            ItemData        =   "frmLogContRegAdendas.frx":037A
            Left            =   3360
            List            =   "frmLogContRegAdendas.frx":037C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   840
            Width           =   1380
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6600
            TabIndex        =   15
            Top             =   840
            Width           =   1005
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "&Quitar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6600
            TabIndex        =   14
            Top             =   1320
            Width           =   1005
         End
         Begin Sicmact.FlexEdit feCronograma 
            Height          =   1755
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   5640
            _ExtentX        =   9948
            _ExtentY        =   3096
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Nº Pago-Fecha de Pago-Moneda-Monto"
            EncabezadosAnchos=   "500-1000-1200-1000-1200"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ColumnasAEditar =   "X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   7
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin MSComCtl2.DTPicker txtFechaPago 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   840
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   122093569
            CurrentDate     =   37156
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cuota:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2880
            TabIndex        =   23
            Top             =   840
            Width           =   465
         End
         Begin VB.Label lblTpoMonedaCro 
            AutoSize        =   -1  'True
            Caption         =   "S/"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   22
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lblNPago 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   21
            Tag             =   "txtcodigo"
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pago"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   600
            TabIndex        =   20
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Frame fraAdendaCOMP 
         Caption         =   "Adenda"
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
         Height          =   660
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   7800
         Begin MSComCtl2.DTPicker txtFecIniCOMP 
            Height          =   315
            Left            =   3120
            TabIndex        =   7
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   129499137
            CurrentDate     =   37156
         End
         Begin MSComCtl2.DTPicker txtFecFinCOMP 
            Height          =   315
            Left            =   5160
            TabIndex        =   8
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   129499137
            CurrentDate     =   37156
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4560
            TabIndex        =   12
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2520
            TabIndex        =   11
            Top             =   270
            Width           =   510
         End
         Begin VB.Label lblNAdendaCOMP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   10
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Nº Adenda:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame fraContratoCOMP 
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
         ForeColor       =   &H8000000D&
         Height          =   1005
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7800
         Begin VB.Label lblProveedorCOMP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Tag             =   "txtnombre"
            Top             =   600
            Width           =   6495
         End
         Begin VB.Label lblNContratoCOMP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   5
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   270
            Width           =   870
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   5520
         Width           =   450
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   615
      Left            =   7920
      TabIndex        =   75
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1085
      Filtro          =   ""
      Altura          =   0
   End
End
Attribute VB_Name = "frmLogContRegAdendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsRuta As String
Dim fsPathFile As String
Dim fsNContrato As String
Dim fnNAdenda2 As Integer
Dim fncontRef As Integer 'PASI20140823 TI-ERS077-2014
Dim fnTipo As Integer
Dim fdFecIni As Date
Dim fdFecFin As Date
Dim fdFecIniCOMP As Date
Dim fdFecFinCOMP As Date
Dim fnMoneda As Integer
Dim fnNPago As Integer
Dim fnNAdenda As Integer
Dim fMatCronograma() As Variant
Dim I As Integer
'WIOR 20130131 ***********************************
Dim pbActivaArchivo As Boolean
Dim psNombreArchivoFinal As String
Dim fsNomFile As String
Dim psRutaContrato As String
'WIOR ********************************************
'PASI***********************
Dim fntpodocorigen As Integer
Dim fncontTipoPago As Integer
Dim fRsAgencia As New ADODB.Recordset
Dim fRsServicio As New ADODB.Recordset
Dim fRsCompra As New ADODB.Recordset
Dim Datoscontrato() As TContratoBS
'end PASI
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cboCuotaDesdeAD_Click()
Dim oLog As DLogGeneral
Set oLog = New DLogGeneral
Dim rsServicios As ADODB.Recordset
Dim row As Integer
Dim lnDesde, lnHasta As Integer

If fntpodocorigen = LogTipoContrato.ContratoServicio Then
    If Not cboCuotaDesdeAD.ListIndex = -1 Then
       If fnNAdenda2 = 0 Then
            cboCuotaHastaAD.ListIndex = IndiceListaCombo(cboCuotaHastaAD, Trim(Me.cboCuotaDesdeAD.Text))
            lnDesde = CInt(Trim(cboCuotaDesdeAD.Text))
            lnHasta = CInt(Trim(cboCuotaDesdeAD.Text))
            Call LimpiaFlex(feOrden)
            Set rsServicios = oLog.ListaServiciosContrato(fsNContrato, fncontRef, lnDesde, lnHasta)
                If Not rsServicios.EOF Then
                        Do While Not rsServicios.EOF
                            feOrden.AdicionaFila
                            row = feOrden.row
                            feOrden.TextMatrix(row, 1) = rsServicios!cAgeDest
                            feOrden.TextMatrix(row, 2) = rsServicios!cCtaContCod
                            feOrden.TextMatrix(row, 3) = rsServicios!cDescripcion
                            rsServicios.MoveNext
                        Loop
                End If
        End If
    End If
End If
If Not (Trim(Me.cboCuotaDesdeAD.Text) = "" Or Trim(Me.cboCuotaHastaAD.Text) = "") Then
    If CInt(Trim(Me.cboCuotaHastaAD.Text)) < CInt(Trim(Me.cboCuotaDesdeAD.Text)) Then
            MsgBox "Cuota Inicial no puede ser mayor a la Cuota Final.", vbInformation, "Aviso"
            cboCuotaDesdeAD.ListIndex = IndiceListaCombo(cboCuotaDesdeAD, Trim(Me.cboCuotaHastaAD.Text))
    End If
End If
    Set rsServicios = Nothing
End Sub
Private Sub cboCuotaDesdeRED_Click()
Dim oLog As DLogGeneral
Set oLog = New DLogGeneral
Dim rsServicios As ADODB.Recordset
Dim row As Integer
Dim lnDesde, lnHasta As Integer
If fntpodocorigen = LogTipoContrato.ContratoServicio Then
 If Not cboCuotaDesdeRED.ListIndex = -1 Then
    cboCuotaHastaRED.ListIndex = IndiceListaCombo(cboCuotaHastaRED, Trim(Me.cboCuotaDesdeRED.Text))
    lnDesde = CInt(Trim(cboCuotaDesdeRED.Text))
    lnHasta = CInt(Trim(cboCuotaDesdeRED.Text))
    If fnNAdenda2 = 0 Then
        Call LimpiaFlex(feOrden)
        Set rsServicios = oLog.ListaServiciosContrato(fsNContrato, fncontRef, lnDesde, lnHasta)
        If Not rsServicios.EOF Then
                Do While Not rsServicios.EOF
                    feOrden.AdicionaFila
                    row = feOrden.row
                    feOrden.TextMatrix(row, 1) = rsServicios!cAgeDest
                    feOrden.TextMatrix(row, 2) = rsServicios!cCtaContCod
                    feOrden.TextMatrix(row, 3) = rsServicios!cDescripcion
                    rsServicios.MoveNext
                Loop
        End If
    End If
 End If

If Not (Trim(Me.cboCuotaDesdeRED.Text) = "" Or Trim(Me.cboCuotaHastaRED.Text) = "") Then
    If CInt(Trim(Me.cboCuotaHastaRED.Text)) < CInt(Trim(Me.cboCuotaDesdeRED.Text)) Then
            MsgBox "Cuota Inicial no puede ser mayor a la Cuota Final.", vbInformation, "Aviso"
            cboCuotaDesdeRED.ListIndex = IndiceListaCombo(cboCuotaDesdeRED, Trim(Me.cboCuotaHastaRED.Text))
    End If
End If
    Set rsServicios = Nothing
End If
         Set rsServicios = Nothing
End Sub
Private Sub cboCuotaHastaAD_Click()

If Me.cboCuotaDesdeAD.ListIndex = -1 Then
    MsgBox "Primero se Debe Seleccionar la Cuota Desde..", vbInformation, "Aviso"
    cboCuotaDesdeAD.SetFocus
    Exit Sub
End If
If Not (Trim(Me.cboCuotaDesdeAD.Text) = "" Or Trim(Me.cboCuotaHastaAD.Text) = "") Then
    If CInt(Trim(Me.cboCuotaHastaAD.Text)) < CInt(Trim(Me.cboCuotaDesdeAD.Text)) Then
        MsgBox "Cuota Final no puede ser menor a la Cuota Inicial.", vbInformation, "Aviso"
        cboCuotaHastaAD.ListIndex = IndiceListaCombo(cboCuotaHastaAD, Trim(Me.cboCuotaDesdeAD.Text))
    End If
End If
End Sub
Private Sub cboCuotaHastaRED_Click()
If Not (Trim(Me.cboCuotaDesdeRED.Text) = "" Or Trim(Me.cboCuotaHastaRED.Text) = "") Then
    If CInt(Trim(Me.cboCuotaHastaRED.Text)) < CInt(Trim(Me.cboCuotaDesdeRED.Text)) Then
        MsgBox "Cuota Final no puede ser menor a la Cuota Inicial.", vbInformation, "Aviso"
        cboCuotaHastaRED.ListIndex = IndiceListaCombo(cboCuotaHastaRED, Trim(Me.cboCuotaDesdeRED.Text))
    End If
End If
End Sub

Private Sub cboMonedaCro_Click()
If cboMonedaCro.Text <> "" Then
    If CInt(Right(cboMonedaCro.Text, 2)) = gMonedaNacional Then
        txtMontoCro.BackColor = vbWhite
        '''Me.lblTpoMonedaCro.Caption = "S/." 'marg ers044-2016
        Me.lblTpoMonedaCro.Caption = gcPEN_SIMBOLO 'marg ers044-2016
        txtMontoCro.BackColor = vbWhite
        Me.txtMontoExtra.BackColor = vbWhite
        Me.txtMontoRED.BackColor = vbWhite
    Else
        txtMontoCro.BackColor = RGB(200, 255, 200)
        Me.txtMontoExtra.BackColor = RGB(200, 255, 200)
        Me.txtMontoRED.BackColor = RGB(200, 255, 200)
        Me.lblTpoMonedaCro.Caption = "$"
    End If
    If Trim(cboMonedaCro.Text) <> "" Then
        cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, fnMoneda)
    End If
    If fntpodocorigen = LogTipoContrato.ContratoServicio Then
        If fncontTipoPago = 2 Then
            cmdAgregar.SetFocus
        End If
    End If
End If
End Sub
'WIOR 20130131 *************************************
Private Sub chkArchivo_Click()
 lblNombreArchivo.Caption = ""
    If chkArchivo.value = 1 Then
        Me.fraArchivoAdenda.Enabled = True
    Else
        Me.fraArchivoAdenda.Enabled = False
    End If
End Sub
Private Function validaIngresoRegistros() As Boolean 'PASI20140110
    Dim I As Long, j As Long
    Dim col As Integer
    Dim Columnas() As String
    Dim lsColumnas As String
    
    lsColumnas = "1,2,6"
    Columnas = Split(lsColumnas, ",")
        
    validaIngresoRegistros = True
    If feOrden.TextMatrix(1, 0) <> "" Then
        For I = 1 To feOrden.Rows - 1
            For j = 1 To feOrden.Cols - 1
                For col = 0 To UBound(Columnas)
                    If j = Columnas(col) Then
                        If Len(Trim(feOrden.TextMatrix(I, j))) = 0 And feOrden.ColWidth(j) <> 0 Then
                            MsgBox "Ud. debe especificar el campo " & feOrden.TextMatrix(0, j), vbInformation, "Aviso"
                            validaIngresoRegistros = False
                            feOrden.TopRow = I
                            feOrden.row = I
                            feOrden.col = j
                            feOrden_RowColChange
                            Exit Function
                        End If
                    End If
                Next
            Next
            If IsNumeric(feOrden.TextMatrix(I, 6)) Then
                If CCur(feOrden.TextMatrix(I, 6)) <= 0 Then
                    MsgBox "El Importe Total debe ser mayor a cero", vbInformation, "Aviso"
                    validaIngresoRegistros = False
                    feOrden.TopRow = I
                    feOrden.row = I
                    feOrden.col = 6
                    Exit Function
                End If
            Else
                MsgBox "El Importe Total debe ser númerico", vbInformation, "Aviso"
                validaIngresoRegistros = False
                feOrden.TopRow = I
                feOrden.row = I
                feOrden.col = 6
                Exit Function
            End If
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
                fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                If Len(Trim(feOrden.TextMatrix(I, 7))) = 0 Then
                    MsgBox "El Objeto " & feOrden.TextMatrix(I, 3) & Chr(10) & "no tiene configurado Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                    feOrden.TopRow = I
                    feOrden.row = I
                    feOrden.col = 2
                    validaIngresoRegistros = False
                    Exit Function
                End If
            End If
        Next
    Else
        MsgBox "Ud. debe agregar los Bienes/Servicios a dar Conformidad", vbInformation, "Aviso"
        validaIngresoRegistros = False
    End If
End Function
Private Sub cmdAgregarItemCont_Click()
    If Not validaBusqueda Then Exit Sub
    If feOrden.TextMatrix(1, 0) <> "" Then
        If Not validaIngresoRegistros Then Exit Sub
    End If
    feOrden.AdicionaFila
    
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
        feOrden.ColumnasAEditar = "X-1-2-3-X-X-6-X" 'PASI20151229
    End If
    feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    feOrden.col = 2
    feOrden.SetFocus
    feOrden_RowColChange
End Sub

Private Sub cmdBuscarArchivo_Click()
Dim I As Integer
    If Trim(cmdBuscarArchivo.Caption) = "E&xaminar" Then
        CdlgFile.nHwd = Me.hwnd
        CdlgFile.Filtro = "Adenda de Contratos Digital (*.pdf)|*.pdf"
        Me.CdlgFile.Altura = 300
        CdlgFile.Show

        fsPathFile = CdlgFile.Ruta
        fsRuta = fsPathFile
                If fsPathFile <> Empty Then
                    For I = Len(fsPathFile) - 1 To 1 Step -1
                            If Mid(fsPathFile, I, 1) = "\" Then
                                fsPathFile = Mid(CdlgFile.Ruta, 1, I)
                                fsNomFile = Mid(CdlgFile.Ruta, I + 1, Len(CdlgFile.Ruta) - I)
                                Exit For
                            End If
                    Next I
                    Screen.MousePointer = 11

                    If pbActivaArchivo Then

                        psNombreArchivoFinal = Trim(fsNContrato)
                        psNombreArchivoFinal = "Contrato_" & psNombreArchivoFinal
                        psNombreArchivoFinal = psNombreArchivoFinal & "_Adenda_" & IIf(Trim(lblNAdendaCOMP.Caption) = Trim(lblNAdendaAD.Caption), IIf(Trim(lblNAdendaAD.Caption) = Trim(lblNAdendaRED.Caption), Trim(lblNAdendaRED.Caption), Trim(lblNAdendaAD.Caption)), Trim(lblNAdendaAD.Caption))

                        lblNombreArchivo.Caption = UCase(psNombreArchivoFinal) & ".pdf"
                    Else
                        lblNombreArchivo.Caption = ""
                    End If
                Else
                   MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
                   lblNombreArchivo.Caption = ""
                   Exit Sub
                End If
            Screen.MousePointer = 0
    Else
        Dim Ruta As String
        Dim Archivo As New Scripting.FileSystemObject

        If Trim(lblNombreArchivo.Caption) = "" Then
            MsgBox "Adenda de Contrato no Cuenta con Archivo Digital.", vbInformation, "Aviso"
        Else
            Ruta = psRutaContrato & Trim(lblNombreArchivo.Caption)
            If Archivo.FileExists(Ruta) = False Then
                MsgBox "Archivo fue eliminado.", vbCritical, "Aviso"
            Else
                ShellExecute Me.hwnd, "open", Ruta, "", "", 4
            End If
        End If
    End If
End Sub
'WIOR FIN *************************************

Private Sub cmdAgregar_Click()
If ValidaCronograma Then
        If MsgBox("Estas seguro de agregar registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
                If feCronograma.TextMatrix(1, 1) <> "" Then
                    MsgBox "No se puede Agregar Mas de una Cuota para el Contrato de Servicio.", vbInformation + vbOKOnly, "Aviso"
                    Exit Sub
                End If
            End If
                ReDim Preserve fMatCronograma(5, 1 To CInt(Trim(lblNPago.Caption)))
                fMatCronograma(1, CInt(lblNPago.Caption)) = Me.lblNPago.Caption
                fMatCronograma(2, CInt(lblNPago.Caption)) = Format(txtFechaPago.value, "DD/MM/YYYY")
                fMatCronograma(3, CInt(lblNPago.Caption)) = Trim(Right(Me.cboMonedaCro.Text, 4))
                fMatCronograma(4, CInt(lblNPago.Caption)) = Trim(Left(Me.cboMonedaCro.Text, 20))
                fMatCronograma(5, CInt(lblNPago.Caption)) = IIf(txtMontoCro.Enabled = True, Format(Me.txtMontoCro.Text, "##00.00"), "-")
                Call LeerMatriz(CInt(lblNPago.Caption))
                Me.lblNPago.Caption = CInt(lblNPago.Caption) + 1
        End If
        cmdAgregar.SetFocus
End If
End Sub
Private Sub cmdCancelarAD_Click()
LimpiarDatosAD
End Sub

Private Sub cmdCancelarCOMP_Click()
LimpiarDatosComp
End Sub

Private Sub cmdCancelarRED_Click()
LimpiarDatosRED
End Sub

Private Sub cmdQuitar_Click()
If CInt(lblNPago.Caption) > 1 Then
    If MsgBox("Estas seguro de quitar el ultimo registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Me.lblNPago.Caption = CInt(lblNPago.Caption) - 1
        If CInt(lblNPago.Caption) > 1 Then
            ReDim Preserve fMatCronograma(5, 1 To (CInt(lblNPago.Caption) - 1))
        End If
        Call LeerMatriz(CInt(lblNPago.Caption) - 1)
    End If
Else
    MsgBox "No hay datos a eliminar", vbInformation, "Aviso"
End If
End Sub
'EJVG20131203 ***
'Private Sub cmdRegistrarAD_Click()
'On Error GoTo ErrorRegistrarAdendaAD
'If ValidaAdendaAdicional Then
'    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Dim oLog As DLogGeneral
'    Set oLog = New DLogGeneral
'
'        If oLog.RegistrarAdenda(Trim(Me.lblNContratoAD.Caption), CInt(Trim(Me.lblNAdendaAD.Caption)), 2, Format(gdFecSis, "DD/MM/YYYY"), _
'        Format(gdFecSis, "DD/MM/YYYY"), fnMoneda, CDbl(Me.txtMontoExtra.Text), Trim(Me.txtRazonAD.Text), 1, _
'        CInt(Trim(Me.cboCuotaDesdeAD.Text)), CInt(Trim(Me.cboCuotaHastaAD.Text)), Trim(lblNombreArchivo.Caption)) = 0 Then 'WIOR 20130131 AGREGO TRIM(lblNombreArchivo.Caption)
'
'            'WIOR 20130131 ********************************
'            If chkArchivo.value = 1 Then
'                GrabarArchivo
'            End If
'            'WIOR *****************************************
'
'            MsgBox "Adenda Adicional registrada Satisfactoriamente", vbInformation, "Aviso"
'            LimpiarDatosAD
'        Else
'            MsgBox "No se grabaron los datos de la Adenda Adicional", vbInformation, "Aviso"
'        End If
'
'    End If
'End If
'Exit Sub
'ErrorRegistrarAdendaAD:
'   MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
'End Sub
Private Sub cmdRegistrarAD_Click()
    Dim oLog As DLogGeneral
    Dim bTrans As Boolean
    Dim lsNContrato As String
    Dim lnNAdenda As Integer
    Dim lnUltimacuota As Integer
    Dim lnDesde As Integer
    Dim lnHasta As Integer
    Dim lnMontoAdenda As Currency
    Dim I, X As Integer
    Dim Datoscontrato() As TContratoBS 'PASIERS0772014
    Dim Index As Integer
    Dim lsSubCta As String
    Dim indexObj As Integer
    Dim lnImporte As Currency
    
    On Error GoTo ErrorRegistrarAdendaAD
    If Not ValidaAdendaAdicional Then Exit Sub
    
    Set oLog = New DLogGeneral
    lsNContrato = Trim(lblNContratoAD.Caption)
    lnNAdenda = CInt(Trim(lblNAdendaAD.Caption))
    
    If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
        lnDesde = CInt(Trim(cboCuotaDesdeAD.Text))
        lnHasta = CInt(Trim(cboCuotaHastaAD.Text))
    End If
    lnMontoAdenda = CCur(txtMontoExtra.Text)
    
    For I = lnDesde To lnHasta
        If oLog.RealizoPagoContratoxCuota(lsNContrato, lnNAdenda - 1, I, fncontRef) Then
            MsgBox "La cuota N° " & Format(I, "00") & " ya esta en proceso de Pago, no se puede continuar, verifique..", vbInformation, "Aviso"
            Set oLog = Nothing
            Exit Sub
        End If
    Next
    
    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Set oLog = New DLogGeneral
    oLog.dBeginTrans
    bTrans = True
    
    If chkArchivo.value = 1 Then
        GrabarArchivo
    End If
    
    If fncontRef = 0 Then
        oLog.ActualizarContratoProveedor lsNContrato, lnNAdenda
    Else
        oLog.ActualizarContratoProveedorNew lsNContrato, fncontRef, lnNAdenda 'PASIERS0772014
    End If
    oLog.RegistrarAdenda_NEW lsNContrato, fncontRef, lnNAdenda, 2, Format(gdFecSis, "DD/MM/YYYY"), _
                            Format(gdFecSis, "DD/MM/YYYY"), fnMoneda, lnMontoAdenda, Trim(txtRazonAD.Text), 1, _
                             lnDesde, lnHasta, Trim(lblNombreArchivo.Caption) 'fncontRef agregado pasi20140823 ti-ers077-2014
    lnUltimacuota = oLog.MigrarCronogramaxAdenda(lsNContrato, lnNAdenda, fncontRef)  'fncontRef agregado pasi20140823 ti-ers077-2014
    If lnUltimacuota <= 0 Then
        oLog.dRollbackTrans
        Set oLog = Nothing
        MsgBox "No se ha podido registrar la Adenda Adicional", vbCritical, "Aviso"
        Exit Sub
    End If
    If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
        For I = lnDesde To lnHasta
            oLog.InsertaContratoAdendaRel lsNContrato, lnNAdenda, I, fncontRef  'fncontRef agregado pasi20140823 ti-ers077-2014
            oLog.ActualizaMontoCronograma lsNContrato, lnNAdenda, I, lnMontoAdenda, fncontRef
        Next
    End If
    'Codigo CONt SERVICIo
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
        For Index = 1 To feOrden.Rows - 1 'ERS0772014
            ReDim Preserve Datoscontrato(Index)
            Datoscontrato(Index).sAgeCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 1))))
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes _
                Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
            ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                lsSubCta = ""
                For indexObj = 1 To feObj.Rows - 1
                    If feObj.TextMatrix(indexObj, 1) = feOrden.TextMatrix(Index, 0) Then
                        lsSubCta = lsSubCta & feObj.TextMatrix(indexObj, 5)
                    End If
                Next
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 2)))) & lsSubCta
            End If
            Datoscontrato(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
            Datoscontrato(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
            Datoscontrato(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
            Datoscontrato(Index).nTotal = IIf(feOrden.TextMatrix(Index, 6) = "", 0, feOrden.TextMatrix(Index, 6))
        Next
        lnUltimacuota = oLog.MigrarServicioxAdenda(lsNContrato, lnNAdenda, fncontRef)
        If lnUltimacuota <= 0 Then
            oLog.dRollbackTrans
            Set oLog = Nothing
            MsgBox "No se ha podido registrar la Adenda Adicional", vbCritical, "Aviso"
            Exit Sub
        End If
        For I = lnDesde To lnHasta
            If UBound(Datoscontrato) > 0 Then
                For X = 1 To UBound(Datoscontrato)
                    oLog.ActualizaMontoServicio lsNContrato, lnNAdenda, I, Datoscontrato(X).nTotal, fncontRef, Datoscontrato(X).sAgeCod, Datoscontrato(X).sCtaContCod
                    oLog.RegistrarContratoAdendaServicioRel Trim(lsNContrato), fncontRef, lnNAdenda, I, Datoscontrato(X).sAgeCod, Datoscontrato(X).sCtaContCod, Datoscontrato(X).sDescripcion, Datoscontrato(X).nTotal
                Next X
                oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnMontoAdenda, lnNAdenda
            End If
        Next
    End If
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2 Then
        lnUltimacuota = oLog.MigrarServicioxAdenda(lsNContrato, lnNAdenda, fncontRef)
        If lnUltimacuota <= 0 Then
            oLog.dRollbackTrans
            Set oLog = Nothing
            MsgBox "No se ha podido registrar la Adenda Adicional", vbCritical, "Aviso"
            Exit Sub
        End If
        oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnMontoAdenda, lnNAdenda
    End If
    If fntpodocorigen = LogTipoContrato.ContratoArrendamiento Then
         For I = lnDesde To lnHasta
             oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnMontoAdenda, lnNAdenda
        Next I
    End If
    oLog.dCommitTrans
    bTrans = False
    Screen.MousePointer = 0
    
    MsgBox "Adenda Adicional registrada Satisfactoriamente", vbInformation, "Aviso"
    LimpiarDatosAD
    Set oLog = Nothing

    Exit Sub
ErrorRegistrarAdendaAD:
    Screen.MousePointer = 0
    If bTrans Then
        oLog.dRollbackTrans
        Set oLog = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'Private Sub cmdRegistrarCOMP_Click()
'On Error GoTo ErrorRegistrarAdenda
'If ValidaAdendaComplementaria Then
'    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Dim oLog As DLogGeneral
'    Set oLog = New DLogGeneral
'
'        If oLog.RegistrarAdenda(Trim(Me.lblNContratoCOMP.Caption), CInt(Trim(Me.lblNAdendaCOMP.Caption)), 1, Format(Me.txtFecIniCOMP.value, "DD/MM/YYYY"), _
'        Format(Me.txtFecFinCOMP.value, "DD/MM/YYYY"), fnMoneda, 0, Trim(Me.txtGlosa.Text), 1, , , Trim(lblNombreArchivo.Caption)) = 0 Then 'WIOR 20130131 AGREGO TRIM(lblNombreArchivo.Caption)
'
'            'WIOR 20130131 ********************************
'            If chkArchivo.value = 1 Then
'                GrabarArchivo
'            End If
'            'WIOR *****************************************
'            'REGISTRAR CRONOGRAMA
'            For i = 0 To (CInt(Me.lblNPago.Caption) - 2)
'                If oLog.RegistrarCronogramaContrato(Trim(Me.lblNContratoCOMP.Caption), CInt(fMatCronograma(1, i + 1)), Format(fMatCronograma(2, i + 1), "DD/MM/YYYY"), _
'                 CInt(fMatCronograma(3, i + 1)), CDbl(fMatCronograma(5, i + 1)), 1, CInt(Me.lblNAdendaCOMP.Caption)) = 1 Then
'                    MsgBox "No se registro el Nº  de Pago: " & fMatCronograma(1, i + 1), vbInformation, "Aviso"
'                    Exit Sub
'                End If
'            Next i
'
'            MsgBox "Adenda Complementario registrada Satisfactoriamente", vbInformation, "Aviso"
'            LimpiarDatosComp
'        Else
'            MsgBox "No se grabaron los datos de Adenda Complementaria", vbInformation, "Aviso"
'        End If
'
'
'    End If
'End If
'Exit Sub
'ErrorRegistrarAdenda:
'   MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
'End Sub
Private Sub cmdRegistrarCOMP_Click()
    On Error GoTo ErrorRegistrarAdenda
    Dim oLog As New DLogGeneral
    Dim bTrans As Boolean
    Dim lsNContrato As String
    Dim lnNAdenda As Integer
    Dim lnUltimacuota As Integer
    Dim Datoscontrato() As TContratoBS 'PASIERS0772014
    Dim Index As Integer
    Dim lsSubCta As String
    Dim indexObj As Integer
    Dim lnImporte As Currency
    
    If Not ValidaAdendaComplementaria Then Exit Sub
    If UBound(fMatCronograma, 2) <= 0 Then
        MsgBox "No hay datos de adenda complementaria para agregar al actual cronograma", vbCritical, "Aviso"
        Exit Sub
    End If
    lsNContrato = Trim(lblNContratoCOMP.Caption)
    lnNAdenda = CInt(Trim(lblNAdendaCOMP.Caption))
    
    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oLog = New DLogGeneral
    oLog.dBeginTrans
    bTrans = True
    If chkArchivo.value = 1 Then
        GrabarArchivo
    End If
    If fncontRef = 0 Then
        oLog.ActualizarContratoProveedor lsNContrato, lnNAdenda
    Else
        oLog.ActualizarContratoProveedorNew lsNContrato, fncontRef, lnNAdenda 'PASIERS0772014
    End If
    oLog.RegistrarAdenda_NEW lsNContrato, fncontRef, lnNAdenda, 1, Format(txtFecIniCOMP.value, "DD/MM/YYYY"), _
                                Format(txtFecFinCOMP.value, "DD/MM/YYYY"), fnMoneda, 0, Trim(txtGlosa.Text), 1, , , Trim(lblNombreArchivo.Caption) 'fnContRef Agregado PASI20140823 Ti-ERS077-2014
    lnUltimacuota = oLog.MigrarCronogramaxAdenda(lsNContrato, lnNAdenda, fncontRef)
    If lnUltimacuota <= 0 Then
        oLog.dRollbackTrans
        Set oLog = Nothing
        MsgBox "No se ha podido registrar la Adenda Complementaria", vbCritical, "Aviso"
        Exit Sub
    End If
    For I = 1 To UBound(fMatCronograma, 2)
        oLog.InsertaContratoAdendaRel lsNContrato, lnNAdenda, lnUltimacuota + I, fncontRef 'fnContRef Agregado PASI20140823 Ti-ERS077-2014
        oLog.RegistrarCronogramaContrato_NEW lsNContrato, lnUltimacuota + I, Format(fMatCronograma(2, I), "DD/MM/YYYY"), _
                                            CInt(fMatCronograma(3, I)), CDbl(IIf(fMatCronograma(5, I) = "-", 0, fMatCronograma(5, I))), 1, lnNAdenda, fncontRef 'fnContRef Agregado PASI20140823 Ti-ERS077-2014
    Next
    'Registra Adenda Comp. Contrato Servicio
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
        For Index = 1 To feOrden.Rows - 1 'ERS0772014
            ReDim Preserve Datoscontrato(Index)
            Datoscontrato(Index).sAgeCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 1))))
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes _
                Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
            ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                lsSubCta = ""
                For indexObj = 1 To feObj.Rows - 1
                    If feObj.TextMatrix(indexObj, 1) = feOrden.TextMatrix(Index, 0) Then
                        lsSubCta = lsSubCta & feObj.TextMatrix(indexObj, 5)
                    End If
                Next
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 2)))) & lsSubCta
            End If
            Datoscontrato(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
            Datoscontrato(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
            Datoscontrato(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
            Datoscontrato(Index).nTotal = feOrden.TextMatrix(Index, 6)
        Next
        lnUltimacuota = oLog.MigrarServicioxAdenda(lsNContrato, lnNAdenda, fncontRef)
        If lnUltimacuota <= 0 Then
            oLog.dRollbackTrans
            Set oLog = Nothing
            MsgBox "No se ha podido registrar la Adenda Complementaria", vbCritical, "Aviso"
            Exit Sub
        End If
        If UBound(Datoscontrato) > 0 Then 'PASIERS0772014
                If fntpodocorigen = LogTipoContrato.ContratoServicio Then
                    For I = 1 To UBound(Datoscontrato)
                        lnImporte = lnImporte + Datoscontrato(I).nTotal
                        oLog.RegistrarContratoServicio Trim(lsNContrato), fncontRef, lnNAdenda, lnUltimacuota + 1, Datoscontrato(I).sAgeCod, Datoscontrato(I).sCtaContCod, Datoscontrato(I).sDescripcion, I, Datoscontrato(I).nTotal
                        oLog.RegistrarContratoAdendaServicioRel Trim(lsNContrato), fncontRef, lnNAdenda, lnUltimacuota + 1, Datoscontrato(I).sAgeCod, Datoscontrato(I).sCtaContCod, Datoscontrato(I).sDescripcion, Datoscontrato(I).nTotal
                    Next I
                    oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnImporte, lnNAdenda
                End If
        End If
    End If
    If (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then
        lnUltimacuota = oLog.MigrarServicioxAdenda(lsNContrato, lnNAdenda, fncontRef)
        If lnUltimacuota <= 0 Then
            oLog.dRollbackTrans
            Set oLog = Nothing
            MsgBox "No se ha podido registrar la Adenda Complementaria", vbCritical, "Aviso"
            Exit Sub
        End If
        For I = 1 To UBound(fMatCronograma, 2)
            oLog.InsertaItemServicioxPagoVariable lsNContrato, fncontRef, lnNAdenda, lnUltimacuota + I
        Next
    End If
    If fntpodocorigen = LogTipoContrato.ContratoArrendamiento Then
        For I = 1 To UBound(fMatCronograma, 2)
            lnImporte = lnImporte + CDbl(fMatCronograma(5, I))
        Next I
        oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnImporte, lnNAdenda
    End If
    'end pasi
    oLog.dCommitTrans
    bTrans = False
    Set oLog = Nothing
        
    MsgBox "Adenda complementaria registrada satisfactoriamente", vbInformation, "Aviso"
    'ARLO 20160126 ***
    gsOpeCod = LogPistaRegistraAdenda
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Grabo la Adenda N° : " & lnNAdenda & " | Del Contrato N° : " & lsNContrato
    Set objPista = Nothing
    '***
    LimpiarDatosComp
    Exit Sub
ErrorRegistrarAdenda:
    If bTrans Then
        oLog.dRollbackTrans
        Set oLog = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdRegistrarRED_Click()
    Dim oLog As DLogGeneral
    Dim bTrans As Boolean
    Dim lsNContrato As String
    Dim lnNAdenda As Integer
    Dim lnUltimacuota As Integer
    Dim lnDesde As Integer
    Dim lnHasta As Integer
    Dim lnMontoAdenda As Currency
    Dim Datoscontrato() As TContratoBS 'PASIERS0772014
    Dim Index As Integer
    Dim lsSubCta As String
    Dim indexObj As Integer
    Dim lnImporte As Currency
    Dim X As Integer
    
    On Error GoTo ErrorRegistrarAdendaRED
    If Not ValidaAdendaReduccion Then Exit Sub
    
    lsNContrato = Trim(lblNContratoRED.Caption)
    lnNAdenda = CInt(Trim(lblNAdendaRED.Caption))
    If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
        lnDesde = CInt(Trim(cboCuotaDesdeRED.Text))
        lnHasta = CInt(Trim(cboCuotaHastaRED.Text))
    End If
    lnMontoAdenda = CCur(txtMontoRED.Text)
    
    Set oLog = New DLogGeneral
    For I = lnDesde To lnHasta
        If oLog.RealizoPagoContratoxCuota(lsNContrato, lnNAdenda - 1, I, fncontRef) Then
            MsgBox "La cuota N° " & Format(I, "00") & " ya ha sido pagada, no se puede continuar, verifique..", vbInformation, "Aviso"
            Set oLog = Nothing
            Exit Sub
        End If
    Next
    
    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Set oLog = New DLogGeneral
    oLog.dBeginTrans
    bTrans = True
    
    If chkArchivo.value = 1 Then
        GrabarArchivo
    End If
    
    If fncontRef = 0 Then
        oLog.ActualizarContratoProveedor lsNContrato, lnNAdenda
    Else
        oLog.ActualizarContratoProveedorNew lsNContrato, fncontRef, lnNAdenda 'PASIERS0772014
    End If
    oLog.RegistrarAdenda_NEW lsNContrato, fncontRef, lnNAdenda, 3, Format(gdFecSis, "DD/MM/YYYY"), _
                            Format(gdFecSis, "DD/MM/YYYY"), fnMoneda, lnMontoAdenda, Trim(txtRazonRED.Text), 1, _
                            lnDesde, lnHasta, Trim(lblNombreArchivo.Caption) 'fnContRef agregado pasi20140823 ti-ers077-2014
    lnUltimacuota = oLog.MigrarCronogramaxAdenda(lsNContrato, lnNAdenda, fncontRef)
    If lnUltimacuota <= 0 Then
        oLog.dRollbackTrans
        Set oLog = Nothing
        MsgBox "No se ha podido registrar la Adenda de Reducción", vbCritical, "Aviso"
        Exit Sub
    End If
    
    For I = lnDesde To lnHasta
        oLog.InsertaContratoAdendaRel lsNContrato, lnNAdenda, I, fncontRef
        oLog.ActualizaMontoCronograma lsNContrato, lnNAdenda, I, -1 * lnMontoAdenda, fncontRef
    Next
    'Codigo cont Servicio
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
        For Index = 1 To feOrden.Rows - 1 'ERS0772014
            ReDim Preserve Datoscontrato(Index)
            Datoscontrato(Index).sAgeCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 1))))
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes _
                Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
            ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                lsSubCta = ""
                For indexObj = 1 To feObj.Rows - 1
                    If feObj.TextMatrix(indexObj, 1) = feOrden.TextMatrix(Index, 0) Then
                        lsSubCta = lsSubCta & feObj.TextMatrix(indexObj, 5)
                    End If
                Next
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 2)))) & lsSubCta
            End If
            Datoscontrato(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
            Datoscontrato(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
            Datoscontrato(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
            Datoscontrato(Index).nTotal = IIf(feOrden.TextMatrix(Index, 6) = "", 0, feOrden.TextMatrix(Index, 6))
        Next
        lnUltimacuota = oLog.MigrarServicioxAdenda(lsNContrato, lnNAdenda, fncontRef)
        If lnUltimacuota <= 0 Then
            oLog.dRollbackTrans
            Set oLog = Nothing
            MsgBox "No se ha podido registrar la Adenda de Recuccion", vbCritical, "Aviso"
            Exit Sub
        End If
        For I = lnDesde To lnHasta
            If UBound(Datoscontrato) > 0 Then
                For X = 1 To UBound(Datoscontrato)
                    oLog.ActualizaMontoServicio lsNContrato, lnNAdenda, I, -1 * Datoscontrato(X).nTotal, fncontRef, Datoscontrato(X).sAgeCod, Datoscontrato(X).sCtaContCod
                    'Modificado PASI20150923
                    'oLog.RegistrarContratoAdendaServicioRel Trim(lsNContrato), fncontRef, lnNAdenda, lnUltimacuota + 1, Datoscontrato(i).sAgeCod, Datoscontrato(i).sCtaContCod, Datoscontrato(i).sDescripcion, Datoscontrato(i).nTotal
                    oLog.RegistrarContratoAdendaServicioRel Trim(lsNContrato), fncontRef, lnNAdenda, I, Datoscontrato(X).sAgeCod, Datoscontrato(X).sCtaContCod, Datoscontrato(X).sDescripcion, Datoscontrato(X).nTotal
                    'end pasi
                Next X
            End If
            oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnMontoAdenda * -1, lnNAdenda
        Next
    End If
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2 Then
        lnUltimacuota = oLog.MigrarServicioxAdenda(lsNContrato, lnNAdenda, fncontRef)
        If lnUltimacuota <= 0 Then
            oLog.dRollbackTrans
            Set oLog = Nothing
            MsgBox "No se ha podido registrar la Adenda Adicional", vbCritical, "Aviso"
            Exit Sub
        End If
        oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnMontoAdenda * -1, lnNAdenda
    End If
    If fntpodocorigen = LogTipoContrato.ContratoArrendamiento Then
         For I = lnDesde To lnHasta
             oLog.ActualizaSaldoContrato lsNContrato, fncontRef, lnMontoAdenda * -1, lnNAdenda
        Next I
    End If
    
    oLog.dCommitTrans
    bTrans = False
    Screen.MousePointer = 0
    Set oLog = Nothing
    
    MsgBox "Adenda de Reducción registrada Satisfactoriamente", vbInformation, "Aviso"
    LimpiarDatosRED
    Exit Sub
ErrorRegistrarAdendaRED:
    If bTrans Then
        oLog.dRollbackTrans
        Set oLog = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'Private Sub cmdRegistrarRED_Click()
'On Error GoTo ErrorRegistrarAdendaRED
'If ValidaAdendaReduccion Then
'    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    Dim oLog As DLogGeneral
'    Set oLog = New DLogGeneral
'
'        If oLog.RegistrarAdenda(Trim(Me.lblNContratoRED.Caption), CInt(Trim(Me.lblNAdendaRED.Caption)), 3, Format(gdFecSis, "DD/MM/YYYY"), _
'        Format(gdFecSis, "DD/MM/YYYY"), fnMoneda, CDbl(Me.txtMontoRED.Text), Trim(Me.txtRazonRED.Text), 1, _
'        CInt(Trim(Me.cboCuotaDesdeRED.Text)), CInt(Trim(Me.cboCuotaHastaRED.Text)), Trim(lblNombreArchivo.Caption)) = 0 Then 'WIOR 20130131 AGREGO TRIM(lblNombreArchivo.Caption)
'
'            'WIOR 20130131 ********************************
'            If chkArchivo.value = 1 Then
'                GrabarArchivo
'            End If
'            'WIOR *****************************************
'            MsgBox "Adenda de Reducción registrada Satisfactoriamente", vbInformation, "Aviso"
'            LimpiarDatosRED
'        Else
'            MsgBox "No se grabaron los datos de la Adenda de Reducción", vbInformation, "Aviso"
'        End If
'
'    End If
'End If
'Exit Sub
'ErrorRegistrarAdendaRED:
'   MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
'End Sub
'END EJVG *******

Private Sub feOrden_OnCellChange(pnRow As Long, pnCol As Long)
    On Error GoTo ErrfeOrden_OnCellChange
    If feOrden.TextMatrix(1, 0) <> "" Then
        If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
            fntpodocorigen = LogTipoContrato.ContratoSuministro Then
            If pnCol = 4 Or pnCol = 5 Then
                feOrden.TextMatrix(pnRow, 6) = Format(Val(feOrden.TextMatrix(pnRow, 4)) * feOrden.TextMatrix(pnRow, 5), gsFormatoNumeroView)
            End If
        End If
    End If
    Exit Sub
ErrfeOrden_OnCellChange:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOrden_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    If psDataCod <> "" Then
        If pnCol = 2 Then
            If fntpodocorigen = LogTipoDocOrigenActaConformidad.Serviciolibre Then
                AsignaObjetosSer psDataCod
            End If
        End If
        If pnCol = 1 Or pnCol = 2 Then
            '*** Si esta vacio el campo de la cuenta contable y si ya eligió agencia y objeto
            If Len(Trim(feOrden.TextMatrix(pnRow, 1))) <> 0 And Len(Trim(feOrden.TextMatrix(pnRow, 2))) <> 0 Then
                feOrden.TextMatrix(pnRow, 7) = DameCtaCont(feOrden.TextMatrix(pnRow, 2), 0, Trim(feOrden.TextMatrix(pnRow, 1)))
            End If
            '***
        End If
    End If
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
Private Sub feOrden_RowColChange()
      If feOrden.col = 1 Then
        feOrden.rsTextBuscar = fRsAgencia
    ElseIf feOrden.col = 2 Then
        If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
            fntpodocorigen = LogTipoContrato.ContratoSuministro Then
            feOrden.rsTextBuscar = fRsCompra
        ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
            feOrden.rsTextBuscar = fRsServicio
        End If
    End If
End Sub

Private Sub Form_Load()
Dim oConst As DConstantes
Dim oConstSist As NConstSistemas 'WIOR 20130131
Set oConst = New DConstantes

CargaCombo oConst.GetConstante(gMoneda), Me.cboMonedaCro
    
'WIOR 20130131 *************************************
Dim oLog As DLogGeneral
Dim NAdenda As String
Set oLog = New DLogGeneral
NAdenda = oLog.NombreArchivoAdenda(fsNContrato, fnNAdenda2)
lblNombreArchivo.Caption = NAdenda

    If Trim(Mid(GetMaquinaUsuario, 1, 2)) = "01" Then
        'OBTENER RUTA DE CONTRATOS
        Set oConstSist = New NConstSistemas
        psRutaContrato = Trim(oConstSist.LeeConstSistema(gsLogContRutaContratos)) & "Adendas\"
        
        If fnNAdenda2 = 0 Then
            Me.cmdBuscarArchivo.Enabled = True
            pbActivaArchivo = True
            cmdBuscarArchivo.Caption = "E&xaminar"
        Else
            cmdBuscarArchivo.Enabled = True
            cmdBuscarArchivo.Caption = "V&er PDF"
            fraArchivoAdenda.Enabled = True
            
            If Trim(lblNombreArchivo.Caption) <> "" Then
                chkArchivo.value = 1
                lblNombreArchivo.Caption = NAdenda
            Else
                chkArchivo.value = 0
                lblNombreArchivo.Caption = ""
            End If
            chkArchivo.Enabled = False
        End If
    Else
        Me.cmdBuscarArchivo.Enabled = False
        fraArchivoAdenda.Enabled = False
        chkArchivo.Enabled = False
        pbActivaArchivo = False
        psRutaContrato = ""
        If fnNAdenda2 = 0 Then
            cmdBuscarArchivo.Caption = "E&xaminar"
        Else
            cmdBuscarArchivo.Caption = "V&er PDF"
        End If
    End If
'WIOR **********************************************
End Sub

Private Sub txtFecFinCOMP_Change()
If CDate(txtFecFinCOMP.value) < CDate(Me.txtFecIniCOMP.value) Then
    MsgBox "Fecha Final no puede ser menor a la Fecha Inicial.", vbInformation, "Aviso"
    txtFecFinCOMP.value = Me.txtFecIniCOMP.value
End If
End Sub

Private Sub txtFechaPago_Change()
If CDate(fdFecFinCOMP) >= CDate(txtFechaPago.value) Then
    MsgBox "Fecha no puede ser menor o igual a la Fecha de Fin del Contrato.", vbInformation, "Aviso"
    Me.txtFechaPago.value = DateAdd("d", 1, fdFecFin)
End If
End Sub


Private Sub txtFecIniCOMP_Change()
If CDate(fdFecFinCOMP) >= CDate(txtFecIniCOMP.value) Then
    MsgBox "Fecha no puede ser menor o igual a la Fecha de Fin del Contrato.", vbInformation, "Aviso"
    Me.txtFecIniCOMP.value = DateAdd("d", 1, fdFecFin)
Else
    If CDate(txtFecFinCOMP.value) < CDate(Me.txtFecIniCOMP.value) Then
        MsgBox "Fecha Inicial no puede ser mayor a la Fecha Final.", vbInformation, "Aviso"
        Me.txtFecIniCOMP.value = txtFecFinCOMP.value
    End If
End If
End Sub

Private Sub txtMontoCro_GotFocus()
fEnfoque txtMontoCro
End Sub

Private Sub txtMontoCro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoCro, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        Me.cmdAgregar.SetFocus
    End If
End Sub

Private Sub txtMontoCro_LostFocus()
If Trim(txtMontoCro.Text) = "" Then
        txtMontoCro.Text = "0.00"
    End If
    txtMontoCro.Text = Format(txtMontoCro.Text, "#0.00")
End Sub


Private Sub txtMontoExtra_GotFocus()
fEnfoque txtMontoExtra
End Sub

Private Sub txtMontoExtra_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoExtra, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        If (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then
            Me.cmdRegistrarAD.SetFocus
        Else
            Me.cboCuotaDesdeAD.SetFocus
        End If
    End If
End Sub
Private Sub txtMontoExtra_LostFocus()
If Trim(txtMontoExtra.Text) = "" Then
        txtMontoExtra.Text = "0.00"
    End If
    txtMontoExtra.Text = Format(txtMontoExtra.Text, "#0.00")
End Sub


Private Sub txtMontoRED_GotFocus()
fEnfoque txtMontoRED
End Sub

Private Sub txtMontoRED_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoRED, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        If (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then
            Me.cmdRegistrarRED.SetFocus
        Else
            Me.cboCuotaDesdeRED.SetFocus
        End If
    End If
End Sub

Private Sub txtMontoRED_LostFocus()
If Trim(txtMontoRED.Text) = "" Then
        txtMontoRED.Text = "0.00"
    End If
    txtMontoRED.Text = Format(txtMontoRED.Text, "#0.00")
End Sub

Public Sub Inicio(ByVal psNContrato As String, Optional ByVal pnNAdenda As Integer = 0, Optional ByVal pnTipo As Integer = 0, Optional ByVal pnTipoContOrigen As Integer = 0, Optional ByVal pnContRef As Integer = 0, Optional ByVal pncontTipoPago As Integer = 1) 'pnContRef Agregado PASI20140823 Ti-ERS077-2014
fsNContrato = psNContrato
fnNAdenda2 = pnNAdenda
fncontRef = pnContRef
fnTipo = pnTipo
fntpodocorigen = pnTipoContOrigen
fncontTipoPago = pncontTipoPago
If fnNAdenda2 = 0 Then
    If fntpodocorigen = LogTipoContrato.ContratoServicio Then 'PASIERS0772014
        Call VerPestana(fnTipo - 1)
    End If
    Call CargaDatos
Else
    Call VerPestana(fnTipo - 1)
    Call CargaDatosDetalle
End If
Me.Show 1
End Sub
Private Sub CargaDatos()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset
Dim rsServicios As ADODB.Recordset
Dim bAdenda As Boolean
Set oLog = New DLogGeneral
If oLog.ExisteAdendaContratos(fsNContrato, fncontRef) > 0 Then 'fnContRef Agregado pasi20140823 ti-ers077-2014
    Set rsLog = oLog.ListarDatosAdenda(fsNContrato, oLog.ExisteAdendaContratos(fsNContrato, fncontRef), fncontRef) 'fnContRef Agregado pasi20140823 ti-ers077-2014
    bAdenda = True
Else
    Set rsLog = oLog.ListarDatosContratos(fsNContrato, fncontRef) 'fnContRef Agregado pasi20140823 ti-ers077-2014
    bAdenda = False
End If
If rsLog.RecordCount > 0 Then
    'COMPLEMENTARIAS
    Me.lblNContratoCOMP.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorCOMP.Caption = Space(1) & rsLog!Proveedor
    'ADICIONALES
    Me.lblNContratoAD.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorAD.Caption = Space(1) & rsLog!Proveedor
    'REDUCCIONES
    Me.lblNContratoRED.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorRED.Caption = Space(1) & rsLog!Proveedor
    '
    fdFecIniCOMP = CDate(rsLog!Desde)
    fdFecFinCOMP = CDate(rsLog!Hasta)
    
    Me.txtFechaPago.value = DateAdd("d", 1, fdFecFinCOMP)
    Me.txtFecIniCOMP.value = DateAdd("d", 1, fdFecFinCOMP)
    Me.txtFecFinCOMP.value = DateAdd("d", 1, fdFecFinCOMP)
    fnMoneda = CInt(rsLog!nMoneda)
    'cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, fnMoneda)
    
    If bAdenda Then
        fdFecIni = CDate(rsLog!DesdeC)
        fdFecFin = CDate(rsLog!HastaC)
    Else
        fdFecIni = CDate(rsLog!Desde)
        fdFecFin = CDate(rsLog!Hasta)
    End If
    

    fnNAdenda = oLog.ObtenerUltAdendaContratos(Trim(rsLog!NContrato), fncontRef) + 1  'fnContRef Agregado pasi20140823 ti-ers077-2014
    Me.lblNPago.Caption = "1"
    Me.lblNAdendaCOMP.Caption = fnNAdenda
    Me.lblNAdendaAD.Caption = fnNAdenda
    Me.lblNAdendaRED.Caption = fnNAdenda
    
    Set rsLog = oLog.ObtenerCuotasContratos(fsNContrato, fncontRef)  'fnContRef Agregado pasi20140823 ti-ers077-2014
    Call CargaCuotas(rsLog)
    
    If fntpodocorigen = LogTipoContrato.ContratoServicio Then 'PASIERS0772014
        Dim row As Integer
        Dim odoc As New DOperacion
        Dim oArea As New DActualizaDatosArea
        Dim oalmacen As New DLogAlmacen
    
        Set fRsAgencia = oArea.GetAgencias(, , True)
        Set fRsCompra = oalmacen.GetBienesAlmacen(, "11','12','13")
        Set fRsServicio = OrdenServicio()
    
        If fnTipo = LogtipoReajusteAdenda.Complementaria Then
            cmdAgregarItemCont.Enabled = True
            cmdQuitarItemCont.Enabled = True
        ElseIf fnTipo = LogtipoReajusteAdenda.Adicional Or fnTipo = LogtipoReajusteAdenda.Reduccion Then
'            Set rsServicios = olog.ListaServiciosContrato(fsNContrato, fnContref) CPASI
'            If Not rsServicios.EOF Then
'                Do While Not rsServicios.EOF
'                    feOrden.AdicionaFila
'                    row = feOrden.row
'                    feOrden.TextMatrix(row, 1) = rsServicios!cAgeDest
'                    feOrden.TextMatrix(row, 2) = rsServicios!cCTaContCod
'                    feOrden.TextMatrix(row, 3) = rsServicios!cCtaContDesc
'                    rsServicios.MoveNext
'                Loop
'            End If
            'If fnTipo = LogtipoReajusteAdenda.Reduccion Then
                cboCuotaHastaAD.Enabled = False
                cboCuotaHastaRED.Enabled = False
                cmdQuitarItemCont.Enabled = False
                cmdAgregarItemCont.Enabled = False
            'End If
            feOrden.ColumnasAEditar = "X-X-X-X-X-X-6-X"
        End If
        If fncontTipoPago = 2 Then
            txtMontoCro.Enabled = False
            cboCuotaDesdeAD.Enabled = False
            cboCuotaHastaAD.Enabled = False
            cboCuotaDesdeRED.Enabled = False
            cboCuotaHastaRED.Enabled = False
        End If
    End If
    If fntpodocorigen = LogTipoContrato.ContratoArrendamiento Then 'PASIERS0772014
        Me.SSTContratos.TabVisible(3) = False
    End If
End If
End Sub
'Private Sub DesHabilitaTabs(ByVal pnTipoAdenda As Integer) 'CPASI
'    Select Case pnTipoAdenda
'        Case LogtipoReajusteAdenda.Complementaria
'            Me.SSTContratos.TabVisible(0) = True
'            Me.SSTContratos.TabVisible(1) = False
'            Me.SSTContratos.TabVisible(2) = False
'            Me.SSTContratos.TabVisible(3) = True
'        Case LogtipoReajusteAdenda.Adicional
'            Me.SSTContratos.TabVisible(0) = False
'            Me.SSTContratos.TabVisible(1) = True
'            Me.SSTContratos.TabVisible(2) = False
'            Me.SSTContratos.TabVisible(3) = True
'        Case LogtipoReajusteAdenda.Reduccion
'            Me.SSTContratos.TabVisible(0) = False
'            Me.SSTContratos.TabVisible(1) = False
'            Me.SSTContratos.TabVisible(2) = True
'            Me.SSTContratos.TabVisible(3) = True
'    End Select
'End Sub
Private Function ValidaCronograma() As Boolean
If Me.cboMonedaCro.Text = "" Then
    MsgBox "Seleccionar moneda para la cuota.", vbInformation, "Aviso"
    ValidaCronograma = False
    Exit Function
End If
If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then
    If Trim(Me.txtMontoCro.Text) = "" Or Trim(Me.txtMontoCro.Text) = "0.00" Then
        MsgBox "Ingrese el monto de la cuota.", vbInformation, "Aviso"
        ValidaCronograma = False
        Exit Function
    End If
End If
ValidaCronograma = True
End Function
Private Sub LeerMatriz(ByVal tamano As Integer)
Dim I As Integer
Call LimpiaFlex(feCronograma)
For I = 0 To tamano - 1
    feCronograma.AdicionaFila
    feCronograma.TextMatrix(I + 1, 0) = I + 1
    feCronograma.TextMatrix(I + 1, 1) = fMatCronograma(1, I + 1)
    feCronograma.TextMatrix(I + 1, 2) = fMatCronograma(2, I + 1)
    feCronograma.TextMatrix(I + 1, 3) = fMatCronograma(4, I + 1)
    feCronograma.TextMatrix(I + 1, 4) = fMatCronograma(5, I + 1)
Next I
End Sub

Private Function ValidaAdendaComplementaria() As Boolean
If Trim(Me.txtGlosa.Text) = "" Then
    MsgBox "Ingrese la glosa de la adenda.", vbInformation, "Aviso"
    ValidaAdendaComplementaria = False
    Exit Function
End If

If Trim(Me.feCronograma.TextMatrix(1, 1)) = "" Then
    MsgBox "Aun no Ingreso Cuotas complementarias", vbInformation, "Aviso"
    ValidaAdendaComplementaria = False
    Exit Function
End If
'WIOR 20130131 *************************************
If chkArchivo.value = 1 Then
    If Trim(lblNombreArchivo.Caption) = "" Then
        MsgBox "Seleccione el archivo de la Adenda", vbInformation, "Aviso"
        ValidaAdendaComplementaria = False
        Exit Function
    End If
End If

'WIOR **********************************************
'PASIERS0772014
If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
    If fnTipo = LogtipoReajusteAdenda.Complementaria Then
        Dim I As Integer
        Dim lnImporte As Currency
        If feOrden.TextMatrix(1, 1) = "" Then
            MsgBox "Asegurese de ingresar corrrectamente los Servicios.", vbInformation, "Aviso"
            ValidaAdendaComplementaria = False
            Exit Function
        End If
        For I = 1 To feOrden.Rows - 1
            lnImporte = lnImporte + (feOrden.TextMatrix(I, 6))
        Next I
        If lnImporte <> feCronograma.TextMatrix(1, 4) Then
            MsgBox "El Monto de la Adenda no coincide con el Monto de los Servicios.", vbInformation, "Aviso"
            ValidaAdendaComplementaria = False
            Exit Function
        End If
    End If
End If
'end PASI
ValidaAdendaComplementaria = True
End Function
Private Function ValidaAdendaAdicional() As Boolean
If Trim(Me.txtMontoExtra.Text) = "" Or Trim(Me.txtMontoExtra.Text) = "0.00" Then
    MsgBox "Ingrese Monto Adicional.", vbInformation, "Aviso"
    ValidaAdendaAdicional = False
    Exit Function
End If

If Trim(Me.txtRazonAD.Text) = "" Then
    MsgBox "Ingrese la razon de la Adenda Adicional.", vbInformation, "Aviso"
    ValidaAdendaAdicional = False
    Exit Function
End If

If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
    If Trim(Me.cboCuotaDesdeAD.Text) = "" Then
        MsgBox "Ingrese la Cuota de inicio de la adenda.", vbInformation, "Aviso"
        ValidaAdendaAdicional = False
        Exit Function
    End If
End If
If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
    If Trim(Me.cboCuotaHastaAD.Text) = "" Then
        MsgBox "Ingrese la Cuota Final de la Adenda.", vbInformation, "Aviso"
        ValidaAdendaAdicional = False
        Exit Function
    End If
End If

'WIOR 20130131 *************************************
If chkArchivo.value = 1 Then
    If Trim(lblNombreArchivo.Caption) = "" Then
        MsgBox "Seleccione el archivo de la Adenda", vbInformation, "Aviso"
        ValidaAdendaAdicional = False
        Exit Function
    End If
End If
'WIOR **********************************************
'PASIERS0772014
If (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1) Then
    If fnTipo = LogtipoReajusteAdenda.Adicional Then
        Dim I As Integer
        Dim lnImporte As Currency
        For I = 1 To feOrden.Rows - 1
            lnImporte = lnImporte + IIf((feOrden.TextMatrix(I, 6)) = "", 0, (feOrden.TextMatrix(I, 6)))
        Next I
        If lnImporte <> CDbl(txtMontoExtra.Text) Then
            MsgBox "El Monto de la Adenda no coincide con el Monto de los Servicios.", vbInformation, "Aviso"
            ValidaAdendaAdicional = False
            Exit Function
        End If
    End If
End If
'end PASI
ValidaAdendaAdicional = True
End Function

Private Function ValidaAdendaReduccion() As Boolean
If Trim(Me.txtMontoRED.Text) = "" Or Trim(Me.txtMontoRED.Text) = "0.00" Then
    MsgBox "Ingrese Monto de Reducción.", vbInformation, "Aviso"
    ValidaAdendaReduccion = False
    Exit Function
End If

If Trim(Me.txtRazonRED.Text) = "" Then
    MsgBox "Ingrese la razon de la Adenda de Reducción.", vbInformation, "Aviso"
    ValidaAdendaReduccion = False
    Exit Function
End If

If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
    If Trim(Me.cboCuotaDesdeRED.Text) = "" Then
        MsgBox "Ingrese la Cuota de Inicio de la adenda.", vbInformation, "Aviso"
        ValidaAdendaReduccion = False
        Exit Function
    End If
End If
If Not (fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2) Then 'PASIERS0772014
    If Trim(Me.cboCuotaHastaRED.Text) = "" Then
        MsgBox "Ingrese la Cuota Final de la Adenda.", vbInformation, "Aviso"
        ValidaAdendaReduccion = False
        Exit Function
    End If
End If
'WIOR 20130131 *************************************
If chkArchivo.value = 1 Then
    If Trim(lblNombreArchivo.Caption) = "" Then
        MsgBox "Seleccione el archivo de la Adenda", vbInformation, "Aviso"
        ValidaAdendaReduccion = False
        Exit Function
    End If
End If
'WIOR **********************************************
'PASIERS0772014
If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 1 Then
    If fnTipo = LogtipoReajusteAdenda.Reduccion Then
        Dim I As Integer
        Dim lnImporte As Currency
        For I = 1 To feOrden.Rows - 1
            lnImporte = lnImporte + (feOrden.TextMatrix(I, 6))
        Next I
        If lnImporte <> CDbl(txtMontoRED.Text) Then
            MsgBox "El Monto de la Adenda no coincide con el Monto de los Servicios.", vbInformation, "Aviso"
            ValidaAdendaReduccion = False
            Exit Function
        End If
    End If
End If
'end PASI
ValidaAdendaReduccion = True
End Function
Sub LimpiarDatosComp()
Me.txtMontoCro.Text = ""
Me.txtGlosa.Text = ""

Call LimpiaFlex(Me.feCronograma)
Call LimpiaFlex(Me.feOrden)
ReDim Preserve fMatCronograma(5, 1 To 1)

Call CargarDatosGenerales
'WIOR 20130131 *************************************
chkArchivo.value = 0
lblNombreArchivo.Caption = ""
fraArchivoAdenda.Enabled = False
'WIOR **********************************************
End Sub
Sub LimpiarDatosAD()
Me.txtMontoExtra.Text = ""
Me.txtRazonAD.Text = ""
Call CargarDatosGenerales
'WIOR 20130131 *************************************
chkArchivo.value = 0
lblNombreArchivo.Caption = ""
fraArchivoAdenda.Enabled = False
'WIOR **********************************************
Call LimpiaFlex(feOrden)
End Sub
Sub LimpiarDatosRED()
Me.txtMontoRED.Text = ""
Me.txtRazonRED.Text = ""
Call CargarDatosGenerales
'WIOR 20130131 *************************************
chkArchivo.value = 0
lblNombreArchivo.Caption = ""
fraArchivoAdenda.Enabled = False
'WIOR **********************************************
Call LimpiaFlex(feOrden)
End Sub

Sub CargarDatosGenerales()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset
Dim bAdenda As Boolean
Set oLog = New DLogGeneral

If oLog.ExisteAdendaContratos(fsNContrato, fncontRef) > 0 Then
    Set rsLog = oLog.ListarDatosAdenda(fsNContrato, fncontRef, oLog.ExisteAdendaContratos(fsNContrato, fncontRef))
    bAdenda = True
Else
    Set rsLog = oLog.ListarDatosContratos(fsNContrato, fncontRef)
    bAdenda = False
End If
If rsLog.RecordCount > 0 Then
    'COMPLEMENTARIAS
    Me.lblNContratoCOMP.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorCOMP.Caption = Space(1) & rsLog!Proveedor
    'ADICIONALES
    Me.lblNContratoAD.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorAD.Caption = Space(1) & rsLog!Proveedor
    'REDUCCIONES
    Me.lblNContratoRED.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorRED.Caption = Space(1) & rsLog!Proveedor
    '
    fdFecIniCOMP = CDate(rsLog!Desde)
    fdFecFinCOMP = CDate(rsLog!Hasta)
    
    Me.txtFechaPago.value = DateAdd("d", 1, fdFecFinCOMP)
    Me.txtFecIniCOMP.value = DateAdd("d", 1, fdFecFinCOMP)
    Me.txtFecFinCOMP.value = DateAdd("d", 1, fdFecFinCOMP)
    fnMoneda = CInt(rsLog!nMoneda)
    cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, fnMoneda)
    
    If bAdenda Then
        fdFecIni = CDate(rsLog!DesdeC)
        fdFecFin = CDate(rsLog!Hasta)
    Else
        fdFecIni = CDate(rsLog!Desde)
        fdFecFin = CDate(rsLog!Hasta)
    End If
    
    fnNAdenda = oLog.ObtenerUltAdendaContratos(Trim(rsLog!NContrato), fncontRef) + 1
    Me.lblNPago.Caption = "1"
    Me.lblNAdendaCOMP.Caption = fnNAdenda
    Me.lblNAdendaAD.Caption = fnNAdenda
    Me.lblNAdendaRED.Caption = fnNAdenda
    
    Set rsLog = oLog.ObtenerCuotasContratos(fsNContrato, fncontRef)
    Call CargaCuotas(rsLog)
End If
End Sub
Sub VerPestana(ByVal I As Integer)
If I = 0 Then
    Me.SSTContratos.TabVisible(I) = False
    Me.SSTContratos.TabVisible(I + 1) = False
    Me.SSTContratos.TabVisible(I + 2) = False
    'Me.SSTContratos.TabVisible(i + 3) = False
    
    'Me.SSTContratos.TabVisible(i + 3) = True
    Me.SSTContratos.TabVisible(I + 2) = True
    Me.SSTContratos.TabVisible(I + 1) = True
    
    Me.SSTContratos.TabVisible(I + 2) = False
    Me.SSTContratos.TabVisible(I + 1) = False
    Me.SSTContratos.TabVisible(I) = True
    
ElseIf I = 1 Then
    Me.SSTContratos.TabVisible(I) = False
    Me.SSTContratos.TabVisible(I - 1) = False
    Me.SSTContratos.TabVisible(I + 1) = False
    Me.SSTContratos.TabVisible(I) = True
    Me.SSTContratos.TabVisible(I - 1) = True
    Me.SSTContratos.TabVisible(I + 1) = True
    
    Me.SSTContratos.TabVisible(I - 1) = False
    Me.SSTContratos.TabVisible(I + 1) = False
ElseIf I = 2 Then
    Me.SSTContratos.TabVisible(I) = False
    Me.SSTContratos.TabVisible(I - 2) = False
    Me.SSTContratos.TabVisible(I - 1) = False
    Me.SSTContratos.TabVisible(I) = True
    Me.SSTContratos.TabVisible(I - 2) = True
    Me.SSTContratos.TabVisible(I - 1) = True
    
    Me.SSTContratos.TabVisible(I - 2) = False
    Me.SSTContratos.TabVisible(I - 1) = False
End If
    If fntpodocorigen = LogTipoContrato.ContratoServicio And fncontTipoPago = 2 Then
        Me.SSTContratos.TabVisible(3) = False
    End If
End Sub

Private Sub CargaDatosDetalle()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset
Dim bAdenda As Boolean
Set oLog = New DLogGeneral
Dim rsServicio As ADODB.Recordset

Set rsLog = oLog.ObtenerCuotasContratos(fsNContrato, fncontRef) 'fnContRef Agregado PASI20140823 Ti-ERS077-2014
Call CargaCuotas(rsLog)

Set rsLog = oLog.ListarDatosAdenda(fsNContrato, fnNAdenda2, fncontRef) 'fnContRef Agregado PASI20140823 Ti-ERS077-2014
 
If rsLog.RecordCount > 0 Then
    If fnTipo = 1 Then
    'COMPLEMENTARIAS
    Me.lblNContratoCOMP.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorCOMP.Caption = Space(1) & rsLog!Proveedor
    Me.fraContratoCOMP.Enabled = False
    
    Me.lblNAdendaCOMP.Caption = Trim(fnNAdenda2)
    txtFecIniCOMP.value = CDate(rsLog!Desde)
    txtFecFinCOMP.value = CDate(rsLog!Hasta)
    Me.fraAdendaCOMP.Enabled = False
    Me.txtGlosa.Text = Trim(rsLog!cGlosa)
    Me.txtGlosa.Enabled = False
    cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, CInt(Trim(rsLog!nMoneda)))
    cboMonedaCro.Enabled = False
    Me.cmdAgregar.Enabled = False
    Me.cmdQuitar.Enabled = False
    Me.txtMontoCro.Enabled = False
    Me.txtFechaPago.value = CDate(rsLog!Desde)
    
    Set rsLog = oLog.ListaCuotasAdenda(fsNContrato, fnNAdenda2, fncontRef)
    Call LimpiaFlex(Me.feCronograma)
        If rsLog.RecordCount > 0 Then
            For I = 0 To rsLog.RecordCount - 1
            Me.feCronograma.AdicionaFila
            feCronograma.TextMatrix(I + 1, 0) = I + 1
            feCronograma.TextMatrix(I + 1, 1) = Trim(rsLog!nNPago)
            feCronograma.TextMatrix(I + 1, 2) = Format(Trim(rsLog!dFecPago), "dd/mm/yyyy")
            feCronograma.TextMatrix(I + 1, 3) = Trim(rsLog!Moneda)
            feCronograma.TextMatrix(I + 1, 4) = Trim(rsLog!nMonto)
            rsLog.MoveNext
            Next I
        End If
    Me.cmdRegistrarCOMP.Enabled = False
    Me.cmdCancelarCOMP.Enabled = False
    ElseIf fnTipo = 2 Then
    'ADICIONALES
    Me.lblNContratoAD.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorAD.Caption = Space(1) & rsLog!Proveedor
    Me.fraContratoAD.Enabled = False
    
    Me.lblNAdendaAD.Caption = Trim(fnNAdenda2)

    cboCuotaDesdeAD.ListIndex = IndiceListaCombo(cboCuotaDesdeAD, Format(Trim(rsLog!nCuotaDesde), "000"))
    cboCuotaHastaAD.ListIndex = IndiceListaCombo(cboCuotaHastaAD, Format(Trim(rsLog!nCuotaHasta), "000"))

    Me.txtMontoExtra.Text = Format(CDbl(Trim(rsLog!monto)), "#0.00")
    Me.fraAdendaAD.Enabled = False
    
    Me.txtRazonAD.Text = Trim(rsLog!cGlosa)
    Me.txtRazonAD.Enabled = False
    Me.cmdRegistrarAD.Enabled = False
    Me.cmdCancelarAD.Enabled = False
    ElseIf fnTipo = 3 Then
    'REDUCCIONES
    Me.lblNContratoRED.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedorRED.Caption = Space(1) & rsLog!Proveedor
     Me.fraContratoRED.Enabled = False
    
    Me.lblNAdendaRED.Caption = Trim(fnNAdenda2)
    cboCuotaDesdeRED.ListIndex = IndiceListaCombo(cboCuotaDesdeRED, Format(Trim(rsLog!nCuotaDesde), "000"))
    cboCuotaHastaRED.ListIndex = IndiceListaCombo(cboCuotaHastaRED, Format(Trim(rsLog!nCuotaHasta), "000"))
    Me.txtMontoRED.Text = Format(CDbl(Trim(rsLog!monto)), "#0.00")
    Me.fraAdendaRED.Enabled = False
    
    Me.txtRazonRED.Text = Trim(rsLog!cGlosa)
    Me.txtRazonRED.Enabled = False
    Me.cmdRegistrarRED.Enabled = False
    Me.cmdCancelarRED.Enabled = False
    End If
    'PASIERS0772014
    If fntpodocorigen = LogTipoContrato.ContratoServicio Then
        Set rsServicio = oLog.ListaContratoServicioAdendaDet(fsNContrato, fncontRef, fnNAdenda2)
        If Not rsServicio.EOF Then
            Do While Not rsServicio.EOF
                feOrden.AdicionaFila
                feOrden.TextMatrix(feOrden.row, 1) = rsServicio!cAgeDest
                feOrden.TextMatrix(feOrden.row, 2) = rsServicio!cCtaContCod
                feOrden.TextMatrix(feOrden.row, 3) = rsServicio!cDescripcion
                feOrden.TextMatrix(feOrden.row, 6) = rsServicio!nMovImporte
                rsServicio.MoveNext
            Loop
            Me.SSTContratos.TabVisible(3) = True
            cmdAgregarItemCont.Enabled = False
            cmdQuitarItemCont.Enabled = False
        End If
    Else
            Me.SSTContratos.TabVisible(3) = False
    End If
    'END PASI
End If
End Sub

Sub CargaCuotas(ByVal pRs As ADODB.Recordset)
On Error GoTo ErrHandler
    cboCuotaDesdeAD.Clear
    cboCuotaHastaAD.Clear
    cboCuotaDesdeRED.Clear
    cboCuotaHastaRED.Clear

        Do Until pRs.EOF
            Me.cboCuotaDesdeAD.AddItem Format(Trim(pRs!NCuota), "000")
            Me.cboCuotaHastaAD.AddItem Format(Trim(pRs!NCuota), "000")
            Me.cboCuotaDesdeRED.AddItem Format(Trim(pRs!NCuota), "000")
            Me.cboCuotaHastaRED.AddItem Format(Trim(pRs!NCuota), "000")
            pRs.MoveNext
        Loop
    Exit Sub
ErrHandler:
    MsgBox "Error al cargar Cuotas del contrato", vbInformation, "AVISO"
End Sub

'WIOR 20130131 *************************************
Sub GrabarArchivo()
If Trim(fsRuta) <> "" Then
Dim RutaFinal As String
RutaFinal = psRutaContrato
Dim a As New Scripting.FileSystemObject

If a.FolderExists(RutaFinal) = False Then
    a.CreateFolder (RutaFinal)
End If

Copiar fsRuta, RutaFinal & Trim(lblNombreArchivo.Caption)
Else
    MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
End If
End Sub

Private Sub Copiar(Archivo As String, Destino As String)
Dim a As New Scripting.FileSystemObject

If a.FileExists(Destino) = False Then
    a.CopyFile Archivo, Destino
Else
    MsgBox "Archivo ya existe", vbInformation, "Aviso"
End If
End Sub
'WIOR **********************************************
'PASIERS0772014
Private Function DameCtaCont(ByVal psObjeto As String, nNiv As Integer, psAgeCod As String) As String
    Dim oCon As New DConecta
    Dim oForm As New frmLogOCompra
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    
    sSql = oForm.FormaSelect(gsOpeCod, psObjeto, 0, psAgeCod)
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    If Not rs.EOF Then
        DameCtaCont = rs!cObjetoCod
    End If
    Set rs = Nothing
    Set oForm = Nothing
    Set oCon = Nothing
End Function
Private Sub AsignaObjetosSer(ByVal sCtaCod As String)
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As New DConecta
    Dim oCtaCont As New DCtaCont
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim oRHAreas As New DActualizaDatosArea
    Dim oCtaIf As New NCajaCtaIF
    Dim oEfect As New Defectivo
    Dim oDescObj As New ClassDescObjeto
    Dim oContFunct As New NContFunciones
    Dim lsRaiz As String, lsFiltro As String, sSql As String
        
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
    EliminaObjeto feOrden.row

    sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSql)
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOK Then
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                            Else
                                EliminaObjeto feOrden.row
                                Exit Do
                            End If
                        Else
                            AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If

    Set rs = Nothing
    Set rs1 = Nothing
    Set oDescObj = Nothing
    Set oCon = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
    Set oContFunct = Nothing
    Set oContFunct = Nothing
    Exit Sub
End Sub
Private Function validaBusqueda()
     validaBusqueda = True
     If fnTipo = LogtipoReajusteAdenda.Complementaria Then
        If feCronograma.TextMatrix(1, 1) = "" Then
            MsgBox "Ud. primero debe de Ingresar el Monto de la Adenda.", vbInformation, "Aviso"
            validaBusqueda = False
        Exit Function
        End If
     ElseIf fnTipo = LogtipoReajusteAdenda.Adicional Then
         If Len(txtMontoExtra.Text) = 0 Then
            MsgBox "Ud. primero debe de Ingresar el Monto de la Adenda.", vbInformation, "Aviso"
            validaBusqueda = False
        Exit Function
        End If
     ElseIf fnTipo = LogtipoReajusteAdenda.Reduccion Then
        If Len(txtMontoRED.Text) = 0 Then
            MsgBox "Ud. primero debe de Ingresar el Monto de la Adenda.", vbInformation, "Aviso"
            validaBusqueda = False
        Exit Function
        End If
     End If
End Function
Private Sub AdicionaObjeto(ByVal pnItem As Integer, ByVal psCtaObjOrden As String, ByVal psCodigo As String, ByVal psDesc As String, ByVal psFiltro As String, ByVal psObjetoCod As String)
    feObj.AdicionaFila
    feObj.TextMatrix(feObj.row, 1) = pnItem
    feObj.TextMatrix(feObj.row, 2) = psCtaObjOrden
    feObj.TextMatrix(feObj.row, 3) = psCodigo
    feObj.TextMatrix(feObj.row, 4) = psDesc
    feObj.TextMatrix(feObj.row, 5) = psFiltro
    feObj.TextMatrix(feObj.row, 6) = psObjetoCod
End Sub
Private Function OrdenServicio() As ADODB.Recordset
    Dim oCon As New DConecta
    Dim sSqlO As String
    Dim lnMoneda As Integer
    If fnMoneda <> 0 Then
        oCon.AbreConexion
        sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
              & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
              & "WHERE b.cCtaContCod = a.cCtaContCod AND (a.cOpeCod='" & IIf(fnMoneda = 1, "501207", "502207") & "' AND (a.cOpeCtaDH='D'))"
        Set OrdenServicio = oCon.CargaRecordSet(sSqlO)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function
'end PASI
