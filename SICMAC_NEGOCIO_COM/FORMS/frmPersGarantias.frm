VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersGarantias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Garantias de Cliente"
   ClientHeight    =   7560
   ClientLeft      =   2745
   ClientTop       =   2430
   ClientWidth     =   10320
   Icon            =   "frmPersGarantias.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSGarant 
      Height          =   5505
      Left            =   120
      TabIndex        =   30
      Top             =   1320
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   9710
      _Version        =   393216
      Tabs            =   10
      Tab             =   6
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Relac. de la Garantia"
      TabPicture(0)   =   "frmPersGarantias.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraPrinc"
      Tab(0).Control(1)=   "FraRelaGar"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos de Garantia"
      TabPicture(1)   =   "frmPersGarantias.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraZonaCbo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "framontos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Garantia Real"
      TabPicture(2)   =   "frmPersGarantias.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraGar"
      Tab(2).Control(1)=   "FraRRPP"
      Tab(2).Control(2)=   "FraDatInm"
      Tab(2).Control(3)=   "fraDatVehic"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Declaración Jurada"
      TabPicture(3)   =   "frmPersGarantias.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tasación"
      TabPicture(4)   =   "frmPersGarantias.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraTasacion"
      Tab(4).Control(1)=   "Frame3"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tabla de Valores"
      TabPicture(5)   =   "frmPersGarantias.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label40"
      Tab(5).Control(1)=   "Label41"
      Tab(5).Control(2)=   "lblCoberturaCredito"
      Tab(5).Control(3)=   "FeTabla"
      Tab(5).Control(4)=   "cmbTipoCreditoTabla"
      Tab(5).Control(5)=   "cmdImprimir"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Póliza de Inmueble"
      TabPicture(6)   =   "frmPersGarantias.frx":03B2
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "fraZonaCbo"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "framontos"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame2"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Documento de Compra"
      TabPicture(7)   =   "frmPersGarantias.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame6"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Póliza Mobiliaria"
      TabPicture(8)   =   "frmPersGarantias.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame7"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Autoliquidables "
      TabPicture(9)   =   "frmPersGarantias.frx":0406
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame9"
      Tab(9).Control(1)=   "Frame8"
      Tab(9).ControlCount=   2
      Begin VB.Frame frBF 
         Caption         =   "Bien Futuro"
         Height          =   735
         Left            =   -74640
         TabIndex        =   212
         Top             =   4320
         Width           =   8175
         Begin VB.CheckBox ckPolizaBF 
            Caption         =   "Poliza de Bien Futuro"
            Height          =   375
            Left            =   240
            TabIndex        =   213
            Top             =   240
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtFechaPBF 
            Height          =   330
            Left            =   2760
            TabIndex        =   214
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "De la Garantia"
         Height          =   735
         Left            =   -74760
         TabIndex        =   203
         Top             =   720
         Width           =   9375
         Begin VB.ComboBox cboTipoGA 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   240
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtFechaBloqueo 
            Height          =   330
            Left            =   4440
            TabIndex        =   206
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Bloqueo"
            Height          =   195
            Left            =   2640
            TabIndex        =   207
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   240
            TabIndex        =   205
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Inscripcion Registros Publicos"
         Height          =   855
         Left            =   -74760
         TabIndex        =   197
         Top             =   1560
         Width           =   9375
         Begin VB.TextBox txtValorGravadoGA 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   15
            TabIndex        =   199
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   240
            Width           =   1245
         End
         Begin VB.CheckBox chkEnTramiteGA 
            Caption         =   "En Tramite"
            Height          =   195
            Left            =   120
            TabIndex        =   198
            Top             =   360
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txtFechaCertifGravamenGA 
            Height          =   330
            Left            =   3840
            TabIndex        =   200
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Valor Gravado:"
            Height          =   195
            Left            =   5520
            TabIndex        =   202
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Certif.Gravamen:"
            Height          =   195
            Left            =   1920
            TabIndex        =   201
            Top             =   360
            Width           =   1680
         End
      End
      Begin VB.Frame Frame7 
         Height          =   3375
         Left            =   -74640
         TabIndex        =   192
         Top             =   840
         Width           =   8175
         Begin VB.TextBox txtAnioFabricacion 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   195
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   720
            Width           =   1245
         End
         Begin VB.ComboBox cboClaseMueble 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            ItemData        =   "frmPersGarantias.frx":0422
            Left            =   2880
            List            =   "frmPersGarantias.frx":0424
            Style           =   2  'Dropdown List
            TabIndex        =   193
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   240
            Width           =   3285
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Año de Fabricación:"
            Height          =   195
            Left            =   240
            TabIndex        =   196
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Clase de Valor Mobiliario:"
            Height          =   195
            Left            =   240
            TabIndex        =   194
            Top             =   240
            Width           =   1770
         End
      End
      Begin VB.Frame Frame6 
         ClipControls    =   0   'False
         Height          =   4020
         Left            =   -74760
         TabIndex        =   178
         Top             =   750
         Width           =   9315
         Begin VB.ComboBox CmbDocCompra 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   720
            Width           =   3285
         End
         Begin VB.TextBox txtNumDocCompra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   188
            Tag             =   "txtPrincipal"
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtValorDocCompra 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   190
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   1140
            Width           =   1365
         End
         Begin VB.CommandButton cmdBuscaEmisorDoc 
            Caption         =   "..."
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
            Height          =   300
            Left            =   8280
            TabIndex        =   179
            Top             =   360
            Width           =   390
         End
         Begin MSMask.MaskEdBox txtFecEmision 
            Height          =   330
            Left            =   1440
            TabIndex        =   189
            Top             =   1140
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "N° Docum:"
            Height          =   195
            Left            =   4800
            TabIndex        =   187
            Top             =   720
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Docum.      :"
            Height          =   195
            Left            =   120
            TabIndex        =   185
            Top             =   720
            Width           =   1320
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Valor asegurado :"
            Height          =   195
            Left            =   4800
            TabIndex        =   184
            Top             =   1155
            Width           =   1380
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión   :"
            Height          =   195
            Left            =   120
            TabIndex        =   183
            Top             =   1150
            Width           =   1215
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Emisor                :"
            Height          =   195
            Left            =   120
            TabIndex        =   182
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label LblEmisorPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   2805
            TabIndex        =   181
            Top             =   360
            Width           =   5400
         End
         Begin VB.Label LblEmisorPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1440
            TabIndex        =   180
            Top             =   360
            Width           =   1350
         End
      End
      Begin VB.Frame FraGar 
         Caption         =   "De la Garantia"
         Enabled         =   0   'False
         Height          =   1530
         Left            =   -74520
         TabIndex        =   73
         Top             =   2520
         Width           =   8490
         Begin VB.ComboBox CboTipo 
            Height          =   315
            Left            =   7080
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox TxtDirecRegPubli 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   121
            Top             =   1080
            Width           =   4770
         End
         Begin VB.TextBox TxtRegNro 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4770
            TabIndex        =   25
            Top             =   600
            Width           =   1785
         End
         Begin MSMask.MaskEdBox TxtFechareg 
            Height          =   330
            Left            =   1845
            TabIndex        =   24
            Top             =   600
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton CmdBuscaNot 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7980
            TabIndex        =   23
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label20 
            Caption         =   "Numero Operacion"
            Height          =   255
            Left            =   3240
            TabIndex        =   170
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   6645
            TabIndex        =   124
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Direcc Reg. Publicos :"
            Height          =   195
            Left            =   105
            TabIndex        =   122
            Top             =   1080
            Width           =   1590
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "de Registro :"
            Height          =   195
            Left            =   3480
            TabIndex        =   78
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Registro    :"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label LblNotaPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1845
            TabIndex        =   76
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label LblNotaPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3210
            TabIndex        =   75
            Top             =   240
            Width           =   4725
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Notaria                       :"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   1590
         End
      End
      Begin VB.Frame FraRRPP 
         Caption         =   "Inscripcion Registros Publicos"
         Height          =   930
         Left            =   -74520
         TabIndex        =   165
         Top             =   4080
         Width           =   8535
         Begin VB.ComboBox cboTpoInscripcion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox chkEnTramite 
            Caption         =   "En Tramite"
            Height          =   195
            Left            =   120
            TabIndex        =   171
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtValorGravado 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   167
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   480
            Width           =   1245
         End
         Begin MSMask.MaskEdBox txtFechaCertifGravamen 
            Height          =   330
            Left            =   3240
            TabIndex        =   166
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Inscripción:"
            Height          =   195
            Left            =   1560
            TabIndex        =   210
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Certif.Gravamen:"
            Height          =   195
            Left            =   3000
            TabIndex        =   169
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Valor Gravado:"
            Height          =   195
            Left            =   4800
            TabIndex        =   168
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3375
         Left            =   -74640
         TabIndex        =   134
         Top             =   840
         Width           =   8175
         Begin VB.ComboBox cmbClaseInmueble 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            ItemData        =   "frmPersGarantias.frx":0426
            Left            =   2880
            List            =   "frmPersGarantias.frx":0428
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   240
            Width           =   3285
         End
         Begin VB.ComboBox cmbCategoria 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            ItemData        =   "frmPersGarantias.frx":042A
            Left            =   2880
            List            =   "frmPersGarantias.frx":042C
            Style           =   2  'Dropdown List
            TabIndex        =   136
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   600
            Width           =   3285
         End
         Begin VB.TextBox txtNumLocales 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   137
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   1080
            Width           =   1245
         End
         Begin VB.TextBox txtNumPisos 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   138
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   1560
            Width           =   1245
         End
         Begin VB.TextBox txtNumSotanos 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   139
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   2040
            Width           =   1245
         End
         Begin VB.TextBox txtAnioConstruccion 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   140
            Tag             =   "txtPrincipal"
            Text            =   "0"
            Top             =   2520
            Width           =   1245
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Clase de Inmueble:"
            Height          =   195
            Left            =   240
            TabIndex        =   146
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Categoría:"
            Height          =   195
            Left            =   240
            TabIndex        =   145
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Locales:"
            Height          =   195
            Left            =   240
            TabIndex        =   144
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Pisos:"
            Height          =   195
            Left            =   240
            TabIndex        =   143
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Sótanos:"
            Height          =   195
            Left            =   240
            TabIndex        =   142
            Top             =   2160
            Width           =   1080
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Año de Construcción:"
            Height          =   195
            Left            =   240
            TabIndex        =   141
            Top             =   2640
            Width           =   1530
         End
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   -74775
         TabIndex        =   118
         Top             =   4410
         Width           =   1125
      End
      Begin VB.ComboBox cmbTipoCreditoTabla 
         Height          =   315
         Left            =   -73245
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   765
         Width           =   3750
      End
      Begin SICMACT.FlexEdit FeTabla 
         Height          =   3165
         Left            =   -74775
         TabIndex        =   113
         Top             =   1170
         Width           =   8610
         _extentx        =   15187
         _extenty        =   5583
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "nTipoEval-Descripcion-Valor-Puntaje"
         encabezadosanchos=   "0-4600-2840-1000"
         font            =   "frmPersGarantias.frx":042E
         font            =   "frmPersGarantias.frx":045A
         font            =   "frmPersGarantias.frx":0486
         font            =   "frmPersGarantias.frx":04B2
         font            =   "frmPersGarantias.frx":04DE
         fontfixed       =   "frmPersGarantias.frx":050A
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-2-X"
         listacontroles  =   "0-0-3-0"
         encabezadosalineacion=   "C-L-L-C"
         formatosedit    =   "0-0-0-0"
         textarray0      =   "nTipoEval"
         lbeditarflex    =   -1
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Frame Frame1 
         Caption         =   "Detalle de Declaración Jurada"
         Height          =   3840
         Left            =   -74880
         TabIndex        =   81
         Top             =   900
         Width           =   8700
         Begin VB.CommandButton CmdImprimirDJ 
            Caption         =   "&Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4080
            TabIndex        =   120
            Top             =   3360
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton CmdDJEliminar 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7440
            TabIndex        =   82
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   3360
            Width           =   1005
         End
         Begin VB.CommandButton CmdDJNuevo 
            Caption         =   "&Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6360
            TabIndex        =   83
            Top             =   3360
            Width           =   1005
         End
         Begin VB.CommandButton CmdDJAceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6360
            TabIndex        =   85
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1005
         End
         Begin SICMACT.FlexEdit FEDeclaracionJur 
            Height          =   2955
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   8460
            _extentx        =   14923
            _extenty        =   5212
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "-Descripción-Cantidad-Valor Actual-Tipo Doc.-Nro Doc.-Aux"
            encabezadosanchos=   "400-5000-1450-1450-1450-1450-0"
            font            =   "frmPersGarantias.frx":0538
            font            =   "frmPersGarantias.frx":0564
            font            =   "frmPersGarantias.frx":0590
            font            =   "frmPersGarantias.frx":05BC
            font            =   "frmPersGarantias.frx":05E8
            fontfixed       =   "frmPersGarantias.frx":0614
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-1-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-3-0-0"
            encabezadosalineacion=   "C-L-R-R-L-L-C"
            formatosedit    =   "0-0-3-2-0-0-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   405
            rowheight0      =   300
         End
         Begin VB.CommandButton CmdDJCancelar 
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
            Height          =   300
            Left            =   7440
            TabIndex        =   86
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label LblTotDJ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   2520
            TabIndex        =   88
            Top             =   3360
            Width           =   1440
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Total Declaración Jurada:"
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
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   3360
            Width           =   2220
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   90
         TabIndex        =   55
         Top             =   2625
         Width           =   8910
         Begin VB.CheckBox chkGarPolizaMob 
            Caption         =   "Póliza Mobiliaria"
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
            Height          =   195
            Left            =   2160
            TabIndex        =   191
            Top             =   2280
            Width           =   1980
         End
         Begin VB.CheckBox chkDocCompra 
            Caption         =   "Con Doc.Compra"
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
            Height          =   195
            Left            =   2160
            TabIndex        =   177
            Top             =   1900
            Width           =   1860
         End
         Begin VB.CheckBox ChkTasacion 
            Caption         =   "Con Tasación"
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
            Height          =   195
            Left            =   120
            TabIndex        =   175
            Top             =   1900
            Width           =   1500
         End
         Begin VB.CheckBox ChkGarPoliza 
            Caption         =   "Póliza de Inmueble"
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
            Left            =   120
            TabIndex        =   149
            Top             =   2280
            Width           =   1980
         End
         Begin VB.CheckBox ChkCF 
            Caption         =   "Carta Fianza"
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
            Height          =   195
            Left            =   2160
            TabIndex        =   119
            Top             =   1500
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox txtcomentarios 
            Height          =   480
            Left            =   900
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   106
            Top             =   157
            Width           =   7845
         End
         Begin VB.Frame FraClase 
            Caption         =   "Clase de Garantia"
            Height          =   615
            Left            =   135
            TabIndex        =   57
            Top             =   690
            Width           =   4005
            Begin VB.OptionButton OptCG 
               Caption         =   "Garantia No Preferida"
               Height          =   240
               Index           =   0
               Left            =   105
               TabIndex        =   14
               Top             =   255
               Value           =   -1  'True
               Width           =   1905
            End
            Begin VB.OptionButton OptCG 
               Caption         =   "Garantia Preferida"
               Height          =   240
               Index           =   1
               Left            =   2025
               TabIndex        =   15
               Top             =   255
               Width           =   1650
            End
         End
         Begin VB.Frame FraTipoRea 
            Caption         =   "Tipo de Realizacion"
            Height          =   615
            Left            =   4365
            TabIndex        =   56
            Top             =   690
            Width           =   4440
            Begin VB.OptionButton OptTR 
               Caption         =   "De Rapida Realizacion"
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2130
               TabIndex        =   17
               Top             =   270
               Width           =   1980
            End
            Begin VB.OptionButton OptTR 
               Caption         =   "De Lenta Realizacion"
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   90
               TabIndex        =   16
               Top             =   255
               Value           =   -1  'True
               Width           =   1950
            End
         End
         Begin VB.CheckBox ChkGarReal 
            Caption         =   "Garantia Real"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1500
            Width           =   1620
         End
         Begin VB.ComboBox CboBanco 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1860
            Width           =   4170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Comentario"
            Height          =   195
            Left            =   60
            TabIndex        =   107
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   4080
            TabIndex        =   58
            Top             =   1890
            Width           =   465
         End
      End
      Begin VB.Frame framontos 
         Height          =   1740
         Left            =   6225
         TabIndex        =   50
         Top             =   810
         Width           =   2775
         Begin VB.TextBox txtMontoxGrav 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   13
            Tag             =   "txtPrincipal"
            Top             =   1215
            Width           =   1185
         End
         Begin VB.TextBox txtMontotas 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   11
            Tag             =   "txtPrincipal"
            Top             =   420
            Width           =   1185
         End
         Begin VB.TextBox txtMontoRea 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   12
            Tag             =   "txtPrincipal"
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Montos"
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
            Index           =   3
            Left            =   870
            TabIndex        =   54
            ToolTipText     =   "Monto Tasación"
            Top             =   15
            Width           =   630
         End
         Begin VB.Label lblMontoGrav 
            AutoSize        =   -1  'True
            Caption         =   "Disponible :"
            Height          =   195
            Left            =   420
            TabIndex        =   53
            Top             =   1245
            Width           =   825
         End
         Begin VB.Label lblrealizacion 
            AutoSize        =   -1  'True
            Caption         =   "Realización :"
            Height          =   195
            Left            =   360
            TabIndex        =   52
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lbltasa 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial :"
            Height          =   195
            Left            =   75
            TabIndex        =   51
            ToolTipText     =   "Monto Tasación"
            Top             =   465
            Width           =   1185
         End
      End
      Begin VB.Frame fraZonaCbo 
         Height          =   1725
         Left            =   105
         TabIndex        =   43
         Top             =   825
         Width           =   6060
         Begin VB.TextBox txtDireccion 
            Height          =   315
            Left            =   900
            TabIndex        =   109
            Top             =   1200
            Width           =   4995
         End
         Begin VB.Frame frazona 
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   330
            TabIndex        =   44
            Top             =   210
            Width           =   5670
            Begin VB.ComboBox cmbPersUbiGeo 
               Height          =   315
               Index           =   0
               Left            =   570
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Zona"
               Top             =   90
               Width           =   1920
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               Height          =   315
               Index           =   2
               Left            =   585
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Provincia"
               Top             =   450
               Width           =   1935
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               Height          =   315
               Index           =   1
               Left            =   3210
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Distrito"
               Top             =   90
               Width           =   1980
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               ForeColor       =   &H80000007&
               Height          =   315
               Index           =   3
               Left            =   3210
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Urbanización"
               Top             =   450
               Width           =   1995
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Dpto :"
               Height          =   195
               Left            =   15
               TabIndex        =   48
               Top             =   150
               Width           =   435
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Prov :"
               Height          =   195
               Left            =   2625
               TabIndex        =   47
               Top             =   150
               Width           =   420
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Distrito :"
               Height          =   195
               Left            =   -15
               TabIndex        =   46
               Top             =   510
               Width           =   570
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Zona :"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   2625
               TabIndex        =   45
               Top             =   510
               Width           =   465
            End
         End
         Begin VB.Label Label34 
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   60
            TabIndex        =   108
            Top             =   1260
            Width           =   795
         End
         Begin VB.Line Line2 
            X1              =   60
            X2              =   5940
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label1 
            Caption         =   " Zona :"
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
            Left            =   195
            TabIndex        =   49
            Top             =   15
            Width           =   555
         End
      End
      Begin VB.Frame fraPrinc 
         Height          =   2175
         Left            =   -74910
         TabIndex        =   35
         ToolTipText     =   "Datos del Cliente"
         Top             =   675
         Width           =   8730
         Begin VB.ComboBox CboNumPF 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   960
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   1485
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo Moneda"
            Top             =   1800
            Width           =   2115
         End
         Begin VB.ComboBox CboGarantia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   570
            Width           =   2715
         End
         Begin VB.TextBox txtDescGarant 
            Height          =   330
            Left            =   1515
            MaxLength       =   60
            TabIndex        =   4
            Tag             =   "txtPrincipal"
            Top             =   1335
            Width           =   5760
         End
         Begin VB.ComboBox CmbTipoGarant 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   570
            Width           =   3225
         End
         Begin VB.ComboBox CmbDocGarant 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1515
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   945
            Width           =   3285
         End
         Begin VB.TextBox txtNumDoc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5730
            MaxLength       =   18
            TabIndex        =   3
            Tag             =   "txtPrincipal"
            Top             =   930
            Width           =   2970
         End
         Begin VB.CommandButton CmdBuscaEmisor 
            Caption         =   "..."
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
            Height          =   300
            Left            =   8265
            TabIndex        =   0
            Top             =   255
            Width           =   390
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   750
            TabIndex        =   105
            Top             =   1860
            Width           =   675
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación"
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
            Left            =   240
            TabIndex        =   103
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   5070
            TabIndex        =   80
            Top             =   1800
            Width           =   2205
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Estado          :"
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
            Left            =   3690
            TabIndex        =   79
            Top             =   1860
            Width           =   1260
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   225
            TabIndex        =   42
            Top             =   1395
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Garantía"
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
            Left            =   4230
            TabIndex        =   41
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Garantía"
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
            Left            =   210
            TabIndex        =   40
            Top             =   1005
            Width           =   1230
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. :"
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
            Left            =   4980
            TabIndex        =   39
            Top             =   1005
            Width           =   810
         End
         Begin VB.Label LblEmisor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   2895
            TabIndex        =   38
            Top             =   270
            Width           =   5340
         End
         Begin VB.Label LblPersCodEmi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1515
            TabIndex        =   37
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Emisor :"
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
            Left            =   225
            TabIndex        =   36
            Top             =   300
            Width           =   690
         End
      End
      Begin VB.Frame FraRelaGar 
         Caption         =   "Representantes Garantia"
         Height          =   1860
         Left            =   -74910
         TabIndex        =   31
         Top             =   2880
         Width           =   8760
         Begin VB.CommandButton CmdCliEliminar 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7560
            TabIndex        =   6
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   1440
            Width           =   1005
         End
         Begin VB.CommandButton CmdCliNuevo 
            Caption         =   "&Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6480
            TabIndex        =   5
            Top             =   1440
            Width           =   1005
         End
         Begin SICMACT.FlexEdit FERelPers 
            Height          =   1110
            Left            =   135
            TabIndex        =   32
            Top             =   270
            Width           =   8520
            _extentx        =   15240
            _extenty        =   2037
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "-Codigo-Nombre-Relacion-Aux"
            encabezadosanchos=   "400-1500-5000-1450-0"
            font            =   "frmPersGarantias.frx":0642
            font            =   "frmPersGarantias.frx":066E
            font            =   "frmPersGarantias.frx":069A
            font            =   "frmPersGarantias.frx":06C6
            font            =   "frmPersGarantias.frx":06F2
            fontfixed       =   "frmPersGarantias.frx":071E
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-1-X-3-X"
            listacontroles  =   "0-1-0-3-0"
            encabezadosalineacion=   "C-C-L-L-C"
            formatosedit    =   "0-0-0-0-0"
            lbeditarflex    =   -1
            colwidth0       =   405
            rowheight0      =   300
         End
         Begin VB.CommandButton CmdCliAceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6480
            TabIndex        =   33
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton CmdCliCancelar 
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
            Height          =   300
            Left            =   7560
            TabIndex        =   34
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1005
         End
      End
      Begin VB.Frame FraDatInm 
         Caption         =   "Datos del Inmueble"
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   1305
         Left            =   -74520
         TabIndex        =   67
         Top             =   840
         Width           =   8475
         Begin VB.ComboBox CboTipoInmueb 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4890
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   660
            Width           =   3390
         End
         Begin VB.TextBox TxtTelefono 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1905
            TabIndex        =   21
            Top             =   690
            Width           =   1395
         End
         Begin VB.CommandButton CmdBuscaInmob 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7800
            TabIndex        =   20
            Top             =   285
            Width           =   390
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Inmueble :"
            Height          =   195
            Left            =   3435
            TabIndex        =   72
            Top             =   705
            Width           =   1320
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Telefono/Inmobiliaria :"
            Height          =   195
            Left            =   180
            TabIndex        =   71
            Top             =   705
            Width           =   1575
         End
         Begin VB.Label LblInmobNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3270
            TabIndex        =   69
            Top             =   300
            Width           =   4500
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Propietario del Bien :"
            Height          =   195
            Left            =   180
            TabIndex        =   68
            Top             =   315
            Width           =   1455
         End
         Begin VB.Label LblInmobCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1905
            TabIndex        =   70
            Top             =   300
            Width           =   1350
         End
      End
      Begin VB.Frame fraDatVehic 
         Caption         =   "Datos del Vehículo"
         Height          =   1695
         Left            =   -74520
         TabIndex        =   110
         Top             =   720
         Visible         =   0   'False
         Width           =   8400
         Begin VB.TextBox txtNumSerie 
            Height          =   285
            Left            =   3720
            TabIndex        =   159
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtNumMotor 
            Height          =   285
            Left            =   1080
            TabIndex        =   158
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtDescrip 
            Height          =   285
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   156
            Top             =   240
            Width           =   5175
         End
         Begin VB.TextBox txtTipoMerca 
            Height          =   285
            Left            =   1080
            TabIndex        =   154
            Top             =   1320
            Width           =   7215
         End
         Begin VB.TextBox txtDireAlma 
            Height          =   285
            Left            =   3720
            TabIndex        =   152
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox txtValorMerca 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   150
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   960
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtPlacaVehic 
            Height          =   315
            Left            =   6960
            TabIndex        =   112
            Top             =   240
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtAnioFab 
            Height          =   330
            Left            =   6960
            TabIndex        =   160
            Top             =   600
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblAnioFab 
            AutoSize        =   -1  'True
            Caption         =   "Año fabricación:"
            Height          =   195
            Left            =   5640
            TabIndex        =   163
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lblNumSerie 
            AutoSize        =   -1  'True
            Caption         =   "Nº serie :"
            Height          =   195
            Left            =   3000
            TabIndex        =   162
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblNumMotor 
            AutoSize        =   -1  'True
            Caption         =   "Nº motor:"
            Height          =   195
            Left            =   120
            TabIndex        =   161
            Top             =   600
            Width           =   660
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   157
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lblTipoMerca 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Mercad:"
            Height          =   195
            Left            =   120
            TabIndex        =   155
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label lblDireAlma 
            AutoSize        =   -1  'True
            Caption         =   "Direcc. almacén :"
            Height          =   195
            Left            =   2400
            TabIndex        =   153
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label lblValorMerca 
            AutoSize        =   -1  'True
            Caption         =   "Valor :"
            Height          =   195
            Left            =   360
            TabIndex        =   151
            Top             =   960
            Width           =   450
         End
         Begin VB.Label lblPlacaVehic 
            AutoSize        =   -1  'True
            Caption         =   "Placa :"
            Height          =   195
            Left            =   6360
            TabIndex        =   111
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraTasacion 
         ClipControls    =   0   'False
         Height          =   4020
         Left            =   -74760
         TabIndex        =   125
         Top             =   750
         Width           =   9315
         Begin SICMACT.FlexEdit FETasacion 
            Height          =   2295
            Left            =   240
            TabIndex        =   164
            Top             =   1440
            Width           =   8895
            _extentx        =   15055
            _extenty        =   2566
            cols0           =   6
            encabezadosnombres=   "-Codigo-Nombre-Fech Tasacion-VRM-Valor Edifc o Valor Aseg"
            encabezadosanchos=   "400-1200-3500-1200-1200-1200"
            font            =   "frmPersGarantias.frx":074C
            font            =   "frmPersGarantias.frx":0778
            font            =   "frmPersGarantias.frx":07A4
            font            =   "frmPersGarantias.frx":07D0
            font            =   "frmPersGarantias.frx":07FC
            fontfixed       =   "frmPersGarantias.frx":0828
            columnasaeditar =   "X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-C-C-C"
            formatosedit    =   "0-0-0-0-0-0"
            colwidth0       =   405
            rowheight0      =   300
         End
         Begin VB.TextBox txtVRM 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            MaxLength       =   15
            TabIndex        =   147
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   600
            Width           =   1245
         End
         Begin VB.CommandButton CmdBuscaTasa 
            Caption         =   "..."
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
            Height          =   300
            Left            =   8400
            TabIndex        =   127
            Top             =   240
            Width           =   390
         End
         Begin VB.TextBox TxtValorEdificacion 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            MaxLength       =   15
            TabIndex        =   126
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   600
            Width           =   1245
         End
         Begin MSMask.MaskEdBox TxtFecTas 
            Height          =   330
            Left            =   1440
            TabIndex        =   128
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "V.R.M. :"
            Height          =   195
            Left            =   2880
            TabIndex        =   148
            Top             =   600
            Width           =   585
         End
         Begin VB.Label LblTasaPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1440
            TabIndex        =   133
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label LblTasaPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   2805
            TabIndex        =   132
            Top             =   240
            Width           =   5400
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tasador            :"
            Height          =   195
            Left            =   120
            TabIndex        =   131
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Fech Tasacion :"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Valor Edificación ó valor asegurado :"
            Height          =   390
            Left            =   5520
            TabIndex        =   129
            Top             =   600
            Width           =   1380
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4065
         Left            =   -74850
         TabIndex        =   89
         Top             =   750
         Visible         =   0   'False
         Width           =   9285
         Begin VB.TextBox TxtMontoPol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1815
            TabIndex        =   98
            Top             =   1215
            Width           =   1395
         End
         Begin VB.TextBox TxtNroPoliza 
            Height          =   285
            Left            =   1815
            TabIndex        =   94
            Top             =   810
            Width           =   1395
         End
         Begin VB.CommandButton CmdBuscaSeg 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7725
            TabIndex        =   90
            Top             =   375
            Width           =   390
         End
         Begin MSMask.MaskEdBox TxtFecVig 
            Height          =   330
            Left            =   5115
            TabIndex        =   96
            Top             =   825
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtFecCons 
            Height          =   330
            Left            =   1815
            TabIndex        =   100
            Top             =   1635
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "F. Constitucion     :"
            Height          =   195
            Left            =   285
            TabIndex        =   101
            Top             =   1665
            Width           =   1320
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Monto Poliza        :"
            Height          =   195
            Left            =   285
            TabIndex        =   99
            Top             =   1230
            Width           =   1320
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vigencia       :"
            Height          =   195
            Left            =   3480
            TabIndex        =   97
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nro Poliza             :"
            Height          =   195
            Left            =   285
            TabIndex        =   95
            Top             =   825
            Width           =   1350
         End
         Begin VB.Label LblSegPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1830
            TabIndex        =   93
            Top             =   390
            Width           =   1350
         End
         Begin VB.Label LblSegPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3195
            TabIndex        =   92
            Top             =   390
            Width           =   4500
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Aseguradora         :"
            Height          =   195
            Left            =   285
            TabIndex        =   91
            Top             =   405
            Width           =   1350
         End
      End
      Begin VB.Label lblCoberturaCredito 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   240
         Left            =   -67575
         TabIndex        =   117
         Top             =   4500
         Width           =   1320
      End
      Begin VB.Label Label41 
         Caption         =   "Cobertura de Credito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69600
         TabIndex        =   116
         Top             =   4500
         Width           =   1905
      End
      Begin VB.Label Label40 
         Caption         =   "Tipo de Credito:"
         Height          =   285
         Left            =   -74595
         TabIndex        =   114
         Top             =   810
         Width           =   1185
      End
   End
   Begin VB.Frame FraBuscaPers 
      Height          =   1275
      Left            =   75
      TabIndex        =   26
      Top             =   -30
      Width           =   10140
      Begin VB.CommandButton cmdVerCred 
         Caption         =   "&Ver Créditos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6710
         TabIndex        =   215
         ToolTipText     =   "Pulse este Boton para Mostrar los Datos de la Garantia"
         Top             =   810
         Width           =   1440
      End
      Begin VB.CommandButton cmdBusGar 
         Caption         =   "Busca"
         Height          =   300
         Left            =   8400
         TabIndex        =   173
         ToolTipText     =   "Busca por número de garantía"
         Top             =   800
         Width           =   1215
      End
      Begin VB.TextBox txtNumGar 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         MaxLength       =   8
         TabIndex        =   172
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   440
         Width           =   1245
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6710
         TabIndex        =   29
         ToolTipText     =   "Pulse este Boton para Mostrar los Datos de la Garantia"
         Top             =   490
         Width           =   1440
      End
      Begin VB.CommandButton CmdBuscaPersona 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6705
         TabIndex        =   28
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   180
         Width           =   1440
      End
      Begin MSComctlLib.ListView LstGaratias 
         Height          =   975
         Left            =   90
         TabIndex        =   27
         Top             =   165
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Garantia"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "codemi"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "nomemi"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "tipodoc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "cnumdoc"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Num.Garantía:"
         Height          =   195
         Left            =   8400
         TabIndex        =   174
         ToolTipText     =   "Monto Tasación"
         Top             =   200
         Width           =   1050
      End
   End
   Begin VB.Frame Frame5 
      Height          =   660
      Left            =   120
      TabIndex        =   59
      Top             =   6840
      Width           =   9880
      Begin VB.CommandButton cmdActAdmCred 
         Caption         =   "Act&ualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5310
         TabIndex        =   211
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar &Garant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3480
         TabIndex        =   208
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
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
         Height          =   390
         Left            =   1230
         TabIndex        =   63
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8640
         TabIndex        =   62
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
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
         Height          =   390
         Left            =   2355
         TabIndex        =   61
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdLimpiar 
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
         Height          =   390
         Left            =   7520
         TabIndex        =   60
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   390
         Left            =   1215
         TabIndex        =   65
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   66
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   60
         TabIndex        =   64
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmPersGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmPersGarantias
'***     Descripcion:       Realiza el Mantenimiento y Registro de Nuevas Garantias
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:         08/06/2001 12:15:13 PM
'***     Ultima Modificacion: Creacion del Formulario
'*****************************************************************************************

Option Explicit
Private Enum TGarantiaTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum
Dim lnTipoGarantiaActual As String
'Enum TGarantiaTipoInicio
'    RegistroGarantia = 1
'    MantenimientoGarantia = 2
'    ConsultaGarant = 3
'End Enum

Public pgcCtaCod As String

Dim Nivel1() As String
Dim ContNiv1 As Long
Dim Nivel2() As String
Dim ContNiv2 As Long
Dim Nivel3() As String
Dim ContNiv3 As Long
Dim Nivel4() As String
Dim ContNiv4 As Long
Dim bEstadoCargando As Boolean
Dim cmdEjecutar As Integer

Dim vTipoInicio As TGarantiaTipoInicio
Dim sNumgarant As String
Dim bCarga As Boolean
Dim bAsignadoACredito As Boolean

'Agregado por LMMD
Dim bCreditoCF As Boolean
Dim bValdiCCF As Boolean

Dim gcPermiteModificar As Boolean 'peac 20071128
Dim lcGar As String 'peac 20071128
Dim objPista As COMManejador.Pista

Dim nxgravar As Double 'madm 20100513
Dim gGarantiaDepPlazoFijoCF As Boolean 'madm 20100817

Dim fbGrupoAdmCred As Boolean 'WIOR 20130122
Dim fnTpoPoliza As Integer 'WIOR 20130122
'*** PEAC 20090904
'SE MODIFICO LA PROPIEDAD MAXLENGTH DEL CONTOL TxtDirecRegPubli A 50
'*** FIN PEAC
Private fbOrigenCF As Boolean 'WIOR 20140628
Private fnCampanaCred As Long 'WIOR 20140628
Private fsGrupoActGarDPF As String 'WIOR 20150608

Public Sub inicio(ByVal pvTipoIni As TGarantiaTipoInicio, Optional ByVal psCodGarantia As String = "", Optional ByVal pbOrigenCF As Boolean = False, Optional ByVal pnCampanaCred As Long = 0)
'WIOR 20130419 AGREGO psCodGarantia
    'WIOR 20140826 AGREGO pbOrigenCF, pnCampanaCred
    fbOrigenCF = pbOrigenCF 'WIOR 20140826
    fnCampanaCred = pnCampanaCred 'WIOR 20140912
    
    vTipoInicio = pvTipoIni
    If vTipoInicio = ConsultaGarant Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
        CmbDocGarant.Enabled = False
        txtNumDoc.Enabled = False
        'WIOR 20130419 **********************************
        If Trim(psCodGarantia) <> "" Then
            txtNumGar.Text = psCodGarantia
            Call cmdBusGar_Click
            txtNumGar.Enabled = False
            cmdBusGar.Enabled = False
            LstGaratias.Enabled = False
            CmdBuscaPersona.Enabled = False
            cmdBuscar.Enabled = False
            cmdLimpiar.Enabled = False
        End If
        'WIOR FIN ***************************************
    End If
    
    If vTipoInicio = RegistroGarantia Then
        cmdNuevo.Enabled = True
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    
    If vTipoInicio = MantenimientoGarantia Then
        cmdNuevo.Enabled = False
        ' AQUI Napo deberia inabilitar
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
        
    End If
    
    Me.Show 1
End Sub

Private Function ValidaBuscar() As Boolean
    ValidaBuscar = True
    
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de Documento de Garantia", vbInformation, "Aviso"
        ValidaBuscar = False
        Exit Function
    End If
    
    If Trim(txtNumDoc.Text) = "" Then
        MsgBox "Ingrese el Numero de Documento", vbInformation, "Aviso"
        ValidaBuscar = False
        Exit Function
    End If
    
End Function

Private Function ValidaDatos() As Boolean
Dim i As Long
Dim Enc As Boolean
Dim oGarantia As COMNCredito.NCOMGarantia
Dim nValor As Double
Dim sCad As String
'Dim odGarantia As COMDCredito.DCOMGarantia
Dim bValidaCta As Boolean
Dim nValorPorcen As Double
Dim nValorCuota As Double

    ValidaDatos = True
    
    Set oGarantia = New COMNCredito.NCOMGarantia
    'MADM 20100513
    If Not Me.CboNumPF.Visible Then
        Call oGarantia.ValidaDatosGarantia(txtNumDoc.Text, bValidaCta, Trim(Right(CmbTipoGarant.Text, 10)), nValorCuota, nValorPorcen)
    Else
        Call oGarantia.ValidaDatosGarantia(Trim(Left(Me.CboNumPF.Text, 18)), bValidaCta, Trim(Right(CmbTipoGarant.Text, 10)), nValorCuota, nValorPorcen)
    End If
    Set oGarantia = Nothing
    'Valida cuando es una garantia de Plazo Fijo o CTS
    If Trim(Right(CmbTipoGarant, 3)) = "6" Then
        'Set odGarantia = New COMDCredito.DCOMGarantia
        'If odGarantia.ValidaPFCTS(txtNumDoc) = False Then
        If bValidaCta = False Then
            'By Capi 20082008 solo se modifico mensaje
            MsgBox "La cuenta de Plazo Fijo o Cts no es valida o se encuentra en estado CANCELADO", vbInformation, "AVISO"
            ValidaDatos = False
            Exit Function
        End If
    End If
    If SSGarant.TabVisible(3) = True Then
        If FEDeclaracionJur.Rows = 1 Then
           MsgBox "Debe digitar la declaracion jurada", vbInformation, "AVISO"
           ValidaDatos = False
           Exit Function
        End If
    End If
    
    If SSGarant.TabVisible(3) = True Then
        If CDbl(LblTotDJ.Caption) <> CDbl(txtMontoRea) Then
            MsgBox "El monto de la realizacion no coincide con el monto de la declaracion jurada", vbInformation, "AVISO"
            ValidaDatos = False
            Exit Function
        End If
    End If
    
    'Verifica seleccion de Documento de Garantia
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Documento de Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    ' verifica seleccion de super tipo de garantia
    If CboGarantia.ListIndex = -1 Then
        MsgBox "Seleccione tipo Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Ingreso de Numero de Documento de Garantia
    If Not Me.CboNumPF.Visible Then
        If Trim(txtNumDoc.Text) = "" Then
            MsgBox "Ingrese el Numero de Documento de la Garantia", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    End If
    
    'Verifica seleccion de Tipo de Garantia
    If CmbTipoGarant.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica seleccion de Moneda
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la Moneda", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica la Zona
    If cmbPersUbiGeo(3).ListIndex = -1 Then
        MsgBox "Seleccione La Zona donde se Ubica la Garantia", vbInformation, "Aviso"
        SSGarant.Tab = 1
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Monto de Tasacion
    If Trim(txtMontotas.Text) = "" Or Trim(txtMontotas.Text) = "0.00" Then
        MsgBox "El Monto de Tasacion debe ser Mayor que Cero", vbInformation, "Aviso"
        SSGarant.Tab = 1
        txtMontotas.SetFocus
        ValidaDatos = False
        Exit Function
    End If

    'Verifica Monto de Realizacion
    If Trim(txtMontoRea.Text) = "" Or Trim(txtMontoRea.Text) = "0.00" Then
        MsgBox "El Monto de Realizacion debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Monto de Realizacion
    If Trim(txtMontoxGrav.Text) = "" Or Trim(txtMontoxGrav.Text) = "0.00" Then
        MsgBox "El Monto de Disponible debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
   Enc = False
   ' Verifica Existencia de Titular de la Garantia
   For i = 1 To FERelPers.Rows - 1
        If Trim(Right(FERelPers.TextMatrix(i, 3), 15)) <> "" Then
            If CInt(Trim(Right(FERelPers.TextMatrix(i, 3), 15))) = gPersRelGarantiaTitular Then
                Enc = True
                Exit For
            End If
        End If
   Next i
   If Not Enc Then
        MsgBox "Ingrese un Titular para la Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        CmdCliNuevo.SetFocus
        Exit Function
   End If
   
   '*** PEAC 20080412
    If Trim(Right(CmbDocGarant, 4)) = "93" Then
    'If Trim(Right(CmbDocGarant, 4)) = "15" Then
        If FEDeclaracionJur.Rows = 2 And FEDeclaracionJur.TextMatrix(1, 1) = "" Then
            MsgBox "Falta digitar la declaracion jurada", vbInformation, "AVISO"
            ValidaDatos = False
            Exit Function
        End If
    End If
   
    ' CMACICA_CSTS - 25/11/2003 -------------------------------------------------
    If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
        '*** PEAC 20080412
        If CInt(Trim(Right(CmbDocGarant, 10))) = 93 Then
        'If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then
            If Trim(FEDeclaracionJur.TextMatrix(i, 1)) = "" Then
               MsgBox "Ingrese el Detalle de la Declaración Jurada", vbInformation, "Aviso"
               ValidaDatos = False
               SSGarant.Tab = 3
               CmdDJNuevo.SetFocus
               Exit Function
            End If
        End If
    End If
'----------------------------------------------------------------------------------

'    ' CMACICA_CSTS - 05/12/2003 ----------------------------------------------------------------------------
'    If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
'        If LblTotDJ.Caption <> "" Then
'            If CDbl(txtMontoRea.Text) > CDbl(LblTotDJ.Caption) Then
'                MsgBox "El Monto de Realización no puede ser mayor al Total de la Declaración Jurada. ", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'    End If
'------------------------------------------------------------------------------------------------------------
    
   'Validacion de Bancos
   'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Or CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaDepositosGarantia Then
   
'arcv 19-07-2006
'   If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Then
'        If CboBanco.ListIndex = -1 Then
'            MsgBox "Seleccione un Banco", vbInformation, "Aviso"
'            SSGarant.Tab = 1
'            CboBanco.SetFocus
'            ValidaDatos = False
'            Exit Function
'        End If
'   End If
   
   'Valida Garantias Reales
   If ChkGarReal.value = 1 Then
    'ALPA 20120320*************************************************************
        If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
            If (txtFechaCertifGravamenGA.Text = "__/__/____" Or txtFechaCertifGravamenGA.Text = "01/01/1950") And chkEnTramiteGA = 0 Then
                MsgBox "Ingrese Fecha Valida de Certificado de Gravamen", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If

            If Val(txtValorGravadoGA.Text) <= 0 And chkEnTramiteGA = 0 Then
                MsgBox "Ingrese Valor Gravado mayor que cero", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If
        Else
            'EJVG20130128 ***
            If cboTpoInscripcion.ListIndex = -1 And chkEnTramite = 0 Then
                MsgBox "Ingrese el Tipo de Inscripción en Registros Públicos de la Garantía", vbInformation, "Aviso"
                If cboTpoInscripcion.Visible And cboTpoInscripcion.Enabled Then cboTpoInscripcion.SetFocus
                If SSGarant.TabVisible(2) Then SSGarant.Tab = 2
                ValidaDatos = False
                Exit Function
            End If
            'END EJVG *******
        'By Capi 03112008
            'If txtFechaCertifGravamen.Text = "__/__/____" Then
            If (txtFechaCertifGravamen.Text = "__/__/____" Or txtFechaCertifGravamen.Text = "01/01/1950") And chkEnTramite = 0 Then
                MsgBox "Ingrese Fecha Valida de Certificado de Gravamen", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If
            
            'By Capi 03112008 para que valide el monto gravado
            If Val(txtValorGravado.Text) <= 0 And chkEnTramite = 0 Then
                MsgBox "Ingrese Valor Gravado mayor que cero", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If
        '**************************************************************************
        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
            If Trim(LblInmobCod.Caption) = "" Then
                MsgBox "Ingrese la Inmobiliaria o el Vendedor", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaInmob.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If CboTipoInmueb.ListIndex = -1 Then
                MsgBox "Seleccione el Tipo del Inmueble", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CboTipoInmueb.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(LblTasaPersCod.Caption) = "" Then
                MsgBox "Ingrese al Tasador", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaTasa.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(LblNotaPersCod.Caption) = "" Then
                MsgBox "Ingrese la Notaria", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaNot.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
'            If Trim(LblSegPersCod.Caption) = "" Then
'                MsgBox "Ingrese la Empresa Aseguradora", vbInformation, "Aviso"
'                SSGarant.Tab = 2
'                CmdBuscaSeg.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If Trim(TxtNroPoliza.Text) = "" Then
'                MsgBox "Ingrese el numero de poliza", vbInformation, "Aviso"
'                SSGarant.Tab = 4
'                TxtNroPoliza.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If Trim(TxtFecVig.Text) = "__/__/____" Then
'                MsgBox "Ingrese la fecha de Vigencia de la poliza", vbInformation, "Aviso"
'                SSGarant.Tab = 4
'                TxtFecVig.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If Trim(TxtMontoPol) = "" Then
'                TxtMontoPol.Text = "0.00"
'            End If
'
'            If Trim(TxtMontoPol) = "0.00" Then
'                MsgBox "Ingrese el Monto de la poliza", vbInformation, "Aviso"
'                SSGarant.Tab = 4
'                TxtMontoPol.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If Trim(TxtFecCons.Text) = "__/__/____" Then
'                MsgBox "Ingrese la fecha de Constitucion de la poliza", vbInformation, "Aviso"
'                SSGarant.Tab = 4
'                TxtFecCons.SetFocus
'                ValidaDatos = False
'                Exit Function
'            End If
            
            If Trim(TxtFecTas.Text) = "__/__/____" Then
                MsgBox "Ingrese la fecha de Tasacion de la Garantia Real", vbInformation, "Aviso"
                SSGarant.Tab = 4
                TxtFecTas.SetFocus
                ValidaDatos = False
                Exit Function
            End If
                        
        End If
                        
        sCad = ValidaFecha(TxtFechareg.Text)
        If Trim(sCad) <> "" Then
            MsgBox sCad, vbInformation, "Aviso"
            TxtFechareg.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(TxtRegNro.Text) = "" Then
            MsgBox "Ingrese el Numero de Registro", vbInformation, "Aviso"
            TxtRegNro.SetFocus
            ValidaDatos = False
            Exit Function
        End If
   End If
End If
     
   'Valida Que Monto Disponible No sea Mayor del 90%
   If ChkGarReal.value = 1 Then
   
        'ARCV 17-07-2006
        If FraDatInm.Visible = True Then
            'ARCV 27-01-2007
            'If cboEstadoTasInm.ListIndex = -1 Then
            '    MsgBox "Ingrese el Estado de Tasación", vbInformation, "Mensaje"
            '    ValidaDatos = False
            '    Exit Function
            'End If
            '-----------
        Else
'            If Me.fraDatMaqEquipo.Visible = True Then
'                If Me.cboEstadoTasMaq.ListIndex = -1 Then
'                    MsgBox "Ingrese el Estado de Tasación", vbInformation, "Mensaje"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'            If Me.fraDatVehic.Visible = True Then
                'By capi 10102008
'                If cboEstadoTasVeh.ListIndex = -1 Then
'                    MsgBox "Ingrese el Estado de Tasación", vbInformation, "Mensaje"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                'End by
'            End If
        
        End If
        If FraDatInm.Visible = True Then
            'EJVG20130128 ***
            If cboTpoInscripcion.ListIndex = -1 And chkEnTramite = 0 Then
                MsgBox "Ingrese el Tipo de Inscripción en Registros Públicos de la Garantía", vbInformation, "Aviso"
                If cboTpoInscripcion.Visible And cboTpoInscripcion.Enabled Then cboTpoInscripcion.SetFocus
                If SSGarant.TabVisible(2) Then SSGarant.Tab = 2
                ValidaDatos = False
                Exit Function
            End If
            'END EJVG *******
'            If txtFechaTasInm.Text = "__/__/____" Then
'                MsgBox "Ingrese Fecha de Tasación del Inmueble", vbInformation, "Mensaje"
'                ValidaDatos = False
'                Exit Function
'            End If

            'peac 20071123
            'By Capi 03112008
            'If txtFechaCertifGravamen.Text = "__/__/____" Then
            If (txtFechaCertifGravamen.Text = "__/__/____" Or txtFechaCertifGravamen.Text = "01/01/1950") And chkEnTramite = 0 Then
                MsgBox "Ingrese Fecha Valida de Certificado de Gravamen", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If
            
            'By Capi 03112008 para que valide el monto gravado
            If Val(txtValorGravado.Text) <= 0 And chkEnTramite = 0 Then
                MsgBox "Ingrese Valor Gravado mayor que cero", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If
            'ALPA 20120320*************************************************************
            If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
                If (txtFechaCertifGravamenGA.Text = "__/__/____" Or txtFechaCertifGravamenGA.Text = "01/01/1950") And chkEnTramiteGA = 0 Then
                    MsgBox "Ingrese Fecha Valida de Certificado de Gravamen", vbInformation, "Mensaje"
                    ValidaDatos = False
                    Exit Function
                End If

                If Val(txtValorGravadoGA.Text) <= 0 And chkEnTramiteGA = 0 Then
                    MsgBox "Ingrese Valor Gravado mayor que cero", vbInformation, "Mensaje"
                    ValidaDatos = False
                Exit Function
            End If
            '***************************************************************************

            End If
            '
            
        End If
        If Me.fraDatVehic.Visible = True Then
'            If txtFecTasVeh.Text = "__/__/____" Then
'                MsgBox "Ingrese Fecha de Tasación.", vbInformation, "Mensaje"
'                ValidaDatos = False
'                Exit Function
'            End If
        End If

        '-------
   
        'by capi 10102008 comentado porque ya no existe txtprecioventa
   
'        If Trim(Right(CmbTipoGarant.Text, 5)) <> "" Then
'            If Trim(TxtPrecioVenta.Text) <> "" Then
'                If CInt(Trim(Right(CmbTipoGarant.Text, 5))) = gPersGarantiaHipotecas Then
'                    'Set oGarantia = New COMNCredito.NCOMGarantia
'                    'nValor = oGarantia.PorcentajeGarantia(gPersGarantia & Trim(Right(CmbTipoGarant.Text, 10)))
'                    nValor = nValorPorcen
'                    'Set oGarantia = Nothing
'
''                    If CDbl(txtMontoxGrav.Text) > CDbl(Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")) Then
''                        MsgBox "Monto Disponible No Puede Exeder al " & Format(nValor * 100, "#0.00") & "% del Precio de Venta", vbInformation, "Aviso"
''                        txtMontoxGrav.Text = Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")
''                        txtMontoxGrav.SetFocus
''                        SSGarant.Tab = 1
''                        ValidaDatos = False
''                        Exit Function
''                    End If
'
'                    'Valida Que la Cuota Inicial No se Menor al 10% del Precio de Venta
'                    'Set oGarantia = New COMNCredito.NCOMGarantia
'                    'nValor = oGarantia.PorcentajeGarantia("3052")
'                    nValor = nValorCuota
'                    'Set oGarantia = Nothing
'
'                    'ARCV 27-01-2007
'                    'If CDbl(TxtHipCuotaIni.Text) < CDbl(Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")) Then
'                    '    MsgBox "Monto de Cuota Inicial No Puede Ser Menor que  el " & Format(nValor * 100, "#0.00") & "% del Precio de Venta", vbInformation, "Aviso"
'                    '    TxtHipCuotaIni.Text = Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")
'                    '    TxtHipCuotaIni.SetFocus
'                    '    SSGarant.Tab = 2
'                    '    ValidaDatos = False
'                    '    Exit Function
'                    'End If
'                    '------------
'                End If
'            End If
'        End If
'    End If
'
   End If
   'ALPA20140206*******************************
   If SSGarant.TabVisible(6) = True Then
        If ckPolizaBF.value = 1 And lnTipoGarantiaActual = 39 Then
            If txtFechaPBF.Text = "__/__/____" Then
                MsgBox "Ingrese Fecha de Bien Futuro", vbInformation, "Mensaje"
                ValidaDatos = False
                Exit Function
            End If
        Else
            ckPolizaBF.value = 0
            txtFechaPBF.Text = "01/01/1900"
        End If
   End If
    ' valida que se haya digitado algun declaracion jurada
    
End Function
Private Sub HabilitaIngresoGarantReal(ByVal pbHabilita As Boolean)
    CmdBuscaInmob.Enabled = pbHabilita
    TxtTelefono.Enabled = pbHabilita
    CboTipoInmueb.Enabled = pbHabilita
    CmdBuscaTasa.Enabled = pbHabilita
    CmdBuscaNot.Enabled = pbHabilita
    CmdBuscaSeg.Enabled = pbHabilita
    TxtFechareg.Enabled = pbHabilita
    TxtRegNro.Enabled = pbHabilita
    FraDatInm.Enabled = pbHabilita
    fraDatVehic.Enabled = pbHabilita
    'ALPA 20120322********************
    Frame9.Enabled = pbHabilita
    Frame8.Enabled = pbHabilita
    '*********************************
    FraRRPP.Enabled = pbHabilita 'EJVG20130131
    '*** PEAC 20080513
    'fraDatMaqEquipo.Enabled = pbHabilita
    
    FraGar.Enabled = pbHabilita
    Me.TxtDirecRegPubli.Enabled = pbHabilita
End Sub
Private Sub HabilitaIngreso(ByVal pbHabilita As Boolean)
        FraClase.Enabled = pbHabilita
        FraTipoRea.Enabled = pbHabilita
        CboBanco.Enabled = pbHabilita
        ChkGarReal.Enabled = pbHabilita
        If ChkGarReal.value = pbHabilita And (ChkGarPoliza.value = True Or chkGarPolizaMob.value = True) Then
            If ChkGarPoliza.value = True Then
                ChkGarPoliza.Enabled = True
                chkGarPolizaMob.Enabled = False
            Else
                ChkGarPoliza.Enabled = False
                chkGarPolizaMob.Enabled = True
            End If
            chkDocCompra.Enabled = pbHabilita
        Else
            ChkGarPoliza.Enabled = False
            chkGarPolizaMob.Enabled = False
            chkDocCompra.Enabled = False
            ChkTasacion.Enabled = False
            'ALPA 20120407 *******Comentado
'            ChkGarPoliza.value = vbUnchecked
'            chkGarPolizaMob.value = vbUnchecked
'            chkDocCompra.value = vbUnchecked
'            ChkTasacion.value = vbUnchecked
'**************************
        End If
        
        If Me.CboGarantia.Text = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" And cmdAceptar.Visible = False Then
            CmbDocGarant.Enabled = Not pbHabilita
            txtNumDoc.Enabled = Not pbHabilita
            CmbTipoGarant.Enabled = Not pbHabilita
            CboGarantia.Enabled = Not pbHabilita
            txtMontotas.Enabled = Not pbHabilita
            txtMontoRea.Enabled = Not pbHabilita
            txtMontoxGrav.Enabled = Not pbHabilita
            cmbMoneda.Enabled = Not pbHabilita
        Else
            CmbDocGarant.Enabled = pbHabilita
            txtNumDoc.Enabled = pbHabilita
            CmbTipoGarant.Enabled = pbHabilita
            CboGarantia.Enabled = pbHabilita
            txtMontotas.Enabled = pbHabilita
            txtMontoRea.Enabled = pbHabilita
            txtMontoxGrav.Enabled = pbHabilita
            cmbMoneda.Enabled = pbHabilita
        End If
        
        If Me.CboNumPF.Visible = True Then
            CboNumPF.Enabled = pbHabilita
        End If
        
        
        cmdBuscar.Enabled = pbHabilita
        txtDescGarant.Enabled = pbHabilita
        cmbPersUbiGeo(0).Enabled = pbHabilita
        cmbPersUbiGeo(1).Enabled = pbHabilita
        cmbPersUbiGeo(2).Enabled = pbHabilita
        cmbPersUbiGeo(3).Enabled = pbHabilita
        txtcomentarios.Enabled = pbHabilita
        txtDireccion.Enabled = pbHabilita 'Campo adicional
        FERelPers.lbEditarFlex = False
        '------------------------------------------
        FEDeclaracionJur.lbEditarFlex = False
        '------------------------------------------
        CmdCliNuevo.Enabled = pbHabilita
        CmdCliEliminar.Enabled = pbHabilita
        
        '---------------------------------
        CmdDJNuevo.Enabled = pbHabilita
        CmdDJEliminar.Enabled = pbHabilita
        '---------------------------------
        
        cmdNuevo.Enabled = Not pbHabilita
        cmdNuevo.Visible = Not pbHabilita
        cmdAceptar.Enabled = pbHabilita
        cmdAceptar.Visible = pbHabilita
        cmdEditar.Enabled = Not pbHabilita
        cmdEditar.Visible = Not pbHabilita
        cmdCancelar.Enabled = pbHabilita
        cmdCancelar.Visible = pbHabilita
        cmdEliminar.Enabled = Not pbHabilita
        cmdEliminar.Visible = Not pbHabilita
        cmdSalir.Enabled = Not pbHabilita
        cmdLimpiar.Enabled = Not pbHabilita
        cmdBuscar.Enabled = Not pbHabilita
        FraBuscaPers.Enabled = Not pbHabilita
        If vTipoInicio = MantenimientoGarantia Then
            cmdNuevo.Enabled = False
        End If
        framontos.Enabled = pbHabilita
        CmdBuscaEmisor.Enabled = pbHabilita
        
        'ARCV 11-07-2006
        FeTabla.lbEditarFlex = pbHabilita
        cmbTipoCreditoTabla.Enabled = pbHabilita
        
        '*** BRGO 20111205 *********************************
        LblEmisorPersCod.Enabled = pbHabilita
        LblEmisorPersNombre.Enabled = pbHabilita
        cmdBuscaEmisorDoc.Enabled = pbHabilita
        CmbDocCompra.Enabled = pbHabilita
        txtNumDocCompra.Enabled = pbHabilita
        txtFecEmision.Enabled = pbHabilita
        txtValorDocCompra.Enabled = pbHabilita
        cboClaseMueble.Enabled = pbHabilita
        txtAnioFabricacion.Enabled = pbHabilita
        '*** END BRGO **************************************
End Sub

Private Sub CargaUbicacionesGeograficas(ByVal prsUbic As ADODB.Recordset)
Dim i As Long
Dim nPos As Integer

    'Carga Niveles
    ContNiv1 = 0
    ContNiv2 = 0
    ContNiv3 = 0
    
    ContNiv4 = 0
    
'    Do While Not prsUbic.EOF
'        Select Case prsUbic!P
'            Case 1 ' Departamento
'                ContNiv1 = ContNiv1 + 1
'                ReDim Preserve Nivel1(ContNiv1)
'                Nivel1(ContNiv1 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'            Case 2 ' Provincia
'                ContNiv2 = ContNiv2 + 1
'                ReDim Preserve Nivel2(ContNiv2)
'                Nivel2(ContNiv2 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'            Case 3 'Distrito
'                ContNiv3 = ContNiv3 + 1
'                ReDim Preserve Nivel3(ContNiv3)
'                Nivel3(ContNiv3 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'            Case 4 'Zona
'                ContNiv4 = ContNiv4 + 1
'                ReDim Preserve Nivel4(ContNiv4)
'                Nivel4(ContNiv4 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'        End Select
'        prsUbic.MoveNext
'    Loop
    
    If prsUbic.EOF Then Exit Sub
    
    Do While prsUbic!P = 1
        ContNiv1 = ContNiv1 + 1
        ReDim Preserve Nivel1(ContNiv1)
        Nivel1(ContNiv1 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
    Loop
        
    Do While prsUbic!P = 2
        ContNiv2 = ContNiv2 + 1
        ReDim Preserve Nivel2(ContNiv2)
        Nivel2(ContNiv2 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
    Loop
    
    Do While prsUbic!P = 3
        ContNiv3 = ContNiv3 + 1
        ReDim Preserve Nivel3(ContNiv3)
        Nivel3(ContNiv3 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
    Loop
    
    Do While prsUbic!P = 4
        ContNiv4 = ContNiv4 + 1
        ReDim Preserve Nivel4(ContNiv4)
        Nivel4(ContNiv4 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
        If prsUbic.EOF Then Exit Do
    Loop
            
    'Carga el Nivel1 en el Control
    cmbPersUbiGeo(0).Clear
    For i = 0 To ContNiv1 - 1
        cmbPersUbiGeo(0).AddItem Nivel1(i)
        If Trim(Right(Nivel1(i), 12)) = "113000000000" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbPersUbiGeo(2).Clear
    cmbPersUbiGeo(3).Clear
    
End Sub

Private Sub LimpiaGarantiaReal()
    
    LblInmobCod.Caption = ""
    LblInmobNombre.Caption = ""
    TxtTelefono.Text = ""
    CboTipoInmueb.ListIndex = -1
    LblTasaPersCod.Caption = ""
    LblTasaPersNombre.Caption = ""
    LblNotaPersCod.Caption = ""
    LblNotaPersNombre.Caption = ""
    LblSegPersCod.Caption = ""
    LblSegPersNombre.Caption = ""
    TxtFechareg.Text = "__/__/____"
    TxtRegNro.Text = ""
    'By Capi 25092008 porque se elimino el control TxtHipCuotaIni y TxtMontoHip
    'TxtHipCuotaIni.Text = "0.00"
    'TxtMontoHip.Text = "0.00"
    '
    
    txtValorGravado.Text = "0.00" 'peac 20071123
    txtValorGravadoGA.Text = "0.00" 'ALPA 20120320
    
    'By Capi 25092008 porque se elimino el control TxtPrecioVenta y TxtValorCConst
    'TxtPrecioVenta.Text = "0.00"
    'TxtValorCConst.Text = "0.00"
    
    TxtNroPoliza.Text = ""
    TxtFecVig.Text = "__/__/____"
    TxtMontoPol.Text = ""
    TxtFecCons.Text = "__/__/____"
    TxtFecTas.Text = "__/__/____"
    
    'By Capi 25092008 porque se elimino el control TxtFecVctoPol
    'TxtFecVctoPol.Text = "__/__/____"
    
    
    Me.TxtDirecRegPubli.Text = ""
    TxtValorEdificacion.Text = "0.00"
    txtValorMerca.Text = "0.00"
    txtVRM.Text = "0.00" 'PEAC 20071122
    '*** PEAC 20080523
    txtAnioFab.Text = "__/__/____"
    txtDescrip.Text = ""
    txtNumMotor.Text = ""
    txtNumSerie.Text = ""
    
    '*** BRGO 20111205 ************
    LblEmisorPersCod = ""
    LblEmisorPersNombre = ""
    cmdBuscaEmisorDoc.Enabled = True
    CmbDocCompra.Enabled = True
    txtNumDocCompra = ""
    txtFecEmision = "__/__/____"
    txtValorDocCompra = "0.00"
    cboClaseMueble.Enabled = True
    txtAnioFabricacion.Text = ""
    '*** END BRGO
    'ALPA 20120322********************
    Frame9.Enabled = False
    Frame8.Enabled = False
    '*********************************
    
End Sub

Private Sub LimpiaPantalla()
    bCarga = True
    LblEmisor.Tag = LblEmisor.Caption
    LblPersCodEmi.Tag = LblPersCodEmi.Caption
    Call LimpiaControles(Me)
    LblEmisor.Caption = LblEmisor.Tag
    LblPersCodEmi.Caption = LblPersCodEmi.Tag
    Call LimpiaFlex(FERelPers)
    '------------------------------------------
    Call LimpiaFlex(FEDeclaracionJur)
    '------------------------------------------
    Call InicializaCombos(Me)
    txtMontotas.BackColor = vbWhite
    txtMontotas.Text = "0.00"
    txtMontoRea.Text = "0.00"
    txtMontoRea.BackColor = vbWhite
    txtMontoxGrav.Text = "0.00"
    txtMontoxGrav.BackColor = vbWhite
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    LblEmisor.Caption = ""
    LblPersCodEmi.Caption = ""
    OptCG(1).value = True
    OptCG(0).value = True
    OptTR(0).value = True
    CboBanco.ListIndex = -1
    ChkGarReal.value = 0
    '*** PEAC 20080513
    ChkGarPoliza.value = 0
    ChkTasacion.value = 0
    Call LimpiaGarantiaReal
    bCarga = False
    bAsignadoACredito = False
    
    'ARCV 11-07-2006
    FeTabla.Clear
    FeTabla.Rows = 2
    FeTabla.FormaCabecera
    cmbTipoCreditoTabla.ListIndex = -1
    lblCoberturaCredito.Caption = "0.00"
    'BRGO 20111205
    chkDocCompra.value = 0
    chkGarPolizaMob.value = 0
    'ALPA20130203****************************
    txtFechaPBF.Text = "__/__/____"
    ckPolizaBF.value = 0
    '****************************************
    
    cmdActAdmCred.Visible = False 'WIOR 20130122
End Sub

Private Function CargaDatos(ByVal psNumGarant As String, Optional ByVal pbBuscaGar As Boolean = False) As Boolean
'***PEAC se agrego el parametro pbBuscaGar
Dim oGarantia As COMDCredito.DCOMGarantia
Dim nTempo As Integer
Dim nLevantada As Boolean

Dim rsGarantia As ADODB.Recordset
Dim pbGarantiaLegal As Boolean
Dim rsRelGarantia As ADODB.Recordset
Dim rsGarantReal As ADODB.Recordset
Dim rsGarantDJ As ADODB.Recordset
Dim rsInmueblePoliza As ADODB.Recordset 'peac 20071222
Dim rsTasador As ADODB.Recordset '*** PEAC 20090722
Dim rsTablaValores As ADODB.Recordset  'ARCV 11-07-2006
Dim rsMueblePoliza As ADODB.Recordset
Dim rsDocCompra As ADODB.Recordset

Dim L As ListItem

    On Error GoTo ErrorCargaDatos
    
    Set oGarantia = New COMDCredito.DCOMGarantia
    bAsignadoACredito = False
    Call oGarantia.CargarDatosGarantia(psNumGarant, rsGarantia, rsRelGarantia, _
                                        rsGarantReal, rsGarantDJ, bAsignadoACredito, rsTablaValores, _
                                        rsInmueblePoliza, rsTasador, pbGarantiaLegal, rsMueblePoliza, rsDocCompra)
    Set oGarantia = Nothing

    'bAsignadoACredito = oGarantia.PerteneceACredito(psNumGarant)
    'Set oGarantia = Nothing

    '***PEAC 20090717
'    If rsGarantia.RecordCount = 0 Then
'        CargaDatos = False
'        Exit Function
'    Else
'        CargaDatos = True
'    End If
      
    If Not (rsGarantia.EOF And rsGarantia.BOF) Then
        CargaDatos = True
    Else
        CargaDatos = False
        Exit Function
    End If
    '***FIN PEAC

    If rsGarantia!nEstado = 5 Then 'Si es levantada
        nLevantada = True
    Else
        nLevantada = False
    End If

    '*** PEAC 20090717
    If pbBuscaGar Then
        LstGaratias.ListItems.Clear
        
        Set L = LstGaratias.ListItems.Add(, , IIf(IsNull(rsGarantia!cDescripcion), "", rsGarantia!cDescripcion))
        L.Bold = True
'        If R!nMoneda = gMonedaExtranjera Then
'            L.ForeColor = RGB(0, 125, 0)
'        Else
'            L.ForeColor = vbBlack
'        End If
        L.SubItems(1) = Trim(psNumGarant)
        
'        L.SubItems(2) = Trim(R!cPersCodEmisor)
'        L.SubItems(3) = PstaNombre(R!cPersNombre)
'        L.SubItems(4) = Trim(R!cTpoDoc)
'        L.SubItems(5) = Trim(R!cNroDoc)
                
        'rsGarantia!cDescripcion psNumGarant
    End If
    
    LblEstado.Caption = Trim(rsGarantia!cEstado)
    nTempo = IIf(IsNull(rsGarantia!nGarClase), 0, rsGarantia!nGarClase)
    OptCG(nTempo).value = True
    nTempo = IIf(IsNull(rsGarantia!nGarTpoRealiz), 0, rsGarantia!nGarTpoRealiz)
    OptTR(nTempo).value = True
    'PosicionSuperGarantias Trim(Str(R!nTpoGarantia))
    PosicionSuperGarantias rsGarantia!IdSupGarant
    CmbTipoGarant.ListIndex = IndiceListaCombo(CmbTipoGarant, Trim(Str(rsGarantia!nTpoGarantia)))

    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, Trim(Str(rsGarantia!nmoneda)))
    txtDescGarant.Text = IIf(IsNull(rsGarantia!cDescripcion), "", Trim(rsGarantia!cDescripcion))
    
    ChkGarReal.value = IIf(IsNull(rsGarantia!nGarantReal), 0, rsGarantia!nGarantReal)
    Call ChkGarReal_Click 'ALPA 20120329
    '*** PEAC 20080513
    ChkGarPoliza.value = IIf(IsNull(rsGarantia!nGarantPoliza), 0, rsGarantia!nGarantPoliza)
    ChkTasacion.value = IIf(IsNull(rsGarantia!nTasador), 0, rsGarantia!nTasador)
    
    '*** BRGO 20111205
    chkGarPolizaMob.value = IIf(Trim(rsGarantia!nGarantMob) = "0", 0, 1)
    chkDocCompra.value = IIf(Trim(rsGarantia!nGarantDoc) = "0", 0, 1)
    
    If Not rsMueblePoliza.EOF And Not rsMueblePoliza.BOF Then
        cboClaseMueble.ListIndex = IndiceListaCombo(cboClaseMueble, rsMueblePoliza!nMobiliario)
        txtAnioFabricacion.Text = rsMueblePoliza!nAnioFabric
    End If

    If Not rsDocCompra.EOF And Not rsDocCompra.BOF Then
        LblEmisorPersCod.Caption = rsDocCompra!cPersCod
        LblEmisorPersNombre.Caption = rsDocCompra!cPersNombre
        txtNumDocCompra.Text = rsDocCompra!cNroDoc
        txtFecEmision.Text = Format(rsDocCompra!dFecEmision, "dd/MM/yyyy")
        txtValorDocCompra.Text = Format(rsDocCompra!nValor, "#,###,##0.00")
        CmbDocCompra.ListIndex = IndiceListaCombo(CmbDocCompra, rsDocCompra!nTpoDoc)
    End If
    '*** END BRGO
    
    
    
    'MADM 20101201 -- llenar labels de emisor
    LblPersCodEmi.Caption = rsGarantia!cPersCodEmisor
    LblEmisor.Caption = rsGarantia!cPersNombre
    'END MADM
    
    ChkTasacion.Enabled = False
    chkDocCompra.Enabled = False
    chkGarPolizaMob.Enabled = False
    ChkGarPoliza.Enabled = False
    
    '*** PEAC 20090327
    If ChkGarPoliza.value = 1 Then
        SSGarant.TabVisible(6) = True
    Else
        SSGarant.TabVisible(6) = False
    End If
    '****** PEAC 20090722
    If ChkTasacion.value = 1 Then
        SSGarant.TabVisible(4) = True
    Else
        SSGarant.TabVisible(4) = False
    End If
    '*******************
    '*** BRGO 20111205 ****************
    If chkGarPolizaMob.value = 1 Then
        SSGarant.TabVisible(8) = True
    Else
        SSGarant.TabVisible(8) = False
    End If
    
    If chkDocCompra.value = 1 Then
        SSGarant.TabVisible(7) = True
    Else
        SSGarant.TabVisible(7) = False
    End If
    '*** END BRGO **********************
    
    Call HabilitaIngresoGarantReal(False)
    CboBanco.ListIndex = IndiceListaCombo(CboBanco, IIf(IsNull(rsGarantia!cBancoPersCod), "", rsGarantia!cBancoPersCod))
    CboBanco.Enabled = False
    
    CmbDocGarant.ListIndex = IndiceListaCombo(CmbDocGarant, rsGarantia!cTpoDoc)
       
    txtNumDoc.Text = rsGarantia!cNroDoc
    
    'Carga Ubicacion Geografica
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "1" & Mid(rsGarantia!cZona, 2, 2) & String(9, "0"))
    cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "2" & Mid(rsGarantia!cZona, 2, 4) & String(7, "0"))
    cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "3" & Mid(rsGarantia!cZona, 2, 6) & String(5, "0"))
    cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & rsGarantia!cZona)
    
    If rsGarantia!nmoneda = gMonedaExtranjera Then
        txtMontotas.BackColor = RGB(200, 255, 200)
        txtMontoRea.BackColor = RGB(200, 255, 200)
        txtMontoxGrav.BackColor = RGB(200, 255, 200)
    Else
        txtMontotas.BackColor = vbWhite
        txtMontoRea.BackColor = vbWhite
        txtMontoxGrav.BackColor = vbWhite
    End If
    
    'By Capi 03112008

    txtMontotas.Text = Format(rsGarantia!nTasacion, "#0.00")
    txtMontoRea.Text = Format(rsGarantia!nRealizacion, "#0.00")
    txtMontoxGrav.Text = Format(rsGarantia!nPorGravar - rsGarantia!nGravament, "#0.00")
    nxgravar = rsGarantia!nGravament 'madm 20100513
    txtcomentarios.Text = Trim(IIf(IsNull(rsGarantia!cComentario), "", rsGarantia!cComentario))
    txtDireccion.Text = rsGarantia!cDireccion 'campo adicional
    '***  PEAC 20090724
    If Me.ChkTasacion.value = 1 Then
    Call LimpiaFlex(FETasacion)
        Do While Not rsGarantReal.EOF
            FETasacion.AdicionaFila
            If Trim(rsGarantReal!cPersCodTasador) = "" Then
                FETasacion.TextMatrix(rsGarantReal.Bookmark, 1) = "No Ingresado"
                FETasacion.TextMatrix(rsGarantReal.Bookmark, 2) = "No Ingresado...Actualice"
            Else
                FETasacion.TextMatrix(rsGarantReal.Bookmark, 1) = IIf(IsNull(rsGarantReal!cPersCodTasador), "", rsGarantReal!cPersCodTasador)
                FETasacion.TextMatrix(rsGarantReal.Bookmark, 2) = IIf(IsNull(rsGarantReal!cTasaPersNombre), "", rsGarantReal!cTasaPersNombre)
            End If
            FETasacion.TextMatrix(rsGarantReal.Bookmark, 3) = IIf(IsNull(rsGarantReal!dTasacion), "", rsGarantReal!dTasacion)
            FETasacion.TextMatrix(rsGarantReal.Bookmark, 4) = IIf(IsNull(rsGarantReal!nVRM), "", rsGarantReal!nVRM)
            FETasacion.TextMatrix(rsGarantReal.Bookmark, 5) = IIf(IsNull(rsGarantReal!nValorEdificacion), "", rsGarantReal!nValorEdificacion)
            rsGarantReal.MoveNext
        Loop
    End If
'***FIN  PEAC

    'Personas Relacionadas con Garantias
'    Set oGarantia = New COMDCredito.DCOMGarantia
'    Set RRelPers = oGarantia.RecuperaRelacPersonaGarantia(psNumGarant)
'    Set oGarantia = Nothing
    
    Call LimpiaFlex(FERelPers)
    Do While Not rsRelGarantia.EOF
        FERelPers.AdicionaFila
        FERelPers.TextMatrix(rsRelGarantia.Bookmark, 1) = rsRelGarantia!cPersCod
        FERelPers.TextMatrix(rsRelGarantia.Bookmark, 2) = rsRelGarantia!cPersNombre
        FERelPers.TextMatrix(rsRelGarantia.Bookmark, 3) = rsRelGarantia!cRelacion
        rsRelGarantia.MoveNext
    Loop
    'RRelPers.Close
    'Set RRelPers = Nothing
    
    'R.Close
    'Set R = Nothing
    
    'Carga Garantias Reales
    If ChkGarReal.value = 1 Then
     '   Set oGarantia = New COMDCredito.DCOMGarantia
     '   Set R = oGarantia.RecuperaGarantiaReal(psNumGarant)
     '   Set oGarantia = Nothing
        rsGarantReal.MoveFirst
        If Not (rsGarantReal.EOF And rsGarantReal.BOF) Then
            LblInmobCod.Caption = IIf(IsNull(rsGarantReal!cPersCodVend), "", rsGarantReal!cPersCodVend)
            LblInmobNombre.Caption = IIf(IsNull(rsGarantReal!cVendPersNombre), "", rsGarantReal!cVendPersNombre)
            TxtTelefono.Text = IIf(IsNull(rsGarantReal!cPersVendTelef), "", rsGarantReal!cPersVendTelef)
            CboTipoInmueb.ListIndex = IndiceListaCombo(CboTipoInmueb, IIf(IsNull(rsGarantReal!nTipVivienda), 0, rsGarantReal!nTipVivienda))
                        
            LblTasaPersCod.Caption = IIf(IsNull(rsGarantReal!cPersCodTasador), "", rsGarantReal!cPersCodTasador)
            LblTasaPersNombre.Caption = IIf(IsNull(rsGarantReal!cTasaPersNombre), "", rsGarantReal!cTasaPersNombre)
            LblNotaPersCod.Caption = IIf(IsNull(rsGarantReal!cPersNotaria), "", rsGarantReal!cPersNotaria)
            LblNotaPersNombre.Caption = IIf(IsNull(rsGarantReal!cNotaPersNombre), "", rsGarantReal!cNotaPersNombre)
            LblSegPersCod.Caption = IIf(IsNull(rsGarantReal!cPersCodSeguro), "", rsGarantReal!cPersCodSeguro)
            LblSegPersNombre.Caption = IIf(IsNull(rsGarantReal!cSegPersNombre), "", rsGarantReal!cSegPersNombre)
            TxtFechareg.Text = IIf(IsNull(rsGarantReal!dEscritura), "__/__/____", rsGarantReal!dEscritura)
            TxtRegNro.Text = IIf(IsNull(rsGarantReal!cRegistro), "", rsGarantReal!cRegistro)
            
            'By Capi 25092008 porque se elimino el control TxtHipCuotaIni y TxtMontoHip
            'TxtHipCuotaIni.Text = Format(IIf(IsNull(rsGarantReal!nCuotaInicial), "0.00", rsGarantReal!nCuotaInicial), "#0.00")
            'TxtMontoHip.Text = Format(IIf(IsNull(rsGarantReal!nMontoHipoteca), "0.00", rsGarantReal!nMontoHipoteca), "#0.00")
            
            cboTpoInscripcion.ListIndex = IndiceListaCombo(cboTpoInscripcion, rsGarantReal!nTipoInscripcion)
            txtValorGravado.Text = Format(IIf(IsNull(rsGarantReal!nValorGravado), "0.00", rsGarantReal!nValorGravado), "#0.00") 'peac 20071123
            txtValorGravadoGA.Text = Format(IIf(IsNull(rsGarantReal!nValorGravado), "0.00", rsGarantReal!nValorGravado), "#0.00") 'ALPA 20120320
            txtFechaCertifGravamen.Text = IIf(rsGarantReal!dCertifGravamen <> "01/01/1900", Format(rsGarantReal!dCertifGravamen, "dd/mm/yyyy"), "__/__/____") 'peac 20071123
            txtFechaCertifGravamenGA.Text = IIf(rsGarantReal!dCertifGravamen <> "01/01/1950", Format(rsGarantReal!dCertifGravamen, "dd/mm/yyyy"), "__/__/____") 'ALPA 20120320
            
            'By Capi 25092008 porque se elimino el control TxtPrecioVenta y TxtValorCConst
            'TxtPrecioVenta.Text = Format(IIf(IsNull(rsGarantReal!nPrecioVenta), "0.00", rsGarantReal!nPrecioVenta), "#0.00")
            'TxtValorCConst.Text = Format(IIf(IsNull(rsGarantReal!nValorConstruccion), "0.00", rsGarantReal!nValorConstruccion), "#0.00")
            
            TxtNroPoliza.Text = rsGarantReal!nNroPoliza
            TxtFecVig.Text = Format(rsGarantReal!dVigenciaPol, "dd/mm/yyyy")
            TxtMontoPol.Text = Format(rsGarantReal!nMontoPoliza, "#0.00")
            TxtFecCons.Text = Format(rsGarantReal!dConstitucion, "dd/mm/yyyy")
            
            '*** PEAC 20090724 - TRASLADADO MAS ABAJO
            'TxtFecTas.Text = Format(rsGarantReal!dTasacion, "dd/mm/yyyy")
            
            '*** PEAC 20080522
            'By Capi 25092008 porque se elimino el control TxtFecVctoPol
            'TxtFecVctoPol.Text = Format(rsGarantReal!dFecVctoPol, "dd/mm/yyyy")
            
            txtValorMerca.Text = Format(rsGarantReal!nValorMerca, "#0.00")
            txtDireAlma.Text = Trim(rsGarantReal!cDireAlma)
            txtTipoMerca.Text = Trim(rsGarantReal!cTipoMerca)
            
            txtDescrip.Text = Trim(rsGarantReal!cDescripBien)
            txtNumMotor.Text = Trim(rsGarantReal!cNumMotor)
            txtNumSerie.Text = Trim(rsGarantReal!cNumSerie)
            txtAnioFab.Text = Format(rsGarantReal!dAnioFab, "dd/mm/yyyy")
            
            Me.TxtDirecRegPubli.Text = rsGarantReal!cdirgarregpubli
            cboTipo.ListIndex = IndiceListaCombo(cboTipo, IIf(IsNull(rsGarantReal!CTIPOCONTRATO), "", rsGarantReal!CTIPOCONTRATO))
            'ALPA 20120313**********************************************************************************************************
            cboTipoGA.ListIndex = IndiceListaCombo(cboTipo, IIf(IsNull(rsGarantReal!CTIPOCONTRATO), "", rsGarantReal!CTIPOCONTRATO))
            txtFechaBloqueo.Text = IIf(Format(rsGarantReal!dFechaBloqueo, "DD/MM/YYYY") = "1900-01-01", gdFecSis, Format(rsGarantReal!dFechaBloqueo, "DD/MM/YYYY"))
            '***********************************************************************************************************************
            '* Mostrar solo los Datos de Vehiculos o Inmuebles **
            Call HabilitarFramesGarantiaReal(CInt(Trim(Right(CmbTipoGarant, 10))))
            
            If fraDatVehic.Visible Then
                txtPlacaVehic.Text = rsGarantReal!cPlacaAuto
                'ARCV 27-01-2007
                'cboEstadoTasVeh.ListIndex = IndiceListaCombo(cboEstadoTasInm, rsGarantReal!nEstadoTasacion)
                'peac 20071128
                'By Capi 25092008 porque se elimino el control cboEstadoTasVeh
                'cboEstadoTasVeh.ListIndex = IndiceListaCombo(cboEstadoTasInm, rsGarantReal!nEstadoTasacion)
                'End
                
                ' CboTipoInmueb.ListIndex = IndiceListaCombo(CboTipoInmueb, IIf(IsNull(rsGarantReal!nTipVivienda), 0, rsGarantReal!nTipVivienda))
                'By Capi 25092008 porque se elimino el control txtFecTasVeh
                'txtFecTasVeh.Text = IIf(rsGarantReal!dTasacionInmueble <> "01/01/1900", Format(rsGarantReal!dTasacionInmueble, "dd/mm/yyyy"), "__/__/____")
                
            End If
            
'            If Me.fraDatMaqEquipo.Visible Then
'                cboEstadoTasMaq.ListIndex = IndiceListaCombo(cboEstadoTasMaq, rsGarantReal!nEstadoTasacion)
'            End If
            
            If Me.FraDatInm.Visible Then
                'ARCV 27-01-2007
                'cboEstadoTasInm.ListIndex = IndiceListaCombo(cboEstadoTasVeh, rsGarantReal!nEstadoTasacion)
               'By Capi 25092008 porque se elimino el control txtFechaTasInm
               'txtFechaTasInm.Text = IIf(rsGarantReal!dTasacionInmueble <> "01/01/1900", Format(rsGarantReal!dTasacionInmueble, "dd/mm/yyyy"), "__/__/____")
            End If
            
            
            
            ' *******
            '*** PEAC 20090724 - TRASLADADO MAS ABAJO
            'TxtValorEdificacion.Text = Format(rsGarantReal!nValorEdificacion, "#0.00")
            
            '*** PEAC 20090724
            'txtVRM.Text = Format(rsGarantReal!nVRM, "#0.00")
            
        End If

        '*** PEAC 20090327
        '*** peac 20071222
        If Not (rsInmueblePoliza.EOF And rsInmueblePoliza.BOF) Then
'            ChkGarPoliza.Enabled = True
'            SSGarant.TabVisible(6) = True

            cmbClaseInmueble.ListIndex = IndiceListaCombo(cmbClaseInmueble, IIf(IsNull(rsInmueblePoliza!nInmueble), 0, rsInmueblePoliza!nInmueble))
            cmbCategoria.ListIndex = IndiceListaCombo(cmbCategoria, IIf(IsNull(rsInmueblePoliza!nCategoria), 0, rsInmueblePoliza!nCategoria))
            txtNumLocales.Text = Format(IIf(IsNull(rsInmueblePoliza!nNumLocales), "0", rsInmueblePoliza!nNumLocales), "#0")
            txtNumPisos.Text = Format(IIf(IsNull(rsInmueblePoliza!nNumPisos), "0", rsInmueblePoliza!nNumPisos), "#0")
            txtNumSotanos.Text = Format(IIf(IsNull(rsInmueblePoliza!nNumSotanos), "0", rsInmueblePoliza!nNumSotanos), "#0")
            txtAnioConstruccion.Text = Format(IIf(IsNull(rsInmueblePoliza!nAnioConstruc), "0", rsInmueblePoliza!nAnioConstruc), "#0")
            'ALPA20140203***************************************************
            ckPolizaBF.value = IIf(IsNull(rsInmueblePoliza!nPolizaBF), "0", rsInmueblePoliza!nPolizaBF)
            If ckPolizaBF.value And Not IsNull(rsInmueblePoliza!dFechaPBF) Then
                txtFechaPBF.Text = Format(rsInmueblePoliza!dFechaPBF, "DD/MM/YYYY")
            End If
            '***************************************************************
        Else
'            ChkGarPoliza.Enabled = False
'            SSGarant.TabVisible(6) = False

        End If

'*** implementar el rsTasador

        'If Not (rsTasador.EOF And rsTasador.BOF) Then
        If Not (rsGarantReal.EOF And rsGarantReal.BOF) Then

            TxtFecTas.Text = Format(rsGarantReal!dTasacion, "dd/mm/yyyy")
            txtVRM.Text = Format(rsGarantReal!nVRM, "#0.00")
            TxtValorEdificacion.Text = Format(rsGarantReal!nValorEdificacion, "#0.00")

            Call LimpiaFlex(FETasacion)
                Do While Not rsGarantReal.EOF
                    FETasacion.AdicionaFila
                    If Trim(rsGarantReal!cPersCodTasador) = "" Then
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 1) = "No Ingresado"
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 2) = "No Ingresado...Actualice"
                    Else
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 1) = IIf(IsNull(rsGarantReal!cPersCodTasador), "", rsGarantReal!cPersCodTasador)
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 2) = IIf(IsNull(rsGarantReal!cTasaPersNombre), "", rsGarantReal!cTasaPersNombre)
                    End If
                    FETasacion.TextMatrix(rsGarantReal.Bookmark, 3) = IIf(IsNull(rsGarantReal!dTasacion), "", rsGarantReal!dTasacion)
                    FETasacion.TextMatrix(rsGarantReal.Bookmark, 4) = IIf(IsNull(rsGarantReal!nVRM), "", rsGarantReal!nVRM)
                    FETasacion.TextMatrix(rsGarantReal.Bookmark, 5) = IIf(IsNull(rsGarantReal!nValorEdificacion), "", rsGarantReal!nValorEdificacion)
                    rsGarantReal.MoveNext
                Loop
        
        End If
        
         'MADM 201104015
        If pbGarantiaLegal Then
            cmdNuevo.Enabled = False
            cmdEditar.Enabled = False
            cmdEliminar.Enabled = False
            CmbDocGarant.Enabled = False
            txtNumDoc.Enabled = False
            FraRRPP.Enabled = False
            TxtFecTas.Enabled = False
            txtVRM.Enabled = False
            TxtValorEdificacion.Enabled = False
            vTipoInicio = ConsultaGarant
            MsgBox "Esta Garantia No podrá ser Modificada, Comuníquese con el Área de Legal / Sup. Créditos en Agencias", vbInformation, "Aviso"
        Else
            cmdNuevo.Enabled = True
            cmdEditar.Enabled = True
            cmdEliminar.Enabled = True
            CmbDocGarant.Enabled = True
            txtNumDoc.Enabled = True
            FraRRPP.Enabled = True
            TxtFecTas.Enabled = True
            txtVRM.Enabled = True
            TxtValorEdificacion.Enabled = True
            'vTipoInicio = MantenimientoGarantia '***Comentado por ELRO el 20120703, según OYP-RFC022-2012
        End If
        'END MADM
                
    End If

    '*** PEAC 20080513
    If ChkGarPoliza.value = 1 Then

        '*** peac 20071222
        If Not (rsInmueblePoliza.EOF And rsInmueblePoliza.BOF) Then
            cmbClaseInmueble.ListIndex = IndiceListaCombo(cmbClaseInmueble, IIf(IsNull(rsInmueblePoliza!nInmueble), 0, rsInmueblePoliza!nInmueble))
            cmbCategoria.ListIndex = IndiceListaCombo(cmbCategoria, IIf(IsNull(rsInmueblePoliza!nCategoria), 0, rsInmueblePoliza!nCategoria))
            txtNumLocales.Text = Format(IIf(IsNull(rsInmueblePoliza!nNumLocales), "0", rsInmueblePoliza!nNumLocales), "#0")
            txtNumPisos.Text = Format(IIf(IsNull(rsInmueblePoliza!nNumPisos), "0", rsInmueblePoliza!nNumPisos), "#0")
            txtNumSotanos.Text = Format(IIf(IsNull(rsInmueblePoliza!nNumSotanos), "0", rsInmueblePoliza!nNumSotanos), "#0")
            txtAnioConstruccion.Text = Format(IIf(IsNull(rsInmueblePoliza!nAnioConstruc), "0", rsInmueblePoliza!nAnioConstruc), "#0")
        End If
    End If

    '*** PEAC 20090724
    If ChkTasacion.value = 1 Then
        'If Not (rsTasador.EOF And rsTasador.BOF) Then
        rsGarantReal.MoveFirst
        If Not (rsGarantReal.EOF And rsGarantReal.BOF) Then

            TxtFecTas.Text = Format(rsGarantReal!dTasacion, "dd/mm/yyyy")
            txtVRM.Text = Format(rsGarantReal!nVRM, "#0.00")
            TxtValorEdificacion.Text = Format(rsGarantReal!nValorEdificacion, "#0.00")

            Call LimpiaFlex(FETasacion)
                Do While Not rsGarantReal.EOF
                    FETasacion.AdicionaFila
                    If Trim(rsGarantReal!cPersCodTasador) = "" Then
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 1) = "No Ingresado"
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 2) = "No Ingresado...Actualice"
                    Else
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 1) = IIf(IsNull(rsGarantReal!cPersCodTasador), "", rsGarantReal!cPersCodTasador)
                        FETasacion.TextMatrix(rsGarantReal.Bookmark, 2) = IIf(IsNull(rsGarantReal!cTasaPersNombre), "", rsGarantReal!cTasaPersNombre)
                    End If
                    FETasacion.TextMatrix(rsGarantReal.Bookmark, 3) = IIf(IsNull(rsGarantReal!dTasacion), "", rsGarantReal!dTasacion)
                    FETasacion.TextMatrix(rsGarantReal.Bookmark, 4) = IIf(IsNull(rsGarantReal!nVRM), "", rsGarantReal!nVRM)
                    FETasacion.TextMatrix(rsGarantReal.Bookmark, 5) = IIf(IsNull(rsGarantReal!nValorEdificacion), "", rsGarantReal!nValorEdificacion)
                    rsGarantReal.MoveNext
                Loop
        End If
        
    End If

    LblTotDJ.Caption = "0.00"

    ' CMACICA_CSTS - 25/11/2003 -------------------------------------------------
    '*** PEAC 20080412
    'If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then
    If CInt(Trim(Right(CmbDocGarant, 10))) = 93 Then
    
       ' Carga Detalle de Garantia DECLARACION JURADA
       'Set oGarantia = New COMDCredito.DCOMGarantia
       'Set RGarDetDJ = oGarantia.RecuperaGarantDeclaracionJur(psNumGarant)
       'Set oGarantia = Nothing
       Call LimpiaFlex(FEDeclaracionJur)
       Do While Not rsGarantDJ.EOF
          FEDeclaracionJur.AdicionaFila
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 1) = rsGarantDJ!cGarDjDescripcion
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 2) = rsGarantDJ!nGarDJCantidad
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 3) = rsGarantDJ!nGarDJPrecioUnit
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 4) = rsGarantDJ!cGarDJTpoDocDes
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 5) = rsGarantDJ!cGarDJNroDoc
          
          LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) + (FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 2) * FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 3)), "#0.00"))
          
          rsGarantDJ.MoveNext
       Loop
       'RGarDetDJ.Close
       'Set RGarDetDJ = Nothing
    End If
    ' ---------------------------------------------------------------------------
     'WIOR 20130122 ****************************
    Dim sGruposUser As String
    Dim sGrupo() As String
    Dim nPosGrup As Integer
    Dim oGrupoPers As comdpersona.UCOMAcceso
    Set oGrupoPers = New comdpersona.UCOMAcceso
    'WIOR 20150225 ***************************
    Dim oConsSist As COMDConstSistema.NCOMConstSistema
    Set oConsSist = New COMDConstSistema.NCOMConstSistema
    
    sGrupo = Split(Trim(oConsSist.LeeConstSistema(499)), ",")
    Set oConsSist = Nothing
    'WIOR FIN ******************************************
    sGruposUser = oGrupoPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
    'sGrupo = Split(sGruposUser, ",")
    fbGrupoAdmCred = False
    For nPosGrup = 0 To UBound(sGrupo)
        If InStr(1, sGruposUser, Trim(sGrupo(nPosGrup))) > 0 Then
        'If "GRUPO ADMINISTRACION DE CREDITOS" = Trim(Replace(sGrupo(nPosGrup), "'", "")) Then 'WIOR 20150225 COMENTO
            fbGrupoAdmCred = True
            Exit For
        End If
    Next nPosGrup
    fnTpoPoliza = 0
    'WIOR FIN ********************************
    'MADM 20110415 - pbGarantiaLegal
    If pbGarantiaLegal = False Then
        If bAsignadoACredito Or nLevantada Then
            framontos.Enabled = False
            cmdEliminar.Enabled = False
            cmdEditar.Enabled = True
            If nLevantada Then
                cmdEditar.Enabled = False
            End If
        Else
            framontos.Enabled = True
            cmdEliminar.Enabled = True
            cmdEditar.Enabled = True
        End If
    'WIOR 20130122 ***************************
        cmdActAdmCred.Visible = False
        fbGrupoAdmCred = False
    Else
        If fbGrupoAdmCred Then
            If Trim(Str(rsGarantia!IdSupGarant)) = "1" Then
                If Trim(Str(rsGarantia!nTpoGarantia)) = "1" Then
                    ChkGarPoliza.Enabled = True
                    chkGarPolizaMob.Enabled = False
                    txtAnioFabricacion.Enabled = False
                    fnTpoPoliza = 1
                    cmdActAdmCred.Visible = True
                ElseIf Trim(Str(rsGarantia!nTpoGarantia)) = "2" Or Trim(Str(rsGarantia!nTpoGarantia)) = "11" Or Trim(Str(rsGarantia!nTpoGarantia)) = "12" Then
                    ChkGarPoliza.Enabled = False
                    chkGarPolizaMob.Enabled = True
                    txtAnioFabricacion.Enabled = True
                    fnTpoPoliza = 2
                    cmdActAdmCred.Visible = True
                Else
                    fbGrupoAdmCred = False
                    cmdActAdmCred.Visible = False
                    txtAnioFabricacion.Enabled = False
                End If
            Else
                fbGrupoAdmCred = False
                cmdActAdmCred.Visible = False
                txtAnioFabricacion.Enabled = False
            End If
        Else
            cmdActAdmCred.Visible = False
            txtAnioFabricacion.Enabled = False
        End If
    'WIOR FIN ********************************
    End If
    'END MADM
    '*** PEAC 20080412
    If Trim(Right(CmbDocGarant, 10)) = "93" Then
    'If Trim(Right(CmbDocGarant, 10)) = "15" Then
        SSGarant.TabVisible(3) = True
    End If
    
    If Not rsTablaValores.EOF Then 'ARCV 11-07-2006
        Dim OCon As COMDConstantes.DCOMConstantes
        Dim rsT As ADODB.Recordset
        Set OCon = New COMDConstantes.DCOMConstantes
        Set rsT = OCon.RecuperaConstantes(9062)
        Set OCon = Nothing
        Call Llenar_Combo_con_Recordset(rsT, cmbTipoCreditoTabla)

        cmbTipoCreditoTabla.ListIndex = IndiceListaCombo(cmbTipoCreditoTabla, rsTablaValores!nTipoTabla)
        FeTabla.Clear
        FeTabla.Rows = 2
        FeTabla.FormaCabecera
        
        While Not rsTablaValores.EOF
            FeTabla.AdicionaFila
            FeTabla.TextMatrix(rsTablaValores.Bookmark, 0) = rsTablaValores!nTipoEval
            FeTabla.TextMatrix(rsTablaValores.Bookmark, 1) = rsTablaValores!cConsDescripcion
            FeTabla.TextMatrix(rsTablaValores.Bookmark, 2) = rsTablaValores!cDescripcion & Space(75) & rsTablaValores!nCodItem
            FeTabla.TextMatrix(rsTablaValores.Bookmark, 3) = rsTablaValores!nValor
            rsTablaValores.MoveNext
        Wend
        Call ActualizaMontoCobertura
    End If
    

    
    Exit Function
    
ErrorCargaDatos:
        MsgBox err.Description, vbCritical, "Aviso"

End Function

Private Sub CargaBancos(ByVal prsBancos As ADODB.Recordset)
    
    CboBanco.Clear
    Do While Not prsBancos.EOF
        CboBanco.AddItem PstaNombre(prsBancos!cPersNombre) & Space(150) & prsBancos!cPersCod
        prsBancos.MoveNext
    Loop
End Sub

Private Sub CargaControles()
Dim oGarant As COMDCredito.DCOMGarantia
Dim rsTContGR As ADODB.Recordset
Dim rsBancos As ADODB.Recordset
Dim rsTInmue As ADODB.Recordset
Dim rsUbic As ADODB.Recordset
Dim rsTGaran As ADODB.Recordset
Dim rsMoneda As ADODB.Recordset
Dim rsRelac As ADODB.Recordset
Dim rsTDocum As ADODB.Recordset
Dim rsSuperG As ADODB.Recordset
Dim rsClaseInmueble As ADODB.Recordset ' peac 20071116
Dim rsCategoria As ADODB.Recordset ' peac 20071116
Dim rsClaseMueble As ADODB.Recordset 'BRGO 20111205
Dim rsTpoInscripcion As ADODB.Recordset 'EJVG20130128
'Dim rsTContGR As adodb.Recordset
Dim oConsSist As COMDConstSistema.NCOMConstSistema 'WIOR 20150608


    On Error GoTo ERRORCargaControles
    
    'Cargar Objetos de los Controles
    Set oGarant = New COMDCredito.DCOMGarantia
    
    'peac 20071130 para acticarlo despues
    Call oGarant.CargarObjetosControles(rsBancos, rsTInmue, rsUbic, rsTGaran, rsMoneda, rsRelac, rsTDocum, rsSuperG, rsTContGR, rsClaseInmueble, rsCategoria, rsClaseMueble, rsTpoInscripcion)
    'EJVG20130128 Se agregó el parametro prsTipoInscripcion
    'Call oGarant.CargarObjetosControles(rsBancos, rsTInmue, rsUbic, rsTGaran, rsMoneda, rsRelac, rsTDocum, rsSuperG, rsTContGR)
    
    Set oGarant = Nothing
    
    'Carga Bancos
    Call CargaBancos(rsBancos)
    
    'Cargar Ubicaciones Geograficas
    'Call CargaUbicacionesGeograficas(rsUbic)
    While Not rsUbic.EOF
        cmbPersUbiGeo(0).AddItem Trim(rsUbic!cUbiGeoDescripcion) & Space(50) & Trim(rsUbic!cUbiGeoCod)
        rsUbic.MoveNext
    Wend
    'Carga Tipos de Inmuebles
    Call Llenar_Combo_con_Recordset(rsTInmue, CboTipoInmueb)
    
    'Carga Tipos de Inmuebles peac 20071112
    'Call Llenar_Combo_con_Recordset(rsTInmue, CboTipoInmueb)
    
    'Carga Tipos de Garantia
    Call CambiaTamañoCombo(CmbTipoGarant)
    Call Llenar_Combo_con_Recordset(rsTGaran, CmbTipoGarant)
    
    'Carga Monedas
    Call Llenar_Combo_con_Recordset(rsMoneda, cmbMoneda)
    
    Dim rsTContGRGA As ADODB.Recordset
    Set rsTContGRGA = rsTContGR.Clone
    'Carga Tipo Contrato Garantia Real
    Call Llenar_Combo_con_Recordset(rsTContGRGA, Me.cboTipoGA) 'ALPA 20120320
    Call Llenar_Combo_con_Recordset(rsTContGR, Me.cboTipo)
    
    

    'Carga Documento 18-05-2006 avmm
    'Call Llenar_Combo_con_Recordset(rsTDocum, CmbDocGarant)
    
    'Carga Relacion de Personas con Garantia
    FERelPers.CargaCombo rsRelac
        
    ' CMACICA_CSTS - 25/11/2003 ----------------------------------------------------------
    'Carga Tipos de Documentos para el detalle de una Declaracion Jurada
    FEDeclaracionJur.CargaCombo rsTDocum

    'MADM 20100825
    If gGarantiaDepPlazoFijoCF Then
        Call CargarSuperGarantias(rsSuperG, True)
    Else
        Call CargarSuperGarantias(rsSuperG)
    End If
    'END MADM
    
    'peac 20071130 activarlo despues
    Call CargarClaseInmueble(rsClaseInmueble) ' peac 20071116
    Call CargarCategoria(rsCategoria) ' peac 20071116
    
    Call CambiaTamañoCombo(CboGarantia)
    '-------------------------------------------------------------------------------------
    
    '*** BRG0 20111125 *********************
    Call Llenar_Combo_con_Recordset(rsClaseMueble, cboClaseMueble)
    Dim R As ADODB.Recordset
        Set oGarant = New COMDCredito.DCOMGarantia
        Set R = oGarant.RecuperaTiposDocumGarantias(11)
        Set oGarant = Nothing
        CmbDocGarant.Clear
        Do While Not R.EOF
            CmbDocCompra.AddItem R!cDocDesc & Space(150) & R!nDocTpo
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Call CambiaTamañoCombo(CmbDocCompra, 300)
    '*** END BRGO
    Call Llenar_Combo_con_Recordset(rsTpoInscripcion, Me.cboTpoInscripcion) 'EJVG20130128
    'WIOR 20150608 ***
    Set oConsSist = New COMDConstSistema.NCOMConstSistema
    fsGrupoActGarDPF = oConsSist.LeeConstSistema(495)
    Set oConsSist = Nothing
    'WIOR FIN ********
    Exit Sub

ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"

End Sub

'Private Sub CargaControles()
'Dim R As ADODB.Recordset
'Dim oConstante As COMDConstantes.DCOMConstantes
'
'    On Error GoTo ERRORCargaControles
'
'    'Carga Bancos
'    Call CargaBancos
'    'Carga Tipos de Inmuebles
'    Call CargaComboConstante(gGarantTpoInmueb, CboTipoInmueb)
'
'    'Carga Ubicaciones Geograficas
'        Call CargaUbicacionesGeograficas
'    'Carga Tipos de Garantia
'        Call CambiaTamañoCombo(CmbTipoGarant)
'
'        Call CargaComboConstante(gPersGarantia, CmbTipoGarant)
'    'Carga Monedas
'        Call CargaComboConstante(gMoneda, cmbMoneda)
'
'    'Carga Relacion de Personas con Garantia
'        Set oConstante = New COMDConstantes.DCOMConstantes
'        FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelGarantia)
'        Set oConstante = Nothing
'
'
'    ' CMACICA_CSTS - 25/11/2003 ----------------------------------------------------------
'    'Carga Tipos de Documentos para el detalle de una Declaracion Jurada
'        Set oConstante = New COMDConstantes.DCOMConstantes
'        FEDeclaracionJur.CargaCombo oConstante.RecuperaConstantes(gColocPigTipoDocumento)
'        Set oConstante = Nothing
'
'    CargarSuperGarantias
'    Call CambiaTamañoCombo(CboGarantia)
'
'    '-------------------------------------------------------------------------------------
'        Exit Sub
'
'ERRORCargaControles:
'        MsgBox Err.Description, vbCritical, "Aviso"
'
'End Sub

Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TGarantiaTipoCombo)
Dim i As Long
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboProv
            cmbPersUbiGeo(1).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(1).Clear
                For i = 0 To ContNiv2 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel2(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(1).AddItem Nivel2(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel3(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel4(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
    End Select
End Sub

Private Sub cboEstadoTasInm_Change()

End Sub

Private Sub cboEstadoTasVeh_Change()

End Sub

Private Sub CboGarantia_Click()
    If CboGarantia.ListIndex <> -1 Then

        ' si son garantias preferidas nose puede cambiar a otra garantia peac 20071112
        If lcGar = "A" Then
            'CboGarantia.ListIndex = 0
        End If

        Call ReLoadCmbTipoGarant(CboGarantia.ItemData(CboGarantia.ListIndex))
        
    'MADM 20100901
    Me.txtMontoRea.Visible = True
    Me.txtMontotas.Visible = True
    'END MADM
    
     End If
     
End Sub

Private Sub CboGarantia_GotFocus()
    
    If CboGarantia.ListIndex = 0 Then
        lcGar = "A" 'Garantía inscrita
    Else
        lcGar = "B"
    End If
End Sub

Private Sub CboGarantia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbTipoGarant.SetFocus
    End If
    'MADM 20100901
    Me.txtMontoRea.Visible = True
    Me.txtMontotas.Visible = True
    'END MADM
End Sub

'MAVM 20100723
Private Sub CargarMonto_GarantJoyas(ByVal cJoyasCod As String)
    Dim oColP As COMDColocPig.DCOMColPContrato
    Dim rsColP As ADODB.Recordset
    Dim nTasacion As Double
    
    Set oColP = New COMDColocPig.DCOMColPContrato
    Set rsColP = oColP.GetDatosRegJoyas(Left(cJoyasCod, 8))
    
    If rsColP.EOF Then Exit Sub
    nTasacion = (rsColP!nTasacion)
    Set oColP = Nothing
               
    llenar_montos_garantias nTasacion, 0.85, 0
    Me.cmbMoneda.ListIndex = 0
    Me.cmbMoneda.Locked = True
    rsColP.Close
End Sub

Sub llenar_montos(ByVal CTA As String)
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
'    Dim oCred As COMDCredito.DCOMGarantia
    Dim rsCap As ADODB.Recordset
'    Dim rscredApro As ADODB.Recordset
'    Dim rscredDes As ADODB.Recordset
    Dim nFormaRetiro As Integer
'    Dim nCredNumVig As Integer
'    Dim nCredNumCancelados As Integer
    Dim nSaldoDis As Double
    Dim nIntPagado As Double
'    Dim nCondicion As Integer
    
    'MADM 20110608
    Dim fnVarTasaGarPF As Double
    Dim loParam As DColPCalculos
    Set loParam = New DColPCalculos
    'END MADM
    
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
'Comentado x MADM 20110608 - RQ11163
'    Set oCred = New COMDCredito.DCOMGarantia
'    Set rscredApro = New ADODB.Recordset
'    Set rscredDes = New ADODB.Recordset
    
    'WIOR 20140912 ****************
'    Dim sCampana As String
'    Dim sCampanaArr() As String
'    Dim i As Long
    Dim bAceptCamp As Boolean
    Dim oConsSist As COMDConstSistema.NCOMConstSistema
    Set oConsSist = New COMDConstSistema.NCOMConstSistema
'    sCampana = oConsSist.LeeConstSistema(479)
'    bAceptCamp = False
'    sCampanaArr = Split(sCampana, ",")
'    For i = 0 To UBound(sCampanaArr)
'       If CLng(sCampanaArr(i)) = fnCampanaCred Then
'            bAceptCamp = True
'            Exit For
'       End If
'    Next i

    
    'WIOR FIN *********************
    'WIOR 20150213 ****************
    Dim nCampanaPF As Integer
    nCampanaPF = oConsSist.LeeConstSistema(488)
    bAceptCamp = IIf(nCampanaPF = 1, True, False)
    Set oConsSist = Nothing
    'WIOR FIN *********************
    Set rsCap = oCap.GetDatosCuentaPF(Left(CTA, 18))
               
          If rsCap.EOF Then Exit Sub
          
               nFormaRetiro = (rsCap!nFormaRetiro)
               nSaldoDis = (rsCap!nSaldoDisp)
               nIntPagado = (rsCap!nIntPag)
               Set oCap = Nothing
               
'               Set rscredApro = oCred.RecuperaCondcionCreditoApro(Me.LblPersCodEmi.Caption, gdFecSis)
'               Set rscredDes = oCred.RecuperaCondicionCreditoDes(Me.LblPersCodEmi.Caption, gdFecSis)
               
'              If rscredApro.EOF Or rscredDes.EOF Then Exit Sub
              
                 'bValidaPF = True
'                 nCredNumVig = (rscredApro!nTotal)
'                 nCredNumCancelados = (rscredDes!nTotal)
                
'                        If nCredNumVig = 0 And nCredNumCancelados = 0 Then
'                              nCondicion = COMDConstantes.gColocCredCondNormal
'                        Else
'                             If nCredNumVig > 0 And nCredNumCancelados = 0 Or nCredNumVig > 0 And nCredNumCancelados > 0 Then
'                                  nCondicion = COMDConstantes.gColocCredCondParalelo
'                              Else
'                                  nCondicion = COMDConstantes.gColocCredCondRecurrente
'                              End If
'                        End If
               
'               Set oCred = Nothing
'               Select Case nformaretiro
'                      Case 1, 2, 3:
'                             If nCondicion = 1 Then
'                                 llenar_montos_garantias nSaldoDis, 0.9
'                             Else
'                                 llenar_montos_garantias nSaldoDis, 0.95
'                             End If
'                       Case 4:
'                                 llenar_montos_garantias nSaldoDis, 0.9, nIntPagado
'                End Select
                'MADM 20110608
                Select Case nFormaRetiro
                      Case 1:
                        'fnVarTasaGarPF = loParam.dObtieneColocParametro(102736)'WIOR 20140912 COMENTO
                        'WIOR 20140912 ************************
                        If bAceptCamp Then
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(1027362)
                        Else
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(102736)
                        End If
                        'WIOR FIN ***************************
                        llenar_montos_garantias nSaldoDis, fnVarTasaGarPF
                      Case 2:
                        'fnVarTasaGarPF = loParam.dObtieneColocParametro(102737) 'WIOR 20140912 COMENTO
                        'WIOR 20140912 ************************
                        If bAceptCamp Then
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(1027372)
                        Else
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(102737)
                        End If
                        'WIOR FIN ***************************
                        llenar_montos_garantias nSaldoDis, fnVarTasaGarPF
                      Case 3:
                        'fnVarTasaGarPF = loParam.dObtieneColocParametro(102738) 'WIOR 20140912 COMENTO
                        'WIOR 20140912 ************************
                        If bAceptCamp Then
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(1027382)
                        Else
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(102738)
                        End If
                        'WIOR FIN ***************************
                        llenar_montos_garantias nSaldoDis, fnVarTasaGarPF, nIntPagado
                      Case 4:
                        'fnVarTasaGarPF = loParam.dObtieneColocParametro(102739) 'WIOR 20140912 COMENTO
                        'WIOR 20140912 ************************
                        If bAceptCamp Then
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(1027392)
                        Else
                            fnVarTasaGarPF = loParam.dObtieneColocParametro(102739)
                        End If
                        'WIOR FIN ***************************
                        llenar_montos_garantias nSaldoDis, fnVarTasaGarPF, nIntPagado
                End Select
                Set loParam = Nothing
                'END MADM
                If Mid(Left(CTA, 18), 9, 1) = "1" Then
                    Me.cmbMoneda.ListIndex = 0
                Else
                    Me.cmbMoneda.ListIndex = 1
                End If
            
                Me.cmbMoneda.Locked = True
            
                rsCap.Close
'                rscredApro.Close
'                rscredDes.Close
                Set rsCap = Nothing
'                Set rscredApro = Nothing
'                Set rscredDes = Nothing
            
End Sub

'madm 20100511
Sub llenar_montos_garantias(ByVal pnSaldo As Double, ByVal pnPorc As Double, Optional pnIntPagado As Double = 0)
Me.txtMontoRea.Visible = False
Me.txtMontotas.Visible = False
Me.txtMontoxGrav.Locked = True
Me.cmbMoneda.Locked = True
'madm 20100817 - variable DPF 1 a 1
Me.txtMontoRea.Text = IIf(gGarantiaDepPlazoFijoCF, Format(pnSaldo, "#0.00"), Format((pnSaldo - pnIntPagado) * pnPorc, "#0.00"))
Me.txtMontotas.Text = Format(pnSaldo, "#0.00")
Me.txtMontoxGrav.Text = IIf(gGarantiaDepPlazoFijoCF, Format(pnSaldo, "#0.00"), Format((pnSaldo - pnIntPagado) * pnPorc, "#0.00"))

End Sub


Private Sub CboNumPF_Click()
        
        Me.txtMontoRea.Text = "0.00"
        Me.txtMontotas.Text = "0.00"
        Me.txtMontoxGrav.Text = "0.00"
        
        'MAVM 20100720 Garant Joyas
        'If Me.CboNumPF.ListIndex <> -1 Then
        If Me.CboNumPF.ListIndex <> -1 And Trim(Right(Me.CmbDocGarant.Text, 4)) <> 145 Then
            llenar_montos Trim(Me.CboNumPF.Text)
        Else
           CargarMonto_GarantJoyas Trim(Me.CboNumPF.Text)
        End If

End Sub
'ALPA 20120320*****************
Private Sub cboTipoGA_Click()
    cboTipo.ListIndex = cboTipoGA.ListIndex 'IndiceListaCombo(CboTipo, IIf(Trim(Right(cboTipoGA.Text, 3)) = "", "", Trim(Right(cboTipoGA.Text, 3))))
End Sub
'******************************


Private Sub CboTipoInmueb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    '    TxtMontoHip.SetFocus
    End If
End Sub

Private Sub ChkCF_Click()
    Dim oDCF As COMDCartaFianza.DCOMCartaFianza
    If ChkCF.value = 1 Then
        If MsgBox("Desea relacionar con Credito C.F", vbInformation + vbYesNo, "AVISO") = vbYes Then
            If Not IsLoadForm("Relacion de Credito Con Garantia") Then
                bCreditoCF = True
                FrmCredRelGarant.Caption = "Credito de la Carta Fianza"
                FrmCredRelGarant.Show vbModal
            End If
            
            'valida el credito
            If pgcCtaCod <> "" Then
                Set oDCF = New COMDCartaFianza.DCOMCartaFianza
                    If oDCF.ValidadCreditoCF(pgcCtaCod) = False Then
                        MsgBox "El Credito no corresponde a una" & vbCrLf & " Carta Fianza", vbInformation, "AVISO"
                        bValdiCCF = False
                    Else
                        bValdiCCF = True
                    End If
                Set oDCF = Nothing
            End If
        End If
    End If
End Sub

Private Sub chkDocCompra_Click()
    If chkDocCompra.value = 1 Then
        SSGarant.TabVisible(7) = True
        cmdBuscaEmisorDoc.Enabled = True
        If CmbTipoGarant.Text <> "" Then
            If CInt(Right(CmbTipoGarant.Text, 2)) = 11 Then
                CmbDocCompra.ListIndex = CmbDocGarant.ListIndex
                Me.txtNumDocCompra.Text = txtNumDoc.Text
                txtValorDocCompra.Text = txtMontotas.Text
            End If
        End If
    Else
        SSGarant.TabVisible(7) = False
        cmdBuscaEmisorDoc.Enabled = False
    End If
End Sub

'By Capi 01102008
Private Sub chkEnTramite_Click()
     txtFechaCertifGravamen = "__/__/____"
     txtFechaCertifGravamenGA = "__/__/____"
     txtValorGravado = "0.00"
     txtValorGravadoGA = "0.00"
     cboTpoInscripcion.ListIndex = -1 'EJVG20130128
     
    If chkEnTramite = 1 Then
        txtFechaCertifGravamen.Enabled = False
         txtFechaCertifGravamenGA.Enabled = False
        txtValorGravado.Enabled = False
        txtValorGravadoGA.Enabled = False
        cboTpoInscripcion.Enabled = False 'EJVG20130128
    Else
        txtFechaCertifGravamen.Enabled = True
        txtFechaCertifGravamenGA.Enabled = True
        txtValorGravado.Enabled = True
        txtValorGravadoGA.Enabled = True
        cboTpoInscripcion.Enabled = True 'EJVG20130128
    End If
End Sub

'ALPA 20120320
Private Sub chkEnTramiteGA_Click()
     txtFechaCertifGravamen = "__/__/____"
     txtFechaCertifGravamenGA = "__/__/____"
     txtValorGravado = "0.00"
     txtValorGravadoGA = "0.00"
     
    If chkEnTramiteGA = 1 Then
        txtFechaCertifGravamen.Enabled = False
        txtFechaCertifGravamenGA.Enabled = False
        txtValorGravado.Enabled = False
        txtValorGravadoGA.Enabled = False
    Else
        txtFechaCertifGravamen.Enabled = True
        txtFechaCertifGravamenGA.Enabled = True
        txtValorGravado.Enabled = True
        txtValorGravadoGA.Enabled = True
    End If
End Sub

Private Sub ChkGarPoliza_Click()
'WIOR 20130122 agrego la condicion fbGrupoAdmCred
If fbGrupoAdmCred Then
    If ChkGarPoliza.value = 1 Then
        SSGarant.TabVisible(6) = True
    Else
        SSGarant.TabVisible(6) = False
    End If
Else
    If ChkGarPoliza.value = 1 Then
        SSGarant.TabVisible(6) = True
        Me.chkGarPolizaMob.Enabled = False
    Else
        SSGarant.TabVisible(6) = False
        Me.chkGarPolizaMob.Enabled = True
    End If
End If
End Sub

Private Sub chkGarPolizaMob_Click()
'WIOR 20130122 agrego la condicion fbGrupoAdmCred
If fbGrupoAdmCred Then
    If chkGarPolizaMob.value = 1 Then
        SSGarant.TabVisible(8) = True
    Else
        SSGarant.TabVisible(8) = False
    End If
Else
    If chkGarPolizaMob.value = 1 Then
        SSGarant.TabVisible(8) = True
        Me.ChkGarPoliza.Enabled = False
    Else
        SSGarant.TabVisible(8) = False
        Me.ChkGarPoliza.Enabled = True
    End If
End If
End Sub

Private Sub ChkGarReal_Click()
    If ChkGarReal.value = 1 Then
        If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
            SSGarant.TabVisible(9) = True 'ALPA 20120320
            SSGarant.TabVisible(2) = False 'ALPA 20120320
            Frame9.Enabled = True
            Frame8.Enabled = True
        Else
            SSGarant.TabVisible(9) = False 'ALPA 20120320
            Call LimpiaGarantiaReal
            SSGarant.TabVisible(2) = True
            '*** PEAC 20090724
            'SSGarant.TabVisible(4) = True
            
            'SSGarant.Tab = 2
            
            ChkGarPoliza.Visible = True
            
            Me.ChkTasacion.Enabled = True
            
            '*** PEAC 20090724
            ChkTasacion.Visible = True
            
            '*** BRGO 20111125 ******
            chkDocCompra.Visible = True
            chkDocCompra.Enabled = True 'BRGO 20111125
            chkGarPolizaMob.Visible = True
            chkGarPolizaMob.Enabled = True 'BRGO 20111125
            ChkGarPoliza.Enabled = True
            '*** END BRGO ************
            
            Call HabilitaIngresoGarantReal(True)
            'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
            '    FraDatInm.Enabled = True
            'Else
            '    FraDatInm.Enabled = False
            'End If
        End If
    Else
        SSGarant.TabVisible(9) = False 'ALPA 20120320
        SSGarant.Tab = 1
        SSGarant.TabVisible(2) = False
        
        '*** PEAC 20090724
        SSGarant.TabVisible(4) = False
        
        ChkGarPoliza.Visible = False
        Me.ChkTasacion.Enabled = False
        Me.ChkTasacion.Enabled = False
        
        '*** BRGO 20111125 *******
        chkDocCompra.Visible = False
        chkDocCompra.Enabled = False
        chkGarPolizaMob.Visible = False
        chkGarPolizaMob.Enabled = False
        ChkGarPoliza.value = vbUnchecked
        chkGarPolizaMob.value = vbUnchecked
        SSGarant.TabVisible(7) = False
        SSGarant.TabVisible(8) = False
        
        '*** END BRGO ************
        
        'ChkTasacion.Visible = False '*** PEAC 20090724
        
        Call HabilitaIngresoGarantReal(False)
    End If
    
End Sub

Private Sub ChkTasacion_Click()
    If ChkTasacion.value = 1 Then
        SSGarant.TabVisible(4) = True
    Else
        SSGarant.TabVisible(4) = False
    End If
End Sub

Private Sub cmbCategoria_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cmbClaseInmueble_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub CmbDocGarant_Click()
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oColP As COMDColocPig.DCOMColPContrato
    Dim rsCtas As ADODB.Recordset
    Dim nSaldoDisC As Double
    
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set oColP = New COMDColocPig.DCOMColPContrato
    Set rsCtas = New ADODB.Recordset
       
    If CmbDocGarant.Enabled = True Then
        If Len(CmbDocGarant) > 0 Then
            'If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then  'Declaracion Jurada
            If CInt(Trim(Right(CmbDocGarant, 10))) = gnDocumentoGarantia Then  'Declaracion Jurada
                    Call LimpiaFlex(FEDeclaracionJur)
                    LblTotDJ.Caption = "0.00"
                    SSGarant.TabVisible(3) = True
                    'SSGarant.Tab = 3
                Else
                    Call LimpiaFlex(FEDeclaracionJur)
                    'SSGarant.Tab = 1
                    SSGarant.TabVisible(3) = False
                End If
        End If
    End If
    
    'madm 20100513 --------------------------------------------------------------
    If cmdEjecutar = 1 Then
        'ALPA 20140315***********************************************************************************
        'If Trim(Right(Me.CmbDocGarant.Text, 4)) = "17" And CmbTipoGarant.ListIndex = 0Then
        If Trim(Right(Me.CmbDocGarant.Text, 4)) = "17" And (CmbTipoGarant.ListIndex = 0 Or lnTipoGarantiaActual = "39") Then
            If Me.LblPersCodEmi.Caption <> "" Then
                'MAVM 20120621 ***
                'Set rsCtas = oCap.GetDatosCuentaPFCodPer(Me.LblPersCodEmi.Caption)
                If FERelPers.Rows = 2 And FERelPers.TextMatrix(1, 1) = "" Then
                    MsgBox "Agregue el Titular de la Garantia", vbInformation
                    CmbDocGarant.ListIndex = -1
                    CmdCliNuevo.SetFocus
                Else
                Dim i As Integer
                For i = 1 To FERelPers.Rows - 1
                    If CInt(Trim(Right(FERelPers.TextMatrix(i, 3), 15))) = gPersRelGarantiaTitular Then
                        Set rsCtas = oCap.GetDatosCuentaPFCodPer(FERelPers.TextMatrix(i, 1))
                        
                        Me.txtNumDoc.Visible = False
                        Me.CboNumPF.Visible = True
                        Me.CboNumPF.Enabled = True
                        If rsCtas.EOF Or rsCtas.BOF Then
                            'MsgBox "Emisor no Registra Plazo Fijos", vbInformation
                            MsgBox "Titular no Registra Plazo Fijos", vbInformation
                        Else
                            CargaCuentas rsCtas
                            Exit Sub
                        End If
                        
                    End If
                Next i
                    If CboNumPF.Enabled = False Then
                        MsgBox "Agregue el Titular de la Garantia", vbInformation
                        CmbDocGarant.ListIndex = -1
                        CmdCliNuevo.SetFocus
                    End If
                End If
                '***
            End If
        ElseIf Trim(Right(Me.CmbDocGarant.Text, 4)) = "17" And CmbTipoGarant.ListIndex = 1 Then
            If Me.LblPersCodEmi.Caption <> "" Then
                'MADM 20111025 optional
                Set rsCtas = oCap.GetDatosCuentaPFCodPer(Me.LblPersCodEmi.Caption, 1)
                Me.txtNumDoc.Visible = False
                Me.CboNumPF.Visible = True
                Me.CboNumPF.Enabled = True
                If rsCtas.EOF Or rsCtas.BOF Then
                    MsgBox "Emisor no Registra Plazo Fijos", vbInformation
                Else
                    CargaCuentas rsCtas
                    Exit Sub
                End If
            End If
        Else
                Me.txtNumDoc.Visible = True
                Me.CboNumPF.Visible = False
                Me.CboNumPF.Enabled = False
        End If
    End If
    
    '*** MAVM 20100720
    If cmdEjecutar = 1 Then
        If Trim(Right(Me.CmbDocGarant.Text, 4)) = "145" Then
            If Me.LblPersCodEmi.Caption <> "" Then
                Set rsCtas = oColP.GetDatosRegJoyasCodPer(Me.LblPersCodEmi.Caption)
                Me.txtNumDoc.Visible = False
                Me.CboNumPF.Visible = True
                Me.CboNumPF.Enabled = True
                If rsCtas.EOF Or rsCtas.BOF Then
                    MsgBox "Emisor no tiene Joyas Registradas", vbInformation
                Else
                    CargaCuentas rsCtas
                End If
            End If
        Else
                Me.txtNumDoc.Visible = True
                Me.CboNumPF.Visible = False
                Me.CboNumPF.Enabled = False
        End If
    End If
    '***
End Sub

Private Sub CargaCuentas(ByVal prsCtas As ADODB.Recordset)
    Me.CboNumPF.Clear
    Do While Not prsCtas.EOF
        CboNumPF.AddItem PstaNombre(prsCtas!cPersNombre) & Space(150) & prsCtas!cPersCod
        prsCtas.MoveNext
    Loop
End Sub

Private Sub CmbDocGarant_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        'txtNumDoc.SetFocus 'cmbMoneda.SetFocus
     End If
End Sub

Private Sub cmbMoneda_Click()
    Call CmbMoneda_KeyPress(13)
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        'If txtNumDoc.Enabled Then
            'txtNumDoc.SetFocus
        'End If
        If CmdCliNuevo.Enabled And CmdCliNuevo.Visible Then
            CmdCliNuevo.SetFocus
        End If
     End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
'        Select Case Index
'            Case 0 'Combo Dpto
'                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboProv)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(2).Clear
'                    cmbPersUbiGeo(3).Clear
'                End If
'            Case 1 'Combo Provincia
'                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboDist)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(3).Clear
'                End If
'            Case 2 'Combo Distrito
'                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboZona)
'        End Select
Dim oUbic As comdpersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim i As Integer

If Index = 3 Then Exit Sub

Set oUbic = New comdpersona.DCOMPersonas

Set rs = oUbic.CargarUbicacionesGeograficas(, Index + 2, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))

For i = Index + 1 To cmbPersUbiGeo.Count - 1
    cmbPersUbiGeo(i).Clear
Next

While Not rs.EOF
    cmbPersUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
    rs.MoveNext
Wend

Set oUbic = Nothing
End Sub


Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
        Select Case Index
            Case 0
                cmbPersUbiGeo(1).SetFocus
            Case 1
                cmbPersUbiGeo(2).SetFocus
            Case 2
                cmbPersUbiGeo(3).SetFocus
            Case 3
                txtDireccion.SetFocus 'txtMontotas.SetFocus
        End Select
     End If
End Sub

Private Sub cmbTipoCreditoTabla_Change()
'Dim oGarant As COMDCredito.DCOMGarantia
'Set oGarant = New COMDCredito.DCOMGarantia
'Dim rs As ADODB.Recordset
'
'Set rs = oGarant.RecuperaTablaValores(CInt(Trim(Right(cmbTipoCreditoTabla, 10))))
'
'While Not rs.EOF
'    With FeTabla
'        .TextMatrix(rs.Bookmark, 0) = rs!nTipoEval
'        .TextMatrix(rs.Bookmark, 1) = rs!cConsDescripcion
'    End With
'    rs.MoveNext
'Wend
'Set oGarant = Nothing
Call cmbTipoCreditoTabla_Click
End Sub

Private Sub cmbTipoCreditoTabla_Click()
Dim oGarant As COMDCredito.DCOMGarantia
Set oGarant = New COMDCredito.DCOMGarantia
Dim rs As ADODB.Recordset

If cmbTipoCreditoTabla.Text = "" Then Exit Sub
Set rs = oGarant.RecuperaTablaValores(CInt(Trim(Right(cmbTipoCreditoTabla, 10))))

FeTabla.Clear
FeTabla.Rows = 2
FeTabla.FormaCabecera

While Not rs.EOF
    With FeTabla
        .AdicionaFila
        .TextMatrix(rs.Bookmark, 0) = rs!nTipoEval
        .TextMatrix(rs.Bookmark, 1) = rs!cConsDescripcion
        .TextMatrix(rs.Bookmark, 3) = 0
    End With
    rs.MoveNext
Wend
Set oGarant = Nothing

End Sub

Private Sub CmbTipoGarant_Change()
    txtMontotas.Text = "0.00"
    txtMontoRea.Text = "0.00"
    txtMontoxGrav.Text = "0.00"
End Sub

Private Sub CmbTipoGarant_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
lnTipoGarantiaActual = "00" 'ALPA201
    'Carga Tipos de Documentos de Garantia
    
        txtMontotas.Text = "0.00"
        txtMontoRea.Text = "0.00"
        txtMontoxGrav.Text = "0.00"
    
        If CmbTipoGarant.ListIndex = -1 Then
            If Not bCarga Then
                MsgBox "Debe Escoger un Tipo de Garantia", vbInformation, "Aviso"
                Exit Sub
            Else
                Exit Sub
            End If
        End If
        Set oGarantia = New COMDCredito.DCOMGarantia
        Set R = oGarantia.RecuperaTiposDocumGarantias(CInt(Right(CmbTipoGarant.Text, 2)))
        Set oGarantia = Nothing
        CmbDocGarant.Clear
        Do While Not R.EOF
            'MADM 20100827
            If gGarantiaDepPlazoFijoCF Then
                If (R!nDocTpo <> 121) Then
                CmbDocGarant.AddItem R!cDocDesc & Space(150) & R!nDocTpo
                End If
            Else
                CmbDocGarant.AddItem R!cDocDesc & Space(150) & R!nDocTpo
            End If
            'END MADM
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Call CambiaTamañoCombo(CmbDocGarant, 300)
        
        'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Or CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaDepositosGarantia Then
        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Then
            CboBanco.Enabled = True
        Else
            CboBanco.Enabled = False
            CboBanco.ListIndex = -1
        End If

        '*** PEAC 20080513
        
        If CInt(Right(CmbTipoGarant.Text, 2)) = 1 Then 'si es garantia hipoteca
            Me.FraDatInm.Visible = True
            fraDatVehic.Visible = False
        End If
        
        If CInt(Right(CmbTipoGarant.Text, 2)) = 2 Then
            fraDatVehic.Caption = "Maquinaria y Equipo"
            Me.FraDatInm.Visible = False
            fraDatVehic.Visible = True
        
            txtPlacaVehic.Visible = False
            lblPlacaVehic.Visible = False
            '***---------------------------
           'By Capi 25092008 porque se elimino el control txtFecTasVeh
            'txtFecTasVeh.Visible = True
            
            txtDescrip.Visible = True
            txtNumMotor.Visible = True
            txtNumSerie.Visible = True
            txtAnioFab.Visible = True
            '-------
            'By Capi 25092008 porque se elimino el control lblFecTasVeh
            'lblFecTasVeh.Visible = True
            
            lblDescrip.Visible = True
            lblNumMotor.Visible = True
            lblNumSerie.Visible = True
            lblAnioFab.Visible = True
            
        
            '***---------------------------
            txtValorMerca.Visible = False
            txtTipoMerca.Visible = False
            txtDireAlma.Visible = False
            '----
            lblValorMerca.Visible = False
            lblTipoMerca.Visible = False
            lblDireAlma.Visible = False
        
        End If
        
        If CInt(Right(CmbTipoGarant.Text, 2)) = 11 Then
            fraDatVehic.Caption = "Vehiculos"
            Me.FraDatInm.Visible = False
            fraDatVehic.Visible = True

            '***---------------------------
            txtPlacaVehic.Visible = True
            
            'By Capi 25092008 porque se elimino el control txtFecTasVeh
            'txtFecTasVeh.Visible = True
            
            txtDescrip.Visible = True
            txtNumMotor.Visible = True
            txtNumSerie.Visible = True
            txtAnioFab.Visible = True
            '-------
            lblPlacaVehic.Visible = True
            
            'By Capi 25092008 porque se elimino el control lblFecTasVeh
            'lblFecTasVeh.Visible = True
            
            
            lblDescrip.Visible = True
            lblNumMotor.Visible = True
            lblNumSerie.Visible = True
            lblAnioFab.Visible = True
            
            '***---------------------------
            txtValorMerca.Visible = False
            txtTipoMerca.Visible = False
            txtDireAlma.Visible = False
            '----
            lblValorMerca.Visible = False
            lblTipoMerca.Visible = False
            lblDireAlma.Visible = False

        End If
        
        If CInt(Right(CmbTipoGarant.Text, 2)) = 12 Then
            fraDatVehic.Caption = "Mercadería"
            Me.FraDatInm.Visible = False
            fraDatVehic.Visible = True
            
            '***---------------------------
            txtPlacaVehic.Visible = False
            
            'By Capi 25092008 porque se elimino el control txtFecTasVeh
            'txtFecTasVeh.Visible = True
            
            txtDescrip.Visible = False
            txtNumMotor.Visible = False
            txtNumSerie.Visible = False
            txtAnioFab.Visible = False
            
            '-------
            lblPlacaVehic.Visible = False
            
            'By Capi 25092008 porque se elimino el control lblFecTasVeh
            'lblFecTasVeh.Visible = True
            
            lblDescrip.Visible = False
            lblNumMotor.Visible = False
            lblNumSerie.Visible = False
            lblAnioFab.Visible = False
            
            '***---------------------------
            txtValorMerca.Visible = True
            txtTipoMerca.Visible = True
            txtDireAlma.Visible = True
            '----
            lblValorMerca.Visible = True
            lblTipoMerca.Visible = True
            lblDireAlma.Visible = True
            
        End If
        

        'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
        '    FraDatInm.Enabled = True
        'Else
        '    FraDatInm.Enabled = False
        'End If
        'If FraDatInm.Visible Then
        '    FraDatInm.Enabled = True
        'End If
        'If fraDatVehic.Visible Then
        '    fraDatVehic.Enabled = True
        'End If
        'CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------
        'DECLARACION JURADA
'        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
'            If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then  'Declaracion Jurada
'                Call LimpiaFlex(FEDeclaracionJur)
'                LblTotDJ.Caption = "0.00"
'                SSGarant.TabVisible(3) = True
'                'SSGarant.Tab = 3
'            Else
'                Call LimpiaFlex(FEDeclaracionJur)
'                'SSGarant.Tab = 1
'                SSGarant.TabVisible(3) = False
'            End If
'        End If
        '--------------------------------------------------------------------------------------------------------
        Call CambiaTamañoCombo(CmbTipoGarant, 300)
        Call HabilitarFramesGarantiaReal(CInt(Trim(Right(CmbTipoGarant, 10))))
        
        'ARCV 11-07-2006
        Call HabilitarTablaDeValores(CInt(Trim(Right(CmbTipoGarant, 10))))
        lnTipoGarantiaActual = CInt(Right(CmbTipoGarant.Text, 2))
        If lnTipoGarantiaActual <> "39" Then
            frBF.Visible = False
        Else
            frBF.Visible = True
        End If
End Sub

Private Sub HabilitarTablaDeValores(ByVal pnSubTipoGarantia As Integer)
Dim oGarant As COMDCredito.DCOMGarantia
Set oGarant = New COMDCredito.DCOMGarantia
Dim bEsTablaValores As Boolean

bEsTablaValores = oGarant.EsSubTipoGarantiaTablaValores(pnSubTipoGarantia)

SSGarant.TabVisible(5) = bEsTablaValores
framontos.Enabled = Not bEsTablaValores
    
Set oGarant = Nothing

End Sub

Private Sub HabilitarFramesGarantiaReal(ByVal pnSubTipoGarantia As Integer)

Dim oGarant As COMDCredito.DCOMGarantia
Dim bEsSubTipoGInmueble As Boolean
Set oGarant = New COMDCredito.DCOMGarantia

bEsSubTipoGInmueble = oGarant.EsSubTipoGarantiaInmueble(pnSubTipoGarantia)

If bEsSubTipoGInmueble Then
    FraDatInm.Visible = True
'    fraDatVehic.Visible = False
Else

'    fraDatVehic.Visible = False
'    Me.fraDatMaqEquipo.Visible = True


    FraDatInm.Visible = False
'   fraDatVehic.Visible = True
    
    'poner aqui condicionante 'peac 20071128
    If gcPermiteModificar Then
        If fraDatVehic.Visible Then
            fraDatVehic.Enabled = True
            
            'By Capi 25092008 porque se elimino el control txtFecTasVeh y cboEstadoTasVeh
            'txtFecTasVeh.Enabled = True
            'cboEstadoTasVeh.Enabled = True
            
            txtPlacaVehic.Enabled = True
        End If
    End If
    
End If

Set oGarant = Nothing
End Sub

Private Sub CmbTipoGarant_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        CmbDocGarant.SetFocus
     End If
End Sub

Private Sub cmdAceptar_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
Dim RelPers() As String
Dim GarDetDJ() As Variant
Dim oMantGarant As COMDCredito.DCOMGarantia
Dim lsNumGarant As String
Dim i As Long
Dim lrs As ADODB.Recordset
'* nuevos campos *
Dim nEstadoTasacion As Integer
Dim dFechaTasacion As Date
Dim dFechaCertifGravamen As Date 'peac 20071123
Dim bVerificaDescobertura As Boolean

'Dim vPol As Polizzas
Dim vPol() As Variant
Dim vDatosGar() As Variant
Dim rsActualizacion As ADODB.Recordset  'WIOR 20120616
'*********

    On Error GoTo ErrorCmdAceptar_Click
    If Not ValidaDatos Then
        Exit Sub
    End If

    If MsgBox("Se va a Grabar los Datos, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    If FERelPers.Rows = 2 And FERelPers.TextMatrix(1, 1) = "" Then
        ReDim RelPers(0, 0)
    
    Else
        ReDim RelPers(FERelPers.Rows - 1, 4)
        For i = 1 To FERelPers.Rows - 1
            RelPers(i - 1, 0) = FERelPers.TextMatrix(i, 1) 'Codigo de Persona
            RelPers(i - 1, 1) = Trim(Right(CmbDocGarant.Text, 10)) 'Tipo de Doc de Garantia
            'madm 20100512
            RelPers(i - 1, 2) = IIf(Me.txtNumDoc.Visible, txtNumDoc.Text, Left(Me.CboNumPF.Text, 18)) 'Numero de Documento
            'end madm
            RelPers(i - 1, 3) = Right("00" & Trim(Right(FERelPers.TextMatrix(i, 3), 10)), 2) 'Relacion
        Next i
    End If

    ' CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------
    If FEDeclaracionJur.Rows = 2 And FEDeclaracionJur.TextMatrix(1, 1) = "" Then
        ReDim GarDetDJ(0, 0)
    Else
        ReDim GarDetDJ(FEDeclaracionJur.Rows - 1, 6)
        For i = 1 To FEDeclaracionJur.Rows - 1
            GarDetDJ(i - 1, 0) = FEDeclaracionJur.TextMatrix(i, 0) 'Item
            'GarDetDJ(i - 1, 1) = FEDeclaracionJur.TextMatrix(i, 1) 'Descripcion del Item
            GarDetDJ(i - 1, 1) = Replace(FEDeclaracionJur.TextMatrix(i, 1), "'", "''") 'Descripcion del Item
            GarDetDJ(i - 1, 2) = FEDeclaracionJur.TextMatrix(i, 2) 'Cantidad del Item
            GarDetDJ(i - 1, 3) = FEDeclaracionJur.TextMatrix(i, 3) 'Precio Unit. del Item
            GarDetDJ(i - 1, 4) = Trim(Right(FEDeclaracionJur.TextMatrix(i, 4), 4)) 'Tipo de Doc. del Item
            GarDetDJ(i - 1, 5) = FEDeclaracionJur.TextMatrix(i, 5) 'Nro. Doc. del Item
        Next i
    End If
    ' --------------------------------------------------------------------------------------------------------

    '** Nuevos Campos **
    If ChkGarReal.value = 1 Then

        'ARCV 27-01-2007
        'If FraDatInm.Visible = True Then
        '    nEstadoTasacion = CInt(Trim(Right(cboEstadoTasInm.Text, 5)))
        'Else
        '    nEstadoTasacion = CInt(Trim(Right(cboEstadoTasVeh.Text, 5)))
        'End If
        
        'By Capi 10102008
        
'        'peac 20071128
'        If fraDatVehic.Visible = True Then
'            nEstadoTasacion = CInt(Trim(Right(cboEstadoTasVeh.Text, 5)))
'        Else
'            nEstadoTasacion = 0
'        End If
        
        '
        
        'By Capi 10102008
'        If FraDatInm.Visible = True Then
'            dFechaTasacion = CDate(IIf(txtFechaTasInm.Text = "__/__/____", "01/01/1950", txtFechaTasInm.Text))
'
'            'dFechaCertifGravamen = CDate(txtFechaCertifGravamen.Text)
'        End If
'        If Me.fraDatVehic.Visible = True Then
'            dFechaTasacion = CDate(IIf(txtFecTasVeh.Text = "__/__/____", "01/01/1950", txtFecTasVeh.Text))
'
'        End If
        'end by

'        If Me.fraDatMaqEquipo.Visible = True Then
'            dFechaTasacion = CDate(Me.txtFecTasMaq.Text)
'        End If
        
       If Trim(CboGarantia.Text) <> "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
            If Me.cboTipo.ListIndex = -1 Then
                MsgBox "Debe de Seleccionar el Tipo de Contrato", vbInformation, "AVISO"
                Exit Sub
            End If
        Else
            'ALPA 20120320
            If Me.cboTipoGA.ListIndex = -1 And Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
                MsgBox "Debe de Seleccionar el Tipo de Contrato", vbInformation, "AVISO"
                Exit Sub
            End If
        End If

        
    End If
    '*****
    
    'ARCV 11-07-2006
    Dim nTipoTablaValor As Integer
    
    If cmbTipoCreditoTabla.Text <> "" Then
        nTipoTablaValor = CInt(Trim(Right(cmbTipoCreditoTabla, 10)))
    Else
        nTipoTablaValor = 0
    End If

    '*** peac 20071221 - variables utilizadas para gravar datos de las polizas
        If SSGarant.TabVisible(6) = True Then
        
            Dim lcClaInm As String, lcCatInm As String, lcNumLoc As String
            Dim lcNumPis As String, lcNumSot As String, lcAniCon As String
            Dim lcPoliBF As String 'ALPA20140203***********************
            Dim lcFecPBF As String 'ALPA20140203***********************
            
            lcClaInm = Trim(Right(Me.cmbClaseInmueble.Text, 2))
            lcCatInm = Trim(Right(Me.cmbCategoria.Text, 2))
            lcNumLoc = Me.txtNumLocales.Text
            lcNumPis = Me.txtNumPisos.Text
            lcNumSot = Me.txtNumSotanos.Text
            lcAniCon = Me.txtAnioConstruccion
            lcPoliBF = IIf(ckPolizaBF.value = 1, 1, 0) 'ALPA20140203
            lcFecPBF = txtFechaPBF.Text 'ALPA20140203
            
            
            If Len(lcClaInm) = 0 Or Len(lcCatInm) = 0 Or Len(lcNumLoc) = 0 Or Len(lcNumPis) = 0 _
                Or Len(lcNumSot) = 0 Or Len(lcAniCon) = 0 Then
                
                MsgBox "Falta ingresar datos en la pestaña Garantía-Póliza Inmueble.", 64, "Aviso"
                Exit Sub
            End If
            
            'ReDim vPol(1, 6)
            ReDim vPol(1, 8)
            vPol(1, 0) = 1
            vPol(1, 1) = CInt(lcClaInm)
            vPol(1, 2) = CInt(lcCatInm)
            vPol(1, 3) = CInt(lcNumLoc)
            vPol(1, 4) = CInt(lcNumPis)
            vPol(1, 5) = CInt(lcNumSot)
            vPol(1, 6) = CInt(lcAniCon)
            vPol(1, 7) = CInt(lcPoliBF) 'ALPA20140203
            vPol(1, 8) = CDate(lcFecPBF) 'ALPA20140203

        ElseIf SSGarant.TabVisible(8) = True Then 'BRGO 20111205
            Dim lcClaMue As String, lcAniFab As String
            
            lcClaMue = Trim(Right(Me.cboClaseMueble.Text, 2))
            lcAniFab = Me.txtAnioFabricacion
            If Len(lcClaMue) = 0 Or Len(lcAniFab) = 0 Then
                MsgBox "Falta ingresar datos en la pestaña Garantía-Póliza Mobiliaria.", 64, "Aviso"
                Exit Sub
            End If
            
            ReDim vPol(1, 2)
            vPol(1, 0) = 2
            vPol(1, 1) = CInt(lcClaMue)
            vPol(1, 2) = CInt(lcAniFab)
            'END BRGO
        Else
            ReDim vPol(1, 0)
            vPol(1, 0) = 0
        End If
    '*******************************************************
    'ALPA 20120322*****************************************************************************************************
    'ReDim vDatosGar(1, 15)
    'ReDim vDatosGar(1, 16)
    ReDim vDatosGar(1, 17) 'EJVG20130129
    vDatosGar(1, 0) = CInt(ChkGarReal.value)
    vDatosGar(1, 1) = CInt(ChkGarPoliza.value)
    vDatosGar(1, 2) = CDbl(txtValorGravado.Text)
    vDatosGar(1, 3) = CDate(IIf(txtFechaCertifGravamen.Text = "__/__/____", "01/01/1950", txtFechaCertifGravamen.Text))
    vDatosGar(1, 4) = CDbl(txtVRM.Text)
    vDatosGar(1, 5) = CDbl(TxtValorEdificacion.Text)
    vDatosGar(1, 6) = Trim(Right(Me.cboTipo.Text, 3))
    vDatosGar(1, 7) = CDbl(Me.txtValorMerca.Text)
    vDatosGar(1, 8) = Trim(Me.txtDireAlma.Text)
    vDatosGar(1, 9) = Trim(Me.txtTipoMerca.Text)
    vDatosGar(1, 10) = Trim(Me.txtDescrip.Text)
    vDatosGar(1, 11) = Trim(Me.txtNumMotor.Text)
    vDatosGar(1, 12) = Trim(Me.txtNumSerie.Text)
    vDatosGar(1, 13) = CDate(IIf(Me.txtAnioFab.Text = "__/__/____", "01/01/1950", Me.txtAnioFab.Text))
    vDatosGar(1, 14) = CInt(ChkTasacion.value) '*** PEAC 20090724
    vDatosGar(1, 15) = CInt(chkDocCompra.value) '*** BRGO 20111205
    vDatosGar(1, 16) = CDate(IIf(txtFechaBloqueo.Text = "__/__/____", "01/01/1950", txtFechaBloqueo.Text)) 'ALPA 20120322
    vDatosGar(1, 17) = CInt(IIf(cboTpoInscripcion.ListIndex = -1, -1, Trim(Right(cboTpoInscripcion.Text, 5)))) 'EJVG20130129

    If cmdEjecutar = 1 Then
    
        Set oGarantia = New COMDCredito.DCOMGarantia
'            Call oGarantia.NuevaGarantia(CStr(Trim(Right(CmbDocGarant.Text, 10))), CStr(Trim(txtNumDoc.Text)), CStr(Trim(Right(CmbTipoGarant.Text, 10))), _
'                                         CboGarantia.ItemData(CboGarantia.ListIndex), CStr(Trim(Right(cmbMoneda.Text, 10))), _
'                                         CStr(Trim(txtDescGarant.Text)), CStr(Trim(Right(cmbPersUbiGeo(3).Text, 15))), _
'                                         CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), CStr(Trim(Txtcomentarios.Text)), _
'                                         RelPers, CDate(gdFecSis), CStr(LblPersCodEmi.Caption), CStr(IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13))), _
'                                         CInt(ChkGarReal.value), CInt(IIf(OptCG(0).value, 0, 1)), CInt(IIf(OptTR(0).value, 0, 1)), CStr(LblSegPersCod.Caption), _
'                                         CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), CStr(TxtRegNro.Text), 0, CStr(LblInmobCod.Caption), _
'                                          CStr(TxtTelefono.Text), CStr(LblTasaPersCod.Caption), _
'                                          CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), _
'                                          "", CStr(LblNotaPersCod.Caption), GarDetDJ, CDbl(TxtHipCuotaIni.Text), CDbl(TxtMontoHip.Text), CDbl(TxtPrecioVenta.Text), CDbl(TxtValorCConst.Text), CStr(TxtNroPoliza.Text), IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons.Text), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), sNumgarant, lrs, gdFecSis, txtDireccion.Text, _
'                                          txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, Trim(Right(Me.cboTipo.Text, 3)), CDbl(TxtValorEdificacion.Text), CDbl(txtVRM.Text), CDate(IIf(txtFechaCertifGravamen.Text = "__/__/____", "01/01/1950", txtFechaCertifGravamen.Text)), CDbl(txtValorGravado.Text), vPol, CInt(ChkGarPoliza.value)) 'Campos adicionales
'                                          'peac 20071122 se agrego CDbl(txtVRM.Text),dFechaCertifGravamen
        
            
            'By capi 10102008 se modifico porque varios controles fueron suprimidos
            
'            Call oGarantia.NuevaGarantia(CStr(Trim(Right(CmbDocGarant.Text, 10))), CStr(Trim(txtNumDoc.Text)), CStr(Trim(Right(CmbTipoGarant.Text, 10))), _
'                                         CboGarantia.ItemData(CboGarantia.ListIndex), CStr(Trim(Right(cmbMoneda.Text, 10))), _
'                                         CStr(Trim(txtDescGarant.Text)), CStr(Trim(Right(cmbPersUbiGeo(3).Text, 15))), _
'                                         CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), CStr(Trim(txtcomentarios.Text)), _
'                                         RelPers, CDate(gdFecSis), CStr(LblPersCodEmi.Caption), CStr(IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13))), _
'                                         CInt(IIf(OptCG(0).value, 0, 1)), _
'                                         CInt(IIf(OptTR(0).value, 0, 1)), CStr(LblSegPersCod.Caption), _
'                                         CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), CStr(TxtRegNro.Text), 0, CStr(LblInmobCod.Caption), _
'                                          CStr(TxtTelefono.Text), CStr(LblTasaPersCod.Caption), _
'                                          CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), _
'                                          "", CStr(LblNotaPersCod.Caption), GarDetDJ, CDbl(TxtHipCuotaIni.Text), CDbl(TxtMontoHip.Text), CDbl(TxtPrecioVenta.Text), CDbl(TxtValorCConst.Text), CStr(TxtNroPoliza.Text), IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons.Text), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), sNumgarant, lrs, gdFecSis, _
'                                          Trim(txtDireccion.Text), _
'                                          txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, vPol, vDatosGar, IIf(TxtFecVctoPol.Text = "__/__/____", "01/01/1950", TxtFecVctoPol.Text))           'Campos adicionales
'                                          'peac 20071122 se agrego CDbl(txtVRM.Text),dFechaCertifGravamen
'                                        'PEAC 20080710 se cambio este "Trim(txtDireccion.Text)" por "Replace(Trim(txtDireccion.Text), Chr$(10), Left(Trim(txtDireccion.Text), Len(Trim(txtDireccion.Text)) - 2))"
'
            'By Capi 03102008
            Dim lbMismaFichaRegistral As Boolean
            lbMismaFichaRegistral = False
            'MADM 20100512
            lbMismaFichaRegistral = oGarantia.ExisteGarantia(CStr(Trim(Right(CmbDocGarant.Text, 10))), IIf(txtNumDoc.Visible, CStr(Trim(txtNumDoc.Text)), Trim(Left(Me.CboNumPF.Text, 18))), CStr(LblPersCodEmi.Caption))
            'END MADM
            If lbMismaFichaRegistral Then
                MsgBox "FICHA REGISTRAL PERTENECE A OTRA GARANTIA...PROCESO CANCELADO"
                Exit Sub
            End If
            'End By
            'madm 20100512
            Call oGarantia.NuevaGarantia(CStr(Trim(Right(CmbDocGarant.Text, 10))), IIf(txtNumDoc.Visible, CStr(Trim(txtNumDoc.Text)), Trim(Left(Me.CboNumPF.Text, 18))), CStr(Trim(Right(CmbTipoGarant.Text, 10))), _
                                         CboGarantia.ItemData(CboGarantia.ListIndex), CStr(Trim(Right(cmbMoneda.Text, 10))), _
                                         CStr(Trim(txtDescGarant.Text)), CStr(Trim(Right(cmbPersUbiGeo(3).Text, 15))), _
                                         CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), CStr(Trim(txtcomentarios.Text)), _
                                         RelPers, CDate(gdFecSis), CStr(LblPersCodEmi.Caption), CStr(IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13))), _
                                         CInt(IIf(OptCG(0).value, 0, 1)), _
                                         CInt(IIf(OptTR(0).value, 0, 1)), CStr(LblSegPersCod.Caption), _
                                         CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), CStr(TxtRegNro.Text), 0, CStr(LblInmobCod.Caption), _
                                          CStr(TxtTelefono.Text), CStr(LblTasaPersCod.Caption), _
                                          CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), _
                                          "", CStr(LblNotaPersCod.Caption), GarDetDJ, , , , , CStr(TxtNroPoliza.Text), IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons.Text), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), sNumgarant, lrs, gdFecSis, _
                                          Trim(txtDireccion.Text), _
                                          txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, vPol, vDatosGar, , IIf(Me.txtFecEmision = "__/__/____", "01/01/1950", Me.txtFecEmision), Me.LblEmisorPersCod, Trim(Right(IIf(CmbDocCompra.Text = "", "-1", CmbDocCompra.Text), 4)), IIf(txtNumDocCompra.Text = "", "0", txtNumDocCompra.Text), CCur(txtValorDocCompra.Text)) 'Campos adicionales
            'end madm
'            Call oGarantia.NuevaGarantia(CStr(Trim(Right(CmbDocGarant.Text, 10))), CStr(Trim(txtNumDoc.Text)), CStr(Trim(Right(CmbTipoGarant.Text, 10))), _
'                                         CboGarantia.ItemData(CboGarantia.ListIndex), CStr(Trim(Right(cmbMoneda.Text, 10))), _
'                                         CStr(Trim(txtDescGarant.Text)), CStr(Trim(Right(cmbPersUbiGeo(3).Text, 15))), _
'                                         CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), CStr(Trim(txtcomentarios.Text)), _
'                                         RelPers, CDate(gdFecSis), CStr(LblPersCodEmi.Caption), CStr(IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13))), _
'                                         CInt(IIf(OptCG(0).value, 0, 1)), _
'                                         CInt(IIf(OptTR(0).value, 0, 1)), CStr(LblSegPersCod.Caption), _
'                                         CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), CStr(TxtRegNro.Text), 0, CStr(LblInmobCod.Caption), _
'                                          CStr(TxtTelefono.Text), CStr(LblTasaPersCod.Caption), _
'                                          CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), _
'                                          "", CStr(LblNotaPersCod.Caption), GarDetDJ, , , , , CStr(TxtNroPoliza.Text), IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons.Text), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), sNumgarant, lrs, gdFecSis, _
'                                          Trim(txtDireccion.Text), _
'                                          txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, vPol, vDatosGar)            'Campos adicionales
                                          'peac 20071122 se agrego CDbl(txtVRM.Text),dFechaCertifGravamen
                                        'PEAC 20080710 se cambio este "Trim(txtDireccion.Text)" por "Replace(Trim(txtDireccion.Text), Chr$(10), Left(Trim(txtDireccion.Text), Len(Trim(txtDireccion.Text)) - 2))"
                                        
        '''*** PEAC 20090126
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, sNumgarant, gCodigoGarantia
                    
        Set oGarantia = Nothing
    Else
        bVerificaDescobertura = False
        Set oGarantia = New COMDCredito.DCOMGarantia
'            Call oGarantia.ActualizaGarantia(sNumgarant, Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text), _
'                    Trim(Right(CmbTipoGarant.Text, 10)), CboGarantia.ItemData(CboGarantia.ListIndex), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
'                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), _
'                    Trim(txtcomentarios.Text), RelPers, gdFecSis, LblPersCodEmi.Caption, IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13)), ChkGarReal.value, _
'                    IIf(OptCG(0).value, 0, 1), IIf(OptTR(0).value, 0, 1), LblSegPersCod.Caption, CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), TxtRegNro.Text, _
'                    0, LblInmobCod.Caption, TxtTelefono.Text, LblTasaPersCod.Caption, CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), "", LblNotaPersCod.Caption, GarDetDJ, CDbl(TxtHipCuotaIni.Text), CDbl(TxtMontoHip.Text), CDbl(TxtPrecioVenta.Text), CDbl(TxtValorCConst.Text), TxtNroPoliza.Text, IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas), lrs, gdFecSis, txtDireccion.Text, _
'                    txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, bVerificaDescobertura, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, Trim(Right(Me.CboTipo.Text, 3)), CDbl(TxtValorEdificacion.Text), CDbl(txtVRM.Text), CDate(IIf(txtFechaCertifGravamen.Text = "__/__/____", "01/01/1950", txtFechaCertifGravamen.Text)), CDbl(txtValorGravado.Text), vPol) 'Se agregaron campos
'                    'peac 20071122 se agrego CDbl(txtVRM.Text)
        
            
            'By capi 10102008 porque se elimino varios controles
            
'            Call oGarantia.ActualizaGarantia(sNumgarant, Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text), _
'                    Trim(Right(CmbTipoGarant.Text, 10)), CboGarantia.ItemData(CboGarantia.ListIndex), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
'                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), _
'                    Trim(txtcomentarios.Text), RelPers, gdFecSis, LblPersCodEmi.Caption, IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13)), ChkGarReal.value, _
'                    IIf(OptCG(0).value, 0, 1), IIf(OptTR(0).value, 0, 1), LblSegPersCod.Caption, CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), TxtRegNro.Text, _
'                    0, LblInmobCod.Caption, TxtTelefono.Text, LblTasaPersCod.Caption, CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), "", LblNotaPersCod.Caption, GarDetDJ, CDbl(TxtHipCuotaIni.Text), CDbl(TxtMontoHip.Text), CDbl(TxtPrecioVenta.Text), CDbl(TxtValorCConst.Text), TxtNroPoliza.Text, IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), lrs, gdFecSis, _
'                    Trim(txtDireccion.Text), _
'                    txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, bVerificaDescobertura, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, _
'                    vPol, vDatosGar, IIf(TxtFecVctoPol.Text = "__/__/____", "01/01/1950", TxtFecVctoPol.Text))      'Se agregaron campos
'                    'PEAC 20071122 se agrego CDbl(txtVRM.Text)
'                    'PEAC 20080710 se cambio este "Trim(txtDireccion.Text)" por "Replace(Trim(txtDireccion.Text), Chr$(10), Left(Trim(txtDireccion.Text), Len(Trim(txtDireccion.Text)) - 2))"
'madm 20100512 monto x gravar no se debe actualizar - calculo

'Comentado por BRGO 20111226... Descomentar en el sgte pase a producción
            Dim vDatosGarMob() As String
            ReDim vDatosGarMob(4)
            vDatosGarMob(0) = IIf(Me.txtFecEmision = "__/__/____", "01/01/1950", Me.txtFecEmision)
            vDatosGarMob(1) = Me.LblEmisorPersCod
            vDatosGarMob(2) = Trim(Right(IIf(CmbDocCompra.Text = "", "-1", CmbDocCompra.Text), 4))
            vDatosGarMob(3) = IIf(txtNumDocCompra.Text = "", "0", txtNumDocCompra.Text)
            vDatosGarMob(4) = Format(CCur(txtValorDocCompra.Text), "0.00")
            
            Call oGarantia.ActualizaGarantia(sNumgarant, Trim(Right(CmbDocGarant.Text, 10)), IIf(txtNumDoc.Visible, CStr(Trim(txtNumDoc.Text)), Left(Me.CboNumPF.Text, 18)), _
                    Trim(Right(CmbTipoGarant.Text, 10)), CboGarantia.ItemData(CboGarantia.ListIndex), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text + nxgravar), _
                    Trim(txtcomentarios.Text), RelPers, gdFecSis, LblPersCodEmi.Caption, IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13)), ChkGarReal.value, _
                    IIf(OptCG(0).value, 0, 1), IIf(OptTR(0).value, 0, 1), LblSegPersCod.Caption, CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), TxtRegNro.Text, _
                    0, LblInmobCod.Caption, TxtTelefono.Text, LblTasaPersCod.Caption, CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), "", LblNotaPersCod.Caption, GarDetDJ, , , , , TxtNroPoliza.Text, IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), lrs, gdFecSis, _
                    Trim(txtDireccion.Text), _
                    txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, bVerificaDescobertura, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, _
                    vPol, vDatosGar, , , vDatosGarMob)   'Se agregaron campos
                    '*** BRGO Se quitó vDatosGarMob
'end madm
'            Call oGarantia.ActualizaGarantia(sNumgarant, Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text), _
'                    Trim(Right(CmbTipoGarant.Text, 10)), CboGarantia.ItemData(CboGarantia.ListIndex), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
'                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), _
'                    Trim(txtcomentarios.Text), RelPers, gdFecSis, LblPersCodEmi.Caption, IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13)), ChkGarReal.value, _
'                    IIf(OptCG(0).value, 0, 1), IIf(OptTR(0).value, 0, 1), LblSegPersCod.Caption, CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), TxtRegNro.Text, _
'                    0, LblInmobCod.Caption, TxtTelefono.Text, LblTasaPersCod.Caption, CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), "", LblNotaPersCod.Caption, GarDetDJ, , , , , TxtNroPoliza.Text, IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas.Text), lrs, gdFecSis, _
'                    Trim(txtDireccion.Text), _
'                    txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, , , gsCodAge, bVerificaDescobertura, FeTabla.GetRsNew(0), nTipoTablaValor, Me.TxtDirecRegPubli.Text, _
'                    vPol, vDatosGar)      'Se agregaron campos
                    'PEAC 20071122 se agrego CDbl(txtVRM.Text)
                    'PEAC 20080710 se cambio este "Trim(txtDireccion.Text)" por "Replace(Trim(txtDireccion.Text), Chr$(10), Left(Trim(txtDireccion.Text), Len(Trim(txtDireccion.Text)) - 2))"
        
            '''*** PEAC 20090126
            objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, , sNumgarant, gCodigoGarantia
        
            'WIOR 20120616*******************
            Set rsActualizacion = oGarantia.ObtenerUltimaActualizacion(sNumgarant)
            If rsActualizacion.RecordCount > 0 Then
                Call oGarantia.ActualizarMovimientoGarantia(1, sNumgarant, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
            Else
                Call oGarantia.RegistrarMovimientoGarantia(sNumgarant, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
            End If
            'WIOR FIN ***********************
        Set oGarantia = Nothing
        
        
        If bVerificaDescobertura Then
            MsgBox "El monto ingresado descobertura el credito", vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
    If SSGarant.TabVisible(3) = True Then
    '    Set oMantGarant = New COMDCredito.DCOMGarantia
        'If cmdEjecutar = 1 Then
           'cuando es una nueva garantia
           'lsNumGarant = oMantGarant.ObtenerMaxcNumGarant
           'Set oMantGarant = Nothing
           'Set lrs = oMantGarant.DJ(lsNumGarant, gdFecSis)
        'Else
            ' cuando es una actualizacion de la garantia
           'Set lrs = oMantGarant.DJ(sNumgarant, gdFecSis)
        'End If
     '   With DRDJ
     '       Set .DataSource = lrs
     '       .DataMember = ""
     '       '.Orientation = rptOrientPortrait
     '       .Inicio sNumgarant, gdFecSis
     '       .Refresh
     '       .Show vbModal
     '   End With
        
     '   Set oMantGarant = Nothing
    End If
    cmdEjecutar = -1
    Call HabilitaIngreso(False)
    Call HabilitaIngresoGarantReal(False)
    
    '***Agregado por ELRO el 20120329, según OYP-RFC022-2012
    If vTipoInicio = ConsultaGarant Then
        If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" And _
           ChkGarReal = 0 And InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0 Then
           'WIOR 20150608 QUITO gsCodCargo = "006024"  Y AGREGO InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0
                cmdActualizar.Visible = True
        Else
                cmdActualizar.Visible = False
        End If
    End If
    '***Fin Agregado por ELRO*******************************
    Exit Sub


ErrorCmdAceptar_Click:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'WIOR 20130122 ***************************
Private Sub cmdActAdmCred_Click()
If MsgBox("Estas Seguro de Guardar Los datos de la Poliza?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim vPoliza() As Variant
    Dim oGarantia As COMDCredito.DCOMGarantia
    Set oGarantia = New COMDCredito.DCOMGarantia
    Dim nRealizado As Boolean
    nRealizado = False
    If fnTpoPoliza = 1 Then
        If ChkGarPoliza.value = 1 Then
            If Trim(cmbClaseInmueble.Text) <> "" And Trim(cmbCategoria.Text) <> "" And Trim(txtNumLocales.Text) <> "" And Trim(txtNumPisos.Text) <> "" _
                And Trim(txtNumSotanos.Text) <> "" And Trim(txtAnioConstruccion.Text) <> "" Then
                ReDim vPoliza(1, 7)
                vPoliza(1, 0) = 1
                vPoliza(1, 1) = gsCodAge
                vPoliza(1, 2) = CInt(Trim(Right(cmbClaseInmueble.Text, 2)))
                vPoliza(1, 3) = CInt(Trim(Right(cmbCategoria.Text, 2)))
                vPoliza(1, 4) = CInt(txtNumLocales.Text)
                vPoliza(1, 5) = CInt(txtNumPisos.Text)
                vPoliza(1, 6) = CInt(txtNumSotanos.Text)
                vPoliza(1, 7) = CInt(txtAnioConstruccion.Text)
            Else
                MsgBox "Ingrese Todos datos de la Poliza de Inmueble.", vbInformation, "Aviso"
                If SSGarant.TabVisible(6) = True Then
                    SSGarant.Tab = 6
                End If
                Exit Sub
            End If
        Else
            ReDim vPoliza(1, 0)
            vPoliza(1, 0) = 1
        End If
        nRealizado = oGarantia.ActualizarGarAdmCred(sNumgarant, IIf(ChkGarPoliza.value = 1, True, False), vPoliza)
    ElseIf fnTpoPoliza = 2 Then
        If chkGarPolizaMob.value = 1 Then
            If Trim(Me.cboClaseMueble.Text) <> "" And Trim(txtAnioFabricacion.Text) <> "" Then
                ReDim vPoliza(1, 3)
                vPoliza(1, 0) = 2
                vPoliza(1, 1) = gsCodAge
                vPoliza(1, 2) = CInt(Trim(Right(Me.cboClaseMueble.Text, 2)))
                vPoliza(1, 3) = CInt(txtAnioFabricacion.Text)
            Else
                MsgBox "Ingrese Todos datos de la Poliza Mobiliaria.", vbInformation, "Aviso"
                If SSGarant.TabVisible(8) = True Then
                    SSGarant.Tab = 8
                End If
                Exit Sub
            End If
        Else
            ReDim vPoliza(1, 0)
            vPoliza(1, 0) = 2
        End If
        nRealizado = oGarantia.ActualizarGarAdmCred(sNumgarant, IIf(chkGarPolizaMob.value = 1, True, False), vPoliza)
    End If
     
    If nRealizado Then
        MsgBox "Los Datos de la Poliza se registraron satisfacoriamente.", vbInformation, "Aviso"
        ReDim vPoliza(1, 0)
        vPoliza(1, 0) = 0
        cmdActAdmCred.Visible = False
        ChkGarPoliza.Enabled = False
        chkGarPolizaMob.Enabled = False
        txtAnioFabricacion.Enabled = False
    End If
End If
End Sub
'WIOR FIN ********************************

'***Agregado por ELRO el 20120329, según OYP-RFC022-2012
Private Sub cmdActualizar_Click()
Dim oDCOMGarantia As COMDCredito.DCOMGarantia
Set oDCOMGarantia = New COMDCredito.DCOMGarantia
Dim ldFechaGarantia As Date

ldFechaGarantia = oDCOMGarantia.devolverFechaGarantiaPF(txtNumDoc)

If ldFechaGarantia <> CDate("01/01/1900") And ldFechaGarantia <= gdFecSis Then
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, "Valor Comercial: " & txtMontotas & " Realización " & txtMontoRea & " Disponible " & txtMontoxGrav, sNumgarant, gCodigoGarantia
    Call oDCOMGarantia.actulizarGarantiaPF(Format(gdFecSis, "yyyyMMdd"), txtNumDoc)
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Valor Comercial: " & txtMontotas & " Realización " & txtMontoRea & " Disponible " & txtMontoxGrav, sNumgarant, gCodigoGarantia
    Call cmdBuscar_Click
End If
End Sub
'***Fin Agregado por ELRO*******************************

Private Sub CmdBuscaEmisor_Click()
Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    If Not oPers Is Nothing Then
        LblPersCodEmi.Caption = oPers.sPerscod
        LblEmisor.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub
'*** BRGO 20111125 ******************************
Private Sub cmdBuscaEmisorDoc_Click()
    Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    If Not oPers Is Nothing Then
        LblEmisorPersCod.Caption = oPers.sPerscod
        LblEmisorPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    If CmbDocCompra.Enabled = True Then
        CmbDocCompra.SetFocus
    End If
End Sub
'*** END BRGO ***********************************

Private Sub CmdBuscaInmob_Click()
    Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    If Not oPers Is Nothing Then
        LblInmobCod.Caption = oPers.sPerscod
        LblInmobNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdBuscaNot_Click()
Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    If Not oPers Is Nothing Then
        LblNotaPersCod.Caption = oPers.sPerscod
        LblNotaPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdBuscaPersona_Click()
    Call cmdCancelar_Click
    ObtieneDocumPersona
    If vTipoInicio = ConsultaGarant Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
End Sub

Private Sub ObtieneDocumPersona()
Dim oGaran As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
Dim oPers As comdpersona.UCOMPersona
Dim L As ListItem
    
    LstGaratias.ListItems.Clear
    Set oPers = New comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    Set oGaran = New COMDCredito.DCOMGarantia
    
    If oPers Is Nothing Then
        Exit Sub
    End If
    
    Set R = oGaran.RecuperaGarantiasPersona(oPers.sPerscod, True)
    Set oGaran = Nothing
    If Not (R.EOF And R.BOF) Then
        'R.RecordCount > 0
        Me.Caption = "Garantias de Cliente : " & oPers.sPersNombre
    End If
    LstGaratias.ListItems.Clear
    Set oPers = Nothing
    Do While Not R.EOF
        Set L = LstGaratias.ListItems.Add(, , IIf(IsNull(R!cDescripcion), "", R!cDescripcion))
        L.Bold = True
        If R!nmoneda = gMonedaExtranjera Then
            L.ForeColor = RGB(0, 125, 0)
        Else
            L.ForeColor = vbBlack
        End If
        L.SubItems(1) = Trim(R!cNumGarant)
        L.SubItems(2) = Trim(R!cPersCodEmisor)
        L.SubItems(3) = PstaNombre(R!cPersNombre)
        L.SubItems(4) = Trim(R!cTpoDoc)
        L.SubItems(5) = Trim(R!cNroDoc)
        
        R.MoveNext
    Loop
End Sub

Private Sub cmdBuscar_Click()
'WIOR 20120616***********************************************
Dim rsGarantia As ADODB.Recordset
Dim sUltActualizacion As String
Dim oGaran As COMDCredito.DCOMGarantia
Dim oAgencia  As COMDConstantes.DCOMAgencias
Dim oFunciones As New COMFunciones.FCOMImpresion
'WIOR FIN ***************************************************
    If SSGarant.TabVisible(3) = True Then
        SSGarant.TabVisible(3) = False
    End If
        
    bAsignadoACredito = False
        
    If Me.LstGaratias.ListItems.Count = 0 Then
        MsgBox "No Existe Garantia que Mostrar ", vbInformation, "Aviso"
        Exit Sub
    End If
 
    CmbDocGarant.Enabled = False
        
    Me.LblPersCodEmi.Caption = Me.LstGaratias.SelectedItem.SubItems(2)
    Me.LblEmisor.Caption = Me.LstGaratias.SelectedItem.SubItems(3)
    'Me.CmbDocGarant.ListIndex = IndiceListaCombo(CmbDocGarant, Me.LstGaratias.SelectedItem.SubItems(4))
    'Me.txtNumDoc.Text = Me.LstGaratias.SelectedItem.SubItems(5)
    chkEnTramite = 0 'EJVG20130201
    Call CargaDatos(Trim(LstGaratias.SelectedItem.SubItems(1)))
    sNumgarant = Trim(Me.LstGaratias.SelectedItem.SubItems(1))
    
    'By Capi 25112008
        If (txtFechaCertifGravamen.Text = "__/__/____" Or txtFechaCertifGravamen.Text = "01/01/1950") Then
            chkEnTramite = 1
        Else
            chkEnTramite = 0
        End If
     'End By
     'ALPA 20120320********************************************************************
     If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
        If (txtFechaCertifGravamenGA.Text = "__/__/____" Or txtFechaCertifGravamen.Text = "01/01/1950") Then
            chkEnTramiteGA = 1
        Else
            chkEnTramiteGA = 0
        End If
     End If
     '*********************************************************************************
        
      
    
    If vTipoInicio = ConsultaGarant Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    'ALPA 20120322********************
    Frame9.Enabled = False
    Frame8.Enabled = False
    '*********************************
    '***Agregado por ELRO el 20120329, según OYP-RFC022-2012
    If vTipoInicio = ConsultaGarant Then
        If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" And _
           ChkGarReal = 0 And InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0 Then
           'WIOR 20150608 QUITO gsCodCargo = "006024"  Y AGREGO InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0
                cmdActualizar.Visible = True
        Else
                cmdActualizar.Visible = False
        End If
    End If
 '***Fin Agregado por ELRO*******************************
  'WIOR 20120616, según OYP-RFC049-2012******************************
 If vTipoInicio = MantenimientoGarantia Or vTipoInicio = ConsultaGarant Then
    Set oGaran = New COMDCredito.DCOMGarantia
    Set oAgencia = New COMDConstantes.DCOMAgencias
    Set rsGarantia = oGaran.ObtenerUltimaActualizacion(sNumgarant)
    
    If rsGarantia.RecordCount > 0 Then
        If IIf(IsNull(rsGarantia!cMovVerificacion), "", Trim(rsGarantia!cMovVerificacion)) = "" And gsCodArea = "040" Then '040 SUPERVISION DE CREDITOS
            sUltActualizacion = rsGarantia!cUltimaActualizacion
            Set rsGarantia = oAgencia.RecuperaAgencias(Trim(Mid(sUltActualizacion, 18, 2)))
            
            MsgBox oFunciones.Centra("Esta garantía a sido modificada por el usuario", 50) & _
            Chr(10) & oFunciones.Centra(Right(sUltActualizacion, 4) & " el día " & _
            Mid(sUltActualizacion, 7, 2) & "/" & Mid(sUltActualizacion, 5, 2) & "/" & Mid(sUltActualizacion, 1, 4) & _
            " a las " & Mid(sUltActualizacion, 9, 2) & ":" & Mid(sUltActualizacion, 11, 2) & ":" & Mid(sUltActualizacion, 13, 2), 50) & _
            IIf(rsGarantia.RecordCount > 0, Chr(10) & oFunciones.Centra(" en la " & Trim(rsGarantia!cAgeDescripcion), 50), ""), vbInformation, "Mensaje"
            Call oGaran.ActualizarMovimientoGarantia(2, sNumgarant, , GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
    End If
    Set oGaran = Nothing
    Set oAgencia = Nothing
    Set rsGarantia = Nothing
 End If
 'WIOR FIN *******************************************************
End Sub

Private Sub CmdBuscaSeg_Click()
Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    If Not oPers Is Nothing Then
        LblSegPersCod.Caption = oPers.sPerscod
        LblSegPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdBuscaTasa_Click()
Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.inicio
    If Not oPers Is Nothing Then
        LblTasaPersCod.Caption = oPers.sPerscod
        LblTasaPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    If TxtFecTas.Enabled = True Then
        TxtFecTas.SetFocus
    End If
End Sub

Private Sub cmdBusGar_Click()
'WIOR 20120616***********************************************
Dim rsGarantia As ADODB.Recordset
Dim sUltActualizacion As String
Dim oAgencia  As COMDConstantes.DCOMAgencias
Dim oFunciones As New COMFunciones.FCOMImpresion
'WIOR FIN ***************************************************
Dim oGaran As COMDCredito.DCOMGarantia
Dim L As ListItem
Dim R As ADODB.Recordset

    If Val(Me.txtNumGar) <= 0 Then
        Exit Sub
    End If

    sNumgarant = Format(Me.txtNumGar, "00000000")

    
    If Not CargaDatos(sNumgarant) Then
        MsgBox "El número de garantía que acaba de ingresar no existe.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    
    
    '------------------------------------
    Set oGaran = New COMDCredito.DCOMGarantia
    
    Set R = oGaran.RecuperaDatosGarantiasPersona(sNumgarant)
    Set oGaran = Nothing
    
    If R.RecordCount > 0 Then
        'R.RecordCount > 0
        Me.Caption = "Garantias de Cliente : " & R!cPersNombre
    End If
    LstGaratias.ListItems.Clear
    
    Do While Not R.EOF
        Set L = LstGaratias.ListItems.Add(, , IIf(IsNull(R!cDescripcion), "", R!cDescripcion))
        L.Bold = True
        If R!nmoneda = gMonedaExtranjera Then
            L.ForeColor = RGB(0, 125, 0)
        Else
            L.ForeColor = vbBlack
        End If
        L.SubItems(1) = Trim(R!cNumGarant)
        L.SubItems(2) = Trim(R!cPersCodEmisor)
        L.SubItems(3) = PstaNombre(R!cPersNombre)
        L.SubItems(4) = Trim(R!cTpoDoc)
        L.SubItems(5) = Trim(R!cNroDoc)
        
        R.MoveNext
    Loop
        
    R.Close
        
    '------------------------------------
    
        If (txtFechaCertifGravamen.Text = "__/__/____" Or txtFechaCertifGravamen.Text = "01/01/1950") Then
            chkEnTramite = 1
        Else
            chkEnTramite = 0
        End If
        'ALPA 20120320**************************************************
        If (txtFechaCertifGravamenGA.Text = "__/__/____" Or txtFechaCertifGravamenGA.Text = "01/01/1950") Then
            chkEnTramiteGA = 1
        Else
            chkEnTramiteGA = 0
        End If
        '****************************************************************
    
    If vTipoInicio = ConsultaGarant Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
 '***Agregado por ELRO el 20120329, según OYP-RFC022-2012
 If vTipoInicio = ConsultaGarant Then
        If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" And _
           ChkGarReal = 0 And InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0 Then
           'WIOR 20150608 QUITO gsCodCargo = "006024"  Y AGREGO InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0
                cmdActualizar.Visible = True
        Else
                cmdActualizar.Visible = False
        End If
 End If
 '***Fin Agregado por ELRO*******************************
'WIOR 20120616, según OYP-RFC049-2012******************************
 If vTipoInicio = MantenimientoGarantia Or vTipoInicio = ConsultaGarant Then
    Set oGaran = New COMDCredito.DCOMGarantia
    Set oAgencia = New COMDConstantes.DCOMAgencias
    Set rsGarantia = oGaran.ObtenerUltimaActualizacion(sNumgarant)
    
    If rsGarantia.RecordCount > 0 Then
        If IIf(IsNull(rsGarantia!cMovVerificacion), "", Trim(rsGarantia!cMovVerificacion)) = "" And gsCodArea = "040" Then '040 SUPERVISION DE CREDITOS
            sUltActualizacion = rsGarantia!cUltimaActualizacion
            Set rsGarantia = oAgencia.RecuperaAgencias(Trim(Mid(sUltActualizacion, 18, 2)))
            
            MsgBox oFunciones.Centra("Esta garantía a sido modificada por el usuario", 50) & _
            Chr(10) & oFunciones.Centra(Right(sUltActualizacion, 4) & " el día " & _
            Mid(sUltActualizacion, 7, 2) & "/" & Mid(sUltActualizacion, 5, 2) & "/" & Mid(sUltActualizacion, 1, 4) & _
            " a las " & Mid(sUltActualizacion, 9, 2) & ":" & Mid(sUltActualizacion, 11, 2) & ":" & Mid(sUltActualizacion, 13, 2), 50) & _
            IIf(rsGarantia.RecordCount > 0, Chr(10) & oFunciones.Centra(" en la " & Trim(rsGarantia!cAgeDescripcion), 50), ""), vbInformation, "Mensaje"
            Call oGaran.ActualizarMovimientoGarantia(2, sNumgarant, , GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
    End If
    Set oGaran = Nothing
    Set oAgencia = Nothing
    Set rsGarantia = Nothing
 End If
 'WIOR FIN *******************************************************
End Sub

Private Sub cmdCancelar_Click()
    If cmdEjecutar = 2 Then
        CargaDatos Trim(sNumgarant)
    Else
        If cmdEjecutar = 1 Then
            Call LimpiaPantalla
        End If
    End If
    Call HabilitaIngreso(False)
    If Me.ChkGarReal.value = 1 Then
        Call HabilitaIngresoGarantReal(False)
    End If
    
    If Me.ChkGarReal.Enabled Then
        Me.ChkTasacion.Enabled = True
    Else
        Me.ChkTasacion.Enabled = False
    End If
    
    
    
    Call LimpiaPantalla
    'CmbDocGarant.Enabled = True
    'txtNumDoc.Enabled = True
    'CmbDocGarant.SetFocus
    cmdEjecutar = -1
    '***Agregado por ELRO el 20120329, según OYP-RFC022-2012
    If vTipoInicio = ConsultaGarant Then
        If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" And _
           ChkGarReal = 0 And InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0 Then
           'WIOR 20150608 QUITO gsCodCargo = "006024"  Y AGREGO InStr(1, fsGrupoActGarDPF, gsCodCargo) > 0
                cmdActualizar.Visible = True
        Else
                cmdActualizar.Visible = False
        End If
    End If
    '***Fin Agregado por ELRO*******************************
End Sub

Private Sub CmdCliAceptar_Click()
Dim i As Long
Dim oGarantia As COMNCredito.NCOMGarantia
Dim RelPers() As String

    For i = 1 To FERelPers.Rows - 2
        If Trim(FERelPers.TextMatrix(i, 1)) = Trim(FERelPers.TextMatrix(FERelPers.Rows - 1, 1)) Then
            MsgBox "Persona Ya Tiene Relacion de la Garantia", vbInformation, "Aviso"
            FERelPers.row = FERelPers.Rows - 1
            FERelPers.Col = 1
            FERelPers.SetFocus
            Exit Sub
        End If
    Next i
    For i = 1 To FERelPers.Rows - 1
        If Len(Trim(FERelPers.TextMatrix(i, 1))) < 13 Then
            MsgBox "Codigo de Persona Incorrecto", vbInformation, "Aviso"
            FERelPers.row = i
            FERelPers.Col = 1
            FERelPers.SetFocus
            Exit Sub
        End If
        If Len(Trim(FERelPers.TextMatrix(i, 3))) = 0 Then
            MsgBox "Relacion de Persona Con la Garantias es Incorrecto", vbInformation, "Aviso"
            FERelPers.row = i
            FERelPers.Col = 3
            FERelPers.SetFocus
            Exit Sub
        End If
    Next i
    ReDim RelPers(FERelPers.Rows - 1)
    For i = 1 To FERelPers.Rows - 1
        RelPers(i - 1) = FERelPers.TextMatrix(i, 3)
    Next i
    Set oGarantia = New COMNCredito.NCOMGarantia
    
    If oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)) <> "" Then
            MsgBox oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)), vbInformation, "Aviso"
            Exit Sub
    End If
     Set oGarantia = Nothing
    
    FERelPers.lbEditarFlex = False
    CmdCliNuevo.Visible = True
    CmdCliEliminar.Visible = True
    CmdCliAceptar.Visible = False
    CmdCliCancelar.Visible = False
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End Sub

Private Sub CmdCliCancelar_Click()
    Call FERelPers.EliminaFila(FERelPers.row)
    FERelPers.lbEditarFlex = False
    CmdCliNuevo.Visible = True
    CmdCliEliminar.Visible = True
    CmdCliAceptar.Visible = False
    CmdCliCancelar.Visible = False
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End Sub

Private Sub CmdCliEliminar_Click()
    If FERelPers.row < 1 Then
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar a la Persona " & FERelPers.TextMatrix(FERelPers.row, 2) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If FERelPers.row = 1 And FERelPers.Rows = 2 Then
            FERelPers.TextMatrix(1, 0) = ""
            FERelPers.TextMatrix(1, 1) = ""
            FERelPers.TextMatrix(1, 2) = ""
            FERelPers.TextMatrix(1, 3) = ""
        Else
            Call FERelPers.EliminaFila(FERelPers.row)
        End If
    End If
End Sub

Private Sub CmdCliNuevo_Click()
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    CmdCliNuevo.Visible = False
    CmdCliEliminar.Visible = False
    CmdCliAceptar.Visible = True
    CmdCliCancelar.Visible = True
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    FERelPers.SetFocus
End Sub

Private Sub CmdDJAceptar_Click()
Dim i As Long
Dim oGarantia As COMNCredito.NCOMGarantia
Dim RelPers() As String

'    For I = 1 To FEDeclaracionJur.Rows - 2
'        If Trim(FEDeclaracionJur.TextMatrix(I, 1)) = Trim(FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Rows - 1, 1)) Then
'            MsgBox "Persona Ya Tiene Relacion de la Garantia", vbInformation, "Aviso"
'            FEDeclaracionJur.Row = FEDeclaracionJur.Rows - 1
'            FEDeclaracionJur.Col = 1
'            FEDeclaracionJur.SetFocus
'            Exit Sub
'        End If
'    Next I
    
    LblTotDJ.Caption = "0.00"
    
    For i = 1 To FEDeclaracionJur.Rows - 1
        If Len(Trim(FEDeclaracionJur.TextMatrix(i, 1))) = 0 Then
            MsgBox "Falta Ingresar la Descripción del Item", vbInformation, "Aviso"
            FEDeclaracionJur.row = i
            FEDeclaracionJur.Col = 1
            FEDeclaracionJur.SetFocus
            Exit Sub
        End If
        If FEDeclaracionJur.TextMatrix(i, 2) = 0 Then
            MsgBox "Falta Ingresar la Cantidad del item", vbInformation, "Aviso"
            FEDeclaracionJur.row = i
            FEDeclaracionJur.Col = 2
            FEDeclaracionJur.SetFocus
            Exit Sub
        End If
        If FEDeclaracionJur.TextMatrix(i, 3) = 0 Then
            MsgBox "Falta Ingresar el Valor Actual del item", vbInformation, "Aviso"
            FEDeclaracionJur.row = i
            FEDeclaracionJur.Col = 3
            FEDeclaracionJur.SetFocus
            Exit Sub
        End If
        
        LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) + (FEDeclaracionJur.TextMatrix(i, 2) * CCur(FEDeclaracionJur.TextMatrix(i, 3))), "#0.00"))
    
    Next i
'    ReDim RelPers(FEDeclaracionJur.Rows - 1)
'    For I = 1 To FEDeclaracionJur.Rows - 1
'        RelPers(I - 1) = FEDeclaracionJur.TextMatrix(I, 3)
'    Next I
'    Set oGarantia = New NGarantia
'    If oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)) <> "" Then
'        MsgBox oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)), vbInformation, "Aviso"
'        Exit Sub
'    End If
'    Set oGarantia = Nothing
    
    FEDeclaracionJur.lbEditarFlex = False
    CmdDJNuevo.Visible = True
    CmdImprimirDJ.Visible = True
    CmdDJEliminar.Visible = True
    CmdDJAceptar.Visible = False
    CmdDJCancelar.Visible = False
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
    
    
End Sub

Private Sub CmdDJCancelar_Click()
    Call FEDeclaracionJur.EliminaFila(FEDeclaracionJur.row)
    FEDeclaracionJur.lbEditarFlex = False
    CmdDJNuevo.Visible = True
    CmdDJEliminar.Visible = True
    CmdImprimirDJ.Visible = True
    CmdDJAceptar.Visible = False
    CmdDJCancelar.Visible = False
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End Sub

Private Sub CmdDJEliminar_Click()
If FEDeclaracionJur.row < 1 Then
    Exit Sub
End If
If MsgBox("Se va a Eliminar al Item " & FEDeclaracionJur.TextMatrix(FEDeclaracionJur.row, 1) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    If FEDeclaracionJur.row = 1 And FEDeclaracionJur.Rows = 2 Then
       LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) - (FEDeclaracionJur.TextMatrix(FEDeclaracionJur.row, 2) * FEDeclaracionJur.TextMatrix(FEDeclaracionJur.row, 3)), "#0.00"))
       
       FEDeclaracionJur.TextMatrix(1, 0) = ""
       FEDeclaracionJur.TextMatrix(1, 1) = ""
       FEDeclaracionJur.TextMatrix(1, 2) = 0
       FEDeclaracionJur.TextMatrix(1, 3) = 0
       FEDeclaracionJur.TextMatrix(1, 4) = ""
       FEDeclaracionJur.TextMatrix(1, 5) = ""
    Else
       LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) - (FEDeclaracionJur.TextMatrix(FEDeclaracionJur.row, 2) * FEDeclaracionJur.TextMatrix(FEDeclaracionJur.row, 3)), "#0.00"))

       Call FEDeclaracionJur.EliminaFila(FEDeclaracionJur.row)
    End If
End If
End Sub

Private Sub CmdDJNuevo_Click()
    FEDeclaracionJur.lbEditarFlex = True
    FEDeclaracionJur.AdicionaFila
    CmdDJNuevo.Visible = False
    CmdDJEliminar.Visible = False
    CmdDJAceptar.Visible = True
    CmdDJCancelar.Visible = True
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    FEDeclaracionJur.SetFocus
End Sub

Private Sub CmdEditar_Click()
'Dim oGarantia As New COMDCredito.DCOMGarantia

    'If oGarantia.GarantiaEnUso(sNumgarant) Then
    '    MsgBox "Solo Puede Editar el Comentario, La Garantia ya esta en uso por un Credito", vbInformation, "Aviso"
    '    Set oGarantia = Nothing
    '    Exit Sub
    'End If
    'Set oGarantia = Nothing
    
    If LstGaratias.ListItems.Count = 0 Then
        MsgBox "Seleccione una Garantia", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    If bAsignadoACredito Then
        fraPrinc.Enabled = False
        FraRelaGar.Enabled = False
        fraZonaCbo.Enabled = True
        framontos.Enabled = False
        Frame2.Enabled = False
        'Para llevar el Historico de los montos (ARCV: 30-04-2007)
        framontos.Enabled = True
        txtMontoxGrav.Enabled = True
        txtMontotas.Enabled = True
        txtMontoRea.Enabled = True
        '***************************
        
        cmbPersUbiGeo(0).Enabled = False
        cmbPersUbiGeo(1).Enabled = False
        cmbPersUbiGeo(2).Enabled = False
        cmbPersUbiGeo(3).Enabled = False
        txtDireccion.Enabled = False 'campo adicional
        txtcomentarios.Enabled = True
        
        '*** PEAC 20090814
        If CboGarantia.ListIndex = 0 Then
            If ChkGarReal.value = 1 Then
                ChkGarReal.Enabled = False
'                Me.ChkTasacion.Enabled = False
            Else
                ChkGarReal.Enabled = True
'                Me.ChkTasacion.Enabled = True
            End If
            If ChkGarPoliza.value = vbChecked Then
                ChkGarPoliza.Enabled = True
                chkGarPolizaMob.Enabled = False
            Else
                ChkGarPoliza.Enabled = False
                chkGarPolizaMob.Enabled = True
            End If
        Else
            ChkGarReal.Enabled = True
        End If
        '*** FIN PEAC
        
        
        FraDatInm.Enabled = False
        
        FraGar.Enabled = False
        
        'txtFechaCertifGravamen.Enabled = True 'peac 20071123
        'txtValorGravado.Enabled = True 'peac 20071123
                
        cmdNuevo.Enabled = False
        cmdNuevo.Visible = False
        cmdAceptar.Enabled = True
        cmdAceptar.Visible = True
        cmdEditar.Enabled = False
        cmdEditar.Visible = False
        cmdCancelar.Enabled = True
        cmdCancelar.Visible = True
        cmdEliminar.Enabled = False
        cmdEliminar.Visible = False
        cmdSalir.Enabled = False
        cmdLimpiar.Enabled = False
        cmdBuscar.Enabled = False
        FraBuscaPers.Enabled = False
        
        '***Agregado por ELRO el 20120329, según OYP-RFC022-2012
        If vTipoInicio = MantenimientoGarantia Or vTipoInicio = ConsultaGarant Then
              cmdActualizar.Visible = False
        End If
        '***Fin Agregado por ELRO*******************************

        'peac 20071127 para modificar todos los campos temporalmente
        If gcPermiteModificar Then
        
            'pagina: Relac. e la Garantia
            
            fraPrinc.Enabled = True
            FraRelaGar.Enabled = True
            
            CmdBuscaEmisor.Enabled = True
            CmbTipoGarant.Enabled = True
            txtNumDoc.Enabled = True
            txtDescGarant.Enabled = True
            CboGarantia.Enabled = True
            CmbDocGarant.Enabled = True
            CmdCliEliminar.Enabled = True
            CmdCliNuevo.Enabled = True
            cmbMoneda.Enabled = True
            
            'pagina: Datos de la Garantía
            
            cmbPersUbiGeo(0).Enabled = True
            cmbPersUbiGeo(1).Enabled = True
            cmbPersUbiGeo(2).Enabled = True
            cmbPersUbiGeo(3).Enabled = True
            
            txtDireccion.Enabled = True
            Frame2.Enabled = True
            
            txtcomentarios.Enabled = True

            If CboGarantia.ListIndex = 0 Then
                'Comentado por MAVM 11112009
                'If ChkGarReal.value = 1 Then
                    'ChkGarReal.Enabled = False
                    'Me.ChkTasacion.Enabled = False
                'Else
                    ChkGarReal.Enabled = True
                    Me.ChkTasacion.Enabled = True
                    
                'End If
                'ChkGarPoliza.Enabled = True
                chkDocCompra.Enabled = True 'BRGO 20111223
            Else
                ChkGarReal.Enabled = True
            End If
            
            OptCG(0).Enabled = True
            OptCG(1).Enabled = True
            
            OptTR(0).Enabled = True
            OptTR(1).Enabled = True
            
             'MADM 20110603
            If CboGarantia.ListIndex = 1 Then
                If txtMontotas.Visible Then
                    txtMontotas.Enabled = False
                End If
                If txtMontoRea.Visible Then
                    txtMontoRea.Enabled = False
                End If
                If txtMontoxGrav.Visible Then
                    txtMontoxGrav.Enabled = False
                End If
            Else
                txtMontotas.Enabled = True
                txtMontoRea.Enabled = True
                txtMontoxGrav.Enabled = True
            End If
            'END MAD
            
'            txtMontotas.Enabled = True
'            txtMontoRea.Enabled = True
'            txtMontoxGrav.Enabled = True
            
            FraClase.Enabled = True
            FraTipoRea.Enabled = True
            
            'pagina: Garantia Real
            
            If FraDatInm.Visible Then
                FraDatInm.Enabled = True
                CmdBuscaInmob.Enabled = True
                CboTipoInmueb.Enabled = True
                TxtTelefono.Enabled = True
            End If
            
            If fraDatVehic.Visible Then
                fraDatVehic.Enabled = True
                'By Capi 10102008
                'txtFecTasVeh.Enabled = True
                'cboEstadoTasVeh.Enabled = True
                '
                txtPlacaVehic.Enabled = True
            End If
            
            FraGar.Enabled = True
            CmdBuscaNot.Enabled = True
            cboTipo.Enabled = True
            cboTipoGA.Enabled = True 'ALPA 20120320
            TxtRegNro.Enabled = True
            TxtFechareg.Enabled = True
            TxtDirecRegPubli.Enabled = True
            
            If SSGarant.TabVisible(3) = True Then
                CmdDJNuevo.Enabled = True
                CmdDJEliminar.Enabled = True
            End If

            'pagina: Tasacion
            
            CmdBuscaTasa.Enabled = True
                        
        End If
        '*** BRGO 20111226 ***************************
        If chkDocCompra.value = vbChecked Then
            Frame6.Enabled = True
            CmbDocCompra.Enabled = True
            txtNumDocCompra.Enabled = True
            txtFecEmision.Enabled = True
            txtValorDocCompra.Enabled = True
            LblEmisorPersCod.Enabled = True
            LblEmisorPersNombre.Enabled = True
        End If
        If chkGarPolizaMob.value = vbChecked Then
            Frame7.Enabled = True
            cboClaseMueble.Enabled = True
            txtAnioFabricacion.Enabled = True
        End If
        '*** END BRGO ********************************
        
    Else
        Call HabilitaIngreso(True)
        If ChkGarReal.value = 1 Then
            Call HabilitaIngresoGarantReal(True)
        End If
        
        If Me.ChkGarReal.Enabled Then
            Me.ChkTasacion.Enabled = True
        Else
            Me.ChkTasacion.Enabled = False
        End If
        
        
        'Activa Controles segun tipo de garantia
        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Or CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaDepositosGarantia Then
            CboBanco.Enabled = True
        Else
            CboBanco.Enabled = False
        End If
        
        If ChkGarReal.value = 1 Then
        '    If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
        '        FraDatInm.Enabled = True
        '    Else
        '        FraDatInm.Enabled = False
        '    End If
            If FraDatInm.Visible Then
                FraDatInm.Enabled = True
            Else
                fraDatVehic.Enabled = True
            End If
        End If
        
        'CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------
'        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
'            SSGarant.TabVisible(3) = True
'            'FEDeclaracionJur.Enabled = True
'        Else
'            'FEDeclaracionJur.Enabled = False
'            SSGarant.TabVisible(3) = False
'        End If
        '--------------------------------------------------------------------------------------------------------
        
        'CmbDocGarant.Enabled = False
        'txtNumDoc.Enabled = False
        '*** PEAC 20080412
        If Trim(Right(CmbDocGarant, 10)) = "93" Then
        'If Trim(Right(CmbDocGarant, 10)) = "15" Then
            SSGarant.TabVisible(3) = True
        Else
            SSGarant.TabVisible(3) = False
        End If
        
        If SSGarant.TabVisible(3) = True Then
            CmdDJNuevo.Enabled = True
            CmdDJEliminar.Enabled = True
        End If
        
        If CmbTipoGarant.Enabled Then CmbTipoGarant.SetFocus
        cmdEjecutar = 2
    End If
End Sub

Private Sub cmdEliminar_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
    If MsgBox("Se va a Eliminar la Garantia, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oGarantia = New COMDCredito.DCOMGarantia
        'If oGarantia.GarantiaEnUso(sNumgarant) Then
        '    MsgBox "No se Puede Eliminar la Garantia porque ya esta en uso por un Credito", vbInformation, "Aviso"
        '    Set oGarantia = Nothing
        '    Exit Sub
        'End If
        Call oGarantia.EliminarGraantia(sNumgarant)
        Set oGarantia = Nothing
        Call LimpiaPantalla
        Call cmdLimpiar_Click
    End If
    cmdCancelar_Click
    cmdBuscar.Enabled = True
End Sub

Private Sub cmdImprimir_Click()
    
    Dim fs As Scripting.FileSystemObject
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim nLineaInicio As Integer
    Dim nLineas As Integer
    Dim nLineasTemp As Integer
    
    Dim i As Integer
    Dim nTotal As Double
    
    Dim glsArchivo As String
    
    
    glsArchivo = "TablaValores_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add


    xlAplicacion.Range("A1:A1").ColumnWidth = 10
    xlAplicacion.Range("B1:B1").ColumnWidth = 40
    xlAplicacion.Range("C1:C1").ColumnWidth = 10
                
                
    nLineas = 1
    xlHoja1.Cells(nLineas, 1) = "RESUMEN DE TABLA DE VALUACION DE GARANTIAS"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "RESUMEN"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    nLineas = nLineas + 2
    
    xlHoja1.Cells(nLineas, 2) = "VARIABLES"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Font.Bold = True
    xlHoja1.Cells(nLineas, 3) = "CALIFICACION"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineas, 3)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineas, 3)).Font.Bold = True
    
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    
    With FeTabla
        For i = 1 To .Rows - 1
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, 2) = .TextMatrix(i, 1)
            xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlLeft
            xlHoja1.Cells(nLineas, 3) = Format(.TextMatrix(i, 3), "#0.00")
            xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineas, 3)).HorizontalAlignment = xlRight
            'xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
        Next
    End With
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = "Total"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlLeft
    xlHoja1.Cells(nLineas, 3) = Format(lblCoberturaCredito.Caption, "#0.00")
    xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineas, 3)).HorizontalAlignment = xlRight
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    nLineas = nLineas + 2
    
    xlHoja1.Cells(nLineas, 2) = "COBERTURA DEL CREDITO"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlLeft
    xlHoja1.Cells(nLineas, 3) = Format(xlHoja1.Cells(nLineas - 2, 3) * 1000, "#0.00")
    xlHoja1.Range(xlHoja1.Cells(nLineas, 3), xlHoja1.Cells(nLineas, 3)).HorizontalAlignment = xlRight
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    nLineas = nLineas + 4
    
    xlHoja1.Cells(nLineas, 2) = "VºBº Analista de Creditos"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas - 1, 2)).Borders(xlEdgeBottom).LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlCenter
    nLineas = nLineas + 3
    xlHoja1.Cells(nLineas, 2) = "VºBº Administrador"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas - 1, 2)).Borders(xlEdgeBottom).LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlCenter
    nLineas = nLineas + 3
    xlHoja1.Cells(nLineas, 2) = "CUSCO " & Day(gdFecSis) & " de " & Format(gdFecSis, "MMMM") & " del " & Year(gdFecSis)
    
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(3, 1)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(nLineas, 6)).Font.Size = 9
    xlHoja1.Cells.EntireColumn.AutoFit
    xlHoja1.Cells.EntireRow.AutoFit
    
    xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsArchivo
               
    MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsArchivo, vbInformation, "Mensaje"
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
        
    Set xlAplicacion = Nothing

End Sub

Private Sub CmdImprimirDJ_Click()
    Call ImprimirDeclaraciónJurada
End Sub

Private Sub cmdLimpiar_Click()
    
    Call LimpiaPantalla
    HabilitaIngreso False
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    
    Me.ChkTasacion.Enabled = False
    
    'CmbDocGarant.Enabled = True
    'txtNumDoc.Enabled = True
    'cmdBuscar.Enabled = True
    'CmbDocGarant.SetFocus
    
    If vTipoInicio = ConsultaGarant Then
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    
    'MADM 20100513
    If Me.CboNumPF.Visible = True Then
        Me.CboNumPF.Visible = False
        Me.txtNumDoc.Visible = True
    End If
    'Agregado por LMMD CF
    bCreditoCF = False
    bValdiCCF = False
    sNumgarant = "" ' MADM 20110426
End Sub

Private Sub cmdNuevo_Click()
    Call HabilitaIngreso(True)
    Call LimpiaPantalla
    Call InicializaCombos(Me)
    cmdEjecutar = 1
    CmdBuscaEmisor.Enabled = True
    Call CmdBuscaEmisor_Click
    'CmbTipoGarant.SetFocus
    If CboGarantia.Enabled Then
        CboGarantia.SetFocus
    End If
    SSGarant.TabVisible(3) = False 'la ficha de la declaracion jurada
    'If Trim(CboGarantia.Text) = "GARANTIAS PREFERIDAS AUTOLIQUIDABLES" Then
        txtFechaBloqueo.Text = Format(gdFecSis, "DD/MM/YYYY")
    'End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'RECO20150421 ERS001-2015**********************************************
Private Sub cmdVerCred_Click()
    If LstGaratias.ListItems.Count > 0 Then Call frmGarantCred.inicio(LstGaratias.SelectedItem.SubItems(1))
End Sub
'RECO FIN *************************************************************
Private Sub FEDeclaracionJur_KeyPress(KeyAscii As Integer)
    Dim c As String
If KeyAscii = 13 Then
    If FEDeclaracionJur.Col = 5 Then
        If FEDeclaracionJur.TextMatrix(FEDeclaracionJur.row, 5) <> "" Then
            CmdDJAceptar.SetFocus
        End If
    End If
End If


If FEDeclaracionJur.Col = 1 Then
    c = Chr(KeyAscii)
    c = UCase(c)
    KeyAscii = Asc(c)
End If
End Sub


Private Sub FEDeclaracionJur_OnCellChange(pnRow As Long, pnCol As Long)
    Dim c As String
    
    If FEDeclaracionJur.Col = 1 Then
        c = FEDeclaracionJur.TextMatrix(pnRow, pnCol)
        c = UCase(c)
        FEDeclaracionJur.TextMatrix(pnRow, pnCol) = c
    End If
End Sub

Private Sub FEDeclaracionJur_RowColChange()
Dim oConstante As COMDConstantes.DCOMConstantes

If FEDeclaracionJur.Col = 4 Then
    'Carga los tipos de documentos del item
    Set oConstante = New COMDConstantes.DCOMConstantes
    FEDeclaracionJur.CargaCombo oConstante.RecuperaConstantes(gColocPigTipoDocumento)
    Set oConstante = Nothing
End If

End Sub

Private Sub FERelPers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FERelPers.Col = 3 Then
        If FERelPers.TextMatrix(FERelPers.row, 3) <> "" Then
            CmdCliAceptar.SetFocus
        End If
    End If
End If

End Sub

Private Sub FERelPers_RowColChange()
Dim oConstante As COMDConstantes.DCOMConstantes

If FERelPers.Col = 3 Then
'Carga Relacion de Personas con Garantia
    Set oConstante = New COMDConstantes.DCOMConstantes
    FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelGarantia)
    Set oConstante = Nothing
End If

End Sub

Private Sub FeTabla_OnChangeCombo()
Dim nTipoEval As Integer
Dim nTipoCredito As Integer
Dim oGar As COMDCredito.DCOMGarantia

With FeTabla
        If .TextMatrix(.row, 2) = "" Then Exit Sub
        'ARCV 14-08-2006
        If Left(.TextMatrix(.row, 2), 7) = "NINGUNO" Then
            .TextMatrix(.row, 3) = 0
            Call ActualizaMontoCobertura
            Exit Sub
        End If
        '----------------
        Set oGar = New COMDCredito.DCOMGarantia
        .TextMatrix(.row, 3) = oGar.RecuperaPuntajeTablaValores(CInt(Trim(Right(cmbTipoCreditoTabla, 10))), CInt(.TextMatrix(.row, 0)), CInt(Trim(Right(.TextMatrix(.row, 2), 20))))
        Set oGar = Nothing
        Call ActualizaMontoCobertura
End With

End Sub

Private Sub ActualizaMontoCobertura()
Dim i As Integer
Dim nSumaCobertura As Double

nSumaCobertura = 0
With FeTabla
        For i = 0 To .Rows - 2
            If .TextMatrix(i + 1, 3) <> "" Then
                nSumaCobertura = nSumaCobertura + CDbl(.TextMatrix(i + 1, 3))
            End If
        Next
End With

lblCoberturaCredito.Caption = Format(nSumaCobertura, "#0.00")

txtMontoRea.Text = Format(nSumaCobertura * 1000, "#0.00")
txtMontotas.Text = Format(nSumaCobertura * 1000, "#0.00")
txtMontoxGrav.Text = Format(nSumaCobertura * 1000, "#0.00")


End Sub

Private Sub FeTabla_RowColChange()
Dim nTipoEval As Integer
Dim nTipoCredito As Integer
Dim oGar As COMDCredito.DCOMGarantia

If FeTabla.TextMatrix(1, 1) = "" Then
    MsgBox "Seleccione un Tipo de Credito", vbInformation, "Aviso"
    Exit Sub
End If

With FeTabla
    If .Col = 2 Then
        Set oGar = New COMDCredito.DCOMGarantia
        .CargaCombo oGar.RecuperaValoresTablaValores(CInt(Trim(Right(cmbTipoCreditoTabla, 10))), CInt(.TextMatrix(.row, 0)))
        Set oGar = Nothing
    End If
End With
End Sub

Private Sub Form_Load()
    gGarantiaDepPlazoFijoCF = False
    
    'madm 20100817 - viene de CF con GPF
    gGarantiaDepPlazoFijoCF = IIf(frmCFSolicitud.Visible, True, False)
    
    CentraForm Me
    SSGarant.Tab = 0
    SSGarant.TabVisible(2) = False
    SSGarant.TabVisible(3) = False
    SSGarant.TabVisible(4) = False
    
    SSGarant.TabVisible(6) = False
    SSGarant.TabVisible(7) = False 'BRGO 20111125
    SSGarant.TabVisible(8) = False 'BRGO 20111125
    SSGarant.TabVisible(9) = False 'ALPA 20120320
    
    bEstadoCargando = True
    Call CargaControles
    Call HabilitaIngreso(False)
    CmbDocGarant.Enabled = False
    txtNumDoc.Enabled = True
    cmdBuscar.Enabled = True
    bEstadoCargando = False
    cmdEjecutar = -1
    cmdEliminar.Enabled = False
    cmdEditar.Enabled = False
    CboGarantia.Enabled = False
    cmdNuevo.Enabled = True
    bCreditoCF = False
    bValdiCCF = False
    
    SSGarant.TabVisible(5) = False
    CboBanco.Enabled = False
    CboBanco.ListIndex = -1
    
    gcPermiteModificar = True 'peac 20071127
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredRegistrarGarantiaCli
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub



Private Sub OptCG_Click(Index As Integer)
    If OptCG(0).value = True Then
        OptTR(0).Enabled = False
        OptTR(1).Enabled = False
        OptTR(0).value = True
    Else
        OptTR(0).Enabled = True
        OptTR(1).Enabled = True
        OptTR(0).value = True
    End If
End Sub

Private Sub Option1_Click()
    
End Sub





Private Sub SSGarant_Click(PreviousTab As Integer)

    If PreviousTab = 0 Then
        If CmdCliAceptar.Visible And CmdCliAceptar.Enabled Then
            MsgBox "Pulse Aceptar para Registrar al Cliente", vbInformation, "Aviso"
            CmdCliAceptar.SetFocus
            SSGarant.Tab = 0
        End If
    End If
   
    If PreviousTab = 3 Then
        If CmdDJAceptar.Visible And CmdDJAceptar.Enabled Then
            MsgBox "Pulse Aceptar para Registrar el Item", vbInformation, "Aviso"
            CmdDJAceptar.SetFocus
            SSGarant.Tab = 3
        End If
    End If
    
    'By Capi 25092008 porque se elimino el control cboEstadoTasInm
    'If SSGarant.Tab = 2 And cboEstadoTasInm.ListCount = 0 Then
    If SSGarant.Tab = 2 Then
    '
        Dim oCons As COMDConstantes.DCOMConstantes
        Dim rs As ADODB.Recordset
        Dim rsTmp As ADODB.Recordset
        Dim rsTmp1 As ADODB.Recordset
        Set oCons = New COMDConstantes.DCOMConstantes
        Set rs = oCons.RecuperaConstantes(gCredGarantEstadoTasacion)
        Set rsTmp = rs.Clone
        Set rsTmp1 = rs.Clone
        Set oCons = Nothing
        'By Capi 25092008 porque se elimino el control cboEstadoTasInm y cboEstadoTasVeh
        'Call Llenar_Combo_con_Recordset(rs, cboEstadoTasInm)
        'Call Llenar_Combo_con_Recordset(rsTmp, cboEstadoTasVeh)
        'End
        
        'Call Llenar_Combo_con_Recordset(rsTmp1, cboEstadoTasMaq)
    End If
    
    'ARCV 11-07-2006
    If SSGarant.Tab = 5 Then
        If cmbTipoCreditoTabla.Text = "" Then
            Dim OCon As COMDConstantes.DCOMConstantes
            Dim rsT As ADODB.Recordset
            Set OCon = New COMDConstantes.DCOMConstantes
            Set rsT = OCon.RecuperaConstantes(9062)
            Set OCon = Nothing
            Call Llenar_Combo_con_Recordset(rsT, cmbTipoCreditoTabla)
        End If
    End If
       
End Sub

Private Sub txtAnioConstruccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtcomentarios_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtDescGarant_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
'        If CmdCliNuevo.Enabled And CmdCliNuevo.Visible Then
'            CmdCliNuevo.SetFocus
'        End If
            If Me.CboNumPF.Visible = True Then
                If CmdCliNuevo.Enabled And CmdCliNuevo.Visible Then
                    CmdCliNuevo.SetFocus
                End If
            Else
                cmbMoneda.SetFocus
            End If
     End If
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPlacaVehic.Enabled And txtPlacaVehic.Visible Then
        txtPlacaVehic.SetFocus
        End If
    End If
End Sub

Private Sub txtDireAlma_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTipoMerca.SetFocus
    End If
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtMontotas.Visible <> False Then
    If txtMontotas.Enabled Then txtMontotas.SetFocus
    End If
End If
End Sub

Private Sub TxtFecCons_GotFocus()
    fEnfoque TxtFecCons
End Sub

Private Sub TxtFecCons_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    '    TxtFecTas.SetFocus
    End If
End Sub

Private Sub txtFechaCertifGravamen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtFechaCertifGravamenGA_Change()
txtFechaCertifGravamen.Text = txtFechaCertifGravamenGA.Text
End Sub

'ALPA 20120320**************************************
Private Sub txtFechaCertifGravamenGA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub
'***************************************************
Private Sub TxtFechareg_GotFocus()
    fEnfoque TxtFechareg
End Sub

Private Sub TxtFechareg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRegNro.SetFocus
    End If
End Sub

Private Sub TxtFechareg_LostFocus()
Dim sCad As String
    If TxtFechareg.Text = "__/__/____" Then Exit Sub
    sCad = ValidaFecha(TxtFechareg.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFechareg.SetFocus
    End If
End Sub

Private Sub TxtFecTas_GotFocus()
    fEnfoque TxtFecTas
End Sub

Private Sub TxtFecTas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtFecTasVeh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAnioFab.SetFocus
    End If

End Sub

'Private Sub TxtFecVctoPol_GotFocus()
'    fEnfoque TxtFecVctoPol
'End Sub
'
'Private Sub TxtFecVctoPol_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub

Private Sub TxtFecVig_GotFocus()
    fEnfoque TxtFecVig
End Sub

Private Sub TxtFecVig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoPol.SetFocus
    End If
End Sub
Private Sub txtNumGar_GotFocus()
    fEnfoque txtNumGar
End Sub

Private Sub txtNumGar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    ElseIf KeyAscii <> 8 Then   'El 8 es la tecla de borrar (backspace)
        'Si después de añadirle la tecla actual no es un número...
        If Not IsNumeric("0" & txtNumGar.Text & Chr(KeyAscii)) Then
        '... se desecha esa tecla y se avisa de que no es correcta
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumGar_LostFocus()
    If Trim(txtNumGar.Text) = "" Then
        txtNumGar.Text = "0"
    Else
        txtNumGar.Text = Format(txtNumGar.Text, "#0")
    End If
End Sub

'Private Sub TxtHipCuotaIni_GotFocus()
'    fEnfoque TxtHipCuotaIni
'End Sub
'
'Private Sub TxtHipCuotaIni_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtHipCuotaIni, KeyAscii)
'    If KeyAscii = 13 Then
'        TxtMontoHip.SetFocus
'    End If
'End Sub
'
'Private Sub TxtHipCuotaIni_LostFocus()
'    If Trim(TxtHipCuotaIni.Text) = "" Then
'        TxtHipCuotaIni.Text = "0.00"
'    Else
'        TxtHipCuotaIni.Text = Format(TxtHipCuotaIni.Text, "#0.00")
'    End If
'End Sub
'
'Private Sub TxtMontoHip_GotFocus()
'    fEnfoque TxtMontoHip
'End Sub

Private Sub txtNumLocales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtNumMotor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNumSerie.SetFocus
    End If

End Sub

Private Sub txtNumPisos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtNumSotanos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtPlacaVehic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNumMotor.SetFocus
    End If
End Sub

'peac 20071123
Private Sub TxtValorGravado_GotFocus()
    fEnfoque txtValorGravado
    fEnfoque txtValorGravadoGA
End Sub

Private Sub TxtValorGravadoGA_GotFocus()
    fEnfoque txtValorGravado
    fEnfoque txtValorGravadoGA
End Sub

'Private Sub TxtMontoHip_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtMontoHip, KeyAscii)
'    If KeyAscii = 13 Then
'        'TxtPrecioVenta.SetFocus
'    End If
'End Sub

'peac 20071123
Private Sub TxtValorGravado_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorGravado, KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

'ALPA 20120320
Private Sub TxtValorGravadoGA_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorGravadoGA, KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

'ALPA 20120320
Private Sub TxtValorGravadoGA_LostFocus()
    If Trim(txtValorGravado.Text) = "" Then
        txtValorGravado.Text = "0.00"
        txtValorGravadoGA.Text = "0.00" 'ALPA 20120320
    Else
        txtValorGravado.Text = Format(txtValorGravadoGA.Text, "#0.00")
        txtValorGravadoGA.Text = Format(txtValorGravadoGA.Text, "#0.00") 'ALPA 20120320
    End If
End Sub

'Private Sub TxtMontoHip_LostFocus()
'    If Trim(TxtMontoHip.Text) = "" Then
'        TxtMontoHip.Text = "0.00"
'    Else
'        TxtMontoHip.Text = Format(TxtMontoHip.Text, "#0.00")
'    End If
'End Sub

' Peac 20071123
Private Sub TxtValorGravado_LostFocus()
    If Trim(txtValorGravado.Text) = "" Then
        txtValorGravado.Text = "0.00"
        txtValorGravadoGA.Text = "0.00" 'ALPA 20120320
    Else
        txtValorGravado.Text = Format(txtValorGravado.Text, "#0.00")
        txtValorGravadoGA.Text = Format(txtValorGravado.Text, "#0.00") 'ALPA 20120320
    End If
End Sub


Private Sub TxtMontoPol_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoPol, KeyAscii)
    If KeyAscii = 13 Then
        TxtFecCons.SetFocus
    End If
End Sub

Private Sub TxtMontoPol_LostFocus()
    If Trim(TxtMontoPol.Text) = "" Then
        TxtMontoPol.Text = "0.00"
    End If
    TxtMontoPol.Text = Format(TxtMontoPol.Text, "#0.00")
End Sub

Private Sub txtMontoRea_Change()
    'txtMontoxGrav.Text = "0.00"
End Sub

Private Sub txtMontoRea_GotFocus()
    fEnfoque txtMontoRea
End Sub

Private Sub txtMontoRea_KeyPress(KeyAscii As Integer)
Dim oGarantia As COMNCredito.NCOMGarantia
Dim oPersona As DPersona
Dim nValor As Double
Dim sCad As String
Dim oMantGarant As COMDCredito.DCOMGarantia
Dim oDCF As COMDCartaFianza.DCOMCartaFianza

     KeyAscii = NumerosDecimales(txtMontoRea, KeyAscii)
     If KeyAscii = 13 Then
        Set oGarantia = New COMNCredito.NCOMGarantia

        If bValdiCCF = True Then
            Set oDCF = New COMDCartaFianza.DCOMCartaFianza
            nValor = oDCF.ValorCoberturaGarantia
            Set oDCF = Nothing
        Else
          nValor = oGarantia.PorcentajeGarantia(gPersGarantia & Trim(Right(CmbTipoGarant.Text, 10)) & IIf(fbOrigenCF And Trim(Right(CmbTipoGarant.Text, 10)) = "1", "4", "")) 'WIOR 20140826 AGREGO IIf(FBORIGEN, "4", "")
        End If
        
        Set oGarantia = Nothing
        If nValor > 1 Then
            txtMontoxGrav.Text = Format(nValor, "#0.00")
        Else
            txtMontoxGrav.Text = Format(nValor * CDbl(txtMontoRea.Text), "#0.00")
        End If
        
        '06-05-2005
        Set oGarantia = New COMNCredito.NCOMGarantia
        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)  ', CLng(Trim(Right(CmbTipoGarant.Text, 10))))
        If Not sCad = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            txtMontoRea.SetFocus
            Exit Sub
        End If
        '*******************************************
'            End If
      ' End If
       Set oPersona = Nothing
        If txtMontoxGrav.Enabled Then
            txtMontoxGrav.SetFocus
        Else
            txtMontoRea.Text = Format(txtMontoRea.Text, "#0.00")
        End If
     End If







End Sub
Function ValidacionCreditoAutomatico(ByVal psCtaCod As String) As Boolean

End Function

Function ObtenerTitularX() As String
    Dim i As Integer
    If FERelPers.Rows = 2 And FERelPers.TextMatrix(1, 1) = "" Then
       ObtenerTitularX = ""
    Else
        ReDim RelPers(FERelPers.Rows - 1, 4)
        For i = 1 To FERelPers.Rows - 1
            If Right("00" & Trim(Right(FERelPers.TextMatrix(i, 3), 10)), 2) = "01" Then
                ObtenerTitularX = FERelPers.TextMatrix(i, 1)   'Codigo de Persona
                Exit For
            End If
        Next i
    End If
End Function

Private Sub txtMontoRea_LostFocus()
    If Trim(txtMontoRea.Text) = "" Then
        txtMontoRea.Text = "0.00"
    Else
        txtMontoRea.Text = Format(txtMontoRea.Text, "#0.00")
    End If
    'Call txtMontoRea_KeyPress(13)
End Sub

Private Sub txtMontotas_GotFocus()
    fEnfoque txtMontotas
End Sub

Private Sub txtMontotas_KeyPress(KeyAscii As Integer)

     KeyAscii = NumerosDecimales(txtMontotas, KeyAscii)
     If KeyAscii = 13 Then
        txtMontoRea.SetFocus
     End If
End Sub

Private Sub txtMontotas_LostFocus()
    If Trim(txtMontotas.Text) = "" Then
        txtMontotas.Text = "0.00"
    Else
        txtMontotas.Text = Format(txtMontotas.Text, "#0.00")
    End If
End Sub

Private Sub txtMontoxGrav_GotFocus()
    fEnfoque txtMontoxGrav
End Sub

Private Sub txtMontoxGrav_KeyPress(KeyAscii As Integer)
Dim oGarantia As COMNCredito.NCOMGarantia
Dim sCad As String

     KeyAscii = NumerosDecimales(txtMontoxGrav, KeyAscii)
     If KeyAscii = 13 Then
        Set oGarantia = New COMNCredito.NCOMGarantia
        If Trim(txtMontoxGrav.Text) = "" Then
            txtMontoxGrav.Text = "0.00"
        End If
        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)
        If Not sCad = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        'txtcomentarios.SetFocus
        cmdAceptar.SetFocus
     End If
     
End Sub

Private Sub txtMontoxGrav_LostFocus()
    If Trim(txtMontoxGrav.Text) = "" Then
        txtMontoxGrav.Text = "0.00"
    Else
        txtMontoxGrav.Text = Format(txtMontoxGrav.Text, "#0.00")
    End If
End Sub


Private Sub TxtNroPoliza_GotFocus()
    fEnfoque TxtNroPoliza
End Sub

Private Sub TxtNroPoliza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtFecVig.SetFocus
    End If
End Sub

Private Sub txtNumDoc_GotFocus()
    fEnfoque txtNumDoc
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        If Trim(Right(cmbMoneda.Text, 2)) = "2" Then
            txtMontotas.BackColor = RGB(200, 255, 200)
            txtMontoRea.BackColor = RGB(200, 255, 200)
            txtMontoxGrav.BackColor = RGB(200, 255, 200)
        Else
            txtMontotas.BackColor = vbWhite
            txtMontoRea.BackColor = vbWhite
            txtMontoxGrav.BackColor = vbWhite
        End If
        If txtDescGarant.Enabled Then
            txtDescGarant.SetFocus
        End If
     End If
End Sub

'Private Sub TxtPrecioVenta_GotFocus()
'    fEnfoque TxtPrecioVenta
'End Sub
'
'Private Sub TxtPrecioVenta_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtPrecioVenta, KeyAscii)
'If KeyAscii = 13 Then
'    SendKeys "{Tab}", True
'End If
'End Sub
'
'Private Sub TxtPrecioVenta_LostFocus()
'Dim oGarantia As COMNCredito.NCOMGarantia
'Dim sCad As String
'Dim nPorc As Double
'
'    If Trim(TxtPrecioVenta.Text) = "" Then
'        TxtPrecioVenta.Text = "0.00"
'    Else
'        TxtPrecioVenta.Text = Format(TxtPrecioVenta.Text, "#0.00")
'    End If
'
'    If Trim(txtMontotas.Text) = "" Then
'        txtMontotas.Text = "0.00"
'    End If
'
'    If Me.ChkGarReal.value = 1 Then
''ARCV 27-01-2007
''        Set oGarantia = New COMNCredito.NCOMGarantia
''        If Trim(Right(CmbTipoGarant.Text, 5)) <> "" Then
''            If CInt(Trim(Right(CmbTipoGarant.Text, 5))) = gPersGarantiaHipotecas Then
''                nPorc = oGarantia.PorcentajeGarantia("3051")
''                If Abs(((CDbl(txtMontotas.Text) - CDbl(TxtPrecioVenta.Text)) / CDbl(txtMontotas.Text))) > nPorc Then
''                    MsgBox "La Diferencia entre el Monto de Tasacion y el Precio de Venta no debe ser mayor a " & Format(nPorc * 100, "#0.00") & "%", vbInformation, "Aviso"
''                    TxtPrecioVenta.Text = Format(((100 - nPorc) / 100) * CDbl(txtMontotas.Text), "#0.00")
''                    TxtPrecioVenta.SetFocus
''                    Exit Sub
''                End If
''            End If
''        End If
''        Set oGarantia = Nothing
'    End If
'End Sub


Private Sub TxtRegNro_GotFocus()
    fEnfoque TxtRegNro
End Sub

Private Sub TxtRegNro_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtTelefono_GotFocus()
    fEnfoque TxtTelefono
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CboTipoInmueb.SetFocus
    End If
End Sub

'Private Sub TxtValorCConst_GotFocus()
'    fEnfoque TxtValorCConst
'End Sub
'
'Private Sub TxtValorCConst_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtValorCConst, KeyAscii)
'    If KeyAscii = 13 Then
'        'CmdBuscaTasa.SetFocus
'        TxtValorEdificacion.SetFocus
'    End If
'End Sub
'
'Private Sub TxtValorCConst_LostFocus()
'    If Trim(TxtValorCConst.Text) = "" Then
'        TxtValorCConst.Text = "0.00"
'    Else
'        TxtValorCConst.Text = Format(TxtValorCConst.Text, "#0.00")
'    End If
'
'End Sub

Sub CargarSuperGarantias(ByVal pRs As ADODB.Recordset, Optional CF As Boolean = False)
'    Dim objDGarantias As COMDCredito.DCOMGarantia
'    Dim rs As ADODB.Recordset
    Dim sDes As String
    Dim nCodigo As Integer
    On Error GoTo ErrHandler
        'Set objDGarantias = New COMDCredito.DCOMGarantia
        'Set rs = objDGarantias.ListaSuperGarantias
        'Set objDGarantias = Nothing
        
        Do Until pRs.EOF
            nCodigo = pRs!nConsValor
            sDes = pRs!cConsDescripcion
            
            If Not (sDes = "GARANTIAS NO PREFERIDAS" And CF) Then
                CboGarantia.AddItem sDes
                CboGarantia.ItemData(CboGarantia.NewIndex) = nCodigo
            End If
            
            pRs.MoveNext
        Loop
    Exit Sub
ErrHandler:
    'If Not objDGarantias Is Nothing Then Set objDGarantias = Nothing
    'If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error al cargar las garantias", vbInformation, "AVISO"
End Sub

'peac 20071116  para activarlo despues
Sub CargarClaseInmueble(ByVal prsClaInm As ADODB.Recordset)
    Dim sDes As String
    Dim nCodigo As Integer
    On Error GoTo ErrHandler

        Do Until prsClaInm.EOF
            nCodigo = prsClaInm!nConsValor
            sDes = prsClaInm!cConsDescripcion

            cmbClaseInmueble.AddItem Trim(sDes) & Space(100) & Trim(Str(nCodigo))
            cmbClaseInmueble.ItemData(cmbClaseInmueble.NewIndex) = nCodigo
            
            prsClaInm.MoveNext
        Loop
    Exit Sub
ErrHandler:
    MsgBox "Error al cargar datos", vbInformation, "AVISO"
End Sub
    
'peac 20071116  'para activarlo despues
Sub CargarCategoria(ByVal prsCate As ADODB.Recordset)
    Dim sDes As String
    Dim nCodigo As Integer
    On Error GoTo ErrHandler

        Do Until prsCate.EOF
            nCodigo = prsCate!nConsValor
            sDes = prsCate!cConsDescripcion

            cmbCategoria.AddItem Trim(sDes) & Space(100) & Trim(Str(nCodigo))
            cmbCategoria.ItemData(cmbCategoria.NewIndex) = nCodigo

            prsCate.MoveNext
        Loop
    Exit Sub
ErrHandler:
    MsgBox "Error al cargar datos", vbInformation, "AVISO"
End Sub
    


'Sub ReconfigurarSubTipoGarant(ByVal pIdTipoGarant As Integer)
'    Dim i As Integer
'
'    For i = 0 To CmbTipoGarant.ListCount - 1
'        If Trim(Left(CmbTipoGarant.List(i), 3)) = pIdTipoGarant Then
'            CmbTipoGarant.RemoveItem (i)
'        End If
'    Next i
'End Sub

Sub ReLoadCmbTipoGarant(ByVal pnIdSuperGarant As Integer)
    Dim rs As ADODB.Recordset
    Dim objDGarantia As COMDCredito.DCOMGarantia
    Dim nCodigo As Integer
    On Error GoTo ErrHandler
        Set objDGarantia = New COMDCredito.DCOMGarantia
        Set rs = objDGarantia.CargarRelGarantia(pnIdSuperGarant)
        Set objDGarantia = Nothing
        If Not rs.EOF And Not rs.BOF Then
            CmbTipoGarant.Clear
        End If
                             
        Do Until rs.EOF
             nCodigo = rs!nConsValor
             'madm 20100826
            If (gGarantiaDepPlazoFijoCF) And Me.CboGarantia.Text = "GARANTIAS PREFERIDAS" Then
                If nCodigo = 1 Then
                    CmbTipoGarant.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim(Str(rs!nConsValor))
                 End If
            Else
                CmbTipoGarant.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim(Str(rs!nConsValor))
            End If
            'end madm
           rs.MoveNext
        Loop
        Set rs = Nothing
    Exit Sub
ErrHandler:
    If Not objDGarantia Is Nothing Then Set objDGarantia = Null
    If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error al cargaer"
End Sub


Sub PosicionSuperGarantias(ByVal pintIndex As Integer)
'    Dim objDGarantia As DGarantia
    Dim i As Integer
    Dim nValorGarant As Integer
    On Error GoTo ErrHandler
'        Set objDGarantia = New DGarantia
'        nValorGarant = objDGarantia.ObtenerIdSuperGarantia(pintIndex)
'        Set objDGarantia = Nothing
        
        For i = 0 To CboGarantia.ListCount - 1
               If CboGarantia.ItemData(i) = pintIndex Then
                  CboGarantia.ListIndex = i
                  Exit For
               End If
        Next i
    Exit Sub
ErrHandler:
    'If Not objDGarantia Is Nothing Then Set objDGarantia = Nothing
    MsgBox "Error a cargar super garantia", vbInformation, "AVISO"
End Sub

Public Function IsLoadForm(ByVal FormCaption As String, Optional Active As Variant) As Boolean
    Dim rtn As Integer, i As Integer
    Dim Name As String
        
    rtn = False
    Name = LCase(FormCaption)
    Do Until i > Forms.Count - 1 Or rtn
        If LCase(Forms(i).Caption) = FormCaption Then

        rtn = True

End If
        i = i + 1
    Loop
    
    If rtn Then
        If Not IsMissing(Active) Then
            If Active Then
                Forms(i - 1).WindowState = vbNormal
            End If
        End If
    End If
    IsLoadForm = rtn
End Function

Sub ImprimirDeclaraciónJurada()
    Dim sCadImp As String
    Dim Prev As previo.clsprevio
    Dim i As Integer
    Dim sTitular As String
    
    If FERelPers.TextMatrix(1, 1) = "" Then
        MsgBox "Ingrese los Titulares de la Garantia", vbInformation, "Aviso"
        Exit Sub
    Else
        For i = 1 To FERelPers.Rows - 1
            If Right(FERelPers.TextMatrix(i, 3), 1) = "1" Then
                sTitular = PstaNombre(FERelPers.TextMatrix(1, 2), False)
            End If
        Next i
    End If
    
    i = 0
    If FEDeclaracionJur.TextMatrix(1, 1) = "" Then
        MsgBox "No existen Datos para Imprmir", vbInformation, "Aviso"
        Exit Sub
    Else
        Set Prev = New clsprevio
        'Impresion del Titulo
        sCadImp = sCadImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sCadImp = sCadImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        'Imprime Cliente
        sCadImp = sCadImp & Space(14) & sTitular & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sCadImp = sCadImp & Space(14) & txtDireccion & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        sCadImp = sCadImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        'Impresion de la Cabecera
        For i = 1 To FEDeclaracionJur.Rows - 1
           sCadImp = sCadImp & Space(7) & ImpreFormat(FEDeclaracionJur.TextMatrix(i, 2), 6) & ImpreFormat(FEDeclaracionJur.TextMatrix(i, 1), 60) & ImpreFormat(FEDeclaracionJur.TextMatrix(i, 4), 20) & ImpreFormat(FEDeclaracionJur.TextMatrix(i, 5), 20) & ImpreFormat(Format(CDbl(FEDeclaracionJur.TextMatrix(i, 3) * FEDeclaracionJur.TextMatrix(i, 2)), "0.00"), 25) & oImpresora.gPrnSaltoLinea
        Next i
            
        Prev.Show sCadImp, "", False
        Set Prev = Nothing
            
    End If
End Sub

Private Sub TxtValorEdificacion_GotFocus()
    fEnfoque TxtValorEdificacion
End Sub

Private Sub txtValorMerca_GotFocus()
    fEnfoque txtValorMerca
End Sub

Private Sub txtValorMerca_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtValorMerca, KeyAscii)
    If KeyAscii = 13 Then
        txtDireAlma.SetFocus
    End If
End Sub

Private Sub txtValorMerca_LostFocus()
    If Trim(txtValorMerca.Text) = "" Then
        txtValorMerca.Text = "0.00"
    Else
        txtValorMerca.Text = Format(txtValorMerca.Text, "#0.00")
    End If
End Sub

'peac 20071123
Private Sub TxtVRM_GotFocus()
    fEnfoque txtVRM
End Sub


Private Sub TxtValorEdificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtValorEdificacion, KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

'peac 20071123
Private Sub TxtVRM_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVRM, KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub


Private Sub TxtValorEdificacion_LostFocus()
    If Trim(TxtValorEdificacion.Text) = "" Then
        TxtValorEdificacion.Text = "0.00"
    Else
        TxtValorEdificacion.Text = Format(TxtValorEdificacion.Text, "#0.00")
    End If
    
End Sub

'PEAC 20071122
Private Sub TxtVRM_LostFocus()
    If Trim(txtVRM.Text) = "" Then
        txtVRM.Text = "0.00"
    Else
        txtVRM.Text = Format(txtVRM.Text, "#0.00")
    End If
End Sub

