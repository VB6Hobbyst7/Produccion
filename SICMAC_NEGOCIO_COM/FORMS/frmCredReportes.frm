VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCredReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Creditos"
   ClientHeight    =   7605
   ClientLeft      =   1170
   ClientTop       =   2055
   ClientWidth     =   11775
   Icon            =   "frmCredReportes.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmMontos 
      Caption         =   "Montos Saldos"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   8730
      TabIndex        =   124
      Top             =   5390
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox TxtMontoFin 
         Height          =   285
         Left            =   120
         TabIndex        =   127
         Top             =   510
         Width           =   1065
      End
      Begin VB.TextBox TxtMontoIni 
         Height          =   285
         Left            =   120
         TabIndex        =   125
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label22 
         Caption         =   "Fin"
         Height          =   285
         Left            =   1230
         TabIndex        =   128
         Top             =   540
         Width           =   420
      End
      Begin VB.Label Label21 
         Caption         =   "Inicio"
         Height          =   270
         Left            =   1215
         TabIndex        =   126
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   5535
      TabIndex        =   112
      Top             =   6900
      Width           =   6210
      Begin ComctlLib.ProgressBar PbProgresoReporte 
         Height          =   375
         Left            =   4680
         TabIndex        =   177
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CheckBox chkMigraExcell 
         Caption         =   "Migra a Excell"
         Height          =   345
         Left            =   150
         TabIndex        =   133
         Top             =   195
         Width           =   1170
      End
      Begin VB.CommandButton CmdImprimirA02 
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
         Height          =   390
         Left            =   1440
         TabIndex        =   69
         Top             =   210
         Width           =   1380
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
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
         Left            =   3120
         TabIndex        =   1
         Top             =   180
         Width           =   1380
      End
      Begin MSComDlg.CommonDialog dlgFileSave 
         Left            =   960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame fraProductos1 
      Height          =   300
      Left            =   12150
      TabIndex        =   90
      Top             =   2775
      Visible         =   0   'False
      Width           =   300
      Begin VB.CheckBox chkHipotecario1 
         Caption         =   "Mi Vivienda"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   109
         Tag             =   "423"
         Top             =   4635
         Width           =   1350
      End
      Begin VB.CheckBox chkHipotecario1 
         Caption         =   "Hipotecaja"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   108
         Tag             =   "401"
         Top             =   4368
         Width           =   1080
      End
      Begin VB.CheckBox chkProducto1 
         Caption         =   "Hipotecario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   75
         TabIndex        =   107
         Top             =   4119
         Width           =   1515
      End
      Begin VB.CheckBox chkConsumo1 
         Caption         =   "Prestamos Admin."
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   106
         Tag             =   "320"
         Top             =   3870
         Width           =   1605
      End
      Begin VB.CheckBox chkConsumo1 
         Caption         =   "Usos Diversos"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   105
         Tag             =   "304"
         Top             =   3621
         Width           =   1470
      End
      Begin VB.CheckBox chkConsumo1 
         Caption         =   "Garantia CTS"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   104
         Tag             =   "303"
         Top             =   3372
         Width           =   1590
      End
      Begin VB.CheckBox chkConsumo1 
         Caption         =   "Garantia Plazo Fijo"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   103
         Tag             =   "302"
         Top             =   3123
         Width           =   1650
      End
      Begin VB.CheckBox chkConsumo1 
         Caption         =   "Descuento x Planilla"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   102
         Tag             =   "301"
         Top             =   2874
         Width           =   1755
      End
      Begin VB.CheckBox chkProducto1 
         Caption         =   "Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   2
         Left            =   75
         TabIndex        =   101
         Top             =   2610
         Width           =   1080
      End
      Begin VB.CheckBox chkMicroEmpresa1 
         Caption         =   "PYME Carta Fianza"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   100
         Tag             =   "221"
         Top             =   2370
         Width           =   1755
      End
      Begin VB.CheckBox chkMicroEmpresa1 
         Caption         =   "PYME Agropecuario"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   99
         Tag             =   "203"
         Top             =   2112
         Width           =   1740
      End
      Begin VB.CheckBox chkMicroEmpresa1 
         Caption         =   "PYME Pesquero"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   98
         Tag             =   "202"
         Top             =   1863
         Width           =   1740
      End
      Begin VB.CheckBox chkMicroEmpresa1 
         Caption         =   "PYME Empresarial"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   97
         Tag             =   "201"
         Top             =   1614
         Width           =   1755
      End
      Begin VB.CheckBox chkProducto1 
         Caption         =   "Micro Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   96
         Top             =   1365
         Width           =   1455
      End
      Begin VB.CheckBox chkComercial1 
         Caption         =   "Carta Fianza"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   95
         Tag             =   "121"
         Top             =   1116
         Width           =   1455
      End
      Begin VB.CheckBox chkComercial1 
         Caption         =   "Agropecuario"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   94
         Tag             =   "103"
         Top             =   867
         Width           =   1380
      End
      Begin VB.CheckBox chkComercial1 
         Caption         =   "Pesquero"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   93
         Tag             =   "102"
         Top             =   618
         Width           =   1080
      End
      Begin VB.CheckBox chkComercial1 
         Caption         =   "Empresarial"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   92
         Tag             =   "101"
         Top             =   369
         Width           =   1200
      End
      Begin VB.CheckBox chkProducto1 
         Caption         =   "Comercial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   91
         Top             =   105
         Width           =   1080
      End
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   225
      Left            =   8640
      TabIndex        =   31
      Top             =   7125
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   397
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCredReportes.frx":030A
   End
   Begin VB.Frame Frame1 
      Height          =   7560
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin MSComctlLib.TreeView TVRep 
         Height          =   7275
         Left            =   100
         TabIndex        =   72
         Top             =   200
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   12832
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   300
         Top             =   5820
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes.frx":038C
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes.frx":06DE
               Key             =   "Bebe"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes.frx":0A30
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes.frx":0D82
               Key             =   "Hijito"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox CR 
         Height          =   480
         Left            =   4125
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   135
         Top             =   5820
         Width           =   1200
      End
   End
   Begin VB.Frame FraAutomaticos 
      Caption         =   "Productos Auto"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   5775
      TabIndex        =   119
      Top             =   5895
      Visible         =   0   'False
      Width           =   1485
      Begin VB.CheckBox ChkConsumo 
         Caption         =   "Consumo"
         Height          =   255
         Left            =   90
         TabIndex        =   121
         Top             =   540
         Width           =   1155
      End
      Begin VB.CheckBox ChkMes 
         Caption         =   "Mes"
         Height          =   255
         Left            =   90
         TabIndex        =   120
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FraConUbiGeo 
      Caption         =   "Ubic.Geografica"
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   5775
      TabIndex        =   122
      Top             =   5985
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CheckBox ChkUbi 
         Caption         =   "Ubicacion Geografica"
         Height          =   315
         Left            =   150
         TabIndex        =   123
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10200
      TabIndex        =   87
      Top             =   6420
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9405
      TabIndex        =   86
      Top             =   6435
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CheckBox ChkRFA 
      Caption         =   "No Incluir Creditos RFA"
      Height          =   225
      Left            =   8760
      TabIndex        =   118
      Top             =   5385
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Frame FraA02 
      Height          =   6900
      Index           =   0
      Left            =   5520
      TabIndex        =   2
      Top             =   0
      Width           =   6195
      Begin VB.Frame FraTipAce 
         Caption         =   "Tipo Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1140
         Left            =   0
         TabIndex        =   173
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
         Begin VB.OptionButton optCredVig 
            Caption         =   "Total"
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   176
            Top             =   840
            Width           =   1470
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Normal (Fechas)"
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   175
            Top             =   300
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Cancelados (Fechas)"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   174
            Top             =   562
            Width           =   1830
         End
      End
      Begin VB.Frame FraACE 
         Caption         =   "Tipo Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   780
         Left            =   0
         TabIndex        =   170
         Top             =   3480
         Visible         =   0   'False
         Width           =   1710
         Begin VB.OptionButton optCredVig 
            Caption         =   "Mancomunos"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   172
            Top             =   562
            Width           =   1470
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Titular "
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   171
            Top             =   300
            Value           =   -1  'True
            Width           =   1470
         End
      End
      Begin VB.Frame frmResAnalista 
         Height          =   2655
         Left            =   3960
         TabIndex        =   161
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
         Begin MSDataListLib.DataCombo dcAgencia 
            Height          =   315
            Left            =   120
            TabIndex        =   162
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcCartera 
            Height          =   315
            Left            =   120
            TabIndex        =   163
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox txtFCierre 
            Height          =   300
            Left            =   120
            TabIndex        =   168
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFCierreAnt 
            Height          =   300
            Left            =   120
            TabIndex        =   169
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label26 
            Caption         =   "T. Cartera:"
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "F. Cierre Ant:"
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label24 
            Caption         =   "F. Cierre:"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Agencia:"
            Height          =   255
            Left            =   120
            TabIndex        =   164
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame fratipo 
         Caption         =   "Tipo"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   120
         TabIndex        =   158
         Top             =   4320
         Visible         =   0   'False
         Width           =   1485
         Begin VB.CheckBox chkcancelados 
            Caption         =   "Cancelados"
            Height          =   255
            Left            =   120
            TabIndex        =   160
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkvigentes 
            Caption         =   "Vigentes"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame fraEstados 
         Caption         =   "Condicion"
         Height          =   1260
         Left            =   4440
         TabIndex        =   153
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
         Begin VB.CheckBox chkEstado 
            Caption         =   "Judicial"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   157
            Top             =   195
            Width           =   1110
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Castigado"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   156
            Top             =   465
            Width           =   1185
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Refinanciado"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   155
            Top             =   720
            Width           =   1305
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Vigente"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   154
            Top             =   960
            Width           =   1185
         End
      End
      Begin VB.Frame fraCondBN 
         Caption         =   "Condicion BN"
         Height          =   855
         Left            =   0
         TabIndex        =   151
         Top             =   4320
         Width           =   1890
         Begin VB.CheckBox chkCondBN 
            Caption         =   "Solo Cred Desmb Banco Nación"
            Height          =   495
            Left            =   120
            TabIndex        =   152
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.Frame FraCondicion 
         Caption         =   "Condicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1440
         Left            =   0
         TabIndex        =   144
         Top             =   2880
         Visible         =   0   'False
         Width           =   1890
         Begin VB.CheckBox ChkCond 
            Caption         =   "Adicional"
            Height          =   210
            Index           =   6
            Left            =   960
            TabIndex        =   178
            Tag             =   "7"
            Top             =   870
            Width           =   990
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Automa"
            Height          =   210
            Index           =   5
            Left            =   975
            TabIndex        =   150
            Tag             =   "6"
            Top             =   600
            Width           =   870
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Ampliad"
            Height          =   210
            Index           =   4
            Left            =   960
            TabIndex        =   149
            Tag             =   "5"
            Top             =   315
            Width           =   870
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Recurren."
            Height          =   210
            Index           =   3
            Left            =   75
            TabIndex        =   148
            Tag             =   "2"
            Top             =   1125
            Width           =   1020
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Refinan."
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   147
            Tag             =   "4"
            Top             =   870
            Width           =   1320
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Paralelo"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   146
            Tag             =   "3"
            Top             =   600
            Width           =   945
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Normal"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   145
            Tag             =   "1"
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.Frame FraUit 
         Caption         =   "Valor U.I.T."
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   4920
         TabIndex        =   142
         Top             =   5520
         Visible         =   0   'False
         Width           =   1185
         Begin VB.TextBox txtUit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   143
            Text            =   "0.00"
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Frame fraDiasAtr2 
         Caption         =   "Dias Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   900
         Left            =   60
         TabIndex        =   32
         Top             =   1005
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtDiaAtrIni 
            Height          =   315
            Left            =   255
            TabIndex        =   34
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TxtDiasAtrFin 
            Height          =   315
            Left            =   930
            TabIndex        =   33
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton CmdGastos 
         Caption         =   "Gastos"
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
         Left            =   3240
         TabIndex        =   141
         Top             =   5640
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Frame FraTipCambio 
         Caption         =   "Tipo Cambio"
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   1920
         TabIndex        =   138
         Top             =   5520
         Visible         =   0   'False
         Width           =   1185
         Begin VB.TextBox TxtTipCambio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   139
            Text            =   "0.00"
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Frame FraOrdenCreditos 
         Caption         =   "Orden Creditos"
         ForeColor       =   &H00800000&
         Height          =   900
         Left            =   90
         TabIndex        =   130
         Top             =   5835
         Width           =   1500
         Begin VB.OptionButton OptNuevos 
            Caption         =   "Nuevos"
            Height          =   255
            Left            =   180
            TabIndex        =   132
            Top             =   510
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptAntiguos 
            Caption         =   "Antiguos"
            Height          =   255
            Left            =   180
            TabIndex        =   131
            Top             =   255
            Width           =   1455
         End
      End
      Begin VB.Frame fraCalif 
         Caption         =   "Calificaciones"
         Height          =   1665
         Left            =   0
         TabIndex        =   136
         Top             =   4920
         Visible         =   0   'False
         Width           =   1665
         Begin VB.ListBox lstCalif 
            Height          =   1185
            ItemData        =   "frmCredReportes.frx":10D4
            Left            =   150
            List            =   "frmCredReportes.frx":10E7
            Style           =   1  'Checkbox
            TabIndex        =   137
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.CommandButton CmdInstitucion 
         Caption         =   "&Instituciones"
         Height          =   450
         Left            =   165
         TabIndex        =   50
         Top             =   4350
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdUbicacion 
         Caption         =   "&Ubic. Geografica"
         Height          =   420
         Left            =   105
         TabIndex        =   43
         Top             =   4335
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame FrmProtesto 
         Caption         =   "Protesto"
         ForeColor       =   &H00800000&
         Height          =   885
         Left            =   45
         TabIndex        =   114
         Top             =   5850
         Visible         =   0   'False
         Width           =   1515
         Begin VB.OptionButton OptAmbos 
            Caption         =   "Ambos"
            Height          =   255
            Left            =   60
            TabIndex        =   117
            Top             =   600
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptProtesto 
            Caption         =   "Con Protesto"
            Height          =   255
            Left            =   60
            TabIndex        =   116
            Top             =   390
            Width           =   1215
         End
         Begin VB.OptionButton OptSinProtesto 
            Caption         =   "Sin Protesto"
            Height          =   255
            Left            =   60
            TabIndex        =   115
            Top             =   180
            Width           =   1215
         End
      End
      Begin MSComCtl2.Animation Logo 
         Height          =   645
         Left            =   5175
         TabIndex        =   110
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1138
         _Version        =   393216
         FullWidth       =   45
         FullHeight      =   43
      End
      Begin VB.Frame fraDatosNota 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   120
         TabIndex        =   35
         Top             =   495
         Visible         =   0   'False
         Width           =   4860
         Begin VB.TextBox TxtNotaFin 
            Height          =   285
            Left            =   3135
            TabIndex        =   42
            Top             =   0
            Width           =   345
         End
         Begin VB.TextBox TxtNotaIni 
            Height          =   285
            Left            =   2490
            TabIndex        =   40
            Top             =   0
            Width           =   390
         End
         Begin VB.CheckBox ChkPorc 
            Alignment       =   1  'Right Justify
            Caption         =   "Por Porcentaje"
            Height          =   210
            Left            =   3495
            TabIndex        =   38
            Top             =   0
            Width           =   1365
         End
         Begin VB.TextBox TxtCuotasPend 
            Height          =   300
            Left            =   1455
            TabIndex        =   37
            Top             =   0
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "Al"
            Height          =   255
            Left            =   2925
            TabIndex        =   41
            Top             =   0
            Width           =   210
         End
         Begin VB.Label Label11 
            Caption         =   "Notas :"
            Height          =   240
            Left            =   1935
            TabIndex        =   39
            Top             =   0
            Width           =   525
         End
         Begin VB.Label Label10 
            Caption         =   "Cuotas Pendiente :"
            Height          =   240
            Left            =   75
            TabIndex        =   36
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.Frame fraProductos 
         Caption         =   "Tipo de crédito"
         ForeColor       =   &H00800000&
         Height          =   4455
         Left            =   1890
         TabIndex        =   84
         Top             =   1005
         Visible         =   0   'False
         Width           =   4050
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4140
            Left            =   60
            TabIndex        =   85
            Top             =   195
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   7303
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            ImageList       =   "imglstFiguras"
            Appearance      =   1
         End
      End
      Begin VB.Frame fraReporte 
         Caption         =   "Tipo de Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1020
         Left            =   45
         TabIndex        =   68
         Top             =   5820
         Visible         =   0   'False
         Width           =   1590
         Begin VB.OptionButton optReporte 
            Caption         =   "Por Credito"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   70
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Por Analista"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   71
            Top             =   600
            Width           =   1230
         End
      End
      Begin MSMask.MaskEdBox TxtFecIniA02 
         Height          =   300
         Left            =   1245
         TabIndex        =   6
         Top             =   495
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecFinA02 
         Height          =   300
         Left            =   3840
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame fraMontoMayor 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   750
         Left            =   435
         TabIndex        =   82
         Top             =   5790
         Visible         =   0   'False
         Width           =   1350
         Begin VB.TextBox txtMontoMayor 
            Height          =   330
            Left            =   90
            TabIndex        =   83
            Top             =   285
            Width           =   1140
         End
      End
      Begin VB.Frame fraCredVig 
         Caption         =   "Crédito Vigente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1140
         Left            =   60
         TabIndex        =   73
         Top             =   1545
         Visible         =   0   'False
         Width           =   1710
         Begin VB.OptionButton optCredVig 
            Caption         =   "Ord. por Cod."
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   74
            Top             =   300
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Analista Esp."
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   76
            Top             =   825
            Width           =   1470
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Agrup. Analista"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   75
            Top             =   562
            Width           =   1470
         End
      End
      Begin VB.Frame FraA02 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   885
         Index           =   3
         Left            =   30
         TabIndex        =   10
         Top             =   1890
         Visible         =   0   'False
         Width           =   1755
         Begin VB.OptionButton OptSaldo 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptSaldo 
            Caption         =   "Con Saldos"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Top             =   495
            Width           =   1455
         End
      End
      Begin VB.Frame fraEstadistica 
         Caption         =   "Est. Mensual por "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1560
         Left            =   60
         TabIndex        =   77
         Top             =   2370
         Visible         =   0   'False
         Width           =   1740
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   30
            Left            =   60
            TabIndex        =   129
            Top             =   1530
            Width           =   1725
         End
         Begin VB.TextBox txtLineaCredito 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   105
            TabIndex        =   81
            Top             =   1125
            Width           =   1500
         End
         Begin VB.OptionButton optEstadistica 
            Caption         =   "Periodo"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   80
            Top             =   285
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton optEstadistica 
            Caption         =   "L.C. Específica"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   78
            Top             =   825
            Width           =   1530
         End
         Begin VB.OptionButton optEstadistica 
            Caption         =   "Líneas de Crédito"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   79
            Top             =   547
            Width           =   1590
         End
      End
      Begin VB.CheckBox ChkDia 
         Caption         =   "DiaFecha"
         Height          =   345
         Left            =   225
         TabIndex        =   134
         Top             =   5370
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Frame FraIncluirMora 
         Height          =   495
         Left            =   165
         TabIndex        =   48
         Top             =   4770
         Visible         =   0   'False
         Width           =   1725
         Begin VB.CheckBox ChkIncluirMora 
            Caption         =   "Incluir Mora"
            Height          =   195
            Left            =   150
            TabIndex        =   49
            Top             =   195
            Width           =   1275
         End
      End
      Begin VB.Frame FraA02 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   780
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   4980
         Visible         =   0   'False
         Width           =   1665
         Begin VB.CheckBox ChkMonA02 
            Caption         =   "Soles"
            Height          =   210
            Index           =   0
            Left            =   330
            TabIndex        =   5
            Top             =   240
            Width           =   915
         End
         Begin VB.CheckBox ChkMonA02 
            Caption         =   "Dolares"
            Height          =   210
            Index           =   1
            Left            =   330
            TabIndex        =   4
            Top             =   480
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filtrar >>>"
         Height          =   675
         Left            =   1920
         TabIndex        =   113
         Top             =   6120
         Width           =   4020
         Begin VB.CommandButton CmdCampanhas 
            Caption         =   "Campañas"
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
            Left            =   2640
            TabIndex        =   140
            Top             =   225
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CommandButton CmdSelecAge 
            Caption         =   "&Agencias"
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
            Left            =   30
            TabIndex        =   67
            Top             =   225
            Width           =   1185
         End
         Begin VB.CommandButton CmdAnalistas 
            Caption         =   "&Analistas"
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
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.Frame FraDiasAtr 
         Caption         =   "Dias Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1680
         Left            =   60
         TabIndex        =   13
         Top             =   1005
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtCar4I 
            Height          =   285
            Left            =   1125
            TabIndex        =   27
            Text            =   "30"
            Top             =   1260
            Width           =   330
         End
         Begin VB.TextBox TxtCar3F 
            Height          =   285
            Left            =   1125
            TabIndex        =   25
            Text            =   "30"
            Top             =   945
            Width           =   330
         End
         Begin VB.TextBox TxtCar3I 
            Height          =   285
            Left            =   480
            TabIndex        =   23
            Text            =   "16"
            Top             =   945
            Width           =   330
         End
         Begin VB.TextBox TxtCar2F 
            Height          =   285
            Left            =   1110
            TabIndex        =   21
            Text            =   "15"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtCar2I 
            Height          =   285
            Left            =   465
            TabIndex        =   19
            Text            =   "8"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtCar1F 
            Height          =   285
            Left            =   1110
            TabIndex        =   17
            Text            =   "7"
            Top             =   255
            Width           =   330
         End
         Begin VB.TextBox TxtCar1I 
            Height          =   285
            Left            =   465
            TabIndex        =   15
            Text            =   "1"
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Mayor De :"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   1305
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   "A"
            Height          =   255
            Left            =   900
            TabIndex        =   24
            Top             =   975
            Width           =   150
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   165
            TabIndex        =   22
            Top             =   975
            Width           =   210
         End
         Begin VB.Label Label6 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   20
            Top             =   630
            Width           =   150
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   630
            Width           =   210
         End
         Begin VB.Label Label4 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   16
            Top             =   285
            Width           =   150
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   285
            Width           =   210
         End
      End
      Begin VB.Frame FraPagCheque 
         Caption         =   "Pago Con Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1230
         Left            =   60
         TabIndex        =   62
         Top             =   1545
         Visible         =   0   'False
         Width           =   1680
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   300
            Left            =   360
            TabIndex        =   65
            Top             =   825
            Width           =   1230
         End
         Begin VB.OptionButton OptPagCheque 
            Caption         =   "Nro Cheque"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   64
            Top             =   570
            Width           =   1215
         End
         Begin VB.OptionButton OptPagCheque 
            Caption         =   "General"
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   63
            Top             =   300
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame fraCredxInstOrden 
         Caption         =   "Ordenar Por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1260
         Left            =   60
         TabIndex        =   44
         Top             =   2835
         Visible         =   0   'False
         Width           =   1740
         Begin VB.OptionButton OptOrdenPagare 
            Caption         =   "&Pagare"
            Height          =   210
            Left            =   135
            TabIndex        =   47
            Top             =   915
            Width           =   1485
         End
         Begin VB.OptionButton OptOrdenAlfabetico 
            Caption         =   "Orden &Alfabetico"
            Height          =   210
            Left            =   135
            TabIndex        =   46
            Top             =   600
            Width           =   1665
         End
         Begin VB.OptionButton OptOrdenCodMod 
            Caption         =   "Codigo &Modular"
            Height          =   210
            Left            =   135
            TabIndex        =   45
            Top             =   300
            Value           =   -1  'True
            Width           =   1590
         End
      End
      Begin VB.Frame FraMoraAnt 
         Height          =   540
         Left            =   60
         TabIndex        =   28
         Top             =   4335
         Visible         =   0   'False
         Width           =   1650
         Begin VB.CheckBox ChkMoraAnt 
            Caption         =   "Mora Anterior"
            Height          =   195
            Left            =   105
            TabIndex        =   29
            Top             =   195
            Width           =   1425
         End
      End
      Begin VB.Frame FraDiasAtrConsumo 
         Caption         =   "Dias Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1290
         Left            =   60
         TabIndex        =   51
         Top             =   1005
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtDiasAtrCons3Ini 
            Height          =   285
            Left            =   1125
            TabIndex        =   56
            Text            =   "30"
            Top             =   915
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons2Fin 
            Height          =   285
            Left            =   1110
            TabIndex        =   55
            Text            =   "15"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons2Ini 
            Height          =   285
            Left            =   465
            TabIndex        =   54
            Text            =   "8"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons1Fin 
            Height          =   285
            Left            =   1110
            TabIndex        =   53
            Text            =   "7"
            Top             =   255
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons1Ini 
            Height          =   285
            Left            =   465
            TabIndex        =   52
            Text            =   "1"
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Mayor De :"
            Height          =   195
            Left            =   195
            TabIndex        =   61
            Top             =   960
            Width           =   780
         End
         Begin VB.Label Label16 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   60
            Top             =   630
            Width           =   150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   630
            Width           =   210
         End
         Begin VB.Label Label18 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   58
            Top             =   285
            Width           =   150
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   57
            Top             =   285
            Width           =   210
         End
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000C0&
         Height          =   60
         Left            =   135
         Top             =   885
         Width           =   4845
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "REPORTES DE CREDITOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   4845
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Left            =   75
         Top             =   225
         Width           =   4950
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Left            =   2760
         TabIndex        =   9
         Top             =   525
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   510
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   630
      SizeMode        =   1  'Stretch
      TabIndex        =   66
      Top             =   1005
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   9315
      TabIndex        =   89
      Top             =   6135
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "1=Existe filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   9930
      TabIndex        =   88
      Top             =   6180
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "frmCredReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MatAgencias() As String
Private MatProductos() As String
Private MatCondicion() As String
Private matAnalista() As String
Private MatInstitucion() As String
Private sUbicacionGeo As String

Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lbLibroOpen As Boolean
 
'ALPA 20080829*************************
Dim sMatrixPosiciones() As String
Dim sMatrixGarantiasAdjud() As String
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
'**************************************


Dim WithEvents loRep As NCredReporte
Attribute loRep.VB_VarHelpID = -1
Dim WithEvents lsRep As nCredRepoFinMes
Attribute lsRep.VB_VarHelpID = -1

Dim loRepFM As nCredRepoFinMes

Dim Progreso As clsProgressBar
Dim Progress As clsProgressBar
Dim vTempo As Boolean
Dim bDiaHora As Boolean

Dim nmoneda As Integer

Dim WithEvents oNCredDoc As NCredDoc
Attribute oNCredDoc.VB_VarHelpID = -1

Dim Index108000 As Integer ' *** MAVM: Auditoria
Dim Index108300 As Integer ' *** MAVM: Auditoria
Dim Index108380 As Integer ' *** MAVM: Auditoria
Dim Index108386 As Integer ' *** MAVM: Auditoria

Dim Index108000OR As Integer ' *** MAVM: Auditoria
Dim Index108300OR As Integer ' *** MAVM: Auditoria
Dim Index108325OR As Integer ' *** MAVM: Auditoria

Dim Index108000CR As Integer ' *** MAVM: Auditoria
Dim Index108100CR As Integer ' *** MAVM: Auditoria
Dim Index108140CR As Integer ' *** MAVM: Auditoria
Dim Index108142CR As Integer ' *** MAVM: Auditoria

Dim Index108000AR As Integer ' *** MAVM: Auditoria
Dim Index108200AR As Integer ' *** MAVM: Auditoria
Dim Index108203AR As Integer ' *** MAVM: Auditoria


'ALPA 20081013*********************************
Dim lsTitProductos As String 'DAOR 20070717
'**********************************************
Private Function DameAgencias() As String
Dim Agencias As String
Dim lnAge As Integer
Dim est As Integer
est = 0
Agencias = ""
For lnAge = 1 To frmSelectAgencias.List1.ListCount
 If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
   est = est + 1
   If est = 1 Then
    Agencias = "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
   Else
    Agencias = Agencias & ", " & "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
   End If
 End If
Next lnAge
DameAgencias = Agencias
End Function

Public Function DescProdConsumoSeleccionado() As String
Dim lsProductos As String
Dim i As Integer
lsProductos = "PRODUCTOS : "
  
  'Cambio Pepe 12
'  For i = 0 To Me.chkConsumo.Count - 1
'    If chkConsumo(i).value Then
'        lsProductos = lsProductos & "/CON-" & Mid(chkConsumo(i).Caption, 1, 5)
'    End If
'  Next i
  
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" And Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(Producto.gColConsuDctoPlan, 1, 1) Then
                lsProductos = lsProductos & "/CON-" & Mid(TreeView1.Nodes(i).Text, 1, 3)
            End If
        End If
    Next
  
  'Fin Cambio Pepe 12
DescProdConsumoSeleccionado = lsProductos

End Function
Public Function ValorProdConsumo() As String
Dim i As Integer
Dim lsCad As String

    lsCad = ""
    'Cambio Pepe 07
'    For i = 0 To Me.chkConsumo.Count - 1
'        If chkConsumo(i).value Then
'            lsCad = lsCad & "'" & chkConsumo(i).Tag & "',"
'        End If
'    Next i
'
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" And Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(Producto.gColConsuDctoPlan, 1, 1) Then
                If Len(Trim(lsCad)) = 0 Then
                    lsCad = "'" & Mid(TreeView1.Nodes(i).Key, 2, 3) & "'"
                Else
                    lsCad = lsCad & ", '" & Mid(TreeView1.Nodes(i).Key, 2, 3) & "'"
                End If
            End If
        End If
    Next
    
    'Fin Cambio Pepe 07
    
    If Len(lsCad) > 0 Then
        'Cambio Pepe 08
        'lsCad = Mid(lsCad, 1, (Len(lsCad) - 1))
        'Fin Cambio Pepe 08
        ValorProdConsumo = " AND substring(Credito.cCtaCod,6,3) IN (" & lsCad & ") "
    Else
        ValorProdConsumo = " AND substring(Credito.cCtaCod,6,1) = '3' "
    End If
End Function
Private Function ValorProducto() As String
Dim i As Integer
Dim lsCad As String

lsCad = ""

    'Cambio Pepe 10
    
    '' Cred. Comercial
    'For i = 0 To chkComercial.Count - 1
    '    If chkComercial(i).value Then
    '        lsCad = lsCad & "'" & chkComercial(i).Tag & "',"
    '    End If
    'Next i
    '
    '' Cred. MicroEmpresarial
    'For i = 0 To chkMicroEmpresa.Count - 1
    '    If chkMicroEmpresa(i).value Then
    '        lsCad = lsCad & "'" & chkMicroEmpresa(i).Tag & "',"
    '    End If
    'Next i
    '
    '' Cred. Consumo
    'For i = 0 To chkConsumo.Count - 1
    '    If chkConsumo(i).value Then
    '        lsCad = lsCad & "'" & chkConsumo(i).Tag & "',"
    '    End If
    'Next i
    '
    ' '  Cred. Hipotecario
    'For i = 0 To chkHipotecario.Count - 1
    '    If chkHipotecario(i).value Then
    '        lsCad = lsCad & "'" & chkHipotecario(i).Tag & "',"
    '    End If
    'Next i

    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                If Len(Trim(lsCad)) = 0 Then
                    lsCad = "'" & Mid(TreeView1.Nodes(i).Key, 2, 3) & "'"
                Else
                    lsCad = lsCad & ", '" & Mid(TreeView1.Nodes(i).Key, 2, 3) & "'"
                End If
            End If
        End If
    Next
    'Fin Cambio Pepe 10
    
If Len(lsCad) > 0 Then
    ValorProducto = " AND substring(Credito.cCtaCod,6,3) IN (" & lsCad & ") "
Else
    ValorProducto = ""
End If
End Function

Private Function DescProdSeleccionado() As String
Dim lsProductos As String
Dim i As Integer
lsProductos = "PRODUCTOS : "
   
  'Cambio Pepe 09
  
'  For i = 0 To chkComercial.Count - 1
'    If chkComercial(i).value Then
'        lsProductos = lsProductos & "/MES-" & Mid(chkComercial(i).Caption, 1, 3)
'    End If
'  Next i
'  For i = 0 To chkMicroEmpresa.Count - 1
'    If chkMicroEmpresa(i).value Then
'        lsProductos = lsProductos & "/MES-" & Mid(chkMicroEmpresa(i).Caption, 1, 3)
'    End If
'  Next i
'  For i = 0 To chkConsumo.Count - 1
'    If chkConsumo(i).value Then
'        lsProductos = lsProductos & "/MES-" & Mid(chkConsumo(i).Caption, 1, 3)
'    End If
'  Next i
'  For i = 0 To chkHipotecario.Count - 1
'    If chkHipotecario(i).value Then
'        lsProductos = lsProductos & "/MES-" & Mid(chkHipotecario(i).Caption, 1, 4)
'    End If
'  Next i
  
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                lsProductos = lsProductos & "/MES" & Mid(TreeView1.Nodes(i).Text, 1, 3)
            End If
        End If
    Next
  
  'Fin Cambio Pepe 09
  
  
DescProdSeleccionado = lsProductos
End Function


Private Function ValorNorRefPar() As String
    ValorNorRefPar = ""
    '**************** CREDITOS NORMALES
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 0 Then
        ValorNorRefPar = " AND (Credito.cRefinan = 'N' and nCondCre in (1,3,2)) "
    End If
    '**************** CREDITOS PARALELOS
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 0 Then
        ValorNorRefPar = " AND (Credito.cRefinan = 'N' and nCondCre = 1) "
    End If
    '*************** CREDITOS REFINANCIADOS
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " AND Credito.cRefinan = 'R' "
    End If
    '*************** CREDITOS NORMALES Y PARALELOS
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 0 Then
        ValorNorRefPar = " AND Credito.cRefinan = 'N' "
    End If
    '*************** CREDITOS NORMALES Y REFINANCIADOS
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " AND Credito.nCondCre <> 3 "
    End If
    '*************** CREDITOS PARALELOS Y REFINANCIADOS
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " AND Credito.nCondCre = 3 OR Credito.cRefinan = 'R' "
    End If
    '*************** TODOS LOS CREDITOS
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " "
    End If
End Function
Public Function DescCondSeleccionado() As String
Dim lsCondic As String
    lsCondic = " %%% CONDICION : "
    If Me.ChkCond(0).value = 1 Then  'Normal
        lsCondic = lsCondic & "Norm"
    End If
    If Me.ChkCond(1).value = 1 Then 'Paralelo
        lsCondic = lsCondic & "/Paral"
    End If
    If Me.ChkCond(2).value = 1 Then  'Refinanciado
        lsCondic = lsCondic & "/Refin"
    End If
'ARCV 12-06-2007
    If Me.ChkCond(3).value = 1 Then  'Recurrente
        lsCondic = lsCondic & "/Recurr"
    End If
    If Me.ChkCond(4).value = 1 Then  'Ampliado
        lsCondic = lsCondic & "/Amplia"
    End If
    If Me.ChkCond(5).value = 1 Then  'Automatico
        lsCondic = lsCondic & "/Automa"
    End If
    'JUEZ 20130604 *********************
    If Me.ChkCond(6).value = 1 Then  'Adicional
        lsCondic = lsCondic & "/Adicional"
    End If
    'END JUEZ **************************
'---------
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 0 Then
        lsCondic = lsCondic & "Norm/Paral/Refin"
    End If
DescCondSeleccionado = lsCondic
End Function

Private Function ValorMoneda() As String

ValorMoneda = ""
If ChkMonA02(0).value = 1 And ChkMonA02(1).value = 0 Then
    ValorMoneda = " AND SUBSTRING(Credito.cCtaCod, 9,1) = '1' "
End If
If ChkMonA02(0).value = 0 And ChkMonA02(1).value = 1 Then
    ValorMoneda = " AND SUBSTRING(Credito.cCtaCod, 9,1) = '2' "
End If
If ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1 Then
    ValorMoneda = " AND SUBSTRING(Credito.cCtaCod, 9,1) in ('1', '2') "
End If

End Function
  

Private Sub CmdCampanhas_Click()
    frmSelectCampanhas.SeleccionaCampanhas
End Sub

Private Sub CmdGastos_Click()
    frmSelectCampanhas.SeleccionaGastos
End Sub

Private Sub loRep_CloseProgress()
    Progress.CloseForm Me
End Sub


Private Sub loRep_ShowProgress()
    Progress.ShowForm Me
End Sub
 
Public Sub inicia(ByVal sCaption As String)
 
    Me.Caption = sCaption
    LlenaArbol
    vTempo = True
    
    LlenaProductos
    Me.Show 0, MDISicmact
    
End Sub

Private Sub LlenaArbol()
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim lsTipREP As String
Dim i As Integer ' *** MAVM Auditoria

    lsTipREP = "108"
    
    Set clsGen = New DGeneral
    'ARCV 20-07-2006
    'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
 
    'JUEZ 20141029 Verifica si tiene permisos *****************
    If gRsOpeRepo.EOF Or gRsOpeRepo.BOF Then
        MsgBox "No tiene permisos para generar reportes", vbInformation, "Aviso"
        Exit Sub
    End If
    'END JUEZ *************************************************
    Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)

    '------------------------
    Set clsGen = Nothing
      
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("cOpeCod")
        sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
        Select Case rsUsu("nOpeNiv")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = TVRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        i = i + 1 ' *** MAVM Auditoria
        If sOpeCod = "108000" Then Index108000 = i ' *** MAVM Auditoria
        If sOpeCod = "108300" Then Index108300 = i ' *** MAVM Auditoria
        If sOpeCod = "108380" Then Index108380 = i ' *** MAVM Auditoria
        If sOpeCod = "108386" Then Index108386 = i ' *** MAVM Auditoria
        
        If sOpeCod = "108300" Then Index108000OR = i ' *** MAVM Auditoria
        If sOpeCod = "108300" Then Index108300OR = i ' *** MAVM Auditoria
        If sOpeCod = "108325" Then Index108325OR = i ' *** MAVM Auditoria
        
        If sOpeCod = "108000" Then Index108000CR = i ' *** MAVM Auditoria Reporte de Creditos Rechazados
        If sOpeCod = "108100" Then Index108100CR = i ' *** MAVM Auditoria
        If sOpeCod = "108140" Then Index108140CR = i ' *** MAVM Auditoria
        If sOpeCod = "108142" Then Index108142CR = i ' *** MAVM Auditoria
        
        If sOpeCod = "108000" Then Index108000AR = i ' *** MAVM Auditoria Reporte de Arqueos
        If sOpeCod = "108200" Then Index108200AR = i ' *** MAVM Auditoria
        If sOpeCod = "108203" Then Index108203AR = i ' *** MAVM Auditoria
        
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
End Sub

Private Sub LlenaProductos()
Dim rs As ADODB.Recordset
Dim oreg As New DCredReporte
Dim sOpePadre As String
Dim sOpeHijo As String
Dim nodOpe As Node
TreeView1.Nodes.Clear
Set rs = New ADODB.Recordset

Set rs = oreg.GetProductos

Do While Not rs.EOF
          
        Select Case rs!cNivel
            Case "1"
                sOpePadre = "P" & rs!cValor
                Set nodOpe = TreeView1.Nodes.Add(, , sOpePadre, rs!cProducto, "Padre")
                nodOpe.Tag = rs!cValor
            Case "2"
                sOpeHijo = "H" & rs!cValor
                Set nodOpe = TreeView1.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, rs!cProducto, "Hijo")
                nodOpe.Tag = rs!cValor
        
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

' peac 20070913 agrego Optional pbCampanhas As Boolean = False

Private Sub HabilitaControleFrame1(ByVal pbTxtFecIni As Boolean, ByVal pbTxtFecFin As Boolean, _
        pbFraMoneda As Boolean, ByVal pbFraSaldos As Boolean, _
        Optional ByVal pbFraDiasAtraso As Boolean = False, Optional pbFraCondicion As Boolean = False, _
        Optional ByVal pbFraMoraAnt As Boolean = False, Optional pbAnalistas As Boolean = False, _
        Optional pbFraDiasAtr2 As Boolean = False, Optional ByVal pbFraDatosNota As Boolean = False, _
        Optional ByVal pbCmdUbicacion As Boolean = False, Optional ByVal pbTipCambio As Boolean = False, _
        Optional ByVal pbfraCredxInstOrden As Boolean = False, Optional ByVal pbFraIncluirMora As Boolean = False, _
        Optional ByVal pbCmdInstitucion As Boolean = False, Optional ByVal pbfradiasatrconsumo As Boolean = False, _
        Optional ByVal pbSoloPrdConsumo As Boolean = False, Optional pbFraPagCheque As Boolean = False, Optional pbFraProductos As Boolean = False, _
        Optional pbFraReporte As Boolean = False, Optional pbcmdAge As Boolean = True, _
        Optional pbFraCredVig As Boolean = False, Optional pbFraEstadistica As Boolean = False, Optional pbFraMontoMayor As Boolean = False, _
        Optional pbFraProtesto As Boolean = False, Optional ByVal pbChkRFA As Boolean = False, _
        Optional pbFraAutomatico As Boolean = False, Optional ByVal pbGeografico As Boolean = False, _
        Optional pbFraMonto As Boolean = False, Optional pbOrdenCreditos As Boolean = False, Optional pbMigraExcel As Boolean = False, _
        Optional ByVal pbDiaFecha As Boolean = False, _
        Optional ByVal pbFraCalificaciones As Boolean = False, _
        Optional pbCampanhas As Boolean = False, _
        Optional pbGastos As Boolean = False, Optional ByVal pbUit As Boolean = False, _
        Optional ByVal pbSoloBN As Boolean = False, Optional pbFraCondBN As Boolean = False, Optional pbFraEstado As Boolean = False, Optional pbFraTipoCred As Boolean = False, _
        Optional ByVal pbfrmResAnalista As Boolean = False, Optional ByVal pbfrmACE As Boolean = False) 'MAVM 20100512
        
        ChkRFA.Visible = pbChkRFA
        CmdSelecAge.Visible = pbcmdAge
        FraA02(3).Visible = pbFraSaldos
        TxtFecFinA02.Visible = pbTxtFecFin
        Label3.Visible = pbTxtFecFin
        TxtFecIniA02.Visible = pbTxtFecIni
        Label2.Visible = pbTxtFecIni
        'FraA02(2).Visible = pbFraProd
        FraA02(1).Visible = pbFraMoneda
        FraDiasAtr.Visible = pbFraDiasAtraso
        FraCondicion.Visible = pbFraCondicion
        FraMoraAnt.Visible = pbFraMoraAnt
        CmdAnalistas.Visible = pbAnalistas
        CmdCampanhas.Visible = pbCampanhas ' peac 20070913
        CmdGastos.Visible = pbGastos ' peac 20071009
        fraDiasAtr2.Visible = pbFraDiasAtr2
        fraDatosNota.Visible = pbFraDatosNota
        CmdUbicacion.Visible = pbCmdUbicacion
        FraTipCambio.Visible = pbTipCambio
        fraCredxInstOrden.Visible = pbfraCredxInstOrden
        FraIncluirMora.Visible = pbFraIncluirMora
        CmdInstitucion.Visible = pbCmdInstitucion
        FraDiasAtrConsumo.Visible = pbfradiasatrconsumo
        fraEstadistica.Visible = pbFraEstadistica
        FraUit.Visible = pbUit  'PEAC 20080219
        
        fraMontoMayor.Visible = pbFraMontoMayor
        
        If pbSoloPrdConsumo = True Then
            'ActFiltra True, Mid(Producto.gColConsuDctoPlan, 1, 1)
            'ActFiltra True, Mid(650, 1, 1)
            ActFiltra True, Mid(751, 1, 1)
        Else
            ActFiltra False
        End If
        
        'FrmProtesto.Visible = pbFraProtesto
        
        FraPagCheque.Visible = pbFraPagCheque
        fraProductos.Visible = pbFraProductos
        fraReporte.Visible = pbFraReporte
        fraCredVig.Visible = pbFraCredVig
        FraAutomaticos.Visible = pbFraAutomatico
        FraConUbiGeo.Visible = pbGeografico
        FrmMontos.Visible = pbFraMonto
        FraOrdenCreditos.Visible = pbOrdenCreditos
        chkMigraExcell.Visible = pbMigraExcel
        ChkDia.Visible = pbDiaFecha
        chkCondBN.Visible = pbSoloBN      'Gitu 03-04-2008
        fraCondBN.Visible = pbFraCondBN   'Gitu 03-04-2008
        
        fraEstados.Visible = pbFraEstado 'MAVM 20091001
        
        fratipo.Visible = pbFraTipoCred 'MADM 20091117
        FraACE.Visible = pbfrmACE 'MADM 20110329
        FraTipAce.Visible = pbfrmACE 'MADM 20110329
        
        OptSaldo(1).Visible = False 'ARCV 27-07-2006
        
        fraCalif.Visible = pbFraCalificaciones  'ARCV 02-02-2007
        frmResAnalista.Visible = pbfrmResAnalista 'MAVM 20100512
        
End Sub
 
Private Sub CmdAnalistas_Click()
    frmSelectAnalistas.SeleccionaAnalistas
End Sub
Private Sub CmdImprimirA02_Click()
Dim i As Integer
Dim nContAge As Integer
Dim nContEstados As Integer 'MAVM 20091001
Dim oNComCre As COMNCredito.NCOMCredito
Set oNComCre = New COMNCredito.NCOMCredito
Dim sCadProd As String
sCadProd = ""
Dim P As previo.clsprevio

Dim sCadImp As String, nValTmp As Integer, dUltimoDia As DCredReporte, nUltimoDia As Integer, sProductos As String
Dim sMoneda As String, sCondicion As String

Dim nContGas As Integer, xCondicion As String, nContCam As Integer ' peac 20070913

Dim sTempo As Integer, lsArchivoN As String, lbLibroOpen As Boolean, sAnalistas As String, nContAna As Integer, sAnalistas2 As String
Dim nContAgencias As Integer, sAgencias As String, nTempoParam As Byte
'--------------------
Dim lsCadenaPar As String, CredRepoMEs As nCredRepoFinMes
Set CredRepoMEs = New nCredRepoFinMes
Dim FMes As Date
Dim fechaini As String, lsCadenaDesPar As String, lsServerCons As String, Rcd As nRCDProceso
Set Rcd = New nRCDProceso
Set lsRep = New nCredRepoFinMes

Dim TipoCambio As Currency, fnRepoSelec As Long, sCadAge As String, sCadMoneda As String

Dim nProtesto As Integer, nProductoAutom As Integer, ntipoCred As Integer, bFecha As Boolean

'******* agregado por LMMD **********
Dim sParam1 As String, sParam2 As String, sParam3 As String, sParam4 As String, sParam5 As String
Dim sParam6 As String, sParam7 As String, sParam8 As String, sParam9 As String, sParam10 As String
'''''''''''''''''''''''''''''''''''''
'Comentado por ALPA 20081013, para declararlo como global*************************
'Dim lsTitProductos As String 'DAOR 20070717
'***************************************************
Dim lnPosI As Integer 'DAOR 20070717
Dim oCOMNCredDoc As COMNCredito.NCOMCredDoc 'DAOR 20070816

'By Capi Set 07
Dim nMES As Integer, nAno As Integer, sOperacion As String

'End By

'ALPA 20080826
Dim sAgenciasTemp As String
'END ALPA

'Cambio Pepe 12
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion

Dim rCargaRuta As ADODB.Recordset

Set rCargaRuta = New ADODB.Recordset

Set rCargaRuta = oCon.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
Call lsRep.inicia(gdFecSis, gsNomCmac, gsNomAge, gsCodUser)
If rCargaRuta.BOF Then
Else
    'sServidorConsolidada = rCargaRuta!nConsSisValor
    lsServerCons = rCargaRuta!nConsSisValor
End If
Set rCargaRuta = Nothing

oCon.CierraConexion

'''Fin Cambio Pepe 12

Dim oTipCambio As nTipoCambio
'sUbicacionGeo = ""

    If CmdUbicacion.Visible And Me.ChkUbi.Visible = False Then
        If Trim(sUbicacionGeo) = "" Then
            MsgBox "Seleccione una Ubicacion Geografica", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
                                
    If TxtFecIniA02.Visible = True Then
        If IsDate(TxtFecIniA02.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            TxtFecIniA02.SetFocus
            Exit Sub
        End If
    End If
    
    If TxtFecFinA02.Visible = True Then
        If IsDate(TxtFecFinA02.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            TxtFecFinA02.SetFocus
            Exit Sub
        End If
    End If
    
    If CmdAnalistas.Visible Then
        ReDim matAnalista(0)
        nContAge = 0
        nContAna = 0
        For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
            If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                nContAge = nContAge + 1
                nContAna = nContAna + 1
                ReDim Preserve matAnalista(nContAge)
                matAnalista(nContAge - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
            End If
        Next i
        If UBound(matAnalista) = 0 Then
            MsgBox "Debe Seleccionar por lo Menos un Analista", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

'peac 20071009
    If CmdGastos.Visible Then
        ReDim MatGastos(0)
        nContAge = 0
        nContGas = 0
        For i = 0 To frmSelectCampanhas.LstCampanha.ListCount - 1
            If frmSelectCampanhas.LstCampanha.Selected(i) = True Then
                nContAge = nContAge + 1
                nContGas = nContGas + 1
                ReDim Preserve MatGastos(nContAge)
                MatGastos(nContAge - 1) = Trim(Right(frmSelectCampanhas.LstCampanha.List(i), 20))
            End If
        Next i
        If UBound(MatGastos) = 0 Then
            MsgBox "Debe Seleccionar por lo Menos un Gasto", vbInformation, "Aviso"
            Exit Sub
        End If
    End If


    'PEAC 20070914
    If CmdCampanhas.Visible Then
        ReDim matCampanha(0)
        nContAge = 0
        nContCam = 0
        For i = 0 To frmSelectCampanhas.LstCampanha.ListCount - 1
            If frmSelectCampanhas.LstCampanha.Selected(i) = True Then
                nContAge = nContAge + 1
                nContCam = nContCam + 1
                ReDim Preserve matCampanha(nContAge)
                matCampanha(nContAge - 1) = Trim(Right(frmSelectCampanhas.LstCampanha.List(i), 20))
            End If
        Next i
        If UBound(matCampanha) = 0 Then
            MsgBox "Debe Seleccionar por lo Menos una Campaña", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
         
         
    '**- PEAC 20071018 - verifica que aya escogido al menos una condicion
    If FraCondicion.Visible = True Then
        Dim Con As Integer, j As Integer
        Con = 0
        'For j = 0 To 5
        For j = 0 To 6 'JUEZ 20130604
            If ChkCond(j).value = 0 Then
                Con = Con + 1
            End If
        Next
        'If Con = 6 Then
        If Con = 7 Then 'JUEZ 20130604
            MsgBox "Seleccione al menos una condición.", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
        
    sTempo = 0
    
    sCadMoneda = ""
    If ChkMonA02(0).value = 1 Then
        sCadMoneda = "'1'"
    End If
    If ChkMonA02(1).value = 1 Then
        If Len(Trim(sCadMoneda)) = 0 Then
            sCadMoneda = "'2'"
        Else
            sCadMoneda = sCadMoneda & ", '2'"
        End If
    End If
    
    If fraDiasAtr2.Visible = True Then
        If IsNumeric(TxtDiaAtrIni.Text) Then
            If IsNumeric(TxtDiasAtrFin.Text) Then
                If val(TxtDiasAtrFin.Text) < val(TxtDiaAtrIni.Text) Then
                    MsgBox "El nro. de dias final no puede ser menor al nro. de dias inicial", vbExclamation, "Aviso"
                    TxtDiasAtrFin.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "Ingrese un nro. de dias válido", vbExclamation, "Aviso"
                TxtDiasAtrCons1Fin.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Ingrese un nro. de dias válido", vbExclamation, "Aviso"
            TxtDiaAtrIni.SetFocus
            Exit Sub
        End If
    End If
        
' ARCV 24-07-2006
    ReDim MatInstitucion(0)
    nContAge = 0
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            ReDim Preserve MatInstitucion(nContAge)
            MatInstitucion(nContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)
        End If
    Next i
'----------------
    ReDim MatAgencias(0)
    nContAge = 0
    sAgenciasTemp = ""
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            nContAgencias = nContAgencias + 1
            ReDim Preserve MatAgencias(nContAge)
            MatAgencias(nContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)

            If Len(Trim(sCadAge)) = 0 Then
                sCadAge = "'" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & "'"
                sAgenciasTemp = "" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ""
            Else
                sCadAge = sCadAge & ", '" & Mid(frmSelectAgencias.List1.List(i), 1, 2) & "'"
                sAgenciasTemp = sAgenciasTemp & ", " & Mid(frmSelectAgencias.List1.List(i), 1, 2) & ""
            End If

        End If
    Next i
    If nContAge = 0 Then
        ReDim MatAgencias(1)
        nContAgencias = 1
        MatAgencias(0) = gsCodAge
    End If

    If CmdInstitucion.Visible Then
        ReDim MatInstitucion(0)
        nContAge = 0
        If frmSelectAnalistas.LstAnalista.ListCount > 0 Then
            For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
                If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                    nContAge = nContAge + 1
                    ReDim Preserve MatInstitucion(nContAge)
                    MatInstitucion(nContAge - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
                End If
            Next i
        End If
        If UBound(MatInstitucion) = 0 Then
            MsgBox "Seleccione una Institucion"
            Exit Sub
        End If
    End If

    lsTitProductos = "" 'DAOR 20070717

    ReDim MatProductos(0)
    nContAge = 0
       
    sCadProd = "0"
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                nContAge = nContAge + 1
                ReDim Preserve MatProductos(nContAge)
                MatProductos(nContAge - 1) = Trim(Mid(TreeView1.Nodes(i).Key, 2, 3))
                sCadProd = sCadProd & "," & MatProductos(nContAge - 1)
                '**DAOR 20070717****
                lnPosI = 0
                lnPosI = InStr(1, TreeView1.Nodes(i).Text, " ")
                If lnPosI > 1 Then
                    lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & " " & Mid(TreeView1.Nodes(i).Text, lnPosI + 1, 3) & "/"
                Else
                    lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & "/"
                End If
                '*******************
            End If
        End If
    Next
    sCadProd = sCadProd & ""
    If Len(lsTitProductos) > 1 Then
        lsTitProductos = Left(lsTitProductos, Len(lsTitProductos) - 1)
    End If

    ReDim MatCondicion(0)
    nContAge = 0
    For i = 0 To ChkCond.Count - 1
        If ChkCond(i).value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatCondicion(nContAge)
            MatCondicion(nContAge - 1) = Trim(ChkCond(i).Tag)
        End If
    Next i
    
    If FraTipCambio.Visible = True Then
        If val(TxtTipCambio.Text) = 0 Then
            MsgBox "Ingrese un tipo de cambio válido " & Chr(13) & "o confirme el que se pondrá por defecto" & Chr(13) & "y que corresponde a la fecha indicada" & Chr(13) & Chr(13) & "Luego presione Imprimir ... ", vbInformation, "Aviso"
            'Saco el tipo de cambio para la fecha que dice solo si el tipo de cambio es vacio
            Set oTipCambio = New nTipoCambio
            TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY"), TCFijoMes), "0.00")
            Set oTipCambio = Nothing
            TxtTipCambio.SetFocus
            Exit Sub
        End If
    End If

    '*** PEAC 20080219
    If FraUit.Visible = True Then
        If val(txtUit.Text) = 0 Then
            txtUit.Text = Format(0, "0.00")
            txtUit.SetFocus
        End If
    End If

    If FrmMontos.Visible = True Then
        If val(TxtMontoIni) = 0 Or val(TxtMontoFin) = 0 Then
            MsgBox "Debe ingresar los rangos de los montos", vbInformation, "AVISO"
            Exit Sub
        End If
        
        If IsNumeric(TxtMontoIni) = False Or IsNumeric(TxtMontoFin) = False Then
            MsgBox "Los montos deben ser formatos numericos", vbInformation, "AVISO"
            Exit Sub
        End If
    End If

    '*** PEAC 20080529
    If TreeView1.Visible = True Then
        If UBound(MatProductos) = 0 Then
                MsgBox "Seleccione por lo menos un producto.", vbInformation, "Aviso"
                Exit Sub
        End If
    End If
    For i = 0 To nContAna - 1
        If i = 0 Then
            sAnalistas2 = "" & matAnalista(i) & ""
        Else
            sAnalistas2 = sAnalistas2 & "," & matAnalista(i) & ""
        End If
    Next
    
    ' MAVM 20091001
    If fraEstados.Visible = True Then
        Dim contador As Integer, m As Integer
        contador = 0
        For m = 0 To 3
            If chkEstado(m).value = 0 Then
                contador = contador + 1
            End If
        Next
        If contador = 4 Then
            MsgBox "Seleccione Al menos una Condición de Credito", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    Dim sEstados As String
    If chkEstado(0).value = 1 Then
        sEstados = gColocEstRecVigJud & "," & gColocEstRecCanJud & "," & "2205"
    End If
    If chkEstado(1).value = 1 Then
        If sEstados = "" Then
            sEstados = gColocEstRecVigCast & "," & gColocEstRecCanJud & "," & "2204"
        Else
            sEstados = sEstados & "," & gColocEstRecVigCast & "," & gColocEstRecCanJud & "," & "2204"
        End If
    End If
    If chkEstado(2).value = 1 Then
        If sEstados = "" Then
            sEstados = gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor
        Else
            sEstados = sEstados & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor
        End If
    End If
    If chkEstado(3).value = 1 Then
        If sEstados = "" Then
            sEstados = gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor
        Else
            sEstados = sEstados & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor
        End If
    End If
'MAVM 20091001
    
''***************************************************************************************************************


    Set oNCredDoc = New NCredDoc
    oNCredDoc.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    Set P = New previo.clsprevio
    
    
    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
    Case 108506 'Reporte de Estadisticas de RFC y DIF
            'FrmCredEstRFCDIF.Show vbModal
    Case 108311 'Reporte de Creditos no Desembolsados
         sCadImp = oNCredDoc.ListaCreditosNoDesembolsados(MatAgencias, TxtFecIniA02, TxtFecFinA02, gdFecSis, gsNomAge, gsCodUser, gsNomCmac)
    Case 108409
         'impre de reporte cancelado por agencia y por institucion
          sCadImp = oNCredDoc.Repor_Imprime108409(ValorMoneda(), TxtFecIniA02, TxtFecFinA02, gsNomAge, gsCodUser, gsNomCmac, MatAgencias, MatInstitucion, gdFecSis)
    Case 108410
         'imprime el reporte de desembolsados por  agencia e institucion
          sCadImp = oNCredDoc.Repor_Impre108410(TxtFecIniA02, TxtFecFinA02, gsNomAge, gsCodUser, gsNomCmac, MatAgencias, MatInstitucion, gdFecSis)
    Case 108307
          'Reporte de Resumen de Comite
          
          Call oNCredDoc.ListaCreditosComite(Me.TxtFecIniA02, gsCodAge)
    Case gColCredRepIngxPagoCred
        If ChkMonA02(0).value = 1 Then
            sCadImp = oNCredDoc.ImprimePagodeCreditos(MatAgencias, CDate(TxtFecIniA02.Text), CDate(Me.TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, ChkRFA.value)
        End If
        If ChkMonA02(1).value = 1 Then
            If Len(sCadImp) > 0 Then sCadImp = sCadImp & Chr(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagodeCreditos(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, ChkRFA.value)
        End If
    'ALPA 20081013********************
    Case gColCredRepDesemEfect
        sCadImp = FColCredRepDesemEfect
    Case gColCredRepSalCarVig
        sCadImp = FColCredRepSalCarVig
    Case gColCredRepCredCancel 'Creditos Cancelados
        sCadImp = FColCredRepCredCancel
    Case gColCredRepResSalCarxAna
         sCadImp = FColCredRepResSalCarxAna
    Case gColCredRepMoraInst
        sCadImp = FColCredRepMoraInst
    Case gColCredRepAtraPagoCuotaLib
        sCadImp = FColCredRepAtraPagoCuotaLib
    Case gColCredRepMoraxAna
        Call ColCredRepMoraxAna(sCadImp)
    '********************************
    Case 108113
        sCadImp = oNCredDoc.ImprimeMoraXAnalitaNewTelefono(MatAgencias, CDate(Me.TxtFecIniA02), gsCodUser, gdFecSis, gsNomAge, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), gsNomCmac)
    Case gColCredRepCredSaldosDiarios
        sCadImp = oNCredDoc.Repo_108118_SaldosDiariosCred(CDate(TxtFecIniA02), gsNomAge, gsCodUser, gsNomCmac, gdFecSis, gsCodAge)
    Case gColCredRepCredProtes
        
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            End If
        End If
    Case 108108

       If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
'        Else
'            If ChkMonA02(0).value = 1 Then
'                sCadImp = oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
'            Else
'                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
'            End If
        End If
    Case gColCredRepCredxUbiGeo
       
        
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            End If
        End If
    
    'Modificado Se agrego una segunda opcion
    Case gColCredRepCredVig, gColCredRepCredVigconCuoLibre

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
            Else
                sCadImp = oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
            End If
        End If
    ''''''''''''''''''''''''''''''''''''
    
        
    Case gColCredRepCredxInst

        If OptOrdenAlfabetico.value Then
            nValTmp = 1
        End If
        If OptOrdenCodMod.value Then
            nValTmp = 0
        End If
        If OptOrdenPagare.value Then
            nValTmp = 2
        End If

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, gsNomCmac, , MatInstitucion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, gsNomCmac, , MatInstitucion)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, gsNomCmac, , MatInstitucion)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, gsNomCmac, , MatInstitucion)
            End If
        End If
    Case gColCredRepMoraxInst
        If OptOrdenAlfabetico.value Then
            nValTmp = 1
        End If
        If OptOrdenCodMod.value Then
            nValTmp = 0
        End If
        If OptOrdenPagare.value Then
            nValTmp = 2
        End If

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatInstitucion, gsNomCmac)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatInstitucion, gsNomCmac)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            End If
        End If
        
    Case 108407 'reporte de pendientes por devolver creditos con convenio
        If Me.chkMigraExcell.value = 0 Then
            sCadImp = oNCredDoc.ImprimePersPendienteDevCredPers(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePersPendienteDevCredPers(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion)
        Else
            sCadImp = oNCredDoc.ImprimePersPendienteDevCredPersExcel(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion)
        End If

    Case 108408 'reporte de creditos devueltos por descuento de creditos con convenio
        If Me.chkMigraExcell.value = 0 Then
            sCadImp = oNCredDoc.ImprimeDevolucionCredPers(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeDevolucionCredPers(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion)
        Else
            sCadImp = oNCredDoc.ImprimeDevolucionCredPersExcel(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion)
        End If
          
    'MADM 20110911
        Case 108411 'reporte de castigo creditos devueltos por descuento de creditos con convenio
        If Me.chkMigraExcell.value = 0 Then
            sCadImp = oNCredDoc.ImprimeDevolucionCredPers(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion, 1)
            If sCadImp = "" Then
                MsgBox "No se encontraron coincidencias con los parametros definidos", vbInformation, "Aviso"
            Else
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeDevolucionCredPers(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion, 1)
            End If
        Else
            sCadImp = oNCredDoc.ImprimeDevolucionCredPersExcel(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "", MatInstitucion, 1)
        End If
    'END MADM
    Case gColCredRepResSaldeCartxInst

        sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXInstitucionConsumo(MatAgencias, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        
    Case gColCredRepLisDesctoPlanilla
        '**DAOR 20070917, Metodo creado para disminuir líneas en el evento click
        Call ColCredRepLisDesctoPlanilla
    Case gColCredRepPagosconCheque

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
            Else
                sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
            End If
        End If
        
    Case gColCredRepPagosdeOtrasAgen

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            End If
        End If
    
    'Case 108312
            'CUSCO
            'frmPaprica.Show 1
    Case 108313
            frmEstadVenMen.Show 1
    
    Case gColCredRepPagosEnOtrasAgen

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            End If
        End If
        
    Case gColCredRepIntEnSusp

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            Else
                sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            End If
        End If
        
'    Case 108308
'        sCadImp = oNCredDoc.ImpLavDineroM(TxtFecIniA02.Text, TxtFecFinA02, gsNomCmac, gsNomAge, sCadAge, gdFecSis, gsCodUser, sCadMoneda)
    Case 108309
        sCadImp = oNCredDoc.ListaDesembolsosXAgenciaXFecha(MatAgencias, gdFecSis, gsNomAge, gdFecSis, gsCodUser, TxtFecIniA02, gsNomCmac)
    Case 108505

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosRefinanciados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRefinanciados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosRefinanciados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosRefinanciados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
            End If
        End If
    Case 108508
        Call GenContaFideicomiso
    Case gColCredRepCredAdjuContabi
        '*** PEAC 20090923
        'Call ReporteGarantiasAdjudicadasContabilidad(CDate(Me.TxtFecIniA02), sAgenciasTemp)
        Call ReporteGarantiasAdjudicadasContabilidad(CDate(Me.TxtFecFinA02), sAgenciasTemp)

'*** PEAC 20080924 - SE CREO UNA OPCION EN EL OPETPO CON ESTE NUMERO PARA UN REPORTE
'    Case 108210
'
'        Dim cMensajex As String
'        Set loRep = New NCredReporte
'        sCadImp = loRep.nRepo108210_MoraxAnalista(cMensajex, gdFecSis, Val(TxtDiasAtrCons1Ini.Text), Val(TxtDiasAtrCons1Fin.Text), Val(TxtDiasAtrCons2Ini.Text), Val(TxtDiasAtrCons2Fin.Text), Val(TxtDiasAtrCons3Ini.Text), sMoneda, sProductos, sCondicion, sAgencias, sAnalistas)

    
'    Case gColCredRepProgPagosxCuota, gColCredRepDatosReqMora, gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsColocxAgencia, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocxFteFinan, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper, _
'        gColCredRepCartaCobMoro1, gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, gColCredRepCartaCobMoro4, gColCredRepCartaCobMoro5, gColCredRepCartaCobMoro6, _
'        gColCredRepCartaInvCredAlt, gColCredRepCartaRecup, gColCredRepCredVigArqueo, _
'        gColCredRepVisitaCobroCuotas, gColCredRepClientesNCuotasPend, gColCredRepIngresosxGasto, gColCredRepCredVigIntDeven, _
'        gColCredRepEstMensual, gColCredRepCredDesmMayores, gColCredRepResSalCartxAna, 108210 ', gColCredRepResSalCarxAna

    '*** SE QUITO LA OPCION 108210
    Case gColCredRepProgPagosxCuota, gColCredRepDatosReqMora, gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsColocxAgencia, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocxFteFinan, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper, _
        gColCredRepCartaCobMoro1, gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, gColCredRepCartaCobMoro4, gColCredRepCartaCobMoro5, gColCredRepCartaCobMoro6, _
        gColCredRepCartaInvCredAlt, gColCredRepCartaRecup, gColCredRepCredVigArqueo, _
        gColCredRepVisitaCobroCuotas, gColCredRepClientesNCuotasPend, gColCredRepIngresosxGasto, gColCredRepCredVigIntDeven, _
        gColCredRepEstMensual, gColCredRepCredDesmMayores, gColCredRepResSalCartxAna, gColCredRepCartaRecup2, gColCredCorresponsaliaPorDebito ', gColCredRepResSalCarxAna'WIOR 20130910 AGREGÓ gColCredRepCartaRecup2
        'WIOR 20140620 AGREGO gColCredCorresponsaliaPorDebito
'*** FIN PEAC 20080924 - SE CREO UNA OPCION EN EL OPETPO CON ESTE NUMERO PARA UN REPORTE

    'gColCredRepResSalCarxAna
        
        Dim cMensaje1 As String
        Dim cMensaje2 As String
        Dim cMensaje As String

        Dim nBandera As Boolean
        
        Dim sAgenciasx As String
            
        sAgenciasx = ListaAgencias()
            
        cMensaje1 = ""
        cMensaje2 = ""
        nBandera = False
        
        For i = 1 To TreeView1.Nodes.Count
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                If TreeView1.Nodes(i).Checked = True Then
                
                    If Len(Trim(sProductos)) = 0 Then
                        sProductos = "'" & Trim(Mid(TreeView1.Nodes(i).Key, 2, 3)) & "'"
                        cMensaje1 = Trim(Mid(TreeView1.Nodes(i).Text, 1, 3))
                    Else
                        sProductos = sProductos & ", '" & Trim(Mid(TreeView1.Nodes(i).Key, 2, 3)) & "'"
                        cMensaje1 = cMensaje1 & "/" & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3))
                    End If
                    nBandera = True
                    
                Else
                    If Len(Trim(cMensaje2)) = 0 Then
                        cMensaje2 = Trim(Mid(TreeView1.Nodes(i).Text, 1, 3))
                    Else
                        cMensaje2 = cMensaje2 & "/" & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3))
                    End If
                End If

            End If
        Next
        
        If nBandera = True Then
            cMensaje = "PRODUCTOS: " & cMensaje1
        Else
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCartxAna Then
                
                cMensaje = cMensaje & " PRODUCTOS: "
                sProductos = ""
                
                For i = 1 To TreeView1.Nodes.Count
                    If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" And Mid(TreeView1.Nodes(i).Key, 2, 1) = "3" Then
                        If Len(Trim(sProductos)) = 0 Then
                            sProductos = "'" & Trim(Mid(TreeView1.Nodes(i).Key, 2, 3)) & "'"
                            cMensaje = cMensaje & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3))
                        Else
                            sProductos = sProductos & ", '" & Trim(Mid(TreeView1.Nodes(i).Key, 2, 3)) & "'"
                            cMensaje = cMensaje & "/" & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3))
                        End If

                    End If
                Next
                
            Else
                cMensaje = "PRODUCTOS: " & cMensaje2
            End If
        End If
        
        cMensaje1 = ""
        cMensaje2 = ""
        nBandera = False
        
        If ChkMonA02(0).value = 1 Then
            If ChkMonA02(1).value = 1 Then
                sMoneda = "'" & gMonedaNacional & "', '" & gMonedaExtranjera & "'"
                cMensaje1 = "Nac./Ext."
            Else
                sMoneda = "'" & gMonedaNacional & "'"
                cMensaje1 = "Nac."
            End If
            nBandera = True
        Else
            If ChkMonA02(1).value = 1 Then
                sMoneda = "'" & gMonedaExtranjera & "'"
                cMensaje1 = "Ext."
                nBandera = True
            Else
                sMoneda = ""
                cMensaje2 = "Nac./Ext."
            End If
        End If
        
        If nBandera = True Then
            cMensaje = cMensaje & " MONEDA: " & cMensaje1
        Else
            cMensaje = cMensaje & " MONEDA: " & cMensaje2
        End If
          
        For i = 0 To nContAna - 1
            If i = 0 Then
                sAnalistas = "'" & matAnalista(i) & "'"
                'cMensaje1 = matAnalista(i)
            Else
                sAnalistas = sAnalistas & ", '" & matAnalista(i) & "'"
                'cMensaje1 = cMensaje1 & "/" & matAnalista(i)
            End If
        Next
         
        'cMensaje = cMensaje & " AGENCIAS: " & cMensaje1
        
        'Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCarxAna Or _

        If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepProgPagosxCuota Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro1 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro2 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro3 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro4 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro5 Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro6 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaInvCredAlt Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup2 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigArqueo Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepVisitaCobroCuotas Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepClientesNCuotasPend Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepIngresosxGasto Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigIntDeven Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredDesmMayores Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCartxAna Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepEstMensual Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredCorresponsaliaPorDebito Then 'WIOR 20130910 AGREGÓ gColCredRepCartaRecup2
           'WIOR 20140620 AGREGO gColCredCorresponsaliaPorDebito
            
            cMensaje1 = ""
            cMensaje2 = ""
            nBandera = False
                
            If ChkCond(0).value = 1 Then
                sCondicion = gColocCredCondNormal
                cMensaje1 = "Norm."
                nBandera = True
            Else
                cMensaje2 = "Norm."
            End If
            If ChkCond(1).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondParalelo
                    cMensaje1 = "Par."
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondParalelo
                    cMensaje1 = cMensaje1 & "/Par."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Par."
            End If
            If ChkCond(3).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondRecurrente
                    cMensaje1 = "Rec."
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondRecurrente
                    cMensaje1 = cMensaje1 & "/Rec."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Rec."
            End If
            
            'ARCV 12-06-2007
            If ChkCond(2).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondRefinan
                    cMensaje1 = "Ref."
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondRefinan
                    cMensaje1 = cMensaje1 & "/Ref."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Ref."
            End If
            If ChkCond(4).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = 5
                    cMensaje1 = "Amp."
                Else
                    sCondicion = sCondicion & ", " & 5
                    cMensaje1 = cMensaje1 & "/Amp."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Amp."
            End If
            If ChkCond(5).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = 6
                    cMensaje1 = "Autom."
                Else
                    sCondicion = sCondicion & ", " & 6
                    cMensaje1 = cMensaje1 & "/Autom."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Autom."
            End If
            'JUEZ 20130604 ********************************
            If ChkCond(6).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = 7
                    cMensaje1 = "Adicional"
                Else
                    sCondicion = sCondicion & ", " & 7
                    cMensaje1 = cMensaje1 & "/Adicional"
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Adicional"
            End If
            'END JUEZ *************************************
            
            '--------
            If nBandera = True Then
                cMensaje = cMensaje & " CONDICION: " & cMensaje1
            Else
                cMensaje = cMensaje & " CONDICION: " & cMensaje2
            End If
            
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepEstMensual Then
                'ARCV 08-05-2007
                'Dim loRep As NCredReporte
    
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    
                For i = 0 To nContAgencias - 1
                    If i = 0 Then
                        sAgencias = "'" & MatAgencias(i) & "'"
                     Else
                        sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                     End If
                Next
    '            sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, sMoneda, sProductos, Trim(txtLineaCredito.Text), sAgencias)
    
                If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
                    sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaNacional, sProductos, Trim(txtLineaCredito.Text), sAgencias)
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaExtranjera, sProductos, Trim(txtLineaCredito.Text), sAgencias)
                Else
                    If ChkMonA02(0).value = 1 Then
                        sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaNacional, sProductos, Trim(txtLineaCredito.Text), sAgencias)
                    Else
                        sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaExtranjera, sProductos, Trim(txtLineaCredito.Text), sAgencias)
                    End If
                End If
            End If
            '-------
            
            'Validando que se escoga lo de ubicacion geografica
            
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro1 Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro2 Or _
               Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro3 Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro4 Or _
               Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro5 Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro6 Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup Or _
               Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup2 Then 'WIOR 20130910 AGREGÓ gColCredRepCartaRecup2
               If ChkUbi.value = vbUnchecked Then
                    sUbicacionGeo = ""
               End If
            End If
                       
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepProgPagosxCuota Then
                Set loRep = New NCredReporte
                Dim nOpcion As Integer
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                If chkMigraExcell Then
                sCadImp = loRep.nRepo108301_ListadoProgramacionResumen(3, cMensaje, TxtFecIniA02.Text, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas, sCadAge)
                Else
                nOpcion = IIf(optReporte(0).value = True, 1, 2)
                sCadImp = loRep.nRepo108301_ListadoProgramacionPagosCuota(nOpcion, cMensaje, TxtFecIniA02.Text, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas, sCadAge)
                End If
            'ARCV 15-02-2007
'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro1 Then
'                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro1, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0, Right(sUbicacionGeo, 12))
'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro2 Then
'                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro2, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0, Right(sUbicacionGeo, 12))
'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro3 Then
'                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro3, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0, Right(sUbicacionGeo, 12))
'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro4 Then
'                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro4, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0, Right(sUbicacionGeo, 12))
'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro5 Then
'                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro5, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0, Right(sUbicacionGeo, 12))
'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup Then
'                sCadImp = Genera_ReporteWORD(gColCredRepCartaRecup, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0, Right(sUbicacionGeo, 12))

'            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCarxAna Then
'                sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCondicion, CInt(TxtCar1I.Text), _
'                                        CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), gsNomCmac, Val(TxtTipCambio.Text), lsTitProductos)

            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro1 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaCobMoro1, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
                
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro2 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaCobMoro2, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro3 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaCobMoro3, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro4 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaCobMoro4, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro5 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaCobMoro5, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro6 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaCobMoro6, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaRecup, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx) 'WIOR 20130910 AGREGÓsAgenciasx
            'WIOR 20130910 *************************************************************
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup2 Then
                sCadImp = Genera_ReporteWORD_NEW(gColCredRepCartaRecup2, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), Right(sUbicacionGeo, 12), sAgenciasx)
            'WIOR FIN *******************************************************************
            'WIOR 20140530 ****************************
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredCorresponsaliaPorDebito Then
                Call GeneraCartaCargoCuentaAhorro(CDate(TxtFecFinA02.Text), sAnalistas, gsCodAge)
            'WIOR FIN *********************************
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaInvCredAlt Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepClientesNCuotasPend Then
                If val(TxtCuotasPend.Text) < 0 Or IsNumeric(TxtCuotasPend.Text) = False Then
                    MsgBox "Ingrese un número de cuotas pendientes válido", vbExclamation, "Aviso"
                    TxtCuotasPend.SetFocus
                    Exit Sub
                Else
                    If val(TxtNotaIni.Text) < 0 Or IsNumeric(TxtNotaIni.Text) = False Then
                        MsgBox "Ingrese una nota válida", vbExclamation, "Aviso"
                        TxtNotaIni.SetFocus
                        Exit Sub
                    Else
                        If val(TxtNotaFin.Text) < 0 Or IsNumeric(TxtNotaFin.Text) = False Then
                            MsgBox "Ingrese una nota válida", vbExclamation, "Aviso"
                            TxtNotaFin.SetFocus
                            Exit Sub
                        Else
                            If val(TxtNotaIni.Text) > val(TxtNotaFin.Text) Then
                                MsgBox "La nota inicial no puede ser mayor que la nota final", vbExclamation, "Aviso"
                                TxtNotaIni.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaInvCredAlt Then
                    sCadImp = Genera_ReporteWORD(gColCredRepCartaInvCredAlt, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, val(TxtDiaAtrIni.Text), val(TxtDiasAtrFin.Text), val(TxtNotaIni.Text), val(TxtNotaFin.Text), ChkPorc.value, val(TxtCuotasPend.Text), sUbicacionGeo)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepClientesNCuotasPend Then
                    Set loRep = New NCredReporte
                    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                    sCadImp = loRep.nRepo108303_ClientesCuotasPend(cMensaje, sMoneda, sProductos, sCondicion, sAnalistas, val(TxtNotaIni.Text), val(TxtNotaFin.Text), ChkPorc.value, val(TxtCuotasPend.Text), sUbicacionGeo, CDbl(TxtCuotasPend.Text))
                     
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigArqueo Then
                If OptProtesto.value = True Then
                    nProtesto = 1
                ElseIf OptSinProtesto.value = True Then
                    nProtesto = 2
                Else
                    nProtesto = 3
                End If
                
                '*** PEAC 20080627
                
'                Set loRep = New NCredReporte
'                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'                sCadImp = loRep.nRepo108203_CreditosVigentes_Arqueo(IIf(optCredVig(0).value = True, 1, IIf(optCredVig(1).value = True, 2, 3)), cMensaje, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas, sAgenciasx, nProtesto, IIf(OptNuevos.value = True, True, False))
'''PARA DESCOMENTAR ALPA
                Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
                    sCadImp = sCadImp & oCOMNCredDoc.nRepo108203_CreditosVigentes_Arqueo(IIf(optCredVig(0).value = True, 1, IIf(optCredVig(1).value = True, 2, 3)), cMensaje, TxtFecIniA02.Text, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas, sAgenciasx, nProtesto, IIf(OptNuevos.value = True, True, False), gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, val(TxtMontoIni), val(TxtMontoFin))
                    'sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepDiasTranscDesdeSoli(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
                Set oCOMNCredDoc = Nothing
                
                '*** FIN PEAC 20080627
                
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepVisitaCobroCuotas Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108302_VisitaCobroCuotas(cMensaje, gdFecSis, sMoneda, sProductos, sCondicion, sAnalistas)
            
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepIngresosxGasto Then
                
                For i = 0 To nContAgencias - 1
                    If i = 0 Then
                        sAgencias = "'" & MatAgencias(i) & "'"
                     Else
                        sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                     End If
                Next
                
                
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108304_IngresosxGastos(Me.TxtFecIniA02.Text, Me.TxtFecFinA02.Text, cMensaje, sMoneda, sProductos, sCondicion, sAnalistas, sAgencias)
                 
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigIntDeven Then
                
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108201_CreditosVigentes_DiasAtraso(cMensaje, gdFecSis, val(TxtDiaAtrIni), val(TxtDiasAtrFin), sMoneda, sProductos, sCondicion, sAnalistas)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredDesmMayores Then
                If TxtTipCambio = "" Or CCur(TxtTipCambio) = 0 Then
                    Set oTipCambio = New nTipoCambio
                    TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY"), TCFijoMes), "0.00")
                    Set oTipCambio = Nothing
                End If
                If val(txtMontoMayor.Text) > 0 Then
                    For i = 0 To nContAgencias - 1
                         If i = 0 Then
                            sAgencias = "'" & MatAgencias(i) & "'"
                        Else
                            sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                        End If
                   Next i
                    Set loRep = New NCredReporte
                    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                    sCadImp = loRep.nRepo108204_CreditosDesembolsadosVigentes(cMensaje, Me.TxtFecIniA02.Text, Me.TxtFecFinA02.Text, val(txtMontoMayor.Text), sMoneda, sProductos, sCondicion, sAgencias, CCur(Me.TxtTipCambio.Text), gdFecSis)
                    'sCadImp = loRep.nRepo108204_CreditosDesembolsadosVigentes(cMensaje, Me.TxtFecIniA02.Text, Me.TxtFecFinA02.Text, Val(txtMontoMayor.Text), sMoneda, sProductos, sCondicion, sAgencias, CCur(Me.TxtTipCambio.Text))
                Else
                    MsgBox "Ud. debe ingresar un monto válido", vbExclamation, "Aviso"
                    txtMontoMayor.SetFocus
                    Exit Sub
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCartxAna Then
                If val(TxtDiasAtrCons1Ini.Text) > val(TxtDiasAtrCons1Fin.Text) Then
                    MsgBox "El rango final debe ser mayor o igual al rango inicial", vbExclamation, "Aviso"
                    TxtDiasAtrCons1Fin.SetFocus
                    Exit Sub
                Else
                    If val(TxtDiasAtrCons2Ini.Text) > val(TxtDiasAtrCons2Fin.Text) Then
                        MsgBox "El rango final debe ser mayor o igual al rango inicial", vbExclamation, "Aviso"
                        TxtDiasAtrCons2Fin.SetFocus
                        Exit Sub
                    Else
                        Set loRep = New NCredReporte
                        loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                        sCadImp = loRep.nRepo108404_SaldosCarteraxAnalista(cMensaje, gdFecSis, val(TxtDiasAtrCons1Ini.Text), val(TxtDiasAtrCons1Fin.Text), val(TxtDiasAtrCons2Ini.Text), val(TxtDiasAtrCons2Fin.Text), val(TxtDiasAtrCons3Ini.Text), sMoneda, sProductos, sCondicion, sAgencias, sAnalistas)
                    End If
                End If
            
            
            ElseIf gColCredRepEstMensual Then
            
            End If
            
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepDatosReqMora Then
         
            
            sCadImp = Genera_Reporte108306(cMensaje, sMoneda, sProductos, sAnalistas)
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsCartAltoRiesgoxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxFteFinan Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsResCartSuper Then
                
            
            cMensaje1 = ""
            cMensaje2 = ""
            nBandera = False
                
            For i = 0 To nContAgencias - 1
                If i = 0 Then
                    sAgencias = "'" & MatAgencias(i) & "'"
                    cMensaje1 = MatAgencias(i)
                Else
                    sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                    cMensaje1 = cMensaje1 & "/" & MatAgencias(i)
                End If
            Next
            
            cMensaje = cMensaje & " AGENCIAS: " & cMensaje1
             
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsCartAltoRiesgoxAna Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108604_CarteraAltoRiesgoxAnalista(cMensaje, val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Then
                If DateDiff("d", Format(gdFecSis, "dd/MM/YYYY"), Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY")) = 0 Then
                    'la fecha que se busca es la fecha actual
                    nTempoParam = 1
                Else
  
                    Set dUltimoDia = New DCredReporte
                    nUltimoDia = dUltimoDia.RecuperaUltimoDiaMes(Me.TxtFecFinA02.Text)
                    If nUltimoDia = val(Mid(TxtFecFinA02.Text, 1, 2)) Then
                        'El dia es el ultimo del mes que se especifica
                        If val(Mid(Format(gdFecSis, "dd/MM/YYYY"), 4, 2)) = val(Mid(Format(Me.TxtFecFinA02, "dd/MM/YYYY"), 4, 2)) And val(Mid(Format(gdFecSis, "dd/MM/YYYY"), 7, 4)) = val(Mid(Format(Me.TxtFecFinA02, "dd/MM/YYYY"), 7, 4)) Then
                            'Es el mismo mes
                            MsgBox "Ud. no puede colocar esta fecha pues en el mes actual solo vale la fecha del sistema", vbExclamation, "Aviso"
                            Me.TxtFecFinA02.SetFocus
                            Exit Sub
                        Else
                            'Es el ultimo dia del mes pasado
                            nTempoParam = 2
                        End If
                    Else
                        MsgBox "La fecha que ud está ingresando no corresponde al último dia de ese mes", vbExclamation, "Aviso"
                        Me.TxtFecFinA02.SetFocus
                        Exit Sub
                    End If
                End If
                 
                If Len(Trim(sProductos)) = 0 Then
                    MsgBox "Ud. debe Seleccionar al menos un producto para buscar", vbInformation, "Aviso"
                    Exit Sub
                End If
                 
'               'Recalculo el tipo de cambio fijo del mes para la fecha especificada
'               Set oTipCambio = New nTipoCambio
'               TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY"), TCFijoMes), "0.00")
'               Set oTipCambio = Nothing
                   
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                
                If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Then
                    sCadImp = loRep.nRepo108602_ConsolidadoColocacionesxAnalista(nTempoParam, cMensaje, val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Then
                    sCadImp = loRep.nRepo108601_ConsolidadoColocacionesxAgencia(nTempoParam, cMensaje, val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Then
                    sCadImp = loRep.nRepo108603_CuadroMetasAlcanzadasxAnalista(nTempoParam, cMensaje, val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Then
                    sCadImp = loRep.nRepo108606_ConsolidadoColocacionesxMoraxAnalista(nTempoParam, cMensaje, val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxFteFinan Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108605_ConsolidadoColocxFteFinanciamiento(cMensaje, val(TxtTipCambio.Text), gdFecSis, sMoneda, sProductos, sAgencias)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsResCartSuper Then
                sCadImp = Genera_Reporte108607(cMensaje, val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
            End If
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepEstMensual Then
            If optEstadistica(2).value = True Then
                If Len(Trim(txtLineaCredito.Text)) = 0 Then
                    MsgBox "Ingrese una linea de crédito", vbExclamation, "Aviso"
                    txtLineaCredito.SetFocus
                    Exit Sub
                End If
            End If
            
            For i = 0 To nContAgencias - 1
                If i = 0 Then
                    sAgencias = "'" & MatAgencias(i) & "'"
                    cMensaje1 = MatAgencias(i)
                Else
                    sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                    cMensaje1 = cMensaje1 & "/" & MatAgencias(i)
                End If
            Next
            
'ARCV 07-05-2007
'            Dim loRep As NCredReporte
'
'            Set loRep = New NCredReporte
'            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'
''            sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, sMoneda, sProductos, Trim(txtLineaCredito.Text), sAgencias)
'
'            If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
'                sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaNacional, sProductos, Trim(txtLineaCredito.Text), sAgencias)
'                sCadImp = sCadImp & Chr$(12)
'                sCadImp = sCadImp & loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaExtranjera, sProductos, Trim(txtLineaCredito.Text), sAgencias)
'            Else
'                If ChkMonA02(0).value = 1 Then
'                    sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaNacional, sProductos, Trim(txtLineaCredito.Text), sAgencias)
'                Else
'                    sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, gMonedaExtranjera, sProductos, Trim(txtLineaCredito.Text), sAgencias)
'                End If
'            End If
             
        End If
          
    '-******   Reportes de Fin de Mes Para Constabilidad y Planeamiento  *******
    '
    '
    '----------------------------------------------------------------------------
    '"1"
    'ALPA 20080619 *************************************************************************
    'MODIFICADO: PEAC 20081003
    Case gColCredRepRenCarAnali:
        sCadImp = ""
        sCadAge = Replace(Replace(sCadAge, "'", ""), " ", "")
        
        If chkMigraExcell Then
            'ALPA 20090324*********************************************************************************************************************************************************************************************************
            'sCadImp = oNComCre.imprimeReporteConsolRentabilidadCarteraXAnalista(TxtFecIniA02.Text, TxtFecFinA02.Text, TxtDiasAtrFin.Text, sAnalistas2, sCadAge, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
            sCadImp = oNComCre.imprimeReporteConsolRentabilidadCarteraXAnalista(TxtFecIniA02.Text, TxtFecFinA02.Text, TxtDiasAtrFin.Text, sAnalistas2, sCadAge, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, sCadProd)
            '**********************************************************************************************************************************************************************************************************************
            Me.chkMigraExcell.value = 0
        Else
            'ALPA 20090324*********************************************************************************************************************************************************************************************************
            'sCadImp = oNComCre.imprimeReporteRentabilidadCarteraXAnalista(TxtFecIniA02.Text, TxtFecFinA02.Text, TxtDiasAtrFin.Text, sAnalistas2, sCadAge)
            sCadImp = oNComCre.imprimeReporteRentabilidadCarteraXAnalista(TxtFecIniA02.Text, TxtFecFinA02.Text, TxtDiasAtrFin.Text, sAnalistas2, sCadAge, sCadProd)
            '**********************************************************************************************************************************************************************************************************************
        End If

    '***************************************************************************************
    'ALPA 20080625 *************************************************************************
    Case gColCredCanXNAtendidos:
        sCadImp = ""
        sCadAge = Replace(Replace(sCadAge, "'", ""), " ", "")
        sCadImp = oNComCre.imprimeReporteClientesConDeudaCanceladaNoVueltosAtender(TxtFecIniA02.Text, TxtFecFinA02.Text, TxtDiasAtrFin.Text, sAnalistas2, sCadAge)
    '***************************************************************************************
    Case 108701:
         sCadImp = ""
         sCadImp = lsRep.nRepo108701_CarteraColocacionesxMoneda("2", lsServerCons)
         
    Case 108702:
        sCadImp = ""
        If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
        sCadImp = lsRep.nRepo108702_ImpCarteraCredConsolidada("2", TxtTipCambio, lsServerCons)
        End If
    Case 108703:
        sCadImp = ""
        If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
        sCadImp = lsRep.nRepo108703_ImpRepCarteraProd_Venc("2", val(TxtTipCambio), lsServerCons)
        End If
        
    Case 108704: '  Reporte por  Producto  Y Agencia (A-2.3)
         sCadImp = ""
         If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
         sCadImp = lsRep.nRepo108704_ImpRepCarteraAgencia_Prod("2", val(TxtTipCambio), "C", lsServerCons)
        End If

    Case 108705: 'Reporte para Reclasificacion de Cartera (A-4)
        sCadImp = ""
        sCadImp = lsRep.nRepo108705_ImpCarteraReclasificacion("2", lsServerCons)
         

    Case 108706: ' Reporte de Intereses Devengados Vigentes (A-5)
         sCadImp = ""
         If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
         sCadImp = lsRep.nRepo108706_ImpRepDevengados_Vigentes(lsServerCons, "2", val(TxtTipCambio))
        End If
         
    Case 108707: ' Reporte de Intereses Devengados Vencidos y Cobranza Judicial (A-6)
         sCadImp = ""
         If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
         sCadImp = lsRep.nRepo108707_ImpRepDevengados_Vencidos(lsServerCons, "2", val(TxtTipCambio))
        End If
         
    Case 108708:  ' Resumen de Garantias  (A-7)
        sCadImp = ""
        If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
        sCadImp = lsRep.nRepo108708_ImpRepResumenGarantias(lsServerCons, "2", val(TxtTipCambio))
        End If
   'Case 108709:  ' Cartera de Alto Riesgo  (A-8)

    Case 108710: '  Colocaciones x Sectores Economicos  (A-9)
          sCadImp = ""
          If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
          sCadImp = lsRep.nRepo108710_ImpRepColocxSectEcon(lsServerCons, "2", val(TxtTipCambio))
        End If
    Case 108711: ' Reporte de Intereses de Créditos (A-4)
            sCadImp = ""
            'sCadImp = lsRep.nRepo108711_ImpCarteraReclasificacion(lsServerCons, "2", "nIntDev")
            sCadImp = lsRep.nRepo108705_ImpCarteraReclasificacion("2", lsServerCons, "nIntDev")
    
    Case 108712: ' Reporte Revision  de Provision de Cartera de Creditos
                sCadImp = ""
                sCadImp = lsRep.nRepo108712_ImpReversionIntDeveng(lsServerCons, "2")
    
    Case 108713:
                sCadImp = ""
                sCadImp = lsRep.nRepo108713_ImpRepCarteraAgencia_Prod(lsServerCons, "2", val(TxtTipCambio), "D")
    
    Case 108714:
                sCadImp = ""
                sCadImp = lsRep.nRepo108713_ImpRepCarteraAgencia_Prod(lsServerCons, "2", val(TxtTipCambio), "S")
    Case 108715:
                Call lsRep.nRepo108715_ImpRepSaldoCartera_Rango(lsServerCons, "2", val(Me.TxtTipCambio))
                sCadImp = "Reporte Ok"
    Case 108721: ' Creditos Vigentes(Garantia) - Pyme
            ' Replace(ValorMoneda, "Credito", "CS") &
            'Lima
            lsCadenaPar = ValorNorRefPar & _
            Replace(ValorProducto, "Credito", "CS") & _
            " And CS.nDiasAtraso >= " & val(Me.TxtDiaAtrIni) & _
            " And CS.nDiasAtraso <= " & val(Me.TxtDiasAtrFin)

             lsCadenaDesPar = DescCondSeleccionado & DescProdSeleccionado
            'AgenciaSeleccionada (False)
            If MatAgencias(0) = "" Then
                MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
               Exit Sub
             End If
             sCadImp = ""
             Dim sTemp As String
             
             For i = 0 To nContAgencias - 1
                sTemp = lsRep.nRepo108721_fgImpCredVigGarant(lsServerCons, val(TxtTipCambio), lsCadenaPar, lsCadenaDesPar, val(Me.TxtDiaAtrIni), val(Me.TxtDiasAtrFin), MatAgencias(i), gdFecSis)
                sCadImp = sCadImp & sTemp & IIf(Len(sTemp) > 50, Chr(12), "")
             Next i
    
    Case 108722: ' Creditos Vigentes (Garantia) - Consumo
            
            lsCadenaPar = ValorNorRefPar & _
            Replace(ValorProdConsumo, "Credito", "CS") & _
            " And CS.nDiasAtraso >= " & val(TxtDiaAtrIni.Text) & _
            " And CS.nDiasAtraso <= " & val(TxtDiasAtrFin.Text)

            lsCadenaDesPar = DescProdConsumoSeleccionado & DescCondSeleccionado
            If MatAgencias(0) = "" Then
                MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
                Exit Sub
            End If
            sCadImp = ""
             For i = 0 To nContAgencias - 1
                    sTemp = lsRep.nRepo108722_fgImpCredPersVigentesGarant(lsServerCons, val(TxtTipCambio), lsCadenaPar, lsCadenaDesPar, val(Me.TxtDiaAtrIni), val(Me.TxtDiasAtrFin), MatAgencias(i), gdFecSis)
                   sCadImp = sCadImp & sTemp & IIf(Len(sTemp) > 50, Chr(12), "")
             Next i
    Case 108723: ' Creditos PIGNORATICIO - Vigentes
            If MatAgencias(0) = "" Then
                MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
                Exit Sub
            End If
            
            sCadImp = ""
            For i = 0 To nContAgencias - 1
                sCadImp = sCadImp & lsRep.nRepo108723_fgImprimeCredPigIntDev(lsServerCons, val(Me.TxtDiaAtrIni), val(Me.TxtDiasAtrFin), MatAgencias(i))
            Next i
            
            
    
    Case 108724:
    
            lsCadenaPar = Replace(ValorMoneda, "Credito", "CS") & _
            Replace(ValorProducto, "Credito", "CS")
            lsCadenaDesPar = DescProdSeleccionado
         If MatAgencias(0) = "" Then
            MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
            Exit Sub
         End If
         
         sCadImp = ""
         For i = 0 To nContAgencias - 1
         
             sTemp = lsRep.nRepo108724_fgImpCredRefinan(lsServerCons, val(Me.TxtTipCambio), lsCadenaPar, lsCadenaDesPar, MatAgencias(i), gdFecSis)
             sCadImp = sCadImp & sTemp & IIf(Len(sTemp) > 50, Chr(12), "")
         Next i
                
    'Case 107825:
            ' reporte de estados de los creditos
            
    Case 108801:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            fechaini = "01" & Mid(CStr(CredRepoMEs.GetFechaCierreMes), 3, 10)
            sCadImp = lsRep.nRepo108801_(lsServerCons, fechaini, CredRepoMEs.GetFechaCierreMes, val(TxtTipCambio), gsCodCMAC)
            End If
    Case 108802:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
                sCadImp = lsRep.nRepo108802_(lsServerCons, val(TxtTipCambio))
            End If
    Case 108803:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            sCadImp = lsRep.nRepo108803_(lsServerCons, val(TxtTipCambio), gsCodCMAC)
            End If
    Case 108804:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            sCadImp = lsRep.nRepo108804_(lsServerCons, val(TxtTipCambio), gsCodCMAC)
            End If
    Case 108806:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
                sCadImp = lsRep.nRepo108806_(lsServerCons, val(TxtTipCambio))
            End If
            
    Case 108808:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            
            TxtFecIniA02 = CredRepoMEs.GetFechaCierreMes
            fechaini = CDate("01" & Mid(CStr(CredRepoMEs.GetFechaCierreMes), 3, 10)) - 1
           Call lsRep.nRepo108808_(lsServerCons, CredRepoMEs.GetFechaCierreMes, val(TxtTipCambio), gsNomCmac, gdFecSis, fechaini, TxtFecIniA02)
            End If
                
    Case 108110:
        sAgenciasx = ListaAgencias()
        sCadImp = oNCredDoc.ImprimeProtestoDia(sAgenciasx, TxtFecIniA02, gsCodUser, "CMAC ICA", gdFecSis)
    Case 108114:
         'reporte de Consolidado por Analista V2.0
          If TxtFecIniA02 < gdFecSis Then
             bFecha = False
          Else
            bFecha = True
          End If
          sCadImp = oNCredDoc.ConsolidadoAnalistaV2(gsCodAge, gsCodUser, TxtFecIniA02, TxtTipCambio, bFecha)
    Case 108111:
        'reporte de consolidado por analista
        
        
        If TxtFecIniA02 < gdFecSis Then
            bFecha = False
        Else
            bFecha = True
        End If
        
        sCadImp = oNCredDoc.ConsolidadoPorAnalista(gsCodAge, TxtFecIniA02, TxtTipCambio, MatAgencias, gsCodUser, matAnalista, MatProductos, bFecha)
    Case 108112
          'reporte de consolidado por producto
                  
        If TxtFecIniA02 < gdFecSis Then
            bFecha = False
        Else
            bFecha = True
        End If
        
       Call oNCredDoc.ConsolidadoPorProductoAna(gsCodAge, TxtFecIniA02, MatAgencias, gsCodUser, gdFecSis, TxtTipCambio, bFecha)
    Case 108310
         ' reporte de los creditos automaticos
          If ChkMes.value = 1 Then
             nProductoAutom = 1
          End If
          If ChkConsumo = 1 Then
            nProductoAutom = nProductoAutom + 2
          End If
          If ChkMes.value = 0 And ChkConsumo.value = 0 Then
            nProductoAutom = 0
          End If
          
          If nProductoAutom = 0 Then
            MsgBox "Debe seleccionar algun producto para este reporte", vbInformation, "AVISO"
            Exit Sub
          End If
          If nProductoAutom = 1 Then
            sCadImp = oNCredDoc.Reporte108310_ListaClientesAutomaticos(MatAgencias, matAnalista, TxtFecIniA02, gsCodAge, gdFecSis, gsCodUser, gsNomCmac, "'201'")
          ElseIf nProductoAutom = 2 Then
            sCadImp = oNCredDoc.Reporte108310_ListaClientesAutomaticos(MatAgencias, matAnalista, TxtFecIniA02, gsCodAge, gdFecSis, gsCodUser, gsNomCmac, "'304'")
          ElseIf nProductoAutom = 3 Then
            sCadImp = oNCredDoc.Reporte108310_ListaClientesAutomaticos(MatAgencias, matAnalista, TxtFecIniA02, gsCodAge, gdFecSis, gsCodUser, gsNomCmac, "'201'")
            sCadImp = sCadImp & oNCredDoc.Reporte108310_ListaClientesAutomaticos(MatAgencias, matAnalista, TxtFecIniA02, gsCodAge, gdFecSis, gsCodUser, gsNomCmac, "'304'")
          End If
    'Case 108205
            'Reporte de Lista de Clientes Vigentes
          '  sCadImp = oNCredDoc.Repor108205_ClientesVigentes(TxtFecIniA02)
    Case 108311
             'reporte de creditos no desembolsados
            sCadImp = oNCredDoc.ListaCreditosNoDesembolsados(MatAgencias, TxtFecIniA02, TxtFecFinA02, gdFecSis, gsNomAge, gsCodUser, gsNomCmac)
    Case 108206
            
            sCadImp = oNCredDoc.Recup_ClientesByAnalista(MatAgencias, matAnalista, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
    
'ARCV 27-07-2006
    Case "108117"

        sCadImp = oNCredDoc.ImprimeProyeccionCancelacionCreditos(TxtFecIniA02.Text, TxtFecFinA02.Text, gsCodUser, gdFecSis, gsNomAge, MatProductos, matAnalista, gsNomCmac)
' COMENTADO X MADM 20111001 - SE UTILIZA EL CODIGO DE OPERACION
'    Case 108411
'        sCadImp = oNCredDoc.ImprimeListadoVigentesXInstitucion(gsCodUser, gdFecSis, gsNomAge, MatInstitucion, MatAgencias, gsNomCmac)
'-----------------------------------
    'ARCV 03-02-2007
    Case 108321
        Dim MatCalificacion As Variant
        Dim nContCalif As Integer
        Dim oCredRep As COMNCredito.NCOMCredDoc
        
        ReDim MatCalificacion(0)
        nContCalif = 0
        For i = 0 To lstCalif.ListCount - 1
            If lstCalif.Selected(i) = True Then
                nContCalif = nContCalif + 1
                ReDim Preserve MatCalificacion(nContCalif)
                MatCalificacion(nContCalif - 1) = lstCalif.List(i)
            End If
        Next i
        
        sCadImp = oNCredDoc.ImprimeListadoCalificacion(gsCodUser, gdFecSis, gsNomAge, MatAgencias, matAnalista, MatCalificacion, gsNomCmac)
    'ARCV 05-02-2007
    Case 108322
        'sCadImp = oNCredDoc.ImprimeListadoPolizas(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
'''PARA DESCOMENTAR ALPA

        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeListadoPolizas(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
        Set oCOMNCredDoc = Nothing
               
        
    'ARCV 07-02-2007
    Case 108323
        sCadImp = oNCredDoc.ImprimeListadoGruposEconomicos(gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
    Case 108324
        sCadImp = oNCredDoc.ImprimeListaReprogramados(gsCodUser, gdFecSis, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsNomAge, MatAgencias, gsNomCmac)
    Case 108381
        sCadImp = oNCredDoc.ImprimeGarantiasXClientes(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
    Case 108382
        sCadImp = oNCredDoc.ImprimeReporteGarantes(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
    Case 108383
        sCadImp = oNCredDoc.ImprimeDetalleGarantiasEnSoles(gsCodUser, gdFecSis, gsNomAge, MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsNomCmac)
    Case 108384
        sCadImp = oNCredDoc.ImprimeGarantiasXTipoDetallado(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
    Case 108385
        sCadImp = oNCredDoc.ImprimeGarantiasXTipoResumido(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
    'By Capi 28122007 se inserto reporte garantias inscritas
    Case gColCredRepGarantiasInscritas
       Call MostrarGarantiasInscritas(CDate(TxtFecFinA02.Text), MatAgencias, sCadMoneda, MatProductos, sEstados)
       
    'By Capi 03112008
    Case gColRepCredRefinan
        Call ImprimirCreditosRefinanciados(CDate(TxtFecFinA02.Text), MatAgencias, val(TxtTipCambio), val(TxtDiasAtrFin.Text))

    '---------
    'ARCV 14-03-2007
    Case 108325
        sCadImp = oNCredDoc.ImprimeListadoCreditosReprogramados(gsCodUser, gdFecSis, MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsNomAge, gsNomCmac)
    '-------------------
    Case gColCredRepCreditosDesBcoBac 'DAOR 20070313, 108331:Reporte de Creditos Desembolsados en Agencias del Banco de la Nación
        If Me.chkMigraExcell.value = 0 Then
            sCadImp = oNCredDoc.nRepo108331_CreditosDesembolsadosBancoNacion(TxtFecIniA02.Text, TxtFecFinA02.Text, MatAgencias, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "")
        Else
            sCadImp = oNCredDoc.nRepo108331_CreditosDesembolsadosBancoNacionExcel(TxtFecIniA02.Text, TxtFecFinA02.Text, MatAgencias, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "")
        End If
    Case gColCredRepCreditosAprExoReglamento 'DAOR 20070418, 108326:Reporte de Creditos Desembolsados con Aprobación de Exoneración de Reglamento
        sCadImp = oNCredDoc.nRepo108326_CreditosDesembConAprobExoneraReglamento(TxtFecIniA02.Text, TxtFecFinA02.Text, MatAgencias, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "108326")
    Case gColCredRepMontoDesembolsadoPorLineas 'DAOR 20070419, Monto Desembolsados por Lineas
        sCadImp = oNCredDoc.nRepo108327_MontoDesembolsadoPorLineas(TxtFecIniA02.Text, TxtFecFinA02.Text, MatAgencias, gsCodUser, gdFecSis, gsNomAge, gsNomCmac, "108327")
    Case gColCredRepEstadosCuentaCredito 'DAOR 20070717, Estados de cuenta de crédito
        Call ImprimeEstadoCuentaCredito(gsCodAge, MatProductos, MatAgencias, TxtFecIniA02.Text, TxtFecFinA02.Text)
    

    Case gColCredRepREULavadoDinero 'By Capi 28012008
        Call mostrarREULavadoDinero(TxtFecIniA02.Text, TxtFecFinA02.Text, MatAgencias)

    Case gColCredRepDUDLavadoDinero 'By Capi 30012008
        Call mostrarDUDLavadoDinero(TxtFecIniA02.Text, TxtFecFinA02.Text, MatAgencias)
    Case gColCredRepConCarxAnalista 'DAOR 20070717, Consolidado de cartera por analista
        sCadImp = sCadImp & oNCredDoc.ImprimeConsolidadoCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac, val(TxtTipCambio.Text), lsTitProductos, IIf(chkCondBN.value = 1, True, False))
'    Case gColCredRepSeguroDesgravConsolida 'DAOR 20071210, Reporte de Seguro de Desgravamen
'        Call MostrarReporteSeguroDesgravamen(CDate(TxtFecFinA02.Text), val(TxtTipCambio.Text))
   Case gColCredRepSeguroDesgravConsolida 'MADM 20110329 -- DAOR 20071210, Reporte de Seguro de Desgravamen
        Call MostrarReporteSeguroDesgravamen(CDate(TxtFecFinA02.Text), val(TxtTipCambio.Text), CDate(Me.TxtFecIniA02.Text), sCadMoneda, IIf(optCredVig(3).value = True, True, False), IIf(optCredVig(6).value = True, True, False), IIf(optCredVig(7).value = True, True, False))
         TxtFecIniA02.Text = gdFecData
         TxtFecFinA02.Text = gdFecData
    Case gColCredRepActaComiteCredAprobados 'PEAC 20070822 reporte de actas de comite de creditos aprobados
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            Call ImprimeActaComiteCredApro(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsCodAge, gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac)
        Set oCOMNCredDoc = Nothing
    Case gColCredRepXTipoCondProd 'PEAC 20070822 reporte por tipo de condicion de producto
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepXTipoCondProd(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsCodAge, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCadMoneda, gsNomCmac, val(TxtTipCambio.Text), matAnalista)
        Set oCOMNCredDoc = Nothing
    Case gColCredRepDetalleCuotasXCobrar  'PEAC 20070905 reporte de detalle de cuotas por cobrar
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepDetalleCuotasXCobrar(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsCodAge, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCadMoneda, gsNomCmac, matAnalista)
        Set oCOMNCredDoc = Nothing
    Case gColCredRepCreditosXCampanha  'PEAC 20070905 reporte de creditos por campañas y/o productos
       'MADM 20091117
        'If chkvigentes.value = 1 And chkcancelados.value = 0 Then
        '    ntipoCred = 1
        If chkvigentes.value = 0 And chkcancelados.value = 1 Then
              ntipoCred = 2
        ElseIf chkvigentes.value = 1 And chkcancelados.value = 1 Then
              ntipoCred = 3
        Else
              ntipoCred = 1
              'MsgBox "Debe seleccionar algun tipo de credito para este reporte", vbInformation, "AVISO"
              'Exit Sub
        End If
       'END MADM
       'MADM 20100719 - CONDICION
        Dim sCondicionC As String
        Dim vCondiC As String
        sCondicionC = ObtieneCondi(vCondiC)
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepCreditosXCampanha(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCadMoneda, gsNomCmac, matAnalista, matCampanha, ntipoCred, sCondicionC)
        Set oCOMNCredDoc = Nothing
        'END MADM
    Case gColCredRepConvCasillero  'PEAC 20070921 reporte de cobros diarios convenio (casillero)
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepConvCasillero(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCadMoneda, MatInstitucion, gsNomCmac, matAnalista, matCampanha, MatGastos)
        Set oCOMNCredDoc = Nothing
    
    'peac 20071228
    Case gColCredRepComVigEEFF '108329
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeComVigEEFF(gsCodUser, gdFecSis, gsNomAge, MatAgencias, gsNomCmac)
        Set oCOMNCredDoc = Nothing

    'peac 20071228
    Case gColCredRepPoliIncendio '108340
        Call ImprimePoliIncendio(gsCodUser, gdFecSis, gsNomAge, MatProductos, MatAgencias, CDate(TxtFecIniA02.Text), CDate(Me.TxtFecFinA02.Text), gsNomCmac)

    Case gColCredRepDiasTranscDesdeSoli 'PEAC 20080215 108141
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepDiasTranscDesdeSoli(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        Set oCOMNCredDoc = Nothing
    Case gColCredRepSoliRechazadas 'PEAC 20080215 108142
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepSoliRechazadas(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        Set oCOMNCredDoc = Nothing
    Case gColCredRepSoloEnEstadoSoli 'PEAC 20080215 108143
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepSoloEstadoSoli(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        Set oCOMNCredDoc = Nothing
    Case gColCredRepSoliProcesadas 'PEAC 20080215 108144
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepSoliProcesadas(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        Set oCOMNCredDoc = Nothing

    Case gColCredRepCredVencPaseCastigo '***PEAC 20080219 108335
            'MADM FRA MONTOS x UIT
            'Call ImprimeRepCredVencPaseCastigo(CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), MatAgencias, txtUit, txtTipCambio, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
            If FrmMontos.Visible = True Then
                    If val(TxtMontoIni) = 0 Or val(TxtMontoFin) = 0 Then
                        MsgBox "Debe ingresar los rangos de los montos", vbInformation, "AVISO"
                        Exit Sub
                    End If
                    
                    If IsNumeric(TxtMontoIni) = False Or IsNumeric(TxtMontoFin) = False Then
                        MsgBox "Los montos deben ser formatos numericos", vbInformation, "AVISO"
                        Exit Sub
                    End If
            End If
           Call ImprimeRepCredVencPaseCastigo(CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), MatAgencias, txtUit, val(TxtMontoIni), val(TxtMontoFin), TxtTipCambio, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
    
    Case gColCredRepCliConDistinTiposCred '***PEAC 20080221 108336
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepCliConDistinTiposCred(MatAgencias, MatProductos, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        Set oCOMNCredDoc = Nothing
        
    '************************************************************'
    '** GITU 20080926 108337 Segun Memo Nº 1705-2008-GM-DI/CMAC *'

    Case gColCredRepRiegCredOfiBN
        Call GeneraArchExcelRepRiesgosCredi(TxtTipCambio, gdFecSis)
    
    '************************************************************'
    
    Case gColCredRepClientesHisNegativo '***PEAC 20080704 108350
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepClientesHisNegativo(MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        Set oCOMNCredDoc = Nothing
        
    Case gColCredRepClientesPotencialesSinCredVig '*** PEAC 20080923 108360
        Call ImprimeRepClientesPotencialesSinCredVig(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
    
    Case gColCredRepNumYSaldoCredPorProductoConsol '*** PEAC 20080924 108210
        Call ImprimeRepNumYSaldoCredPorProductoConsol(TxtTipCambio, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
    
    Case gColCredRepAECIReporte01 '***PEAC 20080303 108921
            Call ImprimeRepAECIReporte01(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsCodUser, gdFecSis, gsNomAge, gsNomCmac)

    Case 108125  'PEAC 20070921 cobranza para gestores
    Dim lxRep  As COMNColocRec.NCOMColRecRConsulta   'peac 20071003
    Dim lsListaAgencias As String
    
         Set lxRep = New COMNColocRec.NCOMColRecRConsulta
            'Screen.MousePointer = 11
            lsListaAgencias = DameAgencias
            sCadImp = sCadImp & lxRep.nRepo138038_RepCobranzaGestores(gsCodAge, TxtDiaAtrIni, TxtDiasAtrFin, TxtFecIniA02.Text, TxtFecFinA02.Text, gsNomAge, gdFecSis, gsCodUser, lsListaAgencias, "", 1)
            'Screen.MousePointer = 0
        Set lxRep = Nothing

    'By Capi Acta 014-2007
    Case gColCredRepFepInforme01
        If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            fechaini = "01" & Mid(CStr(CredRepoMEs.GetFechaCierreMes), 3, 10)
            sCadImp = lsRep.nRepo108801_(lsServerCons, fechaini, CredRepoMEs.GetFechaCierreMes, val(TxtTipCambio), gsCodCMAC)
        End If
'        sOperacion = gColCredRepFepInforme01
'        nMes = Month(CDate(TxtFecFinA02.Text))
'        nAno = Year(CDate(TxtFecFinA02.Text))
'        Call SayInformacionFepCmac(sOperacion, Val(TxtTipCambio.Text), nMes, nAno)
'
    
    Case gCapCredRepFepInforme02
        sOperacion = gCapCredRepFepInforme02
        nMES = Month(CredRepoMEs.GetFechaCierreMes)
        nAno = Year(CredRepoMEs.GetFechaCierreMes)
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
        
    Case gColCredRepFepInforme03
        sOperacion = gColCredRepFepInforme03
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
        

    Case gColCredRepFepInforme3a
        sOperacion = gColCredRepFepInforme3a
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
        
    Case gColCredRepFepInforme3b
        sOperacion = gColCredRepFepInforme3b
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
        
    Case gColCredRepFepInforme3c
        sOperacion = gColCredRepFepInforme3c
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
        

    Case gColCredRepFepInforme3d
        sOperacion = gColCredRepFepInforme3d
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
        
    Case gColCredRepFepInforme04
         If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingrese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            'sCadImp = lsRep.nRepo108804_(lsServerCons, Val(TxtTipCambio), gsCodCMAC)
            sCadImp = lsRep.nRepo108804_(lsServerCons, val(TxtTipCambio), gsCodCMAC)
            End If
'        sOperacion = gColCredRepFepInforme04
'        nMes = Month(CDate(TxtFecFinA02.Text))
'        nAno = Year(CDate(TxtFecFinA02.Text))
'        Call SayInformacionFepCmac(sOperacion, Val(TxtTipCambio.Text), nMes, nAno)
'
    Case gColCredRepFepInforme06
        sOperacion = gColCredRepFepInforme06
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)

    Case gColCredRepFepEntorno
        sOperacion = gColCredRepFepEntorno
        nMES = Month(CDate(TxtFecFinA02.Text))
        nAno = Year(CDate(TxtFecFinA02.Text))
        Call SayInformacionFepCmac(sOperacion, val(TxtTipCambio.Text), nMES, nAno)
  
    'End By
'Para descomentar
    Case gColCredRepResSalCarxAnaConsolida 'DAOR 20070814
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsol(lsServerCons, MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                            CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), gsNomCmac, val(TxtTipCambio.Text), lsTitProductos, IIf(chkCondBN.value = 1, True, False))
        Set oCOMNCredDoc = Nothing
    Case 108730 'ALPA 2008/04/11
            Call GenCanceladosxAmpliacion
    Case 108731 'ALPA 2008/04/11
            Call GenPreCancelaciones
    Case gColCredRepConCarxAnalistaConsolida 'DAOR 20070814
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
            sCadImp = sCadImp & oCOMNCredDoc.ImprimeConsolidadoCarteraXAnalistaConsol(lsServerCons, MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gsCodUser, gdFecSis, gsNomAge, MatProductos, gsNomCmac, val(TxtTipCambio.Text), lsTitProductos)
        Set oCOMNCredDoc = Nothing
    
    'MAVM 16052009 Reporte de Creditos May 90 Dias
    Case 108370
        Call ImprimeRepCredVencMay90Dias(CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), MatAgencias, TxtTipCambio, Format(TxtFecIniA02.Text, "yyyymmdd"), Format(TxtFecFinA02.Text, "yyyymmdd"), gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
    Case 108371
        Call ImprimeRepCredClientesPreferenciales(MatAgencias, MatProductos, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
    'By MAVM
    
    'MAVM 20100510 ***
    Case gColCredRepResXAnalistaBPPR
        If dcAgencia.Enabled = True Then
            ReDim MatAgencias(1)
            MatAgencias(0) = dcAgencia.BoundText
        End If
        
        ReDim Preserve MatCondicion(7) 'JUEZ 20130604
        MatCondicion(0) = "1"
        MatCondicion(1) = "2"
        MatCondicion(2) = "3"
        MatCondicion(3) = "4"
        MatCondicion(4) = "5"
        MatCondicion(5) = "6"
        MatCondicion(6) = "7" 'JUEZ 20130604
        
        If dcCartera.BoundText = 1 Then
            ReDim MatProductos(4)
            MatProductos(0) = "201"
            MatProductos(1) = "202"
            MatProductos(2) = "204"
            MatProductos(3) = "304"
        Else
            ReDim MatProductos(8)
            MatProductos(0) = "101"
            MatProductos(1) = "102"
            MatProductos(2) = "103"
            MatProductos(3) = "301"
            MatProductos(4) = "320"
            MatProductos(5) = "401"
            MatProductos(6) = "403"
            MatProductos(7) = "423"
        End If
        
        sCadImp = ReporteCredResultadoAnalistaBPPR_Excel(lsServerCons, MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, gsNomCmac, val(TxtTipCambio.Text))
    
    Case gColCredRepBonificacionXAnalistaBPPR
        If dcAgencia.Enabled = True Then
            ReDim MatAgencias(1)
            MatAgencias(0) = dcAgencia.BoundText
        End If
        
        ReDim Preserve MatCondicion(7) 'JUEZ 20130604
        MatCondicion(0) = "1"
        MatCondicion(1) = "2"
        MatCondicion(2) = "3"
        MatCondicion(3) = "4"
        MatCondicion(4) = "5"
        MatCondicion(5) = "6"
        MatCondicion(6) = "7" 'JUEZ 20130604
        
        If dcCartera.BoundText = 1 Then
            ReDim MatProductos(4)
            MatProductos(0) = "201"
            MatProductos(1) = "202"
            MatProductos(2) = "204"
            MatProductos(3) = "304"
        Else
            ReDim MatProductos(8)
            MatProductos(0) = "101"
            MatProductos(1) = "102"
            MatProductos(2) = "103"
            MatProductos(3) = "301"
            MatProductos(4) = "320"
            MatProductos(5) = "401"
            MatProductos(6) = "403"
            MatProductos(7) = "423"
        End If
        
        sCadImp = ReporteCredBonificacionAnalistaBPPR_Excel(lsServerCons, MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, gsNomCmac, val(TxtTipCambio.Text))
    'MAVM ***
    
    End Select
    
    If Len(Trim(sCadImp)) <= 1 And Mid(TVRep.SelectedItem.Text, 1, 6) <> "108307" And Mid(TVRep.SelectedItem.Tag, 1, 6) <> "108111" And Mid(TVRep.SelectedItem.Tag, 1, 6) <> "108112" And Mid(TVRep.SelectedItem.Tag, 1, 6) <> "108205" And Mid(TVRep.SelectedItem.Tag, 1, 6) <> "108114" _
       And Mid(TVRep.SelectedItem.Text, 1, 6) <> "108115" And Mid(TVRep.SelectedItem.Text, 1, 6) <> "108202" And (Me.chkMigraExcell.Visible = False Or Me.chkMigraExcell.value = 0) Then 'DAOR 20070313, se agregó la verificación Si  no es un reporte en excel
''      By Capi Set 07
'
'            Case gColCredRepFepInforme03, gColCredRepFepInforme3a, gColCredRepFepInforme3b, gColCredRepFepInforme3c, gColCredRepFepInforme3d, _
'                 gColCredRepFepInforme04, gColCredRepFepInforme06, gColCredRepFepEntorno
'                  MsgBox "Informacion Generada Satisfactoriamente", vbExclamation, "Aviso"

'       MsgBox "No existen datos para el reporte", vbExclamation, "Aviso"
    Else
        Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
                
            'Case gColCredRepDatosReqMora, gColCredRepConsResCartSuper, gColCredRepCartaCobMoro1, _
                 gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, gColCredRepCartaCobMoro4, _
                 gColCredRepCartaCobMoro5, gColCredRepCartaCobMoro6, gColCredRepCartaInvCredAlt, gColCredRepCartaRecup, 108210
                 
'peac 20080303 porque en el opetpo esta visible=0
'            Case 108111
'                'GeneraReporteXls1
'                    Dim oConec As DConecta
'                    Dim sSql As String
'                    Dim rs As ADODB.Recordset
'                    Dim Ix As Integer
'                    Dim sAgenciaX As String
'                    Dim sAgenciaNom As String
'
'                    Set oConec = New DConecta
'                    oConec.AbreConexion
'                    sSql = "select * "
'                    sSql = sSql & " From TmpConsolidadoAnalista"
'                    sSql = sSql & " Where IdCodigo=(Select Max(IdCodigo) From TmpConsolidadoAnalista Where Usuario='" & gsCodUser & "')"
'
'                    Set rs = oConec.CargaRecordSet(sSql)
'                    oConec.CierraConexion
'                    Set oConec = Nothing
'
'                    For Ix = 0 To UBound(MatAgencias) - 1
'                        If Ix = 0 Then
'                            sAgenciaX = MatAgencias(Ix)
'                        Else
'                            sAgenciaX = sAgenciaX & ", " & MatAgencias(Ix)
'                        End If
'
'                    Next Ix
'
'                    Set oConec = New DConecta
'                    sSql = "Select cAgeDescripcion From Agencias Where cAgeCod='" & sAgenciaX & "'"
'                    oConec.AbreConexion
'                    Set rs = oConec.CargaRecordSet(sSql)
'                    oConec.CierraConexion
'                    Set oConec = Nothing
'
'                    If Not rs.EOF And Not rs.BOF Then
'                        sAgenciaNom = rs!cAgeDescripcion
'                    End If
'                    Set rs = Nothing
'
'                    'imprime el crystalreport
'                    Dim odCredDoc1 As DCredDoc
'                    Dim nCodigo As Integer
'                    Dim oConec1 As DConecta
'
'                    Set odCredDoc1 = New DCredDoc
'                    nCodigo = odCredDoc1.ObtenerIdCodigo(TxtFecIniA02, gsCodUser)
'                    Set odCredDoc1 = Nothing
'                    Set oConec1 = New DConecta
'                    oConec1.AbreConexion
                    
                    'CUSCO
'                     With CR
'                         .Connect = oConec1.GetStringConnection
'                         .Connect = "Data Source=" & oConec1.ServerName & ";User Id=USERSICMACCONS;Initial Catalog=" & oConec1.DatabaseName & ";pwd=sicmacicons"
'                         .Connect = "dsn=" & oConec1.ServerName & ";uid=USERSICMACCONS;dsq=" & oConec1.DatabaseName & ";pwd=sicmacicons"
'                         .WindowControls = True
'                         .WindowState = crptMaximized
'                         .SelectionFormula = "{TmpConsolidadoAnalista.IdCodigo}=" & nCodigo
'                         .ParameterFields(0) = "PtCambio;" & TxtTipCambio & ";True"
'                         .ParameterFields(1) = "PAgencia;" & gsNomAge & ";True"
'                         .ParameterFields(2) = "pUsuario;" & gsCodUser & ";True"
'                         .ParameterFields(3) = "pAgenciaConsulta;" & sAgenciaNom & ";True"
'                         .ParameterFields(4) = "pFechaEstadistica;" & TxtFecIniA02 & ";True"
'                         .ReportFileName = App.path & "\Rpts\ConsolAnalista.rpt"
'                         .Destination = crptToWindow
'                         .WindowState = crptNormal
'                         .Action = 1
'
'                        ' .Reset
'                End With

'peac 20080303
'            Case 108114
'                    sSql = "select * "
'                    sSql = sSql & " From TmpConsolidadoAnalista"
'                    sSql = sSql & " Where IdCodigo=(Select Max(IdCodigo) From TmpConsolidadoAnalista Where Usuario='" & gsCodUser & "')"
'
'                    Set oConec = New DConecta
'                    oConec.AbreConexion
'                    Set rs = oConec.CargaRecordSet(sSql)
'                    oConec.CierraConexion
'                    Set oConec = Nothing
'
'                    Set odCredDoc1 = New DCredDoc
'                    nCodigo = odCredDoc1.ObtenerIdCodigo(TxtFecIniA02, gsCodUser)
'                    Set odCredDoc1 = Nothing
'                    Set oConec1 = New DConecta
'                    oConec1.AbreConexion
                    
                    'CUSCO
'                     With CR
'                         .Connect = oConec1.GetStringConnection
'
'                         .Connect = "Data Source=" & oConec1.ServerName & ";User Id=USERSICMACCONS;Initial Catalog=" & oConec1.DatabaseName & ";pwd=sicmacicons"
'                         .Connect = "dsn=" & oConec1.ServerName & ";uid=USERSICMACCONS;dsq=" & oConec1.DatabaseName & ";pwd=sicmacicons"
'                         .WindowControls = True
'                         .WindowState = crptMaximized
'                         .SelectionFormula = "{TmpConsolidadoAnalista.IdCodigo}=" & nCodigo
'                         .ParameterFields(0) = "PtCambio;" & TxtTipCambio & ";True"
'                         .ParameterFields(1) = "PAgencia;" & gsNomAge & ";True"
'                         .ParameterFields(2) = "pUsuario;" & gsCodUser & ";True"
'                         .ParameterFields(3) = "pAgenciaConsulta;" & sAgenciaNom & ";True"
'                         .ParameterFields(4) = "pFechaEstadistica;" & TxtFecIniA02 & ";True"
'                         .ReportFileName = App.path & "\Rpts\ConsolAnalista.rpt"
'                         .Destination = crptToWindow
'                         .WindowState = crptNormal
'                         .Action = 1
'
'                        ' .Reset
'                        End With

'peac 20080303
'            Case 108112:
'                  'Consolidado por producto
'                   Set oConec = New DConecta
'                   oConec.AbreConexion
'                   sSql = "Select *"
'                   sSql = sSql & " From TmpConsolidadoProducto"
'                   sSql = sSql & " Where IdCodigo=(Select Max(IdCodigo) From TmpConsolidadoProducto Where Usuario='" & gsCodUser & "')"
'
'                   Set rs = oConec.CargaRecordSet(sSql)
'                   oConec.CierraConexion
'                   Set oConec = Nothing
'
'                   For Ix = 0 To UBound(MatAgencias) - 1
'                        If Ix = 0 Then
'                            sAgenciaX = sAgenciaX
'                        Else
'                            sAgenciaX = sAgenciaX & "," & MatAgencias(Ix)
'                        End If
'                   Next Ix
'
'                   Set oConec = New DConecta
'                   sSql = "Select cAgeDescripcion From Agencias Where cAgeCod='" & sAgenciaX & "'"
'                   oConec.AbreConexion
'                   Set rs = oConec.CargaRecordSet(sSql)
'                   oConec.CierraConexion
'                   Set oConec = Nothing
'
'                   If Not rs.EOF And Not rs.BOF Then
'                      sAgenciaNom = rs!cAgeDescripcion
'                   End If
'
'                   'Imprime el Crystal Report
'                   Set odCredDoc1 = New DCredDoc
'                   nCodigo = odCredDoc1.ObtenerIDCodigoProductoAnalista(TxtFecIniA02, gsCodUser)
'                   Set odCredDoc1 = Nothing
'                   Set oConec1 = New DConecta
'                   oConec1.AbreConexion
                   'CUSCO
'                   With CR
'                        .Connect = oConec1.GetStringConnection
'                        .Connect = "dsn=" & oConec1.ServerName & ";uid=USERSICMACCONS;dsq=" & oConec1.DatabaseName & ";pwd=sicmacicons"
'                        .WindowControls = True
'                        .WindowState = crptMaximized
'                        .SelectionFormula = "{TmpConsolidadoProducto.IdCodigo}=" & nCodigo
'                        .ParameterFields(1) = "PAgencia;" & gsNomAge & ";True"
'                        .ParameterFields(0) = "PtipoCambio;" & TxtTipCambio & ";True"
'                        .ParameterFields(3) = "pAgenciaFiltro;" & sAgenciaNom & ";True"
'                        .ParameterFields(4) = "pFechaFiltro;" & TxtFecIniA02 & ";True"
'                        .ReportFileName = App.path & "\Rpts\ConsolProducto.rpt"
'                        .Destination = crptToWindow
'                        .WindowState = crptNormal
'                        .Action = 1
'                   End With

'peac 20080303
'                Case 108205
'                        'Lista de Clientes vigentes
'                        If ChkMonA02(0).value = 1 Then
'                            nMoneda = 1
'                        Else
'                            nMoneda = 2
'                        End If
'                        Set oConec1 = New DConecta
'                        oConec1.AbreConexion

                        'CUSCO
'                        With CR
'                            .Connect = oConec1.GetStringConnection
'                            .Connect = "dsn=" & oConec1.ServerName & ";UID=SA;DSQ=" & oConec1.DatabaseName & ";pwd=cmacica"
'                            .WindowControls = True
'                            .WindowState = crptMaximized
'
''                            .ParameterFields(0) = "pMes;" & MonthNamxe(Month(Me.TxtFecIniA02), False) & ";True"
''                            .ParameterFields(1) = "pAno;" & Year(Me.TxtFecIniA02) & ";True"
''                            .ParameterFields(2) = "pUser;" & gsCodUser & ";True"
''                            .ParameterFields(3) = "@dFecha;" & Format(Me.TxtFecIniA02, "MM/dd/yyyy") & ";True"
''                            .ReportFileName = App.path & "\Rpts\ListaClientesGravament.rpt"
''                            .Destination = crptToWindow
'                            '.PrintFileType = crptExcel50
'                            '.PrintFileName = "C:\ListaClientes.xls"
'                            .ParameterFields(0) = "pUser;" & gsCodUser & ";True"
'                            .ParameterFields(1) = "pAno;" & Year(TxtFecIniA02.Text) & ";True"
'                            .ParameterFields(2) = "pMes;" & Month(TxtFecIniA02.Text) & ";True"
'                            .ParameterFields(3) = "pMoneda;" & IIf(nMoneda = 1, "SOLES", "DOLARES") & ";True"
'                            .ParameterFields(4) = "@dFecha;" & Format(TxtFecIniA02, "MM/dd/yyyy") & ";True"
'                            .ParameterFields(5) = "@nMoneda;" & nMoneda & ";True"
'                            .ReportFileName = App.path & "\Rpts\ListClientMonth.rpt"
'                            .Destination = crptToWindow
'                            .WindowState = crptNormal
'                            .Action = 1
'                            .Reset
'                        End With
                          
            Case "108115"  'Reporte de Consolidado de Clientes
                    sParam1 = "pUsuario;" & gsCodUser & ":True"
                    sParam2 = "pAgencia;" & gsNomAge & ";True"
                    sParam3 = "@psCodAgen;" & GetListaAgencias() & ";True"
                    sParam4 = "@psCodAnalista;" & GetListaAnalistas() & ";True"
                    Call ImprimirCrystal(App.Path & "\rpts\ConsolClientes.rpt", , sParam1, sParam2, sParam3, sParam4)
            
            Case Else
                If Me.chkMigraExcell.Visible = False Or Me.chkMigraExcell.value = 0 Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepProgPagosxCuota Then    'DAOR 20070313, Si no es un reporte en excel
                    If Mid(TVRep.SelectedItem.Text, 1, 6) <> "108307" Then
                        'P.Show Chr$(27) & Chr$(77) & sCadImp, "Reportes de Creditos", True
                        P.Show sCadImp, "Reportes de Creditos", True
                    End If
                End If
        End Select
    End If

    
    Set Rcd = Nothing
    Set P = Nothing
    Set oNCredDoc = Nothing
    Set CredRepoMEs = Nothing
    Set lsRep = Nothing
End Sub


Sub ImprimeConsolidadoAnalista()
Dim oConec As DConecta
Dim ssql As String
Dim rs As ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim nFila As Integer
                    
    Set oConec = New DConecta
    oConec.AbreConexion
    ssql = "select * "
    ssql = ssql & " From TmpConsolidadoAnalista"
    ssql = ssql & " Where IdCodigo=(Select Max(IdCodigo) From TmpConsolidadoAnalista Where Usuario='" & gsCodUser & "')"

    Set rs = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
    
    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.Path & "\FormatosCarta\ConsolidadoPorAnalista.xls") Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "FormatosCarta\ConsolidadoPorAnalista.xls")
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets(1)
    
    nFila = 7
    
     Do Until rs.EOF
        With xlHoja1
            .Cells(nFila, 1) = rs!CodAna
            .Cells(nFila, 2) = rs!NumCreVig
            .Cells(nFila, 3) = rs!KDesemVig
            .Cells(nFila, 4) = rs!SaldoKVig
            .Cells(nFila, 5) = rs!SaldoKVig / rs!NumCreVig
            .Cells(nFila, 6) = rs!NumCreNewMes
            .Cells(nFila, 7) = rs!KDesemNewMes
            .Cells(nFila, 8) = rs!NumCreRepreMes
            .Cells(nFila, 9) = rs!KDesemRepreMes
            .Cells(nFila, 10) = rs!NumCreVen1a15
            .Cells(nFila, 11) = rs!K1a15
            .Cells(nFila, 12) = (rs!NumCreVen1a15 / rs!SaldoKVig) * 100
            .Cells(nFila, 13) = rs!NumCreVen16a30
            .Cells(nFila, 14) = rs!K16a30
            .Cells(nFila, 15) = (rs!NumCreVen16a30 / rs!SaldoKVig) * 100
            .Cells(nFila, 16) = rs!NumCreVen31aN
            .Cells(nFila, 17) = rs!K31aN
            .Cells(nFila, 18) = (rs!NumCreVen31aN / rs!SaldoKVig) * 100
            .Cells(nFila, 19) = rs!NumCreJud
            .Cells(nFila, 20) = rs!KJudicial
            .Cells(nFila, 21) = (rs!NumCreJud / rs!SaldoKVig) * 100
            .Cells(nFila, 22) = rs!NumJudVen
            .Cells(nFila, 23) = rs!KJudVen
            .Cells(nFila, 24) = (rs!NumJudVen / rs!SaldoKVig) * 100
            rs.MoveNext
        End With
     Loop
     Set rs = Nothing
     
End Sub

Public Sub nRepo108808_(ByVal psServConsol As String, ByVal pdFechaDesde As Date, ByVal pdFechaHasta As Date, _
ByVal pnTipoCambio As Double)
Dim Co As nCredRepoFinMes
'Dim xlAplicacion As Excel.Application
'Dim xlLibro As Excel.Workbook
'Dim xlHoja1 As Excel.WorksheetDim xlHojaP As Excel.Worksheet
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim Total As Double
Dim Tabula As Integer

Dim lsCond1(11) As String, lsCond2(11) As String
Dim Det As Integer
Dim lnFil As Integer, lnCol As Integer

Dim lnNroCreFMS As Currency, lnNroCreFMD As Currency
Dim lnMonCreFMS As Currency, lnMonCreFMD As Currency
Dim lnNroCreOtorgS As Currency, lnNroCreOtorgD As Currency
Dim lnMonCreOtorgS As Currency, lnMonCreOtorgD As Currency
Dim lnNroCreCancelS As Currency, lnNroCreCancelD As Currency
Dim lnMonCreCancelS As Currency, lnMonCreCancelD As Currency
Dim lnNroCredS As Currency, lnNroCredD As Currency
Dim lnMonCredS As Currency, lnMonCredD As Currency

Dim Titulo As String
Dim lsCreditosVigentes As String
Dim lsPignoraticio As String
Dim lsVig As String
'Dim Tabula As Integer

Set Co = New nCredRepoFinMes
lsCreditosVigentes = gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc
lsPignoraticio = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov
lsVig = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstRecCanJud & "," & gColocEstRecCanCast
'On Error GoTo ErrorExcel
Screen.MousePointer = 11

Total = 4 * 25
'Me.barra.Max = Total
'rtf.Text = ""
Tabula = 20
ReDim lineas(20)
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\INFORME_COLOC_BCR.xls") Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\INFO4.xls")
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
Set xlHoja1 = xlLibro.Worksheets.Add

'--************************** CLIENTES NUEVOS Y CONOCIDOS *****************************
'EncabezadoInfo4Excel
Titulo = " C R E D I T O   E M P R E S A R I A L"
xlAplicacion.Range("A1:E7").Font.Bold = True
xlAplicacion.Range("A1:E7").Font.Size = 9
xlAplicacion.Range("A5:P15").Font.Size = 8
xlAplicacion.Range("A4:E20").Font.Size = 8
xlAplicacion.Range("A7:E7").HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range("A11:E11").HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range("A11:E12").Font.Bold = True
xlHoja1.Cells(1, 3) = "R E P O R T E   I N F O 4"
xlHoja1.Cells(2, 2) = gsNomCmac
xlHoja1.Range("B2:E3").MergeCells = True
xlHoja1.Cells(3, 2) = "INFORMACION AL " & Format(gdFecSis, "dd/mm/yyyy")
xlHoja1.Cells(4, 2) = "T.C.F. :" & Format(pnTipoCambio, "#,#0.000")
xlHoja1.Cells(5, 3) = Titulo

'---------------------------------------
For Det = 1 To 11
    Select Case Det
        Case 1
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('101','201') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 2
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('101','201') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) not in('01') "
        Case 3
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('301') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 4
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('302','303') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 5
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('304') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 6
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('401','423') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 7
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('301') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 8
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('302','303') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 9
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('304') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 10
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('401','423') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 11
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('305') "
            lsCond2(Det) = "  "
    End Select

    Sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumOtorgS , " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumOtorgD , " _
        & " Isnull(Sum ( CASE  WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.NMONTODESEMB End ),  0 ) SKOtorgS,  " _
        & " Isnull(Sum ( CASE  WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.NMONTODESEMB*" & pnTipoCambio & "  End ),  0 ) SKOtorgD  " _
        & " From " & psServConsol & "CreditoConsol  C " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & "," & gColocEstCancelado & "," & lsVig & ") " _
        & "  AND  C.DFECVIG BETWEEN '" & Format(pdFechaDesde, "mm/dd/yyyy") & "' AND '" & Format(pdFechaHasta, "mm/dd/yyyy") & " 23:59' " _
        & lsCond1(Det) & lsCond2(Det)
    
    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = Co.GetQuery(Sql)
    lnNroCreOtorgS = rs!NumOtorgS
    lnNroCreOtorgD = rs!NumOtorgD
    lnMonCreOtorgS = rs!SKOtorgS
    lnMonCreOtorgS = rs!SKOtorgS
    
    rs.Close

    Sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumFinMesS , " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumFinMesD , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '1' THEN (C.nSaldoCap) End ), 0 ) SKFinMesS , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '2' THEN (C.nSaldoCap * " & pnTipoCambio & ") End ), 0 ) SKFinMesD " _
        & " From " & psServConsol & "CreditoSaldoConsol C " _
        & " JOIN " & psServConsol & "CreditoConsol CC on C.cCtaCod = CC.cCtaCod " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & ") And Datediff(d,dFecha,'" & Format(pdFechaDesde, "mm/dd/yyyy") & "') = 0 " _
        & "  " _
        & lsCond1(Det) & Replace(lsCond2(Det), "C", "CC")

    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = Co.GetQuery(Sql)
    lnNroCreFMS = rs!NumFinMesS
    lnNroCreFMD = rs!NumFinMesD
    lnMonCreFMS = rs!SKFinMesS
    lnMonCreFMD = rs!SKFinMesD
    
    rs.Close

    Sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumCredS ,  " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumCredD , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '1' THEN (C.nSaldoCap) End ), 0 ) SKCredS, " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '2' THEN (C.nSaldoCap * " & pnTipoCambio & " ) End ), 0 ) SKCredD " _
        & " From " & psServConsol & "CreditoSaldoConsol C " _
        & " JOIN " & psServConsol & "CreditoConsol CC on C.cCtaCod = CC.cCtaCod " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & ") And Datediff(d,dFecha,'" & Format(pdFechaHasta, "mm/dd/yyyy") & "')=0" _
        & "  " _
        & lsCond1(Det) & Replace(lsCond2(Det), "C", "CC")
    
    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set rs = Co.GetQuery(Sql)
    lnNroCredS = rs!NumCredS
    lnNroCredD = rs!NumCredD
    lnMonCredS = rs!SKCredS
    lnMonCredD = rs!SKCredD
    
    rs.Close

    lnNroCreCancelS = lnNroCreFMS + lnNroCreOtorgS - lnNroCredS
    lnNroCreCancelD = lnNroCreFMD + lnNroCreOtorgD - lnNroCredD
    lnMonCreCancelS = lnMonCreFMS + lnMonCreOtorgS - lnMonCredS
    lnMonCreCancelD = lnMonCreFMD + lnMonCreOtorgD - lnMonCredD
    
    If Det = 1 Or Det = 3 Or Det = 7 Or Det = 10 Or Det = 11 Then
        lnFil = lnFil + 3
        lnCol = 1
        
        xlHoja1.Cells(lnFil, lnCol) = "Nro Cred. Vigentes " & Format(pdFechaDesde, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 1, lnCol) = "Nro Cred. Otorgados   " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 2, lnCol) = "Nro Cred. Cancelados  " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 3, lnCol) = "Nro Cred. Vigentes    " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 4, lnCol) = "Saldo Cred. Vigentes  " & Format(pdFechaDesde, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 5, lnCol) = "Monto Cred. Otorgados " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 6, lnCol) = "Monto Cred. Cancelados" & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 7, lnCol) = "Saldo Cred. Vigentes  " & Format(pdFechaHasta, "dd/mm/yyyy")
    End If
    
    xlHoja1.Cells(lnFil, lnCol + 1) = lnNroCreFMS
    xlHoja1.Cells(lnFil, lnCol + 2) = lnNroCreFMD
    xlHoja1.Cells(lnFil + 1, lnCol + 1) = lnNroCreOtorgS
    xlHoja1.Cells(lnFil + 1, lnCol + 2) = lnNroCreOtorgD
    xlHoja1.Cells(lnFil + 2, lnCol + 1) = lnNroCreCancelS
    xlHoja1.Cells(lnFil + 2, lnCol + 2) = lnNroCreCancelD
    xlHoja1.Cells(lnFil + 3, lnCol + 1) = lnNroCredS
    xlHoja1.Cells(lnFil + 3, lnCol + 2) = lnNroCredD
    xlHoja1.Cells(lnFil + 4, lnCol + 1) = lnMonCreFMS
    xlHoja1.Cells(lnFil + 4, lnCol + 2) = lnMonCreFMD
    xlHoja1.Cells(lnFil + 5, lnCol + 1) = lnMonCreOtorgS
    xlHoja1.Cells(lnFil + 5, lnCol + 2) = lnMonCreOtorgD
    xlHoja1.Cells(lnFil + 6, lnCol + 1) = lnMonCreCancelS
    xlHoja1.Cells(lnFil + 6, lnCol + 2) = lnMonCreCancelD
    xlHoja1.Cells(lnFil + 7, lnCol + 1) = lnMonCredS
    xlHoja1.Cells(lnFil + 7, lnCol + 2) = lnMonCredD

Next Det

xlHoja1.SaveAs App.Path & "\SPOOLER\INFO4.xls"
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
MsgBox "Se ha Generado el Archivo INFO4.XLS Satisfactoriamente", vbInformation, "Aviso"
Exit Sub

ErrorExcel:
    MsgBox "Error Nº [" & str(err.Number) & "] " & err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing

End Sub

Private Sub CmdInstitucion_Click()
    frmSelectAnalistas.SeleccionaInstituciones
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CmdSelecAge_Click()
Dim i As Integer
Dim nContAge As Integer

    frmSelectAgencias.Show 1
    ReDim MatAgencias(0)
    nContAge = 0
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            ReDim Preserve MatAgencias(nContAge)
            MatAgencias(nContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)
        End If
    Next i
End Sub

Private Sub CmdUbicacion_Click()
    sUbicacionGeo = frmUbicacionGeo.Inicio
End Sub

Private Sub Form_Load()
    Unload frmColRecReporte
    ReDim MatAgencias(0)
    ReDim MatProductos(0)
    ReDim matAnalista(0)
    ReDim MatInstitucion(0)

    Set Progress = New clsProgressBar
    Set Progreso = New clsProgressBar

    Dim oTipCambio As nTipoCambio
    

    
    Set oTipCambio = New nTipoCambio
    TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
    Set oTipCambio = Nothing
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Logo.AutoPlay = True
    Logo.Open App.Path & "\videos\LogoA.avi"
    CentraForm Me
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Unload frmSelectAgencias
    Unload frmSelectAnalistas
    Unload frmUbicacionGeo
    Set frmCredReportes = Nothing
End Sub
 
Private Sub lsRep_CloseProgress()
    Progreso.CloseForm Me
End Sub

Private Sub lsRep_Progress(pnValor As Long, pnTotal As Long)
    Progreso.Max = pnTotal
    Progreso.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub lsRep_ShowProgress()
    Progreso.ShowForm Me
End Sub




Private Sub oNCredDoc_FinalizaBarra()
Progress.CloseForm Me
End Sub

Private Sub oNCredDoc_IniciaBarra(ByVal lnTotal As Long)
Progress.Max = lnTotal
Progress.ShowForm Me
End Sub

Private Sub oNCredDoc_ProgresoBarra(ByVal i As Long, ByVal lsTitulo As String, ByVal lsSubtitulo As String)
Progress.Progress i, lsTitulo, lsSubtitulo, lsSubtitulo
End Sub

Private Sub optCredVig_Click(Index As Integer)
    CmdAnalistas.Visible = IIf(Index = 2, True, False)
    If Index = 7 Then
        FraACE.Enabled = False
    Else
        FraACE.Enabled = True
    End If
End Sub

Private Sub optEstadistica_Click(Index As Integer)
txtLineaCredito.Enabled = IIf(Index = 2, True, False)
txtLineaCredito.BackColor = IIf(Index = 2, &H80000005, &H8000000F)
txtLineaCredito.Text = ""
End Sub

Private Sub OptPagCheque_Click(Index As Integer)
    If Index = 0 Then
        TxtNroCheque.Enabled = False
    Else
        TxtNroCheque.Enabled = True
        TxtNroCheque.Text = ""
    End If
    
End Sub

Private Sub Text3_Change()

End Sub

Private Sub TreeView1_Click()
     ActivaDes TreeView1.SelectedItem
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
Dim nExpande As Integer '1 Si no deja expandir a todos 0 Deja expandir a todos
Dim nUnico As Integer 'Valor del unico que se puede expander

nExpande = val(Text2.Text)
nUnico = val(Text1.Text)

If nExpande = 1 Then
    If InStr(Text1.Text, Mid(Node.Key, 2, 1)) > 0 Then
        Node.Expanded = True
    End If
End If

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    
Dim nExpande As Integer '1 Si no deja expandir a todos 0 Deja expandir a todos
Dim nUnico As Integer 'Valor del unico que se puede expander

nExpande = val(Text2.Text)
nUnico = val(Text1.Text)

If nExpande = 1 Then
    If InStr(Text1.Text, Mid(Node.Key, 2, 1)) = 0 Then
        Node.Expanded = False
        Node.Checked = False
    End If
End If

End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    ActivaDes TreeView1.SelectedItem
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)

    Node.Selected = True
    ActivaDes Node
     
End Sub
 
Private Sub ActivaDes(sNode As Node)

    Dim i As Integer
    Dim nExpande As Integer

    nExpande = val(Text2.Text)
         
    If nExpande = 0 Then
'         If Mid(sNode.Key, 2, 1) <> Val(Text1.Text) Then
'            sNode.Checked = False
'         End If
        For i = 1 To TreeView1.Nodes.Count
            If Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(sNode.Key, 2, 1) And Mid(sNode.Key, 1, 1) = "P" Then
                TreeView1.Nodes(i).Checked = sNode.Checked
            End If
        Next
        
    ElseIf nExpande = 1 Then
        If InStr(Text1.Text, Mid(sNode.Key, 2, 1)) = 0 Then
            sNode.Checked = False
            sNode.Expanded = False
        Else
            TreeView1.SelectedItem = sNode
        Select Case Mid(sNode.Key, 1, 1)
        Case "P"
            If sNode.Checked = True Then
                 For i = 1 To TreeView1.Nodes.Count
                     If Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(sNode.Key, 2, 1) Then
                         TreeView1.Nodes(i).Checked = True
                     End If
                 Next
            Else
                 For i = 1 To TreeView1.Nodes.Count
                   If Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(sNode.Key, 2, 1) Then
                     TreeView1.Nodes(i).Checked = False
                   End If
                 Next
            End If

        End Select
        End If
    End If
End Sub

Private Sub ActFiltra(nFiltra As Boolean, Optional nFiltro As String = "")

    Dim i As Integer
    Dim nTempo As Integer
    
    If nFiltra = True Then
        Text2.Text = 1
        Text1.Text = nFiltro
         
        For i = 1 To TreeView1.Nodes.Count
            If InStr(nFiltro, Mid(TreeView1.Nodes(i).Key, 2, 1)) = 0 Then
                TreeView1.Nodes(i).Expanded = False
                TreeView1.Nodes(i).Checked = False
            Else
                TreeView1.Nodes(i).Expanded = True
                TreeView1.Nodes(i).Checked = False
            End If
        Next
        
    Else
        Text2.Text = ""
        Text1.Text = ""
        For i = 1 To TreeView1.Nodes.Count
            TreeView1.Nodes(i).Expanded = False
            TreeView1.Nodes(i).Checked = False
        Next
    End If
    
    
End Sub

Private Function GetProdsMarcados() As String
    Dim i As Integer
    Dim sCad As String
    
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                If Len(Trim(sCad)) = 0 Then
                    sCad = "'" & Mid(TreeView1.Nodes(i).Key, 2, 3)
                Else
                    sCad = sCad & "', '" & Mid(TreeView1.Nodes(i).Key, 2, 3)
                End If
            End If
        End If
    Next
    If Len(Trim(sCad)) > 0 Then
        sCad = "(" & sCad & "')"
    End If
                
    GetProdsMarcados = sCad

End Function

'***MAVM: Modulo de Auditoria 20/08/2008
' Para Mostrar seleccioando El Reporte de Garantias Inscritas
' por defecto en el Modulo de Auditoria
'Public Sub Inicializar_operacion(ByVal sOpeCod As Currency)
'    Limpia
'    TVRep.Nodes(Index108000).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108000).Expanded = True
'
'    TVRep.Nodes(Index108300).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108300).Expanded = True
'
'    TVRep.Nodes(Index108380).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108380).Expanded = True
'
'    TVRep.Nodes(Index108386).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108386).Expanded = True
'
'    TVRep.Enabled = False
'    TVRep.HideSelection = False
'End Sub
'***MAVM: Modulo de Auditoria 20/08/2008

'*** MAVM: Modulo de Auditoria 21/08/2008
' Para Mostrar seleccionado El Reporte de operaciones reprogramadas
' por defecto en el Modulo de Auditoria
'Public Sub Inicializar_OperacionesReprogramadas(ByVal sOpeCod As Currency)
'    TVRep.Nodes(Index108000OR).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108000OR).Expanded = True
'
'    TVRep.Nodes(Index108300OR).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108300OR).Expanded = True
'
'    TVRep.Nodes(Index108325OR).Selected = True
'    TVRep_Click
'    TVRep.Nodes(Index108325OR).Expanded = True
'
'    TVRep.Enabled = False
'    TVRep.HideSelection = False
'End Sub
'*** MAVM: Modulo de Auditoria 21/08/2008

'*** MAVM: Modulo de Auditoria 06/01/2010
' Para Mostrar seleccionado El Reporte de Creditos Rechazados
' Por Defecto en el Modulo de Auditoria

Public Sub Inicializar_CreditosRechazados(ByVal sOpeCod As Currency)
    TVRep.Nodes(Index108000CR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108000CR).Expanded = True
    
    TVRep.Nodes(Index108100CR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108100CR).Expanded = True
    
    TVRep.Nodes(Index108140CR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108140CR).Expanded = True
    
    TVRep.Nodes(Index108142CR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108142CR).Expanded = True
    
    TVRep.Enabled = False
    TVRep.HideSelection = False
End Sub
'*** MAVM: Modulo de Auditoria 06/01/2010

'*** MAVM: Modulo de Auditoria 06/01/2010
' Para Mostrar seleccionado El Reporte de Arqueos
' Por Defecto en el Modulo de Auditoria

Public Sub Inicializar_ReporteArqueos(ByVal sOpeCod As Currency)
    TVRep.Nodes(Index108000AR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108000AR).Expanded = True
    
    TVRep.Nodes(Index108200AR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108200AR).Expanded = True
    
    TVRep.Nodes(Index108203AR).Selected = True
    TVRep_Click
    TVRep.Nodes(Index108203AR).Expanded = True
    
    TVRep.Enabled = False
    TVRep.HideSelection = False
End Sub
'*** MAVM: Modulo de Auditoria 06/01/2010

Private Sub TVRep_Click()
Dim m As Control
Dim i As Integer
Dim sTipo As String
'ALPA 30/06/2008*****************
TxtDiaAtrIni.Visible = True
chkMigraExcell.Caption = "Migra a Excell"
'********************************
    Limpia
    Me.Caption = "Reportes de Créditos " & Mid(TVRep.SelectedItem.Text, 8, Len(TVRep.SelectedItem.Text) - 7)
    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
        Case gColCredRepIngxPagoCred
            Call HabilitaControleFrame1(True, True, True, False, , , , , , , , , , , , , , , , , , , , , , True)
            CmdSelecAge.Visible = True
        Case gColCredRepDesemEfect
            Call HabilitaControleFrame1(True, True, True, False, , , , , , , , , , , , , , , True)
            CmdSelecAge.Visible = True
        Case gColCredRepSalCarVig
            Call HabilitaControleFrame1(False, False, True, False)
        Case gColCredRepCredCancel
            Call HabilitaControleFrame1(True, True, True, True, , , , True)
        Case gColCredRepResSalCarxAna
            'Comentado por DAOR 20070717
            'Call HabilitaControleFrame1(False, False, True, False, True, True, , , , , , , , , , , , , True)
            Call HabilitaControleFrame1(False, False, False, False, True, True, , , , , , True, , , , , , , True, , , , , , , , , , , , , , , , , , True, True) 'Gitu 05/04/2008
        Case gColCredRepMoraInst
            Call HabilitaControleFrame1(True, False, True, False, True, True, True, , , , , , , , , , , , True)
        
        '(Se agrego una segunda opcion con una bandera)
        Case gColCredRepAtraPagoCuotaLib
            Call HabilitaControleFrame1(False, True, True, False, False, True, False, True, , , , , , , , , , , True)
        '''''''''''
        Case gColCredRepMoraxAna
            Call HabilitaControleFrame1(False, True, True, False, False, True, False, True, True, , True, , , , , , , , True, , , , , , , , , True, True, , True)
        Case gColCredRepCredProtes
            Call HabilitaControleFrame1(False, False, True, False, False, False, False, False)
        'Case gColCredRepCredRecha, gColCredRepCredAnula  'gColCredRepCredRetir
         '   Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, , , , , , , , , , , True)
        Case gColCredRepCredxUbiGeo
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, False, False, False, True, , , , , , , , True)
        
        '(Se agrego una segunda opcion con una bandera)
        Case gColCredRepCredVig, gColCredRepCredVigconCuoLibre
            
            Call HabilitaControleFrame1(False, True, True, False, False, True, False, False, True, False, False, True, , , , , IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, True, False), , True)
        '''''''''''''''''''''''
        Case gColCredRepCredSaldosDiarios
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, , , , , , , , , True)
            
        Case gColCredRepCredxInst
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, True, True, True)
        Case gColCredRepMoraxInst
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
' COMENTADO X MADM 20111001
'        Case 108411 'ARCV 28-10-2006
'            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, True, , , , , , True)
        Case 108407
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, , , , , , , , , , , , , True)
        Case 108408
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, True, , , , , , , , , , , , , , , , True)
        'MADM 20111001
        Case 108411
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, True, , , , , , , , , , , , , , , , True)
        'END MADM
        Case gColCredRepResSalCartxAna
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, False, False, False, False, False, False, False, True, True, , True)
        Case gColCredRepResSaldeCartxInst
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True, , True)
            'Ubica el Boton para Seleccionar la Institucion en la posicion Inferior
            'CmdInstitucion.Left = 1965
            'CmdInstitucion.Top = 5535
        Case gColCredRepLisDesctoPlanilla
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, True, True, True, , , , , , , , , , , , , , , , , True)
        Case gColCredRepPagosconCheque
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, True, True, False, False)
            CmdSelecAge.Visible = True
        Case gColCredRepPagosdeOtrasAgen
            Call HabilitaControleFrame1(True, True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepPagosEnOtrasAgen
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepIntEnSusp
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        Case 108505
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        
        ''''''''''''''''
        Case gColCredRepProgPagosxCuota
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True, True, True, , , , , , , , , , True)
            chkMigraExcell.Caption = "Consolidar"
        Case gColCredRepDatosReqMora
            Call HabilitaControleFrame1(False, False, True, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper
            Call HabilitaControleFrame1(False, True, True, False, False, False, False, True, False, False, False, True, False, False, False, False, False, False, True, False)
        Case gColCredRepConsColocxAgencia
            Call HabilitaControleFrame1(False, True, True, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, True, False)
        Case gColCredRepConsColocxFteFinan
            Call HabilitaControleFrame1(False, False, True, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, True, False)
        'WIOR 20140530 ****************************
        Case gColCredCorresponsaliaPorDebito
            Call HabilitaControleFrame1(False, True, False, False, , , , True, , , , , , , , , , , , , False)
        'WIOR FIN *********************************
        Case gColCredRepCartaCobMoro1, gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, _
             gColCredRepCartaCobMoro4, gColCredRepCartaCobMoro5, gColCredRepCartaRecup, gColCredRepCartaCobMoro6, gColCredRepCartaRecup2
             'WIOR AGREGO
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, True, False, True, False, False, False, False, False, False, False, True, , True, , , , , , , True)
        Case gColCredRepCartaInvCredAlt
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, True, True, False, False, False, False, False, False, False, False, True, , False)
        
        Case gColCredRepCredVigArqueo
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True, True, , , True, , , , True, True)

        Case gColCredRepVisitaCobroCuotas
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True, False, False, False)
            
        Case gColCredRepClientesNCuotasPend
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, False, True, False, False, False, False, False, False, False, False, True, , False)
            'FraUit.Caption = "Porcentaje"
        Case gColCredRepIngresosxGasto
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True)
        Case gColCredRepCredVigIntDeven
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, True, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepEstMensual
            Call HabilitaControleFrame1(True, True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True, False, True)
        Case gColCredRepCredDesmMayores
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, True, False, False, False, False, False, False, True, False, True, , , True)
        'ALPA 20080619 *********************************************************************************************************************************************************************************
        Case gColCredRepRenCarAnali
            'ALPA 20090324*****************************************************************
            'Call HabilitaControleFrame1(True, True, False, False, False, False, False, True, True, False, False, False, False, False, False, False, False, False, False, False, True, , , , , , , , , , True)
             Call HabilitaControleFrame1(True, True, False, False, False, False, False, True, True, False, False, False, False, False, False, False, False, False, True, False, True, , , , , , , , , , True)
             '*****************************************************************************
            TxtDiaAtrIni.Visible = False
            chkMigraExcell.Caption = "Consolidar"
        'ALPA 20080625 *********************************************************************************************************************************************************************************
         Case gColCredCanXNAtendidos
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, True, True, False, False, False, False, False, False, False, False, False, False, False, True)
             TxtDiaAtrIni.Visible = False
             
        
       '***********************************************************************************************************************************************************************************************
'        Case 108308
'            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, , , False)
            'FraTipCambio.Visible = False
      '***********************************
        Case 108309
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, , , False)
        Case 108701:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108702:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108703:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108704:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108705:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108706:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108707:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108708:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        'Case 108709:
        Case 108710:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108711:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108712:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108713:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        '********************************************************'
        '** GITU 20081007
        Case 108337:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        '********************************************************'
        Case 108714:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108715:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108721:
                Call HabilitaControleFrame1(False, False, False, False, False, True, False, False, True, False, False, True, False, False, False, False, False, False, True, False, True, False, False, False)
                'ActFiltra True, Mid(Producto.gColPYMEEmp, 1, 1)
                ActFiltra True, "3,4"
        Case 108722:
                Call HabilitaControleFrame1(False, False, False, False, False, True, False, False, True, False, False, True, False, False, False, False, True, False, True, False, True, False, False, False)
        Case 108723:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False)
        Case 108724:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, True, False, False, False)
        Case 108725:
               Call HabilitaControleFrame1(True, True, True, False, False, False, False, True _
                                           , False, False, False, False, False, False, False, False, False, False, True, False, False, False, False)
                CmdSelecAge.Visible = True
        Case 108801:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108802:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108803:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108804:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108806:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108808:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
            
        Case 108810
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108307
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
            FraTipCambio.Visible = False
        Case 108109
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False)
        Case 108110:
              ' reporte de protesto al dia
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, True, False, False, False)
        Case 108111
              'reporte de consolidado por analista
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, True, False, False, False, True, False, False, False, False, False, False, True, False, True, False, False, False, False)
         Case 108114
              'reporte de consolidado por analista
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108310
             ' reporte de creditos automaticos
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, True)
        Case 108205
             ' reporte de clientes vigentes auditoria
             Call HabilitaControleFrame1(True, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108311
             ' reporte de creditos no desembolsados
             Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False)
        Case 108206
            'Arqueo de Clientes por Analista
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False)
        Case 108112
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, True, False, False, False, False)
        Case 108208
            'reporte de calidad de cartera
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108113
             'reporte de mora por analista y telefonno
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, True, False, False, True, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False)
        Case 108409
            Call HabilitaControleFrame1(True, True, True, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False)
        Case 108410
            Call HabilitaControleFrame1(True, True, True, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False)
        Case 108115 'Reporte de Consolidado por Analista
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False)
        Case gColCredRepCredSaldosDiarios
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False)
        Case 108311 'Reporte de Lista de Creditos no Desembolsados
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False)
         Case 108508 'ALPA Reporte de Contabilidad 20080616
         
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
            
        Case 108506 'Reporte de Estadisticas RFC/DIF
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
           
        Case 108117
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False)
        
        'ARCV 03-02-2007
        Case 108321
            Call HabilitaControleFrame1(False, False, False, False, , , , True, , , , , , , , , , , , , True, , , , , , , , , , , , True)
        '05-02-2007
        
        'peac 20071228 se agrego "108329" reporte de creditos comerciales vigentes con fecha EEFF
        Case 108322, 108381, 108382, 108384, 108385, gColCredRepComVigEEFF
            Call HabilitaControleFrame1(False, False, False, False, , , , , , , , , , , , , , , , , True)
        'by Capi 28122007
        Case gColCredRepGarantiasInscritas
            Call HabilitaControleFrame1(False, True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True, False, False, False, , , , , , , , , , , , , , , True)
        '07-02-2007
        Case 108323
            Call HabilitaControleFrame1(False, False, False, False, , , , , , , , , , , , , , , , , False)
        Case 108383, 108325, 108324
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , , , True)
        Case gColCredRepCreditosDesBcoBac 'DAOR 20070313, Creditos Desembolsados en Agencias del Banco de la Nación
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , , , True, , , , , , , , , , True)
        Case gColCredRepCreditosAprExoReglamento 'DAOR 20070418, Creditos Desembolsados Con Aprobación de Exoneración de Reglamento
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , , , True, , , , , , , , , , False)
        Case gColCredRepMontoDesembolsadoPorLineas 'DAOR 20070419, Monto Desembolsados por Lineas
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , , , True, , , , , , , , , , False)
        '-------
        Case gColCredRepEstadosCuentaCredito 'DAOR 20070717, Estados de cuenta de crédito
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , True, , True, , , , , , , , , , False)

        Case gColCredRepREULavadoDinero 'By Capi 28012008, REU Lavado Dinero
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , , , , , , , , , , , , , False)


        Case gColCredRepDUDLavadoDinero 'By Capi 30012008, REU Lavado Dinero
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , , , , , , , , , , , , , False)

            
        Case gColCredRepConCarxAnalista 'DAOR 20070717, Consolidado de cartera por analista
            Call HabilitaControleFrame1(True, True, False, False, False, False, , , , , , True, , , , , , , True)
        Case gColCredRepResSalCarxAnaConsolida 'DAOR 20070814
            Call HabilitaControleFrame1(False, False, False, False, True, True, , , , , , True, , , , , , , True, , , , , , , , , , , , , , , , , , True, True) 'Gitu 03-04-2008
        Case gColCredRepConCarxAnalistaConsolida 'DAOR 20070814
            Call HabilitaControleFrame1(True, True, False, False, False, False, , , , , , True, , , , , , , True)
        Case gColCredRepSeguroDesgravConsolida 'MADM 20110329 - DAOR 20071210, Reporte de Seguro de Desgravamen
            'Call HabilitaControleFrame1(False, True, False, False, False, False, , , , , , True, , , , , , , False, , False)
            Call HabilitaControleFrame1(True, True, True, False, False, False, , , , , , True, , , , , , , False, , False, , , , , , , , , , , , , , , , , , , , , True)
            TxtFecFinA02.Text = gdFecData
            TxtFecIniA02.Text = gdFecData
        Case gColCredRepActaComiteCredAprobados 'PEAC 20070822 reporte acta de comite creditos aprobados
            Call HabilitaControleFrame1(True, True, False, False, , , , , , , , , , , , , , , True, , False)
        Case gColCredRepXTipoCondProd 'peac 20070822 reporte por tipo de condicion de producto
            Call HabilitaControleFrame1(True, True, True, False, False, False, , True, , , , True, , , , , , , True, , False)
        Case gColCredRepDetalleCuotasXCobrar 'peac 20070905 reporte de detalle de cuotas por cobrar
            Call HabilitaControleFrame1(True, True, True, False, False, False, , True, , , , False, , , , , , , True, , False)
        Case gColCredRepCreditosXCampanha 'MADM 20100721 FraCondicion -- peac 20070905 reporte de creditos por campañas y/o productos
            Call HabilitaControleFrame1(True, True, True, False, False, True, , True, , , , False, , , , , , , True, , True, , , , , , , , , , , , , True, , , , , , True)
        Case gColCredRepConvCasillero  'peac 20070921 reporte de cobranza diaria convenios (casillero)
            Call HabilitaControleFrame1(True, True, False, False, False, False, , False, , , , False, , , True, , , , False, , True, , , , , , , , , , , , , False, True)
        Case 108125 'gColRecRepCobranzaGestores ' 108125, PEAC 20070917
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, True, , False, , , , , , , , False, , True, , , , , , , False, False)
        
        Case gColCredRepPoliIncendio 'peac 20071229 reporte de polizas contra insendio
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, , False, , , , , , , , True, , True, , , , , , , False, False)

        Case gColCredRepDiasTranscDesdeSoli 'PEAC 20080215 108141
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, True, , False, , , , , , , , False, , True, , , , , , , False, False)
        Case gColCredRepSoliRechazadas 'PEAC 20080215 108142
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, , False, , , , , , , , False, , True, , , , , , , False, False)
        Case gColCredRepSoloEnEstadoSoli 'PEAC 20080215 108143
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, , False, , , , , , , , False, , True, , , , , , , False, False)
        Case gColCredRepSoliProcesadas 'PEAC 20080215 108144
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, , False, , , , , , , , False, , True, , , , , , , False, False)

        Case gColCredRepCredVencPaseCastigo '*** PEAC 20080219 108335
            'Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, True, , False, True, , , , , , , False, , True, , , , , , , False, False, , , , , , , True)
            'MADM 20110515 FRAMONTOS + UIT
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, True, , False, True, , , , , , , False, , True, , , , , , , False, True, , , , , , , True)

        Case gColCredRepCliConDistinTiposCred '***PEAC 20080221 108336
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, , False, False, , , , , , , True, , True, , , , , , , False, False, , , , , , , False)

        Case gColCredRepAECIReporte01 '***PEAC 20080303 108921
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, , False, False, , , , , , , False, , False, , , , , , , False, False, , , , , , , False)

        Case gColCredRepClientesHisNegativo '***PEAC 20080804 108350
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, , False, False, , , , , , , False, , True, , , , , , , False, False, , , , , , , False)
        
        Case gColCredRepClientesPotencialesSinCredVig '*** PEAC 20080923 108360
          Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, , False, False, , , , , , , False, , True, , , , , , , False, False, , , , , , , False)
        
        Case gColCredRepNumYSaldoCredPorProductoConsol '*** PEAC 20080924 108210
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, , False, True, , , , , , , False, , False, , , , , , , False, False, , , , , , , False)
        
        'By Capi Planeamiento Set 07
        Case gColCredRepFepInforme01
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gCapCredRepFepInforme02
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme03
            Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme3a
            Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme3b
            Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme3c
            Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme3d
            Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme04
           Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepInforme06
           Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepFepEntorno
           Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        'End By
        Case 108730:
           Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        Case 108731:
           Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        'ALPA***20080825********************************************************************************************
        Case gColCredRepCredAdjuContabi:
           Call HabilitaControleFrame1(False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False)
        'By Capi 03112008
         Case gColRepCredRefinan
            Call HabilitaControleFrame1(False, True, False, False, , , , , True, , , True, , , , , , , , , True)
             TxtDiaAtrIni.Visible = False
        
         'By MAVM 16052009
         Case 108370 'Reporte de Creditos Vencidos Mayores a 90 Dias
            TxtDiaAtrIni.Text = 91
            Call HabilitaControleFrame1(True, True, False, False, False, False, False, False, True, , False, True, , , , , , , False, , True, , , , , , , False, False)
        Case 108371 'Reporte de Creditos Vencidos Mayores a 90 Dias
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, , False, False, , , , , , , True, , True, , , , , , , False, False)
                
        'MAVM 20100511 ***
        Case gColCredRepResXAnalistaBPPR 'Reporte de Resultado de Analista BPPR
            CargarAgencias
            CargarCartera
            'Dim oAcceso As COMDPersona.UCOMAcceso
            'Set oAcceso = New COMDPersona.UCOMAcceso
            'If ValidarGrupo(gsCodUser, gsDominio) = True Then
                'dcAgencia.Enabled = True
            'Else
                'dcAgencia.Enabled = False
            'End If
            txtFCierre.Text = gdFecData
            Call HabilitaControleFrame1(True, True, False, False, False, , , , , , , True, , , , , , , , , False, , , , , , , , , , , , , , , , , , , , True)
            
        Case gColCredRepBonificacionXAnalistaBPPR 'Reporte de Bonificacion de Analista BPPR
            CargarAgencias
            CargarCartera
            'Dim oAcceso As COMDPersona.UCOMAcceso
            'Set oAcceso = New COMDPersona.UCOMAcceso
            'If ValidarGrupo(gsCodUser, gsDominio) = True Then
                'dcAgencia.Enabled = True
            'Else
                'dcAgencia.Enabled = False
            'End If
            txtFCierre.Text = gdFecData
            Call HabilitaControleFrame1(True, True, False, False, False, , , , , , , True, , , , , , , , , False, , , , , , , , , , , , , , , , , , , , True)
                    
        'MAVM ***
        
        '***********************************************************************************************************
        Case Else
            Me.Caption = "Reportes de Creditos "
    End Select
End Sub

Private Sub Limpia()

Dim i As Integer
    
    Call HabilitaControleFrame1(False, False, False, False)
    
    Text2.Text = ""
    Text1.Text = ""
    For i = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(i).Checked = False
        'TreeView1.Nodes(i).Expanded = False
    Next
     
    TxtFecIniA02.Text = Format(gdFecSis, "dd/MM/YYYY")
    TxtFecFinA02.Text = Format(gdFecSis, "dd/MM/YYYY")
     
    TxtDiaAtrIni.Text = 0
    TxtDiasAtrFin.Text = 999
    optCredVig(0).value = True
    optEstadistica(0).value = True
   
 End Sub

Private Sub TVRep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TVRep_Click
End If
End Sub

Private Sub TVRep_NodeClick(ByVal Node As MSComctlLib.Node)
    TVRep_Click
End Sub
 
Private Function Genera_Reporte108306(ByVal cSubTit As String, ByVal psMoneda As String, ByVal psProducto As String, ByVal psAnalistas As String) As String
  
    Dim matFilas() As Long
    Dim matCont As Long
    
    Dim nFila As Long
    Dim i As Long
    
    Dim sTempoAnalista As String
    Dim sTempoMoneda As String
    
    Dim loExc As DCredReporte
    Dim reg As New ADODB.Recordset
    
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Set loExc = New DCredReporte
    Set reg = loExc.RecuperaListadoMorosos(psMoneda, psProducto, psAnalistas)
    If reg.BOF Then
        Genera_Reporte108306 = ""
        Exit Function
    Else
        Genera_Reporte108306 = "Reporte_Generado"
        lsArchivoN = App.Path & "\Spooler\SeguimientoMora" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
       
        lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            'Abro...
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
            
            sTempoAnalista = reg!Analista
            sTempoMoneda = reg!nmoneda
            matCont = 0
            nFila = 1
            
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
             
            nFila = 3
            xlHoja1.Cells(nFila, 1) = "REPORTE DE SEGUIMIENTO DE MORA"
            nFila = 4
            xlHoja1.Cells(nFila, 1) = cSubTit
            nFila = 5
            xlHoja1.Cells(nFila, 1) = "Credito"
            xlHoja1.Cells(nFila, 2) = "Cliente"
            xlHoja1.Cells(nFila, 3) = "Direccion"
            xlHoja1.Cells(nFila, 4) = "Zona"
            xlHoja1.Cells(nFila, 5) = "Telefono"
            xlHoja1.Cells(nFila, 6) = "Saldo Capital"
            xlHoja1.Cells(nFila, 7) = "Dias Atraso"
            xlHoja1.Cells(nFila, 8) = "Direc. Fuente Ingreso"
            xlHoja1.Cells(nFila, 9) = "Zona Fuente Ingreso"
            xlHoja1.Cells(nFila, 10) = "Telef. Fuente Ingreso"
            xlHoja1.Cells(nFila, 11) = "Nombre Garante"
            xlHoja1.Cells(nFila, 12) = "Direc. Garante"
            xlHoja1.Cells(nFila, 13) = "Zona Garante"
            
            nFila = nFila + 2
            ReDim Preserve matFilas(0) As Long
            matFilas(0) = nFila
            xlHoja1.Cells(nFila, 1) = "ANALISTA"
            xlHoja1.Cells(nFila, 2) = PstaNombre(reg!cNomAnalista)
            xlHoja1.Cells(nFila, 3) = reg!Analista
            nFila = nFila + 2
            
            xlHoja1.Cells(nFila, 1) = "MONEDA"
            xlHoja1.Cells(nFila, 2) = reg!cmoneda
            xlHoja1.Range("A" & Trim(str(nFila)) & ":C" & Trim(str(nFila))).Font.Bold = True
            With reg
                Do While Not reg.EOF
                    If sTempoAnalista <> !Analista Then
                        nFila = nFila + 2
                        matCont = matCont + 1
                        ReDim Preserve matFilas(matCont) As Long
                        matFilas(matCont) = nFila
                        xlHoja1.Cells(nFila, 1) = "ANALISTA"
                        xlHoja1.Cells(nFila, 2) = PstaNombre(reg!cNomAnalista)
                        xlHoja1.Cells(nFila, 3) = reg!Analista
                        nFila = nFila + 2
                        xlHoja1.Cells(nFila, 1) = "MONEDA"
                        xlHoja1.Cells(nFila, 2) = !cmoneda
                        xlHoja1.Range("A" & Trim(str(nFila)) & ":C" & Trim(str(nFila))).Font.Bold = True
                        sTempoAnalista = !Analista
                        sTempoMoneda = !nmoneda
                    ElseIf sTempoMoneda <> !nmoneda Then
                        nFila = nFila + 2
                        xlHoja1.Cells(nFila, 1) = "MONEDA"
                        xlHoja1.Cells(nFila, 2) = !cmoneda
                        xlHoja1.Range("A" & Trim(str(nFila)) & ":C" & Trim(str(nFila))).Font.Bold = True
                        sTempoMoneda = !nmoneda
                    End If
                    nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = !cCtaCod
                    xlHoja1.Cells(nFila, 2) = PstaNombre(!cPersNombre, False)
                    xlHoja1.Cells(nFila, 3) = !cPersDireccDomicilio
                    xlHoja1.Cells(nFila, 4) = !cUbiGeoDescripcion
                    xlHoja1.Cells(nFila, 5) = CStr(IIf(IsNull(!cPersTelefono), "", !cPersTelefono))
                    xlHoja1.Cells(nFila, 6) = Format(!nSaldo, "#,##0.00")
                    xlHoja1.Cells(nFila, 7) = !nDiasAtraso
                    xlHoja1.Cells(nFila, 8) = !cDirFteIngreso
                    xlHoja1.Cells(nFila, 9) = !cZonaFteIngreso
                    xlHoja1.Cells(nFila, 10) = IIf(IsNull(!cFonoFteIngreso), "", !cFonoFteIngreso)
                    xlHoja1.Cells(nFila, 11) = PstaNombre("" & !cNomGarante, False)
                    xlHoja1.Cells(nFila, 12) = IIf(IsNull(!cDirGarante), "", !cDirGarante)
                    xlHoja1.Cells(nFila, 13) = IIf(IsNull(!cZonaGarante), "", !cZonaGarante)
                    
                    .MoveNext
                Loop
            End With
            reg.Close
            Set reg = Nothing
        
            xlHoja1.Range("A1:B1").MergeCells = True
            xlHoja1.Range("A3:M3").MergeCells = True
            xlHoja1.Range("A4:M4").MergeCells = True

            xlHoja1.Range("A1:B3").Font.Bold = True
            xlHoja1.Range("A4").Font.Bold = True
                        
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
            xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter
            
            For i = 0 To matCont ' - 1
                 xlHoja1.Range("A" & Trim(str(matFilas(i))) & ":C" & Trim(str(matFilas(i)))).Font.Bold = True
                 xlHoja1.Range("A" & Trim(str(matFilas(i))) & ":C" & Trim(str(matFilas(i)))).Interior.ColorIndex = 24
            Next
            
            With xlHoja1.Range("A5:M5")
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.ColorIndex = 0
                .Interior.ColorIndex = 19
            End With
             
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
        
            'Cierro...
            OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
            OleExcel.SourceDoc = lsArchivoN
            OleExcel.Verb = 1
            OleExcel.Action = 1
            OleExcel.DoVerb -1
        End If
    End If
End Function

Private Function Genera_Reporte108607(ByVal cSubTit As String, ByVal pnTipoCambio_ As Currency, ByVal pdFechaFin_ As String, ByVal psMoneda_ As String, ByVal psProductos_ As String, ByVal psAgencias_ As String, ByVal psAnalistas_ As String) As String
     
    Dim loExc As DCredReporte
    Dim reg As New ADODB.Recordset
    
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Dim cMatAgencia() As String
    Dim matAgencia() As String
    Dim cMatAnalista() As String
    Dim matCarNor1() As Long
    Dim matCarNor2() As Currency
    Dim matCarVen1() As Long
    Dim matCarVen2() As Currency
    Dim matCarRef1() As Long
    Dim matCarRef2() As Currency
'    Dim matCarRVen1() As Long
'    Dim matCarRVen2() As Currency
    Dim matCobJud1() As Long
    Dim matCobJud2() As Currency
    Dim matDesemNue1() As Long
    Dim matDesemNue2() As Currency
    Dim matDesemRepre1() As Long
    Dim matDesemRepre2() As Currency
    Dim matOpeRef1() As Long
    Dim matOpeRef2() As Currency
    Dim nContador As Long
    Dim i As Long
    Dim j As Long
    
    Dim sFilaTitulos() As Long 'En este arreglo se graba el numero de la fila en donde escribir los titulos de las agencias
    Dim sFilaSubTotal() As Long 'En este arreglo se graba el numero de filas en donde se deberan llenar con subtotales
    Dim sFilaTextoSubTotal() As String 'En este arreglo se graba el texto de el numero de la filas que se incluiran en las formulas del subtotal
    Dim sFilaTotal As Long 'Fila del total general
    Dim nContadorSubFila As Long 'cuantas subfilas existen
    Dim nFila As Long 'La fila actual
    Dim sTextoTotal As String 'el texto donde se grabara las filas a sumar en el total
    
    Dim sTempoAgencia As String
    Dim sTempoAnalista As String
    
    nContador = 0
    Set loExc = New DCredReporte
    Set reg = loExc.Recupera_ConsolidadoCarteraxAnalista(pnTipoCambio_, pdFechaFin_, psMoneda_, psProductos_, psAgencias_, psAnalistas_)
    If reg.BOF Then
        Genera_Reporte108607 = ""
        Exit Function
    Else
        Genera_Reporte108607 = "Reporte Generado"
        nContador = 0
        sTempoAgencia = reg!CAgencia
        sTempoAnalista = IIf(IsNull(reg!cAnalista), "", reg!cAnalista)
        
        ReDim Preserve cMatAgencia(nContador) As String
        ReDim Preserve matAgencia(nContador) As String
        ReDim Preserve cMatAnalista(nContador) As String
        ReDim Preserve matAnalista(nContador) As String
        ReDim Preserve matCarNor1(nContador) As Long
        ReDim Preserve matCarNor2(nContador) As Currency
        ReDim Preserve matCarVen1(nContador) As Long
        ReDim Preserve matCarVen2(nContador) As Currency
'        ReDim Preserve matCarRVen1(nContador) As Long
'        ReDim Preserve matCarRVen2(nContador) As Currency
        ReDim Preserve matCarRef1(nContador) As Long
        ReDim Preserve matCarRef2(nContador) As Currency
        ReDim Preserve matCobJud1(nContador) As Long
        ReDim Preserve matCobJud2(nContador) As Currency
        ReDim Preserve matDesemNue1(nContador) As Long
        ReDim Preserve matDesemNue2(nContador) As Currency
        ReDim Preserve matDesemRepre1(nContador) As Long
        ReDim Preserve matDesemRepre2(nContador) As Currency
        ReDim Preserve matOpeRef1(nContador) As Long
        ReDim Preserve matOpeRef2(nContador) As Currency
        
        cMatAnalista(0) = IIf(IsNull(reg!cAnalista), "", reg!cAnalista)
        cMatAgencia(0) = reg!CAgencia
        matAgencia(0) = reg!cDesAgencia
        matAnalista(0) = reg!cNomAnalista
        
        Do While Not reg.EOF
            If sTempoAgencia <> reg!CAgencia Or sTempoAnalista <> reg!cAnalista Then
                nContador = nContador + 1
                ReDim Preserve cMatAgencia(nContador) As String
                ReDim Preserve matAgencia(nContador) As String
                ReDim Preserve cMatAnalista(nContador) As String
                ReDim Preserve matAnalista(nContador) As String
                ReDim Preserve matCarNor1(nContador) As Long
                ReDim Preserve matCarNor2(nContador) As Currency
                ReDim Preserve matCarVen1(nContador) As Long
                ReDim Preserve matCarVen2(nContador) As Currency
                ReDim Preserve matCarRef1(nContador) As Long
                ReDim Preserve matCarRef2(nContador) As Currency
'                ReDim Preserve matCarRVen1(nContador) As Long
'                ReDim Preserve matCarRVen2(nContador) As Currency
                ReDim Preserve matCobJud1(nContador) As Long
                ReDim Preserve matCobJud2(nContador) As Currency
                ReDim Preserve matDesemNue1(nContador) As Long
                ReDim Preserve matDesemNue2(nContador) As Currency
                ReDim Preserve matDesemRepre1(nContador) As Long
                ReDim Preserve matDesemRepre2(nContador) As Currency
                ReDim Preserve matOpeRef1(nContador) As Long
                ReDim Preserve matOpeRef2(nContador) As Currency
                
                cMatAnalista(nContador) = reg!cAnalista
                cMatAgencia(nContador) = reg!CAgencia
                matAgencia(nContador) = reg!cDesAgencia
                matAnalista(nContador) = reg!cNomAnalista
                
                sTempoAgencia = reg!CAgencia
                sTempoAnalista = reg!cAnalista
            End If
            
            If reg!Lugar = 1 Then
                'Saldo de Cartera Normal
                matCarNor1(nContador) = matCarNor1(nContador) + reg!Cantidad
                matCarNor2(nContador) = matCarNor2(nContador) + reg!Total
            ElseIf reg!Lugar = 2 Then
                'Saldo de Cartera Vencida
                matCarVen1(nContador) = matCarVen1(nContador) + reg!Cantidad
                matCarVen2(nContador) = matCarVen2(nContador) + reg!Total
            ElseIf reg!Lugar = 3 Then
                'Saldo de Cartera Refinanciada
                matCarRef1(nContador) = matCarRef1(nContador) + reg!Cantidad
                matCarRef2(nContador) = matCarRef2(nContador) + reg!Total
            ElseIf reg!Lugar = 4 Then
                'Cobranza Judicial
                matCobJud1(nContador) = matCobJud1(nContador) + reg!Cantidad
                matCobJud2(nContador) = matCobJud2(nContador) + reg!Total
            ElseIf reg!Lugar = 5 Then
                'Desembolsos Nuevos
                matDesemNue1(nContador) = matDesemNue1(nContador) + reg!Cantidad
                matDesemNue2(nContador) = matDesemNue2(nContador) + reg!Total
            ElseIf reg!Lugar = 6 Then
                'Desembolsos Represtados
                matDesemRepre1(nContador) = matDesemRepre1(nContador) + reg!Cantidad
                matDesemRepre2(nContador) = matDesemRepre2(nContador) + reg!Total
            ElseIf reg!Lugar = 7 Then
                'Operaciones Refinanciadas
                matOpeRef1(nContador) = matOpeRef1(nContador) + reg!Cantidad
                matOpeRef2(nContador) = matOpeRef2(nContador) + reg!Total
'            ElseIf reg!Lugar = 8 Then
'                matCarRVen1(nContador) = matCarRVen1(nContador) + reg!Cantidad
'                matCarRVen2(nContador) = matCarRVen2(nContador) + reg!Total
            End If
            reg.MoveNext
        Loop
        reg.Close
        Set reg = Nothing


         
        lsArchivoN = App.Path & "\Spooler\ConsolCarteraxAnalista" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            'Abro...
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
    
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
    
            nFila = 1
    
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = gsNomAge
    
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "RESUMEN CONSOLIDADO DE CARTERA POR ANALISTA"
    
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "Al " & pdFechaFin_
             
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = cSubTit
            xlHoja1.Cells(nFila, 23) = "T.C.F.= " & Format(pnTipoCambio_, "#,##0.00")
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "Funcionario"
            xlHoja1.Cells(nFila, 2) = "Saldo de"
            xlHoja1.Cells(nFila, 4) = "Saldo de"
            xlHoja1.Cells(nFila, 7) = "Saldo de"
            xlHoja1.Cells(nFila, 10) = "Cobranza"
            xlHoja1.Cells(nFila, 12) = "Resultados Mensuales"
            xlHoja1.Cells(nFila, 17) = "Desembolsos"
            xlHoja1.Cells(nFila, 21) = "Total"
            xlHoja1.Cells(nFila, 23) = "Operaciones"
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "de"
            xlHoja1.Cells(nFila, 2) = "Cartera Vigente"
            xlHoja1.Cells(nFila, 4) = "Cartera Vencida"
            xlHoja1.Cells(nFila, 7) = "Cartera Refinanciada"
            xlHoja1.Cells(nFila, 10) = "Judicial"
            xlHoja1.Cells(nFila, 12) = "Saldo de"
            xlHoja1.Cells(nFila, 13) = "Saldo de"
            xlHoja1.Cells(nFila, 14) = "Indice"
            xlHoja1.Cells(nFila, 15) = "Saldo"
            xlHoja1.Cells(nFila, 16) = "Indice"
            xlHoja1.Cells(nFila, 17) = "Nuevos"
            xlHoja1.Cells(nFila, 19) = "Represtamos"
            xlHoja1.Cells(nFila, 21) = "Desembolso"
            xlHoja1.Cells(nFila, 23) = "Refinanciadas"
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "Credito"
            xlHoja1.Cells(nFila, 2) = "Nro"
            xlHoja1.Cells(nFila, 3) = "Saldo"
            xlHoja1.Cells(nFila, 4) = "Nro"
            xlHoja1.Cells(nFila, 5) = "Saldo"
            xlHoja1.Cells(nFila, 6) = "%"
            xlHoja1.Cells(nFila, 7) = "Nro"
            xlHoja1.Cells(nFila, 8) = "Saldo"
            xlHoja1.Cells(nFila, 9) = "%"
            xlHoja1.Cells(nFila, 10) = "Nro"
            xlHoja1.Cells(nFila, 11) = "Saldo"
            xlHoja1.Cells(nFila, 12) = "Cartera"
            xlHoja1.Cells(nFila, 13) = "Mora"
            xlHoja1.Cells(nFila, 14) = "Mora"
            xlHoja1.Cells(nFila, 15) = "C.A.R."
            xlHoja1.Cells(nFila, 16) = "C.A.R."
            xlHoja1.Cells(nFila, 17) = "Nro"
            xlHoja1.Cells(nFila, 18) = "Saldo"
            xlHoja1.Cells(nFila, 19) = "Nro"
            xlHoja1.Cells(nFila, 20) = "Saldo"
            xlHoja1.Cells(nFila, 21) = "Nro"
            xlHoja1.Cells(nFila, 22) = "Saldo"
            xlHoja1.Cells(nFila, 23) = "Nro"
            xlHoja1.Cells(nFila, 24) = "Saldo"
             
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = cMatAgencia(0) & " " & matAgencia(0)
            
            sTempoAgencia = cMatAgencia(0)
             
            Dim sRef As Byte
            sRef = 1
            nContadorSubFila = 0
            
            'Dimensiono los arreglos para las variables y el subtotal
            ReDim Preserve sFilaSubTotal(nContadorSubFila) As Long
            ReDim Preserve sFilaTitulos(nContadorSubFila) As Long
            ReDim Preserve sFilaTextoSubTotal(nContadorSubFila) As String
            
            'Grabo la fila del primer titulo
            sFilaTitulos(nContadorSubFila) = nFila
            
            For i = 0 To nContador
                If sTempoAgencia <> cMatAgencia(i) Then
                    'Agregar fila en blanco que diga subtotal
                    nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = "SubTotal"
                    'Actualizo el valor de la fila para el subtotal
                    sFilaSubTotal(nContadorSubFila) = nFila
                     
                    nContadorSubFila = nContadorSubFila + 1
                    ReDim Preserve sFilaSubTotal(nContadorSubFila) As Long
                    ReDim Preserve sFilaTextoSubTotal(nContadorSubFila) As String
                       
                    sTempoAgencia = cMatAgencia(i)
                    'Grabo la fila en la que se encuentra el titulo
                    ReDim Preserve sFilaTitulos(nContadorSubFila) As Long
                    
                    'Agrego fila en blanco
                    nFila = nFila + 1
                    
                    'Agrego fila que diga el nombre de la agencia
                    nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = cMatAgencia(i) & " " & matAgencia(i)
                    sFilaTitulos(nContadorSubFila) = nFila
                      
                End If
                 
                nFila = nFila + 1
                
                'LLeno la posicion de cada una de las celdas a sumar en la formula de subtotal
                sFilaTextoSubTotal(nContadorSubFila) = sFilaTextoSubTotal(nContadorSubFila) & "+*" & Trim(str(nFila))
                
                'Imprimo la fila con los valores normales
                xlHoja1.Cells(nFila, 1) = cMatAnalista(i) & " " & matAnalista(i)
                xlHoja1.Cells(nFila, 2) = Format(matCarNor1(i), "#,##0")
                xlHoja1.Cells(nFila, 3) = Format(matCarNor2(i), "#,##0.00")
                xlHoja1.Cells(nFila, 4) = Format(matCarVen1(i), "#,##0")
                xlHoja1.Cells(nFila, 5) = Format(matCarVen2(i), "#,##0.00")
                
                '(6)=(5)/(3) F=E/C
                If matCarNor2(i) <> 0 Then
                    xlHoja1.Range("F" & Trim(str(nFila))).Formula = "=$E$" & Trim(str(nFila)) & "/$C$" & Trim(str(nFila))
                Else
                    xlHoja1.Cells(nFila, 6) = 0
                End If
            
                xlHoja1.Cells(nFila, 7) = Format(matCarRef1(i), "#,##0")
                xlHoja1.Cells(nFila, 8) = Format(matCarRef2(i), "#,##0.00")
                
                '(9)=(8)/(3) I=H/C
                If matCarNor2(i) <> 0 Then
                    xlHoja1.Range("I" & Trim(str(nFila))).Formula = "=$H$" & Trim(str(nFila)) & "/$C$" & Trim(str(nFila))
                Else
                    xlHoja1.Cells(nFila, 9) = 0
                End If
                
                xlHoja1.Cells(nFila, 10) = Format(matCobJud1(i), "#,##0")
                xlHoja1.Cells(nFila, 11) = Format(matCobJud2(i), "#,##0.00")
                
                '(12)=(3)+(11) L=C+K
                xlHoja1.Range("L" & Trim(str(nFila))).Formula = "=$C$" & Trim(str(nFila)) & "+$K$" & Trim(str(nFila))
                
                '(13)=(5)+(11) M=E+K
                
                xlHoja1.Range("M" & Trim(str(nFila))).Formula = "=$E$" & Trim(str(nFila)) & "+$H$" & Trim(str(nFila)) & "+$K$" & Trim(str(nFila))
                
                
                'xlHoja1.Cells(nFila, 13) = matCarVen2(i) + matCobJud2(i) + matCarRef2(i)
                
                
                '(14)=(13)/(12) N=M/L
                
                If xlHoja1.Cells(nFila, 12) <> 0 Then
                    xlHoja1.Range("N" & Trim(str(nFila))).Formula = "=$M$" & Trim(str(nFila)) & "/$L$" & Trim(str(nFila))
                Else
                    xlHoja1.Cells(nFila, 14) = 0
                End If
                 
                '(15)=(5)+(8)+(11) O=E+H+K
                xlHoja1.Range("O" & Trim(str(nFila))).Formula = "=$E$" & Trim(str(nFila)) & "+$H$" & Trim(str(nFila)) & "+$K$" & Trim(str(nFila))
                
                '(16)=(15)/(12) P=O/L
                If xlHoja1.Cells(nFila, 12) <> 0 Then
                    xlHoja1.Range("P" & Trim(str(nFila))).Formula = "=$O$" & Trim(str(nFila)) & "/$L$" & Trim(str(nFila))
                Else
                    xlHoja1.Cells(nFila, 16) = 0
                End If
                 
                xlHoja1.Cells(nFila, 17) = Format(matDesemNue1(i), "#,##0")
                xlHoja1.Cells(nFila, 18) = Format(matDesemNue2(i), "#,##0.00")
                xlHoja1.Cells(nFila, 19) = Format(matDesemRepre1(i), "#,##0")
                xlHoja1.Cells(nFila, 20) = Format(matDesemRepre2(i), "#,##0.00")
                
                '(21)=(17)+(19) U=Q+S
                xlHoja1.Range("U" & Trim(str(nFila))).Formula = "=$Q$" & Trim(str(nFila)) & "+$S$" & Trim(str(nFila))
                
                '(22)=(18)+(20) V=R+T
                xlHoja1.Range("V" & Trim(str(nFila))).Formula = "=$R$" & Trim(str(nFila)) & "+$T$" & Trim(str(nFila))
                
                xlHoja1.Cells(nFila, 23) = Format(matOpeRef1(i), "#,##0")
                xlHoja1.Cells(nFila, 24) = Format(matOpeRef2(i), "#,##0.00")
                     
                With xlHoja1.Range("A" & Trim(str(nFila)) & ":X" & Trim(str(nFila)))
                    '.Font.Bold = True
                    .Borders.LineStyle = xlDash
                    .Borders.Weight = xlThin
                    .Borders.ColorIndex = 0
                    '.Interior.ColorIndex = 19
                End With
                  
            Next
             
            'Imprimo fila del ultimo subtotal
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "SubTotal"
            sFilaSubTotal(nContadorSubFila) = nFila
             
            'imprimo fila en blanco
            nFila = nFila + 1
            
            'imprimo fila que diga total
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "Total"
            sFilaTotal = nFila
            
            xlHoja1.Range("A1:D1").MergeCells = True
            xlHoja1.Range("A2:B2").MergeCells = True
            xlHoja1.Range("A3:X3").MergeCells = True
            xlHoja1.Range("A4:X4").MergeCells = True
            xlHoja1.Range("B6:C6").MergeCells = True
            xlHoja1.Range("D6:F6").MergeCells = True
            xlHoja1.Range("G6:I6").MergeCells = True
            xlHoja1.Range("J6:K6").MergeCells = True
            xlHoja1.Range("L6:P6").MergeCells = True
            xlHoja1.Range("Q6:T6").MergeCells = True
            xlHoja1.Range("U6:V6").MergeCells = True
            xlHoja1.Range("W6:X6").MergeCells = True
            xlHoja1.Range("B7:C7").MergeCells = True
            xlHoja1.Range("D7:F7").MergeCells = True
            xlHoja1.Range("G7:I7").MergeCells = True
            xlHoja1.Range("J7:K7").MergeCells = True
            xlHoja1.Range("Q7:R7").MergeCells = True
            xlHoja1.Range("S7:T7").MergeCells = True
            xlHoja1.Range("U7:V7").MergeCells = True
            xlHoja1.Range("W7:X7").MergeCells = True
     
            xlHoja1.Range("A1:X8").Font.Bold = True
            xlHoja1.Range("A3:X8").HorizontalAlignment = xlCenter
            
            'SubTitulo/Comentario
            xlHoja1.Range("A5").HorizontalAlignment = xlLeft
            
            'De los subtotales
            For i = 0 To nContadorSubFila
                 
                 sTextoTotal = sTextoTotal & sFilaTextoSubTotal(i)
                 
                 xlHoja1.Range("A" & Trim(str(sFilaSubTotal(i))) & ":X" & Trim(str(sFilaSubTotal(i)))).Font.Bold = True
                 
                 'Negrita y Color de los Titulos x Agencia
                 xlHoja1.Range("A" & Trim(str(sFilaTitulos(i)))).Font.Bold = True
                 xlHoja1.Range("A" & Trim(str(sFilaTitulos(i))) & ":C" & Trim(str(sFilaTitulos(i)))).Interior.ColorIndex = 38
                             
                 'Formulas
                For j = Asc("B") To Asc("X")
                    
                    Select Case j
                    Case Asc("F")
                        If xlHoja1.Cells(sFilaSubTotal(i), 3) <> 0 Then
                            xlHoja1.Cells(sFilaSubTotal(i), 6) = xlHoja1.Cells(sFilaSubTotal(i), 5) / xlHoja1.Cells(sFilaSubTotal(i), 3)
                        Else
                            xlHoja1.Cells(sFilaSubTotal(i), 6) = 0
                        End If
                    Case Asc("I")
                        If xlHoja1.Cells(sFilaSubTotal(i), 3) <> 0 Then
                            xlHoja1.Cells(sFilaSubTotal(i), 9) = xlHoja1.Cells(sFilaSubTotal(i), 8) / xlHoja1.Cells(sFilaSubTotal(i), 3)
                        Else
                            xlHoja1.Cells(sFilaSubTotal(i), 9) = 0
                        End If
                    Case Asc("N")
                        If xlHoja1.Cells(sFilaSubTotal(i), 12) <> 0 Then
                            xlHoja1.Cells(sFilaSubTotal(i), 14) = xlHoja1.Cells(sFilaSubTotal(i), 13) / xlHoja1.Cells(sFilaSubTotal(i), 12)
                        Else
                            xlHoja1.Cells(sFilaSubTotal(i), 14) = 0
                        End If
                    Case Asc("P")
                        If xlHoja1.Cells(sFilaSubTotal(i), 12) <> 0 Then
                            xlHoja1.Cells(sFilaSubTotal(i), 16) = xlHoja1.Cells(sFilaSubTotal(i), 15) / xlHoja1.Cells(sFilaSubTotal(i), 12)
                        Else
                            xlHoja1.Cells(sFilaSubTotal(i), 16) = 0
                        End If
                    Case Else
                        xlHoja1.Range(Trim(Chr(j)) & Trim(str(sFilaSubTotal(i)))).Formula = "=" & Replace(sFilaTextoSubTotal(i), "*", Trim(Chr(j)))
                    End Select
                    
                Next j
            Next
            'Bordes y Colores del Total
            With xlHoja1.Range("A" & Trim(str(sFilaTotal)) & ":X" & Trim(str(sFilaTotal)))
                .Font.Bold = True
                .Borders.LineStyle = xlDash
                .Borders.Weight = xlThin
                .Borders.ColorIndex = 0
                .Interior.ColorIndex = 24
            End With
            
            'Calculo de Formulas de la Fila Total
            
            For i = Asc("B") To Asc("X")
            
                    Select Case i
                    Case Asc("F")
                        If xlHoja1.Cells(sFilaTotal, 3) <> 0 Then
                            xlHoja1.Cells(sFilaTotal, 6) = xlHoja1.Cells(sFilaTotal, 5) / xlHoja1.Cells(sFilaTotal, 3)
                        Else
                            xlHoja1.Cells(sFilaTotal, 6) = 0
                        End If
                    Case Asc("I")
                        If xlHoja1.Cells(sFilaTotal, 3) <> 0 Then
                            xlHoja1.Cells(sFilaTotal, 9) = xlHoja1.Cells(sFilaTotal, 8) / xlHoja1.Cells(sFilaTotal, 3)
                        Else
                            xlHoja1.Cells(sFilaTotal, 9) = 0
                        End If
                    Case Asc("N")
                        If xlHoja1.Cells(sFilaTotal, 12) <> 0 Then
                            xlHoja1.Cells(sFilaTotal, 14) = xlHoja1.Cells(sFilaTotal, 13) / xlHoja1.Cells(sFilaTotal, 12)
                        Else
                            xlHoja1.Cells(sFilaTotal, 14) = 0
                        End If
                    Case Asc("P")
                        If xlHoja1.Cells(sFilaTotal, 12) <> 0 Then
                            xlHoja1.Cells(sFilaTotal, 16) = xlHoja1.Cells(sFilaTotal, 15) / xlHoja1.Cells(sFilaTotal, 12)
                        Else
                            xlHoja1.Cells(sFilaTotal, 16) = 0
                        End If
                    Case Else
                        xlHoja1.Range(Trim(Chr(i)) & Trim(str(sFilaTotal))).Formula = "=" & Replace(sTextoTotal, "*", Trim(Chr(i)))
                    End Select
                
             Next
            
            'Bordes del Titulo
            With xlHoja1.Range("A6:X8")
                .Font.Bold = True
                .Borders.LineStyle = xlDash
                .Borders.Weight = xlThin
                .Borders.ColorIndex = 0
            End With
    
            'Colores del Titulo
            xlHoja1.Range("A6:K8").Interior.ColorIndex = 19
            xlHoja1.Range("L6:P8").Interior.ColorIndex = 15
            xlHoja1.Range("Q6:X8").Interior.ColorIndex = 35
    
            'Cierro...
            OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
            OleExcel.SourceDoc = lsArchivoN
            OleExcel.Verb = 1
            OleExcel.Action = 1
            OleExcel.DoVerb -1
        End If
    End If
End Function

Private Function Genera_ReporteWORD(ByVal psModeloCarta As Long, ByVal psMensaje As String, ByVal psCondicion As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAnalistas As String, ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, ByVal psNota1 As Integer, ByVal psNota2 As Integer, ByVal psTipoCuotas As Integer, ByVal psCuotasPend As Integer, _
Optional psUbicacionGeo As String = "")

Dim oDCredDoc As DCredDoc
Dim nMontoAtraso As Double

Dim aLista() As String
Dim vFilas As Integer
Dim vFecAviso As Date
Dim K As Integer
Dim CadenaAna As String

Dim psCtaCod As String

Dim lnDeudaFecha As Currency
 
'A la Fecha
Dim lnSaldoKFecha As Currency
Dim lnIntCompFecha As Currency
Dim lnGastoFecha As Currency
Dim lnIntMorFecha As Currency
Dim lnPenalidadFecha As Currency
 
Dim oNegCred As NCredito
Dim MatCalend As Variant
Dim j As Integer

Dim loExc As DCredReporte

Dim rsCarta As New ADODB.Recordset

Dim lsModeloPlantilla As String
Dim vCont As Integer
Dim lnDeuda As Currency
 
Select Case psModeloCarta
    Case gColCredRepCartaCobMoro1
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso1
    Case gColCredRepCartaCobMoro2
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso2
    Case gColCredRepCartaCobMoro3
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso3
    Case gColCredRepCartaCobMoro4
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso4
    Case gColCredRepCartaCobMoro5
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso5
    Case gColCredRepCartaInvCredAlt
        lsModeloPlantilla = App.Path & cPlantillaCartaInvCredParalelo
    Case gColCredRepCartaRecup
        lsModeloPlantilla = App.Path & cPlantillaCartaRecup
    Case Else
        MsgBox " Error en la definicion de la Plantilla"
        Genera_ReporteWORD = "Error en la definicion de la plantilla"
        Exit Function
End Select

    Set loExc = New DCredReporte
    
    Set rsCarta = loExc.RecuperaDatosCartasWORD(IIf(psModeloCarta = gColCredRepCartaCobMoro1, 0, IIf(psModeloCarta = gColCredRepCartaInvCredAlt, 2, 1)), psCondicion, psMoneda, psProductos, psAnalistas, pnDiasIni, pnDiasFin, psNota1, psNota2, psTipoCuotas, psCuotasPend, psUbicacionGeo, gdFecSis)
     
    If rsCarta.BOF Then
        Genera_ReporteWORD = ""
        Exit Function
    Else
        Genera_ReporteWORD = "Reporte Generado"
    End If
    
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
    
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
    
    'Crea Nuevo Documento
    wApp.Documents.Add
    Dim sTemCuenta As String
    
    Do While Not rsCarta.EOF
        vFilas = vFilas + 1
          
        psCtaCod = rsCarta!cCtaCod
        'sTemCuenta = rsCarta!cCtaCod
        
       
        'Obtener la deuda A LA FECHA
        '===========================
        Set oNegCred = New NCredito
        MatCalend = oNegCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
        Dim nDeudaVencida As Double
        
        lnSaldoKFecha = Format(oNegCred.MatrizCapitalAFecha(psCtaCod, MatCalend), "#0.00")
        If UBound(MatCalend) > 0 Then
            lnIntCompFecha = Format(oNegCred.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, gdFecSis), "#0.00")
            lnGastoFecha = Format(oNegCred.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis), "#0.00")
            lnIntMorFecha = Format(oNegCred.MatrizInteresMorFecha(psCtaCod, MatCalend), "#0.00")
            lnPenalidadFecha = Format(oNegCred.CalculaGastoPenalidadCancelacion(CDbl(lnSaldoKFecha), CInt(Mid(psCtaCod, 9, 1))), "#0.00")
            lnDeudaFecha = Format(CDbl(lnSaldoKFecha) + CDbl(lnIntCompFecha) + CDbl(lnGastoFecha) + CDbl(lnIntMorFecha) + CDbl(lnPenalidadFecha), "#0.00")
            
            nDeudaVencida = oNegCred.MatrizCapitalVencido(MatCalend, gdFecSis) + oNegCred.MatrizIntCompVencido(MatCalend, gdFecSis) + oNegCred.MatrizGastosVencidos(MatCalend, gdFecSis) + oNegCred.MatrizIntGraciaVencido(MatCalend, gdFecSis) + oNegCred.MatrizIntMoratorioCalendario(MatCalend)
            
        End If
 
        '===========================
           Set oDCredDoc = New DCredDoc
           nMontoAtraso = oDCredDoc.Recup_MontoAtrasado(gdFecSis, psCtaCod)
           Set oDCredDoc = Nothing
        '===========================

        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
    
        With wApp.Selection.Find
            .Text = "CampFecha"
            .Replacement.Text = Trim(ImpreFormat(Format(gdFecSis, "dddd, d mmmm yyyy"), 25))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
       With wApp.Selection.Find
            .Text = "CampTitNombre"
            .Replacement.Text = Trim(PstaNombre(rsCarta!cPersNombre, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "CampTitDireccion"
            .Replacement.Text = Trim(rsCarta!cPersDireccDomicilio) & " - " & Trim(rsCarta!cUbiGeoDescripcion)
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "CampTitZonaDir"
            .Replacement.Text = Trim(rsCarta!cUbiGeoDescripcion)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "CampNroCredito"
            .Replacement.Text = Trim(psCtaCod)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "CampAnalista"
            .Replacement.Text = Trim(PstaNombre(rsCarta!cDesAnalista, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        If psModeloCarta <> gColCredRepCartaCobMoro1 Then
            With wApp.Selection.Find
                .Text = "CampTitDirNegocio"
                .Replacement.Text = Trim("" & rsCarta!cRazSocDirecc)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampFiadorNombre"
                .Replacement.Text = Trim(PstaNombre(rsCarta!cDesFiador, True))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampFiadorDireccion"
                .Replacement.Text = Trim(PstaNombre(rsCarta!cDireccionFiador, True))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampDiasAtraso"
                .Replacement.Text = Trim(str(rsCarta!nDiasAtraso))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampCuotasVenc"
                .Replacement.Text = Trim(str(rsCarta!nCuota))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampMoneda"
                .Replacement.Text = Trim(rsCarta!cDesMoneda)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampDeuda"
                '.Replacement.Text = Format(lnDeudaFecha, "#,###.00")
                .Replacement.Text = Format(fgITFCalculaImpuestoNOIncluido(nDeudaVencida), "#,###.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
             End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            
        
            If psModeloCarta = gColCredRepCartaInvCredAlt Then
             
                With wApp.Selection.Find
                    .Text = "CampCuotasPend"
                    .Replacement.Text = Trim(str(rsCarta!nCuotasPend))
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With
                wApp.Selection.Find.Execute Replace:=wdReplaceAll
            
                With wApp.Selection.Find
                    .Text = "CampNota"
                    .Replacement.Text = Trim(str(rsCarta!nColocNota))
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With
                wApp.Selection.Find.Execute Replace:=wdReplaceAll
            End If
            
        End If
        
    
        rsCarta.MoveNext
    Loop
    rsCarta.Close
    Set rsCarta = Nothing

 
wAppSource.ActiveDocument.Close
wApp.Visible = True

End Function

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox err.Description, vbInformation, "Aviso"
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Private Function ListaAgencias() As String
    Dim i As Integer
    Dim sCadena As String
    
    For i = 0 To UBound(MatAgencias) - 1
        If i = 0 Then
            sCadena = MatAgencias(0)
        Else
            sCadena = sCadena & "','" & MatAgencias(i)
        End If
    Next i
    
    ListaAgencias = sCadena
End Function

Private Sub GeneraReporteXls()

Dim lsArchivoN As String

    lsArchivoN = App.Path & "\Spooler\RptXls" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        Set xlHoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
      
        Call GeneraReporteXlsDet

        ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
        xlHoja1.Cells.Select
        xlHoja1.Cells.EntireColumn.AutoFit
    End If
    MousePointer = 0

End Sub

Private Sub GeneraReporteXlsDet()
Dim ssql As String
Dim nFila As Integer
Dim oCon As New DConecta
Dim rs As New ADODB.Recordset
ssql = "select *  From TmpConsolidadoAnalista Where IdCodigo=(Select Max(IdCodigo) From TmpConsolidadoAnalista Where Usuario='LMMD')"

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(ssql)

'Agregar Cabecera

nFila = 6
Do While Not rs.EOF
    nFila = nFila + 1
    xlHoja1.Cells(nFila, 1) = rs!CodAna
    xlHoja1.Cells(nFila, 2) = rs!NumCreVig
    xlHoja1.Cells(nFila, 3) = rs!KDesemVig
    xlHoja1.Cells(nFila, 4) = rs!SaldoKVig
    xlHoja1.Cells(nFila, 5) = rs!SaldoPromVig
    xlHoja1.Cells(nFila, 6) = rs!NumCreNewMes
    xlHoja1.Cells(nFila, 7) = rs!KDesemNewMes
    xlHoja1.Cells(nFila, 8) = rs!NumCreRepreMes
    xlHoja1.Cells(nFila, 9) = rs!KDesemRepreMes
    xlHoja1.Cells(nFila, 10) = rs!NumCreVen1a15
    xlHoja1.Cells(nFila, 11) = rs!K1a15
    xlHoja1.Cells(nFila, 12) = rs!Por1a15
    xlHoja1.Cells(nFila, 13) = rs!NumCreVen16a30
    xlHoja1.Cells(nFila, 14) = rs!K16a30
    xlHoja1.Cells(nFila, 15) = rs!Por16a30
    xlHoja1.Cells(nFila, 16) = rs!NumCreVen31aN
    xlHoja1.Cells(nFila, 17) = rs!K31aN
    xlHoja1.Cells(nFila, 18) = rs!Por31aN
    xlHoja1.Cells(nFila, 19) = rs!NumCreJud
    xlHoja1.Cells(nFila, 20) = rs!KJudicial
    xlHoja1.Cells(nFila, 21) = rs!PorJud
    xlHoja1.Cells(nFila, 22) = rs!NumJudVen
    xlHoja1.Cells(nFila, 23) = rs!KJudVen
    xlHoja1.Cells(nFila, 24) = rs!PorJudVen
    xlHoja1.Cells(nFila, 25) = rs!IdCodigo

    rs.MoveNext
Loop
Set rs = Nothing
oCon.CierraConexion

Set oCon = Nothing

End Sub

Sub ImprimirCrystal(ByVal psFileName As String, Optional ByVal psFormula As String = "", Optional psParam1 As String = "", _
Optional ByVal psParam2 As String = "", Optional ByVal psParam3 As String = "", Optional ByVal psParam4 As String = "", _
Optional ByVal psParam5 As String = "", Optional ByVal psParam6 As String = "", Optional ByVal psParam7 As String = "", _
Optional ByVal psParam8 As String = "", Optional ByVal psParam9 As String, Optional ByVal psParam10 As String = "")
    Dim oConec As DConecta
    Dim sCadenaConexion As String
    Dim sServidor As String
    Dim sBase As String
    Dim sConexion As String
    Dim sConexionAux As String
    
    Set oConec = New DConecta
    oConec.AbreConexion
    sCadenaConexion = oConec.GetStringConnection
    sServidor = oConec.ServerName
    sBase = oConec.DatabaseName
    
    oConec.CierraConexion
    Set oConec = Nothing
    
    sConexion = "Data Source=" & sServidor & ";User Id=USERSICMACCONS;Initial Catalog" & sBase & ";pwd=sicmacicons"
    sConexionAux = "dsn=" & sServidor & ";uid=USERSICMACCONS;dsq=" & sBase & ";pwd=sicmacicons"
    'CUSCO
'    With CR
'         .Connect = sConexion
'         .Connect = sConexionAux
'         .WindowControls = True
'         .WindowState = crptMaximized
'
'         If psFormula <> "" Then
'            .SelectionFormula = psFormula
'         End If
'         If psParam1 <> "" Then
'            .ParameterFields(0) = psParam1
'         End If
'         If psParam2 <> "" Then
'            .ParameterFields(1) = psParam2
'         End If
'         If psParam3 <> "" Then
'            .ParameterFields(2) = psParam3
'         End If
'         If psParam4 <> "" Then
'            .ParameterFields(3) = psParam4
'         End If
'         If psParam5 <> "" Then
'            .ParameterFields(4) = psParam5
'         End If
'         If psParam6 <> "" Then
'            .ParameterFields(5) = psParam6
'         End If
'         If psParam7 <> "" Then
'            .ParameterFields(6) = psParam7
'         End If
'         If psParam8 <> "" Then
'            .ParameterFields(7) = psParam8
'         End If
'         If psParam9 <> "" Then
'            .ParameterFields(8) = psParam9
'         End If
'         If psParam10 <> "" Then
'            .ParameterFields(9) = psParam10
'         End If
'
'         .ReportFileName = psFileName
'         .Destination = crptToWindow
'         .WindowState = crptMaximized
'         .Action = 1
'         .Reset
'      End With
End Sub
Function GetListaAnalistas() As String
    Dim i As Integer
    Dim sAnalistas  As String
    
    For i = 0 To UBound(matAnalista) - 1
        If i = 0 Then
            sAnalistas = "'" & matAnalista(i) & "'"
        Else
            sAnalistas = sAnalistas & ",'" & matAnalista(i) & "'"
        End If
    Next i
    GetListaAnalistas = sAnalistas
End Function

Function GetListaAgencias() As String
    Dim i As Integer
    Dim sAgencias As String
    
    For i = 0 To UBound(MatAgencias) - 1
        If i = 0 Then
            sAgencias = "'" & MatAgencias(i) & "'"
        Else
            sAgencias = sAgencias & ",'" & MatAgencias(i) & "'"
        End If
    Next
    GetListaAgencias = sAgencias
End Function

'ARCV 15-02-2007
Private Function Genera_ReporteWORD_NEW(ByVal psModeloCarta As Long, ByVal psMensaje As String, ByVal psCondicion As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAnalistas As String, ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, ByVal psUbicacionGeo As String, Optional ByVal psAgencias As String)

Dim oDCredDoc As DCredDoc
Dim nMontoAtraso As Double

Dim aLista() As String
Dim vFilas As Integer
Dim vFecAviso As Date
Dim K As Integer
Dim CadenaAna As String

Dim psCtaCod As String

Dim lnDeudaFecha As Currency
 
'A la Fecha
Dim lnSaldoKFecha As Currency
Dim lnIntCompFecha As Currency
Dim lnGastoFecha As Currency
Dim lnIntMorFecha As Currency
Dim lnPenalidadFecha As Currency
 
Dim oNegCred As NCredito
Dim MatCalend As Variant
Dim j As Integer

Dim loExc As DCredReporte
Dim oCOMNCredDoc As COMNCredito.NCOMColocEval


Dim rsCarta As New ADODB.Recordset

Dim lsModeloPlantilla As String
Dim vCont As Integer
Dim lnDeuda As Currency

On Error GoTo ErrGeneraRepo '*** PEAC 20100528

'On Error GoTo ErrEnd
'   If plSave Then
'        xlHoja1.SaveAs psArchivo
'   End If
'   xlLibro.Close
'   xlAplicacion.Quit
'   Set xlAplicacion = Nothing
'   Set xlLibro = Nothing
'   Set xlHoja1 = Nothing
'Exit Sub
'ErrEnd:
'   MsgBox Err.Description, vbInformation, "Aviso"
'


Select Case psModeloCarta
    Case gColCredRepCartaCobMoro1
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso1
    Case gColCredRepCartaCobMoro2
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso2
    Case gColCredRepCartaCobMoro3
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso3
    Case gColCredRepCartaCobMoro4
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso4
    Case gColCredRepCartaCobMoro5
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso5
    Case gColCredRepCartaCobMoro6
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso6
    Case gColCredRepCartaRecup
        'lsModeloPlantilla = App.path & cPlantillaCartaRecup'WIOR 20130910 COMENTÓ
        lsModeloPlantilla = App.Path & cPlantillaCartaRecup1 'WIOR 20130910
    'WIOR 20130910 ********************************************************
    Case gColCredRepCartaRecup2
        lsModeloPlantilla = App.Path & cPlantillaCartaRecup2
    'WIOR FIN **************************************************************
    Case Else
        MsgBox " Error en la definicion de la Plantilla"
        Genera_ReporteWORD_NEW = "Error en la definicion de la plantilla"
        Exit Function
End Select


'    Set loExc = New DCredReporte
'    Set rsCarta = loExc.RecuperaDatosCartasWORD_NEW(IIf(psModeloCarta = gColCredRepCartaCobMoro1 Or psModeloCarta = gColCredRepCartaCobMoro3 Or psModeloCarta = gColCredRepCartaCobMoro5 Or psModeloCarta = gColCredRepCartaRecup, 0, 1), psCondicion, psMoneda, psProductos, psAnalistas, pnDiasIni, pnDiasFin, psUbicacionGeo, gsCodUser, GetMaquinaUsuario, psAgencias)
    
    Screen.MousePointer = 11 '*** PEAC 20100528
    
    Set oCOMNCredDoc = New COMNCredito.NCOMColocEval
'        sCadImp = sCadImp & oCOMNCredDoc.ImprimeRepConvCasillero(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCadMoneda, MatInstitucion, gsNomCmac, matAnalista, matCampanha, MatGastos)
        Set rsCarta = oCOMNCredDoc.RecuperaDatosCartasWORD_NEW(IIf(psModeloCarta = gColCredRepCartaCobMoro1 Or psModeloCarta = gColCredRepCartaCobMoro3 Or psModeloCarta = gColCredRepCartaCobMoro5 Or psModeloCarta = gColCredRepCartaRecup, 0, 1), psCondicion, psMoneda, psProductos, psAnalistas, pnDiasIni, pnDiasFin, psUbicacionGeo, gsCodUser, GetMaquinaUsuario, psAgencias, gsCodAge)
    Set oCOMNCredDoc = Nothing
     
    If rsCarta.BOF Then
        Genera_ReporteWORD_NEW = ""
        Screen.MousePointer = 0 '*** PEAC 20100528
        Exit Function
    Else
        Genera_ReporteWORD_NEW = "Reporte Generado"
    End If

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
    
    Dim oNCred As COMNCredito.NCOMCredito
    Dim nDeudaFecha As Double
    Set oNCred = New COMNCredito.NCOMCredito
    
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
    
    'Crea Nuevo Documento
    wApp.Documents.Add
    Dim sTemCuenta As String
    
    Do While Not rsCarta.EOF
        vFilas = vFilas + 1
          
        psCtaCod = rsCarta!cCtaCod
        
       
        'Obtener la deuda A LA FECHA
        '===========================
  '      If psModeloCarta = gColCredRepCartaCobMoro5 Or psModeloCarta = gColCredRepCartaCobMoro6 Then
            nDeudaFecha = oNCred.ObtenerDeudaFechaTotal(psCtaCod, , gdFecSis)
  '      End If
        '===========================

        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
        'WIOR 20130924**********
        With wApp.Selection.Find
            .Text = "CCIUDADAGE"
            .Replacement.Text = Trim(rsCarta!Ciudad)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        If psModeloCarta = gColCredRepCartaRecup2 Then
        Dim nNumCuotaVencidas  As Long
        Dim nCuotaVencida  As Long
        Dim sCuotasVenc  As String
        Dim iCuota As Long
        sCuotasVenc = ""
        nNumCuotaVencidas = CLng(rsCarta!nNumCuotasVencidas)
        nCuotaVencida = CLng(rsCarta!nNroCuota)
        
        sCuotasVenc = nCuotaVencida
        For iCuota = 2 To nNumCuotaVencidas
            nCuotaVencida = nCuotaVencida + 1
            
            If iCuota = nNumCuotaVencidas Then
                sCuotasVenc = sCuotasVenc & " y " & nCuotaVencida
            Else
                sCuotasVenc = sCuotasVenc & ", " & nCuotaVencida
            End If
            
        Next iCuota
        
            With wApp.Selection.Find
                .Text = "nCuotas"
                .Replacement.Text = Trim(sCuotasVenc)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
            End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
        End If
        'WIOR 20130924**********
        
        With wApp.Selection.Find
            .Text = "dFecha"
            .Replacement.Text = Trim(ImpreFormat(Format(gdFecSis, "d mmmm yyyy"), 25))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
       With wApp.Selection.Find
            .Text = "cCliente"
            .Replacement.Text = Trim(PstaNombre(rsCarta!cTitular, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "cDireccion"
            'ARCV 22-06-2007
            '.Replacement.Text = Trim(rsCarta!cDireccion) & " - " & Trim(rsCarta!cUbiGeoDescripcion)
            .Replacement.Text = Trim(rsCarta!cDireccion)
            '----
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        '*** PEAC 20100514
        
        With wApp.Selection.Find
            .Text = "nDiasAtraso"
            .Replacement.Text = Trim(rsCarta!nDiasAtraso)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "nNumCuotasVenc"
            .Replacement.Text = Trim(rsCarta!nNumCuotasVencidas)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cCiudadAge"
            .Replacement.Text = Trim(rsCarta!Ciudad)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        
        With wApp.Selection.Find
            .Text = "cDireNegocio"
            .Replacement.Text = Trim(rsCarta!cDireNegocio)
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cNomAnalista"
            .Replacement.Text = Trim(PstaNombre(rsCarta!Analista, True))
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "dFecVenc"
            .Replacement.Text = Trim(ImpreFormat(Format(rsCarta!dFecCuotaVenc, "d mmmm yyyy"), 25))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "nNumCuota"
            .Replacement.Text = Trim(rsCarta!nNumCuotaVenc)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "nMontoCuota"
            '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/. ", "US $ ") & Format(rsCarta!nMontoCuotaVenc, "#0.00") 'marg ers044-2016
            .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", StrConv(gcPEN_SIMBOLO, vbProperCase) & " ", "US $ ") & Format(rsCarta!nMontoCuotaVenc, "#0.00") 'marg ers044-2016
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        '*** FIN PEAC
        
        With wApp.Selection.Find
            .Text = "cCredito"
            .Replacement.Text = psCtaCod
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "nNroCuota"
            .Replacement.Text = Trim(rsCarta!nNroCuota)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find ' Add By GITU 31/05/2008
            .Text = "cUser"
            .Replacement.Text = Trim(rsCarta!cUser)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        'If psModeloCarta = gColCredRepCartaCobMoro2 Or psModeloCarta = gColCredRepCartaCobMoro4 Or psModeloCarta = gColCredRepCartaCobMoro6 Then
        'If psModeloCarta = gColCredRepCartaCobMoro3 Or psModeloCarta = gColCredRepCartaCobMoro5 Then
            With wApp.Selection.Find
                .Text = "cGarante"
                .Replacement.Text = Trim(PstaNombre(rsCarta!cGarante, True))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
            End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
        'End If
        
'        If psModeloCarta = gColCredRepCartaCobMoro5 Or psModeloCarta = gColCredRepCartaCobMoro6 Then
            With wApp.Selection.Find
                .Text = "nMonto"
                '''.Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/. ", "US $ ") & Format(nDeudaFecha, "#,#0.00") 'marg ers044-2016
                .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = "1", StrConv(gcPEN_SIMBOLO, vbProperCase) & " ", "US $ ") & Format(nDeudaFecha, "#,#0.00") 'marg ers044-2016
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
 '       End If
    
          With wApp.Selection.Find
            .Text = "cCiuCliente"
            .Replacement.Text = Trim(rsCarta!Ciudad)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        rsCarta.MoveNext
    Loop
    rsCarta.Close
    Set rsCarta = Nothing
    Set oNCred = Nothing
 
Screen.MousePointer = 0 '*** PEAC 20100528
 
wAppSource.ActiveDocument.Close
wApp.Visible = True

Exit Function
ErrGeneraRepo: '*** PEAC 20100528
    Screen.MousePointer = 0
    wAppSource.ActiveDocument.Close
    MsgBox "Error en frmCredReportes.Genera_ReporteWORD_NEW " & err.Description, vbInformation, "Aviso"

End Function

Private Function GetMaquinaUsuario() As String  'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    
    GetMaquinaUsuario = GetComputerName()
End Function

'**DAOR 20070512, Procedimiento que imprime el estado de cuenta de crédito
Public Sub ImprimeEstadoCuentaCredito(ByVal psCodAge As String, ByVal pMatProd As Variant, pMatAgencias As Variant, pdDesembDe As Date, pdDesembHasta As Date)
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim lsModeloPlantilla As String
Dim lMatCabecera As Variant

    lsModeloPlantilla = App.Path & "\FormatoCarta\EstadoCuentaCredito.doc"
    
    ReDim lMatCabecera(24, 2)
    lMatCabecera(0, 0) = "Cuenta": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "Agencia": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "MontoCol": lMatCabecera(2, 1) = "N"
    lMatCabecera(3, 0) = "TasaCompAnual": lMatCabecera(3, 1) = "N"
    lMatCabecera(4, 0) = "TasaCosEfeAnu": lMatCabecera(4, 1) = "N"
    lMatCabecera(5, 0) = "Saldo": lMatCabecera(5, 1) = "N"
    lMatCabecera(6, 0) = "CuotaPag": lMatCabecera(6, 1) = "N"
    lMatCabecera(7, 0) = "VencPag": lMatCabecera(7, 1) = "D"
    lMatCabecera(8, 0) = "FecPag": lMatCabecera(8, 1) = "D"
    lMatCabecera(9, 0) = "CapPag": lMatCabecera(9, 1) = "N"
    lMatCabecera(10, 0) = "IntPag": lMatCabecera(10, 1) = "N"
    lMatCabecera(11, 0) = "MorPag": lMatCabecera(11, 1) = "N"
    lMatCabecera(12, 0) = "SegDesPag": lMatCabecera(12, 1) = "N"
    lMatCabecera(13, 0) = "SegBiePag": lMatCabecera(13, 1) = "N"
    lMatCabecera(14, 0) = "ComiPorPag": lMatCabecera(14, 1) = "N"
    'By Capi 28122007
    lMatCabecera(15, 0) = "OtrPorPag": lMatCabecera(15, 1) = "N"
    
    lMatCabecera(16, 0) = "ItfPag": lMatCabecera(16, 1) = "N"
    lMatCabecera(17, 0) = "CuotaAPag": lMatCabecera(17, 1) = "N"
    lMatCabecera(18, 0) = "VencAPag": lMatCabecera(18, 1) = "D"
    lMatCabecera(19, 0) = "MontoAPag": lMatCabecera(19, 1) = "N"
    lMatCabecera(20, 0) = "NomCli": lMatCabecera(20, 1) = ""
    lMatCabecera(21, 0) = "DirCli": lMatCabecera(21, 1) = ""
    lMatCabecera(22, 0) = "Moneda": lMatCabecera(22, 1) = ""
    lMatCabecera(23, 0) = "Fecha": lMatCabecera(23, 1) = "D"
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosParaEstadoCuentaCredito("", pMatProd, pMatAgencias, pdDesembDe, pdDesembHasta)
    Set oDCred = Nothing
        
    Call GeneraArchivoExcel("ListaEstadoCuentaCredito", lMatCabecera, R, 2, , "Listado")
  
    MsgBox "Proceso finalizado, el documento se encuentra en:" & lsModeloPlantilla
End Sub


'**DAOR 20070917, Método para disminuir lìneas en el evento click
Sub ColCredRepLisDesctoPlanilla()
Dim oNCredDoc As NCredDoc
Dim nValTmp As Integer
    
    Set oNCredDoc = New NCredDoc
        oNCredDoc.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
        
        'Falta el Calculo dela Cuota debe incluir los intereses a la fecha
        'ya que estos creditos son cuota libre
        'Para ello se penso realizar una funcion en sql server para calculo de interes a la fecha
        If OptOrdenAlfabetico.value Then
            nValTmp = 1
        End If
        If OptOrdenCodMod.value Then
            nValTmp = 0
        End If
        If OptOrdenPagare.value Then
            nValTmp = 2
        End If

        If ChkDia.value = 1 Then
           bDiaHora = True
        Else
           bDiaHora = False
        End If

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            'Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, bDiaHora)
            'Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, bDiaHora)
            If bDiaHora = True Then
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, 1)
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, 1)
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, 30)
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, 30)
            Else
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, 0)
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion, 0)
            End If
            'sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion)
            'sCadImp = sCadImp & Chr$(12)
            'sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion)
        Else
            If ChkMonA02(0).value = 1 Then
                'sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion)
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion)
            Else
                'sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion)
                Call oNCredDoc.ImprimeCreditosXInstitucionDet(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value, , , MatInstitucion)
            End If
        End If

'        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
'            sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
'            sCadImp = sCadImp & Chr$(12)
'            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
'        Else
'            If ChkMonA02(0).value = 1 Then
'                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
'            Else
'                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
'            End If
'        End If

    Set oNCredDoc = Nothing
End Sub

'DAOR 20071210, Procedimiento creado para disminuir líneas al evento click
Sub ColCredRepMoraxAna(ByRef sCadImp As String)
Dim oNCredDoc As NCredDoc
'Dim sUbicacionGeo As String
Dim R As ADODB.Recordset

    
    Set oNCredDoc = New NCredDoc
        oNCredDoc.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis

        Dim sMonedaMor As String
        If ChkMonA02(0).value = 1 Then
            sMonedaMor = "1"
        End If
        If ChkMonA02(1).value = 1 Then
            sMonedaMor = "2"
        End If
        If ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1 Then
            sMonedaMor = "1,2"
        End If
        If ChkUbi.value = 1 Then
            If sUbicacionGeo = "" Then
                MsgBox "Debe seleccionar una ubicacion geografica", vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            If sUbicacionGeo <> "" Then
                MsgBox "Debe seleccionar check de ubicacion geografica", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
'            Dim oCredDoc As DCredDoc
'            Dim sAgencias As String
'            Dim sCodUser As String
'            Dim sProductos As String
'            Dim sAnalistas As String
'            Dim sCondicion As String
'            Dim I As Integer
'            Set oCredDoc = New DCredDoc
'            For I = 0 To UBound(matAnalista) - 1
'            If I = 0 Then
'            sAnalistas = "'" & matAnalista(I) & "'"
'            Else
'            sAnalistas = sAnalistas & ",'" & matAnalista(I) & "'"
'            End If
'            Next I
'            For I = 0 To UBound(MatAgencias) - 1
'            If I = 0 Then
'            sAgencias = "'" & MatAgencias(I) & "'"
'            Else
'            sAgencias = sAgencias & ",'" & MatAgencias(I) & "'"
'            End If
'            Next I
'
'            For I = 0 To UBound(MatProductos) - 1
'            If I = 0 Then
'            sProductos = "'" & MatProductos(I) & "'"
'            Else
'            sProductos = sProductos & ",'" & MatProductos(I) & "'"
'            End If
'            Next I
'
'            For I = 0 To UBound(MatCondicion) - 1
'            If I = 0 Then
'            sCondicion = "'" & MatCondicion(I) & "'"
'            Else
'            sCondicion = sCondicion & ",'" & MatCondicion(I) & "'"
'            End If
'            Next I
'            Set R = oCredDoc.ListaMoraXAnalistaCabecera(sAnalistas, sProductos, sAgencias, sCondicion, sMonedaMor, CDate(Me.TxtFecFinA02), CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), Val(TxtMontoIni), Val(TxtMontoFin), Right(sUbicacionGeo, 15))
'            Set oCredDoc = Nothing
        If ChkUbi.value = 1 Then
            
            'R = oNCredDoc.ImprimeMoraXAnalitaNew(MatAgencias, CDate(Me.TxtFecFinA02), sMonedaMor, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), Val(TxtMontoIni), Val(TxtMontoFin), sUbicacionGeo, gsNomCmac)
            If chkMigraExcell.value = 1 Then
                Call MostrarReporteMoraDiarioxAnalista(MatAgencias, CDate(Me.TxtFecFinA02), sMonedaMor, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), val(TxtMontoIni), val(TxtMontoFin), sUbicacionGeo, gsNomCmac)
            Else
                sCadImp = oNCredDoc.ImprimeMoraXAnalitaNew(MatAgencias, CDate(Me.TxtFecFinA02), sMonedaMor, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), val(TxtMontoIni), val(TxtMontoFin), sUbicacionGeo, gsNomCmac)
            End If
        Else
            
            'R = oNCredDoc.ImprimeMoraXAnalitaNew(MatAgencias, CDate(TxtFecFinA02), sMonedaMor, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), Val(TxtMontoIni), Val(TxtMontoFin), "", gsNomCmac)
            If chkMigraExcell.value = 1 Then
                Call MostrarReporteMoraDiarioxAnalista(MatAgencias, CDate(TxtFecFinA02), sMonedaMor, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), val(TxtMontoIni), val(TxtMontoFin), "", gsNomCmac)
            Else
                sCadImp = oNCredDoc.ImprimeMoraXAnalitaNew(MatAgencias, CDate(TxtFecFinA02), sMonedaMor, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, CInt(TxtDiaAtrIni), CInt(TxtDiasAtrFin), val(TxtMontoIni), val(TxtMontoFin), "", gsNomCmac)
            End If
        End If
        sUbicacionGeo = ""
    Set R = Nothing
    Set oNCredDoc = Nothing
End Sub
'By Capi 28122007
Private Sub MostrarGarantiasInscritas(ByVal pdFecFinal As Date, ByVal pMatAgencias As Variant, ByVal psMoneda As String, ByVal pMatProd As Variant, ByVal sEstados As String)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant
Dim lsmensaje As String
Dim i As Integer
Dim sCadAge As String

    If Len(psMoneda) = 0 Then
        MsgBox "Seleccione por lo menos una moneda.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Screen.MousePointer = 11
    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.ObtenerGarantiasInscritas(pdFecFinal, sCadAge, psMoneda, pMatProd, sEstados)
    Set oDCred = Nothing
    
    lsNombreArchivo = "Garantias Inscritas"
    'By Capi 01102008 para que imprima directamente en excel.
    'Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Garantias Inscritas", " Al " & CStr(pdFecFinal), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Garantias Inscritas", " Al " & CStr(pdFecFinal), lsNombreArchivo, lMatCabecera, R, 2, True, , True, True)

    Screen.MousePointer = 0

End Sub
'By Capi 28012008
Private Sub mostrarREULavadoDinero(ByVal pdFecInicial As String, ByVal pdFecFinal As String, ByVal pMatAgencias As Variant)

'**************************************************************
'**Modificado por ELRO 20110714, según acta 158-2011/TI-D
Me.PbProgresoReporte.Visible = True
Me.PbProgresoReporte.value = 0
Me.PbProgresoReporte.value = 10
'**************************************************************

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lsRutaArchivo As String
Dim lMatCabecera(1, 1) As Variant
Dim lsmensaje As String
Dim i As Integer
Dim sCadAge As String
Dim sInicio As String
Dim sTermino As String
Dim lsArregloNombreArchivo() As String
Dim lsArregloRuta() As String
Dim j As Integer
Dim K As Integer





'    sCadAge = "''"
'    For i = 0 To UBound(pMatAgencias) - 1
'    sCadAge = sCadAge & pMatAgencias(i) & "'',''"
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 3)
'
    'By Capi 280012008
      
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
        
    
    
    sInicio = Right(pdFecInicial, 4) & Mid(pdFecInicial, 4, 2) & Left(pdFecInicial, 2)
    sTermino = Right(pdFecFinal, 4) & Mid(pdFecFinal, 4, 2) & Left(pdFecFinal, 2)
    
    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.ObtenerREULavadoDinero(sInicio, sTermino, sCadAge)
    Set oDCred = Nothing
    
    '**************************************************************
    '**Modificado por ELRO 20110714, según acta 158-2011/TI-D
    Me.PbProgresoReporte.value = 80
    
    'lsNombreArchivo = "REU Registro de Efectivo Unico"
         
    Me.dlgFileSave.InitDir = App.Path & "\Spooler\"
    Me.dlgFileSave.Filter = "Archivos de Texto (*.xls)|*.xls"
    Me.dlgFileSave.FilterIndex = 1
    Me.dlgFileSave.ShowSave
  
    lsArregloNombreArchivo() = Split(Me.dlgFileSave.FileTitle, ".")
    lsArregloRuta() = Split(Me.dlgFileSave.FileName, "\")
 
    For j = 0 To UBound(lsArregloRuta()) - 1
    If j = 0 Then
    lsRutaArchivo = lsArregloRuta(j)
    Else
     lsRutaArchivo = lsRutaArchivo & "\" & lsArregloRuta(j)
    End If
    Next
    
    lsRutaArchivo = lsRutaArchivo & "\"
    
    For K = 0 To UBound(lsArregloNombreArchivo()) - 1
    lsNombreArchivo = lsArregloNombreArchivo(K)
    Next
    '**************************************************************
    
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Efectivo Unico", "Desde : " & (pdFecInicial) & "  Hasta : " & (pdFecFinal), lsNombreArchivo, lMatCabecera, R, 2, , , True, True, lsRutaArchivo)

    '***************************************************************
    '**Modificado por ELRO 20110714, según acta 158-2011/TI-D
    Me.PbProgresoReporte.value = 100
    Me.PbProgresoReporte.Visible = False
    '***************************************************************
   
End Sub
'By Capi 03112008
Private Sub ImprimirCreditosRefinanciados(ByVal pdFecFinal As Date, ByVal pMatAgencias As Variant, ByVal pnTipoCambio As Double, ByVal pnDiasAtraso As Integer)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim lMatCabecera(1, 1) As Variant
Dim lsmensaje As String
Dim i As Integer
Dim sCadAge As String
    
    Screen.MousePointer = 11
    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
    sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.ReporteCreditosRefinanciados(pdFecFinal, sCadAge, pnTipoCambio)
    Set oDCred = Nothing
    Call GeneraReporteCreditosRefinanciados(pdFecFinal, gsNomCmac, gsNomAge, gsCodUser, gdFecSis, pnTipoCambio, R, pnDiasAtraso)
    Screen.MousePointer = 0
    
End Sub
'By capi 03112008
Private Sub GeneraReporteCreditosRefinanciados(ByVal pdFecFin As Date, _
                                                ByVal psNomCmac As String, _
                                                ByVal psNomAge As String, _
                                                ByVal psCodUser As String, _
                                                ByVal pdFecSis As Date, _
                                                ByVal pnTipCam As Double, _
                                                ByVal pRs As ADODB.Recordset, _
                                                ByVal pnDiasAtraso As Integer)
                                               

Dim sCadImp As String, i As Integer, j As Integer
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String


Dim lnItem As Integer
Dim lnContador As Integer
Dim lccodcta As String
Dim lcCodCtaRef As String
Dim lnCuotaDificultad As Integer
Dim lcCuentaPivot As String

Dim lnAtraso6Cuotas As Double
Dim lnAtrasoMaximo As Double
Dim lnAtraso6Cuota As Double
Dim lnAtraso5Cuota As Double
Dim lnAtraso4Cuota As Double
Dim lnAtraso3Cuota As Double
Dim lnAtraso2Cuota As Double
Dim lnAtraso1Cuota As Double
   
    
If pRs.RecordCount = 0 Then
    MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
    Screen.MousePointer = 0
    Exit Sub
End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = psNomCmac
    ApExcel.Cells(3, 2).Formula = psNomAge
    ApExcel.Cells(2, 14).Formula = pdFecSis + Time()
    ApExcel.Cells(3, 14).Formula = psCodUser
    ApExcel.Range("B2", "N6").Font.Bold = True
    
    ApExcel.Cells(4, 2).Formula = "REPORTE DE CREDITOS REFINANCIADOS"
    ApExcel.Cells(5, 2).Formula = "Al: " & Format(pdFecFin, "dd/MM/YYYY")
    
    
'    ApExcel.Range("B7", "B8").MergeCells = True
'    ApExcel.Range("C7", "C8").MergeCells = True
'    ApExcel.Range("D7", "D8").MergeCells = True
'    ApExcel.Range("E7", "E8").MergeCells = True
'    ApExcel.Range("F7", "F8").MergeCells = True
'    ApExcel.Range("G7", "G8").MergeCells = True
'    ApExcel.Range("H7", "H8").MergeCells = True
'
'    ApExcel.Range("I7", "K7").MergeCells = True
'    ApExcel.Range("L7", "N7").MergeCells = True
    ApExcel.Range("B4", "N4").MergeCells = True
    ApExcel.Range("B5", "N5").MergeCells = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("N2", "N3").HorizontalAlignment = xlRight
    ApExcel.Range("B4", "N5").HorizontalAlignment = xlCenter
    ApExcel.Range("B7", "N8").VerticalAlignment = xlCenter
    ApExcel.Range("B8", "AA8").Borders.LineStyle = 1

    ApExcel.Cells(6, 2).Formula = "TIPO CAMBIO : " & ImpreFormat(pnTipCam, 5, 3)
    ApExcel.Cells(6, 14).Formula = "Parametro Dias Atraso : " & ImpreFormat(pnDiasAtraso, 5, 0)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    
   
    
    
    ApExcel.Cells(8, 2).Formula = "Item"
    ApExcel.Cells(8, 3).Formula = "Agencia"
    ApExcel.Cells(8, 4).Formula = "Nº Credito"
    ApExcel.Cells(8, 5).Formula = "Cliente"
    'ApExcel.Cells(8, 6).Formula = "Documento"
    ApExcel.Cells(8, 6).Formula = "Moneda"
    ApExcel.Cells(8, 7).Formula = "Fecha"
    ApExcel.Cells(8, 8).Formula = "Monto Refin" '**
    ApExcel.Cells(8, 9).Formula = "Monto Refin MN"
    ApExcel.Cells(8, 10).Formula = "Plazo"
    ApExcel.Cells(8, 11).Formula = "Dias Atraso"
    ApExcel.Cells(8, 12).Formula = "Cuota Dificultad" '**
    ApExcel.Cells(8, 13).Formula = "Saldo Capital"
    ApExcel.Cells(8, 14).Formula = "Saldo Capital MN"
    ApExcel.Cells(8, 15).Formula = "Estado"
    ApExcel.Cells(8, 16).Formula = "Analista"
    'ApExcel.Cells(7, 18).Formula = "PAGO 6 ULTIMAS CUOTAS"
    ApExcel.Cells(8, 17).Formula = "P6"
    ApExcel.Cells(8, 18).Formula = "P5"
    ApExcel.Cells(8, 19).Formula = "P4"
    ApExcel.Cells(8, 20).Formula = "P3"
    ApExcel.Cells(8, 21).Formula = "P2"
    ApExcel.Cells(8, 22).Formula = "P1"
    ApExcel.Cells(8, 23).Formula = "Motivo Refinan"
    'ApExcel.Cells(7, 24).Formula = "DATOS CUENTA ORIGEN"
    ApExcel.Cells(8, 24).Formula = "Cuenta "
    ApExcel.Cells(8, 25).Formula = "Monto "
    ApExcel.Cells(8, 26).Formula = "Moneda"
    ApExcel.Cells(8, 27).Formula = "Plazo"
    
    
    ApExcel.Range("B7", "AA7").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "AA7").Font.Bold = True
    ApExcel.Range("B7", "AA7").HorizontalAlignment = 3
    
    
    
    ApExcel.Range("B8:AA8").HorizontalAlignment = xlGeneral
    ApExcel.Range("B8:AA8").VerticalAlignment = xlBottom
    ApExcel.Range("B8:AA8").WrapText = True
    ApExcel.Range("B8:AA8").Orientation = 0
    ApExcel.Range("B8:AA8").AddIndent = False
    ApExcel.Range("B8:AA8").IndentLevel = 0
    ApExcel.Range("B8:AA8").ShrinkToFit = False
    ApExcel.Range("B8:AA8").ReadingOrder = xlContext
    ApExcel.Range("B8:AA8").MergeCells = False
    

    i = 8
    
    lnItem = 0
    
    Do While Not pRs.EOF
        i = i + 1
        lnAtraso6Cuota = -999
        lnAtraso5Cuota = -999
        lnAtraso4Cuota = -999
        lnAtraso3Cuota = -999
        lnAtraso2Cuota = -999
        lnAtraso1Cuota = -999
        
        lnItem = lnItem + 1
        lnCuotaDificultad = 0
        lnContador = 0
        lnAtrasoMaximo = pnDiasAtraso
        lccodcta = pRs!cCtaCod
        lcCodCtaRef = IIf(IsNull(pRs!sCuentaOrigen), "", pRs!sCuentaOrigen)
        Do While pRs!cCtaCod & pRs!sCuentaOrigen = lccodcta & lcCodCtaRef
            lnAtraso6Cuotas = IIf(IsNull(pRs!nAtraso6cuotas), 0, pRs!nAtraso6cuotas)
            'lnCuotaPivot = IIf(IsNull(prs!nCuotas6Cuotas), 0, prs!nCuotas6Cuotas)
            If lnAtraso6Cuotas >= lnAtrasoMaximo Then
                'If lnAtraso6Cuotas > lnAtrasoMaximo Then
                    lnCuotaDificultad = IIf(IsNull(pRs!nCuotas6Cuotas), 0, pRs!nCuotas6Cuotas)
                    lnAtrasoMaximo = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
                'End If
            End If
            lnContador = lnContador + 1
            Select Case lnContador
                Case 1
                    lnAtraso6Cuota = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
                Case 2
                    lnAtraso5Cuota = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
                Case 3
                    lnAtraso4Cuota = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
                Case 4
                    lnAtraso3Cuota = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
                Case 5
                    lnAtraso2Cuota = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
                Case 6
                    lnAtraso1Cuota = IIf(IsNull(pRs!nAtraso6cuotas), -999, pRs!nAtraso6cuotas)
            End Select
            
               
            If lnContador > 6 Then
                
                Do While True
                    lcCodCtaRef = IIf(IsNull(pRs!sCuentaOrigen), "", pRs!sCuentaOrigen)
                    pRs.MoveNext
                    If pRs!cCtaCod & pRs!sCuentaOrigen <> lccodcta & lcCodCtaRef Then
                        pRs.MovePrevious
                        Exit Do
                    End If
                Loop
                           
             
            End If
            pRs.MoveNext
            If pRs.EOF Then
                Exit Do
            End If
        Loop
         Select Case lnContador
                Case 5
                    lnAtraso1Cuota = lnAtraso2Cuota
                    lnAtraso2Cuota = lnAtraso3Cuota
                    lnAtraso3Cuota = lnAtraso4Cuota
                    lnAtraso4Cuota = lnAtraso5Cuota
                    lnAtraso5Cuota = lnAtraso6Cuota
                    lnAtraso6Cuota = -999
                    
                Case 4
                    lnAtraso1Cuota = lnAtraso3Cuota
                    lnAtraso2Cuota = lnAtraso4Cuota
                    lnAtraso3Cuota = lnAtraso5Cuota
                    lnAtraso4Cuota = lnAtraso6Cuota
                    lnAtraso5Cuota = -999
                    lnAtraso6Cuota = -999
                Case 3
                    lnAtraso1Cuota = lnAtraso4Cuota
                    lnAtraso2Cuota = lnAtraso5Cuota
                    lnAtraso3Cuota = lnAtraso6Cuota
                    lnAtraso4Cuota = -999
                    lnAtraso5Cuota = -999
                    lnAtraso6Cuota = -999
                Case 2
                    lnAtraso1Cuota = lnAtraso5Cuota
                    lnAtraso2Cuota = lnAtraso6Cuota
                    lnAtraso3Cuota = -999
                    lnAtraso4Cuota = -999
                    lnAtraso5Cuota = -999
                    lnAtraso6Cuota = -999
                Case 1
                    lnAtraso1Cuota = lnAtraso6Cuota
                    lnAtraso2Cuota = -999
                    lnAtraso3Cuota = -999
                    lnAtraso4Cuota = -999
                    lnAtraso5Cuota = -999
                    lnAtraso6Cuota = -999
            End Select
            
        If Not pRs.EOF Then
            lcCuentaPivot = pRs!cCtaCod
        Else
            lcCuentaPivot = ""
        End If
        pRs.MovePrevious
        
        
        
            ApExcel.Cells(i, 2).Formula = lnItem
            ApExcel.Cells(i, 3).Formula = pRs!sAgencia
            ApExcel.Cells(i, 4).Formula = "'" & pRs!cCtaCod
            ApExcel.Cells(i, 5).Formula = pRs!sCliente
            'ApExcel.Cells(i, 6).Formula = pRs!sDocumento
            ApExcel.Cells(i, 6).Formula = pRs!sMoneda
            ApExcel.Cells(i, 7).Formula = "'" & Format(pRs!dVigencia, "dd/mm/yyyy")
            ApExcel.Cells(i, 8).Formula = pRs!nMontoRefinan
            ApExcel.Cells(i, 9).Formula = pRs!nMontoRefinanMN
            ApExcel.Cells(i, 10).Formula = pRs!nPlazo
            ApExcel.Cells(i, 11).Formula = pRs!nDiasAtraso
            If lnCuotaDificultad = 0 Then
                ApExcel.Cells(i, 12).Formula = "NA"
            Else
                ApExcel.Cells(i, 12).Formula = lnCuotaDificultad
            End If
            
            ApExcel.Cells(i, 13).Formula = pRs!nSaldoCap
            ApExcel.Cells(i, 14).Formula = pRs!nSaldoCapMN
            ApExcel.Cells(i, 15).Formula = pRs!sEstado
            ApExcel.Cells(i, 16).Formula = pRs!sAnalista
            If lnAtraso1Cuota = -999 Then
                ApExcel.Cells(i, 17).Formula = "NA"
            Else
                ApExcel.Cells(i, 17).Formula = lnAtraso1Cuota
            End If
            If lnAtraso2Cuota = -999 Then
                ApExcel.Cells(i, 18).Formula = "NA"
            Else
               ApExcel.Cells(i, 18).Formula = lnAtraso2Cuota
            End If
            If lnAtraso3Cuota = -999 Then
                ApExcel.Cells(i, 19).Formula = "NA"
            Else
                ApExcel.Cells(i, 19).Formula = lnAtraso3Cuota
            End If
            If lnAtraso4Cuota = -999 Then
                ApExcel.Cells(i, 20).Formula = "NA"
            Else
                ApExcel.Cells(i, 20).Formula = lnAtraso4Cuota
            End If
            If lnAtraso5Cuota = -999 Then
                ApExcel.Cells(i, 21).Formula = "NA"
            Else
                ApExcel.Cells(i, 21).Formula = lnAtraso5Cuota
            End If
            If lnAtraso6Cuota = -999 Then
                ApExcel.Cells(i, 22).Formula = "NA"
            Else
                ApExcel.Cells(i, 22).Formula = lnAtraso6Cuota
            End If
            ApExcel.Cells(i, 23).Formula = pRs!sMotivo
            
            ApExcel.Cells(i, 24).Formula = "'" & pRs!sCuentaOrigen
            ApExcel.Cells(i, 25).Formula = pRs!nMontoOrigen
            ApExcel.Cells(i, 26).Formula = pRs!sMonedaOrigen
            ApExcel.Cells(i, 27).Formula = pRs!nPlazoOrigen
        If pRs!cCtaCod = lcCuentaPivot Then
            Do While True
                
                pRs.MoveNext
                
                'MAVM *** 20091101
                If pRs.EOF Then
                    pRs.MovePrevious
                    Exit Do
                End If
                '***
                
                lcCodCtaRef = IIf(IsNull(pRs!sCuentaOrigen), "", pRs!sCuentaOrigen)
                
                If pRs!cCtaCod & pRs!sCuentaOrigen <> lccodcta & lcCodCtaRef Or pRs.EOF Then
                    pRs.MovePrevious
                    Exit Do
                End If
            Loop
            ApExcel.Cells(i + 1, 24).Formula = "'" & pRs!sCuentaOrigen
            ApExcel.Cells(i + 1, 25).Formula = pRs!nMontoOrigen
            ApExcel.Cells(i + 1, 26).Formula = pRs!sMonedaOrigen
            ApExcel.Cells(i + 1, 27).Formula = pRs!nPlazoOrigen
            i = i + 1
        End If
'        If substring(pRs!cCtaCod, 4, 2) <> Mid(lcCuentaPivot, 4, 2) Then
'            i = i + 2
'            ApExcel.Cells(i, 2).Formula = "Totales por Agencia..: " & Trim(Str(lcCuenta))
'            ApExcel.Cells(i, 2).Formula = "Totales por Agencia..: " & Trim(Str(lcCuenta))
'            ApExcel.Cells(i, 2).Font.Bold = True
'        ElseIf lcCuentaPivot = "" Then
'            i = i + 2
'            ApExcel.Cells(i, 2).Formula = "Totales Generales..: " & Trim(Str(lcCuenta))
'            ApExcel.Cells(i, 2).Font.Bold = True
'        End If
        pRs.MoveNext
        If pRs.EOF Then
            Exit Do
        End If
        
    Loop
    ApExcel.Columns("B:AA").EntireColumn.AutoFit
    ApExcel.Cells(7, 19).Formula = "PAGO 6 ULTIMAS CUOTAS"
    ApExcel.Cells(7, 25).Formula = "DATOS CUENTA ORIGEN"

    
    
'    I = I + 2
'
'    ApExcel.Cells(I, 2).Formula = "Número Total de Clientes: " & Trim(Str(lcCuenta))
'    ApExcel.Cells(I, 2).Font.Bold = True
'
    pRs.Close
    Set pRs = Nothing
   
'    ApExcel.Cells.Select
'    ApExcel.Cells.EntireColumn.AutoFit
''    ApExcel.Columns("B:B").ColumnWidth = 6#
'    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'***PEAC 20080102
Private Sub ImprimePoliIncendio(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pdFecFin As Date, Optional ByVal psNomCmac As String = "")

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant


    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosListadoPoliIncendio(pMatAgencias, pMatProd, pdFecIni, pdFecFin)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Sub
    End If

    lsNombreArchivo = "POLIZAS CONTRA INCENDIO"
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "POLIZAS CONTRA INCENDIO", " Al " & CStr(gdFecSis), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

End Sub

'peac 20080303
Private Sub ImprimeRepAECIReporte01(ByVal pdFecInicial As Date, ByVal pdFecFinal As Date, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psNomAge As String, _
    Optional ByVal psNomCmac As String = "")

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant

    Screen.MousePointer = 11
    
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosAECIReporte01(Format(pdFecInicial, "yyyymmdd"), Format(pdFecFinal, "yyyymmdd"))
    Set oDCred = Nothing

    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    lsNombreArchivo = "AECIReporte01"
    
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, _
    "DATOS DE LAS OPERACIONES - AECI", " Del " & CStr(pdFecInicial) & " Al " & CStr(pdFecFinal), lsNombreArchivo, _
    lMatCabecera, R, 2, , , True, True)

    Screen.MousePointer = 0

End Sub


'*** PEAC 20080102 ' MADM 20110506 - SBS
Private Sub ImprimeRepCredVencPaseCastigo(ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, _
    ByVal pMatAgencias As Variant, ByVal pnUIT As Double, ByVal pnMontoIni As Double, ByVal pnMontoFin As Double, ByVal pnTipCam As Double, _
    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")

'ByVal pMatAgencias As Variant, ByVal pnUit As Double, ByVal pnTipCam As Double, _
TxtMontoIni
Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String, i As Integer, j As Integer
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String

If pnDiaFin = 0 Or pnUIT <= 0 Or pnTipCam <= 0 Then
    MsgBox "Ingrese datos correctos para proseguir.", vbInformation, "Atención"
    Exit Sub
End If

If pnDiaIni > pnDiaFin Then
    MsgBox "Ingrese los Rangos correctamente.", vbInformation, "Atención"
    Exit Sub
End If
    
    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosCredVencPaseCastigo(pnDiaIni, pnDiaFin, pnUIT, pnMontoIni, pnMontoFin, pnTipCam, pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CREDITOS VENCIDOS PARA PASE A CASTIGO AL " & Format(gdFecSis, "dd/MM/YYYY")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B2", "W2").MergeCells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).HorizontalAlignment = 3
    
    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & ImpreFormat(pnTipCam, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "CREDITO"
    ApExcel.Cells(8, 4).Formula = "COD.CLIENTE"
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "FEC.VIGENCIA"
    ApExcel.Cells(7, 7).Formula = "MONTO" '**
    ApExcel.Cells(8, 7).Formula = "COLOCADO"
    ApExcel.Cells(8, 8).Formula = "ANALISTA"
    ApExcel.Cells(8, 9).Formula = "DOC. ID."
    ApExcel.Cells(8, 10).Formula = "DIRECCION"
    ApExcel.Cells(8, 11).Formula = "COD. S.B.S."
    ApExcel.Cells(7, 12).Formula = "DIAS" '**
    ApExcel.Cells(8, 12).Formula = "ATRASO"
    ApExcel.Cells(8, 13).Formula = "PROVISION"
    ApExcel.Cells(8, 14).Formula = "CAPITAL"
    ApExcel.Cells(8, 15).Formula = "INTERES"
    ApExcel.Cells(8, 16).Formula = "MORA"
    ApExcel.Cells(8, 17).Formula = "GASTOS"
    ApExcel.Cells(8, 18).Formula = "TOTAL"
    ApExcel.Cells(8, 19).Formula = "TOTAL MN"
    ApExcel.Cells(8, 20).Formula = "PRODUCTO"
    ApExcel.Cells(8, 21).Formula = "LINEA"
    ApExcel.Cells(8, 22).Formula = "CALIFICACION"
    ApExcel.Cells(8, 23).Formula = "NESTADO"    'MADM 20110503
    ApExcel.Cells(8, 24).Formula = "SALDO"      'MADM 20110503
    ApExcel.Cells(8, 25).Formula = "TASA INT"   'MADM 20110503
    ApExcel.Cells(8, 26).Formula = "CALENDARIO" 'MADM 20110503
    ApExcel.Cells(8, 27).Formula = "RANGO UIT"
'    ApExcel.Cells(8, 28).Formula = "ULT. ACTUALIZACION"
    
'    ApExcel.Range("B7", "W8").Interior.Color = RGB(10, 190, 160)
'    ApExcel.Range("B7", "W8").Font.Bold = True
'    ApExcel.Range("B7", "W8").HorizontalAlignment = 3
    ApExcel.Range("B7", "AA8").Interior.Color = RGB(10, 190, 160) 'MADM 20110503
    ApExcel.Range("B7", "AA8").Font.Bold = True                   'MADM 20110503
    ApExcel.Range("B7", "AA8").HorizontalAlignment = 3            'MADM 20110503

    i = 8

    Do While Not R.EOF
    i = i + 1
    ApExcel.Cells(i, 2).Formula = "AGENCIA : " & R!NomAge
    ApExcel.Cells(i, 2).Font.Bold = True
    vage = R!Agencia
    
        Do While R!Agencia = vage
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "MONEDA : " & R!NomMone
        ApExcel.Cells(i, 2).Font.Bold = True
        vmone = R!Moneda
             j = 0
            Do While R!Moneda = vmone And R!Agencia = vage
                j = j + 1
                i = i + 1
                ApExcel.Cells(i, 2).Formula = j
                ApExcel.Cells(i, 3).Formula = "'" & R!cCtaCod
                ApExcel.Cells(i, 4).Formula = "'" & R!cPersCod
                ApExcel.Cells(i, 5).Formula = R!nomcli
                ApExcel.Cells(i, 6).Formula = Format(R!dVigencia, "mm/dd/yyyy")
                ApExcel.Cells(i, 7).Formula = R!nMontoCol
                ApExcel.Cells(i, 8).Formula = R!Analista
                ApExcel.Cells(i, 9).Formula = R!numdocId
                ApExcel.Cells(i, 10).Formula = R!cPersDireccDomicilio
                ApExcel.Cells(i, 11).Formula = R!CodSbs
                ApExcel.Cells(i, 12).Formula = R!nDiasAtraso
                ApExcel.Cells(i, 13).Formula = R!Provision
                ApExcel.Cells(i, 14).Formula = R!Capital
                ApExcel.Cells(i, 15).Formula = R!Interes
                ApExcel.Cells(i, 16).Formula = R!Mora
                ApExcel.Cells(i, 17).Formula = R!Gastos
                ApExcel.Cells(i, 18).Formula = R!Total
                ApExcel.Cells(i, 19).Formula = R!TotalMN
                ApExcel.Cells(i, 20).Formula = R!Producto
                ApExcel.Cells(i, 21).Formula = R!Linea
                ApExcel.Cells(i, 22).Formula = R!Calificacion
                ApExcel.Cells(i, 23).Formula = R!nPrdEstado 'MADM 20110503
                ApExcel.Cells(i, 24).Formula = R!nSaldo     'MADM 20110503
                ApExcel.Cells(i, 25).Formula = R!nTasaInteres 'MADM 20110503
                ApExcel.Cells(i, 26).Formula = R!nNroCalen  'MADM 20110503
                ApExcel.Cells(i, 27).Formula = R!rangouit
'                ApExcel.Cells(i, 28).Formula = R!cUltimaActualizacion
                
                ApExcel.Range("G" & Trim(str(i)) & ":" & "S" & Trim(str(i))).NumberFormat = "#,##0.00"
                'ApExcel.Range("B" & Trim(str(i)) & ":" & "W" & Trim(str(i))).Borders.LineStyle = 1
                ApExcel.Range("B" & Trim(str(i)) & ":" & "AA" & Trim(str(i))).Borders.LineStyle = 1
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                               
            Loop
            
            i = i + 1
            ApExcel.Cells(i, 13).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(i, 14).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(i, 15).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(i, 16).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(i, 17).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(i, 18).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"
            ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(str(Int(j))) & "]C:R[-1]C)"

'            ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 21).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
            
            i = i + 1
            
            If R.EOF Then
                Exit Do
            End If
            
        Loop
    Loop
    
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing


'**************************************************************************************
    
'    lsNombreArchivo = "CredVcdosPaseCast"
'    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "CREDITOS VENCIDOS PARA PASE A CASTIGO", " Al " & CStr(gdFecSis), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

End Sub

'By MAVM 18052009 Memorandum Nº 024-2009-TI-GA
Private Sub ImprimeRepCredVencMay90Dias(ByVal pnDiaIni As Integer, ByVal pnDiaFin As Integer, _
    ByVal pMatAgencias As Variant, ByVal pnTipCam As Double, ByVal sFI As String, ByVal sFF As String, _
    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String, i As Integer, j As Integer
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String

If pnDiaFin = 0 Or pnTipCam <= 0 Then
    MsgBox "Ingrese datos correctos para proseguir.", vbInformation, "Atención"
    Exit Sub
End If

If pnDiaIni > pnDiaFin Then
    MsgBox "Ingrese los Rangos correctamente.", vbInformation, "Atención"
    Exit Sub
End If
    
    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosCredVencMay90Dias(pnDiaIni, pnDiaFin, pnTipCam, pMatAgencias, sFI, sFF)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "CREDITOS VENCIDOS MAYORES A 90 DIAS AL " & Format(gdFecSis, "dd/MM/YYYY")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B2", "P2").MergeCells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).HorizontalAlignment = 3
    
    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & ImpreFormat(pnTipCam, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "AGENCIA"
    ApExcel.Cells(8, 4).Formula = "CREDITO"
    ApExcel.Cells(8, 5).Formula = "COD.CLIENTE"
    ApExcel.Cells(8, 6).Formula = "CLIENTE"
    ApExcel.Cells(8, 7).Formula = "FEC.VIGENCIA"
    ApExcel.Cells(7, 8).Formula = "MONTO"
    ApExcel.Cells(8, 8).Formula = "DESEMBOLSADO"
    
    ApExcel.Cells(8, 9).Formula = "SALDO K"
    
    ApExcel.Cells(8, 10).Formula = "CUOTAS"
    
    ApExcel.Cells(7, 11).Formula = "DIAS" '**
    ApExcel.Cells(8, 11).Formula = "ATRASO"
    
    ApExcel.Cells(8, 12).Formula = "ANALISTA"
    
    ApExcel.Cells(8, 13).Formula = "PRODUCTO"
    
    ApExcel.Cells(8, 14).Formula = "CONDICION"
    ApExcel.Cells(8, 15).Formula = "CAMPAÑA"
    
    ApExcel.Range("B7", "P8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "P8").Font.Bold = True
    ApExcel.Range("B7", "P8").HorizontalAlignment = 3

    i = 8

    Do While Not R.EOF
    i = i + 1
    ApExcel.Cells(i, 2).Formula = "AGENCIA : " & R!NomAge
    ApExcel.Cells(i, 2).Font.Bold = True
    vage = R!NomAge
    
        Do While R!NomAge = vage
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "MONEDA : " & R!NomMone
        ApExcel.Cells(i, 2).Font.Bold = True
        vmone = R!NomMone
             j = 0
            Do While R!NomMone = vmone And R!NomAge = vage
                j = j + 1
                i = i + 1
                ApExcel.Cells(i, 2).Formula = j
                
                ApExcel.Cells(i, 3).Formula = "'" & R!NomAge
                ApExcel.Cells(i, 4).Formula = "'" & R!cCtaCod
                ApExcel.Cells(i, 5).Formula = "'" & R!cPersCod
                ApExcel.Cells(i, 6).Formula = R!NomCliente
                ApExcel.Cells(i, 7).Formula = Format(R!dVigencia, "mm/dd/yyyy")
                ApExcel.Cells(i, 8).Formula = R!nMonto
                ApExcel.Cells(i, 9).Formula = R!nSaldo
                ApExcel.Cells(i, 10).Formula = R!nCuotas
                ApExcel.Cells(i, 11).Formula = R!nDiasAtraso
                ApExcel.Cells(i, 12).Formula = R!Analista
                ApExcel.Cells(i, 13).Formula = R!Producto
                
                ApExcel.Cells(i, 14).Formula = R!cCondicion
                ApExcel.Cells(i, 15).Formula = R!Campanha
                
                ApExcel.Range("H" & Trim(str(i)) & ":" & "I" & Trim(str(i))).NumberFormat = "#,##0.00"
                ApExcel.Range("B" & Trim(str(i)) & ":" & "P" & Trim(str(i))).Borders.LineStyle = 1
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                               
            Loop
            
            i = i + 1
'            ApExcel.Cells(i, 13).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 14).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 15).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 16).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 17).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 18).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"

'            ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 21).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
            
            i = i + 1
            
            If R.EOF Then
                Exit Do
            End If
            
        Loop
    Loop
    
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing


'**************************************************************************************
    
'    lsNombreArchivo = "CredVcdosPaseCast"
'    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "CREDITOS VENCIDOS PARA PASE A CASTIGO", " Al " & CStr(gdFecSis), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

End Sub

Private Sub ImprimeRepCredClientesPreferenciales(ByVal pMatAgencias As Variant, _
    ByVal pMatProducts As Variant, ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String, i As Integer, j As Integer
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
   
    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaDatosClientesPreferenciales(pMatAgencias, pMatProducts)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "REPORTE DE CLIENTES PREFERENCIALES AL " & Format(gdFecSis, "dd/MM/YYYY")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B2", "K2").MergeCells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).HorizontalAlignment = 3
    
    'ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & ImpreFormat(pntipCam, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "COD.CLIENTE"
    ApExcel.Cells(8, 4).Formula = "CLIENTE"
    ApExcel.Cells(8, 5).Formula = "COD. CUENTA"
    ApExcel.Cells(8, 6).Formula = "DOMICILIO"
    ApExcel.Cells(8, 7).Formula = "DESEMBOLSO"
    ApExcel.Cells(8, 8).Formula = "SALDO K"
    ApExcel.Cells(8, 9).Formula = "CUOTAS"
    ApExcel.Cells(8, 10).Formula = "ANALISTA"
    ApExcel.Cells(8, 11).Formula = "CANT. DEUDA OTROS BCOS"
    
    ApExcel.Range("B7", "K8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "K8").Font.Bold = True
    ApExcel.Range("B7", "K8").HorizontalAlignment = 3

    i = 8

    Do While Not R.EOF
    i = i + 1
    'ApExcel.Cells(i, 2).Formula = "AGENCIA : " & R!NomAge
    'ApExcel.Cells(i, 2).Font.Bold = True
    'vage = R!NomAge
    
        'Do While R!NomAge = vage
        'i = i + 1
        'ApExcel.Cells(i, 2).Formula = "MONEDA : " & R!NomMone
        'ApExcel.Cells(i, 2).Font.Bold = True
        'vmone = R!NomMone
            'J = 0
            'Do While R!NomMone = vmone And R!NomAge = vage
                j = j + 1
                'i = i + 1
                ApExcel.Cells(i, 2).Formula = j
                
                'ApExcel.Cells(i, 3).Formula = "'" & R!NomAge
                'ApExcel.Cells(i, 4).Formula = "'" & R!cCtaCod
                ApExcel.Cells(i, 3).Formula = "'" & R!cPersCod
                ApExcel.Cells(i, 4).Formula = R!Cliente
                
                ApExcel.Cells(i, 5).Formula = R!cCtaCod
                ApExcel.Cells(i, 6).Formula = R!Domicilio
                ApExcel.Cells(i, 7).Formula = R!nMontoCol
                ApExcel.Cells(i, 8).Formula = R!nSaldo
                ApExcel.Cells(i, 9).Formula = R!nCuotas
                ApExcel.Cells(i, 10).Formula = R!Analista
                ApExcel.Cells(i, 11).Formula = R!CantDeudaBanco
                

'                ApExcel.Cells(i, 7).Formula = Format(R!dVigencia, "mm/dd/yyyy")
'                ApExcel.Cells(i, 8).Formula = R!nMonto
'                ApExcel.Cells(i, 9).Formula = R!nSaldo
'                ApExcel.Cells(i, 10).Formula = R!nCuotas
'                ApExcel.Cells(i, 11).Formula = R!nDiasAtraso
'                ApExcel.Cells(i, 12).Formula = R!Analista
'                ApExcel.Cells(i, 13).Formula = R!Producto
'
'                ApExcel.Cells(i, 14).Formula = R!cCondicion
'                ApExcel.Cells(i, 15).Formula = R!Campanha
                
                ApExcel.Range("G" & Trim(str(i)) & ":" & "H" & Trim(str(i))).NumberFormat = "#,##0.00"
                ApExcel.Range("B" & Trim(str(i)) & ":" & "K" & Trim(str(i))).Borders.LineStyle = 1
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                               
            Loop
            
            i = i + 1
'            ApExcel.Cells(i, 13).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 14).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 15).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 16).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 17).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 18).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"

'            ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
'            ApExcel.Cells(i, 21).Formula = "=SUM(R[-" & Trim(Str(Int(J))) & "]C:R[-1]C)"
            
            i = i + 1
            
'            If R.EOF Then
'                Exit Do
'            End If
            
        'Loop
    'Loop
    
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing


'**************************************************************************************
    
'    lsNombreArchivo = "CredVcdosPaseCast"
'    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "CREDITOS VENCIDOS PARA PASE A CASTIGO", " Al " & CStr(gdFecSis), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

End Sub

'By Capi 30012008
Private Sub mostrarDUDLavadoDinero(ByVal pdFecInicial As String, ByVal pdFecFinal As String, ByVal pMatAgencias As Variant)

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant
Dim lsmensaje As String
Dim i As Integer
Dim sCadAge As String
Dim sInicio As String
Dim sTermino As String



'    sCadAge = "''"
'    For i = 0 To UBound(pMatAgencias) - 1
'    sCadAge = sCadAge & pMatAgencias(i) & "'',''"
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 3)
'
    'By Capi 280012008
      
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
        
    
    
    sInicio = Right(pdFecInicial, 4) & Mid(pdFecInicial, 4, 2) & Left(pdFecInicial, 2)
    sTermino = Right(pdFecFinal, 4) & Mid(pdFecFinal, 4, 2) & Left(pdFecFinal, 2)
    
    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.ObtenerDUDLavadoDinero(sInicio, sTermino, sCadAge)
    Set oDCred = Nothing
    
    lsNombreArchivo = "Registro Clientes Dudosos"
            
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Clientes Con Motivos de Lavado de Dinero", "Desde : " & (pdFecInicial) & "  Hasta : " & (pdFecFinal), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

End Sub



' Esta funcion deberá ser puesto en el modulo gCredFunciones

'** PEAC 20071003
Public Function ImprimeActaComiteCredApro( _
ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
ByVal psCodAge As String, ByVal psCodUser As String, _
ByVal pdFecSis As Date, _
ByVal psNomAge As String, _
ByVal pMatProd As Variant, _
Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim vProd As String
Dim vEstado As String
Dim sLetraUlt As String
'Dim sHora As String
'Dim nHora As Integer
'Dim sTHora As String
Dim i As Integer, j As Integer

    Set oDCred = New COMDCredito.DCOMCredDoc 'DCredDoc
    Set R = oDCred.RecuperaActaComiteCredApro(psCodAge, pdFecIni, pdFecFin, pMatProd)
    Set oDCred = Nothing

    If R.RecordCount = 0 Then
        'Print MsgBox("No existen Datos para este Reporte.", vbOKOnly + 48, "Atención")
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If


    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(1, 1).Formula = psNomCmac
    ApExcel.Cells(1, 1).Font.Bold = True
    ApExcel.Cells(2, 1).Formula = psNomAge
    ApExcel.Cells(2, 1).Font.Bold = True
'    sHora = Format(Time(), "HH")
'    sTHora = "AM"
'    If CInt(sHora) >= 12 Then
'        nHora = CInt(sHora) - 12
'        sTHora = "PM"
'        If Len(Trim(nHora)) = 1 Then
'            sHora = "0" & Trim(Str(nHora))
'        Else
'            sHora = Trim(Str(nHora))
'        End If
'    End If
    
    ApExcel.Cells(1, 11).Formula = Format(gdFecSis, "yyyy/mm/dd ") & Format(Time(), "Medium Time")
    ApExcel.Cells(2, 11).Formula = gsCodUser
    'ApExcel.Cells(2, 2).Formula = "ACTAS DE COMITE DE CREDITOS APROBADOS DEL " & Format(pdFecIni, "dd/MM/YYYY") & " AL " & Format(pdFecFin, "dd/MM/YYYY")
    ApExcel.Cells(3, 1).Formula = "REPORTE ACTAS DE COMITE DE CREDITOS DEL " & Format(pdFecIni, "dd/MM/YYYY") & " AL " & Format(pdFecFin, "dd/MM/YYYY")
    ApExcel.Cells(4, 1).Formula = "DEL " & Format(pdFecIni, "dd/MM/YYYY") & " AL " & Format(pdFecFin, "dd/MM/YYYY")
     ApExcel.Cells(3, 1).Font.Bold = True
     ApExcel.Cells(4, 1).Font.Bold = True
    ApExcel.Range("A3:A4").HorizontalAlignment = 3
     ApExcel.Range("A3", "J3").MergeCells = True
     ApExcel.Range("A4", "J4").MergeCells = True
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Cells(2, 1).Font.Bold = True
    ApExcel.Cells(2, 1).HorizontalAlignment = 3
           
    
           
'    ApExcel.Cells(8, 2).Formula = "Nº"
'    ApExcel.Cells(8, 3).Formula = "CREDITO"
'    ApExcel.Cells(8, 4).Formula = "CLIENTE"
'    ApExcel.Cells(7, 5).Formula = "FECHA"
'    ApExcel.Cells(8, 5).Formula = "APROBACION"
'    ApExcel.Cells(7, 6).Formula = "MONTO"
'    ApExcel.Cells(8, 6).Formula = "APROBADO"
'    ApExcel.Cells(8, 7).Formula = "CUOTAS"
'    ApExcel.Cells(8, 8).Formula = "PLAZO"
'    ApExcel.Cells(8, 10).Formula = "LINEA"
'    ApExcel.Cells(8, 11).Formula = "ANALISTA"
    

    i = 7
    vEstado = ""
    Do While Not R.EOF
        If vEstado <> R!nPrdEstado Then
            i = i + 2
            ApExcel.Cells(i, 1).Formula = "CREDITOS : " & R!cPrdEstado
            ApExcel.Cells(i, 1).Font.Bold = True
            ''''ApExcel.Range("A" & i, "A" & i).HorizontalAlignment = 2
            i = i + 1
            'ALPA 20081007***********************************************************************
            If R!nPrdEstado = 2002 Then
                ApExcel.Cells(i, 1).Formula = "Nº"
                ApExcel.Cells(i, 2).Formula = "Cuenta"
                ApExcel.Cells(i, 3).Formula = "Nombre de"
                ApExcel.Cells(i, 3).Formula = "Cliente"
                ApExcel.Cells(i - 1, 4).Formula = "Fecha de"
                ApExcel.Cells(i, 4).Formula = "aprobación"
                ApExcel.Cells(i - 1, 5).Formula = "Monto"
                ApExcel.Cells(i, 5).Formula = "aprobado"
                ApExcel.Cells(i, 6).Formula = "Cuotas"
                ApExcel.Cells(i - 1, 7).Formula = "Tipo"
                ApExcel.Cells(i, 7).Formula = "plazo"
                ApExcel.Cells(i - 1, 8).Formula = "Plazo/"
                ApExcel.Cells(i, 8).Formula = "Fecha fija"
                ApExcel.Cells(i, 9).Formula = "Linea"
                ApExcel.Cells(i, 10).Formula = "Analista"
                sLetraUlt = "J"
            ElseIf R!nPrdEstado = 2003 Then
                ApExcel.Cells(i, 1).Formula = "Nº"
                ApExcel.Cells(i, 2).Formula = "Cuenta"
                ApExcel.Cells(i, 3).Formula = "Nombre de"
                ApExcel.Cells(i, 3).Formula = "Cliente"
                ApExcel.Cells(i - 1, 4).Formula = "Fecha "
                ApExcel.Cells(i, 4).Formula = "rechazo"
                ApExcel.Cells(i - 1, 5).Formula = "Monto"
                ApExcel.Cells(i, 5).Formula = "rechazo"
                ApExcel.Cells(i, 6).Formula = "Cuotas"
                ApExcel.Cells(i - 1, 7).Formula = "Tipo"
                ApExcel.Cells(i, 7).Formula = "plazo"
                ApExcel.Cells(i - 1, 8).Formula = "Plazo/"
                ApExcel.Cells(i, 8).Formula = "Fecha fija"
                ApExcel.Cells(i, 9).Formula = "Linea"
                ApExcel.Cells(i, 10).Formula = "Analista"
                ApExcel.Cells(i, 11).Formula = "Motivo"
                sLetraUlt = "K"
            ElseIf R!nPrdEstado = 2080 Then
                  ApExcel.Cells(i, 1).Formula = "Nº"
                ApExcel.Cells(i, 2).Formula = "Cuenta"
                ApExcel.Cells(i, 3).Formula = "Nombre de"
                ApExcel.Cells(i, 3).Formula = "Cliente"
                ApExcel.Cells(i - 1, 4).Formula = "Fecha "
                ApExcel.Cells(i, 4).Formula = "retiro"
                ApExcel.Cells(i - 1, 5).Formula = "Monto"
                ApExcel.Cells(i, 5).Formula = "retiro"
                ApExcel.Cells(i, 6).Formula = "Cuotas"
                ApExcel.Cells(i - 1, 7).Formula = "Tipo"
                ApExcel.Cells(i, 7).Formula = "plazo"
                ApExcel.Cells(i - 1, 8).Formula = "Plazo/"
                ApExcel.Cells(i, 8).Formula = "Fecha fija"
                ApExcel.Cells(i, 9).Formula = "Linea"
                ApExcel.Cells(i, 10).Formula = "Analista"
                ApExcel.Cells(i, 11).Formula = "Motivo"
                sLetraUlt = "K"
            End If
            '*******************************************************************************************
            'devCelda
            ApExcel.Range("A" & (i - 1), sLetraUlt & i).Interior.Color = RGB(10, 190, 160)
            ApExcel.Range("A" & (i - 1), sLetraUlt & i).Font.Bold = True
            ApExcel.Range("A" & (i - 1), sLetraUlt & i).HorizontalAlignment = 3
            vProd = ""
        End If
        vEstado = R!nPrdEstado
        i = i + 1
                ApExcel.Cells(i, 1).Formula = "PRODUCTO : " & R!Producto
             
             ApExcel.Cells(i, 1).Font.Bold = True
               vProd = R!Producto
             
            
             j = 0
             Do While R!Producto = vProd
                 j = j + 1
                 i = i + 1
                     
                     ApExcel.Cells(i, 1).Formula = j
                     ApExcel.Cells(i, 2).Formula = "'" & R!cCtaCod
                     ApExcel.Cells(i, 3).Formula = R!Cliente
                     ApExcel.Cells(i, 4).Formula = Format(R!dPrdEstado, "mm/dd/yyyy")
                     ApExcel.Cells(i, 5).Formula = R!nMonto
                     ApExcel.Cells(i, 6).Formula = R!nCuotas
                     ApExcel.Cells(i, 7).Formula = R!cConsDescripcion
                     ApExcel.Cells(i, 8).Formula = R!Plazo
                     ApExcel.Cells(i, 9).Formula = R!Descri_Linea
                     ApExcel.Cells(i, 10).Formula = R!Analista
                     ApExcel.Cells(i, 11).Formula = R!MotivoRechazo
                     
                     ApExcel.Range("E" & Trim(str(i)) & ":" & "E" & Trim(str(i))).NumberFormat = "#,##0.00"
                     ApExcel.Range("A" & Trim(str(i)) & ":" & sLetraUlt & Trim(str(i))).Borders.LineStyle = 1
                     
                     R.MoveNext
                     If R.EOF Then
                         Exit Do
                     End If
                                               
                 Loop
            Loop
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("A:A").ColumnWidth = 6#
    ApExcel.Range("A2").Select
    ApExcel.Range("A8:A10000").HorizontalAlignment = 2
    ApExcel.Range("A2:A2").HorizontalAlignment = 2
    ApExcel.Range("K1:K1").HorizontalAlignment = 4
    ApExcel.Range("K2:K2").HorizontalAlignment = 4
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Function

'peac 20071019

Public Function ObtieneCondi(ByVal Condi As String) As String

    If FraCondicion.Visible = True Then
        Dim j As Integer, xCondicion As String
        xCondicion = ""
        For j = 0 To 5
            If ChkCond(j).value = 1 Then
                Select Case j
                    Case 0
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "1"
                        Else
                            xCondicion = xCondicion & "," & "1"
                        End If
                    Case 1
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "3"
                        Else
                            xCondicion = xCondicion & "," & "3"
                        End If
                        
                    Case 2
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "4"
                        Else
                            xCondicion = xCondicion & "," & "4"
                        End If
                        
                    Case 3
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "2"
                        Else
                            xCondicion = xCondicion & "," & "2"
                        End If
                        
                    Case 4
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "5"
                        Else
                            xCondicion = xCondicion & "," & "5"
                        End If
                        
                    Case 5
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "6"
                        Else
                            xCondicion = xCondicion & "," & "6"
                        End If
                        
                    'JUEZ 20130604 ******************************
                    Case 6
                        If Len(Trim(xCondicion)) = 0 Then
                            xCondicion = "7"
                        Else
                            xCondicion = xCondicion & "," & "7"
                        End If
                    'END JUEZ ***********************************
                        
                End Select
            End If
        Next
        ObtieneCondi = xCondicion
    End If

End Function

'**DAOR 20071210
'**Muestra el reporte de seguro de desgravamen en formato Excel
Public Sub MostrarReporteSeguroDesgravamen(pdFechaProc As Date, pnTipCamb As Double, Optional pdFechaProc01 As Date = "31-12-2010", Optional ByVal psMoneda As String = "1", Optional ByVal psIndica As Boolean = False, Optional ByVal psTipIndica As Boolean = False, Optional ByVal psTipCartera As Boolean = False)
Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsmensaje As String

    lsNombreArchivo = "SeguroDesgravamen"
    
    ReDim lMatCabecera(45, 2)
'MADM 20110215 - ASEGURADORA ACE
    lMatCabecera(0, 0) = "TRegistro"
    lMatCabecera(1, 0) = "cCtaCod"
    lMatCabecera(2, 0) = "cCodOfi"
    lMatCabecera(3, 0) = "cCodCli"
    lMatCabecera(4, 0) = "cClaPer"
    lMatCabecera(5, 0) = "nNroCred"
    lMatCabecera(6, 0) = "cCodAseg"
    lMatCabecera(7, 0) = "ApeAseg"
    lMatCabecera(8, 0) = "NomAseg"
    lMatCabecera(9, 0) = "NuDoAseg"
    lMatCabecera(10, 0) = "dNacAseg"
    lMatCabecera(11, 0) = "cEdad_Aseg"
    lMatCabecera(12, 0) = "TipoCober"
    lMatCabecera(13, 0) = "cNomCli"
    lMatCabecera(14, 0) = "cNudoci"
    lMatCabecera(15, 0) = "nMonApr"
    lMatCabecera(16, 0) = "nCuoApr"
    lMatCabecera(17, 0) = "dFecVig"
    lMatCabecera(18, 0) = "nSalCap"
    lMatCabecera(19, 0) = "nIntere"
    lMatCabecera(20, 0) = "Total"
    lMatCabecera(21, 0) = "CapiMN"
    lMatCabecera(22, 0) = "dNacimi"
    lMatCabecera(23, 0) = "cEdad"
    lMatCabecera(24, 0) = "cDirDom"
    lMatCabecera(25, 0) = "cMoneda"
    lMatCabecera(26, 0) = "cConTab"
    lMatCabecera(27, 0) = "cSgrDsg"
    lMatCabecera(28, 0) = "Codeudor"
    lMatCabecera(29, 0) = "cNomCli_Cod"
    lMatCabecera(30, 0) = "cNudoci_Cod"
    lMatCabecera(31, 0) = "dNacimi_Cod"
    lMatCabecera(32, 0) = "cEdad_Cod"
    lMatCabecera(33, 0) = "Conyugue"
    lMatCabecera(34, 0) = "cNomCli_Con"
    lMatCabecera(35, 0) = "cNudci_Con"
    lMatCabecera(36, 0) = "dNacimi_Con"
    lMatCabecera(37, 0) = "cEdad_Con"
    lMatCabecera(38, 0) = "Representante"
    lMatCabecera(39, 0) = "cNomCli_Rep"
    lMatCabecera(40, 0) = "cNudoci_Rep"
    lMatCabecera(41, 0) = "dNacimi_Rep"
    lMatCabecera(42, 0) = "Aprobado_Por"
    lMatCabecera(43, 0) = "cNomApoderado"
    lMatCabecera(44, 0) = "cMancomunada"
    lMatCabecera(45, 0) = "nPrdPersRelac"
'END MADM

'    lMatCabecera(0, 0) = "cCtaCod"
'    lMatCabecera(1, 0) = "cCodOfi"
'    lMatCabecera(2, 0) = "cCodCli"
'    lMatCabecera(3, 0) = "cClaPer"
'    lMatCabecera(4, 0) = "nNroCred"
'    lMatCabecera(5, 0) = "cNomCli"
'    lMatCabecera(6, 0) = "cNudoci"
'    lMatCabecera(7, 0) = "nMonApr"
'    lMatCabecera(8, 0) = "nCuoApr"
'    lMatCabecera(9, 0) = "dFecVig"
'    lMatCabecera(10, 0) = "nSalCap"
'    lMatCabecera(11, 0) = "nIntere"
'    lMatCabecera(12, 0) = "Total"
'    lMatCabecera(13, 0) = "CapiMN"
'    lMatCabecera(14, 0) = "dNacimi"
'    lMatCabecera(15, 0) = "cEdad"
'    lMatCabecera(16, 0) = "cDirDom"
'    lMatCabecera(17, 0) = "cMoneda"
'    lMatCabecera(18, 0) = "cConTab"
'    lMatCabecera(19, 0) = "cSgrDsg"
'    lMatCabecera(20, 0) = "Codeudor"
'    lMatCabecera(21, 0) = "cNomCli_Cod"
'    lMatCabecera(22, 0) = "cNudoci_Cod"
'    lMatCabecera(23, 0) = "dNacimi_Cod"
'    lMatCabecera(24, 0) = "cEdad_Cod"
'    lMatCabecera(25, 0) = "Conyugue"
'    lMatCabecera(26, 0) = "cNomCli_Con"
'    lMatCabecera(27, 0) = "cNudci_Con"
'    lMatCabecera(28, 0) = "dNacimi_Con"
'    lMatCabecera(29, 0) = "cEdad_Con"
'    lMatCabecera(30, 0) = "Representante"
'    lMatCabecera(31, 0) = "cNomCli_Rep"
'    lMatCabecera(32, 0) = "cNudoci_Rep"
'    lMatCabecera(33, 0) = "dNacimi_Rep"
    
    Set oDCred = New COMDCredito.DCOMCreditos
    pdFechaProc = CDate(TxtFecFinA02.Text)
    pdFechaProc01 = CDate(TxtFecIniA02.Text)
    
    If Len(psMoneda) = 0 Then
        MsgBox "Seleccione por lo menos una moneda.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    'MADM 20110627 - 20110527
    'Set R = oDCred.RecuperaSeguroDesgravamenConsol(pdFechaProc, pnTipCamb)
    Set R = oDCred.RecuperaSeguroDesgravamenConsol(pdFechaProc, pnTipCamb, pdFechaProc01, psMoneda, psIndica, psTipIndica, psTipCartera)
    Set oDCred = Nothing
          
    'MADM 20110330
    If Not (R.EOF Or R.BOF) Then
        Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Seguro de Desgravamen", " Al " & pdFechaProc & ", Tipo de Cambio: " & CStr(pnTipCamb), lsNombreArchivo, lMatCabecera, R, 2, , , True)
    Else
        MsgBox "No se encuentran Registros con los Datos Ingresados, Verifíque!!", vbInformation, "Atención"
    End If
    'END MADM
End Sub


Public Sub MostrarReporteMoraDiarioxAnalista(ByVal pMatAgencias As Variant, ByVal psFecIni As String, _
ByVal psMoneda As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal pMatAnalistas As Variant, _
ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, ByVal pnMontoIni As Double, ByVal pnMontoFin As Double, _
Optional psubiGeoCod As String = "", _
Optional ByVal psNomCmac As String)
Dim oCredDoc As clases.DCredDoc
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsmensaje As String

    lsNombreArchivo = "moraDiariaxAnalista"

    ReDim lMatCabecera(30, 2)

    lMatCabecera(0, 0) = "Credito"
    lMatCabecera(1, 0) = "Analista"
    lMatCabecera(2, 0) = "Cliente"
    lMatCabecera(3, 0) = "Telefono"
    lMatCabecera(4, 0) = "Garante"
    lMatCabecera(5, 0) = "TelefonoGarante"
    lMatCabecera(6, 0) = "Prestamo"
    lMatCabecera(7, 0) = "Cuotas Aprobadas"
    lMatCabecera(8, 0) = "Cuota Vencida"
    lMatCabecera(9, 0) = "Dias Atraso"
    lMatCabecera(10, 0) = "Fecha Vencimiento"
    lMatCabecera(11, 0) = "DiasPromPago"
    lMatCabecera(12, 0) = "AtrazMax"
    lMatCabecera(13, 0) = "AtrazMin"
    lMatCabecera(14, 0) = "Capital"
    lMatCabecera(15, 0) = "InteresCuota"
    lMatCabecera(16, 0) = "Mora"
    lMatCabecera(17, 0) = "GastosAdmin"
    lMatCabecera(18, 0) = "MontoAPagar"
    lMatCabecera(19, 0) = "SaldoCapital"
    lMatCabecera(20, 0) = "TasaMora"
    lMatCabecera(21, 0) = "SaldoGastMora"
    lMatCabecera(22, 0) = "DireccionTrabajoTitular"
    lMatCabecera(23, 0) = "DireccionGarante"
    lMatCabecera(24, 0) = "FechaVigencia"
    lMatCabecera(25, 0) = "DireccionCasaTitular"
    lMatCabecera(26, 0) = "S.Cuota"
    lMatCabecera(27, 0) = "Cuotas Vencidas"
    lMatCabecera(28, 0) = "CFR"
    lMatCabecera(29, 0) = "F. Fallec"
    
            Dim sAgencias As String
            'Dim sCodUser As String
            Dim sProductos As String
            Dim sAnalistas As String
            Dim sCondicion As String
            Dim i As Integer
            Set oCredDoc = New DCredDoc
            For i = 0 To UBound(pMatAnalistas) - 1
                If i = 0 Then
                    sAnalistas = "" & pMatAnalistas(i) & ""
                Else
                    sAnalistas = sAnalistas & "," & pMatAnalistas(i) & ""
                End If
            Next i
            For i = 0 To UBound(pMatAgencias) - 1
                If i = 0 Then
                    sAgencias = "" & pMatAgencias(i) & ""
                Else
                    sAgencias = sAgencias & "," & pMatAgencias(i) & ""
                End If
            Next i
            
            For i = 0 To UBound(pMatProd) - 1
                If i = 0 Then
                    sProductos = "" & pMatProd(i) & ""
                Else
                    sProductos = sProductos & "," & pMatProd(i) & ""
                End If
            Next i
            
            For i = 0 To UBound(pMatCond) - 1
                If i = 0 Then
                    sCondicion = "" & pMatCond(i) & ""
                Else
                    sCondicion = sCondicion & "," & pMatCond(i) & ""
                End If
            Next i
    

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.ListaMoraXAnalistaCabecera(sAnalistas, sProductos, sAgencias, sCondicion, psMoneda, psFecIni, pnDiasIni, pnDiasFin, pnMontoIni, pnMontoFin, Right(psubiGeoCod, 15))
    Set oCredDoc = Nothing

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Seguro de Desgravamen", " Al " & psFecIni, lsNombreArchivo, lMatCabecera, R, 2, , , True)
    'R = Nothing
End Sub

Public Sub Reporte_CanceladosxAmpliacion(ByVal psFecIni As Date, ByVal psFecFin As Date)
Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsmensaje As String

    lsNombreArchivo = "reporteCreditosCanXAmpli"

    ReDim lMatCabecera(20, 2)

    lMatCabecera(0, 0) = "Cod.Cliente"
    lMatCabecera(1, 0) = "Cliente"
    lMatCabecera(2, 0) = "Ncredito"
    lMatCabecera(3, 0) = "Tipo"
    lMatCabecera(4, 0) = "Moneda"
    lMatCabecera(5, 0) = "Condicion"
    lMatCabecera(6, 0) = "D.Atraso"
    lMatCabecera(7, 0) = "Cuotas Aprobadas"
    lMatCabecera(8, 0) = "Cuotas Canceladas"
    lMatCabecera(9, 0) = "Calificacion"
    lMatCabecera(10, 0) = "CapitalCancelado"
    lMatCabecera(11, 0) = "Analista"
    lMatCabecera(12, 0) = "Agencia"
    lMatCabecera(13, 0) = "NroCreNu"
    lMatCabecera(14, 0) = "Tipo"
    lMatCabecera(15, 0) = "Moneda"
    lMatCabecera(16, 0) = "MontoDesem"
    lMatCabecera(17, 0) = "Cuotas Aprobadas"
    lMatCabecera(18, 0) = "Analista"
    lMatCabecera(19, 0) = "Agencia"
        
    Set oCredDoc = New DCOMCredDoc
    Set R = oCredDoc.CanceladosxAmpliacion(psFecIni, psFecFin)
    Set oCredDoc = Nothing

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Créditos Cancelados x Ampliación", " Al " & psFecIni, lsNombreArchivo, lMatCabecera, R, 2, , , True)
    'R = Nothing
End Sub
Sub GenCanceladosxAmpliacion()
        If chkMigraExcell.value = 1 Then
                Call Reporte_CanceladosxAmpliacion(CDate(Me.TxtFecIniA02), CDate(Me.TxtFecFinA02))
        End If
    
End Sub
Public Sub Reporte_PreCancelaciones(ByVal psFecIni As Date, ByVal psFecFin As Date)
Dim oCredDoc As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsmensaje As String

    lsNombreArchivo = "reporteCreditosPrecancelados"

    ReDim lMatCabecera(15, 2)

    lMatCabecera(0, 0) = "Cod.Cliente"
    lMatCabecera(1, 0) = "Cliente"
    lMatCabecera(2, 0) = "Ncredito"
    lMatCabecera(3, 0) = "Tipo"
    lMatCabecera(4, 0) = "Moneda"
    lMatCabecera(5, 0) = "Condicion"
    lMatCabecera(6, 0) = "D.Atraso"
    lMatCabecera(7, 0) = "Cuotas Aprobadas"
    lMatCabecera(8, 0) = "Cuotas Canceladas"
    lMatCabecera(9, 0) = "Calificacion"
    lMatCabecera(10, 0) = "CapitalCancelado"
    lMatCabecera(11, 0) = "Analista"
    lMatCabecera(12, 0) = "Agencia"
'    lMatCabecera(13, 0) = "NroCreNu"
'    lMatCabecera(14, 0) = "Tipo"
'    lMatCabecera(15, 0) = "Moneda"
'    lMatCabecera(16, 0) = "MontoDesem"
'    lMatCabecera(17, 0) = "Cuotas Aprobadas"
'    lMatCabecera(18, 0) = "Analista"
'    lMatCabecera(19, 0) = "Agencia"
    'MAVM 20100726 ***
    lMatCabecera(13, 0) = "F_Desembolso"
    lMatCabecera(14, 0) = "F_Cancelacion"
    lMatCabecera(15, 0) = "Tipo_Pago"
    '***
    Set oCredDoc = New DCOMCredDoc
    Set R = oCredDoc.PreCancelaciones(psFecIni, psFecFin)
    Set oCredDoc = Nothing

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Creditos Pre-Cancelados", " Al " & psFecIni, lsNombreArchivo, lMatCabecera, R, 2, , , True)
    'R = Nothing
End Sub
Sub GenPreCancelaciones()
        If chkMigraExcell.value = 1 Then
                Call Reporte_PreCancelaciones(CDate(Me.TxtFecIniA02), CDate(Me.TxtFecFinA02))
        End If
    
End Sub
Public Sub Reporte_ContaFideicomiso(ByVal psFecIni As Date)
Dim oCont As COMDContabilidad.DCOMCtaCont
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsmensaje As String

    lsNombreArchivo = "reporteContaFideicomiso"

    ReDim lMatCabecera(5, 2)

    lMatCabecera(0, 0) = "AgeDescripcion"
    lMatCabecera(1, 0) = "Moratorio"
    lMatCabecera(2, 0) = "Compensatorio "
    lMatCabecera(3, 0) = "InteresGracia"
    lMatCabecera(4, 0) = "Subtotal"
    
        
    Set oCont = New COMDContabilidad.DCOMCtaCont
    Set R = oCont.ContaFideicomiso(psFecIni, CDbl(TxtTipCambio.Text))
    Set oCont = Nothing

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Intereses Compensatorios de Fideicomiso", " Al " & Format(psFecIni, "MM/YYYY"), lsNombreArchivo, lMatCabecera, R, 2, , , True)
    'R = Nothing
End Sub
Sub GenContaFideicomiso()
        Call Reporte_ContaFideicomiso(CDate(Me.TxtFecIniA02))
End Sub
'ALPA*20080825***********************************************************
Private Sub ReporteGarantiasAdjudicadasContabilidad(ByVal pdFecha As Date, ByVal psAgencia As Variant)
 
    Dim fs As Scripting.FileSystemObject
 
'    Dim xlsAplicacion As Excel.Application
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
'    Dim lbLibroOpen As Boolean
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsNomHoja3  As String
    Dim lsNomHoja4  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsTipoGarantia As Integer
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lnContadorMatrix As Integer
    Dim lnPosY1 As Integer
    Dim lnPosY2 As Integer
    Dim lnPosYInicial As Integer
    Dim lsArchivo As String
    Dim lnTipoSalto As Integer
    Dim oGaran As COMDCredito.DCOMGarantia
    Set oGaran = New COMDCredito.DCOMGarantia
    Dim rs As ADODB.Recordset
    Dim lnNumColumns As Integer
    Dim sMatrixTotalHoja1() As String
    
    Dim i As Integer
    Dim j As Integer
    
    Set rs = New ADODB.Recordset
    psAgencia = Replace(psAgencia, " ", "")
    Set rs = oGaran.ReporteGarantiasAdjudicadasContabilidad(pdFecha, psAgencia)
    Set oGaran = Nothing
    'cTipoAdjudicado cDesTipoAdjudicado                                                                                                       cDescrip                                                                                                                                                                                                                                                         nSaldo                                  nInteres                                nAnio       nMes        dFecha                  cAge cTipoBien cOrigen nDifMes     cAgeDescripcion
    ReDim Preserve sMatrixGarantiasAdjud(1 To 13, 1 To 1)
    '******************
    Call actualizarMatrixPosiciones
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "PlantillaAdjudicadoContabilida"
    lsNomHoja1 = "Hoja1"
    lsArchivo1 = "\spooler\ReporteAdjudicadosContabilidad_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja1 Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    '*
    xlHoja1.Cells(10, 25) = Format(pdFecha, "DD.MM.YYYY")
    xlHoja1.Cells(10, 33) = Format(pdFecha, "DD.MM.YYYY")
    lnContador = 23
    lnPosY1 = 24
    lnPosYInicial = lnPosY1
    lsNombreAgencia = ""
    lnTipoSalto = 0
    ReDim Preserve sMatrixTotalHoja1(1 To 17)
    For i = 1 To 17
        sMatrixTotalHoja1(i) = "A1"
    Next i
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                 
                 If lsNombreAgencia <> rs!cAgeDescripcion Then
                    If lsTipoGarantia <> CInt(rs!cTipoAdjudicado) And lsTipoGarantia <> 0 Then
                        If lnTipoSalto = 0 Then
                            lnPosY2 = lnContador - 1
                            
                            xlHoja1.Range(devCelda(7) & lnContador & " :" & devCelda(7) & lnContador).Formula = "=Sum(" & devCelda(7) & lnPosY1 & " :" & devCelda(7) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(8) & lnContador & " :" & devCelda(8) & lnContador).Formula = "=Sum(" & devCelda(8) & lnPosY1 & " :" & devCelda(8) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(9) & lnContador & " :" & devCelda(9) & lnContador).Formula = "=Sum(" & devCelda(9) & lnPosY1 & " :" & devCelda(9) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(11) & lnContador & " :" & devCelda(11) & lnContador).Formula = "=Sum(" & devCelda(11) & lnPosY1 & " :" & devCelda(11) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(12) & lnContador & " :" & devCelda(12) & lnContador).Formula = "=Sum(" & devCelda(12) & lnPosY1 & " :" & devCelda(12) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(14) & lnContador & " :" & devCelda(14) & lnContador).Formula = "=Sum(" & devCelda(14) & lnPosY1 & " :" & devCelda(14) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(17) & lnContador & " :" & devCelda(17) & lnContador).Formula = "=Sum(" & devCelda(17) & lnPosY1 & " :" & devCelda(17) & lnPosY2 & ")"
                            sMatrixTotalHoja1(7) = sMatrixTotalHoja1(7) & "+" & devCelda(7) & lnContador
                            sMatrixTotalHoja1(8) = sMatrixTotalHoja1(8) & "+" & devCelda(8) & lnContador
                            sMatrixTotalHoja1(9) = sMatrixTotalHoja1(9) & "+" & devCelda(9) & lnContador
                            sMatrixTotalHoja1(11) = sMatrixTotalHoja1(11) & "+" & devCelda(11) & lnContador
                            sMatrixTotalHoja1(12) = sMatrixTotalHoja1(12) & "+" & devCelda(12) & lnContador
                            sMatrixTotalHoja1(14) = sMatrixTotalHoja1(14) & "+" & devCelda(14) & lnContador
                            sMatrixTotalHoja1(17) = sMatrixTotalHoja1(17) & "+" & devCelda(17) & lnContador
                            
                            xlHoja1.Range("AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "= " & devCelda(11) & (lnContador)
                            xlHoja1.Range("AC" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AC" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & devCelda(12) & lnPosY1 & " :" & devCelda(12) & lnPosY2 & ")"
                            xlHoja1.Range("AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & sMatrixTotalHoja1(13) & ")"
                            xlHoja1.Range("AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ")"
                            xlHoja1.Range("AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & "-AF" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)
                            Call actualizarNotaContabilidad(CInt(lsCodAgencia), CInt(lsTipoGarantia), "AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1))
                            lnContador = lnContador + 4
                            lnPosY1 = lnContador + 1
                        End If
                    End If
                    If lnContador > 23 Then
                        xlHoja1.Range(devCelda(2) & (lnContador - 1) & " :" & devCelda(17) & (lnContador - 1)).BorderAround xlContinuous, xlThin
                        xlHoja1.Cells(lnContador - 1, 2) = "TOTAL  " & lsNombreAgencia
                        
                        xlHoja1.Range(devCelda(7) & lnContador - 1 & " :" & devCelda(7) & lnContador - 1).Formula = "=Sum(" & sMatrixTotalHoja1(7) & ")"
                        xlHoja1.Range(devCelda(8) & lnContador - 1 & " :" & devCelda(8) & lnContador - 1).Formula = "=Sum(" & devCelda(8) & lnPosYInicial & " :" & devCelda(8) & lnPosY2 & ")"
                        xlHoja1.Range(devCelda(9) & lnContador - 1 & " :" & devCelda(9) & lnContador - 1).Formula = "=Sum(" & devCelda(9) & lnPosYInicial & " :" & devCelda(9) & lnPosY2 & ")"
                        xlHoja1.Range(devCelda(11) & lnContador - 1 & " :" & devCelda(11) & lnContador - 1).Formula = "=Sum(" & devCelda(11) & lnPosYInicial & " :" & devCelda(11) & lnPosY2 & ")"
                        xlHoja1.Range(devCelda(12) & lnContador - 1 & " :" & devCelda(12) & lnContador - 1).Formula = "=Sum(" & devCelda(12) & lnPosYInicial & " :" & devCelda(12) & lnPosY2 & ")"
                        xlHoja1.Range(devCelda(14) & lnContador - 1 & " :" & devCelda(14) & lnContador - 1).Formula = "=Sum(" & devCelda(14) & lnPosYInicial & " :" & devCelda(14) & lnPosY2 & ")"
                        xlHoja1.Range(devCelda(17) & lnContador - 1 & " :" & devCelda(17) & lnContador - 1).Formula = "=Sum(" & devCelda(17) & lnPosYInicial & " :" & devCelda(17) & lnPosY2 & ")"
                       
                        sMatrixTotalHoja1(7) = ""
                        sMatrixTotalHoja1(8) = ""
                        sMatrixTotalHoja1(9) = ""
                        sMatrixTotalHoja1(11) = ""
                        sMatrixTotalHoja1(12) = ""
                        sMatrixTotalHoja1(14) = ""
                        sMatrixTotalHoja1(17) = ""
                        lnPosYInicial = lnContador
                        'lnContador = lnContador + 1*
                        xlHoja1.Cells(lnContador, 2) = rs!cAgeDescripcion
                        If lsTipoGarantia <> CInt(rs!cTipoAdjudicado) Then
                            lnContador = lnContador + 1
                            xlHoja1.Cells(lnContador, 2) = rs!cDesTipoAdjudicado
                        End If
                    Else
                        xlHoja1.Cells(lnContador, 2) = rs!cDesTipoAdjudicado
                    End If
                    lnContador = lnContador + 1
                    lnTipoSalto = 1
                    
                 Else
                    If lsTipoGarantia <> CInt(rs!cTipoAdjudicado) Then
                        If lnTipoSalto = 0 Then
                            lnPosY2 = lnContador - 1
                            xlHoja1.Range(devCelda(7) & lnContador & " :" & devCelda(7) & lnContador).Formula = "=Sum(" & devCelda(7) & lnPosY1 & " :" & devCelda(7) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(8) & lnContador & " :" & devCelda(8) & lnContador).Formula = "=Sum(" & devCelda(8) & lnPosY1 & " :" & devCelda(8) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(9) & lnContador & " :" & devCelda(9) & lnContador).Formula = "=Sum(" & devCelda(9) & lnPosY1 & " :" & devCelda(9) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(11) & lnContador & " :" & devCelda(11) & lnContador).Formula = "=Sum(" & devCelda(11) & lnPosY1 & " :" & devCelda(11) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(12) & lnContador & " :" & devCelda(12) & lnContador).Formula = "=Sum(" & devCelda(12) & lnPosY1 & " :" & devCelda(12) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(14) & lnContador & " :" & devCelda(14) & lnContador).Formula = "=Sum(" & devCelda(14) & lnPosY1 & " :" & devCelda(14) & lnPosY2 & ")"
                            xlHoja1.Range(devCelda(17) & lnContador & " :" & devCelda(17) & lnContador).Formula = "=Sum(" & devCelda(17) & lnPosY1 & " :" & devCelda(17) & lnPosY2 & ")"
                            sMatrixTotalHoja1(7) = sMatrixTotalHoja1(7) & "+" & devCelda(7) & lnContador
                            sMatrixTotalHoja1(8) = sMatrixTotalHoja1(8) & "+" & devCelda(8) & lnContador
                            sMatrixTotalHoja1(9) = sMatrixTotalHoja1(9) & "+" & devCelda(9) & lnContador
                            sMatrixTotalHoja1(11) = sMatrixTotalHoja1(11) & "+" & devCelda(11) & lnContador
                            sMatrixTotalHoja1(12) = sMatrixTotalHoja1(12) & "+" & devCelda(12) & lnContador
                            sMatrixTotalHoja1(14) = sMatrixTotalHoja1(14) & "+" & devCelda(14) & lnContador
                            sMatrixTotalHoja1(17) = sMatrixTotalHoja1(17) & "+" & devCelda(17) & lnContador
                            
                            xlHoja1.Range("AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & devCelda(11) & lnPosY1 & " :" & devCelda(11) & lnPosY2 & ")"
                            xlHoja1.Range("AC" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AC" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & devCelda(12) & lnPosY1 & " :" & devCelda(12) & lnPosY2 & ")"
                            xlHoja1.Range("AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & sMatrixTotalHoja1(13) & ")"
                            xlHoja1.Range("AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ")"
                            xlHoja1.Range("AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & "-AF" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)
                            Call actualizarNotaContabilidad(CInt(lsCodAgencia), CInt(lsTipoGarantia), "AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1))
                            lnContador = lnContador + 4
                            lnPosY1 = lnContador + 1
                        End If
                        xlHoja1.Cells(lnContador, 2) = rs!cDesTipoAdjudicado
                        lnContador = lnContador + 1
                    End If
                 End If
                 
                    lnTipoSalto = 0
                    xlHoja1.Cells(lnContador, 2) = rs!cdescrip
                    xlHoja1.Cells(lnContador, 3) = rs!cTipoBien
                    xlHoja1.Cells(lnContador, 4) = rs!cOrigen
                    xlHoja1.Cells(lnContador, 5) = Format(rs!dfecha, "YYYY/MM/DD")
                    xlHoja1.Cells(lnContador, 7) = Format(Round(rs!nSaldo, 2), "#####,###.00")
                    xlHoja1.Cells(lnContador, 8) = Format(Round(rs!nSaldo, 2), "#####,###.00")
                    xlHoja1.Cells(lnContador, 11) = Format(Round(rs!nSaldo / 5, 2), "#####,###.00")
                    If CInt(rs!cTipoAdjudicado) = 1 Then
                        If rs!nDifMes > 12 Then
                            xlHoja1.Cells(lnContador, 12) = Format(Round((rs!nSaldo - (rs!nSaldo / 5)) / 42 * (rs!nDifMes - 12), 2), "#####,###.00")
                        Else
                            xlHoja1.Cells(lnContador, 12) = 0#
                        End If
                    Else
                        xlHoja1.Cells(lnContador, 12) = Format(Round((rs!nSaldo - (rs!nSaldo / 5)) / 18 * rs!nDifMes, 2), "#####,###.00")
                    End If
                    xlHoja1.Cells(lnContador, 14) = xlHoja1.Cells(lnContador, 11) + xlHoja1.Cells(lnContador, 12)
                    xlHoja1.Cells(lnContador, 17) = xlHoja1.Cells(lnContador, 7) - xlHoja1.Cells(lnContador, 14)
                    ReDim Preserve sMatrixGarantiasAdjud(1 To 13, 1 To lnContador + 1)
                    '**Copiar rs a Matrix*********
                    sMatrixGarantiasAdjud(1, lnContadorMatrix) = rs!cTipoAdjudicado
                    sMatrixGarantiasAdjud(2, lnContadorMatrix) = rs!cDesTipoAdjudicado
                    sMatrixGarantiasAdjud(3, lnContadorMatrix) = rs!cdescrip
                    sMatrixGarantiasAdjud(4, lnContadorMatrix) = rs!nSaldo
                    sMatrixGarantiasAdjud(5, lnContadorMatrix) = rs!nInteres
                    sMatrixGarantiasAdjud(6, lnContadorMatrix) = rs!nAnio
                    sMatrixGarantiasAdjud(7, lnContadorMatrix) = rs!nMES
                    sMatrixGarantiasAdjud(8, lnContadorMatrix) = rs!dfecha
                    sMatrixGarantiasAdjud(9, lnContadorMatrix) = rs!cAge
                    sMatrixGarantiasAdjud(10, lnContadorMatrix) = rs!cTipoBien
                    sMatrixGarantiasAdjud(11, lnContadorMatrix) = rs!cOrigen
                    sMatrixGarantiasAdjud(12, lnContadorMatrix) = rs!nDifMes
                    sMatrixGarantiasAdjud(13, lnContadorMatrix) = rs!cAgeDescripcion
                    '****************************
                    lnContadorMatrix = lnContadorMatrix + 1
                    lnContador = lnContador + 1
                    lsNombreAgencia = rs!cAgeDescripcion
                    lsTipoGarantia = CInt(rs!cTipoAdjudicado)
                    lsCodAgencia = rs!cAge
                    rs.MoveNext
            Wend
    End If
    rs.Close
    Set rs = Nothing
    If lnTipoSalto = 0 Then
        lnPosY2 = lnContador - 1
        xlHoja1.Range(devCelda(7) & lnContador & " :" & devCelda(7) & lnContador).Formula = "=Sum(" & devCelda(7) & lnPosY1 & " :" & devCelda(7) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(8) & lnContador & " :" & devCelda(8) & lnContador).Formula = "=Sum(" & devCelda(8) & lnPosY1 & " :" & devCelda(8) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(9) & lnContador & " :" & devCelda(9) & lnContador).Formula = "=Sum(" & devCelda(9) & lnPosY1 & " :" & devCelda(9) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(11) & lnContador & " :" & devCelda(11) & lnContador).Formula = "=Sum(" & devCelda(11) & lnPosY1 & " :" & devCelda(11) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(12) & lnContador & " :" & devCelda(12) & lnContador).Formula = "=Sum(" & devCelda(12) & lnPosY1 & " :" & devCelda(12) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(14) & lnContador & " :" & devCelda(14) & lnContador).Formula = "=Sum(" & devCelda(14) & lnPosY1 & " :" & devCelda(14) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(17) & lnContador & " :" & devCelda(17) & lnContador).Formula = "=Sum(" & devCelda(17) & lnPosY1 & " :" & devCelda(17) & lnPosY2 & ")"
        xlHoja1.Range(devCelda(2) & (lnContador + 1) & " :" & devCelda(17) & (lnContador + 1)).BorderAround xlContinuous, xlThin
        sMatrixTotalHoja1(7) = sMatrixTotalHoja1(7) & "+" & devCelda(7) & lnContador
        sMatrixTotalHoja1(8) = sMatrixTotalHoja1(8) & "+" & devCelda(8) & lnContador
        sMatrixTotalHoja1(9) = sMatrixTotalHoja1(9) & "+" & devCelda(9) & lnContador
        sMatrixTotalHoja1(11) = sMatrixTotalHoja1(11) & "+" & devCelda(11) & lnContador
        sMatrixTotalHoja1(12) = sMatrixTotalHoja1(12) & "+" & devCelda(12) & lnContador
        sMatrixTotalHoja1(14) = sMatrixTotalHoja1(14) & "+" & devCelda(14) & lnContador
        sMatrixTotalHoja1(17) = sMatrixTotalHoja1(17) & "+" & devCelda(17) & lnContador
        xlHoja1.Cells(lnContador + 1, 2) = "TOTAL  " & lsNombreAgencia
        xlHoja1.Range(devCelda(7) & lnContador + 1 & " :" & devCelda(7) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(7) & ")"
        xlHoja1.Range(devCelda(8) & lnContador + 1 & " :" & devCelda(8) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(8) & ")"
        xlHoja1.Range(devCelda(9) & lnContador + 1 & " :" & devCelda(9) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(9) & ")"
        xlHoja1.Range(devCelda(11) & lnContador + 1 & " :" & devCelda(11) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(11) & ")"
        xlHoja1.Range(devCelda(12) & lnContador + 1 & " :" & devCelda(12) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(12) & ")"
        xlHoja1.Range(devCelda(14) & lnContador + 1 & " :" & devCelda(14) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(14) & ")"
        xlHoja1.Range(devCelda(17) & lnContador + 1 & " :" & devCelda(17) & lnContador + 1).Formula = "=Sum(" & sMatrixTotalHoja1(17) & ")"
        
        '**
        xlHoja1.Range("AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & devCelda(11) & lnPosY1 & " :" & devCelda(11) & lnPosY2 & ")"
        xlHoja1.Range("AC" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AC" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & devCelda(12) & lnPosY1 & " :" & devCelda(12) & lnPosY2 & ")"
        xlHoja1.Range("AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(" & sMatrixTotalHoja1(13) & ")"
        xlHoja1.Range("AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=Sum(AB" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AD" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ")"
        xlHoja1.Range("AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & ":AG" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)).Formula = "=AE" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1) & "-AF" & sMatrixPosiciones(CInt(lsCodAgencia), CInt(lsTipoGarantia), 1)
        '**
    End If
    'sMatrixGarantiasAdjud(13, lnContadorMatrix) = rs!cAgeDescripcion
    If lnContadorMatrix > 0 Then
    lsNomHoja1 = "Hoja2"
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja1 Then
            xlHoja1.Activate
            lbExisteHoja = True
        Exit For
       End If
    Next
    lsNombreAgencia = "xxx"
    If lbExisteHoja = False Then
            Set xlHoja1 = xlsLibro.Worksheets
            xlHoja1.Name = lsNomHoja1
    End If
    lnContador = 8
    ReDim Preserve sMatrixTotalHoja1(1 To 27)
    xlHoja1.Cells(8, 8) = xlHoja1.Cells(8, 8) & "0" & (CInt(Format(pdFecha, "YY")) - 1)
    For i = 9 To 21
        xlHoja1.Cells(8, i) = xlHoja1.Cells(8, i) & Format(pdFecha, "YY")
    Next i
    
    lnPosY1 = 11
    For i = 1 To lnContadorMatrix - 1
    If sMatrixGarantiasAdjud(13, i) <> lsNombreAgencia Then
        lnPosY2 = lnContador - 1
        If lnContador > 8 Then
            For j = 4 To 27
               xlHoja1.Range(devCelda(j) & lnContador & " :" & devCelda(j) & lnContador).Formula = "= sum(" & devCelda(j) & lnPosY1 & ":" & devCelda(j) & lnPosY2 & ")"
            Next j
            lnContador = lnContador + 1
            xlHoja1.Cells(lnContador, 2) = "Total " & lsNombreAgencia
            For j = 4 To 27
               xlHoja1.Range(devCelda(j) & lnContador & " :" & devCelda(j) & lnContador).Formula = "= " & sMatrixTotalHoja1(j)
            Next j
            lnContador = lnContador + 1
        End If
        lnContador = lnContador + 1
        xlHoja1.Cells(lnContador, 2) = sMatrixGarantiasAdjud(13, i)
        lnContador = lnContador + 1
        xlHoja1.Cells(lnContador, 2) = sMatrixGarantiasAdjud(2, i)
        lnContador = lnContador + 1
        lnPosY1 = lnContador
        For j = 4 To 27
            sMatrixTotalHoja1(j) = ""
        Next j
    Else
        If sMatrixGarantiasAdjud(1, i) <> lsTipoGarantia Then
            lnPosY2 = lnContador - 1
            For j = 4 To 27
                xlHoja1.Range(devCelda(j) & lnContador & " :" & devCelda(j) & lnContador).Formula = "= sum(" & devCelda(j) & lnPosY1 & ":" & devCelda(j) & lnPosY2 & ")"
            Next j
            lnContador = lnContador + 1
            xlHoja1.Cells(lnContador, 2) = sMatrixGarantiasAdjud(2, i)
            lnContador = lnContador + 1
            lnPosY1 = lnContador
        End If
    End If
        xlHoja1.Cells(lnContador, 2) = sMatrixGarantiasAdjud(3, i)
        xlHoja1.Cells(lnContador, 3) = Format(sMatrixGarantiasAdjud(8, i), "YYYY/MM/DD")
        xlHoja1.Cells(lnContador, 4) = Format(sMatrixGarantiasAdjud(4, i), "###,###,###.00")
        xlHoja1.Cells(lnContador, 5) = Format(sMatrixGarantiasAdjud(5, i), "###,###,###.00")
        xlHoja1.Cells(lnContador, 6) = 0#
        xlHoja1.Range(devCelda(7) & lnContador & " :" & devCelda(7) & lnContador).Formula = "=ROUND((" & "Sum(" & (devCelda(4) & lnContador) & ":" & (devCelda(6) & lnContador) & ")),2)"
        xlHoja1.Cells(lnContador, 8) = 0#
        xlHoja1.Range(devCelda(9) & lnContador & " :" & devCelda(9) & lnContador).Formula = "=" & (devCelda(4) & lnContador) & "* 0.2"
        
        If sMatrixGarantiasAdjud(1, i) <> "1" Then
            For j = CInt(Format(sMatrixGarantiasAdjud(8, i), "mm")) + 1 To CInt(Format(pdFecha, "mm"))
                xlHoja1.Range(devCelda(j + 9) & lnContador & " :" & devCelda(j + 9) & lnContador).Formula = "=ROUND(((" & (devCelda(7) & lnContador) & "-" & (devCelda(9) & lnContador) & ")/18*1),2)"
            Next j
        ElseIf sMatrixGarantiasAdjud(1, i) = "1" Then
                For j = 1 To 12
                    If sMatrixGarantiasAdjud(12, i) >= 12 Then
                        If sMatrixGarantiasAdjud(12, i) <= 24 Then
                            If j > CInt(Format(sMatrixGarantiasAdjud(8, i), "mm")) And j <= CInt(Format(pdFecha, "mm")) Then
                                xlHoja1.Range(devCelda(j + 9) & lnContador & ":" & devCelda(j + 9) & lnContador).Formula = "=ROUND(((" & (devCelda(7) & lnContador) & "-" & (devCelda(9) & lnContador) & ")/42*(" & sMatrixGarantiasAdjud(12, i) & "-12)),2)"
                            End If
                        Else
                            xlHoja1.Range(devCelda(j + 9) & lnContador & " :" & devCelda(j + 9) & lnContador).Formula = "=ROUND(((" & (devCelda(7) & lnContador) & "-" & (devCelda(9) & lnContador) & ")/42*(" & sMatrixGarantiasAdjud(12, i) & "-12)),2)"
                        End If
                    Else
                        xlHoja1.Range(devCelda(j + 9) & lnContador & " :" & devCelda(j + 9) & lnContador).Formula = "=0"
                    End If
                Next j
        End If
        xlHoja1.Range(devCelda(22) & lnContador & " :" & devCelda(22) & lnContador).Formula = "=ROUND((" & "Sum(" & (devCelda(10) & lnContador) & ":" & (devCelda(21) & lnContador) & ")+ " & devCelda(8) & lnContador & " ),2)"
        xlHoja1.Range(devCelda(24) & lnContador & " :" & devCelda(24) & lnContador).Formula = "= " & devCelda(22) & lnContador & "+" & devCelda(9) & lnContador
        xlHoja1.Range(devCelda(27) & lnContador & " :" & devCelda(27) & lnContador).Formula = "= " & devCelda(7) & lnContador & "-" & devCelda(24) & lnContador
        For j = 4 To 27
            sMatrixTotalHoja1(j) = sMatrixTotalHoja1(j) & "+" & devCelda(j) & lnContador
        Next j
        lnContador = lnContador + 1
        lsNombreAgencia = sMatrixGarantiasAdjud(13, i)
        lsTipoGarantia = sMatrixGarantiasAdjud(1, i)
    Next i
        For j = 4 To 27
           xlHoja1.Range(devCelda(j) & lnContador & " :" & devCelda(j) & lnContador).Formula = "= sum(" & devCelda(j) & lnPosY1 & ":" & devCelda(j) & (lnContador - 1) & ")"
        Next j
        lnContador = lnContador + 1
        xlHoja1.Cells(lnContador, 2) = "Total " & lsNombreAgencia
        For j = 4 To 27
             xlHoja1.Range(devCelda(j) & lnContador & " :" & devCelda(j) & lnContador).Formula = "= " & sMatrixTotalHoja1(j)
        Next j
    End If
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub
'****************************************/***************/**********************
'ALPA 20081013******************************************************************
Private Function FColCredRepDesemEfect() As String
   Dim sCadImp As String
   Dim oNCredDoc As NCredDoc
   Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            Else
                sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            End If
        End If
    Set oNCredDoc = Nothing
    FColCredRepDesemEfect = sCadImp
End Function
'ALPA 20081013******************************************************************
Private Function FColCredRepSalCarVig() As String
Dim sCadImp As String
Dim oNCredDoc As NCredDoc
Set oNCredDoc = New NCredDoc
    If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
        sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
        sCadImp = sCadImp & Chr$(12)
        sCadImp = sCadImp & oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
    Else
        If ChkMonA02(0).value = 1 Then
            sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
        Else
            sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
        End If
    End If
Set oNCredDoc = Nothing
FColCredRepSalCarVig = sCadImp
End Function
'ALPA 20081013******************************************************************
Private Function FColCredRepCredCancel() As String
Dim sCadImp As String
Dim oNCredDoc As NCredDoc
Set oNCredDoc = New NCredDoc
    'HabilitaControleFrame1 True, True, True, True
        If OptSaldo(0).value Then
            If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
                sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, , matAnalista)
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, , matAnalista)
            Else
                If ChkMonA02(0).value = 1 Then
                    sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, , matAnalista)
                Else
                    sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, , matAnalista)
                End If
            End If
        Else
            If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
                sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            Else
                If ChkMonA02(0).value = 1 Then
                    sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                Else
                    sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
                End If
            End If
        End If
        Set oNCredDoc = Nothing
        FColCredRepCredCancel = sCadImp
End Function
'ALPA 20081013***************************************************************
Private Function FColCredRepResSalCarxAna() As String
Dim sCadImp As String
        Dim vCondi As String
        Dim sCondicion As String
        Dim oCOMNCredDoc As COMNCredito.NCOMCredDoc
        
        sCondicion = ObtieneCondi(vCondi)
        Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
        sCadImp = sCadImp & oCOMNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, sCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), gsNomCmac, val(TxtTipCambio.Text), lsTitProductos, IIf(chkCondBN.value = 1, True, False))

        Set oCOMNCredDoc = Nothing
FColCredRepResSalCarxAna = sCadImp
End Function
'ALPA 20081013***************************************************************
Private Function FColCredRepMoraInst() As String
Dim sCadImp As String
Dim oNCredDoc As NCredDoc
Set oNCredDoc = New NCredDoc

        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar2I.Text), CInt(TxtCar2F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
            End If
        End If
 Set oNCredDoc = Nothing
FColCredRepMoraInst = sCadImp
End Function
'ALPA 20081013***************************************************************
Private Function FColCredRepAtraPagoCuotaLib() As String
Dim sCadImp As String
Dim oNCredDoc As NCredDoc
Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True), gsNomCmac)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True), gsNomCmac)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True), gsNomCmac)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True), gsNomCmac)
            End If
        End If
 Set oNCredDoc = Nothing
FColCredRepAtraPagoCuotaLib = sCadImp
End Function
'******************************************************************************
Private Sub actualizarMatrixPosiciones()
    Dim i As Integer
    Dim j As Integer
    Dim nPosic As Integer
    ReDim Preserve sMatrixPosiciones(1 To 25, 0 To 5, 1 To 2)
    'Posicion de Cuadro 1
    For i = 1 To 25
        For j = 0 To 5
                If j = 0 Then
                    If i <= 7 Then
                        nPosic = 11 + 6 * i
                    Else
                        If i = 8 Then
                            nPosic = 59
                        ElseIf i = 10 Then
                            nPosic = 65
                        ElseIf i = 12 Then
                            nPosic = 71
                        ElseIf i = 13 Then
                            nPosic = 77
                        ElseIf i = 24 Then
                            nPosic = 83
                        ElseIf i = 25 Then
                            nPosic = 89
                        End If
                    End If
                End If
                sMatrixPosiciones(i, j, 1) = nPosic
            nPosic = nPosic + 1
        Next j
    Next i
    'Posicion de Cuadro 2
     nPosic = 15
    For j = 0 To 5
        For i = 1 To 25
            If i <= 7 Then
                nPosic = 14 + i + j * 14
            ElseIf i = 9 Then
                nPosic = nPosic = 22 + j * 14
            ElseIf i = 10 Then
                nPosic = nPosic = 23 + j * 14
            ElseIf i = 12 Then
                nPosic = nPosic = 24 + j * 14
            ElseIf i = 13 Then
                nPosic = nPosic = 25 + j * 14
            ElseIf i = 24 Then
                nPosic = nPosic = 26 + j * 14
            ElseIf i = 25 Then
                nPosic = nPosic = 27 + j * 14
            End If
            sMatrixPosiciones(i, j, 2) = nPosic
        Next i
    Next j
End Sub
Private Sub actualizarNotaContabilidad(ByVal nAgencia As Integer, ByVal nGarantia As Integer, ByVal sCelda As String)
    xlHoja1.Range("y" & sMatrixPosiciones(nAgencia, nGarantia, 2) & ":y" & sMatrixPosiciones(nAgencia, nGarantia, 2)).Formula = "=" & sCelda
    xlHoja1.Range("x" & sMatrixPosiciones(nAgencia, nGarantia, 2) + 43 & ":x" & sMatrixPosiciones(nAgencia, nGarantia, 2) + 43).Formula = "=" & sCelda
End Sub

'*** PEAC 20080923
Private Sub ImprimeRepClientesPotencialesSinCredVig(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pMatAgencias As Variant, _
    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim RFec As ADODB.Recordset
Dim sCadImp As String, i As Integer, j As Integer
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String

Dim lcCuenta As Integer
Dim lccodpers As String

'If pnDiaFin = 0 Or pnUit <= 0 Or pnTipCam <= 0 Then
'    MsgBox "Ingrese datos correctos para proseguir.", vbInformation, "Atención"
'    Exit Sub
'End If

'If pnDiaIni > pnDiaFin Then
'    MsgBox "Ingrese los Rangos correctamente.", vbInformation, "Atención"
'    Exit Sub
'End If

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    'Set R = oDCred.RecuperaDatosCredVencPaseCastigo(pnDiaIni, pnDiaFin, pnUit, pnTipCam, pMatAgencias)
    Set R = oDCred.RecuperaClientesPotencialesSinCredVig(pdFecIni, pdFecFin, pMatAgencias)
    Set RFec = oDCred.ObtieneUltimaFechaRCC()
    
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = psNomCmac
    ApExcel.Cells(3, 2).Formula = psNomAge
    ApExcel.Cells(2, 14).Formula = Date + Time()
    ApExcel.Cells(3, 14).Formula = psCodUser
    ApExcel.Range("B2", "N6").Font.Bold = True
    
    ApExcel.Cells(4, 2).Formula = "CLIENTES POTENCIALES SIN CREDITOS VIGENTES"
    ApExcel.Cells(5, 2).Formula = "Del " & Format(pdFecIni, "dd/MM/YYYY") & " al " & Format(pdFecFin, "dd/MM/YYYY")
    ApExcel.Cells(6, 2).Formula = "100 % Calificación Normal, del RCC al " & Format(RFec!Ult_Fec_RCC, "dd/MM/YYYY")
    
    ApExcel.Range("B7", "B8").MergeCells = True
    ApExcel.Range("C7", "C8").MergeCells = True
    ApExcel.Range("D7", "D8").MergeCells = True
    ApExcel.Range("E7", "E8").MergeCells = True
    ApExcel.Range("F7", "F8").MergeCells = True
    ApExcel.Range("G7", "G8").MergeCells = True
    ApExcel.Range("H7", "H8").MergeCells = True
    
    ApExcel.Range("I7", "K7").MergeCells = True
    ApExcel.Range("L7", "N7").MergeCells = True
    ApExcel.Range("B4", "N4").MergeCells = True
    ApExcel.Range("B5", "N5").MergeCells = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("N2", "N3").HorizontalAlignment = xlRight
    ApExcel.Range("B4", "N5").HorizontalAlignment = xlCenter
    ApExcel.Range("B7", "N8").VerticalAlignment = xlCenter
    ApExcel.Range("B7", "N8").Borders.LineStyle = 1
    
'    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & ImpreFormat(pnTipCam, 5, 3)
'    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(7, 2).Formula = "Agencia"
    ApExcel.Cells(7, 3).Formula = "Cod. Cliente"
    ApExcel.Cells(7, 4).Formula = "Nombre Cliente"
    ApExcel.Cells(7, 5).Formula = "DNI"
    ApExcel.Cells(7, 6).Formula = "Dirección Cliente"
    ApExcel.Cells(7, 7).Formula = "Direc. Ult. Fte. Ingreso"
    ApExcel.Cells(7, 8).Formula = "Analista"
    ApExcel.Cells(7, 9).Formula = "Último Crédito" '**
    ApExcel.Cells(8, 9).Formula = "Fecha Cancelación"
    ApExcel.Cells(8, 10).Formula = "Monto Crédito"
    ApExcel.Cells(8, 11).Formula = "Moneda"
    ApExcel.Cells(7, 12).Formula = "Entidades Financieras" '**
    ApExcel.Cells(8, 12).Formula = "Nombre Entidad"
    ApExcel.Cells(8, 13).Formula = "Saldo Capital"
    ApExcel.Cells(8, 14).Formula = "Moneda"
    
    ApExcel.Range("B7", "N8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "N8").Font.Bold = True
    ApExcel.Range("B7", "N8").HorizontalAlignment = 3

    i = 8
    lcCuenta = 0
    lccodpers = "X"
    Do While Not R.EOF
    i = i + 1
            
        If R!cPersCod <> lccodpers Then
            lcCuenta = lcCuenta + 1
        End If
    
        ApExcel.Cells(i, 2).Formula = R!Agencia
        ApExcel.Cells(i, 3).Formula = "'" & R!cPersCod
        ApExcel.Cells(i, 4).Formula = R!Cliente
        ApExcel.Cells(i, 5).Formula = "'" & R!Dni
        ApExcel.Cells(i, 6).Formula = R!Domicilio
        ApExcel.Cells(i, 7).Formula = R!DireccionFte
        ApExcel.Cells(i, 8).Formula = R!Analista
        ApExcel.Cells(i, 9).Formula = "'" & Format(R!Ult_Fec_Cancel, "dd/mm/yyyy")
        ApExcel.Cells(i, 10).Formula = R!nMontoCol
        ApExcel.Cells(i, 11).Formula = R!mone_ult_cancel
        ApExcel.Cells(i, 12).Formula = R!Entidad
        ApExcel.Cells(i, 13).Formula = R!Saldo_entidad
        ApExcel.Cells(i, 14).Formula = R!Moneda_entidad
        
        ApExcel.Range("J" & Trim(str(i)) & ":" & "N" & Trim(str(i))).NumberFormat = "#,##0.00"
        ApExcel.Range("B" & Trim(str(i)) & ":" & "N" & Trim(str(i))).Borders.LineStyle = 1
        
        lccodpers = R!cPersCod
        
        R.MoveNext
        If R.EOF Then
            Exit Do
        End If
        
    Loop
    
    i = i + 2
    ApExcel.Cells(i, 2).Formula = "Número Total de Clientes: " & Trim(str(lcCuenta))
    ApExcel.Cells(i, 2).Font.Bold = True
    
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

'*** PEAC 20080924
Private Sub ImprimeRepNumYSaldoCredPorProductoConsol(ByVal pnTipCam As Double, _
    ByVal psCodAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "")

Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim RFec As ADODB.Recordset
Dim RRes As ADODB.Recordset

Dim sCadImp As String, i As Integer, j As Integer
Dim lsNombreArchivo As String, lcTipProd As String
Dim lMatCabecera(1, 1) As Variant, vmone As String, vage As String
Dim lcValor As Integer

Dim lnNumCredMN As Double, lnNumCredME As Double, lnSaldoMN As Double, lnSaldoME As Double, lnSaldoTot As Double
Dim lnSINCFNumCredMN As Double, lnSINCFNumCredME As Double, lnSINCFSaldoMN As Double, lnSINCFSaldoME As Double, lnSINCFSaldoTot As Double

Dim TOTCONCFNumCredMN As Double, TOTCONCFNumCredME As Double, TOTCONCFSaldoMN As Double, TOTCONCFSaldoME As Double, TOTCONCFSaldoTot As Double
Dim TOTSINCFNumCredMN As Double, TOTSINCFNumCredME As Double, TOTSINCFSaldoMN As Double, TOTSINCFSaldoME As Double, TOTSINCFSaldoTot As Double

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaNumYSaldoCredPorProductoConsol(pnTipCam)
    Set RFec = oDCred.ObtieneUltimaFechaRCC()
    Set RRes = oDCred.ObtieneResCarteraPorProdSinCF(pnTipCam)
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

'**************************************************************************************
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = psNomCmac
    ApExcel.Cells(3, 2).Formula = psNomAge
    ApExcel.Cells(2, 8).Formula = Date + Time()
    ApExcel.Cells(3, 8).Formula = psCodUser
    ApExcel.Range("B2", "H8").Font.Bold = True
    
    ApExcel.Range("B2", "B3").HorizontalAlignment = xlLeft
    ApExcel.Range("B8", "B8").HorizontalAlignment = xlLeft
    ApExcel.Range("H2", "H3").HorizontalAlignment = xlRight
    ApExcel.Range("B5", "H6").HorizontalAlignment = xlCenter
    ApExcel.Range("B9", "H10").VerticalAlignment = xlCenter
    ApExcel.Range("B9", "H10").Borders.LineStyle = 1
    
    ApExcel.Cells(5, 2).Formula = "SALDOS Y NUMEROS DE CREDITOS POR PRODUCTO EN MN Y ME"
    ApExcel.Cells(6, 2).Formula = "Información al " & Format(RFec!Ult_Fec_RCC, "dd/MM/YYYY")
    ApExcel.Cells(8, 2).Formula = "Tipo de Cambio : " & Trim(str(pnTipCam))
    
    ApExcel.Range("B5", "H5").MergeCells = True
    ApExcel.Range("B6", "H6").MergeCells = True

    ApExcel.Range("B9", "C10").MergeCells = True
    ApExcel.Range("D9", "E9").MergeCells = True
    ApExcel.Range("F9", "G9").MergeCells = True
    
    ApExcel.Cells(9, 2).Formula = "Producto"
    ApExcel.Cells(9, 4).Formula = "Nro. CREDITOS"
    ApExcel.Cells(9, 6).Formula = "SALDOS"
    ApExcel.Cells(9, 8).Formula = "SALDO"
    ApExcel.Cells(10, 4).Formula = "M.N."
    ApExcel.Cells(10, 5).Formula = "M.E."
    ApExcel.Cells(10, 6).Formula = "M.N."
    ApExcel.Cells(10, 7).Formula = "M.E."
    ApExcel.Cells(10, 8).Formula = "TOTAL"
    
    ApExcel.Range("B9", "H10").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "H10").Font.Bold = True
    ApExcel.Range("B9", "H10").HorizontalAlignment = 3

    i = 10
    
    TOTCONCFNumCredMN = 0: TOTCONCFNumCredME = 0: TOTCONCFSaldoMN = 0: TOTCONCFSaldoME = 0: TOTCONCFSaldoTot = 0
    TOTSINCFNumCredMN = 0: TOTSINCFNumCredME = 0: TOTSINCFSaldoMN = 0: TOTSINCFSaldoME = 0: TOTSINCFSaldoTot = 0
    
    Do While Not R.EOF
    i = i + 1
            
        lcTipProd = R!TipProd
        lcValor = R!nConsValor
        
        ApExcel.Cells(i, 2).Formula = R!Producto
        ApExcel.Cells(i, 2).Font.Bold = True
                   
        lnNumCredMN = 0: lnNumCredME = 0: lnSaldoMN = 0: lnSaldoME = 0: lnSaldoTot = 0

        lnSINCFNumCredMN = 0: lnSINCFNumCredME = 0: lnSINCFSaldoMN = 0: lnSINCFSaldoME = 0: lnSINCFSaldoTot = 0

        Do While R!TipProd = lcTipProd

            If Mid(R!nConsValor, 2, 2) <> "00" Then
                i = i + 1
                
                ApExcel.Cells(i, 3).Formula = R!Producto
                ApExcel.Cells(i, 4).Formula = R!numcredMN
                ApExcel.Cells(i, 5).Formula = R!numcredME
                ApExcel.Cells(i, 6).Formula = R!saldoMN
                ApExcel.Cells(i, 7).Formula = R!saldoME
                ApExcel.Cells(i, 8).Formula = R!saldotot
                
                ApExcel.Range("D" & Trim(str(i)) & ":" & "H" & Trim(str(i))).NumberFormat = "#,##0.00"
                ApExcel.Range("B" & Trim(str(i)) & ":" & "H" & Trim(str(i))).Borders.LineStyle = 1
                
                lnNumCredMN = lnNumCredMN + R!numcredMN
                lnNumCredME = lnNumCredME + R!numcredME
                lnSaldoMN = lnSaldoMN + R!saldoMN
                lnSaldoME = lnSaldoME + R!saldoME
                lnSaldoTot = lnSaldoTot + R!saldotot
                
                lnSINCFNumCredMN = lnSINCFNumCredMN + IIf(R!nConsValor = "121" Or R!nConsValor = "221", 0, R!numcredMN)
                lnSINCFNumCredME = lnSINCFNumCredME + IIf(R!nConsValor = "121" Or R!nConsValor = "221", 0, R!numcredME)
                lnSINCFSaldoMN = lnSINCFSaldoMN + IIf(R!nConsValor = "121" Or R!nConsValor = "221", 0, R!saldoMN)
                lnSINCFSaldoME = lnSINCFSaldoME + IIf(R!nConsValor = "121" Or R!nConsValor = "221", 0, R!saldoME)
                lnSINCFSaldoTot = lnSINCFSaldoTot + IIf(R!nConsValor = "121" Or R!nConsValor = "221", 0, R!saldotot)
                
            End If
            R.MoveNext
            
            If R.EOF Then
                Exit Do
            End If
        Loop
        
        i = i + 1
        
        ApExcel.Cells(i, 3).Formula = "TOTAL"
        ApExcel.Cells(i, 4).Formula = lnNumCredMN
        ApExcel.Cells(i, 5).Formula = lnNumCredME
        ApExcel.Cells(i, 6).Formula = lnSaldoMN
        ApExcel.Cells(i, 7).Formula = lnSaldoME
        ApExcel.Cells(i, 8).Formula = lnSaldoTot
    
        TOTCONCFNumCredMN = TOTCONCFNumCredMN + lnNumCredMN
        TOTCONCFNumCredME = TOTCONCFNumCredME + lnNumCredME
        TOTCONCFSaldoMN = TOTCONCFSaldoMN + lnSaldoMN
        TOTCONCFSaldoME = TOTCONCFSaldoME + lnSaldoME
        TOTCONCFSaldoTot = TOTCONCFSaldoTot + lnSaldoTot
    
        TOTSINCFNumCredMN = TOTSINCFNumCredMN + lnSINCFNumCredMN
        TOTSINCFNumCredME = TOTSINCFNumCredME + lnSINCFNumCredME
        TOTSINCFSaldoMN = TOTSINCFSaldoMN + lnSINCFSaldoMN
        TOTSINCFSaldoME = TOTSINCFSaldoME + lnSINCFSaldoME
        TOTSINCFSaldoTot = TOTSINCFSaldoTot + lnSINCFSaldoTot

        ApExcel.Cells(i, 3).Font.Bold = True
        ApExcel.Range("D" & Trim(str(i)) & ":" & "H" & Trim(str(i))).Font.Bold = True
    Loop
    
        i = i + 2
        
        ApExcel.Cells(i, 3).Formula = "TOTAL CARTERA"
        ApExcel.Cells(i + 1, 3).Formula = "INCLUYE CARTAS FIANZAS"
        ApExcel.Cells(i + 2, 3).Formula = "SIN CARTAS FIANZAS"
        
        ApExcel.Cells(i + 1, 4).Formula = TOTCONCFNumCredMN
        ApExcel.Cells(i + 1, 5).Formula = TOTCONCFNumCredME
        ApExcel.Cells(i + 1, 6).Formula = TOTCONCFSaldoMN
        ApExcel.Cells(i + 1, 7).Formula = TOTCONCFSaldoME
        ApExcel.Cells(i + 1, 8).Formula = TOTCONCFSaldoTot
    
        ApExcel.Cells(i + 2, 4).Formula = TOTSINCFNumCredMN
        ApExcel.Cells(i + 2, 5).Formula = TOTSINCFNumCredME
        ApExcel.Cells(i + 2, 6).Formula = TOTSINCFSaldoMN
        ApExcel.Cells(i + 2, 7).Formula = TOTSINCFSaldoME
        ApExcel.Cells(i + 2, 8).Formula = TOTSINCFSaldoTot

        ApExcel.Range("C" & Trim(str(i)) & ":" & "H" & Trim(str(i + 2))).Borders.LineStyle = 1
        ApExcel.Range("C" & Trim(str(i)) & ":" & "H" & Trim(str(i + 2))).Font.Bold = True
    i = i + 4
    
'    ApExcel.Cells(I, 2).Formula = "Número Total de Clientes: " & Trim(Str(lcCuenta))
'    ApExcel.Cells(I, 2).Font.Bold = True
    
'**** resumen por prodcuto

    lnSaldoTot = 0

    Do While Not RRes.EOF
    i = i + 1
                            
        ApExcel.Cells(i, 2).Formula = RRes!Producto
        ApExcel.Cells(i, 3).Formula = RRes!saldotot
        
        ApExcel.Range("C" & Trim(str(i)) & ":" & "C" & Trim(str(i))).NumberFormat = "#,##0.00"
        ApExcel.Range("B" & Trim(str(i)) & ":" & "C" & Trim(str(i))).Borders.LineStyle = 1
        
        lnSaldoTot = lnSaldoTot + RRes!saldotot
                                
        RRes.MoveNext
        
        If RRes.EOF Then
            Exit Do
        End If
    Loop
        
        i = i + 1
        
        ApExcel.Cells(i, 2).Formula = "TOTAL (SIN CARTAS FIANZAS)"
        ApExcel.Cells(i, 3).Formula = lnSaldoTot
        
        ApExcel.Cells(i, 2).Font.Bold = True
        ApExcel.Cells(i, 3).Font.Bold = True
    
    R.Close
    RRes.Close
    Set R = Nothing
    Set RRes = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    Screen.MousePointer = 0
    
    ApExcel.Visible = True
    Set ApExcel = Nothing
    
End Sub

Private Sub GeneraArchExcelRepRiesgosCredi(ByVal psTipCamb As String, ByVal pdFecSis As Date)
'************************************************************'
'** GITU 20080926 108337 Segun Memo Nº 1705-2008-GM-DI/CMAC *'
Dim oDCred As COMDCredito.DCOMCredDoc
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lcCuenta As Integer
Dim lccodpers As String

    Screen.MousePointer = 11
    Set oDCred = New COMDCredito.DCOMCredDoc
    Set R = oDCred.RecuperaCreditosRiesgosCrediticios(psTipCamb, pdFecSis)
    
    lsNombreArchivo = "RepRiegosCrediticios"
    
    Set oDCred = Nothing
    
    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If
    '**** BRGO 16/11/2010
    Call GeneraReporte108337(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "REPORTE DE RIESGOS CREDITICIOS OFICINAS COMPARTIDAS", "", lsNombreArchivo, R, "", True)
    '*******************************
    Screen.MousePointer = 0
    
End Sub

'MAVM 20100511 ***
Public Function ReporteCredResultadoAnalistaBPPR_Excel(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Integer, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, _
    Optional ByVal psNomCmac As String = "", _
    Optional ByVal pnTipCam As Double) As String
    
    Dim objCOMDCredito As COMDCredito.DCOMCreditos
    Set objCOMDCredito = New COMDCredito.DCOMCreditos
    Dim objCOMD_Rent As COMDCredito.DCOMCredDoc
    Set objCOMD_Rent = New COMDCredito.DCOMCredDoc
    Dim objCOMDCreditoBPPR As COMDCredito.DCOMBPPR
    Set objCOMDCreditoBPPR = New COMDCredito.DCOMBPPR
    
    Dim rsResumen As ADODB.Recordset
    Dim rsResumenXCierre As ADODB.Recordset
    Dim rsConsolidado As ADODB.Recordset
    Dim rsRentabilidad As ADODB.Recordset
    Dim rsMeta As ADODB.Recordset
    
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim glsArchivo As String
    Dim liLineas As Integer, i, m As Integer
    Dim fs As Scripting.FileSystemObject
    Dim sCadAge, sCadProd, sAn As String
    Dim Casos As Integer
    
    If dcCartera.BoundText = "1" Then
        Set rsResumen = objCOMDCredito.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False)
        Set rsResumenXCierre = objCOMDCreditoBPPR.RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False, CDate(txtFCierreAnt.Text))
        Set rsConsolidado = objCOMDCredito.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
    
        sCadAge = ""
        For i = 0 To UBound(pMatAgencias) - 1
            sCadAge = sCadAge & pMatAgencias(i) & ","
        Next i
        sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
        
        sCadProd = ""
        If UBound(pMatProd) > 0 Then
            For i = 0 To UBound(pMatProd) - 1
                sCadProd = sCadProd & pMatProd(i) & ","
            Next i
            sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
        End If
        
        sAn = ObtenerAnalistasxAgencia(sCadAge)
        Set rsRentabilidad = objCOMD_Rent.RecuperaDatosConsolRentabilidadCarteraXAnalista(pdFecIni, pdFecFin, 8, sAn, sCadAge, sCadProd)

        If Not (rsResumen.EOF And rsResumen.BOF) Then
            glsArchivo = "ResultadoAnalista" & Format(pdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
            Set fs = New Scripting.FileSystemObject
            Set xlAplicacion = New Excel.Application
            If fs.FileExists(App.Path & "\SPOOLER\" & glsArchivo) Then
                Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & glsArchivo)
            Else
                Set xlLibro = xlAplicacion.Workbooks.Add
            End If

            Set xlHoja1 = xlLibro.Worksheets.Add
            lbExisteHoja = False
            lsNomHoja = "Resultado_Analista"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                Exit For
                End If
            Next
        
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 30
            xlAplicacion.Range("B1:B1").ColumnWidth = 10
            xlAplicacion.Range("C1:C1").ColumnWidth = 10
            xlAplicacion.Range("D1:D1").ColumnWidth = 10
            xlAplicacion.Range("E1:E1").ColumnWidth = 10
            xlAplicacion.Range("F1:F1").ColumnWidth = 10
            xlAplicacion.Range("G1:G1").ColumnWidth = 10
            xlAplicacion.Range("H1:H1").ColumnWidth = 10
            xlAplicacion.Range("I1:I1").ColumnWidth = 10
            xlAplicacion.Range("A1:Z2000").Font.Size = 8
            xlHoja1.Cells(1, 1) = "TIPO DE CAMBIO CIERRE"
            xlHoja1.Cells(1, 2) = pnTipCam
            xlHoja1.Cells(4, 1) = "RESULTADOS ANALISTAS" & " " & "MES Y CONSUMO"
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(4, 10)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 10)).Merge True
            xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
            xlHoja1.Cells(5, 1) = "ANALISTAS/CONCEPTOS"
            xlHoja1.Cells(6, 1) = "CALIFIC"
            xlHoja1.Cells(7, 1) = "META CREC. NUEVO"
            xlHoja1.Cells(8, 1) = "META CREC. MONTO"
            xlHoja1.Cells(9, 1) = "NUM. CLIENT. MES ANTER. O ARRASTRE"
            xlHoja1.Cells(10, 1) = "SALDO CARTER MES ANTER. O ARRASTRE"
            xlHoja1.Cells(11, 1) = "MORA > 30 MES ANTERIOR"
            xlHoja1.Cells(12, 1) = "MORA JUDICIAL MES ANTERIOR"
            xlHoja1.Cells(13, 1) = "TOTAL MORA"
            xlHoja1.Cells(14, 1) = "NUM. DESEMBOLSOS"
            xlHoja1.Cells(15, 1) = "NUM. CLIENT. CIERRE"
            xlHoja1.Cells(16, 1) = "NUM. CLIENT. NUEVOS"
            xlHoja1.Cells(17, 1) = "SALDO CARTERA CIERRE"
            xlHoja1.Cells(18, 1) = "CARTA FIANZA"
            xlHoja1.Cells(19, 1) = "MORA > A 30 CIERRE"
            xlHoja1.Cells(20, 1) = "MORA JUDICIAL CIERRE"
            xlHoja1.Cells(21, 1) = "SALDO CAPIT. UTILIZ. SOLES"
            xlHoja1.Cells(22, 1) = "SALDO CAPIT. UTILIZ. DOLARES"
            xlHoja1.Cells(23, 1) = "CAPITAL AMORTIZADO SOLES"
            xlHoja1.Cells(24, 1) = "CAPITAL AMORTIZADO DOLARES"
            xlHoja1.Cells(25, 1) = "N. OPERACIONES O CASOS"
            xlHoja1.Cells(26, 1) = "ING. FINANC. SOLES (INTERESES)"
            xlHoja1.Cells(27, 1) = "ING. FINANC. DOLARES (INTERESES)"
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Cells.Interior.Color = RGB(220, 220, 220)
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
            
            m = 2
            Do Until rsResumen.EOF
                If rsResumen!cUser <> "TOTAL" And rsResumen!nCar1Nro > 1 Then
                'Encabezado Analistas
                xlHoja1.Cells(5, m) = rsResumen!cUser
                xlHoja1.Cells(21, m) = Format(0, "#,##0.00")
                xlHoja1.Cells(22, m) = Format(0, "#,##0.00")
                xlHoja1.Cells(23, m) = Format(0, "#,##0.00")
                xlHoja1.Cells(24, m) = Format(0, "#,##0.00")
                xlHoja1.Cells(25, m) = Format(0, "#,##0.00")
                xlHoja1.Cells(26, m) = Format(0, "#,##0.00")
                xlHoja1.Cells(27, m) = Format(0, "#,##0.00")
                
                'Calif
                If rsResumen!nCar6Saldo <= 1000000 Then
                    xlHoja1.Cells(6, m) = 4
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 4)
                    If rsMeta.RecordCount <> 0 Then
                        'Meta Clie
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Meta Monto
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                        Set rsMeta = Nothing
                    End If
                End If
                If rsResumen!nCar6Saldo > 1000000 And rsResumen!nCar6Saldo <= 1500000 Then
                    xlHoja1.Cells(6, m) = 3
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 3)
                    If rsMeta.RecordCount <> 0 Then
                        'Meta Clie
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Meta Monto
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                        Set rsMeta = Nothing
                    End If
                End If
                If rsResumen!nCar6Saldo > 1500000 And rsResumen!nCar6Saldo <= 2000000 Then
                    xlHoja1.Cells(6, m) = 2
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 2)
                    If rsMeta.RecordCount <> 0 Then
                        'Meta Clie
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Meta Monto
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                        Set rsMeta = Nothing
                    End If
                End If
                If rsResumen!nCar6Saldo > 2000000 Then
                    xlHoja1.Cells(6, m) = 1
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 1)
                    If rsMeta.RecordCount <> 0 Then
                        'Meta Clie
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Meta Monto
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento)
                        Set rsMeta = Nothing
                    End If
                End If
                
                'Cierre Anterior ***
                Do Until rsResumenXCierre.EOF
                    If UCase(rsResumen!cUser) = UCase(rsResumenXCierre!cUser) Then
                        'Num Client Cierre_Ant
                        xlHoja1.Cells(9, m) = rsResumenXCierre!nNumCli6
                        'Saldo Carter Mes Anter Cierre_Ant
                        xlHoja1.Cells(10, m) = Format(rsResumenXCierre!nCar6Saldo, "#,##0.00")
                        'Mora >30 Cierre
                        xlHoja1.Cells(11, m) = Format(rsResumenXCierre!nCar4Saldo, "#,##0.00")
                        'Mora Jud Cierre
                        xlHoja1.Cells(12, m) = Format(val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                        'Total Mora
                        xlHoja1.Cells(13, m) = Format(val(rsResumenXCierre!nCar4Saldo) + val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                        Exit Do
                    End If
                    rsResumenXCierre.MoveNext
                Loop
                
                Do Until rsConsolidado.EOF
                    If UCase(rsResumen!cUser) = UCase(rsConsolidado!cUser) Then
                        'Num Desembolsos Cierre
                        xlHoja1.Cells(14, m) = val(rsConsolidado!nCantNuevo) + val(rsConsolidado!nCantRecurrente) + val(rsConsolidado!nCantParalelo) + val(rsConsolidado!nCantRefinanciado) + val(rsConsolidado!nCantAmpliado) + val(rsConsolidado!nCantAutomatico)
                        'Num Client Nuev Cierre
                        xlHoja1.Cells(16, m) = rsConsolidado!nCantNuevo
                        Exit Do
                    End If
                    rsConsolidado.MoveNext
                Loop
                
                'Num Client Cierre Act
                xlHoja1.Cells(15, m) = rsResumen!nNumCli6
                'Saldo Cart Cierre Act
                xlHoja1.Cells(17, m) = Format(rsResumen!nCar6Saldo, "#,##0.00")
                'Mora >30 Cierre Act
                xlHoja1.Cells(19, m) = Format(rsResumen!nCar4Saldo, "#,##0.00")
                'Mora Jud Cierre Act
                xlHoja1.Cells(20, m) = Format(val(rsResumen!nCar5Saldo1) + val(rsResumen!nCar5Saldo2), "#,##0.00")
                
                'Rentabilidad de Cartera ***
                Do Until rsRentabilidad.EOF
                    Casos = 0
                    If UCase(rsResumen!cUser) = UCase(rsRentabilidad!cUser) Then
                        'Saldo Capit. Utilizado S/.
                        If rsRentabilidad!cMoney = 1 Then
                            xlHoja1.Cells(21, m) = rsRentabilidad!nSaldo
                            xlHoja1.Cells(23, m) = rsRentabilidad!Capital
                            xlHoja1.Cells(26, m) = rsRentabilidad!Interes
                        Else
                            xlHoja1.Cells(22, m) = rsRentabilidad!nSaldo * pnTipCam
                            xlHoja1.Cells(24, m) = rsRentabilidad!Capital * pnTipCam
                            xlHoja1.Cells(27, m) = rsRentabilidad!Interes * pnTipCam
                        End If
                        Casos = Casos + rsRentabilidad!Casos
                        xlHoja1.Cells(25, m) = Casos
                    End If
                    rsRentabilidad.MoveNext
                Loop
                'Rentabilidad de Cartera ***
                
                m = m + 1
                rsConsolidado.MoveFirst
                rsResumenXCierre.MoveFirst
                rsRentabilidad.MoveFirst
            End If
            rsResumen.MoveNext
        Loop
        
            xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsArchivo
            MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsArchivo
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
            Set xlAplicacion = Nothing
            Set xlLibro = Nothing
            Set xlHoja1 = Nothing
            Set objCOMDCredito = Nothing
            Set objCOMD_Rent = Nothing
            Set objCOMDCreditoBPPR = Nothing
            ReporteCredResultadoAnalistaBPPR_Excel = ""
        Else
            MsgBox "No existen datos para generar el reporte"
            ReporteCredResultadoAnalistaBPPR_Excel = ""
        End If
        
        'Comercial y Conv
        Else
            ReDim pMatProd(6)
            pMatProd(0) = "101"
            pMatProd(1) = "102"
            pMatProd(2) = "103"
            pMatProd(3) = "401"
            pMatProd(4) = "403"
            pMatProd(5) = "423"
            
            Set rsResumen = objCOMDCredito.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False)
            Set rsResumenXCierre = objCOMDCreditoBPPR.RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False, CDate(txtFCierreAnt.Text))
            Set rsConsolidado = objCOMDCredito.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
            sCadAge = ""
            For i = 0 To UBound(pMatAgencias) - 1
                sCadAge = sCadAge & pMatAgencias(i) & ","
            Next i
            sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
            sCadProd = ""
            If UBound(pMatProd) > 0 Then
                For i = 0 To UBound(pMatProd) - 1
                    sCadProd = sCadProd & pMatProd(i) & ","
                Next i
                sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
            End If
            sAn = ObtenerAnalistasxAgencia(sCadAge)
            Set rsRentabilidad = objCOMD_Rent.RecuperaDatosConsolRentabilidadCarteraXAnalista(pdFecIni, pdFecFin, 8, sAn, sCadAge, sCadProd)
            
            'If Not (rsResumen.EOF And rsResumen.BOF) Then
                glsArchivo = "ResultadoAnalista" & Format(pdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
                Set fs = New Scripting.FileSystemObject
                Set xlAplicacion = New Excel.Application
                If fs.FileExists(App.Path & "\SPOOLER\" & glsArchivo) Then
                    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & glsArchivo)
                Else
                    Set xlLibro = xlAplicacion.Workbooks.Add
                End If
    
                Set xlHoja1 = xlLibro.Worksheets.Add
                lbExisteHoja = False
                lsNomHoja = "Resultado_Analista"
                For Each xlHoja1 In xlLibro.Worksheets
                    If xlHoja1.Name = lsNomHoja Then
                        xlHoja1.Activate
                        lbExisteHoja = True
                    Exit For
                    End If
                Next
            
                If lbExisteHoja = False Then
                    Set xlHoja1 = xlLibro.Worksheets.Add
                    xlHoja1.Name = lsNomHoja
                End If
    
                xlAplicacion.Range("A1:A1").ColumnWidth = 30
                xlAplicacion.Range("B1:B1").ColumnWidth = 10
                xlAplicacion.Range("C1:C1").ColumnWidth = 10
                xlAplicacion.Range("D1:D1").ColumnWidth = 10
                xlAplicacion.Range("E1:E1").ColumnWidth = 10
                xlAplicacion.Range("F1:F1").ColumnWidth = 10
                xlAplicacion.Range("G1:G1").ColumnWidth = 10
                xlAplicacion.Range("H1:H1").ColumnWidth = 10
                xlAplicacion.Range("I1:I1").ColumnWidth = 10
                xlAplicacion.Range("A1:Z2000").Font.Size = 8
                xlHoja1.Cells(1, 1) = "TIPO DE CAMBIO CIERRE"
                xlHoja1.Cells(1, 2) = pnTipCam
                xlHoja1.Cells(3, 1) = "RESULTADOS ANALISTAS" & " " & "COMERCIAL Y CONVENIO"
                xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(3, 10)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 10)).Merge True
                xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
                xlHoja1.Range("A4:A5").MergeCells = True
                xlHoja1.Cells(4, 2) = "COMERCIAL"
                xlHoja1.Cells(4, 7) = "CONVENIO"
                xlHoja1.Range("B4:F4").MergeCells = True
                xlHoja1.Range("G4:J4").MergeCells = True
                xlHoja1.Cells(4, 1) = "ANALISTAS/CONCEPTOS"
                xlHoja1.Cells(6, 1) = "CALIFIC"
                xlHoja1.Cells(7, 1) = "META CREC. NUEVO"
                xlHoja1.Cells(8, 1) = "META CREC. MONTO"
                xlHoja1.Cells(9, 1) = "NUM. CLIENT. MES ANTER. O ARRASTRE"
                xlHoja1.Cells(10, 1) = "SALDO CARTER MES ANTER. O ARRASTRE"
                xlHoja1.Cells(11, 1) = "MORA > 30 MES ANTERIOR"
                xlHoja1.Cells(12, 1) = "MORA JUDICIAL MES ANTERIOR"
                xlHoja1.Cells(13, 1) = "TOTAL MORA"
                xlHoja1.Cells(14, 1) = "NUM. DESEMBOLSOS"
                xlHoja1.Cells(15, 1) = "NUM. CLIENT. CIERRE"
                xlHoja1.Cells(16, 1) = "NUM. CLIENT. NUEVOS"
                xlHoja1.Cells(17, 1) = "SALDO CARTERA CIERRE"
                xlHoja1.Cells(18, 1) = "CARTA FIANZA"
                xlHoja1.Cells(19, 1) = "MORA > A 30 CIERRE"
                xlHoja1.Cells(20, 1) = "MORA JUDICIAL CIERRE"
                xlHoja1.Cells(21, 1) = "SALDO CAPIT. UTILIZ. SOLES"
                xlHoja1.Cells(22, 1) = "SALDO CAPIT. UTILIZ. DOLARES"
                xlHoja1.Cells(23, 1) = "CAPITAL AMORTIZADO SOLES"
                xlHoja1.Cells(24, 1) = "CAPITAL AMORTIZADO DOLARES"
                xlHoja1.Cells(25, 1) = "N. OPERACIONES O CASOS"
                xlHoja1.Cells(26, 1) = "ING. FINANC. SOLES (INTERESES)"
                xlHoja1.Cells(27, 1) = "ING. FINANC. DOLARES (INTERESES)"
                xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Cells.Interior.Color = RGB(220, 220, 220)
                xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
                            
                m = 2
                Do Until rsResumen.EOF
                    If rsResumen!cUser <> "TOTAL" And rsResumen!nCar1Nro > 1 Then
                        'Encabezado Analistas
                        xlHoja1.Cells(5, m) = rsResumen!cUser
                        xlHoja1.Cells(21, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(22, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(23, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(24, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(25, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(26, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(27, m) = Format(0, "#,##0.00")
                        'Calif
                        If rsResumen!nCar6Saldo <= 5000000 Then
                            xlHoja1.Cells(6, m) = 2
                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 3, 2)
                            If rsMeta.RecordCount <> 0 Then
                                'Meta Clie
                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                                'Meta Monto
                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                                Set rsMeta = Nothing
                            End If
                        End If
                        If rsResumen!nCar6Saldo > 5000000 Then
                            xlHoja1.Cells(6, m) = 1
                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 3, 1)
                            If rsMeta.RecordCount <> 0 Then
                                'Meta Clie
                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                                'Meta Monto
                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento)
                                Set rsMeta = Nothing
                            End If
                        End If
                    
                        'Cierre Anterior ***
                        Do Until rsResumenXCierre.EOF
                            If UCase(rsResumen!cUser) = UCase(rsResumenXCierre!cUser) Then
                                'Num Client Cierre_Ant
                                xlHoja1.Cells(9, m) = rsResumenXCierre!nNumCli6
                                'Saldo Carter Mes Anter Cierre_Ant
                                xlHoja1.Cells(10, m) = Format(rsResumenXCierre!nCar6Saldo, "#,##0.00")
                                'Mora >30 Cierre
                                xlHoja1.Cells(11, m) = Format(rsResumenXCierre!nCar4Saldo, "#,##0.00")
                                'Mora Jud Cierre
                                xlHoja1.Cells(12, m) = Format(val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                                'Total Mora
                                xlHoja1.Cells(13, m) = Format(val(rsResumenXCierre!nCar4Saldo) + val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                                Exit Do
                            End If
                            rsResumenXCierre.MoveNext
                        Loop
                    
                        Do Until rsConsolidado.EOF
                            If UCase(rsResumen!cUser) = UCase(rsConsolidado!cUser) Then
                                'Num Desembolsos Cierre
                                xlHoja1.Cells(14, m) = val(rsConsolidado!nCantNuevo) + val(rsConsolidado!nCantRecurrente) + val(rsConsolidado!nCantParalelo) + val(rsConsolidado!nCantRefinanciado) + val(rsConsolidado!nCantAmpliado) + val(rsConsolidado!nCantAutomatico)
                                'Num Client Nuev Cierre
                                xlHoja1.Cells(16, m) = rsConsolidado!nCantNuevo
                                Exit Do
                            End If
                            rsConsolidado.MoveNext
                        Loop
                    
                        'Num Client Cierre Act
                        xlHoja1.Cells(15, m) = rsResumen!nNumCli6
                        'Saldo Cart Cierre Act
                        xlHoja1.Cells(17, m) = Format(rsResumen!nCar6Saldo, "#,##0.00")
                        'Mora >30 Cierre Act
                        xlHoja1.Cells(19, m) = Format(rsResumen!nCar4Saldo, "#,##0.00")
                        'Mora Jud Cierre Act
                        xlHoja1.Cells(20, m) = Format(val(rsResumen!nCar5Saldo1) + val(rsResumen!nCar5Saldo2), "#,##0.00")
                    
                        'Rentabilidad de Cartera ***
                        Do Until rsRentabilidad.EOF
                            Casos = 0
                            If UCase(rsResumen!cUser) = UCase(rsRentabilidad!cUser) Then
                                'Saldo Capit. Utilizado S/.
                                If rsRentabilidad!cMoney = 1 Then
                                    xlHoja1.Cells(21, m) = rsRentabilidad!nSaldo
                                    xlHoja1.Cells(23, m) = rsRentabilidad!Capital
                                    xlHoja1.Cells(26, m) = rsRentabilidad!Interes
                                Else
                                    xlHoja1.Cells(22, m) = rsRentabilidad!nSaldo * pnTipCam
                                    xlHoja1.Cells(24, m) = rsRentabilidad!Capital * pnTipCam
                                    xlHoja1.Cells(27, m) = rsRentabilidad!Interes * pnTipCam
                                End If
                                Casos = Casos + rsRentabilidad!Casos
                                xlHoja1.Cells(25, m) = Casos
                            End If
                            rsRentabilidad.MoveNext
                        Loop
                        'Rentabilidad de Cartera ***
                        m = m + 1
                        rsConsolidado.MoveFirst
                        rsResumenXCierre.MoveFirst
                        rsRentabilidad.MoveFirst
                        End If
                        rsResumen.MoveNext
                Loop
                
                ReDim pMatProd(2)
                pMatProd(0) = "301"
                pMatProd(1) = "320"
                
                Set rsResumen = objCOMDCredito.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False)
                Set rsResumenXCierre = objCOMDCreditoBPPR.RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False, CDate(txtFCierreAnt.Text))
                Set rsConsolidado = objCOMDCredito.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
                sCadAge = ""
                For i = 0 To UBound(pMatAgencias) - 1
                    sCadAge = sCadAge & pMatAgencias(i) & ","
                Next i
                sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
                sCadProd = ""
                If UBound(pMatProd) > 0 Then
                    For i = 0 To UBound(pMatProd) - 1
                        sCadProd = sCadProd & pMatProd(i) & ","
                    Next i
                    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
                End If
                sAn = ObtenerAnalistasxAgencia(sCadAge)
                Set rsRentabilidad = objCOMD_Rent.RecuperaDatosConsolRentabilidadCarteraXAnalista(pdFecIni, pdFecFin, 8, sAn, sCadAge, sCadProd)

                m = 7
                Do Until rsResumen.EOF
                    If rsResumen!cUser <> "TOTAL" And rsResumen!nCar1Nro > 1 Then
                        'Encabezado Analistas
                        xlHoja1.Cells(5, m) = rsResumen!cUser
                        xlHoja1.Cells(21, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(22, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(23, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(24, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(25, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(26, m) = Format(0, "#,##0.00")
                        xlHoja1.Cells(27, m) = Format(0, "#,##0.00")
                        'Calif
                        If rsResumen!nCar6Saldo <= 3000000 Then
                            xlHoja1.Cells(6, m) = 3
                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 2, 3)
                            If rsMeta.RecordCount <> 0 Then
                                'Meta Clie
                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                                'Meta Monto
                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                                Set rsMeta = Nothing
                            End If
                        End If
                        If rsResumen!nCar6Saldo > 3000000 And rsResumen!nCar6Saldo <= 5500000 Then
                            xlHoja1.Cells(6, m) = 2
                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 2, 2)
                            If rsMeta.RecordCount <> 0 Then
                                'Meta Clie
                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                                'Meta Monto
                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                                Set rsMeta = Nothing
                            End If
                        End If
                        If rsResumen!nCar6Saldo > 5500000 Then
                            xlHoja1.Cells(6, m) = 1
                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 2, 1)
                            If rsMeta.RecordCount <> 0 Then
                                'Meta Clie
                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                                'Meta Monto
                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento)
                                Set rsMeta = Nothing
                            End If
                        End If
                    
                        'Cierre Anterior ***
                        Do Until rsResumenXCierre.EOF
                            If UCase(rsResumen!cUser) = UCase(rsResumenXCierre!cUser) Then
                                'Num Client Cierre_Ant
                                xlHoja1.Cells(9, m) = rsResumenXCierre!nNumCli6
                                'Saldo Carter Mes Anter Cierre_Ant
                                xlHoja1.Cells(10, m) = Format(rsResumenXCierre!nCar6Saldo, "#,##0.00")
                                'Mora >30 Cierre
                                xlHoja1.Cells(11, m) = Format(rsResumenXCierre!nCar4Saldo, "#,##0.00")
                                'Mora Jud Cierre
                                xlHoja1.Cells(12, m) = Format(val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                                'Total Mora
                                xlHoja1.Cells(13, m) = Format(val(rsResumenXCierre!nCar4Saldo) + val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                                Exit Do
                            End If
                            rsResumenXCierre.MoveNext
                        Loop
                    
                        Do Until rsConsolidado.EOF
                            If UCase(rsResumen!cUser) = UCase(rsConsolidado!cUser) Then
                                'Num Desembolsos Cierre
                                xlHoja1.Cells(14, m) = val(rsConsolidado!nCantNuevo) + val(rsConsolidado!nCantRecurrente) + val(rsConsolidado!nCantParalelo) + val(rsConsolidado!nCantRefinanciado) + val(rsConsolidado!nCantAmpliado) + val(rsConsolidado!nCantAutomatico)
                                'Num Client Nuev Cierre
                                xlHoja1.Cells(16, m) = rsConsolidado!nCantNuevo
                                Exit Do
                            End If
                            rsConsolidado.MoveNext
                        Loop
                    
                        'Num Client Cierre Act
                        xlHoja1.Cells(15, m) = rsResumen!nNumCli6
                        'Saldo Cart Cierre Act
                        xlHoja1.Cells(17, m) = Format(rsResumen!nCar6Saldo, "#,##0.00")
                        'Mora >30 Cierre Act
                        xlHoja1.Cells(19, m) = Format(rsResumen!nCar4Saldo, "#,##0.00")
                        'Mora Jud Cierre Act
                        xlHoja1.Cells(20, m) = Format(val(rsResumen!nCar5Saldo1) + val(rsResumen!nCar5Saldo2), "#,##0.00")
                    
                        'Rentabilidad de Cartera ***
                        Do Until rsRentabilidad.EOF
                            Casos = 0
                            If UCase(rsResumen!cUser) = UCase(rsRentabilidad!cUser) Then
                                'Saldo Capit. Utilizado S/.
                                If rsRentabilidad!cMoney = 1 Then
                                    xlHoja1.Cells(21, m) = rsRentabilidad!nSaldo
                                    xlHoja1.Cells(23, m) = rsRentabilidad!Capital
                                    xlHoja1.Cells(26, m) = rsRentabilidad!Interes
                                Else
                                    xlHoja1.Cells(22, m) = rsRentabilidad!nSaldo * pnTipCam
                                    xlHoja1.Cells(24, m) = rsRentabilidad!Capital * pnTipCam
                                    xlHoja1.Cells(27, m) = rsRentabilidad!Interes * pnTipCam
                                End If
                                Casos = Casos + rsRentabilidad!Casos
                                xlHoja1.Cells(25, m) = Casos
                            End If
                            rsRentabilidad.MoveNext
                        Loop
                        'Rentabilidad de Cartera ***
                        m = m + 1
                        rsConsolidado.MoveFirst
                        rsResumenXCierre.MoveFirst
                        rsRentabilidad.MoveFirst
                        End If
                        rsResumen.MoveNext
                Loop
                
                xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsArchivo
                MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsArchivo
                xlAplicacion.Visible = True
                xlAplicacion.Windows(1).Visible = True
                Set xlAplicacion = Nothing
                Set xlLibro = Nothing
                Set xlHoja1 = Nothing
                Set objCOMDCredito = Nothing
                Set objCOMD_Rent = Nothing
                Set objCOMDCreditoBPPR = Nothing
                ReporteCredResultadoAnalistaBPPR_Excel = ""
            'Else
                'MsgBox "No existen datos para generar el reporte"
                'ReporteCredResultadoAnalistaBPPR_Excel = ""
        'End If
        
    End If
End Function

Private Sub CargarAgencias()
    Dim rsAgencia As New ADODB.Recordset
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    Set rsAgencia.DataSource = objCOMDCredito.CargarAgencias
    dcAgencia.BoundColumn = "cAgeCod"
    dcAgencia.DataField = "cAgeCod"
    Set dcAgencia.RowSource = rsAgencia
    dcAgencia.ListField = "cAgeDescripcion"
    dcAgencia.BoundText = gsCodAge
End Sub

Private Sub CargarCartera()
    Dim rsCartera As New ADODB.Recordset
    Dim objCOMDCredito As COMDCredito.DCOMBPPR
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    Set rsCartera.DataSource = objCOMDCredito.CargarTipoCarteraReporte
    dcCartera.BoundColumn = "IdCartera"
    dcCartera.DataField = "IdCartera"
    Set dcCartera.RowSource = rsCartera
    dcCartera.ListField = "Cartera"
    dcCartera.BoundText = 0
End Sub

Private Function ObtenerAnalistasxAgencia(ByVal sAgeCod As String) As String
    Dim rsCargo As New ADODB.Recordset
    Dim rsAnalista As New ADODB.Recordset
    Dim sCargo, sCadAnalista As String
    Dim obj As COMDCredito.DCOMBPPR
    Set obj = New COMDCredito.DCOMBPPR
    Set rsCargo.DataSource = obj.DarCargoAnalistas
    If rsCargo.RecordCount <> 0 Then
        sCargo = Trim(Replace(rsCargo!nConsSisValor, "'", ""))
        sCargo = Trim(Replace(sCargo, "'", ""))
    End If
    'Comentado By JACA 20110707
    'Set rsAnalista.DataSource = obj.CargarAnalistasXAgencia(sAgeCod, sCargo)
    
    Do Until rsAnalista.EOF
        sCadAnalista = sCadAnalista & rsAnalista!cPersCod & ","
        rsAnalista.MoveNext
    Loop
    sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 1)
    ObtenerAnalistasxAgencia = sCadAnalista
    
End Function
'***

'MAVM 20100520
Public Function ReporteCredBonificacionAnalistaBPPR_Excel(psServerCons As String, ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal pnMoneda As Integer, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatCond As Variant, _
    Optional ByVal psNomCmac As String = "", _
    Optional ByVal pnTipCam As Double) As String
    
    Dim objCOMDCredito As COMDCredito.DCOMCreditos
    Set objCOMDCredito = New COMDCredito.DCOMCreditos
    Dim objCOMD_Rent As COMDCredito.DCOMCredDoc
    Set objCOMD_Rent = New COMDCredito.DCOMCredDoc
    Dim objCOMDCreditoBPPR As COMDCredito.DCOMBPPR
    Set objCOMDCreditoBPPR = New COMDCredito.DCOMBPPR
    
    Dim rsResumen As ADODB.Recordset
    Dim rsResumenXCierre As ADODB.Recordset
    Dim rsConsolidado As ADODB.Recordset
    Dim rsRentabilidad As ADODB.Recordset
    Dim rsMeta As ADODB.Recordset
    
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim glsArchivo As String
    Dim liLineas As Integer, i, m As Integer
    Dim fs As Scripting.FileSystemObject
    Dim sCadAge, sCadProd, sAn As String
    Dim Casos As Integer
    
    If dcCartera.BoundText = "1" Then

        Set rsResumen = objCOMDCredito.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False)
        Set rsResumenXCierre = objCOMDCreditoBPPR.RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False, CDate(txtFCierreAnt.Text))
        Set rsConsolidado = objCOMDCredito.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
    
        sCadAge = ""
        For i = 0 To UBound(pMatAgencias) - 1
            sCadAge = sCadAge & pMatAgencias(i) & ","
        Next i
        sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

        sCadProd = ""
        If UBound(pMatProd) > 0 Then
            For i = 0 To UBound(pMatProd) - 1
                sCadProd = sCadProd & pMatProd(i) & ","
            Next i
            sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
        End If

        sAn = ObtenerAnalistasxAgencia(sCadAge)
        Set rsRentabilidad = objCOMD_Rent.RecuperaDatosConsolRentabilidadCarteraXAnalista(pdFecIni, pdFecFin, 8, sAn, sCadAge, sCadProd)

        If Not (rsResumen.EOF And rsResumen.BOF) Then
            glsArchivo = "BonificacionXAnalista" & Format(pdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
            Set fs = New Scripting.FileSystemObject
            Set xlAplicacion = New Excel.Application
            If fs.FileExists(App.Path & "\SPOOLER\" & glsArchivo) Then
                Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & glsArchivo)
            Else
                Set xlLibro = xlAplicacion.Workbooks.Add
            End If

            Set xlHoja1 = xlLibro.Worksheets.Add
            lbExisteHoja = False
            lsNomHoja = "Bonificacion_Analista"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                Exit For
                End If
            Next
        
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 30
            xlAplicacion.Range("B1:B1").ColumnWidth = 10
            xlAplicacion.Range("C1:C1").ColumnWidth = 10
            xlAplicacion.Range("D1:D1").ColumnWidth = 10
            xlAplicacion.Range("E1:E1").ColumnWidth = 10
            xlAplicacion.Range("F1:F1").ColumnWidth = 10
            xlAplicacion.Range("G1:G1").ColumnWidth = 10
            xlAplicacion.Range("H1:H1").ColumnWidth = 10
            xlAplicacion.Range("I1:I1").ColumnWidth = 10
            xlAplicacion.Range("A1:Z2000").Font.Size = 8
            xlHoja1.Cells(4, 1) = "BONIFICACION ANALISTAS" & " " & "MES Y CONSUMO"
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(4, 10)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 10)).Merge True
            xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
            xlHoja1.Cells(5, 1) = "CONCEPTOS"
            xlHoja1.Cells(6, 1) = "Nº creditos desembolsados"
            xlHoja1.Cells(7, 1) = "Nº crecimiento de creditos NUEVOS"
            xlHoja1.Cells(8, 1) = "Monto de crecimiento esperado"
            xlHoja1.Cells(9, 1) = "CRECIMIENTO DEL MES"
            xlHoja1.Cells(10, 1) = "CARTA FIANZA"
            xlHoja1.Cells(11, 1) = "Numero de credito Mes anterior"
            xlHoja1.Cells(12, 1) = "Numero de credito Mes actual"
            xlHoja1.Cells(13, 1) = "CRECIMIENTO Nº NUEVO"
            xlHoja1.Cells(14, 1) = "Saldo cartera mes anterior"
            xlHoja1.Cells(15, 1) = "Saldo cartera mes actual"
            xlHoja1.Cells(16, 1) = "Saldo cartera al dia (S/.)"
            xlHoja1.Cells(17, 1) = "Saldo cartera mora > 30 dias (S/.)"
            xlHoja1.Cells(18, 1) = "Saldo cartera mora > 30 dias (S/.) MES ANTERIOR"
            xlHoja1.Cells(19, 1) = "CAPITAL UTILIZADO"
            xlHoja1.Cells(20, 1) = "Mora > a 30 dias (%)"
            xlHoja1.Cells(21, 1) = "Mora mes anterior vencidos alto riesgo"
            xlHoja1.Cells(22, 1) = "BONO POR CRECIMIENTO EN MONTOS"
            xlHoja1.Cells(23, 1) = "BONO POR CLIENTES nuevos"
            xlHoja1.Cells(24, 1) = "BONO POR CUMPLIMIENTO DE MORA ANALISTA 1"
            xlHoja1.Cells(25, 1) = "BONO POR CUMPLIMIENTO DE MORA ANALISTA 2"
            xlHoja1.Cells(26, 1) = "BONO POR CUMPLIMIENTO DE MORA ANALISTA 3"
            xlHoja1.Cells(27, 1) = "BONO POR CUMPLIMIENTO DE MORA ANALISTA 4"
            xlHoja1.Cells(28, 1) = "PLUS POR CRECIMIENTO 1%"
            xlHoja1.Cells(29, 1) = "PLUS POR CRECIMIENTO NUM. CRED S/ 20"
            xlHoja1.Cells(30, 1) = "Sub Total de BPP (S/.)"
            xlHoja1.Cells(31, 1) = "Deduccion por mora 1"
            xlHoja1.Cells(32, 1) = "Deduccion por mora 2"
            xlHoja1.Cells(33, 1) = "Deduccion por mora 3"
            xlHoja1.Cells(34, 1) = "Deduccion por mora 4"
            xlHoja1.Cells(35, 1) = "Total BPP"
            xlHoja1.Cells(36, 1) = ""
            xlHoja1.Cells(37, 1) = "CARTA FIANZA"
            xlHoja1.Cells(38, 1) = "TOTAL BONIFICACION"
            
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Cells.Interior.Color = RGB(220, 220, 220)
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
            
            m = 2
            Do Until rsResumen.EOF
                If rsResumen!cUser <> "TOTAL" And rsResumen!nCar1Nro > 1 Then
                'Encabezado Analistas
                xlHoja1.Cells(5, m) = rsResumen!cUser
'                xlHoja1.Cells(21, m) = Format(0, "#,##0.00")
'                xlHoja1.Cells(22, m) = Format(0, "#,##0.00")
'                xlHoja1.Cells(23, m) = Format(0, "#,##0.00")
'                xlHoja1.Cells(24, m) = Format(0, "#,##0.00")
'                xlHoja1.Cells(25, m) = Format(0, "#,##0.00")
'                xlHoja1.Cells(26, m) = Format(0, "#,##0.00")
'                xlHoja1.Cells(27, m) = Format(0, "#,##0.00")
                
                Do Until rsConsolidado.EOF
                    If UCase(rsResumen!cUser) = UCase(rsConsolidado!cUser) Then
                        'N Creditos Desembolsados
                        xlHoja1.Cells(6, m) = val(rsConsolidado!nCantNuevo) + val(rsConsolidado!nCantRecurrente) + val(rsConsolidado!nCantParalelo) + val(rsConsolidado!nCantRefinanciado) + val(rsConsolidado!nCantAmpliado) + val(rsConsolidado!nCantAutomatico)
                        'Crecimiento Nº Nuevo
                        xlHoja1.Cells(13, m) = rsConsolidado!nCantNuevo
                        Exit Do
                    End If
                    rsConsolidado.MoveNext
                Loop
                
                Dim CalifAnalist As Integer
                Dim ParamAceptable As Double
                Dim fCMHastaConDescuento1 As Double
                Dim fCMHastaConDescuento2 As Double
                Dim fCMHastaConDescuento3 As Double
                Dim fCMHastaConDescuento4 As Double
                Dim fCMHastaConDescuento5 As Double
                Dim fCMHastaConDescuento6 As Double
                Dim fCMHastaConDescuento7 As Double
                
                Dim fCMHastaConDescuento8 As Double
                Dim fCMHastaConDescuento9 As Double
                Dim fCMHastaConDescuento10 As Double
                Dim fCMHastaConDescuento11 As Double
                
                Dim fCMDesConDescuento2 As Double
                Dim fCMDesConDescuento3 As Double
                Dim fCMDesConDescuento4 As Double
                Dim fCMDesConDescuento5 As Double
                Dim fCMDesConDescuento6 As Double
                Dim fCMDesConDescuento7 As Double
                
                Dim fCMDesConDescuento8 As Double
                Dim fCMDesConDescuento9 As Double
                Dim fCMDesConDescuento10 As Double
                Dim fCMDesConDescuento11 As Double
                Dim fRentabilidad As Double
                
                If rsResumen!nCar6Saldo <= 1000000 Then
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 4)
                    CalifAnalist = 4
                    If rsMeta.RecordCount <> 0 Then
                        'Crecimiento de Creditos Nuevos
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Monto Crecimiento esperado
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                        ParamAceptable = rsMeta!fCMHastaAceptacion
                        
                        fCMHastaConDescuento1 = rsMeta!fCMHastaConDescuento1
                        fCMHastaConDescuento2 = rsMeta!fCMHastaConDescuento2
                        fCMHastaConDescuento3 = rsMeta!fCMHastaConDescuento3
                        fCMHastaConDescuento4 = rsMeta!fCMHastaConDescuento4
                        fCMHastaConDescuento5 = rsMeta!fCMHastaConDescuento5
                        fCMHastaConDescuento6 = rsMeta!fCMHastaConDescuento6
                        fCMHastaConDescuento7 = rsMeta!fCMHastaConDescuento7
                        
                        fCMHastaConDescuento8 = rsMeta!fCMHastaConDescuento8
                        fCMHastaConDescuento9 = rsMeta!fCMHastaConDescuento9
                        fCMHastaConDescuento10 = rsMeta!fCMHastaConDescuento10
                        fCMHastaConDescuento11 = rsMeta!fCMHastaConDescuento11
                        
                        fCMDesConDescuento2 = rsMeta!fCMDescConDescuento2
                        fCMDesConDescuento3 = rsMeta!fCMDescConDescuento3
                        fCMDesConDescuento4 = rsMeta!fCMDescConDescuento4
                        fCMDesConDescuento5 = rsMeta!fCMDescConDescuento5
                        fCMDesConDescuento6 = rsMeta!fCMDescConDescuento6
                        fCMDesConDescuento7 = rsMeta!fCMDescConDescuento7
                        
                        fCMDesConDescuento8 = rsMeta!fCMDescConDescuento8
                        fCMDesConDescuento9 = rsMeta!fCMDescConDescuento9
                        fCMDesConDescuento10 = rsMeta!fCMDescConDescuento10
                        fCMDesConDescuento11 = rsMeta!fCMDescConDescuento11
                        
                        fRentabilidad = rsMeta!fRentabilidad
                        
                        Set rsMeta = Nothing
                    End If
                End If
                If rsResumen!nCar6Saldo > 1000000 And rsResumen!nCar6Saldo <= 1500000 Then
                    CalifAnalist = 3
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 3)
                    If rsMeta.RecordCount <> 0 Then
                        'Crecimiento de Creditos Nuevos
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Monto Crecimiento esperado
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                        ParamAceptable = rsMeta!fCMHastaAceptacion
                        
                        fCMHastaConDescuento1 = rsMeta!fCMHastaConDescuento1
                        fCMHastaConDescuento2 = rsMeta!fCMHastaConDescuento2
                        fCMHastaConDescuento3 = rsMeta!fCMHastaConDescuento3
                        fCMHastaConDescuento4 = rsMeta!fCMHastaConDescuento4
                        fCMHastaConDescuento5 = rsMeta!fCMHastaConDescuento5
                        fCMHastaConDescuento6 = rsMeta!fCMHastaConDescuento6
                        fCMHastaConDescuento7 = rsMeta!fCMHastaConDescuento7
                        
                        fCMHastaConDescuento8 = rsMeta!fCMHastaConDescuento8
                        fCMHastaConDescuento9 = rsMeta!fCMHastaConDescuento9
                        fCMHastaConDescuento10 = rsMeta!fCMHastaConDescuento10
                        fCMHastaConDescuento11 = rsMeta!fCMHastaConDescuento11
                        
                        fCMDesConDescuento2 = rsMeta!fCMDescConDescuento2
                        fCMDesConDescuento3 = rsMeta!fCMDescConDescuento3
                        fCMDesConDescuento4 = rsMeta!fCMDescConDescuento4
                        fCMDesConDescuento5 = rsMeta!fCMDescConDescuento5
                        fCMDesConDescuento6 = rsMeta!fCMDescConDescuento6
                        fCMDesConDescuento7 = rsMeta!fCMDescConDescuento7
                        
                        fCMDesConDescuento8 = rsMeta!fCMDescConDescuento8
                        fCMDesConDescuento9 = rsMeta!fCMDescConDescuento9
                        fCMDesConDescuento10 = rsMeta!fCMDescConDescuento10
                        fCMDesConDescuento11 = rsMeta!fCMDescConDescuento11
                        
                        Set rsMeta = Nothing
                    End If
                End If
                If rsResumen!nCar6Saldo > 1500000 And rsResumen!nCar6Saldo <= 2000000 Then
                    CalifAnalist = 2
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 2)
                    If rsMeta.RecordCount <> 0 Then
                        'Crecimiento de Creditos Nuevos
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Monto Crecimiento esperado
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
                        ParamAceptable = rsMeta!fCMHastaAceptacion
                        
                        fCMHastaConDescuento1 = rsMeta!fCMHastaConDescuento1
                        fCMHastaConDescuento2 = rsMeta!fCMHastaConDescuento2
                        fCMHastaConDescuento3 = rsMeta!fCMHastaConDescuento3
                        fCMHastaConDescuento4 = rsMeta!fCMHastaConDescuento4
                        fCMHastaConDescuento5 = rsMeta!fCMHastaConDescuento5
                        fCMHastaConDescuento6 = rsMeta!fCMHastaConDescuento6
                        fCMHastaConDescuento7 = rsMeta!fCMHastaConDescuento7
                        
                        fCMHastaConDescuento8 = rsMeta!fCMHastaConDescuento8
                        fCMHastaConDescuento9 = rsMeta!fCMHastaConDescuento9
                        fCMHastaConDescuento10 = rsMeta!fCMHastaConDescuento10
                        fCMHastaConDescuento11 = rsMeta!fCMHastaConDescuento11
                        
                        fCMDesConDescuento2 = rsMeta!fCMDescConDescuento2
                        fCMDesConDescuento3 = rsMeta!fCMDescConDescuento3
                        fCMDesConDescuento4 = rsMeta!fCMDescConDescuento4
                        fCMDesConDescuento5 = rsMeta!fCMDescConDescuento5
                        fCMDesConDescuento6 = rsMeta!fCMDescConDescuento6
                        fCMDesConDescuento7 = rsMeta!fCMDescConDescuento7
                        
                        fCMDesConDescuento8 = rsMeta!fCMDescConDescuento8
                        fCMDesConDescuento9 = rsMeta!fCMDescConDescuento9
                        fCMDesConDescuento10 = rsMeta!fCMDescConDescuento10
                        fCMDesConDescuento11 = rsMeta!fCMDescConDescuento11
                        
                        Set rsMeta = Nothing
                    End If
                End If
                If rsResumen!nCar6Saldo > 2000000 Then
                    CalifAnalist = 1
                    Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 1, 1)
                    If rsMeta.RecordCount <> 0 Then
                        'Crecimiento de Creditos Nuevos
                        xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
                        'Monto Crecimiento esperado
                        xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento)
                        ParamAceptable = rsMeta!fCMHastaAceptacion
                        
                        fCMHastaConDescuento1 = rsMeta!fCMHastaConDescuento1
                        fCMHastaConDescuento2 = rsMeta!fCMHastaConDescuento2
                        fCMHastaConDescuento3 = rsMeta!fCMHastaConDescuento3
                        fCMHastaConDescuento4 = rsMeta!fCMHastaConDescuento4
                        fCMHastaConDescuento5 = rsMeta!fCMHastaConDescuento5
                        fCMHastaConDescuento6 = rsMeta!fCMHastaConDescuento6
                        fCMHastaConDescuento7 = rsMeta!fCMHastaConDescuento7
                        
                        fCMHastaConDescuento8 = rsMeta!fCMHastaConDescuento8
                        fCMHastaConDescuento9 = rsMeta!fCMHastaConDescuento9
                        fCMHastaConDescuento10 = rsMeta!fCMHastaConDescuento10
                        fCMHastaConDescuento11 = rsMeta!fCMHastaConDescuento11
                        
                        fCMDesConDescuento2 = rsMeta!fCMDescConDescuento2
                        fCMDesConDescuento3 = rsMeta!fCMDescConDescuento3
                        fCMDesConDescuento4 = rsMeta!fCMDescConDescuento4
                        fCMDesConDescuento5 = rsMeta!fCMDescConDescuento5
                        fCMDesConDescuento6 = rsMeta!fCMDescConDescuento6
                        fCMDesConDescuento7 = rsMeta!fCMDescConDescuento7
                        
                        fCMDesConDescuento8 = rsMeta!fCMDescConDescuento8
                        fCMDesConDescuento9 = rsMeta!fCMDescConDescuento9
                        fCMDesConDescuento10 = rsMeta!fCMDescConDescuento10
                        fCMDesConDescuento11 = rsMeta!fCMDescConDescuento11
                        
                        Set rsMeta = Nothing
                    End If
                End If
                
                'Cierre Anterior ***
                Do Until rsResumenXCierre.EOF
                    If UCase(rsResumen!cUser) = UCase(rsResumenXCierre!cUser) Then
                        'Numero credito mes anterior
                        xlHoja1.Cells(11, m) = rsResumenXCierre!nNumCli6
                        'Saldo Carter Mes Anter
                        xlHoja1.Cells(14, m) = Format(rsResumenXCierre!nCar6Saldo, "#,##0.00")
                        'Saldo cartera mora >30 Cierre
                        xlHoja1.Cells(18, m) = Format(rsResumenXCierre!nCar4Saldo, "#,##0.00")
                        'Mora Jud Cierre
                        'xlHoja1.Cells(12, m) = Format(val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                        'Total Mora
                        'xlHoja1.Cells(13, m) = Format(val(rsResumenXCierre!nCar4Saldo) + val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
                        Exit Do
                    End If
                    rsResumenXCierre.MoveNext
                Loop
                
                Dim dCapUtil As Double
                Dim d26 As Double
                Dim d27 As Double
                
                
                'Num Client Cierre Act
                xlHoja1.Cells(12, m) = rsResumen!nNumCli6
                'Saldo Cart Cierre Act
                xlHoja1.Cells(15, m) = Format(rsResumen!nCar6Saldo, "#,##0.00")
                'Mora >30 Cierre Act
                xlHoja1.Cells(17, m) = Format(rsResumen!nCar4Saldo, "#,##0.00")
                'dCapUtil = rsResumen!nCar5Saldo1 + rsResumen!nCar5Saldo2
                'Mora Jud Cierre Act
                'xlHoja1.Cells(20, m) = Format(val(rsResumen!nCar5Saldo1) + val(rsResumen!nCar5Saldo2), "#,##0.00")
                
                'Rentabilidad de Cartera ***
                Do Until rsRentabilidad.EOF
                    Casos = 0
                    If UCase(rsResumen!cUser) = UCase(rsRentabilidad!cUser) Then
                        'Saldo Capit. Utilizado S/.
                        If rsRentabilidad!cMoney = 1 Then
                            'xlHoja1.Cells(21, m) = rsRentabilidad!nSaldo
                            dCapUtil = dCapUtil + rsRentabilidad!nSaldo
                            'xlHoja1.Cells(23, m) = rsRentabilidad!Capital
                            dCapUtil = dCapUtil + rsRentabilidad!Capital
                            d26 = rsRentabilidad!Interes
                            'xlHoja1.Cells(26, m) = rsRentabilidad!Interes
                        Else
                            'xlHoja1.Cells(22, m) = rsRentabilidad!nSaldo * pnTipCam
                            dCapUtil = dCapUtil + rsRentabilidad!nSaldo * pnTipCam
                            'xlHoja1.Cells(24, m) = rsRentabilidad!Capital * pnTipCam
                            dCapUtil = dCapUtil + rsRentabilidad!Capital * pnTipCam
                            d27 = rsRentabilidad!Interes * pnTipCam
                            'xlHoja1.Cells(27, m) = rsRentabilidad!Interes * pnTipCam
                        End If
                        Casos = Casos + rsRentabilidad!Casos
                        'xlHoja1.Cells(25, m) = Casos
                    End If
                    rsRentabilidad.MoveNext
                Loop
                'Rentabilidad de Cartera ***
                
                xlHoja1.Cells(9, m) = Format(xlHoja1.Cells(15, m) - xlHoja1.Cells(14, m), "#,##0.00")
                xlHoja1.Cells(16, m) = Format(xlHoja1.Cells(15, m) - xlHoja1.Cells(17, m), "#,##0.00")
                xlHoja1.Cells(19, m) = Format(dCapUtil, "#,##0.00")
                xlHoja1.Cells(20, m) = Format(Round((xlHoja1.Cells(17, m) / xlHoja1.Cells(15, m)) * 100, 2), "#,##0.00") & " %"
                xlHoja1.Cells(21, m) = Format(Round((xlHoja1.Cells(18, m) / xlHoja1.Cells(14, m)) * 100, 2), "#,##0.00") & " %"
                xlHoja1.Cells(22, m) = Format(IIf(xlHoja1.Cells(18, m) <= 0, 0, IIf(xlHoja1.Cells(9, m) >= xlHoja1.Cells(8, m), 300, 0)), "#,##0.00")
                xlHoja1.Cells(23, m) = Format(IIf(xlHoja1.Cells(15, m) <= 0, 0, IIf(xlHoja1.Cells(13, m) >= xlHoja1.Cells(7, m), 20, 0)), "#,##0.00")
                xlHoja1.Cells(24, m) = Format(IIf(CalifAnalist = 1, IIf(xlHoja1.Cells(15, m) <= 500000, 0, IIf(xlHoja1.Cells(20, m) <= ParamAceptable / 100, 250, 0)), 0), "#,##0.00")
                xlHoja1.Cells(25, m) = Format(IIf(CalifAnalist = 2, IIf(xlHoja1.Cells(15, m) <= 500000, 0, IIf(xlHoja1.Cells(20, m) <= ParamAceptable / 100, 250, 0)), 0), "#,##0.00")
                xlHoja1.Cells(26, m) = Format(IIf(CalifAnalist = 3, IIf(xlHoja1.Cells(15, m) <= 500000, 0, IIf(xlHoja1.Cells(20, m) <= ParamAceptable / 100, 250, 0)), 0), "#,##0.00")
                xlHoja1.Cells(27, m) = Format(IIf(CalifAnalist = 4, IIf(xlHoja1.Cells(15, m) <= 500000, 0, IIf(xlHoja1.Cells(20, m) <= ParamAceptable / 100, 250, 0)), 0), "#,##0.00")
                xlHoja1.Cells(28, m) = Format(IIf(xlHoja1.Cells(9, m) > xlHoja1.Cells(8, m), (xlHoja1.Cells(9, m) - xlHoja1.Cells(8, m)) * 0.01, 0), "#,##0.00")
                xlHoja1.Cells(29, m) = Format(IIf(xlHoja1.Cells(13, m) > xlHoja1.Cells(7, m), (xlHoja1.Cells(13, m) - xlHoja1.Cells(7, m)) * 20, 0), "#,##0.00")
                xlHoja1.Cells(30, m) = Format(xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m), "#,##0.00")
                
                xlHoja1.Cells(31, m) = Format((IIf(CalifAnalist = 1 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento1, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento2, fCMDesConDescuento2 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento3, fCMDesConDescuento3 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento4, fCMDesConDescuento4 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento5, fCMDesConDescuento5 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento6, fCMDesConDescuento6 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, fCMDesConDescuento7 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , 0))))))), 0)) + _
                                        (IIf(CalifAnalist = 1 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento8, fCMDesConDescuento8 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento9, fCMDesConDescuento9 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento10, fCMDesConDescuento10 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)))))))), 0)), "#,##0.00")
                                        
                xlHoja1.Cells(32, m) = Format((IIf(CalifAnalist = 2 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento1, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento2, fCMDesConDescuento2 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento3, fCMDesConDescuento3 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento4, fCMDesConDescuento4 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento5, fCMDesConDescuento5 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento6, fCMDesConDescuento6 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, fCMDesConDescuento7 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , 0))))))), 0)) + _
                                        (IIf(CalifAnalist = 1 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento8, fCMDesConDescuento8 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento9, fCMDesConDescuento9 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento10, fCMDesConDescuento10 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)))))))), 0)), "#,##0.00")
                                        
                xlHoja1.Cells(33, m) = Format((IIf(CalifAnalist = 3 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento1, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento2, fCMDesConDescuento2 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento3, fCMDesConDescuento3 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento4, fCMDesConDescuento4 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento5, fCMDesConDescuento5 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento6, fCMDesConDescuento6 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, fCMDesConDescuento7 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , 0))))))), 0)) + _
                                        (IIf(CalifAnalist = 1 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento8, fCMDesConDescuento8 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento9, fCMDesConDescuento9 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento10, fCMDesConDescuento10 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)))))))), 0)), "#,##0.00")
                                        
                xlHoja1.Cells(34, m) = Format((IIf(CalifAnalist = 4 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento1, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento2, fCMDesConDescuento2 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento3, fCMDesConDescuento3 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento4, fCMDesConDescuento4 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento5, fCMDesConDescuento5 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento6, fCMDesConDescuento6 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, fCMDesConDescuento7 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , 0))))))), 0)) + _
                                        (IIf(CalifAnalist = 1 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento7, 0 _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento8, fCMDesConDescuento8 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento9, fCMDesConDescuento9 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento10, fCMDesConDescuento10 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , IIf(xlHoja1.Cells(20, m) <= fCMHastaConDescuento11, fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)) _
                                        , fCMDesConDescuento11 * (xlHoja1.Cells(22, m) + xlHoja1.Cells(23, m) + xlHoja1.Cells(24, m) + xlHoja1.Cells(25, m) + xlHoja1.Cells(26, m) + xlHoja1.Cells(27, m) + xlHoja1.Cells(28, m) + xlHoja1.Cells(29, m)))))))), 0)), "#,##0.00")
                
                xlHoja1.Cells(35, m) = Format(xlHoja1.Cells(30, m) + xlHoja1.Cells(31, m) + xlHoja1.Cells(32, m) + xlHoja1.Cells(33, m) + xlHoja1.Cells(34, m), "#,##0.00")
                
                xlHoja1.Cells(36, m) = IIf(IIf(xlHoja1.Cells(22, m) > 0, 1, 0) + IIf(xlHoja1.Cells(23, m) > 0, 1, 0) + IIf(xlHoja1.Cells(24, m) > 0, 1, 0) + IIf(xlHoja1.Cells(25, m) > 0, 1, 0) + IIf(xlHoja1.Cells(26, m) > 0, 1, 0) + IIf(xlHoja1.Cells(27, m) > 0, 1, 0) < 3, ((d26 + (d27 * 1)) - (xlHoja1.Cells(19, m) * 0.01) - (Casos * 15)) * fRentabilidad / 100, 0)
                
                xlHoja1.Cells(38, m) = xlHoja1.Cells(35, m) + xlHoja1.Cells(36, m) + xlHoja1.Cells(37, m)
                
                
                m = m + 1
                rsConsolidado.MoveFirst
                rsResumenXCierre.MoveFirst
                rsRentabilidad.MoveFirst
            End If
            rsResumen.MoveNext
        Loop
        
            xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsArchivo
            MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsArchivo
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
            Set xlAplicacion = Nothing
            Set xlLibro = Nothing
            Set xlHoja1 = Nothing
            Set objCOMDCredito = Nothing
            Set objCOMD_Rent = Nothing
            Set objCOMDCreditoBPPR = Nothing
            ReporteCredBonificacionAnalistaBPPR_Excel = ""
        Else
            MsgBox "No existen datos para generar el reporte"
            ReporteCredBonificacionAnalistaBPPR_Excel = ""
        End If
        
        'Comercial y Conv
'        Else
'            ReDim pMatProd(6)
'            pMatProd(0) = "101"
'            pMatProd(1) = "102"
'            pMatProd(2) = "103"
'            pMatProd(3) = "401"
'            pMatProd(4) = "403"
'            pMatProd(5) = "423"
'
'            Set rsResumen = objCOMDCredito.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False)
'            Set rsResumenXCierre = objCOMDCreditoBPPR.RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False, CDate(txtFCierreAnt.Text))
'            Set rsConsolidado = objCOMDCredito.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
'            sCadAge = ""
'            For i = 0 To UBound(pMatAgencias) - 1
'                sCadAge = sCadAge & pMatAgencias(i) & ","
'            Next i
'            sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
'            sCadProd = ""
'            If UBound(pMatProd) > 0 Then
'                For i = 0 To UBound(pMatProd) - 1
'                    sCadProd = sCadProd & pMatProd(i) & ","
'                Next i
'                sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
'            End If
'            sAn = ObtenerAnalistasxAgencia(sCadAge)
'            Set rsRentabilidad = objCOMD_Rent.RecuperaDatosConsolRentabilidadCarteraXAnalista(pdFecIni, pdFecFin, 8, sAn, sCadAge, sCadProd)

'            'If Not (rsResumen.EOF And rsResumen.BOF) Then
'                glsArchivo = "ResultadoAnalista" & Format(pdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
'                Set fs = New Scripting.FileSystemObject
'                Set xlAplicacion = New Excel.Application
'                If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
'                    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
'                Else
'                    Set xlLibro = xlAplicacion.Workbooks.Add
'                End If
'
'                Set xlHoja1 = xlLibro.Worksheets.Add
'                lbExisteHoja = False
'                lsNomHoja = "Resultado_Analista"
'                For Each xlHoja1 In xlLibro.Worksheets
'                    If xlHoja1.Name = lsNomHoja Then
'                        xlHoja1.Activate
'                        lbExisteHoja = True
'                    Exit For
'                    End If
'                Next
'
'                If lbExisteHoja = False Then
'                    Set xlHoja1 = xlLibro.Worksheets.Add
'                    xlHoja1.Name = lsNomHoja
'                End If
'
'                xlAplicacion.Range("A1:A1").ColumnWidth = 30
'                xlAplicacion.Range("B1:B1").ColumnWidth = 10
'                xlAplicacion.Range("C1:C1").ColumnWidth = 10
'                xlAplicacion.Range("D1:D1").ColumnWidth = 10
'                xlAplicacion.Range("E1:E1").ColumnWidth = 10
'                xlAplicacion.Range("F1:F1").ColumnWidth = 10
'                xlAplicacion.Range("G1:G1").ColumnWidth = 10
'                xlAplicacion.Range("H1:H1").ColumnWidth = 10
'                xlAplicacion.Range("I1:I1").ColumnWidth = 10
'                xlAplicacion.Range("A1:Z2000").Font.Size = 8
'                xlHoja1.Cells(1, 1) = "TIPO DE CAMBIO CIERRE"
'                xlHoja1.Cells(1, 2) = pnTipCam
'                xlHoja1.Cells(3, 1) = "RESULTADOS ANALISTAS" & " " & "COMERCIAL Y CONVENIO"
'                xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(3, 10)).Font.Bold = True
'                xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 10)).Merge True
'                xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
'                xlHoja1.Range("A4:A5").MergeCells = True
'                xlHoja1.Cells(4, 2) = "COMERCIAL"
'                xlHoja1.Cells(4, 7) = "CONVENIO"
'                xlHoja1.Range("B4:F4").MergeCells = True
'                xlHoja1.Range("G4:J4").MergeCells = True
'                xlHoja1.Cells(4, 1) = "ANALISTAS/CONCEPTOS"
'                xlHoja1.Cells(6, 1) = "CALIFIC"
'                xlHoja1.Cells(7, 1) = "META CREC. NUEVO"
'                xlHoja1.Cells(8, 1) = "META CREC. MONTO"
'                xlHoja1.Cells(9, 1) = "NUM. CLIENT. MES ANTER. O ARRASTRE"
'                xlHoja1.Cells(10, 1) = "SALDO CARTER MES ANTER. O ARRASTRE"
'                xlHoja1.Cells(11, 1) = "MORA > 30 MES ANTERIOR"
'                xlHoja1.Cells(12, 1) = "MORA JUDICIAL MES ANTERIOR"
'                xlHoja1.Cells(13, 1) = "TOTAL MORA"
'                xlHoja1.Cells(14, 1) = "NUM. DESEMBOLSOS"
'                xlHoja1.Cells(15, 1) = "NUM. CLIENT. CIERRE"
'                xlHoja1.Cells(16, 1) = "NUM. CLIENT. NUEVOS"
'                xlHoja1.Cells(17, 1) = "SALDO CARTERA CIERRE"
'                xlHoja1.Cells(18, 1) = "CARTA FIANZA"
'                xlHoja1.Cells(19, 1) = "MORA > A 30 CIERRE"
'                xlHoja1.Cells(20, 1) = "MORA JUDICIAL CIERRE"
'                xlHoja1.Cells(21, 1) = "SALDO CAPIT. UTILIZ. SOLES"
'                xlHoja1.Cells(22, 1) = "SALDO CAPIT. UTILIZ. DOLARES"
'                xlHoja1.Cells(23, 1) = "CAPITAL AMORTIZADO SOLES"
'                xlHoja1.Cells(24, 1) = "CAPITAL AMORTIZADO DOLARES"
'                xlHoja1.Cells(25, 1) = "N. OPERACIONES O CASOS"
'                xlHoja1.Cells(26, 1) = "ING. FINANC. SOLES (INTERESES)"
'                xlHoja1.Cells(27, 1) = "ING. FINANC. DOLARES (INTERESES)"
'                xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Cells.Interior.Color = RGB(220, 220, 220)
'                xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).HorizontalAlignment = xlCenter
'
'                m = 2
'                Do Until rsResumen.EOF
'                    If rsResumen!cUser <> "TOTAL" And rsResumen!nCar1Nro > 1 Then
'                        'Encabezado Analistas
'                        xlHoja1.Cells(5, m) = rsResumen!cUser
'                        xlHoja1.Cells(21, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(22, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(23, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(24, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(25, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(26, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(27, m) = Format(0, "#,##0.00")
'                        'Calif
'                        If rsResumen!nCar6Saldo <= 5000000 Then
'                            xlHoja1.Cells(6, m) = 2
'                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 3, 2)
'                            If rsMeta.RecordCount <> 0 Then
'                                'Meta Clie
'                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
'                                'Meta Monto
'                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
'                                Set rsMeta = Nothing
'                            End If
'                        End If
'                        If rsResumen!nCar6Saldo > 5000000 Then
'                            xlHoja1.Cells(6, m) = 1
'                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 3, 1)
'                            If rsMeta.RecordCount <> 0 Then
'                                'Meta Clie
'                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
'                                'Meta Monto
'                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento)
'                                Set rsMeta = Nothing
'                            End If
'                        End If
'
'                        'Cierre Anterior ***
'                        Do Until rsResumenXCierre.EOF
'                            If UCase(rsResumen!cUser) = UCase(rsResumenXCierre!cUser) Then
'                                'Num Client Cierre_Ant
'                                xlHoja1.Cells(9, m) = rsResumenXCierre!nNumCli6
'                                'Saldo Carter Mes Anter Cierre_Ant
'                                xlHoja1.Cells(10, m) = Format(rsResumenXCierre!nCar6Saldo, "#,##0.00")
'                                'Mora >30 Cierre
'                                xlHoja1.Cells(11, m) = Format(rsResumenXCierre!nCar4Saldo, "#,##0.00")
'                                'Mora Jud Cierre
'                                xlHoja1.Cells(12, m) = Format(val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
'                                'Total Mora
'                                xlHoja1.Cells(13, m) = Format(val(rsResumenXCierre!nCar4Saldo) + val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
'                                Exit Do
'                            End If
'                            rsResumenXCierre.MoveNext
'                        Loop
'
'                        Do Until rsConsolidado.EOF
'                            If UCase(rsResumen!cUser) = UCase(rsConsolidado!cUser) Then
'                                'Num Desembolsos Cierre
'                                xlHoja1.Cells(14, m) = val(rsConsolidado!nCantNuevo) + val(rsConsolidado!nCantRecurrente) + val(rsConsolidado!nCantParalelo) + val(rsConsolidado!nCantRefinanciado) + val(rsConsolidado!nCantAmpliado) + val(rsConsolidado!nCantAutomatico)
'                                'Num Client Nuev Cierre
'                                xlHoja1.Cells(16, m) = rsConsolidado!nCantNuevo
'                                Exit Do
'                            End If
'                            rsConsolidado.MoveNext
'                        Loop
'
'                        'Num Client Cierre Act
'                        xlHoja1.Cells(15, m) = rsResumen!nNumCli6
'                        'Saldo Cart Cierre Act
'                        xlHoja1.Cells(17, m) = Format(rsResumen!nCar6Saldo, "#,##0.00")
'                        'Mora >30 Cierre Act
'                        xlHoja1.Cells(19, m) = Format(rsResumen!nCar4Saldo, "#,##0.00")
'                        'Mora Jud Cierre Act
'                        xlHoja1.Cells(20, m) = Format(val(rsResumen!nCar5Saldo1) + val(rsResumen!nCar5Saldo2), "#,##0.00")
'
'                        'Rentabilidad de Cartera ***
'                        Do Until rsRentabilidad.EOF
'                            Casos = 0
'                            If UCase(rsResumen!cUser) = UCase(rsRentabilidad!cUser) Then
'                                'Saldo Capit. Utilizado S/.
'                                If rsRentabilidad!cMoney = 1 Then
'                                    xlHoja1.Cells(21, m) = rsRentabilidad!nSaldo
'                                    xlHoja1.Cells(23, m) = rsRentabilidad!Capital
'                                    xlHoja1.Cells(26, m) = rsRentabilidad!Interes
'                                Else
'                                    xlHoja1.Cells(22, m) = rsRentabilidad!nSaldo * pnTipCam
'                                    xlHoja1.Cells(24, m) = rsRentabilidad!Capital * pnTipCam
'                                    xlHoja1.Cells(27, m) = rsRentabilidad!Interes * pnTipCam
'                                End If
'                                Casos = Casos + rsRentabilidad!Casos
'                                xlHoja1.Cells(25, m) = Casos
'                            End If
'                            rsRentabilidad.MoveNext
'                        Loop
'                        'Rentabilidad de Cartera ***
'                        m = m + 1
'                        rsConsolidado.MoveFirst
'                        rsResumenXCierre.MoveFirst
'                        rsRentabilidad.MoveFirst
'                        End If
'                        rsResumen.MoveNext
'                Loop
'
'                ReDim pMatProd(2)
'                pMatProd(0) = "301"
'                pMatProd(1) = "320"
'
'                Set rsResumen = objCOMDCredito.RecuperaResumenSaldosCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False)
'                Set rsResumenXCierre = objCOMDCreditoBPPR.RecuperaResumenSaldosCarteraPorAnalistaConsolXCierre(psServerCons, pMatAgencias, pdFecIni, pnMoneda, 1, 7, 8, 15, 16, 30, 30, pMatCond, pMatProd, pnTipCam, False, CDate(txtFCierreAnt.Text))
'                Set rsConsolidado = objCOMDCredito.RecuperaConsolidadoCarteraPorAnalistaConsol(psServerCons, pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
'                sCadAge = ""
'                For i = 0 To UBound(pMatAgencias) - 1
'                    sCadAge = sCadAge & pMatAgencias(i) & ","
'                Next i
'                sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)
'                sCadProd = ""
'                If UBound(pMatProd) > 0 Then
'                    For i = 0 To UBound(pMatProd) - 1
'                        sCadProd = sCadProd & pMatProd(i) & ","
'                    Next i
'                    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
'                End If
'                sAn = ObtenerAnalistasxAgencia(sCadAge)
'                Set rsRentabilidad = objCOMD_Rent.RecuperaDatosConsolRentabilidadCarteraXAnalista(pdFecIni, pdFecFin, 8, sAn, sCadAge, sCadProd)
'
'                m = 7
'                Do Until rsResumen.EOF
'                    If rsResumen!cUser <> "TOTAL" And rsResumen!nCar1Nro > 1 Then
'                        'Encabezado Analistas
'                        xlHoja1.Cells(5, m) = rsResumen!cUser
'                        xlHoja1.Cells(21, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(22, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(23, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(24, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(25, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(26, m) = Format(0, "#,##0.00")
'                        xlHoja1.Cells(27, m) = Format(0, "#,##0.00")
'                        'Calif
'                        If rsResumen!nCar6Saldo <= 3000000 Then
'                            xlHoja1.Cells(6, m) = 3
'                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 2, 3)
'                            If rsMeta.RecordCount <> 0 Then
'                                'Meta Clie
'                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
'                                'Meta Monto
'                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
'                                Set rsMeta = Nothing
'                            End If
'                        End If
'                        If rsResumen!nCar6Saldo > 3000000 And rsResumen!nCar6Saldo <= 5500000 Then
'                            xlHoja1.Cells(6, m) = 2
'                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 2, 2)
'                            If rsMeta.RecordCount <> 0 Then
'                                'Meta Clie
'                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
'                                'Meta Monto
'                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento, "#,##0.00")
'                                Set rsMeta = Nothing
'                            End If
'                        End If
'                        If rsResumen!nCar6Saldo > 5500000 Then
'                            xlHoja1.Cells(6, m) = 1
'                            Set rsMeta = objCOMDCreditoBPPR.CargarDatosMeta(sCadAge, 2, 1)
'                            If rsMeta.RecordCount <> 0 Then
'                                'Meta Clie
'                                xlHoja1.Cells(7, m) = rsMeta!iClienteNuevo
'                                'Meta Monto
'                                xlHoja1.Cells(8, m) = Format(rsMeta!fMetaCrecimiento)
'                                Set rsMeta = Nothing
'                            End If
'                        End If
'
'                        'Cierre Anterior ***
'                        Do Until rsResumenXCierre.EOF
'                            If UCase(rsResumen!cUser) = UCase(rsResumenXCierre!cUser) Then
'                                'Num Client Cierre_Ant
'                                xlHoja1.Cells(9, m) = rsResumenXCierre!nNumCli6
'                                'Saldo Carter Mes Anter Cierre_Ant
'                                xlHoja1.Cells(10, m) = Format(rsResumenXCierre!nCar6Saldo, "#,##0.00")
'                                'Mora >30 Cierre
'                                xlHoja1.Cells(11, m) = Format(rsResumenXCierre!nCar4Saldo, "#,##0.00")
'                                'Mora Jud Cierre
'                                xlHoja1.Cells(12, m) = Format(val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
'                                'Total Mora
'                                xlHoja1.Cells(13, m) = Format(val(rsResumenXCierre!nCar4Saldo) + val(rsResumenXCierre!nCar5Saldo1) + val(rsResumenXCierre!nCar5Saldo2), "#,##0.00")
'                                Exit Do
'                            End If
'                            rsResumenXCierre.MoveNext
'                        Loop
'
'                        Do Until rsConsolidado.EOF
'                            If UCase(rsResumen!cUser) = UCase(rsConsolidado!cUser) Then
'                                'Num Desembolsos Cierre
'                                xlHoja1.Cells(14, m) = val(rsConsolidado!nCantNuevo) + val(rsConsolidado!nCantRecurrente) + val(rsConsolidado!nCantParalelo) + val(rsConsolidado!nCantRefinanciado) + val(rsConsolidado!nCantAmpliado) + val(rsConsolidado!nCantAutomatico)
'                                'Num Client Nuev Cierre
'                                xlHoja1.Cells(16, m) = rsConsolidado!nCantNuevo
'                                Exit Do
'                            End If
'                            rsConsolidado.MoveNext
'                        Loop
'
'                        'Num Client Cierre Act
'                        xlHoja1.Cells(15, m) = rsResumen!nNumCli6
'                        'Saldo Cart Cierre Act
'                        xlHoja1.Cells(17, m) = Format(rsResumen!nCar6Saldo, "#,##0.00")
'                        'Mora >30 Cierre Act
'                        xlHoja1.Cells(19, m) = Format(rsResumen!nCar4Saldo, "#,##0.00")
'                        'Mora Jud Cierre Act
'                        xlHoja1.Cells(20, m) = Format(val(rsResumen!nCar5Saldo1) + val(rsResumen!nCar5Saldo2), "#,##0.00")
'
'                        'Rentabilidad de Cartera ***
'                        Do Until rsRentabilidad.EOF
'                            Casos = 0
'                            If UCase(rsResumen!cUser) = UCase(rsRentabilidad!cUser) Then
'                                'Saldo Capit. Utilizado S/.
'                                If rsRentabilidad!cMoney = 1 Then
'                                    xlHoja1.Cells(21, m) = rsRentabilidad!nSaldo
'                                    xlHoja1.Cells(23, m) = rsRentabilidad!Capital
'                                    xlHoja1.Cells(26, m) = rsRentabilidad!Interes
'                                Else
'                                    xlHoja1.Cells(22, m) = rsRentabilidad!nSaldo * pnTipCam
'                                    xlHoja1.Cells(24, m) = rsRentabilidad!Capital * pnTipCam
'                                    xlHoja1.Cells(27, m) = rsRentabilidad!Interes * pnTipCam
'                                End If
'                                Casos = Casos + rsRentabilidad!Casos
'                                xlHoja1.Cells(25, m) = Casos
'                            End If
'                            rsRentabilidad.MoveNext
'                        Loop
'                        'Rentabilidad de Cartera ***
'                        m = m + 1
'                        rsConsolidado.MoveFirst
'                        rsResumenXCierre.MoveFirst
'                        rsRentabilidad.MoveFirst
'                        End If
'                        rsResumen.MoveNext
'                Loop
'
'                xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
'                MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
'                xlAplicacion.Visible = True
'                xlAplicacion.Windows(1).Visible = True
'                Set xlAplicacion = Nothing
'                Set xlLibro = Nothing
'                Set xlHoja1 = Nothing
'                Set objCOMDCredito = Nothing
'                Set objCOMD_Rent = Nothing
'                Set objCOMDCreditoBPPR = Nothing
'                ReporteCredResultadoAnalistaBPPR_Excel = ""
    End If
End Function

'MAVM 20100520
'WIOR 20140620 ***************************************************************
Private Sub GeneraCartaCargoCuentaAhorro(ByVal pdFecha As Date, ByVal psAnalista As String, ByVal psAgencia As String)
Dim j As Integer
Dim oCOMNCredDoc As COMNCredito.NCOMColocEval
Dim rsCarta As New ADODB.Recordset
Dim lsModeloPlantilla As String
Dim nCuotaMin As Long
Dim nCuotaMax As Long
Dim Cuotas As String
Dim fs As Scripting.FileSystemObject
   
On Error GoTo ErrGeneraRepo

lsModeloPlantilla = App.Path & "\FormatoCarta\CartaCargoCuentaAhorro.doc"

    Set fs = New Scripting.FileSystemObject

    If Not fs.FileExists(lsModeloPlantilla) Then
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    Screen.MousePointer = 11
    Set oCOMNCredDoc = New COMNCredito.NCOMColocEval
    Set rsCarta = oCOMNCredDoc.RecuperaCreditoCargoDebito(pdFecha, psAnalista, psAgencia)
    Set oCOMNCredDoc = Nothing
     
    If rsCarta.BOF Then
        Screen.MousePointer = 0
        MsgBox "No hay Datos", vbInformation, "Aviso"
        Exit Sub
    End If

    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
    
    
    Dim RangeSource As Word.Range
    
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
    
    'Crea Nuevo Documento
    wApp.Documents.Add
    
    
    Do While Not rsCarta.EOF
    
        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
     
        With wApp.Selection.Find
            .Text = "dFecha"
            .Replacement.Text = Trim(ImpreFormat(Format(gdFecSis, "d mmmm yyyy"), 25))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "CCIUDADAGE"
            .Replacement.Text = Trim(rsCarta!Ciudad)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cCliAho"
            .Replacement.Text = Trim(PstaNombre(rsCarta!CliAhorro, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cDireccion"
            .Replacement.Text = Trim(rsCarta!Direccion)
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cCliCred"
            .Replacement.Text = Trim(PstaNombre(rsCarta!CliCred, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
         With wApp.Selection.Find
            .Text = "cCredito"
            .Replacement.Text = Trim(rsCarta!CuentaCred)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cSaldoPend"
            .Replacement.Text = IIf(Mid(Trim(rsCarta!CuentaCred), 9, 1) = "1", "S/. ", "US $ ") & Format(rsCarta!PendientePagar, "###," & String(15, "#") & "#0.00")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "dFecVenc"
            .Replacement.Text = Format(CDate(rsCarta!CuentaVence), "dd/mm/yyyy")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cMontoDebito"
            .Replacement.Text = IIf(Mid(Trim(rsCarta!CuentaAho), 9, 1) = "1", "S/. ", "US $ ") & Format(rsCarta!MontoDebito, "###," & String(15, "#") & "#0.00")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "cAhorros"
            .Replacement.Text = Trim(rsCarta!CuentaAho)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        nCuotaMin = CLng(rsCarta!CuotaMin)
        nCuotaMax = CLng(rsCarta!CuotaMax)
        
        Cuotas = ""
        
        If nCuotaMin = nCuotaMax Then
            Cuotas = nCuotaMin
        Else
            For j = nCuotaMin To nCuotaMax
                If j = nCuotaMax - 1 Then
                    Cuotas = Cuotas & j & " y"
                Else
                    Cuotas = Cuotas & j & ", "
                End If
            Next j
            
            Cuotas = Mid(Cuotas, 1, Len(Cuotas) - 2)
        End If
        
        
        With wApp.Selection.Find
            .Text = "cCuotas"
            .Replacement.Text = Trim(Cuotas)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll

        rsCarta.MoveNext
    Loop
    rsCarta.Close
    Set rsCarta = Nothing
 
Screen.MousePointer = 0
 
wAppSource.ActiveDocument.Close
wApp.Visible = True

Exit Sub
ErrGeneraRepo:
    Screen.MousePointer = 0
    wAppSource.ActiveDocument.Close
    MsgBox "Error en frmCredReportes.GeneraCartaCargoCuentaAhorro " & err.Description, vbInformation, "Aviso"

End Sub
'WIOR FIN ********************************************************************
