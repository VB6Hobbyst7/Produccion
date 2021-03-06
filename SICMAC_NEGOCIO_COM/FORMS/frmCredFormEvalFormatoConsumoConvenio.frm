VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormatoConsumoConvenio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cr?ditos - Evaluaci?n - Formato Descuento Por Planilla."
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12420
   Icon            =   "frmCredFormEvalFormatoConsumoConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame16 
      Caption         =   "Ratios e indicadores"
      ForeColor       =   &H8000000D&
      Height          =   650
      Left            =   240
      TabIndex        =   131
      Top             =   9340
      Width           =   12015
      Begin VB.TextBox txtCapPagoConConvenio1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   8640
         TabIndex        =   133
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCapPagoConConvenio2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   9840
         TabIndex        =   132
         Top             =   240
         Width           =   975
      End
      Begin SICMACT.EditMoney txtCapacidadNeta 
         Height          =   300
         Left            =   1560
         TabIndex        =   134
         Top             =   280
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
      End
      Begin VB.Label Label29 
         Caption         =   "Limite Cuota"
         Height          =   255
         Left            =   7680
         TabIndex        =   137
         Top             =   320
         Width           =   945
      End
      Begin VB.Label lblCapacidadPago 
         Caption         =   "Capacidad Pago:"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   320
         Width           =   1300
      End
      Begin VB.Label lblCapaAceptable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2640
         TabIndex        =   135
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.Frame Frame14 
      Height          =   2775
      Left            =   9795
      TabIndex        =   123
      Top             =   600
      Width           =   2535
      Begin VB.CommandButton cmdInformeVisitaConConvenio 
         Caption         =   "Informe de Visita"
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
         Left            =   480
         TabIndex        =   128
         Top             =   2280
         Width           =   1700
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Hoja Evaluaci?n"
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
         Left            =   480
         TabIndex        =   127
         Top             =   1800
         Width           =   1700
      End
      Begin VB.CommandButton cmdCancelarConConvenio 
         Caption         =   "Salir"
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
         Left            =   480
         TabIndex        =   126
         Top             =   840
         Width           =   1700
      End
      Begin VB.CommandButton cmdActualizarConConvenio 
         Caption         =   "Guardar"
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
         Left            =   480
         TabIndex        =   125
         Top             =   360
         Width           =   1700
      End
      Begin VB.CommandButton cmdGuardarConConvenio 
         Caption         =   "Guardar"
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
         Left            =   480
         TabIndex        =   124
         Top             =   360
         Width           =   1700
      End
   End
   Begin TabDlg.SSTab SSTabInfoNego 
      Height          =   3735
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Informaci?n del Negocio"
      TabPicture(0)   =   "frmCredFormEvalFormatoConsumoConvenio.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtInfNegocioFuenteIngreso"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ActXCodCta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtInfNegocioActividad"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtInfNegocioActividad 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5350
         TabIndex        =   100
         Top             =   430
         Width           =   3975
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   9255
         Begin VB.Frame Frame13 
            Enabled         =   0   'False
            Height          =   495
            Left            =   2200
            TabIndex        =   115
            Top             =   1820
            Width           =   3255
            Begin VB.OptionButton optTipoPlanilla 
               Caption         =   "Cesantes"
               Height          =   195
               Index           =   3
               Left            =   1920
               TabIndex        =   118
               Top             =   200
               Width           =   975
            End
            Begin VB.OptionButton optTipoPlanilla 
               Caption         =   "Activos"
               Height          =   255
               Index           =   2
               Left            =   920
               TabIndex        =   117
               Top             =   200
               Width           =   975
            End
            Begin VB.OptionButton optTipoPlanilla 
               Caption         =   "CAS"
               Height          =   255
               Index           =   1
               Left            =   50
               TabIndex        =   116
               Top             =   200
               Width           =   855
            End
         End
         Begin VB.Frame Frame12 
            Height          =   495
            Left            =   2200
            TabIndex        =   114
            Top             =   1340
            Width           =   3255
            Begin VB.CheckBox ChkSectorSalud 
               Caption         =   "Sector Salud"
               Height          =   255
               Left            =   1920
               TabIndex        =   5
               Top             =   200
               Width           =   1300
            End
            Begin VB.OptionButton optTipoInstitucion 
               Caption         =   "Privada"
               Height          =   195
               Index           =   2
               Left            =   920
               TabIndex        =   4
               Top             =   200
               Width           =   855
            End
            Begin VB.OptionButton optTipoInstitucion 
               Caption         =   "P?blica"
               Height          =   195
               Index           =   1
               Left            =   50
               TabIndex        =   3
               Top             =   200
               Width           =   975
            End
         End
         Begin VB.Frame Frame11 
            Height          =   495
            Left            =   2200
            TabIndex        =   113
            Top             =   860
            Width           =   3255
            Begin VB.OptionButton optTipoAportacion 
               Caption         =   "ONP"
               Height          =   255
               Index           =   2
               Left            =   920
               TabIndex        =   2
               Top             =   200
               Width           =   855
            End
            Begin VB.OptionButton optTipoAportacion 
               Caption         =   "AFP"
               Height          =   195
               Index           =   1
               Left            =   50
               TabIndex        =   1
               Top             =   200
               Width           =   855
            End
         End
         Begin VB.TextBox txtInstConv 
            Enabled         =   0   'False
            Height          =   780
            Left            =   5520
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   1515
            Width           =   3615
         End
         Begin VB.TextBox txtInfNegocioCliente 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2200
            TabIndex        =   45
            Top             =   160
            Width           =   6975
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda2 
            Height          =   300
            Left            =   8040
            TabIndex        =   24
            Top             =   1035
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin SICMACT.EditMoney txtInfNegocioCuotas 
            Height          =   300
            Left            =   4440
            TabIndex        =   25
            Top             =   2445
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "00"
         End
         Begin Spinner.uSpinner spnInfNegocioA?o 
            Height          =   315
            Left            =   2200
            TabIndex        =   26
            Top             =   555
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin Spinner.uSpinner spnInfNegocioMes 
            Height          =   315
            Left            =   3480
            TabIndex        =   27
            Top             =   555
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   12
            MaxLength       =   2
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin SICMACT.EditMoney txtInfNegocioMontSolicitado 
            Height          =   300
            Left            =   2160
            TabIndex        =   28
            Top             =   2450
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0.00"
         End
         Begin SICMACT.EditMoney txtInfNegocioExpCredito 
            Height          =   300
            Left            =   8040
            TabIndex        =   39
            Top             =   2445
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0.00"
         End
         Begin SICMACT.EditMoney txtInfNegocioUltDeuda 
            Height          =   300
            Left            =   8040
            TabIndex        =   42
            Top             =   555
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0.00"
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N? Cuotas :"
            Height          =   195
            Left            =   3550
            TabIndex        =   44
            Top             =   2550
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exposici?n con este Cr?dito:"
            Height          =   195
            Left            =   5960
            TabIndex        =   43
            Top             =   2550
            Width           =   2010
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Planilla:"
            Height          =   195
            Left            =   1080
            TabIndex        =   41
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Instituci?n:"
            Height          =   195
            Left            =   840
            TabIndex        =   40
            Top             =   1560
            Width           =   1350
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   1680
            TabIndex        =   36
            Top             =   220
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Antiguedad en actual Empleo:"
            Height          =   195
            Left            =   60
            TabIndex        =   35
            Top             =   620
            Width           =   2130
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Aportaci?n:"
            Height          =   195
            Left            =   780
            TabIndex        =   34
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Solicitado:"
            Height          =   195
            Left            =   930
            TabIndex        =   33
            Top             =   2520
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a?os"
            Height          =   255
            Left            =   2955
            TabIndex        =   32
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   4275
            TabIndex        =   31
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "?ltimo endeudamiento RCC:"
            Height          =   195
            Left            =   6000
            TabIndex        =   30
            Top             =   615
            Width           =   1995
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha ?ltimo endeudamiento RCC:"
            Height          =   195
            Left            =   5520
            TabIndex        =   29
            Top             =   1080
            Width           =   2460
         End
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         Texto           =   "Cr?dito"
      End
      Begin MSMask.MaskEdBox txtInfNegocioFuenteIngreso 
         Height          =   300
         Left            =   8040
         TabIndex        =   121
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Evaluaci?n al :"
         Height          =   195
         Left            =   6240
         TabIndex        =   122
         Top             =   45
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actividad :"
         Height          =   195
         Left            =   4600
         TabIndex        =   38
         Top             =   480
         Width           =   750
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5565
      Left            =   15
      TabIndex        =   119
      Top             =   3765
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   9816
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   4904
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ingresos y Egresos"
      TabPicture(0)   =   "frmCredFormEvalFormatoConsumoConvenio.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Evaluaci?n Aval"
      TabPicture(1)   =   "frmCredFormEvalFormatoConsumoConvenio.frx":0342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Propuesta del Cr?dito"
      TabPicture(2)   =   "frmCredFormEvalFormatoConsumoConvenio.frx":035E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Comentarios y Referidos"
      TabPicture(3)   =   "frmCredFormEvalFormatoConsumoConvenio.frx":037A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(1)=   "Frame10"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame8 
         Caption         =   "Propuesta del Credito"
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
         Height          =   5100
         Left            =   -74880
         TabIndex        =   105
         Top             =   360
         Width           =   12015
         Begin VB.TextBox txtPropCreditoEntornoFamiliar 
            Height          =   550
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   500
            Width           =   11535
         End
         Begin VB.TextBox txtPropCreditoGiroNegocio 
            Height          =   550
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   1250
            Width           =   11535
         End
         Begin VB.TextBox txtPropCreditoExpCrediticia 
            Height          =   550
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   2020
            Width           =   11535
         End
         Begin VB.TextBox txtPropCreditoFormNegocio 
            Height          =   550
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   2820
            Width           =   11535
         End
         Begin VB.TextBox txtPropCreditoColateralesGarantias 
            Height          =   550
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   3650
            Width           =   11535
         End
         Begin VB.TextBox txtPropCreditoDestino 
            Height          =   550
            Left            =   240
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   4450
            Width           =   11535
         End
         Begin MSMask.MaskEdBox txtPropCreditoFechaVista 
            Height          =   345
            Left            =   9720
            TabIndex        =   14
            Top             =   140
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label45 
            Caption         =   "Fecha de Visita:"
            Height          =   270
            Left            =   8520
            TabIndex        =   112
            Top             =   250
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Sobre el Entorno Familiar del Cliente o Representante"
            Height          =   300
            Left            =   240
            TabIndex        =   111
            Top             =   250
            Width           =   4695
         End
         Begin VB.Label Label43 
            Caption         =   "Sobre el Giro y la Ubicacion del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   110
            Top             =   1060
            Width           =   4095
         End
         Begin VB.Label Label42 
            Caption         =   "Sobre la Experiencia Crediticia"
            Height          =   300
            Left            =   240
            TabIndex        =   109
            Top             =   1850
            Width           =   4215
         End
         Begin VB.Label Label41 
            Caption         =   "Sobre la Consistencia de la Informacion y la Formalidad del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   108
            Top             =   2620
            Width           =   6255
         End
         Begin VB.Label Label40 
            Caption         =   "Sobre los Colaterales o Garantias"
            Height          =   300
            Left            =   240
            TabIndex        =   107
            Top             =   3450
            Width           =   3975
         End
         Begin VB.Label Label39 
            Caption         =   "Sobre el Destino y el Impacto del Mismo"
            Height          =   300
            Left            =   240
            TabIndex        =   106
            Top             =   4240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Referidos"
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
         Height          =   2775
         Left            =   -74880
         TabIndex        =   102
         Top             =   2640
         Width           =   11895
         Begin VB.CommandButton cmdQuitarConConvenio 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1320
            TabIndex        =   104
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarConConvenio 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   103
            Top             =   2280
            Width           =   1095
         End
         Begin SICMACT.FlexEdit feReferidosConConvenio 
            Height          =   1695
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   11520
            _ExtentX        =   20320
            _ExtentY        =   2990
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   1
            EncabezadosNombres=   "N?-Nombre-DNI-Telef.-Comentarios-DNI-Aux"
            EncabezadosAnchos=   "400-3500-1100-1100-4500-0-0"
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "R-L-L-L-L-L-C"
            FormatosEdit    =   "3-1-0-0-1-0-3"
            TextArray0      =   "N?"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            TipoBusPersona  =   2
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Comentarios"
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
         Height          =   1935
         Left            =   -74880
         TabIndex        =   101
         Top             =   480
         Width           =   11895
         Begin VB.TextBox txtReferidosComentario 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   11535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Propuesta del Credito"
         Height          =   6495
         Left            =   -74880
         TabIndex        =   77
         Top             =   480
         Width           =   9375
         Begin VB.TextBox txtSustentoIncreVenta 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   5640
            Width           =   9015
         End
         Begin VB.TextBox txtGarantias 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   4680
            Width           =   9015
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   3720
            Width           =   9015
         End
         Begin VB.TextBox txtCrediticia 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   2760
            Width           =   9015
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   79
            Top             =   1800
            Width           =   9015
         End
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   78
            Top             =   840
            Width           =   9015
         End
         Begin MSMask.MaskEdBox txtFechaVista 
            Height          =   345
            Left            =   7920
            TabIndex        =   84
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label33 
            Caption         =   "Sobre el Destino y el Impacto del Mismo"
            Height          =   300
            Left            =   240
            TabIndex        =   91
            Top             =   5400
            Width           =   4575
         End
         Begin VB.Label Label32 
            Caption         =   "Sobre los Colaterales o Garantias"
            Height          =   300
            Left            =   240
            TabIndex        =   90
            Top             =   4440
            Width           =   3975
         End
         Begin VB.Label Label31 
            Caption         =   "Sobre la Consistencia de la Informacion y la Formalidad del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   89
            Top             =   3480
            Width           =   6255
         End
         Begin VB.Label Label30 
            Caption         =   "Sobre la Experiencia Crediticia"
            Height          =   300
            Left            =   240
            TabIndex        =   88
            Top             =   2520
            Width           =   4215
         End
         Begin VB.Label Label27 
            Caption         =   "Sobre el Giro y la Ubicacion del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   87
            Top             =   1560
            Width           =   4095
         End
         Begin VB.Label Label26 
            Caption         =   "Sobre el Entorno Familiar del Cliente o Representante"
            Height          =   300
            Left            =   240
            TabIndex        =   86
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label34 
            Caption         =   "Fecha de Vista:"
            Height          =   300
            Left            =   6720
            TabIndex        =   85
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Calculo de Capacidad de Pago"
         ForeColor       =   &H8000000D&
         Height          =   4860
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   12015
         Begin VB.Frame Frame17 
            Caption         =   "Ingresos :"
            Height          =   2355
            Left            =   6120
            TabIndex        =   138
            Top             =   2400
            Width           =   5775
            Begin SICMACT.FlexEdit fgIngresos 
               Height          =   1815
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   4755
               _ExtentX        =   8387
               _ExtentY        =   3201
               Cols0           =   5
               HighLight       =   1
               EncabezadosNombres=   "Index-N-Concepto-Monto-Aux"
               EncabezadosAnchos=   "0-400-2700-1500-0"
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
               ColumnasAEditar =   "X-X-X-3-X"
               ListaControles  =   "0-0-0-1-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-R-L-R-C"
               FormatosEdit    =   "0-3-0-2-2"
               TextArray0      =   "Index"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   6
               lbBuscaDuplicadoText=   -1  'True
               RowHeight0      =   300
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Gastos Familiares : "
            Height          =   2355
            Left            =   120
            TabIndex        =   129
            Top             =   2400
            Width           =   5895
            Begin SICMACT.FlexEdit feGastosFamiliares 
               Height          =   1755
               Left            =   120
               TabIndex        =   130
               Top             =   240
               Width           =   5745
               _ExtentX        =   10134
               _ExtentY        =   3096
               Cols0           =   5
               HighLight       =   1
               EncabezadosNombres=   "-N-Concepto-Monto-Aux"
               EncabezadosAnchos=   "0-300-3500-1800-0"
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
               ColumnasAEditar =   "X-X-X-3-X"
               ListaControles  =   "0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-L-R-C"
               FormatosEdit    =   "0-0-0-2-0"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               TipoBusqueda    =   6
               lbBuscaDuplicadoText=   -1  'True
               RowHeight0      =   300
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Evaluaci?n mes 2"
            Height          =   2050
            Left            =   3960
            TabIndex        =   92
            Top             =   220
            Width           =   3435
            Begin VB.TextBox txtIngrNetoMes2 
               Alignment       =   1  'Right Justify
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
               Left            =   1350
               TabIndex        =   95
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox txtDescuentoMes2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1350
               TabIndex        =   94
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox txtRemBrutaTotalMes2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1350
               TabIndex        =   93
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtAnoMes2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2535
               TabIndex        =   11
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox cmbFechaMes2 
               Height          =   315
               Left            =   1350
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdLlamaRemBrutaTotalMes2 
               Caption         =   "..."
               Height          =   300
               Left            =   2520
               TabIndex        =   12
               Top             =   720
               Width           =   495
            End
            Begin VB.CommandButton cmdLlamaDescuentoMes2 
               Caption         =   "..."
               Height          =   300
               Left            =   2520
               TabIndex        =   13
               Top             =   1200
               Width           =   495
            End
            Begin VB.Line Line2 
               X1              =   1320
               X2              =   3000
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label Label38 
               Caption         =   "Ingreso Neto :"
               Height          =   225
               Left            =   350
               TabIndex        =   99
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label37 
               Caption         =   "Descuentos :"
               Height          =   300
               Left            =   420
               TabIndex        =   98
               Top             =   1200
               Width           =   1005
            End
            Begin VB.Label Label36 
               Caption         =   "Rem. Bruta Total :"
               Height          =   300
               Left            =   80
               TabIndex        =   97
               Top             =   720
               Width           =   1305
            End
            Begin VB.Label Label35 
               Caption         =   "Mes - A?o :"
               Height          =   300
               Left            =   520
               TabIndex        =   96
               Top             =   315
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Evaluaci?n mes 1"
            Height          =   2050
            Left            =   80
            TabIndex        =   69
            Top             =   220
            Width           =   3560
            Begin VB.CommandButton cmdLlamaDescuentoMes1 
               Caption         =   "..."
               Height          =   300
               Left            =   2640
               TabIndex        =   9
               Top             =   1200
               Width           =   495
            End
            Begin VB.ComboBox cmbFechaMes1 
               Height          =   315
               Left            =   1400
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtAnoMes1 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2620
               TabIndex        =   7
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox txtRemBrutaTotalMes1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1400
               TabIndex        =   72
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtDescuentoMes1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1400
               TabIndex        =   71
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox txtIngrNetoMes1 
               Alignment       =   1  'Right Justify
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
               Left            =   1400
               TabIndex        =   70
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton cmdLlamaRemBrutaTotalMes1 
               Caption         =   "..."
               Height          =   300
               Left            =   2640
               TabIndex        =   8
               Top             =   720
               Width           =   495
            End
            Begin VB.Line Line1 
               X1              =   1440
               X2              =   3120
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label Label15 
               Caption         =   "Mes - A?o :"
               Height          =   300
               Left            =   550
               TabIndex        =   76
               Top             =   320
               Width           =   855
            End
            Begin VB.Label Label19 
               Caption         =   "Rem. Bruta Total :"
               Height          =   300
               Left            =   80
               TabIndex        =   75
               Top             =   720
               Width           =   1305
            End
            Begin VB.Label Label20 
               Caption         =   "Descuentos :"
               Height          =   300
               Left            =   420
               TabIndex        =   74
               Top             =   1200
               Width           =   1005
            End
            Begin VB.Label Label21 
               Caption         =   "Ingreso Neto :"
               Height          =   225
               Left            =   340
               TabIndex        =   73
               Top             =   1680
               Width           =   1005
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Promedios"
            Height          =   2050
            Left            =   7800
            TabIndex        =   60
            Top             =   220
            Width           =   4095
            Begin VB.TextBox txtRemBrutaTotalPromedio 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   64
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtDescuentoPromedio 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   63
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtIngNetolPromedio 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   2640
               TabIndex        =   62
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtMontoMaxIngDescontarPromedio 
               Alignment       =   1  'Right Justify
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
               Left            =   2640
               TabIndex        =   61
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Line Line3 
               X1              =   2640
               X2              =   3720
               Y1              =   1560
               Y2              =   1560
            End
            Begin VB.Label lblMontoMaxIngDes 
               Caption         =   "Monto Max. Ingreso a Descontar :"
               Height          =   300
               Left            =   165
               TabIndex        =   65
               Top             =   1720
               Width           =   2445
            End
            Begin VB.Label lblMontoDispo 
               Caption         =   "Monto Disponible :"
               Height          =   255
               Left            =   1200
               TabIndex        =   120
               Top             =   2.45745e5
               Width           =   1455
            End
            Begin VB.Label Label22 
               Caption         =   "Rem. Bruta Total :"
               Height          =   300
               Left            =   1260
               TabIndex        =   68
               Top             =   315
               Width           =   1500
            End
            Begin VB.Label Label23 
               Caption         =   "Descuentos de Ley :"
               Height          =   300
               Left            =   1080
               TabIndex        =   67
               Top             =   790
               Width           =   1695
            End
            Begin VB.Label Label24 
               Caption         =   "Ingreso Neto :"
               Height          =   300
               Left            =   1560
               TabIndex        =   66
               Top             =   1260
               Width           =   1455
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Evaluacion mes 2"
         Height          =   1815
         Left            =   4320
         TabIndex        =   47
         Top             =   1020
         Width           =   3855
         Begin VB.TextBox txtMonParalelo 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   54
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtResumenIncIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   1560
            TabIndex        =   53
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   1560
            TabIndex        =   52
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2880
            TabIndex        =   51
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            TabIndex        =   50
            Text            =   "Combo1"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   300
            Left            =   2880
            TabIndex        =   49
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   300
            Left            =   2880
            TabIndex        =   48
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "Ingreso Neto :"
            Height          =   225
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Descuentos :"
            Height          =   300
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label17 
            Caption         =   "Rem. Bruta Total :"
            Height          =   300
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "Mes - A?o :"
            Height          =   300
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmCredFormEvalFormatoConsumoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalFormatoConsumoConvenio
'** Descripci?n : Formulario para evaluaci?n de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creaci?n    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'
Option Explicit
Dim gsOpeCod As String

Dim fnTipoRegMant As Integer
Dim fsCtaCod As String

Dim fsGiroNego As String
Dim fsCliente As String
Dim fnAnio As Integer
Dim fnMes As Integer
Dim fnMontoDeudaSbs As Double
Dim fdFechaDeudaSbs As Date
Dim fsInstConv As String
Dim fnMontSolicitado As Double
Dim fnCuota As Integer
Dim fnExpCredito As Double
Dim fdFechaActual As Date

Dim MtrRemuneracionBrutaTotal1 As Variant
Dim MtrRemuneracionBrutaTotal2 As Variant
Dim MtrDescuento1 As Variant
Dim MtrDescuento2 As Variant

Dim MatReferidos As Variant
Dim fnTipoAportacion As Integer
Dim fnTipoInstitucion As Integer
Dim fnSectorSalud As Integer
Dim fnTipoPlanilla As Integer

Dim lnColocCondi As Integer
Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agreg? segun correo: RUSI
Dim nEstado As Integer

Dim fbGrabar As Boolean

Dim nTotalCompraDeu1 As Currency
Dim nTotalCompraDeu2 As Currency
Dim i As Integer, lnFila As Integer

Enum TipoInstitucion
    nTpoPublico = 1
    nTpoPrivado = 2
End Enum

Dim objPista As COMManejador.Pista

Dim fnTipoPermiso As Integer
Dim lcMovNro As String 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018

'CTI320200110 ERS003-2020. Agreg?
Dim j As Integer
Dim fnFormato As Integer
Dim rsFeGastoNeg As ADODB.Recordset
Dim rsFeDatGastoFam As ADODB.Recordset
Dim MatIfiGastoFami As Variant
Dim MatIfiNoSupervisadaGastoFami As Variant
Dim fnTotalRefGastoFami As Currency

Dim rsGastoFam As ADODB.Recordset
Dim rsDatGastoFam As ADODB.Recordset
Dim rsDatIfiGastoFami As ADODB.Recordset
Dim rsDatIfiNoSupervisadaGastoFami As ADODB.Recordset
Dim rsDatRatios As ADODB.Recordset

Dim rsRatiosActual As ADODB.Recordset
Dim rsRatiosAceptableCritico As ADODB.Recordset
Dim rsAceptableCritico As ADODB.Recordset
Dim fbImprimirVB As Boolean
'CTI3 ERS0032020*******************************
Dim pnIngNegocio As Double
Dim pnEgrVenta As Double
Dim pnMargBruto As Double
Dim pnIngNeto As Double
Dim pMtrNegocio As Variant
Dim pMtrBoletaPago As Variant
Dim pMtrReciboHono As Variant
Dim rsFeDatOtrosIng As ADODB.Recordset
Dim pnMontoOtrasIfis As Double

'**********************************************
'Fin CTI320200110

Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                     ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer, _
                     Optional ByVal pbImprimirVB As Boolean = False) As Boolean
    
    gsOpeCod = ""
    lcMovNro = "" 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
    fbImprimirVB = pbImprimirVB 'CTI3ERS0032020
    fnTipoAportacion = 0
    fnTipoInstitucion = 0
    fnTipoPlanilla = 0
    
    nEstado = pnEstado
    fnTipoRegMant = psTipoRegMant
    fsCtaCod = psCtaCod
    
    If nEstado = 2001 Then
        If lnColocCondi <> 4 Then
            cmdImprimir.Enabled = True
            cmdActualizarConConvenio.Enabled = True
        End If
    Else
        cmdInformeVisitaConConvenio.Enabled = False
        cmdImprimir.Enabled = False
    End If
            
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval

    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    'Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredito As ADODB.Recordset
                
    'Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
                
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsDCredito = oDCOMFormatosEval.RecuperarDatosFormatoConConvenio(psCtaCod) ' Recuperar Datos Basico
    Set rsAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(psCtaCod) 'Obtenemos Datos, Aceptable / Critico de los Ratios'CTI320200110 ERS003-2020. Agreg?
    
    lnColocCondi = rsDCredito!nColocCondicion ' para saber si el cliente es NUEVO
    fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses   'Si tiene evaluacion registrada 6 meses (LUCV20171115, agreg? seg?n correo: RUSI)
    
    If lnColocCondi = 4 Then
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.TabEnabled(2) = True
    End If
    
    '(3: Analista, 2: Coordinador, 1: JefeAgencia)
    fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
    fnFormato = pnFormato 'CTI320200110 ERS003-2020. Agreg?
    Call CargarFlexEdit 'CTI320200110 ERS003-2020. Agreg?
    
    If CargaControlesTipoPermiso(fnTipoPermiso) Then

        If fnTipoRegMant = 1 Then
            If Not (rsDCredito.EOF And rsDCredito.BOF) Then
                If (rsDCredito!cActiGiro) = "" Then
                    MsgBox ("Por favor, actualizar los datos del cliente. " & Chr(13) & "(Actividad o Giro del negocio)"), vbInformation, "Alerta"
                    Exit Function
                End If
                               
                fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
                fsCliente = Trim(rsDCredito!cPersNombre)
                fnAnio = rsDCredito!nAnio
                fnMes = rsDCredito!nMes
                fnMontoDeudaSbs = rsDCredito!nMontoUltimaDeudaSBS
                
                fdFechaDeudaSbs = rsDCredito!dFechaUltimaDeudaSBS
                
                fsInstConv = rsDCredito!cInstConv
                fnMontSolicitado = rsDCredito!nMonto
                fnCuota = rsDCredito!nCuotas
                fnExpCredito = rsDCredito!nExpoCred
                fdFechaActual = rsDCredito!dFechaActual
                
                ActXCodCta.NroCuenta = psCtaCod
                txtInfNegocioActividad.Text = fsGiroNego
                txtInfNegocioCliente.Text = fsCliente
                spnInfNegocioA?o.valor = fnAnio
                spnInfNegocioMes.valor = fnMes
                txtInfNegocioUltDeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
                
                'txtFecUltEndeuda2.Text = fdFechaDeudaSbs
                If fdFechaDeudaSbs = "01/01/1900" Then '26
                    txtFecUltEndeuda2.Text = "__/__/____"
                Else
                    txtFecUltEndeuda2.Text = fdFechaDeudaSbs
                End If
                
                txtInstConv = fsInstConv
                txtInfNegocioMontSolicitado.Text = Format(fnMontSolicitado, "#,##0.00")
                txtInfNegocioCuotas.Text = Format(fnCuota, "0#")
                txtInfNegocioExpCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
                txtInfNegocioFuenteIngreso.Text = Format(fdFechaActual, "dd/MM/yyyy")
                
                If rsDCredito!cT_Plani = "A" Then
                    optTipoPlanilla(2) = 2
                ElseIf rsDCredito!cT_Plani = "CA" Then
                    optTipoPlanilla(1) = 1
                ElseIf rsDCredito!cT_Plani = "C" Then
                    optTipoPlanilla(3) = 3
                End If
        
                cmdGuardarConConvenio.Visible = True
                cmdActualizarConConvenio.Visible = False
                
                'cmdImprimir.Enabled = False
                'cmdInformeVisitaConConvenio.Enabled = False
                
                Call Registro
             End If
        ElseIf fnTipoRegMant = 2 Then
        
            If fnTipoRegMant = 2 And Mantenimineto(IIf(fnTipoRegMant = 2, False, True)) = False Then
               MsgBox "No Cuenta con Registros", vbInformation, "Aviso"
               Exit Function
            End If
            
            cmdGuardarConConvenio.Visible = False
            cmdActualizarConConvenio.Visible = True
                           
            Call Registro
            
            If fnTipoInstitucion = 1 And fnSectorSalud = 0 Then
                lblMontoDispo.Visible = False
                lblMontoMaxIngDes.Visible = True
            ElseIf fnTipoInstitucion = 2 And fnSectorSalud = 0 Then
                lblMontoDispo.Visible = False
                lblMontoMaxIngDes.Visible = True
            ElseIf fnTipoInstitucion = 1 And fnSectorSalud = 1 Then
                lblMontoMaxIngDes.Visible = False
                lblMontoDispo.Visible = True
            End If
        
        ElseIf fnTipoRegMant = 3 Then
            Call Mantenimineto(IIf(fnTipoRegMant = 3, False, True))
            Call Consultar
                
            'Activar Boton InformeVisita y HojaEvaluacion
            If pnEstado = 2001 Or pnEstado = 2002 Then
                cmdInformeVisitaConConvenio.Enabled = True
                cmdImprimir.Enabled = True
            End If
            
            'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            gsOpeCod = gCredConsultarEvaluacionCred
            lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 8 - Consumo Con Convenio", fsCtaCod, gCodigoCuenta
            Set objPista = Nothing
            'Fin LUCV20181220
        End If
    Else
        Unload Me
        Exit Function
        'Me.Show 1
    End If

    'Para la Impresion -> LUCV Agrego
    fbGrabar = False
    If Not pbImprimir Then
        If fbImprimirVB Then
            Call cmdActualizarConConvenio_Click
        End If
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
    'Fin LUCV
    
End Function

Private Sub cmdAgregarConConvenio_Click()
    If feReferidosConConvenio.rows - 1 < 25 Then
        feReferidosConConvenio.lbEditarFlex = True
        feReferidosConConvenio.AdicionaFila
        feReferidosConConvenio.SetFocus
        feReferidosConConvenio.AvanceCeldas = Horizontal
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdCancelarConConvenio_Click()
    Unload Me
End Sub

Private Sub cmdActualizarConConvenio_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval    'CTI320200110 ERS003-2020. Agreg?
    Dim ActualizarDatos As Boolean
    Dim i As Integer
    Dim rsIngresos As ADODB.Recordset
    
If Validar Then
    Set rsIngresos = IIf(fgIngresos.rows - 1 > 0, fgIngresos.GetRsNew(), Nothing)
    gsOpeCod = gCredMantenimientoEvaluacionCred
    
    Set objPista = New COMManejador.Pista

        'If lnColocCondi = 1 Then 'LUCV2017115, Seg?n correo: RUSI
        If Not fbTieneReferido6Meses Then
            'Flex a Matriz Referidos **********->
                ReDim MatReferidos(feReferidosConConvenio.rows - 1, 6)
                For i = 1 To feReferidosConConvenio.rows - 1
                    MatReferidos(i, 0) = feReferidosConConvenio.TextMatrix(i, 0)
                    MatReferidos(i, 1) = feReferidosConConvenio.TextMatrix(i, 1)
                    MatReferidos(i, 2) = feReferidosConConvenio.TextMatrix(i, 2)
                    MatReferidos(i, 3) = feReferidosConConvenio.TextMatrix(i, 3)
                    MatReferidos(i, 4) = feReferidosConConvenio.TextMatrix(i, 4)
                    MatReferidos(i, 5) = feReferidosConConvenio.TextMatrix(i, 5)
                 Next i
        Else
                ReDim MatReferidos(0)
        End If
        'Fin Referidos
        
        Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing) 'CTI320200110 ERS003-2020. Agreg?
        If Not fbImprimirVB Then
            If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        End If
        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
        
                                                                            'IIf(txtFecUltEndeuda2.Text = "__/__/____", "01/01/1900", txtFecUltEndeuda2.Text)
        ActualizarDatos = oNCOMFormatosEval.ActualizarConsumoConConvenio_InfCliente(fsCtaCod, 8, txtInfNegocioActividad.Text, spnInfNegocioA?o.valor, spnInfNegocioMes.valor, _
                                                        txtInfNegocioUltDeuda.Text, IIf(txtFecUltEndeuda2.Text = "__/__/____", "01/01/1900", txtFecUltEndeuda2.Text), _
                                                        fnTipoAportacion, fnTipoInstitucion, fnSectorSalud, fnTipoPlanilla, txtInstConv.Text, txtInfNegocioMontSolicitado.Text, _
                                                        txtInfNegocioCuotas.Text, txtInfNegocioExpCredito.Text, CDate(txtInfNegocioFuenteIngreso.Text), _
                                                        (cmbFechaMes1.ItemData(cmbFechaMes1.ListIndex)), txtAnoMes1.Text, txtRemBrutaTotalMes1, _
                                                        txtDescuentoMes1, txtIngrNetoMes1, (cmbFechaMes2.ItemData(cmbFechaMes2.ListIndex)), txtAnoMes2.Text, _
                                                        txtRemBrutaTotalMes2, txtDescuentoMes2, txtIngrNetoMes2, txtRemBrutaTotalPromedio.Text, txtDescuentoPromedio.Text, _
                                                        txtIngNetolPromedio.Text, txtMontoMaxIngDescontarPromedio.Text, Replace(txtCapPagoConConvenio1.Text, "%", ""), _
                                                        txtCapPagoConConvenio2.Text, IIf(txtPropCreditoFechaVista.Text = "__/__/____", CDate(gdFecSis), _
                                                        txtPropCreditoFechaVista.Text), txtPropCreditoEntornoFamiliar.Text, txtPropCreditoGiroNegocio.Text, _
                                                        txtPropCreditoExpCrediticia.Text, txtPropCreditoFormNegocio.Text, txtPropCreditoColateralesGarantias.Text, _
                                                        txtPropCreditoDestino.Text, txtReferidosComentario.Text, MatReferidos, MtrRemuneracionBrutaTotal1, _
                                                        MtrDescuento1, MtrRemuneracionBrutaTotal2, MtrDescuento2, lnColocCondi, _
                                                        rsGastoFam, MatIfiGastoFami, MatIfiNoSupervisadaGastoFami, rsIngresos, pMtrNegocio)
                                                        'rsGastoFam, MatIfiGastoFami, MatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?
        
                                              
        If ActualizarDatos Then
            'CTI320200110 ERS003-2020. Agreg?
             Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval 'CTI320200110 ERS003-2020. Agreg?
             Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(fsCtaCod)
             Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod)
             Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(fsCtaCod)
            'Fin CTI320200110
         
            fbGrabar = True
            'LUCV20181220 Coment? y Agreg?, Anexo01 de Acta 199-2018
'                objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato Con Convenio", fsCtaCod, gCodigoCuenta
'                If fnTipoRegMant = 1 Then
'                    MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
'                Else
'                    MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
'                End If
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 8 - Consumo Con Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            Set objPista = Nothing 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            If Not fbImprimirVB Then
                MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
            End If
            Dim objCredito As COMDCredito.DCOMCredito
            Set objCredito = New COMDCredito.DCOMCredito
            Call objCredito.ActualizarEstadoxVB(ActXCodCta.NroCuenta, 1)
            
            'Fin LUCV20181220
            
            cmdActualizarConConvenio.Enabled = False
            cmdGuardarConConvenio.Visible = False
 
            If lnColocCondi <> 4 Then
                cmdInformeVisitaConConvenio.Enabled = True
            End If
            
            If (nEstado = 2001) Then
                cmdImprimir.Enabled = True
            End If
            
            'CTI320200110 ERS003-2020. Agreg?
            'Actualizacion de los Ratios
            txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"

            'Ratios: Aceptable / Critico ->*****
            If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                    Me.lblCapaAceptable.Caption = "Aceptable"
                    Me.lblCapaAceptable.ForeColor = &H8000&
                Else
                    Me.lblCapaAceptable.Caption = "Cr?tico"
                    Me.lblCapaAceptable.ForeColor = vbRed
                End If
                
            Else
                lblCapaAceptable.Visible = False
            End If
            'Fin Ratios <-****
            
            Set rsRatiosActual = Nothing
            Set rsRatiosAceptableCritico = Nothing
            'Fin CTI320200110
            If fbImprimirVB Then
                cmdActualizarConConvenio.Enabled = True
                fbImprimirVB = False
            End If
        Else
            MsgBox "Hubo errores al grabar la informaci?n", vbError, "Error"
        End If
End If
End Sub

Private Sub cmdGuardarConConvenio_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval    'CTI320200110 ERS003-2020. Agreg?
    Dim GrabarDatos As Boolean
    Dim i As Integer
    
    If Validar Then
    gsOpeCod = gCredRegistrarEvaluacionCred
    Set objPista = New COMManejador.Pista
    Dim rsIngresos As ADODB.Recordset
    'If lnColocCondi = 1 Then 'LUCV2017115, Seg?n correo: RUSI
    If Not fbTieneReferido6Meses Then
    'Flex a Matriz Referidos **********->
            ReDim MatReferidos(feReferidosConConvenio.rows - 1, 6)
            For i = 1 To feReferidosConConvenio.rows - 1
                MatReferidos(i, 0) = feReferidosConConvenio.TextMatrix(i, 0)
                MatReferidos(i, 1) = feReferidosConConvenio.TextMatrix(i, 1)
                MatReferidos(i, 2) = feReferidosConConvenio.TextMatrix(i, 2)
                MatReferidos(i, 3) = feReferidosConConvenio.TextMatrix(i, 3)
                MatReferidos(i, 4) = feReferidosConConvenio.TextMatrix(i, 4)
                MatReferidos(i, 5) = feReferidosConConvenio.TextMatrix(i, 5)
             Next i
     Else
        ReDim MatReferidos(0)
     End If
    'Fin Referidos
    Set rsIngresos = IIf(fgIngresos.rows - 1 > 0, fgIngresos.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing) 'CTI320200110 ERS003-2020. Agreg?
   
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        
                                                                                'IIf(txtFecUltEndeuda2.Text = "__/__/____", "01/01/1900", txtFecUltEndeuda2.Text)
        GrabarDatos = oNCOMFormatosEval.GuardarConsumoConConvenio_InfCliente(fsCtaCod, 8, txtInfNegocioActividad.Text, spnInfNegocioA?o.valor, spnInfNegocioMes.valor, _
                                            txtInfNegocioUltDeuda.Text, IIf(txtFecUltEndeuda2.Text = "__/__/____", "01/01/1900", txtFecUltEndeuda2.Text), _
                                            fnTipoAportacion, fnTipoInstitucion, fnSectorSalud, fnTipoPlanilla, txtInstConv.Text, txtInfNegocioMontSolicitado.Text, _
                                            txtInfNegocioCuotas.Text, txtInfNegocioExpCredito.Text, CDate(txtInfNegocioFuenteIngreso.Text), _
                                            (cmbFechaMes1.ItemData(cmbFechaMes1.ListIndex)), txtAnoMes1.Text, txtRemBrutaTotalMes1, txtDescuentoMes1, _
                                            txtIngrNetoMes1, (cmbFechaMes2.ItemData(cmbFechaMes2.ListIndex)), txtAnoMes2.Text, txtRemBrutaTotalMes2, _
                                            txtDescuentoMes2, txtIngrNetoMes2, txtRemBrutaTotalPromedio.Text, txtDescuentoPromedio.Text, _
                                            txtIngNetolPromedio.Text, txtMontoMaxIngDescontarPromedio.Text, Replace(txtCapPagoConConvenio1.Text, "%", ""), txtCapPagoConConvenio2.Text, _
                                            IIf(txtPropCreditoFechaVista.Text = "__/__/____", CDate(gdFecSis), txtPropCreditoFechaVista.Text), _
                                            txtPropCreditoEntornoFamiliar.Text, txtPropCreditoGiroNegocio.Text, txtPropCreditoExpCrediticia.Text, _
                                            txtPropCreditoFormNegocio.Text, txtPropCreditoColateralesGarantias.Text, txtPropCreditoDestino.Text, _
                                            txtReferidosComentario.Text, MatReferidos, MtrRemuneracionBrutaTotal1, MtrDescuento1, _
                                            MtrRemuneracionBrutaTotal2, MtrDescuento2, lnColocCondi, _
                                            rsGastoFam, MatIfiGastoFami, MatIfiNoSupervisadaGastoFami, rsIngresos, pMtrNegocio)
                                            
                                            'rsGastoFam, MatIfiGastoFami, MatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?
                                                      
            If GrabarDatos Then
                'CTI320200110 ERS003-2020. Agreg?
                Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval 'CTI320200110 ERS003-2020. Agreg?
                Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(fsCtaCod)
                Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod)
                Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(fsCtaCod)
                'Fin CTI320200110
                
                fbGrabar = True
                Dim oNCOMColocEval As New NCOMColocEval
                'RECO20161020 ERS060-2016 **********************************************************
                'Dim lcMovNro As String 'LUCV20181220 Coment?, Anexo01 de Acta 199-2018
                lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
                
                If Not ValidaExisteRegProceso(fsCtaCod, gTpoRegCtrlEvaluacion) Then
                   lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                   'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato Con Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
                   Call oNCOMColocEval.insEstadosExpediente(fsCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
                   Set oNCOMColocEval = Nothing
                End If
                'RECO FIN **************************************************************************
                'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato Con Convenio", fsCtaCod, gCodigoCuenta 'RECO20161020 ERS060-2016
                
                If fnTipoRegMant = 1 Then
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 8 - Consumo Con Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
                    MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
                Else
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 8 - Consumo Con Convenio", fsCtaCod, gCodigoCuenta 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
                    MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                End If
                
                cmdGuardarConConvenio.Enabled = False
                cmdActualizarConConvenio.Visible = False
                
                If lnColocCondi <> 4 Then
                    cmdInformeVisitaConConvenio.Enabled = True
                End If
                
                If (nEstado = 2001) Then
                    cmdImprimir.Enabled = True
                End If
                
                'CTI320200110 ERS003-2020. Agreg?
                'Actualizacion de los Ratios
                txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"

                'Ratios: Aceptable / Critico ->*****
                If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                    If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                        Me.lblCapaAceptable.Caption = "Aceptable"
                        Me.lblCapaAceptable.ForeColor = &H8000&
                    Else
                        Me.lblCapaAceptable.Caption = "Cr?tico"
                        Me.lblCapaAceptable.ForeColor = vbRed
                    End If
                    
                Else
                    lblCapaAceptable.Visible = False
                End If
                'Fin Ratios <-****
                
                Set rsRatiosActual = Nothing
                Set rsRatiosAceptableCritico = Nothing
                'Fin CTI320200110
            Else
                MsgBox "Hubo errores al grabar la informaci?n", vbError, "Error"
            End If
End If
End Sub

Private Sub cmdImprimir_Click()
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsImformeVisitaConsumoConConvenio As ADODB.Recordset
    Dim rsMostrarIngresos As ADODB.Recordset
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    Dim nTotalDescuento As Currency
    Dim nDescuentoLey As Currency
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsImformeVisitaConsumoConConvenio = New ADODB.Recordset
    
    'Set rsImformeVisitaConsumoConConvenio = oDCOMFormatosEval.MostrarDatosInformeVisitaFormatoConConvenio(fsCtaCod)
    Set rsImformeVisitaConsumoConConvenio = oDCOMFormatosEval.MostrarFormatoSinConvenioInfVisCabecera(fsCtaCod, 8)
    Set rsMostrarIngresos = oDCOMFormatosEval.MostrarIngresos(fsCtaCod, 8)
    Dim a As Integer
    Dim B As Integer
    Dim nFila As Integer
    Dim nFila1 As Integer
    Dim n As Currency
    a = 50
    B = 29

    'Creaci?n del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Cr?dito de Maynas S.A."
    oDoc.Subject = "Informe de Visita N? " & fsCtaCod
    oDoc.Title = "Informe de Visita N? " & fsCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoConvenio_HojaEvaluacion" & fsCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
If Not (rsImformeVisitaConsumoConConvenio.BOF Or rsImformeVisitaConsumoConConvenio.EOF) Then

    'Tama?o de hoja A4
    oDoc.NewPage A4_Vertical

    '---------- cabecera ---------------
    oDoc.WImage 45, 45, 45, 113, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(rsImformeVisitaConsumoConConvenio!cAgeDescripcion), "F2", 7.5, hLeft

    oDoc.WTextBox 40, 30, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight
    oDoc.WTextBox 60, 440, 10, 200, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
    oDoc.WTextBox 70, 440, 10, 200, "ANALISTA: " & Trim(rsImformeVisitaConsumoConConvenio!cUser), "F2", 7.5, hLeft
    
    oDoc.WTextBox 65, 100, 10, 400, "HOJA DE EVALUACION", "F2", 12, hCenter
    oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & fsCtaCod, "F2", 7.5, hLeft
    oDoc.WTextBox 90, 440, 10, 300, "MONEDA: " & IIf(Mid(fsCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersCod), "F2", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsImformeVisitaConsumoConConvenio!cPersNombre), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 440, 10, 200, "DNI: " & Trim(rsImformeVisitaConsumoConConvenio!cPersDni) & "   ", "F2", 7.5, hLeft
    oDoc.WTextBox 110, 440, 10, 200, "RUC: " & Trim(IIf(rsImformeVisitaConsumoConConvenio!cPersRuc = "-", Space(11), rsImformeVisitaConsumoConConvenio!cPersRuc)), "F2", 7.5, hLeft

    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 120, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 130, 55, 1, 160, "Evaluacion Mes 1", "F2", 7.5, hjustify
    oDoc.WTextBox 140, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = 140
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "Mes", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 80, 1, 160, cmbFechaMes1.Text, "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 10, 1, 160, "A?o", "F2", 7.5, hRight
    oDoc.WTextBox nFila, 40, 1, 160, txtAnoMes1.Text, "F2", 7.5, hRight
        
    oDoc.WTextBox 160, 55, 1, 160, "Remuneracion Bruta Total", "F1", 7.5, hjustify
    oDoc.WTextBox 160, 130, 1, 160, txtRemBrutaTotalMes1.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox 170, 55, 1, 160, "Descuento Total", "F1", 7.5, hjustify
    oDoc.WTextBox 170, 130, 1, 160, txtDescuentoMes1.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox 180, 55, 1, 160, "Ingreso Neto Total", "F1", 7.5, hjustify
    oDoc.WTextBox 180, 130, 1, 160, txtIngrNetoMes1.Text, "F1", 7.5, hRight
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 200, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 210, 55, 1, 190, "Detalle de Remuneracion Bruta Total", "F2", 7.5, hjustify
    oDoc.WTextBox 220, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        
    a = 0
    n = 0
    
If IsArray(MtrRemuneracionBrutaTotal1) Then
    For i = 1 To UBound(MtrRemuneracionBrutaTotal1, 2)
    oDoc.WTextBox 230 + a, 55, 1, 160, MtrRemuneracionBrutaTotal1(1, i), "F1", 7.5, hjustify
    oDoc.WTextBox 230 + a, 70, 1, 160, MtrRemuneracionBrutaTotal1(2, i), "F1", 7.5, hjustify
    oDoc.WTextBox 230 + a, 130, 1, 160, MtrRemuneracionBrutaTotal1(3, i), "F1", 7.5, hRight
    n = n + MtrRemuneracionBrutaTotal1(3, i)
    a = a + 10
    Next i
    oDoc.WTextBox 280, 80, 1, 160, "Total", "F2", 7.5, hRight
    oDoc.WTextBox 280, 130, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
End If
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 290, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 300, 55, 1, 160, "Detalle de Descuento", "F2", 7.5, hjustify
    oDoc.WTextBox 310, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        
    a = 0
    n = 0

If IsArray(MtrDescuento1) Then
    For i = 1 To UBound(MtrDescuento1, 2)
        oDoc.WTextBox 320 + a, 55, 1, 160, MtrDescuento1(0, i), "F1", 7.5, hjustify
        oDoc.WTextBox 320 + a, 80, 1, 160, MtrDescuento1(1, i), "F1", 7.5, hjustify
        oDoc.WTextBox 320 + a, 130, 1, 160, Format(MtrDescuento1(2, i), "#,##0.00"), "F1", 7.5, hRight
    
        If (MtrDescuento1(1, i) = "TOTAL DESCUENTOS") Then
            nTotalDescuento = Format(MtrDescuento1(2, i), "#,##0.00")
        End If
    
        If MtrDescuento1(1, i) = "DESCUENTO POR LEY" Then
            nDescuentoLey = Format(MtrDescuento1(2, i), "#,##0.00")
        End If
    
        a = a + 10
    Next i
    
    If UBound(MtrDescuento1, 2) > 0 Then
        n = n + MtrDescuento1(2, 2) - MtrDescuento1(2, 1) - MtrDescuento1(2, 3) - MtrDescuento1(2, 4) - MtrDescuento1(2, 5) - MtrDescuento1(2, 6)
        oDoc.WTextBox 380, 80, 1, 160, "Total", "F2", 7.5, hRight
        oDoc.WTextBox 380, 130, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
    End If
End If
     '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 390, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 400, 55, 1, 160, "Evaluacion Mes 2", "F2", 7.5, hjustify
    oDoc.WTextBox 410, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    
    oDoc.WTextBox 420, 55, 1, 160, "Mes", "F2", 7.5, hjustify
    oDoc.WTextBox 420, 80, 1, 160, cmbFechaMes2.Text, "F2", 7.5, hjustify
    
    oDoc.WTextBox 420, 10, 1, 160, "A?o", "F2", 7.5, hRight
    oDoc.WTextBox 420, 40, 1, 160, txtAnoMes2.Text, "F2", 7.5, hRight
        
    oDoc.WTextBox 430, 55, 1, 160, "Remuneracion Bruta Total", "F1", 7.5, hjustify
    oDoc.WTextBox 430, 130, 1, 160, txtRemBrutaTotalMes2.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox 440, 55, 1, 160, "Descuento Total", "F1", 7.5, hjustify
    oDoc.WTextBox 440, 130, 1, 160, txtDescuentoMes2.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox 450, 55, 1, 160, "Ingreso Neto Total", "F1", 7.5, hjustify
    oDoc.WTextBox 450, 130, 1, 160, txtIngrNetoMes2.Text, "F1", 7.5, hRight
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 470, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 480, 55, 1, 190, "Detalle de Remuneracion Bruta Total", "F2", 7.5, hjustify
    oDoc.WTextBox 490, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
       
    a = 0
    n = 0
    
If IsArray(MtrRemuneracionBrutaTotal2) Then
    For i = 1 To UBound(MtrRemuneracionBrutaTotal2, 2)
        oDoc.WTextBox 500 + a, 55, 1, 160, MtrRemuneracionBrutaTotal2(1, i), "F1", 7.5, hjustify
        oDoc.WTextBox 500 + a, 70, 1, 160, MtrRemuneracionBrutaTotal2(2, i), "F1", 7.5, hjustify
        oDoc.WTextBox 500 + a, 130, 1, 160, MtrRemuneracionBrutaTotal2(3, i), "F1", 7.5, hRight
        n = n + MtrRemuneracionBrutaTotal2(3, i)
        a = a + 10
    Next i
    oDoc.WTextBox 550, 80, 1, 160, "Total", "F2", 7.5, hRight
    oDoc.WTextBox 550, 130, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
End If
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 560, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 570, 55, 1, 160, "Detalle de Descuento", "F2", 7.5, hjustify
    oDoc.WTextBox 580, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        
    a = 0
    n = 0
If IsArray(MtrDescuento2) Then
    For i = 1 To UBound(MtrDescuento2, 2)
        oDoc.WTextBox 590 + a, 55, 1, 160, MtrDescuento2(0, i), "F1", 7.5, hjustify
        oDoc.WTextBox 590 + a, 80, 1, 160, MtrDescuento2(1, i), "F1", 7.5, hjustify
        oDoc.WTextBox 590 + a, 130, 1, 160, Format(MtrDescuento2(2, i), "#,##0.00"), "F1", 7.5, hRight
        a = a + 10
    Next i
    If UBound(MtrDescuento2, 2) > 0 Then
        n = n + MtrDescuento2(2, 2) - MtrDescuento2(2, 1) - MtrDescuento2(2, 3) - MtrDescuento2(2, 4) - MtrDescuento2(2, 5) - MtrDescuento2(2, 6)
        oDoc.WTextBox 650, 80, 1, 160, "Total", "F2", 7.5, hRight
        oDoc.WTextBox 650, 130, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
    End If
End If

    nFila = 660
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    a = 0
    For i = 1 To feGastosFamiliares.rows - 1
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 15, 250, feGastosFamiliares.TextMatrix(i, 2), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosFamiliares.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
        a = a + feGastosFamiliares.TextMatrix(i, 3)
    Next i
    nFila = nFila + 10
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
    
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10

 '--------------------------------------------------------------------------------------------------------------------------
    oDoc.NewPage A4_Vertical
    nFila = 35 + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox nFila + 10, 55, 1, 160, "Ingresos", "F2", 7.5, hjustify
    nFila = nFila + 20
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    a = 0
    n = 0
    oDoc.WTextBox nFila, 55, 1, 160, "ITEM", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 80, 1, 300, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    nFila = nFila + 10
    
    If Not (rsMostrarIngresos.BOF And rsMostrarIngresos.EOF) Then
        For i = 1 To rsMostrarIngresos.RecordCount
            If rsMostrarIngresos!nCodIngr <> 4 Then
                oDoc.WTextBox nFila, 55, 1, 160, rsMostrarIngresos!nCodIngr, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 80, 1, 300, rsMostrarIngresos!cConsDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarIngresos!nMonto, "#,##0.00"), "F1", 7.5, hRight
                nFila = nFila + 10
            End If
            a = a + 10
            n = n + rsMostrarIngresos!nMonto
            rsMostrarIngresos.MoveNext
        Next i
        nFila = nFila + 10
        oDoc.WTextBox nFila, 80, 1, 160, "Total", "F2", 7.5, hRight
        oDoc.WTextBox nFila, 140, 1, 160, Format(n, "#,##0.00"), "F2", 7.5, hRight
    End If
    nFila = nFila + 10
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox nFila + 10, 55, 1, 160, "Promedios", "F2", 7.5, hjustify
    oDoc.WTextBox nFila + 20, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft

    oDoc.WTextBox nFila + 30, 55, 1, 160, "Rem. Bruta Total", "F1", 7.5, hjustify
    oDoc.WTextBox nFila + 30, 130, 1, 160, txtRemBrutaTotalPromedio.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox nFila + 40, 55, 1, 160, "Descuentos", "F1", 7.5, hjustify
    oDoc.WTextBox nFila + 40, 130, 1, 160, txtDescuentoPromedio.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox nFila + 50, 55, 1, 160, "Ingreso Neto", "F1", 7.5, hjustify
    oDoc.WTextBox nFila + 50, 130, 1, 160, txtIngNetolPromedio.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox nFila + 60, 55, 1, 100, "Monto Maximo. Ingreso a Descontar", "F1", 7.5, hjustify
    oDoc.WTextBox nFila + 60, 130, 1, 160, txtMontoMaxIngDescontarPromedio.Text, "F1", 7.5, hRight
    
    oDoc.WTextBox nFila + 80, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox nFila + 90, 55, 1, 160, "Capacidad de Pago", "F2", 7.5, hjustify
    oDoc.WTextBox nFila + 100, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    
    oDoc.WTextBox nFila + 110, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
    'oDoc.WTextBox 770, 135, 1, 160, txtCapPagoConConvenio1.Text, "F1", 7.5, hRight CTI320200110 ERS003-2020. Coment?
    oDoc.WTextBox nFila + 110, 135, 1, 160, txtCapacidadNeta.Text, "F1", 7.5, hRight 'CTI320200110 ERS003-2020. Agreg?
    
    If fnTipoInstitucion = 1 And fnSectorSalud = 0 Then
        oDoc.WTextBox nFila + 110, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hjustify
    ElseIf fnTipoInstitucion = 2 And fnSectorSalud = 0 Then
        If rsImformeVisitaConsumoConConvenio!cTpIngr = "D" Then
            oDoc.WTextBox nFila + 110, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hjustify
        Else
            oDoc.WTextBox nFila + 110, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hRight
        End If
    ElseIf fnTipoInstitucion = 1 And fnSectorSalud = 1 Then
        oDoc.WTextBox nFila + 110, 330, 1, 160, "EN RELACION A SU INGRESO NETO", "F1", 7.5, hjustify
    End If
    
    Dim nCuotaLimite As Currency
    
    If CDbl(txtIngNetolPromedio.Text) <= 0 Then
        nCuotaLimite = 0
    Else
        'JOEP20210818 Mejora Limite de descuento
        nCuotaLimite = rsImformeVisitaConsumoConConvenio!nLimiteDescuento
        'JOEP20210818 Mejora Limite de descuento
        'nCuotaLimite = (nTotalDescuento + CDbl(rsImformeVisitaConsumoConConvenio!nMontoCuota) - nDescuentoLey) / CDbl(txtIngNetolPromedio.Text) 'Comento JOEP20210818 Mejora Limite de descuento
    End If
    '(total descuento + cuota - ingreso de ley)/ingreso neto
    oDoc.WTextBox nFila + 120, 55, 1, 160, "Cuota l?mite de descuento", "F1", 7.5, hjustify
    oDoc.WTextBox nFila + 120, 130, 1, 160, CStr(Round(nCuotaLimite * 100, 2)) & "%", "F1", 7.5, hRight
    '--------------------------------------------------------------------------------------------------------------------------
    
    oDoc.PDFClose
    oDoc.Show
Else
        MsgBox "Los Datos de Hoja de Evaluacion se mostrara despues de GRABAR la Sugerencia", vbInformation, "Aviso"
End If

End Sub

Private Sub cmdInformeVisitaConConvenio_Click()

    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(fsCtaCod)
    
    Me.cmdInformeVisitaConConvenio.Enabled = False
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atenci?n"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes
    Me.cmdInformeVisitaConConvenio.Enabled = True
End Sub

Private Sub cmdLlamaDescuentoMes1_Click()
    Dim psTotal As Double
    Dim nIngNeto As Currency
    
    If fnTipoInstitucion = 0 Then
        MsgBox "Ud. debe Seleccionar el Tipo de Institucion", vbInformation, "Aviso"
    Exit Sub
    End If
    
    If txtDescuentoMes1.Text = 0 Then
        Set MtrDescuento1 = Nothing
        
        frmCredFormEvalDescuento.Inicio1 psTotal, MtrDescuento1, 1, ActXCodCta.NroCuenta
        
        If psTotal <= 0 Then
            Set MtrDescuento1 = Nothing
        End If
        
     Else
        frmCredFormEvalDescuento.Inicio1 psTotal, MtrDescuento1, 1, ActXCodCta.NroCuenta
    End If
    
If IsArray(MtrDescuento1) Then
        
    txtDescuentoMes1.Text = Format(psTotal, "#,##0.00")
    nIngNeto = val(Replace(txtRemBrutaTotalMes1.Text, ",", "")) - MtrDescuento1(2, 1)
    
    If optTipoInstitucion(1).value = True Then
        txtIngrNetoMes1.Text = Format(nIngNeto, "#,##0.00")
    ElseIf optTipoInstitucion(2).value = True Then
        txtIngrNetoMes1.Text = Format(nIngNeto, "#,##0.00")
    End If
    
    Call CalculoTotal(3)
    Call CalculoTotal(4)
    Call CalculoTotal(5)
    Call CalculoTotal(6)
    Call CalculoTotal(7)
End If

End Sub

Private Sub cmdLlamaDescuentoMes2_Click()
    Dim psTotal As Double
    Dim nIngNeto As Currency
    If fnTipoInstitucion = 0 Then
        MsgBox "Ud. debe Seleccionar el Tipo de Institucion", vbInformation, "Aviso"
    Exit Sub
    End If
    
    If txtDescuentoMes2.Text = 0 Then
        Set MtrDescuento2 = Nothing
                
        frmCredFormEvalDescuento.Inicio2 psTotal, MtrDescuento2, 2, ActXCodCta.NroCuenta
        
        If psTotal <= 0 Then
         Set MtrDescuento2 = Nothing
        End If
        
     Else
        frmCredFormEvalDescuento.Inicio2 psTotal, MtrDescuento2, 2, ActXCodCta.NroCuenta
    End If
    
If IsArray(MtrDescuento2) Then

    txtDescuentoMes2.Text = Format(psTotal, "#,##0.00")
    nIngNeto = val(Replace(txtRemBrutaTotalMes2.Text, ",", "")) - MtrDescuento2(2, 1)
    
    If optTipoInstitucion(1).value = True Then
        txtIngrNetoMes2.Text = Format(nIngNeto, "#,##0.00")
    ElseIf optTipoInstitucion(2).value = True Then
        txtIngrNetoMes2.Text = Format(nIngNeto, "#,##0.00")
    End If
    
    Call CalculoTotal(3)
    Call CalculoTotal(4)
    Call CalculoTotal(6)
    Call CalculoTotal(5)
    Call CalculoTotal(7)
End If
End Sub

Private Sub cmdLlamaRemBrutaTotalMes1_Click()
    Dim psTotal As Double
    Dim psFilaPrimero As Double
    Dim nTpoInst As Integer
    
    If fnTipoInstitucion = 0 Then
        MsgBox "Ud. debe Seleccionar el Tipo de Institucion", vbInformation, "Aviso"
    Exit Sub
    End If
    
    nTpoInst = IIf(optTipoInstitucion(1).value = True, TipoInstitucion.nTpoPublico, TipoInstitucion.nTpoPrivado)
    
    If txtRemBrutaTotalMes1.Text = 0 Then
             
        Set MtrRemuneracionBrutaTotal1 = Nothing
        
        frmCredFormEvalRemuneracionBrutaTotal.Inicio1 psTotal, psFilaPrimero, MtrRemuneracionBrutaTotal1, 1, nTotalCompraDeu1, fsCtaCod, nTpoInst, IIf(ChkSectorSalud.value = 1, True, False)
    Else
        frmCredFormEvalRemuneracionBrutaTotal.Inicio1 psTotal, psFilaPrimero, MtrRemuneracionBrutaTotal1, 1, nTotalCompraDeu1, fsCtaCod, nTpoInst, IIf(ChkSectorSalud.value = 1, True, False)
    End If
    
   
    
    txtRemBrutaTotalMes1.Text = Format(psFilaPrimero, "#,##0.00")
    
    If Trim(txtDescuentoMes1.Text) = "" Then
        txtDescuentoMes1.Text = 0#
    End If
    
    If Trim(txtRemBrutaTotalMes1.Text) = "" Then
        txtRemBrutaTotalMes1.Text = 0#
    End If
    
    psTotal = val(Replace(txtRemBrutaTotalMes1.Text, ",", "")) - val(Replace(txtDescuentoMes1.Text, ",", ""))
    
'If optTipoInstitucion(2).value = True Then
    txtIngrNetoMes1.Text = Format(psTotal, "#,##0.00")
'End If

    Call CalculoTotal(3)
    Call CalculoTotal(4)
    Call CalculoTotal(5)
    Call CalculoTotal(6)
    Call CalculoTotal(7)
End Sub

Private Sub cmdLlamaRemBrutaTotalMes2_Click()

    Dim psTotal As Double
    Dim psFilaPrimero As Double
    Dim nTpoInst As Integer
    
    If fnTipoInstitucion = 0 Then
        MsgBox "Ud. debe Seleccionar el Tipo de Institucion", vbInformation, "Aviso"
    Exit Sub
    End If
    
    nTpoInst = IIf(optTipoInstitucion(1).value = True, TipoInstitucion.nTpoPublico, TipoInstitucion.nTpoPrivado)
    
    If txtRemBrutaTotalMes2.Text = 0 Then
       
        Set MtrRemuneracionBrutaTotal2 = Nothing
        
        frmCredFormEvalRemuneracionBrutaTotal.Inicio2 psTotal, psFilaPrimero, MtrRemuneracionBrutaTotal2, 2, nTotalCompraDeu2, fsCtaCod, nTpoInst, IIf(ChkSectorSalud.value = 1, True, False)
    Else
        frmCredFormEvalRemuneracionBrutaTotal.Inicio2 psTotal, psFilaPrimero, MtrRemuneracionBrutaTotal2, 2, nTotalCompraDeu2, fsCtaCod, nTpoInst, IIf(ChkSectorSalud.value = 1, True, False)
    End If
    
    txtRemBrutaTotalMes2.Text = Format(psFilaPrimero, "#,##0.00")
    
    If Trim(txtDescuentoMes2.Text) = "" Then
        txtDescuentoMes2.Text = 0#
    End If
    
    If Trim(txtRemBrutaTotalMes2.Text) = "" Then
        txtRemBrutaTotalMes2.Text = 0#
    End If
    
    psTotal = val(Replace(txtRemBrutaTotalMes2.Text, ",", "")) - val(Replace(txtDescuentoMes2.Text, ",", ""))

'    If optTipoInstitucion(2).value = True Then
        txtIngrNetoMes2.Text = Format(psTotal, "#,##0.00")
'    End If
    
    Call CalculoTotal(3)
    Call CalculoTotal(4)
    Call CalculoTotal(6)
    Call CalculoTotal(7)
    Call CalculoTotal(5)
    
End Sub

Private Sub CalculoTotal(ByVal pnTipo As Integer)

    Dim nTotalDescuento As Currency
    Dim nTotalDescuento1 As Currency
    Dim nTotalDescuento2 As Currency
    
    Select Case pnTipo
    
        'Promedio de Remuneracion Bruta Total
        Case 3:
                If txtRemBrutaTotalMes2.Text <> "0.00" Then
                txtRemBrutaTotalPromedio.Text = Format((CDbl(txtRemBrutaTotalMes1.Text) + CDbl(txtRemBrutaTotalMes2.Text)) / 2, "#,##0.00")
                End If
                
        'Promedio de Decuentos
        Case 4:
                If txtDescuentoMes2.Text <> "0.00" Then
                txtDescuentoPromedio.Text = Format((CDbl(txtDescuentoMes1.Text) + CDbl(txtDescuentoMes2.Text)) / 2, "#,##0.00")
                End If
                
        
        Case 5:
                'Publico al 50%
                'ElseIf fnTipoInstitucion = 1 Then
                If fnTipoInstitucion = 1 And ChkSectorSalud = 0 Then
                    If txtIngNetolPromedio.Text > 0 Then
                        txtMontoMaxIngDescontarPromedio.Text = Format(CDbl(txtIngNetolPromedio.Text) * 0.5, "#,##0.00")
                        txtCapPagoConConvenio1.Text = 50 & "%"
                        
                                If IsArray(MtrDescuento1) And IsArray(MtrDescuento2) Then
                                    If UBound(MtrDescuento1, 2) > 0 And UBound(MtrDescuento2, 2) > 0 Then
                                    txtCapPagoConConvenio2.Text = 0
                                        nTotalDescuento1 = (MtrDescuento1(2, 2) - MtrDescuento1(2, 1) - MtrDescuento1(2, 3) - MtrDescuento1(2, 4) - MtrDescuento1(2, 5) - MtrDescuento1(2, 6))
                                        nTotalDescuento2 = (MtrDescuento2(2, 2) - MtrDescuento2(2, 1) - MtrDescuento2(2, 3) - MtrDescuento2(2, 4) - MtrDescuento2(2, 5) - MtrDescuento2(2, 6))
                                        nTotalDescuento = (nTotalDescuento1 + nTotalDescuento2) / 2
                                        txtCapPagoConConvenio2.Text = Format(Replace(txtMontoMaxIngDescontarPromedio.Text, ",", "") - nTotalDescuento, "#,##0.00")
                                    End If
                                End If
                        
                        'txtCapPagoConConvenio2.Text = Format(txtMontoMaxIngDescontarPromedio.Text, "#,##0.00")
                    End If
                End If
                
                'Publico y Sector Salud al 50%
                If fnTipoInstitucion = 1 And ChkSectorSalud = 1 Then
                    'txtMontoMaxIngDescontarPromedio.Text = Format(CDbl(txtRemBrutaTotalPromedio.Text) * 0.5 - CDbl(txtDescuentoPromedio.Text), "#,##0.00") + nTotalCompraDeu1 + nTotalCompraDeu2
                    'txtCapPagoConConvenio1.Text = 50 & "%"
                    'txtCapPagoConConvenio2.Text = txtMontoMaxIngDescontarPromedio.Text
                    If txtIngNetolPromedio.Text <> "" Then
                        Call CalcularSectorSalud
                        txtCapPagoConConvenio1.Text = 50 & "%"
                        
                                If IsArray(MtrDescuento1) And IsArray(MtrDescuento2) Then
                                    If UBound(MtrDescuento1, 2) > 0 And UBound(MtrDescuento2, 2) > 0 Then
                                    txtCapPagoConConvenio2.Text = 0
                                        nTotalDescuento1 = (MtrDescuento1(2, 2) - MtrDescuento1(2, 1) - MtrDescuento1(2, 3) - MtrDescuento1(2, 4) - MtrDescuento1(2, 5) - MtrDescuento1(2, 6))
                                        nTotalDescuento2 = (MtrDescuento2(2, 2) - MtrDescuento2(2, 1) - MtrDescuento2(2, 3) - MtrDescuento2(2, 4) - MtrDescuento2(2, 5) - MtrDescuento2(2, 6))
                                        nTotalDescuento = (nTotalDescuento1 + nTotalDescuento2) / 2
                                        txtCapPagoConConvenio2.Text = Format(Replace(txtMontoMaxIngDescontarPromedio.Text, ",", "") - nTotalDescuento - txtDescuentoPromedio.Text, "#,##0.00")
                                    End If
                                End If
                        
                        'txtCapPagoConConvenio2.Text = Format(txtMontoMaxIngDescontarPromedio.Text - nTotalDescuento, "#,##0.00")
                    End If
                End If
        'Promedio del Ingreso Neto
        Case 6:
            If txtIngrNetoMes2.Text > 0 Then
                    txtIngNetolPromedio.Text = Format((CDbl(txtIngrNetoMes1.Text) + CDbl(txtIngrNetoMes2.Text)) / 2, "#,##0.00")
            End If
            
        'Privado
        Case 7:
                If fnTipoInstitucion = 2 Then '2
                    If txtIngNetolPromedio.Text >= 0 And txtIngNetolPromedio.Text <= 1000 Then
                        txtMontoMaxIngDescontarPromedio.Text = Format(txtIngNetolPromedio.Text * 0.35, "#,##0.00")
                        
                        txtCapPagoConConvenio1.Text = 35 & "%"
                        txtCapPagoConConvenio2.Text = txtMontoMaxIngDescontarPromedio.Text
                    
                    ElseIf txtIngNetolPromedio.Text >= 1000.01 And txtIngNetolPromedio.Text <= 2000 Then
                        txtMontoMaxIngDescontarPromedio.Text = Format(txtIngNetolPromedio.Text * 0.45, "#,##0.00")
                        
                        txtCapPagoConConvenio1.Text = 45 & "%"
                        txtCapPagoConConvenio2.Text = txtMontoMaxIngDescontarPromedio.Text
                    
                    ElseIf txtIngNetolPromedio.Text >= 2000.01 Then
                        txtMontoMaxIngDescontarPromedio.Text = Format(txtIngNetolPromedio.Text * 0.5, "#,##0.00")
                        
                        txtCapPagoConConvenio1.Text = 50 & "%"
                        txtCapPagoConConvenio2.Text = txtMontoMaxIngDescontarPromedio.Text
                    
                    End If
                End If
    End Select
    
'On Error GoTo ErrorCalculo

'If optTipoInstitucion(2).value = False Then ' no entra si es Privada
'    If IsArray(MtrDescuento1) And IsArray(MtrDescuento2) Then
'        If UBound(MtrDescuento1, 2) > 0 And UBound(MtrDescuento2, 2) > 0 Then
'        txtCapPagoConConvenio2.Text = 0
'            nTotalDescuento1 = (MtrDescuento1(2, 2) - MtrDescuento1(2, 1) - MtrDescuento1(2, 3) - MtrDescuento1(2, 4) - MtrDescuento1(2, 5) - MtrDescuento1(2, 6))
'            nTotalDescuento2 = (MtrDescuento2(2, 2) - MtrDescuento2(2, 1) - MtrDescuento2(2, 3) - MtrDescuento2(2, 4) - MtrDescuento2(2, 5) - MtrDescuento2(2, 6))
'            nTotalDescuento = (nTotalDescuento1 + nTotalDescuento2) / 2
'            txtCapPagoConConvenio2.Text = Format(Replace(txtMontoMaxIngDescontarPromedio.Text, ",", "") - nTotalDescuento, "#,##0.00")
'        End If
'    End If
'End If

'ErrorCalculo:

 '   If optTipoInstitucion(1).value = True And ChkSectorSalud.value = 1 Then ' solo entra si es Publico y Sector Salud
'        If txtIngNetolPromedio.Text <> "" Then
'            Call CalcularSectorSalud
'            txtCapPagoConConvenio1.Text = 50 & "%"
'            txtCapPagoConConvenio2.Text = Format(txtMontoMaxIngDescontarPromedio.Text - nTotalDescuento, "#,##0.00")
'        End If
 '   End If
    
    Exit Sub
    
End Sub

Private Sub cmdQuitarConConvenio_Click()
    If MsgBox("Esta Seguro de Eliminar  a " & feReferidosConConvenio.TextMatrix(feReferidosConConvenio.row, 1) & " del Registro?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    feReferidosConConvenio.EliminaFila (feReferidosConConvenio.row)
End Sub

'CTI320200110 ERS003-2020, Agreg?
Private Sub feGastosFamiliares_Click()
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
        Or CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodDeudaLCNUGastoFami Then
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Me.feGastosFamiliares.ForeColorRow vbBlack, True
    Else
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End If
End Sub

Private Sub feGastosFamiliares_EnterCell()
    If feGastosFamiliares.Col = 3 Or (feGastosFamiliares.Col = 3 And feGastosFamiliares.row = 1) Then
            feGastosFamiliares.AvanceCeldas = Vertical
    Else
            feGastosFamiliares.AvanceCeldas = Horizontal
    End If
        
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
            Or (CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami) Then
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
        
    If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodDeudaLCNUGastoFami Then
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
    Else
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End If
End Sub

Private Sub feGastosFamiliares_RowColChange()
    If feGastosFamiliares.Col = 3 Or (feGastosFamiliares.Col = 3 And feGastosFamiliares.row = 1) Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.Col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami _
        Or (CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiNoSupervisadaGastoFami) Then 'CTI320200110 ERS003-2020, Agreg?
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
        feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodDeudaLCNUGastoFami Then
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
       
    Else
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End If
End Sub

Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String)
    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
     If CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then 'CTI320200110 ERS003-2020. Agreg?
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiNoSupervisadaGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                             gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                            gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    End If
End Sub

Private Sub feGastosFamiliares_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosFamiliares.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosFamiliares.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
    End If

    If (Me.feGastosFamiliares.Col = 3 And Me.feGastosFamiliares.row = 4) Then
        SSTab1.Tab = 2
        SendKeys "{TAB}"
    End If
End Sub
'Fin CTI320200110
Private Sub feReferidosConConvenio_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case pnCol
    Case 2
        If IsNumeric(feReferidosConConvenio.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidosConConvenio.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidosConConvenio.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                    feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
            feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 3
        If IsNumeric(feReferidosConConvenio.TextMatrix(pnRow, pnCol)) Then

        Else
            MsgBox "Telefono Mal Ingresado", vbInformation, "Alerta"
            feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 5
        If IsNumeric(feReferidosConConvenio.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidosConConvenio.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidosConConvenio.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                    feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
                feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "DNI Mal Ingresado", vbInformation, "Alerta"
            feReferidosConConvenio.TextMatrix(pnRow, pnCol) = 0
        End If
    End Select
End Sub

Private Sub feReferidosConConvenio_RowColChange()
If feReferidosConConvenio.Col = 1 Then
    feReferidosConConvenio.MaxLength = "200"
ElseIf feReferidosConConvenio.Col = 2 Then
    feReferidosConConvenio.MaxLength = "8"
ElseIf feReferidosConConvenio.Col = 3 Then
    feReferidosConConvenio.MaxLength = "9"
ElseIf feReferidosConConvenio.Col = 4 Then
    feReferidosConConvenio.MaxLength = "200"
ElseIf feReferidosConConvenio.Col = 5 Then
    feReferidosConConvenio.MaxLength = "8"
End If
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
CentraForm Me
SSTab1.TabVisible(1) = False
    cmdActualizarConConvenio.Visible = False
    ChkSectorSalud.Visible = False
        Call MostraComboFechasMes1
        Call MostraComboFechasMes2
        Call ControlText
    nTotalCompraDeu1 = 0
    nTotalCompraDeu2 = 0
    lblMontoDispo.Visible = False
End Sub

Private Sub ControlText()
    txtIngrNetoMes1.Text = "0.00"
    txtRemBrutaTotalMes1.Text = "0.00"
    txtDescuentoMes1.Text = "0.00"

    txtIngrNetoMes2.Text = "0.00"
    txtRemBrutaTotalMes2.Text = "0.00"
    txtDescuentoMes2.Text = "0.00"

    txtRemBrutaTotalPromedio.Text = "0.00"
    
    txtDescuentoPromedio.Text = "0.00"
    
    txtIngNetolPromedio.Text = "0.00"
    
    txtMontoMaxIngDescontarPromedio.Text = "0.00"
    txtCapPagoConConvenio1.Text = "00"
    txtCapPagoConConvenio2.Text = "0.00"
    
    nTotalCompraDeu1 = 0
    nTotalCompraDeu2 = 0
    
    Set MtrDescuento1 = Nothing
    Set MtrDescuento2 = Nothing
    
    Set MtrRemuneracionBrutaTotal1 = Nothing
    Set MtrRemuneracionBrutaTotal2 = Nothing
    
End Sub

Public Sub MostraComboFechasMes1()

Dim oComboFecha As COMDCredito.DCOMFormatosEval
Dim rsComboFecha As ADODB.Recordset

Set oComboFecha = New COMDCredito.DCOMFormatosEval
Set rsComboFecha = oComboFecha.MostrarComboFecha()

CargarComboBox rsComboFecha, cmbFechaMes1

'Para guardar Dato = cmbFechaMes1.ItemData(cmbFechaMes1.ListIndex))

End Sub

Public Sub MostraComboFechasMes2()

Dim oComboFecha As COMDCredito.DCOMFormatosEval
Dim rsComboFecha As ADODB.Recordset

Set oComboFecha = New COMDCredito.DCOMFormatosEval
Set rsComboFecha = oComboFecha.MostrarComboFecha()

CargarComboBox rsComboFecha, cmbFechaMes2
'Para guardar Dato = cmbFechaMes2.ItemData(cmbFechaMes2.ListIndex))

End Sub

Private Sub ChkSectorSalud_Click()

If ChkSectorSalud = vbChecked Then
    
    lblMontoMaxIngDes.Visible = False
    
    ChkSectorSalud.value = 1
    
    Call ControlText
    
    lblMontoDispo.Visible = True

ElseIf ChkSectorSalud = vbUnchecked Then
    
    lblMontoDispo.Visible = False
    
    ChkSectorSalud.value = 0
    
    Call ControlText
    
    lblMontoMaxIngDes.Visible = True

End If
    fnSectorSalud = ChkSectorSalud.value
End Sub



Private Sub optTipoAportacion_Click(Index As Integer)
    'Tipo Aportacion
    '1: AFP ; 2: ONP
    fnTipoAportacion = Index
End Sub

Private Sub optTipoInstitucion_Click(Index As Integer)
    fnTipoInstitucion = Index
    
If fnTipoInstitucion = 1 Then

    lblMontoDispo.Visible = False
    
    Call ControlText
    
    cmdLlamaDescuentoMes1.Enabled = True
    cmdLlamaDescuentoMes2.Enabled = True
    
    ChkSectorSalud.Visible = True
    
    If ChkSectorSalud = vbChecked Then
        ChkSectorSalud = vbUnchecked
    End If
    
    fnSectorSalud = 0
    
    lblMontoMaxIngDes.Visible = True
    
ElseIf fnTipoInstitucion = 2 Then

    lblMontoDispo.Visible = False
    
    Call ControlText

    ChkSectorSalud.Visible = False
    
    cmdLlamaDescuentoMes1.Enabled = False
    cmdLlamaDescuentoMes2.Enabled = False
    
    fnSectorSalud = 0
    
    lblMontoMaxIngDes.Visible = True
    
End If

End Sub

Private Sub optTipoPlanilla_Click(Index As Integer)
    'Tipo Aportacion
    '1: CAS ; 2: Activos ; 3: Cesantes
    fnTipoPlanilla = Index
End Sub

Public Function Mantenimineto(ByVal pbMantenimiento As Boolean) As Boolean
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsMantenimientoConConvenio As ADODB.Recordset
    Dim rsMantenimientoSinConvenioPropuestaCredito As ADODB.Recordset
    Dim rsMantenimientoConConvenioEvalMeses As ADODB.Recordset
    Dim rsMantenimientoConConvenioPromedios As ADODB.Recordset
    Dim rsMantenimientoConConvenioReferidos As ADODB.Recordset
    Dim rsMantenimientoConConvenioRemBrutaTotal_1 As ADODB.Recordset
    Dim rsMantenimientoConConvenioRemBrutaTotal_2 As ADODB.Recordset
    Dim rsMantenimientoConConvenioDescuento_1 As ADODB.Recordset
    Dim rsMantenimientoConConvenioDescuento_2 As ADODB.Recordset
    Dim rsMantenimientoSinConvenioIngresos As ADODB.Recordset
    Dim rsMantenimientoSinConvenioGastoNegocio As ADODB.Recordset
    
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsMantenimientoConConvenio = New ADODB.Recordset
    Set rsMantenimientoSinConvenioPropuestaCredito = New ADODB.Recordset
    Set rsMantenimientoConConvenioEvalMeses = New ADODB.Recordset
    Set rsMantenimientoConConvenioPromedios = New ADODB.Recordset
    Set rsMantenimientoConConvenioReferidos = New ADODB.Recordset
    Set rsMantenimientoConConvenioRemBrutaTotal_1 = New ADODB.Recordset
    Set rsMantenimientoConConvenioRemBrutaTotal_2 = New ADODB.Recordset
    Set rsMantenimientoConConvenioDescuento_1 = New ADODB.Recordset
    Set rsMantenimientoConConvenioDescuento_2 = New ADODB.Recordset
    Set rsMantenimientoSinConvenioIngresos = New ADODB.Recordset
    Set rsMantenimientoSinConvenioGastoNegocio = New ADODB.Recordset
    
    Set rsMantenimientoConConvenio = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenio(fsCtaCod)
    Set rsMantenimientoSinConvenioPropuestaCredito = oDCOMFormatosEval.RecuperarConsumoSinConvenioPropuestaCredito(fsCtaCod, 8)
    Set rsMantenimientoConConvenioEvalMeses = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioEvalMeses(fsCtaCod)
    Set rsMantenimientoConConvenioPromedios = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioPromedios(fsCtaCod)
    Set rsMantenimientoConConvenioReferidos = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioReferidos(fsCtaCod)
    Set rsMantenimientoConConvenioRemBrutaTotal_1 = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioRemuBrutaTota_1(fsCtaCod)
    Set rsMantenimientoConConvenioRemBrutaTotal_2 = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioRemuBrutaTota_2(fsCtaCod)
    Set rsMantenimientoConConvenioDescuento_1 = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioDescuento_1(fsCtaCod)
    Set rsMantenimientoConConvenioDescuento_2 = oDCOMFormatosEval.RecuperarDatosTotalConsumoConConvenioDescuento_2(fsCtaCod)
    Set rsMantenimientoSinConvenioIngresos = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioIngresos(fsCtaCod)
    Set rsMantenimientoSinConvenioGastoNegocio = oDCOMFormatosEval.RecuperarDatosTotalConsumoSinConvenioGastoNegocio(fsCtaCod)

    
    'CTI320200110 ERS003-2020. Agreg?:
    Set rsDatGastoFam = oDCOMFormatosEval.RecuperaDatosCredEvalGastosFam(fsCtaCod)
    pnMontoOtrasIfis = 0
    Set rsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, fnFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami)
    If Not (rsDatIfiGastoFami.BOF Or rsDatIfiGastoFami.EOF) Then
     For i = 1 To rsDatIfiGastoFami.RecordCount
        pnMontoOtrasIfis = pnMontoOtrasIfis + rsDatIfiGastoFami!nMonto
        rsDatIfiGastoFami.MoveNext
     Next i
     rsDatIfiGastoFami.MoveFirst
    End If
   
    
    
    Set rsDatIfiNoSupervisadaGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, fnFormato, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
    Set rsDatRatios = oDCOMFormatosEval.RecuperaDatosRatios(fsCtaCod)
    'Fin CTI320200110
    
    If Not (rsMantenimientoConConvenio.BOF And rsMantenimientoConConvenio.EOF) Then
        
        ActXCodCta.NroCuenta = rsMantenimientoConConvenio!cCtaCod
        txtInfNegocioActividad.Text = rsMantenimientoConConvenio!cActividad
        txtInfNegocioCliente.Text = rsMantenimientoConConvenio!cPersNombre
        spnInfNegocioA?o.valor = rsMantenimientoConConvenio!nAntgAnios
        spnInfNegocioMes.valor = rsMantenimientoConConvenio!nAntgMes
        txtInfNegocioUltDeuda.Text = Format(rsMantenimientoConConvenio!nUltEndeSBS, "#,##0.00")
        optTipoAportacion(rsMantenimientoConConvenio!nTipoAportacion).value = 1
                
        'txtFecUltEndeuda2.Text = rsMantenimientoConConvenio!dUltEndeuSBS
        
        If rsMantenimientoConConvenio!dUltEndeuSBS = "01/01/1900" Then
            txtFecUltEndeuda2.Text = "__/__/____"
        Else
            txtFecUltEndeuda2.Text = rsMantenimientoConConvenio!dUltEndeuSBS
        End If
        
        optTipoInstitucion(rsMantenimientoConConvenio!nTipoInstitucion).value = 1
        ChkSectorSalud.value = rsMantenimientoConConvenio!nSectorSalud
        optTipoPlanilla(rsMantenimientoConConvenio!ntipoPlanilla).value = 1
        txtInstConv.Text = rsMantenimientoConConvenio!cinstitConvenio
        txtInfNegocioMontSolicitado.Text = Format(rsMantenimientoConConvenio!nMontoSol, "#,##0.00")
        txtInfNegocioCuotas.Text = rsMantenimientoConConvenio!nNumCuotas
        txtInfNegocioExpCredito.Text = Format(rsMantenimientoConConvenio!nExposiCred, "#,##0.00")
        txtInfNegocioFuenteIngreso.Text = Format(rsMantenimientoConConvenio!dFecEval, "dd/mm/yyyy")
                        
        txtReferidosComentario.Text = rsMantenimientoConConvenio!cComentario
        Mantenimineto = True
    End If
    
    If lnColocCondi <> 4 Then
        If Not (rsMantenimientoSinConvenioPropuestaCredito.BOF And rsMantenimientoSinConvenioPropuestaCredito.EOF) Then
            txtPropCreditoFechaVista.Text = rsMantenimientoSinConvenioPropuestaCredito!dFecVisita
            txtPropCreditoEntornoFamiliar.Text = rsMantenimientoSinConvenioPropuestaCredito!cEntornoFami
            txtPropCreditoGiroNegocio.Text = rsMantenimientoSinConvenioPropuestaCredito!cGiroUbica
            txtPropCreditoExpCrediticia.Text = rsMantenimientoSinConvenioPropuestaCredito!cExpeCrediticia
            txtPropCreditoFormNegocio.Text = rsMantenimientoSinConvenioPropuestaCredito!cFormalNegocio
            txtPropCreditoColateralesGarantias.Text = rsMantenimientoSinConvenioPropuestaCredito!cColateGarantia
            txtPropCreditoDestino.Text = rsMantenimientoSinConvenioPropuestaCredito!cDestino
        Mantenimineto = True
        End If
    End If
    
    'Evaluacion Meses
      If Not (rsMantenimientoConConvenioEvalMeses.BOF And rsMantenimientoConConvenioEvalMeses.EOF) Then
        For i = 1 To rsMantenimientoConConvenioEvalMeses.RecordCount
            If rsMantenimientoConConvenioEvalMeses!nEvalMes = 1 Then
                'Mes 1
                cmbFechaMes1.ListIndex = (rsMantenimientoConConvenioEvalMeses!nMes) - 1
                txtAnoMes1 = rsMantenimientoConConvenioEvalMeses!nAnio
                txtRemBrutaTotalMes1 = Format(rsMantenimientoConConvenioEvalMeses!nRemBrutaTotal, "#,##0.00")
                txtDescuentoMes1 = Format(rsMantenimientoConConvenioEvalMeses!nDescuento, "#,##0.00")
                txtIngrNetoMes1 = Format(rsMantenimientoConConvenioEvalMeses!nIngNeto, "#,##0.00")
            Else
                'Mes 2
                cmbFechaMes2.ListIndex = (rsMantenimientoConConvenioEvalMeses!nMes) - 1
                txtAnoMes2 = rsMantenimientoConConvenioEvalMeses!nAnio
                txtRemBrutaTotalMes2 = Format(rsMantenimientoConConvenioEvalMeses!nRemBrutaTotal, "#,##0.00")
                txtDescuentoMes2 = Format(rsMantenimientoConConvenioEvalMeses!nDescuento, "#,##0.00")
                txtIngrNetoMes2 = Format(rsMantenimientoConConvenioEvalMeses!nIngNeto, "#,##0.00")
            End If
            rsMantenimientoConConvenioEvalMeses.MoveNext
        Next i
      Mantenimineto = True
    End If
    
    
    'Promedios
      If Not (rsMantenimientoConConvenioPromedios.BOF And rsMantenimientoConConvenioPromedios.EOF) Then
                  
                txtRemBrutaTotalPromedio.Text = Format(rsMantenimientoConConvenioPromedios!nRemBrutaTotal, "#,##0.00")
                txtDescuentoPromedio.Text = Format(rsMantenimientoConConvenioPromedios!nDescuento, "#,##0.00")
                txtIngNetolPromedio.Text = Format(rsMantenimientoConConvenioPromedios!nIngNeto, "#,##0.00")
                txtMontoMaxIngDescontarPromedio.Text = Format(rsMantenimientoConConvenioPromedios!nMontoMax, "#,##0.00")
                
                txtCapPagoConConvenio1.Text = rsMantenimientoConConvenioPromedios!nCapPagoPorc & "%"
                txtCapPagoConConvenio2.Text = Format(rsMantenimientoConConvenioPromedios!nCapPagoTotal, "#,##0.00")
                       
      Mantenimineto = True
    End If
    
    'Referidos
      If Not (rsMantenimientoConConvenioReferidos.EOF And rsMantenimientoConConvenioReferidos.BOF) Then
    Do While Not rsMantenimientoConConvenioReferidos.EOF
        feReferidosConConvenio.AdicionaFila
        lnFila = feReferidosConConvenio.row
        
        feReferidosConConvenio.TextMatrix(lnFila, 1) = rsMantenimientoConConvenioReferidos!cNombre
        feReferidosConConvenio.TextMatrix(lnFila, 2) = rsMantenimientoConConvenioReferidos!cDniNom
        feReferidosConConvenio.TextMatrix(lnFila, 3) = rsMantenimientoConConvenioReferidos!cTelf
        feReferidosConConvenio.TextMatrix(lnFila, 4) = rsMantenimientoConConvenioReferidos!cReferido
        feReferidosConConvenio.TextMatrix(lnFila, 5) = rsMantenimientoConConvenioReferidos!cDNIRef
                
        rsMantenimientoConConvenioReferidos.MoveNext
    Loop
        rsMantenimientoConConvenioReferidos.Close
        Set rsMantenimientoConConvenioReferidos = Nothing
    End If
    
    'Matriz Remuneracion Bruta Total 1
    If Not (rsMantenimientoConConvenioRemBrutaTotal_1.EOF And rsMantenimientoConConvenioRemBrutaTotal_1.BOF) Then
          ReDim MtrRemuneracionBrutaTotal1(3, 0)
             For i = 1 To (rsMantenimientoConConvenioRemBrutaTotal_1.RecordCount)
          ReDim Preserve MtrRemuneracionBrutaTotal1(3, i)
             MtrRemuneracionBrutaTotal1(1, i) = rsMantenimientoConConvenioRemBrutaTotal_1!nCodRemBruTot
             MtrRemuneracionBrutaTotal1(2, i) = rsMantenimientoConConvenioRemBrutaTotal_1!CDescripcion
             MtrRemuneracionBrutaTotal1(3, i) = Format(rsMantenimientoConConvenioRemBrutaTotal_1!nMonto, "#,##0.00")
             rsMantenimientoConConvenioRemBrutaTotal_1.MoveNext
             Next i
    End If
    
    'Matriz Remuneracion Bruta Total 2
    If Not (rsMantenimientoConConvenioRemBrutaTotal_2.EOF And rsMantenimientoConConvenioRemBrutaTotal_2.BOF) Then
          ReDim MtrRemuneracionBrutaTotal2(3, 0)
             For i = 1 To (rsMantenimientoConConvenioRemBrutaTotal_2.RecordCount)
          ReDim Preserve MtrRemuneracionBrutaTotal2(3, i)
             MtrRemuneracionBrutaTotal2(1, i) = rsMantenimientoConConvenioRemBrutaTotal_2!nCodRemBruTot
             MtrRemuneracionBrutaTotal2(2, i) = rsMantenimientoConConvenioRemBrutaTotal_2!CDescripcion
             MtrRemuneracionBrutaTotal2(3, i) = Format(rsMantenimientoConConvenioRemBrutaTotal_2!nMonto, "#,##0.00")
             rsMantenimientoConConvenioRemBrutaTotal_2.MoveNext
             Next i
    End If
    
    'Matriz Descuento Total 1
    ReDim MtrDescuento1(3, 0)
        For i = 1 To (rsMantenimientoConConvenioDescuento_1.RecordCount)
            ReDim Preserve MtrDescuento1(3, i)
              MtrDescuento1(0, i) = rsMantenimientoConConvenioDescuento_1!nCodDesc
              MtrDescuento1(1, i) = rsMantenimientoConConvenioDescuento_1!CDescripcion
              MtrDescuento1(2, i) = Format(rsMantenimientoConConvenioDescuento_1!nMonto, "#,##0.00")
              rsMantenimientoConConvenioDescuento_1.MoveNext
        Next i

    'Matriz Descuento Total 2
    ReDim MtrDescuento2(3, 0)
        For i = 1 To (rsMantenimientoConConvenioDescuento_2.RecordCount)
            ReDim Preserve MtrDescuento2(3, i)
              MtrDescuento2(0, i) = rsMantenimientoConConvenioDescuento_2!nCodDesc
              MtrDescuento2(1, i) = rsMantenimientoConConvenioDescuento_2!CDescripcion
              MtrDescuento2(2, i) = Format(rsMantenimientoConConvenioDescuento_2!nMonto, "#,##0.00")
              rsMantenimientoConConvenioDescuento_2.MoveNext
        Next i


    'CTI320200110 ERS003-2020. Agreg?:
        'Call FormatearGrillas(feGastosFamiliares)
        Call LimpiaFlex(feGastosFamiliares)
            Do While Not rsDatGastoFam.EOF
                feGastosFamiliares.AdicionaFila
                lnFila = feGastosFamiliares.row
                feGastosFamiliares.TextMatrix(lnFila, 1) = rsDatGastoFam!nConsValor
                feGastosFamiliares.TextMatrix(lnFila, 2) = rsDatGastoFam!cConsDescripcion
                feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsDatGastoFam!nMonto, "#,##0.00")
                If rsDatGastoFam!nConsValor = 5 Then
                    feGastosFamiliares.TextMatrix(lnFila, 3) = Format(pnMontoOtrasIfis, "#,##0.00")
                End If
                 Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
                    Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?: gCodCuotaIfiNoSupervisadaGastoFami
                        'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
                        Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                        Me.feGastosFamiliares.ForeColorRow vbBlack, True
                    Case gCodDeudaLCNUGastoFami, gCodCuotaCmac
                        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                        Me.feGastosFamiliares.ForeColorRow vbBlack, True
                    Case Else
                        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                 End Select
                rsDatGastoFam.MoveNext
            Loop
        rsDatGastoFam.Close
        Set rsDatGastoFam = Nothing
        
        'Carga de rsDatIfiGastoFami -> Matrix
        ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
        j = 0
        Do While Not rsDatIfiGastoFami.EOF
            MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
            MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
            MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#0.00")
            rsDatIfiGastoFami.MoveNext
        j = j + 1
        Loop
        rsDatIfiGastoFami.Close
        Set rsDatIfiGastoFami = Nothing
        
        'Carga de rsDatIfiNoSupervisadaGastoFami -> Matrix
        ReDim MatIfiNoSupervisadaGastoFami(rsDatIfiNoSupervisadaGastoFami.RecordCount, 4)
        j = 0
        Do While Not rsDatIfiNoSupervisadaGastoFami.EOF
            MatIfiNoSupervisadaGastoFami(j, 0) = rsDatIfiNoSupervisadaGastoFami!nNroCuota
            MatIfiNoSupervisadaGastoFami(j, 1) = rsDatIfiNoSupervisadaGastoFami!CDescripcion
            MatIfiNoSupervisadaGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiNoSupervisadaGastoFami!nMonto), 0, rsDatIfiNoSupervisadaGastoFami!nMonto), "#0.00")
            rsDatIfiNoSupervisadaGastoFami.MoveNext
        j = j + 1
        Loop
        rsDatIfiNoSupervisadaGastoFami.Close
        Set rsDatIfiNoSupervisadaGastoFami = Nothing
        
        'Ratios: Aceptable / Critico ->*****
         If Not (rsAceptableCritico.EOF Or rsAceptableCritico.BOF) Then
            If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                Me.lblCapaAceptable.Caption = "Aceptable"
                Me.lblCapaAceptable.ForeColor = &H8000&
            Else
                Me.lblCapaAceptable.Caption = "Cr?tico"
                Me.lblCapaAceptable.ForeColor = vbRed
            End If
        Else
            Me.lblCapaAceptable.Visible = False
        End If
        
        'Ratios e Indicadores
        If CDbl(rsDatRatios!nCapPagNeta) > 0 Then
            txtCapacidadNeta.Text = CStr(rsDatRatios!nCapPagNeta * 100) & "%"
            lblCapacidadPago.Visible = True
            txtCapacidadNeta.Visible = True
            lblCapaAceptable.Visible = True
        Else
            lblCapacidadPago.Visible = False
            txtCapacidadNeta.Visible = False
            lblCapaAceptable.Visible = False
        End If
        
        
       If Not (rsMantenimientoSinConvenioIngresos.EOF And rsMantenimientoSinConvenioIngresos.BOF) Then
            FormateaFlex fgIngresos
            Do While Not rsMantenimientoSinConvenioIngresos.EOF
                fgIngresos.AdicionaFila
                lnFila = fgIngresos.row
                
                fgIngresos.TextMatrix(lnFila, 1) = rsMantenimientoSinConvenioIngresos!nCodIngr
                fgIngresos.TextMatrix(lnFila, 2) = rsMantenimientoSinConvenioIngresos!cConsDescripcion
                fgIngresos.TextMatrix(lnFila, 3) = Format(rsMantenimientoSinConvenioIngresos!nMonto, "#,##0.00")
                               
                rsMantenimientoSinConvenioIngresos.MoveNext
                
                Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
                    'Case 1, 5, 6 'CTI320200110 ERS003-2020. Coment?
                    Case 1, 6 'CTI320200110 ERS003-2020. Agreg?
                            fgIngresos.BackColorRow (&HC0FFFF)
                    'CTI320200110 ERS003-2020. Agreg?
                    Case 4
                        If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
                            Me.fgIngresos.RowHeight(lnFila) = 1
                        End If
                    Case 5
    '                    If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
    '                        Me.fgIngresos.RowHeight(lnFila) = 1
    '                    Else
                            fgIngresos.BackColorRow (&HC0FFFF)
    '                    End If
                    'Fin CTI320200110
                End Select
            Loop
                rsMantenimientoSinConvenioIngresos.Close
                Set rsMantenimientoSinConvenioIngresos = Nothing
        End If
        
        ReDim pMtrNegocio(3, 0)
        For i = 1 To (rsMantenimientoSinConvenioGastoNegocio.RecordCount)
            ReDim Preserve pMtrNegocio(3, i)
              pMtrNegocio(0, i) = rsMantenimientoSinConvenioGastoNegocio!nCodGasto
              pMtrNegocio(1, i) = rsMantenimientoSinConvenioGastoNegocio!cConsDescripcion
              pMtrNegocio(2, i) = Format(rsMantenimientoSinConvenioGastoNegocio!nMonto, "#,##0.00")
              rsMantenimientoSinConvenioGastoNegocio.MoveNext
        Next i
        RSClose rsMantenimientoSinConvenioGastoNegocio
    'Fin CTI320200110 ERS003-2020
End Function

Private Sub txtPropCreditoDestino_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTab1.Tab = 3 'LUCV20171115, Agreg? segun correo: RUSI
        If Not fbTieneReferido6Meses Then
            txtReferidosComentario.SetFocus
        Else
            cmdGuardarConConvenio.SetFocus
        End If
    End If
End Sub

'Popuesta de Credito
Private Sub txtPropCreditoFechaVista_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        EnfocaControl txtPropCreditoEntornoFamiliar
    End If
End Sub

Private Sub txtPropCreditoEntornoFamiliar_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtPropCreditoGiroNegocio
    End If
End Sub

Private Sub txtPropCreditoFechaVista_LostFocus()
    If Not IsDate(txtPropCreditoFechaVista) Then
       MsgBox "Verifique Dia,Mes,A?o , Fecha Incorrecta", vbInformation, "Aviso"
       'txtPropCreditoFechaVista.SetFocus
    End If
End Sub

Private Sub txtPropCreditoGiroNegocio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtPropCreditoExpCrediticia
    End If
End Sub

Private Sub txtPropCreditoExpCrediticia_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtPropCreditoFormNegocio
    End If
End Sub

Private Sub txtPropCreditoFormNegocio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtPropCreditoColateralesGarantias
    End If
End Sub
Private Sub txtPropCreditoColateralesGarantias_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtPropCreditoDestino
    End If
End Sub

'Comentarios y Referidos
Private Sub txtReferidosComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAgregarConConvenio
    End If
End Sub

'Ingresos y Egresos
Private Sub cmbFechaMes1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtAnoMes1
    End If
End Sub

Private Sub txtAnoMes1_KeyPress(KeyAscii As Integer)
txtAnoMes1.MaxLength = "4"
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdLlamaRemBrutaTotalMes1
    End If
End Sub

Private Sub cmbFechaMes2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtAnoMes2
    End If
End Sub

Private Sub txtAnoMes2_KeyPress(KeyAscii As Integer)
txtAnoMes2.MaxLength = "4"
KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdLlamaRemBrutaTotalMes2
    End If
End Sub

Private Function Validar() As Boolean
Dim i As Integer
Dim j As Integer
Dim lsFecha As String

Validar = True

'Informacion del Negocio
    If fnTipoAportacion = 0 Then
        MsgBox "Seleccione Tipo de Aportacion", vbInformation, "Aviso"
        SSTabInfoNego.Tab = 0
        Validar = False
        Exit Function
    End If
    If fnTipoInstitucion = 0 Then
        MsgBox "Seleccione Tipo de Institucion", vbInformation, "Aviso"
        SSTabInfoNego.Tab = 0
        Validar = False
        Exit Function
    End If
    If fnTipoPlanilla = 0 Then
        MsgBox "Seleccione Tipo de Planilla", vbInformation, "Aviso"
        SSTabInfoNego.Tab = 0
        Validar = False
        Exit Function
    End If

'Ingresos y Egresos
    If txtAnoMes1.Text = "" Then
        MsgBox "Ingrese el A?o de la Evaluacion MES 1", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtAnoMes1.SetFocus
        Validar = False
        Exit Function
    End If
    If IsArray(MtrRemuneracionBrutaTotal1) Then
    Else
        MsgBox "Debe Ingresar Remuneraciones de la Evaluacion del Mes 1", vbInformation, "Aviso"
        SSTab1.Tab = 0
        cmdLlamaRemBrutaTotalMes1.SetFocus
        Validar = False
        Exit Function
    End If
    
If fnTipoInstitucion = 1 Then
    If IsArray(MtrDescuento1) Then
    Else
        MsgBox "Debe Ingresar Descuentos de la Evaluacion del Mes 1", vbInformation, "Aviso"
        SSTab1.Tab = 0
        cmdLlamaDescuentoMes1.SetFocus
        Validar = False
        Exit Function
    End If
End If

    If txtAnoMes2.Text = "" Then
        MsgBox "Ingrese el A?o de la Evaluacion MES 2", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtAnoMes2.SetFocus
        Validar = False
        Exit Function
    End If
    If IsArray(MtrRemuneracionBrutaTotal2) Then
    Else
        MsgBox "Debe Ingresar Remuneraciones de la Evaluacion del Mes 2", vbInformation, "Aviso"
        cmdLlamaRemBrutaTotalMes2.SetFocus
        Validar = False
        Exit Function
    End If
    
If fnTipoInstitucion = 1 Then
    If IsArray(MtrDescuento2) Then
    Else
        MsgBox "Debe Ingresar Descuentos de la Evaluacion del Mes 2", vbInformation, "Aviso"
        cmdLlamaDescuentoMes2.SetFocus
        Validar = False
        Exit Function
    End If
End If

If lnColocCondi <> 4 Then
'Propuesta de Credito
    If txtPropCreditoFechaVista.Text = "__/__/____" Then
        MsgBox "Ingrese Fecha de Visita", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoFechaVista.SetFocus
        Validar = False
        Exit Function
    End If
    
    lsFecha = ValidaFecha(txtPropCreditoFechaVista)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        SSTab1.Tab = 2
        EnfocaControl txtPropCreditoFechaVista
        fEnfoque txtPropCreditoFechaVista
        Validar = False
        Exit Function
    End If
    
    If txtPropCreditoEntornoFamiliar.Text = "" Then
        MsgBox "Ingrese Sobre el Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoEntornoFamiliar.SetFocus
        Validar = False
        Exit Function
    End If
    
    If txtPropCreditoGiroNegocio.Text = "" Then
        MsgBox "Ingrese Sobre el Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoGiroNegocio.SetFocus
        Validar = False
        Exit Function
    End If
     
     If txtPropCreditoExpCrediticia.Text = "" Then
        MsgBox "Ingrese Sobre la Experiencia Crediticia", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoExpCrediticia.SetFocus
        Validar = False
        Exit Function
    End If
       
       If txtPropCreditoFormNegocio.Text = "" Then
        MsgBox "Ingrese Sobre la Consistencia de la Informacion y la Formalidad del Negocio", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoFormNegocio.SetFocus
        Validar = False
        Exit Function
    End If
       
       If txtPropCreditoColateralesGarantias.Text = "" Then
        MsgBox "Ingrese Sobre los Colaterales o Garantias", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoColateralesGarantias.SetFocus
        Validar = False
        Exit Function
    End If
       
       If txtPropCreditoDestino.Text = "" Then
        MsgBox "Ingrese Sobre el Destino y el Impacto del Mismo", vbInformation, "Aviso"
        SSTab1.Tab = 2
        txtPropCreditoDestino.SetFocus
        Validar = False
        Exit Function
    End If
End If

'If lnColocCondi = 1 Then 'LUCV2017115, Seg?n correo: RUSI
If Not fbTieneReferido6Meses Then
'Comentario y referidos
    If txtReferidosComentario.Text = "" Then
        MsgBox "Ingrese Comentarios", vbInformation, "Aviso"
        SSTab1.Tab = 3
        txtReferidosComentario.SetFocus
        Validar = False
        Exit Function
    End If
            
    If feReferidosConConvenio.TextMatrix(1, 1) = "" Then
        MsgBox "Ingrese Referidos", vbInformation, "Aviso"
        SSTab1.Tab = 3
        feReferidosConConvenio.SetFocus
        Validar = False
        Exit Function
    End If
    
    If feReferidosConConvenio.rows - 1 < 2 Then
        MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
        SSTab1.Tab = 3
        cmdAgregarConConvenio.SetFocus
        Validar = False
        Exit Function
    End If
    
    For i = 1 To feReferidosConConvenio.rows - 1
        If feReferidosConConvenio.TextMatrix(i, 2) = 0 Then
            MsgBox "DNI mal Ingresado", vbInformation, "Alerta"
            SSTab1.Tab = 3
            feReferidosConConvenio.SetFocus
            Validar = False
            Exit Function
        ElseIf feReferidosConConvenio.TextMatrix(i, 3) = 0 Then
            MsgBox "Telefono mal Ingresado", vbInformation, "Alerta"
            SSTab1.Tab = 3
            feReferidosConConvenio.SetFocus
            Validar = False
            Exit Function
        End If
    Next i
    
    For i = 1 To feReferidosConConvenio.rows - 1 'Verfica ambos DNI que no sean iguales
            For j = 1 To feReferidosConConvenio.rows - 1
                If i <> j Then
                    If feReferidosConConvenio.TextMatrix(i, 2) = feReferidosConConvenio.TextMatrix(j, 2) Then
                        MsgBox "No se puede ingresar el mismo DNI mas de una vez en los referidos", vbInformation, "Alerta"
                        Validar = False
                        Exit Function
                    End If
                End If
            Next
        Next
    
End If
    Dim lsMensajeIfi As String 'LUCV20161115
    If Not ValidaIfiExisteCompraDeuda(fsCtaCod, MatIfiGastoFami, , lsMensajeIfi, MatIfiNoSupervisadaGastoFami) Or Len(Trim(lsMensajeIfi)) > 0 Then
        MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
        Me.SSTab1.Tab = 0
        Exit Function
    End If

    'Ingresos
'    Dim nCon As Currency
'    nCon = 0
'    For i = 1 To fgIngresos.rows - 1
'        nCon = nCon + fgIngresos.TextMatrix(i, 3)
'    Next i
'    If nCon = 0 Then
'        MsgBox "Ingrese Datos en INGRESOS", vbInformation, "Aviso"
'        SSTab1.Tab = 0
'        fgIngresos.SetFocus
'        Validar = False
'        Exit Function
'    End If
    
End Function

Private Sub CalcularSectorSalud()
    Dim nRemBrutTot As Currency
    Dim nDescuento As Currency
    Dim nRestanteDescuento As Currency
    Dim nCapacidad As Currency
    
    nRemBrutTot = val(Replace(txtRemBrutaTotalPromedio.Text, ",", ""))
    nDescuento = nRemBrutTot * 0.5
        
    txtMontoMaxIngDescontarPromedio.Text = Format(nDescuento, "#,##0.00")
End Sub

Private Function Registro()
    'si el cliente es nuevo-> referido obligatorio
    
    'If  lnColocCondi = 1 Then 'LUCV2017115, Seg?n correo: RUSI
    If Not fbTieneReferido6Meses Then
        txtReferidosComentario.Enabled = True
        feReferidosConConvenio.Enabled = True
        cmdAgregarConConvenio.Enabled = True
        cmdQuitarConConvenio.Enabled = True
    Else
        Frame9.Enabled = False
        Frame10.Enabled = False
        txtReferidosComentario.Enabled = False
        feReferidosConConvenio.Enabled = False
        cmdAgregarConConvenio.Enabled = False
        cmdQuitarConConvenio.Enabled = False
    End If
    
    'CTI320200110 ERS003-2020. Agreg?:
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
    Set rsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(fsCtaCod, fnFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami)

    'Carga de rsDatIfiGastoFami -> Matrix
    ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
    j = 0
    Do While Not rsDatIfiGastoFami.EOF
        MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
        MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
        MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#0.00")
        rsDatIfiGastoFami.MoveNext
    j = j + 1
    Loop
    rsDatIfiGastoFami.Close
    Set rsDatIfiGastoFami = Nothing
        
End Function

Public Sub Consultar()
    optTipoAportacion(1).Enabled = False
    optTipoAportacion(2).Enabled = False

    optTipoInstitucion(1).Enabled = False
    optTipoInstitucion(2).Enabled = False

    ChkSectorSalud.Enabled = False

    optTipoPlanilla(1).Enabled = False
    optTipoPlanilla(2).Enabled = False
    optTipoPlanilla(3).Enabled = False

    cmbFechaMes1.Enabled = False
    txtAnoMes1.Enabled = False

    cmbFechaMes2.Enabled = False
    txtAnoMes2.Enabled = False

    txtPropCreditoFechaVista.Enabled = False
    txtPropCreditoEntornoFamiliar.Enabled = False
    txtPropCreditoGiroNegocio.Enabled = False
    txtPropCreditoExpCrediticia.Enabled = False
    txtPropCreditoFormNegocio.Enabled = False
    txtPropCreditoColateralesGarantias.Enabled = False
    txtPropCreditoDestino.Enabled = False

    txtReferidosComentario.Enabled = False
    feReferidosConConvenio.Enabled = False
    cmdAgregarConConvenio.Enabled = False
    cmdQuitarConConvenio.Enabled = False

    cmdInformeVisitaConConvenio.Enabled = False
    cmdImprimir.Enabled = False

    cmdGuardarConConvenio.Enabled = False
    cmdActualizarConConvenio.Enabled = False
End Sub

'*
Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este m?dulo", vbInformation, "Aviso"
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = False
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean)
    optTipoAportacion(1).Enabled = pbHabilitaA
    optTipoAportacion(2).Enabled = pbHabilitaA
    
    optTipoInstitucion(1).Enabled = pbHabilitaA
    optTipoInstitucion(2).Enabled = pbHabilitaA
    
    ChkSectorSalud.Enabled = pbHabilitaA
    
    optTipoPlanilla(1).Enabled = pbHabilitaA
    optTipoPlanilla(2).Enabled = pbHabilitaA
    optTipoPlanilla(3).Enabled = pbHabilitaA
    
    cmbFechaMes1.Enabled = pbHabilitaA
    txtAnoMes1.Enabled = pbHabilitaA
    
    cmbFechaMes2.Enabled = pbHabilitaA
    txtAnoMes2.Enabled = pbHabilitaA
    
    cmdLlamaRemBrutaTotalMes1.Enabled = pbHabilitaA
    cmdLlamaRemBrutaTotalMes2.Enabled = pbHabilitaA
    
    cmdLlamaDescuentoMes1.Enabled = pbHabilitaA
    cmdLlamaDescuentoMes2.Enabled = pbHabilitaA
    
    txtPropCreditoFechaVista.Enabled = pbHabilitaA
    txtPropCreditoEntornoFamiliar.Enabled = pbHabilitaA
    txtPropCreditoGiroNegocio.Enabled = pbHabilitaA
    txtPropCreditoExpCrediticia.Enabled = pbHabilitaA
    txtPropCreditoFormNegocio.Enabled = pbHabilitaA
    txtPropCreditoColateralesGarantias.Enabled = pbHabilitaA
    txtPropCreditoDestino.Enabled = pbHabilitaA
            
    cmdGuardarConConvenio.Enabled = pbHabilitaA
    cmdActualizarConConvenio.Enabled = pbHabilitaA
    
    fgIngresos.Enabled = pbHabilitaA 'CTI3 ERS0032020
End Function

'CTI320200110 ERS003-2020. Agreg?
Private Sub CargarFlexEdit() 'Registrar New Formato Evaluacion
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMonto As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    nMonto = Format(0, "00.00")
    
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(fnFormato, fsCtaCod, rsFeGastoNeg, rsFeDatGastoFam, rsFeDatOtrosIng)
   'Otros Ingresos
    fgIngresos.Clear
    fgIngresos.FormaCabecera
    'fgIngresos.Rows = 3
    Call LimpiaFlex(fgIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            If rsFeDatOtrosIng!nConsValor <> 1 Then
                fgIngresos.AdicionaFila
                lnFila = fgIngresos.row
                fgIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
                fgIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
                fgIngresos.TextMatrix(lnFila, 3) = Format(rsFeDatOtrosIng!nMonto, "#,##0.00")
                
                Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
                    'Case 1, 5, 6 'CTI320200110 ERS003-2020. Coment?
                    Case 6 'CTI320200110 ERS003-2020. Agreg?
                        fgIngresos.BackColorRow (&HC0FFFF)
                    'CTI320200110 ERS003-2020. Agreg?
                    Case 4
                        If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
                            Me.fgIngresos.RowHeight(lnFila) = 1
                        End If
                    Case 5
'                        If (CDbl(fgIngresos.TextMatrix(lnFila, 3)) <= 0) Then
'                            Me.fgIngresos.RowHeight(lnFila) = 1
'                        Else
                            fgIngresos.BackColorRow (&HC0FFFF)
'                        End If
                    'Fin CTI320200110
                End Select
            End If
            rsFeDatOtrosIng.MoveNext
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing
                                                            
    'Gastos Negocio
'    feGastosNegocio.Clear
'    feGastosNegocio.FormaCabecera
'    feGastosNegocio.rows = 2
'    Call LimpiaFlex(feGastosNegocio)
'        Do While Not rsFeGastoNeg.EOF
'            feGastosNegocio.AdicionaFila
'            lnFila = feGastosNegocio.row
'            feGastosNegocio.TextMatrix(lnFila, 1) = rsFeGastoNeg!nConsValor
'            feGastosNegocio.TextMatrix(lnFila, 2) = rsFeGastoNeg!cConsDescripcion
'            feGastosNegocio.TextMatrix(lnFila, 3) = Format(rsFeGastoNeg!nMonto, "#,##0.00")
'
'                Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
'                    Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego
'                        Me.feGastosNegocio.BackColorRow &HC0FFFF, True
'                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
'                        Me.feGastosNegocio.ForeColorRow vbBlack, True
'                    Case gCodCuotaCmac
'                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
'                        Me.feGastosNegocio.ForeColorRow vbBlack, True
'                    Case Else
'                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
'                End Select
'            rsFeGastoNeg.MoveNext
'        Loop
'    rsFeGastoNeg.Close
'    Set rsFeGastoNeg = Nothing
    
    'Gastos Familiares
    feGastosFamiliares.Clear
    feGastosFamiliares.FormaCabecera
    feGastosFamiliares.rows = 2
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsFeDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,##0.00")
                
            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
                Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami
                   Me.feGastosFamiliares.BackColorRow &HC0FFFF
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case gCodDeudaLCNUGastoFami, gCodCuotaCmac
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case Else
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
             End Select
            rsFeDatGastoFam.MoveNext
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
End Sub
'Fin CTI320200110
'CTI3 ERS0032020*************************************
Private Sub fgIngresos_Click()
    If fgIngresos.Col = 3 Then
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            Else
                fgIngresos.ListaControles = "0-0-0-0-0"
            End If
    End If
   
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
            Case 1, 5, 6
                Me.fgIngresos.ColumnasAEditar = "X-X-X-3-X"
            Case 2, 3, 4
                Me.fgIngresos.ColumnasAEditar = "X-X-X-X-X"
        End Select
    End If
 
End Sub

Private Sub fgIngresos_EnterCell()
        If fgIngresos.Col = 3 Then
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            Else
                fgIngresos.ListaControles = "0-0-0-0-0"
            End If
        End If
    
        If fnTipoRegMant = 3 Then
           Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
               Case 1, 5, 6
                   Me.fgIngresos.ColumnasAEditar = "X-X-X-3-X"
               Case 2, 3, 4
                   Me.fgIngresos.ColumnasAEditar = "X-X-X-X-X"
               End Select
        End If
 
End Sub

Private Sub fgIngresos_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)

Dim pnTotal As Double

    If fgIngresos.Col = 3 Then
    
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                
                If fgIngresos.TextMatrix(fgIngresos.row, 3) = "0.00" Then
                   Set pMtrBoletaPago = Nothing
                       frmCredFormEvalBoletaPago.Inicio pnTotal, pMtrBoletaPago
                Else
                       frmCredFormEvalBoletaPago.Inicio pnTotal, pMtrBoletaPago
                End If
                       psCodigo = Format(pnTotal, "#,##0.00")
                
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                
                If fgIngresos.TextMatrix(fgIngresos.row, 3) = "0.00" Then
                    Set pMtrNegocio = Nothing
                        frmCredFormEvalNegocio.Inicio fsCtaCod, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto, pMtrNegocio
                Else
                        frmCredFormEvalNegocio.Inicio fsCtaCod, pnIngNegocio, pnEgrVenta, pnMargBruto, pnIngNeto, pMtrNegocio
                End If
                    psCodigo = Format(pnIngNeto, "#,##0.00")
                
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                                    
                If fgIngresos.TextMatrix(fgIngresos.row, 3) = "0.00" Then
                    Set pMtrReciboHono = Nothing
                        frmCredFormEvalReciboHonorarios.Inicio pnTotal, pMtrReciboHono
                Else
                        frmCredFormEvalReciboHonorarios.Inicio pnTotal, pMtrReciboHono
                End If
                    psCodigo = Format(pnTotal, "#,##0.00")
                    
            End If
        End If

End Sub

Private Sub fgIngresos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
        
    Editar = Split(Me.fgIngresos.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{Tab}"
        Exit Sub
    End If
    
End Sub

Private Sub fgIngresos_RowColChange()

'Se usa para que se direccione hacia abajo en una sola Columna
    If fgIngresos.Col = 3 Then
        fgIngresos.AvanceCeldas = Vertical
    Else
        fgIngresos.AvanceCeldas = Horizontal
    End If
    
'Se usa para activar el control en la celda indicada
    If fgIngresos.Col = 3 Then
            If CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 1 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 5 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            ElseIf CInt(fgIngresos.TextMatrix(fgIngresos.row, 1)) = 6 Then
                fgIngresos.ListaControles = "0-0-0-1-0"
            Else
                fgIngresos.ListaControles = "0-0-0-0-0"
            End If
    End If

'Se usa solo cuando entra a consultar el credito
    If fnTipoRegMant = 3 Then
        Select Case CInt(fgIngresos.TextMatrix(fgIngresos.row, 1))
            Case 1, 5, 6
                Me.fgIngresos.ColumnasAEditar = "X-X-X-3-X"
            Case 2, 3, 4
                Me.fgIngresos.ColumnasAEditar = "X-X-X-X-X"
            End Select
    End If
    
End Sub
Public Sub Calcular(ByVal pnTipo As Integer)
  '  Dim nTotalGastos As Currency
    Dim nIngNeta As Currency
    'Dim nExcedente As Currency
   ' Dim nCapPago As Currency
    
   ' nTotalGastos = 0
    nIngNeta = 0
   ' nExcedente = 0
   ' nCapPago = 0
    
 '   nTotalGastos = fgEgresos.SumaRow(3) - val(Replace(fgEgresos.TextMatrix(5, 3), ",", "")) - val(Replace(fgEgresos.TextMatrix(7, 3), ",", "")) - val(Replace(fgEgresos.TextMatrix(8, 3), ",", ""))
    nIngNeta = SumarCampo(fgIngresos, 3)
   ' nExcedente = SumarCampo(fgIngresos, 3) - nTotalGastos
    
'    txtRatioIngNeto.Text = Format(nIngNeta, "#,##0.00")
'
'    txtRatioExcedente.Text = Format(0, "#,##0.00") 'CTI320200110 ERS003-2020. Agreg?
    
End Sub
Private Sub fgIngresos_OnCellChange(pnRow As Long, pnCol As Long)

    If pnCol = 3 Then
        If IsNumeric(fgIngresos.TextMatrix(pnRow, pnCol)) Then
            Call Calcular(1)
        Else
            fgIngresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End If
    
    Select Case pnRow
    Case 2, 3, 4, 5
        If IsNumeric(fgIngresos.TextMatrix(pnRow, pnCol)) Then
           Select Case CCur(fgIngresos.TextMatrix(pnRow, pnCol))
            Case Is >= 0
                Case Else
                    MsgBox "Monto Incorrecto", vbInformation, "Alerta"
                    fgIngresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                    Exit Sub
            End Select
        Else
            fgIngresos.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End Select
    
End Sub
'****************************************************
