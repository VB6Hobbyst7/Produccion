VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgosInforme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Riesgos: Registrar"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14685
   Icon            =   "frmCredRiesgosInforme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      DragIcon        =   "frmCredRiesgosInforme.frx":030A
      Height          =   8895
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmCredRiesgosInforme.frx":0614
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDatosGeneral"
      Tab(0).Control(1)=   "fraAntecedente"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos de la Operación"
      TabPicture(1)   =   "frmCredRiesgosInforme.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSaldoDeudor"
      Tab(1).Control(1)=   "FrDatosColoc2"
      Tab(1).Control(2)=   "fraResultadosCredito"
      Tab(1).Control(3)=   "fraGarantias"
      Tab(1).Control(4)=   "fraEndCuota"
      Tab(1).Control(5)=   "fraEndMonto"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Evaluación Eco-Fin"
      TabPicture(2)   =   "frmCredRiesgosInforme.frx":064C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameConsumo"
      Tab(2).Control(1)=   "FrameNoMinorista"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Nivel de Riesgo"
      TabPicture(3)   =   "frmCredRiesgosInforme.frx":0668
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label81"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label80"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label71"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label83"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label63"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtFecha4"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "fraNivelRiesgo"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdCancelar4"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdGrabar4"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtAnalisis4"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "txtRecomendaciones4"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtConclusiones4"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "fraNumeroInforme"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtPersCodAnalista4"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "cmdImprimir4"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtConclusionGen"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).ControlCount=   16
      Begin VB.Frame FrameNoMinorista 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   7455
         Left            =   -74880
         TabIndex        =   149
         Top             =   360
         Width           =   14295
         Begin VB.TextBox txtCalidadEvaluacion3_1 
            Height          =   1815
            Left            =   7320
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   5040
            Width           =   6855
         End
         Begin VB.Frame Frame12 
            Caption         =   "Promedio de Declaraciones Anuales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   3960
            TabIndex        =   177
            Top             =   2640
            Width           =   6735
            Begin VB.TextBox txtUtilidades3_2 
               Height          =   375
               Left            =   5160
               TabIndex        =   75
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtUtilidades3_1 
               Height          =   375
               Left            =   1560
               TabIndex        =   73
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtVentasAnt3_2 
               Height          =   375
               Left            =   5160
               TabIndex        =   74
               Top             =   640
               Width           =   1335
            End
            Begin VB.TextBox txtVentasAnt3_1 
               Height          =   375
               Left            =   1560
               TabIndex        =   72
               Top             =   640
               Width           =   1335
            End
            Begin VB.Label lblAnioAnterior3_1 
               Caption         =   "Año Anterior:2011"
               Height          =   255
               Left            =   3720
               TabIndex        =   183
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label lblAnioActual3_1 
               Caption         =   "Año Anterior:2012"
               Height          =   255
               Left            =   120
               TabIndex        =   182
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label62 
               Caption         =   "Ventas:"
               Height          =   255
               Left            =   3960
               TabIndex        =   181
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label61 
               Caption         =   "Ventas:"
               Height          =   255
               Left            =   240
               TabIndex        =   180
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label60 
               Caption         =   "Utilidades:"
               Height          =   255
               Left            =   240
               TabIndex        =   179
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label59 
               Caption         =   "Utilidades:"
               Height          =   255
               Left            =   3960
               TabIndex        =   178
               Top             =   1200
               Width           =   855
            End
         End
         Begin VB.TextBox txtCondicionSector3_1 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   5040
            Width           =   6855
         End
         Begin VB.Frame Frame11 
            Caption         =   "EE.FF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   10800
            TabIndex        =   172
            Top             =   2640
            Width           =   3375
            Begin VB.TextBox txtPorcentaje3 
               Height          =   375
               Left            =   1560
               TabIndex        =   79
               Top             =   1500
               Width           =   1335
            End
            Begin VB.TextBox txtUtilidadesEEFF3 
               Height          =   375
               Left            =   1560
               TabIndex        =   78
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtVentasEEFF3 
               Height          =   375
               Left            =   1560
               TabIndex        =   77
               Top             =   640
               Width           =   1335
            End
            Begin MSMask.MaskEdBox txtFecha3 
               Height          =   375
               Left            =   1560
               TabIndex        =   76
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label58 
               Caption         =   "Fecha:"
               Height          =   255
               Left            =   240
               TabIndex        =   176
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label57 
               Caption         =   "Ventas:"
               Height          =   255
               Left            =   240
               TabIndex        =   175
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label56 
               Caption         =   "Utilidades:"
               Height          =   255
               Left            =   240
               TabIndex        =   174
               Top             =   1120
               Width           =   1215
            End
            Begin VB.Label Label55 
               Caption         =   "%:"
               Height          =   255
               Left            =   240
               TabIndex        =   173
               Top             =   1560
               Width           =   1815
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Declaración SUNAT - Régimen General"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   167
            Top             =   2640
            Width           =   3615
            Begin VB.TextBox txtMesAnt3_1_1 
               Height          =   375
               Left            =   1560
               TabIndex        =   69
               Top             =   640
               Width           =   1335
            End
            Begin VB.TextBox txtPromedio3 
               Height          =   375
               Left            =   1560
               TabIndex        =   68
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtMesAnt3_1_2 
               Height          =   375
               Left            =   1560
               TabIndex        =   70
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtMesAnt3_1_3 
               Height          =   375
               Left            =   1560
               TabIndex        =   71
               Top             =   1500
               Width           =   1335
            End
            Begin VB.Label Label54 
               Caption         =   "Promedio:"
               Height          =   255
               Left            =   240
               TabIndex        =   171
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label51 
               Caption         =   "Mes Anterior 1:"
               Height          =   255
               Left            =   240
               TabIndex        =   170
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label52 
               Caption         =   "Mes Anterior 2:"
               Height          =   255
               Left            =   240
               TabIndex        =   169
               Top             =   1120
               Width           =   1215
            End
            Begin VB.Label Label53 
               Caption         =   "Mes Anterior 3:"
               Height          =   255
               Left            =   240
               TabIndex        =   168
               Top             =   1560
               Width           =   1815
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Evaluación Economica / Financiera EE.FF."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   150
            Top             =   240
            Width           =   14055
            Begin VB.TextBox txtEvVentas3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1320
               TabIndex        =   54
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtEvUtilidades3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1320
               TabIndex        =   55
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtunVentas3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1320
               TabIndex        =   151
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox txtRazon3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1320
               TabIndex        =   56
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtCapPag3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   4920
               TabIndex        =   58
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtSens3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   375
               Left            =   4920
               TabIndex        =   59
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox txtCapTra 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   4920
               TabIndex        =   60
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtPatrimonio3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   8520
               TabIndex        =   61
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtApalan3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   8520
               TabIndex        =   62
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtMoraSector3_1 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   8520
               TabIndex        =   63
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox txtCapSoc3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   8520
               TabIndex        =   64
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtPasivo3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   12240
               TabIndex        =   65
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtLineaCred3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   12240
               TabIndex        =   66
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtOtrosIngr3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   12240
               TabIndex        =   67
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox txtSaldoDis3_1 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   375
               Left            =   4920
               TabIndex        =   57
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label50 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Ventas"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   240
               TabIndex        =   166
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label36 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Utilidades"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   240
               TabIndex        =   165
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label37 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Un/Ventas"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   240
               TabIndex        =   164
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Razón CTE"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   240
               TabIndex        =   163
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label39 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Saldo Disponible"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   3120
               TabIndex        =   162
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label40 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Capacidad de Pago"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   3120
               TabIndex        =   161
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label41 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Sensibilización"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   3120
               TabIndex        =   160
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label Label42 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Capital de Trabajo"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   3120
               TabIndex        =   159
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label43 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Patrimonio"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   6720
               TabIndex        =   158
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label44 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Apalancamiento"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   6720
               TabIndex        =   157
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label45 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Mora del Sector"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   6720
               TabIndex        =   156
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label Label46 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Capital Social"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   6720
               TabIndex        =   155
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label47 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pasivo"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   10440
               TabIndex        =   154
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label48 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Linea de Créd."
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   10440
               TabIndex        =   153
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label49 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Otros Ingresos"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   10440
               TabIndex        =   152
               Top             =   1080
               Width           =   1815
            End
         End
         Begin VB.Label Label66 
            Caption         =   "Calidad de Evaluación:"
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
            Left            =   7320
            TabIndex        =   185
            Top             =   4800
            Width           =   2295
         End
         Begin VB.Label Label65 
            Caption         =   "Condición Sector:"
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
            Left            =   120
            TabIndex        =   184
            Top             =   4800
            Width           =   1815
         End
      End
      Begin VB.Frame FrameConsumo 
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         Height          =   7095
         Left            =   -74760
         TabIndex        =   186
         Top             =   720
         Width           =   14055
         Begin VB.TextBox txtCalidadEvaluacion3_2 
            Height          =   855
            Left            =   6240
            MultiLine       =   -1  'True
            TabIndex        =   91
            Top             =   6120
            Width           =   7815
         End
         Begin VB.TextBox txtCondicionSector3_2 
            Height          =   855
            Left            =   6240
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   4800
            Width           =   7815
         End
         Begin VB.Frame fraOtrosDatos 
            Caption         =   "Otros Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   9840
            TabIndex        =   201
            Top             =   120
            Width           =   4215
            Begin VB.TextBox txtLineaCredito3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   43
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtPasivo3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   41
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtPatrimonio3_2 
               Height          =   375
               Left            =   1920
               TabIndex        =   42
               Top             =   640
               Width           =   1335
            End
            Begin VB.Label Label72 
               Caption         =   "Pasivo:"
               Height          =   255
               Left            =   240
               TabIndex        =   204
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label73 
               Caption         =   "Patrimonio:"
               Height          =   255
               Left            =   240
               TabIndex        =   203
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label74 
               Caption         =   "Linea de Credito:"
               Height          =   255
               Left            =   240
               TabIndex        =   202
               Top             =   1120
               Width           =   1215
            End
         End
         Begin VB.Frame fraDeclaSUNAT 
            Caption         =   "Declaración SUNAT - Régimen General"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   196
            Top             =   4800
            Width           =   3855
            Begin VB.TextBox txtMesAnt3_2_1 
               Height          =   375
               Left            =   1560
               TabIndex        =   50
               Top             =   640
               Width           =   1335
            End
            Begin VB.TextBox txtPromedio3_2 
               Height          =   375
               Left            =   1560
               TabIndex        =   49
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtMesAnt3_2_2 
               Height          =   375
               Left            =   1560
               TabIndex        =   51
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtMesAnt3_2_3 
               Height          =   375
               Left            =   1560
               TabIndex        =   52
               Top             =   1500
               Width           =   1335
            End
            Begin VB.Label Label76 
               Caption         =   "Promedio:"
               Height          =   255
               Left            =   240
               TabIndex        =   200
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label77 
               Caption         =   "Mes Anterior 1:"
               Height          =   255
               Left            =   240
               TabIndex        =   199
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label78 
               Caption         =   "Mes Anterior 2:"
               Height          =   255
               Left            =   240
               TabIndex        =   198
               Top             =   1120
               Width           =   1215
            End
            Begin VB.Label Label79 
               Caption         =   "Mes Anterior 3:"
               Height          =   255
               Left            =   240
               TabIndex        =   197
               Top             =   1560
               Width           =   1815
            End
         End
         Begin VB.Frame fraResultados 
            Caption         =   "Resultados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   9840
            TabIndex        =   187
            Top             =   1800
            Width           =   4215
            Begin VB.TextBox txtSensibilizacion3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   47
               Top             =   2040
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtApalancamiento3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   46
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtSaldoDis3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   44
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtCapPag3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   45
               Top             =   640
               Width           =   1335
            End
            Begin VB.TextBox txtMoraSector3_2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   48
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Label Label70 
               Caption         =   "Saldo Disponible :"
               Height          =   255
               Left            =   480
               TabIndex        =   192
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label69 
               Caption         =   "Capacidad de Pago :"
               Height          =   255
               Left            =   240
               TabIndex        =   191
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label68 
               Caption         =   "Apalancamiento :"
               Height          =   255
               Left            =   480
               TabIndex        =   190
               Top             =   1125
               Width           =   1335
            End
            Begin VB.Label Label67 
               Caption         =   "Sensibilización :"
               Height          =   255
               Left            =   600
               TabIndex        =   189
               Top             =   2160
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label Label75 
               Caption         =   "Mora del sector :"
               Height          =   255
               Left            =   600
               TabIndex        =   188
               Top             =   1680
               Width           =   1215
            End
         End
         Begin VB.Frame fraEvalEco 
            Caption         =   "Evaluación Economica (consumo)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   120
            TabIndex        =   193
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtEvaluacion3 
               Height          =   375
               Left            =   1320
               TabIndex        =   40
               Top             =   240
               Width           =   4575
            End
            Begin SICMACT.FlexEdit FEEvaluacionEco3 
               Height          =   3735
               Left            =   240
               TabIndex        =   194
               Top             =   720
               Width           =   9135
               _ExtentX        =   16113
               _ExtentY        =   6588
               Cols0           =   7
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "#-Tipo Eval-Titulo Eval-Descripción-Personal-Negocio-Unico"
               EncabezadosAnchos=   "0-1700-3000-2500-1800-0-0"
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
               ColumnasAEditar =   "X-X-X-X-X-X-X"
               ListaControles  =   "0-0-0-0-0-0-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-L-L-R-R-R"
               FormatosEdit    =   "0-0-0-0-0-0-0"
               TextArray0      =   "#"
               lbUltimaInstancia=   -1  'True
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
            Begin VB.Label Label82 
               Caption         =   "Evaluación:"
               Height          =   375
               Left            =   240
               TabIndex        =   195
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Label Label95 
            Caption         =   "Calidad de Evaluación:"
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
            Left            =   4200
            TabIndex        =   206
            Top             =   6360
            Width           =   2055
         End
         Begin VB.Label Label96 
            Caption         =   "Condición Sector:"
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
            Left            =   4680
            TabIndex        =   205
            Top             =   5040
            Width           =   1575
         End
      End
      Begin VB.TextBox txtConclusionGen 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   226
         Top             =   6720
         Width           =   13695
      End
      Begin VB.CommandButton cmdImprimir4 
         Caption         =   "Imprimir"
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
         Left            =   240
         TabIndex        =   217
         Top             =   8160
         Width           =   1455
      End
      Begin VB.TextBox txtPersCodAnalista4 
         Height          =   375
         Left            =   2160
         TabIndex        =   86
         Top             =   7560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame fraNumeroInforme 
         Caption         =   "Numero de Informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         TabIndex        =   216
         Top             =   480
         Width           =   6975
         Begin VB.TextBox txtInforme 
            Height          =   375
            Left            =   150
            TabIndex        =   83
            Top             =   180
            Width           =   2055
         End
      End
      Begin VB.TextBox txtConclusiones4 
         Height          =   2295
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   1560
         Width           =   13695
      End
      Begin VB.TextBox txtRecomendaciones4 
         Height          =   2175
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   85
         Top             =   4200
         Width           =   13695
      End
      Begin VB.TextBox txtAnalisis4 
         Height          =   375
         Left            =   2160
         TabIndex        =   87
         Top             =   7560
         Width           =   6495
      End
      Begin VB.CommandButton cmdGrabar4 
         Caption         =   "Grabar"
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
         Left            =   11160
         TabIndex        =   89
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar4 
         Caption         =   "Cancelar"
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
         Left            =   12600
         TabIndex        =   90
         Top             =   8160
         Width           =   1335
      End
      Begin VB.Frame fraNivelRiesgo 
         Caption         =   "Nivel  de Riesgo de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   207
         Top             =   480
         Width           =   6375
         Begin VB.ComboBox cmdNivelRiesgo4 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame fraEndMonto 
         Caption         =   "Escalonamiento de Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -69360
         TabIndex        =   142
         Top             =   3840
         Width           =   8535
         Begin VB.CommandButton cmdAgregarMontoPropuesto2 
            Caption         =   "Agregar"
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
            Left            =   5400
            TabIndex        =   223
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdQuitarMontoPropuesto2 
            Caption         =   "Quitar"
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
            Left            =   6480
            TabIndex        =   222
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtEscalonamiento2_2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5880
            TabIndex        =   39
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtTotal2_2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   38
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CommandButton cmdQuitar2_2 
            Caption         =   "Quitar"
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
            Left            =   1560
            TabIndex        =   37
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregar2_2 
            Caption         =   "Agregar"
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
            TabIndex        =   36
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtMontoPropuesto2 
            Enabled         =   0   'False
            Height          =   405
            Left            =   6480
            TabIndex        =   31
            Top             =   120
            Width           =   1575
         End
         Begin SICMACT.FlexEdit FEMontoCredito2 
            Height          =   1455
            Left            =   120
            TabIndex        =   143
            Top             =   600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2566
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Montos de Crédito"
            EncabezadosAnchos=   "0-2200"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1"
            ListaControles  =   "0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R"
            FormatosEdit    =   "0-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit FEMontoPropuesto2 
            Height          =   1455
            Left            =   3000
            TabIndex        =   144
            Top             =   600
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2566
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Monto Credito-Pagado"
            EncabezadosAnchos=   "0-1500-1200"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R"
            FormatosEdit    =   "0-2-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label35 
            Caption         =   "Escalonamiento:"
            Height          =   255
            Left            =   4560
            TabIndex        =   148
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label34 
            Caption         =   "Monto Propuesto:"
            Height          =   375
            Left            =   3000
            TabIndex        =   147
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label33 
            Caption         =   "Total:"
            Height          =   255
            Left            =   2400
            TabIndex        =   146
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label32 
            Caption         =   "Montos de crédito>30%"
            Height          =   375
            Left            =   120
            TabIndex        =   145
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraEndCuota 
         Caption         =   "Estacionamiento de cuota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74880
         TabIndex        =   135
         Top             =   3840
         Width           =   5415
         Begin VB.CommandButton cmdAgregarCuoPro2 
            Caption         =   "Agregar"
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
            TabIndex        =   225
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdQuitarCuoPro2 
            Caption         =   "Quitar"
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
            Left            =   4320
            TabIndex        =   224
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtEscalonamiento2_1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3960
            TabIndex        =   33
            Top             =   3405
            Width           =   1335
         End
         Begin VB.TextBox txtTotal2_1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3960
            TabIndex        =   32
            Top             =   3015
            Width           =   1335
         End
         Begin VB.TextBox txtCuotaPropuesta2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            TabIndex        =   30
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdQuitar2_1 
            Caption         =   "Quitar"
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
            Left            =   1440
            TabIndex        =   35
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregar2_1 
            Caption         =   "Agregar"
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
            Left            =   360
            TabIndex        =   34
            Top             =   2160
            Width           =   975
         End
         Begin SICMACT.FlexEdit FECuoPro2 
            Height          =   1455
            Left            =   2520
            TabIndex        =   136
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   2566
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cuota-Pagado"
            EncabezadosAnchos=   "0-1300-1100"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2"
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R"
            FormatosEdit    =   "0-2-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit FEEsCuaMa2 
            Height          =   1455
            Left            =   240
            TabIndex        =   137
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   2566
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Cuotas"
            EncabezadosAnchos=   "0-2000"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1"
            ListaControles  =   "0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R"
            FormatosEdit    =   "0-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label31 
            Caption         =   "Escalonamiento"
            Height          =   255
            Left            =   2520
            TabIndex        =   141
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Total:"
            Height          =   255
            Left            =   2520
            TabIndex        =   140
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Cuota propuesta"
            Height          =   255
            Left            =   2520
            TabIndex        =   139
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label28 
            Caption         =   "Cuotas>30%"
            Height          =   375
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraGarantias 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   133
         Top             =   2160
         Width           =   14055
         Begin SICMACT.FlexEdit FEGarantias2 
            Height          =   1095
            Left            =   120
            TabIndex        =   134
            Top             =   240
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   1931
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Tipo garantía-VRM-V.Gravamen-Descripción-Cobertura-cNumGarant-Indice"
            EncabezadosAnchos=   "0-2500-1200-1500-6000-1200-0-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-5-X-X"
            ListaControles  =   "0-0-0-0-0-3-3-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraResultadosCredito 
         Caption         =   "Resultado del Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -64920
         TabIndex        =   129
         Top             =   360
         Width           =   4095
         Begin VB.TextBox txtVgExpTotal2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   29
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtExposicionRiUn2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   28
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtExposicionMN2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label27 
            Caption         =   "VG/Exposic.Total:"
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   1300
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Exp. Riesgo Unico:"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   820
            Width           =   1455
         End
         Begin VB.Label Label25 
            Caption         =   "Exp.Total MN:"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame FrDatosColoc2 
         Caption         =   "Datos del Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -68280
         TabIndex        =   123
         Top             =   360
         Width           =   3255
         Begin VB.TextBox txtTEM2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2520
            TabIndex        =   25
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtTEMAnt2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   24
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtTEA2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2520
            TabIndex        =   22
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNroCuotas2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   21
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMontoPropuest2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtMontoPropuestTM2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtTipoCambio2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   23
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtComisionTrim2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   26
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label24 
            Caption         =   "TEM"
            Height          =   255
            Left            =   1800
            TabIndex        =   128
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label23 
            Caption         =   "TEM Ant."
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "T.E.A."
            Height          =   255
            Left            =   1800
            TabIndex        =   126
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblCuTC 
            Caption         =   "Nº de Cuotas"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Monto Propuesto"
            Height          =   615
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraSaldoDeudor 
         Caption         =   "Saldo Deudor MN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   121
         Top             =   360
         Width           =   6495
         Begin SICMACT.FlexEdit FECredVinculados2 
            Height          =   1215
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2143
            Cols0           =   5
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Titular-Moneda-Monto-cPersVinculados"
            EncabezadosAnchos=   "0-3500-1200-1200-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-R-R-L"
            FormatosEdit    =   "0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraAntecedente 
         Caption         =   "Antecedente Crediticio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   -74880
         TabIndex        =   111
         Top             =   2280
         Width           =   14175
         Begin VB.CommandButton cmdAgregarEV 
            Caption         =   "Agregar"
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
            Left            =   11880
            TabIndex        =   221
            Top             =   6000
            Width           =   975
         End
         Begin VB.CommandButton cmdQuitarEV 
            Caption         =   "Quitar"
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
            Left            =   12960
            TabIndex        =   220
            Top             =   6000
            Width           =   975
         End
         Begin VB.TextBox txtEvolDeudaSF1 
            Height          =   375
            Left            =   3600
            TabIndex        =   18
            Top             =   1680
            Width           =   10455
         End
         Begin VB.TextBox txtHistCMACM1 
            Height          =   375
            Left            =   3600
            TabIndex        =   17
            Top             =   1200
            Width           =   10455
         End
         Begin VB.TextBox txtNroCredOtorgados1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13200
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtAntigCred1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   9480
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtNroEntidades1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5280
            TabIndex        =   16
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtClasifSBS1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   4800
            TabIndex        =   12
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtTotalDeuda1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtClasifInterna1 
            Enabled         =   0   'False
            Height          =   405
            Left            =   2040
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin SICMACT.FlexEdit FEVinculados1 
            Height          =   1575
            Left            =   120
            TabIndex        =   112
            Top             =   2160
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   2778
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Apellidos y Nombres /Razo-Relación-Endeudamiento-Nº IFIs-Calificación SBS-Evolución Endeudamiento-cPersCodVinculado-nPersRelac"
            EncabezadosAnchos=   "400-4000-1200-1800-800-1600-3000-0-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-6-X-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C-C-C-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feEmpresasVinc 
            Height          =   1935
            Left            =   120
            TabIndex        =   219
            Top             =   3960
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   3413
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Codigo-Razon Social-Vinculado A:-Grupo Econ.-Relación-Endeudamiento-Nº IFIs-Calificación SBS-Evolución Endeudamiento-Aux"
            EncabezadosAnchos=   "400-2000-3500-2800-2000-1200-1800-800-1600-3000-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3-4-5-6-7-8-9-X"
            ListaControles  =   "0-1-0-3-3-0-0-0-3-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-0-2-3-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label21 
            Caption         =   "Empresas Vinculadas:"
            Height          =   255
            Left            =   240
            TabIndex        =   218
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Evolución deuda sistema financiero"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label Label18 
            Caption         =   "Historial Crediticio en CMACM"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Label Label17 
            Caption         =   "Nº Créd.Directos Otorgados"
            Height          =   495
            Left            =   11520
            TabIndex        =   118
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Antiguedad Cred. En CMACM(Mes)"
            Height          =   495
            Left            =   8040
            TabIndex        =   117
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Nº de entidades"
            Height          =   255
            Left            =   3720
            TabIndex        =   116
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Calificación SBS"
            Height          =   495
            Left            =   3720
            TabIndex        =   115
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Total Deuda S.F."
            Height          =   375
            Left            =   240
            TabIndex        =   114
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Calificación Interna:"
            Height          =   375
            Left            =   240
            TabIndex        =   113
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraDatosGeneral 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         TabIndex        =   100
         Top             =   360
         Width           =   14175
         Begin VB.TextBox txtAgeCodAct1 
            Height          =   375
            Left            =   11160
            TabIndex        =   215
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtCodAnalista1 
            Height          =   375
            Left            =   11640
            TabIndex        =   214
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtTpoCredCod1 
            Height          =   375
            Left            =   10680
            TabIndex        =   213
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtPersCIIU1 
            Height          =   375
            Left            =   10200
            TabIndex        =   212
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtCliente1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Width           =   6735
         End
         Begin VB.TextBox txtSector1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtAntigNeg1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   6240
            TabIndex        =   4
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtModalidad1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   6
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtNroCredito1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5400
            TabIndex        =   7
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtAgencia1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   9
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtAnalista1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   6240
            TabIndex        =   10
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtActividad1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   10200
            TabIndex        =   2
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtDestino1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   10200
            TabIndex        =   5
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtTipCredito1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   10200
            TabIndex        =   8
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label Label11 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   240
            TabIndex        =   110
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Sector"
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Antiguedad Neg:"
            Height          =   255
            Left            =   4440
            TabIndex        =   108
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Modalidad: "
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   1100
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Nº Crédito"
            Height          =   255
            Left            =   4440
            TabIndex        =   106
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Agencia"
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Analista"
            Height          =   255
            Left            =   4440
            TabIndex        =   104
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Actividad:"
            Height          =   255
            Left            =   8280
            TabIndex        =   103
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Destino"
            Height          =   255
            Left            =   8280
            TabIndex        =   102
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Crédito:"
            Height          =   255
            Left            =   8280
            TabIndex        =   101
            Top             =   1080
            Width           =   1215
         End
      End
      Begin MSMask.MaskEdBox txtFecha4 
         Height          =   375
         Left            =   12600
         TabIndex        =   88
         Top             =   7560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label63 
         Caption         =   "Conclusión General"
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
         Left            =   240
         TabIndex        =   227
         Top             =   6480
         Width           =   2535
      End
      Begin VB.Label Label83 
         Caption         =   "Conclusiones:"
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
         Left            =   240
         TabIndex        =   211
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Label71 
         Caption         =   "Recomendaciones y/o Observaciones:"
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
         Left            =   240
         TabIndex        =   210
         Top             =   3960
         Width           =   4935
      End
      Begin VB.Label Label80 
         Caption         =   "Analista de Riesgo:"
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
         Left            =   240
         TabIndex        =   209
         Top             =   7680
         Width           =   2055
      End
      Begin VB.Label Label81 
         Caption         =   "Fecha:"
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
         Left            =   10560
         TabIndex        =   208
         Top             =   7680
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1B 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   11160
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   661
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Registrar"
      TabPicture(0)   =   "frmCredRiesgosInforme.frx":0684
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frm"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ActxCta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbNivel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtGlosa"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtGlosa 
         Height          =   1815
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   1800
         Width           =   4935
      End
      Begin VB.ComboBox cmbNivel 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   1320
         Width           =   3225
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "Registrar"
         Height          =   375
         Left            =   -3960
         TabIndex        =   92
         Top             =   4080
         Width           =   1215
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   360
         TabIndex        =   96
         Top             =   600
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   741
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   960
         TabIndex        =   98
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label frm 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Riesgo:"
         Height          =   195
         Left            =   480
         TabIndex        =   97
         Top             =   1320
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCredRiesgosInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTotalCuotasPagadasIFIS As Currency
Dim lnTotalCuotasPagadasCMAC As Currency
Dim lnTotalCreditPagadasIFIS As Currency
Dim lnTotalCreditPagadasCMAC As Currency
Dim lnSumaGarantia As Currency
Dim lnExposicionTotal As Currency
Dim gbEstado As Boolean 'ALPA 20140725
Dim gnNumDec As Integer 'ALPA 20140725
Public Event Change() 'ALPA 20140725
Public Event KeyPress(KeyAscii As Integer) 'ALPA 20140725
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'Public Sub Inicio(ByVal psCtaCod As String)
'Me.ActxCta.NroCuenta = psCtaCod
'ActxCta.Enabled = False
'Me.Show 1
'End Sub
'

'
'Public Function ValidarDatos() As Boolean
'
'If Trim(Me.cmbNivel.Text) = "" Then
'    MsgBox "Seleccione el nivel de Riesgo.", vbInformation, "Aviso"
'    ValidarDatos = False
'    Exit Function
'End If
'
'
'If Trim(Me.txtGlosa.Text) = "" Then
'    MsgBox "Ingrese Descripcion del Informe.", vbInformation, "Aviso"
'    ValidarDatos = False
'    Exit Function
'End If
'
'
'If Len(Me.txtGlosa.Text) > 500 Then
'    MsgBox "El texto de la Descripcion del Informe no debe superar los 500 caracteres.", vbInformation, "Aviso"
'    ValidarDatos = False
'    Exit Function
'End If
'
'ValidarDatos = True
'End Function
'
'Private Sub Form_Load()
'Call CentraForm(Me)
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
'Call Llenarcontroles
'End Sub
'Sub Llenarcontroles()
'Dim oCons As COMDConstantes.DCOMConstantes
'Dim rsConstante As ADODB.Recordset
'Set oCons = New COMDConstantes.DCOMConstantes
'Set rsConstante = oCons.RecuperaConstantes(9999)
'Call Llenar_Combo_con_Recordset(rsConstante, cmbNivel)
'End Sub

Dim lsCtaCod As String
Dim lnContadorGada As Integer
Dim lnExposicionTotal_1 As Currency
Dim lnPasivoCorriente_1 As Currency
Dim lnPasivo As Currency
Dim lnNivelRiesgo As Integer
Dim lnCodForm As Integer 'LUCV20160906, Según ERS052-2016
Dim lnMontoPropuesto As Currency 'LUCV20160919, Según ERS052-2016

Private Sub cmdregistro()
Dim oCredito As COMDCredito.DCOMCredito
Dim lsMovNro As String 'EJVG20160531
Set oCredito = New COMDCredito.DCOMCredito
    lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) 'EJVG20160531
'    If MsgBox("Esta seguro de grabar el Informe de Riesgo de Credito Nº " & Trim(lsCtaCod), vbYesNo + vbInformation, "Aviso") = vbYes Then
        'Call oCredito.OpeInformeRiesgo(Trim(lsCtaCod), 2, , GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), Trim(Right(Me.cmdNivelRiesgo4.Text, 3)), Trim(Me.txtConclusionGen.Text), 2)
        oCredito.ActualizarInformeRiesgo Trim(lsCtaCod), 0, 2, , Trim(Right(cmdNivelRiesgo4.Text, 3)), , lsMovNro, , lsMovNro, Trim(txtConclusionGen.Text) 'EJVG20160531
        frmCredRiesgos.LLenarGrilla
        MsgBox "Informe de Riesgo registrado satisfactoriamente.", vbInformation, "Aviso"
        Unload Me
'    End If

End Sub
Public Sub Inicio(ByVal psCtaCod As String, Optional ByVal pbRegistro As Boolean = False)
'WIOR 20141010 AGREGO pnInformeID Y pbRegistro
    lsCtaCod = psCtaCod
    '*****-> LUCV20160906, Según ERS052-2016
    Dim rsFormatoEval As New ADODB.Recordset
    Dim oDCOMFormatosEval As New COMDCredito.DCOMFormatosEval
    Set rsFormatoEval = oDCOMFormatosEval.RecuperaFormatoEvaluacion(lsCtaCod)
    
    If Not (rsFormatoEval.BOF Or rsFormatoEval.EOF) Then
        lnCodForm = rsFormatoEval!nCodForm
    Else
        MsgBox "No se ha registrado la evaluación del crédito" & Chr(10) & " - Por favor registrar el formato de evaluación correspondiente.", vbInformation, "Alerta"
        Screen.MousePointer = 0 'RECO20161019 ERS060-2016
        Exit Sub
    End If
    '<-***** Fin LUCV20160906
    
    Call ConsNivelRiesgos
    Call Cuadro1(lsCtaCod)
    Call Cuadro1Vinculados(lsCtaCod)
    Call Cuadro1EmpVinculados(lsCtaCod) 'WIOR 20141016
    Call Cuadro2(lsCtaCod)
    Call Cuadro2SaldoDeudor(lsCtaCod)
    Call Cuadro2Garantias(lsCtaCod)
    Call Cuadro2Creditos(lsCtaCod)
    Call Cuadro2CartaFianza(lsCtaCod)
    Call Cuadro2EscaCuotasC1(lsCtaCod)
    Call Cuadro2EscaCuotasC2(lsCtaCod)
    Call Cuadro2EscaCreditosC1(lsCtaCod)
    Call Cuadro2EscaCreditosC2(lsCtaCod)
    Call Cuadro3(lsCtaCod)
    Call Cuadro3EvaluacionConsumo(lsCtaCod)
    Call Cuadro4(lsCtaCod)
    If pbRegistro Then 'Trim(txtRecomendaciones4.Text) = "" Then'WIOR 20141010
        cmdGrabar4.Enabled = True
        cmdCancelar4.Enabled = True
        cmdImprimir4.Enabled = False
        'WIOR 20141013*********
        fraDatosGeneral.Enabled = True
        fraAntecedente.Enabled = True
        fraSaldoDeudor.Enabled = True
        FrDatosColoc2.Enabled = True
        fraResultadosCredito.Enabled = True
        fraGarantias.Enabled = True
        fraEndCuota.Enabled = True
        fraEndMonto.Enabled = True
        fraEvalEco.Enabled = True
        fraOtrosDatos.Enabled = True
        fraResultados.Enabled = True
        fraDeclaSUNAT.Enabled = True
        txtCondicionSector3_2.Enabled = True
        txtCalidadEvaluacion3_2.Enabled = True
        fraNivelRiesgo.Enabled = True
        fraNumeroInforme.Enabled = True
        txtConclusiones4.Enabled = True
        txtRecomendaciones4.Enabled = True
        txtAnalisis4.Enabled = True
        txtAnalisis4.Enabled = True
        txtFecha4.Enabled = True
        FrameNoMinorista.Enabled = True
        FrameConsumo.Enabled = True
        txtConclusionGen.Enabled = True
        'WIOR FIN *************
    Else
        cmdGrabar4.Enabled = False
        cmdCancelar4.Enabled = False
        cmdImprimir4.Enabled = True
        'WIOR 20141013*********
        fraDatosGeneral.Enabled = False
        fraAntecedente.Enabled = False
        fraSaldoDeudor.Enabled = False
        FrDatosColoc2.Enabled = False
        fraResultadosCredito.Enabled = False
        fraGarantias.Enabled = False
        fraEndCuota.Enabled = False
        fraEndMonto.Enabled = False
        fraEvalEco.Enabled = False
        fraOtrosDatos.Enabled = False
        fraResultados.Enabled = False
        fraDeclaSUNAT.Enabled = False
        txtCondicionSector3_2.Enabled = False
        txtCalidadEvaluacion3_2.Enabled = False
        fraNivelRiesgo.Enabled = False
        fraNumeroInforme.Enabled = False
        txtConclusiones4.Enabled = False
        txtRecomendaciones4.Enabled = False
        txtAnalisis4.Enabled = False
        txtAnalisis4.Enabled = False
        txtFecha4.Enabled = False
        FrameNoMinorista.Enabled = False
        FrameConsumo.Enabled = False
        txtConclusionGen.Enabled = False
        'WIOR FIN *************
    End If
    lblAnioAnterior3_1.Caption = "Año Anterior:" & CStr(Year(DateAdd("M", -12, gdFecSis)))
    lblAnioActual3_1.Caption = "Año Anterior:" & CStr(Year(gdFecSis))
    'WIOR 20150225 ***************
    txtSaldoDis3_1.Text = Format(IIf(Trim(txtSaldoDis3_1.Text) = "", 0, txtSaldoDis3_1.Text), "###,###,###,##0.00")
    txtSaldoDis3_2.Text = Format(IIf(Trim(txtSaldoDis3_2.Text) = "", 0, txtSaldoDis3_2.Text), "###,###,###,##0.00")
    'WIOR 20150225 ***************
    Screen.MousePointer = 0 'RECO20161019 ERS060-2016
    Show 1
End Sub

Private Sub Cuadro1(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro1(psCtaCod, gdFecSis)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
        txtCliente1.Text = oRS!cPersNombreTitular
        'INICIO EAAS 20180202
        txtSector1.Text = IIf(IsNull(oRS!cSector), "", oRS!cSector)
        txtAntigNeg1.Text = IIf(IsNull(oRS!Antiguedad_Neg), "", oRS!Antiguedad_Neg)
        txtModalidad1.Text = IIf(IsNull(oRS!cModalidad), "", oRS!cModalidad)
        txtNroCredito1.Text = psCtaCod
        txtAgencia1.Text = IIf(IsNull(oRS!cAgeDescripcion), "", oRS!cAgeDescripcion)
        txtAnalista1.Text = IIf(IsNull(oRS!cPersNombreAnalista), "", oRS!cPersNombreAnalista)
        txtActividad1.Text = IIf(IsNull(oRS!cActividad), "", oRS!cActividad)
        txtDestino1.Text = IIf(IsNull(oRS!cDestino), "", oRS!cDestino)
        txtTipCredito1.Text = IIf(IsNull(oRS!cTipoCredito), "", oRS!cTipoCredito)
        txtClasifInterna1.Text = IIf(IsNull(oRS!Calificacion), "", oRS!Calificacion)
        txtClasifSBS1.Text = IIf(IsNull(oRS!CalificacionSBS), "", oRS!CalificacionSBS)
        'FIN EAAS 20180202
        txtAntigCred1.Text = IIf(IsNull(oRS!Antiguedad_CMACM), "", oRS!Antiguedad_CMACM)
        txtNroCredOtorgados1.Text = IIf(IsNull(oRS!CantidadCreditos), 0, oRS!CantidadCreditos)
        txtTotalDeuda1.Text = IIf(IsNull(oRS!nSumSalRCC), 0, oRS!nSumSalRCC)
        txtNroEntidades1.Text = IIf(IsNull(oRS!Can_EntsRCC), 0, oRS!Can_EntsRCC)
        txtHistCMACM1.Text = IIf(IsNull(oRS!cHisCredCMACM), "", oRS!cHisCredCMACM)
        txtEvolDeudaSF1.Text = IIf(IsNull(oRS!cEvolSistFina), "", oRS!cEvolSistFina)
        
        'If (Left(oRs!cTpoCredCod, 1) = "1" Or Left(oRs!cTpoCredCod, 1) = "2" Or Left(oRs!cTpoCredCod, 1) = "3" Or Left(oRs!cTpoCredCod, 1) = "4" Or Left(oRs!cTpoCredCod, 1) = "5") Then 'LUCV20160906, Comentó
        '*****-> LUCV20160906, Según ERS052-2016
        If (oRS!nCodForm = 7 Or oRS!nCodForm = 8) Then  'Formato de evaluación ->Consumo (7: Sin convenio  8: Con Convenio)
            FrameNoMinorista.Visible = False
            FrameConsumo.Visible = True
        Else
            FrameNoMinorista.Visible = True                 'Formato de evaluación no Minoristas (formato 5,6)
            FrameConsumo.Visible = False
        End If
        '<-***** Fin LUCV20160906
        txtPersCIIU1.Text = oRS!cPersCIIU
        txtTpoCredCod1.Text = oRS!cTpoCredCod
        txtAgeCodAct1.Text = oRS!cAgeCodAct
        txtCodAnalista1 = oRS!cAnalista
        lnNivelRiesgo = oRS!cNivelRiesgo
        oRS.MoveNext
    Loop
    End If
End Sub
Private Sub Cuadro1Vinculados(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro1Vinculados(psCtaCod)
    FormateaFlex FEVinculados1
    lnContadorGada = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            FEVinculados1.AdicionaFila
            FEVinculados1.TextMatrix(oRS.Bookmark, 0) = lnContadorGada + 1
            FEVinculados1.TextMatrix(oRS.Bookmark, 1) = oRS!cPersNombre
            FEVinculados1.TextMatrix(oRS.Bookmark, 2) = oRS!cConsDescripcion
            FEVinculados1.TextMatrix(oRS.Bookmark, 3) = Format(oRS!Val_Saldo, "###,###,###,##0.00")
            FEVinculados1.TextMatrix(oRS.Bookmark, 4) = oRS!nCantidadESF
            FEVinculados1.TextMatrix(oRS.Bookmark, 5) = oRS!cCalSBS
            FEVinculados1.TextMatrix(oRS.Bookmark, 6) = oRS!cEvoEndeu
            FEVinculados1.TextMatrix(oRS.Bookmark, 7) = oRS!cPersCod
            FEVinculados1.TextMatrix(oRS.Bookmark, 8) = oRS!nPrdPersRelac
            lnContadorGada = lnContadorGada + 1
            oRS.MoveNext
    Loop
    End If
End Sub

'WIOR 20141016 *************************************
Private Sub Cuadro1EmpVinculados(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro1EmpVinculados(psCtaCod)
    FormateaFlex feEmpresasVinc
    lnContadorGada = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            feEmpresasVinc.AdicionaFila
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 0) = lnContadorGada + 1
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 1) = oRS!cPersCodEmp
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 2) = oRS!cPersNombreEmp
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 3) = oRS!cPersNombreVinculado & Space(75) & oRS!cPersCodVinculado
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 4) = oRS!cGrupo & Space(75) & oRS!nGrupoCod
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 5) = oRS!cRelacion
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 6) = oRS!nMontoEnde
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 7) = oRS!nNroIFIS
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 8) = oRS!cCalSistF & Space(75) & oRS!nCalSistF
            feEmpresasVinc.TextMatrix(oRS.Bookmark, 9) = oRS!cEvolucion
            lnContadorGada = lnContadorGada + 1
            oRS.MoveNext
    Loop
    End If
End Sub
'WIOR FIN ******************************************

Private Sub Cuadro2(ByVal psCtaCod As String)
    If Mid(psCtaCod, 6, 3) = "514" Or Mid(psCtaCod, 6, 3) = "121" Or Mid(psCtaCod, 6, 3) = "221" Then
        FrDatosColoc2.Caption = "Datos del Crédito Indirecto"
        lblCuTC.Caption = "Tipo Cambio:"
        Label23.Caption = "Comisión Trim:"
        txtTipoCambio2.Visible = True
        txtNroCuotas2.Visible = False
        txtTEA2.Visible = False
        txtComisionTrim2.Visible = True
    Else
        FrDatosColoc2.Caption = "Datos del Crédito"
        lblCuTC.Caption = "NºCuotas:"
        Label23.Caption = "TEM Ant:"
        txtTipoCambio2.Visible = False
        txtNroCuotas2.Visible = True
        txtTEA2.Visible = True
        txtComisionTrim2.Visible = False
    End If
End Sub
Private Sub Cuadro2SaldoDeudor(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2SaldoDeudor(psCtaCod)
    FormateaFlex FECredVinculados2
    lnContadorGada = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            FECredVinculados2.AdicionaFila
            FECredVinculados2.TextMatrix(oRS.Bookmark, 0) = oRS!cPersCod
            FECredVinculados2.TextMatrix(oRS.Bookmark, 1) = oRS!cPersNombre
            FECredVinculados2.TextMatrix(oRS.Bookmark, 2) = oRS!cMoneda
            FECredVinculados2.TextMatrix(oRS.Bookmark, 3) = Format(oRS!nSaldo, "###,###,###,##0.00")
            FECredVinculados2.TextMatrix(oRS.Bookmark, 4) = oRS!nPrdPersRelac
'            cPersVinculados
            oRS.MoveNext
    Loop
    End If
End Sub

Private Sub Cuadro2Creditos(ByVal psCtaCod As String)
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    
    
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2Creditos(psCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
         lnMontoPropuesto = Format(oRS!nMontoPropuesto, "###,###,###,##0.00") 'LUCV20160919
         txtMontoPropuestTM2.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$")
         txtMontoPropuest2.Text = (lnMontoPropuesto)
         txtNroCuotas2.Text = oRS!nNrocuotas
         txtTEA2.Text = oRS!nTEA
         txtTEMAnt2.Text = IIf(IsNull(oRS!nTEMA), 0#, oRS!nTEMA)
         txtTEM2.Text = IIf(IsNull(oRS!nTEM), 0, oRS!nTEM)
         txtExposicionMN2.Text = Format(IIf(IsNull(oRS!nExMN), 0, oRS!nExMN), "###,###,###,##0.00")
         txtExposicionRiUn2.Text = Format(IIf(IsNull(oRS!nExRU), 0, oRS!nExRU), "###,###,###,##0.00")
         txtVgExpTotal2.Text = Format(IIf(IsNull(oRS!nVGET), 0, oRS!nVGET), "###,###,###,##0.00")
         lnExposicionTotal = oRS!nVGET
         txtCuotaPropuesta2.Text = Format(IIf(IsNull(oRS!nCuoPropuesta), 0, oRS!nCuoPropuesta), "###,###,###,##0.00")
         txtMontoPropuesto2.Text = Format(IIf(IsNull(oRS!nMonPropuesta), 0, oRS!nMonPropuesta), "###,###,###,##0.00")
        oRS.MoveNext
    Loop
    End If
End Sub
Private Sub Cuadro2CartaFianza(ByVal psCtaCod As String)
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2CartasFianzas(psCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
         txtMontoPropuestTM2.Text = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$")
         txtMontoPropuest2.Text = oRS!nMontoPropuesto
         txtTipoCambio2.Text = Format(oRS!nTipoCambio, "###,###,###,##0.00")
         txtComisionTrim2.Text = Format(oRS!nComisionTr, "###,###,###,##0.00")
         txtExposicionMN2.Text = Format(oRS!nExMN, "###,###,###,##0.00")
         txtExposicionRiUn2.Text = Format(oRS!nExRU, "###,###,###,##0.00")
         txtVgExpTotal2.Text = Format(oRS!nVGET, "###,###,###,##0.00")
         lnExposicionTotal = oRS!nVGET
        oRS.MoveNext
    Loop
    End If
End Sub

Private Sub Cuadro2Garantias(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2Garantias(psCtaCod)
    FormateaFlex FEGarantias2
    lnContadorGada = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            FEGarantias2.AdicionaFila
            FEGarantias2.TextMatrix(oRS.Bookmark, 0) = psCtaCod
            FEGarantias2.TextMatrix(oRS.Bookmark, 1) = oRS!cConsDescripcion
            FEGarantias2.TextMatrix(oRS.Bookmark, 2) = Format(oRS!nVRM, "###,###,###,##0.00")
            FEGarantias2.TextMatrix(oRS.Bookmark, 3) = Format(oRS!nVGravamen, "###,###,###,##0.00")
            FEGarantias2.TextMatrix(oRS.Bookmark, 4) = oRS!CDescripcion
            FEGarantias2.TextMatrix(oRS.Bookmark, 5) = oRS!cCober
            FEGarantias2.TextMatrix(oRS.Bookmark, 6) = oRS!cNumGarant
            FEGarantias2.TextMatrix(oRS.Bookmark, 7) = oRS!nTipoGarantia
            oRS.MoveNext
    Loop
    End If
End Sub

'WIOR 20141016 ***********************
Private Sub cmdAgregarCuoPro2_Click()
 FECuoPro2.AdicionaFila
End Sub

Private Sub cmdAgregarEV_Click()
 feEmpresasVinc.AdicionaFila
End Sub

Private Sub cmdAgregarMontoPropuesto2_Click()
 FEMontoPropuesto2.AdicionaFila
End Sub
'WIOR FIN ****************************

Private Sub cmdCancelar4_Click()
Unload Me
End Sub

Private Sub cmdImprimir4_Click()
    Call GenerarExcel
End Sub

'WIOR 20141016 ***********************
Private Sub cmdQuitarCuoPro2_Click()
    If FECuoPro2.TextMatrix(FECuoPro2.row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(FECuoPro2.row) + " de la lista de cuotas ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FECuoPro2.EliminaFila FECuoPro2.row
        End If
    End If
    lnTotalCuotasPagadasCMAC = SumarPagadasCMAC
    Call EscalonamientoCuotas
End Sub

Private Sub cmdQuitarEV_Click()
If feEmpresasVinc.TextMatrix(feEmpresasVinc.row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(feEmpresasVinc.row) + " ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            feEmpresasVinc.EliminaFila feEmpresasVinc.row
        End If
    End If
End Sub

Private Sub cmdQuitarMontoPropuesto2_Click()
    If FEMontoPropuesto2.TextMatrix(FEMontoPropuesto2.row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(FEMontoPropuesto2.row) + " de la lista de montos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FEMontoPropuesto2.EliminaFila FEMontoPropuesto2.row
        End If
    End If
    
    lnTotalCreditPagadasCMAC = SumarCreditosPagadasCMAC
    Call EscalonamientoCreditos
End Sub

Private Sub feEmpresasVinc_RowColChange()
Dim oConstante As COMDConstantes.DCOMConstantes
    Select Case feEmpresasVinc.Col
        Case 3 'VINCULADOS
            feEmpresasVinc.CargaCombo rsIntervinientes
        Case 4 'GRUPO ECONOMICO
            feEmpresasVinc.CargaCombo rsGrupoEconomico
        Case 8 'CALIFICACION
            Set oConstante = New COMDConstantes.DCOMConstantes
            feEmpresasVinc.CargaCombo oConstante.RecuperaConstantes(3010)
    End Select
End Sub

Private Function rsGrupoEconomico() As ADODB.Recordset
Dim oCons As COMDPersona.DCOMGrupoE
Dim rs As ADODB.Recordset
Dim rsGE As ADODB.Recordset
Dim i As Long

Set oCons = New COMDPersona.DCOMGrupoE
Set rsGE = oCons.ListarGrupoEconomico(1)

Set rs = New ADODB.Recordset
With rs
    'Crear RecordSet
    .Fields.Append "Descripcion", adVarChar, 100    '1
    .Fields.Append "Codigo", adInteger              '2
    .Open

    If Not (rsGE.EOF Or rsGE.BOF) Then
        For i = 1 To rsGE.RecordCount
            .AddNew
            .Fields("Descripcion") = UCase(Trim(rsGE!cConsDescripcion))
            .Fields("Codigo") = rsGE!nConsValor
            
            rsGE.MoveNext
        Next i
    End If
    If Not .EOF Then .MoveFirst
End With

Set rsGrupoEconomico = rs

End Function

Private Function rsIntervinientes() As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim i As Long

Set rs = New ADODB.Recordset
With rs
    'Crear RecordSet
    .Fields.Append "Descripcion", adVarChar, 200    '1
    .Fields.Append "Codigo", adVarChar, 13          '2
    .Open
    
    For i = 1 To FEVinculados1.rows - 1
        .AddNew
        .Fields("Descripcion") = FEVinculados1.TextMatrix(i, 1)
        .Fields("Codigo") = FEVinculados1.TextMatrix(i, 7)
    Next i
    If Not .EOF Then .MoveFirst
End With

Set rsIntervinientes = rs
End Function
'WIOR FIN ****************************

Private Sub FECuoPro2_RowColChange()
    lnTotalCuotasPagadasIFIS = SumarPagadasIFIS
    lnTotalCuotasPagadasCMAC = SumarPagadasCMAC 'WIOR 20141016
    Call EscalonamientoCuotas
End Sub



Private Sub FEEsCuaMa2_RowColChange()
    lnTotalCuotasPagadasIFIS = SumarPagadasIFIS
    Call EscalonamientoCuotas
End Sub

Private Sub FEGarantias2_DblClick()
    If FEGarantias2.TextMatrix(1, 0) <> "" And FEGarantias2.Col <> 5 Then
        frmGarantia.Consultar FEGarantias2.TextMatrix(FEGarantias2.row, 6)
    End If
End Sub

Private Sub FEGarantias2_OnChangeCombo()
    Call subcoberturaGarantia
End Sub

Private Sub FEMontoCredito2_RowColChange()
    lnTotalCreditPagadasIFIS = SumarCreditosPagadasIFIS
    Call EscalonamientoCreditos
End Sub

Private Sub FEMontoPropuesto2_RowColChange()
    lnTotalCreditPagadasIFIS = SumarCreditosPagadasIFIS
    lnTotalCreditPagadasCMAC = SumarCreditosPagadasCMAC 'WIOR 20141016
    Call EscalonamientoCreditos
End Sub


Private Sub EscalonamientoCuotas()
    If (lnTotalCuotasPagadasCMAC + lnTotalCuotasPagadasIFIS) <> 0# Then
    txtEscalonamiento2_1.Text = Round((CDbl(txtCuotaPropuesta2.Text) - (lnTotalCuotasPagadasCMAC + lnTotalCuotasPagadasIFIS)) / (lnTotalCuotasPagadasCMAC + lnTotalCuotasPagadasIFIS), 4) * 100
    End If
End Sub
Private Sub EscalonamientoCreditos()
If (lnTotalCreditPagadasCMAC + lnTotalCreditPagadasIFIS) <> 0# Then
    txtEscalonamiento2_2.Text = Round((CDbl(txtMontoPropuesto2.Text) - (lnTotalCreditPagadasCMAC + lnTotalCreditPagadasIFIS)) / (lnTotalCreditPagadasCMAC + lnTotalCreditPagadasIFIS), 4) * 100
End If
End Sub
Private Sub FEGarantias2_RowColChange()
If FEGarantias2.Col = 5 Then
    FEGarantias2.CargaCombo CargarCobertura(lnNivelRiesgo)
    Call subcoberturaGarantia
End If
End Sub
Private Sub subcoberturaGarantia()
    lnSumaGarantia = 0
    Dim i As Integer
    For i = 1 To FEGarantias2.rows - 1
        If Trim(Right(FEGarantias2.TextMatrix(i, 5), 5)) = 1 Then
            lnSumaGarantia = lnSumaGarantia + CDbl(FEGarantias2.TextMatrix(i, 2))
        ElseIf Trim(Right(FEGarantias2.TextMatrix(i, 5), 5)) = 2 Then
            lnSumaGarantia = lnSumaGarantia + CDbl(FEGarantias2.TextMatrix(i, 3))
        End If
    Next i
    If lnExposicionTotal > 0 Then
        txtVgExpTotal2.Text = Round(CCur(lnSumaGarantia / lnExposicionTotal), 4)
    Else
        txtVgExpTotal2.Text = 0#
    End If
End Sub

'
'Private Sub FEGarantias2_RowColChange()
'If FEGarantias2.Col = 5 Then
'    FEGarantias2.CargaCombo CargarCobertura
'End If
'End Sub
Private Function CargarCobertura(ByVal pnCodigoCobertura As Integer) As ADODB.Recordset
Dim rsCobertura As ADODB.Recordset
Set rsCobertura = New ADODB.Recordset

Dim oRsCons As ADODB.Recordset
Set oRsCons = New ADODB.Recordset

Dim ObjCons As COMDConstantes.DCOMConstantes
Set ObjCons = New COMDConstantes.DCOMConstantes
If pnCodigoCobertura = 1 Then
    Set oRsCons = ObjCons.RecuperaConstantes(gConsNivelRiesgosCobNR1)
Else
    Set oRsCons = ObjCons.RecuperaConstantes(gConsNivelRiesgosCobNR2)
End If

With rsCobertura
    'Crear RecordSet
     .Fields.Append "desc", adVarChar, 50
     .Fields.Append "value", adVarChar, 3
    .Open
    'Llenar Recordset
    If Not (oRsCons.BOF Or oRsCons.EOF) Then
    Do While Not oRsCons.EOF
        .AddNew
        .Fields("desc") = oRsCons!cConsDescripcion
        .Fields("value") = oRsCons!nConsValor
'        .AddNew
'        .Fields("desc") = "V.Grav."
'        .Fields("value") = "2"
'        .AddNew
'        .Fields("desc") = "SGR"
'        .Fields("value") = "3"
        oRsCons.MoveNext
    Loop
    End If
    
End With
rsCobertura.MoveFirst
Set CargarCobertura = rsCobertura
Set rsCobertura = Nothing
End Function

Private Sub cmdAgregar2_1_Click()
    FEEsCuaMa2.AdicionaFila
End Sub

Private Sub cmdAgregar2_2_Click()
    FEMontoCredito2.AdicionaFila
End Sub

Private Sub cmdGrabar4_Click()
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim i As Integer
    Set objDPersona = New COMDPersona.DCOMPersona
    If Len(Trim(ValidarDatos)) > 0 Then
        MsgBox ValidarDatos, vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Esta seguro de grabar el Informe de Riesgo del Credito Nº " & Trim(lsCtaCod), vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    
    
    Call objDPersona.InsertarInformeRiesgoCuadro1(gdFecSis, lsCtaCod, txtPersCIIU1.Text, txtAntigNeg1.Text, txtTpoCredCod1.Text, txtAgeCodAct1.Text, txtCodAnalista1.Text, txtClasifInterna1.Text, txtClasifSBS1.Text, txtAntigCred1.Text, txtNroCredOtorgados1.Text, txtTotalDeuda1.Text, txtNroEntidades1.Text, txtHistCMACM1.Text, txtEvolDeudaSF1.Text)
    For i = 1 To FEVinculados1.rows - 1
        Call objDPersona.InsertarInformeRiesgoCuadro1Vinculados(gdFecSis, lsCtaCod, FEVinculados1.TextMatrix(i, 7), FEVinculados1.TextMatrix(i, 8), FEVinculados1.TextMatrix(i, 3), FEVinculados1.TextMatrix(i, 4), FEVinculados1.TextMatrix(i, 5), FEVinculados1.TextMatrix(i, 6))
    Next i
    
    'WIOR 20141016 ************
    Dim cPersCodVinc  As String
    Dim nGrupoCod As Long
    Dim nCalif As Integer
    
    For i = 1 To feEmpresasVinc.rows - 1
        If Len(feEmpresasVinc.TextMatrix(i, 0)) > 0 Then
            cPersCodVinc = IIf(Trim(feEmpresasVinc.TextMatrix(i, 3)) = "", "", Trim(Right(feEmpresasVinc.TextMatrix(i, 3), 13)))
            nGrupoCod = CLng(IIf(Trim(feEmpresasVinc.TextMatrix(i, 4)) = "", 0, Trim(Right(feEmpresasVinc.TextMatrix(i, 4), 10))))
            nCalif = CInt(IIf(Trim(feEmpresasVinc.TextMatrix(i, 8)) = "", -1, Trim(Right(feEmpresasVinc.TextMatrix(i, 8), 10))))
    
            Call objDPersona.InsertarInformeRiesgoCuadro1EmpVinculados(gdFecSis, lsCtaCod, feEmpresasVinc.TextMatrix(i, 1), cPersCodVinc, nGrupoCod, feEmpresasVinc.TextMatrix(i, 5), _
            feEmpresasVinc.TextMatrix(i, 6), feEmpresasVinc.TextMatrix(i, 7), nCalif, feEmpresasVinc.TextMatrix(i, 9))
        End If
    Next i
    'WIOR FIN *****************
    
    If Mid(lsCtaCod, 6, 3) = "514" Or Mid(lsCtaCod, 6, 3) = "121" Or Mid(lsCtaCod, 6, 3) = "221" Then
        Call objDPersona.InsertarInformeRiesgoCuadro2CartasFianzas(gdFecSis, lsCtaCod, txtMontoPropuest2.Text, txtTipoCambio2.Text, txtComisionTrim2.Text, txtExposicionMN2.Text, txtExposicionRiUn2.Text, txtVgExpTotal2.Text)
    Else
        Call objDPersona.InsertarInformeRiesgoCuadro2Creditos(gdFecSis, lsCtaCod, txtMontoPropuest2.Text, txtNroCuotas2.Text, txtTEA2.Text, txtTEMAnt2.Text, txtTEM2.Text, txtExposicionMN2.Text, txtExposicionRiUn2.Text, txtVgExpTotal2.Text, txtCuotaPropuesta2.Text, txtMontoPropuesto2.Text)
    End If
    
    For i = 1 To FECredVinculados2.rows - 1
        If Len(FECredVinculados2.TextMatrix(i, 0)) > 0 Then
            Call objDPersona.InformeRiesgoCuadro2SaldoDeudor(gdFecSis, lsCtaCod, FECredVinculados2.TextMatrix(i, 0), FECredVinculados2.TextMatrix(i, 4), IIf(FECredVinculados2.TextMatrix(i, 2) = "SOLES", 1, 2), FECredVinculados2.TextMatrix(i, 3))
        End If
    Next i
    
    For i = 1 To FEGarantias2.rows - 1
        If Len(FEGarantias2.TextMatrix(i, 6)) > 0 Then
            Call objDPersona.InformeRiesgoCuadro2Garantias(gdFecSis, lsCtaCod, FEGarantias2.TextMatrix(i, 6), FEGarantias2.TextMatrix(i, 7), FEGarantias2.TextMatrix(i, 2), FEGarantias2.TextMatrix(i, 3), FEGarantias2.TextMatrix(i, 4), Trim(Right(FEGarantias2.TextMatrix(i, 5), 5)))
        End If
    Next i
    
    For i = 1 To FEEsCuaMa2.rows - 1
    If Len(FEEsCuaMa2.TextMatrix(i, 1)) > 0 Then
        Call objDPersona.InformeRiesgoCuadro2EscaCuotasC1(gdFecSis, lsCtaCod, FEEsCuaMa2.TextMatrix(i, 1))
    End If
    Next i
    
    For i = 1 To FECuoPro2.rows - 1
    If Len(FECuoPro2.TextMatrix(i, 1)) > 0 Then
        Call objDPersona.InformeRiesgoCuadro2EscaCuotasC2(gdFecSis, lsCtaCod, FECuoPro2.TextMatrix(i, 1), FECuoPro2.TextMatrix(i, 2))
    End If
    Next i
    
    For i = 1 To FEMontoCredito2.rows - 1
        If Len(FEMontoCredito2.TextMatrix(i, 1)) > 0 Then
            Call objDPersona.InformeRiesgoCuadro2EscaCreditosC1(gdFecSis, lsCtaCod, FEMontoCredito2.TextMatrix(i, 1))
        End If
    Next i
    
    For i = 1 To FEMontoPropuesto2.rows - 1
        If Len(FEMontoPropuesto2.TextMatrix(i, 1)) > 0 Then
            Call objDPersona.InformeRiesgoCuadro2EscaCreditosC2(gdFecSis, lsCtaCod, i, FEMontoPropuesto2.TextMatrix(i, 1), FEMontoPropuesto2.TextMatrix(i, 2))
        End If
    Next i
    'If (Left(txtTpoCredCod1.Text, 1) = "1" Or Left(txtTpoCredCod1.Text, 1) = "2" Or Left(txtTpoCredCod1.Text, 1) = "3" Or Left(txtTpoCredCod1.Text, 1) = "4" Or Left(txtTpoCredCod1.Text, 1) = "5") Then 'LUCV20160915, Comentó Según ERS052-2016
    If lnCodForm = 1 Or lnCodForm = 2 Or lnCodForm = 3 Or lnCodForm = 4 Or lnCodForm = 5 Or lnCodForm = 6 Or lnCodForm = 9 Then   'Formato->No minoristas
        Call objDPersona.InformeRiesgoCuadro3(gdFecSis, lsCtaCod, txtEvVentas3.Text, CDbl(txtEvUtilidades3.Text), CDbl(txtunVentas3.Text), CDbl(txtRazon3.Text), CDbl(IIf(Trim(txtSaldoDis3_1.Text) = "", 0, txtSaldoDis3_1.Text)), CDbl(txtCapPag3.Text), CDbl(IIf(Trim(txtSens3.Text) = "", 0, txtSens3.Text)), CDbl(txtCapTra.Text), CDbl(txtPatrimonio3.Text), CDbl(txtApalan3.Text), CDbl(txtMoraSector3_1.Text), CDbl(txtCapSoc3.Text), CDbl(txtPasivo3.Text), CDbl(txtLineaCred3.Text), CDbl(txtOtrosIngr3.Text), CDbl(txtPromedio3.Text), CDbl(txtMesAnt3_1_1.Text), CDbl(txtMesAnt3_1_2.Text), CDbl(txtMesAnt3_1_3.Text), CDbl(txtVentasAnt3_1.Text), CDbl(txtUtilidades3_1.Text), CDbl(txtVentasAnt3_2.Text), CDbl(txtUtilidades3_2.Text), CDate(txtFecha3.Text), CDbl(txtVentasEEFF3.Text), CDbl(txtUtilidadesEEFF3.Text), CDbl(txtPorcentaje3.Text), txtCondicionSector3_1.Text, txtCalidadEvaluacion3_1.Text, "")
    Else
        Call objDPersona.InformeRiesgoCuadro3(gdFecSis, lsCtaCod, 0, 0, 0, 0, CDbl(txtSaldoDis3_2.Text), CDbl(txtCapPag3_2.Text), CDbl(txtSensibilizacion3_2.Text), 0, CDbl(txtPatrimonio3_2.Text), CDbl(txtApalancamiento3_2.Text), CDbl(txtMoraSector3_2.Text), 0, CDbl(txtPasivo3_2.Text), CDbl(txtLineaCredito3_2.Text), 0, CDbl(txtPromedio3_2.Text), CDbl(txtMesAnt3_2_1.Text), CDbl(txtMesAnt3_2_2.Text), CDbl(txtMesAnt3_2_3.Text), 0, 0, 0, 0, CDate("1900/01/01"), 0, 0, 0, txtCondicionSector3_2.Text, txtCalidadEvaluacion3_2.Text, "")
        For i = 1 To FEEvaluacionEco3.rows - 1 'Formato->Consumos
            'Call objDPersona.InsertarInformeRiesgoCuadro3EvaluacionConsumo(gdFecSis, lsCtaCod, FEEvaluacionEco3.TextMatrix(i, 1), FEEvaluacionEco3.TextMatrix(i, 2), FEEvaluacionEco3.TextMatrix(i, 3), FEEvaluacionEco3.TextMatrix(i, 4), FEEvaluacionEco3.TextMatrix(i, 5), FEEvaluacionEco3.TextMatrix(i, 6)) 'LUCV20161509, Comentó Según ERS052-2016
            Call objDPersona.InsertarInformeRiesgoCuadro3EvaluacionConsumo(gdFecSis, lsCtaCod, FEEvaluacionEco3.TextMatrix(i, 1), FEEvaluacionEco3.TextMatrix(i, 2), FEEvaluacionEco3.TextMatrix(i, 3), FEEvaluacionEco3.TextMatrix(i, 4))
        Next i
    End If
    Call objDPersona.InsertarInformeRiesgoCuadro4(gdFecSis, lsCtaCod, Trim(Right(cmdNivelRiesgo4.Text, 4)), txtConclusiones4.Text, txtRecomendaciones4.Text, txtPersCodAnalista4.Text, txtFecha4.Text, txtInforme.Text)
    Call objDPersona.InsertarInformePersVinculados(lsCtaCod) 'EJVG20160610
    Call cmdregistro
    
    'RECO20161019 ERS060-2016 ******************************************************
    Dim oNCOMColocEval As New NCOMColocEval
    Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
    Dim rsCredSugAprob As New ADODB.Recordset
    Dim lcMovNro As String

    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(lsCtaCod, gTpoRegCtrlRiesgos)    'BY ARLO 20171027
    'Call oNCOMColocEval.updateEstadoExpediente(lsCtaCod) 'COMENTADO POR ARLO20171129
    Call oNCOMColocEval.insEstadosExpediente(lsCtaCod, "", "", "", "", lcMovNro, 2, 2001, gTpoRegCtrlRiesgos)
    MsgBox "Expediente Salio de Gerencia de Riesgs", vbInformation, "Aviso"
    Set oNCOMColocEval = Nothing
    'RECO FIN **********************************************************************
    MsgBox "Proceso se realizó correctamente", vbInformation
End Sub

Private Sub cmdQuitar2_1_Click()
    If FEEsCuaMa2.TextMatrix(FEEsCuaMa2.row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(FEEsCuaMa2.row) + " de la lista de cuotas?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FEEsCuaMa2.EliminaFila FEEsCuaMa2.row
        End If
    End If
    lnTotalCuotasPagadasIFIS = SumarPagadasIFIS
    Call EscalonamientoCuotas
End Sub

Private Sub cmdQuitar2_2_Click()
    If FEMontoCredito2.TextMatrix(FEMontoCredito2.row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(FEMontoCredito2.row) + " de la lista de montos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FEMontoCredito2.EliminaFila FEMontoCredito2.row
        End If
    End If
    
    lnTotalCreditPagadasIFIS = SumarCreditosPagadasIFIS
    Call EscalonamientoCreditos
End Sub

Private Sub Cuadro2EscaCuotasC1(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2EscaCuotasC1(psCtaCod)
    Dim lnContadorIFIS As Integer
    FormateaFlex FEEsCuaMa2
    lnContadorIFIS = 0
    lnTotalCuotasPagadasIFIS = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            lnContadorIFIS = lnContadorIFIS + 1
            FEEsCuaMa2.AdicionaFila
            FEEsCuaMa2.TextMatrix(oRS.Bookmark, 0) = psCtaCod
            FEEsCuaMa2.TextMatrix(oRS.Bookmark, 1) = oRS!cCuotas
            lnTotalCuotasPagadasIFIS = lnTotalCuotasPagadasIFIS + IIf(Len(Trim(oRS!cCuotas)) = 0, 0, oRS!cCuotas)
            oRS.MoveNext
    Loop
    End If
    If lnContadorIFIS > 0 Then
    lnTotalCuotasPagadasIFIS = SumarPagadasIFIS
    Call EscalonamientoCuotas
    End If
End Sub

Private Sub Cuadro2EscaCuotasC2(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2EscaCuotasC2(psCtaCod)
    FormateaFlex FECuoPro2
    lnContadorGada = 0
    lnTotalCuotasPagadasCMAC = 0
    txtTotal2_1.Text = CDbl(IIf(Len(txtCuotaPropuesta2.Text) = 0, 0, txtCuotaPropuesta2.Text))
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            FECuoPro2.AdicionaFila
            FECuoPro2.TextMatrix(oRS.Bookmark, 0) = psCtaCod
            FECuoPro2.TextMatrix(oRS.Bookmark, 1) = Format(oRS!nMontoPagado, "###,###,###,##0.00")
            FECuoPro2.TextMatrix(oRS.Bookmark, 2) = oRS!cCuotasPagadas
            lnTotalCuotasPagadasCMAC = lnTotalCuotasPagadasCMAC + oRS!nMontoPagado
            oRS.MoveNext
    Loop
    End If
    lnTotalCuotasPagadasCMAC = SumarPagadasCMAC

    Call EscalonamientoCuotas
End Sub

Private Sub Cuadro2EscaCreditosC1(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Dim lnContadorIFICRED As Integer
    Set oRS = New ADODB.Recordset
    Dim lnTotalCuotasPagadas As Currency
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2EscaCreditosC1(psCtaCod)
    FormateaFlex FEMontoCredito2
    lnContadorIFICRED = 0
    lnTotalCreditPagadasIFIS = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            FEMontoCredito2.AdicionaFila
            FEMontoCredito2.TextMatrix(oRS.Bookmark, 0) = psCtaCod
            FEMontoCredito2.TextMatrix(oRS.Bookmark, 1) = Format(oRS!nMontosCred, "###,###,###,##0.00")
            lnTotalCreditPagadasIFIS = lnTotalCreditPagadasIFIS + oRS!nMontosCred
            oRS.MoveNext
    Loop
    End If
    If lnContadorIFICRED > 0 Then
    'lnTotalCuotasPagadasCMAC = SumarPagadasCMAC
    lnTotalCreditPagadasIFIS = SumarCreditosPagadasIFIS
    Call EscalonamientoCreditos
    End If
End Sub

Private Sub Cuadro2EscaCreditosC2(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2EscaCreditosC2(psCtaCod)
    FormateaFlex FEMontoPropuesto2
    lnContadorGada = 0
    lnTotalCreditPagadasCMAC = 0
    txtTotal2_2.Text = CDbl(IIf(Len(txtMontoPropuesto2.Text) = 0, 0, txtMontoPropuesto2.Text))
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            FEMontoPropuesto2.AdicionaFila
            FEMontoPropuesto2.TextMatrix(oRS.Bookmark, 0) = psCtaCod
            FEMontoPropuesto2.TextMatrix(oRS.Bookmark, 1) = Format(oRS!nMontoCredito, "###,###,###,##0.00")
            FEMontoPropuesto2.TextMatrix(oRS.Bookmark, 2) = Format(oRS!nMontoPagado, "###,###,###,##0.00")
            txtTotal2_2.Text = CDbl(IIf(Len(txtTotal2_2.Text) = 0, 0, txtTotal2_2.Text)) + oRS!nMontoPagado
            lnTotalCreditPagadasCMAC = lnTotalCreditPagadasCMAC + oRS!nMontoPagado
            oRS.MoveNext
    Loop
    End If
    lnTotalCreditPagadasCMAC = SumarCreditosPagadasCMAC
   ' lnTotalCuotasPagadasIFIS = SumarPagadasIFIS
    'lnTotalCuotasPagadasCMAC = SumarPagadas
    Call EscalonamientoCreditos
End Sub

Private Sub Cuadro3Calculo(ByVal pnTipo As Integer)
Select Case pnTipo
    Case 1:
        If (txtPatrimonio3_2.Text) <> "" Then  'LUCV20160916
            If (txtPatrimonio3_2.Text) <= 0 Then
                txtApalancamiento3_2.Enabled = False
                txtApalancamiento3_2.Text = 0
            Else
                txtApalancamiento3_2.Enabled = True
                txtApalancamiento3_2.Text = Format(((txtPasivo3_2.Text) + (lnMontoPropuesto)) / (txtPatrimonio3_2.Text), "###,###,###,##0.00")
            End If
        Else
            txtApalancamiento3_2.Enabled = False
            txtApalancamiento3_2.Text = 0
        End If
    Case 2: 'Apalancamiento-> Paralelo
        If (txtPatrimonio3.Text) <> "" Then  'LUCV20160919
            If (txtPatrimonio3.Text) <= 0 Then
                txtApalan3.Enabled = False
                txtApalan3.Text = 0
            Else
                txtApalan3.Enabled = True
                txtApalan3.Text = Format(((txtPasivo3.Text) + (lnMontoPropuesto)) / (txtPatrimonio3.Text), "###,###,###,##0.00")
            End If
        Else
            txtApalan3.Enabled = False
            txtApalan3.Text = 0
        End If
End Select
End Sub

Private Sub Cuadro3(ByVal psCtaCod As String)
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
'    Dim lnExposicionTotal_1 As Currency
'    Dim lnPasivoCorriente_1 As Currency
    
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro3(psCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
        lnPasivo = Format(oRS!nPasivo, "###,###,###,##0.00")
        lnExposicionTotal_1 = Format(oRS!ExposicionTotal_1, "###,###,###,##0.00")
        lnPasivoCorriente_1 = Format(oRS!nPasivoCorriente_1, "###,###,###,##0.00")
        If lnCodForm = 9 Then
        txtPatrimonio3.Enabled = True
        End If
        
        '*****->Formato Evaluación: No minoristas Evaluación Económica / Financiera EEFF
        txtEvVentas3.Text = Format(oRS!nVentas, "###,###,###,##0.00")
        txtEvUtilidades3.Text = Format(oRS!nUtilidades, "###,###,###,##0.00")
        txtunVentas3.Text = Format(oRS!nUnVentas, "###,###,###,##0.00") '*100
        txtRazon3.Text = Format(oRS!nRazonCTE, "###,###,###,##0.00") '*100
        
        txtSaldoDis3_1.Text = Format(oRS!nSaldoDisponible, "###,###,###,##0.00")
        txtCapPag3.Text = Format(oRS!nCapaPago, "###,###,###,##0.00") '*100
        txtSens3.Text = Format(oRS!nSensibli, "###,###,###,##0.00")
        txtCapTra.Text = Format(oRS!nCapiTrab, "###,###,###,##0.00")
        
        txtPatrimonio3.Text = Format(oRS!nPatrimonio, "###,###,###,##0.00")
        txtApalan3.Text = Format(oRS!nApalancamiento, "###,###,###,##0.00") '*100
        txtMoraSector3_1.Text = Format(oRS!nMoraDelSector, "###,###,###,##0.00")
        txtCapSoc3.Text = Format(oRS!nCapitalSocial, "###,###,###,##0.00")
        
        txtPasivo3.Text = Format(oRS!nPasivo, "###,###,###,##0.00")
        txtLineaCred3.Text = Format(oRS!nLineaCred, "###,###,###,##0.00")
        txtOtrosIngr3.Text = Format(oRS!nOtrosIngr, "###,###,###,##0.00")
        
        'Declaración SUNAT - Regimen General
        txtPromedio3.Text = Format(oRS!nPromedio, "###,###,###,##0.00")
        txtMesAnt3_1_1.Text = Format(oRS!nMesAnterior1, "###,###,###,##0.00")
        txtMesAnt3_1_2.Text = Format(oRS!nMesAnterior2, "###,###,###,##0.00")
        txtMesAnt3_1_3.Text = Format(oRS!nMesAnterior3, "###,###,###,##0.00")
        'Promedio de declaraciones anuales
        txtVentasAnt3_1.Text = Format(oRS!nVentasAnter1, "###,###,###,##0.00")
        txtUtilidades3_1.Text = Format(oRS!nUtilidAnter1, "###,###,###,##0.00")
        txtVentasAnt3_2.Text = Format(oRS!nVentasAnter2, "###,###,###,##0.00")
        txtUtilidades3_2.Text = Format(oRS!nUtilidAnter2, "###,###,###,##0.00")
        'EEFF
        txtFecha3.Text = Format(oRS!dFechaEEFF, "DD/MM/YYYY")
        txtVentasEEFF3.Text = Format(oRS!nVentasActual, "###,###,###,##0.00")
        txtUtilidadesEEFF3.Text = Format(oRS!nUtilidActual, "###,###,###,##0.00")
        txtPorcentaje3.Text = Format(oRS!nPorcentaje, "###,###,###,##0.00")
        'Condicion Sector
        txtCondicionSector3_1.Text = oRS!cCondicionS
        'Calidad Evaluación
        txtCalidadEvaluacion3_1.Text = oRS!cCalidadEvS
        
        'LUCV20160909, Comentó según ERS052-2016
'        If lnExposicionTotal_1 = -1 Then
'            txtPasivo3.Text = oRs!nPasivo
'            If Len(txtPatrimonio3.Text) > 0 And txtPatrimonio3.Text > 0 Then
'                txtApalan3.Text = Format(oRs!nApalancamiento, "###,###,###,##0.00")
'
'            Else
'                txtApalan3.Text = 0#
'            End If
'        Else
'            txtApalan3.Text = 0#
'            txtPasivo3.Text = 0#
'        End If
        'Fin LUCV20160909
        
        '*****-> Formatos de Evaluación Consumo
        txtEvaluacion3.Text = oRS!cEvaluacioS 'Evaluación económica
        txtPasivo3_2.Text = Format(oRS!nPasivo, "###,###,###,##0.00")
        txtPatrimonio3_2.Text = Format(oRS!nPatrimonio, "###,###,###,##0.00")
        txtLineaCredito3_2.Text = Format(oRS!nLineaCred, "###,###,###,##0.00")
        'Resultados
        txtSaldoDis3_2.Text = Format(oRS!nSaldoDisponible, "###,###,###,##0.00")
        txtCapPag3_2.Text = Format(oRS!nCapaPago, "##0.00") '*100
        txtApalancamiento3_2.Text = Format(oRS!nApalancamiento, "###,###,###,##0.00")
        txtSensibilizacion3_2.Text = Format(oRS!nSensibli, "###,###,###,##0.00")
        txtMoraSector3_2.Text = Format(oRS!nMoraDelSector, "###,###,###,##0.00")
        'Declaración SUNAT - Regimen general
        txtPromedio3_2.Text = Format(oRS!nPromedio, "###,###,###,##0.00")
        txtMesAnt3_2_1.Text = Format(oRS!nMesAnterior1, "###,###,###,##0.00")
        txtMesAnt3_2_2.Text = Format(oRS!nMesAnterior2, "###,###,###,##0.00")
        txtMesAnt3_2_3.Text = Format(oRS!nMesAnterior3, "###,###,###,##0.00")
        'Condición sector
        txtCondicionSector3_2.Text = oRS!cCondicionS
        'Calidad Evaluación
        txtCalidadEvaluacion3_2.Text = oRS!cCalidadEvS
        
        '->*****LUCV20160916, Comentó: Según ERS052-2016
'        If lnExposicionTotal_1 = -1 Then
'            txtPasivo3_2.Text = oRs!nPasivo
'            If Len(txtPatrimonio3_2.Text) > 0 And txtPatrimonio3_2.Text > 0 Then
'                txtApalancamiento3_2.Text = Format(oRs!nApalancamiento, "###,###,###,##0.00")
'            Else
'                txtApalancamiento3_2.Text = 0
'            End If
'        Else
'            txtApalancamiento3_2.Text = 0#
'            txtPasivo3_2.Text = 0#
'        End If
        '<-***** Fin LUCV20160916
        oRS.MoveNext
    Loop
    End If
End Sub

Private Sub txtApalancamiento3_2_LostFocus()
    With txtApalancamiento3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtCalidadEvaluacion3_1_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub


Private Sub txtCalidadEvaluacion3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub

Private Sub txtCapPag3_2_LostFocus()
    With txtCapPag3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtConclusiones4_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub


Private Sub txtCondicionSector3_1_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub


Private Sub txtCondicionSector3_2_KeyPress(KeyAscii As Integer)
KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub

Private Sub txtEvaluacion3_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub

Private Sub txtEvolDeudaSF1_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub

Private Sub txtHistCMACM1_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub

Private Sub txtLineaCred3_ActualizaPatrimonio(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If val(txtPatrimonio3.Text) > 0 Then
            txtApalan3.Text = Round(IIf(CDbl(IIf(Len(txtPatrimonio3.Text) = 0, 0, txtPatrimonio3.Text)) = 0, 0, (lnExposicionTotal_1 + lnPasivoCorriente_1 + CDbl(IIf(Len(txtLineaCred3.Text) = 0, 0, txtLineaCred3.Text))) / CDbl(IIf(Len(txtPatrimonio3.Text) = 0, 0, txtPatrimonio3.Text))), 2) * 100
        End If
        txtPasivo3.Text = lnPasivo + CDbl(IIf(Len(txtLineaCred3.Text) = 0, 0, txtLineaCred3.Text))
    End If
End Sub

Private Sub txtLineaCredito3_2_ActualizaPatrimonio(KeyAscii As Integer)
If KeyAscii = 13 Then
        'txtApalan3.Text = IIf(CDbl(IIf(Len(txtPatrimonio3.Text) = 0, 0, txtPatrimonio3.Text)) = 0, 0, (lnExposicionTotal_1 + lnPasivoCorriente_1 + CDbl(IIf(Len(txtLineaCred3.Text) = 0, 0, txtLineaCred3.Text))) / CDbl(IIf(Len(txtPatrimonio3.Text) = 0, 0, txtPatrimonio3.Text)))
        'txtPasivo3.Text = lnPasivo + CDbl(IIf(Len(txtLineaCred3.Text) = 0, 0, txtLineaCred3.Text))
        'Consumo
        txtPasivo3_2.Text = lnPasivo + CDbl(IIf(Len(txtLineaCred3.Text) = 0, 0, txtLineaCred3.Text))
        If Len(txtPatrimonio3_2.Text) > 0 And txtPatrimonio3_2.Text > 0 Then
            txtApalancamiento3_2.Text = IIf(CDbl(IIf(Len(txtPatrimonio3.Text) = 0, 0, txtPatrimonio3_2.Text)) = 0, 0, (lnExposicionTotal_1 + lnPasivoCorriente_1 + CDbl(IIf(Len(txtLineaCredito3_2.Text) = 0, 0, txtLineaCredito3_2.Text))) / CDbl(IIf(Len(txtPatrimonio3_2.Text) = 0, 0, txtPatrimonio3_2.Text)))
        Else
            txtApalancamiento3_2.Text = 0
        End If
    End If
End Sub

Private Sub ConsNivelRiesgos()
Dim oRsCons As ADODB.Recordset
Set oRsCons = New ADODB.Recordset

Dim ObjCons As COMDConstantes.DCOMConstantes
Set ObjCons = New COMDConstantes.DCOMConstantes

Set oRsCons = ObjCons.RecuperaConstantes(gConsNivelRiesgos)
If Not (oRsCons.BOF Or oRsCons.EOF) Then
    Do While Not oRsCons.EOF
        cmdNivelRiesgo4.AddItem oRsCons!cConsDescripcion & Space(100) & oRsCons!nConsValor
        oRsCons.MoveNext
    Loop
End If
End Sub

Private Sub Cuadro4(ByVal psCtaCod As String)
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    'Set clsTC = New COMDConstSistema.NCOMTipoCambio
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim objDPersonas As COMDPersona.DCOMPersonas
    Dim oCredito As COMDCredito.DCOMCredito 'WIOR 20141017
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    
    Dim oRs2 As ADODB.Recordset
    Set oRs2 = New ADODB.Recordset
    
    Set objDPersona = New COMDPersona.DCOMPersona
    Set objDPersonas = New COMDPersona.DCOMPersonas
    Set oRs2 = objDPersonas.dDatosPersonas(gsCodPersUser)
    
'    FEGarantias2.Col = 5
'    FEGarantias2.CargaCombo CargarCobertura
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro4(psCtaCod)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
         cmdNivelRiesgo4.ListIndex = IndiceListaCombo(cmdNivelRiesgo4, oRS!nNivelRiesgo)
         txtConclusiones4.Text = oRS!cConclusione
         txtRecomendaciones4.Text = oRS!cRecomendaci
         txtAnalisis4.Text = oRS!cPersNombre
         txtPersCodAnalista4.Text = oRS!cAnalisisRie
         txtFecha4.Text = Format(oRS!dFechaNR, "DD/MM/YYYY")
         txtInforme.Text = oRS!cNumeroInforme
        oRS.MoveNext
    Loop
    End If
    If Not (oRs2.BOF Or oRs2.EOF) Then
    If Trim(txtAnalisis4.Text) = "" Then
        txtAnalisis4.Text = oRs2!cPersNombre
        txtPersCodAnalista4.Text = oRs2!cPersCod
    End If
    End If
    
    'WIOR 20141017 *****************
    Set oCredito = New COMDCredito.DCOMCredito
    Set oRS = oCredito.ObtenerInformeRiesgo(Trim(psCtaCod), 1)
    
    If Not (oRS.EOF And oRS.BOF) Then
        Me.txtConclusionGen.Text = oRS!Glosa
    End If
     Set oCredito = Nothing
    Set oRS = Nothing
    'WIOR FIN **********************
End Sub
'
Private Sub Cuadro3EvaluacionConsumo(ByVal psCtaCod As String)
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro3EvaluacionConsumo(psCtaCod)
    FormateaFlex FEEvaluacionEco3
    If lnCodForm = 7 Or lnCodForm = 8 Then 'LUCV20160906
        If Not (oRS.BOF Or oRS.EOF) Then
        Do While Not oRS.EOF
            FEEvaluacionEco3.AdicionaFila
            FEEvaluacionEco3.TextMatrix(oRS.Bookmark, 0) = psCtaCod
            FEEvaluacionEco3.TextMatrix(oRS.Bookmark, 1) = oRS!cTipoEval
            FEEvaluacionEco3.TextMatrix(oRS.Bookmark, 2) = oRS!cTituloEval
            FEEvaluacionEco3.TextMatrix(oRS.Bookmark, 3) = oRS!CDescripcion
            FEEvaluacionEco3.TextMatrix(oRS.Bookmark, 4) = Format(oRS!nPersonal, "###,###,###,##0.00")
            'FEEvaluacionEco3.TextMatrix(oRs.Bookmark, 5) = Format(oRs!nNegocio, "###,###,###,##0.00") 'Comentado por: LUCV20160906, Según ERS052-2016
            'FEEvaluacionEco3.TextMatrix(oRs.Bookmark, 6) = Format(oRs!nUnico, "###,###,###,##0.00")   'Comentado por: LUCV20160906, Según ERS052-2016
            oRS.MoveNext
        Loop
        End If
    End If
End Sub
Private Function ValidarDatos() As String
Dim lsmensaje As String
Dim i As Integer
If Trim(txtHistCMACM1) = "" Then
    lsmensaje = "Ingresar el historial Crediticio en la CMACM"
    ValidarDatos = lsmensaje
    SSTab1.Tab = 0
    txtHistCMACM1.SetFocus
    Exit Function
End If
If InStr(Trim(txtHistCMACM1), "'") > 0 Then
    lsmensaje = "El caracter ' no es permitido para el historial Crediticio en la CMACM"
    ValidarDatos = lsmensaje
    SSTab1.Tab = 0
    txtHistCMACM1.SetFocus
    Exit Function
End If

If Trim(txtEvolDeudaSF1) = "" Then
    lsmensaje = "Ingresar la evolución deuda sistema financiero"
    ValidarDatos = lsmensaje
    SSTab1.Tab = 0
    txtEvolDeudaSF1.SetFocus
    Exit Function
End If
If InStr(Trim(txtEvolDeudaSF1), "'") > 0 Then
    lsmensaje = "El caracter ' no es permitido para la evolución deuda sistema financiero"
    ValidarDatos = lsmensaje
    SSTab1.Tab = 0
    txtEvolDeudaSF1.SetFocus
    Exit Function
End If

'WIOR 20150724 ***
For i = 1 To FEVinculados1.rows - 1
     If Len(FEVinculados1.TextMatrix(i, 0)) > 0 Then
        If InStr(Trim(FEVinculados1.TextMatrix(i, 6)), "'") > 0 Then
            lsmensaje = "El caracter ' no es permitido para la Evolución de los vinculados de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            FEVinculados1.SetFocus
            Exit Function
        End If
     End If
Next i
'WIOR FIN ********

'WIOR 20141016 ************
For i = 1 To feEmpresasVinc.rows - 1
    If Len(feEmpresasVinc.TextMatrix(i, 0)) > 0 Then
        If Trim(feEmpresasVinc.TextMatrix(i, 1)) = "" Then
            lsmensaje = "Seleccione a la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Trim(feEmpresasVinc.TextMatrix(i, 3)) = "" And Trim(feEmpresasVinc.TextMatrix(i, 4)) = "" Then
            lsmensaje = "Seleccione el vinculado o grupo economico a la que esta relacionada la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Trim(feEmpresasVinc.TextMatrix(i, 5)) = "" Then
            lsmensaje = "Ingrese la relacion a la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Trim(feEmpresasVinc.TextMatrix(i, 6)) = "" Then
            lsmensaje = "Ingrese el endeudamiento de la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Not IsNumeric(Trim(feEmpresasVinc.TextMatrix(i, 6))) Then
            lsmensaje = "Ingrese correctamente el endeudamiento de la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Trim(feEmpresasVinc.TextMatrix(i, 7)) = "" Then
            lsmensaje = "Ingrese el Nro de IFIS de la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Not IsNumeric(Trim(feEmpresasVinc.TextMatrix(i, 7))) Then
            lsmensaje = "Ingrese correctamente el Nro de IFIS de la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        If Trim(feEmpresasVinc.TextMatrix(i, 9)) = "" Then
            lsmensaje = "Ingrese la Evolución de la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        
        'WIOR 20150724 ***
        If InStr(Trim(feEmpresasVinc.TextMatrix(i, 9)), "'") > 0 Then
            lsmensaje = "El caracter ' no es permitido para la Evolución de la empresa vinculada de la fila " & i
            ValidarDatos = lsmensaje
            SSTab1.Tab = 0
            feEmpresasVinc.SetFocus
            Exit Function
        End If
        'WIOR FIN ********
    End If
Next i
'WIOR FIN *****************

For i = 1 To FEGarantias2.rows - 1
    If Trim(Right(FEGarantias2.TextMatrix(i, 5), 5)) = "" Then
        lsmensaje = "Ingresar la cobertura de la Garantia"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 1
        Exit Function
    End If
Next i
'If Not ((Left(txtTpoCredCod1.Text, 1) = "1" Or Left(txtTpoCredCod1.Text, 1) = "2" Or Left(txtTpoCredCod1.Text, 1) = "3" Or Left(txtTpoCredCod1.Text, 1) = "4" Or Left(txtTpoCredCod1.Text, 1) = "5")) Then
If Not (lnCodForm = 4 Or lnCodForm = 5 Or lnCodForm = 6 Or lnCodForm = 9) Then 'Formato->No minoristas
    If Len(Trim(txtPasivo3_2.Text)) = 0 Then
        lsmensaje = "Ingresar el valor del pasivo"
        SSTab1.Tab = 2
        txtPasivo3_2.SetFocus
        ValidarDatos = lsmensaje
        Exit Function
    End If
    If Len(Trim(txtApalancamiento3_2.Text)) = 0 Then
        lsmensaje = "Ingresar el valor del apalancamiento"
        SSTab1.Tab = 2
        txtApalancamiento3_2.SetFocus
        ValidarDatos = lsmensaje
        Exit Function
    End If
    If Len(Trim(txtLineaCredito3_2.Text)) = 0 Then
        lsmensaje = "Ingresar el valor del Linea de Credito"
        SSTab1.Tab = 2
        txtLineaCredito3_2.SetFocus
        ValidarDatos = lsmensaje
        Exit Function
    End If
    If Len(Trim(txtLineaCredito3_2.Text)) = 0 Then
        lsmensaje = "Ingresar el valor del Linea de Credito"
        SSTab1.Tab = 2
        txtLineaCredito3_2.SetFocus
        ValidarDatos = lsmensaje
        Exit Function
    End If
'    If Trim(txtEvaluacion3.Text) = "" Then
'          lsMensaje = "Favor ingresar a evaluacion del informe de riesgo"
'          ValidarDatos = lsMensaje
'          SSTab1.Tab = 2
'          txtEvaluacion3.SetFocus
'          Exit Function
'      End If
      If InStr(Trim(txtEvaluacion3), "'") > 0 Then
          lsmensaje = "El caracter ' no es permitido para la evaluacion del informe de riesgo"
          ValidarDatos = lsmensaje
          SSTab1.Tab = 2
          txtEvaluacion3.SetFocus
          Exit Function
      End If
      If Trim(txtCondicionSector3_2.Text) = "" Then
          lsmensaje = "Favor ingresar el texto de la condición sector del informe del riesgo"
          ValidarDatos = lsmensaje
          SSTab1.Tab = 2
          txtCondicionSector3_2.SetFocus
          Exit Function
      End If
      If InStr(Trim(txtCondicionSector3_2), "'") > 0 Then
          lsmensaje = "El caracter ' no es permitido para la condición sector del informe de riesgo"
          ValidarDatos = lsmensaje
          SSTab1.Tab = 2
          txtCondicionSector3_2.SetFocus
          Exit Function
      End If
      If Trim(txtCalidadEvaluacion3_2.Text) = "" Then
          lsmensaje = "Favor ingresar el texto de la calidad de evaluacion del informe del riesgo"
          ValidarDatos = lsmensaje
          SSTab1.Tab = 2
          txtCalidadEvaluacion3_2.SetFocus
          Exit Function
      End If
      If InStr(Trim(txtCalidadEvaluacion3_2), "'") > 0 Then
          lsmensaje = "El caracter ' no es permitido para la calidad de evaluacion del informe del riesgo"
          ValidarDatos = lsmensaje
          SSTab1.Tab = 2
          txtCalidadEvaluacion3_2.SetFocus
          Exit Function
      End If
Else
    '/////
    If Trim(txtCondicionSector3_1) = "" Then
        lsmensaje = "Ingresar la condición en el sector"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 2
        txtCondicionSector3_1.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtCondicionSector3_1), "'") > 0 Then
        lsmensaje = "El caracter ' no es permitido para la condicion en el sector"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 2
        txtCondicionSector3_1.SetFocus
        Exit Function
    End If
    If Trim(txtCalidadEvaluacion3_1) = "" Then
        lsmensaje = "Ingresar la calidad de evaluacion"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 2
        txtCalidadEvaluacion3_1.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtCalidadEvaluacion3_1), "'") > 0 Then
        lsmensaje = "El caracter ' no es permitido para la calidad de la evaluación"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 2
        txtCalidadEvaluacion3_1.SetFocus
        Exit Function
    End If
    If (txtFecha3.Text = "__/__/____") Then
        lsmensaje = "Favor ingresar correctamente la Fecha de la evaluacion económica-financiera"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 2
        txtFecha3.SetFocus
        Exit Function
    End If
End If
    '/////
    If Trim(cmdNivelRiesgo4.Text) = "" Then
        lsmensaje = "Favor seleccionar el nivel de Riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        cmdNivelRiesgo4.SetFocus
        Exit Function
    End If
    If Trim(txtInforme.Text) = "" Then
        lsmensaje = "Favor ingresar el numero del informe de riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtInforme.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtInforme), "'") > 0 Then
        lsmensaje = "El caracter ' no es permitido para el numero del informe de riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtInforme.SetFocus
        Exit Function
    End If
    If Trim(txtConclusiones4.Text) = "" Then
        lsmensaje = "Favor ingresar las conclusiones"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtConclusiones4.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtConclusiones4), "'") > 0 Then
        lsmensaje = "El caracter ' no es permitido para las conclusiones del informe de riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtConclusiones4.SetFocus
        Exit Function
    End If
    If Trim(txtRecomendaciones4.Text) = "" Then
        lsmensaje = "Favor ingresar las recomendaciones"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtRecomendaciones4.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtRecomendaciones4), "'") > 0 Then
        lsmensaje = "El caracter ' no es permitido para las recomendaciones del informe de riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtRecomendaciones4.SetFocus
        Exit Function
    End If
    If Trim(txtPersCodAnalista4.Text) = "" Then
        lsmensaje = "Favor ingresar el codigo del analista de riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtRecomendaciones4.SetFocus
        Exit Function
    End If
    'WIOR 20141016 ********************************
    If Trim(txtConclusionGen.Text) = "" Then
        lsmensaje = "Favor Ingrese la Conclusión General"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtConclusionGen.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtConclusionGen.Text), "'") > 0 Then
        lsmensaje = "El caracter ' no es permitido para la Conclusion General del Informe de riesgo"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtConclusionGen.SetFocus
        Exit Function
    End If
    'WIOR FIN **************************************
If (txtFecha4.Text = "__/__/____") Then
    lsmensaje = "Favor ingresar correctamente la Fecha del informe de riesgos"
        ValidarDatos = lsmensaje
        SSTab1.Tab = 3
        txtFecha4.SetFocus
        Exit Function
End If
ValidarDatos = lsmensaje
End Function

Private Function SumarPagadasCMAC() As Currency
    Dim i As Integer
    Dim nMontoCMAMC As Currency
    nMontoCMAMC = 0
    For i = 1 To FECuoPro2.rows - 1
    If FECuoPro2.TextMatrix(i, 1) <> "" Then
        nMontoCMAMC = nMontoCMAMC + FECuoPro2.TextMatrix(i, 1)
    End If
    Next i
    SumarPagadasCMAC = nMontoCMAMC
End Function

Private Function SumarPagadasIFIS() As Currency
    Dim i As Integer
    Dim nMontoIFIS As Currency
    nMontoIFIS = 0
    For i = 1 To FEEsCuaMa2.rows - 1
        If FEEsCuaMa2.TextMatrix(i, 1) <> "" Then
            nMontoIFIS = nMontoIFIS + FEEsCuaMa2.TextMatrix(i, 1)
        End If
    Next i
    SumarPagadasIFIS = nMontoIFIS
'    Call EscalonamientoCuotas
End Function
Private Function SumarCreditosPagadasCMAC() As Currency
    Dim i As Integer
    Dim nMontoCMAMC As Currency
    nMontoCMAMC = 0
    For i = 1 To FEMontoPropuesto2.rows - 1
    If FEMontoPropuesto2.TextMatrix(i, 1) <> "" Then
        nMontoCMAMC = nMontoCMAMC + FEMontoPropuesto2.TextMatrix(i, 1)
    End If
    Next i
    SumarCreditosPagadasCMAC = nMontoCMAMC
End Function

Private Function SumarCreditosPagadasIFIS() As Currency
    Dim i As Integer
    Dim nMontoIFIS As Currency
    nMontoIFIS = 0
    For i = 1 To FEMontoCredito2.rows - 1
        If FEMontoCredito2.TextMatrix(i, 1) <> "" Then
            nMontoIFIS = nMontoIFIS + FEMontoCredito2.TextMatrix(i, 1)
        End If
    Next i
    SumarCreditosPagadasIFIS = nMontoIFIS
End Function
Public Function SoloCaracteresValidos(intTecla As Integer, _
                           Optional lbMayusculas As Boolean = False) As Integer
Dim cValidar  As String
    'cValidar = "'<>?_=+[]{}|!@#$%ç¨-,´`¡¿Çºª""·"
    cValidar = "'"
    If intTecla > 26 Then
        If InStr(cValidar, Chr(intTecla)) <> 0 Then
            intTecla = 0
            Beep
        End If
    End If
    If lbMayusculas Then
        SoloCaracteresValidos = Asc((Chr(intTecla)))
    Else
        SoloCaracteresValidos = intTecla
    End If
End Function

Private Sub txtLineaCred3_LostFocus()
    With txtLineaCred3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtLineaCredito3_2_LostFocus()
    With txtLineaCredito3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtMesAnt3_1_1_LostFocus()
    With txtMesAnt3_1_1
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtMesAnt3_1_2_LostFocus()
    With txtMesAnt3_1_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtMesAnt3_1_3_LostFocus()
    With txtMesAnt3_1_3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtMesAnt3_2_1_LostFocus()
    With txtMesAnt3_2_1
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtMesAnt3_2_2_LostFocus()
    With txtMesAnt3_2_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With

End Sub

Private Sub txtMesAnt3_2_3_LostFocus()
    With txtMesAnt3_2_3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtMoraSector3_2_LostFocus()
    With txtMoraSector3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtOtrosIngr3_LostFocus()
    With txtOtrosIngr3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtPasivo3_2_LostFocus()
    With txtPasivo3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtPatrimonio3_2_LostFocus()
    With txtPatrimonio3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
      Call Cuadro3Calculo(1) 'LUCV20160915
End Sub

Private Sub txtPatrimonio3_KeyPress(KeyAscii As Integer) 'LUCV20160919
    KeyAscii = NumerosDecimales(txtPatrimonio3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_1_1.SetFocus
    End If
End Sub

Private Sub txtPatrimonio3_LostFocus() 'LUCV20160919
     With txtPatrimonio3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
      Call Cuadro3Calculo(2)
End Sub

Private Sub txtPorcentaje3_LostFocus()
    With txtPorcentaje3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtPromedio3_2_LostFocus()
    With txtPromedio3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtPromedio3_LostFocus()
    With txtPromedio3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtRecomendaciones4_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloCaracteresValidos(KeyAscii, True)
End Sub

Private Function TienePunto(psCadena As String) As Boolean
If InStr(1, psCadena, ".", vbTextCompare) > 0 Then
    TienePunto = True
Else
    TienePunto = False
End If
End Function

Private Function NumDecimal(psCadena As String) As Integer
Dim lnPos As Integer
lnPos = InStr(1, psCadena, ".", vbTextCompare)
If lnPos > 0 Then
    NumDecimal = Len(psCadena) - lnPos
Else
    NumDecimal = 0
End If
End Function

Private Sub txtMesAnt3_2_1_Change()
txtMesAnt3_2_1.SelStart = Len(txtMesAnt3_2_1)
gnNumDec = NumDecimal(txtMesAnt3_2_1)
If gbEstado And txtMesAnt3_2_1 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMesAnt3_2_1 = Format(txtMesAnt3_2_1, "#,###,###,##0")
        Case 1
                txtMesAnt3_2_1 = Format(txtMesAnt3_2_1, "#,###,###,##0.0")
        Case 2
                txtMesAnt3_2_1 = Format(txtMesAnt3_2_1, "#,###,###,##0.00")
        Case 3
                txtMesAnt3_2_1 = Format(txtMesAnt3_2_1, "#,###,###,##0.000")
        Case Else
                txtMesAnt3_2_1 = Format(txtMesAnt3_2_1, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
txtPromedio3_2.Text = Format(PromedioMes3(IIf(txtMesAnt3_2_1 = "", 0, txtMesAnt3_2_1), IIf(txtMesAnt3_2_2 = "", 0, txtMesAnt3_2_2), IIf(txtMesAnt3_2_3 = "", 0, txtMesAnt3_2_3)), "#,###,###,##0.00")
End Sub
Private Sub txtMesAnt3_2_1_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMesAnt3_2_1, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_2_2.SetFocus
    End If
End Sub

Private Sub txtMesAnt3_2_1_GotFocus()
With txtMesAnt3_2_1
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMesAnt3_2_2_Change()
txtMesAnt3_2_2.SelStart = Len(txtMesAnt3_2_2)
gnNumDec = NumDecimal(txtMesAnt3_2_2)
If gbEstado And txtMesAnt3_2_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMesAnt3_2_2 = Format(txtMesAnt3_2_2, "#,###,###,##0")
        Case 1
                txtMesAnt3_2_2 = Format(txtMesAnt3_2_2, "#,###,###,##0.0")
        Case 2
                txtMesAnt3_2_2 = Format(txtMesAnt3_2_2, "#,###,###,##0.00")
        Case 3
                txtMesAnt3_2_2 = Format(txtMesAnt3_2_2, "#,###,###,##0.000")
        Case Else
                txtMesAnt3_2_2 = Format(txtMesAnt3_2_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
txtPromedio3_2.Text = Format(PromedioMes3(IIf(txtMesAnt3_2_1 = "", 0, txtMesAnt3_2_1), IIf(txtMesAnt3_2_2 = "", 0, txtMesAnt3_2_2), IIf(txtMesAnt3_2_3 = "", 0, txtMesAnt3_2_3)), "#,###,###,##0.00")
End Sub
Private Sub txtMesAnt3_2_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMesAnt3_2_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_2_3.SetFocus
    End If
End Sub

Private Sub txtMesAnt3_2_2_GotFocus()
With txtMesAnt3_2_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMesAnt3_2_3_Change()
txtMesAnt3_2_3.SelStart = Len(txtMesAnt3_2_3)
gnNumDec = NumDecimal(txtMesAnt3_2_3)
If gbEstado And txtMesAnt3_2_3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMesAnt3_2_3 = Format(txtMesAnt3_2_3, "#,###,###,##0")
        Case 1
                txtMesAnt3_2_3 = Format(txtMesAnt3_2_3, "#,###,###,##0.0")
        Case 2
                txtMesAnt3_2_3 = Format(txtMesAnt3_2_3, "#,###,###,##0.00")
        Case 3
                txtMesAnt3_2_3 = Format(txtMesAnt3_2_3, "#,###,###,##0.000")
        Case Else
                txtMesAnt3_2_3 = Format(txtMesAnt3_2_3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
txtPromedio3_2.Text = Format(PromedioMes3(IIf(txtMesAnt3_2_1 = "", 0, txtMesAnt3_2_1), IIf(txtMesAnt3_2_2 = "", 0, txtMesAnt3_2_2), IIf(txtMesAnt3_2_3 = "", 0, txtMesAnt3_2_3)), "#,###,###,##0.00")
End Sub
Private Sub txtMesAnt3_2_3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMesAnt3_2_3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_2_3.SetFocus
    End If
End Sub

Private Sub txtMesAnt3_2_3_GotFocus()
With txtMesAnt3_2_3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Function PromedioMes3(ByVal nMes1 As Currency, ByVal nMes2 As Currency, ByVal nMes3 As Currency)
            PromedioMes3 = (nMes1 + nMes2 + nMes3) / 3
End Function
Private Sub txtPasivo3_2_Change()
txtPasivo3_2.SelStart = Len(txtPasivo3_2)
gnNumDec = NumDecimal(txtPasivo3_2)
If gbEstado And txtPasivo3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtPasivo3_2 = Format(txtPasivo3_2, "#,###,###,##0")
        Case 1
                txtPasivo3_2 = Format(txtPasivo3_2, "#,###,###,##0.0")
        Case 2
                txtPasivo3_2 = Format(txtPasivo3_2, "#,###,###,##0.00")
        Case 3
                txtPasivo3_2 = Format(txtPasivo3_2, "#,###,###,##0.000")
        Case Else
                txtPasivo3_2 = Format(txtPasivo3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtPasivo3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPasivo3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtPatrimonio3_2.SetFocus
    End If
End Sub

Private Sub txtPasivo3_2_GotFocus()
With txtPasivo3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtPatrimonio3_2_Change()
txtPatrimonio3_2.SelStart = Len(txtPatrimonio3_2)
gnNumDec = NumDecimal(txtPatrimonio3_2)
If gbEstado And txtPatrimonio3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtPatrimonio3_2 = Format(txtPatrimonio3_2, "#,###,###,##0")
        Case 1
                txtPatrimonio3_2 = Format(txtPatrimonio3_2, "#,###,###,##0.0")
        Case 2
                txtPatrimonio3_2 = Format(txtPatrimonio3_2, "#,###,###,##0.00")
        Case 3
                txtPatrimonio3_2 = Format(txtPatrimonio3_2, "#,###,###,##0.000")
        Case Else
                txtPatrimonio3_2 = Format(txtPatrimonio3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
If (txtPatrimonio3_2.Text) <> "" Then  'LUCV20160916
    If (txtPatrimonio3_2.Text) <= 0 Then
        txtApalancamiento3_2.Enabled = False
    Else
        txtApalancamiento3_2.Enabled = True
    End If
Else
        txtApalancamiento3_2.Enabled = False
End If
RaiseEvent Change
End Sub
Private Sub txtPatrimonio3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPatrimonio3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
       'txtLineaCredito3_2.SetFocus 'LUCV20160916, Comentó
       txtMesAnt3_2_1.SetFocus
    End If
End Sub

Private Sub txtPatrimonio3_2_GotFocus()
With txtPatrimonio3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtLineaCredito3_2_Change()
txtLineaCredito3_2.SelStart = Len(txtLineaCredito3_2)
gnNumDec = NumDecimal(txtLineaCredito3_2)
If gbEstado And txtLineaCredito3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtLineaCredito3_2 = Format(txtLineaCredito3_2, "#,###,###,##0")
        Case 1
                txtLineaCredito3_2 = Format(txtLineaCredito3_2, "#,###,###,##0.0")
        Case 2
                txtLineaCredito3_2 = Format(txtLineaCredito3_2, "#,###,###,##0.00")
        Case 3
                txtLineaCredito3_2 = Format(txtLineaCredito3_2, "#,###,###,##0.000")
        Case Else
                txtLineaCredito3_2 = Format(txtLineaCredito3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtLineaCredito3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtLineaCredito3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        Call txtLineaCredito3_2_ActualizaPatrimonio(KeyAscii)
        txtSaldoDis3_2.SetFocus
    End If
End Sub

Private Sub txtLineaCredito3_2_GotFocus()
With txtLineaCredito3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtSaldoDis3_2_Change()
txtSaldoDis3_2.SelStart = Len(txtSaldoDis3_2)
gnNumDec = NumDecimal(txtSaldoDis3_2)
If gbEstado And txtSaldoDis3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtSaldoDis3_2 = Format(txtSaldoDis3_2, "#,###,###,##0")
        Case 1
                txtSaldoDis3_2 = Format(txtSaldoDis3_2, "#,###,###,##0.0")
        Case 2
                txtSaldoDis3_2 = Format(txtSaldoDis3_2, "#,###,###,##0.00")
        Case 3
                txtSaldoDis3_2 = Format(txtSaldoDis3_2, "#,###,###,##0.000")
        Case Else
                txtSaldoDis3_2 = Format(txtSaldoDis3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtSaldoDis3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtSaldoDis3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtCapPag3_2.SetFocus
    End If
End Sub

Private Sub txtSaldoDis3_2_GotFocus()
With txtSaldoDis3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtCapPag3_2_Change()
txtCapPag3_2.SelStart = Len(txtCapPag3_2)
gnNumDec = NumDecimal(txtCapPag3_2)
If gbEstado And txtCapPag3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtCapPag3_2 = Format(txtCapPag3_2, "#,###,###,##0")
        Case 1
                txtCapPag3_2 = Format(txtCapPag3_2, "#,###,###,##0.0")
        Case 2
                txtCapPag3_2 = Format(txtCapPag3_2, "#,###,###,##0.00")
        Case 3
                txtCapPag3_2 = Format(txtCapPag3_2, "#,###,###,##0.000")
        Case Else
                txtCapPag3_2 = Format(txtCapPag3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtCapPag3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCapPag3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        'txtApalancamiento3_2.SetFocus 'LUCV20160916
    End If
End Sub

Private Sub txtCapPag3_2_GotFocus()
With txtCapPag3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtApalancamiento3_2_Change()
txtApalancamiento3_2.SelStart = Len(txtApalancamiento3_2)
gnNumDec = NumDecimal(txtApalancamiento3_2)
If gbEstado And txtApalancamiento3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtApalancamiento3_2 = Format(txtApalancamiento3_2, "#,###,###,##0")
        Case 1
                txtApalancamiento3_2 = Format(txtApalancamiento3_2, "#,###,###,##0.0")
        Case 2
                txtApalancamiento3_2 = Format(txtApalancamiento3_2, "#,###,###,##0.00")
        Case 3
                txtApalancamiento3_2 = Format(txtApalancamiento3_2, "#,###,###,##0.000")
        Case Else
                txtApalancamiento3_2 = Format(txtApalancamiento3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
'Private Sub txtApalancamiento3_2_KeyPress(KeyAscii As Integer) 'LUCV20160615
'    KeyAscii = NumerosDecimales(txtApalancamiento3_2, KeyAscii, , 4)
'    If KeyAscii = 13 Then
'        txtSensibilizacion3_2.SetFocus
'    End If
'End Sub

Private Sub txtApalancamiento3_2_GotFocus()
With txtApalancamiento3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtSaldoDis3_2_LostFocus()
    With txtSaldoDis3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtSensibilizacion3_2_Change()
txtSensibilizacion3_2.SelStart = Len(txtSensibilizacion3_2)
gnNumDec = NumDecimal(txtSensibilizacion3_2)
If gbEstado And txtSensibilizacion3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtSensibilizacion3_2 = Format(txtSensibilizacion3_2, "#,###,###,##0")
        Case 1
                txtSensibilizacion3_2 = Format(txtSensibilizacion3_2, "#,###,###,##0.0")
        Case 2
                txtSensibilizacion3_2 = Format(txtSensibilizacion3_2, "#,###,###,##0.00")
        Case 3
                txtSensibilizacion3_2 = Format(txtSensibilizacion3_2, "#,###,###,##0.000")
        Case Else
                txtSensibilizacion3_2 = Format(txtSensibilizacion3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtSensibilizacion3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtSensibilizacion3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMoraSector3_2.SetFocus
    End If
End Sub

Private Sub txtSensibilizacion3_2_GotFocus()
With txtSensibilizacion3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtMoraSector3_2_Change()
txtMoraSector3_2.SelStart = Len(txtMoraSector3_2)
gnNumDec = NumDecimal(txtMoraSector3_2)
If gbEstado And txtMoraSector3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMoraSector3_2 = Format(txtMoraSector3_2, "#,###,###,##0")
        Case 1
                txtMoraSector3_2 = Format(txtMoraSector3_2, "#,###,###,##0.0")
        Case 2
                txtMoraSector3_2 = Format(txtMoraSector3_2, "#,###,###,##0.00")
        Case 3
                txtMoraSector3_2 = Format(txtMoraSector3_2, "#,###,###,##0.000")
        Case Else
                txtMoraSector3_2 = Format(txtMoraSector3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtMoraSector3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMoraSector3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMoraSector3_2.SetFocus
    End If
End Sub

Private Sub txtMoraSector3_2_GotFocus()
With txtMoraSector3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
'////
Private Sub txtMesAnt3_1_1_Change()
txtMesAnt3_1_1.SelStart = Len(txtMesAnt3_1_1)
gnNumDec = NumDecimal(txtMesAnt3_1_1)
If gbEstado And txtMesAnt3_1_1 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMesAnt3_1_1 = Format(txtMesAnt3_1_1, "#,###,###,##0")
        Case 1
                txtMesAnt3_1_1 = Format(txtMesAnt3_1_1, "#,###,###,##0.0")
        Case 2
                txtMesAnt3_1_1 = Format(txtMesAnt3_1_1, "#,###,###,##0.00")
        Case 3
                txtMesAnt3_1_1 = Format(txtMesAnt3_1_1, "#,###,###,##0.000")
        Case Else
                txtMesAnt3_1_1 = Format(txtMesAnt3_1_1, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
txtPromedio3.Text = Format(PromedioMes3(IIf(txtMesAnt3_1_1 = "", 0, txtMesAnt3_1_1), IIf(txtMesAnt3_1_2 = "", 0, txtMesAnt3_1_2), IIf(txtMesAnt3_1_3 = "", 0, txtMesAnt3_1_3)), "#,###,###,##0.00")
End Sub
Private Sub txtMesAnt3_1_1_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMesAnt3_1_1, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_1_2.SetFocus
    End If
End Sub

Private Sub txtMesAnt3_1_1_GotFocus()
With txtMesAnt3_1_1
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMesAnt3_1_2_Change()
txtMesAnt3_1_2.SelStart = Len(txtMesAnt3_1_2)
gnNumDec = NumDecimal(txtMesAnt3_1_2)
If gbEstado And txtMesAnt3_1_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMesAnt3_1_2 = Format(txtMesAnt3_1_2, "#,###,###,##0")
        Case 1
                txtMesAnt3_1_2 = Format(txtMesAnt3_1_2, "#,###,###,##0.0")
        Case 2
                txtMesAnt3_1_2 = Format(txtMesAnt3_1_2, "#,###,###,##0.00")
        Case 3
                txtMesAnt3_1_2 = Format(txtMesAnt3_1_2, "#,###,###,##0.000")
        Case Else
                txtMesAnt3_1_2 = Format(txtMesAnt3_1_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
txtPromedio3.Text = Format(PromedioMes3(IIf(txtMesAnt3_1_1 = "", 0, txtMesAnt3_1_1), IIf(txtMesAnt3_1_2 = "", 0, txtMesAnt3_1_2), IIf(txtMesAnt3_1_3 = "", 0, txtMesAnt3_1_3)), "#,###,###,##0.00")
End Sub
Private Sub txtMesAnt3_1_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMesAnt3_1_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_1_2.SetFocus
    End If
End Sub

Private Sub txtMesAnt3_1_2_GotFocus()
With txtMesAnt3_1_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMesAnt3_1_3_Change()
txtMesAnt3_1_3.SelStart = Len(txtMesAnt3_1_3)
gnNumDec = NumDecimal(txtMesAnt3_1_3)
If gbEstado And txtMesAnt3_1_3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtMesAnt3_1_3 = Format(txtMesAnt3_1_3, "#,###,###,##0")
        Case 1
                txtMesAnt3_1_3 = Format(txtMesAnt3_1_3, "#,###,###,##0.0")
        Case 2
                txtMesAnt3_1_3 = Format(txtMesAnt3_1_3, "#,###,###,##0.00")
        Case 3
                txtMesAnt3_1_3 = Format(txtMesAnt3_1_3, "#,###,###,##0.000")
        Case Else
                txtMesAnt3_1_3 = Format(txtMesAnt3_1_3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
txtPromedio3.Text = Format(PromedioMes3(IIf(txtMesAnt3_1_1 = "", 0, txtMesAnt3_1_1), IIf(txtMesAnt3_1_2 = "", 0, txtMesAnt3_1_2), IIf(txtMesAnt3_1_3 = "", 0, txtMesAnt3_1_3)), "#,###,###,##0.00")
End Sub
Private Sub txtMesAnt3_1_3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMesAnt3_1_3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtMesAnt3_1_3.SetFocus
    End If
End Sub

Private Sub txtMesAnt3_1_3_GotFocus()
With txtMesAnt3_1_3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtSensibilizacion3_2_LostFocus()
    With txtSensibilizacion3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtUtilidades3_1_LostFocus()
    With txtUtilidades3_1
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtUtilidades3_2_LostFocus()
    With txtUtilidades3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtUtilidadesEEFF3_LostFocus()
    With txtUtilidadesEEFF3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtVentasAnt3_1_Change()
txtVentasAnt3_1.SelStart = Len(txtVentasAnt3_1)
gnNumDec = NumDecimal(txtVentasAnt3_1)
If gbEstado And txtVentasAnt3_1 <> "" Then
    Select Case gnNumDec
        Case 0
                txtVentasAnt3_1 = Format(txtVentasAnt3_1, "#,###,###,##0")
        Case 1
                txtVentasAnt3_1 = Format(txtVentasAnt3_1, "#,###,###,##0.0")
        Case 2
                txtVentasAnt3_1 = Format(txtVentasAnt3_1, "#,###,###,##0.00")
        Case 3
                txtVentasAnt3_1 = Format(txtVentasAnt3_1, "#,###,###,##0.000")
        Case Else
                txtVentasAnt3_1 = Format(txtVentasAnt3_1, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtVentasAnt3_1_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentasAnt3_1, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtUtilidades3_1.SetFocus
    End If
End Sub

Private Sub txtVentasAnt3_1_GotFocus()
With txtVentasAnt3_1
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtUtilidades3_1_Change()
txtUtilidades3_1.SelStart = Len(txtUtilidades3_1)
gnNumDec = NumDecimal(txtUtilidades3_1)
If gbEstado And txtUtilidades3_1 <> "" Then
    Select Case gnNumDec
        Case 0
                txtUtilidades3_1 = Format(txtUtilidades3_1, "#,###,###,##0")
        Case 1
                txtUtilidades3_1 = Format(txtUtilidades3_1, "#,###,###,##0.0")
        Case 2
                txtUtilidades3_1 = Format(txtUtilidades3_1, "#,###,###,##0.00")
        Case 3
                txtUtilidades3_1 = Format(txtUtilidades3_1, "#,###,###,##0.000")
        Case Else
                txtUtilidades3_1 = Format(txtUtilidades3_1, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtUtilidades3_1_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtUtilidades3_1, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtVentasAnt3_2.SetFocus
    End If
End Sub

Private Sub txtUtilidades3_1_GotFocus()
With txtUtilidades3_1
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVentasAnt3_1_LostFocus()
    With txtVentasAnt3_1
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtVentasAnt3_2_Change()
txtVentasAnt3_2.SelStart = Len(txtVentasAnt3_2)
gnNumDec = NumDecimal(txtVentasAnt3_2)
If gbEstado And txtVentasAnt3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtVentasAnt3_2 = Format(txtVentasAnt3_2, "#,###,###,##0")
        Case 1
                txtVentasAnt3_2 = Format(txtVentasAnt3_2, "#,###,###,##0.0")
        Case 2
                txtVentasAnt3_2 = Format(txtVentasAnt3_2, "#,###,###,##0.00")
        Case 3
                txtVentasAnt3_2 = Format(txtVentasAnt3_2, "#,###,###,##0.000")
        Case Else
                txtVentasAnt3_2 = Format(txtVentasAnt3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtVentasAnt3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentasAnt3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtUtilidades3_2.SetFocus
    End If
End Sub

Private Sub txtVentasAnt3_2_GotFocus()
With txtVentasAnt3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtUtilidades3_2_Change()
txtUtilidades3_2.SelStart = Len(txtUtilidades3_2)
gnNumDec = NumDecimal(txtUtilidades3_2)
If gbEstado And txtUtilidades3_2 <> "" Then
    Select Case gnNumDec
        Case 0
                txtUtilidades3_2 = Format(txtUtilidades3_2, "#,###,###,##0")
        Case 1
                txtUtilidades3_2 = Format(txtUtilidades3_2, "#,###,###,##0.0")
        Case 2
                txtUtilidades3_2 = Format(txtUtilidades3_2, "#,###,###,##0.00")
        Case 3
                txtUtilidades3_2 = Format(txtUtilidades3_2, "#,###,###,##0.000")
        Case Else
                txtUtilidades3_2 = Format(txtUtilidades3_2, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtUtilidades3_2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtUtilidades3_2, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtVentasEEFF3.SetFocus
    End If
End Sub

Private Sub txtUtilidades3_2_GotFocus()
With txtUtilidades3_2
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVentasAnt3_2_LostFocus()
    With txtVentasAnt3_2
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Private Sub txtVentasEEFF3_Change()
txtVentasEEFF3.SelStart = Len(txtVentasEEFF3)
gnNumDec = NumDecimal(txtVentasEEFF3)
If gbEstado And txtVentasEEFF3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtVentasEEFF3 = Format(txtVentasEEFF3, "#,###,###,##0")
        Case 1
                txtVentasEEFF3 = Format(txtVentasEEFF3, "#,###,###,##0.0")
        Case 2
                txtVentasEEFF3 = Format(txtVentasEEFF3, "#,###,###,##0.00")
        Case 3
                txtVentasEEFF3 = Format(txtVentasEEFF3, "#,###,###,##0.000")
        Case Else
                txtVentasEEFF3 = Format(txtVentasEEFF3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtVentasEEFF3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentasEEFF3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtUtilidadesEEFF3.SetFocus
    End If
End Sub

Private Sub txtVentasEEFF3_GotFocus()
With txtVentasEEFF3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtUtilidadesEEFF3_Change()
txtUtilidadesEEFF3.SelStart = Len(txtUtilidadesEEFF3)
gnNumDec = NumDecimal(txtUtilidadesEEFF3)
If gbEstado And txtUtilidadesEEFF3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtUtilidadesEEFF3 = Format(txtUtilidadesEEFF3, "#,###,###,##0")
        Case 1
                txtUtilidadesEEFF3 = Format(txtUtilidadesEEFF3, "#,###,###,##0.0")
        Case 2
                txtUtilidadesEEFF3 = Format(txtUtilidadesEEFF3, "#,###,###,##0.00")
        Case 3
                txtUtilidadesEEFF3 = Format(txtUtilidadesEEFF3, "#,###,###,##0.000")
        Case Else
                txtUtilidadesEEFF3 = Format(txtUtilidadesEEFF3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtUtilidadesEEFF3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtUtilidadesEEFF3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtPorcentaje3.SetFocus
    End If
End Sub

Private Sub txtUtilidadesEEFF3_GotFocus()
With txtUtilidadesEEFF3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtPorcentaje3_Change()
txtPorcentaje3.SelStart = Len(txtPorcentaje3)
gnNumDec = NumDecimal(txtPorcentaje3)
If gbEstado And txtPorcentaje3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtPorcentaje3 = Format(txtPorcentaje3, "#,###,###,##0")
        Case 1
                txtPorcentaje3 = Format(txtPorcentaje3, "#,###,###,##0.0")
        Case 2
                txtPorcentaje3 = Format(txtPorcentaje3, "#,###,###,##0.00")
        Case 3
                txtPorcentaje3 = Format(txtPorcentaje3, "#,###,###,##0.000")
        Case Else
                txtPorcentaje3 = Format(txtPorcentaje3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtPorcentaje3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPorcentaje3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtPorcentaje3.SetFocus
    End If
End Sub

Private Sub txtPorcentaje3_GotFocus()
With txtPorcentaje3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtLineaCred3_Change()
txtLineaCred3.SelStart = Len(txtLineaCred3)
gnNumDec = NumDecimal(txtLineaCred3)
If gbEstado And txtLineaCred3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtLineaCred3 = Format(txtLineaCred3, "#,###,###,##0")
        Case 1
                txtLineaCred3 = Format(txtLineaCred3, "#,###,###,##0.0")
        Case 2
                txtLineaCred3 = Format(txtLineaCred3, "#,###,###,##0.00")
        Case 3
                txtLineaCred3 = Format(txtLineaCred3, "#,###,###,##0.000")
        Case Else
                txtLineaCred3 = Format(txtLineaCred3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtLineaCred3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtLineaCred3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        Call txtLineaCred3_ActualizaPatrimonio(KeyAscii)
        txtOtrosIngr3.SetFocus
    End If
End Sub

Private Sub txtLineaCred3_GotFocus()
With txtLineaCred3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtOtrosIngr3_Change()
txtOtrosIngr3.SelStart = Len(txtOtrosIngr3)
gnNumDec = NumDecimal(txtOtrosIngr3)
If gbEstado And txtOtrosIngr3 <> "" Then
    Select Case gnNumDec
        Case 0
                txtOtrosIngr3 = Format(txtOtrosIngr3, "#,###,###,##0")
        Case 1
                txtOtrosIngr3 = Format(txtOtrosIngr3, "#,###,###,##0.0")
        Case 2
                txtOtrosIngr3 = Format(txtOtrosIngr3, "#,###,###,##0.00")
        Case 3
                txtOtrosIngr3 = Format(txtOtrosIngr3, "#,###,###,##0.000")
        Case Else
                txtOtrosIngr3 = Format(txtOtrosIngr3, "#,###,###,##0.0000")
    End Select
End If
gbEstado = False
RaiseEvent Change
End Sub
Private Sub txtOtrosIngr3_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtOtrosIngr3, KeyAscii, , 4)
    If KeyAscii = 13 Then
        txtOtrosIngr3.SetFocus
    End If
End Sub

Private Sub txtOtrosIngr3_GotFocus()
With txtOtrosIngr3
    .SelStart = 0#
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtVentasEEFF3_LostFocus()
    With txtVentasEEFF3
        If .Text = "" Then
            .Text = "0.00"
        End If
    End With
End Sub

Public Sub GenerarExcel()
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet



    Dim oCreditos As New DCreditos

    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim n As Integer
    Dim pnLinPage As Integer
    Dim nMES As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim dFechaCP As Date
    Dim lsCelda As String
    Dim objDPersona As COMDPersona.DCOMPersona
    Dim oRS As ADODB.Recordset
    'WIOR 20141017 ******************
    Dim sCabVinculado As String
    Dim sCabVinculadoAct As String
    Dim nTipoCabVinc As Integer
    Dim nCantRegVinculados As Integer
    'WIOR FIN ***********************
    
    'Dim lnContador As Integer
    Set oRS = New ADODB.Recordset
    Set objDPersona = New COMDPersona.DCOMPersona
'On Error GoTo GeneraExcelErr
    Dim clsTC As COMDConstSistema.NCOMTipoCambio
    Dim nTC As Double
    Set clsTC = New COMDConstSistema.NCOMTipoCambio
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "InformeRiesgo"
    'Primera Hoja ******************************************************
    lsNomHoja = "OPINIÓN DE RIESGOS"
    '*******************************************************************
    lsArchivo1 = "\spooler\" & lsArchivo & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If

    nSaltoContador = 6

    'Cuadro1
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro1(lsCtaCod, gdFecSis)
    nPase = 1
    If (oRS Is Nothing) Then
        nPase = 0
    End If
    If nPase = 1 Then
    Do While Not oRS.EOF
            xlHoja1.Cells(6, 8) = oRS!cPersNombreTitular
            xlHoja1.Cells(6, 36) = oRS!cAgeDescripcion
            xlHoja1.Cells(7, 5) = oRS!cActividad
            xlHoja1.Cells(7, 36) = oRS!cPersNombreAnalista
            xlHoja1.Cells(8, 8) = oRS!Antiguedad_Neg
            xlHoja1.Cells(8, 36) = oRS!cModalidad
            xlHoja1.Cells(9, 5) = oRS!cSector
            xlHoja1.Cells(9, 37) = "'" & oRS!cCtaCod
            xlHoja1.Cells(10, 5) = oRS!cDestino
            xlHoja1.Cells(10, 37) = oRS!cTipoCredito

            xlHoja1.Cells(15, 8) = IIf(IsNull(oRS!Calificacion), "", oRS!Calificacion)
            xlHoja1.Cells(15, 23) = IIf(IsNull(oRS!Antiguedad_CMACM), "", oRS!Antiguedad_CMACM)
            xlHoja1.Cells(15, 36) = IIf(IsNull(oRS!CantidadCreditos), "", oRS!CantidadCreditos)

            xlHoja1.Cells(17, 7) = IIf(IsNull(oRS!CalificacionSBS), "", oRS!CalificacionSBS)
            xlHoja1.Cells(17, 23) = IIf(Mid(oRS!cCtaCod, 9, 1) = "1", "S/.", "$") & Format(IIf(IsNull(oRS!nSumSalRCC), 0, oRS!nSumSalRCC), "###,###,###,###0.00")
            xlHoja1.Cells(17, 34) = IIf(IsNull(oRS!Can_EntsRCC), 0, oRS!Can_EntsRCC)

            xlHoja1.Cells(19, 13) = IIf(IsNull(oRS!cHisCredCMACM), "", oRS!cHisCredCMACM)
            xlHoja1.Cells(24, 13) = IIf(IsNull(oRS!cEvolSistFina), "", oRS!cEvolSistFina)

            nSaltoContador = nSaltoContador + 1
            oRS.MoveNext
        If oRS.EOF Then
           Exit Do
        End If
    Loop
    End If

    'Vinculados
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro1Vinculados(lsCtaCod)
    lnContador = 30
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
        xlHoja1.Cells(lnContador, 2) = oRS!cPersNombre
        xlHoja1.Cells(lnContador, 13) = oRS!cConsDescripcion
        xlHoja1.Cells(lnContador, 18) = oRS!Val_Saldo
        xlHoja1.Cells(lnContador, 22) = oRS!nCantidadESF
        xlHoja1.Cells(lnContador, 25) = oRS!cCalSBS
        xlHoja1.Cells(lnContador, 30) = oRS!cEvoEndeu
        lnContador = lnContador + 1
        oRS.MoveNext
    Loop
    End If
    
    'WIOR 20141017 ******************
    nTipoCabVinc = 0
    sCabVinculado = ""
    sCabVinculadoAct = ""
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro1EmpVinculados(lsCtaCod)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
        nTipoCabVinc = IIf(Trim(oRS!cPersNombreVinculado) = "", 2, 1)
        sCabVinculadoAct = IIf(nTipoCabVinc = 1, Trim(oRS!cPersNombreVinculado), Trim(oRS!cGrupo))
        
        If Trim(sCabVinculado) <> Trim(sCabVinculadoAct) Then
            If nTipoCabVinc = 1 Then
                xlHoja1.Cells(lnContador, 2) = "EMPRESAS VINCULADAS A: " & sCabVinculadoAct
            ElseIf nTipoCabVinc = 2 Then
                xlHoja1.Cells(lnContador, 2) = "PRINCIPALES EMPRESAS VINCULADAS AL " & sCabVinculadoAct
            End If
            
            xlHoja1.Range(xlHoja1.Cells(lnContador, 2), xlHoja1.Cells(lnContador, 30)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(lnContador, 2), xlHoja1.Cells(lnContador, 30)).Merge
            lnContador = lnContador + 1
        End If
     
        xlHoja1.Cells(lnContador, 2) = oRS!cPersNombreEmp
        xlHoja1.Cells(lnContador, 13) = oRS!cRelacion
        xlHoja1.Cells(lnContador, 18) = oRS!nMontoEnde
        xlHoja1.Cells(lnContador, 22) = oRS!nNroIFIS
        xlHoja1.Cells(lnContador, 25) = IIf(oRS!nCalSistF = -1, "", oRS!nCalSistF) & " " & UCase(oRS!cCalSistF)
        xlHoja1.Cells(lnContador, 30) = oRS!cEvolucion
        lnContador = lnContador + 1
        sCabVinculado = sCabVinculadoAct
        oRS.MoveNext
    Loop
    End If
    nCantRegVinculados = lnContador
    'WIOR FIN ***********************
    
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2SaldoDeudor(lsCtaCod)
    lnContador = 118
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            xlHoja1.Cells(lnContador, 6) = oRS!nSaldo
            lnContador = lnContador + 1
            oRS.MoveNext
    Loop
    End If

    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2Creditos(lsCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
         xlHoja1.Cells(118, 15) = oRS!nMontoPropuesto
         xlHoja1.Cells(118, 24) = oRS!nNrocuotas
         xlHoja1.Cells(120, 23) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = 1, "SOLES", "DOLARES")
         xlHoja1.Cells(120, 31) = oRS!nTEA
         xlHoja1.Cells(122, 31) = IIf(IsNull(oRS!nTEMA), 0#, oRS!nTEMA)
         xlHoja1.Cells(124, 52) = IIf(IsNull(oRS!nVGET), 0, oRS!nVGET)
         xlHoja1.Cells(131, 34) = IIf(IsNull(oRS!nCuoPropuesta), 0, oRS!nCuoPropuesta)
        oRS.MoveNext
    Loop
    End If

    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2CartasFianzas(lsCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
        xlHoja1.Cells(118, 15) = oRS!nMontoPropuesto
         xlHoja1.Cells(122, 23) = oRS!nTipoCambio
         xlHoja1.Cells(118, 31) = oRS!nComisionTr
         xlHoja1.Cells(120, 23) = IIf(Mid(ActxCta.NroCuenta, 9, 1) = 1, "SOLES", "DOLARES")
         xlHoja1.Cells(124, 52) = oRS!nVGET
        oRS.MoveNext
    Loop
    End If

    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2Garantias(lsCtaCod)
'    FormateaFlex FEGarantias2
'    lnContadorGada = 0
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            xlHoja1.Cells(124, 15) = oRS!nVRM
            xlHoja1.Cells(124, 22) = oRS!nVGravamen
            xlHoja1.Cells(124, 6) = oRS!cConsDescripcion
            xlHoja1.Cells(126, 9) = oRS!CDescripcion
            oRS.MoveNext
    Loop
    End If


    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro2EscaCuotasC2(lsCtaCod)
'    FormateaFlex FECuoPro2
    lnContador = 131

    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
            xlHoja1.Cells(lnContador, 21) = oRS!nMontoPagado
            oRS.MoveNext
    Loop
    End If

    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro3(lsCtaCod, nTC)
    If Not (oRS.BOF Or oRS.EOF) Then
    Do While Not oRS.EOF
        xlHoja1.Cells(141, 5) = oRS!nVentas
        xlHoja1.Cells(142, 5) = oRS!nUtilidades 'LUCV->16
        xlHoja1.Cells(143, 5) = oRS!nUnVentas
        xlHoja1.Cells(144, 5) = oRS!nRazonCTE '
        xlHoja1.Cells(141, 16) = oRS!nSaldoDisponible
        xlHoja1.Cells(142, 16) = oRS!nCapaPago
        xlHoja1.Cells(144, 16) = oRS!nCapiTrab
        xlHoja1.Cells(141, 27) = oRS!nPatrimonio
        xlHoja1.Cells(141, 38) = oRS!nPasivo
        xlHoja1.Cells(142, 38) = oRS!nLineaCred
        xlHoja1.Cells(142, 27) = oRS!nApalancamiento
        xlHoja1.Cells(143, 16) = oRS!nSensibli
        xlHoja1.Cells(143, 27) = (oRS!nMoraDelSector / 100) 'LUCV20160920
        xlHoja1.Cells(143, 38) = oRS!nOtrosIngr
        xlHoja1.Cells(144, 27) = oRS!nCapitalSocial
        xlHoja1.Cells(148, 5) = oRS!nMesAnterior1
        xlHoja1.Cells(148, 9) = oRS!nMesAnterior2
        xlHoja1.Cells(148, 13) = oRS!nMesAnterior3
        xlHoja1.Cells(149, 22) = oRS!nVentasAnter1
        xlHoja1.Cells(149, 26) = oRS!nUtilidAnter1
        xlHoja1.Cells(149, 30) = oRS!nVentasAnter2
        xlHoja1.Cells(149, 34) = oRS!nUtilidAnter2
        xlHoja1.Cells(151, 9) = oRS!cCondicionS
        xlHoja1.Cells(153, 8) = oRS!cCalidadEvS
        oRS.MoveNext
    Loop
    End If

    'Cuadro4
    Set oRS = objDPersona.ObtenerInformeRiesgoCuadro4(lsCtaCod)
    nPase = 1
    If (oRS Is Nothing) Then
        nPase = 0
    End If
    If nPase = 1 Then
        Do While Not oRS.EOF
                xlHoja1.Cells(2, 2) = "INFORME DE OPINIÓN DE RIESGOS  Nº " & oRS!cNumeroInforme
                If oRS!nNivelRiesgo = 1 Then
                    xlHoja1.Cells(165, 6) = "SI"
                ElseIf oRS!nNivelRiesgo = 2 Then
                    xlHoja1.Cells(165, 17) = "SI"
                ElseIf oRS!nNivelRiesgo = 3 Then
                    xlHoja1.Cells(165, 28) = "SI"
                ElseIf oRS!nNivelRiesgo = 4 Then
                    xlHoja1.Cells(165, 39) = "SI"
                End If
                xlHoja1.Cells(167, 6) = oRS!cConclusione
                xlHoja1.Cells(178, 7) = oRS!cRecomendaci
                xlHoja1.Cells(193, 7) = oRS!cAnalisisRie
                xlHoja1.Cells(193, 37) = Format(oRS!dFechaNR, "YYYY/MM/DD")
                nSaltoContador = nSaltoContador + 1
            oRS.MoveNext
            If oRS.EOF Then
               Exit Do
            End If
        Loop
    End If

    xlHoja1.Range(xlHoja1.Cells(nCantRegVinculados, 2), xlHoja1.Cells(114, 45)).Delete 'WIOR 20141017
    Set objDPersona = Nothing
    Set oRS = Nothing

    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub
