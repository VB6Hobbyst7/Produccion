VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmCreOpeRFA 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos en RFA"
   ClientHeight    =   7560
   ClientLeft      =   3795
   ClientTop       =   2355
   ClientWidth     =   8400
   Icon            =   "FrmCreOpeRFA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   15
      TabIndex        =   62
      Top             =   6960
      Width           =   8370
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   6300
         TabIndex        =   67
         Top             =   150
         Width           =   1275
      End
      Begin VB.CommandButton cmdmora 
         Caption         =   "&Mora"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4920
         TabIndex        =   66
         Top             =   150
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   3510
         TabIndex        =   65
         Top             =   150
         Width           =   1275
      End
      Begin VB.CommandButton CmdPlanPagos 
         Caption         =   "&Plan Pagos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2160
         TabIndex        =   63
         Top             =   150
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   780
         TabIndex        =   64
         Top             =   150
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   30
      TabIndex        =   47
      Top             =   5460
      Width           =   8340
      Begin VB.ComboBox CmbForPag 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmCreOpeRFA.frx":030A
         Left            =   1335
         List            =   "FrmCreOpeRFA.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   1785
      End
      Begin VB.TextBox TxtMonPag 
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
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   5
         Top             =   570
         Width           =   1770
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Deuda Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   4710
         TabIndex        =   82
         Top             =   1035
         Width           =   1710
      End
      Begin VB.Label LblDeudaFecha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   6450
         TabIndex        =   81
         ToolTipText     =   "Monto a pagar por el cliente"
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblTotalPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   6450
         TabIndex        =   61
         ToolTipText     =   "Monto a pagar por el cliente"
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Deuda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   4890
         TabIndex        =   60
         Top             =   555
         Width           =   1560
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3210
         TabIndex        =   58
         Top             =   255
         Width           =   1050
      End
      Begin VB.Label LblProxfec 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4575
         TabIndex        =   57
         Top             =   975
         Width           =   45
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Prox. fecha Pag :"
         Height          =   195
         Left            =   3180
         TabIndex        =   56
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Pagar"
         Height          =   195
         Left            =   135
         TabIndex        =   55
         Top             =   615
         Width           =   1050
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Saldo de Capital"
         Height          =   195
         Left            =   90
         TabIndex        =   54
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label LblNewSalCap 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1890
         TabIndex        =   53
         Top             =   900
         Width           =   45
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Cuota Pendiente"
         Height          =   195
         Left            =   90
         TabIndex        =   52
         Top             =   1230
         Width           =   1710
      End
      Begin VB.Label LblNewCPend 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1905
         TabIndex        =   51
         Top             =   1230
         Width           =   45
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Estado Credito"
         Height          =   195
         Left            =   3180
         TabIndex        =   50
         Top             =   1215
         Width           =   1035
      End
      Begin VB.Label LblEstado 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4560
         TabIndex        =   49
         Top             =   1215
         Width           =   75
      End
      Begin VB.Label LblNumDoc 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4425
         TabIndex        =   4
         Top             =   225
         Width           =   1665
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "I.T.F. :"
         Height          =   195
         Left            =   3225
         TabIndex        =   48
         Top             =   615
         Width           =   465
      End
      Begin VB.Label LblItf 
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
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   3750
         TabIndex        =   6
         Top             =   570
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Cuota"
      Height          =   2940
      Left            =   15
      TabIndex        =   12
      Top             =   2520
      Width           =   8370
      Begin VB.TextBox txtTotaLRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6750
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   2460
         Width           =   1245
      End
      Begin VB.TextBox txtTotalRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   2475
         Width           =   1245
      End
      Begin VB.TextBox txtTotalDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   2475
         Width           =   1245
      End
      Begin VB.TextBox txtIntMorRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1485
         Width           =   1245
      End
      Begin VB.TextBox txtIntMorRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   1485
         Width           =   1245
      End
      Begin VB.TextBox txtIntMorDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1485
         Width           =   1245
      End
      Begin VB.TextBox txtComCofideDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2085
         Width           =   1245
      End
      Begin VB.TextBox txtGastosDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1785
         Width           =   1245
      End
      Begin VB.TextBox txtIntCompDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1185
         Width           =   1245
      End
      Begin VB.TextBox txtCapitalDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   885
         Width           =   1245
      End
      Begin VB.TextBox txtnCuotaDIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   585
         Width           =   1245
      End
      Begin VB.TextBox txtComCofideRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2085
         Width           =   1245
      End
      Begin VB.TextBox txtGastoRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1785
         Width           =   1245
      End
      Begin VB.TextBox txtIntCompRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1185
         Width           =   1245
      End
      Begin VB.TextBox txtCapitalRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   885
         Width           =   1245
      End
      Begin VB.TextBox txtnCuotaRFC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   1245
      End
      Begin VB.TextBox txtComCofideRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2085
         Width           =   1245
      End
      Begin VB.TextBox txtGastosRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1785
         Width           =   1245
      End
      Begin VB.TextBox txtIntCompRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1185
         Width           =   1245
      End
      Begin VB.TextBox txtCapitalRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   885
         Width           =   1245
      End
      Begin VB.TextBox txtnCuotaRFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5790
         TabIndex        =   79
         Top             =   2505
         Width           =   615
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3060
         TabIndex        =   77
         Top             =   2505
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   75
         Top             =   2505
         Width           =   615
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Int-Mor"
         Height          =   195
         Left            =   5715
         TabIndex        =   73
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Int.Mor"
         Height          =   195
         Left            =   3015
         TabIndex        =   72
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Int.Mor."
         Height          =   195
         Left            =   195
         TabIndex        =   71
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "DIF"
         Height          =   195
         Left            =   1545
         TabIndex        =   46
         Top             =   180
         Width           =   255
      End
      Begin VB.Line Line3 
         X1              =   1035
         X2              =   2505
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Com.Cofide:"
         Height          =   195
         Left            =   195
         TabIndex        =   44
         Top             =   2130
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Left            =   210
         TabIndex        =   42
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Int.Comp."
         Height          =   195
         Left            =   195
         TabIndex        =   40
         Top             =   1230
         Width           =   675
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
         Height          =   195
         Left            =   195
         TabIndex        =   38
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "NroCuota:"
         Height          =   195
         Left            =   165
         TabIndex        =   36
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "RFC"
         Height          =   195
         Left            =   4395
         TabIndex        =   35
         Top             =   210
         Width           =   315
      End
      Begin VB.Line Line2 
         X1              =   3945
         X2              =   5415
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Com.Cofide:"
         Height          =   195
         Left            =   3045
         TabIndex        =   33
         Top             =   2130
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Left            =   3045
         TabIndex        =   31
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Int.Comp."
         Height          =   195
         Left            =   3015
         TabIndex        =   29
         Top             =   1230
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
         Height          =   195
         Left            =   3015
         TabIndex        =   27
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "NroCuota:"
         Height          =   195
         Left            =   3015
         TabIndex        =   25
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "RFA"
         Height          =   195
         Left            =   7095
         TabIndex        =   24
         Top             =   210
         Width           =   315
      End
      Begin VB.Line Line1 
         X1              =   6525
         X2              =   7995
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Com.Cofide:"
         Height          =   195
         Left            =   5745
         TabIndex        =   22
         Top             =   2130
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Left            =   5745
         TabIndex        =   20
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Int.Comp:"
         Height          =   195
         Left            =   5715
         TabIndex        =   18
         Top             =   1230
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
         Height          =   195
         Left            =   5715
         TabIndex        =   16
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "NroCuota:"
         Height          =   195
         Left            =   5715
         TabIndex        =   14
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pendiente:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   225
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   1605
      Left            =   30
      TabIndex        =   11
      Top             =   930
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   2831
      _Version        =   393216
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      Height          =   915
      Left            =   0
      TabIndex        =   7
      Top             =   30
      Width           =   8355
      Begin VB.CheckBox ChkRefinanciados 
         Caption         =   "Pago Refinanciados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3630
         TabIndex        =   80
         Top             =   180
         Width           =   2535
      End
      Begin SICMACT.TxtBuscar TxtBuscar1 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   210
         Width           =   2325
         _ExtentX        =   4948
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.TextBox txtDNIPersona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5535
         TabIndex        =   2
         Top             =   525
         Width           =   1665
      End
      Begin VB.TextBox txtNombrePersona 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   525
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DNI"
         Height          =   195
         Left            =   5025
         TabIndex        =   9
         Top             =   585
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   285
         TabIndex        =   8
         Top             =   615
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmCreOpeRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nValor As Integer ' variable que determina si el cliente posee creditos en rfa
Dim bPago As Boolean ' variable que determina si el formulario esta listo para realizar un pago
Const nColorBlanco = &HFFFFFF
Const nColorPendiente = &H80C0FF
Const nColorAzul = &HFF0000
Const nColorNegro = &H0&
Dim cCtaCodRFC As String
Dim cCtaCodDIF As String
Dim cCtaCodRFA As String
Dim nMontoRFC As Double
Dim nMontoDIF As Double
Dim nMontoRFA As Double
Dim nMotoTotal As Double

Dim nMontoProDIF As Double
Dim nMontoProxRFC As Double
Dim nMontoProxRFA As Double


Dim nMontoRFCCancel As Double
Dim nMontoDIFCancel As Double
Dim nMontoRFACancel As Double
        
Dim nInteresDesaguioRFC As Double
Dim nInteresDesaguioDIF As Double
Dim nInteresDesaguioRFA As Double

' variables de pago de rfa por recepcion
Dim bRecepcionCmact As Boolean
Dim sPersCmac As String

Dim MatTempCalendRFA As Variant
Dim MatTempDistribuidoRFA As Variant
Dim MatTempCalendRFC As Variant
Dim MatTempDistribuidoRFC As Variant
Dim MatTempCalendDIF As Variant
Dim MatTempDistribuidoDIF As Variant
Dim nITFRFA As Double
Dim nITFRFC As Double
Dim nITFDIF As Double
Dim nRedondeoITF As Double 'BRGO 20110914

'Sub GeneraTempRFA(ByVal pMatCalend As Variant, ByVal pMatDistribuido As Variant, ByVal pnITF As Double)
'    MatTempCalendRFA = pMatCalend
'    MatTempDistribuidoRFA = pMatDistribuido
'    nITFRFA = pnITF
'End Sub
'
'Sub GeneraTempRFC(ByVal pMatCalend As Variant, ByVal pMatDistribuido As Variant, ByVal pnITF As Double)
'    MatTempCalendRFC = pMatCalend
'    MatTempDistribuidoRFC = pMatDistribuido
'    nITFRFC = pnITF
'End Sub
'
'Sub GenereTempDIF(ByVal pMatCalend As Variant, ByVal pMatDistribuido As Variant, ByVal pnITF As Double)
'    MatTempCalendDIF = pMatCalend
'    MatTempDistribuidoDIF = pMatDistribuido
'    nITFDIF = pnITF
'End Sub

Public Sub RecepcionCmac(ByVal psPersCodCMAC As String)
    bRecepcionCmact = True
    sPersCmac = psPersCodCMAC
    Me.Show 1
End Sub

Sub Configurar_MSH()
    With MSH
        .Clear
        .Cols = 10
        .Rows = 2
        .TextMatrix(0, 0) = "Des.Credito"
        .TextMatrix(0, 1) = "Nro.Cuota"
        .TextMatrix(0, 2) = "Fec.Ven"
        .TextMatrix(0, 3) = "Atraso"
        .TextMatrix(0, 4) = "Capital"
        .TextMatrix(0, 5) = "Int.Comp."
        .TextMatrix(0, 6) = "Int.Mor"
        .TextMatrix(0, 7) = "Com.Cofide"
        .TextMatrix(0, 8) = "Gastos"
        .TextMatrix(0, 9) = "Total Cuota"
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
    End With
    
End Sub
Private Sub CmdBuscaPersona_Click()
'    Dim oPersona As COMDPersona.DCOMPersona
'    On Error GoTo ErrHandler

'        If Len(txtNombrePersona) > 0 Then
'            frmBuscaPersona.Inicio (True)
'        Else
'            frmBuscaPersona.Inicio (False)
'        End If
    '    CargaPersona (frmBuscaPersona.Inicio)
        
'    Exit Sub
'ErrHandler:
'If Not oPersona Is Nothing Then Set oPersona = Nothing
End Sub

Private Sub cmdCancelar_Click()
 'Form_Load
 Configurar_MSH
 bPago = False
    
 cmdGrabar.Enabled = False
 CmdPlanPagos.Enabled = False
 cmdmora.Enabled = False
 TxtBuscar1.Text = ""
 txtNombrePersona.Text = ""
 txtDNIPersona.Text = ""
 LimpiarCajasTexto
 CmbForPag.Enabled = False
 TxtMonPag.Text = ""
 lblTotalPago.Caption = ""
 lblITF.Caption = "0.00"
 LblDeudaFecha.Caption = "0.00"
 nMontoRFCCancel = 0
 nMontoDIFCancel = 0
 nMontoRFACancel = 0
 nRedondeoITF = 0
End Sub

Sub LimpiarCajasTexto()
    txtnCuotaDIF.Text = ""
    txtCapitalDIF.Text = ""
    txtIntCompDIF.Text = ""
    txtIntMorDIF.Text = ""
    txtGastosDIF.Text = ""
    txtComCofideDIF.Text = ""
    txtTotalDIF.Text = ""
    txtnCuotaRFC.Text = ""
    txtCapitalRFC.Text = ""
    txtIntCompRFC.Text = ""
    txtIntMorRFC.Text = ""
    txtGastoRFC.Text = ""
    txtComCofideRFC.Text = ""
    txtTotalRFC.Text = ""
    txtnCuotaRFA.Text = ""
    txtCapitalRFA.Text = ""
    txtIntCompRFA.Text = ""
    txtIntMorRFA.Text = ""
    txtGastosRFA.Text = ""
    txtComCofideRFA.Text = ""
    txtTotaLRFA.Text = ""
End Sub

Private Sub CmdGrabar_Click()

Dim oRFA As COMNCredito.NCOMRFA
Dim sImpresion() As String
Dim bProceso As Boolean
Dim sError As String
Dim nMonPag As Double
Dim nCuotaDIF As Integer

Set oRFA = New COMNCredito.NCOMRFA

nMonPag = CDbl(IIf(TxtMonPag.Text = "", 0, TxtMonPag.Text))
nCuotaDIF = CInt(IIf(txtnCuotaDIF.Text = "", 0, txtnCuotaDIF.Text))

bProceso = oRFA.GrabarPagoRFA(CDbl(LblDeudaFecha.Caption), nMonPag, nCuotaDIF, TxtBuscar1.Text, gdFecSis, IIf(CHKRefinanciados.value = 1, True, False), _
                    txtNombrePersona.Text, gsNomAge, gsCodUser, sLpt, gsCodCMAC, bRecepcionCmact, sPersCmac, gsCodAge, cCtaCodDIF, cCtaCodRFC, cCtaCodRFA, _
                    nMontoRFCCancel, nMontoDIFCancel, nMontoRFACancel, nInteresDesaguioRFC, nInteresDesaguioDIF, nInteresDesaguioRFA, nMontoProDIF, nMontoProxRFC, _
                    nMontoProxRFA, nMotoTotal, nMontoRFC, nMontoDIF, nMontoRFA, sError, sImpresion)

Set oRFA = Nothing

If bProceso = False Then
    MsgBox sError, vbInformation, "Mensaje"
Else
    Call ImprimeBoletas(sImpresion)
End If

Call cmdCancelar_Click

'Dim I As Integer
'Dim MatCalend As Variant
'Dim MatCalendDistribuido As Variant
'Dim nMontoPagoCred As Double
'Dim sCtaCodPago As String
'Dim oNCred As COMNCredito.NCOMCredito
'Dim oNRFA As COMNCredito.NCOMRFA
'Dim sError As String
'Dim oDoc As COMNCredito.NCOMCredDoc
'Dim nitf As Double
'Dim oBase As COMDCredito.DCOMCredActBD
'Dim sMovNro As String
'Dim nmovnro As Long
'Dim nCuota As Integer
'Dim sTitulo As String
'Dim nMontoPago As Double
'Dim odRFa As COMDCredito.DCOMRFA
'Dim sRFA As String
'Dim oRFA As COMDCredito.DCOMRFA
'Dim bProcesoRFA As Boolean
'Dim bProcesoRFC As Boolean
'Dim bProcesoDIF As Boolean
'
'    If MsgBox("Se va a Efectuar la Operacion, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'    bProcesoRFA = False
'    bProcesoRFC = False
'    bProcesoDIF = False
'    'Valida el monto de pago con la cancelacion adelantada...
'    If Val(LblDeudaFecha) <> Val(TxtMonPag) Then
'            'Es una amortizacion de creditos...
'            'Distribuyo Montos
'            nMotoTotal = CDbl(IIf(TxtMonPag.Text = "", 0, TxtMonPag.Text))
'            nMontoRFC = CDbl(IIf(txtTotalRFC.Text = "", 0, txtTotalRFC.Text))
'            nMontoDIF = CDbl(IIf(txtTotalDIF.Text = "", 0, txtTotalDIF.Text))
'            nMontoRFA = CDbl(IIf(txtTotaLRFA.Text = "", 0, txtTotaLRFA.Text))
'
'            nMontoDIF = 0
'            nMontoRFA = 0
'            nMontoRFC = 0
'            nMotoTotal = CDbl(IIf(TxtMonPag.Text = "", 0, TxtMonPag.Text))
'            nCuota = CInt(IIf(txtnCuotaDIF = "", 0, txtnCuotaDIF))
'            Do While nMotoTotal > 0
'               Call MontoProximos(nCuota)
'                If nMotoTotal >= nMontoProDIF Then
'                    nMotoTotal = nMotoTotal - nMontoProDIF
'                    nMontoDIF = nMontoDIF + nMontoProDIF
'                Else
'                    nMontoDIF = nMontoDIF + nMotoTotal
'                    nMotoTotal = 0
'                    Exit Do
'                End If
'                   nMotoTotal = Format(nMotoTotal, "#0.00")
'
'                    If nMotoTotal = 0 Then
'                        Exit Do
'                    End If
'
'                If nMotoTotal >= nMontoProxRFC Then
'                    nMotoTotal = nMotoTotal - nMontoProxRFC
'                    nMontoRFC = nMontoRFC + nMontoProxRFC
'
'                Else
'                    nMontoRFC = nMontoRFC + nMotoTotal
'                    nMotoTotal = 0
'                    Exit Do
'                End If
'
'                    nMotoTotal = Format(nMotoTotal, "#0.00")
'                    If nMotoTotal = 0 Then
'                        Exit Do
'                    End If
'
'                If nMotoTotal >= nMontoProxRFA Then
'                    nMotoTotal = nMotoTotal - nMontoProxRFA
'                    nMontoRFA = nMontoRFA + nMontoProxRFA
'                Else
'                    nMontoRFA = nMontoRFA + nMotoTotal
'                    nMotoTotal = 0
'
'                End If
'                nMotoTotal = Format(nMotoTotal, "#0.00")
'                If nMotoTotal = 0 Then
'                    Exit Do
'                End If
'                nCuota = nCuota + 1
'            Loop
'
'
'            'Cargamos los calendarios
'            Set oBase = New COMDCredito.DCOMCredActBD
'            oBase.coConex.AbreConexion
'            oBase.dBeginTrans
'
'            sMovNro = oBase.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'            Call oBase.dInsertMov(sMovNro, "100201", "", gMovEstContabMovContable, gMovFlagVigente, False)
'            nmovnro = oBase.dGetnMovNro(sMovNro)
'
'            For I = 1 To 3
'                If I = 1 Then
'                    sCtaCodPago = cCtaCodDIF
'                    nMontoPagoCred = nMontoDIF
'                    sTitulo = "DIF"
'                End If
'                If I = 2 Then
'                    sCtaCodPago = cCtaCodRFC
'                    nMontoPagoCred = nMontoRFC
'                    sTitulo = "RFC"
'                End If
'                If I = 3 Then
'                    sCtaCodPago = cCtaCodRFA
'                    nMontoPagoCred = nMontoRFA
'                    sTitulo = "RFA"
'                End If
'
'
'                If nMontoPagoCred > 0 And sCtaCodPago <> "" Then
'
'                    Set oNCred = New COMNCredito.NCOMCredito
'                    MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(sCtaCodPago, , True)
'                    MatCalendDistribuido = oNCred.CrearMatrizparaAmortizacion(MatCalend)
'
'                    nMontoPago = fgITFCalculaImpuestoNOIncluido(nMontoPagoCred)
'                    nitf = Format(nMontoPago - nMontoPagoCred, "0.00")
'                    nMontoPago = nMontoPagoCred
'
'                    'Distribuye el Pago
'                    MatCalendDistribuido = oNCred.MatrizDistribuirMonto(MatCalend, nMontoPago, "GMIC")
'
'                     Set odRFa = New COMDCredito.DCOMRFA
'                     sRFA = odRFa.GetCRFA(sCtaCodPago)
'                     Set odRFa = Nothing
'
'                   'sError = oNCred.AmortizarCredito(sCtaCodPago, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, "GMIC", gColocTipoPagoEfectivo, gsCodAge, gsCodUser, , oBase, nMovNro, , , , , , sMovNro, , , , , , , , nitf)
'                    Set oNRFA = New COMNCredito.NCOMRFA
'                    sError = oNRFA.AmortizarCredito(sCtaCodPago, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, "GMIC", gColocTipoPagoEfectivo, gsCodAge, gsCodUser, , oBase, nmovnro, , bRecepcionCmact, sPersCmac, , , sMovNro, , , , , , , , nitf, True, , sRFA)
'                    If sError <> "" Then
'                        MsgBox sError, vbInformation, "Aviso"
'                    Else
'                        If sTitulo = "DIF" Then
'                           If ChkRefinanciados.value = 0 Then
'                                bProcesoDIF = True
'                                Call GenereTempDIF(MatCalend, MatCalendDistribuido, nitf)
'                           End If
'                        ElseIf sTitulo = "RFC" Then
'                           Call GeneraTempRFC(MatCalend, MatCalendDistribuido, nitf)
'                           bProcesoRFC = True
'                        ElseIf sTitulo = "RFA" Then
'                           bProcesoRFA = True
'                           Call GeneraTempRFA(MatCalend, MatCalendDistribuido, nitf)
'                        End If
'
''                        Set oDoc = New COMNCredito.NCOMCredDoc
''                        Call oDoc.ImprimeBoleta(sCtaCodPago, txtNombrePersona, gsNomAge & " -" & sTitulo, "DOLARES", _
''                            oNCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(Date, "hh:mm:ss"), 0, "", _
''                            oNCred.MatrizCapitalPagado(MatCalendDistribuido), oNCred.MatrizIntCompPagado(MatCalendDistribuido), _
''                            oNCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
''                            oNCred.MatrizIntMorPagado(MatCalendDistribuido), oNCred.MatrizGastoPag(MatCalendDistribuido), _
''                            oNCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
''                            oNCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNCred.MatrizIntReprogPag(MatCalendDistribuido), _
''                            oNCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), oNCred.MatrizFechaCuotaPendiente(MatCalend, MatCalendDistribuido), _
''                            gsCodUser, sLpt, gsNomAge, , , gsCodCMAC, nITF)
''
''                         Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
''                            Call oDoc.ImprimeBoleta(sCtaCodPago, txtNombrePersona, gsNomAge & " -" & sTitulo, "DOLARES", _
''                            oNCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(Date, "hh:mm:ss"), 0, "", _
''                            oNCred.MatrizCapitalPagado(MatCalendDistribuido), oNCred.MatrizIntCompPagado(MatCalendDistribuido), _
''                            oNCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
''                            oNCred.MatrizIntMorPagado(MatCalendDistribuido), oNCred.MatrizGastoPag(MatCalendDistribuido), _
''                            oNCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
''                            oNCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNCred.MatrizIntReprogPag(MatCalendDistribuido), _
''                            oNCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), oNCred.MatrizFechaCuotaPendiente(MatCalend, MatCalendDistribuido), _
''                            gsCodUser, sLpt, gsNomAge, , , gsCodCMAC, nITF)
''                         Loop
'                    End If
'                End If
'            Next I
'            oBase.dCommitTrans
'            Set oNCred = Nothing
'            Set oBase = Nothing
'
'            'Imprimiendo rfc
'           If bProcesoRFC = True Then
'              Call ImprimeBoletas(MatTempCalendRFC, MatTempDistribuidoRFC, nITFRFC, "RFC", cCtaCodRFC, "RFC")
'           End If
'
'           If bProcesoDIF = True Then
'             Call ImprimeBoletas(MatTempCalendDIF, MatTempDistribuidoDIF, nITFDIF, "DIF", cCtaCodDIF, "DIF")
'           End If
'
'           If bProcesoRFA = True Then
'             Call ImprimeBoletas(MatTempCalendRFA, MatTempDistribuidoRFA, nITFRFA, "RFA", cCtaCodRFA, "RFA")
'           End If
'
'    Else
'            ' es un cancelacion adelantada....
'            'Modificado por LMMD
'            'Cancelando cada uno de los creditos...
'             Set oBase = New COMDCredito.DCOMCredActBD
'             oBase.coConex.AbreConexion
'             oBase.dBeginTrans
'
'            sMovNro = oBase.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'            Call oBase.dInsertMov(sMovNro, "100201", "", gMovEstContabMovContable, gMovFlagVigente, False)
'            nmovnro = oBase.dGetnMovNro(sMovNro)
'
'            'RFC
'              If ValidaCredito(cCtaCodRFC) = True Then
'                   Set oNCred = New COMNCredito.NCOMCredito
'
'                   MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(cCtaCodRFC, , True)
'                   MatCalendDistribuido = oNCred.MatrizDistribuirCancelacion(cCtaCodRFC, MatCalend, nMontoRFCCancel, "GMIC", gdFecSis)
'                   nMontoPago = fgITFCalculaImpuestoNOIncluido(nMontoRFCCancel, True)
'                   nitf = Format(nMontoPago - nMontoRFCCancel, "0.00")
'                   nMontoPago = nMontoRFCCancel
'
'                   'Obteniendo el concepto del RFA
'                   Set odRFa = New COMDCredito.DCOMRFA
'                   sRFA = odRFa.GetCRFA(cCtaCodRFC)
'                   Set odRFa = Nothing
'
'                   Set oNRFA = New COMNCredito.NCOMRFA
'                   sError = oNRFA.AmortizarCredito(cCtaCodRFC, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, "GMIC", gColocTipoPagoEfectivo, gsCodAge, gsCodUser, , oBase, nmovnro, , bRecepcionCmact, sPersCmac, , , sMovNro, , , , , , , , nitf, nInteresDesaguioRFC, , sRFA)
'                   If sError <> "" Then
'                      MsgBox sError, vbInformation, "Aviso"
'                   Else
'                      bProcesoRFC = True
'                      Call GeneraTempRFC(MatCalend, MatCalendDistribuido, nitf)
'                   End If
'            End If
'            'Se verifica que esto solo se realize para pagos de RFA Normales...
'            'DIF
'            If ChkRefinanciados.value = 0 Then
'                If ValidaCredito(cCtaCodDIF) = True Then
'                   Set oNCred = New COMNCredito.NCOMCredito
'                   MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(cCtaCodDIF, , True)
'                   MatCalendDistribuido = oNCred.MatrizDistribuirCancelacion(cCtaCodDIF, MatCalend, nMontoDIFCancel, "GMIC", gdFecSis)
'                   nMontoPago = fgITFCalculaImpuestoNOIncluido(nMontoDIFCancel, True)
'                   nitf = Format(nMontoPago - nMontoDIFCancel, "0.00")
'                   nMontoPago = nMontoDIFCancel
'
'                   'Obteniendo el concepto del RFA
'                   Set odRFa = New COMDCredito.DCOMRFA
'                   sRFA = odRFa.GetCRFA(cCtaCodDIF)
'                   Set odRFa = Nothing
'
'                   Set oNRFA = New COMNCredito.NCOMRFA
'                   sError = oNRFA.AmortizarCredito(cCtaCodDIF, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, "GMIC", gColocTipoPagoEfectivo, gsCodAge, gsCodUser, , oBase, nmovnro, , bRecepcionCmact, sPersCmac, , , sMovNro, , , , , , , , nitf, nInteresDesaguioDIF, , sRFA)
'
'                   If sError <> "" Then
'                      MsgBox sError, vbInformation, "Aviso"
'                   Else
'                    '  bProceso = True
'                       bProcesoDIF = True
'                      Call GenereTempDIF(MatCalend, MatCalendDistribuido, nitf)
'                   End If
'
'                End If
'             End If
'             'RFA
'
'            If ValidaCredito(cCtaCodRFA) = True Then
'               Set oNCred = New COMNCredito.NCOMCredito
'               MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(cCtaCodRFA, , True)
'               Set odRFa = New COMDCredito.DCOMRFA
'               MatCalendDistribuido = odRFa.MatrizDistribuirCancelacion(cCtaCodRFA, MatCalend, nMontoRFACancel, "GMIC", gdFecSis)
'               Set odRFa = Nothing
'               nMontoPago = fgITFCalculaImpuestoNOIncluido(nMontoRFACancel, True)
'               nitf = Format(nMontoPago - nMontoRFACancel, "0.00")
'               nMontoPago = nMontoRFACancel
'
'               'Obteniendo el concepto del RFA
'               Set odRFa = New COMDCredito.DCOMRFA
'               sRFA = odRFa.GetCRFA(cCtaCodRFA)
'               Set odRFa = Nothing
'
'               Set oNRFA = New COMNCredito.NCOMRFA
'               sError = oNRFA.AmortizarCredito(cCtaCodRFA, MatCalend, MatCalendDistribuido, nMontoPago, gdFecSis, "GMIC", gColocTipoPagoEfectivo, gsCodAge, gsCodUser, , oBase, nmovnro, , bRecepcionCmact, sPersCmac, , , sMovNro, , , , , , , , nitf, nInteresDesaguioRFA, , sRFA)
'               If sError <> "" Then
'                  MsgBox sError, vbInformation, "Aviso"
'               Else
'                  bProcesoRFA = True
'                  Call GeneraTempRFA(MatCalend, MatCalendDistribuido, nitf)
'               End If
'            End If
'
'            oBase.dCommitTrans
'            Set oNCred = Nothing
'            Set oBase = Nothing
'
'        'Imprimiendo rfc
'         If bProcesoRFC = True Then
'             Call ImprimeBoletas(MatTempCalendRFC, MatTempDistribuidoRFC, nITFRFC, "RFC", cCtaCodRFC, "RFC")
'         End If
'
'         If bProcesoDIF = True Then
'            Call ImprimeBoletas(MatTempCalendDIF, MatTempDistribuidoDIF, nITFDIF, "DIF", cCtaCodDIF, "DIF")
'         End If
'
'         If bProcesoRFA = True Then
'            Call ImprimeBoletas(MatTempCalendRFA, MatTempDistribuidoRFA, nITFRFA, "RFA", cCtaCodRFA, "RFA")
'         End If
'    End If
'
'    Call cmdCancelar_Click
    
End Sub

Sub ImprimeBoletas(ByVal psImpresion As Variant)
Dim oPrevio As previo.clsprevio
Dim sDescripcion As String
Dim i As Integer
    
For i = 0 To UBound(psImpresion) - 1
    Select Case i
        Case 0
            sDescripcion = "RFC"
        Case 1
            sDescripcion = "DIF"
        Case 2
            sDescripcion = "RFA"
    End Select
    
    MsgBox "Prepare el Papel para el Credito " & sDescripcion, vbInformation, "Mensaje"
    
    Set oPrevio = New previo.clsprevio
    'oPrevio.PrintSpool sLpt, CStr(psImpresion(I))
    oPrevio.Show CStr(psImpresion(i)), sDescripcion
    
    Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
        'oPrevio.PrintSpool sLpt, CStr(psImpresion(I))
        oPrevio.Show CStr(psImpresion(i)), sDescripcion
    Loop
Next i

Set oPrevio = Nothing

End Sub


'Sub ImprimeBoletas(ByVal pMatCalend As Variant, ByVal pMatCalendDistribuido As Variant, ByVal pnITF As Double, _
'                   ByVal psTitulo As String, ByVal psCtaCod As String, ByVal psDescripcion As String)
'
'    Dim oDoc As COMNCredito.NCOMCredDoc
'    Dim oNCred As COMNCredito.NCOMCredito
'    Set oNCred = New COMNCredito.NCOMCredito
'    Set oDoc = New COMNCredito.NCOMCredDoc
'                        MsgBox "Prepare el Papel para el Credito " & psDescripcion
'
'    Call oDoc.ImprimeBoleta(psCtaCod, txtNombrePersona, gsNomAge & " -" & psTitulo, "DOLARES", _
'                            oNCred.MatrizCuotasPagadas(pMatCalendDistribuido), gdFecSis, Format(Date, "hh:mm:ss"), 0, "", _
'                            oNCred.MatrizCapitalPagado(pMatCalendDistribuido), oNCred.MatrizIntCompPagado(pMatCalendDistribuido), _
'                            oNCred.MatrizIntCompVencPagado(pMatCalendDistribuido), _
'                            oNCred.MatrizIntMorPagado(pMatCalendDistribuido), oNCred.MatrizGastoPag(pMatCalendDistribuido), _
'                            oNCred.MatrizIntGraciaPagado(pMatCalendDistribuido), _
'                            oNCred.MatrizIntSuspensoPag(pMatCalendDistribuido) + oNCred.MatrizIntReprogPag(pMatCalendDistribuido), _
'                            oNCred.MatrizSaldoCapital(pMatCalend, pMatCalendDistribuido), oNCred.MatrizFechaCuotaPendiente(pMatCalend, pMatCalendDistribuido), _
'                            gsCodUser, sLpt, gsNomAge, , , gsCodCMAC, pnITF)
'
'   Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
'                            Call oDoc.ImprimeBoleta(psCtaCod, txtNombrePersona, gsNomAge & " -" & psTitulo, "DOLARES", _
'                            oNCred.MatrizCuotasPagadas(pMatCalendDistribuido), gdFecSis, Format(Date, "hh:mm:ss"), 0, "", _
'                            oNCred.MatrizCapitalPagado(pMatCalendDistribuido), oNCred.MatrizIntCompPagado(pMatCalendDistribuido), _
'                            oNCred.MatrizIntCompVencPagado(pMatCalendDistribuido), _
'                            oNCred.MatrizIntMorPagado(pMatCalendDistribuido), oNCred.MatrizGastoPag(pMatCalendDistribuido), _
'                            oNCred.MatrizIntGraciaPagado(pMatCalendDistribuido), _
'                            oNCred.MatrizIntSuspensoPag(pMatCalendDistribuido) + oNCred.MatrizIntReprogPag(pMatCalendDistribuido), _
'                            oNCred.MatrizSaldoCapital(pMatCalend, pMatCalendDistribuido), oNCred.MatrizFechaCuotaPendiente(pMatCalend, pMatCalendDistribuido), _
'                            gsCodUser, sLpt, gsNomAge, , , gsCodCMAC, pnITF)
' Loop
' Set oDoc = Nothing
'End Sub

'Function ValidaCredito(ByVal psCtaCod As String) As Boolean
'
'    Dim odRFa As COMDCredito.DCOMRFA
'
'    Set odRFa = New COMDCredito.DCOMRFA
'    ValidaCredito = odRFa.ValidaCredito(psCtaCod)
'    Set odRFa = Nothing
'End Function

Private Sub CmdPlanPagos_Click()
    FrmPlanPagosRFA.nCodCli = TxtBuscar1.Text
    FrmPlanPagosRFA.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Configurar_MSH
    CargaControles
    bPago = False
End Sub


Private Sub TxtBuscar1_EmiteDatos()
    Dim objRFA As COMNCredito.NCOMRFA
    Dim sDNI As String
    Dim rs As ADODB.Recordset
    Dim nDeuda As Double
    'txtDNIPersona.Text = objRFA.BuscaDNICliente(TxtBuscar1.Text)
    
    If TxtBuscar1.Text = "" Then Exit Sub
    
    Set objRFA = New COMNCredito.NCOMRFA
    Call objRFA.CargarDatosClienteRFA(TxtBuscar1.Text, IIf(CHKRefinanciados.value = 1, True, False), gdFecSis, sDNI, nValor, rs, nDeuda, nMontoRFCCancel, nMontoDIFCancel, nMontoRFACancel, _
                                      nInteresDesaguioRFC, nInteresDesaguioDIF, nInteresDesaguioRFA, cCtaCodDIF, cCtaCodRFC, cCtaCodRFA)
    Set objRFA = Nothing
    
    txtDNIPersona.Text = sDNI
    
    txtNombrePersona.Text = TxtBuscar1.psDescripcion
    txtNombrePersona.Locked = True
    txtDNIPersona.Locked = True
    VerificarClienteRFA
    
    If nValor = 1 Then
        CargarCuotasPendientes rs
        CargarCajasTexto
        CalculoMontoPagar
        PintarCuotasPendientes
        'Calcular_ITF
        ControlesTransaccion
        PonerColorTotales
        LblDeudaFecha.Caption = Format(nDeuda, "#0.00") 'Format(ObtenerDeudaFecha(cCtaCodDIF, cCtaCodRFC, cCtaCodRFA, IIf(ChkRefinanciados.value = 1, True, False)), "#0.00")
        If bPago = True Then
           CmbForPag.Enabled = True
           TxtMonPag.Enabled = True
           If CmbForPag.ListCount > 0 Then
             CmbForPag.ListIndex = 0
           End If
           lblNumdoc.Enabled = True
        End If
'    Else
'        MsgBox "El cliente no posee creditos en rfa", vbInformation, "AVISO"
    End If
End Sub

Sub Calcular_ITF()
Dim nMontoPago As Currency
Dim nITF As Currency
  
    nMontoPago = fgITFCalculaImpuestoIncluido(CDbl(TxtMonPag.Text))
    nITF = Format(CDbl(TxtMonPag.Text) - nMontoPago, "0.00")
    lblITF.Caption = Format(nITF, "0.00")
End Sub

Sub VerificarClienteRFA()
    'Dim cCtaCod As String
    'Dim objRFA As COMDCredito.DCOMRFA
    
'    On Error GoTo ErrHandler
    '    cCtaCod = TxtBuscar1.Text
    '    Set objRFA = New COMDCredito.DCOMRFA
    '    nValor = objRFA.VerificaCred(cCtaCod)
        If nValor = 0 Then
            MsgBox "El cliente no posee creditos de RFA...", vbInformation, "AVISO"
            Call Configurar_MSH
        ElseIf nValor = -1 Then
            MsgBox "Error al verificar al cliente" & vbCrLf & _
                   "Porfavor informar a la oficina de TI", vbInformation, "AVISO"
            Call Configurar_MSH
        End If
    Exit Sub
'ErrHandler:
'   If Not objRFA Is Nothing Then Set objRFA = Nothing
End Sub

Sub CargarCuotasPendientes(ByVal prs As ADODB.Recordset)
    'Dim objRFA As COMDCredito.DCOMRFA
    'Dim rs As ADODB.Recordset
    Dim strCredito As String
    Dim nCuota As Integer
    Dim nCapital As Currency
    Dim nIntComp As Currency
    Dim nIntMor As Currency
    Dim nGastos As Currency
    Dim nCofide As Currency
    Dim dDias As Date
    Dim nAtraso As Integer
    Dim bAdd As Boolean 'variable que determina si ingreso al bucle
    On Error GoTo ErrHandler
    Configurar_MSH
    'Set objRFA = New COMDCredito.DCOMRFA
    'If ChkRefinanciados.value = 1 Then
    '    Set rs = objRFA.ListaCreditosPendientes(TxtBuscar1.Text, gdFecSis, True)
    'Else
    '    Set rs = objRFA.ListaCreditosPendientes(TxtBuscar1.Text, gdFecSis, False)
    'End If
    'Set objRFA = Nothing
    If Not prs.EOF And Not prs.BOF Then
            bAdd = False
            Do Until prs.EOF
                'If prs!indice = "1" Then
                '    cCtaCodDIF = prs!cCtaCod
                'End If
                'If prs!indice = "2" Then
                '    cCtaCodRFC = prs!cCtaCod
                'End If
                'If prs!indice = "3" Then
                '    cCtaCodRFA = prs!cCtaCod
                'End If
                bAdd = True
                strCredito = prs!cCredito
                nCuota = prs!nCuota
                nCapital = prs!Capital
                nIntComp = prs!IntComp
                nIntMor = prs!IntMor
                nGastos = prs!Gastos
                nCofide = prs!ComCofide
                dDias = Format(prs!dVenc, "dd/MM/yyyy")
                nAtraso = prs!DiasAtraso
                MSH.Rows = MSH.Rows + 1
                
                With MSH
                    .TextMatrix(.Rows - 2, 0) = strCredito
                    .TextMatrix(.Rows - 2, 1) = nCuota
                    .TextMatrix(.Rows - 2, 2) = dDias
                    .TextMatrix(.Rows - 2, 3) = nAtraso
                    .TextMatrix(.Rows - 2, 4) = Format(nCapital, "#.00")
                    .TextMatrix(.Rows - 2, 5) = Format(nIntComp, "#.00")
                    .TextMatrix(.Rows - 2, 6) = Format(nIntMor, "#.00")
                    .TextMatrix(.Rows - 2, 7) = Format(nCofide, "#.00")
                    .TextMatrix(.Rows - 2, 8) = Format(nGastos, "#.00")
                    .TextMatrix(.Rows - 2, 9) = Format(nCapital + nIntComp + nIntMor + nGastos + nCofide, "#.00")
                    
                 End With
              prs.MoveNext
            Loop
            If bAdd = True Then
                MSH.Rows = MSH.Rows - 1
            End If
'            TxtMonPag_KeyPress 13
    Else
        MsgBox "El cliente no posee cuotas pendientes...", vbInformation, "Aviso"
    End If
    'Set rs = Nothing
    Exit Sub
ErrHandler:
    'If Not objRFA Is Nothing Then Set objRFA = Nothing
    'If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error en el momento de Cargar las cuotas pendientes" & vbCrLf & _
           "Por favor llame a la oficina de TI", vbInformation, "AVISO"
End Sub

Sub CargarCajasTexto()
    Dim i As Integer
    Dim cCredito As String
    If MSH.Rows >= 2 Then
        If MSH.TextMatrix(1, 0) <> "" Then
            For i = 0 To MSH.Rows - 1
                cCredito = MSH.TextMatrix(i, 0)
                Select Case cCredito
                    Case "DIF"
                        txtnCuotaDIF.Text = MSH.TextMatrix(i, 1)
                        txtCapitalDIF.Text = Format(MSH.TextMatrix(i, 4), "#0.00")
                        txtIntCompDIF.Text = Format(MSH.TextMatrix(i, 5), "#0.00")
                        txtIntMorDIF.Text = Format(MSH.TextMatrix(i, 6), "#0.00")
                        txtGastosDIF.Text = Format(MSH.TextMatrix(i, 8), "#0.00")
                        txtComCofideDIF.Text = Format(MSH.TextMatrix(i, 7), "#0.00")
                        txtTotalDIF.Text = Format(MSH.TextMatrix(i, 9), "#0.00")
                    Case "RFC"
                        txtnCuotaRFC.Text = MSH.TextMatrix(i, 1)
                        txtCapitalRFC.Text = Format(MSH.TextMatrix(i, 4), "#0.00")
                        txtIntCompRFC.Text = Format(MSH.TextMatrix(i, 5), "#0.00")
                        txtIntMorRFC.Text = Format(MSH.TextMatrix(i, 6), "#0.00")
                        txtGastoRFC.Text = Format(MSH.TextMatrix(i, 8), "#0.00")
                        txtComCofideRFC.Text = Format(MSH.TextMatrix(i, 7), "#0.00")
                        txtTotalRFC.Text = MSH.TextMatrix(i, 9)
                    Case "RFA"
                        txtnCuotaRFA.Text = MSH.TextMatrix(i, 1)
                        txtCapitalRFA.Text = Format(MSH.TextMatrix(i, 4), "#0.00")
                        txtIntCompRFA.Text = Format(MSH.TextMatrix(i, 5), "#0.00")
                        txtIntMorRFA.Text = Format(MSH.TextMatrix(i, 6), "#0.00")
                        txtGastosRFA.Text = Format(MSH.TextMatrix(i, 8), "#0.00")
                        txtComCofideRFA.Text = Format(MSH.TextMatrix(i, 7), "#0.00")
                        txtTotaLRFA.Text = Format(MSH.TextMatrix(i, 9), "#0.00")
                End Select
            Next i
        End If
    End If
End Sub

Private Sub CargaControles()

Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    'Call CargaComboConstante(gColocTipoPago, CmbForPag)
    Set R = oCons.RecuperaConstantes(gColocTipoPago)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Exit Sub

ERRORCargaControles:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Sub CalculoMontoPagar()
    Dim i As Integer
    Dim nSumaPendientes As Currency
    Dim nSumaNoPendientes As Currency
    If MSH.Rows > 1 Then
        bPago = True
        nSumaPendientes = 0
        nSumaNoPendientes = 0
        For i = 1 To MSH.Rows - 1
            If Val(MSH.TextMatrix(i, 3)) > 0 Then
                nSumaPendientes = nSumaPendientes + Val(MSH.TextMatrix(i, 9))
            Else
                nSumaNoPendientes = nSumaNoPendientes + Val(MSH.TextMatrix(i, 9))
            End If
            
          '  nSuma = nSuma + MSH.TextMatrix(i, 9)
        Next i
    End If
    If nSumaPendientes > 0 Then
        lblTotalPago.Caption = nSumaPendientes
    Else
        lblTotalPago.Caption = nSumaNoPendientes
    End If
    
    TxtMonPag.Text = lblTotalPago.Caption
    TxtMonPag_KeyPress 13
    
        'nMontoPago = fgITFCalculaImpuestoIncluido(CDbl(TxtMonPag.Text))
        'nITF = Format(CDbl(TxtMonPag.Text) - nMontoPago, "0.00")
        'LblItf.Caption = Format(nITF, "0.00")
        
End Sub

Sub PintarCuotasPendientes()
    Dim i As Integer
            With MSH
                For i = 1 To .Rows - 1 ' recorre cada fila
                  If Val(.TextMatrix(i, 3)) > 0 Then
                      PintarFilaCuotaPendiente i
                   End If
                Next i
        End With
End Sub

Sub PintarFilaCuotaPendiente(ByVal pnIndice As Integer)
    Dim i As Integer
    For i = 0 To MSH.Cols - 1
        MSH.FillStyle = flexFillSingle
        MSH.Col = i
        MSH.Row = pnIndice
        MSH.CellBackColor = nColorPendiente
        MSH.CellForeColor = nColorAzul
    Next i
End Sub

Private Sub TxtMonPag_KeyPress(KeyAscii As Integer)

Dim nMontoPago As Double
Dim nITF As Double
KeyAscii = NumerosDecimales(TxtMonPag, KeyAscii)
    If KeyAscii = 13 Then
        If CDbl(TxtMonPag.Text) = 0 Then
            MsgBox "Monto de Pago Debe ser mayor que Cero", vbQuestion, "Aviso"
            Exit Sub
        End If
'        If CDbl(TxtMonPag.Text) > CDbl(LblTotDeuda.Caption) Then
'            MsgBox "Monto de Pago No Puede ser Mayor que la Deuda", vbQuestion, "Aviso"
'            Exit Sub
'        End If
        
        TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
        
 'NO ES NECESARIO
 '       LblDeudaFecha = Format(ObtenerDeudaFecha(cCtaCodDIF, cCtaCodRFC, cCtaCodRFA, IIf(ChkRefinanciados.value = 1, True, False)), "#0.00")
        
        If Val(LblDeudaFecha) <> Val(TxtMonPag) Then
              nMontoPago = fgITFCalculaImpuestoNOIncluido(CDbl(TxtMonPag.Text), False)
         Else
              nMontoPago = fgITFCalculaImpuestoNOIncluido(CDbl(TxtMonPag.Text), True)
         End If
        
        nITF = Format(nMontoPago - CDbl(TxtMonPag.Text), "0.00")
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(nITF)
        If nRedondeoITF > 0 Then
            nITF = nITF - nRedondeoITF
        End If
        '*** END BRGO
        lblITF.Caption = Format(nITF, "0.00")
        lblTotalPago.Caption = Format(Val(TxtMonPag.Text) + nITF, "#0.00")
        'NO ES NECESARIO
        'LblDeudaFecha = Format(ObtenerDeudaFecha(cCtaCodDIF, cCtaCodRFC, cCtaCodRFA, IIf(ChkRefinanciados.value = 1, True, False)), "#0.00")
        If cmdGrabar.Enabled And cmdGrabar.Visible Then cmdGrabar.SetFocus
        
     End If
End Sub

Sub ControlesTransaccion()
 cmdGrabar.Enabled = True
 CmdPlanPagos.Enabled = True
 cmdmora.Enabled = True
End Sub

Sub PonerColorTotales()
    Dim i As Integer
    Dim j As Integer
        
        For i = 0 To MSH.Rows - 1
            MSH.Row = i
            MSH.Col = 9
            MSH.FillStyle = flexFillSingle
            MSH.CellBackColor = vbGreen
                        
        Next i
End Sub

'Sub MontoProximos(ByVal pnCuota As Integer)
'Dim objDRFA As COMDCredito.DCOMRFA
'Dim rs As ADODB.Recordset
'
'    Set objDRFA = New COMDCredito.DCOMRFA
'    Set rs = objDRFA.ListaCreditosPendientesCuota(TxtBuscar1.Text, gdFecSis, pnCuota, IIf(ChkRefinanciados.value = 1, True, False))
'    Set objDRFA = Nothing
'    Do Until rs.EOF
'        If rs!cCredito = "DIF" Then
'            nMontoProDIF = rs!Capital + rs!IntComp + rs!IntMor + rs!Gastos + rs!ComCofide
'        ElseIf rs!cCredito = "RFC" Then
'            nMontoProxRFC = rs!Capital + rs!IntComp + rs!IntMor + rs!Gastos + rs!ComCofide
'        ElseIf rs!cCredito = "RFA" Then
'            nMontoProxRFA = rs!Capital + rs!IntComp + rs!IntMor + rs!Gastos + rs!ComCofide
'        End If
'    rs.MoveNext
'    Loop
'    Set rs = Nothing
'
'End Sub

'Public Function ObtenerDeudaFecha(Optional ByVal psCtaCodDIF As String, Optional ByVal psCtaCodRFC As String, Optional ByVal psCtaCodRFA As String, _
'                                Optional ByVal pbRefinac As Boolean = False) As Double
'
'    Dim LnMontoRFC As Double
'    Dim LnMontoDIF As Double
'    Dim LnMontoRFA As Double
'    Dim lnMonto As Double
'    Dim LnInteresFecha As Double
'    Dim LMatCalend As Variant
'    Dim oNCredito As COMNCredito.NCOMCredito
'    Dim oD  As COMDCredito.DCOMRFA
'
'    Set oNCredito = New COMNCredito.NCOMCredito
'
'    If pbRefinac = False Then
'        If psCtaCodDIF <> "" Then
'            LMatCalend = oNCredito.RecuperaMatrizCalendarioPendiente(psCtaCodDIF, , True)
'            LnInteresFecha = oNCredito.MatrizInteresGastosAFecha(psCtaCodDIF, LMatCalend, gdFecSis, True, False)
'            lnMonto = oNCredito.MatrizCapitalAFecha(psCtaCodDIF, LMatCalend)
'            LnMontoDIF = LnInteresFecha + lnMonto
'            nMontoDIFCancel = LnMontoDIF
'            'verificando si tiene desaguio
'            nInteresDesaguioDIF = 0
'            If LnInteresFecha < 0 Then
'                nInteresDesaguioDIF = Abs(LnInteresFecha)
'            End If
'        End If
'   End If
'
'    If psCtaCodRFC <> "" Then
'        LMatCalend = oNCredito.RecuperaMatrizCalendarioPendiente(psCtaCodRFC, , True)
'        LnInteresFecha = oNCredito.MatrizInteresGastosAFecha(psCtaCodRFC, LMatCalend, gdFecSis, True, False)
'        lnMonto = oNCredito.MatrizCapitalAFecha(psCtaCodRFC, LMatCalend)
'        LnMontoRFC = LnInteresFecha + lnMonto
'        nMontoRFCCancel = LnMontoRFC
'        'verificando si tiene desaguio
'        nInteresDesaguioRFC = 0
'        If LnInteresFecha < 0 Then
'            nInteresDesaguioRFC = Abs(LnInteresFecha)
'        End If
'
'    End If
'    If psCtaCodRFA <> "" Then
'        Set oD = New COMDCredito.DCOMRFA
'        LMatCalend = oNCredito.RecuperaMatrizCalendarioPendiente(psCtaCodRFA, , True)
'        LnInteresFecha = oD.MatrizInteresGastosAFecha(psCtaCodRFA, LMatCalend, gdFecSis, True, False)
'        Set oD = Nothing
'        lnMonto = oNCredito.MatrizCapitalAFecha(psCtaCodRFA, LMatCalend)
'        LnMontoRFA = LnInteresFecha + lnMonto
'        nMontoRFACancel = LnMontoRFA
'        nInteresDesaguioRFA = 0
'        If LnInteresFecha < 0 Then
'            nInteresDesaguioRFA = Abs(LnInteresFecha)
'        End If
'    End If
'    Set oNCredito = Nothing
'    ObtenerDeudaFecha = LnMontoRFC + LnMontoDIF + LnMontoRFA
'End Function
