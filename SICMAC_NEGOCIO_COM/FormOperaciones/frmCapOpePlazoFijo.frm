VERSION 5.00
Begin VB.Form frmCapOpePlazoFijo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmCapOpePlazoFijo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraComisionTransf 
      Height          =   375
      Left            =   4620
      TabIndex        =   122
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CheckBox chkComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   123
         Top             =   60
         Width           =   705
      End
      Begin VB.Label lblComisionL 
         Caption         =   "Comision: "
         Height          =   255
         Left            =   0
         TabIndex        =   125
         Top             =   60
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   315
         Left            =   1560
         TabIndex        =   124
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame fraTransBco 
      Caption         =   "Documento"
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
      ForeColor       =   &H00000080&
      Height          =   3555
      Left            =   120
      TabIndex        =   110
      Top             =   5325
      Visible         =   0   'False
      Width           =   4980
      Begin VB.ComboBox cboPlaza 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   1150
         Width           =   1935
      End
      Begin VB.CheckBox chkMismoTitular 
         Caption         =   "Mismo Titular"
         Height          =   255
         Left            =   960
         TabIndex        =   114
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtCuentaDestino 
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   113
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtGlosaTransBco 
         Height          =   945
         Left            =   960
         TabIndex        =   112
         Top             =   2445
         Width           =   3855
      End
      Begin VB.TextBox txtTitular 
         Height          =   315
         Left            =   2925
         TabIndex        =   111
         Top             =   1530
         Width           =   1770
      End
      Begin SICMACT.TxtBuscar txtBancoDestino 
         Height          =   315
         Left            =   960
         TabIndex        =   116
         Top             =   360
         Width           =   1935
         _extentx        =   3413
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmCapOpePlazoFijo.frx":030A
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label Label28 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   390
         Width           =   615
      End
      Begin VB.Label lblBancoDestino 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   120
         Top             =   750
         Width           =   3855
      End
      Begin VB.Label Label27 
         Caption         =   "Plaza:"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label30 
         Caption         =   "CCI:"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   1965
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.Frame FraCargoCta 
      Caption         =   "Cuenta Cargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   106
      Top             =   4680
      Width           =   8730
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   0
         TabIndex        =   107
         Top             =   220
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcta      =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblTitularCargoCta 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4440
         TabIndex        =   109
         Top             =   240
         Width           =   4185
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
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
         Left            =   3720
         TabIndex        =   108
         Top             =   285
         Width           =   675
      End
   End
   Begin VB.Frame FRMedio 
      Height          =   495
      Left            =   5160
      TabIndex        =   102
      Top             =   6810
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ComboBox cboMedioRetiro 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   120
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label lblMedioRetiro 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Retiro :"
         Height          =   195
         Left            =   120
         TabIndex        =   104
         Top             =   150
         Visible         =   0   'False
         Width           =   1240
      End
   End
   Begin VB.Frame fraTranferecia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1785
      Left            =   120
      TabIndex        =   83
      Top             =   6000
      Visible         =   0   'False
      Width           =   4980
      Begin VB.ComboBox cboTransferMoneda 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   165
         Width           =   1575
      End
      Begin VB.CommandButton cmdTranfer 
         Height          =   350
         Left            =   2520
         Picture         =   "frmCapOpePlazoFijo.frx":0336
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   480
         Width           =   475
      End
      Begin VB.Label lblEtiMonTra 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transacción"
         Height          =   195
         Left            =   30
         TabIndex        =   101
         Top             =   1380
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblSimTra 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1560
         TabIndex        =   100
         Top             =   1350
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblMonTra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1920
         TabIndex        =   99
         Top             =   1320
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lbltransferBco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   855
         TabIndex        =   94
         Top             =   900
         Width           =   3450
      End
      Begin VB.Label lbltransferN 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   30
         TabIndex        =   93
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lbltransferBcol 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   30
         TabIndex        =   92
         Top             =   975
         Width           =   555
      End
      Begin VB.Label lblTrasferND 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   855
         TabIndex        =   91
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label lblTransferMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   30
         TabIndex        =   90
         Top             =   225
         Width           =   585
      End
      Begin VB.Label lblTTCC 
         Caption         =   "TCC"
         Height          =   285
         Left            =   3315
         TabIndex        =   89
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label25 
         Caption         =   "TCV"
         Height          =   285
         Left            =   3300
         TabIndex        =   88
         Top             =   540
         Width           =   345
      End
      Begin VB.Label lblTTCCD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3690
         TabIndex        =   87
         Top             =   165
         Width           =   630
      End
      Begin VB.Label lblTTCVD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3675
         TabIndex        =   86
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.Frame fraAumDismCap 
      Caption         =   "Datos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3255
      Left            =   120
      TabIndex        =   73
      Top             =   5310
      Visible         =   0   'False
      Width           =   4980
      Begin VB.CommandButton cmdCheque 
         Height          =   375
         Left            =   2250
         Picture         =   "frmCapOpePlazoFijo.frx":0778
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1110
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox txtPlazo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   765
         TabIndex        =   2
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txtGlosaAumDism 
         Height          =   660
         Left            =   600
         TabIndex        =   3
         Top             =   2475
         Width           =   4260
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   1605
         Width           =   510
      End
      Begin VB.Label lblNombreIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   765
         TabIndex        =   81
         Top             =   1515
         Width           =   4065
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label lblNroDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   765
         TabIndex        =   79
         Top             =   1110
         Width           =   1410
      End
      Begin VB.Line Line2 
         X1              =   45
         X2              =   4950
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4950
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label lblTEA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   2700
         TabIndex        =   77
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "TEA(%):"
         Height          =   195
         Left            =   2025
         TabIndex        =   76
         Top             =   345
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   120
         TabIndex        =   75
         Top             =   345
         Width           =   435
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   2715
         Width           =   555
      End
   End
   Begin VB.Frame fraTran 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3255
      Left            =   90
      TabIndex        =   59
      Top             =   5310
      Visible         =   0   'False
      Width           =   4980
      Begin VB.TextBox txtglosaTrans 
         Height          =   1080
         Left            =   1140
         TabIndex        =   63
         Top             =   1170
         Width           =   3765
      End
      Begin VB.ComboBox cboMonedaBanco 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   705
         Width           =   1050
      End
      Begin VB.TextBox txtCtaBanco 
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   60
         Top             =   300
         Width           =   1935
      End
      Begin SICMACT.TxtBuscar txtBanco 
         Height          =   315
         Left            =   3000
         TabIndex        =   62
         Top             =   300
         Width           =   1755
         _extentx        =   3096
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmCapOpePlazoFijo.frx":0A82
         psraiz          =   "BANCOS"
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   393
         Width           =   915
      End
      Begin VB.Label lblOrdenPago 
         AutoSize        =   -1  'True
         Caption         =   "Orden Pago :"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label lblBanco 
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
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   1140
         TabIndex        =   66
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblEtqBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   2520
         TabIndex        =   65
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblCtaBanco 
         Caption         =   "Cta Banco :"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame fraMonto 
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3570
      Left            =   5130
      TabIndex        =   35
      Top             =   5325
      Width           =   3720
      Begin VB.CheckBox chkVBEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   945
         TabIndex        =   97
         Top             =   1680
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   945
         TabIndex        =   50
         Top             =   2160
         Width           =   705
      End
      Begin VB.Frame fraDatCancelacion 
         Height          =   945
         Left            =   360
         TabIndex        =   36
         Top             =   195
         Width           =   2595
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Int. Pagar:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblIntGanado 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   960
            TabIndex        =   39
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblCapital 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   960
            TabIndex        =   38
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Capital :"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   195
            Width           =   570
         End
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   1005
         TabIndex        =   5
         Top             =   1185
         Width           =   1815
         _extentx        =   3201
         _extenty        =   661
         font            =   "frmCapOpePlazoFijo.frx":0AAE
         backcolor       =   12648447
         forecolor       =   192
         text            =   "0"
      End
      Begin VB.Frame fraAbonoOtraMoneda 
         Height          =   945
         Left            =   345
         TabIndex        =   43
         Top             =   195
         Width           =   2595
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Monto Abono :"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   570
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio :"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   975
         End
         Begin VB.Label lblMontoAbono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1260
            TabIndex        =   45
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label lblTipoCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1260
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Total Efect:"
         Height          =   195
         Left            =   30
         TabIndex        =   127
         Top             =   3210
         Width           =   825
      End
      Begin VB.Label lblTotalEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1020
         TabIndex        =   126
         Top             =   3180
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label lblEtqComi 
         Caption         =   "Comision :"
         Height          =   195
         Left            =   165
         TabIndex        =   96
         Top             =   1680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblMonComision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   1725
         TabIndex        =   95
         Top             =   1627
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Final:"
         Height          =   195
         Left            =   165
         TabIndex        =   72
         Top             =   2865
         Width           =   825
      End
      Begin VB.Label lblSaldoFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1020
         TabIndex        =   71
         Top             =   2820
         Width           =   1800
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   165
         TabIndex        =   54
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total Efect:"
         Height          =   195
         Left            =   90
         TabIndex        =   53
         Top             =   2505
         Width           =   825
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   300
         Left            =   1725
         TabIndex        =   52
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1020
         TabIndex        =   51
         Top             =   2445
         Width           =   1800
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
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
         Height          =   240
         Left            =   2880
         TabIndex        =   42
         Top             =   1245
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   165
         TabIndex        =   41
         Top             =   1245
         Width           =   540
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1680
      Left            =   90
      TabIndex        =   20
      Top             =   0
      Width           =   8730
      Begin VB.Frame fraDatos 
         Height          =   945
         Left            =   105
         TabIndex        =   21
         Top             =   600
         Width           =   8580
         Begin VB.Label lblEtqTasaCanc 
            AutoSize        =   -1  'True
            Caption         =   "TEA Canc (%) :"
            Height          =   195
            Left            =   7005
            TabIndex        =   49
            Top             =   255
            Width           =   1080
         End
         Begin VB.Label lblTasaCanc 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   7230
            TabIndex        =   48
            Top             =   540
            Width           =   930
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   135
            TabIndex        =   33
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lblApertura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   1140
            TabIndex        =   32
            Top             =   195
            Width           =   2040
         End
         Begin VB.Label lblEtqUltCnt 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (días) :"
            Height          =   195
            Left            =   3360
            TabIndex        =   31
            Top             =   255
            Width           =   930
         End
         Begin VB.Label lblPlazo 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4395
            TabIndex        =   30
            Top             =   195
            Width           =   705
         End
         Begin VB.Label lblDuplicados 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   4395
            TabIndex        =   29
            Top             =   540
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "# Duplicados :"
            Height          =   195
            Left            =   3300
            TabIndex        =   28
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lblVencimiento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   1140
            TabIndex        =   27
            Top             =   540
            Width           =   2040
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento :"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   593
            Width           =   960
         End
         Begin VB.Label lblTasa 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   6015
            TabIndex        =   25
            Top             =   195
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "TEA (%) :"
            Height          =   195
            Left            =   5280
            TabIndex        =   24
            Top             =   255
            Width           =   660
         End
         Begin VB.Label lblDias 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000001&
            Height          =   300
            Left            =   6015
            TabIndex        =   23
            Top             =   540
            Width           =   855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "# Días :"
            Height          =   195
            Left            =   5220
            TabIndex        =   22
            Top             =   600
            Width           =   585
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   3630
         _extentx        =   6403
         _extenty        =   661
         texto           =   "Cuenta N°:"
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   3840
         TabIndex        =   34
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2895
      Left            =   90
      TabIndex        =   15
      Top             =   1710
      Width           =   8730
      Begin VB.CommandButton cmdVerRegla 
         Caption         =   "&Ver Regla"
         Height          =   315
         Left            =   7200
         TabIndex        =   105
         Top             =   2060
         Visible         =   0   'False
         Width           =   1440
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1755
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   8535
         _extentx        =   15055
         _extenty        =   3096
         cols0           =   9
         highlight       =   1
         allowuserresizing=   3
         visiblepopmenu  =   -1
         encabezadosnombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig-Grupo-Presente"
         encabezadosanchos=   "250-1400-3400-1200-0-0-0-700-900"
         font            =   "frmCapOpePlazoFijo.frx":0ADA
         font            =   "frmCapOpePlazoFijo.frx":0B02
         font            =   "frmCapOpePlazoFijo.frx":0B2A
         font            =   "frmCapOpePlazoFijo.frx":0B52
         font            =   "frmCapOpePlazoFijo.frx":0B7A
         fontfixed       =   "frmCapOpePlazoFijo.frx":0BA2
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-X-X-X-X-X-X-X-8"
         listacontroles  =   "0-0-0-0-0-0-0-0-4"
         encabezadosalineacion=   "C-C-L-L-C-C-C-C-L"
         formatosedit    =   "0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         colwidth0       =   255
         rowheight0      =   300
         tipobuspersona  =   1
      End
      Begin VB.Label lblMinFirmas 
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   6030
         TabIndex        =   58
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Mínimo Firmas :"
         Height          =   195
         Left            =   4815
         TabIndex        =   57
         Top             =   2130
         Width           =   1110
      End
      Begin VB.Label lblAlias 
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1515
         TabIndex        =   56
         Top             =   2445
         Width           =   7080
      End
      Begin VB.Label Label17 
         Caption         =   "Alias de la Cuenta:"
         Height          =   225
         Left            =   135
         TabIndex        =   55
         Top             =   2505
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   2130
         Width           =   960
      End
      Begin VB.Label lblTipoCuenta 
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   1515
         TabIndex        =   18
         Top             =   2070
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "# Firmas :"
         Height          =   195
         Left            =   3285
         TabIndex        =   17
         Top             =   2130
         Width           =   690
      End
      Begin VB.Label lblFirmas 
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
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   4110
         TabIndex        =   16
         Top             =   2070
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7830
      TabIndex        =   7
      Top             =   8985
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6780
      TabIndex        =   6
      Top             =   9000
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   9000
      Width           =   1000
   End
   Begin VB.PictureBox pctCheque 
      Height          =   345
      Left            =   2025
      Picture         =   "frmCapOpePlazoFijo.frx":0BC8
      ScaleHeight     =   285
      ScaleWidth      =   120
      TabIndex        =   10
      Top             =   6420
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pctNotaAbono 
      Height          =   300
      Left            =   2445
      Picture         =   "frmCapOpePlazoFijo.frx":129A
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   5310
      Width           =   4980
      Begin VB.CheckBox ckbPorAfectacion 
         Caption         =   "Con Afectación"
         Height          =   195
         Left            =   1200
         TabIndex        =   98
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtGlosa 
         Height          =   1260
         Left            =   1140
         TabIndex        =   4
         Top             =   960
         Width           =   3735
      End
      Begin SICMACT.TxtBuscar txtCtaAhoAboInt 
         Height          =   315
         Left            =   1130
         TabIndex        =   70
         Top             =   360
         Width           =   2475
         _extentx        =   4366
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmCapOpePlazoFijo.frx":181C
         psraiz          =   "BANCOS"
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label lblCuentaAho 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Aho :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   930
      End
      Begin VB.Label lblDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   393
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCapOpePlazoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nmoneda As COMDConstantes.Moneda
Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nDocumento As COMDConstantes.tpoDoc
Dim nMontoRetiro As Double
Dim nTCC As Double, nTCV As Double
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim sOperacion As String
Dim sCaption As String
Dim gsPersCod As String
Dim nMontoMinAumentoCap As Double
Dim nTasaNominal As Double
'By Capi 07082008
Dim lbTasaPactada As Boolean
'
Dim lnTpoPrograma As Integer 'ande ers021-2018

'***********Datos para los cheques************
Public sCodIF As String
Public dFechaValorizacion As Date
Public lnDValoriza As Integer
Public sIFCuenta As String
'*********************************************

'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String
'*****************
Dim sMovNroAut As String
Dim lbITFCtaExonerada As Boolean
Dim nMontoPremio As Double

'******************** Transferencia**********************

Dim lnMovNroTransfer As Long
Dim lnTransferSaldo As Currency
Dim vnMontoDOC As Double
Dim fsPersCodTransfer As String '***Agregado por ELRO el 20120706, según OYP-RFC074-2012
Dim fsOpeCod As String '***Agregado por ELRO el 20120706, según OYP-RFC074-2012
Dim fnMovNroRVD As Long '***Agregado por ELRO el 20120706, según OYP-RFC074-2012
Dim fsPersNombreCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
Dim fsPersDireccionCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
Dim fsdocumentoCVME As String '***Agregado por ELRO el 20120717, según OYP-RFC024-2012

Dim sNumTarj As String
Dim sCuenta As String
Dim cGetValorOpe As String

Dim nRedondeoITF As Double 'BRGO 20110914
Dim fnDepositoPersRealiza As Boolean 'WIOR 20121114
Dim fnCondicion As Integer 'WIOR 20121114

' ***** Agregado Por RIRO el 20130501, Proyecto Ahorro - Poderes *****
Dim bProcesoNuevo As Boolean
Dim strReglas As String
' ***** Fin RIRO *****
Dim fnTpoCtaCargo As Integer 'JUEZ 20131218
Dim rsRelPersCtaCargo As ADODB.Recordset 'JUEZ 20131318
Dim nMontoPlazoFijo As Double
Dim oDocRec As UDocRec 'EJVG20140210
Dim bInstFinanc As Boolean 'JUEZ 20140414
'JUEZ 20141114 Nuevos parámetros **********
Dim bValidaAumCap As Boolean
Dim bParAumCap As Boolean
Dim nParAumCapMinSol As Double
Dim nParAumCapMinDol As Double
Dim nParAumCapCantMaxMes As Integer
Dim nCantAumCap As Integer
'END JUEZ *********************************

Private Sub ValidaTasaInteres()
If nOperacion = gPFAumCapTasaPactEfec Or nOperacion = gPFAumCapTasaPactChq Or nOperacion = gPFAumCapTasaPactTrans Then Exit Sub
If Trim(txtPlazo.Text) <> "" Then
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim nMonto As Double
    Dim nPlazo As Long
    If Trim(lblSaldoFinal.Caption) <> "" Then
        nMonto = CDbl(lblSaldoFinal.Caption)
        nPlazo = CLng(txtPlazo)
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nTasaNominal = clsDef.GetCapTasaInteres(gCapPlazoFijo, nmoneda, gCapTasaNormal, nPlazo, nMonto, gsCodAge)
        lblTEA.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
        Set clsDef = Nothing
    End If
End If
End Sub

'Funcion de Impresion de Boletas
Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta '& oImpresora.gPrnSaltoLinea '& oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String, sTipoCuenta As String
Dim nMonto As Double
Dim sCuenta As String
For i = 1 To grdCliente.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(i, 3), 4)))
    If nPersoneria = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.TitPersLavDineroDir = grdCliente.TextMatrix(i, 3)
            poLavDinero.TitPersLavDineroDoc = grdCliente.TextMatrix(i, 4)
            Exit For
        End If
    Else
        'WIOR 20121108
        If nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCliente.TextMatrix(i, 2)
        End If
        'WIOR FIN
        If nRelacion = gCapRelPersRepTitular Then
            poLavDinero.ReaPersLavDinero = grdCliente.TextMatrix(i, 1)
            poLavDinero.ReaPersLavDineroNom = grdCliente.TextMatrix(i, 2)
            poLavDinero.ReaPersLavDineroDir = grdCliente.TextMatrix(i, 3)
            poLavDinero.ReaPersLavDineroDoc = grdCliente.TextMatrix(i, 4)
            Exit For
        End If
    End If
Next i
nMonto = txtMonto.value
sCuenta = TxtCuenta.NroCuenta
sTipoCuenta = Trim(lblTipoCuenta.Caption)
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, False, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, False, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'End If
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsGen As COMDConstSistema.DCOMGeneral
    'Dim nTasaNominal As Double
    Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset
    Dim rsPar As ADODB.Recordset 'JUEZ 20141114
    Dim nEstado As COMDConstantes.CaptacEstado
    Dim nRow As Long
    Dim sMsg As String, sMoneda As String, sPersona As String
    Dim dUltRetInt As Date
    Dim bGarantia As Boolean
    Dim dRenovacion As Date, dApeReal As Date
    Dim nTasaCancelacion As Double
    
    Dim oDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    '----- MADM
    Dim lafirma As frmPersonaFirma
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim Rf As ADODB.Recordset
    '----- MADM
    Dim nTpoPrograma As Integer 'BRGO 20111220
    Dim nMontoAbono As Double 'BRGO 20111220
    grdCliente.lbEditarFlex = True ' RIRO SE CAMBIO A TRUE

    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    bGarantia = False
    If nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Then
        bGarantia = True
    End If
    '***Modificación ELRO 20120307, según Acta N° 039-2012/TI-D
    'sMsg = clsCap.ValidaCuentaOperacion(sCuenta, , bGarantia)
    sMsg = clsCap.ValidaCuentaOperacion(sCuenta, , bGarantia, nOperacion)
    '***Fin Modificación ELRO**********************************

    If sMsg = "" Then
        sMsg = ""
        Select Case nOperacion
            Case gPFCancEfec, gPFAumCapEfec, gPFAumCapchq, gPFAumCapTasaPactEfec, _
                    gPFAumCapTasaPactChq, gPFDismCapEfec, gPFAumCapTrans, gPFAumCapTasaPactTrans, gPFAumCapCargoCta 'JUEZ 20131218 gPFAumCapCargoCta
                If clsCap.TieneChequesValorizacion(sCuenta) Then
                    MsgBox "La cuenta posee cheques en valorización.", vbInformation, "Aviso"
                    Set clsCap = Nothing
                    Exit Sub
                End If
            Case gPFRetInt, gPFRetIntAboAho, gPFRetIntAdelantado
                sMsg = clsCap.ValidaRetiroIntPF(sCuenta, gdFecSis, nOperacion)
        End Select
        If sMsg = "" Then
            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set rsCta = New ADODB.Recordset
            Set rsCta = clsMant.GetDatosCuenta(sCuenta)
        
            If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
                '-- AVMM -- 16-06-2006 -- validar fecha de cancelacion
                If CDate(Format$(rsCta("dApertura"), "dd mmm yyyy ")) = CDate(gdFecSis) Then
                    MsgBox "Cuenta no puede ser Cancelada el mismo día de la Apertura", vbInformation, "Aviso"
                    Exit Sub
                End If
                '-- Validar si tiene plazo Fijo como Garantia -- AVMM --12-02-2007
                Dim cGarPF As COMDCaptaGenerales.DCOMCaptaMovimiento
                Set cGarPF = New COMDCaptaGenerales.DCOMCaptaMovimiento
                If cGarPF.BuscaGarantiaCreditosPlazoFijo(sCuenta) Then
                    MsgBox "Plazo Fijo no puede ser Cancelado por ser Garantia de un Crédito", vbInformation, "Aviso"
                    Exit Sub
                End If
                Set cGarPF = Nothing
            End If
        
            If Not (rsCta.EOF And rsCta.BOF) Then
                'JUEZ 20141114 Nuevos parámetros ******************************************
                Set oDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
                Set rsPar = oDef.GetCapParametroNew(gCapPlazoFijo, rsCta("nTpoPrograma"))
                Set oDef = Nothing
                bParAumCap = rsPar!bAumCap
                nParAumCapMinSol = rsPar!nAumCapMinSol
                nParAumCapMinDol = rsPar!nAumCapMinDol
                nParAumCapCantMaxMes = rsPar!nAumCapCantMaxMes
                Set rsPar = Nothing
        
                If bValidaAumCap Then
                    If bParAumCap Then
                        'Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
                            nCantAumCap = clsMant.ObtenerCantidadOperaciones(sCuenta, gCapMovDeposito, gdFecSis)
                        'Set clsMant = Nothing
                        
                        If nCantAumCap >= nParAumCapCantMaxMes Then
                            MsgBox "Se ha realizado el número máximo de Aumento de Capital para este Plazo Fijo", vbInformation, "Aviso"
                            Exit Sub
                        End If
                    Else
                        MsgBox "El Tipo de Plazo Fijo no permite Aumento de Capital", vbInformation, "Aviso"
                        Exit Sub
                    End If
                End If
                'END JUEZ *****************************************************************
                nEstado = rsCta("nPrdEstado")
                lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm")
                nmoneda = CLng(Mid(sCuenta, 9, 1))
                strReglas = IIf(IsNull(rsCta!cReglas), "", rsCta!cReglas) 'Agregado por RIRO el 20130501, Proyecto Ahorro - Poderes
                     
                'ITF INICIO
                lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
                fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
            
                If gbITFAsumidoPF Then
                    Me.chkITFEfectivo.Visible = False
                    Me.chkITFEfectivo.value = 0
                Else
                    Me.chkITFEfectivo.Visible = True
                    Me.chkITFEfectivo.value = 1
                End If
                'ITF FIN
            
                If nmoneda = gMonedaNacional Then
                    sMoneda = "MONEDA NACIONAL"
                    txtMonto.BackColor = &HC0FFFF
                    lblMon.Caption = "S/."
                Else
                    sMoneda = "MONEDA EXTRANJERA"
                    txtMonto.BackColor = &HC0FFC0
                    lblMon.Caption = "$"
                End If
                Me.lblITF.BackColor = txtMonto.BackColor
                Me.lblTotal.BackColor = txtMonto.BackColor
            
                'lblMensaje = "DEPOSITO A PLAZO FIJO" & Chr$(13) & sMoneda
                lblPlazo = Format$(rsCta("nPlazo"), "#,##0")
                txtPlazo.Text = Format$(rsCta("nPlazo"), "#,##0")
                lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
                lblVencimiento = Format(DateAdd("d", rsCta("nPlazo"), rsCta("dRenovacion")), "dd mmm yyyy")
                nTasaNominal = rsCta("nTasaInteres")
                lblTasa = Format$(ConvierteTNAaTEA(nTasaNominal), "#0.00")
                lblTEA = Format$(ConvierteTNAaTEA(nTasaNominal), "#0.00")
                nTipoCuenta = rsCta("nPrdCtaTpo")
                nPersoneria = rsCta("nPersoneria")
                lblFirmas = Format$(rsCta("nFirmas"), "#0")
                lblDuplicados = rsCta("nDuplicado")
                lnTpoPrograma = rsCta("nTpoPrograma") 'APRI20191128
                Me.lblAlias = IIf(IsNull(rsCta("cAlias")), "", rsCta("cAlias"))
                Me.lblMinFirmas = IIf(IsNull(rsCta("nFirmasMin")), "", rsCta("nFirmasMin"))
                'By Capi 07082008
                lbTasaPactada = rsCta("bTasaPactada")
                
                '*** BRGO 20111220 ***********************************
                nTpoPrograma = rsCta("nTpoPrograma")
                nMontoAbono = rsCta("nMontoAbono")
                Set clsGen = New COMDConstSistema.DCOMGeneral
                Set rsRel = clsGen.GetConstante(gCaptacSubProdPlazoFijo, , CStr(nTpoPrograma))
                lblMensaje = rsRel!cDescripcion
                Set clsGen = Nothing
                Set rsRel = Nothing
                '*** END BRGO ****************************************
                
                'Add By Gitu 23-08-2011 para cobro de comision por operacion sin tarjeta
                If sNumTarj = "" Then
                    cGetValorOpe = ""
                    If nmoneda = gMonedaNacional Then
                        cGetValorOpe = GetMontoDescuento(2117, 1, 1)
                    Else
                        cGetValorOpe = GetMontoDescuento(2118, 1, 2)
                    End If
                    lblMonComision = Format(cGetValorOpe, "#,##0.00")
                End If
                'End Gitu
                txtPlazo.Locked = False 'BRGO 20111220
                Select Case nOperacion
                    'RIRO20131219 ERS137, Se agregó "gPFRetIntAboCtaBanco"
                    Case gPFRetInt, gPFRetIntAboAho, gPFRetIntAdelantado, gPFRetIntAboCtaBanco
                        If nOperacion = gPFRetIntAdelantado Then
                            nMontoRetiro = clsCap.GetMontoRetiroIntPF(rsCta("nSaldo"), rsCta("dRenovacion"), nTasaNominal, gdFecSis, 0, , True, rsCta("nPlazo"))
                        Else
                            nMontoRetiro = clsCap.GetMontoRetiroIntPF(rsCta("nSaldo"), rsCta("dUltCierre"), nTasaNominal, gdFecSis, rsCta("nIntAcum"))
                        End If
                        fraMonto.Enabled = True
                        txtCtaAhoAboInt.Text = rsCta("cCtaCodAbono")
                        dUltRetInt = clsCap.GetFechaUltimoRetiroIntPF(sCuenta)
                        lblDias = Format$(DateDiff("d", dUltRetInt, gdFecSis), "#0")
                    
                    Case gPFCancEfec, gPFCancTransf
                        'By Capi 07082008
                            'nMontoRetiro = clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, gsCodAge, nTasaCancelacion, , , , nMontoPremio)
                            nMontoRetiro = clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, gsCodAge, nTasaCancelacion, , , lbTasaPactada, nMontoPremio)
                        '
                        If nMontoPremio > 0 Then
                            fraDatCancelacion.Height = 1280
                            Label9.Top = Label9.Top + 345
                            txtMonto.Top = txtMonto.Top + 345
                            Label15.Top = Label15.Top + 345
                            chkITFEfectivo.Top = chkITFEfectivo.Top + 345
                            lblITF.Top = lblITF.Top + 345
                            Label16.Top = Label16.Top + 345
                            lblTotal.Top = lblTotal.Top + 345
                        End If
                        
                        lblTasaCanc = Format$(ConvierteTNAaTEA(nTasaCancelacion), "#,##0.00")
                        lblCapital = Format$(rsCta("nSaldo"), "#,##0.00")
                        lblIntGanado = Format$(nMontoRetiro - rsCta("nSaldo"), "#,##0.00")
                        lblDias = Format$(DateDiff("d", rsCta("dRenovacion"), gdFecSis), "#0")
                        fraMonto.Enabled = False
                        
                    Case gPFAumCapEfec, gPFAumCapchq, gPFAumCapTasaPactEfec, gPFAumCapTasaPactChq, gPFDismCapEfec, gPFAumCapTrans, gPFAumCapTasaPactTrans, gPFAumCapCargoCta 'JUEZ 20131218 gPFAumCapCargoCta
                        '*** BRGO 20111220 ***********************************
                        If nTpoPrograma = 0 Then
                            MsgBox "Esta operación no se aplica a cuentas de Plazo Fijo Clásico", vbInformation, "Aviso"
                            TxtCuenta.Age = ""
                            TxtCuenta.Cuenta = ""
                            TxtCuenta.SetFocusAge
                            Exit Sub
                        End If
                        txtPlazo.Locked = True
                        '*** END BRGO ****************************************
                        If nOperacion = gPFAumCapTasaPactEfec Or nOperacion = gPFAumCapTasaPactChq Then
                            nMontoRetiro = clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, Mid(sCuenta, 4, 2), nTasaCancelacion, , , True)
                        Else
                            nMontoRetiro = clsCap.GetSaldoCancelacion(sCuenta, gdFecSis, Mid(sCuenta, 4, 2), nTasaCancelacion, , , False, , True)
                        End If
                        lblTasaCanc = Format$(ConvierteTNAaTEA(nTasaCancelacion), "#,##0.00")
                        lblCapital = Format$(rsCta("nSaldo"), "#,##0.00")
                        lblIntGanado = Format$(nMontoRetiro - rsCta("nSaldo"), "#,##0.00")
                        lblDias = Format$(DateDiff("d", rsCta("dRenovacion"), gdFecSis), "#0")
                        If nOperacion <> gPFAumCapchq And nOperacion <> gPFAumCapTasaPactChq _
                           And nOperacion <> gPFAumCapTrans And nOperacion <> gPFAumCapTasaPactTrans Then
                            fraMonto.Enabled = True
                        End If
                        nMontoRetiro = 0
                        Set oDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
                        '*** BRGO 20111226 *****************************
                        If nTpoPrograma <> 2 And nTpoPrograma <> 3 Then
                            'nMontoMinAumentoCap = oDef.GetMontoMinimoAperturaPersoneria(gCapPlazoFijo, nMoneda, nPersoneria, False)
                            nMontoMinAumentoCap = IIf(nmoneda = gMonedaNacional, nParAumCapMinSol, nParAumCapMinDol) 'JUEZ 20141114
                        Else
                            nMontoMinAumentoCap = nMontoAbono
                        End If
                        '*** END BRGO **********************************
                        Set oDef = Nothing
                        fraAumDismCap.Enabled = True
                        'txtMonto.SetFocus
                        '***Agregado por ELRO el 20121015, según OYP-RFC024-2012
                        If nOperacion = gPFAumCapTrans Then
                            If nmoneda = gMonedaNacional Then
                                cboTransferMoneda.ListIndex = 0
                            Else
                                cboTransferMoneda.ListIndex = 1
                            End If
                            cboTransferMoneda.Enabled = False
                        End If
                        '***Fin Agregado por ELRO el 20121015*******************
                        If nOperacion = gPFAumCapCargoCta Then FraCargoCta.Enabled = True 'JUEZ 20131218
                End Select
                
                '***Agregado por ELRO el 20130722, según TI-ERS079-2013****
                'RIRO20131218 ERS137 - se incluyó a gPFRetIntAboCtaBanco
                If nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Or nOperacion = gPFCancEfec Or _
                nOperacion = gPFRetIntAboCtaBanco Then
                     
                     cargarMediosRetiros
                     
                End If
                '***Fin Agregado por ELRO el 20130722, según TI-ERS079-2013
                
                 ' RIRO20131210 ERS137
                If nOperacion = gPFRetIntAboCtaBanco Or nOperacion = gPFCancTransf Then
                    lblMedioRetiro.Visible = True
                    cboMedioRetiro.Visible = True
                    cargarMediosRetiros
                    cboMedioRetiro.Text = "TRANSFERENCIA BANCO                                                                                                    3"
                    cboMedioRetiro.Enabled = False
                    CalculaComision
                    fraTransBco.Enabled = True
                    txtBancoDestino.SetFocus
                End If
                ' FIN RIRO
                
                txtMonto.Text = Format$(nMontoRetiro, "#,##0.00")
                Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
                Set clsMant = Nothing
            
                sPersona = ""
                Do While Not rsRel.EOF
                    If rsRel("cPersCod") = gsCodPersUser Then
                        MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                        Unload Me
                        Exit Sub
                    End If
                    If sPersona <> rsRel("cPersCod") Then
                        grdCliente.AdicionaFila
                        nRow = grdCliente.Rows - 1
                        grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                        gsPersCod = rsRel("cPersCod")
                        grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                        grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & space(50) & Trim(rsRel("nPrdPersRelac"))
                        grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion")
                        grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                        'grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("bobligatorio")) Or rsRel("bobligatorio") = False, "NO", "SI")
                        
                    ' ***** Agregado por RIRO *****
                    
                        If rsRel("cGrupo") <> "" Then
                            bProcesoNuevo = True
                            grdCliente.TextMatrix(nRow, 7) = rsRel("cGrupo")
                            
                        Else
                            bProcesoNuevo = False
                            grdCliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("bobligatorio")) Or rsRel("bobligatorio") = False, "NO", "SI")
                            
                        End If

                    ' ***** Fin RIRO *****
                        
                        sPersona = rsRel("cPersCod")
                    End If
                    rsRel.MoveNext
                Loop
                
                'JUEZ 20140414 ****************************************
                If nOperacion = gPFCancEfec Or nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Or _
                   nOperacion = gPFRetIntAdelantado Or nOperacion = gPFAumCapEfec Or nOperacion = gPFAumCapTrans Then
                    Dim i As Integer
                    For i = 1 To grdCliente.Rows - 1
                        If Trim(Left(grdCliente.TextMatrix(i, 3), 10)) = "TITULAR" Then
                            Dim oDInstFinan As COMDPersona.DCOMInstFinac
                            Set oDInstFinan = New COMDPersona.DCOMInstFinac
                            bInstFinanc = oDInstFinan.VerificaEsInstFinanc(Trim(grdCliente.TextMatrix(i, 1)))
                            Set oDInstFinan = Nothing
                            txtMonto_Change
                        End If
                    Next
                End If
                'END JUEZ *********************************************
                
                ' COMENTADO POR RIRO SEGUN PODERES
                  '********* firma madm
                        'If sPersona <> "" Then
                        ' Set lafirma = New frmPersonaFirma
                        ' Set ClsPersona = New COMDPersona.DCOMPersonas
                        '
                        ' Set Rf = ClsPersona.BuscaCliente(gsPersCod, BusquedaCodigo)
                        '
                        ' If Not Rf.BOF And Not Rf.EOF Then
                        '    If Rf!nPersPersoneria = 1 Then
                        '    Call frmPersonaFirma.Inicio(Trim(gsPersCod), Mid(gsPersCod, 4, 2), False, True)
                        '    End If
                        ' End If
                        ' Set Rf = Nothing
                        'End If
                     '*******************
                
                If bProcesoNuevo Then
                    cmdVerRegla.Visible = True
                    Label19.Visible = False
                    lblMinFirmas.Visible = False
                    Label8.Visible = False
                    lblFirmas.Visible = False
                    grdCliente.ColWidth(6) = 0
                    grdCliente.ColWidth(7) = 800
                    grdCliente.ColWidth(8) = 900
                
                Else
                    cmdVerRegla.Visible = False
                    Label19.Visible = True
                    lblMinFirmas.Visible = True
                    Label8.Visible = True
                    lblFirmas.Visible = True
                    grdCliente.ColWidth(6) = 1200
                    grdCliente.ColWidth(7) = 0
                    grdCliente.ColWidth(8) = 0
                    
                    MsgBox "Se recomienda actualizar los grupos y reglas de la cuenta", vbExclamation, "Aviso"
                    Set lafirma = New frmPersonaFirma
                    Set ClsPersona = New COMDPersona.DCOMPersonas

                    Set Rf = ClsPersona.BuscaCliente(gsPersCod, BusquedaCodigo)

                    If Not Rf.BOF And Not Rf.EOF Then
                       If Rf!nPersPersoneria = 1 Then
                       Call frmPersonaFirma.Inicio(Trim(gsPersCod), Mid(gsPersCod, 4, 2), False, True)
                       End If
                    End If
                    Set Rf = Nothing
                    
                End If
                
                rsRel.Close
                Set rsRel = Nothing
                fraCuenta.Enabled = False
                fraCliente.Enabled = True
                fraDocumento.Enabled = True
            
                If txtPlazo.Visible Then
                    If Not cmdCheque.Enabled Then
                        txtMonto.Enabled = True
                    End If
                    'txtPlazo.SetFocus
                Else
                    If txtCtaAhoAboInt.Visible Then
                        'txtCtaAhoAboInt.SetFocus
                    Else
                        If fraDocumento.Visible Then
                            'txtGlosa.SetFocus
                        Else
                            'Me.txtglosaTrans.SetFocus
                        End If
                    End If
                End If
                cmdGrabar.Enabled = True
                cmdCancelar.Enabled = True
                
                '***Agregado por ELRO el 20120823, según OYP-RFC024-2012
                If nOperacion = gPFAumCapTrans Then
                    fraTranferecia.Enabled = True
                    chkITFEfectivo.value = 0
                ElseIf nOperacion = gPFRetIntAboCtaBanco Or nOperacion = gPFCancTransf Then ' RIRO20131219 ERS137
                    fraTransBco.Enabled = True
                End If
                '***Fin Agregado por ELRO el 20120823*******************
            
                'MuestraFirmas sCuenta      MADM ARRIBA IMPLMENTADO
' AVMM -- 23-03-2007 -- NUNCA SE BLOQUEA EN LA CMACT-MAYNAS

'                'CONTROL DE CANCELACION DE PLAZOS FIJOS SI SUS TIRULARES TIENEN CREDITOS
'                If Not ValidaCreditosPendientes(sCuenta) Then
'                    MsgBox "La cuenta ha sido bloqueada para retiro, pues posee créditos pendientes de pago.", vbInformation, "Operacion"
'                    LimpiaControles
'                    txtCuenta.SetFocus
'                    Exit Sub
'                End If
'
'                '*********************************
                If Me.FraCargoCta.Visible Then Me.txtCuentaCargo.SetFocusCuenta 'JUEZ 20131218
            End If
        Else
            MsgBox sMsg, vbInformation, "Operacion"
            TxtCuenta.SetFocus
            Exit Sub
        End If
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        TxtCuenta.SetFocus
        Exit Sub
    End If
    Set clsCap = Nothing
End Sub

Private Sub LimpiaControles()
    grdCliente.Clear
    grdCliente.Rows = 2
    grdCliente.FormaCabecera
    txtGlosa = ""
    txtMonto.value = 0
    cmdGrabar.Enabled = False
    TxtCuenta.Age = ""
    TxtCuenta.Cuenta = ""
    txtCtaAhoAboInt.Text = ""
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    lblApertura = ""
    lblFirmas = ""
    lblTipoCuenta = ""
    lblPlazo = ""
    lblDias = ""
    lblVencimiento = ""
    lblDuplicados = ""
    lblTasa = ""
    lblTasaCanc = ""
    lblCapital = ""
    lblMensaje = ""
    lblIntGanado = ""
    fraCuenta.Enabled = True
    fraCliente.Enabled = False
    fraDatos.Enabled = False
    fraDocumento.Enabled = False
    fraAumDismCap.Enabled = False
    
    'RIRO20131219 ERS137
    If fraTransBco.Visible Then
    
    fraTransBco.Enabled = False
    txtCuentaDestino.Text = ""
    txtTitular.Text = ""
    lblBancoDestino.Caption = ""
    chkMismoTitular.value = 0
    txtBancoDestino.Text = ""
    txtGlosaTransBco.Text = ""
    cboMedioRetiro.Enabled = False
    cboPlaza.ListIndex = 1
    nMontoPlazoFijo = -1
    End If
    ' ENN RIRO
    
    nMontoRetiro = 0
    TxtCuenta.SetFocus
    sMovNroAut = ""
    lblSaldoFinal.Caption = ""
    lblAlias.Caption = ""
    lblMinFirmas.Caption = ""
    lblNroDoc.Caption = ""
    lblNombreIF.Caption = ""
    lnDValoriza = 0
    sCodIF = ""
    dFechaValorizacion = gdFecSis
    txtPlazo.Text = "0"
    lblTEA.Caption = ""
    txtGlosaAumDism.Text = ""
    'Controles de Dscto Premio
    lnMovNroTransfer = -1
    lnTransferSaldo = 0
    '***Agregado por ELRO el 20120810, según OYP-RFC024-2012
    Me.cboTransferMoneda.Enabled = False
    lblTrasferND = ""
    lbltransferBco = ""
    lblMonTra = "0.00"
    fsPersCodTransfer = ""
    fsOpeCod = ""
    fnMovNroRVD = 0
    fsPersNombreCVME = ""
    fsPersDireccionCVME = ""
    fsdocumentoCVME = ""
    '***Fin Agregado por ELRO el 20120810, según OYP-RFC024-2012
    fraDatCancelacion.Height = 945
    nRedondeoITF = 0
    '***Agregado por ELRO el 20130724, según TI-ERS079-2013****
    If cboMedioRetiro.Visible Then
        cargarMediosRetiros
    End If
    '***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013
    'EJVG20130916 ***
    chkITFEfectivo.value = 0
    fraMonto.Enabled = False
    'END EJVG *******
    'JUEZ 20131218 ******************
    If FraCargoCta.Visible Then
        FraCargoCta.Enabled = False
        LimpiaControlesCargoCta
    End If
    'END JUEZ ***********************
    bInstFinanc = False 'JUEZ 20140414
End Sub

Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOperacion As String, _
            Optional sCodCmac As String = "", Optional sNomCmac As String)
    nOperacion = nOpe

    sNombreCMAC = sNomCmac
    sPersCodCMAC = sCodCmac
    sOperacion = sDescOperacion
    nMontoPlazoFijo = -1 ' RIRO20131212ERS137

    If sPersCodCMAC = "" Then
        Me.Caption = "Captaciones -  Plazo Fijo - " & sDescOperacion
    Else
        Me.Caption = "Captaciones -  Plazo Fijo - " & sDescOperacion & " - " & sNomCmac
    End If
    sCaption = Me.Caption

    Label17.Visible = True
    Label19.Visible = False ' RIRO CAMBIO A FLASE
    lblAlias.Visible = True
    lblMinFirmas.Visible = False ' RIRO CAMBIO A FLASE
    
    ' Agregado Por RIRO el 20130501, Proyecto Ahorros
    Label8.Visible = False
    lblFirmas.Visible = False
    ' Fin RIRO
    
    Label21.Visible = False
    lblSaldoFinal.Visible = False
    fraAbonoOtraMoneda.Visible = False
    fraAumDismCap.Visible = False

    Label24.Enabled = False
    Label26.Enabled = False
    lblNroDoc.Enabled = False
    lblNombreIF.Enabled = False
    cmdCheque.Enabled = False
    fraAumDismCap.Enabled = False
    FraCargoCta.Visible = False 'JUEZ 20131218

    Select Case nOperacion
        Case gPFRetInt, gPFRetIntAdelantado
            fraDatCancelacion.Visible = False
            lblCuentaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            fraTranferecia.Visible = False
        Case nOperacion = gPFRetIntAboAho
            fraMonto.Enabled = True
            fraDatCancelacion.Visible = False
            lblCuentaAho.Visible = True
            txtCtaAhoAboInt.Visible = True
            fraTranferecia.Visible = False
        Case gPFCancEfec
            fraDatCancelacion.Visible = True
            lblCuentaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            fraTranferecia.Visible = False
            '***Agregado por ELRO el 20120327, según RFC-018-2012
            ckbPorAfectacion.Visible = True
            '***Fin Agregado por ELRO****************************
        Case gPFAumCapEfec, gPFAumCapTasaPactEfec, gPFDismCapEfec
            fraDatCancelacion.Visible = True
            lblCuentaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            Label21.Visible = True
            lblSaldoFinal.Visible = True
            fraAumDismCap.Visible = True
            fraTranferecia.Visible = False
        Case gPFAumCapchq, gPFAumCapTasaPactChq
            fraDatCancelacion.Visible = True
            lblCuentaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            Label21.Visible = True
            lblSaldoFinal.Visible = True
            fraAumDismCap.Visible = True
            Label24.Enabled = True
            Label26.Enabled = True
            lblNroDoc.Enabled = True
            lblNombreIF.Enabled = True
            cmdCheque.Enabled = True
            fraTranferecia.Visible = False
        
        ' RIRO20131218 ERS137 Comentado
        'Case gPFCancTransf
        '    Dim oCon As COMDConstantes.DCOMConstantes, rsBanco As New ADODB.Recordset, clsBanco As COMNCajaGeneral.NCOMCajaCtaIF  'NCajaCtaIF
        '    Set oCon = New COMDConstantes.DCOMConstantes
        '    Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
        '    Set rsBanco = clsBanco.CargaCtasIF(gMonedaNacional, "0[123]%", MuestraInstituciones)
        '    Set clsBanco = Nothing
        '    fraTran.Visible = True
        '    fraDocumento.Visible = False
        '    txtBanco.rs = rsBanco
        '    CargaCombo Me.cboMonedaBanco, oCon.RecuperaConstantes(gMoneda)
        '    Set oCon = Nothing
        '    fraTranferecia.Visible = False
        
        ' RIRO20131218 ERS137 Agregado
        Case gPFRetIntAboCtaBanco, gPFCancTransf
        
            'Cargando Bancos
            Dim rsBancoDest As New ADODB.Recordset, clsBancoDest As COMNCajaGeneral.NCOMCajaCtaIF  'NCajaCtaIF
            
            Set clsBancoDest = New COMNCajaGeneral.NCOMCajaCtaIF
            Set rsBancoDest = clsBancoDest.CargaCtasIF(gMonedaNacional, "01%", MuestraInstituciones)
            Set clsBancoDest = Nothing
            txtBancoDestino.rs = rsBancoDest
            
            'Cargando Plaza
            Dim rsConstante As ADODB.Recordset
            Dim oConstante As COMDConstSistema.DCOMGeneral
            Set oConstante = New COMDConstSistema.DCOMGeneral
            Set rsConstante = oConstante.GetConstante("10032", , "'20[^0]'")
            CargaCombo cboPlaza, rsConstante
            If Len(Trim(TxtCuenta.Prod)) > 0 Then cboPlaza.ListIndex = 0
            Set rsConstante = Nothing
            Set oConstante = Nothing
           
            'Ordenando controles del lado derecha
            fraAbonoOtraMoneda.Visible = False
            FRMedio.Top = 5580
            FRMedio.Enabled = True
            FRMedio.Visible = True
            FRMedio.Left = 5250
            
            fraDatCancelacion.Top = 750
            fraDatCancelacion.Left = 420
            
            If nOperacion = gPFRetIntAboCtaBanco Then
                fraDatCancelacion.Visible = False
            Else
                fraDatCancelacion.Visible = True
            End If
                        
            txtMonto.Left = 1020
            txtMonto.Top = 1740
            txtMonto.Height = 315
            
            Label9.Top = 1830
            Label9.Left = 165
            lblMon.Top = 1800
            lblMon.Left = 2940
            
            'lblComision.Top = 2130
            'lblComision.Left = 1725
            '
            'chkComision.Top = 2205
            'chkComision.Left = 945
            'lblComisionL.Top = 2190
            'lblComisionL.Left = 165
            
            fraComisionTransf.Top = 7450
            fraComisionTransf.BorderStyle = 0
            fraComisionTransf.Height = 315
            fraComisionTransf.Left = 5295
            fraComisionTransf.Visible = True
            fraComisionTransf.Width = 2805
                        
            lblITF.Top = 2520
            lblITF.Left = 1725
            
            chkITFEfectivo.Top = 2580
            chkITFEfectivo.Left = 945
            Label15.Top = 2580
            Label15.Left = 165
            
            lblTotal.Top = 2865
            lblTotal.Left = 1020
            Label16.Top = 2925
            Label16.Left = 50
            
            lblSaldoFinal.Top = 3360
            lblSaldoFinal.Left = 1020
            
            Label21.Top = 3285
            Label21.Left = 165
            
            lblComision.Visible = True
            lblComisionL.Visible = True
            chkVBEfectivo.Visible = False
            lblEtqComi.Visible = False
            
            fraTransBco.Visible = True
            fraTranferecia.Visible = False
            fraTran.Visible = False
            fraDocumento.Visible = False
            Label16.Caption = "Total Transf: "
            lblTotalEfectivo.Left = 1020
            lblTotalEfectivo.Top = 3230
            
            lblTotalEfectivo.Visible = True
        Case gPFAumCapTasaPactTrans, gPFAumCapTrans
        
            IniciaCombo cboTransferMoneda, gMoneda
            fraDatCancelacion.Visible = True
            lblCuentaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            Label21.Visible = True
            lblSaldoFinal.Visible = True
            fraAumDismCap.Visible = True
            fraTranferecia.Visible = True
            '***Agregado por ELRO el 20120814, según OYP-RFC024-2012
            fraTranferecia.Enabled = False
            lblEtiMonTra.Visible = True
            lblSimTra.Visible = True
            lblMonTra.Visible = True
            '***Fin Agregado por ELRO el 20120814*******************
            'JUEZ 20131218 *******************
        Case gPFAumCapCargoCta
            FraCargoCta.Visible = True
            fraDatCancelacion.Visible = True
            lblCuentaAho.Visible = False
            txtCtaAhoAboInt.Visible = False
            Label21.Visible = True
            lblSaldoFinal.Visible = True
            fraAumDismCap.Visible = True
            fraTranferecia.Visible = False
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Age = gsCodAge
            txtCuentaCargo.Prod = gCapAhorros
            'END JUEZ ************************
    End Select
    'Verifica si la operacion necesita algun documento
    Dim rsDoc As ADODB.Recordset 'Recordset
    Dim clsTip As COMDConstSistema.NCOMTipoCambio 'nTipoCambio
    Set clsTip = New COMDConstSistema.NCOMTipoCambio
    nTCC = clsTip.EmiteTipoCambio(gdFecSis, TCCompra)
    nTCV = clsTip.EmiteTipoCambio(gdFecSis, TCVenta)
    Set clsTip = Nothing
    TxtCuenta.Prod = Trim(gCapPlazoFijo)
    TxtCuenta.EnabledProd = False
    TxtCuenta.CMAC = gsCodCMAC
    TxtCuenta.EnabledCMAC = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    fraCliente.Enabled = False
    fraDocumento.Enabled = False
    fraMonto.Enabled = False
    chkITFEfectivo.Enabled = False
    sMovNroAut = ""
    chkITFEfectivo.value = 1
    txtPlazo.Text = "0"
    lblNroDoc.Caption = ""
    lblNombreIF.Caption = ""
    
    '***Agregado por ELRO el 20120823, según OYP-RFC024-2012
    If nOperacion = gPFAumCapTrans Then
        chkITFEfectivo.value = 0
        chkITFEfectivo.Enabled = True
    End If
    '***Fin Agregado por ELRO el 20120823*******************
    '***Agregado por ELRO el 20130723, según TI-ERS079-2013****
    If nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Or nOperacion = gPFCancEfec Or nOperacion = gPFRetIntAboCtaBanco Or nOperacion = gPFCancTransf Then
        FRMedio.Visible = True
        lblMedioRetiro.Visible = True
        cboMedioRetiro.Visible = True
    End If
    '***Fin Agregado por ELRO el 20130723, según TI-ERS079-2013
    
    bInstFinanc = False 'JUEZ 20140414
    
    'JUEZ 20141114 Verificar si operación valida cantidad de depositos en mes ****
    Dim oCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    bValidaAumCap = oCapDef.ValidaCantOperaciones(nOperacion, gCapPlazoFijo, gCapMovDeposito)
    Set oCapDef = Nothing
    'END JUEZ ********************************************************************
    
    'ADD By GITU para el uso de las operaciones con tarjeta
    If gnCodOpeTarj = 1 And (gsOpeCod = "210201" Or gsOpeCod = "210202" Or gsOpeCod = "210206" Or gsOpeCod = "210301" Or gsOpeCod = "210302") Then
        sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(nOperacion), sNumTarj, 233)
        If sCuenta <> "123456789" Then
            If Val(Mid(sCuenta, 6, 3)) <> 233 And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                TxtCuenta.NroCuenta = sCuenta
                'txtCuenta.SetFocusCuenta
                ObtieneDatosCuenta sCuenta
            End If
            If sCuenta <> "" Then
                Me.Show 1
            End If
        Else
            lblEtqComi.Visible = True
            chkVBEfectivo.Visible = True
            lblMonComision.Visible = True
            
            Me.Show 1
        End If
    Else
        Me.Show 1
    End If
    'End GITU
    
End Sub
'***Agregado por ELRO el 20130723, según TI-ERS079-2013
Private Sub cboMedioRetiro_KeyPress(KeyAscii As Integer)
    'JUEZ 20130907 **************
    'If cmdGrabar.Enabled Then
    '    cmdGrabar.SetFocus
    'End If
    If KeyAscii = 13 Then
        If txtMonto.Enabled Then
            txtMonto.SetFocus
        Else
            cmdGrabar.SetFocus
        End If
    End If
    'END JUEZ *******************
End Sub
'***Fin Agregado por ELRO el 20130723, según TI-ERS079-2013

'***Agregado por ELRO el 20120823, según OYP-RFC024-2012
Private Sub cboTransferMoneda_Click()
        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
        lblSimTra.Caption = "S/."
        lblMonTra.BackColor = &HC0FFFF
    Else
        lblSimTra.Caption = "$"
        lblMonTra.BackColor = &HC0FFC0
    End If
End Sub
'***Fin Agregado por ELRO el 20120823*******************
Private Sub cboTransferMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdTranfer.SetFocus
    End If
End Sub

'EJVG20130916 ***
Private Sub chkITFEfectivo_Click()
    txtMonto_Change
End Sub
'END EJVG *******
Private Sub chkVBEfectivo_Click()
Dim nMonto As Double
nMonto = txtMonto.value
    'GITU 20110829
    If chkVBEfectivo.value = 1 And chkITFEfectivo.value = 1 Then
        If (nOperacion = "210201" Or nOperacion = "210202" Or nOperacion = "210206") Then
            lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption) + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
    ElseIf chkVBEfectivo.value = 1 And chkITFEfectivo.value = 0 Then
        If (nOperacion = "210201" Or nOperacion = "210202" Or nOperacion = "210206") Then
            lblTotal.Caption = Format(nMonto + lblMonComision, "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    ElseIf chkVBEfectivo.value = 0 And chkITFEfectivo.value = 1 Then
        If (nOperacion = "210201" Or nOperacion = "210202" Or nOperacion = "210206") Then
            lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
            Exit Sub
        End If
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    Else
        lblTotal.Caption = Format(nMonto, "#,##0.00")
    End If
    'lblTotal.Caption = Format(nMonto + CDbl(lblITF.Caption), "#,##0.00")
    'End GITU
End Sub
'***Agredo por ELRO el 20120327, según RFC-018-2012
Private Sub ckbPorAfectacion_Click()
    If ckbPorAfectacion Then
        txtGlosa = ckbPorAfectacion.Caption
        txtGlosa.SetFocus
    Else
        txtGlosa = ""
    End If
End Sub
'***Fin Agredo por ELRO****************************

Private Sub cmdCancelar_Click()
    LimpiaControles
End Sub

Private Sub cmdCheque_Click()
    'EJVG20140210 ***
    On Error GoTo ErrCmdCheque_Click
    'frmCapAperturaListaChq.Inicia frmCapOpePlazoFijo, nOperacion, nMoneda, gCapPlazoFijo, True
    'lblTotal.Caption = Format(txtMonto.value + CDbl(lblITF.Caption), "#,##0.00")
    Dim oForm As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque

    Set oDocRec = oForm.iniciarBusqueda(nmoneda, TipoOperacionCheque.DPF_AumentoCapital, TxtCuenta.NroCuenta)
    Set oForm = Nothing
    lblNroDoc.Caption = oDocRec.fsNroDoc
    sCodIF = oDocRec.fsPersCod
    sIFCuenta = oDocRec.fsIFCta
    lblNombreIF.Caption = oDocRec.fsPersNombre
    txtGlosaAumDism.Text = oDocRec.fsGlosa
    txtMonto.Text = Format(oDocRec.fnMonto, gsFormatoNumeroView)
    fraMonto.Enabled = True
    'END EJVG *******
    lblTotal.Caption = Format(CDbl(lblITF.Caption), "#,##0.00")
    txtGlosaAumDism.SetFocus
    Exit Sub
ErrCmdCheque_Click:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdGrabar_Click()

    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

'***Agregado por ELRO el 20130723, según TI-ERS079-2013****
If cboMedioRetiro.Visible Then
    If Trim(cboMedioRetiro) = "" Then
        MsgBox "Debe seleccionar el medio de retiro.", vbInformation, "Aviso"
        cboMedioRetiro.SetFocus
        Exit Sub
    End If
End If
'***Fin Agregado por ELRO el 20130723, según TI-ERS079-2013

 'ANDE 20180419 ERS021-2018 camapaña mundialito
Dim cperscod As String
Dim nTitularCod As Integer, nTipoPersona As Integer, bParticipaCamp As Boolean, cTextoDatos As String
Dim ix As Integer
Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
For ix = 1 To grdCliente.Rows - 1
    nTitularCod = Val(Right(Trim(grdCliente.TextMatrix(ix, 3)), 2))
    If nTitularCod = 10 Then
        'nTipoPersona = grdCliente.TextMatrix(ix, 1)
        cperscod = grdCliente.TextMatrix(ix, 1)
        nTipoPersona = oCaptaLN.getVerificarPersonaNatJur(cperscod)
    End If
Next ix
    'end ande


'WIOR 20130301 **************************
Dim fbPersonaReaAhorros As Boolean
Dim fnCondicion As Integer
Dim nI As Integer
nI = 0
'WIOR FIN *******************************
    Dim sNroDoc As String
    Dim nMonto As Double
    Dim sCuenta As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCapD As COMDCaptaGenerales.DCOMCaptaGenerales
    
    Dim sCtaAboInt As String
    'Dim nI As Integer

    Dim nMontoPremio As Double
    Dim nCantPremio As Integer
    Dim lsBoletaPremioCanc As String
    Dim oCred As COMDCredito.DCOMCredito
    Dim loLavDinero As frmMovLavDinero
    
    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    
    Set loLavDinero = New frmMovLavDinero

    sCuenta = TxtCuenta.NroCuenta
    nMonto = txtMonto.value
    
    'JUEZ 20131212 *****************************************************
    If nOperacion = gPFAumCapCargoCta Then
        If Len(txtCuentaCargo.NroCuenta) <> 18 Then
            MsgBox "Debe ingresar la cuenta de ahorros a la que se va a debitar el monto del Aumento a Capital", vbInformation, "Aviso"
            txtCuentaCargo.SetFocusCuenta
            Exit Sub
        End If
    End If
    'END JUEZ **********************************************************
    'FRHU ERS077-2015 20151203
    Dim nFila As Integer
    For nFila = 1 To Me.grdCliente.Rows - 1
        Call VerSiClienteActualizoAutorizoSusDatos(grdCliente.TextMatrix(nFila, 1), nOperacion)
    Next nFila
    'FIN FRHU 20151203
    ' ***** Agregado por RIRO el 20130501 *****
    If bProcesoNuevo = True And InStr(1, "210201,210202,210206,210209,210301,210302", nOperacion) > 0 Then 'RIRO20140910 Solo operaciones re retiro y cancelacion.
        If Not validarReglasPersonas Then
            MsgBox "Las personas seleccionadas no tienen suficientes poderes para realizar la operación", vbInformation 'JUEZ 20131212
            Exit Sub
        End If
        'Validando si es mayor de edad
        Dim oPersonaTemp As COMNPersona.NCOMPersona
        Dim iTemp, nMenorEdad As Integer
        Set oPersonaTemp = New COMNPersona.NCOMPersona
        For iTemp = 1 To grdCliente.Rows - 1
            If grdCliente.TextMatrix(iTemp, 7) <> "PJ" Then 'JUEZ 20150204
                If oPersonaTemp.validarPersonaMayorEdad(grdCliente.TextMatrix(iTemp, 1), Format(gdFecSis, "dd/mm/yyyy")) = False Then
                    nMenorEdad = nMenorEdad + 1
                End If
            End If
        Next
        If nMenorEdad > 0 Then
            If MsgBox("Uno de los intervinientes en la cuenta es menor de edad, SOLO podrá disponer de los fondos con autorización del Juez " & vbNewLine & "Desea continuar?", vbInformation + vbYesNo, "AVISO") = vbYes Then
                Dim loVistoElectronico As frmVistoElectronico
                Dim lbResultadoVisto As Boolean
                Set loVistoElectronico = New frmVistoElectronico
                lbResultadoVisto = False
                lbResultadoVisto = loVistoElectronico.Inicio(3, nOperacion)
                If Not lbResultadoVisto Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End If
    ' ***** Fin RIRO *****

    'Mody By Gitu 2010-06-08 Para que permita cancelar la cuentas con monto cero
    If nMonto = 0 Then
        '***Agregado por ELRO el 20120824, según OYP-RFC024-2012
        'If MsgBox("El Monto de la cancelacion es igual a cero ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        '    If txtMonto.Enabled Then txtMonto.SetFocus
        '    Exit Sub
        'End If
        If nOperacion = gPFAumCapTrans Then
            MsgBox "El monto de Aumento es 0.", vbInformation, "Aviso"
            cmdTranfer.SetFocus
            Exit Sub
        Else
            If MsgBox("El Monto de la cancelacion es igual a cero ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                If txtMonto.Enabled Then txtMonto.SetFocus
                Exit Sub
            End If
        End If
        '***Fin Agregado por ELRO el 20120824*******************
    End If
    'End Gitu

    'Validación del Monto
    Select Case nOperacion
        Case gPFAumCapEfec, gPFAumCapchq, gPFAumCapTasaPactEfec, gPFAumCapTasaPactChq, gPFAumCapTrans, gPFAumCapTasaPactTrans, gPFAumCapCargoCta 'JUEZ 20131218 gPFAumCapCargoCta
            
            If nMonto < nMontoMinAumentoCap Then
                'MsgBox "Monto es menor que Monto Mínimo de Apertura", vbInformation, "Aviso"
                MsgBox "Monto es menor que Monto Mínimo de Apertura o de Aumento de Capital", vbInformation, "Aviso" 'JUEZ 20141114
                '***Agregado por ELRO el 20120824, según OYP-RFC024-2012
                If nOperacion = gPFAumCapTrans Then
                    Exit Sub
                End If
                '***Fin Agregado por ELRO el 20120824*******************
                txtMonto.value = nMontoMinAumentoCap
                If nOperacion <> gPFAumCapTrans And nOperacion <> gPFAumCapTasaPactTrans Then
                    txtMonto.SetFocus
                End If
                Exit Sub
            End If

        Case gPFDismCapEfec
            Dim nTotalDisp As Double
            nTotalDisp = CDbl(lblCapital.Caption) + CDbl(lblIntGanado.Caption)
            If nTotalDisp - nMonto < nMontoMinAumentoCap Then
                MsgBox "El Saldo Final es menor que Monto Mínimo de Apertura", vbInformation, "Aviso"
                txtMonto.value = nTotalDisp - nMontoMinAumentoCap
                txtMonto.SetFocus
                Exit Sub
            End If
        
        'RIRO20131212 ERS137
        Case gPFRetIntAboCtaBanco, gPFCancTransf
            If Round(lblTotal.Caption, 2) > Round(nMontoRetiro, 2) Then
                MsgBox "Monto es mayor que el disponible para retiro", vbInformation, "Aviso"
                If txtMonto.Enabled And fraMonto.Enabled Then txtMonto.SetFocus
                Exit Sub
            End If
        'END RIRO
        
        Case Else
            If Round(nMonto, 2) > Round(nMontoRetiro, 2) Then
                MsgBox "Monto es mayor que el disponible para retiro", vbInformation, "Aviso"
                txtMonto.Text = Format$(nMontoRetiro, "#,##0.00")
                txtMonto.SetFocus
                Exit Sub
            End If
    End Select

    If nOperacion = gPFRetIntAboAho Then
        sCtaAboInt = Trim(txtCtaAhoAboInt)
        If Len(sCtaAboInt) <> 18 Then
            MsgBox "Cuenta de Abono NO Válida.", vbInformation
            If txtCtaAhoAboInt.Enabled = True Then
                txtCtaAhoAboInt.SetFocus
            End If
            Exit Sub
        End If
    End If

    For nI = 1 To Me.grdCliente.Rows - 1
        If grdCliente.TextMatrix(nI, 1) = gsCodPersUser Then
            MsgBox "Ud no puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
            Exit Sub
        End If
    Next nI

    '-- AUTORIZACION AVMM -- 13/04/2006 -----------------------------
    If nOperacion <> gPFAumCapchq And nOperacion <> gPFAumCapEfec And _
        nOperacion <> gPFAumCapTasaPactEfec And nOperacion <> gPFAumCapTasaPactChq _
        And nOperacion <> gPFAumCapTrans And nOperacion <> gPFAumCapTasaPactTrans Then
        If VerificarAutorizacion = False Then Exit Sub
    End If
'----------------------------------------------------------------

    'Valida datos del cheque para las operaciones con Cheque
    If nOperacion = gPFAumCapchq Or nOperacion = gPFAumCapTasaPactChq Then
        If Trim(lblNroDoc.Caption) = "" Then
            MsgBox "Debe registrar un cheque válido", vbInformation, "Aviso"
            cmdCheque.SetFocus
            Exit Sub
        End If
    End If
    
    '---------------Valida si tiene creditos Pendientes de Pago AVMM -- 23-03-2007----------
    Dim sMensaCred As String
    sMensaCred = ""
    If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
        Set oCred = New COMDCredito.DCOMCredito
        If oCred.VerificarClienteCreditos(gsPersCod) Then
            sMensaCred = "...Cliente posee pagos de Creditos Pendientes"
        End If
    End If
    '---------------------------------------------------------------------------------------
    
    'JUEZ 20131218 **************************************************************
    If nOperacion = gPFAumCapCargoCta Then
        If nTipoCuenta <> fnTpoCtaCargo Then
            MsgBox "Cuenta a debitar debe tener el mismo Tipo de Cuenta del Plazo Fijo", vbInformation, "Aviso"
            Exit Sub
        End If
        If Not ValidaRelPersonasCtaCargo Then
            MsgBox "Las personas y relaciones de la cuenta a debitar deben ser las mismas que las del Plazo Fijo", vbInformation, "Aviso"
            'txtCuentaCargo.SetFocusCuenta
            LimpiaControlesCargoCta
            Exit Sub
        End If
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, nMonto) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "Aviso"
            Set clsCap = Nothing
            Exit Sub
        End If
        Set clsCap = Nothing
        
        If VerificarAutorizacion = False Then Exit Sub
    End If
    
    'Verifica actualización Persona
    If nOperacion = gPFAumCapCargoCta Then
        Dim i As Integer
        Dim lsDireccionActualizada As String
        For i = 1 To grdCliente.Rows - 1
            Dim oPersona As New COMNPersona.NCOMPersona
          
            If oPersona.NecesitaActualizarDatos(grdCliente.TextMatrix(i, 1), gdFecSis) Then
                 MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & Trim(Left(grdCliente.TextMatrix(i, 3), 50)) & ": " & grdCliente.TextMatrix(i, 2), vbInformation, "Aviso"
                 Dim foPersona As New frmPersona
                 If Not foPersona.realizarMantenimiento(grdCliente.TextMatrix(i, 1), lsDireccionActualizada) Then
                     MsgBox "No se ha realizado la actualización de los datos de " & grdCliente.TextMatrix(i, 2) & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
                     Exit Sub
                 End If
                 If Trim(lsDireccionActualizada) <> "" Then
                    grdCliente.TextMatrix(i, 8) = lsDireccionActualizada
                 End If
            End If
            lsDireccionActualizada = ""
        Next
    End If
    'END JUEZ *******************************************************************
    
    'WIOR 20121009**********************************************************
    If nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Or nOperacion = gPFRetIntAdelantado Or nOperacion = gPFCancEfec Or nOperacion = gPFAumCapEfec Or nOperacion = gPFAumCapCargoCta Then 'JUEZ 20131218 gPFAumCapCargoCta
        Dim oDPersona As COMDPersona.DCOMPersona
        Dim rsPersona As ADODB.Recordset
        Dim sCodPersona As String
        Dim Cont As Integer
        
        Set oDPersona = New COMDPersona.DCOMPersona
        
        For Cont = 0 To grdCliente.Rows - 2
            If Trim(Right(grdCliente.TextMatrix(Cont + 1, 3), 5)) = gCapRelPersTitular Then
                sCodPersona = Trim(grdCliente.TextMatrix(Cont + 1, 1))
                Set rsPersona = oDPersona.ObtenerUltimaVisita(sCodPersona)
                If rsPersona.RecordCount > 0 Then
                    If Not (rsPersona.EOF And rsPersona.BOF) Then
                        If Trim(rsPersona!sUsual) = "3" Then
                            MsgBox Trim(grdCliente.TextMatrix(Cont + 1, 2)) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                            Call frmPersona.Inicio(sCodPersona, PersonaActualiza)
                        End If
                    End If
                End If
                Set rsPersona = Nothing
            End If
        Next Cont
    End If
    'WIOR FIN ***************************************************************
    
        
    'AMDO 20130702 TI-ERS063-2013 ****************************************************
        If nOperacion = gPFRetInt Or nOperacion = gPFAumCapEfec Then
            Dim oDPersonaAct As COMDPersona.DCOMPersona
            Dim conta As Integer
            Dim sPersCod As String
            
            Set oDPersonaAct = New COMDPersona.DCOMPersona
            For conta = 0 To grdCliente.Rows - 2
            sPersCod = Trim(grdCliente.TextMatrix(conta + 1, 1))
                            If oDPersonaAct.VerificaExisteSolicitudDatos(sPersCod) Then
                                MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & grdCliente.TextMatrix(conta + 1, 2)) & "." & Chr(10), vbInformation, "Aviso"
                                Call frmActInfContacto.Inicio(sPersCod)
                            End If
            Next conta
        End If
    'AMDO FIN ********************************************************************************
    
    'JUEZ 20131212 *******************************************************************
    If nOperacion = gPFAumCapCargoCta Then
        If nTipoCuenta = gPrdCtaTpoIndist Or nTipoCuenta = gPrdCtaTpoMancom Then
            If Not frmCapConfirmPoderes.Inicia(txtCuentaCargo.NroCuenta, gCapAhorros, nOperacion, "Débito Aumento Capital") Then
                Exit Sub
            End If
        End If
    End If
    'END JUEZ ************************************************************************
    
    ' RIRO20131212 ERS137
    If Len(Trim(txtTitular.Text)) = 0 And txtTitular.Visible Then
        MsgBox "Debe Ingresar el nombre del titular de la cuenta destino", vbInformation, "Aviso"
        Exit Sub
    End If
    ' END RIRO
    'EJVG20140210 ***
    If nOperacion = gPFAumCapchq Or nOperacion = gPFAumCapTasaPactChq Then
        If Not ValidaSeleccionCheque Then
            MsgBox "Ud. debe seleccionar un Cheque para continuar", vbInformation, "Aviso"
            If cmdCheque.Visible And cmdCheque.Enabled Then cmdCheque.SetFocus
            Exit Sub
        End If
    End If
    'END EJVG *******
    If MsgBox("¿Está seguro de grabar la información? " & sMensaCred, vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Dim sMovNro As String, sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        Dim sMovNroCom As String
        Dim clsMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
        Dim nSaldo As Double
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMontoLavDinero As Double, nTC As Double
        Dim lsBoleta As String, lsBoletaITF As String, lsBoletaPremio As String
        Dim nFicSal As Integer
        Dim lsBoletaCVME As String '***Agregado por ELRO el 20120814, según OYP-RFC024-2012
        Dim oNCOMContImprimir As COMNContabilidad.NCOMContImprimir '***Agregado por ELRO el 20120717, según OYP-RFC024-2012
        Set oNCOMContImprimir = New COMNContabilidad.NCOMContImprimir '***Agregado por ELRO el 20120717, según OYP-RFC024-2012

        'Realiza la Validación para el Lavado de Dinero
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        If clsLav.EsOperacionEfectivo(nOperacion) Then
            Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
            If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
                Set clsExo = Nothing
                sPersLavDinero = ""
                nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
                Set clsLav = Nothing
                If nmoneda = gMonedaNacional Then
                    Dim clsTC As COMDConstSistema.NCOMTipoCambio 'Tipo Cambio
                    Set clsTC = New COMDConstSistema.NCOMTipoCambio
                    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set clsTC = Nothing
                Else
                    nTC = 1
                End If
                If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                 
                'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    'ALPA 20081009****************************************************
                    'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "", , , , , nMoneda)
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End

                
                End If
            Else
                Set clsExo = Nothing
            End If
        Else
            Set clsLav = Nothing
        End If
    'WIOR 20130301 **SEGUN TI-ERS005-2013*************************************
    fbPersonaReaAhorros = False
    If (loLavDinero.OrdPersLavDinero = "" Or loLavDinero.OrdPersLavDinero = "Exit") _
            And (nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Or nOperacion = gPFRetIntAdelantado _
            Or nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf _
            Or nOperacion = gPFAumCapEfec Or nOperacion = gPFAumCapchq Or nOperacion = gPFAumCapTrans _
            Or nOperacion = gPFAumCapTasaPactTrans Or nOperacion = gPFAumCapTasaPactEfec Or nOperacion = gPFAumCapTasaPactChq) Then
            
            Dim oPersonaSPR As UPersona_Cli
            Dim oPersonaU As COMDPersona.UCOMPersona
            Dim nTipoConBN As Integer
            Dim sConPersona As String
            Dim pbClienteReforzado As Boolean
            Dim rsAgeParam As Recordset
            Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
            Dim lnMontoX As Double, lnTC As Double
            Dim ObjTc As COMDConstSistema.NCOMTipoCambio
            
            
            Set oPersonaU = New COMDPersona.UCOMPersona
            Set oPersonaSPR = New UPersona_Cli
            
            fbPersonaReaAhorros = False
            pbClienteReforzado = False
            fnCondicion = 0
            
            For nI = 0 To grdCliente.Rows - 2
                oPersonaSPR.RecuperaPersona Trim(grdCliente.TextMatrix(nI + 1, 1))
                                    
                If oPersonaSPR.Personeria = 1 Then
                    If oPersonaSPR.Nacionalidad <> "04028" Then
                        sConPersona = "Extranjera"
                        fnCondicion = 1
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.Residencia <> 1 Then
                        sConPersona = "No Residente"
                        fnCondicion = 2
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaSPR.RPeps = 1 Then
                        sConPersona = "PEPS"
                        fnCondicion = 4
                        pbClienteReforzado = True
                        Exit For
                    ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                Else
                    If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                        If nTipoConBN = 1 Or nTipoConBN = 3 Then
                            sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                            fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                            pbClienteReforzado = True
                            Exit For
                        End If
                    End If
                End If
            Next nI
            
            If pbClienteReforzado Then
                MsgBox "El Cliente: " & Trim(grdCliente.TextMatrix(nI + 1, 2)) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
                frmPersRealizaOpeGeneral.Inicia sOperacion & " (Persona " & sConPersona & ")", nOperacion
                fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                
                If Not fbPersonaReaAhorros Then
                    MsgBox "Se va a proceder a Anular la Operacion del Plazo Fijo", vbInformation, "Aviso"
                    cmdGrabar.Enabled = True
                    Exit Sub
                End If
            Else
                fnCondicion = 0
                lnMontoX = nMonto
                pbClienteReforzado = False
                
                Set ObjTc = New COMDConstSistema.NCOMTipoCambio
                lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set ObjTc = Nothing
            
            
                Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
                Set rsAgeParam = objCap.getCapAbonoAgeParam(gsCodAge)
                Set objCap = Nothing
                
                If Mid(Trim(TxtCuenta.NroCuenta), 9, 1) = 1 Then
                    lnMontoX = Round(lnMontoX / lnTC, 2)
                End If
            
                If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                    If lnMontoX >= rsAgeParam!nMontoMin And lnMontoX <= rsAgeParam!nMontoMax Then
                        frmPersRealizaOpeGeneral.Inicia sOperacion, nOperacion
                        fbPersonaReaAhorros = frmPersRealizaOpeGeneral.PersRegistrar
                        If Not fbPersonaReaAhorros Then
                            MsgBox "Se va a proceder a Anular la Operacion del Plazo Fijo", vbInformation, "Aviso"
                            cmdGrabar.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
    End If
    'WIOR FIN ***************************************************************
    If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140212
        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Sleep (1000) 'RIRO20131212 ERS137
        sMovNroCom = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) 'RIRO20131212 ERS137
        Set clsMov = Nothing
        On Error GoTo ErrGraba
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        clsCap.IniciaImpresora gImpresora
    
        'ANDE 20180419 ERS021-2018 participanción en campaña mundialto
        Dim nCondicion As Integer, nPuntosRef As Integer, nPTotalAcumulado As Integer
        'If nOperacion = gPFCancEfec Then
        If nOperacion = gPFCancEfec Or nOperacion = gPFAumCapEfec Then 'APRI20191128
            Dim nOpeTipo As Integer
'            If nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho Or nOperacion = gPFRetIntAdelantado Or nOperacion = gPFRetIntAboCtaBanco Then
'                nOpeTipo = 3 '3:Retiro
            If nOperacion = gPFAumCapEfec Then 'APRI20191128
                nOpeTipo = 2 '2:Abono
            Else
                nOpeTipo = 4 '4: Cancelación
            End If
            nmoneda = Mid(sCuenta, 9, 1)
            If nmoneda = gMonedaNacional Then
                'Dim oCampanhaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
                Call oCaptaLN.VerificarParticipacionCampMundial(cperscod, sCuenta, nOperacion, nOpeTipo, nmoneda, nMonto, nTipoPersona, bParticipaCamp, sMovNro, , gdFecSis, lnTpoPrograma, TxtCuenta.Age, nPuntosRef, nCondicion, , nPTotalAcumulado)
                If bParticipaCamp Then
                    cTextoDatos = "#" & IIf(bParticipaCamp, "1", "0") & "." & CStr(nPuntosRef) & "$" & CStr(nCondicion) & "_" & CStr(nPTotalAcumulado) & "&"
                End If
            End If
        End If
    
        Select Case nOperacion
            'RIRO20131220 ERS137 - Se Agregó gPFRetIntAboCtaBanco
            Case gPFRetInt, gPFRetIntAdelantado, gPFRetIntAboCtaBanco
                '***Agregado por ELRO el 20130723, según TI-ERS079-2013****
                'nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(val(lblDias.Caption)), Trim(txtGlosa.Text), nOperacion, , , , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                If nOperacion = gPFRetInt Then
                    nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), Trim(txtGlosa.Text), nOperacion, , , , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , , , cTextoDatos) 'ande ers021-2018 agregué cTextoDatos
                
                ' *** RIRO20131212 ERS137
                ElseIf nOperacion = gPFRetIntAboCtaBanco Then
                    nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), getGlosa, nOperacion, , , , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, CDbl(lblComision.Caption), Mid(txtBancoDestino.Text, 4, 13), sMovNroCom, Trim(txtCuentaDestino.Text), getTitular, chkComision.value, cTextoDatos) 'ande ers021-2018 agregué cTextoDatos
                ' *** END RIRO
                
                Else
                    'nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), Trim(txtGlosa.Text), nOperacion, , , , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, ObtenerRegla, , , , , ,  , cTextoDatos) 'ande ers021-2018 agregué cTextoDatos
                    nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), Trim(txtGlosa.Text), nOperacion, , , , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, , ObtenerRegla, , , , , , , cTextoDatos) 'ANDE GIPO APRI INC1805280003 20180529
                End If
                '***Fin Agregado por ELRO el 20130723, según TI-ERS079-2013
                 'ALPA 20081010
                 If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
                     Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                 End If
            Case gPFCancEfec
                'By Capi 19082008 se adiciono valor opcion lbtasapactada
                '***Modificado por ELRO el 20120327, según OYP-RFC023-2012
                'nSaldo = clsCap.CapCancelacionPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosa.Text), CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , , , , lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lbTasaPactada, gnMovNro)
                nSaldo = clsCap.CapCancelacionPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosa.Text), CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , , , , lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lbTasaPactada, gnMovNro, ckbPorAfectacion, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , cTextoDatos) 'ande ers021-2018 agregué cTextoDatos
                '***Fin Modificado por ELRO*******************************
                Set oCred = Nothing
                 'ALPA 20081010
                If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
                     Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If
            Case gPFCancTransf
                
                Set oCred = New COMDCredito.DCOMCredito
                'By Capi 19082008 se adiciono valor opcion lbtasapactada
                'nSaldo = clsCap.CapCancelacionPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosa.Text), CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, True, Right(Me.txtBanco.Text, 13), Me.txtCtaBanco.Text, CInt(Right(Me.cboMonedaBanco.Text, 3)), , lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lbTasaPactada, gnMovNro, , , ObtenerRegla) RIRO20131212 ERS137 - Comentado
                nSaldo = clsCap.CapCancelacionPF(sCuenta, sMovNro, nOperacion, Trim(getGlosa), CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, True, Right(Me.txtBancoDestino.Text, 13), Me.txtCuentaDestino.Text, , , lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, lbTasaPactada, gnMovNro, , , ObtenerRegla, lblComision, sMovNroCom, getTitular, chkComision.value, , Trim(lblBancoDestino.Caption)) 'RIRO20131212 ERS137 - Comentado
                Set oCred = Nothing
                'ALPA 20081010
                If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
                     Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If
            Case gPFRetIntAboAho
                '***Parametro CInt(Trim(Right(cboMedioRetiro, 3))) agregado por ELRO el 20130724, según TI-ERS079-2013
                If Mid(sCuenta, 9, 1) <> Mid(sCtaAboInt, 9, 1) Then
                    If CLng(Mid(sCtaAboInt, 9, 1)) = gMonedaExtranjera Then
                        nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), Trim(txtGlosa.Text), gPFRetIntAboAho, sCtaAboInt, nTCV, , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla)
                    Else
                        nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), Trim(txtGlosa.Text), gPFRetIntAboAho, sCtaAboInt, nTCC, , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , , , cTextoDatos) 'ande ers021-2018 agregué cTextoDatos
                    End If
                Else
                    nSaldo = clsCap.CapRetiroInteresPF(sCuenta, sMovNro, nMonto, CLng(Val(lblDias.Caption)), Trim(txtGlosa.Text), gPFRetIntAboAho, sCtaAboInt, , , , gsNomAge, sLpt, sPersLavDinero, gsCodCMAC, gbITFAplica, Me.lblITF.Caption, gbITFAsumidoPF, gITFCobroEfectivo, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, CInt(Trim(Right(cboMedioRetiro, 3))), ObtenerRegla, , , , , , , cTextoDatos) 'ande ers021-2018 agregué cTextoDatos
                End If
              'ALPA 20081010
              If gnMovNro > 0 Then
                'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
                 Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
              End If
            Case gPFAumCapEfec, gPFAumCapTasaPactEfec
                'ValidaTasaInteres
                'nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, , , , gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro)
                nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, , , , gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro, , , , , , , , , , , , cTextoDatos) 'APRI20191128
                'ALPA 20081010
                If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sPersLavDinero, , , , sPersLavDinero, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                     Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If

            Case gPFDismCapEfec
                ValidaTasaInteres
                nMonto = CDbl(lblTotal.Caption)
                nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, , , , gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro)
                'ALPA 20081010
                If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sPersLavDinero, , , , sPersLavDinero, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If

            Case gPFAumCapchq, gPFAumCapTasaPactChq
                ValidaTasaInteres
                'nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, Trim(lblNroDoc.Caption), sCodIF, sIFCuenta, gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro)
                nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFCta, gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro, , , , , , , , , , oDocRec.fnTpoDoc, oDocRec.fsIFTpo) 'EJVG20140210
                'ALPA 20081010
                If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sPersLavDinero, , , , sPersLavDinero, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                     Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If
            Case gPFAumCapTrans, gPFAumCapTasaPactTrans
                'ValidaTasaInteres '***Comentado por ELRO el 20130304, SATI INC1303050011
                '***Modificado por ELRO el 20120824, según OYP-RFC024-2012
                'fnSaldo = nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, , , , gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF , gITFCobroEfectivo , sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro, lnMovNroTransfer, Right(cboTransferMoneda, 3), fnMovNroRVD, lnTransferSaldo)
                nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, , , , gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, IIf(chkITFEfectivo.value = 1, gITFCobroEfectivo, gITFCobroCargoPF), sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro, lnMovNroTransfer, Right(cboTransferMoneda, 3), fnMovNroRVD, lnTransferSaldo)
                '***Fin Modificado por ELRO el 20120824*******************
                'ALPA 20081010
                If gnMovNro > 0 Then
                    'Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, sPersLavDinero, , , , sPersLavDinero, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If
            Case gPFAumCapCargoCta 'JUEZ 20131218
                nSaldo = clsCap.CapAumDismCapPF(sCuenta, sMovNro, nOperacion, Trim(txtGlosaAumDism.Text), nMonto, CDbl(lblCapital.Caption), CDbl(lblIntGanado.Caption), CLng(txtPlazo.Text), nTasaNominal, , , , gsNomAge, sLpt, gsCodCMAC, sPersLavDinero, gbITFAplica, CDbl(lblITF.Caption), gbITFAsumidoPF, gITFCobroEfectivo, sPersLavDinero, , , lsBoleta, lsBoletaITF, gbImpTMU, gnMovNro, , , , , txtCuentaCargo.NroCuenta, ObtenerRegla, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.VisPersLavDinero)
                If gnMovNro > 0 Then
                     Call loLavDinero.InsertarLavDinero(sPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
                End If
        End Select
        '*****BRGO 20110914 ***********
        If gbITFAplica = True And CCur(lblITF.Caption) > 0 Then
            Call oMov.InsertaMovRedondeoITF(sMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
        End If
        Set oMov = Nothing
        '*** End BRGO *****************
    'CAAU Cambio para Maynas
    If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
        If CDate(Me.lblVencimiento) > CDate(gdFecSis) Then
            nMontoPremio = 0
            nCantPremio = 0
            lsBoletaPremioCanc = ""
            Dim lnMonto As Currency
            Dim lnMovNro As Long
            Dim lsMov As String
            Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
            Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
            Dim RsPremios As ADODB.Recordset
            Set RsPremios = New ADODB.Recordset
            Set RsPremios = clsCapMov.Get_PremioPF(sCuenta)
            If Not (RsPremios.BOF And RsPremios.EOF) Then
'                Call FrmCapOpeCancPremioPF.Inicion(sCuenta, RsPremios, lnMonto)
'                Set FrmCapOpeCancPremioPF = Nothing
'                Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
'                sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                Set clsMov = Nothing
'
'                lsMov = sMovNro
'                lnMovNro = clsCapMov.OtrasOperaciones(lsMov, gOtrOpePremioCancPF, lnMonto, "", "DSCTO X CANC DE PF PF", nmoneda, grdCliente.TextMatrix(1, 1))
'                Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
'                Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
'                    lsBoletaPremio = oBol.ImprimeBoleta("OTRAS OPERACIONES", "DCTO X PREMIO", gOtrOpePremioCancPF, str(lnMonto), Me.grdCliente.TextMatrix(1, 2), txtCuenta.NroCuenta, "", CDbl(lnMonto), "", "", 0, 0, 0, 0, , , , False _
'                    , , , , gdFecSis, , gsCodUser, sLpt, False, False, 0, False)
'                Set oBol = Nothing
'
'                Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
'                Call clsCapMov.CapCancelacionPremioPF(gsNomAge, gsCodUser, sMovNro, FechaHora(gdFecSis), nSaldo, sCuenta, lsBoletaPremioCanc)
            End If
        End If
    End If
    'APRI20190109 ERS077-2018
    If nOperacion = gPFRetInt Or nOperacion = gPFCancEfec Then
        Dim cDCap As New COMDCaptaGenerales.DCOMCaptaGenerales
        Dim rs As ADODB.Recordset
        Dim bComunica As Boolean
        Dim cFechaApli As String
        nI = 0
        For nI = 1 To grdCliente.Rows - 1
            Set rs = cDCap.AplicaComunicacionCaptaciones(grdCliente.TextMatrix(nI, 1), "")
            bComunica = rs!bComun
            If Not bComunica Then
               Exit For
            End If
        Next nI
         
        If Not bComunica Then
            Set rs = cDCap.ObtenerFechaAplicacionTarifario
            cFechaApli = rs!dFecApli
            MsgBox "En Cumplimiento al Reglamento de Gestión de Conducta de Mercado del Sistema Financiero Res. SBS N° 3274-2017 y sus modificatorias, tenemos el agrado de comunicarle que a partir del " & cFechaApli & " entró en vigencia las nuevas condiciones de nuestros productos pasivos.", vbInformation, "COMUNICACIÓN POR CAMBIOS CONTRACTUALES"
                    
            ImpreCartaNotificacionTarifario "", sCuenta, gdFecSis
            
        End If
        rs.Close
        Set rs = Nothing
        Set cDCap = Nothing
    End If
    'END APRI
    Set clsCap = Nothing
    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
    If Trim(lsBoletaPremio) <> "" Then ImprimeBoleta lsBoletaPremio, "Boleta Premio"
        
    Dim nTasa As Double
    nTasa = CDbl(lblTasa)
    
    If nOperacion = gPFAumCapEfec Or nOperacion = gPFAumCapTasaPactEfec Or nOperacion = gPFDismCapEfec Or _
       nOperacion = gPFAumCapchq Or nOperacion = gPFAumCapTasaPactChq _
       Or nOperacion = gPFAumCapTrans Or nOperacion = gPFAumCapTasaPactTrans Then
       nMonto = CDbl(lblSaldoFinal)
       '*** Comentado por BRGO 20111226
       'EmiteCalendarioRetiroIntPFMensual nMonto, nTasa, CInt(txtPlazo.Text), gdFecSis, nMoneda, lnDValoriza
       '***Agregado por ELRO el 20120814, según OYP-RFC24-2012
       If nOperacion = gPFAumCapTrans Then
        If Mid(sCuenta, 9, 1) <> Trim(Right(cboTransferMoneda, 3)) Then
          MsgBox "Coloque papel para la Boleta de Compra/Venta Moneda Extranjera.", vbInformation, "Aviso"
          lsBoletaCVME = oNCOMContImprimir.ImprimeBoletaCompraVentaME("Compra/Venta Moneda Extranjera", "", _
                                                                      fsPersNombreCVME, _
                                                                      fsPersDireccionCVME, _
                                                                      fsdocumentoCVME, _
                                                                      IIf(Trim(Right(cboTransferMoneda, 3)) = Moneda.gMonedaExtranjera, CCur(lblTTCCD), CCur(lblTTCVD)), _
                                                                      IIf(Trim(Right(cboTransferMoneda, 3)) = Moneda.gMonedaExtranjera, gOpeCajeroMECompra, gOpeCajeroMEVenta), _
                                                                      CCur(lblMonTra), _
                                                                      CCur(txtMonto), _
                                                                      gsNomAge, _
                                                                      sMovNro, _
                                                                      sLpt, _
                                                                      gsCodCMAC, _
                                                                      gsNomCmac, _
                                                                      gbImpTMU)
              Do
               If Trim(lsBoletaCVME) <> "" Then
                  nFicSal = FreeFile
                  Open sLpt For Output As nFicSal
                     Print #nFicSal, lsBoletaCVME
                     Print #nFicSal, ""
                  Close #nFicSal
                End If
                
            Loop Until MsgBox("¿Desea reimprimir Boleta de Compra/Venta Moneda Extranjera? ", vbQuestion + vbYesNo, Me.Caption) = vbNo
        End If
        fsPersNombreCVME = ""
        fsPersDireccionCVME = ""
        fsdocumentoCVME = ""
        lsBoletaCVME = ""
       End If
       '***Fin Agregado por ELRO el 20120814

    End If
    Set loLavDinero = Nothing
    'WIOR 20130301 comento
'    'WIOR 20121114 ************************************
'    If fnDepositoPersRealiza Then
'        frmPersRealizaOperacion.InsertaPersonaRealizaOperacion gnMovNro, sCuenta, frmPersRealizaOperacion.PersTipoCliente, _
'        frmPersRealizaOperacion.PersCod, frmPersRealizaOperacion.PersTipoDOI, frmPersRealizaOperacion.PersDOI, frmPersRealizaOperacion.PersNombre, _
'        frmPersRealizaOperacion.TipoOperacion, frmPersRealizaOperacion.Origen, fnCondicion
'
'        fnDepositoPersRealiza = False
'        fnCondicion = 0
'    End If
'    'WIOR FIN *****************************************
    'WIOR 20130301 ************************************************************
    If fbPersonaReaAhorros And gnMovNro > 0 Then
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(sCuenta), fnCondicion
        fbPersonaReaAhorros = False
    End If
    'WIOR FIN *****************************************************************

    '***Agregado por ELRO el 20120718, según OYP-RFC024-2012
    fnMovNroRVD = 0
    lblMonTra = "0.00"
    Set oNCOMContImprimir = Nothing
    '***Fin Agregado por ELRO el 20120718*******************
    'INICIO JHCU ENCUESTA 16-10-2019
    Dim nOpeEnc As String
    nOpeEnc = nOperacion
    Encuestas gsCodUser, gsCodAge, "ERS0292019", nOpeEnc
    'FIN
    cmdCancelar_Click
End If
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
    Exit Sub
End Sub


Private Sub cmdsalir_Click()
Unload Me
End Sub
'EJVG20130916 ***
'Private Sub cmdTranfer_Click()
'Dim lsGlosa As String
'Dim lsDoc As String
'Dim lsInstit As String
'
'Dim oform As New frmCapRegVouDepBus '***Agregado por ELRO el 20120810, según OYP-RFC024-2012
'Dim lnTipMot As Integer '***Agregado por ELRO el 20120810, según OYP-RFC024-2012
'Dim i As Integer '***Agregado por ELRO el 20120810, según OYP-RFC024-2012
'
'
'    If Me.cboTransferMoneda.Text = "" Then
'        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
'        cboTransferMoneda.SetFocus
'        Exit Sub
'    End If
'
'    '***Agregado por ELRO el 20120810, según OYP-RFC024-2012
'    If gsOpeCod = gPFAumCapTrans Then
'        lnTipMot = 4
'    End If
'    '***Fin Agregado por ELRO*******************************
'
'    '***Modificado por ELRO el 20120810, según OYP-RFC024-2012
'    'lnMovNroTransfer = frmTransfpendientes.Ini(Right(Me.cboTransferMoneda.Text, 2), lnTransferSaldo, lsGlosa, lsInstit, lsDoc)
'    oform.iniciarFormularioDeposito Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, fsPersNombreCVME, fsPersDireccionCVME, fsdocumentoCVME
'    '***Fin Modificado por ELRO el 20120706*******************
'
'    '***Comentado por ELRO el 20120810, según OYP-RFC024-2012
'    'If lnMovNroTransfer = -1 Then
'    '    Me.cboTransferMoneda.Enabled = True
'    '    lnTransferSaldo = 0
'    'Else
'    '    Me.cboTransferMoneda.Enabled = False
'    'End If
'    '***Fin Comentado por ELRO*******************************
'
'    Me.txtGlosaAumDism.Text = lsGlosa
'    Me.lbltransferBco.Caption = lsInstit
'    Me.lblTrasferND.Caption = lsDoc
'    'sNroDoc = lsDoc
'
'    Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
'
'    If Mid(txtCuenta.NroCuenta, 9, 1) = Moneda.gMonedaNacional Then
'        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
'            Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
'        Else
'            Me.txtMonto.Text = Format(lnTransferSaldo * CCur(Me.lblTTCCD.Caption), "#,##0.00")
'        End If
'    Else
'        If Right(cboTransferMoneda, 3) = Moneda.gMonedaNacional Then
'            Me.txtMonto.Text = Format(lnTransferSaldo / CCur(Me.lblTTCVD.Caption), "#,##0.00")
'        Else
'            Me.txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
'        End If
'    End If
'
'    '***Agregado por ELRO el 20120810, según OYP-RFC024-2012
'    If lnTransferSaldo > 0# Then
'        cboTransferMoneda.Enabled = False
'    Else
'        cboTransferMoneda.Enabled = True
'    End If
'    txtGlosaAumDism.Locked = True
'    txtMonto.Enabled = False
'    lblMonTra = Format(lnTransferSaldo, "#,##0.00")
'    '***Fin Agregado por ELRO el 20120810*******************
'
'
''    If txtCuenta.Prod = "234" Then
''        vnMontoDOC = CDbl(txtMonto.Text)
''        lblTotTran.Caption = vnMontoDOC
''    End If
'
'   ' Me.LblTotal.Caption = Format(txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00")
'
''    If lnMovNroTransfer <> -1 Then
''        Me.txtGlosaAumDism.SetFocus
''    End If
'
'
'End Sub
Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oForm As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim i As Integer

    If Len(TxtCuenta.NroCuenta) <> 18 Then
        MsgBox "Ud. debe especificar el Nro. de Cuenta", vbInformation, "Aviso"
        If TxtCuenta.Visible And TxtCuenta.Enabled Then TxtCuenta.SetFocusCuenta
        Exit Sub
    End If
    If cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        cboTransferMoneda.SetFocus
        Exit Sub
    End If
    
    If gsOpeCod = gPFAumCapTrans Then
        lnTipMot = 4
    End If
    
    'txtMonto.value = "0.00"
    Set oForm = New frmCapRegVouDepBus
    oForm.iniciarFormularioDeposito Trim(Right(cboTransferMoneda.Text, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, fsPersNombreCVME, fsPersDireccionCVME, fsdocumentoCVME, TxtCuenta.NroCuenta

    txtGlosaAumDism.Text = lsGlosa
    lbltransferBco.Caption = lsInstit
    lblTrasferND.Caption = lsDoc
    
    txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
    
    If Mid(TxtCuenta.NroCuenta, 9, 1) = Moneda.gMonedaNacional Then
        If Right(cboTransferMoneda.Text, 3) = Moneda.gMonedaNacional Then
            txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
        Else
            txtMonto.Text = Format(lnTransferSaldo * CCur(lblTTCCD.Caption), "#,##0.00")
        End If
    Else
        If Right(cboTransferMoneda.Text, 3) = Moneda.gMonedaNacional Then
            txtMonto.Text = Format(lnTransferSaldo / CCur(lblTTCVD.Caption), "#,##0.00")
        Else
            txtMonto.Text = Format(lnTransferSaldo, "#,##0.00")
        End If
    End If

    txtGlosaAumDism.Locked = True
    txtMonto.Enabled = False
    lblMonTra = Format(lnTransferSaldo, "#,##0.00")
    fraMonto.Enabled = True
    Set oForm = Nothing
End Sub

Private Sub cmdVerRegla_Click()
    If strReglas <> "" Then
        Call frmCapVerReglas.Inicia(strReglas)
    Else
        MsgBox "Cuenta no tiene reglas definidas", vbInformation
    End If
End Sub

'END EJVG *******
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral
Set clsGen = New COMDConstSistema.DCOMGeneral
fraDocumento.Enabled = True
    If KeyCode = vbKeyF12 And TxtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        Dim bRetSinTarjeta As Boolean
        bRetSinTarjeta = clsGen.GetPermisoEspecialUsuario(gCapPermEspRetSinTarj, gsCodUser, gsDominio)
        sCuenta = frmValTarCodAnt.Inicia(gCapPlazoFijo, bRetSinTarjeta)
        If Val(Mid(sCuenta, 6, 3)) <> Producto.gCapPlazoFijo Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            TxtCuenta.NroCuenta = sCuenta
            TxtCuenta.SetFocusCuenta
        End If
    End If
    
    
'-------------------------------------------------------------------------------------------


    If KeyCode = vbKeyF11 And TxtCuenta.Enabled = True Then 'F11
    
        
        Dim nPuerto As COMDConstantes.TipoPuertoSerial
        
        Dim sNumTar As String
        Dim sClaveTar As String
        Dim nErr As Integer
        Dim nEstado As COMDConstantes.CaptacTarjetaEstado
        Dim sMaquina As String
        sMaquina = GetComputerName

        
        Set clsGen = New COMDConstSistema.DCOMGeneral 'DGeneral
        
        nPuerto = clsGen.GetPuertoPeriferico(gPerifPINPAD, sMaquina)
                
        If nPuerto < 0 Then nPuerto = gPuertoSerialCOM1
        
        'ppoa Modificacion
        If Not IniciaPinPad(nPuerto) Then
            MsgBox "No Inicio Dispositivo" & ". Consulte con Servicio Tecnico.", vbInformation, "Aviso"
            Exit Sub
        End If
        
        'ppoa Modificacion
        If Not WriteToLcd("Pase su Tarjeta por la Lectora.") Then
            FinalizaPinPad
            MsgBox "No se Realizó Envío", vbInformation, "Aviso"
            Exit Sub
        End If
        
        sCaption = Me.Caption
        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
        
        'ppoa Modificacion
        sNumTar = GetNumTarjeta
        
        'ppoa neceita cambio el formateo ??
         'trasladado a la funcion  GetNumTarjeta
        'sNumTar = Trim(Replace(sNumTar, "-", "", 1, , vbTextCompare))
        
        If Len(sNumTar) <> 16 Then
            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
            
            'ppoa Modificacion
            FinalizaPinPad
            
            Me.Caption = sCaption
            Exit Sub
        End If
        
        Me.Caption = "Ingrese la Clave de la Tarjeta."
                        
        
        'ppoa Modificacion
        sClaveTar = GetClaveTarjeta("INGRESE CLAVE")
        
        Dim lnResult As COMDConstantes.ResultVerificacionTarjeta
        
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Select Case clsGen.ValidaTarjeta(sNumTar, sClaveTar)

            Case gClaveValida
                    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
                    Dim rsTarj As New ADODB.Recordset
                    
                    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
                    If nOperacion = gAhoRetEfec Or nOperacion = gAhoRetOP Then
                        Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, gCapAhorros)
                    ElseIf nOperacion = gCTSRetEfec Then
                        Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, gCapCTS)
                    Else
                        Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar, gCapPlazoFijo)
                    End If
                    
                                        
                    If rsTarj.EOF And rsTarj.BOF Then
                        MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
                        FinalizaPinPad
                        Me.Caption = sCaption
                        Exit Sub
                    Else
                        nEstado = rsTarj("nEstado")
                        If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
                            If nEstado = gCapTarjEstBloqueada Then
                                MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            ElseIf nEstado = gCapTarjEstCancelada Then
                                MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
                            End If
                                                        
                            FinalizaPinPad
                            Me.Caption = sCaption
                            Exit Sub
                        End If
                        Dim rsPers As New ADODB.Recordset
                        Dim sCta As String, sProducto As String, sMoneda As String
                        Dim clsCuenta As UCapCuenta
                                                
                        
                        'Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
                        Set rsPers = clsMant.GetTarjetaCuentas(sNumTar, gCapPlazoFijo)
                        
                        
                        Set clsMant = Nothing
                        If Not (rsPers.EOF And rsPers.EOF) Then
                            Do While Not rsPers.EOF
                                sCta = rsPers("cCtaCod")
                                sProducto = rsPers("Producto")
                                sMoneda = Trim(rsPers("Moneda"))
                                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & space(2) & sProducto & space(2) & sMoneda
                                rsPers.MoveNext
                            Loop
                            Set clsCuenta = New UCapCuenta
                            Set clsCuenta = frmCapMantenimientoCtas.Inicia
                            If clsCuenta.sCtaCod <> "" Then
                                TxtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                                TxtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
                                TxtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                                TxtCuenta.SetFocusCuenta
                                Call txtCuenta_KeyPress(13)
                                'SendKeys "{Enter}"
                            End If
                            Set clsCuenta = Nothing
                        Else
                            MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
                        End If
                        rsPers.Close
                        Set rsPers = Nothing
                    End If
                    Set rsTarj = Nothing
                    Set clsMant = Nothing
                
                
            Case gTarjNoRegistrada
            
                'ppoa Modificacion
                If Not WriteToLcd("Espere Por Favor") Then
                    FinalizaPinPad
                    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
                    Exit Sub
                End If
                MsgBox "Tarjeta no Registrada", vbInformation, "Aviso"
                
            Case gClaveNOValida
                'ppoa Modificacion
                If Not WriteToLcd("Clave Incorrecta") Then
                    FinalizaPinPad
                    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
                    Exit Sub
                End If
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
                
        End Select
                                                                        
        Set clsGen = Nothing
        
        FinalizaPinPad
        Me.Caption = "Captaciones -  Cargo - Ahorros " & sOperacion
        
    End If
    
End Sub

Private Sub Form_Load()
    nMontoRetiro = 0
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    lnMovNroTransfer = -1
    lnTransferSaldo = 0
    GetTipCambio gdFecSis
    
    '***Modificado por ELRO el 20120828, según OYP-RFC024-2012
    'lblTTCCD.Caption = Format(gnTipCambioC, "#.00")
    lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
    'lblTTCVD.Caption = Format(gnTipCambioV, "#.00")
    lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
    '***Fin Modificado por ELRO el 20120828*******************
    
    ' ****** Agregado Por RIRO el 20130501 *****
    grdCliente.ColWidth(1) = 1300
    grdCliente.ColWidth(2) = 3200
    grdCliente.ColWidth(3) = 1500
    grdCliente.ColWidth(7) = 800
    grdCliente.ColWidth(8) = 900
    ' ***** Fin RIRO *****

End Sub



Private Sub grdCliente_DblClick()
    Dim R As ADODB.Recordset
    Dim ssql As String
    Dim clsFirma As COMDCaptaGenerales.DCOMCaptaMovimiento 'DCapMovimientos
    Set clsFirma = New COMDCaptaGenerales.DCOMCaptaMovimiento
     
    If Me.grdCliente.TextMatrix(grdCliente.row, 1) = "" Then Exit Sub
    
    Set R = New ADODB.Recordset
    Set R = clsFirma.GetFirma(Me.grdCliente.TextMatrix(grdCliente.row, 1))
    If R.BOF Or R.EOF Then
       Set R = Nothing
       MsgBox "La visualización del DNI no esta Disponible", vbOKOnly + vbInformation, "AVISO"
       Exit Sub
    End If
            
    If R.RecordCount > 0 Then
       If IsNull(R!iPersFirma) = True Then
          MsgBox "El cliente no posee Firmas", vbInformation, "Aviso"
          Exit Sub
       End If
       frmMuestraFirma.psCodCli = Me.grdCliente.TextMatrix(grdCliente.row, 1)
       Set frmMuestraFirma.rs = R
    End If
    Set clsFirma = Nothing
    frmMuestraFirma.Show 1
End Sub

Private Sub l_Click()

End Sub

Private Sub grdCliente_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If pnCol = 8 And (Trim(grdCliente.TextMatrix(pnRow, 7)) = "PJ" Or _
                      Trim(grdCliente.TextMatrix(pnRow, 7)) = "AP") Then
        grdCliente.TextMatrix(grdCliente.row, 8) = False
    End If
End Sub

Private Sub grdCliente_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdCliente.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Private Sub txtBanco_EmiteDatos()
lblBanco = Trim(txtBanco.psDescripcion)
If lblBanco <> "" Then
    If Me.cboMonedaBanco.Visible And Me.cboMonedaBanco.Enabled Then
        cboMonedaBanco.SetFocus
    Else
        txtglosaTrans.SetFocus
    End If
End If
End Sub

Private Sub txtCtaAhoAboInt_EmiteDatos()
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim rsCta As New ADODB.Recordset
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = clsMant.GetCuentasPersona(grdCliente.TextMatrix(grdCliente.row, 1), gCapAhorros, True)
    Set clsMant = Nothing
    If rsCta.RecordCount = 0 Then
       MsgBox "Cliente no posee cuentas para Abonar Intereses", vbInformation, "Aviso"
       txtCtaAhoAboInt.Enabled = False
       Exit Sub
    Else
        '***Modificado por ELRO el 20121102, según OYP-RFC098-2012
        If rsCta.RecordCount = 1 Then
            txtCtaAhoAboInt.Text = rsCta(0)
        Else
            txtCtaAhoAboInt.rs = rsCta
            Set rsCta = Nothing
        End If
        '***Fin Modificado por ELRO el 20121102*******************
        txtCtaAhoAboInt.Enabled = True
        If Mid(txtCtaAhoAboInt, 9, 1) <> "" Then
            If CLng(Mid(txtCtaAhoAboInt, 9, 1)) = gMonedaExtranjera Then
                txtCtaAhoAboInt.BackColor = &HC0FFC0
                lblTipoCambio = Format$(nTCV, "#0.0000")
                lblMontoAbono = Format$(txtMonto.value / nTCV, "#,##0.00")
            Else
                txtCtaAhoAboInt.BackColor = &H80000005
                lblTipoCambio = Format$(nTCC, "#0.0000")
                lblMontoAbono = Format$(txtMonto.value * nTCC, "#,##0.00")
            End If
        End If
    End If

 
End Sub

'Private Sub txtCtaAhoAboInt_EmiteDatos()
'If Mid(txtCtaAhoAboInt, 9, 1) <> Mid(txtCuenta.NroCuenta, 9, 1) Then
'    fraAbonoOtraMoneda.Visible = True
'Else
'    fraAbonoOtraMoneda.Visible = False
'End If
'
'Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
'Dim rsCta As New ADODB.Recordset
'Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'Set rsCta = clsMant.GetCuentasPersona(grdCliente.TextMatrix(grdCliente.Row, 1), gCapAhorros, True)
'Set clsMant = Nothing
'txtCtaAhoAboInt.rs = rsCta
'Set rsCta = Nothing
'txtCtaAhoAboInt.Enabled = True
'
'If Mid(txtCtaAhoAboInt, 9, 1) <> "" Then
'    If CLng(Mid(txtCtaAhoAboInt, 9, 1)) = gMonedaExtranjera Then
'        txtCtaAhoAboInt.BackColor = &HC0FFC0
'        lblTipoCambio = Format$(nTCV, "#0.0000")
'        lblMontoAbono = Format$(txtMonto.value / nTCV, "#,##0.00")
'    Else
'        txtCtaAhoAboInt.BackColor = &H80000005
'        lblTipoCambio = Format$(nTCC, "#0.0000")
'        lblMontoAbono = Format$(txtMonto.value * nTCC, "#,##0.00")
'    End If
'Else
'    txtCtaAhoAboInt.BackColor = &H80000005
'End If
'txtGlosa.SetFocus
'End Sub


Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = TxtCuenta.NroCuenta
        ObtieneDatosCuenta sCta
        chkComision_Click 'RIRO20131212 ERS137
    End If
End Sub

'JUEZ 20131212 *************************
Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosaAumDism.SetFocus
    End If
End Sub

Private Sub txtCuentaCargo_LostFocus()
    ValidaCargoCta
End Sub
'END JUEZ ******************************

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
        '***Modificado por ELRO el 20130723, según TI-ERS079-2013****
        'cmdGrabar.SetFocus
        If cboMedioRetiro.Visible Then
            cboMedioRetiro.Enabled = True
            cboMedioRetiro.SetFocus
        Else
            cmdGrabar.SetFocus
        End If
        '***Fin Modificado por ELRO el 20130723, según TI-ERS079-2013
    Else
        '***Modificado por ELRO el 20130723, según TI-ERS079-2013****
        'txtMonto.Enabled = True
        'txtMonto.SetFocus
        If nOperacion <> gPFRetIntAdelantado Then txtMonto.Enabled = True 'JUEZ 20130907
        If cboMedioRetiro.Visible Then
            cboMedioRetiro.SetFocus
        Else
            'JUEZ 20130907 **************
            'txtMonto.Enabled = True
            If txtMonto.Enabled Then
                txtMonto.SetFocus
            Else
                cmdGrabar.SetFocus
            End If
            'END JUEZ *******************
        End If
        '***Fin Modificado por ELRO el 20130723, según TI-ERS079-2013
    End If
End If
End Sub

Private Sub txtGlosaAumDism_GotFocus()
txtGlosaAumDism.SelStart = 0
txtGlosaAumDism.SelLength = Len(txtGlosaAumDism.Text)
End Sub

Private Sub txtGlosaAumDism_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
    If txtMonto.Enabled Then
     If fraMonto.Enabled = True Then
        txtMonto.SetFocus
     End If
    Else
        cmdGrabar.SetFocus
    End If
End If
End Sub

Private Sub txtglosaTrans_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub

Private Sub txtMonto_Change()

' Agregado Por RIRO el 20130501, Proyecto Recaudo
If Trim(txtMonto.Text) = "." Then
    txtMonto.Text = 0
    Exit Sub
End If

If Mid(txtCtaAhoAboInt, 9, 1) <> "" Then
    If Mid(TxtCuenta.NroCuenta, 9, 1) <> Mid(txtCtaAhoAboInt, 9, 1) Then
        If CLng(Mid(txtCtaAhoAboInt, 9, 1)) = gMonedaExtranjera Then
            lblMontoAbono = Format$(txtMonto.value / nTCV, "#,##0.00")
        Else
            lblMontoAbono = Format$(txtMonto.value * nTCC, "#,##0.00")
        End If
    End If
End If

Dim nMonto As Double
Dim nTotalDisp As Double, nITF As Double, nComisionTransf As Double 'RIRO20131212 Se Agrego nComisionTransf
nMonto = txtMonto.value
nComisionTransf = CDbl(lblComision.Caption)
lblITF.Caption = "0.00"

'If gbITFAplica And Me.txtCuenta.Prod <> gCapCTS Then       'Filtra para CTS
If gbITFAplica And Me.TxtCuenta.Prod <> gCapCTS And nOperacion <> gPFAumCapCargoCta Then 'JUEZ 20131218 gPFAumCapCargoCta para no cobra ITF en Aum. Cap. Cargo Cuenta
    If nMonto > gnITFMontoMin Then
        If Not lbITFCtaExonerada Then
            lblITF.Caption = Format(fgITFCalculaImpuesto(nMonto), "#,##0.00")
            '*** BRGO 20110908 ************************************************
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
            If nRedondeoITF > 0 Then
                Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
            End If
            '*** END BRGO
        End If
    End If
End If
            
lblTotal.Caption = Format(0, "#,##0.00")

'RIRO20131212 ERS137
If chkMismoTitular.value = 1 Then
    lblITF.Caption = "0.00"
End If
'END RIRO
If bInstFinanc Then lblITF.Caption = "0.00" 'JUEZ 20140414

nITF = CDbl(lblITF.Caption)
If gbITFAsumidoPF Then nITF = 0

If lblCapital <> "" And lblIntGanado <> "" Then
    nTotalDisp = CDbl(lblCapital) '+ CDbl(lblIntGanado) 'BRGO 20111229
End If
If nOperacion = gPFDismCapEfec Then
    lblSaldoFinal.Caption = Format$(nTotalDisp - nMonto, "#,##0.00")
    lblTotal.Caption = Format$(nMonto - nITF, "#,##0.00")
ElseIf nOperacion = gPFAumCapEfec Or nOperacion = gPFAumCapTasaPactEfec Then
    lblSaldoFinal.Caption = Format$(nTotalDisp + nMonto, "#,##0.00")
    lblTotal.Caption = Format$(nMonto + nITF, "#,##0.00")
ElseIf nOperacion = gPFAumCapchq Or nOperacion = gPFAumCapTasaPactChq Then
    lblSaldoFinal.Caption = Format$(nTotalDisp + nMonto, "#,##0.00")
    lblTotal.Caption = Format$(nITF, "#,##0.00")
ElseIf nOperacion = gPFAumCapTrans Or nOperacion = gPFAumCapTasaPactTrans Then
    
    '***Modificado por ELRO el 20120823, según OYP-RFC024-2012
    'lblSaldoFinal.Caption = Format$(nTotalDisp + nMonto, "#,##0.00")
    'lblTotal.Caption = Format$(nITF, "#,##0.00")
    If nOperacion = gPFAumCapTrans Then
        If chkITFEfectivo.value = 1 Then
            lblTotal.Caption = Format$(nITF, "#,##0.00")
            lblSaldoFinal.Caption = Format$(nTotalDisp + nMonto, "#,##0.00")
        Else
            lblTotal.Caption = Format$(0, "#,##0.00")
            lblSaldoFinal.Caption = Format$(nTotalDisp + (nMonto - nITF), "#,##0.00")
        End If
    Else
        lblTotal.Caption = Format$(nITF, "#,##0.00")
        lblSaldoFinal.Caption = Format$(nTotalDisp + nMonto, "#,##0.00")
    End If
    '***Fin Modificado por ELRO el 20120823*******************

'RIRO20131212 ERS137 ***
ElseIf nOperacion = gPFRetIntAboCtaBanco Then
    lblSaldoFinal.Caption = Format$(nTotalDisp - nMonto, "#,##0.00")
    lblTotal.Caption = Format$(nMonto + IIf(chkComision.value, 0, nComisionTransf), "#,##0.00")
ElseIf nOperacion = gPFCancTransf Then
    lblSaldoFinal.Caption = Format$(nTotalDisp - nMonto, "#,##0.00")
    lblTotal.Caption = Format$(nMonto - IIf(chkComision.value, 0, nComisionTransf), "#,##0.00")
'END RIRO **************
Else
    lblSaldoFinal.Caption = "0.00"
    lblTotal.Caption = Format(nMonto - nITF, "#,##0.00")
End If

'RIRO20131212 ERS137
lblTotalEfectivo.Caption = Format(IIf(chkITFEfectivo.value, Val(lblITF), 0) + IIf(chkComision.value, Val(lblComision), 0), "#,##0.00")
'END RIRO

End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub MuestraFirmas(ByVal sCuenta As String)
    Dim sql As String
    Dim nI As Integer
    Dim sPersona As String
        
    'Determinar la personeria de la cuenta
    If nPersoneria <> PersPersoneria.gPersonaNat Then
        For nI = 1 To Me.grdCliente.Rows - 1
            If Trim(Right(grdCliente.TextMatrix(nI, 3), 5)) = gCapRelPersRepSuplente Or Trim(Right(grdCliente.TextMatrix(nI, 3), 5)) = gCapRelPersRepTitular Then
                sPersona = grdCliente.TextMatrix(nI, 1)
                MuestraFirma sPersona, gsCodAge
            End If
        Next nI
        Exit Sub
    End If
    For nI = 1 To Me.grdCliente.Rows - 1
        If Trim(Right(grdCliente.TextMatrix(nI, 3), 5)) = gCapRelPersTitular Then
            sPersona = grdCliente.TextMatrix(nI, 1)
            MuestraFirma sPersona, gsCodAge
        End If
    Next nI
End Sub

Private Function ValidaCreditosPendientes(psCtaCod As String) As Boolean
    Dim rsCred As ADODB.Recordset
    Dim lnI As Long
    Dim bPendiente As Boolean
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento 'NCapMovimientos
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCapMant As COMDCaptaGenerales.DCOMCaptaGenerales 'DCapMantenimiento
    Set clsCapMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim clsCont As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lsMovNro As String
    
    If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
        If clsCapMov.BuscaCreditosPendientesPago(psCtaCod) Then
            bPendiente = False
            For lnI = 1 To Me.grdCliente.Rows - 1
                If Trim(Right(grdCliente.TextMatrix(lnI, 3), 5)) = gCapRelPersTitular Then
                    
                    Set rsCred = New ADODB.Recordset
                    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
                    Set rsCred = clsCapMov.GetCreditosPendientes(Me.grdCliente.TextMatrix(lnI, 1), gdFecSis)
                    Set clsCapMov = Nothing
                    If Not rsCred Is Nothing Then
                        If Not (rsCred.EOF And rsCred.BOF) Then
                            bPendiente = True
                            'frmCredPendPago.Inicia rsCred
                        End If
                    End If
                    Set rsCred = Nothing
                End If
            Next
            
            If bPendiente Then
                lsMovNro = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                'Bloquea la cuenta
                    clsCapMant.NuevoBloqueoRetiro psCtaCod, gCapMotBlqRetGarantia, "POR CREDITO PENDIENTES DE PAGO", lsMovNro
                    clsCapMant.ActualizaEstadoCuenta psCtaCod, gCapEstBloqRetiro
                Set clsCapMant = Nothing
                Set clsCont = Nothing
                ValidaCreditosPendientes = False
                Exit Function
            End If
        End If
    End If

    ValidaCreditosPendientes = True
End Function

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing
 'rs.Close
 Set oCons = Nothing
End Function

Private Function VerificarAutorizacion() As Boolean
Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim oPers As COMDPersona.UCOMAcceso
Dim rs As New ADODB.Recordset
Dim lnMonTopD As Double
Dim lnMonTopS As Double
Dim lsmensaje As String, lsOpeTpo As String
Dim gsGrupo As String
Dim sCuenta As String, sNivel As String
Dim lbEstadoApr As Boolean
Dim nMonto As Double
Dim nmoneda As Moneda
Dim lsOpeCod As String 'JUEZ 20131218

sCuenta = TxtCuenta.NroCuenta
nMonto = txtMonto.value
nmoneda = CLng(Mid(sCuenta, 9, 1))
'Obtiene los grupos al cual pertenece el usuario
Set oPers = New COMDPersona.UCOMAcceso
    gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
Set oPers = Nothing
 
If nOperacion = gPFCancEfec Or nOperacion = gPFCancTransf Then
    lsOpeTpo = 2
    lsOpeCod = gOpeAutorizacionCancelacion 'JUEZ 20131218
 ElseIf nOperacion = gPFRetInt Or nOperacion = gPFRetIntAboAho _
        Or nOperacion = gPFDismCapEfec Or nOperacion = gPFRetIntAdelantado Or _
        nOperacion = gPFRetIntAboCtaBanco Then ' RIRO20131215 ERS137 gPFRetIntAboCtaBanco
        
    lsOpeTpo = 1
    lsOpeCod = gOpeAutorizacionRetiro
 ElseIf nOperacion = gPFAumCapCargoCta Then 'JUEZ 20131218
    lsOpeTpo = 3
    lsOpeCod = gOpeAutorizacionCargoCuenta
 End If
 

'Verificar Montos
Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    'Set rs = ocapaut.ObtenerMontoTopNivAutRetCan(gsGrupo, lsOpeTpo, gsCodAge)
    Set rs = oCapAut.ObtenerMontoTopNivAutRetCan(gsGrupo, lsOpeTpo, gsCodAge, gsCodPersUser) 'RIRO20141106 ERS159
Set oCapAut = Nothing
 
If Not (rs.EOF And rs.BOF) Then
    lnMonTopD = rs("nTopDol")
    lnMonTopS = rs("nTopSol")
    sNivel = rs("cNivCod")
Else
    MsgBox "Usuario no Autorizado para realizar Operacion", vbInformation, "Aviso"
    VerificarAutorizacion = False
    Exit Function
End If

If nmoneda = gMonedaNacional Then
    If nMonto <= lnMonTopS Then
        VerificarAutorizacion = True
        Exit Function
    End If
Else
    If nMonto <= lnMonTopD Then
        VerificarAutorizacion = True
        Exit Function
    End If
End If
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra
    'oCapAutN.NuevaSolicitudAutorizacion sCuenta, lsOpeTpo, nMonto, gdFecSis, gsCodAge, gsCodUser, nmoneda, gOpeAutorizacionRetiro, sNivel, sMovNroAut
    oCapAutN.NuevaSolicitudAutorizacion sCuenta, lsOpeTpo, nMonto, gdFecSis, gsCodAge, gsCodUser, nmoneda, lsOpeCod, sNivel, sMovNroAut 'JUEZ 20131218
    MsgBox "Solicitud Registrada, comunique a su Admnistrador para la Aprobación..." & Chr$(10) & _
        " No salir de esta operación mientras se realice el proceso..." & Chr$(10) & _
        " Porque sino se procedera a grabar otra Solicitud...", vbInformation, "Aviso"
    VerificarAutorizacion = False
Else
    'Valida el estado de la Solicitud
    If Not oCapAutN.VerificarAutorizacion(sCuenta, lsOpeTpo, nMonto, sMovNroAut, lsmensaje) Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        VerificarAutorizacion = False
    Else
        VerificarAutorizacion = True
    End If
End If
Set oCapAutN = Nothing
End Function

Private Sub txtMonto_LostFocus()
    If nOperacion <> gPFAumCapEfec And nOperacion <> gPFAumCapTasaPactEfec Then
        ValidaTasaInteres
    End If
    If cboPlaza.Visible Then CalculaComision    'RIRO20131212 ERS137
End Sub

Private Sub txtPlazo_GotFocus()
txtPlazo.SelStart = 0
txtPlazo.SelLength = Len(txtPlazo.Text)
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdCheque.Enabled Then
        cmdCheque.SetFocus
    Else
        txtGlosaAumDism.SetFocus
    End If
Else
    KeyAscii = NumerosEnteros(KeyAscii)
End If
End Sub

Private Sub txtPlazo_LostFocus()
'ValidaTasaInteres
End Sub

Private Sub EmiteCalendarioRetiroIntPFMensual(ByVal nCapital As Double, ByVal nTasa As Double, ByVal nPlazo As Long, _
            ByVal dApertura As Date, ByVal nmoneda As Moneda, Optional ByVal nDiasVal As Integer = 0)

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim nIntMens As Double, nIntFinal As Double
Dim dFecVenc As Date, dFecVal As Date
    
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
dFecVenc = DateAdd("d", nPlazo + nDiasVal, dApertura)
dFecVal = DateAdd("d", nDiasVal, dApertura)
nIntMens = clsMant.GetInteresPF(nTasa, nCapital, 30)
nIntFinal = clsMant.GetInteresPF(nTasa, nCapital, nPlazo)

Set clsMant = Nothing

Dim clsPrev As previo.clsprevio
Dim sCad As String
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales

sCad = clsMant.GetPFPlanRetInt(dApertura, nIntMens, nPlazo, nmoneda, nIntFinal, nCapital, nTasa, nDiasVal, dFecVal)
    
Set clsMant = Nothing

Set clsPrev = New previo.clsprevio
    clsPrev.Show sCad, "Plazo Fijo", True, , gImpresora
Set clsPrev = Nothing
End Sub


Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As ConstanteCabecera)
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsConst = clsGen.GetConstante(nCapConst)
Set clsGen = Nothing
Do While Not rsConst.EOF
    
    If nCapConst = gProductoCuentaTipo Then
        If TxtCuenta.Prod = "234" Then
            If rsConst("nConsValor") = 0 Then
                cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
            End If
            rsConst.MoveNext
        Else
           cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
           rsConst.MoveNext
        End If
    Else
        cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    End If
Loop
cboConst.ListIndex = 0
End Sub

'MADM 20101112
Private Function GetMontoDescuento(pnTipoDescuento As CaptacParametro, Optional pnCntPag As Integer = 0, _
                                   Optional pnMoneda As Integer = 1) As Double
Dim oParam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As New ADODB.Recordset

Set oParam = New COMNCaptaGenerales.NCOMCaptaDefinicion
'Modi By Gitu 29-08-2011
    Set rsPar = oParam.GetTarifaParametro(nOperacion, pnMoneda, pnTipoDescuento)
'End Gitu
Set oParam = Nothing

If rsPar.EOF And rsPar.BOF Then
    GetMontoDescuento = 0
Else
    GetMontoDescuento = rsPar("nParValor") * pnCntPag
End If
rsPar.Close
Set rsPar = Nothing
End Function
'END MADM
'***Agregado por ELRO el 20130724, según TI-ERS079-2013****
Private Sub cargarMediosRetiros()
Dim ODCOMConstantes As COMDConstantes.DCOMConstantes
Set ODCOMConstantes = New COMDConstantes.DCOMConstantes
Dim rsMedio As ADODB.Recordset
Set rsMedio = New ADODB.Recordset

Set rsMedio = ODCOMConstantes.devolverMediosRetiros
cboMedioRetiro.Clear

If Not (rsMedio.BOF And rsMedio.EOF) Then
 Do While Not rsMedio.EOF
     cboMedioRetiro.AddItem rsMedio!cConsDescripcion & space(100) & rsMedio!nConsValor
     rsMedio.MoveNext
 Loop
End If
Set rsMedio = Nothing
Set ODCOMConstantes = Nothing
End Sub
'***Fin Agregado por ELRO el 20130724, según TI-ERS079-2013

' ***** Agregado por RIRO 20131102 *****

Private Function validarReglasPersonas() As Boolean
    Dim sReglas() As String
    Dim sGrupos() As String
    Dim sTemporal As String
    Dim v1, v2 As Variant
    Dim bAprobado As Boolean
    Dim intRegla, i, J As Integer
    
    If Trim(strReglas) = "" Then
        validarReglasPersonas = False
        Exit Function
    End If
    sReglas = Split(strReglas, "-")
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 8) = "." Then
            If J = 0 Then
               sTemporal = sTemporal & grdCliente.TextMatrix(i, 7)
            Else
               sTemporal = sTemporal & "," & grdCliente.TextMatrix(i, 7)
            End If
            J = J + 1
        End If
    Next
    If Trim(sTemporal) = "" Then
        validarReglasPersonas = False
        Exit Function
    End If
    sGrupos = Split(sTemporal, ",")
    For Each v1 In sReglas
        bAprobado = True
        For Each v2 In sGrupos
            If InStr(CStr(v1), CStr(v2)) = 0 Then
                bAprobado = False
                Exit For
            End If
        Next
        If bAprobado Then
            If UBound(sGrupos) = UBound(Split(CStr(v1), "+")) Then
                Exit For
            Else
                bAprobado = False
            End If
        End If
    Next
    validarReglasPersonas = bAprobado
End Function

Private Function ObtenerRegla() As String

    Dim nLetraMin, nMedio, i, J As Integer
    Dim sRegla As String
    Dim nReglas() As Integer
    
    nLetraMin = 65
    nMedio = 90
    J = 0
    ReDim Preserve nReglas(0)
    For i = 1 To grdCliente.Rows - 1
        If Trim(grdCliente.TextMatrix(i, 8)) = "." Then
            If Trim(grdCliente.TextMatrix(i, 7)) <> "AP" And Trim(grdCliente.TextMatrix(i, 7)) <> "N/A" Then
                If Len(Trim(grdCliente.TextMatrix(i, 7))) > 0 Then
                    ReDim Preserve nReglas(J)
                    nReglas(J) = CInt(AscW(grdCliente.TextMatrix(i, 7)))
                    J = J + 1
                End If
            End If
        End If
    Next
    
    nLetraMin = 0
    nMedio = 90
    
    If J > 0 Then
        For i = 0 To UBound(nReglas)
            nMedio = 90
            For J = 0 To UBound(nReglas)
                If nReglas(J) > nLetraMin And nReglas(J) <= nMedio Then
                    nMedio = nReglas(J)
                End If
            Next
            nLetraMin = nMedio
            sRegla = sRegla & "+" & ChrW(nMedio)
        Next
        sRegla = Mid(sRegla, 2, Len(sRegla) - 1)
    Else
        sRegla = ""
    End If
    
    ObtenerRegla = sRegla

End Function

' ***** Fin RIRO *****

'JUEZ 20131218 ****************************************************************
Private Sub ValidaCargoCta()
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rs As ADODB.Recordset
Dim sMsg As String
    
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    sMsg = oNCapMov.ValidaCuentaOperacion(txtCuentaCargo.NroCuenta)
    If sMsg <> "" Then
        MsgBox sMsg, vbInformation, "Aviso"
        txtCuentaCargo.SetFocusCuenta
        LimpiaControlesCargoCta
        Exit Sub
    End If
    If Mid(TxtCuenta.NroCuenta, 9, 1) <> Mid(txtCuentaCargo.NroCuenta, 9, 1) Then
        MsgBox "Cuenta debe ser de la misma moneda que del Plazo Fijo", vbInformation, "Aviso"
        txtCuentaCargo.SetFocusCuenta
        LimpiaControlesCargoCta
        Exit Sub
    End If
    
    Set rs = oDCapGen.GetDatosCuentaAho(txtCuentaCargo.NroCuenta)
    fnTpoCtaCargo = rs("nPrdCtaTpo")
    
    If nTipoCuenta <> fnTpoCtaCargo Then
        MsgBox "Cuenta debe ser del mismo tipo de cuenta de la apertura", vbInformation, "Aviso"
'        txtCuentaCargo.SetFocusCuenta
        txtMonto.SetFocus
        LimpiaControlesCargoCta
        Exit Sub
    End If
    Set rs = Nothing
    
    Set rsRelPersCtaCargo = oDCapGen.GetPersonaCuenta(txtCuentaCargo.NroCuenta)
    Set oDCapGen = Nothing
    If Not ValidaRelPersonasCtaCargo Then
        MsgBox "La personas y relaciones de la cuenta a debitar deben ser las mismas que las de la apertura", vbInformation, "Aviso"
        txtMonto.SetFocus
        LimpiaControlesCargoCta
        Exit Sub
    End If
    rsRelPersCtaCargo.MoveFirst
    lblTitularCargoCta.Caption = UCase(PstaNombre(rsRelPersCtaCargo("Nombre")))
    
End Sub

Private Sub LimpiaControlesCargoCta()
    txtCuentaCargo.Age = gsCodAge
    txtCuentaCargo.Cuenta = ""
    lblTitularCargoCta.Caption = ""
    Set rsRelPersCtaCargo = Nothing
    fnTpoCtaCargo = 0
End Sub

Private Function ValidaRelPersonasCtaCargo() As Boolean
    Dim bExisteRelPers As Boolean
    Dim i As Integer
    
    ValidaRelPersonasCtaCargo = False
    
    rsRelPersCtaCargo.MoveFirst
    Do While Not rsRelPersCtaCargo.EOF
        bExisteRelPers = False
        For i = 1 To grdCliente.Rows - 1
            If grdCliente.TextMatrix(i, 1) = rsRelPersCtaCargo("cPersCod") And Trim(Right(grdCliente.TextMatrix(i, 3), 2)) = Trim(Right(rsRelPersCtaCargo("Relacion"), 2)) Then
                bExisteRelPers = True
                Exit For
            End If
        Next i
        If Not bExisteRelPers Then Exit Function
        rsRelPersCtaCargo.MoveNext
    Loop
    
    ValidaRelPersonasCtaCargo = True
End Function

Public Function ValidaFlexVacio() As Boolean
    Dim i As Integer
    For i = 1 To grdCliente.Rows - 1
        If grdCliente.TextMatrix(i, 1) = "" Then
            ValidaFlexVacio = True
            Exit Function
        End If
    Next i
End Function
'END JUEZ *********************************************************************

' RIRO20131212 ERS137
Private Sub CalculaComision()

    Dim idBanco As String
    Dim nPlaza As Integer
    Dim nmoneda As Integer
    Dim nTipo As Integer
    Dim nMonto As Double
    Dim nComision As Double
    Dim oDefinicion As COMNCaptaGenerales.NCOMCaptaDefinicion
    
    If fraTransBco.Enabled And fraTransBco.Visible Then
        Set oDefinicion = New COMNCaptaGenerales.NCOMCaptaDefinicion
        idBanco = Mid(txtBancoDestino.Text, 4, 13)
        nPlaza = Val(Trim(Right(cboPlaza.Text, 8)))
        nmoneda = Val(Mid(TxtCuenta.NroCuenta, 9, 1))
        nTipo = 102 ' Emision
        nMonto = txtMonto.value
        nComision = oDefinicion.getCalculaComision(idBanco, nPlaza, nmoneda, nTipo, nMonto, gdFecSis)
    End If
    lblComision.Caption = Format(Round(nComision, 2), "#0.00")

End Sub

Private Function getTitular() As String

    Dim sTitular As String
    Dim nI As Integer
    
    If txtTitular.Visible And txtTitular.Enabled Then
        sTitular = Trim(txtTitular.Text)
        
    Else
        For nI = 1 To grdCliente.Rows - 1
            If Val(Trim(Right(grdCliente.TextMatrix(nI, 3), 5))) = 10 Then
                sTitular = Trim(grdCliente.TextMatrix(nI, 2))
                nI = grdCliente.Rows
            End If
        Next
    End If
    
    getTitular = sTitular
    
End Function

Private Function getGlosa() As String
    Dim sGlosa As String
    sGlosa = "Banco destino: " & lblBancoDestino.Caption & ", Titular: " & getTitular & ", " & Trim(txtGlosaTransBco.Text)
    getGlosa = UCase(sGlosa)
End Function

Private Sub txtTitular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCuentaDestino.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtCuentaDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosaTransBco.SetFocus
        
    Else
        KeyAscii = Letras(KeyAscii)
        
    End If
End Sub

Private Sub chkMismoTitular_Click()
    If chkMismoTitular.value And chkMismoTitular.Enabled Then
        txtTitular.Text = ""
        txtTitular.Visible = False
    Else
        txtTitular.Text = ""
        txtTitular.Visible = True
        If txtTitular.Enabled And txtTitular.Visible Then txtTitular.SetFocus
    End If
    txtMonto_Change
End Sub

Private Sub txtGlosaTransBco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboMedioRetiro.Visible Then
            txtMonto.Enabled = True
            If cboMedioRetiro.Enabled Then
                cboMedioRetiro.SetFocus
            ElseIf txtMonto.Enabled And fraMonto.Enabled Then
                txtMonto.SetFocus
            End If
        End If
        If Not fraMonto.Enabled Then
            If cmdGrabar.Enabled Then cmdGrabar.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtBancoDestino_Click(psCodigo As String, psDescripcion As String)
CalculaComision
chkComision_Click
End Sub

Private Sub txtBancoDestino_EmiteDatos()
    lblBancoDestino = Trim(txtBancoDestino.psDescripcion)
    If lblBancoDestino <> "" Then
        cboPlaza.SetFocus
    End If
    CalculaComision
    chkComision_Click
End Sub

Private Sub cboPlaza_Click()
    If chkMismoTitular.Visible And chkMismoTitular.Enabled Then chkMismoTitular.SetFocus
    CalculaComision
    txtMonto_Change
End Sub

Private Sub chkComision_Click()
    If nMontoPlazoFijo = -1 Then nMontoPlazoFijo = txtMonto.value
    txtMonto_Change
End Sub

' END RIRO
'EJVG20140207 ***
Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
'END EJVG *******

