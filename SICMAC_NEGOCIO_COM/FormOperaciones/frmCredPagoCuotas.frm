VERSION 5.00
Begin VB.Form frmCredPagoCuotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Cuotas"
   ClientHeight    =   8580
   ClientLeft      =   3330
   ClientTop       =   2220
   ClientWidth     =   7065
   ForeColor       =   &H8000000F&
   Icon            =   "frmCredPagoCuotas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Credito"
      Height          =   1920
      Left            =   30
      TabIndex        =   58
      Top             =   -15
      Width           =   6990
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4800
         TabIndex        =   59
         Top             =   150
         Width           =   2115
         Begin VB.ListBox LstCred 
            Height          =   450
            Left            =   75
            TabIndex        =   3
            Top             =   225
            Width           =   1980
         End
      End
      Begin VB.CommandButton CmdBuscar 
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
         Height          =   360
         Left            =   3840
         TabIndex        =   1
         Top             =   315
         Width           =   900
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto"
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   1515
         Width           =   1005
      End
      Begin VB.Label lblTipoProd 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1365
         TabIndex        =   74
         Top             =   1500
         Width           =   4950
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Crédito"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   1215
         Width           =   855
      End
      Begin VB.Label lblTipoCred 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1365
         TabIndex        =   72
         Top             =   1200
         Width           =   4950
      End
      Begin VB.Label LblAgencia 
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
         Left            =   240
         TabIndex        =   60
         Top             =   780
         Width           =   3465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Pago"
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   7005
      Begin VB.CheckBox ckbPorAfectacion 
         Caption         =   "Afect"
         Height          =   375
         Left            =   6030
         TabIndex        =   76
         Top             =   3840
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton CmdGastos 
         Caption         =   "&Gastos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4695
         TabIndex        =   71
         Top             =   6180
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Frame Frame3 
         Height          =   1920
         Left            =   60
         TabIndex        =   10
         Top             =   4230
         Width           =   6810
         Begin VB.CheckBox chkPreParaAmpliacion 
            Caption         =   "Preparar Crédito para Ampliación"
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
            Left            =   3600
            TabIndex        =   81
            Top             =   1490
            Width           =   3135
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   5
            Top             =   600
            Width           =   1380
         End
         Begin VB.ComboBox CmbForPag 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   210
            Width           =   1785
         End
         Begin SICMACT.ActXCodCta txtCuentaCargo 
            Height          =   375
            Left            =   3120
            TabIndex        =   84
            Top             =   180
            Visible         =   0   'False
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   661
            Texto           =   "Cuenta N°:"
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Pago:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   1530
            Width           =   1005
         End
         Begin VB.Label lblTipoPago 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   82
            Top             =   1500
            Width           =   1710
         End
         Begin VB.Label lblMontoPreCanc 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1335
            TabIndex        =   80
            Top             =   960
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblComPreCanc2 
            AutoSize        =   -1  'True
            Caption         =   "PreCancelación:"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lblComPreCanc1 
            AutoSize        =   -1  'True
            Caption         =   "Comisión"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   900
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Pag.Total. :"
            Height          =   195
            Left            =   4620
            TabIndex        =   70
            Top             =   630
            Width           =   825
         End
         Begin VB.Label lblPagoTotal 
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5475
            TabIndex        =   69
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label lblITF 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3405
            TabIndex        =   68
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "I.T.F. :"
            Height          =   195
            Left            =   2820
            TabIndex        =   67
            Top             =   630
            Width           =   465
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
            TabIndex        =   63
            Top             =   225
            Width           =   1665
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
            TabIndex        =   21
            Top             =   1215
            Width           =   75
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Estado Credito"
            Height          =   195
            Left            =   3180
            TabIndex        =   20
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label LblNewCPend 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1905
            TabIndex        =   19
            Top             =   1230
            Width           =   45
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Cuota Pendiente"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   1230
            Width           =   1710
         End
         Begin VB.Label LblNewSalCap 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1890
            TabIndex        =   17
            Top             =   900
            Width           =   45
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Saldo de Capital"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   915
            Width           =   1680
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto a Pagar"
            Height          =   195
            Left            =   135
            TabIndex        =   15
            Top             =   615
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Prox. fecha Pag :"
            Height          =   195
            Left            =   3180
            TabIndex        =   14
            Top             =   945
            Width           =   1230
         End
         Begin VB.Label LblProxfec 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4575
            TabIndex        =   13
            Top             =   975
            Width           =   45
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Left            =   3210
            TabIndex        =   12
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton CmdPlanPagos 
         Caption         =   "&Plan Pagos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1455
         TabIndex        =   66
         Top             =   6180
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   240
         TabIndex        =   22
         Top             =   195
         Width           =   6570
         Begin VB.CheckBox chkCancelarCred 
            Caption         =   "Cancelación total del crédito"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2820
            TabIndex        =   77
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label LblCalMiViv 
            Appearance      =   0  'Flat
            Caption         =   "Mal Pagador"
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
            Height          =   195
            Left            =   1110
            TabIndex        =   65
            Top             =   1710
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Calificacion :"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   1695
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label LblNomCli 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1110
            TabIndex        =   39
            Top             =   195
            Width           =   4950
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   210
            Width           =   480
         End
         Begin VB.Label LblMonCred 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4290
            TabIndex        =   37
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto del Credito"
            Height          =   195
            Left            =   2835
            TabIndex        =   36
            Top             =   810
            Width           =   1245
         End
         Begin VB.Label LblSalCap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   35
            Top             =   1065
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   105
            TabIndex        =   34
            Top             =   1095
            Width           =   930
         End
         Begin VB.Label LblLinCred 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1110
            TabIndex        =   33
            Top             =   495
            Width           =   4950
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Linea Credito"
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   510
            Width           =   930
         End
         Begin VB.Label LblMoneda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   31
            Top             =   750
            Width           =   1155
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   105
            TabIndex        =   30
            Top             =   780
            Width           =   585
         End
         Begin VB.Label LblTotDeuda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
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
            Height          =   270
            Left            =   4290
            TabIndex        =   29
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Deuda a la Fecha : "
            Height          =   195
            Left            =   2835
            TabIndex        =   28
            Top             =   1095
            Width           =   1410
         End
         Begin VB.Label Lbl2 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1395
            Width           =   870
         End
         Begin VB.Label LblForma 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   26
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label LblTotCuo 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            Height          =   195
            Left            =   1650
            TabIndex        =   25
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Monto Cuota :"
            Height          =   195
            Left            =   2835
            TabIndex        =   24
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label LblMontoCuota 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
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
            Height          =   270
            Left            =   4290
            TabIndex        =   23
            Top             =   1410
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   360
         TabIndex        =   6
         Top             =   6180
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   2535
         TabIndex        =   7
         Top             =   6180
         Width           =   1050
      End
      Begin VB.CommandButton cmdmora 
         Caption         =   "&Mora"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3615
         TabIndex        =   8
         Top             =   6180
         Width           =   1050
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   5775
         TabIndex        =   9
         Top             =   6180
         Width           =   1050
      End
      Begin VB.Label lblICVTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5000
         TabIndex        =   88
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label Label33 
         Caption         =   "Int.Com.Vencido"
         Height          =   285
         Left            =   3480
         TabIndex        =   87
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label lblICV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2100
         TabIndex        =   86
         Top             =   2925
         Width           =   1155
      End
      Begin VB.Label Label34 
         Caption         =   "Int.Com.Vencido"
         Height          =   195
         Left            =   420
         TabIndex        =   85
         Top             =   2925
         Width           =   1245
      End
      Begin VB.Label LblMonCalDin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2100
         TabIndex        =   62
         Top             =   3915
         Width           =   1155
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Monto Calen. Din:"
         Height          =   195
         Left            =   390
         TabIndex        =   61
         Top             =   3915
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Pendiente"
         Height          =   210
         Left            =   435
         TabIndex        =   57
         Top             =   2340
         Width           =   1185
      End
      Begin VB.Label LblCPend 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2100
         TabIndex        =   56
         Top             =   2310
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Monto de Pago"
         Height          =   195
         Left            =   435
         TabIndex        =   55
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label LblMonPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2100
         TabIndex        =   54
         Top             =   2625
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Mora"
         Height          =   195
         Left            =   3495
         TabIndex        =   53
         Top             =   2625
         Width           =   360
      End
      Begin VB.Label LblMora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5000
         TabIndex        =   52
         Top             =   2610
         Width           =   1275
      End
      Begin VB.Label LblGastos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5000
         TabIndex        =   51
         Top             =   2310
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
         Height          =   195
         Left            =   3495
         TabIndex        =   50
         Top             =   2355
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Venc."
         Height          =   195
         Left            =   3480
         TabIndex        =   49
         Top             =   2925
         Width           =   915
      End
      Begin VB.Label LblFecVec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Height          =   270
         Left            =   5000
         TabIndex        =   48
         Top             =   2910
         Width           =   1275
      End
      Begin VB.Label LblDiasAtraso 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4995
         TabIndex        =   47
         Top             =   3510
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dias Atrasados"
         Height          =   195
         Left            =   3480
         TabIndex        =   46
         Top             =   3570
         Width           =   1065
      End
      Begin VB.Label lblMoraC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2100
         TabIndex        =   45
         Top             =   3255
         Width           =   1155
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Mora"
         Height          =   195
         Left            =   405
         TabIndex        =   44
         Top             =   3270
         Width           =   360
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas en Mora"
         Height          =   195
         Left            =   405
         TabIndex        =   43
         Top             =   3585
         Width           =   1125
      End
      Begin VB.Label lblCuotasMora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2100
         TabIndex        =   42
         Top             =   3570
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Met Liq."
         Height          =   195
         Left            =   3480
         TabIndex        =   41
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblMetLiq 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4995
         TabIndex        =   40
         Top             =   3795
         Width           =   795
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu mnuVerCredito 
         Caption         =   "Ver Credito"
      End
   End
End
Attribute VB_Name = "frmCredPagoCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nProducto As Producto

'**** 23-03-06 (Los Gastos se manejaran en el Componente)
'Private MatGastosCancelacion As Variant
'Private nNumGastosCancel As Integer
'Private MatGastosFinal As Variant
'Private nNumGastosFinal As Integer
'*************************************

'ARCV 04-07-2006 (performance)
'Private MatCalendTmp As Variant
'Private MatCalend As Variant
'Private MatCalend_2 As Variant
'Private MatCalendNormalT1 As Variant
'Private MatCalendParalelo As Variant
'Private MatCalendMiVivResult As Variant
'Private MatCalendDistribuido As Variant
'Private MatCalendDistribuido_2 As Variant
'Private MatCalendDistribuidoParalelo As Variant
'Private MatCalendDistribuidoTempo As Variant
Private oCredito As COMNCredito.NCOMCredito
'----------------------

Private nNroTransac As Long
Private bCalenDinamic As Boolean
Private bCalenCuotaLibre As Boolean
Private bRecepcionCmact As Boolean
Private sPersCmac As String
Private vnIntPendiente As Double
Private vnIntPendientePagado As Double
Dim nCalPago As Integer
Dim bDistrib As Boolean
Dim bPrepago As Integer
Dim nCalendDinamTipo As Integer
Dim nMiVivienda As Integer
Dim MatDatos As Variant
Dim sOperacion As String
Dim sPersCod As String
Dim nInteresDesagio As Double

Dim nMontoPago As Double
Dim nITF As Double

' CMACICA_CSTS - 08/11/2003 -------------------------------------------------------------------
Dim nPrestamo As Double
Dim bCuotaCom As Integer
Dim nCalendDinamico As Integer

' RFA
Dim bRFA As Boolean
'----------------------------------------------------------------------------------------------

'Lavado de Dinero
Dim bExoneradaLavado As Boolean
Dim sPerscodLav As String
Dim sNombreLav As String
Dim sDireccionLav As String
Dim sDocIdLav As String

'Variables agregadas para el uso de los Componentes
Private bOperacionEfectivo As Boolean
Private nMontoLavDinero As Double
Private nTC As Double

Private bantxtmonpag As Boolean
'ARCV

Private bActualizaMontoPago As Boolean
'AVMM
Dim pnValorChq As Double

Dim lsTemp As String 'DAOR 20080415
Dim lsAgeCodAct As String
Dim lsTpoProdCod As String
Dim lsTpoCredCod As String
Dim nRedondeoITF As Double 'BRGO 20110914

'** Juez 20120528 ****************
Dim nMontoPag2CuotxVenc As Double
Dim nMontoPagoFecha As Double
Dim bCuotasVencidas As Boolean
'** End Juez *********************

Dim fnPersPersoneria As Integer 'JUEZ 20130925
Dim oDocRec As UDocRec 'EJVG20140408
Dim bInstFinanc As Boolean 'JUEZ 20140411

'*** RIRO 20140530 ERS017 ***
Private nMovNroRVD() As Variant 'Mov del voucher y de la pendiente
Private nMontoVoucher As Currency
'*** Fin RIRO ***************
'ALPA 20150317****************************
Dim lnMontoPendienteIntGracia As Double
Dim lnMontoPagoInicio As Double
'ALPA 20150318****************************pnMontGasto
Dim lnMontGasto As Double
Dim lnMontIntComp As Double
'*****************************************lnMontIntComp
'WIOR 20150404 *********************
Dim fnHayPerdonCamp As Double
Dim fnMontoMinPerdon As Double
Dim fnMontoMaxPerdon As Double
'WIOR FIN **************************
'JUEZ 20150415 *************************
Dim dFecVencCuotaProx As Date
Dim nMontoPagoAlDia As Double
Dim nTipoAdelantoCuota As Integer
'END JUEZ ******************************
Dim fRsPerdonCampRecup As ADODB.Recordset 'WIOR 20150602

'WIOR 20151220 ***
Dim fbEsCredMIVIVIENDA As Boolean
Dim fbMIVIVIENDAAnt As Boolean
'WIOR FIN ********
Dim nAmpl As Integer 'EAAS20180911

Dim MatDatVivienda(1) As Variant 'CTI3 28122018

Dim nTipoPagoAnticipado As Integer 'JOEP20200705 Cambio ReactivaCovid
Dim bValidaActualizacionLiq As Boolean 'RIRO 20200911 Actualización Liquidación

Public Sub RecepcionCmac(ByVal psPersCodCMAC As String)
    bRecepcionCmact = True
    sPersCmac = psPersCodCMAC
    Me.Show 1
End Sub

Private Function HabilitaActualizacion(ByVal pbHabilita As Boolean) As Boolean
    cmdmora.Enabled = pbHabilita
    Frame4.Enabled = Not pbHabilita
    CmbForPag.Enabled = pbHabilita
    LblNumDoc.Enabled = pbHabilita
    TxtMonPag.Enabled = pbHabilita
    chkPreParaAmpliacion.Enabled = pbHabilita
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Or Trim(Mid(ActxCta.NroCuenta, 9, 1)) = "" Then
        TxtMonPag.BackColor = vbWhite
        lblITF.BackColor = vbWhite
    Else
        TxtMonPag.BackColor = vbGreen
        lblITF.BackColor = vbGreen
    End If
    Frame3.Enabled = pbHabilita
    If CmbForPag.ListCount > 0 Then
        CmbForPag.ListIndex = 0
    End If
    If pbHabilita Then
        If TxtMonPag.Enabled And TxtMonPag.Visible Then
            TxtMonPag.SetFocus
        End If
    End If
End Function


Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    If CInt(Trim(Right(CmbForPag, 10))) = gColocTipoPagoCheque Then
        If Trim(LblNumDoc.Caption) = "" Then
            ValidaDatos = False
            MsgBox "Ingrese Numero de Documento", vbInformation, "Aviso"
        End If
    End If
    
End Function

Private Sub LimpiaPantalla()
    LimpiaControles Me, True
    InicializaCombos Me
    Frame3.Enabled = False
    LblEstado.Caption = ""
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    LblNewSalCap.Caption = ""
    LblProxfec.Caption = ""
    LblNewCPend.Caption = ""
    LblEstado.Caption = ""
    bCalenDinamic = False
    LblAgencia.Caption = ""
    Label23.Visible = False
    LblCalMiViv.Visible = False
    nRedondeoITF = 0
    '***Agregado por ELRO el 20120329, según RFC023-2012
    ckbPorAfectacion.Visible = False
    ckbPorAfectacion.value = False
    '***Fin Agregado por ELRO***************************
    '** Juez 20120601
    chkCancelarCred.Enabled = False
    chkCancelarCred.value = False
    '** End Juez
    HabControlesPreCanc True, False 'JUEZ 20130925
    txtCuentaCargo.NroCuenta = ""
    txtCuentaCargo.Visible = False 'JUEZ 20131227
    bInstFinanc = False 'JUEZ 20140411
    'WIOR 20150404 *********************
    fnHayPerdonCamp = 0
    fnMontoMinPerdon = 0
    fnMontoMaxPerdon = 0
    'WIOR FIN **************************
    'JUEZ 20150415 ****************
    dFecVencCuotaProx = "31/12/1899"
    nMontoPagoAlDia = 0
    nTipoAdelantoCuota = 0
    'END JUEZ *********************
    Set fRsPerdonCampRecup = Nothing 'WIOR 20150602
    'WIOR 20151220 ***
    fbEsCredMIVIVIENDA = False
    fbMIVIVIENDAAnt = False
    'WIOR FIN ********
    nTipoPagoAnticipado = 0 'JOEP20200705 Cambio ReactivaCovid
    bValidaActualizacionLiq = False
    lblICV.Caption = "" 'RIRO 20210523
    lblICVTotal.Caption = "" 'RIRO 20210523
End Sub
Private Sub CargaControles()
'Dim oCredGeneral As DCredGeneral
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    'Set R = oCons.RecuperaConstantes(gColocTipoPago)
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 2)  'JUEZ 20131227 Para visualizar la forma de pago Cargo a Cuenta
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    'Call CargaComboConstante(gColocTipoPago, CmbForPag)
    Set fRsPerdonCampRecup = Nothing 'WIOR 20150602
    Exit Sub

ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean

'Dim oCredito As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim oNegCredito As COMNCredito.NCOMCredito
'Dim oGastos As COMNCredito.NCOMGasto
'Dim dParam As COMDCredito.DCOMParametro
'Dim nAnios As Integer
'Dim oAge As COMDConstantes.DCOMAgencias

'Dim clsExo As COMNCaptaServicios.NCOMCaptaServicios
'Dim objRFA As COMDCredito.DCOMRFA

'Dim oPersona As COMDCaptaGenerales.DCOMCaptaGenerales
'ARCV 04-07-2006
'Dim oCredito As COMNCredito.NCOMCredito

Dim rsPers As ADODB.Recordset
Dim rsCredVig As ADODB.Recordset
Dim sAgencia As String
Dim nGastos As Double
Dim nMonPago As Double
Dim nMora As Double
Dim nCuotasMora As Integer
Dim nTotalDeuda As Currency
Dim nInteresDesagio As Double
Dim nMonCalDin As Double
Dim sMensaje As String

'ARCV
Dim nNewSalCap As Double
Dim nNewCPend As Integer
Dim dProxFec As Date
Dim sEstado As String
'ARCV
Dim nCuotaPendiente As Integer
Dim nMoraCalculada As Double
Dim dFechaVencimiento As Date
'---------------
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM
Dim sMensajePerd As String 'WIOR 20150602
Dim pbExcluyeGastos As Boolean 'CTI2 20190101 ERS075-2018

    On Error GoTo ErrorCargaDatos
    nInteresDesagio = 0
    lnMontIntComp = 0
    lnMontGasto = 0
    lnMontoPendienteIntGracia = 0
    pbExcluyeGastos = True  'CTI2 20190101 ERS075-2018
    
    Set oCredito = New COMNCredito.NCOMCredito
    'Call oCredito.CargaDatosPagoCuotas(psCtaCod, gdFecSis, bPrepago, gsCodAge, rsCredVig, sAgencia, nCalendDinamico, bCalenDinamic, bCalenCuotaLibre, _
                                    nMiVivienda, nCalPago, MatCalend, MatCalend_2, MatCalendNormalT1, MatCalendParalelo, MatCalendMiVivResult, _
                                    MatCalendTmp, MatCalendDistribuido, nGastos, nMonPago, nMora, nCuotasMora, nTotalDeuda, nInteresDesagio, _
                                    nMonCalDin, sMensaje, sPersCod, sOperacion, bExoneradaLavado, bRFA, rsPers, bOperacionEfectivo, nMontoLavDinero, nTC)
    Call oCredito.CargaDatosPagoCuotas(psCtaCod, gdFecSis, bPrepago, gsCodAge, rsCredVig, sAgencia, nCalendDinamico, bCalenDinamic, bCalenCuotaLibre, _
                                    nMiVivienda, nCalPago, nGastos, nMonPago, nMora, nCuotasMora, nTotalDeuda, nInteresDesagio, _
                                    nMonCalDin, sMensaje, sPersCod, sOperacion, bExoneradaLavado, bRFA, rsPers, bOperacionEfectivo, nMontoLavDinero, nTC, _
                                    nMontoPago, nITF, vnIntPendientePagado, nNewSalCap, nNewCPend, dProxFec, sEstado, nCuotaPendiente, nMoraCalculada, dFechaVencimiento, _
                                    nMontoPag2CuotxVenc, bCuotasVencidas, lnMontoPendienteIntGracia, lnMontIntComp, lnMontGasto, _
                                    fnHayPerdonCamp, fnMontoMinPerdon, fnMontoMaxPerdon, fRsPerdonCampRecup, nMontoPagoAlDia, pbExcluyeGastos)   '** Juez 20120528 Se agregó nMontoPag2CuotxVenc y bCuotasVencidas
                                    'WIOR 20150404 AGREGO fnHayPerdonCamp, fnMontoMinPerdon, fnMontoMaxPerdon
                                    'JUEZ 20150415 Se agregó nMontoPagoAlDia
                                    'WIOR 20150602 AGREGO fRsPerdonCampRecup
                                    'CTI2 ADD pbExcluyeGastos
    'Set oCredito = Nothing
    
    'RIRO20200911 VALIDA LIQUIDACION ***************
    bValidaActualizacionLiq = oCredito.VerificaActualizacionLiquidacion(psCtaCod)
    If Not bValidaActualizacionLiq Then
        MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar cancelaciones " & _
        "anticipadas a menos que actualice estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
    End If
    'END RIRO **************************************
        
    If Not rsCredVig.BOF And Not rsCredVig.EOF Then
        'EJVG20120719 Valida Pagos Leasing ***
        Select Case Mid(psCtaCod, 6, 3)
            Case "515", "516"
                MsgBox "Ud. debe realizar el pago de este crédito por la opción de PAGO CUOTA ARRENDAMIENTO FINANCIERO", vbInformation, "Aviso"
                CargaDatos = False
                Exit Function
        End Select
        'END EJVG ****************************
        LblAgencia.Caption = sAgencia
        lblMetLiq.Caption = Trim(rsCredVig!cMetLiquidacion)
        LblMontoCuota.Caption = Format(IIf(IsNull(rsCredVig!CuotaAprobada), 0, rsCredVig!CuotaAprobada), "#0.00")
        nCalendDinamTipo = rsCredVig!nCalendDinamTipo
        
        'WIOR 20151220 ***
        fbEsCredMIVIVIENDA = oCredito.EsCredMIVIVENDA(rsCredVig!cTpoProdCod, rsCredVig!cTpoCredCod, 3)
        fbMIVIVIENDAAnt = oCredito.EsCredMIVIVENDA(rsCredVig!cTpoProdCod, rsCredVig!cTpoCredCod)
        'WIOR FIN ********
        
        'If nMiVivienda Then 'WIOR 20160107 - COMENTO
        If fbMIVIVIENDAAnt Then 'WIOR 20160107
            Label23.Visible = True
            LblCalMiViv.Visible = True
            If nCalPago = 1 Then
                LblCalMiViv.Caption = "Buen Pagador"
            Else
                LblCalMiViv.Caption = "Mal Pagador"
            End If
        Else
            Label23.Visible = False
            LblCalMiViv.Visible = False
        End If
                
        CargaDatos = True
        'ALPA 20100607**************************************
        lblTipoCred.Caption = rsCredVig!cTpoCredDes
        lblTipoProd.Caption = rsCredVig!cTpoProdDes
        lsAgeCodAct = rsCredVig!cAgeCodAct
        lsTpoProdCod = rsCredVig!cTpoProdCod
        lsTpoCredCod = rsCredVig!cTpoCredCod
        '***************************************************
        
        
        
        vnIntPendiente = IIf(IsNull(rsCredVig!nintPend), 0, rsCredVig!nintPend)
        vnIntPendientePagado = 0
        nNroTransac = IIf(IsNull(rsCredVig!nTransacc), 0, rsCredVig!nTransacc)
        fnPersPersoneria = rsCredVig!nPersPersoneria 'JUEZ 20130925
        LblNomCli.Caption = PstaNombre(rsCredVig!cPersNombre)
        LblLinCred.Caption = Trim(rsCredVig!cLineaCred)
        LblMoneda.Caption = Trim(rsCredVig!cmoneda)
        LblMonCred.Caption = Format(rsCredVig!nMontoCol, "#0.00")
        
        ' CMACICA_CSTS - 08/11/2003 -------------------------------------------------------------------
        nPrestamo = Format(rsCredVig!nMontoCol, "#0.00")
        bCuotaCom = IIf(IsNull(rsCredVig!bCuotaCom), 0, rsCredVig!bCuotaCom)
        '----------------------------------------------------------------------------------------------
        
        LblSalCap.Caption = Format(rsCredVig!nSaldo, "#0.00")
        LblForma.Caption = Trim(Str(rsCredVig!nCuotasApr))
                
        LblGastos.Caption = nGastos
        LblMonPago.Caption = nMonPago
        
        LblMora.Caption = nMora
        
        'ARCV
        lblMoraC.Caption = nMoraCalculada  'Format(CDbl(MatCalend(0, 6)), "#0.00")
        LblFecVec.Caption = dFechaVencimiento 'MatCalend(0, 0)
        LblCPend.Caption = nCuotaPendiente 'MatCalend(0, 1)
        lblICV.Caption = oCredito.nICVCuota 'RIRO 20210523
        lblICVTotal.Caption = oCredito.nICVTotal 'RIRO 20210523
        '-----------------------
        lblCuotasMora.Caption = nCuotasMora
        LblDiasAtraso.Caption = Trim(Str(rsCredVig!nDiasAtraso))
        lblMetLiq.Caption = Trim(rsCredVig!cMetLiquidacion)
        TxtMonPag.Text = nMonPago
        
        'WIOR 20150404 **********************
        If fnHayPerdonCamp > 0 Then
            TxtMonPag.Text = Round(fnMontoMinPerdon, 2)
        End If
        'WIOR FIN ***************************
    
        LblTotDeuda.Caption = nTotalDeuda
        'ALPA 20150317******************
        lnMontoPagoInicio = nMonPago
        '*******************************
        'JUEZ 20140411 **************************************
        Dim oDInstFinan As COMDPersona.DCOMInstFinac
        Set oDInstFinan = New COMDPersona.DCOMInstFinac
        bInstFinanc = oDInstFinan.VerificaEsInstFinanc(rsCredVig!cPersCod)
        Set oDInstFinan = Nothing
        'END JUEZ *******************************************
        
        TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
        nMontoPagoFecha = Format(TxtMonPag.Text, "#0.00") '** Juez 20120528
        chkCancelarCred.Enabled = True
        
        'Para Generar el calendario Dinamico
        'Si es mivivienda
        LblMonCalDin.Caption = nMonCalDin
        
        If nMiVivienda = 1 And bPrepago = 0 Then
            TxtMonPag.Locked = True
        Else
            TxtMonPag.Locked = False
        End If
        
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Mensaje"
            Exit Function
        End If
        
    bantxtmonpag = True
    
    If bInstFinanc Then nITF = 0 'JUEZ 20140411
    lblITF.Caption = Format(nITF, "#0.00")
    '*** BRGO 20110908 ************************************************
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
    If nRedondeoITF > 0 Then
        Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
    End If
    If Trim(Right(CmbForPag.Text, 10)) = gColocTipoPagoCargoCta Then lblITF.Caption = "0.00" 'JUEZ 20131227
    '*** END BRGO
    lblMontoPreCanc.Caption = "0.00" 'JUEZ 20130925
    lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + CCur(Me.lblITF.Caption), "#0.00")
    LblNewSalCap.Caption = nNewSalCap
    LblNewCPend.Caption = nNewCPend
    If dProxFec <> 0 Then LblProxfec.Caption = dProxFec
    If dProxFec <> 0 Then dFecVencCuotaProx = dProxFec 'JUEZ 20150415
    nTipoAdelantoCuota = 0 'JUEZ 20150415
    nTipoPagoAnticipado = 0 'JOEP20200705 Cambio ReactivaCovid
    LblEstado.Caption = sEstado
    '***Agregado por ELRO el 20120407, según RFC023-2012
    If sEstado = "CANCELADO" Then
        ckbPorAfectacion.Visible = True
    'End If
    '***Fin Agregado por ELRO***************************
    Else '** Juez 20120601
        ckbPorAfectacion.Visible = False
    End If
    
    'chkCancelarCred.Enabled = True '** Juez 20120601
    
    bantxtmonpag = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus

    bActualizaMontoPago = False
    '-----------------------

        
        If Not rsPers.EOF Then
            sPerscodLav = sPersCod
            sNombreLav = rsPers!Nombre
            sDireccionLav = rsPers!Direccion
            sDocIdLav = rsPers!id & " " & rsPers![ID N°]
        End If
        
         '************ firma madm
         Set lafirma = New frmPersonaFirma
         Set ClsPersona = New COMDPersona.DCOMPersonas
        
         Set Rf = ClsPersona.BuscaCliente(sPersCod, BusquedaCodigo)
         If Not Rf.BOF And Not Rf.EOF Then
            If Rf!nPersPersoneria = 1 Then
                ' Call frmPersonaFirma.Inicio(Trim(sPersCod), Mid(sPersCod, 4, 2), False, True) 'COM BY JATO 20210331
                Call frmPersonaFirma.inicio(Trim(sPersCod), Mid(sPersCod, 4, 2), False, False) 'ADD BY JATO 20210331
            End If

         Set Rf = Nothing

        '************ firma madm
        'ALPA 20150316**************************
        chkPreParaAmpliacion.value = 0
        '***************************************
        End If
        'WIOR 20150602 ***
        sMensajePerd = ""
        If Not fRsPerdonCampRecup Is Nothing Then
            If Not (fRsPerdonCampRecup.EOF And fRsPerdonCampRecup.BOF) Then
                If fRsPerdonCampRecup.RecordCount > 0 Then
                sMensajePerd = Trim(fRsPerdonCampRecup!cMovNroSol)
                sMensajePerd = Right(sMensajePerd, 4) & " hoy " & Mid(sMensajePerd, 7, 2) & "/" & Mid(sMensajePerd, 5, 2) & "/" & Mid(sMensajePerd, 1, 4) & " a las " & Mid(sMensajePerd, 9, 2) & ":" & Mid(sMensajePerd, 11, 2) & ":" & Mid(sMensajePerd, 13, 2)
                    If CInt(fRsPerdonCampRecup!nEstado) = 1 Then
                        bantxtmonpag = True
                        TxtMonPag.Text = Format(CDbl(fRsPerdonCampRecup!nMontoPagar), "###," & String(15, "#") & "#0.00")
                        bActualizaMontoPago = True
                        CmbForPag.ListIndex = IndiceListaCombo(CmbForPag, 1)
                        Call TxtMonPag_KeyPress(13)
                        MsgBox "El credito tiene un Perdón por Campaña de Recuperaciones solicitado por el Usuario " & sMensajePerd & " y solo se podra realizar el pago de lo indicado.", vbInformation, "Aviso"
                    Else
                        If MsgBox("El credito tiene un Perdón por Campaña de Recuperaciones solicitado por el Usuario " & sMensajePerd & " pendiente por aprobar." & _
                            IIf(CInt(fRsPerdonCampRecup!nNivel) = 5, Chr(10) & "Falta" & IIf(CInt(fRsPerdonCampRecup!nFaltaApr) = 1, " " & Trim(fRsPerdonCampRecup!nFaltaApr) & " aprobación.", "n " & Trim(fRsPerdonCampRecup!nFaltaApr) & " aprobaciones."), "") & _
                            Chr(10) & "Va a proceder el pago sin el perdón?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                            Call cmdCancelar_Click
                        End If
                    End If
                End If
            End If
        End If
        'WIOR FIN ********
        'RECO20151204 ERS073-2015*****************************
            frmSegSepelioAfiliacion.inicio psCtaCod
        'RECO FIN ********************************************
    Else
        CargaDatos = False
    End If
    
    Exit Function

ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"

End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'VerificaRFA (ActxCta.NroCuenta)
        If bRFA Then
            MsgBox "Este credito es RFA " & vbCrLf & _
                   "Por favor ingrese a la opción de Pagos en RFA", vbInformation, ""
        Else
            '*** CONTROLAR TIEMPO DE CARGA
            'Dim oSeguridad As New COMManejador.Pista
            'Call oSeguridad.InsertarPista("123456", GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, 7, "CARGAR DATOS DEL CREDITO", ActxCta.NroCuenta)
            '***
            'FRHU 20150415 ERS022-2015
            If VerificarSiEsUnCreditoTransferido(ActxCta.NroCuenta) Then
                MsgBox "El Credito seleccionado se encuentra en estado Transferido"
                Exit Sub
            End If
            'FIN FRHU 20150415
            If Not CargaDatos(ActxCta.NroCuenta) Then
                HabilitaActualizacion False
                MsgBox "No se pudo encontrar el Credito, o el Credito No esta Vigente", vbInformation, "Aviso"
                CmdPlanPagos.Enabled = False
                CmdGastos.Enabled = False
            Else
                '*** CONTROLAR TIEMPO DE CARGA
                'Call oSeguridad.InsertarPista("123456", GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, 8, "CARGAR DATOS DEL CREDITO", ActxCta.NroCuenta)
                'Set oSeguridad = Nothing
                '***
                CmdPlanPagos.Enabled = True
                CmdGastos.Enabled = True
                HabilitaActualizacion True
                'WIOR 20150602 ***
                If Not fRsPerdonCampRecup Is Nothing Then
                    If Not (fRsPerdonCampRecup.EOF And fRsPerdonCampRecup.BOF) Then
                        If fRsPerdonCampRecup.RecordCount > 0 Then
                            If CInt(fRsPerdonCampRecup!nEstado) = 1 Then
                                TxtMonPag.Enabled = False
                                chkPreParaAmpliacion.Enabled = False
                                chkCancelarCred.Enabled = False 'APRI MEJORA INC1708110005
                                CmbForPag.Enabled = False
                            End If
                        End If
                    End If
                End If
                'WIOR FIN ********
                
                'APRI20171010 ERS028 - 2017
                Dim oNSegSep As COMNCaptaGenerales.NCOMSeguros
                Set oNSegSep = New COMNCaptaGenerales.NCOMSeguros
                If oNSegSep.SepelioVerificaPagoEfectivo(sPersCod) Then
                   If MsgBox("El cliente registra prima pendiente de pago del Microseguro Vida y Sepelio, ¿Desea realizar la operación?", vbYesNo, "Aviso") = vbYes Then
                        frmSegSepelioCobroPrima.Inicia sPersCod, 1
                   End If
                End If
                Set oNSegSep = Nothing
                'END APRI
                                 'APRI20200430 POR COVID-19
                If CInt(LblCPend.Caption) = 1 Then
                    Dim oCred As COMDCredito.DCOMCredito
                    Set oCred = New COMDCredito.DCOMCredito
                    Dim R As ADODB.Recordset
                    Set R = oCred.ExisteReprogramacionxDesastreNat(ActxCta.NroCuenta, "100929")
                    If R.RecordCount > 0 Then
                        If MsgBox("El Crédito fue Reprogramado por Covid-19, ¿Desea realizar el pago?", vbYesNo, "Aviso") = vbNo Then
                            cmdCancelar_Click
                        End If
                    End If
                    Set oCred = Nothing
                End If
                'END APRI
            End If
        End If
    End If
End Sub

'** Juez 20120523 **********************************
Private Sub chkCancelarCred_Click()
cmdGrabar.Enabled = False
If chkCancelarCred.value = 1 Then
    TxtMonPag.Text = LblTotDeuda.Caption
    TxtMonPag.SetFocus
Else
    If TxtMonPag.Enabled = True Then
        'EJVG20140408 ***
        'TxtMonPag.Text = nMontoPagoFecha
        If Val(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCheque Then
            TxtMonPag.Text = Format(DeducirMontoxITF(oDocRec.fnMonto), "#0.00") 'Restamos el ITF al disponible
        Else
            TxtMonPag.Text = nMontoPagoFecha
        End If
        'END EJVG *******
        TxtMonPag.SetFocus
    End If
    HabControlesPreCanc True, False 'JUEZ 20130925
End If
End Sub
'** End Juez **************************************
'ALPA 20150317*****************************************************************
Private Sub chkPreParaAmpliacion_Click()

'RIRO20200911 VALIDA LIQUIDACION ***************
    Dim oCreditoTmp As COMNCredito.NCOMCredito
    Set oCreditoTmp = New COMNCredito.NCOMCredito
    bValidaActualizacionLiq = False
    bValidaActualizacionLiq = oCreditoTmp.VerificaActualizacionLiquidacion(ActxCta.NroCuenta)
    If chkPreParaAmpliacion.value = 1 Then
        If Not bValidaActualizacionLiq Then
            
            MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar cancelaciones " & _
            "anticipadas a menos que actualice estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
            
            Set oCreditoTmp = Nothing
            chkPreParaAmpliacion.value = 0
            Exit Sub
            
        End If
    End If
'END RIRO **************************************

    If CInt(lblCuotasMora.Caption) > 0 Then
        If chkPreParaAmpliacion.value = 1 Then
            MsgBox "No puede realizar la preparación para ampliación de crédito si tiene alguna cuota en mora", vbInformation, "Aviso!"
        End If
        chkPreParaAmpliacion.value = 0
    End If
    

    
'    If chkPreParaAmpliacion.value = 1 Then
'        If lblMetLiq.Caption <> "GMiC" Then
'            chkPreParaAmpliacion.value = 0
'            MsgBox "Para realizar la operación de preparación para ampliación de crédito", vbInformation, "Aviso!"
'            Exit Sub
'        End If
'    End If
    If chkPreParaAmpliacion.value = 1 Then
        TxtMonPag.Text = lnMontoPendienteIntGracia + lnMontIntComp + lnMontGasto
        TxtMonPag.Enabled = False
    Else
        TxtMonPag.Text = lnMontoPagoInicio
        TxtMonPag.Enabled = True
    End If
     bantxtmonpag = True
    Call TxtMonPag_KeyPress(13)
End Sub
'******************************************************************************
Private Sub CmbForPag_Click()

    LblNumDoc.Caption = ""
    txtCuentaCargo.NroCuenta = ""
    TxtMonPag.Locked = False
    If CmbForPag.ListIndex <> -1 Then
        cmdGrabar.Enabled = False 'JUEZ 20140731
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCheque Then
            'EJVG20140408 ***
            'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstEnValorizacion, CInt(Mid(ActxCta.NroCuenta, 9, 1)))
            'If MatDatos(0) <> "" Then
            '    LblNumDoc.Caption = MatDatos(4)
            '    TxtMonPag.Text = MatDatos(3)
            '    pnValorChq = MatDatos(3)
            'Else
            '    LblNumDoc.Caption = ""
            'End If
            Dim oform As New frmChequeBusqueda
            Set oDocRec = oform.iniciarBusqueda(Val(Mid(ActxCta.NroCuenta, 9, 1)), TipoOperacionCheque.CRED_Pago, ActxCta.NroCuenta)
            Set oform = Nothing
            LblNumDoc.Caption = oDocRec.fsNroDoc
            TxtMonPag.Text = Format(DeducirMontoxITF(oDocRec.fnMonto), "#0.00") 'Restamos el ITF al disponible
            TxtMonPag.Locked = True
            'END EJVG *******
            LblNumDoc.Visible = True
            txtCuentaCargo.Visible = False 'JUEZ 20131227
            ReDim nMovNroRVD(6) 'RIRO 20140530 ERS017
        'JUEZ 20131227 **********************************************
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            If Not SeleccionarCtaCargo Then
                CmbForPag.ListIndex = 0
                Exit Sub
            End If
            LblNumDoc.Visible = False
            txtCuentaCargo.Visible = True
        'END JUEZ ***************************************************
        '*** RIRO 20140530 ERS017 ***
            ReDim nMovNroRVD(6)
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            Dim nMovNro As Long, nMovNroPend As Long
                        
            lnTipMot = 10 ' Pago Credito
            oformVou.iniciarFormularioDeposito CInt(Mid(ActxCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNro, nMovNroPend, sNombre, sDireccion, sDocumento, ActxCta.NroCuenta
            LblNumDoc.Visible = True
            ReDim nMovNroRVD(5)
            nMovNroRVD(0) = nMovNro
            nMovNroRVD(1) = nMovNroPend
            nMovNroRVD(2) = lnMontoPendienteIntGracia
            nMovNroRVD(3) = IIf(chkPreParaAmpliacion.value, 1, 0)
            nMovNroRVD(4) = lnMontIntComp
            nMovNroRVD(5) = lnMontGasto
            If Len(sVaucher) = 0 Then
                LblNumDoc.Caption = sVaucher
            Else
                LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            End If
            TxtMonPag.Text = Format(DeducirMontoxITF(nMontoVoucher), "#,##0.00") 'Restamos el ITF al disponible
           'TxtMonPag.Text = Format(DeducirMontoxITF(oDocRec.fnMonto), "#0.00") 'Restamos el ITF al disponible
        '*** END RIRO ***************
        Else
            ReDim nMovNroRVD(6) 'RIRO 20140530 ERS017
            nMovNroRVD(0) = nMovNro
            nMovNroRVD(1) = nMovNroPend
            nMovNroRVD(2) = lnMontoPendienteIntGracia
            nMovNroRVD(3) = IIf(chkPreParaAmpliacion.value, 1, 0)
            nMovNroRVD(4) = lnMontIntComp
            nMovNroRVD(5) = lnMontGasto
            LblNumDoc.Visible = False
            txtCuentaCargo.Visible = False 'JUEZ 20131227
        End If
        bActualizaMontoPago = True 'JUEZ 20140731
        If TxtMonPag.Enabled Then TxtMonPag.SetFocus 'JUEZ 20140731
    End If
End Sub

Private Sub CmbForPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CInt(Trim(Right(CmbForPag, 10))) = gColocTipoPagoCheque Then
            TxtMonPag.SetFocus
        Else
            'WIOR 20150822 ***
            If TxtMonPag.Visible And TxtMonPag.Enabled Then
                TxtMonPag.SetFocus
            End If
            'WIOR FIN ********
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona

    
    LstCred.Clear
    Set oPers = frmBuscaPersona.inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        'FRHU 20150415 ERS022-2015: Se agrego gColocEstTransferido
        'Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm))
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm, gColocEstTransferido))
        'FIN FRHU 20150415
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
        'FRHU 20150415 ERS022-2015
        'FrmVerCredito.Inicio (oPers.sPersCod)
        Call FrmVerCredito.inicio(oPers.sPersCod, , , , , True)
        'FIN FRHU 20150415
        
        Me.ActxCta.SetFocusCuenta
        
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    'ventana = 1
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
    Call HabilitaActualizacion(False)
    cmdGrabar.Enabled = False
    CmdPlanPagos.Enabled = False
    If Not (oCredito Is Nothing) Then Set oCredito = Nothing
End Sub

Private Sub CmdGastos_Click()
    frmCredAdmiGastos.inicio (ActxCta.NroCuenta)
End Sub

Private Sub cmdGrabar_Click()

'Dim oNegCred As COMNCredito.NCOMCredito
'Dim oDoc As COMNCredito.NCOMCredDoc
'Dim oConstante As COMDConstantes.DCOMConstantes

'ARCV
'Dim oCredito As COMNCredito.NCOMCredito

Dim sError As String
'Dim sTipoCred As String
'Dim MatCalDinam As Variant
'Dim MatCalDinam_2 As Variant
'Dim sCad As String
'Dim sCad2 As String
Dim vPrevio As previo.clsprevio
'Dim oCal As COMDCredito.DCOMCalendario
Dim sImprePlanPago As String
Dim sImpreBoleta As String
Dim oCredD As COMDCredito.DCOMCreditos
Dim sVisPersLavDinero As String 'DAOR 20070511
Dim loLavDinero As frmMovLavDinero
Dim oCred As COMDCredito.DCOMCredActBD 'madm 20091211
Set loLavDinero = New frmMovLavDinero

Dim objPersona As COMDPersona.DCOMPersonas 'JACA 20110512
Set objPersona = New COMDPersona.DCOMPersonas 'JACA 20110512

Dim oMov As COMDMov.DCOMMov 'BRGO 20110914
Set oMov = New COMDMov.DCOMMov 'BRGO 20110914
Dim fnCondicion As Integer 'WIOR 20130301
Dim regPersonaRealizaPago As Boolean 'WIOR 20130301
Dim pnMotCancAnt As Integer 'AMDO 20130408
Dim oDCred As COMDCredito.DCOMCredito 'JUEZ 20140709

'LUCV20160613, Según: ERS004-2016 **********
Dim rsRefinanciado As ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCreditos
Dim bRefinanciado As Boolean
'Fin LUCV20160613 **********

'JIPR 20180625 INICIO
Dim oCredCodFactElect As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim scPersCod As String
Dim scPersIDTpo As String
Dim sSerie As String
Dim sCorrelativo As String
Set R = New ADODB.Recordset
'JIPR 20180625 FIN

'RIRO20200911 VALIDA LIQUIDACION ***************

Set oDCred = New COMDCredito.DCOMCredito
If ValidaPagoAnticipado Then
    bValidaActualizacionLiq = False
    bValidaActualizacionLiq = oCredito.VerificaActualizacionLiquidacion(ActxCta.NroCuenta)
    If Not bValidaActualizacionLiq Then
        MsgBox "El crédito no tiene actualizados sus datos de liquidación, no podrá realizar cancelaciones ni pagos " & _
        "anticipados a menos que actualice estos datos. Deberá comunicarse con el área de T.I.", vbExclamation, "Aviso"
        Exit Sub
    End If
End If


'END RIRO **************************************

Dim pnCancelado(3) As Variant 'CTI3 28122018
pnCancelado(3) = chkCancelarCred.value  'CTI3 28122018
    On Error GoTo ErrorCmdGrabar_Click
    Me.cmdGrabar.Enabled = False 'ALPA 20160427
    Call VerSiClienteActualizoAutorizoSusDatos(sPersCod, sOperacion) 'FRHU ERS077-2015 20151204
    
    'MADM 20110915
    Dim PerAut As COMDPersona.DCOMPersonas
    Dim rs As New ADODB.Recordset
    Dim nValorBloqueo As Boolean
    Set PerAut = New COMDPersona.DCOMPersonas
    Set rs = New ADODB.Recordset
    
    Set rs = PerAut.DevuelvePersBloqueaRecuperaCred(Trim(sPersCod), Trim(ActxCta.NroCuenta))
    Set PerAut = Nothing
    If Not (rs.EOF And rs.BOF) Then
        nValorBloqueo = IIf(rs!dVigente, True, False)
         rs.Close
         Set rs = Nothing
        If nValorBloqueo Then
            MsgBox "Ud. NO podrá continuar, persona registra Bloqueo en Recuperaciones, Comuniquese con el Area de Recuperaciones", vbCritical, "Aviso"
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
    End If
    'END MADM
    'ALPA 20150318*******************************************
    
    'LUCV20160613, ERS004-2016 **********
    Set rsRefinanciado = New ADODB.Recordset
     Set rsRefinanciado = oCredD.DevuelveAprobacionCreditoRefinanciado(Trim(ActxCta.NroCuenta), Trim(gsCodAge))
     Set oCredD = Nothing
     If Not (rsRefinanciado.EOF And rsRefinanciado.BOF) Then
        bRefinanciado = IIf(rsRefinanciado!cCtaCod, True, False)
        rsRefinanciado.Close
        Set rsRefinanciado = Nothing
        If bRefinanciado Then
        MsgBox "Ud. NO podrá continuar, Este credito ha sido Aprobado para ser refinanciado, Comuniquese con el Area de Creditos", vbCritical, "Aviso"
        Me.cmdGrabar.Enabled = True
        Exit Sub
        End If
    End If
    'Fin LUCV20160613 **********
    
'JOEP20200705 Cambio ReactivaCovid
Dim objReactCovid As COMDCredito.DCOMCredito
Dim rsReactCovid As ADODB.Recordset
Set objReactCovid = New COMDCredito.DCOMCredito
Set rsReactCovid = objReactCovid.RestrincionCovidReact(Trim(ActxCta.NroCuenta), Trim(gsCodAge), sOperacion, Me.chkCancelarCred.value, nTipoAdelantoCuota, nTipoPagoAnticipado)
If Not (rsReactCovid.EOF And rsReactCovid.BOF) Then
    If rsReactCovid!MsgBox <> "" Then
        MsgBox rsReactCovid!MsgBox, vbInformation, "Aviso"
        Exit Sub
    End If
End If
'JOEP20200705 Cambio ReactivaCovid
    
    
    'WIOR 20150404 **********************
    If fnHayPerdonCamp > 0 Then
        If CDbl(TxtMonPag.Text) < Round(fnMontoMinPerdon, 2) Then
            MsgBox "El monto del pago no puede ser menor a " & fnMontoMinPerdon & ", ya que realiza el Perdon de Mora Campaña de Recuperaciones.", vbInformation, "Mensaje"
            TxtMonPag.Text = fnMontoMinPerdon
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
        
        If CDbl(TxtMonPag.Text) > Round(fnMontoMaxPerdon, 2) Then
            MsgBox "El monto del pago no puede ser mayor a " & fnMontoMaxPerdon & ", ya que realiza el Perdon de Mora Campaña de Recuperaciones.", vbInformation, "Mensaje"
            TxtMonPag.Text = fnMontoMaxPerdon
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
    End If
    'WIOR FIN ***************************
    
    'WIOR 20151220 ***
    'If fbEsCredMIVIVIENDA Then
    If fbEsCredMIVIVIENDA Or lsTpoCredCod = "854" Or lsTpoCredCod = "853" Then   'CTI3 ERS085-2018
        Dim bPrepagoCanAticip As Boolean
        bPrepagoCanAticip = False
                
        If Trim(LblEstado.Caption) = "CANCELADO" Then
            Dim lbCancAntPA As Boolean
            Set oDCred = New COMDCredito.DCOMCredito
            lbCancAntPA = oDCred.VerificaSiEsCancelacionAnticipada(ActxCta.NroCuenta, gdFecSis)
            If lbCancAntPA Then
                bPrepagoCanAticip = True
            End If
        End If
        
        If Not bPrepagoCanAticip Then
            If ValidaPagoAnticipado Then
                bPrepagoCanAticip = True
            End If
        End If
        
        If bPrepagoCanAticip Then
            'CTI3: ERS085-2018
            'If Not frmCredMiViviendaAlertas.Inicio(Trim(ActxCta.NroCuenta), CDbl(TxtMonPag.Text)) Then
            If Not frmCredMiViviendaAlertas.inicio(Trim(ActxCta.NroCuenta), CDbl(TxtMonPag.Text), MatDatVivienda) Then
                Me.cmdGrabar.Enabled = True 'ALPA 20160427
                Exit Sub
            End If
        Else
            Call frmCredMiViviendaAlertas.DarBajaCanPagoAnticipado(Trim(ActxCta.NroCuenta))
        End If
    End If
    'WIOR FIN ********
    
    'JUEZ 20150415 **************************************************************
    If chkPreParaAmpliacion.value = 0 Then
        If lblTipoPago.Caption <> "Pago Anticipado" Then
            '** Juez 20120601 ***************************************
            If Not bCuotasVencidas Then
                If Me.chkCancelarCred.value = 0 And nCalendDinamico = 0 And bPrepago = 0 Then
                    Dim nCuotasPagadas As Integer
                    Dim dFechPuedePagar As Date
                    Dim rsFec As ADODB.Recordset
                    Dim oCredFec As COMNCredito.NCOMCredito
                    Set oCredFec = New COMNCredito.NCOMCredito
            
                    Set rsFec = oCredFec.obtieneNroCuotasPorVencUltimaCuotaPagada(Trim(ActxCta.NroCuenta), gdFecSis)
                    Set oCredFec = Nothing
                    nCuotasPagadas = rsFec!CuotasPagadas
                    If rsFec!FecPuedePagar = "01/01/1900" Then
                        dFechPuedePagar = rsFec!dVenc
                    Else
                        dFechPuedePagar = rsFec!FecPuedePagar
                    End If
                    'Set rsFec = Nothing
                    If nCuotasPagadas >= 2 Then
                        MsgBox "NO puede continuar con el pago, ya pagó las 2 cuotas permitidas." & Chr(13) & _
                                "Este pago podrá ser realizado a partir del " & dFechPuedePagar + 1, vbInformation, "Aviso"
                        Me.cmdGrabar.Enabled = True 'ALPA 20160427
                        Exit Sub
                    End If
                End If
            End If
            '** End Juez ********************************************
        End If
    End If
    'END JUEZ *******************************************************************
    
    If CInt(Trim(Right(CmbForPag.Text, 2))) = gColocTipoPagoCheque Then
        If Trim(Me.LblNumDoc.Caption) = "" Then
            MsgBox "Cheque No es Valido", vbInformation, "Aviso"
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Me.CmbForPag.SetFocus
            Exit Sub
        End If
        If Not ValidaSeleccionCheque Then
            MsgBox "Ud. debe seleccionar el Cheque para continuar", vbInformation, "Aviso"
            If CmbForPag.Visible And CmbForPag.Enabled Then CmbForPag.SetFocus
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
        'If IsArray(MatDatos) Then
        '    If Trim(MatDatos(3)) = "" Then
        '        MatDatos(3) = "0.00"
        '    End If
        '    If Trim(TxtMonPag.Text) = "" Then
        '        TxtMonPag.Text = "0.00"
        '    End If
        '    If CDbl(TxtMonPag.Text) > CDbl(MatDatos(3)) Then
        '        MsgBox "Monto de Pago No Puede Ser Mayor que el Monto de Cheque", vbInformation, "Aviso"
        '        TxtMonPag.SetFocus
        '        Exit Sub
        '    End If
        'End If
        'Validar Monto de Cheque... AVMM  19-10-2006
        'Dim nValorCh As Double
        Dim nDifValorCh As Double
        Dim nDifTotalCh As Double
        Dim nPagadoTotal As Double
        'Set oCredD = New COMDCredito.DCOMCreditos
        'nValorCh = oCredD.ObtenerMontoCheque(LblNumDoc.Caption)
        'ALPA 20090805****************************************
        'nDifValorCh = Format((CDbl(pnValorChq) - CDbl(nValorCh)), "0.00")
        'nDifValorCh = Format(CDbl(MatDatos(3)), "0.00")
        nDifValorCh = Format(CDbl(oDocRec.fnMonto), "0.00")
        '*****************************************************
        nPagadoTotal = CDbl(lblPagoTotal.Caption)
        nDifTotalCh = (CDbl(nDifValorCh) - CDbl(nPagadoTotal))
        If nDifTotalCh < 0 Then
            'MsgBox "No se puede realizar el Pago con Cheque solo dispone de: " & nDifValorCh, vbInformation, "Aviso"
            MsgBox "No se puede realizar el Pago con Cheque solo dispone de: " & Format(nDifValorCh, gsFormatoNumeroView), vbInformation, "Aviso"
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
    End If
    
    'JUEZ 20150415 ***************************************************************
    If lsTpoCredCod = "853" Or lsTpoCredCod = "854" Then
        'JUEZ 20140709 *********************************************
        If (bCalenDinamic Or bPrepago = 1) And (nMontoPago < CDbl(LblTotDeuda.Caption)) And (nMontoPago > CDbl(LblMonCalDin.Caption)) Then
            Set oDCred = New COMDCredito.DCOMCredito
            Set rs = oDCred.RecuperaCredMantPrepago(ActxCta.NroCuenta, gdFecSis)
            If rs.EOF And rs.BOF Then
                MsgBox "Este pago fue configurado como prepago y calendario dinámico, sin embargo la fecha de configuración fue en dias anteriores, " & Chr(13) & _
                        "es necesario que esta configuración sea el mismo dia del pago. Favor de verificarlo con el Jefe de Agencia", vbInformation, "Aviso"
                Me.cmdGrabar.Enabled = True 'ALPA 20160427
                Exit Sub
            End If
            Set oDCred = Nothing
        End If
        'END JUEZ **************************************************
    End If
    'END JUEZ ********************************************************************
    
    'JUEZ 20131227 ******************************************************************
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Len(txtCuentaCargo.NroCuenta) <> 18 Then
            MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "Aviso"
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
        
        Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, nMontoPago) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "Aviso"
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
        
        'Verifica actualización Persona
        Dim lsDireccionActualizada As String
        Dim oPersona As New COMNPersona.NCOMPersona
        
        If oPersona.NecesitaActualizarDatos(sPersCod, gdFecSis) Then
             MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & LblNomCli.Caption, vbInformation, "Aviso"
             Dim foPersona As New frmPersona
             If Not foPersona.realizarMantenimiento(sPersCod, lsDireccionActualizada) Then
                 MsgBox "No se ha realizado la actualización de los datos de " & LblNomCli.Caption & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
                 Me.cmdGrabar.Enabled = True 'ALPA 20160427
                 Exit Sub
             End If
        End If
        lsDireccionActualizada = ""
    End If
    'END JUEZ ***********************************************************************

    'RIRO 20140530 ERS017 ***********************************************************
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
        If Trim(Me.LblNumDoc.Caption) = "" Then
            MsgBox "Voucher No es Valido", vbInformation, "Aviso"
            Me.CmbForPag.SetFocus
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
        Dim nPagadoTotalV As Double
        nPagadoTotalV = CDbl(lblPagoTotal.Caption)
        If nPagadoTotalV > nMontoVoucher Then
            MsgBox "No se puede realizar el Pago con Voucher solo dispone de: " & Format(nMontoVoucher, "#0.00"), vbInformation, "Aviso"
            Me.cmdGrabar.Enabled = True 'ALPA 20160427
            Exit Sub
        End If
    End If
    'END RIRO ***********************************************************************

    'JACA 20110512 *****VERIFICA SI LAS PERSONAS CUENTAN CON OCUPACION E INGRESO PROMEDIO
        Dim rsPersVerifica As Recordset
        Dim i As Integer
        Set rsPersVerifica = New Recordset
        
            Set rsPersVerifica = objPersona.ObtenerDatosPersona(sPersCod)
            If rsPersVerifica!nPersIngresoProm = 0 Or rsPersVerifica!cActiGiro1 = "" Then
                If MsgBox("Necesita Registrar la Ocupacion e Ingreso Promedio de: " + LblNomCli, vbYesNo) = vbYes Then
                    'frmPersona.Inicio Me.grdCliente.TextMatrix(i, 1), PersonaActualiza
                    frmPersOcupIngreProm.inicio sPersCod, LblNomCli, rsPersVerifica!cActiGiro1, rsPersVerifica!nPersIngresoProm
                End If
            End If
      
    'JACA END***************************************************************************
    'WIOR 20121019 Clientes Observados *************************************
            Dim oDPersona As COMDPersona.DCOMPersona
            Dim rsPersona As ADODB.Recordset
            Set oDPersona = New COMDPersona.DCOMPersona
            Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(sPersCod))
         
            If rsPersona.RecordCount > 0 Then
                If Not (rsPersona.EOF And rsPersona.BOF) Then
                    If Trim(rsPersona!sUsual) = "3" Then
                        MsgBox "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                        Call frmPersona.inicio(Trim(sPersCod), PersonaActualiza)
                    End If
                End If
            End If
    'WIOR FIN ***************************************************************
    
     '*** AMDO20130705 TI-ERS063-2013
        Dim oDPersonaAct As COMDPersona.DCOMPersona
        Set oDPersonaAct = New COMDPersona.DCOMPersona
                            If oDPersonaAct.VerificaExisteSolicitudDatos(sPersCod) Then
                                MsgBox Trim("SE SOLICITA DATOS DEL CLIENTE: " & LblNomCli.Caption) & "." & Chr(10), vbInformation, "Aviso"
                                Call frmActInfContacto.inicio(sPersCod)
                            End If
    '***END AMDO
    
    If MsgBox("Se va a Efectuar el Pago del Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoEfectivo Then
'            nMovNroRVD(0) = 0
'            nMovNroRVD(1) = 0
            nMovNroRVD(2) = lnMontoPendienteIntGracia
            nMovNroRVD(3) = IIf(chkPreParaAmpliacion.value, 1, 0)
            nMovNroRVD(4) = lnMontIntComp
            nMovNroRVD(5) = lnMontGasto
    End If
'    If chkPreParaAmpliacion.value = 1 Then
'        TxtMonPag.Text = lnMontoPagoInicio
'        Call TxtMonPag_KeyPress(13)
'    End If
    'WIOR 20130301 ***SEGUN TI-ERS005-2013  SE COMENTO - INICIO ************************************
    ''** Juez 20120323 ************************************ Modif Juez 20120514
    'Dim regPersonaRealizaPago As Boolean
    'Dim rsAgeParam As Recordset
    'regPersonaRealizaPago = False
    'Dim lnMonto As Double
    'lnMonto = TxtMonPag.Text
    
    'Dim ObjTc As COMDConstSistema.NCOMTipoCambio
    'Set ObjTc = New COMDConstSistema.NCOMTipoCambio
    'nTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
    'Set ObjTc = Nothing
    
    'Dim objCred As COMNCredito.NCOMCredito
    
    'Set objCred = New COMNCredito.NCOMCredito
    'Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
    'Set objCred = Nothing
    
    'If Mid(ActxCta.NroCuenta, 9, 1) = 2 Then
    '    lnMonto = Round(lnMonto * nTC, 2)
    'End If
    'If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
    '    If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
    '        frmCredPagoCuotasPersRealiza.Inicia
    '        regPersonaRealizaPago = frmCredPagoCuotasPersRealiza.PersRegistrar
    '        If Not regPersonaRealizaPago Then
    '                MsgBox "Se va a proceder a Anular el Pago de la Cuota"
    '            Exit Sub
    '        End If
    '    End If
    'End If
    ''** End Juez *****************************************
    'WIOR 20130301 ***SEGUN TI-ERS005-2013  SE COMENTO -FIN **************************************
    
    'JUEZ 20120921 ********************* VERIFICA CANCELACION ANTICIPADA
    If Trim(LblEstado.Caption) = "CANCELADO" Then
        Dim lbCancAnt As Boolean, lbRegMotivo As Boolean
        'Dim oDCredCancAnt As COMDCredito.DCOMCredito
        pnCancelado(0) = Trim(LblEstado.Caption) 'CTI3 28122018
        'Set oDCredCancAnt = New COMDCredito.DCOMCredito
        Set oDCred = New COMDCredito.DCOMCredito 'JUEZ 20140709
        'lbCancAnt = oDCredCancAnt.VerificaSiEsCancelacionAnticipada(ActxCta.NroCuenta, gdFecSis)
        lbCancAnt = oDCred.VerificaSiEsCancelacionAnticipada(ActxCta.NroCuenta, gdFecSis) 'JUEZ 20140709
        If lbCancAnt Then
            MsgBox "Este pago es una cancelación anticipada. Deberá registrar el motivo", vbInformation, "Aviso"
            frmCredPagoMotivoCancAnt.Inicia
            lbRegMotivo = frmCredPagoMotivoCancAnt.RegistraMotivo
            If Not lbRegMotivo Then
                MsgBox "Se procederá a Anular la Cancelación de Crédito", vbInformation, "Aviso"
                Me.cmdGrabar.Enabled = True 'ALPA 20160427
                Exit Sub
            Else
                Call frmCredPagoMotivoCancAnt.dInsertaMotivoCancAnticipada(ActxCta.NroCuenta, frmCredPagoMotivoCancAnt.MotivoCanc, frmCredPagoMotivoCancAnt.MotivoCancOtros)
                 'CTI3
                'pnMotCancAnt = frmCredPagoMotivoCancAnt.MotivoCanc 'AMDO 20130408
                 pnCancelado(1) = lbCancAnt
                 pnCancelado(2) = frmCredPagoMotivoCancAnt.MotivoCanc 'CTI3 28122018
            End If
        Else 'CTI3
            pnCancelado(1) = lbCancAnt 'CTI3
            pnCancelado(2) = ""        'CTI3
        End If
        Set oDCred = Nothing
    End If
    'END JUEZ **********************************************************
        
    'JUEZ 20150415 ********************************************************
    If Trim(LblEstado.Caption) <> "CANCELADO" Then
      '  If lsTpoCredCod <> "853" And lsTpoCredCod <> "854" Then 'CTI3:ERS085-2018
            If ValidaPagoAnticipado Then
                'If frmCredMntPagoAnticipado.Registrar(ActxCta.NroCuenta, CInt(LblForma.Caption) - CInt(LblCPend.Caption)) Then
                If frmCredMntPagoAnticipado.Registrar(ActxCta.NroCuenta, CInt(LblForma.Caption) - (CInt(LblCPend.Caption) - 1)) Then 'JUEZ 20150625
                    nCalendDinamico = 1
                    bCalenDinamic = True
                    Call TxtMonPag_KeyPress(13)
                Else
                    MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                    Me.cmdGrabar.Enabled = True 'ALPA 20160427
                    'MARG20180619 Pag. Ant.******************
                    Call AnularPagoAnticipado
                    'END MARG********************************
                    Exit Sub
                End If
            End If
        'End If
    End If
    'END JUEZ *************************************************************
    
    'CTI3 28122018
    If Trim(LblEstado.Caption) <> "CANCELADO" Then pnCancelado(0) = Trim(LblEstado.Caption)

'    Dim nPorcRetCTS As Double, nMontoLavDinero As Double, nTC As Double
'    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim nMonto As Double
    Dim nmoneda As Integer
    'nMonto = CDbl(TxtMonPag.Text)
    nMonto = CDbl(nMontoPago) 'JUEZ 20140220
    Dim sPersLavDinero As String
    nmoneda = CLng(Mid(ActxCta.NroCuenta, 9, 1))
      
    
    sPersLavDinero = ""
    If bOperacionEfectivo Then
        If Not bExoneradaLavado Then
            If CDbl(TxtMonPag.Text) >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 1402208
                 Call IniciaLavDinero(loLavDinero)
                 'ALPA 20081009*******************************************************
                 'sperslavdinero = loLavDinero.Inicia(, , , , False, True, nMonto, ActxCta.NroCuenta, Mid(Me.Caption, 15), True, "", , , , , nMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                 sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, ActxCta.NroCuenta, Me.Caption, True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                 '********************************************************************
                 If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End
            End If
        End If
    End If
    'WIOR 20130301 ***SEGUN TI-ERS005-2013 ************************************************************
    If loLavDinero.OrdPersLavDinero = "Exit" Then
        Dim oPersonaSPR As UPersona_Cli
        Dim oPersonaU As COMDPersona.UCOMPersona
        Dim nTipoConBN As Integer
        Dim sConPersona As String
        Dim pbClienteReforzado As Boolean
        Dim rsAgeParam As Recordset
        Dim objCred As COMNCredito.NCOMCredito
        Dim lnMonto As Double, lnTC As Double
        Dim ObjTc As COMDConstSistema.NCOMTipoCambio
        
        
        Set oPersonaU = New COMDPersona.UCOMPersona
        Set oPersonaSPR = New UPersona_Cli
        
        regPersonaRealizaPago = False
        pbClienteReforzado = False
        fnCondicion = 0

        oPersonaSPR.RecuperaPersona sPersCod
                            
        If oPersonaSPR.Personeria = 1 Then
            If oPersonaSPR.Nacionalidad <> "04028" Then
                sConPersona = "Extranjera"
                fnCondicion = 1
                pbClienteReforzado = True
            ElseIf oPersonaSPR.Residencia <> 1 Then
                sConPersona = "No Residente"
                fnCondicion = 2
                pbClienteReforzado = True
            ElseIf oPersonaSPR.RPeps = 1 Then
                sConPersona = "PEPS"
                fnCondicion = 4
                pbClienteReforzado = True
            ElseIf oPersonaU.ValidaEnListaNegativaCondicion(IIf(Trim(oPersonaSPR.ObtenerDNI) = "", oPersonaSPR.ObtenerNumeroDoc(0), oPersonaSPR.ObtenerDNI), oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                If nTipoConBN = 1 Or nTipoConBN = 3 Then
                    sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                    fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                    pbClienteReforzado = True
                End If
            End If
        Else
            If oPersonaU.ValidaEnListaNegativaCondicion(oPersonaSPR.ObtenerDNI, oPersonaSPR.ObtenerRUC, nTipoConBN, oPersonaSPR.NombreCompleto) Then
                If nTipoConBN = 1 Or nTipoConBN = 3 Then
                    sConPersona = IIf(nTipoConBN = 1, "Negativa", "PEPS")
                    fnCondicion = IIf(nTipoConBN = 1, 3, 4)
                    pbClienteReforzado = True
                End If
            End If
        End If
        
        If pbClienteReforzado Then
            MsgBox "El Cliente: " & Trim(LblNomCli.Caption) & " es un Cliente de Procedimiento Reforzado (Persona " & sConPersona & ")", vbInformation, "Aviso"
            frmPersRealizaOpeGeneral.Inicia Me.Caption & " (Persona " & sConPersona & ")", sOperacion
            regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
            
            If Not regPersonaRealizaPago Then
                MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                Me.cmdGrabar.Enabled = True 'ALPA 20160427
                'MARG20180619 Pag. Ant.******************
                Call AnularPagoAnticipado
                'END MARG********************************
                Exit Sub
            End If
        Else
            fnCondicion = 0
            lnMonto = nMonto
            pbClienteReforzado = False
            
            Set ObjTc = New COMDConstSistema.NCOMTipoCambio
            lnTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set ObjTc = Nothing
        
        
            Set objCred = New COMNCredito.NCOMCredito
            Set rsAgeParam = objCred.obtieneCredPagoCuotasAgeParam(gsCodAge)
            Set objCred = Nothing
            
            If Mid(ActxCta.NroCuenta, 9, 1) = 2 Then
                lnMonto = Round(lnMonto * lnTC, 2)
            End If
        
            If Not (rsAgeParam.EOF And rsAgeParam.BOF) Then
                If lnMonto >= rsAgeParam!nMontoMin And lnMonto <= rsAgeParam!nMontoMax Then
                    frmPersRealizaOpeGeneral.Inicia Me.Caption, sOperacion
                    regPersonaRealizaPago = frmPersRealizaOpeGeneral.PersRegistrar
                    If Not regPersonaRealizaPago Then
                        MsgBox "Se va a proceder a Anular el Pago de la Cuota", vbInformation, "Aviso"
                        Me.cmdGrabar.Enabled = True 'ALPA 20160427
                        'MARG20180619 Pag. Ant.******************
                        Call AnularPagoAnticipado
                        'END MARG********************************
                        Exit Sub
                    End If
                End If
            End If
            
        End If
    End If
    'WIOR FIN ***************************************************************
    '**DAOR 20080415, Para evitar registros duplicados *******************
    lsTemp = MDISicmact.SBBarra.Panels(1).Text
    MDISicmact.SBBarra.Panels(1).Text = "Procesando ....."
    Me.cmdGrabar.Enabled = False
    '********************************************************************

    ''''''''''''''''''''''''''''''
    'Set oCredito = New COMNCredito.NCOMCredito
                                
    'MatGastosFinal, nNumGastosFinal(23-03-06)
    
    '*** CONTROLAR TIEMPO DE PAGO
'    Dim oSeguridad As New COMManejador.Pista
'    Call oSeguridad.InsertarPista("654321", GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, 7, "PAGO DE CUOTA", ActxCta.NroCuenta)

    oCredito.pbExcluyeGastos = True 'CTI2 20181229
    
    Call oCredito.GrabarPagoCuotas(ActxCta.NroCuenta, IIf(fbMIVIVIENDAAnt = True, 1, 0), nCalPago, nMontoPago, _
                            gdFecSis, lblMetLiq.Caption, CInt(Trim(Right(CmbForPag.Text, 10))), gsCodAge, gsCodUser, gsCodCMAC, Trim(LblNumDoc.Caption), _
                            bRecepcionCmact, sPersCmac, vnIntPendiente, vnIntPendientePagado, bPrepago, sPersLavDinero, CCur(lblITF.Caption), _
                            nInteresDesagio, CDbl(LblTotDeuda.Caption), bCalenDinamic, CDbl(LblMonCalDin.Caption), nCalendDinamTipo, gsNomAge, CInt(ActxCta.Prod), _
                            LblNomCli.Caption, LblMoneda.Caption, nNroTransac, LblProxfec.Caption, sLpt, gsInstCmac, IIf(Trim(Right(Me.CmbForPag.Text, 2)) = "2", True, False), _
                            Me.LblNumDoc.Caption, sError, sImprePlanPago, sImpreBoleta, CInt(LblDiasAtraso.Caption), gsProyectoActual, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, lsAgeCodAct, lsTpoProdCod, lsTpoCredCod, ckbPorAfectacion, , pnCancelado, nTipoAdelantoCuota, CDbl(lblMontoPreCanc.Caption), txtCuentaCargo.NroCuenta, _
                            oDocRec.fnTpoDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta, nMovNroRVD)
                            'By Capi 28012007 se agrego parametros sOrdPersLavDinero, sReaPersLavDinero, sBenPersLavDinero
                            'Parametro ckbPorAfectacion agregado por ELRO el 20120329, según RFC023-2012
                            'Parametro pnMotCancAnt agregado por AMDO 20130315 CrediPremiazo
                            'lblMontoPreCanc.Caption JUEZ 20130925
                            'txtCuentaCargo.NroCuenta JUEZ 20131227
                            'EJVG20140408 Se agregó oDocRec
                            'WIOR 20151224 SE CAMBIO nMivivienda POR IIf(fbMIVIVIENDAAnt = True, 1, 0)
    'Set oCredito = Nothing
    'ALPA 20081010*********************************************
    If gnMovNro > 0 Then
     'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
     Call oMov.InsertaMovRedondeoITF("", 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption), gnMovNro) 'BRGO 20110914
     Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
    End If
    Set oMov = Nothing
    'madm 20091211 --------------------------
    Set oCred = New COMDCredito.DCOMCredActBD
    Call oCred.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , IIf(bCalenDinamic, 0, -1))
    '----------------------------------------
    
    'WIOR 20150404 **************************
    If gnMovNro > 0 Then
        If fnHayPerdonCamp > 0 Then
            Set oDCred = New COMDCredito.DCOMCredito
            Call oDCred.ActualizaPerdonMoraCampRecupXPago(fnHayPerdonCamp, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gnMovNro)
        End If
    End If
    'WIOR FIN *******************************
    
    'WIOR 20150602 ***
    If gnMovNro > 0 Then
        If Not fRsPerdonCampRecup Is Nothing Then
            If Not (fRsPerdonCampRecup.EOF And fRsPerdonCampRecup.BOF) Then
                If fRsPerdonCampRecup.RecordCount > 0 Then
                    Set oDCred = New COMDCredito.DCOMCredito
                    If CInt(fRsPerdonCampRecup!nEstado) = 1 Then
                        Call oDCred.CampanasRecupPagoConPerdon(CLng(fRsPerdonCampRecup!nId), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gnMovNro, 2)
                    ElseIf CInt(fRsPerdonCampRecup!nEstado) = 0 Then
                        Call oDCred.CampanasRecupPagoConPerdon(CLng(fRsPerdonCampRecup!nId), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gnMovNro, 3)
                    End If
                End If
            End If
        End If
    End If
    'WIOR FIN *******
    
    'CONTROLAR TIEMPO DE PAGO
'    Call oSeguridad.InsertarPista("654321", GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, 8, "PAGO DE CUOTA", ActxCta.NroCuenta)
'    Set oSeguridad = Nothing
    
    'JACA 20110510***********************************************************
                
        'Dim objPersona As COMDPersona.DCOMPersonas
        Dim rsPersOcu As Recordset
        Dim nAcumulado As Currency
        Dim nMontoPersOcupacion As Currency
        Dim clsMov As COMNContabilidad.NCOMContFunciones
        Dim sMovNro As String
        
        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Dim clsTC As COMDConstSistema.NCOMTipoCambio
        Set clsTC = New COMDConstSistema.NCOMTipoCambio
        nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
        Set clsTC = Nothing
        
        Set rsPersOcu = New Recordset
        'Set objPersona = New COMDPersona.DCOMPersonas
                        
        Set rsPersOcu = objPersona.ObtenerDatosPersona(sPersCod)
        nAcumulado = objPersona.ObtenerPersAcumuladoMontoOpe(nTC, Mid(Format(gdFecSis, "yyyymmdd"), 1, 6), rsPersOcu!cPersCod)
        nMontoPersOcupacion = objPersona.ObtenerParamPersAgeOcupacionMonto(Mid(rsPersOcu!cPersCod, 4, 2), CInt(Mid(rsPersOcu!cPersCIIU, 2, 2)))
    
        If nAcumulado >= nMontoPersOcupacion Then
            If Not objPersona.ObtenerPersonaAgeOcupDatos_Verificar(rsPersOcu!cPersCod, gdFecSis) Then
                objPersona.insertarPersonaAgeOcupacionDatos gnMovNro, rsPersOcu!cPersCod, IIf(nmoneda = 1, nMonto, nMonto * nTC), nAcumulado, gdFecSis, sMovNro
            End If
        End If
       

    'JACA END*****************************************************************
    
    '** Juez 20120323 ************************************
    If regPersonaRealizaPago And gnMovNro > 0 Then 'WIOR 20130301 AGREGO gnMovNro>0
        'WIOR 20130301 COMENTO
        'frmCredPagoCuotasPersRealiza.insertaPersonaRealizaPagoCuotas gnMovNro, frmCredPagoCuotasPersRealiza.PersCod, frmCredPagoCuotasPersRealiza.PersDNI, frmCredPagoCuotasPersRealiza.PersNombre
        'WIOR 20130301 **************************************************
        frmPersRealizaOpeGeneral.InsertaPersonasOperacion gnMovNro, Trim(ActxCta.NroCuenta), fnCondicion
        regPersonaRealizaPago = False
        'WIOR FIN ***************************************************
    End If
    '** End Juez *****************************************
	
       'INICIO JIPR 20180625
     Set oCredCodFactElect = New COMDCredito.DCOMCredito
     Call oCredCodFactElect.VerificaInteresFactElect(gnMovNro)
     'ANPS COMENTADO 20210820
'     If Not (R.BOF And R.EOF) Then
'     If R!exist = True Then
'     Set R = oCredCodFactElect.RecuperaFactElectDetalle(ActxCta.NroCuenta)
'        If R.RecordCount > 0 Then
'            If Not (R.BOF And R.EOF) Then
'         scPersCod = R!cPersCod
'         scPersIDTpo = R!cPersIDTpo
'         Call oCredCodFactElect.InsertaFactElect(gnMovNro, ActxCta.NroCuenta, scPersCod, scPersIDTpo, gdFecSis)
'
'         Set R = oCredCodFactElect.GenerarCorrelativo(scPersIDTpo, gsCodAge, ActxCta.NroCuenta)  'ANPS 16022021 add codigo de cuenta
'                If R.RecordCount > 0 Then
'                    If Not (R.BOF And R.EOF) Then
'         sSerie = R!cSerie
'         sCorrelativo = R!cNro
'
'         Call oCredCodFactElect.UpdateFactElect(gnMovNro, sSerie, sCorrelativo)
'         Call oCredCodFactElect.InsertarRegVentaFactElect(gdFecSis, scPersCod, ActxCta.NroCuenta, gnMovNro, IIf(nmoneda = 1, 1, 2), gsCodAge)
'
            'MsgBox "Coordinar con el Supervisor de Operaciones, para la Emisión de Facturación Electrónica del Pago de Crédito.", vbInformation, "Aviso" COMENTADO ANPS 20210531
          
'                    End If
'                End If
         Set oCredCodFactElect = Nothing
          
'            End If
'        End If
'     End If
'     End If
     Set oCredCodFactElect = Nothing
     'FIN JIPR 20180625
    
    Set vPrevio = New clsprevio
    If sImprePlanPago <> "" Then
        vPrevio.PrintSpool sLpt, sImprePlanPago
        'vPrevio.Show sImprePlanPago, "Impresión de Plan de Pagos"
    End If
    
    vPrevio.PrintSpool sLpt, sImpreBoleta
    'vPrevio.Show sImpreBoleta, "Impresión de Plan de Pagos"
    Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
        vPrevio.PrintSpool sLpt, sImpreBoleta
        'vPrevio.Show sImpreBoleta, "Impresión de Plan de Pagos"
    Loop
    Set vPrevio = Nothing
    
    gVarPublicas.LimpiaVarLavDinero 'DAOR 20070511
    
    Call cmdCancelar_Click
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", sOperacion
    'FIN
    '**DAOR 20080415, ***************************************************
    MDISicmact.SBBarra.Panels(1).Text = lsTemp
    '********************************************************************
    
    
Exit Sub

'        If nMiVivienda = 1 Then
'            'Salvo porque lo voy a utilizar luego para imprimir las boletas
'            MatCalendDistribuidoTempo = MatCalendDistribuido
'            MatCalend_2 = MatCalend
'            MatCalendDistribuido_2 = oNegCred.CrearMatrizparaAmortizacion(MatCalend_2)
'            MatCalendDistribuidoParalelo = oNegCred.CrearMatrizparaAmortizacion(MatCalendParalelo)
'            Call oNegCred.DistribuirMatrizMiVivEnDosCalendarios(MatCalendDistribuidoParalelo, MatCalendDistribuido_2, MatCalendDistribuido, MatCalendParalelo, MatCalendNormalT1, nCalPago)
'            MatCalendDistribuido = MatCalendDistribuido_2
'
'        End If
'
''        If nMontoPago <> CDbl(LblTotDeuda.Caption) Then
''            nInteresDesagio = 0
''        End If
'
'        sError = oNegCred.AmortizarCredito(ActxCta.NroCuenta, MatCalend, MatCalendDistribuido, _
'                nMontoPago, gdFecSis, lblMetLiq.Caption, CInt(Trim(Right(CmbForPag.Text, 10))), _
'                gsCodAge, gsCodUser, Trim(LblNumDoc.Caption), , , , bRecepcionCmact, sPersCmac, _
'                vnIntPendiente, vnIntPendientePagado, , MatGastosFinal, nNumGastosFinal, MatCalendDistribuidoParalelo, _
'                nCalPago, MatCalendParalelo, bPrepago, sPersLavDinero, nITF, nInteresDesagio)
'
'        If sError <> "" Then
'            MsgBox sError, vbInformation, "Aviso"
'        Else
'            'Verifica si fue un pago para Calendario Dinamico
'            'bPrepago nCalendDinamTipo
'            If (bCalenDinamic Or bPrepago = 1) And (nMontoPago < CDbl(LblTotDeuda.Caption)) Then
'                If nMontoPago > CDbl(LblMonCalDin.Caption) Then
'                    If nMiVivienda = 1 Then
'                        MatCalDinam = oNegCred.ReprogramarCreditoenMemoriaTotalMiVivienda(ActxCta.NroCuenta, gdFecSis, MatCalDinam_2, IIf(nCalendDinamTipo = 1, True, False))
'                        'Reporgramacion 2 de otorgar un nuevo calendario en basae al saldo de capital pendiente
'                        'Como si fueera un nuevo credito bajo las cuotas pendientes
'                        oNegCred.ReprogramarCredito ActxCta.NroCuenta, MatCalDinam, 2, True, MatCalDinam_2, gdFecSis, , gsCodUser, gsCodAge
'                        Call oNegCred.ActualizarCalificacionMIVivienda(ActxCta.NroCuenta)
'                        Set oDoc = New COMNCredito.NCOMCredDoc
'                        sCad = oDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, oNegCred.MatrizCapitalCalendario(MatCalDinam) + oNegCred.MatrizCapitalCalendario(MatCalDinam_2), True)
'                        sCad2 = oDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, oNegCred.MatrizCapitalCalendario(MatCalDinam) + oNegCred.MatrizCapitalCalendario(MatCalDinam_2), True, True)
'                        Set vPrevio = New clsPrevio
'                        vPrevio.PrintSpool sLpt, sCad & sCad2
'                        Set vPrevio = Nothing
'                        Set oDoc = Nothing
'                    Else
'                        MatCalDinam = oNegCred.ReprogramarCreditoenMemoriaTotal(ActxCta.NroCuenta, gdFecSis)
'                        'Reporgramacion 2 de otorgar un nuevo calendario en basae al saldo de capital pendiente
'                        'Como si fueera un nuevo credito bajo las cuotas pendientes
'                        oNegCred.ReprogramarCredito ActxCta.NroCuenta, MatCalDinam, 2, , , gdFecSis, , gsCodUser, gsCodAge
'                        Set oDoc = New COMNCredito.NCOMCredDoc
'                        sCad = oDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, oNegCred.MatrizCapitalCalendario(MatCalDinam), False)
'                        Set vPrevio = New clsPrevio
'                        vPrevio.ShowImpreSpool sCad
'                        Set vPrevio = Nothing
'                        Set oDoc = Nothing
'                    End If
'
'
'                End If
'            End If
'
'            Set oConstante = New COMDConstantes.DCOMConstantes
'            sTipoCred = oConstante.DameDescripcionConstante(gProducto, CInt(ActxCta.Prod))
'            Set oConstante = Nothing
'            Set oDoc = New COMNCredito.NCOMCredDoc
'
'            If nMiVivienda = 1 Then
'                'Recupero para imprimir las boletas
'                MatCalendDistribuido = MatCalendDistribuidoTempo
'            End If
'
'            Set oCal = New COMDCredito.DCOMCalendario
'
'            Call oDoc.ImprimeBoleta(ActxCta.NroCuenta, LblNomCli.Caption, gsNomAge, LblMoneda, _
'                oNegCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(FechaHora(gdFecSis), "hh:mm:ss"), nNroTransac + 1, Mid(sTipoCred, 1, 18), _
'                oNegCred.MatrizCapitalPagado(MatCalendDistribuido), oNegCred.MatrizIntCompPagado(MatCalendDistribuido), _
'                oNegCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
'                oNegCred.MatrizIntMorPagado(MatCalendDistribuido), oNegCred.MatrizGastoPag(MatCalendDistribuido), _
'                oNegCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
'                oNegCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNegCred.MatrizIntReprogPag(MatCalendDistribuido), _
'                oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), LblProxfec.Caption, _
'                gsCodUser, sLpt, gsInstCmac, IIf(Trim(Right(Me.CmbForPag.Text, 2)) = "2", True, False), Me.LblNumDoc.Caption, gsCodCMAC, nITF, nInteresDesagio, bRecepcionCmact)
'
'            Do While MsgBox("Desea Reimprimir el Comprobante de Pago?", vbInformation + vbYesNo, "Aviso") = vbYes
'                Call oDoc.ImprimeBoleta(ActxCta.NroCuenta, LblNomCli.Caption, gsNomAge, LblMoneda, _
'                oNegCred.MatrizCuotasPagadas(MatCalendDistribuido), gdFecSis, Format(FechaHora(gdFecSis), "hh:mm:ss"), nNroTransac + 1, Mid(sTipoCred, 1, 18), _
'                oNegCred.MatrizCapitalPagado(MatCalendDistribuido), oNegCred.MatrizIntCompPagado(MatCalendDistribuido), _
'                oNegCred.MatrizIntCompVencPagado(MatCalendDistribuido), _
'                oNegCred.MatrizIntMorPagado(MatCalendDistribuido), oNegCred.MatrizGastoPag(MatCalendDistribuido), _
'                oNegCred.MatrizIntGraciaPagado(MatCalendDistribuido), _
'                oNegCred.MatrizIntSuspensoPag(MatCalendDistribuido) + oNegCred.MatrizIntReprogPag(MatCalendDistribuido), _
'                oNegCred.MatrizSaldoCapital(MatCalend, MatCalendDistribuido), LblProxfec.Caption, _
'                gsCodUser, sLpt, gsInstCmac, IIf(Trim(Right(Me.CmbForPag.Text, 2)) = "2", True, False), Me.LblNumDoc.Caption, gsCodCMAC, nITF, nInteresDesagio, bRecepcionCmact)
'            Loop
'            Set oDoc = Nothing
'
'            Set oCal = Nothing
'
''            ''''''''''''''''''''
''            Dim clsImp As NCapImpBoleta
''            Dim oCta As dCredito
''            Dim oPersona As DCapMantenimiento
''            Dim sBoletaLavDinero As String
''            Dim sPersCod As String
''            Dim sNombre As String
''            Dim sDireccion As String
''            Dim sDocId As String
''            Dim rsPers As New ADODB.Recordset
''
''            Set oCta = New dCredito
''            sPersCod = oCta.RecuperaTitularCredito(ActxCta.NroCuenta)
''            Set oCta = Nothing
''
''            Set oPersona = New DCapMantenimiento
''
''            Set rsPers = oPersona.GetDatosPersona(sPersCod)
''            If rsPers.BOF Then
''            Else
''
''                sPersCod = sPersCod
''                sNombre = rsPers!Nombre
''                sDireccion = rsPers!Direccion
''                sDocId = rsPers!id & " " & rsPers![ID N°]
''            End If
''
''            If sPersLavDinero <> "" Then
''
''                Do
''                    Set clsImp = New NCapImpBoleta
''                    If sBoletaLavDinero <> "" Then sBoletaLavDinero = sBoletaLavDinero & Chr$(12)
''                    sBoletaLavDinero = sBoletaLavDinero & clsImp.ImprimeBoletaLavadoDinero(gsNomCmac, gsNomAge, gdFecSis, ActxCta.NroCuenta, sNombre, sDocId, sDireccion, _
''                                 sNombre, sDocId, sDireccion, sNombre, sDocId, sDireccion, sOperacion, CDbl(TxtMonPag.Text), sLpt, , False, "COLOCACIONES")
''
''                    Set clsImp = Nothing
''                Loop While MsgBox("Desea Reimprimir Boleta de lavado de dinero?", vbInformation + vbYesNo, "Aviso") = vbYes
''            End If
''            '''''''''''''''''''
'
'            Call cmdcancelar_Click
'        End If
'        Set oNegCred = Nothing
'    End If

'    Exit Sub

ErrorCmdGrabar_Click:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdmora_Click()
'    Call TxtMonPag_KeyPress(13)
    Call frmCredMoraCuotas.MostarMoraDetalle(Me.ActxCta.NroCuenta, gdFecSis)
End Sub

Private Sub CmdPlanPagos_Click()

'Dim odCred As DCredito
'Dim oCredDoc As NCredDoc
'Dim sCadImp As String
'Dim sCadImp_2 As String
'Dim Prev As Previo.clsPrevio
'
    On Error GoTo ErrorCmdPlanPagos_Click
'            Set oCredDoc = New NCredDoc
'            Set Prev = New clsPrevio
'            sCadImp = oCredDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, , gsNomCmac, bCuotaCom, nCalendDinamico)
'            sCadImp_2 = ""
'            If nMiVivienda Then
'                sCadImp_2 = oCredDoc.ImprimePlandePagos(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, nPrestamo, nMiVivienda, True, gsNomCmac, bCuotaCom, nCalendDinamico)
'            End If
'            Prev.Show sCadImp & sCadImp_2, "", False
'            Set Prev = Nothing
'            Set oCredDoc = Nothing
    
    Call frmCredHistCalendario.PagoCuotas(ActxCta.NroCuenta)

    Exit Sub

ErrorCmdPlanPagos_Click:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Inicia(sCodOpe As String)
    bRecepcionCmact = False
    sOperacion = sCodOpe
    Me.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sNumTar As String
    Dim sClaveTar As String
    Dim nErr As Integer
    Dim sCaption As String
    'Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim nEstado  As CaptacTarjetaEstado
    'Set clsGen = New COMDConstSistema.DCOMGeneral
    
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
    
'    If KeyCode = vbKeyF11 And txtCuenta.Enabled = True Then 'F11
'        Dim nPuerto As TipoPuertoSerial
'        nPuerto = clsGen.GetPuertoPeriferico(gPerifPINPAD)
'        If nPuerto < 0 Then nPuerto = gPuertoSerialCOM1
'        IniciaPinPad nPuerto
'
'        WriteToLcd "Pase su Tarjeta por la Lectora."
'        sCaption = Me.Caption
'        Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
'        sNumTar = GetNumTarjeta
'        sNumTar = Trim(Replace(sNumTar, "-", "", 1, , vbTextCompare))
'        If Len(sNumTar) <> 16 Then
'            MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
'            WriteToLcd "                                       "
'            WriteToLcd "Gracias por su  Preferencia..."
'            FinalizaPinPad
'            Me.Caption = sCaption
'            Exit Sub
'        End If
'
'        Me.Caption = "Ingrese la Clave de la Tarjeta."
'        WriteToLcd "                                       "
'        WriteToLcd "Ingrese Clave"
'        sClaveTar = GetClaveTarjeta
'        If clsGen.ValidaClaveTarjeta(sNumTar, sClaveTar) Then
'            Dim clsMant As NCapMantenimiento
'            Dim rsTarj As Recordset
'
'            Set clsMant = New NCapMantenimiento
'            Set rsTarj = clsMant.GetTarjetaCuentas(sNumTar)
'            If rsTarj.EOF And rsTarj.BOF Then
'                MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
'                WriteToLcd "                                       "
'                WriteToLcd "Gracias por su  Preferencia..."
'                FinalizaPinPad
'                Me.Caption = sCaption
'                Exit Sub
'            Else
'                nEstado = rsTarj("nEstado")
'                If nEstado = gCapTarjEstBloqueada Or nEstado = gCapTarjEstCancelada Then
'                    If nEstado = gCapTarjEstBloqueada Then
'                        MsgBox "Número de Tarjeta Bloqueada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
'                    ElseIf nEstado = gCapTarjEstCancelada Then
'                        MsgBox "Número de Tarjeta Cancelada, consulte con el Administrador de la Agencia.", vbInformation, "Aviso"
'                    End If
'                    WriteToLcd "                                       "
'                    WriteToLcd "Gracias por su  Preferencia..."
'                    FinalizaPinPad
'                    Me.Caption = sCaption
'                    Exit Sub
'                End If
'
'                Dim rsPers As Recordset
'                Dim sCta As String, sRelac As String, sEstado As String
'                Dim clsCuenta As UCapCuentas
'
'                Set rsPers = clsMant.GetCuentasPersona(rsTarj("cPersCod"), nProducto)
'                Set clsMant = Nothing
'                If Not (rsPers.EOF And rsPers.EOF) Then
'                    Do While Not rsPers.EOF
'                        sCta = rsPers("cCtaCod")
'                        sRelac = rsPers("cRelacion")
'                        sEstado = Trim(rsPers("cEstado"))
'                        frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
'                        rsPers.MoveNext
'                    Loop
'                    Set clsCuenta = New UCapCuentas
'                    Set clsCuenta = frmCapMantenimientoCtas.Inicia
'                    If clsCuenta.sCtaCod <> "" Then
'                        txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
'                        txtCuenta.Prod = Mid(clsCuenta.sCtaCod, 6, 3)
'                        txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
'                        txtCuenta.SetFocusCuenta
'                        SendKeys "{Enter}"
'                    End If
'                    Set clsCuenta = Nothing
'                Else
'                    MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
'                End If
'                rsPers.Close
'                Set rsPers = Nothing
'            End If
'            Set rsTarj = Nothing
'            Set clsMant = Nothing
'
'        Else
'            WriteToLcd "                                       "
'            WriteToLcd "Clave Incorrecta"
'            MsgBox "Clave Incorrecta", vbInformation, "Aviso"
'        End If
'        Set clsGen = Nothing
'        WriteToLcd "                                       "
'        WriteToLcd "Gracias por su  Preferencia..."
'        FinalizaPinPad
'    End If
End Sub


Private Sub Form_Load()
    Call CargaControles
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    bCalenDinamic = False
    CentraSdi Me
    bantxtmonpag = False
    'ventana = 0
    Set oDocRec = New UDocRec 'EJVG20140408
    chkPreParaAmpliacion.Enabled = False 'ALPA 20150326
    
    'WIOR 20151220 ***
    fbEsCredMIVIVIENDA = False
    fbMIVIVIENDAAnt = False
    'WIOR FIN ********
    bValidaActualizacionLiq = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set oDocRec = New UDocRec 'EJVG20140408
End Sub

Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub TxtMonPag_Change()

    If Not bantxtmonpag Then
        'ARCV
        'MatCalendDistribuido = CrearMatrizparaAmortizacion(MatCalend)
        bActualizaMontoPago = True
        LblNewSalCap.Caption = ""
        LblNewCPend.Caption = ""
        LblProxfec.Caption = ""
        LblEstado.Caption = ""
        lblTipoPago.Caption = "" 'JUEZ 20150415
        'LblItf.Caption = "0.00"
        cmdGrabar.Enabled = False
    End If
    
End Sub

'Para evitar conectarse a un COMPONENTE solo para formar la Matriz

Private Function CrearMatrizparaAmortizacion(ByVal MatCalend As Variant) As Variant
Dim MatCalendAmortiz() As String
Dim i As Integer
    ReDim MatCalendAmortiz(UBound(MatCalend), 12)

    For i = 0 To UBound(MatCalend) - 1
        MatCalendAmortiz(i, 0) = MatCalend(i, 0)
        MatCalendAmortiz(i, 1) = MatCalend(i, 1)
        MatCalendAmortiz(i, 2) = MatCalend(i, 2)
        MatCalendAmortiz(i, 3) = "0.00"
        MatCalendAmortiz(i, 4) = "0.00"
        MatCalendAmortiz(i, 5) = "0.00"
        MatCalendAmortiz(i, 6) = "0.00"
        MatCalendAmortiz(i, 7) = "0.00"
        MatCalendAmortiz(i, 8) = "0.00"
        MatCalendAmortiz(i, 9) = "0.00"
        MatCalendAmortiz(i, 11) = "0.00"
        MatCalendAmortiz(i, 10) = MatCalend(i, 10)
    Next i
    CrearMatrizparaAmortizacion = MatCalendAmortiz
End Function


Private Sub TxtMonPag_GotFocus()
    fEnfoque TxtMonPag
End Sub

Private Sub TxtMonPag_KeyPress(KeyAscii As Integer)

'Dim oNegCredito As COMNCredito.NCOMCredito
'Dim oGastos As COMNCredito.NCOMGasto
'Dim nMontoGastoGen As Double
'Dim oDCredito As COMDCredito.DCOMCredito
'Dim nInteresFecha As Double

'Dim oCredito As COMNCredito.NCOMCredito
Dim bValorProceso As Boolean
Dim sMensaje As String
Dim nMonIntGra As Double
Dim nNewSalCap As Double
Dim nNewCPend As Integer
Dim dProxFec As Date
Dim sEstado As String

    KeyAscii = NumerosDecimales(TxtMonPag, KeyAscii, 15)
    
    '** Juez 20120523 ************************
    If KeyAscii = 13 And Me.chkCancelarCred.value = 1 Then
            TxtMonPag.Text = LblTotDeuda.Caption
        'Exit Sub
    End If
    '** End Juez *****************************
    
    If KeyAscii <> 13 Then Exit Sub
    
    
    '23-06
    
    If bActualizaMontoPago = False Then
        bActualizaMontoPago = True
        If cmdGrabar.Enabled Then
            cmdGrabar.SetFocus
        End If
        Exit Sub
    End If
    '--------
    
    
    '09-06-2006
    
    If Not IsNumeric(TxtMonPag.Text) Then
        MsgBox "Ingrese un monto válido", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    'WIOR 20150404 **********************
    If fnHayPerdonCamp > 0 Then
        If CDbl(TxtMonPag.Text) < Round(fnMontoMinPerdon, 2) Then
            MsgBox "El monto del pago no puede ser menor a " & Round(fnMontoMinPerdon, 2) & ", ya que realiza el Perdon de Mora Campaña de Recuperaciones.", vbInformation, "Mensaje"
            TxtMonPag.Text = Round(fnMontoMinPerdon, 2)
            Exit Sub
        End If
        
        If CDbl(TxtMonPag.Text) > Round(fnMontoMaxPerdon, 2) Then
            MsgBox "El monto del pago no puede ser mayor a " & Round(fnMontoMaxPerdon, 2) & ", ya que realiza el Perdon de Mora Campaña de Recuperaciones.", vbInformation, "Mensaje"
            TxtMonPag.Text = Round(fnMontoMaxPerdon, 2)
            Exit Sub
        End If
    End If
    'WIOR FIN ***************************
    
    '-------------------------------
    'Set oCredito = New COMNCredito.NCOMCredito
    'nNumGastosFinal, MatGastosFinal (23-03-06)
    'ALPA20130805 Se agregó el campo nMiVivienda
    'bValorProceso = oCredito.ActualizaMontoPago(CDbl(TxtMonPag.Text), CDbl(LblTotDeuda.Caption), ActxCta.NroCuenta, gdFecSis, LblMetLiq.Caption, vnIntPendiente, vnIntPendientePagado, _
                                        bCalenCuotaLibre, bCalenDinamic, bPrepago, nMontoPago, CDbl(LblMonCalDin.Caption), sMensaje, nitf, MatCalend, _
                                        nInteresDesagio, MatCalendDistribuido, MatCalendTmp, nNewSalCap, nNewCPend, dProxFec, sEstado)
    Dim bExcluyeGastos As Boolean
    bExcluyeGastos = True
    bValorProceso = oCredito.ActualizaMontoPago(CDbl(TxtMonPag.Text), CDbl(LblTotDeuda.Caption), ActxCta.NroCuenta, gdFecSis, lblMetLiq.Caption, vnIntPendiente, vnIntPendientePagado, _
                                        bCalenCuotaLibre, bCalenDinamic, bPrepago, nMontoPago, CDbl(LblMonCalDin.Caption), sMensaje, nITF, _
                                        nInteresDesagio, nNewSalCap, nNewCPend, dProxFec, sEstado, nMonIntGra, , , , IIf(fbMIVIVIENDAAnt = True, 1, 0), , lnMontoPendienteIntGracia, IIf(chkPreParaAmpliacion.value, 1, 0), lnMontIntComp, lnMontGasto, , nAmpl, _
                                        bExcluyeGastos)
    'pbExcluyeGastos CTI2 20190306
    'Set oCredito = Nothing
    'ALPA20130815
    'WIOR 20160107 SE CAMBIO nMiVivienda POR IIf(fbMIVIVIENDAAnt = True, 1, 0)
    
    'INICIO EAAS20180516
    If (nAmpl = 1) Then
        MsgBox "Se realizó la preparación de la ampliación. Por favor proceder con la cancelación de gastos de los demás créditos a ampliar. De no haber más créditos, proceder con el desembolso.", vbInformation, "Mensaje"
        nAmpl = 0
        Call cmdCancelar_Click
    Exit Sub
    End If
    'FIN EAAS20180516
    
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
    End If
    If bValorProceso = False Then
    cmdGrabar.Enabled = False 'INICIO EAAS20180516
    Exit Sub
    End If
       
    'JUEZ 20150415 ***********************************************************************
    If lsTpoCredCod = "853" Or lsTpoCredCod = "854" Then
        '** Juez 20120528 **********************************************
        If Me.chkCancelarCred.value = 0 And nCalendDinamico = 0 And bPrepago = 0 Then
            If ValidaPagoAnticipado Then
'comentado por CTI3: ers085-2018
'                If Not bCuotasVencidas Then
'                    MsgBox "El monto a pagar excede lo permitido. No puede pagar mas de 2 cuotas.", vbExclamation, "Mensaje"
'                Else
'                    MsgBox "El monto adicionalmente agregado excede lo permitido. No puede pagar mas de 2 cuotas Por Vencer.", vbExclamation, "Mensaje"
'                End If
'                Exit Sub
            End If
        End If
        '** End Juez ***************************************************
        'JUEZ 20141212 *****************************************************
        If nCalendDinamico = 1 Or bPrepago = 1 Then
            If Round((Trim(TxtMonPag.Text)), 2) < Round(nMontoPag2CuotxVenc, 2) Then
                MsgBox "El monto a pagar debe ser mas de 2 cuotas Por Vencer, ya que el crédito fue configurado como prepago y calendario dinámico", vbExclamation, "Mensaje"
                Exit Sub
            End If
        End If
        'END JUEZ **********************************************************
    End If
    'END JUEZ ****************************************************************************
    
    'JUEZ 20130925 *****************************************************
    Dim oAmpliacion As New COMDCredito.DCOMAmpliacion
    If Not oAmpliacion.ValidaCreditoaAmpliar(ActxCta.NroCuenta) Then
        If fnPersPersoneria <> 1 And (Left(lsTpoCredCod, 1) = "1" Or Left(lsTpoCredCod, 1) = "2" Or Left(lsTpoCredCod, 1) = "3") Then 'Comisión sólo para Pers Jur y Tipo Cred Corporativa, Grande o Mediana Empresa
            Dim oDCred As New COMDCredito.DCOMCredito
            If sEstado = "CANCELADO" And oDCred.VerificaSiEsCancelacionAnticipada(ActxCta.NroCuenta, gdFecSis) Then
                If oDCred.ExisteExoneracionPreCancelacion(ActxCta.NroCuenta, gdFecSis) Then
                    MsgBox "El crédito está Exanorado de comisión por PreCancelación", vbInformation, "Aviso"
                    lblMontoPreCanc.Caption = "0.00"
                    HabControlesPreCanc True, False
                Else
                    lblMontoPreCanc.Caption = Format(CalculaComisionPreCancelacion(CDbl(LblTotDeuda.Caption), ActxCta.NroCuenta), "#,##0.00")
                    HabControlesPreCanc False, True
                End If
            Else
                lblMontoPreCanc.Caption = "0.00"
                HabControlesPreCanc True, False
            End If
        Else
            lblMontoPreCanc.Caption = "0.00"
            HabControlesPreCanc True, False
        End If
    Else
        lblMontoPreCanc.Caption = "0.00"
        HabControlesPreCanc True, False
    End If
    Set oAmpliacion = Nothing
    'END JUEZ **************************************************************
    
'    If CDbl(TxtMonPag.Text) = 0 Then
'        MsgBox "Monto de Pago Debe ser mayor que Cero", vbQuestion, "Aviso"
'        Exit Sub
'    End If
'    If CDbl(TxtMonPag.Text) > CDbl(LblTotDeuda.Caption) Then
'        Set oDCredito = New COMDCredito.DCOMCredito
'        If oDCredito.ObtieneMonto_Validate(ActxCta.NroCuenta, TxtMonPag) = True Then
'            MsgBox "Monto de Pago es mayor que la deuda", vbQuestion, "Aviso"
'        Else
'            MsgBox "Monto de pago sobrepasa el total", vbInformation, "Aviso"
'        Exit Sub
'    End If
'        End If
'
'        TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
'
'        nMontoPago = fgITFCalculaImpuestoNOIncluido(CDbl(TxtMonPag.Text))
'        'nITF = Format(nMontoPago - CDbl(TxtMonPag.Text), "0.00")
'        'nITF = CalculoSinRedondeo(nMontoPago - CDbl(TxtMonPag.Text))
'        If Mid(ActxCta.NroCuenta, 6, 3) = "423" Then
'            nITF = 0
'        Else
'            nITF = CalculoSinRedondeo(CDbl(TxtMonPag.Text))
'        End If
'        'LblItf.Caption = Format(nITF, "0.00")
'        LblItf.Caption = Format(nITF, "#0.00") 'CalculoSinRedondeo(nITF)
'        lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + nITF, "#0.00")
'        nMontoPago = Val(TxtMonPag.Text)
'
'        'Si el monto es igual a la deuda a la fecha adicionar le desagio
'        'If nMontoPago = CDbl(LblTotDeuda.Caption) Then
'            'nMontoPago = nMontoPago + nInteresDesagio
'        'End If
'
'        Set oNegCredito = New COMNCredito.NCOMCredito
'        nInteresFecha = oNegCredito.MatrizInteresGastosAFecha(ActxCta.NroCuenta, MatCalend, gdFecSis, True)
'        nInteresDesagio = 0
'        If nInteresFecha < 0 Then
'            nInteresDesagio = Abs(nInteresFecha)
'        End If
'
'        If bCalenCuotaLibre Then
'            nInteresDesagio = 0
'            MatCalendDistribuido = oNegCredito.MatrizDistribuirCalendCuotaLibre(ActxCta.NroCuenta, MatCalend, nMontoPago, Trim(lblMetLiq.Caption), gdFecSis, vnIntPendiente, vnIntPendientePagado)
'        End If
'        If (bCalenDinamic Or bPrepago = 1) And (nMontoPago < CDbl(LblTotDeuda.Caption)) Then
'            nInteresDesagio = 0
'            If nMontoPago > CDbl(LblMonCalDin.Caption) Then
'                'Genera Gastos por PrePago
'                Set oGastos = New COMNCredito.NCOMGasto
'                MatGastosFinal = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosFinal, gdFecSis, ActxCta.NroCuenta, 1, "PP", , , nMontoPago, nMontoPago - CDbl(LblMonCalDin.Caption), oNegCredito.MatrizCuotaPendiente(MatCalend, MatCalendDistribuido))
'                'obtener el total del gastos MatGastosPrepago
'                nMontoGastoGen = MontoTotalGastosGenerado(MatGastosFinal, nNumGastosFinal, Array("PP", "PA", ""))
'                MatCalend = MatCalendTmp
'                MatCalend(0, 9) = Format(CDbl(MatCalend(0, 9)) + nMontoGastoGen, "#0.00")
'                Set oGastos = Nothing
'
'                'Distribuye Monto
'                MatCalendDistribuido = oNegCredito.MatrizDistribuirCalendDinamico(ActxCta.NroCuenta, MatCalend, nMontoPago, Trim(lblMetLiq.Caption), gdFecSis)
'            Else
'                Set oGastos = New COMNCredito.NCOMGasto
'                MatGastosFinal = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosFinal, gdFecSis, ActxCta.NroCuenta, 1, "PA", , , nMontoPago, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), oNegCredito.MatrizCuotaPendiente(MatCalend, MatCalendDistribuido))
'                'obtener el total del gastos MatGastosPrepago
'                nMontoGastoGen = MontoTotalGastosGenerado(MatGastosFinal, nNumGastosFinal, Array("PA", "", ""))
'                MatCalend = MatCalendTmp
'                MatCalend(0, 9) = Format(CDbl(MatCalend(0, 9)) + nMontoGastoGen, "#0.00")
'                Set oGastos = Nothing
'
'                'Distribuye Monto
'                MatCalendDistribuido = oNegCredito.MatrizDistribuirMonto(MatCalend, nMontoPago, Trim(lblMetLiq.Caption))
'            End If
'        Else
'            'Si es Pago Normal
'            If nMontoPago <> CDbl(LblTotDeuda.Caption) Then
'                nInteresDesagio = 0
'                Set oGastos = New COMNCredito.NCOMGasto
'                MatGastosFinal = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosFinal, gdFecSis, ActxCta.NroCuenta, 1, "PA", , , nMontoPago, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), oNegCredito.MatrizCuotaPendiente(MatCalend, MatCalendDistribuido))
'                'obtener el total del gastos
'                nMontoGastoGen = MontoTotalGastosGenerado(MatGastosFinal, nNumGastosFinal, Array("PA", "", ""))
'                MatCalend = MatCalendTmp
'                MatCalend(0, 9) = Format(CDbl(MatCalend(0, 9)) + nMontoGastoGen, "#0.00")
'                Set oGastos = Nothing
'
'                'Distribuye Monto
'
'                If Mid(Trim(lblMetLiq.Caption), 3, 1) = "i" Or Mid(Trim(lblMetLiq.Caption), 3, 1) = "Y" Then
'                   If CDbl(TxtMonPag) >= CDbl(LblTotDeuda) Then
'                        MatCalendDistribuido = oNegCredito.MatrizDistribuirCancelacion(ActxCta.NroCuenta, MatCalend, nMontoPago, Trim(lblMetLiq.Caption), gdFecSis, True)
'                   Else
'                        MatCalendDistribuido = oNegCredito.MatrizDistribuirCancelacion(ActxCta.NroCuenta, MatCalend, nMontoPago, Trim(lblMetLiq.Caption), gdFecSis, False, , False)
'                   End If
'                Else
'                    MatCalendDistribuido = oNegCredito.MatrizDistribuirMonto(MatCalend, nMontoPago, Trim(lblMetLiq.Caption))
'                End If
'            Else 'Si es una Cancelacion del Credito
'                'Obtener los Gastos Generados y Agregarlos al Credito
'                MatGastosFinal = MatGastosCancelacion
'                nNumGastosFinal = nNumGastosCancel
'                nMontoGastoGen = MontoTotalGastosGenerado(MatGastosFinal, nNumGastosFinal, Array("PA", "CA", ""))
'                MatCalend = MatCalendTmp
'                MatCalend(0, 9) = Format(CDbl(MatCalend(0, 9)) + nMontoGastoGen, "#0.00")
'                Set oGastos = Nothing
'
'                'Distribuye Monto
'                MatCalendDistribuido = oNegCredito.MatrizDistribuirCancelacion(ActxCta.NroCuenta, MatCalend, nMontoPago, Trim(lblMetLiq.Caption), gdFecSis, , bCalenDinamic)
'            End If
'        End If
'        LblNewSalCap.Caption = oNegCredito.MatrizSaldoCapital(MatCalend, MatCalendDistribuido)
'        LblNewCPend.Caption = oNegCredito.MatrizCuotaPendiente(MatCalend, MatCalendDistribuido)
'        LblProxfec.Caption = Format(oNegCredito.MatrizFechaCuotaPendiente(MatCalend, MatCalendDistribuido), "dd/mm/yyyy")
'        LblEstado.Caption = IIf(oNegCredito.MatrizEstadoCalendario(MatCalendDistribuido) = gColocCalendEstadoPagado, "CANCELADO", "VIGENTE")
'        If LblEstado.Caption = "CANCELADO" Then
'            LblProxfec.Caption = ""
'        End If
'        Set oNegCredito = Nothing
    bantxtmonpag = True
    TxtMonPag.Text = Format(TxtMonPag.Text, "#0.00")
    
    If bInstFinanc Then nITF = 0 'JUEZ 20140411
    lblITF.Caption = Format(nITF, "#0.00")
    '*** BRGO 20110908 ************************************************
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
    If nRedondeoITF > 0 Then
        Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then lblITF.Caption = "0.00" 'JUEZ 20131227
    '*** END BRGO
    'JUEZ 20130925 *************************************************************************
    'lblPagoTotal.Caption = Format(val(TxtMonPag.Text) + CCur(Me.LblItf.Caption), "#0.00")
    lblPagoTotal.Caption = Format(Val(TxtMonPag.Text) + CCur(Me.lblITF.Caption) + CCur(lblMontoPreCanc.Caption), "#0.00")
    nMontoPago = nMontoPago + CCur(lblMontoPreCanc.Caption)
    'END JUEZ ******************************************************************************
    LblNewSalCap.Caption = nNewSalCap
    LblNewCPend.Caption = nNewCPend
    If dProxFec <> 0 Then LblProxfec.Caption = dProxFec
    LblEstado.Caption = sEstado
    '***Agregado por ELRO el 20120329, según RFC023-2012
    If sEstado = "CANCELADO" Then
        ckbPorAfectacion.Visible = True
    'End If
    '***Fin Agregado por ELRO***************************
    Else '** Juez 20120601
        ckbPorAfectacion.Visible = False
    End If
    
    'JUEZ 20150415 ****************************************
    nTipoAdelantoCuota = 0
    nTipoPagoAnticipado = 0 'JOEP20200705 Cambio ReactivaCovid
    If ValidaAdelantoDeCuota Then
        lblTipoPago.Caption = "Adelanto de Cuota"
        nTipoAdelantoCuota = 1
    ElseIf ValidaPagoAnticipado Then
        lblTipoPago.Caption = "Pago Anticipado"
        nTipoPagoAnticipado = 1 'JOEP20200705 Cambio ReactivaCovid
    Else
        lblTipoPago.Caption = ""
    End If
    'END JUEZ *********************************************
    
      'MARG ERS004-2017
    If ValidaPagoAnticipado = True Then
        LblNewSalCap.Visible = False
        LblNewCPend.Visible = False
        LblProxfec.Visible = False
    Else
        LblNewSalCap.Visible = True
        LblNewCPend.Visible = True
        LblProxfec.Visible = True
    End If
    'END MARG
    
    'CTI3 0212018
    If Me.chkCancelarCred.value = 1 Then MatDatVivienda(0) = lsTpoCredCod: MatDatVivienda(1) = LblEstado.Caption
    If Me.chkCancelarCred.value = 0 Then MatDatVivienda(0) = lsTpoCredCod: MatDatVivienda(1) = lblTipoPago.Caption
    
    'MARG 20180618 Pag. Ant.---------------------------------------------------------
    Dim oDecisionPagAnt As COMDCredito.DCOMCredito
    Dim rsDecisionPagAnt As ADODB.Recordset
    Dim cMensajePagAnt As String
    Dim bMuestraMensajePagAnt As Boolean
    Dim bPermitePagAnt As Boolean
    Dim bAplicaValidacionPagAnt As Boolean
    
    Set oDecisionPagAnt = New COMDCredito.DCOMCredito
    Set rsDecisionPagAnt = oDecisionPagAnt.getDecisionPagoAnticipado(ActxCta.NroCuenta, CInt(LblDiasAtraso.Caption), CBool(chkCancelarCred.value), ValidaPagoAnticipado, CBool(chkPreParaAmpliacion.value))
    If Not rsDecisionPagAnt.BOF And Not rsDecisionPagAnt.EOF Then
        bAplicaValidacionPagAnt = CBool(rsDecisionPagAnt!bAplicaValidacionPagAnt)
        bPermitePagAnt = CBool(rsDecisionPagAnt!bPermitePagAnt)
        bMuestraMensajePagAnt = CBool(rsDecisionPagAnt!bMuestraMensajePagAnt)
        cMensajePagAnt = rsDecisionPagAnt!cMensajePagAnt
        If bAplicaValidacionPagAnt Then
            If bMuestraMensajePagAnt Then
                MsgBox cMensajePagAnt, vbInformation, "AVISO"
            End If
            If Not bPermitePagAnt Then
                rsDecisionPagAnt.Close
                Set rsDecisionPagAnt = Nothing
                Exit Sub
            End If
        End If
    End If
    rsDecisionPagAnt.Close
    Set rsDecisionPagAnt = Nothing
    'END MARG-----------------------------------------------------------------------
    
    bantxtmonpag = False
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
       
'    End If
End Sub
Function CalculoSinRedondeo(ByVal pnMonto As Double) As Double
    Dim sCadena As String
    Dim intpos  As Integer
    Dim nEntera As Integer
    Dim nDecimal As Integer
    Dim lnValor As Double
    
        lnValor = pnMonto * gnITFPorcent
        lnValor = CortaDosITF(lnValor)
        lnValor = Format(lnValor, "#0.00")
        CalculoSinRedondeo = lnValor
       
End Function


Private Sub TxtMonPag_LostFocus()
    If Trim(TxtMonPag.Text) = "" Then
        TxtMonPag.Text = "0.00"
    End If

End Sub

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
'Dim i As Long
 
'Dim oPersona As COMDCaptaGenerales.DCOMCaptaGenerales
'Dim oCta As COMDCredito.DCOMCredito
'Dim rsPers As Recordset
'Dim sPerscod As String
'Dim sNombre As String
'Dim sDireccion As String
'Dim sDocId As String
Dim nMonto As Double
'Set oCta = New COMDCredito.DCOMCredito
'sPerscod = oCta.RecuperaTitularCredito(ActxCta.NroCuenta)
'Set oCta = Nothing

'Set oPersona = New COMDCaptaGenerales.DCOMCaptaGenerales

'Set rsPers = oPersona.GetDatosPersona(sPerscod)
'If rsPers.BOF Then
'Else

'    sPerscod = sPerscod
'    sNombre = rsPers!Nombre
'    sDireccion = rsPers!Direccion
'    sDocId = rsPers!id & " " & rsPers![ID N°]
'End If
'rsPers.Close
'Set rsPers = Nothing

nMonto = CDbl(TxtMonPag.Text)
poLavDinero.TitPersLavDinero = sPersCod
poLavDinero.OrdPersLavDinero = sPersCod

'IniciaLavDinero = frmMovLavDinero.Inicia(sPerscodLav, sNombreLav, sDireccionLav, sDocIdLav, True, True, nMonto, (ActxCta.NroCuenta), sOperacion, True, "PAGO CREDITO", , , , , Mid(ActxCta.NroCuenta, 9, 1))

End Sub

'Private Function EsExoneradaLavadoDinero() As Boolean
'Dim bExito As Boolean
'Dim clsExo As COMNCaptaServicios.NCOMCaptaServicios
'bExito = True
'
'    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
'
'    If Not clsExo.EsPersonaExoneradaLavadoDinero(sPerscod) Then bExito = False
'
'    Set clsExo = Nothing
'    EsExoneradaLavadoDinero = bExito
'
'End Function


'Sub VerificaRFA(ByVal psCtaCod As String)
'    Dim objRFA As COMDCredito.DCOMRFA
'    Set objRFA = New COMDCredito.DCOMRFA
'    bRFA = objRFA.VerificaCreditoRFA(psCtaCod)
'    Set objRFA = Nothing
'End Sub

'JUEZ 20130925 *****************************************************************
Private Sub HabControlesPreCanc(ByVal pbNoPreCanc As Boolean, ByVal pbPreCanc As Boolean)
    Label14.Visible = pbNoPreCanc
    LblNewSalCap.Visible = pbNoPreCanc
    Label16.Visible = pbNoPreCanc
    LblNewCPend.Visible = pbNoPreCanc
    
    lblComPreCanc1.Visible = pbPreCanc
    lblComPreCanc2.Visible = pbPreCanc
    lblMontoPreCanc.Visible = pbPreCanc
End Sub
'END JUEZ **********************************************************************
'JUEZ 20131227 *****************************************************************
Private Function SeleccionarCtaCargo() As Boolean
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim loCuentas As COMDPersona.UCOMProdPersona
Dim rsCuentas As ADODB.Recordset

    SeleccionarCtaCargo = False

    If Trim(sPersCod) <> "" Then
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCuentas = oDCapGen.GetCuentasPersona(sPersCod, gCapAhorros, True, True, Mid(ActxCta.NroCuenta, 9, 1), , , , True, gPrdCtaTpoIndiv)
    Set oDCapGen = Nothing
    End If
    
    If rsCuentas.RecordCount > 0 Then
        Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.inicio(LblNomCli.Caption, rsCuentas)
        If loCuentas.sCtaCod <> "" Then
            SeleccionarCtaCargo = True
            txtCuentaCargo.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            txtCuentaCargo.SetFocusCuenta
        End If
        Set loCuentas = Nothing
    Else
        MsgBox "El cliente no tiene cuentas de ahorro activas", vbInformation, "Aviso"
        SeleccionarCtaCargo = False
        Exit Function
    End If
End Function
'END JUEZ **********************************************************************
'EJVG20140408 ***
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
'JUEZ 20150415 **************************************************************************************
Private Function ValidaAdelantoDeCuota() As Boolean
Dim dFecSis As Date, dFecVencCuotaPag As Date
Dim nMesFecSis As Integer, nMesFecVencCuotaPag As Integer, nMesFecVencCuotaProx As Integer

    ValidaAdelantoDeCuota = False
    dFecSis = CDate(gdFecSis): dFecVencCuotaPag = CDate(LblFecVec.Caption)
    nMesFecSis = Month(dFecSis): nMesFecVencCuotaPag = Month(dFecVencCuotaPag): nMesFecVencCuotaProx = Month(dFecVencCuotaProx)
    nMontoPagoFecha = Round(nMontoPagoFecha, 2)
    nMontoPagoAlDia = Round(nMontoPagoAlDia, 2)
    nMontoPag2CuotxVenc = Round(nMontoPag2CuotxVenc, 2)
    
    If dFecSis <= dFecVencCuotaPag And CDbl(TxtMonPag.Text) > nMontoPagoFecha And CDbl(TxtMonPag.Text) <= nMontoPag2CuotxVenc Then 'Caso 1
        ValidaAdelantoDeCuota = True
    ElseIf dFecSis > dFecVencCuotaPag And dFecSis <= dFecVencCuotaProx And nMesFecSis = nMesFecVencCuotaPag And CDbl(TxtMonPag.Text) > nMontoPagoAlDia And CDbl(TxtMonPag.Text) <= nMontoPag2CuotxVenc + nMontoPagoFecha Then 'Caso 2
        ValidaAdelantoDeCuota = True
    ElseIf dFecSis > dFecVencCuotaPag And dFecSis <= dFecVencCuotaProx And nMesFecSis = nMesFecVencCuotaProx And CDbl(TxtMonPag.Text) > nMontoPagoAlDia And CDbl(TxtMonPag.Text) <= nMontoPag2CuotxVenc + nMontoPagoFecha Then 'Caso 3
        ValidaAdelantoDeCuota = True
    ElseIf dFecSis > dFecVencCuotaPag And CDbl(TxtMonPag.Text) > nMontoPagoAlDia And CDbl(TxtMonPag.Text) <= nMontoPag2CuotxVenc + nMontoPagoFecha Then 'Caso 4 y 5
        ValidaAdelantoDeCuota = True
    Else
        ValidaAdelantoDeCuota = False
    End If
End Function
Private Function ValidaPagoAnticipado() As Boolean

    ValidaPagoAnticipado = False
    
    If Not bCuotasVencidas Then
        If CDbl(TxtMonPag.Text) > nMontoPag2CuotxVenc Then
            ValidaPagoAnticipado = True
        End If
    Else
        If CDbl(TxtMonPag.Text) > nMontoPagoFecha Then
            If (CDbl(TxtMonPag.Text) - nMontoPagoFecha) > nMontoPag2CuotxVenc Then
                ValidaPagoAnticipado = True
            End If
        End If
    End If
End Function
'END JUEZ *******************************************************************************************
'FRHU 20150415 ERS022-2015
Private Function VerificarSiEsUnCreditoTransferido(ByVal psCtaCod As String) As Boolean
    Dim oCredito As COMDCredito.DCOMCredito
    
    Set oCredito = New COMDCredito.DCOMCredito
    VerificarSiEsUnCreditoTransferido = oCredito.VerificaSiEsCreditoTransferido(psCtaCod)
    Set oCredito = Nothing
End Function
'FIN FRHU 20150415

'MARG20180619 Pag. Ant.******************
Private Sub AnularPagoAnticipado()
    Dim oCredPagAnt As COMDCredito.DCOMCredActBD
    Set oCredPagAnt = New COMDCredito.DCOMCredActBD
    If ValidaPagoAnticipado Then
        nCalendDinamico = 0
        bCalenDinamic = False
        Call oCredPagAnt.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , , , , , , , 0)
        oCredPagAnt.dInsertCredMantPrepago ActxCta.NroCuenta, Format(gdFecSis, "yyyymmdd"), , , 1
    End If
    Set oCredPagAnt = Nothing
End Sub
'END MARG********************************
