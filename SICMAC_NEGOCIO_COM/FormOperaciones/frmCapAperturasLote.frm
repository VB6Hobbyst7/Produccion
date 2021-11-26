VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCapAperturasLote 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "frmCapAperturasLote.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDocumento 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2040
      Left            =   80
      TabIndex        =   20
      Top             =   6240
      Width           =   4425
      Begin VB.CommandButton cmdDocumento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Picture         =   "frmCapAperturasLote.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   2280
         TabIndex        =   77
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label lblDocMonto 
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
         Height          =   345
         Left            =   3060
         TabIndex        =   76
         Top             =   1200
         Width           =   1255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nro Ope :"
         Height          =   195
         Left            =   180
         TabIndex        =   75
         Top             =   1275
         Width           =   690
      End
      Begin VB.Label lblDocNroOpe 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   74
         Top             =   1200
         Width           =   735
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
         Height          =   345
         Left            =   960
         TabIndex        =   26
         Top             =   360
         Width           =   1575
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
         Height          =   345
         Left            =   960
         TabIndex        =   25
         Top             =   765
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   375
         Width           =   690
      End
   End
   Begin VB.Frame fraTranferecia 
      Caption         =   "Transferencia"
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
      Height          =   2040
      Left            =   80
      TabIndex        =   57
      Top             =   6270
      Width           =   4425
      Begin VB.TextBox txtTransferGlosa 
         Height          =   315
         Left            =   840
         MaxLength       =   255
         TabIndex        =   60
         Top             =   1290
         Width           =   3465
      End
      Begin VB.CommandButton cmdTranfer 
         Height          =   315
         Left            =   2520
         Picture         =   "frmCapAperturasLote.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   555
         Width           =   475
      End
      Begin VB.ComboBox cboTransferMoneda 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label lblEtiMonTra 
         AutoSize        =   -1  'True
         Caption         =   "Monto Transacción"
         Height          =   195
         Left            =   870
         TabIndex        =   73
         Top             =   1710
         Width           =   1380
      End
      Begin VB.Label lblSimTra 
         AutoSize        =   -1  'True
         Caption         =   "S/"
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
         Left            =   2370
         TabIndex        =   72
         Top             =   1680
         Width           =   240
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
         Left            =   2850
         TabIndex        =   71
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblTTCVD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3570
         TabIndex        =   70
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblTTCCD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3570
         TabIndex        =   69
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "TCV"
         Height          =   285
         Left            =   3120
         TabIndex        =   68
         Top             =   480
         Width           =   390
      End
      Begin VB.Label lblTTCC 
         Caption         =   "TCC"
         Height          =   285
         Left            =   3120
         TabIndex        =   67
         Top             =   180
         Width           =   390
      End
      Begin VB.Label lblTransferGlosa 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   60
         TabIndex        =   66
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblTransferMoneda 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   60
         TabIndex        =   65
         Top             =   225
         Width           =   585
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
         Height          =   315
         Left            =   840
         TabIndex        =   64
         Top             =   555
         Width           =   1575
      End
      Begin VB.Label lbltransferBcol 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
         Height          =   195
         Left            =   60
         TabIndex        =   63
         Top             =   930
         Width           =   555
      End
      Begin VB.Label lbltransferN 
         AutoSize        =   -1  'True
         Caption         =   "Nro Doc :"
         Height          =   195
         Left            =   60
         TabIndex        =   62
         Top             =   570
         Width           =   690
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
         Height          =   315
         Left            =   840
         TabIndex        =   61
         Top             =   930
         Width           =   3465
      End
   End
   Begin ComctlLib.ProgressBar pvCuentasSueldo 
      Height          =   225
      Left            =   4800
      TabIndex        =   55
      Top             =   8490
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   1500
      Top             =   8340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame FraITFAsume 
      Height          =   615
      Left            =   7320
      TabIndex        =   42
      Top             =   6450
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton OptAsuITF 
         Caption         =   "No Asume ITF"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptAsuITF 
         Caption         =   "Asume ITF"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1095
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
      Height          =   2040
      Left            =   7200
      TabIndex        =   12
      Top             =   6270
      Width           =   3255
      Begin VB.CheckBox chkITFEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Efect"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   810
         TabIndex        =   33
         Top             =   1320
         Width           =   705
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   825
         Width           =   1905
         _extentx        =   3360
         _extenty        =   661
         font            =   "frmCapAperturasLote.frx":0B8E
         backcolor       =   12648447
         forecolor       =   12582912
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   1710
         Width           =   450
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
         Left            =   1620
         TabIndex        =   35
         Top             =   1260
         Width           =   1125
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   840
         TabIndex        =   34
         Top             =   1650
         Width           =   1905
      End
      Begin VB.Label lblDispCTS 
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
         Height          =   375
         Left            =   1725
         TabIndex        =   22
         Top             =   285
         Width           =   795
      End
      Begin VB.Label lblCTS 
         AutoSize        =   -1  'True
         Caption         =   "Disponibilidad (%) :"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   345
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   900
         Width           =   540
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/"
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
         Left            =   2835
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2040
      Left            =   4560
      TabIndex        =   10
      Top             =   6270
      Width           =   2595
      Begin VB.TextBox txtGlosa 
         Height          =   1770
         Left            =   90
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   210
         Width           =   2415
      End
   End
   Begin VB.Frame fraITF 
      Height          =   615
      Left            =   80
      TabIndex        =   29
      Top             =   1590
      Width           =   10390
      Begin VB.CheckBox chkExoITF 
         Caption         =   "Exonerado ITF"
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
         Height          =   225
         Left            =   105
         TabIndex        =   31
         Top             =   30
         Width           =   1590
      End
      Begin VB.ComboBox cboTipoExoneracion 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   195
         Width           =   2925
      End
      Begin VB.Label Label17 
         Caption         =   "Tipo de Exoneracion :"
         Height          =   225
         Left            =   2685
         TabIndex        =   32
         Top             =   255
         Width           =   1620
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   80
      TabIndex        =   16
      Top             =   0
      Width           =   10390
      Begin VB.TextBox txtPlazo 
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
         Height          =   300
         Left            =   8640
         MaxLength       =   5
         TabIndex        =   78
         Text            =   "0"
         Top             =   1080
         Width           =   1425
      End
      Begin VB.ComboBox cboInstConvDep 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CheckBox chkRelConv 
         Alignment       =   1  'Right Justify
         Caption         =   "Relacion con Convenio"
         Height          =   375
         Left            =   4380
         TabIndex        =   47
         Top             =   300
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cboPrograma 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.ComboBox cboPeriodoCTS 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   2355
      End
      Begin VB.ComboBox cboTipoTasa 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   690
         Width           =   1995
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   1995
      End
      Begin SICMACT.TxtBuscar txtInstitucion 
         Height          =   375
         Left            =   7860
         TabIndex        =   45
         Top             =   300
         Width           =   2175
         _extentx        =   3836
         _extenty        =   661
         appearance      =   1
         appearance      =   1
         font            =   "frmCapAperturasLote.frx":0BBA
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
         tipobuspers     =   1
      End
      Begin VB.Label lblPlazo 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   8040
         TabIndex        =   79
         Top             =   1140
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblInst 
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
         Left            =   4380
         TabIndex        =   46
         Top             =   690
         Width           =   5655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tasa EA (%) :"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label lblTasa 
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1200
         TabIndex        =   40
         Top             =   1087
         Width           =   1995
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
         Height          =   195
         Left            =   3300
         TabIndex        =   39
         Top             =   1140
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblInstEtq 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         Height          =   195
         Left            =   6840
         TabIndex        =   28
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblPeriodo 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   3300
         TabIndex        =   19
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Tasa :"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   390
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   80
      TabIndex        =   8
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4065
      Left            =   80
      TabIndex        =   9
      Top             =   2190
      Width           =   10390
      Begin VB.CommandButton cmdFormato 
         Caption         =   "Formato"
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   3600
         Visible         =   0   'False
         Width           =   915
      End
      Begin SICMACT.FlexEdit grdCuenta 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   10155
         _extentx        =   17912
         _extenty        =   5530
         cols0           =   22
         highlight       =   1
         allowuserresizing=   1
         rowsizingmode   =   1
         encabezadosnombres=   $"frmCapAperturasLote.frx":0BE2
         encabezadosanchos=   "350-1550-3400-400-1500-600-700-700-1200-0-0-0-900-0-0-0-0-0-0-0-0-0"
         font            =   "frmCapAperturasLote.frx":0CA7
         font            =   "frmCapAperturasLote.frx":0CCF
         font            =   "frmCapAperturasLote.frx":0CF7
         font            =   "frmCapAperturasLote.frx":0D1F
         font            =   "frmCapAperturasLote.frx":0D47
         fontfixed       =   "frmCapAperturasLote.frx":0D6F
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-1-X-X-4-5-6-X-8-X-X-X-X-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-1-0-0-3-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-L-C-R-R-R-C-C-C-R-R-R-C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-3-2-2-0-0-0-0-0-0-0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         lbbuscaduplicadotext=   -1
         colwidth0       =   345
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit grdClientes 
         Height          =   1695
         Left            =   120
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   9975
         _extentx        =   17595
         _extenty        =   2990
         cols0           =   10
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Código-Nombre-RE-OP-Tasa-Monto-cPersoneria-nTasaNominal-ITF"
         encabezadosanchos=   "350-1550-3900-400-600-700-1200-0-0-900"
         font            =   "frmCapAperturasLote.frx":0D95
         font            =   "frmCapAperturasLote.frx":0DBD
         font            =   "frmCapAperturasLote.frx":0DE5
         font            =   "frmCapAperturasLote.frx":0E0D
         font            =   "frmCapAperturasLote.frx":0E35
         fontfixed       =   "frmCapAperturasLote.frx":0E5D
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-4-5-X-X-X-X-7"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-3-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-C-L-R-R-C-R"
         formatosedit    =   "0-0-0-0-0-3-2-0-0-2"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbflexduplicados=   0
         lbbuscaduplicadotext=   -1
         colwidth0       =   345
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.TextBox txtArchivo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3600
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7050
         TabIndex        =   51
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "..."
         Height          =   315
         Left            =   6420
         TabIndex        =   50
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Left            =   9210
         TabIndex        =   5
         Top             =   3600
         Width           =   1035
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   8100
         TabIndex        =   4
         Top             =   3600
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Archivo :"
         Height          =   195
         Left            =   1170
         TabIndex        =   49
         Top             =   3660
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Label lblCuentaSueldo 
      Alignment       =   2  'Center
      Caption         =   "Verifica cuentas caja sueldo:"
      Height          =   255
      Left            =   2490
      TabIndex        =   56
      Top             =   8490
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "frmCapAperturasLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProducto As COMDConstantes.Producto
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim bDocumento As Boolean
Dim nmoneda As COMDConstantes.Moneda
Dim nTipoTasa As COMDConstantes.CaptacTipoTasa
Public sCodIF As String
Public dFechaValorizacion As Date
Public lnDValoriza As Integer
Dim nTasaNominal As Double
Dim sTipoCuenta  As String, sPersCod As String
Dim lnTpoPrograma As Integer
Dim nRedondeoITF As Double
Dim nTpoProgramaCTS As Integer '***Agregado por ELRO el 20121129, según OYP-RFC101-2012
Dim oDocRec As UDocRec 'EJVG20140409
'RIRO20140430 ERS017 *******
Dim bCargaLote As Boolean
Dim nNroApertura As Integer ' Contiene el numero de cuentas a aperturar
Dim fnMovNroRVD As Long
Dim lnMovNroTransfer As Long
Dim lsDetalle As String
Dim bLoad As Boolean
'End RIRO ******************
'JUEZ 20141010 Nuevos Parametros **********
Dim bParPersJur As Boolean
Dim bParPersNat As Boolean
Dim bParMonedaSol As Boolean
Dim bParMonedaDol As Boolean
Dim nParMontoMinSol As Double
Dim nParMontoMinDol As Double
Dim nParOrdPag As Integer
Dim nParPlazoMin As Integer
Dim nParPlazoMax As Integer
Dim bParFormaRetFinPlazo As Boolean
Dim bParFormaRetMensual As Boolean
Dim bParFormaRetIniPlazo As Boolean
'END JUEZ *********************************
Dim fnCampanaCod As Long 'JUEZ 20160420

Private Function ValidaDatosCuentas() As Boolean
Dim i As Long
Dim nMonto As Double, nMontoTotal As Double
Dim nTasa As Double
Dim nFormaRetiro As COMDConstantes.CaptacPFFormaRetiro
Dim nPersoneria As COMDConstantes.PersPersoneria
Dim nPlazo As Long
Dim sTitular As String
Dim oCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim lsErrMonto As String
Dim fs As Scripting.FileSystemObject
Set fs = New Scripting.FileSystemObject
Dim nMontoMinProd As Double 'JUEZ 20141010

'***Agregado por ELRO el 20121204, según OYP-RFC101-2012
If grdCuenta.TextMatrix(1, 0) = "" Then Exit Function
'***Fin Agregado por ELRO el 20121204*******************
 
nMontoTotal = grdCuenta.SumaRow(8)
nMontoMinProd = IIf(CInt(Right(cboMoneda.Text, 1)) = gMonedaNacional, nParMontoMinSol, nParMontoMinDol) 'JUEZ 20141010

If nMontoTotal <> txtMonto.value Then
    MsgBox "Monto total no coincide con la suma de los montos de cada cuenta.", vbInformation, "Aviso"
    grdCuenta.SetFocus
    ValidaDatosCuentas = False
    Exit Function
End If
For i = 1 To grdCuenta.Rows - 1
    sTitular = Trim(grdCuenta.TextMatrix(i, 2))
    nMonto = CDbl(grdCuenta.TextMatrix(i, 8))
    nTasa = CDbl(grdCuenta.TextMatrix(i, 7))
    nPlazo = CLng(IIf(grdCuenta.TextMatrix(i, 6) = "", 0, grdCuenta.TextMatrix(i, 6))) 'JUEZ 20141010
    nFormaRetiro = CInt(IIf(grdCuenta.TextMatrix(i, 4) = "", 0, Trim(Right(grdCuenta.TextMatrix(i, 4), 2)))) 'JUEZ 20141010
    '***Modificado por ELRO el 20121129, según OYP-RFC101-2012
    'If nMonto <= 0 And (cboPrograma.ListIndex <> 6 And cboPrograma.ListIndex <> 5 And cboPrograma.ListIndex <> 0) Then
    '    MsgBox "Monto debe ser mayor a cero. Titular : " & sTitular, vbInformation, "Aviso"
    '    grdCuenta.Row = i
    '    grdCuenta.Col = 8
    '    grdCuenta.SetFocus
    '    ValidaDatosCuentas = False
    '    Exit Function
    'End If
    If nProducto <> gCapCTS Then
        'JUEZ 20141010 Nuevos parámetros ************************************************
        'If nMonto <= 0 And (cboPrograma.ListIndex <> 6 And cboPrograma.ListIndex <> 5 And cboPrograma.ListIndex <> 0 And cboPrograma.ListIndex <> 8) Then
        ''***Condición cboPrograma.ListIndex <> 8 agregado por ELRO el 20130131
        '    MsgBox "Monto debe ser mayor a cero. Titular : " & sTitular, vbInformation, "Aviso"
        '    grdCuenta.row = i
        '    grdCuenta.Col = 8
        '    grdCuenta.SetFocus
        '    ValidaDatosCuentas = False
        '    Exit Function
        'End If
        If Trim(Right(cboPrograma.Text, 1)) <> "" Then
            If nMonto < nMontoMinProd Then
                MsgBox "El monto no debe ser menor de " & Format(nMontoMinProd, "#,##0.00") & ". Titular : " & sTitular, vbInformation, "Aviso"
                grdCuenta.row = i
                grdCuenta.Col = 8
                grdCuenta.SetFocus
                ValidaDatosCuentas = False
                Exit Function
            End If
        End If
        'END JUEZ ***********************************************************************
    End If
    '***Modificado por ELRO el 20121129***********************
    If nTasa <= 0 Then
        MsgBox "Tasa debe ser mayor a cero. Titular : " & sTitular, vbInformation, "Aviso"
        grdCuenta.row = i
        grdCuenta.Col = 7
        grdCuenta.SetFocus
        ValidaDatosCuentas = False
        Exit Function
    End If
    If grdCuenta.TextMatrix(i, 9) = "" Then
        MsgBox "Personería No Definida para este cliente. Defina nuevamente a la persona.", vbInformation, "Aviso"
        grdCuenta.row = i
        grdCuenta.Col = 1
        grdCuenta.SetFocus
        ValidaDatosCuentas = False
        Exit Function
    End If
    
    If nProducto = gCapAhorros Then
        'CAAU - MAYNAS Valida para los otros sub Pruductos diferentes al ahorro normal
        'JUEZ 20141010 Comentado para nuevos parámetros *************************************************
'        If Trim(Right(Me.cboPrograma.Text, 3)) <> "0" And Trim(Right(Me.cboPrograma.Text, 3)) <> "5" And Trim(Right(Me.cboPrograma.Text, 3)) <> "6" And Trim(Right(Me.cboPrograma.Text, 3)) <> "8" Then
'        '***Condición And Trim(Right(Me.cboPrograma.Text, 3)) <> "8" agregado por ELRO el 20130131
''            If grdCuenta.TextMatrix(i, 13) = "" Then
''                MsgBox "No se registro Plazo Minimo", vbInformation, "Aviso"
''                grdCuenta.Row = i
''                grdCuenta.Col = 1
''                grdCuenta.SetFocus
''                ValidaDatosCuentas = False
''                Exit Function
''            End If
'
''            If grdCuenta.TextMatrix(i, 14) = "" Then
''                MsgBox "No se registro el Monto Min de Deposito", vbInformation, "Aviso"
''                grdCuenta.Row = i
''                grdCuenta.Col = 1
''                grdCuenta.SetFocus
''                ValidaDatosCuentas = False
''                Exit Function
''            End If
'
'            Dim sErrorV As String
'            Dim nPlazoMin As Integer
'            Dim nMontoMinDep As Currency
'            Dim objCap As COMNCaptaGenerales.NCOMCaptaDefinicion
'            Set objCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
'            sErrorV = ""
'            sErrorV = objCap.ValidaPlazoMin_SubProductoAHO(grdCuenta.TextMatrix(i, 13), grdCuenta.TextMatrix(i, 9), nmoneda, CInt(Trim(Right(Me.cboPrograma.Text, 3))), nPlazoMin, IIf(grdCuenta.TextMatrix(i, 5) = ".", True, False))
'            If Trim(sErrorV) <> "" Then
'                MsgBox sErrorV, vbInformation, "AVISO"
'                Set objCap = Nothing
'                grdCuenta.TextMatrix(i, 13) = IIf(IsNull(nPlazoMin), 0, nPlazoMin)
'                Exit Function
'            End If
'            sErrorV = ""
'            sErrorV = objCap.ValidaMontoMinDep_SubProductoAHO(grdCuenta.TextMatrix(i, 14), grdCuenta.TextMatrix(i, 9), nmoneda, CInt(Trim(Right(Me.cboPrograma.Text, 3))), nMontoMinDep, IIf(grdCuenta.TextMatrix(i, 5) = ".", True, False))
'            If Trim(sErrorV) <> "" And cboPrograma.ListIndex <> 6 Then
'                MsgBox sErrorV, vbInformation, "AVISO"
'                Set objCap = Nothing
'                grdCuenta.TextMatrix(i, 14) = IIf(IsNull(nMontoMinDep), 0, nMontoMinDep)
'                Exit Function
'            End If
'            'Set clsCap = Nothing
'        End If
        'END JUEZ ***************************************************************************************
    End If
     
    If nProducto = gCapPlazoFijo Then
        If grdCuenta.TextMatrix(i, 4) = "" Then
            MsgBox "Forma de retiro no Definida.", vbInformation, "Aviso"
            grdCuenta.row = i
            grdCuenta.Col = 4
            grdCuenta.SetFocus
            SendKeys "{Enter}"
            ValidaDatosCuentas = False
            Exit Function
        'JUEZ 20141010 Nuevos parámetros **************************************
'        ElseIf CLng(grdCuenta.TextMatrix(i, 6)) < 30 And Right(grdCuenta.TextMatrix(i, 4), 1) = 1 Then
'            MsgBox "Si la forma de Retiro es mensual, el Plazo debe ser mayor o igual a Treinta dias.", vbInformation, "Aviso"
'            grdCuenta.row = i
'            grdCuenta.Col = 6
'            grdCuenta.SetFocus
'            SendKeys "{Enter}"
'            ValidaDatosCuentas = False
'            Exit Function
        End If
        If Trim(Right(cboPrograma.Text, 1)) <> "" Then
            If Not ValidarPlazoPF(nPlazo, sTitular) Then
                ValidaDatosCuentas = False
                Exit Function
            End If
            If Not ValidarMedioRetiroPF(nFormaRetiro, sTitular) Then
                ValidaDatosCuentas = False
                Exit Function
            End If
        End If
        'END JUEZ *************************************************************
    End If
    
    
    'JUEZ 20141010 Comentado para nuevos parámetros ***********************
    'lsErrMonto = oCap.ValidaMontoApertura(nProducto, grdCuenta.TextMatrix(i, 9), nMonto, nMoneda, IIf(grdCuenta.TextMatrix(i, 5) = ".", True, False))
    'If lsErrMonto <> "" Then
    '    MsgBox lsErrMonto & " para : " & sTitular, vbInformation, "Aviso"
    '    grdCuenta.row = i
    '    grdCuenta.Col = 8
    '    grdCuenta.SetFocus
    '    SendKeys "{Enter}"
    '    ValidaDatosCuentas = False
    '    Exit Function
    'End If
    'END JUEZ *************************************************************
Next i
'***Agregado por ELRO el 20121129, según OYP-RFC020-2013
If fs.FileExists(App.Path & "\FormatoCarta\CARTILLACTS2.doc") = False Then
   MsgBox "No existe la plantilla CARTILLACTS2.doc en la carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
   ValidaDatosCuentas = False
   Exit Function
End If
'***Fin Agregado por ELRO el 20121129

'RIRO20140610 ERS017 ********************************
'If nProducto = gCapAhorros And CInt(Trim(Right(Me.cboPrograma.Text, 3))) = 6 Then RIRO 20141002
If nProducto = gCapAhorros Then
    If CInt(Trim(Right(Me.cboPrograma.Text, 3))) = 6 Then 'RIRO 20141002
    
    If Not ValidarLoteTitular Then
        MsgBox "Cliente ya dispone de una cuenta en la institucion seleccionada", vbInformation, "Aviso"
        ValidaDatosCuentas = False
        Exit Function
    End If
    
    End If
End If
If Len(Trim(Replace(txtGlosa.Text, vbNewLine, " "))) = 0 Then
    MsgBox "Debe ingresar una observacion a la glosa", vbInformation, "Aviso"
    ValidaDatosCuentas = False
    Exit Function
End If
'Validando Ordenes de Pago
'JUEZ 20141010 Comentado para nuevos parámetros ******************
'For i = 1 To grdCuenta.Rows - 1
'    If grdCuenta.TextMatrix(i, 5) = "." Then
'        If nProducto <> gCapAhorros Or lnTpoPrograma <> 0 Then
'            MsgBox "Solo cuentas Ahorro corriente pueden tener ordenes de pago", vbExclamation, "Aviso"
'            ValidaDatosCuentas = False
'            Exit Function
'        End If
'    End If
'Next
'END JUEZ ********************************************************
'END RIRO *******************************************

Set oCap = Nothing
ValidaDatosCuentas = True
End Function

Private Sub IniciaComboCTSPeriodo()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsConst As New ADODB.Recordset
Dim sCodigo As String * 2
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsConst = clsMant.GetCTSPeriodo()
Set clsMant = Nothing
Do While Not rsConst.EOF
    sCodigo = rsConst("nItem")
    cboPeriodoCTS.AddItem sCodigo & space(2) & UCase(rsConst("cDescripcion")) & space(100) & rsConst("nPorcentaje")
    rsConst.MoveNext
Loop
cboPeriodoCTS.ListIndex = 0
End Sub

Private Sub ClearScreen()
grdCuenta.Rows = 2
grdCuenta.FormaCabecera

Select Case nProducto
    Case gCapAhorros
        lblPeriodo.Visible = False
        cboPeriodoCTS.Visible = False
        grdCuenta.ColWidth(4) = 0
        grdCuenta.ColWidth(6) = 0
        lblDispCTS.Visible = False
        lblCTS.Visible = False
        lblInst.Visible = False
        lblInstEtq.Visible = False
        txtInstitucion.Visible = False
        chkRelConv.value = 0
        txtArchivo.Text = ""
        cboPrograma.ListIndex = 0
        lblDispCTS.Caption = "" '****Agregado por ELRO el 20121205, según OYP-RFC101-2012
    Case gCapPlazoFijo
        lblPeriodo.Visible = False
        cboPeriodoCTS.Visible = False
        grdCuenta.ColWidth(5) = 0
        Dim clsGen As DGeneral
        Dim rsRel As Recordset
        Set clsGen = New DGeneral
        Set rsRel = clsGen.GetConstante(gCaptacPFFormaRetiro)
        Set clsGen = Nothing
        grdCuenta.CargaCombo rsRel
        Set rsRel = Nothing
        lblDispCTS.Visible = False
        lblCTS.Visible = False
        lblInst.Visible = False
        lblInstEtq.Visible = False
        txtInstitucion.Visible = False
        lblDispCTS.Caption = "" '****Agregado por ELRO el 20121205, según OYP-RFC101-2012
    Case gCapCTS
        If nOperacion = gCTSApeLoteChq Then Me.Caption = "Captaciones - Apertura Lote Cheque - CTS"
        lblPeriodo.Visible = True
        cboPeriodoCTS.Visible = True
        grdCuenta.ColWidth(4) = 0
        grdCuenta.ColWidth(5) = 0
        grdCuenta.ColWidth(6) = 0
        lblDispCTS.Visible = True
        lblCTS.Visible = True
        lblInst.Visible = True
        txtInstitucion.Visible = True
        lblInstEtq.Visible = True
        lblDispCTS.Caption = "" '****Agregado por ELRO el 20121205, según OYP-RFC101-2012
        IniciaComboCTSPeriodo
        cboPeriodoCTS_Click '****Agregado por ELRO el 20121205, según OYP-RFC101-2012
        txtArchivo = ""
End Select

txtMonto.Text = "0.00"
cboMoneda.ListIndex = 0
cboTipoTasa.ListIndex = 0
If nProducto = gCapCTS Then
    cboPeriodoCTS.ListIndex = 0
    txtInstitucion.Text = ""
    lblInst.Caption = ""
End If
'lblNroDoc = ""
'lblNombreIF = ""
Call limpiarDatosCheque 'CTI6 20210607
txtGlosa.Text = ""
'lblDispCTS.Caption = ""'****Comentado por ELRO el 20121205, según OYP-RFC101-2012
OptAsuITF(0).value = False
OptAsuITF(1).value = False
    
'RIRO20140430 ERS017 *************
bCargaLote = False
If nOperacion = gAhoApeLoteTransfBanco Or nOperacion = gAhoApeLoteEfec Or nOperacion = gPFApeLoteTransf Or nOperacion = gCTSApeLoteTransfNew Then
    SetDatosTransferencia "", "", "", 0, -1, ""
    nNroApertura = 0
    fnMovNroRVD = 0
    lnMovNroTransfer = 0
    grdCuenta.ColumnasAEditar = "X-1-X-X-4-5-6-X-8-X-X-X-X-X-X-X-X"
End If
'END IF **************************
    
End Sub

'CTI6 20210607***
Private Sub setDatosCheque()
    lblNroDoc.Caption = oDocRec.fsNroDoc
    lblNombreIF.Caption = oDocRec.fsPersNombre
    lblDocNroOpe.Caption = oDocRec.fnNroCliLote
    lblDocMonto.Caption = Format(oDocRec.fnMonto, gsFormatoNumeroView)
    txtGlosa.Text = oDocRec.fsGlosa
End Sub

Private Sub limpiarDatosCheque()
    Set oDocRec = New UDocRec
    Call setDatosCheque
End Sub

Private Function validarApertLoteCheque() As String
    Dim sValidacion As String
    Dim nRegistrosApertura As Integer
    Dim i As Integer
        
    nRegistrosApertura = grdCuenta.Rows - 1
    If Trim(grdCuenta.TextMatrix(nRegistrosApertura, 1)) = "" And Trim(grdCuenta.TextMatrix(nRegistrosApertura, 2)) = "" Then
        nRegistrosApertura = nRegistrosApertura - 1
    End If
    sValidacion = ""
    'Validando que grid tenga registros
    If nRegistrosApertura = 0 Then
        sValidacion = "El Grid de cuentas a aperturar, no tiene registrado ningún cliente." & vbNewLine
    End If
    'Validando la cantidad de registros del voucher con la cantidad de registros del grid
    If oDocRec.fnNroCliLote <> nRegistrosApertura Then
        sValidacion = sValidacion & "El número de registros del grid es diferente al número de declarado en el cheque." & vbNewLine
    End If
    If CDbl(oDocRec.fnMonto) <> CDbl(txtMonto.value) Then
        sValidacion = sValidacion & "El Monto del total de aperturas debe ser igual al monto del cheque." & vbNewLine
    End If
   
    validarApertLoteCheque = sValidacion
End Function
'END CTI6 *******
Private Sub DefineTasaGrid(Optional ByVal pnTpoPrograma As Integer)
Dim i As Long
'RIRO20140407 ERS017  ****************
pvCuentasSueldo.Min = 1
pvCuentasSueldo.Max = grdCuenta.Rows
pvCuentasSueldo.value = 1
Evento
lblCuentaSueldo.Caption = "Actualizando datos del Grid"
pvCuentasSueldo.Visible = True
lblCuentaSueldo.Visible = True
'END RIRO ********************************
For i = 1 To grdCuenta.Rows - 1
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim bOrdPag As Boolean
    Dim nMonto As Double, nTasa As Double
    Dim nPlazo As Long
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    
    bOrdPag = IIf(grdCuenta.TextMatrix(i, 5) = ".", True, False)
    If grdCuenta.TextMatrix(i, 8) <> "" Then
        nMonto = CDbl(grdCuenta.TextMatrix(i, 8))
    Else
        nMonto = 0
    End If
    If nProducto = gCapPlazoFijo Then
        If grdCuenta.TextMatrix(i, 6) <> "" Then
            nPlazo = CLng(grdCuenta.TextMatrix(i, 6))
            nTasa = clsDef.GetCapTasaInteresCamp(nProducto, pnTpoPrograma, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, False, fnCampanaCod) 'JUEZ 20160420
            If nTasa = 0 Then nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge) 'JUEZ 20160420
        End If
    Else
        nTasa = clsDef.GetCapTasaInteresCamp(nProducto, pnTpoPrograma, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, bOrdPag, fnCampanaCod) 'JUEZ 20160420
        If nTasa = 0 Then nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , nMonto, gsCodAge, bOrdPag, pnTpoPrograma) 'JUEZ 20160420
    End If
    'grdCuenta.TextMatrix(i, 4) = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00") 'JUEZ 20141010
    grdCuenta.TextMatrix(i, 7) = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
    grdCuenta.TextMatrix(i, 11) = Format$(nTasa, "#,##0.00")
    grdCuenta.TextMatrix(i, 21) = fnCampanaCod 'JUEZ 20160420
    Set clsDef = Nothing
    'RIRO20140407 ERS017  **********
    Evento
    If pvCuentasSueldo.Max >= i Then
        pvCuentasSueldo.value = i
    End If
    'END RIRO **************************
Next i
pvCuentasSueldo.Visible = False 'RIRO20140407 ERS017
lblCuentaSueldo.Visible = False 'RIRO20140407 ERS017
End Sub

'RIRO20140430 ERS017 Se agregaron los parámetros sFiltro,sFiltro2,sFiltro3
Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As ConstanteCabecera, _
                        Optional ByVal sFiltro As String = "", Optional ByVal sFiltro2 As String = "", _
                        Optional ByVal sFiltro3 As String = "")
                        
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConst As New ADODB.Recordset
Set clsGen = New COMDConstSistema.DCOMGeneral
'Set rsConst = clsGen.GetConstante(nCapConst) ' RIRO ERS017 Comentado
Set rsConst = clsGen.GetConstante(nCapConst, sFiltro, sFiltro2, sFiltro3) 'RIRO ERS017 Agregado
Set clsGen = Nothing
Do While Not rsConst.EOF
    cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
    rsConst.MoveNext
Loop
cboConst.ListIndex = 0
End Sub

Public Sub Inicia(ByVal nProd As Producto, ByVal nOpe As CaptacOperacion, ByVal sDescOperacion As String)
nProducto = nProd
nOperacion = nOpe
bLoad = True
fgITFParamAsume gsCodAge, CStr(nProd)

Me.cboPrograma.Visible = False
Label20.Visible = False
Me.lblTasa.Visible = False
Me.lblPlazo.Visible = False
Me.txtPlazo.Visible = False
Select Case nProd
    Case gCapAhorros
        Me.Caption = "Captaciones - Ahorros - " & sDescOperacion
        lblPeriodo.Visible = False
        cboPeriodoCTS.Visible = False
        grdCuenta.ColWidth(4) = 0
        grdCuenta.ColWidth(6) = 0
        lblDispCTS.Visible = False
        lblCTS.Visible = False
        lblInst.Visible = False
        lblInstEtq.Visible = False
        txtInstitucion.Visible = False
        Me.fraITF.Visible = True
        Me.chkITFEfectivo.value = 0
        Me.chkITFEfectivo.Enabled = True
                
        If gbITFAsumidoAho Then
            chkITFEfectivo.value = 1
            chkITFEfectivo.Visible = False
        Else
            chkITFEfectivo.value = 0
            chkITFEfectivo.Visible = True
        End If
        
        'IniciaCombo cboPrograma, 2030 RIRO ERS017 Comentado
        IniciaCombo cboPrograma, 2030, "4", , " " 'RIRO20140430 ERS 017-2014 Agregado
        Me.cboPrograma.Visible = True
        Label20.Visible = True
        Me.lblTasa.Visible = True
        FraITFAsume.Visible = False
        
        'Add By GITU 11-09-2012
        'chkRelConv.Visible = True RIRO ERS017 Comentado
        cboPrograma_Click 'RIRO ERS017 Agregado
        
        IniciaComboConvDep 9
        'grdClientes.Visible = True
        Label6.Visible = True
        txtArchivo.Visible = True
        cmdBusca.Visible = True
        cmdCargar.Visible = True
        'End GITU
        cboTipoTasa.Enabled = False '***Agregado por ELRO el 20130131, según TI-ERS020-2013
        grdCuenta.ColWidth(17) = 0
        grdCuenta.ColWidth(18) = 0
        grdCuenta.ColWidth(19) = 0
        cmdFormato.Visible = True 'RIRO20140610 ERS017
    Case gCapPlazoFijo
        
        Me.Caption = "Captaciones - Plazo Fijo - " & sDescOperacion
        lblPeriodo.Visible = False
        cboPeriodoCTS.Visible = False
        grdCuenta.ColWidth(5) = 0
        Dim clsGen As COMDConstSistema.DCOMGeneral
        Dim rsRel As ADODB.Recordset
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Set rsRel = clsGen.GetConstante(gCaptacPFFormaRetiro)
        Set clsGen = Nothing
        grdCuenta.CargaCombo rsRel
        Set rsRel = Nothing
        lblDispCTS.Visible = False
        lblCTS.Visible = False
        lblInst.Visible = False
        lblInstEtq.Visible = False
        txtInstitucion.Visible = False
        Me.fraITF.Visible = True
        Me.chkITFEfectivo.value = 1
        '***Modificado por ELRO la fecha 20110912, según Acta 245-2011/TI-D
        'Me.chkITFEfectivo.Enabled = False  'comentado por ELRO el 20110912
        Me.chkITFEfectivo.Enabled = True
        ''***Fin Modificado por ELRO***************************************
    
        If gbITFAsumidoPF Then
            chkITFEfectivo.value = 0
        Else
            chkITFEfectivo.value = 1
        End If
        '***Modificado por ELRO la fecha 20110912, según Acta 245-2011/TI-D
        'chkITFEfectivo.Visible = True  'comentado por ELRO el 20110912
        'chkITFEfectivo.Enabled = False 'comentado por ELRO el 20110912
        'FraITFAsume.Visible = True 'comentado por ELRO el 20110912
        'CTI7 OPv2********************************************************
        IniciaCombo cboPrograma, 2032, "", , " " 'RIRO20140430 ERS 017-2014 Agregado
        Me.cboPrograma.Visible = True
        Label20.Visible = True
        Me.lblTasa.Visible = True
        Me.lblPlazo.Visible = False
        Me.txtPlazo.Visible = False
        cboPrograma_Click
        '*****************************************************************
        FraITFAsume.Visible = False
        '***Fin Modificado por ELRO****************************************
        
    Case gCapCTS
        Me.Caption = "Captaciones - CTS - " & sDescOperacion
        If nOperacion = gCTSApeLoteChq Then Me.Caption = "Captaciones - Apertura Lote Cheque - CTS"
        lblPeriodo.Visible = True
        cboPeriodoCTS.Visible = True
        grdCuenta.ColWidth(4) = 0
        grdCuenta.ColWidth(5) = 0
        grdCuenta.ColWidth(6) = 0
        lblDispCTS.Visible = True
        lblCTS.Visible = True
        lblInst.Visible = True
        txtInstitucion.Visible = True
        lblInstEtq.Visible = True
        IniciaComboCTSPeriodo
        Me.fraITF.Visible = False
        Me.fraITF.Visible = False
        chkITFEfectivo.Visible = False
        FraITFAsume.Visible = False
        '***Agregado por ELRO el 20121130, según OYP-RFC101-2012
        Label4.Visible = False
        Label6.Visible = True
        txtArchivo.Visible = True
        cmdBusca.Visible = True
        cmdCargar.Visible = True
        cboTipoTasa.Enabled = False
        grdCuenta.ColWidth(17) = 0
        grdCuenta.ColWidth(18) = 0
        grdCuenta.ColWidth(19) = 0
        '***Fin Agregado por ELRO el 20121130*******************
End Select

Select Case nOperacion
    Case gAhoApeLoteEfec, gPFApeLoteEfec, gCTSApeLoteEfec
        fraDocumento.Visible = False
    Case gAhoApeLoteChq, gPFApeLoteChq, gCTSApeLoteChq
        fraDocumento.Visible = True
'        Me.chkITFEfectivo.value = 1
'        Me.chkITFEfectivo.Enabled = False
    'RIRO20140407 ERS017
    Case gAhoApeLoteTransfBanco, gCTSApeLoteTransfNew, gPFApeLoteTransf
        fraDocumento.Visible = False
        fraTranferecia.Visible = True
        
    'END RIRO
End Select

If nProducto = gCapAhorros And gbITFAsumidoAho Then
    Me.chkITFEfectivo.value = 0
ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
    Me.chkITFEfectivo.value = 0
End If

chkExoITF_Click

IniciaCombo cboMoneda, gMoneda
IniciaCombo cboTransferMoneda, gMoneda 'RIRO20140430 ERS017
IniciaCombo cboTipoTasa, gCaptacTipoTasa
cmdAgregar.Enabled = True
cmdEliminar.Enabled = False
txtMonto.Enabled = False

Me.Show 1
End Sub

Private Sub cboMoneda_Click()
Dim i As Long '***Agregado por ELRO el 20130131, según TI-ERS020-2013
nmoneda = CLng(Right(cboMoneda.Text, 1))
'JUEZ 20141010 VERIFICAR PARAMETRO MONEDA *****************
If Trim(Right(cboPrograma.Text, 1)) <> "" Then
    If nmoneda = gMonedaNacional And Not bParMonedaSol Then
        '''MsgBox "El producto no permite apertura de cuentas en soles", vbInformation, "Aviso" 'marg ers044-2016
        MsgBox "El producto no permite apertura de cuentas en " & StrConv(gcPEN_PLURAL, vbLowerCase), vbInformation, "Aviso" 'marg ers044-2016
        cboMoneda.ListIndex = 1
    End If
    If nmoneda = gMonedaExtranjera And Not bParMonedaDol Then
        MsgBox "El producto no permite apertura de cuentas en dólares", vbInformation, "Aviso"
        cboMoneda.ListIndex = 0
    End If
End If
'END JUEZ *************************************************

If nmoneda = gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
    '''lblMon.Caption = "S/." 'marg ers044-2016
    lblMon.Caption = gcPEN_SIMBOLO 'marg ers044-2016
ElseIf nmoneda = gMonedaExtranjera Then
    txtMonto.BackColor = &HC0FFC0
    lblMon.Caption = "US$"
End If

Me.lblITF.BackColor = txtMonto.BackColor
Me.lblTotal.BackColor = txtMonto.BackColor

'***Modificado por ELRO el 20130131, según TI-ERS020-2013
'DefineTasaGrid
If nProducto <> gCapCTS Then
    'DefineTasaGrid (CInt(Right(Trim(cboPrograma.Text), 2)))
    DefineTasaGrid (Val(Right(Trim(cboPrograma.Text), 2))) 'EJVG20140408
Else
    If Trim(grdCuenta.TextMatrix(1, 0)) <> "" Then
        For i = 1 To CLng(grdCuenta.Rows - 1)
            Call grdCuenta_OnCellChange(i, 8)
        Next i
    End If
End If
'***Fin Modificado por ELRO el 20130131******************

If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
    Me.txtMonto.value = 0
    Me.lblNroDoc.Caption = ""
    Me.lblNombreIF.Caption = ""
End If

ValidaTasaInteres

'RIRO20140430 ERS017 **************************
If nOperacion = gAhoApeLoteTransfBanco Or nOperacion = gPFApeLoteTransf Or nOperacion = gCTSApeLoteTransfNew Then
    cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, Trim(Right(cboMoneda.Text, 5)))
    SetDatosTransferencia "", "", "", 0, -1, ""
End If
'END RIRO *************************************
limpiarDatosCheque 'CTI6 20210607
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    '***Modificado por ELRO el 20121205, según OYP-RFC101-2012
    'cboTipoTasa.SetFocus
    If cboTipoTasa.Enabled Then
        cboTipoTasa.SetFocus
    End If
    '***Fin Modificado por ELRO el 20121205*******************
End If
End Sub

Private Sub cboPeriodoCTS_Click()
lblDispCTS = Format$(CDbl(Trim(Right(cboPeriodoCTS.Text, 5))) * 100, "#,##0.00")
End Sub

Private Sub cboPeriodoCTS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtInstitucion.SetFocus
End If
End Sub

Private Sub cboPrograma_Click()
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim bOrdPag As Boolean
Dim nMonto As Double
Dim nPlazo As Long
Dim nTpoPrograma As Integer
Dim i As Integer

Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
'bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
'nMonto = txtMonto.value
'nTpoPrograma = 1  RIRO ERS017
nTpoPrograma = 0  'RIRO ERS017

If cboPrograma.Visible Then
    nTpoPrograma = CInt(Right(Trim(cboPrograma.Text), 2))
End If

lnTpoPrograma = nTpoPrograma

'JUEZ 20141010 ******************************************************
If nProducto <> gCapCTS Then
    Dim rsPar As ADODB.Recordset
    Set rsPar = clsDef.GetCapParametroNew(nProducto, CInt(Trim(Right(cboPrograma.Text, 2))))
    bParPersNat = rsPar!bPersNat
    bParPersJur = rsPar!bPersJur
    bParMonedaSol = rsPar!bMonSol
    bParMonedaDol = rsPar!bMonDol
    nParMontoMinSol = rsPar!nMontoMinApertSol
    nParMontoMinDol = rsPar!nMontoMinApertDol
    If nProducto = gCapAhorros Then
        nParOrdPag = rsPar!nOrdPago
    ElseIf nProducto = gCapPlazoFijo Then
        nParPlazoMin = rsPar!nPlazoMin
        nParPlazoMax = rsPar!nPlazoMax
        bParFormaRetFinPlazo = rsPar!bFormaRetFinPlazo
        bParFormaRetMensual = rsPar!bFormaRetMensual
        bParFormaRetIniPlazo = rsPar!bFormaRetInicioPlazo
    End If
End If
'END JUEZ ***********************************************************

If nTpoPrograma = 4 Then
    lblInst.Visible = True
    txtInstitucion.Visible = True
    lblInstEtq.Visible = True
Else
    lblInst.Visible = False
    txtInstitucion.Visible = False
    lblInstEtq.Visible = False
End If

'If chkTasaPreferencial.value = vbUnchecked Then
    If nProducto = gCapPlazoFijo Then
        If txtPlazo <> "" Then
            nPlazo = CLng(txtPlazo)
            nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
            lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
        End If
    ElseIf nProducto = gCapAhorros Then
        nTasaNominal = clsDef.GetCapTasaInteresCamp(nProducto, nTpoPrograma, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, bOrdPag, fnCampanaCod) 'JUEZ 20160420
        If nTasaNominal = 0 Then nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, bOrdPag, nTpoPrograma) 'JUEZ 20160420
        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    Else
        nTasaNominal = clsDef.GetCapTasaInteresCamp(nProducto, nTpoPrograma, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, False, fnCampanaCod) 'JUEZ 20160420
        If nTasaNominal = 0 Then nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma) 'JUEZ 20160420
        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    End If
'If cboPrograma.ListIndex = 6 Then
      
    'If cboPrograma.ListIndex = 6 Then RIRO ERS017 Comentado
    If nTpoPrograma = 6 Then 'RIRO ERS017 Agregado
        Me.txtInstitucion.Visible = True
        Me.lblInst.Visible = True
        Me.lblInstEtq.Visible = True
        chkExoITF.value = 1
        chkExoITF.Enabled = False
        Me.chkITFEfectivo.value = 1
        cboTipoExoneracion.ListIndex = IndiceListaCombo(cboTipoExoneracion, 3)
        cboTipoExoneracion.Enabled = False 'RIRO20140407 ERS017
        Me.cboInstConvDep.Visible = False
    Else
        Me.txtInstitucion.Visible = False
        Me.lblInst.Visible = False
        Me.lblInstEtq.Visible = False
        chkExoITF.value = 0
        chkExoITF.Enabled = True
        Me.chkITFEfectivo.value = 0
        cboTipoExoneracion.ListIndex = -1
    End If
    
    If nTpoPrograma = 0 Or nTpoPrograma = 5 Or nTpoPrograma = 6 Or nTpoPrograma = 8 Then
    '***Condición nTpoPrograma = 8 agregado por ELRO el 20130131, según TI-ERS020-2013
        chkRelConv.Visible = True
        If chkRelConv.value = 1 And nTpoPrograma <> 6 Then
            Me.cboInstConvDep.Visible = True
        End If
    Else
        chkRelConv.Visible = False
        chkRelConv.value = 0
        Me.cboInstConvDep.Visible = False
    End If

    ' *** RIRO20140331 ERS017 2014 ************************************
    If nProducto = gCapAhorros And nTpoPrograma = 8 Then
        chkRelConv.Visible = True
        chkRelConv.value = 1
        chkRelConv_Click

    ElseIf nProducto = gCapAhorros And nTpoPrograma = 6 Then
        chkITFEfectivo.value = 0
        lblITF.Caption = "0.00"
        chkITFEfectivo.Enabled = False
        chkITFEfectivo_Click
        'JUEZ 20141010 Comentado para nuevos parámetros **************
        'For i = 1 To grdCuenta.Rows - 1
        '    If grdCuenta.TextMatrix(i, 5) = "." Then
        '        grdCuenta.TextMatrix(i, 5) = "0"
        '    End If
        '    grdCuenta.TextMatrix(i, 12) = "0.00"
        'Next
        'grdCuenta.ColumnasAEditar = "X-1-X-X-4-X-6-X-8-X-X-X-X-X-X-X-X"
        'END JUEZ ****************************************************
        chkRelConv.Visible = False
        chkRelConv.value = 0
        ValidarLoteTitular
    Else
        chkITFEfectivo.Enabled = True
        txtMonto_Change
        'grdCuenta.ColumnasAEditar = "X-1-X-X-4-5-6-X-8-X-X-X-X-X-X-X-X" 'JUEZ 20141010 Comentado para nuevos parámetros
        chkRelConv.Visible = False
        chkRelConv.value = 0
        chkRelConv_Click
        If chkExoITF.value = 1 Then
            chkExoITF.value = 0
        Else
            chkExoITF_Click
        End If
        chkExoITF.Enabled = True
        
    End If
    ' *** END RIRO ****************************************************
    'JUEZ 20141010 Nuevos parámetros ****************************
    If nParOrdPag = 0 Or nParOrdPag = 1 Then
        For i = 1 To grdCuenta.Rows - 1
            If grdCuenta.TextMatrix(i, 1) = "" Then Exit For
            grdCuenta.TextMatrix(i, 5) = IIf(nParOrdPag = 1, "1", "0")
        Next
        grdCuenta.ColumnasAEditar = "X-1-X-X-4-X-6-X-8-X-X-X-X-X-X-X-X"
    Else
        grdCuenta.ColumnasAEditar = "X-1-X-X-4-5-6-X-8-X-X-X-X-X-X-X-X"
    End If
    'END JUEZ ***************************************************
'End If
'End If

Set clsDef = Nothing

DefineTasaGrid nTpoPrograma
End Sub

Private Sub cboTipoTasa_Click()
nTipoTasa = CLng(Right(cboTipoTasa.Text, 4))
'***Modificado por ELRO el 20130131, según TI-ERS020-2013
'DefineTasaGrid
If nProducto <> gCapCTS Then
    'DefineTasaGrid (CInt(Right(Trim(cboPrograma.Text), 2)))
    DefineTasaGrid (Val(Right(Trim(cboPrograma.Text), 2))) 'EJVG20140408
End If
'***Fin Modificado por ELRO el 20130131******************
End Sub

Private Sub cboTipoTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cboPeriodoCTS.Visible Then
        cboPeriodoCTS.SetFocus
    Else
        cmdAgregar.SetFocus
    End If
End If
End Sub

'RIRO20140430 ERS017 **************************
Private Sub cboTransferMoneda_Click()
    If nmoneda = gMonedaNacional Then
        lblMonTra.BackColor = &HC0FFFF
        '''lblSimTra.Caption = "S/." 'marg ers044-2016
        lblSimTra.Caption = gcPEN_SIMBOLO 'marg ers044-2016
    ElseIf nmoneda = gMonedaExtranjera Then
        lblMonTra.BackColor = &HC0FFC0
        lblSimTra.Caption = "US$"
    End If
End Sub
'END RIRO *************************************

Private Sub chkExoITF_Click()
    Dim lnI As Integer
    Dim lnMontoCta As Currency
    
    Me.cboTipoExoneracion.ListIndex = -1
    If gbITFAplica Or (nProducto = gCapAhorros Or nProducto = gCapPlazoFijo) Then
        If chkExoITF.value = 1 Then
            Me.cboTipoExoneracion.Enabled = True
            For lnI = 1 To Me.grdCuenta.Rows - 1
                grdCuenta.TextMatrix(lnI, 12) = Format(0, "#,##0.00")
            Next lnI
        Else
            Me.cboTipoExoneracion.Enabled = False
            For lnI = 1 To Me.grdCuenta.Rows - 1
                If grdCuenta.TextMatrix(lnI, 8) = "" Then
                    lnMontoCta = 0
                Else
                    lnMontoCta = CCur(grdCuenta.TextMatrix(lnI, 8))
                End If
            
                If IsNumeric(grdCuenta.TextMatrix(lnI, 8)) And lnMontoCta >= 1000 Then
                    grdCuenta.TextMatrix(lnI, 12) = Format(fgITFCalculaImpuesto(CCur(grdCuenta.TextMatrix(lnI, 8))), "#,##0.00")
                    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(grdCuenta.TextMatrix(lnI, 12)))
                    If nRedondeoITF > 0 Then
                        grdCuenta.TextMatrix(lnI, 12) = Format(CCur(grdCuenta.TextMatrix(lnI, 12)) - nRedondeoITF, "#,##0.00")
                    End If
                Else
                    grdCuenta.TextMatrix(lnI, 12) = Format(0, "#,##0.00")
                End If
            Next lnI
        End If
    End If
    
    txtMonto_Change
End Sub

Private Sub chkITFEfectivo_Click()
    If chkITFEfectivo.value = 1 Then
        Me.lblTotal.Caption = Format(Me.txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00") 'RIRO ERS017 Comentado
        'Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00") ' RIRO ERS017 Agregado
    Else
        Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00") 'RIRO ERS017 Comentado
        'Me.lblTotal.Caption = Format(Me.txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00") ' RIRO ERS017 Agregado
    End If
End Sub

Private Sub chkRelConv_Click()
    If lnTpoPrograma <> 6 Then
        If chkRelConv.value = 1 Then
            cboInstConvDep.Visible = True
        Else
            cboInstConvDep.Visible = False
        End If
    Else
        If Not ValidaInstConv(txtInstitucion.Text) And txtInstitucion.Text <> "" Then
            MsgBox "La Institucion no esta para convenio de Depositos", vbInformation, "SISTEMA"
            txtInstitucion.Text = ""
            lblInst.Caption = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdAgregar_Click()
'***Agregado por ELRO el 20121205, según OYP-RFC101-2012
If nProducto = gCapCTS Then
    If Trim(txtInstitucion) = "" Then
        MsgBox "Debe seleccionar la Empresa Empleadora.", vbInformation, "¡Aviso!"
        Exit Sub
    End If
End If
'***Fin Agregado por ELRO el 20121205*******************

'RIRO20140407 ERS017 ***********************************
If bCargaLote Then
    If MsgBox("Al usar esta opción, se limpiará el Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        LimpiarGrdCuenta
        txtArchivo.Text = ""
        bCargaLote = False
        txtMonto.Text = "0.00"
        lblITF.Caption = "0.00"
        lblTotal.Caption = "0.00"
        grdCuenta.ColumnasAEditar = "X-1-X-X-4-5-6-X-8-X-X-X-X-X-X-X-X"
    Else
        Exit Sub
    End If
End If
'END RIRO **********************************************

'EJVG20140219 ***
If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
    If Len(Trim(lblNroDoc.Caption)) = 0 Then
        MsgBox "Ud. debe seleccionar primero el Cheque", vbInformation, "Aviso"
        If cmdDocumento.Visible And cmdDocumento.Enabled Then cmdDocumento.SetFocus
        Exit Sub
    End If
    'If IIf(grdCuenta.TextMatrix(1, 0) = "", 0, grdCuenta.Rows - 1) >= oDocRec.fnNroCliLote Then
    '    MsgBox "No se puede agregar más clientes, ya que se ha alcanzado el máximo nro. de clientes registrados en el Cheque", vbInformation, "Aviso"
    '    Exit Sub
    'End If
End If
'END EJVG ******
Dim nFila As Long
grdCuenta.Col = 1 'RIRO20140430 ERS017
grdCuenta_RowColChange 'RIRO20140430 ERS017
grdCuenta.AdicionaFila
nFila = grdCuenta.Rows - 1
grdCuenta.SetFocus
SendKeys "{Enter}"
grdCuenta.TextMatrix(nFila, 3) = "TI"
grdCuenta.TextMatrix(nFila, 5) = IIf(nParOrdPag = 1, "1", "0") 'JUEZ 20141010
'***Modificado por ELRO el 20121205, según OYP-RFC101-2012
'grdCuenta.TextMatrix(nFila, 7) = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
grdCuenta.TextMatrix(nFila, 7) = IIf(nProducto = gCapCTS, Format$(ConvierteTNAaTEA(0), "#,##0.00"), Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00"))
'***Fin Modificado por ELRO el 20121205*******************
grdCuenta.TextMatrix(nFila, 8) = "0.00"
grdCuenta.TextMatrix(nFila, 12) = "0.00"
grdCuenta.TextMatrix(nFila, 13) = "0.00"
grdCuenta.TextMatrix(nFila, 14) = "0.00"
grdCuenta.TextMatrix(nFila, 11) = nTasaNominal
grdCuenta.TextMatrix(nFila, 21) = fnCampanaCod 'JUEZ 20160420
cmdEliminar.Enabled = True
End Sub

Private Sub cmdBusca_Click()
    
    On Error GoTo error_handler
    
    txtArchivo.Text = Empty
    
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
        'TxtFecha.SetFocus
        '***Agregado por ELRO el 20121203, según OYP-RFC101-2012
        If nProducto = gCapCTS Then
            txtArchivo.Locked = True
            cmdAgregar.Enabled = False
        End If
        '***Agregado por ELRO el 20121203***********************
    Else
        txtArchivo.Text = "NO SE ABRIO NINGUN ARCHIVO"
        Exit Sub
    End If
    cmdCargar.Enabled = True
    
     Exit Sub
error_handler:
    
    If err.Number = 32755 Then
        'MsgBox "Se ha cancelado formulario", vbInformation, "Aviso"
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        MsgBox "Error al momento de seleccionar el archivo", vbCritical, "Aviso"
    End If
End Sub

Private Sub cmdCancelar_Click()
ClearScreen
End Sub

Private Sub cmdCargar_Click()
     '***Agregado por ELRO el 20121129, según OYP-RFC101-2012

'RIRO20140407 ERS017 ************************
If grdCuenta.Rows >= 2 And Len(Trim(grdCuenta.TextMatrix(1, 1))) > 0 Then
    If MsgBox("Al cargar la trama, se limpiaran los registros del Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    Else
        LimpiarGrdCuenta
    End If
End If
'END RIRO ***********************************

If nProducto = gCapCTS Then

    If Trim(txtInstitucion) = "" Then
        MsgBox "Debe seleccionar la Empresa Empleadora.", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    If InStr(Trim(txtArchivo), ".xls") = 0 And InStr(Trim(txtArchivo), ".xlsx") = 0 Then
        MsgBox "Debe seleccionar el archivo .xls o .xlsx para que sea cargado los datos."
        cmdBusca.SetFocus
       Exit Sub
    End If

    LimpiaFlex grdCuenta
    cboMoneda.Enabled = False
    cmdGrabar.Enabled = False
    
    Dim lsNroDoc2 As String
    Dim lsPersCod2 As String
    Dim lsNombre2 As String
    Dim lnPersoneria2 As Integer
    Dim lsTipDOI2 As String
    Dim lnMonApeCli2 As Double
    Dim lsDire2 As String
    Dim bMayorEdadCTS As Boolean 'RIRO20140407 ERS017
    
    'Variable de tipo Aplicación de Excel
    Dim oExcel2 As Excel.Application
    Dim lnFila1, lnFila2, lnFilasFormato As Integer
    Dim lsMoneda As String
    Dim lnMonedaSueldos As Integer
    Dim lnMontoSueldos As Currency
    Dim lbExisteCTS As Boolean
    Dim lbExisteError As Boolean

  
    'Una variable de tipo Libro de Excel
    Dim oLibro As Excel.Workbook
    Dim oHoja As Excel.Worksheet

    'creamos un nuevo objeto excel
    Set oExcel2 = New Excel.Application
     
    lnFilasFormato = 1036
    bCargaLote = True 'RIRO20140407 ERS017
     
    'Usamos el método open para abrir el archivo que está en el directorio del programa llamado archivo.xls
    Set oLibro = oExcel2.Workbooks.Open(txtArchivo)
    
    'Hacemos referencia a la Hoja
    Set oHoja = oLibro.Sheets(1)
    
    'Hacemos el Excel Visible
    'oLibro.Visible = False
    
    grdCuenta.lbEditarFlex = True
    
    lsMoneda = oHoja.Cells(27, 4)
    
    cboMoneda.ListIndex = 0
    If Trim(Left(cboMoneda.Text, 10)) <> UCase(lsMoneda) Then
        cboMoneda.ListIndex = 1
    End If
    
    With oHoja
        For lnFila1 = 36 To lnFilasFormato
            lsTipDOI2 = .Cells(lnFila1, 2)
            lsNroDoc2 = .Cells(lnFila1, 3)
            lnMonApeCli2 = .Cells(lnFila1, 10)
            lnMonedaSueldos = IIf(UCase(.Cells(lnFila1, 11)) = "SOLES", 1, 2)
            lnMontoSueldos = CCur(.Cells(lnFila1, 12))
            If Len(lsNroDoc2) > 0 Then
                '***Agrega nueva fila****
                grdCuenta.AdicionaFila
                lnFila2 = grdCuenta.row
                    grdCuenta.TextMatrix(lnFila2, 1) = ""
                    grdCuenta.TextMatrix(lnFila2, 2) = ""
                    grdCuenta.TextMatrix(lnFila2, 3) = "TI"
                    grdCuenta.TextMatrix(lnFila2, 8) = "0.00"
                    grdCuenta_OnCellChange CInt(lnFila2), 8
                    grdCuenta.TextMatrix(lnFila2, 7) = "0.00"
                    grdCuenta.TextMatrix(lnFila2, 9) = ""
                    grdCuenta.TextMatrix(lnFila2, 11) = "0.00"
                    grdCuenta.TextMatrix(lnFila2, 12) = "0.00"
                    grdCuenta.TextMatrix(lnFila2, 15) = ""
                    grdCuenta.TextMatrix(lnFila2, 16) = ""
                    grdCuenta.TextMatrix(lnFila2, 17) = ""
                    grdCuenta.TextMatrix(lnFila2, 18) = ""
                    grdCuenta.TextMatrix(lnFila2, 19) = "0.00"
                    grdCuenta.TextMatrix(lnFila2, 21) = 0 'JUEZ 20160420
                '***Fin Agrega nueva fila
                                
                DevuelveDatosCliente lsNroDoc2, lsPersCod2, lsNombre2, lnPersoneria2, lsDire2, bMayorEdadCTS 'RIRO20140430 ERS017
                
                If lsPersCod2 <> "" Then
                    grdCuenta.TextMatrix(lnFila2, 1) = lsPersCod2
                    grdCuenta.TextMatrix(lnFila2, 2) = lsNombre2
                    grdCuenta.TextMatrix(lnFila2, 3) = "TI"
                    grdCuenta.TextMatrix(lnFila2, 8) = Format$(lnMonApeCli2, "###,##0.00")
                    grdCuenta_OnCellChange CInt(lnFila2), 8
                    grdCuenta.TextMatrix(lnFila2, 7) = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
                    grdCuenta.TextMatrix(lnFila2, 9) = lnPersoneria2
                    grdCuenta.TextMatrix(lnFila2, 11) = nTasaNominal
                    grdCuenta.TextMatrix(lnFila2, 15) = lsNroDoc2
                    grdCuenta.TextMatrix(lnFila2, 16) = lsDire2
                    grdCuenta.TextMatrix(lnFila2, 17) = nTpoProgramaCTS
                    grdCuenta.TextMatrix(lnFila2, 18) = lnMonedaSueldos
                    grdCuenta.TextMatrix(lnFila2, 19) = lnMontoSueldos
                    grdCuenta.TextMatrix(lnFila2, 21) = fnCampanaCod 'JUEZ 20160420
                Else
                    .Range("A" & lnFila1, "M" & lnFila1).Interior.Color = RGB(255, 255, 0)
                    .Cells(lnFila1, 1) = "LA PERSONA NO EXISTE EN EL SISTEMA."
                    lbExisteError = True
                End If
          End If
        Next lnFila1
    End With
    
    grdCuenta.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
    
    If lbExisteError = False Then
        cmdGrabar.Enabled = True
    Else
        oExcel2.Visible = True
        Exit Sub
    End If
    
    Set oHoja = Nothing
    Set oLibro = Nothing
    oExcel2.Quit
    Set oExcel2 = Nothing

Else
'***Fin Agregado por ELRO el 20121129*******************
    Dim lsNroDoc As String
    Dim lsPersCod As String
    Dim lsNombre As String
    Dim lnPersoneria As Integer
    Dim lnOP As Integer
    Dim lnTasaCli As Double
    Dim lnMonApeCli As Double
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, fila As Integer
    Dim psArchivoAGrabar As String
    'Dim fs As New Scripting.FileSystemObject
    Dim sCad As String
    Dim nFila As Long
    Dim lsNomArch As String
    Dim lsDire As String
    Dim lsTipDOI As String
    Dim X As Integer
    
    'RIRO20140407 ERS017 *****************
    Dim Y As Integer, Z As Integer
    Dim bMayorEdad As Boolean, bFormato As Boolean
    Dim psArchivoAGrabarMenores As String
    Dim psArchivoAGrabarPersJurid As String
    
    Dim oBookMenores As Object
    Dim oSheetMenores As Object
    
    Dim oBookPersJurid As Object
    Dim oSheetPersJurid As Object
    Dim bValidTrama As Boolean
    'END RIRO ****************************
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    
        If txtArchivo.Text = "" Then
            MsgBox "No selecciono ningun archivo", vbExclamation, "Aviso"
            Exit Sub
        End If
    
        Set objExcel = New Excel.Application
        Set xLibro = objExcel.Workbooks.Open(txtArchivo.Text)
        psArchivoAGrabar = App.Path & "\SPOOLER\No_ClientesSinRegistrar_" & Format(gdFecSis, "yyyymmdd") & ".xls"
        psArchivoAGrabarMenores = App.Path & "\SPOOLER\clientesMenoresEdad_" & Format(CDate(gdFecSis), "yyyyMMdd") & ".xls" 'RIRO20140407 ERS017
        psArchivoAGrabarPersJurid = App.Path & "\SPOOLER\clientesPersonaJuridica_" & Format(CDate(gdFecSis), "yyyyMMdd") & ".xls" 'RIRO20140407 ERS017
        
        grdCuenta.SetFocus
        'SendKeys "{Enter}" RIRO ERS017 Comentado
                    
        cmdEliminar.Enabled = True
        X = 1
        Y = 1: Z = 1 'RIRO20140407 ERS017
        
        If Dir(psArchivoAGrabar) <> "" Then
            Kill psArchivoAGrabar
        End If

        'RIRO20140407 ERS017 ***************************
        If Dir(psArchivoAGrabarMenores) <> "" Then
            Kill psArchivoAGrabarMenores
        End If
        If Dir(psArchivoAGrabarPersJurid) <> "" Then
            Kill psArchivoAGrabarPersJurid
        End If
        'END RIRO **************************************
    
       'Start a new workbook in Excel
       Set oExcel = CreateObject("Excel.Application")
       Set oBook = oExcel.Workbooks.Add
    
       'Add data to cells of the first worksheet in the new workbook
       Set oSheet = oBook.Worksheets(1)
       
       'oBook.SaveAs psArchivoAGrabar
        
        'Creacion del Archivo
        'Open psArchivoAGrabar For Output As #
    '    Dim ArcSal As Integer
    '    ArcSal = FreeFile
    '    Open psArchivoAGrabar For Output As ArcSal
    '
    '    Print #ArcSal, "Clientes no registrados en el sistema"
    '    sCad = ""
    
       'RIRO20140430 ERS017 *******************************
        Set oBookMenores = oExcel.Workbooks.Add
        Set oSheetMenores = oBookMenores.Worksheets(1)

        Set oBookPersJurid = oExcel.Workbooks.Add
        Set oSheetPersJurid = oBookPersJurid.Worksheets(1)
        
        bFormato = True
        bCargaLote = True
        'grdCuenta.ColumnasAEditar = "X-1-X-X-4-X-6-X-X-X-X-X-X-X-X-X-X"
        grdCuenta.ColumnasAEditar = "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 1))) <> "CLIENTE" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 2))) <> "TIPO DOI" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 3))) <> "Nº D.O.I." Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 4))) <> "OP" Then bFormato = False
        If UCase(Trim(xLibro.Sheets(1).Cells(1, 5))) <> "MONTO S/." Then bFormato = False
        If bFormato = False Then
            MsgBox "El archivo seleccionado no tiene el formato adecuado para la carga en lote, verifíquelo e inténtelo de nuevo", vbInformation, "Aviso"
            If Not objExcel Is Nothing Then
                objExcel.Workbooks.Close
                Set objExcel = Nothing
            End If
            If Not oExcel Is Nothing Then
                oExcel.Workbooks.Close
                Set oExcel = Nothing
            End If
            Exit Sub
        End If
        If Len(Trim(xLibro.Sheets(1).Cells(1002, 1))) > 0 Then
            MsgBox "La cantidad de aperturas supera el límite de 1000", vbInformation, "Aviso"
            If Not objExcel Is Nothing Then
                objExcel.Workbooks.Close
                Set objExcel = Nothing
            End If
            If Not oExcel Is Nothing Then
                oExcel.Workbooks.Close
                Set oExcel = Nothing
            End If
            Exit Sub
        End If
       'END RIRO ******************************************
    
        With xLibro
            With .Sheets(1)
                For fila = 2 To 1002
                    Evento
                    bValidTrama = True
                    
                    lsNomArch = .Cells(fila, 1)
                    'Tipo de DOI
                    If IsNumeric(.Cells(fila, 2)) Then
                        lsTipDOI = .Cells(fila, 2)
                    Else
                        bValidTrama = False
                    End If
                    
                    lsNroDoc = .Cells(fila, 3)
                    
                    'Orden de Pago
                    If IsNumeric(.Cells(fila, 4)) Then
                        lnOP = .Cells(fila, 4)
                    Else
                        bValidTrama = False
                    End If
                    'Monto de Apertura.
                    If IsNumeric(.Cells(fila, 5)) Then
                        lnMonApeCli = .Cells(fila, 5)
                    Else
                        bValidTrama = False
                    End If
                    If Not bValidTrama Then
                        MsgBox "Los datos ingresados en la trama no son los adecuados.", vbExclamation, "Aviso"
                        LimpiarGrdCuenta
                        Exit Sub
                    End If
                    
                    If lsNroDoc <> "" Then
                        DevuelveDatosCliente lsNroDoc, lsPersCod, lsNombre, lnPersoneria, lsDire, bMayorEdad 'RIRO20140430 ERS017, Se agrego parametro bMayorEdad
                        
                        If lsPersCod <> "" Then
                            grdCuenta.AdicionaFila
                            nFila = grdCuenta.Rows - 1
                            grdCuenta.TextMatrix(nFila, 0) = nFila
                            grdCuenta.TextMatrix(nFila, 1) = lsPersCod
                            grdCuenta.TextMatrix(nFila, 2) = lsNombre
                            grdCuenta.TextMatrix(nFila, 3) = "TI"
                            'grdCuenta.TextMatrix(nFila, 5) = lnOP
                            grdCuenta.TextMatrix(nFila, 5) = IIf(nParOrdPag = 1, "1", IIf(nParOrdPag = 0, "0", lnOP)) 'JUEZ 20141010
                            'grdCuenta.TextMatrix(nFila, 7) = Format$(lnTasaCli, "#,##0.00")
                            grdCuenta.TextMatrix(nFila, 7) = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
                            grdCuenta.TextMatrix(nFila, 8) = Format$(lnMonApeCli, "###,##0.00")
                            grdCuenta.TextMatrix(nFila, 9) = lnPersoneria
                            'grdCuenta.TextMatrix(nFila, 11) = Format$(ConvierteTEAaTNA(lnTasaCli), "#,##0.00")
                            grdCuenta.TextMatrix(nFila, 11) = nTasaNominal
                            grdCuenta.TextMatrix(nFila, 15) = lsNroDoc
                            grdCuenta.TextMatrix(nFila, 16) = lsDire
                            grdCuenta.TextMatrix(nFila, 20) = lsTipDOI 'RIRO ERS017
                            grdCuenta.TextMatrix(nFila, 21) = fnCampanaCod 'JUEZ 20160420
                            
                            If nProducto = gCapAhorros Or nProducto = gCapPlazoFijo Then
                                If Me.chkExoITF.value <> 1 And lnMonApeCli >= 1000 And (nProducto = gCapAhorros And lnTpoPrograma <> 6) Then 'RIRO20140407 ERS017 Se agrego condicion "And lnTpoPrograma <> 6"
                                    grdCuenta.TextMatrix(nFila, 12) = Format(fgITFCalculaImpuesto(CCur(lnMonApeCli)), "#,##0.00")
                                    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(grdCuenta.TextMatrix(nFila, 12)))
                                    If nRedondeoITF > 0 Then
                                        grdCuenta.TextMatrix(nFila, 12) = Format(CCur(grdCuenta.TextMatrix(nFila, 12)) - nRedondeoITF, "#,##0.00")
                                    End If
                                Else
                                    grdCuenta.TextMatrix(nFila, 12) = "0.00" '0#
                                End If
                            End If
                            
                            'RIRO20140430 ERS017 Se agrego "gAhoApeLoteTransfBanco"
                            If nOperacion = gAhoApeLoteEfec Or nOperacion = gPFApeLoteEfec Or nOperacion = gCTSApeLoteEfec Or nOperacion = gAhoApeLoteTransfBanco _
                                Or nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Or nOperacion = gCTSApeLoteTransfNew Or nOperacion = gPFApeLoteTransf Then
                                txtMonto.Text = Format$(grdCuenta.SumaRow(8), "#,##0.00")
                            End If

                            'RIRO20140407 ERS017 ********************************
                            If (Not bMayorEdad) And lnPersoneria = 1 Then
                                Y = Y + 1
                                oSheetMenores.Range("A1:F1").Font.Bold = True
                                oSheetMenores.Columns("A:A").ColumnWidth = 56
                                oSheetMenores.Columns("C:C").NumberFormat = "@"
                                
                                oSheetMenores.Range("A1").value = "Cliente"
                                oSheetMenores.Range("B1").value = "Tipo DOI"
                                oSheetMenores.Range("C1").value = "N° D.N.I"
                                                                
                                oSheetMenores.Range("A" & Y).value = lsNomArch
                                oSheetMenores.Range("B" & Y).value = lsTipDOI
                                oSheetMenores.Range("C" & Y).value = lsNroDoc
                            End If
                            If lnPersoneria > 1 Then
                                Z = Z + 1
                                oSheetPersJurid.Range("A1:F1").Font.Bold = True
                                oSheetPersJurid.Columns("A:A").ColumnWidth = 56
                                oSheetPersJurid.Columns("C:C").NumberFormat = "@"
                                
                                oSheetPersJurid.Range("A1").value = "Cliente"
                                oSheetPersJurid.Range("B1").value = "Tipo DOI"
                                oSheetPersJurid.Range("C1").value = "N° D.N.I"
                                                  
                                oSheetPersJurid.Range("A" & Z).value = lsNomArch
                                oSheetPersJurid.Range("B" & Z).value = lsTipDOI
                                oSheetPersJurid.Range("C" & Z).value = lsNroDoc
                            End If
                            'END RIRO *******************************************

                        Else
                            X = X + 1
                            'sCad = lsNomArch & Space(50 - Len(lsNomArch)) & "|" & lsTipDOI & "|" & lsNroDoc & "|" & CStr(lnOP) & "|" & CStr(lnTasaCli) & "|" & CStr(lnMonApeCli)
                            oSheet.Range("A1:F1").Font.Bold = True
                            oSheet.Columns("A:A").ColumnWidth = 56 'RIRO ERS017
                            oSheet.Columns("D:D").ColumnWidth = 5.29 'RIRO ERS017
                            oSheet.Columns("C:C").NumberFormat = "@" 'RIRO ERS017
                            oSheet.Range("A1").value = "Cliente"
                            oSheet.Range("B1").value = "Tipo DOI"
                            oSheet.Range("C1").value = "N° D.N.I"
                            oSheet.Range("D1").value = "OP"
                            oSheet.Range("E1").value = "MONTO S/."
                                
                            oSheet.Range("A" & X).value = lsNomArch
                            oSheet.Range("B" & X).value = lsTipDOI
                            oSheet.Range("C" & X).value = "'" & lsNroDoc
                            oSheet.Range("D" & X).value = CStr(lnOP)
                            oSheet.Range("E" & X).value = CStr(lnMonApeCli)
                            
                            'Print #1, sCad; ""
                        End If
                    Else
                        Exit For
                    End If
                Next
    
            End With
        End With
        'RIRO20130430 ERS017 COMENTADO ************************************************
        ''Eliminamos los objetos si ya no los usamos
        'objExcel.Quit
        'Set objExcel = Nothing
        'Set xLibro = Nothing
        '
        ''Close ArcSal
        ''Save the Workbook and Quit Excel
        'oBook.SaveAs psArchivoAGrabar
        'Set oBook = Nothing
        'oExcel.Quit
        '
        'Dim m_Excel As New Excel.Application
        'm_Excel.Workbooks.Open (psArchivoAGrabar)
        'm_Excel.Visible = True
        'END RIRO ERS017 COMENTADO ****************************************************
          
        
        'RIRO20140430 ERS017 **********************************************************
        
        'Eliminamos los objetos si ya no los usamos
        objExcel.Quit
        Set objExcel = Nothing
        Set xLibro = Nothing
        
        If X > 1 Then
            oBook.SaveAs psArchivoAGrabar
            If MsgBox("Existen " & (X - 1) & " clientes no registrados, el proceso no puede continuar, ¿desea mostrar la lista de los clientes pendientes de registro?", vbQuestion + vbYesNo, "Aviso") = vbYes Then 'RIRO20140407 ERS017
                Dim m_Excel As New Excel.Application
                m_Excel.Workbooks.Open (psArchivoAGrabar)
                m_Excel.Visible = True
            End If
            LimpiarGrdCuenta
        Else
            If lnTpoPrograma = 6 Then
            ValidarLoteTitular
            End If
        End If
        If Y > 1 Then
            oBookMenores.SaveAs psArchivoAGrabarMenores
            If MsgBox("Existen " & (Y - 1) & " clientes menores de edad, el proceso no puede continuar, ¿desea mostrar la lista de los clientes menores de edad?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                Dim m_ExcelMenores As New Excel.Application
                m_ExcelMenores.Workbooks.Open (psArchivoAGrabarMenores)
                m_ExcelMenores.Visible = True
            End If
            LimpiarGrdCuenta
        End If
        If Z > 1 Then
            oBookPersJurid.SaveAs psArchivoAGrabarPersJurid
            If MsgBox("No se puede realizar la apertura de " & (Z - 1) & " clientes por ser personas juridicas ¿Desea mostrar la lista de los clientes con personería jurídica?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                Dim m_ExcelPersJurid As New Excel.Application
                m_ExcelPersJurid.Workbooks.Open (psArchivoAGrabarPersJurid)
                m_ExcelPersJurid.Visible = True
            End If
            LimpiarGrdCuenta
        End If
                
        Set oBook = Nothing
        Set m_Excel = Nothing
        Set m_ExcelMenores = Nothing
        Set m_ExcelPersJurid = Nothing
        oExcel.Quit
        'END RIRO *********************************************************************
        
    '    If sCad <> "" Then
    '        'MsgBox "Hay Clientes no registrados en el sistema", vbInformation, "MENSAJE DEL SISTEMA"
    '        'Dim strArchivo As String
    '        'strArchivo = Dir(psArchivoAGrabar)
    '        Shell "NotePad " & psArchivoAGrabar, vbMaximizedFocus
    '        'strArchivo = Dir
    '    End If
End If
End Sub

'RIRO20140407 ERS017 2014 ****************************************
'Funcion valida si los clientes del grid poseen
'cuentas caja sueldo vigentes
Private Function ValidarLoteTitular() As Boolean

Dim slCuentas() As String
Dim slTmp() As String
Dim sCodEmpleador As String
Dim nmoneda As Integer
Dim sPersCod As String
Dim sRuta As String
Dim i As Integer, J As Integer, nCuentas As Integer

sCodEmpleador = txtInstitucion.Text
nmoneda = Trim(Right(cboMoneda.Text, 5))
Dim oPer As COMDPersona.DCOMPersonas
Set oPer = New COMDPersona.DCOMPersonas

pvCuentasSueldo.Min = 1
pvCuentasSueldo.Max = grdCuenta.Rows
pvCuentasSueldo.value = 1
Evento
lblCuentaSueldo.Caption = "Verifica cuentas caja sueldo:"
pvCuentasSueldo.Visible = True
lblCuentaSueldo.Visible = True

For i = 1 To grdCuenta.Rows - 1
    Evento
    sPersCod = grdCuenta.TextMatrix(i, 1)
    If oPer.ValidaTitularSueldo(sCodEmpleador, nmoneda, sPersCod) > 0 Then
        nCuentas = nCuentas + 1
        ReDim Preserve slCuentas(nCuentas)
        slCuentas(nCuentas) = slCuentas(nCuentas) & grdCuenta.TextMatrix(i, 2) & "|"
        slCuentas(nCuentas) = slCuentas(nCuentas) & grdCuenta.TextMatrix(i, 20) & "|"
        slCuentas(nCuentas) = slCuentas(nCuentas) & grdCuenta.TextMatrix(i, 15) & "|"
        slCuentas(nCuentas) = slCuentas(nCuentas) & IIf(grdCuenta.TextMatrix(i, 5) = "", 0, 1) & "|"
        slCuentas(nCuentas) = slCuentas(nCuentas) & grdCuenta.TextMatrix(i, 7) & "|"
        slCuentas(nCuentas) = slCuentas(nCuentas) & grdCuenta.TextMatrix(i, 8)
    End If
    If pvCuentasSueldo.Max < i Then
        pvCuentasSueldo.value = pvCuentasSueldo.Max
    Else
        pvCuentasSueldo.value = i
    End If
Next
pvCuentasSueldo.value = pvCuentasSueldo.Max
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

If nCuentas > 0 Then
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    oSheet.Range("A1:F1").Font.Bold = True
    oSheet.Columns("A:A").ColumnWidth = 56
    oSheet.Columns("D:D").ColumnWidth = 5.29
    oSheet.Columns("C:C").NumberFormat = "@"
    
    oSheet.Range("A1").value = "Cliente"
    'oSheet.Range("A1").HorizontalAlignment = xlCenter
    'HorizontalAlignment = xlCenter
    oSheet.Range("B1").value = "Tipo DOI"
    oSheet.Range("C1").value = "N° D.N.I"
    oSheet.Range("D1").value = "OP"
    oSheet.Range("E1").value = "Tasa"
    oSheet.Range("F1").value = "Monto S/."
    
    For i = 1 To nCuentas
        Evento
        slTmp() = Split(slCuentas(i), "|")
        oSheet.Range("A" & (i + 1)).value = slTmp(0)
        oSheet.Range("B" & (i + 1)).value = slTmp(1)
        oSheet.Range("C" & (i + 1)).value = slTmp(2)
        oSheet.Range("D" & (i + 1)).value = slTmp(3)
        oSheet.Range("E" & (i + 1)).value = slTmp(4)
        oSheet.Range("F" & (i + 1)).value = slTmp(5)
    Next
    sRuta = App.Path & "\SPOOLER\ClientesConCuentaSueldo" & Format(Now, "yyyyMMddhhnnss") & ".xls"
    If Dir(sRuta) <> "" Then
        Kill sRuta
    End If
    oBook.SaveAs sRuta
    Dim m_Excel As New Excel.Application
    m_Excel.Workbooks.Open (sRuta)
    m_Excel.Visible = True
    oExcel.Quit
    ValidarLoteTitular = False
Else
    ValidarLoteTitular = True
End If

pvCuentasSueldo.Visible = False
lblCuentaSueldo.Visible = False

End Function
'END RIRO

Private Sub cmdDocumento_Click()
'EJVG20140408 ***
'frmCapAperturaListaChq.Inicia frmCapAperturasLote, nOperacion, nmoneda, nProducto
    Dim oForm As New frmChequeBusqueda
    Dim lnOperacion As TipoOperacionCheque
    Dim oDocRecTmp As UDocRec

    On Error GoTo ErrCargaDocumento
    If nOperacion = gAhoApeLoteChq Then
        lnOperacion = AHO_AperturaLote
    ElseIf nOperacion = gPFApeLoteChq Then
        lnOperacion = DPF_AperturaLote
    ElseIf nOperacion = gCTSApeLoteChq Then
        lnOperacion = Ninguno
    Else
        lnOperacion = Ninguno
    End If

    Set oDocRecTmp = oForm.Iniciar(nmoneda, lnOperacion)
    Set oForm = Nothing
    
    If Len(Trim(oDocRecTmp.fsNroDoc)) = 0 Then Exit Sub
                                              
                                        
                              
    
    Set oDocRec = oDocRecTmp
    setDatosCheque

    Exit Sub
ErrCargaDocumento:
    MsgBox "Ha sucedido un error al cargar los datos del Documento", vbCritical, "Aviso"
'END EJVG *******
End Sub

Private Sub CmdEliminar_Click()
Dim nFila As Long
nFila = grdCuenta.row
If bCargaLote Then Exit Sub 'RIRO20140430 ERS017
If MsgBox("¿Desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    grdCuenta.EliminaFila nFila
    If nOperacion <> gAhoApeLoteChq And nOperacion <> gPFApeLoteChq And nOperacion <> gCTSApeLoteChq Then
        txtMonto.Text = Format$(grdCuenta.SumaRow(8), "#,##0.00")
    End If
End If
End Sub

Private Sub cmdFormato_Click()

'Dim m_Excel As New Excel.Application
'Dim oBook As Object
'm_Excel.Workbooks.Open (App.path & "\FormatoCarta\AperturasLote.xlsx") ' & Format(gdFecSis, "yyyymmdd") & ".xls")
'Set oBook = m_Excel.Workbooks.Add
    
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nFila, i As Double
        
On Error GoTo error_handler
        
    dlgArchivo.FileName = Empty
    'dlgArchivo.Filter = "Archivo *.xlsx"
    dlgArchivo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx| Archivos de Excel (*.xls)|*.xls"
    dlgArchivo.FileName = "AperturaEnLote" & Format(Now, "yyyyMMddhhnnss") & ".xlsx"
    dlgArchivo.ShowSave
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    If fs.FileExists(dlgArchivo.FileName) Then
        
        MsgBox "El archivo '" & dlgArchivo.FileTitle & "' ya existe, debe asignarle un nombre diferente", vbExclamation, ""
        Exit Sub
    
    End If
        
    If nProducto = gCapAhorros Then
        lsArchivo = App.Path & "\FormatoCarta\AperturasLoteAhorro.xlsx"
        lsNomHoja = "Apertura en Lote"
    ElseIf nProducto = gCapCTS Then
        lsArchivo = App.Path & "\FormatoCarta\AperturasLoteCTS.xlsx"
        lsNomHoja = "Plantilla Deposito CTS"
    End If
    
    lsArchivo1 = dlgArchivo.FileName

    If fs.FileExists(lsArchivo) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(lsArchivo)
    Else
        MsgBox "No Existe Plantilla en la Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    
    For Each xlHoja1 In xlsLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
        
    xlHoja1.SaveAs lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
     
    
    MsgBox "Se exportó el formato para la carga de archivos", vbInformation, "Aviso"
               
Exit Sub
    
error_handler:
    
    If err.Number = 32755 Then
        'MsgBox "Se ha cancelado formulario", vbInformation, "Aviso"
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Error al momento de generar el archivo", vbCritical, "Aviso"
    End If
End Sub

'RIRO20140430 ERS017 **************************************
Private Function validarApertLoteTransferencia() As String

    Dim sValidacion As String
    Dim nRegistrosApertura As Integer
    Dim i As Integer
        
    nRegistrosApertura = grdCuenta.Rows - 1
    If Trim(grdCuenta.TextMatrix(nRegistrosApertura, 1)) = "" And Trim(grdCuenta.TextMatrix(nRegistrosApertura, 2)) = "" Then
        nRegistrosApertura = nRegistrosApertura - 1
    End If
    sValidacion = ""
    'Validando que grid tenga registros
    If nRegistrosApertura = 0 Then
        sValidacion = "El Grid de cuentas a aperturar, no tiene registrado ningun cliente." & vbNewLine
    End If
    'Validando la seleccion del voucher
    If fnMovNroRVD = 0 Then
        sValidacion = sValidacion & "Debe seleccionar un voucher para concretar la operación." & vbNewLine
    End If
    'Validando la cantidad de registros del voucher con la cantidad de registros del grid
    If nNroApertura <> nRegistrosApertura Then
        sValidacion = sValidacion & "El número de registros del voucher es diferente al numero de registros del grid." & vbNewLine
    End If
    'Validando la moneda del voucher y de la operación.
    If Trim(Right(cboMoneda.Text, 5)) <> Trim(Right(cboTransferMoneda.Text, 5)) Then
        sValidacion = sValidacion & "La moneda del voucher es diferente a la modeda de la operación." & vbNewLine
    End If
    If CDbl(lblMonTra.Caption) <> CDbl(lblTotal.Caption) Then
        sValidacion = sValidacion & "El Monto de Transacción debe ser igual al Monto Total." & vbNewLine
    End If
   
    validarApertLoteTransferencia = sValidacion
    
End Function
'END RIRO **************************************

Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
                
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
               
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII

Dim nMontoTotal As Double
Dim previo As New previo.clsprevio
Dim previo2 As New previo.clsprevio
Dim lsCadCartilla As String
Dim nTpoPrograma As Integer
Dim lsPersCodConv As String
Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim sPersLavDinero As String
Dim nMontoLavDinero As Double
Dim nTC As Double
Dim nMonto As Double
Dim loLavDinero As frmMovLavDinero
Dim sCuenta As String
Set loLavDinero = New frmMovLavDinero
Dim bImprimirCartillas As Boolean 'RIRO20140430 ERS017

lsCadCartilla = String(78000, "aa")

nTpoPrograma = -1
If cboPrograma.Visible Then
    nTpoPrograma = CInt(Trim(Right(Me.cboPrograma.Text, 3)))
End If

If Me.chkExoITF.value = 1 And fraITF.Visible Then
    If Me.cboTipoExoneracion.Text = "" Then
        MsgBox "Debe ingresa un tipo de Exoneracion.", vbInformation, "Aviso"
        cboTipoExoneracion.SetFocus
        Exit Sub
    End If
End If

If chkRelConv.value = 1 Then
    If Trim(Right(cboPrograma, 3)) <> "1" And Trim(Right(cboPrograma, 3)) <> "4" And _
       Trim(Right(cboPrograma, 3)) <> "7" Then
    '***Condición cboPrograma.ListIndex <> 6 modificado por ELRO el 20130201
        lsPersCodConv = Trim(Right(Me.cboInstConvDep.Text, 13))
    Else
        lsPersCodConv = txtInstitucion.Text
    End If
End If

nMontoTotal = txtMonto.value

If nProducto = gCapCTS Then
    If Me.txtInstitucion.Text = "" Then
        MsgBox "Debe Ingresar una Institucion Valida.", vbInformation, "Aviso"
        txtInstitucion.SetFocus
        Exit Sub
    End If
End If

If nProducto = gCapAhorros Then
    If cboPrograma.ListIndex = -1 Then
        MsgBox "Debe de Seleccionar un tipo de Sub Producto para AHORRO", vbInformation
    End If
    'If cboPrograma.ListIndex = 6 Then
    If nTpoPrograma = 6 Then 'APRI20190109 ERS077-2018 - MEJORA
        If Me.txtInstitucion.Text = "" Then
            MsgBox "Debe Ingresar una Institucion Valida.", vbInformation, "Aviso"
            txtInstitucion.SetFocus
            Exit Sub
        End If
    End If
    '***Agregado por ELRO el 20130201, según TI-ERS020-2013
    If Trim(Right(cboPrograma, 3)) = "8" Then
        If chkRelConv.value = 0 Then
            MsgBox "Debe seleccionar empresa convenio", vbInformation, "Aviso"
            chkRelConv.SetFocus
            Exit Sub
        Else
            If Trim(cboInstConvDep) = "" Then
                MsgBox "Debe seleccionar empresa convenio", vbInformation, "Aviso"
                cboInstConvDep.SetFocus
                Exit Sub
            End If
        End If
    End If
    '***Agregado por ELRO el 20130201**********************
End If

'***Modificado por ELRO el 20121129, según OYP-RFC101-2012
'If nMontoTotal <= 0 And (cboPrograma.ListIndex <> 6 And cboPrograma.ListIndex <> 5 And cboPrograma.ListIndex <> 0) Then
'    MsgBox "Debe registrar al menos una persona para la apertura", vbInformation, "Aviso"
'    cmdAgregar.SetFocus
'    Exit Sub
'End If
If nProducto <> gCapCTS Then
    'MIOL 20130506, SEGUN SATI INC1303140003 - SE AGREGO And cboPrograma.ListIndex <> 8
    'If nMontoTotal <= 0 And cboPrograma.ListIndex <> 6 And cboPrograma.ListIndex <> 5 And cboPrograma.ListIndex <> 0 And cboPrograma.ListIndex <> 8) Then
    If nMontoTotal <= 0 And (Right(cboPrograma.Text, 1) <> 6 And Right(cboPrograma.Text, 1) <> 5 And cboPrograma.ListIndex <> 0 And Right(cboPrograma.Text, 1) <> 8) Then 'APRI20170526, SEGUN SATI INC1705260003
        MsgBox "Debe registrar al menos una persona para la apertura", vbInformation, "Aviso"
        cmdAgregar.SetFocus
    Exit Sub
End If
End If
'***Modificado por ELRO el 20121129***********************


If Not ValidaDatosCuentas() Then Exit Sub

'JUEZ 20130723 *******************************************
If nProducto = gCapCTS Then
    Dim c As Integer, nCantExiste As Integer, MatLista As Variant
    Dim oCap As New COMNCaptaGenerales.NCOMCaptaGenerales
    ReDim MatLista(grdCuenta.Rows - 1, 6)
    For c = 1 To grdCuenta.Rows - 1
        If oCap.VerificarExisteCuentaCTS(grdCuenta.TextMatrix(c, 1), txtInstitucion.Text, CInt(Trim(Right(cboMoneda.Text, 1)))) Then
            nCantExiste = nCantExiste + 1
            MatLista(nCantExiste, 1) = grdCuenta.TextMatrix(c, 1)
            MatLista(nCantExiste, 2) = grdCuenta.TextMatrix(c, 2)
            MatLista(nCantExiste, 3) = grdCuenta.TextMatrix(c, 3)
            MatLista(nCantExiste, 4) = grdCuenta.TextMatrix(c, 7)
            MatLista(nCantExiste, 5) = grdCuenta.TextMatrix(c, 8)
            MatLista(nCantExiste, 6) = grdCuenta.TextMatrix(c, 12)
        End If
    Next c
    If nCantExiste > 0 Then
        MsgBox "No puede realizarse la operación porque existen " & nCantExiste & " registros que no pueden procesarse debido a que los clientes ya poseen cuentas CTS con la empresa y moneda seleccionada", vbInformation, "Aviso"
        GeneraExcelCTSExistentes (MatLista)
        Exit Sub
    End If
End If
'END JUEZ ************************************************
'EJVG20140219 ***
If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
    If Len(Trim(lblNroDoc.Caption)) = 0 Then
        MsgBox "Ud. debe seleccionar el cheque.", vbExclamation, "Aviso"
        EnfocaControl cmdDocumento
        Exit Sub
    End If
    
    Dim sValidacionChq As String
    sValidacionChq = Trim(validarApertLoteCheque)
    If Len(sValidacionChq) > 0 Then
        MsgBox "Se presentaron las siguientes observaciones: " & vbNewLine & sValidacionChq, vbExclamation, "Aviso"
        Exit Sub
    End If
End If
'END EJVG *******

' RIRO20140407 ERS017 ***
If nOperacion = gAhoApeLoteTransfBanco Or nOperacion = gCTSApeLoteTransfNew Or nOperacion = gPFApeLoteTransf Then
    Dim sValidacion As String
    sValidacion = Trim(validarApertLoteTransferencia)
    If Len(sValidacion) > 0 Then
        MsgBox "Se presentaron las siguientes observaciones: " & vbNewLine & sValidacion, vbInformation, "Aviso"
        Exit Sub
    End If
End If
' END RIRO            ***

If MsgBox("¿Desea grabar la operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub

Dim rsCuenta As New ADODB.Recordset
Dim clsMov As COMNContabilidad.NCOMContFunciones
Dim sMovNro As String, sGlosa As String, sMovNro2 As String
Dim clsApe As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim nPorcentajeCTS As Double
Dim sCodInstitucion As String
Dim psCroRetInt As String
Dim psCadImp As String
Dim psImpBoleta As String
Dim psImpBoletaITF As String
Dim psImpBoletaRes As String
Dim i As Integer, J As Integer
Dim nTipo As Integer
Dim iCue As Integer

'***Agregado por ELRO el 20121106, según OYP-RFC101-2012
Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
Dim lnSueldoMinimo As Currency
Dim lnUltimasRemuneracionesBruta As Integer
Dim lnSumaSueldosMinimoMN As Currency
Dim lnSumaSueldosMinimoME As Currency
'***Fin Agregado por ELRO el 20121106*******************

'RIRO20140430 ERS017 ******
If MsgBox("¿Desea imprimir las cartillas para las aperturas?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    bImprimirCartillas = True
Else
    bImprimirCartillas = False
End If
'END RIRO *****************

Set rsCuenta = New ADODB.Recordset
Set clsMov = New COMNContabilidad.NCOMContFunciones
sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Sleep (1000)
sMovNro2 = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set clsMov = Nothing
Set rsCuenta = grdCuenta.GetRsNew()
sGlosa = Trim(txtGlosa.Text)
nPorcentajeCTS = 0
sCodInstitucion = ""
If nProducto = gCapCTS Then
    nPorcentajeCTS = CDbl(lblDispCTS)
    sCodInstitucion = txtInstitucion
End If
'ALPA 201001
'If cboPrograma.ListIndex = 6 Then Comentado
If nTpoPrograma = 6 Then 'RIRO ERS017 Agregado
    sCodInstitucion = txtInstitucion
End If

If nOperacion = gAhoApeLoteEfec Or nOperacion = gAhoApeLoteChq Then
    nTipo = 1
ElseIf nOperacion = gPFApeLoteEfec Or nOperacion = gPFApeLoteChq Or nOperacion = gPFApeLoteTransf Then
    nTipo = 2
ElseIf nOperacion = gCTSApeLoteEfec Or nOperacion = gCTSApeLoteChq Or nOperacion = gCTSApeLoteTransfNew Then
    nTipo = 3
End If

Dim rsNroCuenta() As String
Dim nCuenta As Integer
Dim psImpPlazo As String
Dim lbok As Boolean

Set clsApe = New COMNCaptaGenerales.NCOMCaptaMovimiento
nCuenta = 1
  
Dim MatTitular() As String

Dim sOpeITFPlazoFijo As String

If chkITFEfectivo.value = 1 Then
    If OptAsuITF(0).value = True Then
        sOpeITFPlazoFijo = gITFCobroEfectivoAsumidoPF
    Else
        sOpeITFPlazoFijo = gITFCobroEfectivo
    End If
End If
If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140408
For iCue = 1 To grdCuenta.Rows - 1

    nMonto = grdCuenta.TextMatrix(iCue, 8)
'    Dim oPersona As New COMNPersona.NCOMPersona
'    If oPersona.NecesitaActualizarDatos(grdCuenta.TextMatrix(iCue, 1), gdFecSis) Then
'         MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & grdCuenta.TextMatrix(iCue, 2), vbInformation, "Aviso"
'         Dim foPersona As New frmPersona
'         If Not foPersona.realizarMantenimiento(grdCuenta.TextMatrix(iCue, 1)) Then
'             MsgBox "No se ha realizado la actualización de los datos de " & grdCuenta.TextMatrix(iCue, 2) & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
'             Exit Sub
'         End If
'    End If

'Realiza la Validación para el Lavado de Dinero
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
            If Not EsExoneradaLavadoDinero() Then
                sPersLavDinero = ""
                nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
                Set clsLav = Nothing
                If nmoneda = gMonedaNacional Then
                    Dim clsTC As COMDConstSistema.NCOMTipoCambio
                    Set clsTC = New COMDConstSistema.NCOMTipoCambio
                    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                    Set clsTC = Nothing
                Else
                    nTC = 1
                End If
                If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                    'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    'ALPA 20081009***********************************************************************************************
                    'sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), False, sTipoCuenta, , , , , nmoneda)
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), False, sTipoCuenta, , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    'ALPA
                    'If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                    'End
                End If
            End If
Next

If nProducto = gCapPlazoFijo And nTpoPrograma = -1 Then nTpoPrograma = 0 'JUEZ 20160420

If nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then
    Dim sNroDoc As String
    sNroDoc = Trim(lblNroDoc)
    
    Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    dFechaValorizacion = oNCapMov.ObtenerFechaValorizaCheque(oDocRec.fsNroDoc, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta) 'PASI20140530
    'clsApe.CapAperturaCuentaLote nProducto, nmoneda, rsCuenta, nOperacion, sGlosa, sMovNro, nTipoTasa, True, sNroDoc, sCodIF, dFechaValorizacion, nPorcentajeCTS, sCodInstitucion, gsNomCmac, gsNomAge, sLpt, gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), IIf(nProducto = gCapAhorros, gbITFAsumidoAho, IIf(nProducto = gCapPlazoFijo, gbITFAsumidoPF, False)), IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargo), psCadImp, psImpBoleta, psImpBoletaITF, rsNroCuenta, psImpPlazo, nTpoPrograma
    clsApe.CapAperturaCuentaLote nProducto, nmoneda, rsCuenta, nOperacion, sGlosa, sMovNro, nTipoTasa, True, oDocRec.fsNroDoc, oDocRec.fsPersCod, dFechaValorizacion, nPorcentajeCTS, sCodInstitucion, gsNomCmac, gsNomAge, sLpt, gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), IIf(nProducto = gCapAhorros, gbITFAsumidoAho, IIf(nProducto = gCapPlazoFijo, gbITFAsumidoPF, False)), IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargo), psCadImp, psImpBoleta, psImpBoletaITF, rsNroCuenta, psImpPlazo, nTpoPrograma, , , , , , oDocRec.fnTpoDoc, oDocRec.fsIFTpo, oDocRec.fsIFCta, lnMovNroTransfer, fnMovNroRVD, , , gbImpTMU 'EJVG20140408
    'RIRO ERS017 Se agregó "lnMovNroTransfer, fnMovNroRVD"
      
    'imprime firmas
    'FRHU 20140927 ERS099-2014
    'MsgBox "Coloque Papel para el Registro de Firmas", vbInformation, "Aviso"
    MsgBox "Coloque Papel para la Solicitud de Apertura", vbInformation, "Aviso" 'imprime solicitud de apertura
    'FIN FRHU 20140927
    'ALPA 20100202*********************************************
    'previo.Show psCadImp, "Apertura en Lote", True
    previo.Show psCadImp, "Apertura en Lote", True, , gImpresora
    Set previo = Nothing
    
      'imprime Certificado plazo fijo
    If psImpPlazo <> "" Then
        MsgBox "Coloque Papel para Certificado Plazo Fijo", vbInformation, "Aviso"
    End If
    If Trim(psImpPlazo) <> "" Then
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
           Print #nFicSal, psImpPlazo
           Print #nFicSal, ""
        Close #nFicSal
    End If
    
    
    '*** Impresion de Cartillas **** AVMM-13-08-2006
    MsgBox "Coloque Papel para Cartillas", vbInformation, "Aviso"
    

    CargaTitulares MatTitular
    MsgBox "Coloque Papel para Cartillas", vbInformation, "Aviso"
    If nOperacion = gAhoApeLoteChq Then
        If Trim(Right(cboPrograma.Text, 1)) = 0 Or Trim(Right(cboPrograma.Text, 1)) = 6 Then
            'ImpreCartillaAHLote MatTitular, rsNroCuenta
            ImpreCartillaAHLote MatTitular, rsNroCuenta, Trim(Right(cboPrograma.Text, 1)), , sMovNro 'APRI20190109 ERS077-2018
        ElseIf Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
            ImpreCartillaAHPanderoLote MatTitular, rsNroCuenta, nTpoPrograma, lblInst
        End If
    ElseIf nOperacion = gPFApeLoteChq Then
        ImpreCartillaPFLote MatTitular, rsNroCuenta
    End If
    '************************************************
    
    'imprime Boleta
    lbok = True
    MsgBox "Coloque Papel para Imprimir Boletas", vbInformation, "Aviso"
    Do While lbok
        If Trim(psImpBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
               Print #nFicSal, psImpBoleta
               Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbok = False
            End If
        End If
    Loop
    
    ' imprime BoletaITF
    If Trim(psImpBoletaITF) <> "" Then
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
           Print #nFicSal, psImpBoletaITF
           Print #nFicSal, ""
        Close #nFicSal
    End If
Else
    '***Modificado por ELRO en la fecha 20110912, según Acta 245-2011/TI-D
    'clsApe.CapAperturaCuentaLote nProducto, nmoneda, rsCuenta, nOperacion, sGlosa, sMovNro, nTipoTasa, False, , , , nPorcentajeCTS, sCodInstitucion, gsNomCmac, gsNomAge, sLpt, gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), IIf(nProducto = gCapAhorros, gbITFAsumidoAho, IIf(nProducto = gCapPlazoFijo, gbITFAsumidoPF, False)), IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargo), psCadImp, psImpBoleta, psImpBoletaITF, rsNroCuenta, psImpPlazo, nTpoPrograma    'comentado por ELRO el 20110912
    clsApe.CapAperturaCuentaLote nProducto, nmoneda, rsCuenta, nOperacion, sGlosa, sMovNro, nTipoTasa, False, , , , nPorcentajeCTS, sCodInstitucion, gsNomCmac, gsNomAge, sLpt, gbITFAplica, IIf(Me.chkExoITF.value = 1, Right(Me.cboTipoExoneracion.Text, 3), "0"), IIf(nProducto = gCapAhorros, gbITFAsumidoAho, IIf(nProducto = gCapPlazoFijo, gbITFAsumidoPF, False)), IIf(Me.chkITFEfectivo.value = 1, sOpeITFPlazoFijo, gITFCobroCargo), psCadImp, psImpBoleta, psImpBoletaITF, rsNroCuenta, psImpPlazo, nTpoPrograma, Me.chkITFEfectivo.value, psCroRetInt, lsPersCodConv, psImpBoletaRes, nMontoTotal, , , , lnMovNroTransfer, fnMovNroRVD, IIf(Not IsNumeric(lsDetalle), 0, lsDetalle), sMovNro2, gbImpTMU  ' RIRO ERS017 lnMovNroTransfer, fnMovNroRVD
    '***Fin Modificado por ELRO
       
     'imprime firmas
    'FRHU 20140927 ERS099-2014
    'MsgBox "Coloque Papel para el Registro de Firmas", vbInformation, "Aviso"
    MsgBox "Coloque Papel para la Solicitud de Apertura", vbInformation, "Aviso" 'imprime solicitud de apertura
    'FIN FRHU 20140927
    'ALPA 20100202*****************************************
    'previo.Show psCadImp, "Apertura en Lote", True
    previo.Show psCadImp, "Apertura en Lote", True, , gImpresora
    Set previo = Nothing
    
    'Modificado por ELRO el 20110915, según Acta 252-2011/TI-D
    'imprime cronograma de retiro de interes
    If psCroRetInt <> "" Then
    MsgBox "Coloque Papel para el Cronograma de Retiro de Intereses", vbInformation, "Aviso"
   
    previo2.Show psCroRetInt, "Apertura en Lote", True, , gImpresora
    Set previo2 = Nothing
    End If
    '***Fin Modificado por ELRO
    
    'imprime Certificado plazo fijo
    If psImpPlazo <> "" Then
        MsgBox "Coloque Papel para Certificado Plazo Fijo", vbInformation, "Aviso"
    End If
    If Trim(psImpPlazo) <> "" Then
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
           Print #nFicSal, psImpPlazo
           Print #nFicSal, ""
        Close #nFicSal
    End If
            
    If bImprimirCartillas Then 'RIRO 20140531 - ERS017, Se dio la opcion de elegir imprimir o no las cartillas.
    
    '*** Impresion de Cartillas **** AVMM-13-08-2006
        CargaTitulares MatTitular
        MsgBox "Coloque Papel para Cartillas", vbInformation, "Aviso"
         If nOperacion = gAhoApeLoteEfec Or nOperacion = gAhoApeLoteTransfBanco _
            Or nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Then 'RIRO20140430 ERS017 Se agrego "gAhoApeLoteTransfBanco"
            If Trim(Right(cboPrograma.Text, 1)) = 0 Or Trim(Right(cboPrograma.Text, 1)) = 5 Or Trim(Right(cboPrograma.Text, 1)) = 6 Or Trim(Right(cboPrograma.Text, 1)) = 8 Then
            '***Condición Trim(Right(cboPrograma.Text, 1)) = 8 agregado por ELRO el 20130131, según TI-ERS020-2013
            'ALPA 20100108***************************
                 'ImpreCartillaAHLote MatTitular, rsNroCuenta
                '***Modificado por ELRO el 20130131, según TI-ERS020-2013
                If Trim(Right(cboPrograma.Text, 1)) = 8 Then
                    If Trim(cboInstConvDep) <> "" Then
                        ImpreCartillaAHLote MatTitular, rsNroCuenta, Trim(Right(cboPrograma.Text, 1)), Left(cboInstConvDep, Len(cboInstConvDep) - 13), sMovNro
                    Else
                        ImpreCartillaAHLote MatTitular, rsNroCuenta, Trim(Right(cboPrograma.Text, 1)), , sMovNro
                    End If
                Else
                    ImpreCartillaAHLote MatTitular, rsNroCuenta, Trim(Right(cboPrograma.Text, 1)), , sMovNro
                End If
                '***Fin Modificado por ELRO el 20130131******************
            '****************************************
            ElseIf Trim(Right(cboPrograma.Text, 1)) = 3 Or Trim(Right(cboPrograma.Text, 1)) = 4 Then
                 ImpreCartillaAHPanderoLote MatTitular, rsNroCuenta, nTpoPrograma, lblInst.Caption
            End If
        ElseIf nOperacion = gPFApeLoteEfec Or nOperacion = gPFApeLoteTransf Then
            ImpreCartillaPFLote MatTitular, rsNroCuenta
        End If
        '***Agregado por ELRO el 20121201, según OYP-RFC101-2012
        If nProducto = gCapCTS Then
            ImpreCartillaCTSLote rsCuenta, rsNroCuenta, sMovNro
        End If
        '***Fin Agregado por ELRO el 20121201*******************
    'INICIO EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
     AhorroApertura_ContratosAutomaticosLote MatTitular, rsNroCuenta, sMovNro
    'FIN EAAS20190523 Memorándum Nº 756-2019-GM-DI/CMACM
    End If
    '************************************************
    
    'imprime Boleta
    MsgBox "Coloque Papel para Imprimir Boletas", vbInformation, "Aviso"
        
    lbok = True
    Do While lbok
        If Trim(psImpBoletaRes) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
               Print #nFicSal, psImpBoletaRes
               Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta Resumen ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbok = False
            End If
        '***Agregado por ELRO el 20121129, según OYP-RFC101-2012
        Else
            lbok = False
        '***Agregado por ELRO el 20121129***********************
        End If
    Loop
    
    If MsgBox("Desea Imprimir detalle de las operaciones ??", vbYesNo + vbQuestion, "AVISO") = vbYes Then
        lbok = True
        Do While lbok
            If Trim(psImpBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                   Print #nFicSal, psImpBoleta
                   Print #nFicSal, ""
                Close #nFicSal
                If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbok = False
                End If
            End If
        Loop
        
        ' imprime BoletaITF
    
        If Trim(psImpBoletaITF) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
               Print #nFicSal, psImpBoletaITF
               Print #nFicSal, ""
            Close #nFicSal
        End If
    End If
    '***Agregado por ELRO el 20121205, según OYP-RFC101-2012
    If nProducto = gCapCTS Then
        Dim nK As Integer
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        lnSueldoMinimo = CCur(clsLav.GetCapParametro(2128))
        lnUltimasRemuneracionesBruta = CCur(clsLav.GetCapParametro(2129))
        lnSumaSueldosMinimoMN = lnSueldoMinimo * lnUltimasRemuneracionesBruta
        lnSumaSueldosMinimoME = lnSumaSueldosMinimoMN / oDCOMGeneral.GetTipCambio(gdFecSis, TCFijoMes)
        Set clsLav = Nothing
    
        rsCuenta.MoveFirst
        nK = 1
        Do While Not rsCuenta.EOF
            If rsCuenta.Fields(17) <> "" And rsCuenta.Fields(18) <> "" Then
                If (CInt(rsCuenta.Fields(17)) = 1 And CCur(rsCuenta.Fields(18)) > lnSumaSueldosMinimoMN) Or _
                   (CInt(rsCuenta.Fields(17)) = 2 And CCur(rsCuenta.Fields(18)) > lnSumaSueldosMinimoME) Then
                    actualizarSeisUltimasRemuneraciones rsNroCuenta(nK), CInt(rsCuenta.Fields(17)), CCur(rsCuenta.Fields(18))
                End If
            End If
        nK = nK + 1
        rsCuenta.MoveNext
        Loop
    End If
    '***Fin Agregado por ELRO el 20121205*******************
End If
'CTI7 OPEv2**********************
cboTransferMoneda.ListIndex = IndiceListaCombo(cboTransferMoneda, Trim(Right(cboMoneda.Text, 5)))
lblTrasferND.Caption = ""
lbltransferBco.Caption = ""
txtTransferGlosa.Text = ""
lblMonTra.Caption = "0.00"
'********************************
Set clsApe = Nothing
ClearScreen
'INICIO JHCU ENCUESTA 16-10-2019
Dim cOpecodEncuesta As String
cOpecodEncuesta = nOperacion
Encuestas gsCodUser, gsCodAge, "ERS0292019", cOpecodEncuesta
'FIN
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

'RIRO20140407 ERS017 **********************************************************
Private Sub cmdTranfer_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oForm As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lnTransferSaldo As Currency
    Dim fsPersCodTransfer As String

    If cboTransferMoneda.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMoneda.Visible And cboTransferMoneda.Enabled Then cboTransferMoneda.SetFocus
        Exit Sub
    End If
    If gsOpeCod = gAhoApeLoteTransfBanco Then
        lnTipMot = 8
    ElseIf gsOpeCod = gCTSApeLoteTransfNew Then
        lnTipMot = 23
    ElseIf gsOpeCod = gPFApeLoteTransf Then
        lnTipMot = 24
    End If
    fnMovNroRVD = 0
    Set oForm = New frmCapRegVouDepBus
    SetDatosTransferencia "", "", "", 0, -1, "" 'Limpiamos datos y variables globales
    oForm.iniciarFormulario Trim(Right(cboTransferMoneda, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    If IsNumeric(Trim(lsDetalle)) Then
        nNroApertura = CInt(lsDetalle)
    Else
        nNroApertura = 0
    End If
    SetDatosTransferencia lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Me.grdCuenta.row = 1
    Set oForm = Nothing
    Exit Sub
End Sub

Private Sub SetDatosTransferencia(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosa.Text = psGlosa
    lbltransferBco.Caption = psInstit
    lblTrasferND.Caption = psDoc
    If psDetalle <> "" Then

    End If
    
    txtMonto_Change

    If pnMovNroTransfer <> -1 Then
        txtTransferGlosa.SetFocus
    End If
    
    txtTransferGlosa.Locked = True
    txtMonto.Enabled = False
    lblMonTra = Format(pnTransferSaldo, "#,##0.00")
    
    Set rsPersona = Nothing
    Set oPersona = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
'END RIRO *****************************************

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)

Dim oCons As COMDConstantes.DCOMConstantes 'DConstante
Set oCons = New COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Set rs = oCons.GetConstante(1044, , , True, , "0','1044")
Me.cboTipoExoneracion.Clear
While Not rs.EOF
    cboTipoExoneracion.AddItem rs.Fields(1) & space(100) & rs.Fields(0)
    rs.MoveNext
Wend

If nOperacion = gPFApeLoteEfec Or nOperacion = gPFApeLoteChq Or nOperacion = gPFApeLoteTransf Then 'CTI7 OPEv2
    grdCuenta.ColWidth(13) = 0
    grdCuenta.ColWidth(14) = 0
ElseIf nOperacion = gCTSApeLoteEfec Or nOperacion = gCTSApeLoteChq Or nOperacion = gCTSApeLoteTransfNew Then
    grdCuenta.ColWidth(13) = 0
    grdCuenta.ColWidth(14) = 0
Else
'    grdCuenta.ColWidth(13) = 1200
'    grdCuenta.ColWidth(14) = 1200
End If

'RIRO201404 ERS017 *********************
GetTipCambio gdFecSis
lblTTCCD.Caption = Format(gnTipCambioC, "#,#0.0000")
lblTTCVD.Caption = Format(gnTipCambioV, "#,#0.0000")
bCargaLote = False
fraTranferecia.Visible = False
'END RIRO ******************************

Set oCons = Nothing
Set rs = Nothing

End Sub

Private Sub grdClientes_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 4 Or pnCol = 5 Or pnCol = 6 Then
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
    Dim bOrdPag As Boolean
    Dim nMonto As Double, nTasa As Double
    Dim nPlazo As Long
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    bOrdPag = IIf(grdClientes.TextMatrix(pnRow, 4) = ".", True, False)
    nMonto = CDbl(grdClientes.TextMatrix(pnRow, 6))
    
    'ALPA 20100125**************************************
    If Trim(Right(cboPrograma.Text, 4)) = "6" Then
        nTasa = nTasaNominal
    Else
        nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , nMonto, gsCodAge, bOrdPag, IIf(nProducto = gCapAhorros, Right(cboPrograma.Text, 4), 0))
    End If
    '***************************************************
    
    grdClientes.TextMatrix(pnRow, 6) = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
    '***Modificado por ELRO el 20110915, según el Acta 252-2011/TI-D
    'grdCuenta.TextMatrix(pnRow, 11) = Format$(nTasa, "#,##0.00")   'comentado por ELRO el 20110915
    'grdCuenta.TextMatrix(pnRow, 11) = nTasa
    '***Fin Modificado por ELRO
    Set clsDef = Nothing
    
    If nProducto = gCapAhorros Or nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
        If Me.chkExoITF.value <> 1 Then
            grdClientes.TextMatrix(pnRow, 9) = Format(fgITFCalculaImpuesto(CCur(grdClientes.TextMatrix(pnRow, 7))), "#,##0.00")
        End If
    End If
    
    If nOperacion = gAhoApeLoteEfec Or nOperacion = gPFApeLoteEfec Or nOperacion = gCTSApeLoteEfec Or nOperacion = gCTSApeLoteTransfNew Or nOperacion = gPFApeLoteTransf Then
        txtMonto.Text = Format$(grdClientes.SumaRow(7), "#,##0.00")
    End If
    
    If pnCol = 7 Then
        '***Modificado por ELRO el 20110913, segun Acta 252-2011/TI-D
        If grdClientes.Rows > 1 And grdClientes.TextMatrix(1, 0) <> "" Then
    
        Dim nRow As Integer
        Dim bIncluidoLavDin As Boolean
        nRow = grdClientes.Rows - 1
         
        If grdClientes.TextMatrix(nRow, 1) <> "" Then
            Call llenarFormularioLavDin
        End If
        
        End If
        '***Fin Modificado por ELRO *********************************
    End If
    
End If
End Sub

Private Sub grdClientes_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim bOrdPag As Boolean
Dim nMonto As Double, nTasa As Double
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    
    bOrdPag = IIf(grdClientes.TextMatrix(pnRow, 4) = ".", True, False)
    nMonto = CDbl(grdClientes.TextMatrix(pnRow, 7))
    nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , nMonto, gsCodAge, bOrdPag)
    grdClientes.TextMatrix(pnRow, 5) = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
    grdCuenta.TextMatrix(pnRow, 8) = Format$(nTasa, "#,##0.00")
        
    Set clsDef = Nothing
End Sub

Private Sub grdCuenta_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 5 Or pnCol = 6 Or pnCol = 8 Then
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion 'nCapDefinicion
    Dim bOrdPag As Boolean
    Dim nMonto As Double, nTasa As Double
    Dim nPlazo As Long
    '***Agregado por ELRO el 20121129, según OYP-RFC101-2012
    Dim lsTitular As String
    Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
    '***Fin Agregado por ELRO el 20121129*******************
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    bOrdPag = IIf(grdCuenta.TextMatrix(pnRow, 5) = ".", True, False)
    nMonto = CDbl(grdCuenta.TextMatrix(pnRow, 8))
    
    If nProducto = gCapPlazoFijo Then
        If grdCuenta.TextMatrix(pnRow, 6) <> "" Then
            nPlazo = CLng(grdCuenta.TextMatrix(pnRow, 6))
            
            nTasa = clsDef.GetCapTasaInteresCamp(nProducto, lnTpoPrograma, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, False, fnCampanaCod) 'JUEZ 20160420
   '        nTasa = clsDef.GetCapTasaInteresCamp(nProducto, 0, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, False, fnCampanaCod) 'JUEZ 20160420
            If nTasa = 0 Then nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , lnTpoPrograma)
        End If
    '***Agregado por ELRO el 20121129, según OYP-RFC101-2012
    ElseIf nProducto = gCapCTS Then
        nTpoProgramaCTS = 1 'Por defecto se asigna Tasa de CTS sin Cta Sueldo
        lsTitular = grdCuenta.TextMatrix(pnRow, 1)
        If lsTitular <> "" Then
            Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
            If oNCOMCaptaGenerales.TieneCuentasCaptacxSubProducto(lsTitular, gCapAhorros, 6) Then 'Verifica si Cliente tiene Cta Sueldo
                nTpoProgramaCTS = 0 ' Si tiene CtaSueldo se cambia de Sub Producto
            End If
            Set oNCOMCaptaGenerales = Nothing
        End If
        nTasa = clsDef.GetCapTasaInteresCamp(nProducto, nTpoProgramaCTS, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, False, fnCampanaCod) 'JUEZ 20160420
        If nTasa = 0 Then nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoProgramaCTS) 'JUEZ 20160420
        nTasaNominal = nTasa
        grdCuenta.TextMatrix(pnRow, 17) = nTpoProgramaCTS
    '***Agregado por ELRO el 20121129***********************
    Else
        'ALPA 20100125**************************************
        If Trim(Right(cboPrograma.Text, 4)) = "6" Then
            nTasa = nTasaNominal
        Else
            nTasa = clsDef.GetCapTasaInteresCamp(nProducto, IIf(nProducto = gCapAhorros, Right(cboPrograma.Text, 4), 0), nmoneda, 0, nMonto, gsCodAge, gdFecSis, False, bOrdPag, fnCampanaCod) 'JUEZ 20160420
            If nTasa = 0 Then nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , nMonto, gsCodAge, bOrdPag, IIf(nProducto = gCapAhorros, Right(cboPrograma.Text, 4), 0)) 'JUEZ 20160420
        End If
        '***************************************************
    End If
    grdCuenta.TextMatrix(pnRow, 7) = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
    '***Modificado por ELRO el 20110915, según el Acta 252-2011/TI-D
    'grdCuenta.TextMatrix(pnRow, 11) = Format$(nTasa, "#,##0.00")   'comentado por ELRO el 20110915
    grdCuenta.TextMatrix(pnRow, 11) = nTasa
    '***Fin Modificado por ELRO
    grdCuenta.TextMatrix(pnRow, 21) = fnCampanaCod 'JUEZ 20160420
    Set clsDef = Nothing
    
    If nProducto = gCapAhorros Or nProducto = gCapPlazoFijo Or nProducto = gCapCTS Then
        If Me.chkExoITF.value <> 1 And grdCuenta.TextMatrix(pnRow, 8) >= 1000 Then
            grdCuenta.TextMatrix(pnRow, 12) = Format(fgITFCalculaImpuesto(CCur(grdCuenta.TextMatrix(pnRow, 8))), "#,##0.00")
            'Redondeo del ITF
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(grdCuenta.TextMatrix(pnRow, 12)))
            If nRedondeoITF > 0 Then
                grdCuenta.TextMatrix(pnRow, 12) = Format(CCur(grdCuenta.TextMatrix(pnRow, 12)) - nRedondeoITF, "#,##0.00")
            End If
            'Fin del Redondeo
        Else
            grdCuenta.TextMatrix(pnRow, 12) = Format(0, "#,##0.00")
        End If
    End If
    
    'RIRO20140430 ERS017, Se agrego "gAhoApeLoteTransfBanco"
    If nOperacion = gAhoApeLoteEfec Or nOperacion = gPFApeLoteEfec Or nOperacion = gCTSApeLoteEfec Or nOperacion = gAhoApeLoteTransfBanco _
        Or nOperacion = gAhoApeLoteChq Or nOperacion = gPFApeLoteChq Or nOperacion = gCTSApeLoteChq Or nOperacion = gCTSApeLoteTransfNew Or nOperacion = gPFApeLoteTransf Then
        txtMonto.Text = Format$(grdCuenta.SumaRow(8), "#,##0.00")
    End If
    
    If pnCol = 8 Then
    '***Modificado por ELRO el 20110913, segun Acta 252-2011/TI-D
If grdCuenta.Rows > 1 And grdCuenta.TextMatrix(1, 0) <> "" Then


    Dim nRow As Integer
    Dim bIncluidoLavDin As Boolean
    nRow = grdCuenta.Rows - 1
     
'   If grdCuenta.TextMatrix(nRow, 1) <> "" Then
'        Call llenarFormularioLavDin
'    End If
    
End If
'***Fin Modificado por ELRO *********************************
    
    End If
    
End If
End Sub

Private Sub grdCuenta_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim bOrdPag As Boolean
    Dim nMonto As Double, nTasa As Double
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        bOrdPag = IIf(grdCuenta.TextMatrix(pnRow, 5) = ".", True, False)
        nMonto = CDbl(grdCuenta.TextMatrix(pnRow, 8))
        nTasa = clsDef.GetCapTasaInteresCamp(nProducto, lnTpoPrograma, nmoneda, 0, nMonto, gsCodAge, gdFecSis, False, bOrdPag, fnCampanaCod) 'JUEZ 20160420
        If nTasa = 0 Then nTasa = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , nMonto, gsCodAge, bOrdPag, lnTpoPrograma) 'JUEZ 20160420
        grdCuenta.TextMatrix(pnRow, 7) = Format$(ConvierteTNAaTEA(nTasa), "#,##0.00")
        grdCuenta.TextMatrix(pnRow, 11) = Format$(nTasa, "#,##0.00")
        grdCuenta.TextMatrix(pnRow, 21) = fnCampanaCod 'JUEZ 20160420
    Set clsDef = Nothing

End Sub

Private Sub grdCuenta_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
If psDataCod <> "" Then
    grdCuenta.TextMatrix(pnRow, 9) = grdCuenta.PersPersoneria
Else
    grdCuenta.EliminaFila pnRow
    Exit Sub 'RIRO ERS017
End If
If pbEsDuplicado Then
    grdCuenta.EliminaFila pnRow
    Exit Sub 'RIRO ERS017
End If

'RIRO20140407 ERS017
Dim oPersona As COMNPersona.NCOMPersona
Dim sMensaje As String
'END RIRO
    
'-- Para agregar  La Direccion y el documento que se Utilizara para la Cartilla -- AVMM -- 06-02-2006
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim R As New ADODB.Recordset
Set ClsPersona = New COMDPersona.DCOMPersonas
Set oPersona = New COMNPersona.NCOMPersona 'RIRO20140430 ERS017
Set R = ClsPersona.BuscaCliente(grdCuenta.TextMatrix(grdCuenta.row, 1), BusquedaCodigo)
    If Not (R.EOF And R.BOF) Then
'       grdCuenta.TextMatrix(grdCuenta.Row, 13) = IIf(IsNull(R!cPersIDnroDNI), "", IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI))
'       grdCuenta.TextMatrix(grdCuenta.Row, 14) = R!cPersDireccDomicilio
       grdCuenta.TextMatrix(grdCuenta.row, 15) = IIf(IsNull(R!cPersIDnroDNI), "", IIf(R!cPersIDnroDNI = "", R!cPersIDnroRUC, R!cPersIDnroDNI))
       grdCuenta.TextMatrix(grdCuenta.row, 16) = R!cPersDireccDomicilio
       'RIRO20140407 ERS017 *******************************
       If (Not oPersona.validarPersonaMayorEdad(R!cperscod, CDate(gdFecSis)) And R!nPersPersoneria = 1) Or R!nPersPersoneria > 1 Then
        If R!nPersPersoneria > 1 Then
            sMensaje = "El cliente seleccionado tiene Personería Jurídica."
        Else
            sMensaje = "El cliente seleccionado es menor de edad."
        End If
        MsgBox sMensaje, vbInformation, "Aviso"
        grdCuenta.EliminaFila pnRow
       End If
       'END RIRO ******************************************
    End If
Set ClsPersona = Nothing
Set oPersona = Nothing 'RIRO20140407 ERS017
Set R = Nothing 'RIRO20140407 ERS017
'----------------------------------------------------------------------------------------------------
End Sub

'RIRO20140430 ERS017 ***************************
Private Sub grdCuenta_OnRowDelete()
    If nOperacion = gAhoApeLoteTransfBanco Then
        txtMonto.Text = Format$(grdCuenta.SumaRow(8), "#,##0.00")
    End If
End Sub
Private Sub grdCuenta_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdCuenta.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub grdCuenta_RowColChange()
    Dim nRow As Long
    Dim nCol As Long
    
    nRow = grdCuenta.row
    nCol = grdCuenta.Col
    If bCargaLote Then
        If nCol = 1 Then
            grdCuenta.lbEditarFlex = False
            Me.KeyPreview = True
        Else
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = True
        End If
    Else
        If nCol = 1 Then
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = False
        Else
            grdCuenta.lbEditarFlex = True
            Me.KeyPreview = True
        End If
    End If
    If Not IsNumeric(Trim(grdCuenta.TextMatrix(nRow, 1))) Then
        grdCuenta.TextMatrix(nRow, 1) = ""
    End If
End Sub
Private Sub LimpiarGrdCuenta()
    grdCuenta.Clear
    grdCuenta.Rows = 2
    grdCuenta.FormaCabecera
    grdCuenta.ColWidth(4) = 0
    grdCuenta.ColWidth(6) = 0
    grdCuenta.ColWidth(17) = 0
    grdCuenta.ColWidth(18) = 0
    grdCuenta.ColWidth(19) = 0
End Sub
'RIRO *****************************************
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    'RIRO ERS017 ***
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled And cmdGrabar.Visible Then cmdGrabar.SetFocus
    End If
    'END RIRO ******
End Sub

Private Sub txtInstitucion_EmiteDatos()
If txtInstitucion.Text <> "" Then
    If chkRelConv.value = 0 Then
        lblInst.Caption = txtInstitucion.psDescripcion
        cmdGrabar.SetFocus
    Else
        If Not ValidaInstConv(txtInstitucion.Text) Then
            MsgBox "La Institucion no esta para convenio de Depositos", vbInformation, "SISTEMA"
            txtInstitucion.Text = ""
        Else
            lblInst.Caption = txtInstitucion.psDescripcion
            cmdGrabar.SetFocus
        End If
    End If
    '*** RIRO20140430 ERS017
    If nProducto = gCapAhorros Then
        ValidarLoteTitular
    End If
    '*** END RIRO
End If
End Sub

Private Sub txtMonto_Change()
    ValidaTasaInteres
    If gbITFAplica And nProducto <> gCapCTS Then       'Filtra para CTS
        If txtMonto.value > gnITFMontoMin Then
            If Me.chkExoITF.value = 0 Then
                If nOperacion = gAhoApeTransf Or nOperacion = gPFApeTransf Then
                    Me.lblITF.Caption = Format(0, "#,##0.00")
                ElseIf nProducto = gCapAhorros And (nOperacion <> gAhoApeChq) Then
                    'Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    Me.lblITF.Caption = Format(grdCuenta.SumaRow(12), "#,##0.00") 'RIRO20140430 ERS017
                Else
                    'Me.lblITF.Caption = Format(fgITFCalculaImpuesto(txtMonto.value), "#,##0.00")
                    Me.lblITF.Caption = Format(grdCuenta.SumaRow(12), "#,##0.00") 'RIRO20140430 ERS017
                End If
                nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
                If nRedondeoITF > 0 Then
                   Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
                End If
            Else
                Me.lblITF.Caption = "0.00"
            End If
                                                                                                                        
                                                           
            If nOperacion = gPFApeLoteChq Or nOperacion = gAhoApeLoteChq Then
                    If nProducto = gCapAhorros And gbITFAsumidoAho Then
                        Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                    ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                        Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                    Else
                        If Me.chkITFEfectivo.value = 1 Then
                            Me.lblTotal.Caption = Format(txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00")
                        Else
                            Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                        End If
                    End If
            Else
                If nProducto = gCapAhorros And gbITFAsumidoAho Then
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                ElseIf nProducto = gCapPlazoFijo And gbITFAsumidoPF Then
                    Me.lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
                Else
                    '***Modificado por ELRO en la fecha 20110912, según Acta 245-2011/TI-D
                    'Me.lblTotal.Caption = Format(txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00") 'comentado por ELRO el 20110912
                    Me.lblTotal.Caption = IIf(Me.chkITFEfectivo.value = 1, Format(txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00"), Format(txtMonto.value, "#,##0.00")) 'RIRO20140407 ERS017 Comentado
                    'Me.lblTotal.Caption = IIf(Me.chkITFEfectivo.value = 1, Format(txtMonto.value, "#,##0.00"), Format(txtMonto.value + CCur(Me.lblITF.Caption), "#,##0.00")) 'RIRO20140407 ERS017 Agregado
                    '***Fin Modificado por ELRO
                End If
            End If
        End If
    Else
        Me.lblITF.Caption = Format(0, "#,##0.00")
        If nOperacion = gCTSDepChq Then
            Me.lblTotal.Caption = Format(0, "#,##0.00")
        Else
            Me.lblTotal.Caption = Format(Me.txtMonto.value, "#,##0.00")
        End If
        chkITFEfectivo_Click
    End If
    If txtMonto.value = 0 Then
        Me.lblITF.Caption = "0.00"
        Me.lblTotal.Caption = "0.00"
        chkITFEfectivo_Click
    End If
End Sub

Public Sub CargaTitulares(ByRef MatTitular() As String)
    Dim nNumTit As Integer
    Dim i As Integer
    nNumTit = grdCuenta.Rows - 1
    
    'ALPA 20100109***************************
    'ReDim MatTitular(nNumTit, 8)
    'ReDim MatTitular(nNumTit, 9)
    ReDim MatTitular(nNumTit, 10) 'JUEZ 20150121
    '****************************************
    For i = 1 To grdCuenta.Rows - 1
        MatTitular(i, 1) = grdCuenta.TextMatrix(i, 2)
        MatTitular(i, 2) = grdCuenta.TextMatrix(i, 7)
        MatTitular(i, 3) = grdCuenta.TextMatrix(i, 8)
        MatTitular(i, 4) = grdCuenta.TextMatrix(i, 15)
        MatTitular(i, 5) = grdCuenta.TextMatrix(i, 16)
        MatTitular(i, 6) = IIf(grdCuenta.TextMatrix(i, 6) = "", 0, grdCuenta.TextMatrix(i, 6))
        MatTitular(i, 7) = IIf(grdCuenta.TextMatrix(i, 13) = "", 0, grdCuenta.TextMatrix(i, 13))
        MatTitular(i, 8) = IIf(grdCuenta.TextMatrix(i, 14) = "", 0, grdCuenta.TextMatrix(i, 14))
        'ALPA 20100109***************************
        MatTitular(i, 9) = grdCuenta.TextMatrix(i, 1)
        '****************************************
        MatTitular(i, 10) = Trim(Right(grdCuenta.TextMatrix(i, 4), 2)) 'JUEZ 20150121
    Next i
End Sub
'***Modificado por ELRO el 20110913, según Acta 252-2011/TI-D
Private Function EsExoneradaLavadoDinero() As Boolean
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim bExito As Boolean
Dim clsExo As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
bExito = True
Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
For i = 1 To grdCuenta.Rows - 1
    nRelacion = 10
    If nRelacion = gCapRelPersTitular Then
        sPersCod = grdCuenta.TextMatrix(i, 1)
        If Not clsExo.EsPersonaExoneradaLavadoDinero(sPersCod) Then
            bExito = False
            Exit For
        End If
    End If
Next i
Set clsExo = Nothing
EsExoneradaLavadoDinero = bExito
End Function

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)
Dim i As Long
Dim nRelacion As COMDConstantes.CaptacRelacPersona
Dim oPersona As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsPers As New ADODB.Recordset

For i = grdCuenta.Rows - 1 To grdCuenta.Rows - 1
    nRelacion = 10
    If grdCuenta.TextMatrix(i, 9) = gPersonaNat Then
        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
            poLavDinero.TitPersLavDinero = grdCuenta.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCuenta.TextMatrix(i, 2)
            Exit For
        End If
    Else
        If nRelacion = gCapRelPersTitular And grdCuenta.TextMatrix(i, 3) = "TI" Then
            poLavDinero.TitPersLavDinero = grdCuenta.TextMatrix(i, 1)
            poLavDinero.TitPersLavDineroNom = grdCuenta.TextMatrix(i, 2)
        End If
        If nRelacion = gCapRelPersRepTitular And grdCuenta.TextMatrix(i, 3) = "TI" Then
            poLavDinero.TitPersLavDinero = grdCuenta.TextMatrix(i, 1)
            poLavDinero.ReaPersLavDineroNom = grdCuenta.TextMatrix(i, 2)
            If poLavDinero.TitPersLavDinero <> "" Then Exit For
        End If
    End If
Next i
sTipoCuenta = Me.Caption
End Sub

Private Sub llenarFormularioLavDin()
    Dim nRow As Integer
    Dim sPersLavDinero As String, sCuenta As String
    Dim nMonto As Double, nMontoLavDinero As Double, nTC As Double

    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim loLavDinero As frmMovLavDinero
    Set loLavDinero = New frmMovLavDinero


    nRow = grdCuenta.Rows - 1
    nMonto = grdCuenta.TextMatrix(nRow, 8)
    
     If Not EsExoneradaLavadoDinero() Then
        sPersLavDinero = ""
        nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
        Set clsLav = Nothing
        If nmoneda = gMonedaNacional Then
            Dim clsTC As COMDConstSistema.NCOMTipoCambio
            Set clsTC = New COMDConstSistema.NCOMTipoCambio
            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set clsTC = Nothing
        Else
            nTC = 1
        End If
        If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                  
            Call IniciaLavDinero(loLavDinero)
        
            sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), False, sTipoCuenta, , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen, False)
                        
             If sPersLavDinero = "Aviso" Then
                MsgBox "El monto que desea aperturar no se realiza por esta operación", vbInformation, "Aviso"
                grdCuenta.TextMatrix(nRow, 8) = "0.00"
                grdCuenta.TextMatrix(nRow, 12) = "0.00"
                grdCuenta.TextMatrix(nRow, 13) = "0.00"
                grdCuenta.TextMatrix(nRow, 14) = "0.00"
                grdCuenta.TextMatrix(nRow, 11) = nTasaNominal
                grdCuenta.TextMatrix(nRow, 21) = fnCampanaCod 'JUEZ 20160420
                Exit Sub
             End If
                        
            If loLavDinero.OrdPersLavDinero = "" Then
                Exit Sub
            End If
        End If
    End If
    
End Sub

'***Fin Modificado por ELRO

'**Create By GITU 11-09-2012
Private Sub IniciaComboConvDep(ByVal pnTipoRol As Integer)
Dim lRegPers As New ADODB.Recordset
Dim oPers As COMDPersona.DCOMRoles

    Set oPers = New COMDPersona.DCOMRoles
    Set lRegPers = oPers.CargaPersonas(pnTipoRol)
    Set oPers = Nothing

    If Not lRegPers.BOF And Not lRegPers.EOF Then
        Do While Not lRegPers.EOF
            If lRegPers("PersEstado") = "ACTIVO" Then
                cboInstConvDep.AddItem lRegPers("cPersNombre") & space(100) & lRegPers("cPersCod")
            End If
            lRegPers.MoveNext
        Loop
        cboInstConvDep.ListIndex = 0
    End If
    lRegPers.Close
    Set lRegPers = Nothing
 End Sub
 
Private Function ValidaInstConv(ByVal psCodPers As String) As Boolean
    Dim oPers As COMDPersona.DCOMRoles

    Set oPers = New COMDPersona.DCOMRoles
    If oPers.ExistePersonaRol(psCodPers, 9) Then
        ValidaInstConv = True
    Else
        ValidaInstConv = False
    End If
    Set oPers = Nothing
End Function
Private Function ValidaTasaInteres()
Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'BRGO 20111020
Dim bOrdPag As Boolean
Dim nMonto As Double
Dim nPlazo As Long
Dim nTpoPrograma As Integer
Dim sTitular As String

Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales 'BRGO 20111020
'bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
nMonto = txtMonto.value
nTpoPrograma = 0

If cboPrograma.Visible Then
    nTpoPrograma = CInt(Right(Trim(cboPrograma.Text), 2))
End If
'If chkTasaPreferencial.value = vbUnchecked Then
    nTipoTasa = gCapTasaNormal
    If nProducto = gCapPlazoFijo Then
'        If txtPlazo <> "" Then
'            nPlazo = CLng(txtPlazo)
'            'Add by Gitu 2010-08-06
'            If lnValOpePF = 1 And (nPersoneria <> gPersonaNat Or lnTitularPJ = 1) Then
'                        If chkDepGar.value = 1 Then 'MADM 20111022
'                             'nTasaNominal = (clsDef.GetCapTasaInteresPF(gCapPlazoFijo, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma) / 2)
'                            nTasaNominalTemp = clsDef.GetCapTasaInteresPF(gCapPlazoFijo, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
'                            nTasaEfectivaTemp = Format$(ConvierteTNAaTEA(nTasaNominalTemp), "#,##0.00") / 2
'                            nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.00")
'                        Else
'                             nTasaNominal = clsDef.GetCapTasaInteresPF(gCapPlazoFijo, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
'                        End If
'            Else
'                        If chkDepGar.value = 1 Then 'MADM 20111022
'                            'nTasaNominal = (clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma) / 2)
'                            nTasaNominalTemp = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
'                            nTasaEfectivaTemp = Format$(ConvierteTNAaTEA(nTasaNominalTemp), "#,##0.00") / 2
'                            nTasaNominal = Format$(ConvierteTEAaTNA(nTasaEfectivaTemp), "#,##0.00")
'                        Else
'                            nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoPrograma)
'                        End If
'            End If
'            'End Gitu
'            lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
'        End If
    ElseIf nProducto = gCapAhorros Then
        nTasaNominal = clsDef.GetCapTasaInteresCamp(nProducto, nTpoPrograma, nmoneda, nPlazo, nMonto, gsCodAge, gdFecSis, False, bOrdPag, fnCampanaCod) 'JUEZ 20160420
        If nTasaNominal = 0 Then nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, bOrdPag, nTpoPrograma) 'JUEZ 20160420
        lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    Else
        'nTpoProgramaCTS = 1 'Por defecto se asigna Tasa de CTS sin Cta Sueldo
        'sTitular = ObtTitular
'        If sTitular <> "" Then
'            If clsCap.TieneCuentasCaptacxSubProducto(sTitular, gCapAhorros, 6) Then 'Verifica si Cliente tiene Cta Sueldo
'                nTpoProgramaCTS = 0 ' Si tiene CtaSueldo se cambia de Sub Producto
'            End If
'        End If
        'nTasaNominal = clsDef.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nMonto, gsCodAge, , nTpoProgramaCTS)
        'lblTasa.Caption = Format$(ConvierteTNAaTEA(nTasaNominal), "#,##0.00")
    End If
'End If
Set clsDef = Nothing
End Function

'RIRO20140430 ERS017 Se agrego el parámetro pbMayorEdad
Private Function DevuelveDatosCliente(ByVal psNroDoc As String, ByRef psPersCod As String, ByRef psNombre As String, ByRef pnPersoneria As Integer, _
                                      ByRef psDirec As String, ByRef pbMayorEdad As Boolean)
Dim oPer As COMDPersona.DCOMPersonas
Dim rsPers As ADODB.Recordset
    
    Set oPer = New COMDPersona.DCOMPersonas
    Set rsPers = New ADODB.Recordset
    Set rsPers = oPer.RecuperaDatosPersonaApeLote(psNroDoc)
    
    If Not rsPers.BOF And Not rsPers.EOF Then
        psPersCod = rsPers!cperscod
        psNombre = rsPers!cPersNombre
        pnPersoneria = rsPers!nPersPersoneria
        psDirec = rsPers!cPersDireccDomicilio
        If rsPers!nMayorEdad = 1 Then
            pbMayorEdad = True
        Else
            pbMayorEdad = False
        End If
    Else
        psPersCod = ""
        psNombre = ""
        pnPersoneria = 0
        psDirec = ""
    End If
    
    Set oPer = Nothing
    Set rsPers = Nothing
End Function
'**End GITU
'***Agregado por ELRO el 20121205, según OYP-RFC101-2012
Private Sub actualizarSeisUltimasRemuneraciones(ByVal pcCtaCod As String, ByVal pnMoneda As Integer, ByVal pnSueldos As Currency)

Dim oNCOMCaptaMovimiento As New COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNCOMCaptaDefinicion As New COMNCaptaGenerales.NCOMCaptaDefinicion
Dim oDCOMCaptaMovimiento As New COMDCaptaGenerales.DCOMCaptaMovimiento
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones
Dim oDCOMGeneral As New COMDConstSistema.DCOMGeneral
 
Dim rsCta As New ADODB.Recordset

Dim lsMovNro As String
Dim lnPorcDisp As Double
Dim lnExcedente As Double
Dim lnIntSaldo As Double
Dim ldUltMov As Date
Dim lnTasa As Double
Dim lnDiasTranscurridos As Integer
Dim lnSaldoRetiro As Currency

Set rsCta = oDCOMCaptaGenerales.GetDatosCuentaCTS(pcCtaCod)
lnSaldoRetiro = rsCta("nSaldRetiro")
lnTasa = rsCta("nTasaInteres")
ldUltMov = rsCta("dUltCierre")

lnDiasTranscurridos = DateDiff("d", ldUltMov, gdFecSis) - 1
If lnDiasTranscurridos < 0 Then
    lnDiasTranscurridos = 0
End If
lnIntSaldo = oNCOMCaptaMovimiento.GetInteres(lnSaldoRetiro, lnTasa, lnDiasTranscurridos, TpoCalcIntSimple)

lnPorcDisp = oNCOMCaptaDefinicion.GetCapParametro(gPorRetCTS)
lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
lnExcedente = 0


oDCOMCaptaMovimiento.AgregaDatosSueldosClientesCTS lsMovNro, pcCtaCod, pnMoneda, pnSueldos

Set rsCta = oDCOMCaptaMovimiento.ObtenerCapSaldosCuentasCTS(pcCtaCod, oDCOMGeneral.GetTipCambio(gdFecSis, TCFijoMes))
lnExcedente = rsCta!TotalSaldos - rsCta!TotalSueldos
If lnExcedente > 0 Then
    lnSaldoRetiro = lnExcedente * lnPorcDisp / 100
Else
    lnSaldoRetiro = 0
End If
oDCOMCaptaMovimiento.ActualizaSaldoRetiroCTS pcCtaCod, lnSaldoRetiro, lnIntSaldo
        
Set oNCOMCaptaDefinicion = Nothing
Set oNCOMCaptaMovimiento = Nothing
Set oNCOMContFunciones = Nothing
Set oDCOMCaptaMovimiento = Nothing
Set oDCOMCaptaGenerales = Nothing

End Sub
'***Fin Agregado por ELRO el 20121205*******************
'JUEZ 20130723 *************************************************
Public Sub GeneraExcelCTSExistentes(ByVal pMatLista As Variant)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim lsHoja As String
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet
Dim nLin As Long
Dim nItem As Long
Dim sColumna As String

    lsArchivo = App.Path & "\SPOOLER\CTSExistentes_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
    lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If Not lbLibroOpen Then
        Exit Sub
    End If
    nLin = 1
    lsHoja = "CTS"
    ExcelAddHoja lsHoja, xlLibro, xlHoja1
    
    xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.Range("A1:Y1").EntireColumn.Font.Size = 9
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 14
    xlHoja1.Range("A1:A1").ColumnWidth = 14
    xlHoja1.Range("B1:B1").ColumnWidth = 40
    xlHoja1.Range("C1:C1").ColumnWidth = 7
    xlHoja1.Range("D1:D1").ColumnWidth = 8
    xlHoja1.Range("E1:E1").ColumnWidth = 8
    xlHoja1.Range("F1:F1").ColumnWidth = 8
    
    xlHoja1.Cells(nLin, 1) = "CODIGO"
    xlHoja1.Cells(nLin, 2) = "NOMBRES"
    xlHoja1.Cells(nLin, 3) = "RE"
    xlHoja1.Cells(nLin, 4) = "TASA"
    xlHoja1.Cells(nLin, 5) = "MONTO"
    xlHoja1.Cells(nLin, 6) = "ITF"
    
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A" & nLin & ":F" & nLin).Interior.Color = RGB(217, 217, 217)
    
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
    
    nItem = 1
    nLin = nLin + 1
    For nItem = 1 To UBound(pMatLista)
        xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignLeft
        xlHoja1.Range("D" & nLin & ":F" & nLin).NumberFormat = "#,##0.00"
        xlHoja1.Cells(nLin, 1) = "'" & pMatLista(nItem, 1)
        xlHoja1.Cells(nLin, 2) = pMatLista(nItem, 2)
        xlHoja1.Cells(nLin, 3) = pMatLista(nItem, 3)
        xlHoja1.Cells(nLin, 4) = Format(pMatLista(nItem, 4), "#,##0.00")
        xlHoja1.Cells(nLin, 5) = Format(pMatLista(nItem, 5), "#,##0.00")
        xlHoja1.Cells(nLin, 6) = Format(pMatLista(nItem, 6), "#,##0.00")
        nLin = nLin + 1
    Next nItem
    
    gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
End Sub
'END JUEZ ******************************************************

'RIRO 20140528 ERS017
Private Sub Evento()
    If Not bLoad Then
        DoEvents
    End If
End Sub
'END RIRO
'JUEZ 20141010 Nuevos parámetros *****************************************
Private Function ValidarPlazoPF(ByVal pnPlazo As Integer, ByVal psTitular As String) As Boolean
ValidarPlazoPF = True
If pnPlazo < nParPlazoMin Or pnPlazo > nParPlazoMax Then
    If nParPlazoMin = nParPlazoMax Then
        MsgBox "El plazo debe ser " & nParPlazoMin & " días. Titular: " & psTitular, vbInformation, "Aviso"
    Else
        MsgBox "El plazo debe estar entre " & nParPlazoMin & " y " & nParPlazoMax & " días. Titular: " & psTitular, vbInformation, "Aviso"
    End If
    ValidarPlazoPF = False
End If
End Function
Private Function ValidarMedioRetiroPF(ByVal pnFormaRetiro As Integer, ByVal psTitular As String) As Boolean
ValidarMedioRetiroPF = True
If pnFormaRetiro = gCapPFFormRetMensual And Not bParFormaRetMensual Then
    MsgBox "El producto no permite la forma de retiro seleccionada. Titular: " & psTitular, vbInformation, "Aviso"
    ValidarMedioRetiroPF = False
End If
If pnFormaRetiro = gCapPFFormRetFinalPlazo And Not bParFormaRetFinPlazo Then
    MsgBox "El producto no permite la forma de retiro seleccionada. Titular: " & psTitular, vbInformation, "Aviso"
    ValidarMedioRetiroPF = False
End If
If pnFormaRetiro = gCapPFFormRetAdelantado And Not bParFormaRetIniPlazo Then
    MsgBox "El producto no permite la forma de retiro seleccionada. Titular: " & psTitular, vbInformation, "Aviso"
    ValidarMedioRetiroPF = False
End If
End Function
'END JUEZ ****************************************************************

