VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredCalendPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendario de Pagos"
   ClientHeight    =   8280
   ClientLeft      =   1530
   ClientTop       =   825
   ClientWidth     =   10200
   Icon            =   "frmCredCalendPagos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdResumen 
      Caption         =   "&Resumen"
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
      Left            =   7800
      TabIndex        =   66
      ToolTipText     =   "Resumen del Calendario de Pagos"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame FraFechaPago 
      Height          =   855
      Left            =   7515
      TabIndex        =   63
      Top             =   50
      Width           =   2535
      Begin MSMask.MaskEdBox txtFechaPago 
         Height          =   315
         Left            =   1320
         TabIndex        =   65
         ToolTipText     =   "Presione Enter"
         Top             =   360
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha de Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   400
         Width           =   1095
      End
   End
   Begin VB.Frame fraTasaAnuales 
      Caption         =   " &Tasas Anuales "
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
      Height          =   1480
      Left            =   7515
      TabIndex        =   58
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label lblTasaCostoEfectivoAnual 
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
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblTCEA 
         Caption         =   "Tasa Costo Efectivo Anual"
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblTasaEfectivaAnual 
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
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTEA 
         Caption         =   "Tasa Efectiva Anual"
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   "Condiciones"
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
      Height          =   3075
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame4 
         Caption         =   "Gracia"
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
         Height          =   585
         Left            =   3000
         TabIndex        =   21
         Top             =   1440
         Width           =   3705
         Begin VB.CheckBox ChkPerGra 
            Caption         =   "Periodo &Gracia"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1350
         End
         Begin VB.CheckBox chkPagoInteres 
            Caption         =   "Pago de Interes"
            Height          =   195
            Left            =   2160
            TabIndex        =   57
            Top             =   900
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton optTipoGracia 
            Caption         =   "Gracia en Cuotas"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   56
            Top             =   900
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optTipoGracia 
            Caption         =   "Capitalizar"
            Height          =   375
            Index           =   0
            Left            =   420
            TabIndex        =   55
            Top             =   180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CheckBox chkIncremenK 
            Caption         =   "Incrementa Capital"
            Height          =   255
            Left            =   2160
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox TxtTasaGracia 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton CmdGracia 
            Caption         =   "-->"
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
            Height          =   315
            Left            =   3480
            TabIndex        =   13
            Top             =   210
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox TxtPerGra 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "0"
            Top             =   225
            Width           =   750
         End
         Begin VB.Label LblPorcGracia 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   3240
            TabIndex        =   40
            Top             =   270
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label LblTasaGracia 
            Caption         =   "Tasa :"
            Height          =   165
            Left            =   2070
            TabIndex        =   38
            Top             =   255
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame fraGastoCom 
         Caption         =   "Gastos/Comisión"
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
         Height          =   900
         Left            =   3000
         TabIndex        =   79
         Top             =   2100
         Visible         =   0   'False
         Width           =   4155
         Begin VB.Frame fraEnvioEst 
            Height          =   615
            Left            =   1920
            TabIndex        =   82
            Top             =   240
            Width           =   2175
            Begin VB.CheckBox chkEnvioEst 
               Caption         =   "Envío Estado Cuenta"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   0
               Width           =   1935
            End
            Begin VB.ComboBox cmbEnvioEst 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   240
               Width           =   1860
            End
         End
         Begin VB.ComboBox cmbSeguroDes 
            Height          =   315
            ItemData        =   "frmCredCalendPagos.frx":030A
            Left            =   120
            List            =   "frmCredCalendPagos.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Seg. Desgravamen:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   1410
         End
      End
      Begin Spinner.uSpinner SpnCuotas 
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   1270
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   450
         Max             =   300
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
      End
      Begin MSComCtl2.DTPicker DTFecDesemb 
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   2080
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         Format          =   133890049
         CurrentDate     =   37054
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtInteres 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   1
         Top             =   840
         Width           =   615
      End
      Begin VB.Frame FraTipoPeriodo 
         Caption         =   "Tipo Periodo"
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
         Height          =   675
         Left            =   3000
         TabIndex        =   20
         Top             =   720
         Width           =   3675
         Begin VB.OptionButton OptTipoPeriodo 
            Caption         =   "Fecha Fija"
            Height          =   315
            Index           =   1
            Left            =   1470
            TabIndex        =   9
            Top             =   260
            Width           =   1035
         End
         Begin VB.Frame Frame6 
            Height          =   460
            Left            =   2550
            TabIndex        =   35
            Top             =   120
            Width           =   1065
            Begin VB.TextBox TxtDiaFijo2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   735
               MaxLength       =   2
               TabIndex        =   52
               Text            =   "00"
               Top             =   480
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.TextBox TxtDiaFijo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   615
               MaxLength       =   2
               TabIndex        =   37
               Top             =   150
               Width           =   330
            End
            Begin VB.CheckBox ChkProxMes 
               Caption         =   "Prox Mes"
               Enabled         =   0   'False
               Height          =   210
               Left            =   1020
               TabIndex        =   10
               Top             =   180
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lblDia2 
               AutoSize        =   -1  'True
               Caption         =   "&Dia 2:"
               Height          =   195
               Left            =   90
               TabIndex        =   53
               Top             =   510
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.Label LblDia 
               AutoSize        =   -1  'True
               Caption         =   "&Dia :"
               Enabled         =   0   'False
               Height          =   195
               Left            =   90
               TabIndex        =   36
               Top             =   180
               Width           =   330
            End
         End
         Begin VB.OptionButton OptTipoPeriodo 
            Caption         =   "Periodo Fijo"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   280
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.Frame FraTipoCuota 
         Caption         =   "Tipo Cuota"
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
         Height          =   540
         Left            =   3000
         TabIndex        =   19
         Top             =   165
         Width           =   3705
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Creciente"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Fijo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptTipoCuota 
            Caption         =   "Decreciente"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin Spinner.uSpinner SpnPlazo 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1650
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         Max             =   360
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
      End
      Begin VB.Frame FraComodin 
         Height          =   580
         Left            =   3000
         TabIndex        =   50
         Top             =   1380
         Visible         =   0   'False
         Width           =   2415
         Begin VB.CheckBox ChkCuotaCom 
            Caption         =   "Calendario con Cuota Comodin"
            Height          =   345
            Left            =   120
            TabIndex        =   51
            Top             =   150
            Width           =   2150
         End
      End
      Begin VB.Frame fraCuotaBalon 
         ForeColor       =   &H80000007&
         Height          =   500
         Left            =   3120
         TabIndex        =   76
         Top             =   1440
         Visible         =   0   'False
         Width           =   2745
         Begin VB.CheckBox chkCuotaBalon 
            Caption         =   "Cuotas con Periodo de Gracia con Pago de Intereses"
            Height          =   350
            Left            =   60
            TabIndex        =   77
            Top             =   -120
            Width           =   4215
         End
         Begin Spinner.uSpinner uspCuotaBalon 
            Height          =   255
            Left            =   4200
            TabIndex        =   78
            Top             =   180
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   450
            Max             =   300
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
         End
      End
      Begin VB.Frame Frame3 
         Height          =   500
         Left            =   3000
         TabIndex        =   42
         Top             =   910
         Visible         =   0   'False
         Width           =   2415
         Begin VB.CheckBox ChkQuincenal 
            Caption         =   "Calendario de Trabajadores y Directores"
            Height          =   345
            Left            =   150
            TabIndex        =   49
            Top             =   660
            Width           =   2595
         End
         Begin VB.CheckBox ChkCalMiViv 
            Caption         =   "Calendario Mi Vivienda"
            Height          =   350
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   2115
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   2100
         TabIndex        =   27
         Top             =   870
         Width           =   150
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Desembolso"
         Height          =   435
         Left            =   360
         TabIndex        =   26
         Top             =   1995
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Periodo (Dias)"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1670
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cuotas"
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   1300
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Interes (Mensual)"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label lblmonto 
         AutoSize        =   -1  'True
         Caption         =   "&Monto"
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   405
         Width           =   450
      End
   End
   Begin VB.CommandButton CmdDesembParcial 
      Caption         =   "&Desembolsos"
      Enabled         =   0   'False
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
      Left            =   2925
      TabIndex        =   41
      ToolTipText     =   "Desembolsos Parciales"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   6270
      TabIndex        =   17
      ToolTipText     =   "Salir del Calendario de Pagos"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Left            =   4590
      TabIndex        =   16
      ToolTipText     =   "Imprimir el Calendario de Pagos"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
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
      Left            =   1245
      TabIndex        =   15
      ToolTipText     =   "Generar el Calendario de Pagos"
      Top             =   7785
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Height          =   4095
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Width           =   10035
      Begin SICMACT.FlexEdit FECalend 
         Height          =   3825
         Left            =   120
         TabIndex        =   84
         Top             =   160
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   6747
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gast/Comis-Seg. Desg-Seg. Mult.-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-0-1000-1000-0-1200-1000"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-C-R"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin TabDlg.SSTab SSCalend 
      Height          =   3465
      Left            =   120
      TabIndex        =   44
      Top             =   3720
      Visible         =   0   'False
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6112
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Buen Pagador"
      TabPicture(0)   =   "frmCredCalendPagos.frx":030E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FECalBPag"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mal Pagador"
      TabPicture(1)   =   "frmCredCalendPagos.frx":032A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FECalMPag"
      Tab(1).ControlCount=   1
      Begin SICMACT.FlexEdit FECalBPag 
         Height          =   2880
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   5080
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gastos-Seg Desg-Seg Inmueb-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-1000-1000-1000-1000-1200-1000"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit FECalMPag 
         Height          =   2880
         Left            =   -74760
         TabIndex        =   46
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   5080
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gastos-Seg Desg-Seg Inmueb-Saldo Capital-Cuota + ITF"
         EncabezadosAnchos=   "400-1000-600-1200-1000-1000-1000-1000-1000-1000-1200-1200"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-R-R-R-R-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-2-3-2-2-2-2-2-2-2"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
   Begin VB.Frame FraMivivienda 
      Caption         =   "Mivivienda"
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
      Height          =   2820
      Left            =   120
      TabIndex        =   67
      Top             =   15
      Visible         =   0   'False
      Width           =   1310
      Begin VB.TextBox txtBonoBuenPagador 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   74
         Top             =   2400
         Width           =   1095
      End
      Begin VB.ComboBox cmbSegDes 
         Height          =   315
         ItemData        =   "frmCredCalendPagos.frx":0346
         Left            =   120
         List            =   "frmCredCalendPagos.frx":0350
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCuotaInicial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   69
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtValorInmueble 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   68
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "&Bono"
         Height          =   195
         Left            =   120
         TabIndex        =   75
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "&Seg Desg"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Cuota Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Valor Inmueble"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Label lblTotalCONITF 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8490
      TabIndex        =   48
      Top             =   7320
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total+ITF:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7560
      TabIndex        =   47
      Top             =   7365
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Capital:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   300
      TabIndex        =   33
      Top             =   7365
      Width           =   630
   End
   Begin VB.Label lblCapital 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   960
      TabIndex        =   32
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label lblInteres 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3720
      TabIndex        =   31
      Top             =   7320
      Width           =   1410
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5440
      TabIndex        =   30
      Top             =   7365
      Width           =   480
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6000
      TabIndex        =   29
      Top             =   7320
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "Intereses:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   2800
      TabIndex        =   28
      Top             =   7365
      Width           =   855
   End
End
Attribute VB_Name = "frmCredCalendPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredCalendPagos
'***     Descripcion:       Simulador de Calendario de Pagos a diferentes condiciones de pago
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:         12/06/2001 12:03:12 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************

Option Explicit
Dim nTipoGracia As Integer
Dim MatGracia As Variant
Dim psCtaCod As String
Dim lsCtaCodLeasing As String
Dim bGraciaGenerada As Boolean
Dim MatCalend As Variant
Dim MatCalend_2 As Variant
Dim MatResul As Variant
Dim MatResulDiff As Variant
Dim MatCalendDesParcial As Variant
Dim bTipoDesembolso As TCalendTipoDesemb
Dim bDesembParcialGenerado As Boolean
Dim nSugerAprob As Integer '1 suegerencia 2 aprobacion
Dim bMuestraGastos As Boolean
Dim bDesemParcial As Boolean
Dim MatDesPar As Variant
Dim cCtaCodG As String
Dim lnLeasing As Integer
Enum TCalendTipoDesemb
    DesembolsoTotal = 0
    DesembolsoParcial = 1
End Enum

' Para evitar la Caida en el caso que surjan Errores de Validacion en
' la Generacion del Calendario
Dim bErrorValidacion As Boolean
'Para trabajar los Gastos con Componentes
'Dim MatGastos As Variant
'Dim nNumgastos As Integer

'Para almacenar el valor del Capital antes de Capitalizar la Gracia
Dim nMontoCapInicial As Double
Dim bRenovarCredito As Boolean 'ARCV 24-10-2006
Dim nInteresAFecha As Double   '-----------------
Dim nTasaCostoEfectivoAnual As Double 'DAOR 20070402
Dim nTasaEfectivaAnual As Double 'DAOR 20070402
Dim nCuotMensBono As Double 'MAVM 20121113
Dim nCuotMens As Double 'MAVM 20121113
Private MatGastos As Variant 'DAOR 20070410
Private MatDesemb As Variant 'DAOR 20070410
Private sCtaCodRep As String 'DAOR 20070410
Dim nTotalcuotasLeasing As Currency
Dim nIntGraInicial As Double 'MAVM 20130209 ***
Dim lbLogicoBF As Integer 'ALPA 20140206 ***
Dim ldFechaBF As Date 'ALPA 20140206 ***
Dim lnMontoMivivienda As Currency 'ALPA 20141106 ***
Dim lnCuotaMivienda As Integer 'ALPA 20141106 ***
Private fbInicioSim As Boolean 'WIOR 20150210
Dim lnTasaSegDes As Double 'LUCV20180601, Agregó según ERS022-2018
Dim lnExoSeguroDesgravamen As Integer 'LUCV20180601, Agregó según ERS022-2018
Dim MatCalendSegDes As Variant 'LUCV20180601, Agregó según ERS022-2018
Dim lnMontoPoliza As Double 'LUCV20180601, Agregó según ERS022-2018
Dim lnTasaMensualSegInc As Double 'LUCV20180601, Agregó según ERS022-2018
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Public Function Renovar(ByVal nMonto As Double, ByVal dFecDes As Date, ByVal nTasa As Double, _
                        ByVal pnInteresAFecha As Double, Optional psCtaCod As String) As Variant

    TxtMonto.Text = Format(nMonto, "#0.00")
    TxtInteres.Text = Format(nTasa, "#0.00")
    DTFecDesemb.value = Format(dFecDes, "dd/mm/yyyy")
    
    bRenovarCredito = True
    nInteresAFecha = pnInteresAFecha
    sCtaCodRep = psCtaCod 'DAOR 20070410
    TxtMonto.Enabled = False
    TxtInteres.Enabled = False
    DTFecDesemb.Enabled = False
    Me.Show 1
    
    bRenovarCredito = False
    nInteresAFecha = 0
    Renovar = MatCalend
    
End Function

Public Sub Simulacion(ByVal pTipoSimulacion As TCalendTipoDesemb)
    DTFecDesemb.value = Format(gdFecSis, "dd/mm/yyyy")
    Set MatCalendDesParcial = Nothing
    If pTipoSimulacion = DesembolsoTotal Then
        CmdDesembParcial.Enabled = False
        TxtMonto.Text = "0.00"
        TxtMonto.Enabled = True
        SpnCuotas.valor = 1
        SpnCuotas.Enabled = True
        bDesemParcial = False
    Else
        bDesemParcial = True
        CmdDesembParcial.Enabled = True
        TxtMonto.Text = "0.00"
        TxtMonto.Enabled = False
        SpnCuotas.valor = 1
        SpnCuotas.Enabled = False
    End If
    bTipoDesembolso = pTipoSimulacion
    Me.Show 1
End Sub

'->***** LUCV20180601, Comentó según ERS022-2018
'Public Sub SoloMuestraMatrices(ByVal pMatCalend As Variant, ByVal pMatResul As Variant, ByVal MatGastos As Variant, ByVal nNumGasto As Integer, _
'                ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
'                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
'                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
'                ByVal pnDiasGracia As Integer, ByVal pnTasaGracia As Double, ByVal pnDiaFijo As Integer, _
'                ByVal bProxMes As Boolean, ByVal pMatGracia As Variant, ByVal pnMiViv As Integer, _
'                ByVal pnCuotaCom As Integer, ByRef MatMiViv_2 As Variant, Optional ByVal pnSugerAprob As Integer = 0, _
'                Optional ByVal pbNoMostrarCalendario As Boolean = False, _
'                Optional ByVal pnDiaFijo2 As Integer = 0, Optional ByVal pbIncrementaCapital As Boolean = False, _
'                Optional ByRef pnTasCosEfeAnu As Double, Optional ByRef psCtaCodLeasing As String = "", Optional pbLogicoBF As Integer = 0, Optional pdFechaBF As Date = CDate("1900-01-01")) ' DAOR 20070419, pnTasCosEfeAnu:Tasa Costo Efectivo Anual
'
'Dim i, j As Integer
'Dim nTotalInteres As Double
'Dim nTotalCapital As Double
'Dim nTotalGasto As Double
'Dim nTotalGastoSeg As Double
'Dim lnSalCap As Double
'Dim oCredito As COMNCredito.NCOMCredito 'DAOR 20070402
'Dim nRedondeoITF As Double 'MAVM 20121113
'Dim nTotalcuotasCONItF As Double 'MAVM 20121113
'
'        nSugerAprob = pnSugerAprob
'        txtMonto.Text = Format(pnMonto, "#0.00")
'        Txtinteres.Text = Format(pnTasaInt, "#0.0000")
'        spnCuotas.valor = Trim(str(pnNroCuotas))
'        SpnPlazo.valor = Trim(str(pnPeriodo))
'        lbLogicoBF = pbLogicoBF   'ALPA20140206
'        ldFechaBF = pdFechaBF           'ALPA20140206
'
'        'MAVM 28112010 ***
'        ChkPerGra.value = IIf(pnDiasGracia <> "0", 1, 0)
'        '***
'
'        DTFecDesemb.value = Format(pdFecDesemb, "dd/mm/yyyy")
'        OptTipoCuota(pnTipoCuota - 1).value = True
'        OptTipoPeriodo(pnTipoPeriodo - 1).value = True
'        If pnTipoPeriodo = FechaFija Then
'            TxtDiaFijo.Text = Trim(str(pnDiaFijo))
'            'If txtDiafijo.Enabled And txtDiafijo.Visible And FraDatos.Enabled Then
'                ChkProxMes.value = IIf(bProxMes, 1, 0)
'            'End If
'            'Se agrego para manejar la opcion de 2 dias fijos
'            TxtDiaFijo2.Text = Trim(str(pnDiaFijo2))
'            '*************
'        End If
'        nTipoGracia = pnTipoGracia
'        txtPerGra.Text = Trim(str(pnDiasGracia))
'        TxtTasaGracia.Text = Format(pnTasaGracia, "#0.0000")
'        Set MatGracia = Nothing
'        MatGracia = pMatGracia
'        cmdAplicar.Enabled = False
'        FraDatos.Enabled = False
'        FraFechaPago.Enabled = False 'MAVM 25102010
'        ChkCalMiViv.value = pnMiViv
'        ChkCuotaCom.value = pnCuotaCom
'
'        'Cambios para las opciones de gracia
'        If pnTipoGracia = EnCuotas - 1 Then
'            optTipoGracia(1).value = True
'        End If
'        If pnTipoGracia = Capitalizada - 1 Then
'            optTipoGracia(0).value = True
'        End If
'        'If pbIncrementaCapital = True Then chkIncremenK.value = 1
'        '***********************************
'
'        '***************************************************************************
'        'Adicionamos los Gastos
'        '***************************************************************************
'
'        '*********
'         'ReDim pMatCalend(UBound(MatGastos), 15) 'reco
'        '*********
'        If Len(Trim(psCtaCodLeasing)) = 0 Then
'            If IsArray(MatGastos) Then
'                For j = 0 To UBound(pMatCalend) - 1
'                    nTotalGasto = 0
'                    nTotalGastoSeg = 0
'                    For i = 0 To UBound(MatGastos) - 1 'nNumGasto - 1
'                    'Comentado por MAVM 20100320
'    '                    If Trim(Right(MatGastos(I, 0), 2)) = "1" And _
'    '                       (Trim(MatGastos(I, 1)) = Trim(pMatCalend(J, 1)) _
'    '                         Or Trim(MatGastos(I, 1)) = "*") Then
'    '                        nTotalGasto = nTotalGasto + CDbl(MatGastos(I, 3))
'    '                    End If
'
'                    'If ChkCalMiViv.value = 0 Then 'MAVM 20121113'WIOR 20151223- COMENTO
'                        If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
'                           (Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1217") Then
'                            nTotalGastoSeg = nTotalGastoSeg + CDbl(MatGastos(i, 3))
'                            pMatCalend(j, 6) = Format(nTotalGastoSeg, "#0.00")
'                        ElseIf (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1272") Then 'RECO20160408
'                            'nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))'RECO20160408
'                            pMatCalend(j, 14) = Format(CDbl(MatGastos(i, 3)), "#0.00") 'RECO20160408
'                        Else
'                            If Trim(MatGastos(i, 1)) = "*" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217") Then
'                                nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
'                                pMatCalend(j, 8) = Format(nTotalGasto, "#0.00")
'                            End If
'                        End If
'                    'WIOR 20151223- COMENTO
'                    '                    'MAVM 20121113 Two
'                    '                    Else
'                    '                        If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
'                    '                           (Trim(MatGastos(i, 1)) = Trim(pMatCalend(J, 1))) Then
'                    '                            If Left(Trim(MatGastos(i, 2)), 2) <> "PO" Then
'                    '                                nTotalGasto = CDbl(MatGastos(i, 3))
'                    '                                pMatCalend(J, 6) = Format(nTotalGasto, "#0.00")
'                    '                            Else
'                    '                                nTotalGasto = CDbl(MatGastos(i, 3))
'                    '                                pMatCalend(J, 9) = Format(nTotalGasto, "#0.00")
'                    '                                'ALPA20140619****************************************
'                    '                                pMatResul(J, 9) = Format(nTotalGasto, "#0.00")
'                    '                            End If
'                    '                        Else
'                    '                            If Trim(MatGastos(i, 1)) = "*" Then
'                    '                                If Left(Trim(MatGastos(i, 2)), 2) = "CO" And i <> UBound(MatGastos) - 1 Then
'                    '                                    nTotalGasto = CDbl(MatGastos(i, 3)) '+ 9 'CDbl(MatGastos(UBound(MatGastos) - 1, 3))
'                    '                                    pMatCalend(J, 8) = Format(nTotalGasto, "#0.00")
'                    '                                Else
'                    '                                    'If i <> UBound(MatCalend) + 2 Then
'                    '                                    If i <> UBound(MatGastos) - 1 Then
'                    '                                        nTotalGasto = CDbl(MatGastos(i, 3))
'                    '                                        pMatCalend(J, 9) = Format(nTotalGasto, "#0.00")
'                    '                                        'ALPA20140619****************************************
'                    '                                        pMatResul(J, 9) = Format(nTotalGasto, "#0.00")
'                    '                                    End If
'                    '                                End If
'                    '                            Else
'                    '                                If Left(Trim(MatGastos(i, 2)), 2) <> "PO" And _
'                    '                                   (Trim(MatGastos(i, 1)) = Trim(pMatCalend(J, 1))) Then
'                    '                                    nTotalGasto = CDbl(MatGastos(i, 3))
'                    '                                    pMatCalend(J, 6) = Format(nTotalGasto, "#0.00")
'                    '                                Else
'                    '                                    If (Trim(MatGastos(i, 1)) = Trim(pMatCalend(J, 1))) Then
'                    '                                        nTotalGasto = CDbl(MatGastos(i, 3))
'                    '                                        pMatCalend(J, 9) = Format(nTotalGasto, "#0.00")
'                    '                                        'ALPA20140619****************************************
'                    '                                        pMatResul(J, 9) = Format(nTotalGasto, "#0.00")
'                    '                                    End If
'                    '                                End If
'                    '                            End If
'                    '                        End If
'                    '                    End If
'                    Next i
'                    'Add By GITU 06-08-2008
'                    'Descomentar cuando esten seguros de los cambios GITU
'    '                If j > 0 Then
'    '                    If InStr(Trim(Str(nTotalGasto)), ".") > 0 Then
'    '                        pMatCalend(0, 6) = Format(pMatCalend(0, 6) + (nTotalGasto - Val(Left(Trim(Str(nTotalGasto)), InStr(Trim(nTotalGasto), ".") - 1))), "#0.00")
'    '                        pMatCalend(j, 6) = Format(Val(Left(Trim(nTotalGasto), InStr(Trim(nTotalGasto), ".") - 1)), "#0.00")
'    '                    Else
'    '                        pMatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'    '                    End If
'    '                Else
'    '                    pMatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'    '                End If
'                    'End GITU
'
'                    'MAVM 20100320
'                    'pMatCalend(J, 6) = Format(nTotalGasto, "#0.00")
'                Next j
'            End If
'        End If
'
'        'WIOR 20151223 - COMENTO
'        '        If Me.ChkCalMiViv.value = 1 Then
'        '            'MAVM 20121113 ***
'        '                For i = 0 To UBound(pMatCalend) - 1
'        '                    pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 3)) + CDbl(pMatCalend(i, 4)) + CDbl(pMatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(pMatCalend(i, 8)) + CDbl(pMatCalend(i, 9)), "#0.00")
'        '                Next i
'        '            '***
'        '            'Cargar Matrices a FlexGrid
'        '            nTotalInteres = 0
'        '            nTotalCapital = 0
'        '            LimpiaFlex FECalBPag
'        '            For i = 0 To UBound(pMatCalend) - 1
'        '                FECalBPag.AdicionaFila
'        '                FECalBPag.TextMatrix(i + 1, 1) = Trim(pMatCalend(i, 0))
'        '                FECalBPag.TextMatrix(i + 1, 2) = Trim(pMatCalend(i, 1))
'        '                'MAVM 20121113 ***
'        '                'pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 6)), "#0.00")
'        '                pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)), "#0.00")
'        '                '***
'        ''                pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "#0.00")
'        '
'        '                FECalBPag.TextMatrix(i + 1, 3) = Format(CDbl(pMatCalend(i, 2)), "#0.00")
'        ''                If i < 6 Then
'        ''                    FECalBPag.TextMatrix(i + 1, 3) = CCur(pMatCalend(i, 2)) + CCur(Trim(MatResul(i, 6))) - CCur(Trim(MatCalend(i, 6)))
'        ''                Else
'        ''                    FECalBPag.TextMatrix(i + 1, 3) = Trim(MatResul(i, 2)) - Trim(MatResul(i, 3)) + Trim(MatCalend(i, 3)) + Trim(MatCalend(i, 4)) - Trim(MatResul(i, 4)) - Trim(MatResul(i, 6)) + Trim(MatCalend(i, 6))
'        ''                End If
'        '
'        '                FECalBPag.row = i + 1
'        '                FECalBPag.col = 3
'        '                FECalBPag.CellForeColor = vbBlue
'        '                FECalBPag.TextMatrix(i + 1, 4) = Trim(pMatCalend(i, 3))
'        '                FECalBPag.TextMatrix(i + 1, 5) = Trim(pMatCalend(i, 4))
'        '                FECalBPag.TextMatrix(i + 1, 6) = Trim(pMatCalend(i, 5))
'        '
'        '                'MAVM 20121113 ***
'        '                'FECalBPag.TextMatrix(i + 1, 7) = Trim(pMatCalend(i, 6))
'        '                'FECalBPag.TextMatrix(i + 1, 8) = Trim(pMatCalend(i, 7))
'        '                'nTotalCapital = nTotalCapital + CDbl(Trim(pMatCalend(i, 3)))
'        '                ''nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(I, 4)))
'        '                ''11-05-2006
'        '                'nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(i, 4))) + CDbl(Trim(pMatCalend(i, 5))) '+ CDbl(Trim(pMatCalend(i, 6)))
'        '                'FECalBPag.TextMatrix(i + 1, 9) = Format(CDbl(pMatCalend(i, 2))) + gITF.fgITFCalculaImpuesto(Format(CDbl(pMatCalend(i, 2))))
'        '                FECalBPag.TextMatrix(i + 1, 7) = "0.00" 'Trim(pMatCalend(i, 8))
'        '                pMatCalend(i, 6) = pMatResul(i, 6)
'        '                FECalBPag.TextMatrix(i + 1, 8) = Trim(pMatCalend(i, 6))
'        '                FECalBPag.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 9))
'        '                nTotalCapital = nTotalCapital + CDbl(Trim(pMatCalend(i, 3)))
'        '                nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(i, 4))) + CDbl(Trim(pMatCalend(i, 5))) '+ CDbl(Trim(pMatCalend(i, 6)))
'        '                FECalBPag.TextMatrix(i + 1, 10) = Trim(pMatCalend(i, 7))
'        '
'        '                nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))))
'        '                If nRedondeoITF > 0 Then
'        '                    FECalBPag.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))) - nRedondeoITF + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "0.00")
'        '                Else
'        '                    FECalBPag.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))) + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "0.00")
'        '                End If
'        '                '***
'        '            Next i
'        '            lblCapital.Caption = Format(nTotalCapital, "#0.00")
'        '            lblInteres.Caption = Format(nTotalInteres, "#0.00")
'        '            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
'        '            'Carga Mal Pagador
'        '
'        '            LimpiaFlex FECalMPag
'        '            'pMatResul = UnirMatricesCalendarioIguales(pMatCalend, pMatResul, pnMonto) 'MAVM 20121113 ***
'        '                For i = 0 To UBound(pMatCalend) - 1
'        '                    pMatResul(i, 2) = Format(CDbl(pMatResul(i, 3)) + CDbl(pMatResul(i, 4)) + CDbl(pMatResul(i, 5)) + CDbl(pMatResul(i, 6)) + CDbl(pMatResul(i, 8)) + CDbl(pMatResul(i, 9)), "#0.00")
'        '                Next i
'        '            For i = 0 To UBound(pMatCalend) - 1
'        '                FECalMPag.AdicionaFila
'        '                FECalMPag.TextMatrix(i + 1, 1) = Trim(pMatResul(i, 0))
'        '                FECalMPag.TextMatrix(i + 1, 2) = Trim(pMatResul(i, 1))
'        '                If Trim(pMatCalend(i, 9)) <> Trim(pMatResul(i, 9)) Then
'        '                pMatResul(i, 2) = pMatResul(i, 2) - Trim(pMatResul(i, 9)) + Trim(pMatCalend(i, 9))
'        '                End If
'        '                FECalMPag.TextMatrix(i + 1, 3) = Trim(pMatResul(i, 2))
'        '                FECalMPag.row = i + 1
'        '                FECalMPag.col = 3
'        '                FECalMPag.CellForeColor = vbBlue
'        '                FECalMPag.TextMatrix(i + 1, 4) = Trim(pMatResul(i, 3))
'        '                FECalMPag.TextMatrix(i + 1, 5) = Trim(pMatResul(i, 4))
'        '                FECalMPag.TextMatrix(i + 1, 6) = Trim(pMatResul(i, 5))
'        '                FECalMPag.TextMatrix(i + 1, 7) = "0.00" 'Trim(pMatResul(i, 8))
'        '
'        '                'MAVM 20121113 ***
'        '                'FECalMPag.TextMatrix(i + 1, 8) = Trim(pMatResul(i, 7))
'        '                'FECalMPag.TextMatrix(i + 1, 9) = Format(CDbl(pMatResul(i, 2))) + gITF.fgITFCalculaImpuesto(Format(CDbl(pMatResul(i, 2))))
'        '                FECalMPag.TextMatrix(i + 1, 8) = Trim(pMatResul(i, 6))
'        '                FECalMPag.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 9)) 'Trim(pMatResul(i, 9))
'        '                FECalMPag.TextMatrix(i + 1, 10) = Trim(pMatResul(i, 7))
'        '
'        '                nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(pMatResul(i, 2))))
'        '                If nRedondeoITF > 0 Then
'        '                    FECalMPag.TextMatrix(i + 1, 11) = Format(CDbl(pMatResul(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatResul(i, 2))) - nRedondeoITF, "0.00")
'        '                Else
'        '                    FECalMPag.TextMatrix(i + 1, 11) = Format(CDbl(pMatResul(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatResul(i, 2))), "0.00")
'        '                End If
'        ''                If Trim(pMatCalend(i, 9)) <> Trim(pMatResul(i, 9)) Then
'        ''                    FECalMPag.TextMatrix(i + 1, 11) = FECalMPag.TextMatrix(i + 1, 11) - Trim(pMatResul(i, 9)) + Trim(pMatCalend(i, 9))
'        ''                End If
'        '                '***
'        '
'        '            Next i
'        '            FECalBPag.row = 1
'        '            FECalBPag.TopRow = 1
'        '            FECalMPag.TopRow = 1
'        '            SSCalend.Tab = 0
'        '
'        '        Else
'            nTotalInteres = 0
'            nTotalCapital = 0
'            'lnSalCap = Val(TxtMonto.Text)
'            LimpiaFlex FECalend
'            For i = 0 To UBound(pMatCalend) - 1
'                FECalend.AdicionaFila
'                FECalend.TextMatrix(i + 1, 1) = Trim(pMatCalend(i, 0))
'                FECalend.TextMatrix(i + 1, 2) = Trim(pMatCalend(i, 1))
'
'                'MAVM 201004 ***
'                'pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 6)), "#0.00")
'                'ALPA 20111025
'                If Mid(psCtaCodLeasing, 6, 3) = "515" Or Mid(psCtaCodLeasing, 6, 3) = "516" Then
'                    pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)), "#0.00")
'                Else
'                    pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 6)) + CDbl(pMatCalend(i, 8)), "#0.00")
'                End If
'                '***
'                'Modify by Gitu 22-08-08
'                'Descomentar cuando esten seguros de los cambios GITU
''                If Val(pMatCalend(I, 2)) <> (Val(Trim(pMatCalend(I, 3))) + Val(Trim(pMatCalend(I, 4))) + Val(Trim(pMatCalend(I, 5))) + Val(Trim(pMatCalend(I, 6)))) Then
''                    pMatCalend(I, 2) = Format(Val(Trim(pMatCalend(I, 3))) + Val(Trim(pMatCalend(I, 4))) + Val(Trim(pMatCalend(I, 5))) + Val(Trim(pMatCalend(I, 6))), "##0.00")
''                Else
''                    pMatCalend(I, 2) = pMatCalend(I, 2)
''                End If
'
'                'MAVM 20130312
'                'FECalend.TextMatrix(i + 1, 3) = pMatCalend(i, 2)
'                If nTipoGracia = 6 Then
'                    pMatCalend(i, 2) = Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 11)))
'                    'FECalend.TextMatrix(i + 1, 3) = Format(Trim(pMatCalend(i, 2)), "#0.00")
'                    FECalend.TextMatrix(i + 1, 3) = Format(Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 14))), "#0.00") 'RECO
'                Else
'                    'FECalend.TextMatrix(i + 1, 3) = Trim(pMatCalend(i, 2))
'                    FECalend.TextMatrix(i + 1, 3) = Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 14)))
'                End If
'                '***
'
'                FECalend.row = i + 1
'                FECalend.Col = 3
'                FECalend.CellForeColor = vbBlue
'                FECalend.TextMatrix(i + 1, 4) = Trim(pMatCalend(i, 3))
'
'                'MAVM 20130209 *** 'Interes Comp
'                'FECalend.TextMatrix(i + 1, 5) = Trim(pMatCalend(i, 4))
'                'If nTipoGracia = 6 Then
'                '    pMatCalend(i, 4) = Trim(CDbl(pMatCalend(i, 4)) + CDbl(pMatCalend(i, 11)))
'                '    FECalend.TextMatrix(i + 1, 5) = Format(Trim(pMatCalend(i, 4)), "#0.00")
'                'Else
'                    FECalend.TextMatrix(i + 1, 5) = Trim(pMatCalend(i, 4))
'                'End If
'
'                FECalend.TextMatrix(i + 1, 6) = Trim(pMatCalend(i, 5))
'                'ALPA 20110526 *Leasing
'                If Len(Trim(psCtaCodLeasing)) = 18 Then
'                        FECalend.TextMatrix(i + 1, 7) = Trim(pMatCalend(i, 6))
'                        FECalend.TextMatrix(i + 1, 8) = Format(Trim(pMatCalend(i, 8)), "#0.00")
'                Else
'                        FECalend.TextMatrix(i + 1, 7) = Trim(pMatCalend(i, 8))
'                        FECalend.TextMatrix(i + 1, 8) = Trim(pMatCalend(i, 6)) 'RECO20160408
'                End If
'                FECalend.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 14)) 'RECO20160408
'                'FECalend.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 8)) 'RECO20160408
'                'FECalend.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 7)) 'RECO20160408
'                FECalend.TextMatrix(i + 1, 10) = Trim(pMatCalend(i, 7)) 'RECO20160408
'                'Descomentar cuando esten seguros de los cambios GITU
''                FECalend.TextMatrix(I + 1, 8) = lnSalCap - Trim(pMatCalend(I, 3))
''                lnSalCap = lnSalCap - Trim(pMatCalend(I, 3))
'
'                'MAVM 20130209 ***
'                If Not (i = 0 And nTipoGracia = 6) Then
'                    nTotalCapital = nTotalCapital + CDbl(Trim(pMatCalend(i, 3)))
'                    nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(i, 4))) + CDbl(Trim(pMatCalend(i, 5))) '+ CDbl(Trim(pMatCalend(i, 6)))
'                End If
'                '***
'
'                'nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(I, 4)))
'                'Modify By Gitu 22-08-08 se agrega los decimales de las
'                'cuotas mayores a la primera a la esta.
'                'FECalend.TextMatrix(I + 1, 9) = Format(CDbl(pMatCalend(I, 2))) + gITF.fgITFCalculaImpuesto(Format(CDbl(pMatCalend(I, 2))))
'
'                'MAVM 20121113 ***
'                'FECalend.TextMatrix(i + 1, 10) = Format(CDbl(pMatCalend(i, 2))) + gITF.fgITFCalculaImpuesto(Format(CDbl(pMatCalend(i, 2))))
'                nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))))
'                If nRedondeoITF > 0 Then
'                    'FECalend.TextMatrix(i + 1, 10) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))) - nRedondeoITF, "0.00") 'RECO20160408
'                    FECalend.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))) - nRedondeoITF, "0.00") 'RECO20160408
'                Else
'                    'FECalend.TextMatrix(i + 1, 10) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))), "0.00") 'RECO20160408
'                    FECalend.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))), "0.00") 'RECO20160408
'                End If
'                '***
'
'                'Descomentar cuando esten seguros de los cambios GITU
''                If I > 0 Then
''                    If InStr(FECalend.TextMatrix(I + 1, 9), ".") <> 0 Then
''                        FECalend.TextMatrix(1, 9) = Format(Val(FECalend.TextMatrix(1, 9)) + Round(Val(FECalend.TextMatrix(I + 1, 9)) - Val(Left(FECalend.TextMatrix(I + 1, 9), InStr(FECalend.TextMatrix(I + 1, 9), ".") - 1)), 2))
''                        FECalend.TextMatrix(I + 1, 9) = Format(Val(Left(FECalend.TextMatrix(I + 1, 9), InStr(FECalend.TextMatrix(I + 1, 9), ".") - 1)))
''                    End If
''                End If
'
'            If Not (pnTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
'                'MAVM 20121113 ***
'                'nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 10)) 'RECO20160408
'                nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 11)) 'RECO20160408
'                '***
'            End If
'
'            Next i
'
'            If pnTipoGracia = 6 Then
'                FECalend.TextMatrix(1, 3) = ""
'                FECalend.TextMatrix(1, 5) = ""
'                FECalend.TextMatrix(1, 7) = ""
'                FECalend.TextMatrix(1, 8) = ""
'                'FECalend.TextMatrix(1, 10) = "" 'reco
'                FECalend.TextMatrix(1, 11) = "" 'reco
'            End If
'            '***
'
'            lblCapital.Caption = Format(nTotalCapital, "#0.00")
'
'            LblInteres.Caption = Format(nTotalInteres, "#0.00")
'            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
'            FECalend.row = 1
'            FECalend.TopRow = 1
'        'End If 'WIOR 20151223 - COMENTO
'
'        '**DAOR 20070402, Desarrollo de la Tasa Costo Efectivo Anual (Según el SIAFC)
'        fraTasaAnuales.Visible = True
'        Set oCredito = New COMNCredito.NCOMCredito
'            nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(pnTasaInt, 360) * 100, 2)
'            'MAVM 28112009 ***
'            'nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFecDesemb, pnMonto, pMatCalend)
'            'nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFecDesemb, pnMonto, pMatCalend, pnTasaInt, lsCtaCodLeasing)
'
'        'MAVM 20130305 ***
'        If pnTipoGracia = 6 Then
'            Dim Y As Integer
'            Dim MatCalendTemp() As String
'            ReDim MatCalendTemp(UBound(pMatCalend) - 1, 13)
'            For i = 0 To UBound(pMatCalend) - 2
'                For Y = 0 To 13
'                    MatCalendTemp(i, Y) = pMatCalend(i + 1, Y)
'                Next Y
'            Next i
'            Erase pMatCalend
'            ReDim pMatCalend(UBound(MatCalendTemp), 13)
'
'            For i = 0 To UBound(MatCalendTemp)
'                For Y = 0 To 13
'                    pMatCalend(i, Y) = MatCalendTemp(i, Y)
'                Next Y
'            Next i
'            Erase MatCalendTemp
'        End If
'        '***
'
'            'nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFecDesemb, IIf(ChkCalMiViv.value = 0, pnMonto, pnMonto - 12500), pMatCalend, pnTasaInt, lsCtaCodLeasing) 'MAVM 20121113
'            nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFecDesemb, IIf(ChkCalMiViv.value = 0, pnMonto, pnMonto - 12500), pMatCalend, pnTasaInt, lsCtaCodLeasing, pnTipoPeriodo) 'JUEZ 20140814
'            '***
'            lblTasaCostoEfectivoAnual.Caption = nTasaCostoEfectivoAnual & " %"
'            lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
'
'            pnTasCosEfeAnu = nTasaCostoEfectivoAnual
'        Set oCredito = Nothing
'        '***********************************************************
'        If UBound(pMatCalend) = 0 Then
'            cmdImprimir.Enabled = False
'        Else
'            cmdImprimir.Enabled = True
'        End If
'        MatCalend = pMatCalend
'        MatResul = pMatResul
'
'
'        If Mid(psCtaCodLeasing, 6, 3) = "515" Or Mid(psCtaCodLeasing, 6, 3) = "516" Then
'            If nTotalcuotasLeasing > 0 Then
'                nTotalcuotasLeasing = Format(nTotalcuotasLeasing + fgITFCalculaImpuesto(CDbl(nTotalcuotasLeasing)), "0.00")
'            Else
'                lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
'            End If
'        Else
'        'MAVM 20121113 ***
'        'lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
'        lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
'        '***
'        End If
'
'        Me.Show 1
'End Sub
'<-***** Fin LUCV20180601
'->***** LUCV20180601, Agregó según ERS022-2018
Public Sub SoloMuestraMatrices(ByVal pMatCalend As Variant, ByVal pMatResul As Variant, _
                ByVal MatGastos As Variant, ByVal nNumGasto As Integer, _
                ByVal pnMonto As Double, ByVal pnTasaInt As Double, _
                ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, _
                ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, _
                ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, _
                ByVal pnTasaGracia As Double, _
                ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, _
                ByVal pMatGracia As Variant, _
                ByVal pnMiViv As Integer, _
                ByVal pnCuotaCom As Integer, _
                ByRef MatMiViv_2 As Variant, _
                Optional ByVal pnSugerAprob As Integer = 0, _
                Optional ByVal pbNoMostrarCalendario As Boolean = False, _
                Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional ByVal pbIncrementaCapital As Boolean = False, _
                Optional ByRef pnTasCosEfeAnu As Double, _
                Optional ByRef psCtaCodLeasing As String = "", _
                Optional pbLogicoBF As Integer = 0, _
                Optional pdFechaBF As Date = CDate("1900-01-01"))
                
Dim i, j As Integer
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim lnSalCap As Double
Dim oCredito As COMNCredito.NCOMCredito
Dim nRedondeoITF As Double
Dim nTotalcuotasCONItF As Double
    
    nSugerAprob = pnSugerAprob
    TxtMonto.Text = Format(pnMonto, "#0.00")
    TxtInteres.Text = Format(pnTasaInt, "#0.0000")
    SpnCuotas.valor = Trim(str(pnNroCuotas))
    SpnPlazo.valor = Trim(str(pnPeriodo))
    lbLogicoBF = pbLogicoBF   'ALPA20140206
    ldFechaBF = pdFechaBF     'ALPA20140206
    ChkPerGra.value = IIf(pnDiasGracia <> "0", 1, 0)
    
    DTFecDesemb.value = Format(pdFecDesemb, "dd/mm/yyyy")
    OptTipoCuota(pnTipoCuota - 1).value = True
    OptTipoPeriodo(pnTipoPeriodo - 1).value = True
    If pnTipoPeriodo = FechaFija Then
        TxtDiaFijo.Text = Trim(str(pnDiaFijo))
        ChkProxMes.value = IIf(bProxMes, 1, 0)
        'Se agrego para manejar la opcion de 2 dias fijos
        TxtDiaFijo2.Text = Trim(str(pnDiaFijo2))
        '*************
    End If
    nTipoGracia = pnTipoGracia
    TxtPerGra.Text = Trim(str(pnDiasGracia))
    TxtTasaGracia.Text = Format(pnTasaGracia, "#0.0000")
    Set MatGracia = Nothing
    MatGracia = pMatGracia
    cmdAplicar.Enabled = False
    FraDatos.Enabled = False
    FraFechaPago.Enabled = False 'MAVM 25102010
    ChkCalMiViv.value = pnMiViv
    ChkCuotaCom.value = pnCuotaCom
    
    'Cambios para las opciones de gracia
    If pnTipoGracia = EnCuotas - 1 Then
        optTipoGracia(1).value = True
    End If
    If pnTipoGracia = Capitalizada - 1 Then
        optTipoGracia(0).value = True
    End If

    'Adicionamos los Gastos
    If Len(Trim(psCtaCodLeasing)) = 0 Then
        If IsArray(MatGastos) Then
            For j = 0 To UBound(pMatCalend) - 1
                nTotalGasto = 0
                nTotalGastoSeg = 0
                For i = 0 To UBound(MatGastos) - 1
                    '1217-Seguro de Desgravamen
                    'If Trim(Right(MatGastos(i, 0), 2)) = "1" And (Trim(MatGastos(i, 1)) = Trim(pMatCalend(J, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1217") Then 'LUCV20180601, Comentó según ERS022-2018
                        'nTotalGastoSeg = nTotalGastoSeg + CDbl(MatGastos(i, 3)) 'LUCV20180601, Comentó según ERS022-2018
                        'pMatCalend(j, 6) = Format(nTotalGastoSeg, "#0.00") 'LUCV20180601, Comentó según ERS022-2018
                    '1272-Seguro Multiriesgo MYPE
                    'ElseIf (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(J, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1272") Then 'LUCV20180601, Comentó según ERS022-2018
                    If (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1272") Then 'LUCV20180601, Comentó según ERS022-2018
                        pMatCalend(j, 14) = Format(CDbl(MatGastos(i, 3)), "#0.00")
                    '(*)-Todos los Gastos. (No incluye 1217)
                    Else
                        'If Trim(MatGastos(i, 1)) = "*" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217") Then'LUCV20180601, Comentó Según ERS022-2018
                        If (Trim(MatGastos(i, 1)) = "*" And (Trim(Right(MatGastos(i, 2), 4)) <> "1231" And Trim(Right(MatGastos(i, 2), 4)) <> "1279")) _
                            Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(pMatCalend(j, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217" _
                            And (Trim(Right(MatGastos(i, 2), 4)) <> "1231" And Trim(Right(MatGastos(i, 2), 4)) <> "1279")) Then 'LUCV20180601, Comentó Según ERS022-2018
                            nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
                            'pMatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                            pMatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                        End If
                    End If
                 Next i
            Next j
        End If
    End If
        
    nTotalInteres = 0
    nTotalCapital = 0
    LimpiaFlex FECalend
    For i = 0 To UBound(pMatCalend) - 1
        FECalend.AdicionaFila
        FECalend.TextMatrix(i + 1, 1) = Trim(pMatCalend(i, 0)) 'Fecha Venc.
        FECalend.TextMatrix(i + 1, 2) = Trim(pMatCalend(i, 1)) 'Nro. Cuota
        
    '**ARLO20180712 ERS042 - 2018
        Set objProducto = New COMDCredito.DCOMCredito
        If objProducto.GetResultadoCondicionCatalogo("N0000087", Mid(psCtaCodLeasing, 6, 3)) Then
        'If Mid(psCtaCodLeasing, 6, 3) = "515" Or Mid(psCtaCodLeasing, 6, 3) = "516" Then
    '**ARLO20180712 ERS042 - 2018
            pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)), "#0.00")
        Else
            'pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 6)) + CDbl(pMatCalend(i, 8)), "#0.00") 'LUCV20180601, Comentó según ERS022-2018
            'pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl( (i, 6)) + CDbl(pMatCalend(i, 8)) + CDbl(pMatCalend(i, 14)), "#0.00") 'APRI20180821 ERS061-2018 'LUCV20180601, Comentó según ERS022-2018
            pMatCalend(i, 2) = Format(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 6)) + CDbl(pMatCalend(i, 14)), "#0.00")  'LUCV20180601, según ERS022-2018 (SegDesg. Forma parte del importe de cuota)
        End If
        
        If nTipoGracia = 6 Then ' Capitalizada
            pMatCalend(i, 2) = Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 11)))
            'FECalend.TextMatrix(i + 1, 3) = Format(Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 14))), "#0.00") 'Monto Cuota
            FECalend.TextMatrix(i + 1, 3) = Format(Trim(CDbl(pMatCalend(i, 2))), "#0.00") 'APRI20180821 ERS061-2018
        Else
            'FECalend.TextMatrix(i + 1, 3) = Trim(CDbl(pMatCalend(i, 2)) + CDbl(pMatCalend(i, 14))) 'Monto Cuota
            FECalend.TextMatrix(i + 1, 3) = Trim(CDbl(pMatCalend(i, 2))) 'APRI20180821 ERS061-2018
        End If
        
        FECalend.row = i + 1
        FECalend.col = 3
        FECalend.CellForeColor = vbBlue
        FECalend.TextMatrix(i + 1, 4) = Trim(pMatCalend(i, 3)) 'Amortizacion capital
        FECalend.TextMatrix(i + 1, 5) = Trim(CDbl(pMatCalend(i, 4)) + CDbl(pMatCalend(i, 5))) 'Interes Comp. + Interes Gracia
        'FECalend.TextMatrix(i + 1, 6) = Trim(pMatCalend(i, 5)) 'Interes Gracia 'LUCV20180601, Comentó según ERS022-2018
        
        'ALPA 20110526 *Leasing
        If Len(Trim(psCtaCodLeasing)) = 18 Then
            FECalend.TextMatrix(i + 1, 7) = Trim(CDbl(pMatCalend(i, 6))) 'Gast./Comis.
            FECalend.TextMatrix(i + 1, 8) = Format(Trim(pMatCalend(i, 8)), "#0.00") 'Seg. Desg.
        Else
            'FECalend.TextMatrix(i + 1, 7) = Trim(pMatCalend(i, 8)) 'Gast./Comis. 'LUCV20180601, Comentó Según ERS022-2018
            'FECalend.TextMatrix(i + 1, 8) = Trim(pMatCalend(i, 6)) 'Seg. Desg.   'LUCV20180601, Comentó Según ERS022-2018
            FECalend.TextMatrix(i + 1, 7) = Trim(CDbl(pMatCalend(i, 6)) + CDbl(pMatCalend(i, 15)) + CDbl(pMatCalend(i, 16))) 'Gast./Comis.  'LUCV20180601, Modificó posición de Gastos (ERS022-2018)
            FECalend.TextMatrix(i + 1, 8) = Trim(pMatCalend(i, 8)) 'Seg. Desg.    'LUCV20180601, Modificó posición de Gastos (ERS022-2018)
        End If
        FECalend.TextMatrix(i + 1, 9) = Trim(pMatCalend(i, 14)) 'Seg. Mult.
        FECalend.TextMatrix(i + 1, 10) = Trim(pMatCalend(i, 7)) 'Saldo capital
        If Not (i = 0 And nTipoGracia = 6) Then
            nTotalCapital = nTotalCapital + CDbl(Trim(pMatCalend(i, 3)))
            nTotalInteres = nTotalInteres + CDbl(Trim(pMatCalend(i, 4))) + CDbl(Trim(pMatCalend(i, 5)))
        End If

        'MAVM 20121113 ***
        nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))))
        If nRedondeoITF > 0 Then
            FECalend.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))) - nRedondeoITF, "0.00") 'Cuota+ITF
        Else
            FECalend.TextMatrix(i + 1, 11) = Format(CDbl(pMatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(pMatCalend(i, 2))), "0.00") 'Cuota+ITF
        End If
        '***
        If Not (pnTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
            nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 11))
        End If
    Next i
        
    lblCapital.Caption = Format(nTotalCapital, "#0.00")
    lblInteres.Caption = Format(nTotalInteres, "#0.00")
    lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
    FECalend.row = 1
    FECalend.TopRow = 1
        
    '**DAOR 20070402, Desarrollo de la Tasa Costo Efectivo Anual (Según el SIAFC)
    fraTasaAnuales.Visible = True
    Set oCredito = New COMNCredito.NCOMCredito
    
    nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(pnTasaInt, 360) * 100, 2)
    nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(pdFecDesemb, IIf(ChkCalMiViv.value = 0, pnMonto, pnMonto - 12500), pMatCalend, pnTasaInt, lsCtaCodLeasing, pnTipoPeriodo) 'JUEZ 20140814
    lblTasaCostoEfectivoAnual.Caption = nTasaCostoEfectivoAnual & " %"
    lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
    
    pnTasCosEfeAnu = nTasaCostoEfectivoAnual
    Set oCredito = Nothing
    
    If UBound(pMatCalend) = 0 Then
        cmdImprimir.Enabled = False
    Else
        cmdImprimir.Enabled = True
    End If
    MatCalend = pMatCalend
    MatResul = pMatResul
    
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000088", Mid(psCtaCodLeasing, 6, 3)) Then
    'If Mid(psCtaCodLeasing, 6, 3) = "515" Or Mid(psCtaCodLeasing, 6, 3) = "516" Then
    '**ARLO20180712 ERS042 - 2018
        If nTotalcuotasLeasing > 0 Then
            nTotalcuotasLeasing = Format(nTotalcuotasLeasing + fgITFCalculaImpuesto(CDbl(nTotalcuotasLeasing)), "0.00")
        Else
            lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
        End If
    Else
    'MAVM 20121113 ***
    lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
    '***
    End If
    
    Me.Show 1
End Sub
'<-***** Fin LUCV20180601


'Modificado CACV para trabajar los Gastos con los Componentes
Public Function Inicio(ByVal pnMonto As Double, ByVal pnTasaInt As Double, ByVal pnNroCuotas As Integer, _
                ByVal pnPeriodo As Integer, ByVal pdFecDesemb As Date, ByVal pnTipoCuota As TCalendTipoCuota, _
                ByVal pnTipoPeriodo As TCalendTipoPeriodo, ByVal pnTipoGracia As TCalendTipoGracia, _
                ByVal pnDiasGracia As Integer, ByVal pnTasaGracia As Double, ByVal pnDiaFijo As Integer, _
                ByVal bProxMes As Boolean, ByVal pMatGracia As Variant, ByVal pnMiViv As Integer, _
                ByVal pnCuotaCom As Integer, ByRef MatMiViv_2 As Variant, _
                Optional ByVal pnSugerAprob As Integer = 0, Optional ByVal pbNoMostrarCalendario As Boolean = False, _
                Optional ByVal pbDesemParcial As Boolean = False, Optional ByVal pMatDesPar As Variant = "", _
                Optional ByVal bQuincenal As Boolean, Optional ByVal psCtaCod As String, _
                Optional ByRef pbErrorValidacion As Boolean = False, Optional ByVal pnDiaFijo2 As Integer = 0, _
                Optional ByVal pbIncrementaCapital As Boolean = False, Optional ByVal bGracia As Boolean = False, _
                Optional ByVal dFechaPago As Date, Optional pnLeasing As Integer = 0, _
                Optional psCredCodLeasing As String = "", Optional ByVal pnValorInmueb As Double, _
                Optional ByRef pnIntGraInicial As Double = 0, Optional ByVal pnCuotaBalon As Integer = 0, _
                Optional pbLogicoBF As Boolean = False, Optional pdFechaBF As Date = CDate("1900-01-01"), _
                Optional pnMontoMivivienda As Currency = 0#, Optional pnCuotaMivienda As Integer, _
                Optional ByVal pArrMIVIVIENDA As Variant, Optional pnTasaSegDes As Double = 0, _
                Optional ByRef pMatCalendSegDes As Variant = Nothing, _
                Optional ByVal pnExoSeguroDesgravamen As Integer = 0, _
                Optional ByVal pnMontoPoliza As Double, _
                Optional ByVal pnTasaMensualSegInc As Double) As Variant
                ' Agregado Por MAVM 25102010
                'MAVM 20121113: pnIntGraInicial
                'WIOR 20131111 AGREGO pnCuotaBalon
                'Optional ByRef pMatGastos As Variant = "", Optional ByRef pnNumGastos As Integer = 0
                'ALPA 20110524-Se agrego el parametro pnLeasing, para determinar si es un credito leasing
                'ALPA 20140206-Se agregó el parametro pnTipoEntrada
                'Donde:
                    '1  :   Aprobacion
                    '0  :   Otros
                'WIOR 20151223 AGREGO - Optional ByVal pArrMIVIVIENDA As Variant
                'LUCV20180510, Agregó lnTasaSegDes, pMatCalendSegDes, pnMontoPoliza, pnExoSeguroDesgravamen según ERS022-2018
        Dim oParam As COMDCredito.DCOMParametro
        Set oParam = New COMDCredito.DCOMParametro
        Dim nTramoNoConsMonto As Double
        Dim nTramoConsMonto As Double
        Dim nTramoNoConsPorcen As Double
        lnCuotaMivienda = pnCuotaMivienda
        ChkCalMiViv.value = pnMiViv 'ALPA 20141125
        lnMontoMivivienda = pnMontoMivivienda 'ALPA20140206
        lnTasaSegDes = pnTasaSegDes 'LUCV20180601, Agregó según ERS022-2018
        lnExoSeguroDesgravamen = pnExoSeguroDesgravamen 'Agregó según ERS022-2018
        MatCalendSegDes = pMatCalendSegDes 'LUCV20180601, Agregó Según ERS022-2018
        lnMontoPoliza = pnMontoPoliza 'LUCV20180601, Agregó Según ERS022-2018
        lnTasaMensualSegInc = pnTasaMensualSegInc 'LUCV20180601, Agregó Según ERS022-2018
        'WIOR 20160112
        '        'ALPA 20140511****************************************************
        '        If lnMontoMivivienda > oParam.RecuperaValorParametro(2001) * 50 Then
        '                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador2)
        '        Else
        '                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador)
        '        End If
        lnLeasing = pnLeasing
        lsCtaCodLeasing = psCredCodLeasing
        lbLogicoBF = pbLogicoBF   'ALPA20140206
        ldFechaBF = pdFechaBF     'ALPA20140206
        lnMontoMivivienda = pnMontoMivivienda 'ALPA20140206
        txtCuotaInicial.Text = Format(lnMontoMivivienda - pnMonto, "#0.00")  'ALPA20140206
        Me.txtValorInmueble.Text = Format(lnMontoMivivienda, "#0.00") 'ALPA20140206
        txtBonoBuenPagador.Text = Format(nTramoNoConsPorcen, "#0.00") 'ALPA20140206
        bDesemParcial = pbDesemParcial
        MatDesPar = pMatDesPar
        cCtaCodG = psCtaCod
        nSugerAprob = pnSugerAprob
        TxtMonto.Text = Format(pnMonto, "#0.00")
        TxtInteres.Text = Format(pnTasaInt, "#0.0000")
        SpnCuotas.valor = Trim(str(pnNroCuotas))
        SpnPlazo.valor = Trim(str(pnPeriodo))
        TxtPerGra.Text = Trim(str(pnDiasGracia))
        DTFecDesemb.value = Format(pdFecDesemb, "dd/mm/yyyy")
        
        OptTipoCuota(pnTipoCuota - 1).value = True
        OptTipoPeriodo(pnTipoPeriodo - 1).value = True
        If pnTipoPeriodo = FechaFija Then
            TxtDiaFijo.Text = Trim(str(pnDiaFijo))
            'If txtDiafijo.Enabled And txtDiafijo.Visible And FraDatos.Enabled Then
                ChkProxMes.value = IIf(bProxMes, 1, 0)
            'End If
            'Se agrego para manejar la opcion de 2 dias fijos
            TxtDiaFijo2.Text = Trim(str(pnDiaFijo2))
            '*************
        End If
        
        'MAVM 25102010 ***
            txtFechaPago.Text = CDate(dFechaPago)
        '***
        nTipoGracia = pnTipoGracia
        
        'MAVM 28112010 ***
        ChkPerGra.value = IIf(bGracia, 1, 0)
        '***

        TxtTasaGracia.Text = Format(pnTasaGracia, "#0.0000")
        bGraciaGenerada = True
        Set MatGracia = Nothing
        cmdAplicar.Enabled = False
        FraDatos.Enabled = False
        FraFechaPago.Enabled = False 'MAVM 28112010 ***
        'ChkCalMiViv.value = pnMiViv 'Comentado por ALPA 20141125
        ChkCuotaCom.value = pnCuotaCom
        MatGracia = pMatGracia
        
        'Cambios para las opciones de gracia
        If pnTipoGracia = EnCuotas - 1 Then
            optTipoGracia(1).value = True
        End If
        If pnTipoGracia = Capitalizada - 1 Then
            optTipoGracia(0).value = True
        End If
        'If pbIncrementaCapital = True Then chkIncremenK.value = 1
        '***********************************
        
        If bQuincenal = True Then
            ChkQuincenal.value = 1
        End If
        
        'WIOR 20131111 ***********************
        If pnCuotaBalon > 0 Then
            chkCuotaBalon.value = 1
            uspCuotaBalon.valor = pnCuotaBalon
        End If
        'WIOR FIN ****************************
        
        'WIOR 20151223 ***
        txtValorInmueble.Text = ""
        txtCuotaInicial.Text = ""
        txtBonoBuenPagador.Text = ""
        If pnMiViv = 1 Then
            If IsArray(pArrMIVIVIENDA) Then
                If Trim(pArrMIVIVIENDA(0)) <> "" Then
                    txtValorInmueble.Text = Format(CDbl(pArrMIVIVIENDA(0)), "###," & String(15, "#") & "#0.00")
                    txtCuotaInicial.Text = Format(CDbl(pArrMIVIVIENDA(1)), "###," & String(15, "#") & "#0.00")
                    txtBonoBuenPagador.Text = Format(CDbl(pArrMIVIVIENDA(2)), "###," & String(15, "#") & "#0.00")
                End If
            End If
        End If
        'WIOR FIN ********
        
        Call cmdAplicar_Click
        
        'MAVM 20130209 ***
        If pnTipoGracia = 6 Then
            pnIntGraInicial = nIntGraInicial
        End If
        '***
        
        'MAVM 20121113
        cmdResumen.Enabled = False
        FraMivivienda.Enabled = False
        
        'WIOR 20151223 - COMENTO
        '        If lnMontoMivivienda = 0# Then
        '        txtValorInmueble.Text = ""
        '        txtCuotaInicial.Text = ""
        '        End If
        
        cmbSegDes.Visible = False
        'txtBonoBuenPagador.Text = ""
        '***
        
        If bErrorValidacion = True Then
            pbNoMostrarCalendario = True
        End If
        
        If Not pbNoMostrarCalendario Then
            Me.Show 1
        End If
        
        'pMatGastos = MatGastos
        'pnNumGastos = nNumGastos
        
        Inicio = MatCalend
        'MAVM 20121113 ***
        'MatMiViv_2 = MatResulDiff
        MatMiViv_2 = MatResul
        '***
        
        pbErrorValidacion = bErrorValidacion
        'MAVM 20130209 ***
        cCtaCodG = ""
        nSugerAprob = 0
        '***
        pMatCalendSegDes = MatCalendSegDes 'LUCV20180601, Agregó Según ERS022-2018
End Function

Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    
    'Monto de Prestamo
    If Trim(TxtMonto.Text) = "" Then
        MsgBox "Ingrese el Monto del Prestamo", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtMonto.Enabled Then TxtMonto.SetFocus
        Exit Function
    End If
    
    'Interes
    If Trim(TxtInteres.Text) = "" Then
        MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtInteres.Enabled Then TxtInteres.SetFocus
        Exit Function
    Else
        If CDbl(TxtInteres.Text) = 0 Then
            MsgBox "Ingrese la Tasa de Interes", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtInteres.Enabled Then TxtInteres.SetFocus
            Exit Function
        End If
    End If
    
    'Numero de Cuotas
    If Trim(SpnCuotas.valor) = "" Or CInt(SpnCuotas.valor) <= 0 Then
        MsgBox "Ingrese el Numero de Cuotas del Prestamo", vbInformation, "Aviso"
        ValidaDatos = False
        If SpnCuotas.Enabled Then SpnCuotas.SetFocus
        Exit Function
    End If
    
    'Plazo de Cuotas
    If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") And OptTipoPeriodo(0).value = True Then
        MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
        ValidaDatos = False
        If SpnPlazo.Enabled Then SpnPlazo.SetFocus
        Exit Function
    End If
    
    'Fecha de Desembolso
    If ValidaFecha(DTFecDesemb.value) <> "" Then
        MsgBox ValidaFecha(DTFecDesemb.value), vbInformation, "Aviso"
        ValidaDatos = False
        If DTFecDesemb.Enabled Then DTFecDesemb.SetFocus
        Exit Function
    End If
    
    'Valida dia de Fecha Fija
    If OptTipoPeriodo(1).value And (Trim(TxtDiaFijo.Text) = "" Or Trim(TxtDiaFijo.Text) = "0" Or Trim(TxtDiaFijo.Text) = "00") Then
        MsgBox "Ingrese el Dia del Mes que Venceran todas las cuotas", vbInformation, "Aviso"
        ValidaDatos = False
        If TxtDiaFijo.Enabled Then TxtDiaFijo.SetFocus
        Exit Function
    End If
    If CInt(TxtDiaFijo2.Text) > 0 And (CInt(TxtDiaFijo2.Text) <= CInt(TxtDiaFijo.Text)) Then
        MsgBox "El 2do dia tiene que ser mayor al 1ro", vbInformation, "Mensaje"
        ValidaDatos = False
        TxtDiaFijo2.Text = "00"
        TxtDiaFijo2.SetFocus
        Exit Function
    End If
    
    'Valida Generacion de Tipos de Periodo de Gracia
    If ChkPerGra.value = 1 Then
        If (TxtPerGra.Text = "00" Or TxtPerGra.Text = "0") Then
            MsgBox "Ingrese los Dias de Gracia", vbInformation, "Aviso"
            ValidaDatos = False
            If TxtPerGra.Enabled Then TxtPerGra.SetFocus
            Exit Function
        '->***** LUCV20180601, Comentó según ERS022-2018
        Else
            '->*****LUCV20180601, Comentó y agregó según ERS022-2018
'            If (TxtTasaGracia.Text = "0.00" Or TxtTasaGracia.Text = "") Then
'                MsgBox "Ingrese la Tasa de Gracia ", vbInformation, "Aviso"
'                ValidaDatos = False
'                If TxtTasaGracia.Enabled Then TxtTasaGracia.SetFocus
'                    Exit Function
'                Else
'                    If Not bGraciaGenerada And (optTipoGracia(0).value = False And optTipoGracia(1).value = False) Then
'                        ValidaDatos = False
'                        MsgBox "Seleccione un Tipo de Gracia", vbInformation, "Aviso"
'                    If CmdGracia.Enabled Then
'                        CmdGracia.SetFocus
'                    End If
'                    Exit Function
'                End If
'            End If
            If (TxtTasaGracia.Text = "0.00" Or TxtTasaGracia.Text = "") Then
                MsgBox "Ingrese la Tasa de Interés ", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
            '<-***** Fin LUCV20180601
        End If
    End If
    
    'Valida Generacion de Desembolsos
    If bTipoDesembolso = DesembolsoParcial Then
        If Not bDesembParcialGenerado Then
            ValidaDatos = False
            MsgBox "Ingrese los Desembolsos Parciales", vbInformation, "Aviso"
            If CmdDesembParcial.Enabled Then CmdDesembParcial.SetFocus
            Exit Function
        End If
    End If
    
    'ARCV 03-03-2007
    If CInt(TxtPerGra.Text) > 0 Then
        Dim dFechaGracia As Date
        Dim nDiasGraciaPermitido As Integer
        If OptTipoPeriodo(1).value Then 'Fecha Fija
            ''If ChkProxMes.value = 1 Then
            '    dFechaGracia = DateAdd("m", 1, CDate(DTFecDesemb.value)) 'CDate(DTFecDesemb.value) + 30
            ''Else
            ''    dFechaGracia = CDate(DTFecDesemb.value)
            ''End If
            'dFechaGracia = CInt(TxtDiaFijo.Text) & "/" & Month(dFechaGracia) & "/" & Year(dFechaGracia)
            
            'nDiasGraciaPermitido = dFechaGracia - DateAdd("m", 1, CDate(DTFecDesemb.value)) + 1 '+ IIf(ChkProxMes.value = 1, 30, 0)
            'If nDiasGraciaPermitido < 0 Then
            '    ValidaDatos = False
            '    MsgBox "Es necesario utilizar la opcion de proximo mes", vbInformation, "Aviso"
            '    Exit Function
            'End If
            'If nDiasGraciaPermitido <> CInt(txtPerGra.Text) Then
            '    ValidaDatos = False
            '    MsgBox "El numero de dias de gracia permitido es " & nDiasGraciaPermitido, vbInformation, "Aviso"
            '    Exit Function
            'End If
        End If
    End If
    '---------
    
    If OptTipoPeriodo(0).value = True Then
        If CDate(CDate(DTFecDesemb.value) + CDate(SpnPlazo.valor) + CDate(TxtPerGra.Text)) <> CDate(txtFechaPago.Text) Then
            MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
            DTFecDesemb.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    Else
        If Len(Trim(lsCtaCodLeasing)) = 0 Then
            If CDate(TxtPerGra.Text) <> "0" Then
                If CDate(CDate(DTFecDesemb.value) + CDate(30) + CDate(TxtPerGra.Text)) <> CDate(txtFechaPago.Text) Then
                    MsgBox "Falto Presionar Enter en el Campo Fecha de Desembolso", vbInformation, "Aviso"
                    DTFecDesemb.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
    End If
    'ALPA 20140620***********************************************************************
    If ChkCalMiViv.value = 1 Then
        If val(txtValorInmueble.Text) = 0# Or Trim(txtValorInmueble.Text) = "" Then
            MsgBox "Falto ingresar valor del inmueble", vbInformation, "Aviso"
            txtValorInmueble.SetFocus
            ValidaDatos = False
            Exit Function
        End If
'        If Val(txtCuotaInicial.Text) = 0# Or Trim(txtCuotaInicial.Text) = "" Then
'            MsgBox "Falto ingresar valor de la cuota inicial", vbInformation, "Aviso"
'            txtCuotaInicial.SetFocus
'            ValidaDatos = False
'            Exit Function
'        End If
        'If val(txtBonoBuenPagador.Text) = 0# Or Trim(txtBonoBuenPagador.Text) = "" Then
        If val(txtBonoBuenPagador.Text) < 0# Or Trim(txtBonoBuenPagador.Text) = "" Then
            MsgBox "Falto ingresar valor del bono de buen pagador", vbInformation, "Aviso"
            'txtBonoBuenPagador.SetFocus
            EnfocaControl txtBonoBuenPagador
            ValidaDatos = False
            Exit Function
        End If
    End If
    '************************************************************************************
    
    'WIOR 20150210 **********************************
    If fbInicioSim Then
        If fraGastoCom.Visible Then
            If Trim(cmbSeguroDes.Text) = "" Then
                MsgBox "Ingrese el tipo de Seguro Desgravamen", vbInformation, "Aviso"
                cmbSeguroDes.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If chkEnvioEst.value = 1 Then
                If Trim(cmbEnvioEst.Text) = "" Then
                    MsgBox "Ingrese el Tipo de Envío del estado de cuenta", vbInformation, "Aviso"
                    cmbEnvioEst.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
    End If
    'WIOR FIN ***************************************
End Function

Private Sub HabilitaFechaFija(ByVal pbHabilita As Boolean)
    'DTFecDesemb.Enabled = Not pbHabilita
    SpnPlazo.Enabled = Not pbHabilita
    SpnPlazo.valor = IIf(pbHabilita, "0", SpnPlazo.valor)
    LblDia.Enabled = pbHabilita
    TxtDiaFijo.Enabled = pbHabilita
    ChkProxMes.Enabled = pbHabilita
    ChkProxMes.value = 0
    TxtDiaFijo2.Enabled = pbHabilita
    lblDia2.Enabled = pbHabilita
    If pbHabilita Then
        TxtDiaFijo.Text = "00"
        If TxtDiaFijo.Enabled And TxtDiaFijo.Visible And FraDatos.Enabled Then
            TxtDiaFijo.SetFocus
        End If
    End If
End Sub

Private Sub CargaControles()
    FECalend.RowHeight(0) = 250
    FECalend.RowHeight(1) = 250
    TxtDiaFijo.Text = "00"
    TxtTasaGracia.Text = "00"
End Sub

Private Sub ChkCalMiViv_Click()
    ReDim MatGracia(0)
    If ChkCalMiViv.value = 1 Then
        OptTipoCuota(0).value = True
        OptTipoCuota(1).Enabled = False
        OptTipoCuota(2).Enabled = False
        OptTipoPeriodo(1).value = True
        OptTipoPeriodo(0).Enabled = False
        ChkCuotaCom.Enabled = False
        'Frame5.Visible = False 'WIOR 20151223 - COMENTO
        Frame5.Visible = True 'WIOR 20151223
        SSCalend.Visible = True
        optTipoGracia(0).value = False
        optTipoGracia(1).value = False
        'MAVM 20121113 ***
        cmbSegDes.Visible = True
        cmbSegDes.ListIndex = 0
        FraMivivienda.Enabled = True
        'txtValorInmueble.SetFocus
        cmdResumen.Visible = True
        cmdResumen.Enabled = False 'MAVM 20121115
        '***
        TxtMonto.Text = 0# 'ALPA20140620**********************
        
        'WIOR 20150210 ***********************
        If fbInicioSim Then
            fraGastoCom.Visible = False
        End If
        'WIOR FIN ****************************
    Else
        OptTipoCuota(0).value = True
        OptTipoCuota(1).Enabled = True
        OptTipoCuota(2).Enabled = True
        OptTipoPeriodo(0).value = True
        OptTipoPeriodo(0).Enabled = True
        ChkCuotaCom.Enabled = True
        Frame5.Visible = True
        SSCalend.Visible = False
        optTipoGracia(0).value = True
        optTipoGracia(1).value = True
        'MAVM 20121113 ***
        FraMivivienda.Enabled = False
        cmdResumen.Visible = False
        txtValorInmueble.Text = ""
        txtCuotaInicial.Text = ""
        cmbSegDes.Visible = False
        txtBonoBuenPagador.Text = ""
        TxtMonto.Text = ""
        TxtInteres.Text = ""
        '***
        TxtMonto.Text = 0# 'ALPA20140620**********************
          'WIOR 20150210 ***********************
        If fbInicioSim Then
            fraGastoCom.Visible = True
        End If
        'WIOR FIN ****************************
    End If
End Sub

'WIOR 20131109 ********************************
Private Sub chkCuotaBalon_Click()
If chkCuotaBalon.value = 1 Then
    uspCuotaBalon.Enabled = True
    uspCuotaBalon.valor = 1
Else
    uspCuotaBalon.Enabled = False
    uspCuotaBalon.valor = 0
End If
End Sub
'WIOR FIN *************************************

'WIOR 20150210 ***************************
Private Sub chkEnvioEst_Click()
    If chkEnvioEst.value = 1 Then
        cmbEnvioEst.Enabled = True
    Else
        cmbEnvioEst.Enabled = False
    End If
End Sub
'WIOR FIN ********************************

'Private Sub chkIncremenK_Click()
    'txtMonto.Text = ""
'End Sub

'Private Sub chkCapitalizar_Click()
'If chkCapitalizar.value = 1 Then
'    CmdGracia.Enabled = False
'    chkIncremenK.Visible = True
'Else
'    CmdGracia.Enabled = True
'    chkIncremenK.Visible = False
'End If
'End Sub

Private Sub ChkPerGra_Click()
Dim i As Integer

    ReDim MatGracia(CInt(SpnCuotas.valor))

'    Comentado Por MAVM 19102010
'    ChkProxMes.Enabled = True

    For i = 0 To CInt(SpnCuotas.valor) - 1
        MatGracia(i) = "0.00"
    Next i
    Call LimpiaFlex(FECalend)
    If ChkPerGra.value = 1 Then
        LblTasaGracia.Enabled = True
        'TxtTasaGracia.Enabled = True 'LUCV20180601, Comentó, segun ERS022-2018
        LblPorcGracia.Enabled = True
        
'        Comentado Por MAVM 05102010 ***
'        txtPerGra.Enabled = True
'        txtPerGra.Text = "0"
'        ***

        'TxtTasaGracia.Text = "0.00" 'LUCV20180601, Comentó según ERS022-2018
        TxtTasaGracia.Text = Format(TxtInteres.Text, "#0.00") 'LUCV20180601, Agregó según ERS022-2018
        CmdGracia.Enabled = True
        optTipoGracia(0).value = False
        optTipoGracia(1).value = False
        
        'Para Fecha Fija no Aplica
        If OptTipoPeriodo(1).value = True Then
            optTipoGracia(0).Enabled = True
            optTipoGracia(1).Enabled = True 'False '10-05-2006
        Else
            optTipoGracia(0).Enabled = True
            optTipoGracia(1).Enabled = True
        End If
        'Para los Calendarios Mi Vivienda y de Trabajadores y Directores no aplica
        If ChkCalMiViv.value = 1 Or ChkQuincenal.value = 1 Then
            optTipoGracia(0).value = False
            optTipoGracia(1).value = False
            optTipoGracia(0).Enabled = False
            optTipoGracia(1).Enabled = False
        Else
        '    optTipoGracia(0).Enabled = True
        '    optTipoGracia(1).Enabled = True
        End If
        
'        Comentado y Agregado Por MAVM 05102010 ***
'        txtPerGra.SetFocus
'        txtFechaPago.SetFocus 'Descomentar
'        ***
        
        If Me.ChkCalMiViv.value = 1 Then
            TxtTasaGracia.Text = TxtInteres.Text
            TxtTasaGracia.Enabled = False
            'cmdgracia.Enabled = False
            bGraciaGenerada = True
        End If
    Else
        LblTasaGracia.Enabled = False
        TxtTasaGracia.Enabled = False
        LblPorcGracia.Enabled = False
        TxtPerGra.Enabled = False
        TxtPerGra.Text = "0"
        TxtTasaGracia.Text = "0.00"
        CmdGracia.Enabled = False
        optTipoGracia(0).Enabled = False
        optTipoGracia(1).Enabled = False
        optTipoGracia(0).value = False
        optTipoGracia(1).value = False
        
        'MAVM 30092010 ***
        GenerarFechaPago
        If OptTipoPeriodo(1).value = True Then
            ChkPerGra.Enabled = False
        End If
        '***
        Call txtFechaPago_KeyPress(13) 'JUEZ 20150307
    End If
End Sub

Private Sub ChkProxMes_Click()
     If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
End Sub

Private Sub ChkQuincenal_Click()
    If ChkQuincenal.value = 1 Then
        FraTipoCuota.Enabled = False
        FraTipoPeriodo.Enabled = False
        FraComodin.Enabled = False
        
    Else
        FraTipoCuota.Enabled = True
        FraTipoPeriodo.Enabled = True
        FraComodin.Enabled = True
        
    End If
End Sub

'->*****LUCV20180601, Comentó y Agregó según ERS022-2018
'Private Sub cmdAplicar_Click()
'Dim i As Integer
'Dim nTipoCuota As Integer
'Dim nTipoPeriodo As Integer
'Dim nTotalInteres As Double
'Dim nTotalCapital As Double
'Dim oParam As COMDCredito.DCOMParametro
'Dim nTramoNoConsMonto As Double
'Dim nTramoConsMonto As Double
'Dim nTramoNoConsPorcen As Double
''Dim nTramoConsTasa As Double
'Dim nPlazoMiViv As Integer
'Dim nPlazoMiVivMax As Integer 'MAVM 20121113
'Dim nRedondeoITF As Double 'MAVM 20121113
'Dim nTotalcuotasCONItF As Double
'Dim nTotalcuotasLeasing As Double
'Dim lnSalCapital As Double
''Dim lrsCalendLeasing As ADODB.Recordset
''Set lrsCalendLeasing = New ADODB.Recordset
'
'    nTotalcuotasLeasing = 0
'    nIntGraInicial = 0   ' MAVM 20130209
'    nMontoCapInicial = 0 ' MAVM 20130509
'
'    Call LimpiaFlex(FECalend)
'    Call LimpiaFlex(FECalBPag)
'    Call LimpiaFlex(FECalMPag)
'    MatResul = Array(0)
'    MatResulDiff = Array(0)
'    MatCalend = Array(0)
'
'    If Not ValidaDatos Then
'        bErrorValidacion = True
'        Exit Sub
'    Else
'        bErrorValidacion = False
'    End If
'
'    If ChkCalMiViv.value = 1 Then   'Esos Parametros se deben cargar solo para la Opcion MiVivienda
'        'txtMonto.Text = val(txtValorInmueble.Text) - val(txtCuotaInicial)  'RECO20140813 152-2013'WIOR 20160111 - COMENTO
'        'Porcentaje Real sin dividir enter 100
'        Set oParam = New COMDCredito.DCOMParametro
'        'nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivTramo)
'        'nPlazoMiViv = oParam.RecuperaValorParametro(3056)
'        'MAVM 20121113 ***
'        'Call oParam.RecuperaParametrosCalendario(nTramoNoConsPorcen, nPlazoMiViv)
'        Call oParam.RecuperaParametrosCalendario(nTramoNoConsPorcen, nPlazoMiViv, nPlazoMiVivMax)
'        '***
'        Set oParam = Nothing
'
'        If OptTipoPeriodo(0).value Then 'Periodo Fijo
'           If (CInt(SpnCuotas.valor) * CInt(SpnPlazo.valor)) / 360 <= nPlazoMiViv Then
'                MsgBox "El Plazo del Credito debe ser Mayor a " & nPlazoMiViv & " Años", vbInformation, "Aviso"
'                bErrorValidacion = True
'                Exit Sub
'           End If
'        Else 'Fecha Fija
''            If (CInt(SpnCuotas.valor) * 30) / 360 <= nPlazoMiViv Then
''                MsgBox "El Plazo del Credito debe ser Mayor a " & nPlazoMiViv & " Años", vbInformation, "Aviso"
''                bErrorValidacion = True
''                Exit Sub
''           End If
'            'MAVM 20121113 ***
'            If (CInt(SpnCuotas.valor) * 30) / 360 < nPlazoMiViv Then
'                MsgBox "El Plazo del Credito debe ser Minimo " & nPlazoMiViv & " Años", vbInformation, "Aviso"
'                bErrorValidacion = True
'                Exit Sub
'            End If
'
'            If Format((CInt(SpnCuotas.valor) * 30) / 360, "####.00") > Format(nPlazoMiVivMax, "####.00") Then
'                MsgBox "El Plazo del Credito debe ser Maximo " & nPlazoMiVivMax & " Años", vbInformation, "Aviso"
'                bErrorValidacion = True
'                Exit Sub
'            End If
'            '***
'        End If
'    End If
'
'    Call LimpiaFlex(FECalend)
'    Call LimpiaFlex(FECalBPag)
'    Call LimpiaFlex(FECalMPag)
'    For i = 0 To 2
'        If OptTipoCuota(i).value Then
'            nTipoCuota = i + 1
'            Exit For
'        End If
'    Next i
'    For i = 0 To 1
'        If OptTipoPeriodo(i).value Then
'            nTipoPeriodo = i + 1
'            Exit For
'        End If
'    Next i
'
'    'Set oCalendario = New COMNCredito.NCOMCalendario
'    'WIOR 20151223 - COMENTO
'    '    If Me.ChkCalMiViv.value = 1 Then
'    '        'MAVM 20121113 ***
'    '
'    '        'Dim oParam As COMDCredito.DCOMParametro
'    '        Set oParam = New COMDCredito.DCOMParametro
'    '        'ALPA 20140511****************************************************
'    '        If lnMontoMivivienda > oParam.RecuperaValorParametro(2001) * 50 Then
'    '                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador2)
'    '        Else
'    '                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador)
'    '        End If
'    '        'nTramoNoConsMonto = Format((nTramoNoConsPorcen / 100) * CDbl(Me.TxtMonto.Text), "#0.00")
'    '        'nTramoConsMonto = Format(CDbl(Me.TxtMonto.Text) - nTramoNoConsMonto, "#0.00")
'    '        nTramoNoConsMonto = Format(CDbl(Me.TxtMonto.Text) - nTramoNoConsPorcen, "#0.00")
'    '        nTramoConsMonto = Format(nTramoNoConsPorcen, "#0.00")
'    '        '***
'    '
'    '        Call GeneraDobleCalendario(MatCalend, MatCalend_2, nTramoConsMonto, nTramoNoConsMonto, CDbl(Txtinteres.Text), CInt(spnCuotas.valor), CInt(SpnPlazo.valor), _
'    '                                        CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, nTipoGracia, CInt(txtPerGra.Text), _
'    '                                         CInt(TxtDiaFijo.Text), IIf(ChkProxMes.value = 1, True, False), MatGracia, IIf(Me.ChkCalMiViv.value = 1, True, False), _
'    '                                         optTipoGracia(0).value, CDbl(TxtTasaGracia.Text))
'    '        'Set oCalendario = Nothing
'    '
'    '        'MAVM 20121113 ***
'    '        'MatResul = UnirMatricesMiVivienda(MatCalend, MatCalend_2, CDbl(Me.TxtMonto.Text))
'    '        Call ObtenerDesgravamenHipot(CDbl(IIf(txtValorInmueble.Text = "", 0, txtValorInmueble.Text)), nTramoNoConsMonto + nTramoNoConsPorcen, lnMontoMivivienda, CDbl(lnCuotaMivienda))
'    '        MatResul = UnirMatricesMiVivienda(MatCalend, MatCalend_2, CDbl(Me.TxtMonto.Text), IIf(ChkCalMiViv.value, Left(cmbSegDes.Text, 1), ""), nTramoNoConsPorcen)
'    '        '***
'    '        MatResulDiff = DiferencialMatricesMiVivienda(MatCalend, MatResul)
'    '
'    '        'Cargar Matrices a FlexGrid
'    '        nTotalInteres = 0
'    '        nTotalCapital = 0
'    '        nTotalcuotasCONItF = 0
'    '        For i = 0 To UBound(MatCalend) - 1
'    '            FECalBPag.AdicionaFila
'    '            FECalBPag.TextMatrix(i + 1, 1) = Trim(MatCalend(i, 0))
'    '            FECalBPag.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))
'    '            'MAVM 20121113 ***
'    '            If i = 0 Then
'    '                nCuotMensBono = Trim(MatCalend(i, 2))
'    '            End If
'    '            '***
'    '            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 2)) + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "#0.00")
'    '            FECalBPag.TextMatrix(i + 1, 3) = MatCalend(i, 2) 'Format(CDbl(MatCalend(i, 2)) + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "#0.00")
'    ''            If i < 6 Then
'    ''                FECalBPag.TextMatrix(i + 1, 3) = Trim(MatResul(i, 2)) + Trim(MatResul(i, 6)) - Trim(MatCalend(i, 6))
'    ''            Else
'    ''                FECalBPag.TextMatrix(i + 1, 3) = Trim(MatResul(i, 2)) - Trim(MatResul(i, 3)) + Trim(MatCalend(i, 3)) + Trim(MatCalend(i, 4)) - Trim(MatResul(i, 4)) - Trim(MatResul(i, 6)) + Trim(MatCalend(i, 6))
'    ''            End If
'    '            FECalBPag.row = i + 1
'    '            FECalBPag.col = 3
'    '            FECalBPag.CellForeColor = vbBlue
'    '            FECalBPag.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 3))
'    '            FECalBPag.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4))
'    '            FECalBPag.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 5))
'    '            'MAVM 20121113 ***
'    '            'FECalBPag.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 6))
'    '            'FECalBPag.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 7))
'    '
'    '            'FECalBPag.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00")
'    '
'    '            FECalBPag.TextMatrix(i + 1, 7) = "0.00" 'Trim(MatCalend(i, 8))
'    '            MatCalend(i, 6) = MatResul(i, 6)
'    '            FECalBPag.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 6))
'    '            FECalBPag.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 9))
'    '            FECalBPag.TextMatrix(i + 1, 10) = Trim(MatCalend(i, 7))
'    '
'    '            nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))))
'    '            If nRedondeoITF > 0 Then
'    '                FECalBPag.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) - nRedondeoITF + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "0.00")
'    '            Else
'    '                FECalBPag.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) + CDbl(Trim(MatResul(i, 6))) - CDbl(Trim(MatCalend(i, 6))), "0.00")
'    '            End If
'    '            '***
'    '
'    '            nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
'    '            'MAVM 20121113 ***
'    '            'nTotalcuotasCONItF = nTotalCapital + FECalBPag.TextMatrix(i + 1, 9)
'    '            nTotalcuotasCONItF = nTotalcuotasCONItF + FECalBPag.TextMatrix(i + 1, 11)
'    '            '***
'    '
'    '            'nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(I, 4)))
'    '            '11-05-2006
'    '            nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5))) '+ CDbl(Trim(MatCalend(i, 6)))
'    '        Next i
'    '        lblCapital.Caption = Format(nTotalCapital, "#0.00")
'    '        lblInteres.Caption = Format(nTotalInteres, "#0.00")
'    '        lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
'    '        lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
'    '
'    '        'Carga Mal Pagador
'    '        For i = 0 To UBound(MatCalend) - 1
'    '            FECalMPag.AdicionaFila
'    '            FECalMPag.TextMatrix(i + 1, 1) = Trim(MatResul(i, 0))
'    '            FECalMPag.TextMatrix(i + 1, 2) = Trim(MatResul(i, 1))
'    '            'MAVM 20121113 ***
'    '            If i = 6 Then
'    '                nCuotMens = Trim(MatResul(i, 2))
'    '            End If
'    '            '***
'    '            FECalMPag.TextMatrix(i + 1, 3) = Trim(MatResul(i, 2))
'    '            FECalMPag.row = i + 1
'    '            FECalMPag.col = 3
'    '            FECalMPag.CellForeColor = vbBlue
'    '            FECalMPag.TextMatrix(i + 1, 4) = Trim(MatResul(i, 3))
'    '            FECalMPag.TextMatrix(i + 1, 5) = Trim(MatResul(i, 4))
'    '            FECalMPag.TextMatrix(i + 1, 6) = Trim(MatResul(i, 5))
'    '
'    '            'MAVM 20121113 ***
'    '            'FECalMPag.TextMatrix(i + 1, 7) = Trim(MatResul(i, 6))
'    '            'FECalMPag.TextMatrix(i + 1, 8) = Trim(MatResul(i, 7))
'    '
'    '            'FECalMPag.TextMatrix(i + 1, 9) = Format(CDbl(MatResul(i, 2)) + fgITFCalculaImpuesto(CDbl(MatResul(i, 2))), "0.00")
'    '            FECalMPag.TextMatrix(i + 1, 7) = "0.00" 'Trim(MatResul(i, 8))
'    '            FECalMPag.TextMatrix(i + 1, 8) = Trim(MatResul(i, 6))
'    '            FECalMPag.TextMatrix(i + 1, 9) = Trim(MatResul(i, 9))
'    '            FECalMPag.TextMatrix(i + 1, 10) = Trim(MatResul(i, 7))
'    '
'    '            nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(MatResul(i, 2))))
'    '            If nRedondeoITF > 0 Then
'    '                FECalMPag.TextMatrix(i + 1, 11) = Format(CDbl(MatResul(i, 2)) + fgITFCalculaImpuesto(CDbl(MatResul(i, 2))) - nRedondeoITF, "0.00")
'    '            Else
'    '                FECalMPag.TextMatrix(i + 1, 11) = Format(CDbl(MatResul(i, 2)) + fgITFCalculaImpuesto(CDbl(MatResul(i, 2))), "0.00")
'    '            End If
'    '            '***
'    '        Next i
'    '        FECalBPag.row = 1
'    '        FECalBPag.TopRow = 1
'    '        FECalMPag.TopRow = 1
'    '        SSCalend.Tab = 0
'    '
'    '    Else
'    'WIOR 20151223 FIN
'    If ChkQuincenal.value = 1 Then  'Esta opcion esta deshabilitada
'            MatCalend = GeneraCalendario(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
'                         CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, _
'                        nTipoGracia, CInt(TxtPerGra.Text), CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, True, IIf(ChkCuotaCom.value = 1, True, False), bDesemParcial, MatDesPar, , , True, _
'                         optTipoGracia(0).value, CDbl(TxtTasaGracia.Text))
'
'            'Set oCalendario = Nothing
'            'Carga Flex Edit
'        nTotalInteres = 0
'        nTotalCapital = 0
'        nTotalcuotasCONItF = 0
'
'        For i = 0 To UBound(MatCalend) - 1
'            FECalend.AdicionaFila
'            FECalend.TextMatrix(i + 1, 1) = Trim(MatCalend(i, 0))
'            FECalend.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))
'            FECalend.TextMatrix(i + 1, 3) = Trim(MatCalend(i, 2))
'            FECalend.row = i + 1
'            FECalend.Col = 3
'            FECalend.CellForeColor = vbBlue
'            FECalend.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 3))
'            FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4))
'            FECalend.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 5))
'            FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 6))
'
'            FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 7))
'
'            nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
'            'nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(I, 4)))
'            '11-05-2006
'            nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5))) '+ CDbl(Trim(MatCalend(i, 6)))
'
'            'FECalend.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00")'RECO
'            FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") 'RECO
'
'            'nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 9)) 'RECO
'            nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 10)) 'RECO
'
'        Next i
'        lblCapital.Caption = Format(nTotalCapital, "#0.00")
'        lblInteres.Caption = Format(nTotalInteres, "#0.00")
'        lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
'        lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
'        FECalend.row = 1
'        FECalend.TopRow = 1
'    Else
'        'Se Agrego para manejar la Capitalizacion de la Gracia
'        If optTipoGracia(0).value Then 'LUCV20180601, Se deshabilitó según ERS022-2018
'            'MAVM 20130209
'            Dim oCredito As COMNCredito.NCOMCredito
'            Set oCredito = New COMNCredito.NCOMCredito
'            '***
'            'Para realizar los cálculos
'            nMontoCapInicial = CDbl(TxtMonto.Text)
'            'MAVM 20130209 ***
'            'If chkIncremenK.value = 1 Then
'            nIntGraInicial = oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), CInt(TxtPerGra.Text), CDbl(TxtMonto.Text))
'            '***
'            nTipoGracia = gColocTiposGraciaCapitalizada
'            'End If
'            Set oCredito = Nothing
'            'Los calculos se hacen con el Nuevo Monto
'            'If chkIncremenK.value = 1 Then
'            'nMontoCapInicial = 0
'        End If
'        '*********************************************lsCtaCodLeasing
'        If Len(Trim(lsCtaCodLeasing)) = 0 Then
'            'MAVM 20130305: nInteresAFecha, nCapitalInicial
'
'            '->***** LUCV20180601, Comentó y agregó según ERS022-2018
''               MatCalend = GeneraCalendario(CDbl(txtMonto.Text), CDbl(Txtinteres.Text), _
''                                            CInt(spnCuotas.valor), CInt(SpnPlazo.valor), _
''                                            CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, _
''                                            nTipoPeriodo, nTipoGracia, CInt(txtPerGra.Text), _
''                                            CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, _
''                                            True, IIf(ChkCuotaCom.value = 1, True, False), bDesemParcial, _
''                                            MatDesPar, , , , _
''                                            optTipoGracia(1).value, CDbl(TxtTasaGracia.Text), _
''                                            CInt(TxtDiaFijo2.Text), nMontoCapInicial, _
''                                            IIf(chkPagoInteres.value = 1, True, False), bRenovarCredito, _
''                                            nInteresAFecha, nIntGraInicial, _
''                                            IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0))
'                                            'WIOR 20131111 AGREGO IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0)
'
'            MatCalend = GeneraCalendarioNuevo(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), _
'                                            CInt(SpnCuotas.valor), CInt(SpnPlazo.valor), _
'                                            CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, _
'                                            nTipoPeriodo, nTipoGracia, CInt(TxtPerGra.Text), _
'                                            CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, _
'                                            True, IIf(ChkCuotaCom.value = 1, True, False), bDesemParcial, _
'                                            MatDesPar, , , , _
'                                            optTipoGracia(1).value, CDbl(TxtTasaGracia.Text), _
'                                            CInt(TxtDiaFijo2.Text), nMontoCapInicial, _
'                                            IIf(chkPagoInteres.value = 1, True, False), bRenovarCredito, _
'                                            nInteresAFecha, nIntGraInicial, _
'                                            IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0), _
'                                            lnTasaSegDes, lnExoneraSegDes, cCtaCodG, MatCalendSegDes)
'            '<-***** Fin LUCV20180601
'
'        Else
'            MatCalend = GeneraCalendarioLeasing(CDbl(TxtMonto.Text), -CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
'                        CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, nTipoPeriodo, _
'                        nTipoGracia, CInt(TxtPerGra.Text), CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, True, IIf(ChkCuotaCom.value = 1, True, False), bDesemParcial, MatDesPar, , , , _
'                        optTipoGracia(1).value, CDbl(TxtTasaGracia.Text), CInt(TxtDiaFijo2.Text), nMontoCapInicial, IIf(chkPagoInteres.value = 1, True, False), bRenovarCredito, nInteresAFecha, lsCtaCodLeasing)
'        End If
'        'Set oCalendario = Nothing
'        'Carga Flex Edit
'
'
'        '**DAOR 20070410, Obtener Gastos en Reprogramación**********************
'        If bRenovarCredito Then
'            Call ObtenerGastosEnReprogramacion
'        End If
''        '***********************************************************************
''
''        '**PEAC 20080815, Obtener Desgravamen **********************
'        If cmdAplicar.Enabled Then
'            Call ObtenerDesgravamen
'        Else
'            For i = 0 To UBound(MatCalend) - 1
'                'Cuota          = Capital + Interes Compensatorio + Interes Gracia + Gasto + Seguro Desgramen
'                MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)), "#0.00")
'            Next i
'        End If
'        '***********************************************************************
'
'        nTotalInteres = 0
'        nTotalCapital = 0
'        nTotalcuotasCONItF = 0
'        lnSalCapital = val(TxtMonto.Text)
'
'        For i = 0 To UBound(MatCalend) - 1
'            FECalend.AdicionaFila
'            FECalend.TextMatrix(i + 1, 1) = Trim(MatCalend(i, 0))
'            If Len(Trim(lsCtaCodLeasing)) = 18 Then
'                txtFechaPago.Text = CDate(Trim(MatCalend(0, 0)))
'            End If
'            'MAVM 25102010 ***
'            If i = 0 Then
'                If CDate(Trim(MatCalend(i, 0))) <> CDate(txtFechaPago.Text) And nTipoGracia <> 1 Then
'                    If nTipoGracia <> 6 Then 'MAVM 20130312
'                        Call LimpiaFlex(FECalend)
'                        MsgBox "Falto Presionar Enter en el Campo Fecha de Pago", vbInformation, "Aviso"
'                        bErrorValidacion = True
'                        Exit Sub
'                    End If 'MAVM 20130312
'                End If
'            End If
'            '***
'
'            FECalend.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))
'            'Modify Gitu 22-08-08
'            'Descomentar cuando esten seguros de los cambios GITU
''            If Val(MatCalend(I, 2)) <> (Val(Trim(MatCalend(I, 3))) + Val(Trim(MatCalend(I, 4))) + Val(Trim(MatCalend(I, 5))) + Val(Trim(MatCalend(I, 6)))) Then
''                MatCalend(I, 2) = Format(Val(Trim(MatCalend(I, 3))) + Val(Trim(MatCalend(I, 4))) + Val(Trim(MatCalend(I, 5))) + Val(Trim(MatCalend(I, 6))), "##0.00")
''            Else
''                MatCalend(I, 2) = MatCalend(I, 2)
'            'End If
'
'            'MAVM 20130312
'            'FECalend.TextMatrix(i + 1, 3) = Trim(MatCalend(i, 2))
'            If nTipoGracia = 6 Then
'                MatCalend(i, 2) = Trim(CDbl(MatCalend(i, 2)) + CDbl(MatCalend(i, 11)))
'                FECalend.TextMatrix(i + 1, 3) = Format(Trim(MatCalend(i, 2)), "#0.00")
'            Else
'                FECalend.TextMatrix(i + 1, 3) = Trim(MatCalend(i, 2))
'            End If
'            '***
'
'            FECalend.row = i + 1
'            FECalend.Col = 3
'            FECalend.CellForeColor = vbBlue
'            FECalend.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 3)) 'Amort Cap
'
'            'MAVM 20130209 *** 'Interes Comp
'            'FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4))
'            If nTipoGracia = 6 Then
'                MatCalend(i, 4) = Trim(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 11)))
'                FECalend.TextMatrix(i + 1, 5) = Format(Trim(MatCalend(i, 4)), "#0.00")
'            Else
'                FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4))
'            End If
'
'            'Interes Gracia
'            FECalend.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 5))
'
'            If Len(Trim(lsCtaCodLeasing)) = 0 Then
'                FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 8))
'                'MAVM 20102003 ok
'                FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 6))
'            Else
'                FECalend.TextMatrix(i + 1, 7) = Trim(MatCalend(i, 6))
'                FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 8))
'            End If
'
'            'FECalend.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 7)) 'RECO
'            FECalend.TextMatrix(i + 1, 10) = Trim(MatCalend(i, 7)) 'RECO
'            'Descomentar cuando esten seguros de los cambios GITU
''            FECalend.TextMatrix(I + 1, 8) = lnSalCapital - Trim(MatCalend(I, 3))
''            lnSalCapital = lnSalCapital - Trim(MatCalend(I, 3))
'            FECalend.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 13)) 'RECO20150512
'            'MAVM 20130209 ***
'            'nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(I, 4)))
'            'nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5))) '+ CDbl(Trim(MatCalend(i, 6)))
'            If Not (i = 0 And nTipoGracia = gColocTiposGraciaCapitalizada) Then
'                nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
'                nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5)))
'            End If
'            '***
'
'            'MAVM 20100320
'            'FECalend.TextMatrix(i + 1, 9) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00")
'
'            'MAVM 20121113 ***
'            'FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") 'RECO
'            FECalend.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") 'RECO
'            nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))))
'            If nRedondeoITF > 0 Then
'                'FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) - nRedondeoITF, "0.00") 'RECO
'                FECalend.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) - nRedondeoITF, "0.00") 'RECO
'            Else
'                'FECalend.TextMatrix(i + 1, 10) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") 'RECO
'                FECalend.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") 'RECO
'            End If
'            '***
'
'            'Descomentar cuando esten seguros de los cambios GITU
''            If I > 0 Then
''                If InStr(FECalend.TextMatrix(I + 1, 9), ".") <> 0 Then
''                    FECalend.TextMatrix(1, 9) = Format(Val(FECalend.TextMatrix(1, 9)) + Round(Val(FECalend.TextMatrix(I, 9)) - Val(Left(FECalend.TextMatrix(I + 1, 9), InStr(FECalend.TextMatrix(I + 1, 9), ".") - 1)), 2))
''                    FECalend.TextMatrix(I + 1, 9) = Format(Val(Left(FECalend.TextMatrix(I + 1, 9), InStr(FECalend.TextMatrix(I + 1, 9), ".") - 1)))
''                End If
''            End If
'            'End Gitu
'            'nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(I + 1, 9))
'
'            'MAVM 20130209 ***
'            If Not (nTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
'                'nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 10)) 'RECO
'                nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 11)) 'RECO
'                'nTotalcuotasLeasing = nTotalcuotasLeasing + CDbl(FECalend.TextMatrix(i + 1, 10))'RECO
'                nTotalcuotasLeasing = nTotalcuotasLeasing + CDbl(FECalend.TextMatrix(i + 1, 11)) 'RECO
'            End If
'            '***
'        Next i
'
'        'MAVM 20130209 ***
'        Set oCredito = Nothing
'        If nTipoGracia = gColocTiposGraciaCapitalizada Then
'            FECalend.TextMatrix(1, 3) = ""
'            FECalend.TextMatrix(1, 5) = ""
'            FECalend.TextMatrix(1, 7) = ""
'            FECalend.TextMatrix(1, 8) = ""
'            FECalend.TextMatrix(1, 10) = ""
'        End If
'        '***
'
'        lblCapital.Caption = Format(nTotalCapital, "#0.00")
'        lblInteres.Caption = Format(nTotalInteres, "#0.00")
'        lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
'        lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
'        FECalend.row = 1
'        FECalend.TopRow = 1
'    End If
'
'    '*** PEAC 20080819, Desarrollo de la Tasa Costo Efectivo Anual **********************
''    lblTEA.Visible = False
''    lblTasaEfectivaAnual.Visible = False
''    lblTCEA.Visible = True
''    lblTasaCostoEfectivoAnual.Visible = True
'
'    fraTasaAnuales.Visible = True
'    Set oCredito = New COMNCredito.NCOMCredito
'        nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(CDbl(TxtInteres), 360) * 100, 2)
'        'MAVM 20121113 ***
'        'nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), CDbl(TxtMonto.Text), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing)
'
'        'MAVM 20130305 ***
'        If nTipoGracia = 6 Then
'            Dim Y As Integer
'            Dim MatCalendTemp() As String
'            ReDim MatCalendTemp(UBound(MatCalend) - 1, 14)
'            For i = 0 To UBound(MatCalend) - 1
'                For Y = 0 To 14
'                    MatCalendTemp(i, Y) = MatCalend(i + 1, Y)
'                Next Y
'            Next i
'            Erase MatCalend
'            ReDim MatCalend(UBound(MatCalendTemp), 14)
'
'            For i = 0 To UBound(MatCalendTemp)
'                For Y = 0 To 14
'                    MatCalend(i, Y) = MatCalendTemp(i, Y)
'                Next Y
'            Next i
'            Erase MatCalendTemp
'        End If
'        '***
'
'        'nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), IIf(ChkCalMiViv.value = 0, CDbl(TxtMonto.Text), nTramoNoConsMonto), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing) 'MAVM 20121113
'        'nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), IIf(ChkCalMiViv.value = 0, CDbl(TxtMonto.Text), nTramoNoConsMonto), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing, nTipoPeriodo) 'JUEZ 20140814 'WIOR 20151223 - COMENTO
'        nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), CDbl(TxtMonto.Text), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing, nTipoPeriodo) 'WIOR 20151223
'        '***
'        lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
'        lblTasaCostoEfectivoAnual.Caption = nTasaCostoEfectivoAnual & " %"
'    Set oCredito = Nothing
'    '**********************************************************************************
'    'MAVM 20121115 ***
'    If Me.ChkCalMiViv.value = 1 Then
'        cmdResumen.Enabled = True
'    End If
'    'END MAVM ********
'    If UBound(MatCalend) = 0 Then
'        cmdImprimir.Enabled = False
'    Else
'        cmdImprimir.Enabled = True
'    End If
'
'End Sub
'<-***** Fin LUCV20180601

Private Sub cmdAplicar_Click() '[Se comentó el evento [cmdAplicar] según ERS022-2018 (firma:LUCV20180601)]
Dim i As Integer
Dim nTipoCuota As Integer
Dim nTipoPeriodo As Integer
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim nRedondeoITF As Double
Dim nTotalcuotasCONItF As Double
Dim lnSalCapital As Double

    nIntGraInicial = 0
    nMontoCapInicial = 0
    
    Call LimpiaFlex(FECalend)
    Call LimpiaFlex(FECalBPag)
    Call LimpiaFlex(FECalMPag)
    MatCalend = Array(0)
       
    If Not ValidaDatos Then
        bErrorValidacion = True
        Exit Sub
    Else
        bErrorValidacion = False
    End If
    
    Call LimpiaFlex(FECalend)
    Call LimpiaFlex(FECalBPag)
    Call LimpiaFlex(FECalMPag)
    
    For i = 0 To 2
        If OptTipoCuota(i).value Then
            nTipoCuota = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 1
        If OptTipoPeriodo(i).value Then
            nTipoPeriodo = i + 1
            Exit For
        End If
    Next i
    
    '->***** LUCV20180601, Comentó y agregó según ERS022-2018
'     MatCalend = GeneraCalendario(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), _
'                                    CInt(SpnCuotas.valor), CInt(SpnPlazo.valor), _
'                                    CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), nTipoCuota, _
'                                    nTipoPeriodo, nTipoGracia, CInt(TxtPerGra.Text), _
'                                    CInt(TxtDiaFijo.Text), ChkProxMes.value, MatGracia, _
'                                    True, IIf(ChkCuotaCom.value = 1, True, False), bDesemParcial, _
'                                    MatDesPar, , , , _
'                                    optTipoGracia(1).value, CDbl(TxtTasaGracia.Text), _
'                                    CInt(TxtDiaFijo2.Text), nMontoCapInicial, _
'                                    IIf(chkPagoInteres.value = 1, True, False), bRenovarCredito, _
'                                    nInteresAFecha, nIntGraInicial, _
'                                    IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0))
'                                    'WIOR 20131111 AGREGO IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0)
        
        Screen.MousePointer = vbHourglass
        MatCalend = GeneraCalendarioNuevo(CDbl(TxtMonto.Text), _
                                        CDbl(TxtInteres.Text), _
                                        CInt(SpnCuotas.valor), _
                                        CInt(SpnPlazo.valor), _
                                        CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), _
                                        nTipoCuota, _
                                        nTipoPeriodo, _
                                        nTipoGracia, _
                                        CInt(TxtPerGra.Text), _
                                        CInt(TxtDiaFijo.Text), _
                                        ChkProxMes.value, _
                                        MatGracia, _
                                        True, IIf(ChkCuotaCom.value = 1, True, False), bDesemParcial, _
                                        MatDesPar, , , , _
                                        optTipoGracia(1).value, _
                                        CDbl(TxtTasaGracia.Text), _
                                        CInt(TxtDiaFijo2.Text), nMontoCapInicial, _
                                        IIf(chkPagoInteres.value = 1, True, False), bRenovarCredito, _
                                        nInteresAFecha, nIntGraInicial, _
                                        IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0), _
                                        cCtaCodG, lnTasaSegDes, MatCalendSegDes, lnExoSeguroDesgravamen, _
                                        lnMontoPoliza, lnTasaMensualSegInc)
                                        'LUCV20180601, Agregó: cCtaCod, lnTasaSegDes, MatCalendSegDes, lnExoSeguroDesgravamen, pnMontoPoliza, lnTasaMensualSegInc
    '<-***** Fin LUCV20180601
   
    '**DAOR 20070410, Obtener Gastos en Reprogramación**********************
    If bRenovarCredito Then
        Call ObtenerGastosEnReprogramacion
    End If
    '***********************************************************************

    '**PEAC 20080815, Obtener Desgravamen **********************
     If cmdAplicar.Enabled Then
         Call ObtenerDesgravamen
     Else
         For i = 0 To UBound(MatCalend) - 1
             'Monto de Cuota = Capital + Interes Compensatorio + Interes Gracia + Gasto +(Seguro Poliza Incendio + Gracia Poliza Incendio) + Seguro Desgramen
             MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)) + CDbl(MatCalend(i, 8)), "#0.00")
         Next i
     End If
     '***********************************************************************

    nTotalInteres = 0
    nTotalCapital = 0
    nTotalcuotasCONItF = 0
    lnSalCapital = val(TxtMonto.Text)
    
    'Cargamos Grilla [FECalend] del calendario a mostrar
    For i = 0 To UBound(MatCalend) - 1
        FECalend.AdicionaFila
        FECalend.TextMatrix(i + 1, 1) = Trim(MatCalend(i, 0))  'Fecha Venc.
        If i = 0 Then
            If CDate(Trim(MatCalend(i, 0))) <> CDate(txtFechaPago.Text) And nTipoGracia <> 1 Then
                If nTipoGracia <> 6 Then
                    Call LimpiaFlex(FECalend)
                    MsgBox "Falto Presionar Enter en el Campo Fecha de Pago", vbInformation, "Aviso"
                    bErrorValidacion = True
                    Exit Sub
                End If
            End If
        End If
    
        FECalend.TextMatrix(i + 1, 2) = Trim(MatCalend(i, 1))  'Nro. Cuota
        FECalend.TextMatrix(i + 1, 3) = Trim(MatCalend(i, 2))  'Importe Cuota
        FECalend.TextMatrix(i + 1, 4) = Trim(MatCalend(i, 3))  'Amortización Cap.
        
        FECalend.TextMatrix(i + 1, 5) = Format(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)), "0.00") 'Interes Comp. + Gracia
        'FECalend.TextMatrix(i + 1, 5) = Trim(MatCalend(i, 4)) 'Interés Comp.
        'FECalend.TextMatrix(i + 1, 6) = Trim(MatCalend(i, 5)) 'Interes Gracia
        FECalend.TextMatrix(i + 1, 7) = Format(CDbl(MatCalend(i, 15)) + CDbl(MatCalend(i, 16)), "0.00") 'Trim((MatCalend(i, 6))) 'Gastos Comis.
        FECalend.TextMatrix(i + 1, 8) = Trim(MatCalend(i, 8))  'Seg. Desg.
        FECalend.TextMatrix(i + 1, 9) = Trim(MatCalend(i, 13)) 'Seg. Multiriesgo
        FECalend.TextMatrix(i + 1, 10) = Trim(MatCalend(i, 7)) 'Saldo Cap.
        FECalend.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00") 'Cuota + ITF
        
        FECalend.row = i + 1
        FECalend.col = 3
        FECalend.CellForeColor = vbBlue
        
        If Not (i = 0 And nTipoGracia = gColocTiposGraciaCapitalizada) Then
            nTotalCapital = nTotalCapital + CDbl(Trim(MatCalend(i, 3)))
            nTotalInteres = nTotalInteres + CDbl(Trim(MatCalend(i, 4))) + CDbl(Trim(MatCalend(i, 5)))
        End If
        
        nRedondeoITF = fgDiferenciaRedondeoITF(fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))))
        If nRedondeoITF > 0 Then
            FECalend.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))) - nRedondeoITF, "0.00")
        Else
            FECalend.TextMatrix(i + 1, 11) = Format(CDbl(MatCalend(i, 2)) + fgITFCalculaImpuesto(CDbl(MatCalend(i, 2))), "0.00")
        End If
        If Not (nTipoGracia = gColocTiposGraciaCapitalizada And i = 0) Then
            nTotalcuotasCONItF = nTotalcuotasCONItF + CDbl(FECalend.TextMatrix(i + 1, 11))
        End If
    Next i
        
    lblCapital.Caption = Format(nTotalCapital, "#0.00")
    lblInteres.Caption = Format(nTotalInteres, "#0.00")
    lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
    lblTotalCONITF.Caption = Format(nTotalcuotasCONItF, "#0.00")
    FECalend.row = 1
    FECalend.TopRow = 1

    '*** PEAC 20080819, Desarrollo de la Tasa Costo Efectivo Anual **********************
    fraTasaAnuales.Visible = True
    Dim oCredito As COMNCredito.NCOMCredito
    Set oCredito = New COMNCredito.NCOMCredito
    nTasaEfectivaAnual = Round(oCredito.TasaIntPerDias(CDbl(TxtInteres), 360) * 100, 2)
        
    nTasaCostoEfectivoAnual = oCredito.GeneraTasaCostoEfectivoAnual(CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), CDbl(TxtMonto.Text), MatCalend, CDbl(TxtInteres), lsCtaCodLeasing, nTipoPeriodo) 'WIOR 20151223
    lblTasaEfectivaAnual.Caption = nTasaEfectivaAnual & " %"
    lblTasaCostoEfectivoAnual.Caption = nTasaCostoEfectivoAnual & " %"
    Set oCredito = Nothing
    '**********************************************************************************
    
    Screen.MousePointer = vbDefault
    
    If Me.ChkCalMiViv.value = 1 Then
        cmdResumen.Enabled = True
    End If
    
    If UBound(MatCalend) = 0 Then
        cmdImprimir.Enabled = False
    Else
        cmdImprimir.Enabled = True
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdDesembParcial_Click()
Dim nSumaDesPar As Double
Dim i As Integer
    MatCalendDesParcial = frmCredDesembParcial.Inicio(gdFecSis) ', MatCalendDesParcial)
    MatDesPar = MatCalendDesParcial
    If UBound(MatCalendDesParcial) > 0 Then
        bDesembParcialGenerado = True
        nSumaDesPar = 0
        For i = 0 To UBound(MatCalendDesParcial) - 1
            nSumaDesPar = nSumaDesPar + CDbl(MatCalendDesParcial(i, 1))
        Next i
        TxtMonto.Text = Format(nSumaDesPar, "#0.00")
    Else
        bDesembParcialGenerado = False
    End If
End Sub

Private Sub CmdGracia_Click()
Dim oCredito As COMNCredito.NCOMCredito

Set oCredito = New COMNCredito.NCOMCredito

'MAVM 25102010 ***
If CDbl(TxtTasaGracia.Text) <= 0# Then
    MsgBox "Ingrese la Tasa de Interes de Gracia", vbInformation, "Aviso"
    TxtTasaGracia.SetFocus
    Exit Sub
End If
'***

    '11-05-2006
    'Para Generar el interes de Gracia
'    If OptTipoPeriodo(1).value Then
'        Dim nMes As Integer
'        Dim nAnio As Integer
'        Dim nDia As Integer
'        Dim dFecTemp As Date
'        Dim dDesembolso As Date
'
'        dDesembolso = CDate(Format(DTFecDesemb, "dd/mm/yyyy"))
'
'        nMes = Month(dDesembolso)
'        nAnio = Year(dDesembolso)
'        nDia = CInt(TxtDiaFijo.Text)
'
'        If Not (nDia > Day(dDesembolso) And (Not ChkProxMes.value)) Then
'                 nMes = nMes + 1
'                    If nMes > 12 Then
'                        nAnio = nAnio + 1
'                        nMes = 1
'                    End If
'                Else
'                    If nDia > 30 Then
'                        If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                            nMes = nMes + 1
'                        End If
'                    End If
'                End If
'            If nMes = 2 Then
'                If nDia > 28 Then
'                    If nAnio Mod 4 = 0 Then
'                        nDia = 29
'                    Else
'                        nDia = 28
'                    End If
'                End If
'            Else
'                If nDia > 30 Then
'                    If nMes = 4 Or nMes = 6 Or nMes = 9 Or nMes = 11 Then
'                        nDia = 30
'                    End If
'                End If
'            End If
'
'            dFecTemp = CDate(Right("0" & Trim(Str(nDia)), 2) & "/" & Right("0" & Trim(Str(nMes + 1)), 2) & "/" & Trim(Str(nAnio)))
'            MatGracia = frmCredGracia.Inicio(dFecTemp - dDesembolso, oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), dFecTemp - dDesembolso, CDbl(TxtMonto.Text)), CInt(SpnCuotas.Valor), nTipoGracia, psCtaCod)
'    Else
        'Metodo antiguo
        MatGracia = frmCredGracia.Inicio(CInt(TxtPerGra.Text), oCredito.MontoIntPerDias(CDbl(TxtTasaGracia.Text), CInt(TxtPerGra.Text), CDbl(TxtMonto.Text)), CInt(SpnCuotas.valor), nTipoGracia, psCtaCod)
'    End If
    '***********************
    
    Set oCredito = Nothing
    bGraciaGenerada = True
    Call LimpiaFlex(FECalend)
End Sub

Private Sub EjecutaReporte()
Dim loRep As COMNCredito.NCOMCalendario
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String   'cadena q forma
Dim x As Integer
Dim loPrevio As previo.clsprevio
Dim Agencia As String
Dim lnAge As Integer

Dim TCuota As String
Dim Periodo As String
Dim tIndex As Integer
Dim Interes As Double
Dim Monto As Double
Dim Cuota As Double
Dim Plazo As Integer
Dim Vigencia As Date
Dim pnCuotas As Integer
Dim ArrSimulador() As Variant 'WIOR 20150211

If OptTipoCuota(0).value Then
    tIndex = 0
Else
    If OptTipoCuota(1).value Then
        tIndex = 1
    Else
        tIndex = 2
    End If
End If
Select Case tIndex
    Case 0: TCuota = "Cuota Fija"
    Case 1: TCuota = "Cuota Creciente"
    Case 2: TCuota = "Cuota Decreciente"
End Select

'WIOR 20150211 **************************
ReDim ArrSimulador(3)
ArrSimulador(0) = 0 'Para decir que es del simulador
ArrSimulador(1) = 0 'Para decir el envio de estado de cuenta
ArrSimulador(2) = 0 'Seguro Desgravamen

If fbInicioSim Then
    If fraGastoCom.Visible Then
        ArrSimulador(0) = 1
        ArrSimulador(2) = CInt(Trim(Right(Me.cmbSeguroDes.Text, 4)))
        If chkEnvioEst = 1 Then
            If Trim(Right(Me.cmbEnvioEst.Text, 4)) = "2" Then
                ArrSimulador(1) = 1
            End If
        End If
    End If
End If
'WIOR FIN *******************************

    Periodo = IIf(OptTipoPeriodo(0).value = True, "Periodo Fijo - ", "Fecha Fija - ")
    TCuota = Periodo & TCuota
    'MAVM 20121219 ***
    pnCuotas = SpnCuotas.valor
    'pnCuotas = IIf(nTipoGracia = 1, SpnCuotas.valor + 1, SpnCuotas.valor)
    '***
    Set loRep = New COMNCredito.NCOMCalendario
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    lsCadImp = lsCadImp & Chr(10) & loRep.ReporteCalendario(ChkCalMiViv.value + 1, MatCalend, MatResul, _
    TCuota, CDbl(TxtInteres.Text), TxtMonto.Text, SpnCuotas.valor, SpnPlazo.valor, DTFecDesemb.value, nSugerAprob, IIf(bDesemParcial, MatDesPar, ""), gbITFAplica, gnITFPorcent, gnITFMontoMin, cCtaCodG, pnCuotas, _
    nTasaEfectivaAnual, nTasaCostoEfectivoAnual, lsCtaCodLeasing, nTipoGracia, nIntGraInicial, CInt(TxtPerGra.Text), , ArrSimulador) 'DAOR 20070403
    'WIOR 20150211 AGREGO ArrSimulador
    'MAVM 20130209

lsDestino = "P"
Set loRep = Nothing
If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            '*** PEAC 20080723
            'loPrevio.Show lsCadImp, "Calendario de Pagos - Simulacion", True
            loPrevio.Show oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & lsCadImp, "Calendario de Pagos - Simulacion"
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
ElseIf lsDestino = "A" Then
End If

End Sub

Private Sub cmdImprimir_Click()
    'If ChkCalMiViv.value = 0 Then 'WIOR 20151223 - COMENTO
        If Len(Trim(FECalend.TextMatrix(1, 1))) = 0 Then
            MsgBox "No existen datos para imprimir", vbExclamation, "Aviso"
            Exit Sub
        Else
            EjecutaReporte
        End If
    'WIOR 20151223 - COMENTO
    '    Else
    '        If Len(Trim(FECalBPag.TextMatrix(1, 1))) = 0 Then
    '            MsgBox "No existen datos para imprimir", vbExclamation, "Aviso"
    '            Exit Sub
    '        Else
    '            EjecutaReporte
    '        End If
    '    End If
End Sub

'MAVM 20121113 ***
Private Sub cmdResumen_Click()
Dim oGastos As COMDCredito.DCOMGasto
Dim nGastoAdministracion As Double
Dim loRep As COMNCredito.NCOMCalendario
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String   'cadena q forma
Dim x As Integer
Dim loPrevio As previo.clsprevio
Dim Agencia As String
Dim lnAge As Integer

Dim TCuota As String
Dim Periodo As String
Dim tIndex As Integer
Dim Interes As Double
Dim Monto As Double
Dim Cuota As Double
Dim Plazo As Integer
Dim Vigencia As Date
Dim pnCuotas As Integer

If OptTipoCuota(0).value Then
    tIndex = 0
Else
    If OptTipoCuota(1).value Then
        tIndex = 1
    Else
        tIndex = 2
    End If
End If
Select Case tIndex
    Case 0: TCuota = "Cuota Fija"
    Case 1: TCuota = "Cuota Creciente"
    Case 2: TCuota = "Cuota Decreciente"
End Select
 
    Periodo = IIf(OptTipoPeriodo(0).value = True, "Periodo Fijo - ", "Fecha Fija - ")
    TCuota = Periodo & TCuota
    pnCuotas = SpnCuotas.valor
    Set loRep = New COMNCredito.NCOMCalendario
    
    Set oGastos = New COMDCredito.DCOMGasto
    nGastoAdministracion = oGastos.GetGastoAdmMiViv(1224)
    
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    lsCadImp = lsCadImp & Chr(10) & loRep.ReporteResumenMiVivienda(ChkCalMiViv.value + 1, MatCalend, MatResul, _
    TCuota, CDbl(TxtInteres.Text), TxtMonto.Text, SpnCuotas.valor, SpnPlazo.valor, DTFecDesemb.value, nSugerAprob, IIf(bDesemParcial, MatDesPar, ""), gbITFAplica, gnITFPorcent, gnITFMontoMin, cCtaCodG, pnCuotas, _
    nTasaEfectivaAnual, nTasaCostoEfectivoAnual, lsCtaCodLeasing, txtValorInmueble.Text, txtCuotaInicial.Text, txtBonoBuenPagador.Text, TxtPerGra.Text, cmbSegDes.Text, nGastoAdministracion, nCuotMens, nCuotMensBono)

lsDestino = "P"
Set loRep = Nothing
If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            '*** PEAC 20080723
            'loPrevio.Show lsCadImp, "Calendario de Pagos - Simulacion", True
            loPrevio.Show oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & lsCadImp, "Calendario de Pagos - Simulacion"
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
ElseIf lsDestino = "A" Then
End If
    'cmdResumen.Enabled = False 'MAVM 20121115
End Sub
'***
 
Private Sub DTFecDesemb_Change()
     If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
    GenerarFechaPago 'MAVM 30092010
End Sub


Private Sub Form_Load()
    CentraForm Me
    nTasaEfectivaAnual = 0: nTasaCostoEfectivoAnual = 0 'DAOR 20070403
    Call CargaControles
    Set MatCalendDesParcial = Nothing
    bGraciaGenerada = False
    DTFecDesemb.value = gdFecSis 'MAVM 30092010
    Call HabilitaFechaFija(False) 'MAVM 30092010
    'cmbSegDes.ListIndex = 0 ' MAVM 20121113
End Sub
Private Sub OptTipoCuota_Click(Index As Integer)
    Call LimpiaFlex(FECalend)
End Sub

Private Sub optTipoGracia_Click(Index As Integer)
If Index = 0 Then
    'chkIncremenK.Visible = True
    ChkProxMes.Enabled = True
    'chkPagoInteres.Visible = False
Else
    'ARCV 20-07-2006
    If OptTipoPeriodo(1).value Then
        MsgBox "Gracia en Cuotas no es aplicable para este Periodo", vbInformation, "Mensaje"
        Exit Sub
    End If
    '-----------------
    'chkIncremenK.Visible = False
    ChkProxMes.Enabled = False
    ChkProxMes.value = 0
    'chkPagoInteres.Visible = True
End If

CmdGracia.Enabled = False
End Sub

Private Sub OptTipoPeriodo_Click(Index As Integer)
    Call LimpiaFlex(FECalend)
    If Index = 1 Then
        'ARCV 20-07-2006
        If optTipoGracia(1).value Then
            MsgBox "Gracia en Cuotas no es aplicable para este Periodo", vbInformation, "Mensaje"
            Exit Sub
        End If
        '-----------------
        Call HabilitaFechaFija(True)
        optTipoGracia(0).Enabled = False
        optTipoGracia(1).Enabled = False
        'MAVM 30092010 ***
        Frame6.Enabled = False
        ChkPerGra.Enabled = False
        txtFechaPago.Text = DTFecDesemb.value
        GenerarFechaPago
        '***
    Else
        Call HabilitaFechaFija(False)
        optTipoGracia(0).Enabled = True
        optTipoGracia(1).Enabled = True
        GenerarFechaPago 'MAVM 30092010
        ChkPerGra.value = 0 'MAVM 15102010
        TxtDiaFijo.Text = "00" 'MAVM 15102010
    End If
End Sub

Private Sub SpnCuotas_Change()
     If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
    ValidaCuotaBalon 'WIOR 20131129
End Sub

Private Sub spnCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If SpnPlazo.Enabled Then
            SpnPlazo.SetFocus
        Else
            If DTFecDesemb.Enabled Then DTFecDesemb.SetFocus
        End If
    End If
End Sub

Private Sub SpnPlazo_Change()
     If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
    GenerarFechaPago 'MAVM 30092010
    ChkPerGra.value = 0 'MAVM 30092010
End Sub

Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And OptTipoCuota(0).Enabled Then
        OptTipoCuota(0).SetFocus
    End If
End Sub

Private Sub SSCalend_Click(PreviousTab As Integer)
Dim nTotalInteres As Double
Dim nTotalCapital As Double
Dim nTotalCONITF As Double
Dim i As Integer
'Para mostrar los totales de MiVivienda cuando no se muestre Cuota+ITF
Dim bTieneTotalITF As Boolean
'***************************

    If FECalBPag.rows <= 2 Then
        Exit Sub
        lblCapital.Caption = "0.00"
        lblInteres.Caption = "0.00"
        lblTotal.Caption = "0.00"
        lblTotalCONITF.Caption = "0.00"
    End If
    If SSCalend.Tab = 0 Then
            'MAVM 20121113 ***
            'If FECalBPag.TextMatrix(1, 9) = "" Then bTieneTotalITF = True
            If FECalBPag.TextMatrix(1, 11) = "" Then bTieneTotalITF = True
            '***
            If FECalBPag.TextMatrix(1, 9) = "" Then bTieneTotalITF = True
            For i = 0 To UBound(MatCalend) - 1
                nTotalCapital = nTotalCapital + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 4)))
                'MAVM 20121113 ***
                'If Not bTieneTotalITF Then nTotalCONITF = nTotalCONITF + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 9)))
                If Not bTieneTotalITF Then nTotalCONITF = nTotalCONITF + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 11)))
                '***
                nTotalInteres = nTotalInteres + CDbl(Trim(FECalBPag.TextMatrix(i + 1, 5)))
            Next i
            lblCapital.Caption = Format(nTotalCapital, "#0.00")
            lblInteres.Caption = Format(nTotalInteres, "#0.00")
            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
            If Not bTieneTotalITF Then
                lblTotalCONITF.Caption = Format(nTotalCONITF, "#0.00")
            Else
                lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
            End If
            
    Else
            'MAVM 20121113 ***
            'If FECalMPag.TextMatrix(1, 9) = "" Then bTieneTotalITF = True
            If FECalMPag.TextMatrix(1, 11) = "" Then bTieneTotalITF = True
            '***
            For i = 0 To UBound(MatCalend) - 1
                nTotalCapital = nTotalCapital + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 4)))
                'MAVM 20121113 ***
                'If Not bTieneTotalITF Then nTotalCONITF = nTotalCONITF + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 9)))
                If Not bTieneTotalITF Then nTotalCONITF = nTotalCONITF + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 11)))
                '***
                nTotalInteres = nTotalInteres + CDbl(Trim(FECalMPag.TextMatrix(i + 1, 5)))
            Next i
            lblCapital.Caption = Format(nTotalCapital, "#0.00")
            lblInteres.Caption = Format(nTotalInteres, "#0.00")
            lblTotal.Caption = Format(nTotalCapital + nTotalInteres, "#0.00")
            If Not bTieneTotalITF Then
                lblTotalCONITF.Caption = Format(nTotalCONITF, "#0.00")
            Else
                lblTotalCONITF.Caption = Format(CDbl(lblTotal.Caption) + fgITFCalculaImpuesto(CDbl(lblTotal.Caption)), "0.00")
            End If
            
    End If
End Sub

Private Sub TxtDiaFijo_Change()
     If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    If TxtDiaFijo.Text = "" Then TxtDiaFijo.Text = "00"
    If CInt(TxtDiaFijo.Text) > 31 Then
        TxtDiaFijo.Text = "00"
    End If
    Call LimpiaFlex(FECalend)
End Sub

Private Sub TxtDiaFijo_GotFocus()
    fEnfoque TxtDiaFijo
End Sub

Private Sub TxtDiaFijo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        ChkPerGra.SetFocus
    End If
End Sub

Private Sub TxtDiaFijo_LostFocus()
    TxtDiaFijo.Text = Right("00" + Trim(TxtDiaFijo.Text), 2)
End Sub

Private Sub TxtDiaFijo2_Change()
If TxtDiaFijo2.Text = "" Then TxtDiaFijo2.Text = "00"

If CInt(TxtDiaFijo2.Text) = 0 Then
    ChkProxMes.Enabled = True
Else
    ChkProxMes.Enabled = False
End If

If CInt(TxtDiaFijo2.Text) > 31 Then
    TxtDiaFijo2.Text = "00"
End If
Call LimpiaFlex(FECalend)
End Sub

Private Sub TxtDiaFijo2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

'MAVM 30092010 ***
Private Sub txtFechaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        nTipoGracia = 0
'        cmdFechaPago.SetFocus
        If Not Trim(ValidaFecha(txtFechaPago.Text)) = "" Then
            MsgBox Trim(ValidaFecha(txtFechaPago.Text)), vbInformation, "Aviso"
            Exit Sub
        End If
        
        If OptTipoPeriodo(0).value = True Then 'Periodo Fijo
            If (Trim(SpnPlazo.valor) = "" Or SpnPlazo.valor = "0" Or SpnPlazo.valor = "00") Then
                MsgBox "Ingrese el Plazo de las Cuotas del Prestamo", vbInformation, "Aviso"
                If SpnPlazo.Enabled Then SpnPlazo.SetFocus
                Exit Sub
            End If
            If CDate(DTFecDesemb.value + SpnPlazo.valor) > CDate(txtFechaPago.Text) Then
                MsgBox "La Fecha de Pago No puede ser Menor que el Plazo", vbInformation, "Aviso"
                txtFechaPago.Text = CDate(DTFecDesemb.value + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
                txtFechaPago.SetFocus
                ChkPerGra.value = 0
                Exit Sub
            End If
            If txtFechaPago.Text > CDate(DTFecDesemb.value + SpnPlazo.valor) Then
                ChkPerGra.Enabled = True
                ChkPerGra.value = 1
                TxtPerGra.Text = CInt(CDate(txtFechaPago.Text) - CDate(DTFecDesemb.value + SpnPlazo.valor))
            Else
                TxtPerGra.Text = 0
                ChkPerGra.value = 0
            End If
        End If
        If OptTipoPeriodo(1).value = True Then 'Fecha Fija
            If CDate(DTFecDesemb.value) > CDate(txtFechaPago.Text) Then
                MsgBox "La Fecha de Pago No puede ser Menor que la F. Desembolso", vbInformation, "Aviso"
                txtFechaPago.Text = CDate(DTFecDesemb.value + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor))
                txtFechaPago.SetFocus
                ChkPerGra.value = 0
                Exit Sub
            End If
            If Month(DTFecDesemb.value) = Month(txtFechaPago.Text) And Year(DTFecDesemb.value) = Year(txtFechaPago.Text) Then
                ChkProxMes.value = 0
            Else
                ChkProxMes.value = 1
            End If
            If CDate(DTFecDesemb.value + 30) < txtFechaPago.Text Then
                ChkPerGra.Enabled = True
                ChkPerGra.value = 1
                TxtPerGra.Text = CInt(CDate(txtFechaPago.Text) - CDate(DTFecDesemb.value + 30))
            Else
                ChkPerGra.value = 0
            End If
            TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaPago.Text)), 2)
        End If
    End If
End Sub

Private Sub txtInteres_Change()
    If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    TxtTasaGracia.Text = Format(TxtInteres.Text, "#0.00") 'LUCV20180601, Agregó según ERS022-2018
    Call LimpiaFlex(FECalend)
End Sub

Private Sub txtinteres_GotFocus()
    fEnfoque TxtInteres
    TxtTasaGracia.Text = Format(TxtInteres.Text, "#0.00") 'LUCV20180601, Agregó según ERS022-2018
End Sub

Private Sub txtinteres_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtInteres, KeyAscii, , 4)
    If KeyAscii = 13 Then
        If SpnCuotas.Enabled Then
            SpnCuotas.SetFocus
        Else
            If SpnPlazo.Enabled Then SpnPlazo.SetFocus
        End If
        TxtTasaGracia.Text = Format(TxtInteres.Text, "#0.00") 'LUCV20180601, Agregó según ERS022-2018
    End If
End Sub

Private Sub txtinteres_LostFocus()
    If Trim(TxtInteres.Text) = "" Then
        TxtInteres.Text = "0.00"
    Else
        TxtInteres.Text = Format(TxtInteres.Text, "#0.0000")
        TxtTasaGracia.Text = Format(TxtInteres.Text, "#0.00") 'LUCV20180601, Agregó según ERS022-2018
    End If
End Sub

Private Sub txtMonto_Change()
     If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    Call LimpiaFlex(FECalend)
End Sub

Private Sub txtMonto_GotFocus()
    fEnfoque TxtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMonto, KeyAscii)
    If KeyAscii = 13 Then
        TxtInteres.SetFocus
        If ChkCalMiViv.value = 1 Then
            Dim nValVenMin As Double
            Dim nValVenMax As Double
            Dim nPorcentCI As Double
            Dim nBonoBuenPagador As Double
            Dim oParam As COMDCredito.DCOMParametro
            Set oParam = New COMDCredito.DCOMParametro
            If val(txtValorInmueble.Text) = 0# Then
                MsgBox "Favor primero el valor del inmueble ", vbInformation, "Aviso"
                TxtMonto.Text = 0#
                txtValorInmueble.SetFocus
                Exit Sub
            End If
            If TxtMonto < oParam.RecuperaValorParametro(2001) * 14 Then
                MsgBox "Favor ingresar el minino  S/. " & oParam.RecuperaValorParametro(2001) * 14, vbInformation, "Aviso"
                TxtMonto.Text = 0#
                TxtMonto.SetFocus
                Exit Sub
            End If
            
            Call oParam.RecuperaParametrosCalendarioMiViv(nValVenMin, nValVenMax, nPorcentCI, nBonoBuenPagador)
            
            If val(txtValorInmueble.Text) - val(TxtMonto.Text) < 0 Then
                txtCuotaInicial.Text = Format(val(txtValorInmueble.Text) * nPorcentCI, "#0.00")
                TxtMonto.Text = Format(val(txtValorInmueble.Text) - val(txtCuotaInicial.Text), "#0.00")
                Exit Sub
            End If
            
            txtCuotaInicial.Text = Format((val(txtValorInmueble.Text) - val(TxtMonto.Text)), "#0.00")
        End If
    End If
End Sub

Private Sub txtMonto_LostFocus()
    If Trim(TxtMonto.Text) = "" Then
        TxtMonto.Text = "0.00"
    Else
        TxtMonto.Text = Format(TxtMonto.Text, "#0.00")
    End If
End Sub

Private Sub TxtPerGra_Change()
    If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    
    Call LimpiaFlex(FECalend)
End Sub

Private Sub TxtPerGra_GotFocus()
    fEnfoque TxtPerGra
End Sub

Private Sub TxtPerGra_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        'LUCV20180601, Modificó según ERS022-2018
'        If TxtTasaGracia.Enabled Then
'            TxtTasaGracia.SetFocus
'        Else
'            cmdAplicar.SetFocus
'        End If
        cmdAplicar.SetFocus
        'Fin LUCV20180601
    End If
End Sub

Private Sub TxtTasaGracia_Change()
    If Me.ChkCalMiViv.value = 1 And ChkPerGra.value = 1 Then
        bGraciaGenerada = True
    Else
        bGraciaGenerada = False
    End If
    
    Call LimpiaFlex(FECalend)
End Sub

'->*****LUCV2018601, Comentó según ERS022-2018
'Private Sub TxtTasaGracia_GotFocus()
'    fEnfoque TxtTasaGracia
'End Sub
'
'Private Sub TxtTasaGracia_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosDecimales(TxtTasaGracia, KeyAscii, , 4)
'    If KeyAscii = 13 And CmdGracia.Enabled Then
'        CmdGracia.SetFocus
'    End If
'End Sub
'
'Private Sub TxtTasaGracia_LostFocus()
'    If Trim(TxtTasaGracia.Text) = "" Then
'        TxtTasaGracia.Text = "0.00"
'    Else
'        TxtTasaGracia.Text = Format(TxtTasaGracia.Text, "#0.0000")
'    End If
'End Sub
'<-***** Fin LUCV20180601

'**DAOR 20070410, Función que devuelve el tipo de cuota
Private Function getTipoCuota() As Integer
Dim i As Integer
    For i = 0 To 2
        If OptTipoCuota(i).value Then
            getTipoCuota = i + 1
            Exit For
        End If
    Next i
End Function

'**DAOR 20070410, Función que obtiene los gastos en la reprogramación
Private Sub ObtenerGastosEnReprogramacion()
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim i, j As Integer

        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    
        Set oDCredito = New COMDCredito.DCOMCredito
        Set rsCredito = oDCredito.RecuperaDatosParaGenerarGastosEnReprog(sCtaCodRep)
        Set oDCredito = Nothing
                    
        Set oNGasto = New COMNCredito.NCOMGasto
            MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                ChkProxMes.value, MatGracia, ChkCalMiViv.value, ChkCuotaCom.value, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "RP", rsCredito("cTipoGasto"), _
                                CDbl(MatCalend(0, 2)), CDbl(TxtMonto.Text), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                CInt(TxtDiaFijo2.Text), True, , _
                                gnITFMontoMin, gnITFPorcent, gbITFAplica, rsCredito("nExoSeguroDes"))
        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
        '***************************************************************************
        'Adicionamos los Gastos
        '***************************************************************************
        If IsArray(MatGastos) Then
            For j = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                For i = 0 To UBound(MatGastos) - 1 'nNumGasto - 1
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1)) _
                         Or Trim(MatGastos(i, 1)) = "*") Then
                        nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
                    End If
                Next i
                'Add By GITU 06-08-2008
                'Descomentar cuando esten seguros de los cambios GITU
'                If j > 0 Then
'                    If InStr(Trim(Str(nTotalGasto)), ".") > 0 Then
'                        MatCalend(0, 6) = Format(MatCalend(0, 6) + (nTotalGasto - Val(Left(Trim(Str(nTotalGasto)), InStr(Trim(nTotalGasto), ".") - 1))), "#0.00")
'                        MatCalend(j, 6) = Format(Val(Left(Trim(nTotalGasto), InStr(Trim(nTotalGasto), ".") - 1)), "#0.00")
'                    Else
'                        MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'                    End If
'                Else
'                    MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'                End If
                'End GITU
                MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                'MatCalend(j, 2) = Format(CDbl(MatCalend(j, 2)) + CDbl(MatCalend(j, 6)), "#0.00")
            Next j
        End If
    
End Sub

'**PEAC 20080815, Función que obtiene los gastos de desgravamen
Private Sub ObtenerDesgravamen()
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim i, j As Integer
'WIOR 20150210 *********************************
Dim oGastos As COMDCredito.DCOMGasto
'Dim RGastos  As ADODB.Recordset
Dim RGastosSegDes As ADODB.Recordset
Dim RGastosEnvCue As ADODB.Recordset
Dim nMunMesesPorDia, nNumMesesPorMes As Integer
'Dim nIntGraciaCap As Double 'JUEZ 20150307
                    
Set oGastos = New COMDCredito.DCOMGasto

If fbInicioSim And fraGastoCom.Visible Then
    Set RGastosSegDes = oGastos.RecuperaGastosCabecera(1)
    RGastosSegDes.Filter = " nPrdConceptoCod = 1217"
    
    Set RGastosEnvCue = oGastos.RecuperaGastosCabecera(1)
    RGastosEnvCue.Filter = " nPrdConceptoCod = 1249"
End If

Set oGastos = Nothing
'WIOR FIN **************************************

        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    
'        Set oDCredito = New COMDCredito.DCOMCredito
'        Set rsCredito = oDCredito.RecuperaDatosParaGenerarGastosEnReprog(sCtaCodRep)
'        Set oDCredito = Nothing
'
        Set oNGasto = New COMNCredito.NCOMGasto
        MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                ChkProxMes.value, MatGracia, ChkCalMiViv.value, ChkCuotaCom.value, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "SI", "F", _
                                CDbl(MatCalend(0, 2)), CDbl(TxtMonto.Text), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                CInt(TxtDiaFijo2.Text), , , _
                                gnITFMontoMin, gnITFPorcent, gbITFAplica, 0, , , , , , nIntGraInicial, , , IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0))
                                'WIOR 20131111 AGREGO IIf(chkCuotaBalon.value = 1, CInt(uspCuotaBalon.valor), 0)
        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
        '***************************************************************************
        'Adicionamos los Gastos
        '***************************************************************************
               
        If IsArray(MatGastos) Then
            For j = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                nTotalGastoSeg = 0
                'WIOR 20150210 *****************************
                If fbInicioSim And fraGastoCom.Visible Then
                    'If J = 0 Then
                    If j = 0 Or CInt(MatCalend(j, 1)) = 1 Then 'JUEZ 20150307
                        nMunMesesPorDia = Round(DateDiff("d", DTFecDesemb.value, MatCalend(j, 0)) / 30, 0)
                        nNumMesesPorMes = DateDiff("m", DTFecDesemb.value, MatCalend(j, 0))
                        'nIntGraciaCap = IIf(optTipoGracia(0).value And MatCalend(J, 1) = 1, MatCalend(0, 5), 0) 'JUEZ 20150307
                        
                        'MatCalend(i, 6) = Format(TxtMonto.Text * (IIf(Trim(Right(cmbSeguroDes.Text, 4)) = "1", RGastosSegDes!nValor, RGastosSegDes!nValorDosTit) / 100) * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes), "#0.00")
                        MatCalend(j, 6) = Format(TxtMonto.Text * (IIf(Trim(Right(cmbSeguroDes.Text, 4)) = "1", RGastosSegDes!nValor, RGastosSegDes!nValorDosTit) / 100) * IIf(nMunMesesPorDia >= nNumMesesPorMes, nMunMesesPorDia, nNumMesesPorMes), "#0.00")  'JUEZ 20150307
                        'MatCalend(J, 6) = Format(CDbl(MatCalend(J, 6)) + (nIntGraciaCap * (IIf(Trim(Right(cmbSeguroDes.Text, 4)) = "1", RGastosSegDes!nValor, RGastosSegDes!nValorDosTit) / 100)), "#0.00") 'JUEZ 20150307
                    Else
                         MatCalend(j, 6) = Format(MatCalend(j - 1, 7) * (IIf(Trim(Right(cmbSeguroDes.Text, 4)) = "1", RGastosSegDes!nValor, RGastosSegDes!nValorDosTit) / 100), "#0.00")
                    End If
                    
                    FECalend.EncabezadosNombres = "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gast/Comis-Seg.Desg(1)-Seg. Mult.-Saldo Capital-Cuota + ITF"
                    
                    If chkEnvioEst = 1 Then
                        If Trim(Right(Me.cmbEnvioEst.Text, 4)) = "2" Then
                            FECalend.EncabezadosNombres = "-Fecha Venc.-Cuota-Cuotas-Capital-Interes-Int. Gracia-Gast/Com(2)-Seg.Desg(1)-Seg. Mult.-Saldo Capital-Cuota + ITF"
                            MatCalend(j, 8) = Format(RGastosEnvCue!nValor, "#0.00")
                        End If
                    End If
                Else
                'WIOR FIN **********************************
                    For i = 0 To UBound(MatGastos) - 1
                    'Comentado por MAVM para separar los gastos 20100320
    '                    If Trim(Right(MatGastos(I, 0), 2)) = "1" And _
    '                       (Trim(MatGastos(I, 1)) = Trim(MatCalend(J, 1)) _
    '                         Or Trim(MatGastos(I, 1)) = "*") Then
    '                        nTotalGasto = nTotalGasto + CDbl(MatGastos(I, 3))
    '                    End If
                        
                        If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                           (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                        Else
                            If Trim(MatGastos(i, 1)) = "*" Then
                                nTotalGasto = CDbl(MatGastos(i, 3))
                                MatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                            End If
    
                        End If
    
    '                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
    '                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(J, 1)) And Trim(Right(MatGastos(i, 2), 4)) = "1217") Then
    '                        nTotalGastoSeg = nTotalGastoSeg + CDbl(MatGastos(i, 3))
    '                        MatCalend(J, 6) = Format(nTotalGastoSeg, "#0.00")
    '                    Else
    '                        If Trim(MatGastos(i, 1)) = "*" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And Trim(MatGastos(i, 1)) = Trim(MatCalend(J, 1)) And Trim(Right(MatGastos(i, 2), 4)) <> "1217") Then
    '                            nTotalGasto = nTotalGasto + CDbl(MatGastos(i, 3))
    '                            MatCalend(J, 8) = Format(nTotalGasto, "#0.00")
    '                        End If
    '
    '                    End If
                        
                    Next i
                End If  'WIOR 20150210
                'Add By GITU 06-08-2008
                'Descomentar cuando esten seguros de los cambios GITU
'                If j > 0 Then
'                    If InStr(Trim(Str(nTotalGasto)), ".") > 0 Then
'                        MatCalend(0, 6) = Format(MatCalend(0, 6) + (nTotalGasto - Val(Left(Trim(Str(nTotalGasto)), InStr(Trim(nTotalGasto), ".") - 1))), "#0.00")
'                        MatCalend(j, 6) = Format(Val(Left(Trim(nTotalGasto), InStr(Trim(nTotalGasto), ".") - 1)), "#0.00")
'                    Else
'                        MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'                    End If
'                Else
'                    MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
'                End If
                'End GITU
                
                'Comentado por MAVM 20100320
                'MatCalend(J, 6) = Format(nTotalGasto, "#0.00")
                
                'MatCalend(j, 2) = Format(CDbl(MatCalend(j, 2)) + CDbl(MatCalend(j, 6)), "#0.00")
            Next j
        End If
               
        For i = 0 To UBound(MatCalend) - 1
            'MAVM 20100320
            'MatCalend(I, 2) = Format(CDbl(MatCalend(I, 3)) + CDbl(MatCalend(I, 4)) + CDbl(MatCalend(I, 5)) + CDbl(MatCalend(I, 6)), "#0.00")
            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)), "#0.00")
        Next i

End Sub

'MAVM 20121113
Private Sub ObtenerDesgravamenHipot(Optional ByVal pnValorInmueble As Double, Optional ByVal pnTramoConsMonto As Double, Optional pnMontoMivivienda As Currency = 0#, Optional ByVal pnCuotaMivienda As Double = -1)
Dim oNGasto As COMNCredito.NCOMGasto
Dim oDCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nNumGastos As Integer
Dim nTotalGasto As Double
Dim nTotalGastoSeg As Double
Dim i, j As Integer

        ReDim MatDesemb(1, 2)
        MatDesemb(0, 0) = Format(DTFecDesemb.value, "dd/mm/yyyy")
        MatDesemb(0, 1) = Format(TxtMonto.Text, "#0.00")
    
        Set oNGasto = New COMNCredito.NCOMGasto
        MatGastos = oNGasto.GeneraCalendarioGastos_NEW(CDbl(TxtMonto.Text), CDbl(TxtInteres.Text), CInt(SpnCuotas.valor), _
                                CInt(SpnPlazo.valor), CDate(Format(DTFecDesemb.value, "dd/mm/yyyy")), getTipoCuota, _
                                IIf(OptTipoPeriodo(0).value, 1, 2), nTipoGracia, CInt(TxtPerGra.Text), _
                                CDbl(TxtTasaGracia.Text), CInt(IIf(Trim(TxtDiaFijo.Text) = "", "00", TxtDiaFijo.Text)), _
                                ChkProxMes.value, MatGracia, ChkCalMiViv.value, ChkCuotaCom.value, MatCalend_2, _
                                MatDesemb, nNumGastos, gdFecSis, _
                                sCtaCodRep, 1, "SI", "F", _
                                CDbl(MatCalend(0, 2)), IIf(ChkCalMiViv.value, pnTramoConsMonto, CDbl(TxtMonto.Text)), , , , , , , , True, _
                                2, True, False, MatDesemb, False, , _
                                CInt(TxtDiaFijo2.Text), True, , _
                                pnCuotaMivienda, gnITFPorcent, gbITFAplica, 0, , , , IIf(ChkCalMiViv.value, Left(cmbSegDes.Text, 1), ""), pnValorInmueble, , , , , , , , pnMontoMivivienda)
        'ALPA 20141127 pnCuotaMivienda--gnITFMontoMin
        Set oNGasto = Nothing
        Set rsCredito = Nothing
        Call frmCredReprogCred.EstablecerGastos(MatGastos, True, nNumGastos, IIf(OptTipoPeriodo(0).value, 1, 2), CInt(SpnPlazo.valor))
                
        If IsArray(MatGastos) Then
            For j = 0 To UBound(MatCalend) - 1
                nTotalGasto = 0
                nTotalGastoSeg = 0
                For i = 0 To UBound(MatGastos) - 1
                
                If ChkCalMiViv.value = 0 Then 'MAVM 20121113
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                        nTotalGasto = CDbl(MatGastos(i, 3))
                        MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                    Else
                        If Trim(MatGastos(i, 1)) = "*" Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                        End If
                    End If
                'MAVM 20121113
                Else
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And _
                       (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                        If Left(Trim(MatGastos(i, 2)), 2) <> "PO" Then
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                        Else
                            nTotalGasto = CDbl(MatGastos(i, 3))
                            MatCalend(j, 9) = Format(nTotalGasto, "#0.00")
                        End If
                    Else
                        If Trim(MatGastos(i, 1)) = "*" Then
                            If Left(Trim(MatGastos(i, 2)), 2) = "CO" And i <> UBound(MatGastos) - 1 Then
                                nTotalGasto = CDbl(MatGastos(i, 3)) '+ 9 'CDbl(MatGastos(UBound(MatGastos) - 1, 3))
                                MatCalend(j, 8) = Format(nTotalGasto, "#0.00")
                            Else
                                If i <> UBound(MatGastos) - 1 Then
                                    nTotalGasto = CDbl(MatGastos(i, 3))
                                    MatCalend(j, 9) = Format(nTotalGasto, "#0.00")
                                End If
                            End If
                        Else
                            If Left(Trim(MatGastos(i, 2)), 2) <> "PO" And _
                               (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                                nTotalGasto = CDbl(MatGastos(i, 3))
                                MatCalend(j, 6) = Format(nTotalGasto, "#0.00")
                            Else
                                If (Trim(MatGastos(i, 1)) = Trim(MatCalend(j, 1))) Then
                                    nTotalGasto = CDbl(MatGastos(i, 3))
                                    MatCalend(j, 9) = Format(nTotalGasto, "#0.00")
                                End If
                            End If
                        End If
                    End If
                End If
                '***
                Next i
            Next j
        End If
        
        For i = 0 To UBound(MatCalend) - 1
            MatCalend(i, 2) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)), "#0.00")
        Next i
End Sub
'***

'MAVM 28102010 ***
Private Sub GenerarFechaPago()
    If OptTipoPeriodo(0).value = True Then
        txtFechaPago.Text = CDate(DTFecDesemb.value + IIf(SpnPlazo.valor = "", 0, SpnPlazo.valor) + TxtPerGra.Text)
    End If
    If OptTipoPeriodo(1).value = True Then
        If SpnPlazo.Enabled = True Then
            txtFechaPago.Text = CDate(DTFecDesemb.value)
            TxtDiaFijo.Text = Right("00" + Trim(Day(txtFechaPago.Text)), 2)
                                        
            If Month(DTFecDesemb.value) = Month(CDate(txtFechaPago.Text)) And Year(gdFecSis) = Year(CDate(txtFechaPago.Text)) Then
                ChkProxMes.value = 0
            Else
                ChkProxMes.value = 1
            End If
        End If
    End If
End Sub
'***

'MAVM 20121113
Private Sub txtValorInmueble_Change()
    Dim nMontoMV As Double
        If txtValorInmueble.Text = "" Or txtValorInmueble.Text = "." Then
            nMontoMV = 0
        Else
            If IsNumeric(txtValorInmueble.Text) Then
                nMontoMV = CDbl(txtValorInmueble.Text)
            Else
                nMontoMV = 0
            End If
        End If
End Sub
Private Sub txtValorInmueble_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Me.txtValorInmueble.Text) Then Exit Sub
    KeyAscii = NumerosDecimales(txtValorInmueble, KeyAscii, 10)
    Dim nTramoNoConsPorcen As Currency
    lnMontoMivivienda = CDbl(Me.txtValorInmueble.Text)
    If KeyAscii = 13 Then
        Dim nValVenMin As Double
        Dim nValVenMax As Double
        Dim nPorcentCI As Double
        Dim nBonoBuenPagador As Double
        Dim oParam As COMDCredito.DCOMParametro
        Set oParam = New COMDCredito.DCOMParametro
        
        Call oParam.RecuperaParametrosCalendarioMiViv(nValVenMin, nValVenMax, nPorcentCI, nBonoBuenPagador)
        
        If txtValorInmueble.Text < nValVenMin Then
            MsgBox ("El valor del Inmueble no cobertura el minimo"), vbInformation
            Exit Sub
        End If
        If txtValorInmueble.Text > nValVenMax Then
            MsgBox ("El valor del Inmueble supera el maximo"), vbInformation
            Exit Sub
        End If
        
        'WIOR 20151223 *** comento
        ''        txtBonoBuenPagador.Text = Format(nBonoBuenPagador, "#0.00")
        ''        txtValorInmueble.Text = Format(txtValorInmueble.Text, "#0.00")
        ''        txtCuotaInicial.Text = Format((txtValorInmueble.Text) * nPorcentCI, "#0.00")
        '        'Set oParam = New COMDCredito.DCOMParametro
        '        'ALPA 20140511****************************************************
        '        If lnMontoMivivienda > oParam.RecuperaValorParametro(2001) * 50 Then
        '                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador2)
        '        Else
        '                nTramoNoConsPorcen = oParam.RecuperaValorParametro(gColocMiVivBonoBuenPagador)
        '        End If
        '        txtBonoBuenPagador.Text = Format(nTramoNoConsPorcen, "#0.00")
        '        txtCuotaInicial.Text = Format((txtValorInmueble.Text) * nPorcentCI, "#0.00")
        '        TxtMonto.Text = Format(txtValorInmueble.Text - txtCuotaInicial.Text, "#0.00")
        '        'nTramoNoConsMonto = Format((nTramoNoConsPorcen / 100) * CDbl(Me.TxtMonto.Text), "#0.00")
        '        'nTramoConsMonto = Format(CDbl(Me.TxtMonto.Text) - nTramoNoConsMonto, "#0.00")
        '        'nTramoNoConsMonto = Format(CDbl(Me.txtMonto.Text) - nTramoNoConsPorcen, "#0.00")
        '        'nTramoConsMonto = Format(nTramoNoConsPorcen, "#0.00")
        
        'WIOR 20151223 ***
        Dim oDCredito As COMDCredito.DCOMCredito
        Set oDCredito = New COMDCredito.DCOMCredito
        Dim rs As ADODB.Recordset
        
        Set rs = oDCredito.ObtenerValoresNuevoMIVIVIENDA("1", gdFecSis, CDbl(txtValorInmueble.Text))
        'txtCuotaInicial.Enabled = False
        txtBonoBuenPagador.Enabled = False
        TxtMonto.Enabled = False
        
        If Not (rs.EOF And rs.BOF) Then
            If (CInt(rs!nValida) = 1) Then
                txtCuotaInicial.Text = Format(CDbl(rs!nCuotaInicial), "###," & String(15, "#") & "#0.00")
                txtBonoBuenPagador.Text = Format(CDbl(rs!nBonoOtorgado), "###," & String(15, "#") & "#0.00")
                TxtMonto.Text = Format(CDbl(rs!nMOntoCred), "###," & String(15, "#") & "#0.00")
            Else
                MsgBox "El valor del crédito (" & IIf("1" = "1", "S/. ", "$. ") & Format(CDbl(rs!nMOntoCred), "###," & String(15, "#") & "#0.00") & ") tiene que se mayor o igual al " & CDbl(rs!nMinCredUIT) & " de la UIT(S/. " & Format(CDbl(rs!nUIT), "###," & String(15, "#") & "#0.00") & ").", vbInformation, "Aviso"
                txtCuotaInicial.Text = "0.00"
                txtBonoBuenPagador.Text = "0.00"
                TxtMonto.Text = "0.00"
            End If
        End If
        Set oDCredito = Nothing
        'WIOR FIN ********
        txtCuotaInicial.SetFocus
    End If
End Sub

Private Sub txtValorInmueble_LostFocus()
    If Trim(txtValorInmueble.Text) = "" Then
        txtValorInmueble.Text = "0.00"
    Else
        txtValorInmueble.Text = Format(txtValorInmueble.Text, "#0.00")
    End If
End Sub

Private Sub txtValorInmueble_GotFocus()
    fEnfoque TxtMonto
End Sub

Private Sub txtCuotaInicial_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Me.txtCuotaInicial.Text) Then Exit Sub
KeyAscii = NumerosDecimales(txtCuotaInicial, KeyAscii, 10)
    If KeyAscii = 13 Then
        'WIOR 20151223 ***
        Dim oParam As COMDCredito.DCOMParametro
        Dim nValPorCI As Double
        Set oParam = New COMDCredito.DCOMParametro
        nValPorCI = oParam.RecuperaValorParametro(3060)
        'WIOR FIN ********
        
        If Round(txtCuotaInicial.Text, 2) < Round((txtValorInmueble.Text * nValPorCI), 2) Then 'WIOR 20151223
            MsgBox "La cuota inicial minimo debe ser: " & Format(nValPorCI, "#0.00%"), vbInformation, "Aviso"  'WIOR 20151223
            Exit Sub 'WIOR 20151223
        End If
        txtCuotaInicial.Text = Format(txtCuotaInicial.Text, "#0.00")
        'TxtMonto.Text = Format(txtValorInmueble.Text - txtCuotaInicial.Text, "#0.00") 'WIOR 20151223 - COMENTO
        
        'WIOR 20151223 ***
        Dim oDCredito As COMDCredito.DCOMCredito
        Set oDCredito = New COMDCredito.DCOMCredito
        Dim rs As ADODB.Recordset
        
        Set rs = oDCredito.ObtenerValoresNuevoMIVIVIENDA("1", gdFecSis, CDbl(txtValorInmueble.Text), CDbl(txtCuotaInicial.Text))
        txtBonoBuenPagador.Enabled = False
        TxtMonto.Enabled = False
        
        If Not (rs.EOF And rs.BOF) Then
            If (CInt(rs!nValida) = 1) Then
                txtCuotaInicial.Text = Format(CDbl(rs!nCuotaInicial), "###," & String(15, "#") & "#0.00")
                txtBonoBuenPagador.Text = Format(CDbl(rs!nBonoOtorgado), "###," & String(15, "#") & "#0.00")
                TxtMonto.Text = Format(CDbl(rs!nMOntoCred), "###," & String(15, "#") & "#0.00")
            Else
                MsgBox "El valor del crédito (" & IIf("1" = "1", "S/. ", "$. ") & Format(CDbl(rs!nMOntoCred), "###," & String(15, "#") & "#0.00") & ") tiene que se mayor o igual al " & CDbl(rs!nMinCredUIT) & " de la UIT(S/. " & Format(CDbl(rs!nUIT), "###," & String(15, "#") & "#0.00") & ").", vbInformation, "Aviso"
                txtCuotaInicial.Text = "0.00"
                txtBonoBuenPagador.Text = "0.00"
                TxtMonto.Text = "0.00"
            End If
        End If
        Set oDCredito = Nothing
        'WIOR FIN ********
        
        cmbSegDes.SetFocus
    End If
End Sub
Private Sub txtCuotaInicial_Change()
    Dim nMontoMV As Double
        If txtCuotaInicial.Text = "" Or txtCuotaInicial.Text = "." Then
            nMontoMV = 0
        Else
            If IsNumeric(txtCuotaInicial.Text) Then
                nMontoMV = CDbl(txtCuotaInicial.Text)
            Else
                nMontoMV = 0
            End If
        End If
End Sub
'***
'WIOR 20131111 **********************
Private Sub uspCuotaBalon_Change()
ValidaCuotaBalon
End Sub
'WIOR FIN ****************************

'WIOR 20131115 ********************************************************
Private Sub ValidaCuotaBalon()
Dim valor As Integer
Dim valorCB As Integer

If uspCuotaBalon.Visible And chkCuotaBalon.Visible Then
    If chkCuotaBalon.value = 0 Then Exit Sub
    
    If CInt(SpnCuotas.valor) < 2 Then
        chkCuotaBalon.value = 0
        uspCuotaBalon.valor = "0"
        Exit Sub
    End If
    
    If SpnCuotas.valor = 0 Or SpnCuotas.valor = "" Then
        valor = 0
    Else
        valor = CInt(SpnCuotas.valor) - 1
    End If
    
    If uspCuotaBalon.valor = "0" Or uspCuotaBalon.valor = "" Then
        valorCB = 0
    Else
        valorCB = CInt(uspCuotaBalon.valor)
    End If
    
    
    If valor < valorCB Then
        uspCuotaBalon.valor = valor
    End If
End If

End Sub
'WIOR FIN *************************************************************

'WIOR 201050209 *******************************************************
Public Sub InicioSim()
Dim oCons As COMDConstantes.DCOMConstantes
  
fraGastoCom.Visible = True
fbInicioSim = True
Set oCons = New COMDConstantes.DCOMConstantes

Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(7096), cmbSeguroDes)
Call Llenar_Combo_con_Recordset(oCons.RecuperaConstantes(9110), cmbEnvioEst)

Set oCons = Nothing

Me.Show 1
End Sub
'WIOR FIN *************************************************************
    
