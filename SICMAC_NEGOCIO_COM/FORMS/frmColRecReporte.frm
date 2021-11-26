VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColRecReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   1515
   ClientTop       =   1920
   ClientWidth     =   11775
   Icon            =   "frmColRecReporte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmTransExcel 
      Caption         =   "Cargar Excel Transferido"
      Height          =   615
      Left            =   4560
      TabIndex        =   80
      Top             =   6960
      Width           =   2055
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "&Archivo"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame FraAgencias 
      Caption         =   "Agencias"
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
      Height          =   825
      Left            =   5400
      TabIndex        =   51
      Top             =   5640
      Width           =   3405
      Begin VB.CommandButton CmdAgencias 
         Caption         =   "Agencias"
         Height          =   435
         Left            =   180
         TabIndex        =   52
         Top             =   270
         Width           =   1365
      End
      Begin VB.CommandButton CmdAnalistas 
         Caption         =   "Analistas"
         Height          =   405
         Left            =   1740
         TabIndex        =   57
         Top             =   270
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin VB.CheckBox chkTodosLosCastigados 
      Caption         =   "Todos los Castigados"
      Height          =   195
      Left            =   4680
      TabIndex        =   79
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Frame fraPagos 
      Height          =   1110
      Left            =   5400
      TabIndex        =   41
      Top             =   5370
      Width           =   2805
      Begin VB.CheckBox chkPago 
         Caption         =   "Castigados"
         Height          =   210
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         Top             =   180
         Width           =   1125
      End
      Begin VB.CheckBox chkPago 
         Caption         =   "Judiciales"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   1125
      End
      Begin VB.Frame Frame3 
         Caption         =   "Moneda"
         ForeColor       =   &H00800000&
         Height          =   570
         Left            =   120
         TabIndex        =   42
         Top             =   435
         Width           =   2595
         Begin VB.CheckBox chkPagosMoneda 
            Caption         =   "Dólares"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   23
            Top             =   210
            Width           =   990
         End
         Begin VB.CheckBox chkPagosMoneda 
            Caption         =   "Soles"
            Height          =   255
            Index           =   0
            Left            =   255
            TabIndex        =   22
            Top             =   210
            Width           =   990
         End
      End
   End
   Begin VB.Frame FraCastigado 
      Height          =   615
      Left            =   5400
      TabIndex        =   67
      Top             =   4680
      Width           =   2805
      Begin VB.CheckBox ChkTipoBusCastigado 
         Caption         =   "Castigado"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   69
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkTipoBusCastigado 
         Caption         =   "Judicial"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkExportaExcel 
      Caption         =   "Exportar a Excel"
      Height          =   195
      Left            =   4680
      TabIndex        =   78
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Frame FrameMora 
      Caption         =   "Rango de Mora"
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
      Height          =   780
      Left            =   8880
      TabIndex        =   70
      Top             =   5640
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox TxtDiaAtrFin 
         Height          =   330
         Left            =   1320
         TabIndex        =   73
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox TxtDiaAtrIni 
         Height          =   330
         Left            =   480
         TabIndex        =   71
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "de:           a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   480
         Width           =   1110
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10125
      TabIndex        =   77
      Top             =   6360
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9360
      TabIndex        =   76
      Top             =   6375
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   6570
      TabIndex        =   38
      Top             =   6510
      Width           =   5055
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   360
         Left            =   1965
         TabIndex        =   27
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Imprimir"
         Height          =   360
         Left            =   480
         TabIndex        =   26
         Top             =   195
         Width           =   1155
      End
   End
   Begin VB.Frame fraProductos 
      Caption         =   "Tipos de créditos"
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
      Height          =   3975
      Left            =   8280
      TabIndex        =   74
      Top             =   750
      Visible         =   0   'False
      Width           =   3450
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3600
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   6350
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imglstFiguras"
         Appearance      =   1
      End
   End
   Begin VB.Frame FraActProcesales 
      Caption         =   "Periodo"
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
      Height          =   735
      Left            =   5400
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   2955
      Begin MSMask.MaskEdBox MskFecVen 
         Height          =   315
         Left            =   1200
         TabIndex        =   65
         Top             =   280
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Vencimiento:"
         Height          =   435
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo"
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
      Height          =   705
      Left            =   5400
      TabIndex        =   29
      Top             =   30
      Visible         =   0   'False
      Width           =   6225
      Begin MSMask.MaskEdBox mskPeriodo1Al 
         Height          =   330
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         TabIndex        =   31
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame fraAbogado 
      Caption         =   "Abogado"
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
      Height          =   885
      Left            =   8400
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   3285
      Begin VB.OptionButton optAbogado 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optAbogado 
         Caption         =   "Individual"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   18
         Top             =   555
         Width           =   1125
      End
      Begin VB.TextBox txtAbogado 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   19
         Top             =   450
         Width           =   1770
      End
   End
   Begin VB.Frame FraMoneda1 
      Caption         =   "Moneda"
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
      Height          =   825
      Left            =   5460
      TabIndex        =   58
      Top             =   1950
      Width           =   2025
      Begin VB.OptionButton ChkDolares1 
         Caption         =   "Dolares"
         Height          =   285
         Left            =   150
         TabIndex        =   60
         Top             =   480
         Width           =   1665
      End
      Begin VB.OptionButton ChkSoles1 
         Caption         =   "Soles"
         Height          =   285
         Left            =   150
         TabIndex        =   59
         Top             =   210
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.Frame FraEstadoJud 
      Caption         =   "Estados"
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
      Height          =   885
      Left            =   5460
      TabIndex        =   54
      Top             =   990
      Width           =   2055
      Begin VB.OptionButton OptCastigado 
         Caption         =   "Castigado"
         Height          =   285
         Left            =   210
         TabIndex        =   56
         Top             =   540
         Width           =   1755
      End
      Begin VB.OptionButton OptJud 
         Caption         =   "Judicial"
         Height          =   315
         Left            =   210
         TabIndex        =   55
         Top             =   210
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin VB.Frame fraTC 
      Caption         =   "Tipo de Cambio Fijo"
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
      Height          =   780
      Left            =   2595
      TabIndex        =   49
      Top             =   6360
      Width           =   1950
      Begin VB.TextBox txtTipoCambio 
         Height          =   330
         Left            =   135
         TabIndex        =   50
         Top             =   345
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impresión"
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
      Height          =   765
      Left            =   120
      TabIndex        =   37
      Top             =   6375
      Width           =   2475
      Begin VB.OptionButton optOpcionImpresion 
         Caption         =   "Pantalla"
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   24
         Top             =   330
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optOpcionImpresion 
         Caption         =   "Impresora"
         Height          =   210
         Index           =   1
         Left            =   1275
         TabIndex        =   25
         Top             =   330
         Width           =   1020
      End
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Seleccione Reporte"
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
      Height          =   6285
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   5220
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   270
         Top             =   5355
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
               Picture         =   "frmColRecReporte.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmColRecReporte.frx":065C
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmColRecReporte.frx":09AE
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmColRecReporte.frx":0D00
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwOperacion 
         Height          =   5895
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   10398
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
   End
   Begin VB.Frame fraOpciones1 
      Height          =   2925
      Left            =   5400
      TabIndex        =   32
      Top             =   750
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox chkCero 
         Caption         =   "Saldo Cap = 0"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   2400
         Width           =   1005
      End
      Begin VB.CheckBox chkMasCero 
         Caption         =   "Saldo C > 0"
         Height          =   255
         Left            =   200
         TabIndex        =   12
         Top             =   1470
         Width           =   1170
      End
      Begin VB.Frame fraMoneda 
         Height          =   840
         Left            =   1560
         TabIndex        =   35
         Top             =   150
         Width           =   1125
         Begin VB.CheckBox chkMoneda 
            Caption         =   "Soles"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   6
            Top             =   195
            Width           =   990
         End
         Begin VB.CheckBox chkMoneda 
            Caption         =   "Dolares"
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   7
            Top             =   465
            Width           =   900
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Expediente"
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   1275
      End
      Begin VB.Frame fraExpediente 
         Height          =   840
         Left            =   50
         TabIndex        =   36
         Top             =   150
         Width           =   1470
         Begin VB.OptionButton optExpediente 
            Caption         =   "Sin Exp."
            Enabled         =   0   'False
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   540
            Width           =   1095
         End
         Begin VB.OptionButton optExpediente 
            Caption         =   "Con Exp."
            Enabled         =   0   'False
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   285
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.Frame fraEstados 
         Height          =   1020
         Left            =   50
         TabIndex        =   34
         Top             =   1800
         Width           =   1395
         Begin VB.CheckBox chkRefinanciado 
            Caption         =   "Refinanciado"
            Height          =   225
            Left            =   120
            TabIndex        =   53
            Top             =   750
            Width           =   1245
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Vigentes"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   195
            Width           =   1110
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "No Vigentes"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   465
            Width           =   1185
         End
      End
      Begin VB.Frame fraJudicialCastigado 
         Height          =   780
         Left            =   50
         TabIndex        =   33
         Top             =   990
         Width           =   1440
         Begin VB.CheckBox chkJudicialCastigado 
            Caption         =   "Castigados"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1230
         End
         Begin VB.CheckBox chkJudicialCastigado 
            Caption         =   "Judiciales"
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   8
            Top             =   195
            Width           =   1170
         End
      End
   End
   Begin VB.Frame FraTipoC 
      Caption         =   "Tipos de Cobranza"
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
      Height          =   735
      Left            =   5880
      TabIndex        =   61
      Top             =   3960
      Width           =   2295
      Begin VB.CheckBox ChkTipoC 
         Caption         =   "Judicial"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   63
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox ChkTipoC 
         Caption         =   "GECO"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraAnalista 
      Caption         =   "Analista"
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
      Height          =   900
      Left            =   5400
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   2805
      Begin VB.TextBox txtAnalista 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   16
         Top             =   450
         Width           =   990
      End
      Begin VB.OptionButton optAnalista 
         Caption         =   "Individual"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   555
         Width           =   1125
      End
      Begin VB.OptionButton optAnalista 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.Frame fraTop 
      Height          =   1005
      Left            =   5400
      TabIndex        =   43
      Top             =   2400
      Visible         =   0   'False
      Width           =   2805
      Begin VB.CheckBox chkEstadoRecup 
         Caption         =   "Judiciales"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   705
         Width           =   1170
      End
      Begin VB.CheckBox chkEstadoRecup 
         Caption         =   "Castigados"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   47
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtTop 
         Height          =   300
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   45
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "créditos"
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
         Left            =   2040
         TabIndex        =   48
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Los Primeros"
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
         TabIndex        =   44
         Top             =   255
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Archivo de Excel (*.xls)|*.xls"
      FilterIndex     =   1
   End
End
Attribute VB_Name = "frmColRecReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MatAgencias() As String
Private MatProductos() As String

Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Private matAnalista() As String

Dim loRep As COMNColocRec.NCOMColRecRConsulta
Attribute loRep.VB_VarHelpID = -1
Dim Progress As clsProgressBar
Dim nContAna As Integer
Dim P As previo.clsprevio

Private MatCreditos() As String 'MADM 20110513
Dim bValorCredJud As Boolean 'MADM 20110513

Private Function ValorProducto() As String
Dim i As Integer
Dim lsCad As String

lsCad = ""

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

Private Sub chkTodosLosCastigados_Click()
    If chkTodosLosCastigados.value = 1 Then
        CmdAnalistas.Visible = False
    Else
        CmdAnalistas.Visible = True
    End If
End Sub

Private Sub CmdAgencias_Click()
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

Private Sub CmdAnalistas_Click()
    frmSelectAnalistas.SeleccionaAnalistas
End Sub

Private Sub loRep_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub loRep_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub loRep_ShowProgress()
    Progress.ShowForm Me
End Sub

Public Sub inicia(ByVal sCaption As String)
    
'agregado por peac 20070922
    Me.Caption = sCaption
    LlenaArbol
'   vTempo = True
    
    LlenaProductos
    Me.Show 0, MDISicmact
    
    
'comentado por peac 20070922
'    Me.Caption = sCaption
'    LlenaArbol
'    Me.Show 0, MDISicmact
    
End Sub

Private Sub LlenaArbol()
Dim clsGen As DGeneral  'COMDConstSistema.DCOMGeneral ARCV 25-10-200
Dim rsUsu As New ADODB.Recordset
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim lsTipREP As String

    lsTipREP = "1380"
    
    Set clsGen = New DGeneral 'COMDConstSistema.DCOMGeneral
    'ARCV 20-07-2006
    'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
    Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)

    Set clsGen = Nothing
      
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("cOpeCod")
        sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
        Select Case rsUsu("nOpeNiv")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = tvwOperacion.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = tvwOperacion.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
End Sub

Private Sub Check1_Click()
    optExpediente(0).value = True
    optExpediente(0).Enabled = IIf(Check1.value = 1, True, False)
    optExpediente(1).Enabled = IIf(Check1.value = 1, True, False)
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Check1.value = 1 Then
            optExpediente(0).SetFocus
        Else
            chkMoneda(0).SetFocus
        End If
    End If
End Sub

Private Sub chkCero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optOpcionImpresion(0).SetFocus
    End If
End Sub

Private Sub chkEstado_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            chkEstado(1).SetFocus
        ElseIf Index = 1 Then
            chkMasCero.SetFocus
        End If
    End If
End Sub

Private Sub chkEstadoRecup_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        chkEstadoRecup(1).SetFocus
    Else
        Me.optOpcionImpresion(0).SetFocus
    End If
End If
End Sub

Private Sub chkJudicialCastigado_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            chkJudicialCastigado(1).SetFocus
        Else
            chkEstado(0).SetFocus
        End If
    End If
End Sub
 
Private Sub chkMasCero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkCero.SetFocus
    End If
End Sub

Private Sub chkMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            chkMoneda(1).SetFocus
        Else
            chkJudicialCastigado(0).SetFocus
        End If
    End If
End Sub
  
Private Sub chkPago_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        chkPago(1).SetFocus
    Else
        chkPagosMoneda(0).SetFocus
    End If
End If
End Sub
 
Private Sub chkPagosMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        chkPagosMoneda(1).SetFocus
    Else
        optOpcionImpresion(0).SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
Dim oTipCambio As COMDConstSistema.NCOMTipoCambio
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set oTipCambio = New COMDConstSistema.NCOMTipoCambio
txtTipoCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.000")
Set oTipCambio = Nothing
HabilitaControles False, False, False, False, False, False, False, False, False, False, False, False, False, False
End Sub

Private Sub mskPeriodo1Al_KeyPress(KeyAscii As Integer)
Dim nodOpe As Node
Dim sOpe As String
         
    If KeyAscii = 13 Then
        Set nodOpe = tvwOperacion.SelectedItem
        sOpe = Mid(nodOpe.Text, 1, 6)
        Select Case sOpe
            Case gColRecRepPagoCredito, gColRecRepPagoPorFecha
                chkPago(0).SetFocus
            Case gColRecRepPagoAbogado
                optAbogado(0).SetFocus
            Case gColRecRepPagoPorAnalista
                optAnalista(0).SetFocus
            Case Else
                optOpcionImpresion(0).SetFocus
        End Select
        
    End If
End Sub

Private Sub mskPeriodo1Del_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskPeriodo1Al.SetFocus
    End If
End Sub

Private Sub optAbogado_Click(Index As Integer)
txtAbogado.Enabled = IIf(Index = 0, False, True)
txtAbogado.Text = ""

End Sub

Private Sub optAbogado_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            optOpcionImpresion(0).SetFocus
        ElseIf Index = 1 Then
            txtAbogado.SetFocus
        End If
    End If
End Sub

Private Sub OptAnalista_Click(Index As Integer)
txtAnalista.Enabled = IIf(Index = 0, False, True)
txtAnalista.Text = ""

End Sub

Private Sub optAnalista_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            optOpcionImpresion(0).SetFocus
        ElseIf Index = 1 Then
            txtAnalista.SetFocus
        End If
    End If
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
 Dim psNomArchivo As String
 psNomArchivo = ""
 CommonDialog1.ShowOpen
 psNomArchivo = CommonDialog1.Filename
If MsgBox("Seguro de Procesar el Archivo ?", vbYesNo, "Aviso") = vbYes Then
    
    Dim xlApp As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim varMatriz As Variant
    Dim cNombreHoja As String
    Dim i As Long, n As Long, nContCred As Long, nNoContCred As Long
    ReDim Preserve MatCreditos(0)
     
    Set xlApp = New Excel.Application
    nContCred = 0
    nNoContCred = 0
    bValorCredJud = False
    
    If Trim(psNomArchivo) = "" Then
        MsgBox "Debe indicar la ruta del Archivo Excel", vbInformation, "Mensaje"
        Exit Sub
   Else
'        Set rsTrans = New ADODB.Recordset
        Set xlLibro = xlApp.Workbooks.Open(psNomArchivo, True, True, , "")
        cNombreHoja = "Hoja1"
        'validar nombre de hoja
        
        Set xlHoja = xlApp.Worksheets(cNombreHoja)
        varMatriz = xlHoja.Range("A2:AA2000").value
        xlLibro.Close SaveChanges:=False
        xlApp.Quit
        
        Set xlHoja = Nothing
        Set xlLibro = Nothing
        Set xlApp = Nothing
        n = UBound(varMatriz)
       
'        Set rsTrans = New ADODB.Recordset
'         With rsTrans
'            'Crear RecordSet
'            .Fields.Append "cCtaCod", adVarChar, 20
'            .Open
'            'Llenar Recordset
    For i = 8 To n 'MADM 20110513
        If varMatriz(i, 1) = "" Then
            Exit For
        Else
'   .AddNew
'   .Fields("csCtaCod") = lstJudicial.ListItems.Add(, , varMatriz(i, 2))          'Nro de Cred Pig
             If Mid(Trim(varMatriz(i, 2)), 1, 3) = "109" Then
                nContCred = nContCred + 1
                ReDim Preserve MatCreditos(nContCred)
                MatCreditos(nContCred - 1) = varMatriz(i, 2)
                bValorCredJud = True
            Else
                nNoContCred = nNoContCred + 1
                MsgBox "La Estructura Incorrecta: Debe Comenza con NºCredito Celda (A9) , Verifique!!", vbCritical, "Aviso"
                bValorCredJud = False
                ReDim Preserve MatCreditos(0)
                Exit For
            End If
        End If
    Next i
'  End With
   psNomArchivo = ""
   End If
End If
End Sub

Private Sub optExpediente_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkMoneda(0).SetFocus
    End If
End Sub

Private Sub optOpcionImpresion_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
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
    If Mid(Node.Key, 2, 1) = val(Text1.Text) Then
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
    If Mid(Node.Key, 2, 1) <> val(Text1.Text) Then
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
        If Mid(sNode.Key, 2, 1) <> val(Text1.Text) Then
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
            If Mid(TreeView1.Nodes(i).Key, 2, 1) <> nFiltro Then
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

Private Sub tvwOperacion_DblClick()
    
    Dim nodOpe As Node
    Dim sDesc As String
    Set nodOpe = tvwOperacion.SelectedItem
    sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
    EjecutaOperacion CLng(nodOpe.Tag), sDesc
    Set nodOpe = Nothing
End Sub

Private Sub tvwOperacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim nodOpe As Node
        Dim sDesc As String
        Set nodOpe = tvwOperacion.SelectedItem
        sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
        
        EjecutaOperacion CLng(nodOpe.Tag), sDesc
        
        Select Case CLng(nodOpe.Tag)
            Case gColRecRepPagoAbogado, gColRecRepPagoCredito, gColRecRepPagoPorFecha, gColRecRepInformeCreditosNuevos, gColRecRepPagoPorAnalista, gColRecRepResumenMovCobranza
                mskPeriodo1Del.SetFocus
            Case gColRecRepInformeCreditosJudicial
                Check1.SetFocus
            Case gColRecRepCreditosenCJxAnalista
                optAnalista(0).SetFocus
            Case gColRecRepCreditosenCJxAbogado, gColRecRepExpedientesxAbogado
                optAbogado(0).SetFocus
            Case gColRecRepCreditossegunSaldo
                txtTop.SetFocus
        End Select
        
        Set nodOpe = Nothing
    End If
End Sub
Private Sub EjecutaOperacion(ByVal nOperacion As CaptacOperacion, ByVal sDescripcion As String)
Me.Caption = "Reportes de Recuperaciones " & sDescripcion
Me.fraJudicialCastigado.Visible = True
Label2.Visible = True
mskPeriodo1Al.Visible = True
Select Case nOperacion
    Case gColRecRepPagoAbogado '= 138001
        HabilitaControles True, True, False, False, True, False, False
    Case gColRecRepPagoCredito '= 138002
        HabilitaControles True, True, False, False, False, True, False, , , , , , True
    Case gColRecRepPagoPorFecha '= 138003
        '*** PEAC 20080724 - SE AGREGÓ UN PARAMETRO PARA ENVIAR INFORMACION A AEXCEL
        HabilitaControles True, True, False, False, False, True, False, , , , , , True, , , , , True
    Case gColRecRepPagoPorAnalista '= 138004
        HabilitaControles True, True, False, True, False, False, False
    Case gColRecRepInformeCreditosNuevos '= 138011
        HabilitaControles True, True, False, False, False, False, False, False
    Case gColRecRepInformeCreditosJudicial '= 138012
        HabilitaControles True, False, True, False, False, False, False, , True
    
    Case gColRecRepCreditosCastigadosFechas '= 138013
        HabilitaControles True, True, True, False, False, False, False
        Me.mskPeriodo1Al.Text = gdFecDataFM
        Me.mskPeriodo1Del.Text = DateAdd("m", -1, gdFecDataFM)
        Me.chkJudicialCastigado(1).value = 1
        Me.chkJudicialCastigado(0).value = 0
        Me.fraJudicialCastigado.Visible = False
    
    Case gColRecRepCreditosenCJxAnalista '= 138014
        HabilitaControles True, False, False, True, False, False, False, , , , , , , , True
    
    Case gColRecRepCarteraJudCas '138015, peac 20071022 cartera de cred jud y cas
        HabilitaControles True, True, False, False, False, False, False, False, True, True, False, True, False, False, False, False, True
        
    Case gColRecRepCreditosenCJxAbogado '= 138021
        'HabilitaControles True, False, False, False, True, False, False
        HabilitaControles True, True, False, False, True, False, False, True, True
    Case gColRecRepCreditosVigentes '= 138022
        HabilitaControles True, False, False, False, False, False, False
    Case gColRecRepCreditossegunSaldo '= 138023
        HabilitaControles True, False, False, False, False, False, True, , , , , , True
    Case gColRecRepExpedientesxAbogado '= 138024
        HabilitaControles True, True, False, False, True, False, False, True, True
    Case gColRecRepResumenMovCobranza '= 138025
        HabilitaControles True, True, False, False, False, False, False
    Case 138026  'reporte de creditos vencidos
        HabilitaControles True, False, False, False, False, False, False, False, True
    Case 138027 'reporte de creditos en judiciales o castigados por agencia
        HabilitaControles True, False, False, False, False, False, False, False, True, True, , , True
    Case 138028 'Reporte de Clientes  por analista
        HabilitaControles True, True, False, False, False, False, False, False, True, False
    Case 138029 'Reporte por Montos  por analista
        HabilitaControles True, True, False, False, False, False, False, True, True, False
    Case 138030 'Reporte por Montos  por analista
        HabilitaControles True, True, False, False, False, False, False, True, True, False
    Case 138031 'Reporte de Lista de Creditos por Analista
        HabilitaControles True, True, False, False, False, False, False, False, True, True, True, True, True
'        fraJudicialCastigado.Visible = False
'        fraEstados.Visible = False
'        chkMasCero.Visible = False
'        chkCero.Visible = False
    Case 138041, 138042, 138043, 138044, 138045, 138046
        HabilitaControles True, False, False, False, False, False, False, False, True, False, True, False, False, False, False
    Case gColRecRepActuacionesProcesales
        HabilitaControles True, False, False, False, False, False, False, False, False, False, False, False, False, True
    Case gColRecRepActuacionesProcesales2 ' 138035, DAOR 20070122
        HabilitaControles True, True, False, False, True, False, False, False, False, False, False, False, False, False
    Case gColRecRepCreditosCastigados ' 138036, DAOR 20070124 'MODIFICADO POR PEAC
        HabilitaControles True, True, False, False, False, False, False, True, True, False, True, False, False, False, False, False, True, , True
    Case gColRecRepCredJudToCastigar ' 138037, PEAC 20070912
        HabilitaControles True, False, False, False, False, False, False, True, True, False, False, False, False, False, False, True, , , , True
    Case gColRecRepGastosJudCas ' 138038, PEAC 20071010
        HabilitaControles True, True, False, False, False, False, False, False, True, True, False, True, False, False, False, False, True, True
    Case gColRecRepCredSaldosCondo ' 138039, PEAC 20071011
        HabilitaControles True, True, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False
    Case 138005  'ALPA 20080820
        Label2.Visible = False
        mskPeriodo1Al.Visible = False
        HabilitaControles True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True
    Case Else
        HabilitaControles False, False, False, False, False, False, False
        Me.Caption = "Reportes de Recuperaciones"
End Select
    
End Sub

Private Sub cmdAceptar_Click()
Dim nodOpe As Node
Dim sDesc As String
Dim nContAge As Integer
Dim lsTitProductos As String 'peac 20070922
Dim lnPosI As Integer 'peac 20070922

Dim i As Integer
    Set nodOpe = tvwOperacion.SelectedItem
    If Not nodOpe Is Nothing Then
        
        '***PEAC 20080731
        If IsDate(mskPeriodo1Del.Text) Then
            If IsDate(mskPeriodo1Al.Text) Then
                If CDate(mskPeriodo1Del.Text) > CDate(mskPeriodo1Al.Text) Then
                    MsgBox "La fecha inicial no puede ser mayor a la fercha final.", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        
        If mskPeriodo1Del.Visible = True Then
            If Not IsDate(mskPeriodo1Del.Text) Then
                Screen.MousePointer = 0
                MsgBox "La fecha inicial no parece correcta", vbExclamation, "Aviso"
                Exit Sub
            End If
        End If
                
        If mskPeriodo1Al.Visible = True Then
            If Not IsDate(mskPeriodo1Al.Text) Then
                Screen.MousePointer = 0
                MsgBox "La fecha final no parece correcta", vbExclamation, "Aviso"
                Exit Sub
            End If
        End If
                
        Select Case CLng(nodOpe.Tag)
            Case gColRecRepPagoAbogado
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
                If optAbogado(1).value = True Then
                    If Len(Trim(txtAbogado.Text)) <> 13 Then
                        MsgBox "Ud. debe ingresar un abogado para buscar", vbExclamation, "Aviso"
                        txtAbogado.SetFocus
                        Exit Sub
                    End If
                End If
            Case gColRecRepPagoCredito, gColRecRepPagoPorFecha
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If chkPagosMoneda(0).value = 0 And chkPagosMoneda(1).value = 0 Then
                    MsgBox "Seleccione una moneda", vbExclamation, "Aviso"
                    chkPagosMoneda(0).SetFocus
                    Exit Sub
                End If
                If chkPago(0).value = 0 And chkPago(1).value = 0 Then
                    MsgBox "Ud. debe seleccionar una opcion", vbExclamation, "Aviso"
                    chkPago(0).SetFocus
                    Exit Sub
                End If
            Case gColRecRepPagoPorAnalista
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
                If optAnalista(1).value = True Then
                    If Len(Trim(txtAnalista.Text)) <> 4 Then
                        MsgBox "Ud. debe ingresar un analista para buscar", vbExclamation, "Aviso"
                        txtAnalista.SetFocus
                        Exit Sub
                    End If
                End If
            Case gColRecRepInformeCreditosNuevos, gColRecRepResumenMovCobranza
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
                If CLng(nodOpe.Tag) = gColRecRepResumenMovCobranza Then
                    If val(txtTipoCambio.Text) = 0 Then
                        MsgBox "Ingrese un tipo de cambio válido", vbExclamation, "Aviso"
                        txtTipoCambio.SetFocus
                        Exit Sub
                    End If
                End If
            Case gColRecRepInformeCreditosJudicial
                If chkMoneda(0).value = 0 And chkMoneda(1).value = 0 Then
                    MsgBox "Seleccione una moneda", vbExclamation, "Aviso"
                    chkMoneda(0).SetFocus
                    Exit Sub
                End If
                If chkJudicialCastigado(0).value = 0 And chkJudicialCastigado(1).value = 0 Then
                    MsgBox "Ud. debe seleccionar una opcion", vbExclamation, "Aviso"
                    chkJudicialCastigado(0).SetFocus
                    Exit Sub
                End If
                If chkEstado(0).value = 0 And chkEstado(1).value = 0 Then
                    MsgBox "Ud. debe seleccionar una opcion", vbExclamation, "Aviso"
                    chkEstado(0).SetFocus
                    Exit Sub
                End If
                
            Case gColRecRepCreditosCastigadosFechas '138013
                If chkMoneda(0).value = 0 And chkMoneda(1).value = 0 Then
                    MsgBox "Seleccione una moneda", vbExclamation, "Aviso"
                    chkMoneda(0).SetFocus
                    Exit Sub
                End If
                If chkJudicialCastigado(0).value = 0 And chkJudicialCastigado(1).value = 0 Then
                    MsgBox "Ud. debe seleccionar una opcion", vbExclamation, "Aviso"
                    chkJudicialCastigado(0).SetFocus
                    Exit Sub
                End If
                If chkEstado(0).value = 0 And chkEstado(1).value = 0 Then
                    MsgBox "Ud. debe seleccionar una opcion", vbExclamation, "Aviso"
                    chkEstado(0).SetFocus
                    Exit Sub
                End If
            
            Case gColRecRepCreditosenCJxAnalista
                If optAnalista(1).value = True Then
                    If Len(Trim(txtAnalista.Text)) <> 4 Then
                        MsgBox "Ud. debe ingresar un analista para buscar", vbExclamation, "Aviso"
                        txtAnalista.SetFocus
                        Exit Sub
                    End If
                End If
            Case gColRecRepCreditosenCJxAbogado, gColRecRepExpedientesxAbogado
                If optAbogado(1).value = True Then
                    If Len(Trim(txtAbogado.Text)) <> 13 Then
                        MsgBox "Ud. debe ingresar un abogado para buscar", vbExclamation, "Aviso"
                        txtAbogado.SetFocus
                        Exit Sub
                    End If
                End If
            Case gColRecRepCreditossegunSaldo
                If val(txtTipoCambio.Text) = 0 Then
                    MsgBox "Ud. debe ingresar un tipo de cambio válido", vbExclamation, "Aviso"
                    txtTipoCambio.SetFocus
                    Exit Sub
                End If
                If val(txtTop.Text) = 0 Then
                    MsgBox "Ud. debe ingresar un tope válido", vbExclamation, "Aviso"
                    txtTop.SetFocus
                    Exit Sub
                End If
                If chkEstadoRecup(0).value = 0 And chkEstadoRecup(1).value = 0 Then
                    MsgBox "Ud. debe seleccionar una opción", vbExclamation, "Aviso"
                    chkEstadoRecup(0).SetFocus
                    Exit Sub
                End If
                
          Case 138031
          
                 If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
          
                If CmdAnalistas.Visible Then
                    ReDim matAnalista(0)
                    nContAna = 0
                    For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
                        If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                            nContAna = nContAna + 1
                            ReDim Preserve matAnalista(nContAna)
                           matAnalista(nContAna - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
                       End If
                    Next i
                    If UBound(matAnalista) = 0 Then
                         MsgBox "Debe Seleccionar por lo Menos un Analista", vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
          Case gColRecRepActuacionesProcesales
                'Validacion
                If IsDate(MskFecVen.Text) = False Then
                   MsgBox "Ingrese un Fecha Correcta", vbInformation, "Aviso"
                   MskFecVen.SetFocus
                   Exit Sub
                End If
          Case gColRecRepActuacionesProcesales2 ' 138035 , DAOR 20070122
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
                If optAbogado(1).value = True Then
                    If Len(Trim(txtAbogado.Text)) <> 13 Then
                        MsgBox "Ud. debe ingresar un abogado para buscar", vbExclamation, "Aviso"
                        txtAbogado.SetFocus
                        Exit Sub
                    End If
                End If
            Case gColRecRepCreditosCastigados ' 138036 , DAOR 20070124 'modificado por peac
                
                ReDim MatProductos(0)
                nContAge = 0
            
                For i = 1 To TreeView1.Nodes.Count
                    If TreeView1.Nodes(i).Checked = True Then
                        If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                            nContAge = nContAge + 1
                            ReDim Preserve MatProductos(nContAge)
                            MatProductos(nContAge - 1) = Trim(Mid(TreeView1.Nodes(i).Key, 2, 3))
                            lnPosI = 0
                            lnPosI = InStr(1, TreeView1.Nodes(i).Text, " ")
                            If lnPosI > 1 Then
                                lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & " " & Mid(TreeView1.Nodes(i).Text, lnPosI + 1, 3) & "/"
                            Else
                                lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & "/"
                            End If
                        End If
                    End If
                Next
                If Len(lsTitProductos) > 1 Then
                    lsTitProductos = Left(lsTitProductos, Len(lsTitProductos) - 1)
                End If
                
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
                
                If CmdAnalistas.Visible Then
                    ReDim matAnalista(0)
                    nContAna = 0
                    For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
                        If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                            nContAna = nContAna + 1
                            ReDim Preserve matAnalista(nContAna)
                           matAnalista(nContAna - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
                       End If
                    Next i
                    If UBound(matAnalista) = 0 Then
                         MsgBox "Debe Seleccionar por lo Menos un Analista", vbInformation, "Aviso"
                         Exit Sub
                    End If
                End If
                
            Case gColRecRepCredJudToCastigar ' 138037 , PEAC 20070912
            
                If FrameMora.Visible = True Then
                    If IsNumeric(TxtDiaAtrIni.Text) Then
                        If IsNumeric(TxtDiaAtrFin.Text) Then
                            If val(TxtDiaAtrFin.Text) < val(TxtDiaAtrIni.Text) Then
                                MsgBox "El nro. de dias final no puede ser menor al nro. de dias inicial", vbExclamation, "Aviso"
                                TxtDiaAtrFin.SetFocus
                                Exit Sub
                            End If
                        Else
                            MsgBox "Ingrese un nro. de dias válido", vbExclamation, "Aviso"
                            TxtDiaAtrFin.SetFocus
                            Exit Sub
                        End If
                    Else
                        MsgBox "Ingrese un nro. de dias válido", vbExclamation, "Aviso"
                        TxtDiaAtrIni.SetFocus
                        Exit Sub
                    End If
                End If
            
            Case gColRecRepGastosJudCas ' 138038 , PEAC 20071010

                ReDim MatProductos(0)
                nContAge = 0
            
                For i = 1 To TreeView1.Nodes.Count
                    If TreeView1.Nodes(i).Checked = True Then
                        If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                            nContAge = nContAge + 1
                            ReDim Preserve MatProductos(nContAge)
                            MatProductos(nContAge - 1) = Trim(Mid(TreeView1.Nodes(i).Key, 2, 3))
                            lnPosI = 0
                            lnPosI = InStr(1, TreeView1.Nodes(i).Text, " ")
                            If lnPosI > 1 Then
                                lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & " " & Mid(TreeView1.Nodes(i).Text, lnPosI + 1, 3) & "/"
                            Else
                                lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & "/"
                            End If
                        End If
                    End If
                Next
                If Len(lsTitProductos) > 1 Then
                    lsTitProductos = Left(lsTitProductos, Len(lsTitProductos) - 1)
                End If

                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If

            Case gColRecRepCredSaldosCondo ' 138039 , PEAC 20071011

                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If
            Case 138005 'ALPA
                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
            Call GenAdjudicados138005
            Case gColRecRepCarteraJudCas ' 138015 , PEAC 20071023

                ReDim MatProductos(0)
                nContAge = 0
            
                For i = 1 To TreeView1.Nodes.Count
                    If TreeView1.Nodes(i).Checked = True Then
                        If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                            nContAge = nContAge + 1
                            ReDim Preserve MatProductos(nContAge)
                            MatProductos(nContAge - 1) = Trim(Mid(TreeView1.Nodes(i).Key, 2, 3))
                            lnPosI = 0
                            lnPosI = InStr(1, TreeView1.Nodes(i).Text, " ")
                            If lnPosI > 1 Then
                                lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & " " & Mid(TreeView1.Nodes(i).Text, lnPosI + 1, 3) & "/"
                            Else
                                lsTitProductos = lsTitProductos & Trim(Mid(TreeView1.Nodes(i).Text, 1, 3)) & "/"
                            End If
                        End If
                    End If
                Next
                If Len(lsTitProductos) > 1 Then
                    lsTitProductos = Left(lsTitProductos, Len(lsTitProductos) - 1)
                End If

                If IsDate(mskPeriodo1Del.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Del.SetFocus
                    Exit Sub
                End If
                If IsDate(mskPeriodo1Al.Text) = False Then
                    MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
                    mskPeriodo1Al.SetFocus
                    Exit Sub
                End If

        End Select
          
        sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
        EjecutaReporte CLng(nodOpe.Tag), sDesc
        
    End If
    Set nodOpe = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub HabilitaControles(ByVal pbcmdImprimir As Boolean, _
        ByVal pbFraPeriodo1 As Boolean, ByVal pbFraopciones1 As Boolean, _
        ByVal pbFraAnalista1 As Boolean, ByVal pbFraAbogado1 As Boolean, _
        ByVal pbFraPagos1 As Boolean, ByVal pbFraTop1 As Boolean, _
        Optional ByVal pbfraTC As Boolean = True, Optional ByVal pbFraAgencias As Boolean = False, _
        Optional ByVal pbFraJudicast As Boolean = False, Optional ByVal pbAnalistas2 As Boolean = False, _
        Optional ByVal pbMoneda1 As Boolean = False, Optional pbTipoC As Boolean = False, Optional pbActProcesales As Boolean = False, Optional pbFraCastigado As Boolean = False, _
        Optional pbMora As Boolean = False, Optional pbFraProductos As Boolean = False, Optional pbExportaExcel As Boolean = False, Optional pbTodosLosCastigados As Boolean = False, Optional pFrmTransExcel = False)

        '*** PEAC 20080724 - SE ADICIONO EL PARAMETRO "pbExportaExcel" PARA INDICAR SI EXPORTA A EXCEL LA INFORMACION

        FrameMora.Visible = pbMora ' peac 20070917
        fraProductos.Visible = pbFraProductos ' peac 20070921
        
        Me.chkExportaExcel.Visible = pbExportaExcel '*** PEAC 20080724
        Me.chkTodosLosCastigados.Visible = pbTodosLosCastigados ''*** PEAC 20090126
        
        Me.cmdAceptar.Visible = pbcmdImprimir
        Me.fraPeriodo1.Visible = pbFraPeriodo1
        Me.fraOpciones1.Visible = pbFraopciones1
        Me.fraAnalista.Visible = pbFraAnalista1
        Me.fraAbogado.Visible = pbFraAbogado1
        Me.fraPagos.Visible = pbFraPagos1
        Me.fraTop.Visible = pbFraTop1
        Me.fraTC.Visible = pbfraTC
        Me.FraEstadoJud.Visible = pbFraJudicast
        FraAgencias.Visible = pbFraAgencias
        CmdAnalistas.Visible = pbAnalistas2
        FraMoneda1.Visible = pbMoneda1
        FraTipoC.Visible = pbTipoC
        FraActProcesales.Visible = pbActProcesales
        FraCastigado.Visible = pbFraCastigado
        FrmTransExcel.Visible = pFrmTransExcel
        Limpia
     
End Sub

Private Sub Limpia()
Dim i As Integer
    mskPeriodo1Del.Text = "__/__/____"
    mskPeriodo1Al.Text = "__/__/____"
    Check1.value = 0
    optExpediente(0).value = True
    chkJudicialCastigado(0).value = 0
    chkJudicialCastigado(1).value = 0
    chkMoneda(0).value = 0
    chkMoneda(1).value = 0
    chkEstado(0).value = 0
    chkEstado(1).value = 0
    chkMasCero.value = 0
    chkCero.value = 0
    optOpcionImpresion(0).value = True
    optAnalista(0).value = True
    optAbogado(0).value = True
    chkPago(0).value = 0
    chkPagosMoneda(0).value = 0
    chkPago(1).value = 0
    chkPagosMoneda(1).value = 0
    chkEstadoRecup(0).value = 0
    chkEstadoRecup(1).value = 0
    ChkTipoC(0).value = 0
    ChkTipoC(1).value = 0
    txtTop.Text = ""
    
    'peac 20070921
     For i = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(i).Checked = False
        'TreeView1.Nodes(i).Expanded = False
    Next
    
    'peac 20070920
    TxtDiaAtrIni.Text = 16
    TxtDiaAtrFin.Text = 90
    
End Sub
Private Sub EjecutaReporte(ByVal nOperacion As CaptacOperacion, ByVal sDescOperacion As String)
Dim sMoneda As Byte
Dim sEstados As String
Dim sConSinExp As Byte

Dim STIPOC As String

Dim lscadimp As String
Dim lsCadTempo As String
Dim loPrevio As previo.clsprevio
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim x As Integer

Dim sRotulo As String
Dim sNomAgeFiltro As String
    
Dim oDAgencias As COMDConstantes.DCOMAgencias

Dim lsmensaje As String
Dim loRep  As COMNColocRec.NCOMColRecRConsulta
Dim rsAP As ADODB.Recordset
Dim lsListaAgencias As String
Dim LslistaCreditos As String
Dim i As Integer

Dim nEstadoJud As String, nmoneda As Integer

    sEstados = ""
    If ChkTipoC(0).value = 1 And ChkTipoC(1).value = 1 Then
        STIPOC = "('" & gColRecTipCobJudicial & "'," & "'" & gColRecTipCobExtraJudi & "')"
    ElseIf ChkTipoC(1).value = 1 Then
        STIPOC = "('" & gColRecTipCobExtraJudi & "')"
    ElseIf ChkTipoC(0).value = 1 Then
        STIPOC = "('" & gColRecTipCobJudicial & "')"
    Else
        STIPOC = ""
    End If
    Select Case nOperacion
        Case gColRecRepPagoAbogado
            Set loRep = New COMNColocRec.NCOMColRecRConsulta 'NColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lscadimp = loRep.nRepo138001_ListadoPagosxAbogado(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, IIf(optAbogado(0).value = True, 2, 1), Trim(txtAbogado.Text), gImpresora)
        
        Case gColRecRepPagoCredito, gColRecRepPagoPorFecha
            If chkPagosMoneda(0).value = 1 Then
                If chkPagosMoneda(1).value = 1 Then
                    sMoneda = 3 'ambas
                Else
                    sMoneda = 1 'soles
                End If
            Else
                If chkPagosMoneda(1).value = 1 Then
                    sMoneda = 2 'dolares
                Else
                End If
            End If
            
            
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            
            If nOperacion = gColRecRepPagoCredito Then
                 If chkPago(0).value = 1 Then  'Judicial
                    sEstados = "'" & gColocEstRecVigJud & "', '" & gColocEstRecCanJud & "'"
                    lscadimp = loRep.nRepo138002_ListadoPagosdeCreditos(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, sEstados, sMoneda, , STIPOC, gImpresora)
                End If
                If chkPago(1).value = 1 Then
                    sEstados = "'" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "'"
                    If chkPago(0).value = 1 Then
                        lscadimp = lscadimp '& Chr(12) & " "
                    Else
                        lscadimp = ""
                    End If
                    lscadimp = lscadimp & loRep.nRepo138002_ListadoPagosdeCreditos(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, sEstados, sMoneda, , STIPOC, gImpresora)
                End If
            ElseIf nOperacion = gColRecRepPagoPorFecha Then
                
                '*** PEAC 20080724
                If Me.chkExportaExcel.value = 1 Then
                
                    If chkPago(0).value = 1 And chkPago(1).value = 1 Then  'Judicial y castigado
                        sEstados = "'" & gColocEstRecVigJud & "', '2205', '" & gColocEstRecCanJud & "','" & gColocEstRecVigCast & "','2206', '" & gColocEstRecCanCast & "'"
                    ElseIf chkPago(0).value = 1 Then 'judicial
                        sEstados = "'" & gColocEstRecVigJud & "', '2205', '" & gColocEstRecCanJud & "'"
                    Else ' castigado
                        sEstados = "'" & gColocEstRecVigCast & "','2206', '" & gColocEstRecCanCast & "'"
                    End If
                
                    Call MostrarPagosPorFechas(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, sEstados, sMoneda, STIPOC)
                    
                Else
                
                    If chkPago(0).value = 1 Then  'Judicial
                        
                        sEstados = "'" & gColocEstRecVigJud & "', '2205', '" & gColocEstRecCanJud & "'"
                        lsCadTempo = loRep.nRepo138003_ListadoPagosxFechas(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, sEstados, sMoneda, STIPOC, gImpresora)
                        
                    End If
                    If chkPago(1).value = 1 Then
                        sEstados = "'" & gColocEstRecVigCast & "','2206', '" & gColocEstRecCanCast & "'"
                        If chkPago(0).value = 1 Then
                            lsCadTempo = lsCadTempo '& Chr(12) & " "
                        Else
                            lsCadTempo = ""
                        End If
                        
                        lsCadTempo = lsCadTempo & loRep.nRepo138003_ListadoPagosxFechas(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, sEstados, sMoneda, STIPOC, gImpresora)
                        
                    End If
                End If
                
                lscadimp = lsCadTempo
                
            End If
        Case gColRecRepPagoPorAnalista
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lscadimp = loRep.nRepo138004_ListadoPagosxAnalista(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, IIf(optAnalista(0).value = True, 2, 1), Trim(txtAnalista.Text), gImpresora)
        Case gColRecRepInformeCreditosNuevos
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lscadimp = loRep.nRepo138011_ListadoCreditosNuevos(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gImpresora)
        Case gColRecRepInformeCreditosJudicial
            sRotulo = ""
            If Check1.value = 0 Then
                sConSinExp = 0 'sin considerar expedientes
                sRotulo = "Sin Exp "
            Else
                If optExpediente(0).value = True Then
                    sConSinExp = 1 'con expedientes
                    sRotulo = " Con Exp "
                Else
                    sConSinExp = 2 'sin expedientes
                    sRotulo = " Sin Exp "
                End If
            End If
            If chkMoneda(0).value = 1 Then
                If chkMoneda(1).value = 1 Then
                    sMoneda = 3 'ambas
                Else
                    sMoneda = 1 'soles
                End If
            Else
                If chkMoneda(1).value = 1 Then
                    sMoneda = 2 'dolares
                Else
                End If
            End If
            
            
            Dim sR1 As String
            sR1 = ""
            If chkJudicialCastigado(0).value = 1 Then 'Judicial
                If chkEstado(0).value = 1 Then 'vigente
                    sEstados = "'" & gColocEstRecVigJud & "','2205'"
                    sR1 = "- Vigente "
                End If
                If chkEstado(1).value = 1 Then 'No vigente
                    sR1 = sR1 & "- No Vigente "
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'" & gColocEstRecCanJud & "'"
                    Else
                        sEstados = sEstados & ", '" & gColocEstRecCanJud & "'"
                    End If
                End If
                
                If chkRefinanciado.value = 1 Then
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'2205'"
                    Else
                        sEstados = sEstados & ", '2205'"
                    End If
                End If
                sRotulo = sRotulo & "[Judicial " & sR1 & "- ]"
            End If
            
            sR1 = ""
            If chkJudicialCastigado(1).value = 1 Then 'Castigado
                If chkEstado(0).value = 1 Then 'vigente
                    sR1 = "- Vigente "
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'" & gColocEstRecVigCast & "','2206'"
                    Else
                        sEstados = sEstados & ", '" & gColocEstRecVigCast & "','2206'"
                    End If
                End If
                If chkEstado(1).value = 1 Then 'No vigente
                    sR1 = sR1 & "- No Vigente "
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'" & gColocEstRecCanCast & "'"
                    Else
                        sEstados = sEstados & ", '" & gColocEstRecCanCast & "'"
                    End If
                End If
                If chkRefinanciado.value = 1 Then
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'2206'"
                    Else
                        sEstados = sEstados & ", '2206'"
                    End If
                End If
                
                sRotulo = Trim(sRotulo & "  [Castigado " & sR1 & " - ]")
            End If
            
            Dim sAgencia As String
            
            
            For i = 0 To UBound(MatAgencias) - 1
                If i = 0 Then
                    sAgencia = "'" & MatAgencias(i) & "'"
                Else
                    sAgencia = sAgencia & "'" & MatAgencias(i) & "'"
                End If
            Next i
             
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            
            lscadimp = loRep.nRepo138012_ListadoCreditosEnJudicial(IIf(sConSinExp = 0, 0, 1), sConSinExp, sMoneda, sEstados, chkMasCero.value, chkCero.value, sRotulo, sAgencia, gImpresora)
            'lsCadImp = loRep.nRepo138012_ListadoCreditosEnJudicial(IIf(sConSinExp = 0, 0, 1), sConSinExp, sMoneda, sEstados, chkMasCero.value, chkCero.value)
        
        Case gColRecRepCreditosenCJxAnalista
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            If val(txtTipoCambio.Text) = 0 Then
                MsgBox "Ingrese un T.C. válido", vbExclamation, "Aviso"
                txtTipoCambio.SetFocus
                Exit Sub
            End If
            Dim sCadena As String
            sCadena = Trim(txtAnalista.Text)
            If ChkTipoBusCastigado(0).value = 1 And ChkTipoBusCastigado(1).value = 1 Then ' 1= Judicial+Castigado /2=Judicial/3=Castigado
                sCadena = sCadena & "1"
            ElseIf ChkTipoBusCastigado(0).value = 1 Then
                sCadena = sCadena & "2"
            ElseIf ChkTipoBusCastigado(1).value = 1 Then
                sCadena = sCadena & "3"
            End If
            
            lscadimp = loRep.nRepo138014_ListadoCreditosxAnalista(IIf(optAnalista(0).value = True, 2, 1), Trim(sCadena), val(txtTipoCambio.Text), gImpresora)
            
        
        Case 138013
            
            If Check1.value = 0 Then
                sConSinExp = 0 'sin considerar expedientes
            Else
                If optExpediente(0).value = True Then
                    sConSinExp = 1 'con expedientes
                Else
                    sConSinExp = 2 'sin expedientes
                End If
            End If
            If chkMoneda(0).value = 1 Then
                If chkMoneda(1).value = 1 Then
                    sMoneda = 3 'ambas
                Else
                    sMoneda = 1 'soles
                End If
            Else
                If chkMoneda(1).value = 1 Then
                    sMoneda = 2 'dolares
                Else
                End If
            End If
            If chkJudicialCastigado(0).value = 1 Then 'Judicial
                If chkEstado(0).value = 1 Then 'vigente
                    sEstados = "'" & gColocEstRecVigJud & "','2205'"
                End If
                If chkEstado(1).value = 1 Then 'No vigente
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'" & gColocEstRecCanJud & "'"
                    Else
                        sEstados = sEstados & ", '" & gColocEstRecCanJud & "'"
                    End If
                End If
            End If
            If chkJudicialCastigado(1).value = 1 Then 'Castigado
                If chkEstado(0).value = 1 Then 'vigente
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'" & gColocEstRecVigCast & "','2206'"
                    Else
                        sEstados = sEstados & ", '" & gColocEstRecVigCast & "'"
                    End If
                End If
                If chkEstado(1).value = 1 Then 'No vigente
                    If Len(Trim(sEstados)) = 0 Then
                        sEstados = "'" & gColocEstRecCanCast & "'"
                    Else
                        sEstados = sEstados & ", '" & gColocEstRecCanCast & "'"
                    End If
                End If
            End If
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
        lscadimp = loRep.nRepo138013_(IIf(sConSinExp = 0, 0, 1), sConSinExp, sMoneda, sEstados, _
        chkMasCero.value, chkCero.value, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gImpresora)
          
        Case gColRecRepCreditosenCJxAbogado
            
            If val(txtTipoCambio.Text) = 0 Then
                MsgBox "Ingrese un T.C. válido", vbExclamation, "Aviso"
                txtTipoCambio.SetFocus
                Exit Sub
            End If
        
'            Set loRep = New COMNColocRec.NCOMColRecRConsulta
'            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'            lscadimp = loRep.nRepo138021_ListadoCreditosEnCJxAbogado(IIf(optAbogado(0).value = True, 2, 1), Trim(txtAbogado.Text), Val(txtTipoCambio.Text), gImpresora)

            '*** PEAC 20080929
           Call nRepo138021_ListadoExcelCreditosEnCJxAbogado(IIf(optAbogado(0).value = True, 2, 1), val(txtTipoCambio.Text), Trim(txtAbogado.Text), Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
        
        
        Case gColRecRepCreditossegunSaldo
            sEstados = ""
            If chkEstadoRecup(0).value = 1 Then  'Judicial
                sEstados = "'" & gColocEstRecVigJud & "', '" & gColocEstRecCanJud & "'"
            End If
            If chkEstadoRecup(1).value = 1 Then
                lsCadTempo = "'" & gColocEstRecVigCast & "', '" & gColocEstRecCanCast & "'"
                If sEstados = "" Then
                    sEstados = lsCadTempo
                Else
                    sEstados = sEstados & ", " & lsCadTempo
                End If
            End If
             
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lscadimp = loRep.nRepo138023_ListadoSaldoCapital(val(txtTop.Text), val(txtTipoCambio.Text), sEstados, gdFecSis, STIPOC, gImpresora)
        Case gColRecRepExpedientesxAbogado
            
            If val(txtTipoCambio.Text) = 0 Then
                MsgBox "Ingrese un T.C. válido", vbExclamation, "Aviso"
                txtTipoCambio.SetFocus
                Exit Sub
            End If
            
'            Set loRep = New COMNColocRec.NCOMColRecRConsulta
'            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'            lscadimp = loRep.nRepo138024_ListadoExpedientesxAbogado(IIf(optAbogado(0).value = True, 2, 1), Trim(txtAbogado.Text), Val(txtTipoCambio.Text), gImpresora)

            '*** PEAC 20080926
           Call nRepo138024_ListadoExcelExpedientesxAbogado(IIf(optAbogado(0).value = True, 2, 1), val(txtTipoCambio.Text), Trim(txtAbogado.Text), Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, MatAgencias, gsCodAge, gsCodUser, gdFecSis, gsNomAge, gsNomCmac)
            
        Case gColRecRepResumenMovCobranza
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lscadimp = loRep.nRepo138025_ResumenMovimientoCJ(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, val(txtTipoCambio.Text), gImpresora)
        
        Case 138031
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            
            If OptJud.value = True Then
                nEstadoJud = 1
            Else
                nEstadoJud = 2
            End If
            
            If ChkSoles1.value = True Then
                nmoneda = 1
            Else
                nmoneda = 2
            End If
            Call loRep.Inicio(gsNomCmac, gsNomAge, gsCodUser, gdFecSis)
            lscadimp = loRep.Repo_138031(nEstadoJud, MatAgencias, matAnalista, nmoneda, , gImpresora, mskPeriodo1Del.Text, mskPeriodo1Al.Text)
        Case 138041, 138042, 138043, 138044, 138045, 138046
         'Imprimir modelo de Cartas
        
        Case gColRecRepActuacionesProcesales
         'funcion de actuacion procesales
          'Dim rsAP As New ADODB.Recordset
          'Dim lsMensaje As String
          Set loRep = New COMNColocRec.NCOMColRecRConsulta
            Set rsAP = New ADODB.Recordset
            Set rsAP = loRep.nRepo138034_ListadoActuacionesProcesales(MskFecVen.Text, lsmensaje)
          Set loRep = Nothing
          If lsmensaje = "" Then
            If fgImprimeActuacionesProcesales(rsAP) = False Then
             MsgBox "No existen Datos para el Reporte", vbInformation, "AVISO"
             Exit Sub
            End If
          Else
            MsgBox "No existen Datos para el Reporte", vbInformation, "Aviso"
            Exit Sub
          End If
        Case gColRecRepActuacionesProcesales2 ' 138035 , DAOR 20070122
            Dim sCondiciones As String
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
              Set rsAP = New ADODB.Recordset
              Set rsAP = loRep.nRepo138035_ListadoActuacionesProcesales(mskPeriodo1Del, mskPeriodo1Al, IIf(optAbogado(0).value = True, 2, 1), Trim(txtAbogado.Text), lsmensaje)
            Set loRep = Nothing
            If lsmensaje = "" Then
            sCondiciones = "Del " & mskPeriodo1Del.Text & " Al " & mskPeriodo1Al.Text
              If fgImprimeActuacionesProcesales2(rsAP, sCondiciones) = False Then
                MsgBox "No existen Datos para el Reporte", vbInformation, "AVISO"
                Exit Sub
              End If
            Else
              MsgBox lsmensaje, vbInformation, "Aviso"
              Exit Sub
            End If
        Case gColRecRepCreditosCastigados '138036, DAOR 20070124 -- PEAC 20070922
            Screen.MousePointer = 11
            'Set loRep = New COMNColocRec.NCOMColRecRConsulta
            'loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsListaAgencias = DameAgencias
            'lsCadImp = loRep.nRepo138036_CreditosCastigados(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, Val(txtTipoCambio.Text), gsCodAge, lsListaAgencias, MatProductos, lsmensaje, gImpresora)
            
            'Call nRepo138036_CreditosCastigados(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, Val(txtTipoCambio.Text), gsCodAge, lsListaAgencias, MatProductos, matAnalista, lsmensaje, gImpresora)
             Call nRepo138036_CreditosCastigados(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, val(txtTipoCambio.Text), gsCodAge, MatAgencias, MatProductos, matAnalista, lsmensaje, gImpresora)
             If lsmensaje <> "" Then
                 MsgBox "No existen Datos para el Reporte", vbInformation, "Aviso"
                 Screen.MousePointer = 0
                 Exit Sub
             End If
            'MatAgencias
            
            Screen.MousePointer = 0
        Case gColRecRepCredJudToCastigar '138037, PEAC 20070912
            Screen.MousePointer = 11
            'Set loRep = New COMNColocRec.NCOMColRecRConsulta
            'loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsListaAgencias = DameAgencias
            If Me.OptBusqueda(3).value = True Then
                If bValorCredJud Then
                    LslistaCreditos = DameCreditos
                    bValorCredJud = False
                End If
                OptBusqueda.iTem(3).value = False
            End If
            
            'lsCadImp = loRep.nRepo138037_RepCredJudToCastigar(Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gsCodAge, lsListaAgencias, lsmensaje, gImpresora)
            'lsCadImp = loRep.nRepo138037_RepCredJudToCastigar(gsCodAge, gdFecSis, Me.TxtDiaAtrIni.Text, Me.TxtDiaAtrFin.Text, Val(txtTipoCambio.Text), lsListaAgencias)
            Call nRepo138037_RepCredJudToCastigar(gsCodAge, gdFecSis, Me.TxtDiaAtrIni.Text, Me.TxtDiaAtrFin.Text, val(txtTipoCambio.Text), lsListaAgencias, LslistaCreditos)
            Screen.MousePointer = 0
            
            If lsmensaje <> "" Then
                 MsgBox "No existen Datos para el Reporte", vbInformation, "Aviso"
                 Screen.MousePointer = 0
                 Exit Sub
             End If
            
        Case gColRecRepGastosJudCas '138038, PEAC 20071010
        
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            If OptJud.value = True Then
                nEstadoJud = 1
            Else
                nEstadoJud = 2
            End If
            If ChkSoles1.value = True Then
                nmoneda = 1
            Else
                nmoneda = 2
            End If
            
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsListaAgencias = DameAgencias
            
            
            If Me.chkExportaExcel.value = 1 Then
                Call Repo138038_RepGastosJudCas_Excel(gsCodAge, nEstadoJud, nmoneda, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gsNomAge, gdFecSis, gsCodUser, lsListaAgencias, MatProductos, lsmensaje)
            Else
                Screen.MousePointer = 11
                lscadimp = loRep.nRepo138038_RepGastosJudCas(gsCodAge, nEstadoJud, nmoneda, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gsNomAge, gdFecSis, gsCodUser, lsListaAgencias, MatProductos, lsmensaje, gImpresora)
                Screen.MousePointer = 0
            End If
           
        Case gColRecRepCredSaldosCondo '138039, PEAC 20071011
            
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsListaAgencias = DameAgencias
            
            Screen.MousePointer = 11
            lscadimp = loRep.nRepo138039_RepCredSaldosCondo(gsCodAge, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gsNomAge, gdFecSis, gsCodUser, lsListaAgencias, lsmensaje, gImpresora)
            Screen.MousePointer = 0
                        
        Case gColRecRepCarteraJudCas '138015, PEAC 20071023
        
            Set loRep = New COMNColocRec.NCOMColRecRConsulta
            If OptJud.value = True Then
                nEstadoJud = 1
            Else
                nEstadoJud = 2
            End If
            If ChkSoles1.value = True Then
                nmoneda = 1
            Else
                nmoneda = 2
            End If
            
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsListaAgencias = DameAgencias
            
            Screen.MousePointer = 11
            lscadimp = loRep.nRepo138015_RepCarteraJudCas(gsCodAge, nEstadoJud, nmoneda, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, gsNomAge, gdFecSis, gsCodUser, lsListaAgencias, MatProductos, lsmensaje, gImpresora)
            Screen.MousePointer = 0
                                   
    End Select
    If optOpcionImpresion(0).value = True Then
        lsDestino = "P"
    ElseIf optOpcionImpresion(1).value = True Then
        lsDestino = "A"
    End If
    
Set loRep = Nothing
    If nOperacion <> 138034 Then
        If Len(Trim(lscadimp)) > 0 And nOperacion <> 138026 Then
            Set loPrevio = New previo.clsprevio
                If lsDestino = "P" Then
                    loPrevio.Show lscadimp, sDescOperacion, True
                ElseIf lsDestino = "A" Then
                    frmImpresora.Show 1
                    loPrevio.PrintSpool sLpt, lscadimp, True
                End If
            Set loPrevio = Nothing
        Else
            If nOperacion <> 138026 Then
'                MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
            End If
        End If
    End If
End Sub

'MAVM 20090706 Reporte 138038
Public Sub Repo138038_RepGastosJudCas_Excel( _
ByVal psCodAge As String, _
ByVal pEstado As String, _
ByVal pMoneda As String, _
ByVal pdDel As Date, _
ByVal pdAl As Date, _
ByVal psNomAge As String, _
ByVal pdFecSis As Date, _
ByVal psCodUser As String, _
ByVal psListaAgencias As String, _
ByVal pMatProd As Variant, _
Optional ByRef psMensaje As String)

Dim rsDatos As ADODB.Recordset
Dim oDCred As COMDColocRec.DCOMColRecCredito
Dim R As ADODB.Recordset
Dim i As Integer, J As Integer
Dim vmone As String, vage As String, vana As String, vEstado As String, vMoneda As String, vproducto As String
Dim vCuenta As String, vNombre As String
Dim nTipoCambio As Currency
Dim lscadimp As String
Dim lnPage As Integer
Dim lnLineas As Integer
Dim lnTotPago As Double, lnTotCapi As Double, lnTotInte As Double
Dim lnTotMora As Double, lnTotGtos As Double

Dim lnMonto As Double, lnMontoPag As Double, lnSaldo As Double
Dim lnMonto1 As Double, lnMontoPag1 As Double, lnSaldo1 As Double
Dim lnMonto2 As Double, lnMontoPag2 As Double, lnSaldo2 As Double

'Dim xlAplicacion As Excel.Application
'Dim xlLibro As Excel.Workbook
'Dim xlHoja1 As Excel.Worksheet
'Dim liLineas As Integer, i As Integer
'Dim fs As Scripting.FileSystemObject

Set oDCred = New COMDColocRec.DCOMColRecCredito
Set R = oDCred.dObtienegRepGastosJudCas(psCodAge, pEstado, pMoneda, psListaAgencias, pdDel, pdAl, pMatProd)
Set oDCred = Nothing

If R.RecordCount = 0 Then
    MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
    Exit Sub
End If

Dim lMatCabecera As Variant
Dim lsNombreArchivo As String
lsNombreArchivo = "Report"
ReDim lMatCabecera(7, 2)

lMatCabecera(0, 0) = "ITEM": lMatCabecera(0, 1) = ""
lMatCabecera(1, 0) = "CONCEPTO": lMatCabecera(1, 1) = ""
lMatCabecera(2, 0) = "MONTO INGRESADO": lMatCabecera(2, 1) = ""
lMatCabecera(3, 0) = "FECHA INGRESADA": lMatCabecera(3, 1) = ""
lMatCabecera(4, 0) = "MONTO PAGADO": lMatCabecera(4, 1) = ""
lMatCabecera(5, 0) = "FECHA PAGADA": lMatCabecera(5, 1) = ""
lMatCabecera(6, 0) = "SALDO": lMatCabecera(6, 1) = ""
    
Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "REPORTE DE GASTOS DE CREDITOS (Vigente Juducial)", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

'lnPage = 1
'
'    lsCadImp = ""
'
'    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE GASTOS DE CREDITOS (" & R!Estado & ")", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 165, "", 138038)
'
'    lsCadImp = lsCadImp & "MONEDA : " & R!Moneda & oFunI.gPrnSaltoLinea
'    lsCadImp = lsCadImp & "ESTADO : " & R!Estado & oFunI.gPrnSaltoLinea
'
'    lnLineas = 7
'
'    Do While Not R.EOF
'        lsCadImp = lsCadImp & "AGENCIA  : " & R!Agencia & oFunI.gPrnSaltoLinea
'        vage = R!Agencia
'
'        lnMonto2 = 0
'        lnMontoPag2 = 0
'        lnSaldo2 = 0
'
'            Do While R!Agencia = vage
'                lsCadImp = lsCadImp & "   PRODUCTO : " & R!Producto & oFunI.gPrnSaltoLinea
'                vproducto = R!Producto
'
'                    Do While R!Producto = vproducto
'                    lsCadImp = lsCadImp & "CREDITO ANTIGUO: " & R!cCtaCodAnt & "CREDITO NUEVO: " & R!cCtaCod & "NOMBRES: " & R!cPersNombre & oFunI.gPrnSaltoLinea
'                    vCuenta = R!cCtaCod
'
'                        lnMonto = 0
'                        lnMontoPag = 0
'                        lnSaldo = 0
'
'                        j = 0
'
'                            Do While R!cCtaCod = vCuenta
'
'                                j = j + 1
'                                lnLineas = lnLineas + 1
'
'                                lsCadImp = lsCadImp & Space(5) & oFun.ImpreFormat(j, 4, 0)
'                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!cMotivoGasto, 20)
'                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!nMonto, 12, 2, True)
'                                lsCadImp = lsCadImp & oFun.ImpreFormat(Format(R!dAsigna, "dd/mm/yyyy"), 10)
'                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!nMontoPagado, 12, 2, True)
'                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Pago, 8)
'                                lsCadImp = lsCadImp & oFun.ImpreFormat(R!Saldo, 12, 2, True)
'                                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'                                 lnMonto = lnMonto + R!nMonto
'                                 lnMontoPag = lnMontoPag + R!nMontoPagado
'                                 lnSaldo = lnSaldo + R!Saldo
'
'                                If lnLineas >= 55 Then
'                                    lnPage = lnPage + 1
'                                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                                    lsCadImp = lsCadImp & nRepoCabecera(" REPORTE DE GASTOS DE CREDITOS (" & R!Estado & ")", " DEL " & Format(pdDel, "dd/mm/yyyy") & " AL " & Format(pdAl, "dd/mm/yyyy"), lnPage, 165, "", 138038)
'                                    lnLineas = 7
'                                End If
'
'                                R.MoveNext
'                                If R.EOF Then
'                                    Exit Do
'                                End If
'                            Loop
'
'                            lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
'                            lsCadImp = lsCadImp & Space(15) & "TOT. CUENTA :" & Space(14)
'                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnMonto, 12, 2, True)
'                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnMontoPag, 12, 2, True)
'                            lsCadImp = lsCadImp & oFun.ImpreFormat(lnSaldo, 12, 2, True)
'                            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'                            lnMonto1 = lnMonto1 + lnMonto
'                            lnMontoPag1 = lnMontoPag1 + lnMontoPag
'                            lnSaldo1 = lnSaldo1 + lnSaldo
'
'                    If R.EOF Then
'                        Exit Do
'                    End If
'                Loop
'
'                lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
'                lsCadImp = lsCadImp & Space(15) & "TOT. PRODUCTO :" & Space(14)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(lnMonto1, 12, 2, True)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(lnMontoPag1, 12, 2, True)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(lnSaldo1, 12, 2, True)
'                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'                lnMonto2 = lnMonto2 + lnMonto1
'                lnMontoPag2 = lnMontoPag2 + lnMontoPag1
'                lnSaldo2 = lnSaldo2 + lnSaldo1
'
'            If R.EOF Then
'                Exit Do
'            End If
'        Loop
'
'        lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & Space(15) & "TOT. AGENCIA  :" & Space(14)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(lnMonto2, 12, 2, True)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(lnMontoPag2, 12, 2, True)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(lnSaldo2, 12, 2, True) & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & String(165, "-") & oFunI.gPrnSaltoLinea
'
'    Loop

    R.Close
    Set R = Nothing
    'nRepo138038_RepGastosJudCas = lsCadImp
    'Set oFunI = Nothing

End Sub

'Private Sub GeneraReporte(prRs As ADODB.Recordset)
'    Dim xlAplicacion As Excel.Application
'    Dim xlLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim liLineas As Integer, i As Integer
'    Dim fs As Scripting.FileSystemObject
'    Dim lnNumColumns As Integer
'
'    Dim j As Integer
'
'    Dim lnSI As Currency '(6) Saldo Inicial
'    Dim lnVH As Currency '(11) Valor Historico
'    Dim lnVA As Currency '(13) Valor Ajustado
'    Dim lnDACEA As Currency '(19) Depre Acum Cier Ejer Ant
'    Dim lnDE As Currency '(20) Depre Ejerc
'
'    Dim lnDAH As Currency '(23) Depre Acum Hist
'    Dim lnDAAI As Currency '(25) Depre Acum Ajust Inflac
'
'    i = 8
'    'prRs.MoveFirst
'    While Not prRs.EOF
'        i = i + 1
'        For j = 0 To prRs.Fields.Count - 1
'
'            If IsNumeric(prRs.Fields(j)) And (j = 6 Or j = 11 Or j = 13 Or j = 23 Or j = 25) Then
'                xlHoja1.Cells(i + 1, j + 1) = Format(prRs.Fields(j), "#,##0.00")
'                Select Case j
'                Case 6
'                    lnSI = lnSI + CCur(prRs.Fields(j))
'                Case 11
'                    lnVH = lnVH + CCur(prRs.Fields(j))
'                Case 13
'                    lnVA = lnVA + CCur(prRs.Fields(j))
'                Case 23
'                    lnDAH = lnDAH + CCur(prRs.Fields(j))
'                Case 25
'                    lnDAAI = lnDAAI + CCur(prRs.Fields(j))
'                End Select
'            Else
'                If j = 19 Then
'                    If prRs!nBSPerDeprecia <> prRs!PeriodosDeprecia Then
'                        xlHoja1.Cells(i + 1, j + 1) = Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * DepreAcumuladaEjeAnt(prRs!FAdquisicion, prRs!nBSPerDeprecia), 2), "#,##0.00")
'                    Else
'                        xlHoja1.Cells(i + 1, j + 1) = Format(prRs!nBSValor, "#,##0.00")
'                    End If
'                    lnDACEA = lnDACEA + CCur(xlHoja1.Cells(i + 1, j + 1))
'                Else
'                    If j = 20 Then
'                        If prRs!nBSPerDeprecia <> prRs!PeriodosDeprecia Then
'                            xlHoja1.Cells(i + 1, j + 1) = Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * DepreEjercicio, 2), "#,##0.00")
'                            lnDE = lnDE + CCur(Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * DepreEjercicio, 2), "#,##0.00"))
'                        End If
'                    Else
'                        xlHoja1.Cells(i + 1, j + 1) = prRs.Fields(j)
'                    End If
'                End If
'            End If
'        Next j
'        prRs.MoveNext
'    Wend
'
'    xlHoja1.Cells(prRs.RecordCount + 10, 6) = "TOTALES"
'    xlHoja1.Cells(prRs.RecordCount + 10, 7) = Format(lnSI, "#,##0.00")
'    xlHoja1.Cells(prRs.RecordCount + 10, 12) = Format(lnVH, "#,##0.00")
'    xlHoja1.Cells(prRs.RecordCount + 10, 14) = Format(lnVA, "#,##0.00")
'
'    xlHoja1.Cells(prRs.RecordCount + 10, 20) = Format(lnDACEA, "#,##0.00")
'    xlHoja1.Cells(prRs.RecordCount + 10, 21) = Format(lnDE, "#,##0.00")
'    xlHoja1.Cells(prRs.RecordCount + 10, 24) = Format(lnDAH, "#,##0.00")
'    xlHoja1.Cells(prRs.RecordCount + 10, 26) = Format(lnDAAI, "#,##0.00")
'
'    xlHoja1.Range("B10:B" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
'    xlHoja1.Range("O10:O" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
'    xlHoja1.Range("P10:P" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
'    xlHoja1.Range("Q10:Q" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
'    xlHoja1.Range("S10:S" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
'    xlHoja1.Range("V10:V" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
'
'    'Border's Tabla
'    xlHoja1.Range("A10:Z" & prRs.RecordCount + 9).BorderAround xlContinuous, xlMedium
'    xlHoja1.Range("A10:Z" & prRs.RecordCount + 9).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    xlHoja1.Range("A10:Z" & prRs.RecordCount + 9).Borders(xlInsideVertical).LineStyle = xlContinuous
'
'    'Border's Totales
'    xlHoja1.Range("G" & prRs.RecordCount + 10 & ":Z" & prRs.RecordCount + 10).BorderAround xlContinuous, xlMedium
'    xlHoja1.Range("G" & prRs.RecordCount + 10 & ":Z" & prRs.RecordCount + 10).Borders(xlInsideVertical).LineStyle = xlContinuous
'
'    xlHoja1.Range("J8:J9").Cells.VerticalAlignment = xlJustify
'
'     xlHoja1.Range("L8:L9").Cells.VerticalAlignment = xlJustify
'
'    xlHoja1.Range("H8:H9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("N8:N9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("M8:M9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("O8:O9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("P8:P9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("Q9:Q9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("R9:R9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("V8:V9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("S8:S9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("T8:T9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("U8:U9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("W8:W9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("X8:X9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("Y8:Y9").Cells.VerticalAlignment = xlJustify
'    xlHoja1.Range("Z8:Z9").Cells.VerticalAlignment = xlJustify
'End Sub

Private Sub tvwOperacion_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim nodOpe As Node
    Dim sDesc As String
    Set nodOpe = tvwOperacion.SelectedItem
    sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
    EjecutaOperacion CLng(nodOpe.Tag), sDesc
    Set nodOpe = Nothing
End Sub

Private Sub txtAbogado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optOpcionImpresion(0).SetFocus
    End If
End Sub

Private Sub txtAnalista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optOpcionImpresion(0).SetFocus
    End If
End Sub

Private Sub txtTop_GotFocus()
    fEnfoque Me.txtTop
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
If KeyAscii <> 45 Then
    KeyAscii = NumerosEnteros(KeyAscii)
End If
If KeyAscii = 13 Then
    chkEstadoRecup(0).SetFocus
End If
End Sub

Private Function DameAgencias() As String
Dim Agencias As String
Dim lnAge As Integer
Dim est As Integer
est = 0
Agencias = ""
ReDim MatAgencias(0)
For lnAge = 1 To frmSelectAgencias.List1.ListCount
 If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
    est = est + 1
    ReDim Preserve MatAgencias(est)
    If est = 1 Then
        Agencias = "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
        MatAgencias(0) = Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2)
    Else
        Agencias = Agencias & ", " & "'" & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2) & "'"
        MatAgencias(est - 1) = Mid(frmSelectAgencias.List1.List(lnAge), 1, 2)
    End If
 End If
Next lnAge
DameAgencias = Agencias
End Function
'MADM 20110515
Private Function DameCreditos() As String
Dim sCreditos As String
Dim iCre As Integer
sCreditos = ""

  If UBound(MatCreditos) > 0 Then
     For iCre = 0 To UBound(MatCreditos) - 1
            If UBound(MatCreditos) = 1 Then
                sCreditos = "'" & MatCreditos(iCre) & "'"
            Else
                sCreditos = sCreditos & "'" & MatCreditos(iCre) & "'" & ","
            End If
        Next iCre
   End If
    DameCreditos = sCreditos
    
End Function

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
Public Function DescProdConsumoSeleccionado() As String
Dim lsProductos As String
Dim i As Integer
lsProductos = "PRODUCTOS : "
   
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" And Mid(TreeView1.Nodes(i).Key, 2, 1) = Mid(Producto.gColConsuDctoPlan, 1, 1) Then
                lsProductos = lsProductos & "/CON-" & Mid(TreeView1.Nodes(i).Text, 1, 3)
            End If
        End If
    Next
  
DescProdConsumoSeleccionado = lsProductos

End Function

Public Function ValorProdConsumo() As String
Dim i As Integer
Dim lsCad As String

    lsCad = ""
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
        
    If Len(lsCad) > 0 Then
        ValorProdConsumo = " AND substring(Credito.cCtaCod,6,3) IN (" & lsCad & ") "
    Else
        ValorProdConsumo = " AND substring(Credito.cCtaCod,6,1) = '3' "
    End If
End Function

Private Function DescProdSeleccionado() As String
Dim lsProductos As String
Dim i As Integer
lsProductos = "PRODUCTOS : "
     
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Checked = True Then
            If Mid(TreeView1.Nodes(i).Key, 1, 1) = "H" Then
                lsProductos = lsProductos & "/MES" & Mid(TreeView1.Nodes(i).Text, 1, 3)
            End If
        End If
    Next
  
DescProdSeleccionado = lsProductos
End Function


'** PEAC 20071003 'MADM 20110505 - SBS
Public Function nRepo138037_RepCredJudToCastigar( _
ByVal psCodAge As String, _
ByVal gdFecSis As Date, _
ByVal pnAtrIni As Integer, _
ByVal pnAtrFin As Integer, _
ByVal nTipoCambio As Currency, _
ByVal psListaAgencias As String, Optional ByVal psListaCreditos As String = "") As String

Dim oDCred As COMDColocRec.DCOMColRecCredito
Dim oFun As New COMFunciones.FCOMImpresion
Dim R As ADODB.Recordset
Dim i As Integer, J As Integer
Dim vmone As String, vage As String
'Dim nTipoCambio As Currency
'nTipoCambio = 3.161
    
    Set oDCred = New COMDColocRec.DCOMColRecCredito
    Set R = oDCred.dObtienegRepCredJudToCastigar(psCodAge, gdFecSis, pnAtrIni, pnAtrFin, psListaAgencias, psListaCreditos)
    Set oDCred = Nothing

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "REPORTE DE CREDITOS JUDICIALES PARA PASE A CASTIGO AL " & Format(gdFecSis, "dd/MM/YYYY")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B2", "Z2").MergeCells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).HorizontalAlignment = 3
    
    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & oFun.ImpreFormat(nTipoCambio, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "CREDITO"
    ApExcel.Cells(8, 4).Formula = "ESTADO"
    ApExcel.Cells(8, 5).Formula = "ANALISTA"
    ApExcel.Cells(8, 6).Formula = "LINEA"
    ApExcel.Cells(7, 7).Formula = "CODIGO" '**
    ApExcel.Cells(8, 7).Formula = "CLIENTE"
    ApExcel.Cells(8, 8).Formula = "CLIENTE"
    ApExcel.Cells(8, 9).Formula = "DOC. ID."
    ApExcel.Cells(8, 10).Formula = "DIRECCION"
    ApExcel.Cells(8, 11).Formula = "COD. S.B.S."
    ApExcel.Cells(7, 12).Formula = "FECHA" '**
    ApExcel.Cells(8, 12).Formula = "VIGENCIA"
    ApExcel.Cells(8, 13).Formula = "ATRASO"
    ApExcel.Cells(7, 14).Formula = "MONTO" '**
    ApExcel.Cells(8, 14).Formula = "DESEMBOLSADO"
    ApExcel.Cells(7, 15).Formula = "MONTO PROV." '**
    ApExcel.Cells(8, 15).Formula = "CONSTITUIDA"
    ApExcel.Cells(7, 16).Formula = "SALDO" '**
    ApExcel.Cells(8, 16).Formula = "CAPITAL"
    ApExcel.Cells(8, 17).Formula = "INTERES"
    ApExcel.Cells(8, 18).Formula = "MORA"
    ApExcel.Cells(8, 19).Formula = "GASTOS"
    ApExcel.Cells(8, 20).Formula = "TOTAL"
    ApExcel.Cells(8, 21).Formula = "TOTAL MN"
    ApExcel.Cells(8, 22).Formula = "CALIFICACION"
    ApExcel.Cells(8, 23).Formula = "DEMANDA"
    ApExcel.Cells(8, 24).Formula = "TASAINTCOMP"
    ApExcel.Cells(8, 25).Formula = "DiasTransUltPgo"
    ApExcel.Cells(8, 26).Formula = "CODAGE"
    
    ApExcel.Range("B7", "Z8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "Z8").Font.Bold = True
    ApExcel.Range("B7", "Z8").HorizontalAlignment = 3

    i = 8

    Do While Not R.EOF
    i = i + 1
    
        If psListaCreditos = "" Then
            ApExcel.Cells(i, 2).Formula = "AGENCIA : " & R!NomAge
            ApExcel.Cells(i, 2).Font.Bold = True
            vage = R!CODAge
                J = 0
               Do While R!CODAge = vage
                i = i + 1
                ApExcel.Cells(i, 2).Formula = "MONEDA : " & R!cMoneda
                ApExcel.Cells(i, 2).Font.Bold = True
                vmone = R!cMoney
                   
                    Do While R!cMoney = vmone And R!CODAge = vage
                        J = J + 1
                        i = i + 1
                        ApExcel.Cells(i, 2).Formula = J
                        ApExcel.Cells(i, 3).Formula = "'" & R!Credito
                        ApExcel.Cells(i, 4).Formula = R!Estado
                        ApExcel.Cells(i, 5).Formula = R!Analista
                        ApExcel.Cells(i, 6).Formula = R!Linea
                        ApExcel.Cells(i, 7).Formula = "'" & R!CodCliente
                        ApExcel.Cells(i, 8).Formula = R!Cliente
                        ApExcel.Cells(i, 9).Formula = R!Dni
                        ApExcel.Cells(i, 10).Formula = R!Direccion
                        ApExcel.Cells(i, 11).Formula = R!CodSbs
                        ApExcel.Cells(i, 12).Formula = Format(R!FechaVige, "mm/dd/yyyy")
                        ApExcel.Cells(i, 13).Formula = R!Atraso
                        ApExcel.Cells(i, 14).Formula = R!MontoDesem
                        ApExcel.Cells(i, 15).Formula = R!ProviConst
                        ApExcel.Cells(i, 16).Formula = R!SaldoCap
                        ApExcel.Cells(i, 17).Formula = R!IntCom + ((((1 + (R!TasaIntComp / 100)) ^ (R!nDiasTransUltPgo / 30)) - 1) * R!SaldoCap)
                        ApExcel.Cells(i, 18).Formula = R!IntMor + ((((1 + (R!TasaIntMora / 100)) ^ (R!nDiasTransUltPgo / 30)) - 1) * R!SaldoCap)
                        ApExcel.Cells(i, 19).Formula = R!Gastos
                        ApExcel.Cells(i, 20).Formula = "=+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
                        ApExcel.Cells(i, 21).Formula = IIf(R!cMoney = "2", "=+RC[-1]*" & nTipoCambio, 0)
                        ApExcel.Cells(i, 22).Formula = "'" & R!Calificacion
                        ApExcel.Cells(i, 23).Formula = "'" & R!nDemanda 'MADM 20110504
                        ApExcel.Cells(i, 24).Formula = "'" & R!TasaIntComp 'MADM 20110504
                        ApExcel.Cells(i, 25).Formula = "'" & R!nDiasTransUltPgo 'MADM 20110504
                        ApExcel.Cells(i, 26).Formula = "'" & R!CODAge 'MADM 20110504
                        
                        ApExcel.Range("N" & Trim(str(i)) & ":" & "U" & Trim(str(i))).NumberFormat = "#,##0.00"
                        ApExcel.Range("B" & Trim(str(i)) & ":" & "Z" & Trim(str(i))).Borders.LineStyle = 1
                        
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                                       
                    Loop
                    
                    i = i + 1
                    ApExcel.Cells(i, 14).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 15).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 16).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 17).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 18).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 19).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 20).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    ApExcel.Cells(i, 21).Formula = "=SUM(R[-" & Trim(str(Int(J))) & "]C:R[-1]C)"
                    i = i + 1
                    
                    If R.EOF Then
                        Exit Do
                    End If
                    
                Loop
        Else
                        J = J + 1
                        ApExcel.Cells(i, 2).Formula = J
                        ApExcel.Cells(i, 3).Formula = "'" & R!Credito
                        ApExcel.Cells(i, 4).Formula = R!Estado
                        ApExcel.Cells(i, 5).Formula = R!Analista
                        ApExcel.Cells(i, 6).Formula = R!Linea
                        ApExcel.Cells(i, 7).Formula = "'" & R!CodCliente
                        ApExcel.Cells(i, 8).Formula = R!Cliente
                        ApExcel.Cells(i, 9).Formula = R!Dni
                        ApExcel.Cells(i, 10).Formula = R!Direccion
                        ApExcel.Cells(i, 11).Formula = R!CodSbs
                        ApExcel.Cells(i, 12).Formula = Format(R!FechaVige, "mm/dd/yyyy")
                        ApExcel.Cells(i, 13).Formula = R!Atraso
                        ApExcel.Cells(i, 14).Formula = R!MontoDesem
                        ApExcel.Cells(i, 15).Formula = R!ProviConst
                        ApExcel.Cells(i, 16).Formula = R!SaldoCap
                        ApExcel.Cells(i, 17).Formula = R!IntCom + ((((1 + (R!TasaIntComp / 100)) ^ (R!nDiasTransUltPgo / 30)) - 1) * R!SaldoCap)
                        ApExcel.Cells(i, 18).Formula = R!IntMor + ((((1 + (R!TasaIntMora / 100)) ^ (R!nDiasTransUltPgo / 30)) - 1) * R!SaldoCap)
                        ApExcel.Cells(i, 19).Formula = R!Gastos
                        ApExcel.Cells(i, 20).Formula = "=+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
                        ApExcel.Cells(i, 21).Formula = IIf(R!cMoney = "2", "=+RC[-1]*" & nTipoCambio, 0)
                        ApExcel.Cells(i, 22).Formula = "'" & R!Calificacion
                        ApExcel.Cells(i, 23).Formula = "'" & R!nDemanda 'MADM 20110504
                        ApExcel.Cells(i, 24).Formula = "'" & R!TasaIntComp 'MADM 20110504
                        ApExcel.Cells(i, 25).Formula = "'" & R!nDiasTransUltPgo 'MADM 20110504
                        ApExcel.Cells(i, 26).Formula = "'" & R!CODAge 'MADM 20110504
                        
                        ApExcel.Range("N" & Trim(str(i)) & ":" & "U" & Trim(str(i))).NumberFormat = "#,##0.00"
                        ApExcel.Range("B" & Trim(str(i)) & ":" & "Z" & Trim(str(i))).Borders.LineStyle = 1
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
            End If
    Loop
'    If psListaCreditos <> "" Then
'    ApExcel.Cells(i, 14).Formula = "=SUM(N9:N" & (i - 1) & ")"
'    ApExcel.Cells(i, 15).Formula = "=SUM(O9:O" & (i - 1) & ")"
'    ApExcel.Cells(i, 16).Formula = "=SUM(P9:P" & (i - 1) & ")"
'    ApExcel.Cells(i, 17).Formula = "=SUM(Q9:Q" & (i - 1) & ")"
'    ApExcel.Cells(i, 18).Formula = "=SUM(R9:R" & (i - 1) & ")"
'    ApExcel.Cells(i, 19).Formula = "=SUM(S9:S" & (i - 1) & ")"
'    ApExcel.Cells(i, 20).Formula = "=SUM(T9:T" & (i - 1) & ")"
'    ApExcel.Cells(i, 21).Formula = "=SUM(U9:U" & (i - 1) & ")"
'    End If
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    ApExcel.Visible = True
    Set ApExcel = Nothing

    'Set oFunI = Nothing

End Function

'**DAOR 20070124, Reporte de créditos castigados 'modificado por peac 20071019
Public Function nRepo138036_CreditosCastigados( _
ByVal pdFecCastDe As Date, pdFecCastHasta As Date, ByVal nTipoCambio As Currency, _
psCodAge As String, ByVal pMatAgencias As Variant, ByVal pMatProd As Variant, _
ByVal pMatAnalistas As Variant, Optional ByRef psMensaje As String, _
Optional ByVal psImpresora As Impresoras = gEPSON) As String

'*** PEAC 20080707 se reeamplazo el parametro "psListaAgencias" por "pMatAgencias"

Dim lsSQL As String
Dim lrDataRep As New ADODB.Recordset
Dim lscadimp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim oFun As New COMFunciones.FCOMImpresion
Dim P As String
Dim oConec As New COMConecta.DCOMConecta
Dim whereAgencias As String
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.inicia psImpresora
Dim nTotalD As Double, nTotalS As Double
Dim nTotalCas As Integer, i As Integer, J As Integer
Dim sTempMoneda As String
Dim sCadProd As String, sCadAnalista As String, sCadAge As String
Dim vmone As String
Dim vage As String

'*** PEAC 20080707
'-----------------------------------------------------------------------------------
    sCadProd = ""
    For i = 0 To UBound(pMatProd) - 1
        sCadProd = sCadProd & pMatProd(i) & ","
    Next i
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 1)
    
    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

    
    sCadAnalista = ""
    If Me.chkTodosLosCastigados.value <> 1 Then  '' *** PEAC 20090126
        For i = 0 To UBound(pMatAnalistas) - 1
            sCadAnalista = sCadAnalista & pMatAnalistas(i) & ","
        Next i
        sCadAnalista = Mid(sCadAnalista, 1, Len(sCadAnalista) - 1)
    End If
    
   lsSQL = " exec stp_sel_ObtieneCreditosCastigados '" & Format(pdFecCastDe, "yyyymmdd") & "','" & Format(pdFecCastHasta, "yyyymmdd") & "','" & sCadAge & "','" & sCadAnalista & "','" & sCadProd & "'," & nTipoCambio

    oConec.AbreConexion
    Set lrDataRep = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
        
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psMensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
'---------------------------****************************** inicia excel

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    ApExcel.Cells(2, 2).Formula = "REPORTE DE CREDITOS CASTIGADOS DEL " & Format(pdFecCastDe, "dd/MM/YYYY") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy")
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B2", "R2").MergeCells = True
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(2, 2).HorizontalAlignment = 3
    
    ApExcel.Cells(5, 2).Formula = "TIPO CAMBIO : " & oFun.ImpreFormat(nTipoCambio, 5, 3)
    ApExcel.Cells(5, 2).Font.Bold = True
    
    ApExcel.Cells(8, 2).Formula = "Nº"
    ApExcel.Cells(8, 3).Formula = "CREDITO"
    ApExcel.Cells(7, 4).Formula = "CUENTA" '**
    ApExcel.Cells(8, 4).Formula = "ANTIGUA"
    ApExcel.Cells(7, 5).Formula = "CODIGO" '**
    ApExcel.Cells(8, 5).Formula = "CLIENTE"
    ApExcel.Cells(8, 6).Formula = "NOMBRE"
    ApExcel.Cells(8, 7).Formula = "DIRECCION"
    ApExcel.Cells(8, 8).Formula = "DOCUMENTO"
    
    ApExcel.Cells(8, 9).Formula = "PERS.RELACION"
    ApExcel.Cells(8, 10).Formula = "NOMBRE PERS.RELACIONADA"
    ApExcel.Cells(8, 11).Formula = "DIRECCION PERS.RELACIONADA"
    
    ApExcel.Cells(7, 12).Formula = "FECHA"
    ApExcel.Cells(8, 12).Formula = "CASTIGO"
    ApExcel.Cells(7, 13).Formula = "MONTO"
    ApExcel.Cells(8, 13).Formula = "COLOCADO"
    ApExcel.Cells(7, 14).Formula = "MONTO"
    ApExcel.Cells(8, 14).Formula = "PAGADO"
    ApExcel.Cells(7, 15).Formula = "SALDO"
    ApExcel.Cells(8, 15).Formula = "CAPITAL"
    ApExcel.Cells(8, 16).Formula = "INTERES"
    ApExcel.Cells(8, 17).Formula = "MORA"
    ApExcel.Cells(8, 18).Formula = "GASTOS"
    ApExcel.Cells(8, 19).Formula = "TOTAL"
    ApExcel.Cells(8, 20).Formula = "TOTAL MN"
    ApExcel.Cells(7, 21).Formula = "CAPITAL"
    ApExcel.Cells(8, 21).Formula = "CASTIGADO"
    ApExcel.Cells(7, 22).Formula = "INTERES"
    ApExcel.Cells(8, 22).Formula = "CASTIGADO"
    ApExcel.Cells(7, 23).Formula = "MORA"
    ApExcel.Cells(8, 23).Formula = "CASTIGADO"
    ApExcel.Cells(7, 24).Formula = "GASTO"
    ApExcel.Cells(8, 24).Formula = "CASTIGADO"
    ApExcel.Cells(7, 25).Formula = "TOTAL"
    ApExcel.Cells(8, 25).Formula = "CASTIGADO"
    ApExcel.Cells(7, 26).Formula = "TOTAL"
    ApExcel.Cells(8, 26).Formula = "CAST. MN"
    ApExcel.Cells(8, 27).Formula = "ANALISTA"
    ApExcel.Cells(7, 28).Formula = "DIAS"
    ApExcel.Cells(8, 28).Formula = "ATRASO"
    ApExcel.Cells(8, 29).Formula = "AGENCIA"
    
    ApExcel.Cells(7, 30).Formula = "FECHA" 'MAVM 06042009
    ApExcel.Cells(8, 30).Formula = "FALLECIMIENTO" 'MAVM 06042009

    ApExcel.Range("B7", "AD8").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B7", "AD8").Font.Bold = True
    ApExcel.Range("B7", "AD8").HorizontalAlignment = 3

    i = 8
    Do While Not lrDataRep.EOF
        i = i + 1
        'ApExcel.Cells(i, 2).Formula = "MONEDA : " & lrDataRep!Money
        ApExcel.Cells(i, 2).Formula = "AGENCIA : " & lrDataRep!NombAgencia
        ApExcel.Cells(i, 2).Font.Bold = True
        vage = lrDataRep!cAgeCodAct
        
            J = 0
            Do While lrDataRep!cAgeCodAct = vage
            i = i + 1
            ApExcel.Cells(i, 2).Formula = "MONEDA : " & lrDataRep!Money
            ApExcel.Cells(i, 2).Font.Bold = True
            vmone = lrDataRep!Moneda
            i = i + 1
                Do While lrDataRep!Moneda = vmone And lrDataRep!cAgeCodAct = vage
                        If ApExcel.Cells(i - 1, 3).Formula <> lrDataRep!cCtaCod Then
                            J = J + 1
                            ApExcel.Cells(i, 2).Formula = J
                            ApExcel.Cells(i, 3).Formula = "'" & lrDataRep!cCtaCod
                            ApExcel.Cells(i, 4).Formula = "'" & lrDataRep!CtaAnt
                            ApExcel.Cells(i, 5).Formula = lrDataRep!cPersCod
                            ApExcel.Cells(i, 6).Formula = lrDataRep!Cliente
                            ApExcel.Cells(i, 7).Formula = lrDataRep!Direccion
                            ApExcel.Cells(i, 8).Formula = "'" & lrDataRep!DocID
                            
                            ApExcel.Cells(i, 9).Formula = "'" & Trim(lrDataRep!PersRel)
                            ApExcel.Cells(i, 10).Formula = "'" & Trim(lrDataRep!PersRelNombre)
                            ApExcel.Cells(i, 11).Formula = "'" & Trim(lrDataRep!PersRelDire)
                            
                            ApExcel.Cells(i, 12).Formula = Format(lrDataRep!Fecha, "mm/dd/yyyy")
                            ApExcel.Cells(i, 13).Formula = lrDataRep!nMontoCol
                            ApExcel.Cells(i, 14).Formula = lrDataRep!nMontoCol - lrDataRep!nSaldo
                            ApExcel.Cells(i, 15).Formula = lrDataRep!nSaldo
                            ApExcel.Cells(i, 16).Formula = lrDataRep!Interes
                            ApExcel.Cells(i, 17).Formula = lrDataRep!Mora
                            ApExcel.Cells(i, 18).Formula = lrDataRep!Gastos
                            ApExcel.Cells(i, 19).Formula = lrDataRep!TotalDCast
                            'ApExcel.Cells(i, 17).Formula = lrDataRep!TotalACastMN
                            ApExcel.Cells(i, 20).Formula = lrDataRep!TotalDCastMN
                            ApExcel.Cells(i, 21).Formula = lrDataRep!CapiCas
                            ApExcel.Cells(i, 22).Formula = lrDataRep!InteCas
                            ApExcel.Cells(i, 23).Formula = lrDataRep!MoraCas
                            ApExcel.Cells(i, 24).Formula = lrDataRep!GastCas
                            ApExcel.Cells(i, 25).Formula = lrDataRep!TotalACast
                            ApExcel.Cells(i, 26).Formula = lrDataRep!TotalACastMN
                            ApExcel.Cells(i, 27).Formula = UCase(lrDataRep!Analista)
                            ApExcel.Cells(i, 28).Formula = lrDataRep!nDiasAtraso
                            ApExcel.Cells(i, 29).Formula = lrDataRep!NombAgencia
                            
                            'MAVM 06042009
                            ApExcel.Cells(i, 30).Formula = IIf(IsNull(lrDataRep!dPersFallec), "", IIf(lrDataRep!dPersFallec = "", "", Format(lrDataRep!dPersFallec, "mm/dd/yyyy")))
                                            
                            ApExcel.Range("M" & Trim(str(i)) & ":" & "Z" & Trim(str(i))).NumberFormat = "#,##0.00"
            '                ApExcel.RANGE("B" & Trim(Str(I)) & ":" & "R" & Trim(Str(I))).borders.linestyle = 1
        
                        Else
                            ApExcel.Cells(i, 3).Formula = "'" & lrDataRep!cCtaCod
                            ApExcel.Cells(i, 9).Formula = "'" & Trim(lrDataRep!PersRel)
                            ApExcel.Cells(i, 10).Formula = "'" & Trim(lrDataRep!PersRelNombre)
                            ApExcel.Cells(i, 11).Formula = "'" & Trim(lrDataRep!PersRelDire)
                        End If
                       i = i + 1
                       lrDataRep.MoveNext
                       If lrDataRep.EOF Then
                            Exit Do
                       End If
                       
                    Loop
                      If lrDataRep.EOF Then
                            Exit Do
                       End If
                     
        Loop
    Loop
    
    lrDataRep.Close
    Set lrDataRep = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    ApExcel.Visible = True
    Set ApExcel = Nothing
       
    End If
    'nRepo138036_CreditosCastigados = lsCadBuffer
Set oFunI = Nothing
      
       
'----------------------------****************************** finaliza excel
'        lnLineas = 0
'        lnPage = 1
'        sTempMoneda = "0"
'        nTotalCas = 0
'        '**Configuración de página**********************************************************
'        'lsCadImp = lsCadImp & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
'        'lsCadImp = lsCadImp & Chr$(27) & Chr$(67) & Chr$(66)   'Longitud de página a 22 líneas'
'        lsCadImp = lsCadImp & oFunI.gPrnTamLetra10CPI    'Tamaño 10 cpi
'        'lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
'        'lsCadImp = lsCadImp & Chr$(27) + Chr$(18)  ' cancela condensada
'        lsCadImp = lsCadImp & oFunI.gPrnBoldOFF  ' desactiva negrita
'        '***********************************************************************************
'
        
'        lsCadImp = lsCadImp & nRepoCabecera(" LISTADO DE CRÉDITOS CASTIGADOS ", " DEL " & Format(pdFecCastDe, "dd/mm/yyyy") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy"), lnPage, 150, "", 138036)

'
'        lnIndice = 0:  lnLineas = 7
'
'        With lrDataRep
'            Do While Not lrDataRep.EOF
'                lnIndice = lnIndice + 1
'                lnLineas = lnLineas + 1
'                If sTempMoneda <> !Moneda And sTempMoneda <> "0" Then
'                    lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
'                    lsCadImp = lsCadImp & oFun.ImpreFormat(">TOTAL " & IIf(sTempMoneda = "1", "SOLES", "DOLARES"), 20, 0)
'                    lsCadImp = lsCadImp & Space(55) & oFun.ImpreFormat(nTotalD, 12, 2, True) & Space(1)
'                    lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalD - nTotalS, 12, 2, True) & Space(1)
'                    lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalS, 12, 2, True) & oFunI.gPrnSaltoLinea
'                    lsCadImp = lsCadImp & ">Total de Creditos Castigados en  " & IIf(sTempMoneda = "1", "SOLES", "DOLARES") & ": " & nTotalCas & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'                End If
'                If sTempMoneda <> !Moneda Then
'                    sTempMoneda = !Moneda
'                    lsCadImp = lsCadImp & ">MONEDA : " & IIf(sTempMoneda = "1", "SOLES", "DOLARES") & oFunI.gPrnSaltoLinea
'                    lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
'                    nTotalD = 0
'                    nTotalS = 0
'                    nTotalCas = 0
'                End If
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!cCtaCod, 18, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!CtaAnt, 18, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!cPersCod, 13, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Cliente, 30, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Direccion, 30, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!DocID, 15, 0) & Space(1)
'                lsCadImp = lsCadImp & !Fecha & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nMontoCol, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nMontoCol - !nSaldo, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nSaldo, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Interes, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Mora, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Gastos, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalDCast, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalDCastMN, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!CapiCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!InteCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!MoraCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!GastCas, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalACast, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!TotalACastMN, 12, 2, True) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!Analista, 4, 0) & Space(1)
'                lsCadImp = lsCadImp & oFun.ImpreFormat(!nDiasAtraso, 8, 0) & Space(1)
'                If psListaAgencias <> "" Then
'                    lsCadImp = lsCadImp & Space(1) & oFun.ImpreFormat(!NombAgencia, 12, 0)
'                End If
'                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'                nTotalD = nTotalD + !nMontoCol
'                nTotalS = nTotalS + !nSaldo
'                nTotalCas = nTotalCas + 1
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera(" LISTADO DE CRÉDITOS CASTIGADOS ", " DEL " & Format(pdFecCastDe, "dd/mm/yyyy") & " AL " & Format(pdFecCastHasta, "dd/mm/yyyy"), lnPage, 150, "", 138036)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop
'        End With
'        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & oFun.ImpreFormat(">TOTAL " & IIf(sTempMoneda = "1", "SOLES", "DOLARES"), 20, 0)
'        lsCadImp = lsCadImp & Space(55) & oFun.ImpreFormat(nTotalD, 12, 2, True) & Space(1)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalD - nTotalS, 12, 2, True) & Space(1)
'        lsCadImp = lsCadImp & oFun.ImpreFormat(nTotalS, 12, 2, True) & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & ">Total de Creditos Castigados en  " & IIf(sTempMoneda = "1", "SOLES", "DOLARES") & ": " & nTotalCas & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
'
'
'        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & "Total de Creditos Castigados: " & oFun.ImpreFormat(lnIndice, 5, 0)
'        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
'
'    End If
'    nRepo138036_CreditosCastigados = lsCadBuffer
'Set oFunI = Nothing
End Function

'*** PEAC 20080724
Private Sub MostrarPagosPorFechas(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psEstados As String, ByVal pnMoneda As Byte, Optional ByVal psTipoC As String = "")
'(ByVal pdFecIni As String, ByVal pdFecFin As String, ByVal psEstados As String, ByVal pnMoneda As Byte, Optional ByVal psTipoC As String = "", Optional ByVal psImpresora As Impresoras = gEPSON) As String
'Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, sEstados, sMoneda, sTipoC

Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera(1, 1) As Variant
Dim lsmensaje As String
Dim i As Integer
Dim sCadAge As String
Dim lsSQL As String
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim sEstados As String
Dim cMoneda As String
Dim cTipo As String

    If Len(pnMoneda) = 0 Then
        MsgBox "Seleccione por lo menos una moneda.", vbInformation, "Atención"
        Screen.MousePointer = 0
        Exit Sub
    End If

    Screen.MousePointer = 11
    
'    sCadAge = ""
'    For i = 0 To UBound(pMatAgencias) - 1
'    sCadAge = sCadAge & pMatAgencias(i) & ","
'    Next i
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)


    If pnMoneda = 3 Then
        cMoneda = "%"
    Else
        cMoneda = pnMoneda
    End If

    sEstados = Replace(psEstados, "'", "")

    If Trim(psTipoC) <> "" Then
        cTipo = Replace(Replace(Replace(Trim(psTipoC), "(", ""), ")", ""), "'", "")
        lsSQL = "exec stp_sel_ListadoPagosPorFechas '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sEstados & "','" & cMoneda & "'," & CDbl(txtTipoCambio.Text) & ",'" & cTipo & "'"
    Else
        lsSQL = "exec stp_sel_ListadoPagosPorFechas '" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sEstados & "','" & cMoneda & "'," & CDbl(txtTipoCambio.Text)
    End If

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set R = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing

    If R.EOF And R.BOF Then
        MsgBox "No existe datos para este Reporte.", vbInformation, "Atención"
        Exit Sub
    End If

'    Set oDCred = New COMDCredito.DCOMCreditos
'    Set R = oDCred.ObtenerGarantiasInscritas(pdFecFinal, sCadAge, psMoneda, pMatProd)
'    Set oDCred = Nothing
    
'    If R.EOF And R.BOF Then
'        MsgBox "No existe datos para este Reporte.", vbInformation, "Atención"
'        Exit Sub
'    End If
    
    lsNombreArchivo = "Pagos_Realizados"
            
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Pagos", " Del " & CStr(pdFecIni) & " Al " & CStr(pdFecFin), lsNombreArchivo, lMatCabecera, R, 2, , , True, True)

    Screen.MousePointer = 0

End Sub
'ALPA**20080820********************************************
Public Sub Reporte_Adjudicados138005(ByVal psFecIni As Date)
Dim oCont As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
Dim lsNombreArchivo As String
Dim lMatCabecera As Variant
Dim lsmensaje As String

    lsNombreArchivo = "Reporte_Adjudicados138005"

    ReDim lMatCabecera(10, 2) '4

    lMatCabecera(0, 0) = "CodigoPersona"
    lMatCabecera(1, 0) = "Persona"
    lMatCabecera(2, 0) = "Cuenta"
    lMatCabecera(3, 0) = "NGarantia"
    lMatCabecera(4, 0) = "Descripcion Garantia"
    lMatCabecera(5, 0) = "Direccion"
    lMatCabecera(6, 0) = "Saneamiento"
    lMatCabecera(7, 0) = "CodigoComprador"
    lMatCabecera(8, 0) = "NombreComprador"
    lMatCabecera(9, 0) = "MontoVenta"
    
        
    Set oCont = New COMDCredito.DCOMGarantia
    Set R = oCont.Reporte_Adjudicados138005(psFecIni)
    Set oCont = Nothing

    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Intereses Compensatorios de Fideicomiso", " Al " & Format(psFecIni, "MM/YYYY"), lsNombreArchivo, lMatCabecera, R, 2, , , True)
    'R = Nothing
End Sub
Sub GenAdjudicados138005()
        Call Reporte_Adjudicados138005(CDate(Me.mskPeriodo1Del))
End Sub

'*** PEAC 20080926

Public Function nRepo138024_ListadoExcelExpedientesxAbogado(ByVal sConAbogado As Byte, ByVal sTipC As Double, ByVal sCodAbogado As String, ByVal pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, ByVal psCodAge As String, ByVal psCodUser As String, _
    ByVal pdFecSis As Date, ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim oFun As New COMFunciones.FCOMImpresion
Dim oConec As New COMConecta.DCOMConecta
Dim lcCodAbogado As String, lcCodProducto As String, lcCodMoneda As String
Dim psMensaje  As String
Dim i As Integer, J As Integer

Dim lnTotPrestamo As Double, lnTotSaldo As Double, lnTotInteres As Double, lnTotMora As Double, lnTotGastos As Double, lnTotTotal As Double
Dim lnTotTipCredPrestamo As Double, lnTotTipCredSaldo As Double, lnTotTipCredInteres As Double, lnTotTipCredMora As Double, lnTotTipCredGastos As Double, lnTotTipCredTotal As Double
Dim lnTotAboPrestamo As Double, lnTotAboSaldo As Double, lnTotAboInteres As Double, lnTotAboMora As Double, lnTotAboGastos As Double, lnTotAboTotal As Double

Dim sCadAge As String
Dim vmone As String
      
    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

    If sConAbogado = 1 Then
        lsSQL = " exec stp_sel_ReporteCredEnCJporAbogadoDetalle " & sConAbogado & ",'" & Format(pdFecSis, "yyyymmdd") & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "','" & sCodAbogado & "'"
    Else
        lsSQL = " exec stp_sel_ReporteCredEnCJporAbogadoDetalle " & sConAbogado & ",'" & Format(pdFecSis, "yyyymmdd") & "','" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "'"
    End If

    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
        
    If R Is Nothing Or (R.BOF And R.EOF) Then
        psMensaje = "Lo sentimos, No Existen Datos para este reporte."
        Exit Function
    Else
'---------------------------****************************** inicia excel

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    
    ApExcel.Cells(1, 2).Formula = psNomCmac
    ApExcel.Cells(2, 2).Formula = psNomAge
    
    ApExcel.Cells(1, 13).Formula = Date + Time
    ApExcel.Cells(2, 13).Formula = psCodUser
    
    ApExcel.Cells(4, 2).Formula = "CREDITOS EN C.J POR ABOGADO DETALLE"
    ApExcel.Cells(5, 2).Formula = "DEL " & Format(pdFecIni, "dd/MM/YYYY") & " AL " & Format(pdFecFin, "dd/mm/yyyy")
    
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B4", "M4").MergeCells = True
    ApExcel.Range("B5", "M5").MergeCells = True
        
    ApExcel.Cells(7, 2).Formula = "TIPO CAMBIO : " & oFun.ImpreFormat(sTipC, 5, 3)
    ApExcel.Cells(7, 2).Font.Bold = True
    
    ApExcel.Range("B1", "M8").Font.Bold = True
    
    ApExcel.Range("B1", "B2").HorizontalAlignment = xlLeft
    ApExcel.Range("L1", "M2").HorizontalAlignment = xlRight
    ApExcel.Range("B4", "M5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(9, 2).Formula = "Nº"
    ApExcel.Cells(9, 3).Formula = "CUENTA"
    ApExcel.Cells(9, 4).Formula = "F.ING.REC."
    ApExcel.Cells(9, 5).Formula = "CLIENTE"
    ApExcel.Cells(9, 6).Formula = "PRESTAMO"
    ApExcel.Cells(9, 7).Formula = "SALDO CAP."
    ApExcel.Cells(9, 8).Formula = "INTERES"
    ApExcel.Cells(9, 9).Formula = "MORA"
    ApExcel.Cells(9, 10).Formula = "GASTOS"
    ApExcel.Cells(9, 11).Formula = "TOTAL"
    ApExcel.Cells(9, 12).Formula = "ANALISTA"
    ApExcel.Cells(9, 13).Formula = "ESTADO"
    
    ApExcel.Range("B9", "M9").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "M9").Font.Bold = True
    ApExcel.Range("B9", "M9").HorizontalAlignment = 3

    i = 9

    Do While Not R.EOF

        i = i + 1
        
        lcCodAbogado = R!codanalista
        
        ApExcel.Cells(i, 2).Formula = "ABOGADO : " & R!Abogado & " : " & R!codanalista
        ApExcel.Cells(i, 2).Font.Bold = True
        
        
        lnTotAboPrestamo = 0: lnTotAboSaldo = 0: lnTotAboInteres = 0: lnTotAboMora = 0: lnTotAboGastos = 0: lnTotAboTotal = 0
        
        
        Do While R!codanalista = lcCodAbogado
        
        i = i + 1
                
        lcCodProducto = R!cCodTipoCredito
                        
        ApExcel.Cells(i, 2).Formula = "TIPO DE CREDITO : " & R!cDesTipoCredito
        ApExcel.Cells(i, 2).Font.Bold = True
        
        lnTotTipCredPrestamo = 0: lnTotTipCredSaldo = 0: lnTotTipCredInteres = 0: lnTotTipCredMora = 0: lnTotTipCredGastos = 0: lnTotTipCredTotal = 0
                
        Do While R!cCodTipoCredito = lcCodProducto And R!codanalista = lcCodAbogado
        
        i = i + 1
        
        lcCodMoneda = R!cMoneda
                        
        ApExcel.Cells(i, 2).Formula = "MONEDA : " & R!sMoneda
        ApExcel.Cells(i, 2).Font.Bold = True
        
        lnTotPrestamo = 0: lnTotSaldo = 0: lnTotInteres = 0: lnTotMora = 0: lnTotGastos = 0: lnTotTotal = 0
        
        J = 0
        Do While R!cMoneda = lcCodMoneda And R!cCodTipoCredito = lcCodProducto And R!codanalista = lcCodAbogado

                J = J + 1
                i = i + 1
                ApExcel.Cells(i, 2).Formula = J
                ApExcel.Cells(i, 3).Formula = "'" & R!cCtaCod
                ApExcel.Cells(i, 4).Formula = Format(R!dIngRecup, "dd/mm/yyyy")
                ApExcel.Cells(i, 5).Formula = R!cPersNombre
                ApExcel.Cells(i, 6).Formula = R!prestamo
                ApExcel.Cells(i, 7).Formula = R!Saldo
                ApExcel.Cells(i, 8).Formula = R!Interes
                ApExcel.Cells(i, 9).Formula = R!Mora
                ApExcel.Cells(i, 10).Formula = R!Gasto
                ApExcel.Cells(i, 11).Formula = (R!Saldo + R!Interes + R!Mora + R!Gasto)
                ApExcel.Cells(i, 12).Formula = R!Analista
                ApExcel.Cells(i, 13).Formula = R!cEstado
                
                lnTotPrestamo = lnTotPrestamo + R!prestamo
                lnTotSaldo = lnTotSaldo + R!Saldo
                lnTotInteres = lnTotInteres + R!Interes
                lnTotMora = lnTotMora + R!Mora
                lnTotGastos = lnTotGastos + R!Gasto
                lnTotTotal = lnTotTotal + R!Saldo + R!Interes + R!Mora + R!Gasto
                
                ApExcel.Range("F" & Trim(str(i)) & ":" & "K" & Trim(str(i))).NumberFormat = "#,##0.00"
                ApExcel.Range("B" & Trim(str(i)) & ":" & "M" & Trim(str(i))).Borders.LineStyle = 1

                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
        Loop
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "TOTAL POR MONEDA"
        ApExcel.Cells(i, 6).Formula = lnTotPrestamo
        ApExcel.Cells(i, 7).Formula = lnTotSaldo
        ApExcel.Cells(i, 8).Formula = lnTotInteres
        ApExcel.Cells(i, 9).Formula = lnTotMora
        ApExcel.Cells(i, 10).Formula = lnTotGastos
        ApExcel.Cells(i, 11).Formula = lnTotTotal
        ApExcel.Range("B" & Trim(str(i)) & ":" & "K" & Trim(str(i))).Font.Bold = True
        
        
        lnTotTipCredPrestamo = lnTotTipCredPrestamo + (lnTotPrestamo * IIf(lcCodMoneda = 1, 1, sTipC))
        lnTotTipCredSaldo = lnTotTipCredSaldo + (lnTotSaldo * IIf(lcCodMoneda = 1, 1, sTipC))
        lnTotTipCredInteres = lnTotTipCredInteres + (lnTotInteres * IIf(lcCodMoneda = 1, 1, sTipC))
        lnTotTipCredMora = lnTotTipCredMora + (lnTotMora * IIf(lcCodMoneda = 1, 1, sTipC))
        lnTotTipCredGastos = lnTotTipCredGastos + (lnTotGastos * IIf(lcCodMoneda = 1, 1, sTipC))
        lnTotTipCredTotal = lnTotTipCredTotal + (lnTotTotal * IIf(lcCodMoneda = 1, 1, sTipC))
        
        
        i = i + 1
        
        If R.EOF Then
            Exit Do
        End If
    Loop
    
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "TOTAL POR TIPO DE CREDITO"
        ApExcel.Cells(i, 6).Formula = lnTotTipCredPrestamo
        ApExcel.Cells(i, 7).Formula = lnTotTipCredSaldo
        ApExcel.Cells(i, 8).Formula = lnTotTipCredInteres
        ApExcel.Cells(i, 9).Formula = lnTotTipCredMora
        ApExcel.Cells(i, 10).Formula = lnTotTipCredGastos
        ApExcel.Cells(i, 11).Formula = lnTotTipCredTotal
        ApExcel.Range("B" & Trim(str(i)) & ":" & "K" & Trim(str(i))).Font.Bold = True
        
        lnTotAboPrestamo = lnTotAboPrestamo + lnTotTipCredPrestamo
        lnTotAboSaldo = lnTotAboSaldo + lnTotTipCredSaldo
        lnTotAboInteres = lnTotAboInteres + lnTotTipCredInteres
        lnTotAboMora = lnTotAboMora + lnTotTipCredMora
        lnTotAboGastos = lnTotAboGastos + lnTotTipCredGastos
        lnTotAboTotal = lnTotAboTotal + lnTotTipCredTotal
        
        i = i + 1
        
        If R.EOF Then
            Exit Do
        End If
    
    Loop
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "TOTAL POR ABOGADO"
        ApExcel.Cells(i, 6).Formula = lnTotAboPrestamo
        ApExcel.Cells(i, 7).Formula = lnTotAboSaldo
        ApExcel.Cells(i, 8).Formula = lnTotAboInteres
        ApExcel.Cells(i, 9).Formula = lnTotAboMora
        ApExcel.Cells(i, 10).Formula = lnTotAboGastos
        ApExcel.Cells(i, 11).Formula = lnTotAboTotal
        ApExcel.Range("B" & Trim(str(i)) & ":" & "K" & Trim(str(i))).Font.Bold = True
        
        i = i + 1
    
    Loop
        
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 6#
    ApExcel.Range("B2").Select

    ApExcel.Visible = True
    Set ApExcel = Nothing
       
    End If
End Function

'*** PEAC 20080929
Public Function nRepo138021_ListadoExcelCreditosEnCJxAbogado(ByVal sConAbogado As Byte, ByVal sTipC As Double, ByVal sCodAbogado As String, ByVal pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, ByVal psCodAge As String, ByVal psCodUser As String, _
    ByVal pdFecSis As Date, ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim oFun As New COMFunciones.FCOMImpresion
Dim oConec As New COMConecta.DCOMConecta
Dim lcCodAbogado As String, lcCodProducto As String, lcCodMoneda As String
Dim psMensaje  As String
Dim i As Integer, J As Integer

Dim lnTotPrestamo As Double, lnTotSaldo As Double, lnTotInteres As Double, lnTotMora As Double, lnTotGastos As Double, lnTotTotal As Double
Dim lnTotTipCredPrestamo As Double, lnTotTipCredSaldo As Double, lnTotTipCredInteres As Double, lnTotTipCredMora As Double, lnTotTipCredGastos As Double, lnTotTipCredTotal As Double
Dim lnTotAboPrestamo As Double, lnTotAboSaldo As Double, lnTotAboInteres As Double, lnTotAboMora As Double, lnTotAboGastos As Double, lnTotAboTotal As Double

Dim sCadAge As String
Dim vmone As String
      
    sCadAge = ""
    For i = 0 To UBound(pMatAgencias) - 1
        sCadAge = sCadAge & pMatAgencias(i) & ","
    Next i
    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)

    If sConAbogado = 1 Then
        lsSQL = " exec stp_sel_ReporteCredEnCJporAbogadoResumen " & sConAbogado & "," & sTipC & ",'" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "','" & sCodAbogado & "'"
    Else
        lsSQL = " exec stp_sel_ReporteCredEnCJporAbogadoResumen " & sConAbogado & "," & sTipC & ",'" & Format(pdFecIni, "yyyymmdd") & "','" & Format(pdFecFin, "yyyymmdd") & "','" & sCadAge & "'"
    End If

    oConec.AbreConexion
    Set R = oConec.CargaRecordSet(lsSQL)
    oConec.CierraConexion
    Set oConec = Nothing
        
    If R Is Nothing Or (R.BOF And R.EOF) Then
        psMensaje = "Lo sentimos, No Existen Datos para este reporte."
        Exit Function
    Else
'---------------------------****************************** inicia excel

    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
    'Poner Titulos
    
    ApExcel.Cells(1, 2).Formula = psNomCmac
    ApExcel.Cells(2, 2).Formula = psNomAge
    
    ApExcel.Cells(1, 8).Formula = Date + Time
    ApExcel.Cells(2, 8).Formula = psCodUser
    
    ApExcel.Cells(4, 2).Formula = "CREDITOS EN C.J POR ABOGADO (CONSOLIDADO)"
    ApExcel.Cells(5, 2).Formula = "DEL " & Format(pdFecIni, "dd/MM/YYYY") & " AL " & Format(pdFecFin, "dd/mm/yyyy")
    
    'ApExcel.cells(2, 2).Font.Size = 12
    ApExcel.Range("B4", "H4").MergeCells = True
    ApExcel.Range("B5", "H5").MergeCells = True
        
    ApExcel.Cells(7, 2).Formula = "TIPO CAMBIO : " & oFun.ImpreFormat(sTipC, 5, 3)
    ApExcel.Cells(7, 2).Font.Bold = True
    
    ApExcel.Range("B1", "H8").Font.Bold = True
    
    ApExcel.Range("B1", "B2").HorizontalAlignment = xlLeft
    ApExcel.Range("L1", "H2").HorizontalAlignment = xlRight
    ApExcel.Range("B4", "H5").HorizontalAlignment = xlCenter
    
    ApExcel.Cells(9, 2).Formula = "DESCRIPCION"
    ApExcel.Cells(9, 3).Formula = "PRESTAMO"
    ApExcel.Cells(9, 4).Formula = "CAPITAL"
    ApExcel.Cells(9, 5).Formula = "INTERES"
    ApExcel.Cells(9, 6).Formula = "MORA"
    ApExcel.Cells(9, 7).Formula = "GASTOS"
    ApExcel.Cells(9, 8).Formula = "TOTAL"

    ApExcel.Range("B9", "H9").Interior.Color = RGB(10, 190, 160)
    ApExcel.Range("B9", "H9").Font.Bold = True
    ApExcel.Range("B9", "H9").HorizontalAlignment = 3

    i = 9

    lnTotPrestamo = 0: lnTotSaldo = 0: lnTotInteres = 0: lnTotMora = 0: lnTotGastos = 0: lnTotTotal = 0
        
    Do While Not R.EOF
        i = i + 1
        
        lcCodAbogado = R!CodAbogado
        
        ApExcel.Cells(i, 2).Formula = "ABOGADO : " & R!Abogado & " : " & R!CodAbogado
        ApExcel.Cells(i, 2).Font.Bold = True
        
        lnTotAboPrestamo = 0: lnTotAboSaldo = 0: lnTotAboInteres = 0: lnTotAboMora = 0: lnTotAboGastos = 0: lnTotAboTotal = 0
                
        Do While R!CodAbogado = lcCodAbogado
        
                J = J + 1
                i = i + 1
                ApExcel.Cells(i, 2).Formula = R!desestadorecuperacion
                ApExcel.Cells(i, 3).Formula = R!prestamo
                ApExcel.Cells(i, 4).Formula = R!Saldo
                ApExcel.Cells(i, 5).Formula = R!Interes
                ApExcel.Cells(i, 6).Formula = R!Mora
                ApExcel.Cells(i, 7).Formula = R!Gasto
                ApExcel.Cells(i, 8).Formula = R!Total
                
                lnTotAboPrestamo = lnTotAboPrestamo + R!prestamo
                lnTotAboSaldo = lnTotAboSaldo + R!Saldo
                lnTotAboInteres = lnTotAboInteres + R!Interes
                lnTotAboMora = lnTotAboMora + R!Mora
                lnTotAboGastos = lnTotAboGastos + R!Gasto
                lnTotAboTotal = lnTotAboTotal + R!Total
                
                ApExcel.Range("C" & Trim(str(i)) & ":" & "H" & Trim(str(i))).NumberFormat = "#,##0.00"
                ApExcel.Range("B" & Trim(str(i)) & ":" & "H" & Trim(str(i))).Borders.LineStyle = 1

                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
        Loop
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "TOTAL POR ABOGADO"
        ApExcel.Cells(i, 3).Formula = lnTotAboPrestamo
        ApExcel.Cells(i, 4).Formula = lnTotAboSaldo
        ApExcel.Cells(i, 5).Formula = lnTotAboInteres
        ApExcel.Cells(i, 6).Formula = lnTotAboMora
        ApExcel.Cells(i, 7).Formula = lnTotAboGastos
        ApExcel.Cells(i, 8).Formula = lnTotAboTotal
        ApExcel.Range("B" & Trim(str(i)) & ":" & "H" & Trim(str(i))).Font.Bold = True
        
        lnTotPrestamo = lnTotPrestamo + lnTotAboPrestamo
        lnTotSaldo = lnTotSaldo + lnTotAboSaldo
        lnTotInteres = lnTotInteres + lnTotAboInteres
        lnTotMora = lnTotMora + lnTotAboMora
        lnTotGastos = lnTotGastos + lnTotAboGastos
        lnTotTotal = lnTotTotal + lnTotAboTotal
        
        i = i + 1
           
    Loop
    
        i = i + 1
        ApExcel.Cells(i, 2).Formula = "TOTAL"
        ApExcel.Cells(i, 3).Formula = lnTotPrestamo
        ApExcel.Cells(i, 4).Formula = lnTotSaldo
        ApExcel.Cells(i, 5).Formula = lnTotInteres
        ApExcel.Cells(i, 6).Formula = lnTotMora
        ApExcel.Cells(i, 7).Formula = lnTotGastos
        ApExcel.Cells(i, 8).Formula = lnTotTotal
        ApExcel.Range("B" & Trim(str(i)) & ":" & "H" & Trim(str(i))).Font.Bold = True
        
        i = i + 1
    
    R.Close
    Set R = Nothing
   
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("B:B").ColumnWidth = 27#
    ApExcel.Range("B2").Select

    ApExcel.Visible = True
    Set ApExcel = Nothing
       
    End If
End Function


