VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   8250
   ClientLeft      =   1920
   ClientTop       =   1740
   ClientWidth     =   8445
   HelpContextID   =   210
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdValorizacion 
      Caption         =   "Registrar Valorización"
      Height          =   495
      Left            =   6840
      TabIndex        =   58
      Top             =   980
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraSemestre 
      Height          =   1335
      Left            =   5040
      TabIndex        =   53
      Top             =   1920
      Visible         =   0   'False
      Width           =   3360
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmReportes.frx":030A
         Left            =   1440
         List            =   "frmReportes.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboSemestre 
         Height          =   315
         ItemData        =   "frmReportes.frx":032F
         Left            =   1440
         List            =   "frmReportes.frx":0339
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   600
         TabIndex        =   57
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Semestre:"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraFechaRango 
      Caption         =   "Rango de Fechas"
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
      Left            =   5010
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   3360
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   300
         Left            =   510
         TabIndex        =   27
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
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
         Caption         =   "Del"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1770
         TabIndex        =   29
         Top             =   315
         Width           =   135
      End
   End
   Begin VB.CommandButton txtInst 
      Caption         =   "Inst."
      Height          =   330
      Left            =   7760
      TabIndex        =   46
      Top             =   1110
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txtNroUITs 
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
      Height          =   345
      Left            =   7200
      TabIndex        =   44
      Text            =   "0"
      Top             =   2640
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdFTP 
      Caption         =   "FTP"
      Height          =   330
      Left            =   6840
      TabIndex        =   43
      Top             =   7845
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   7680
      TabIndex        =   38
      Top             =   7845
      Width           =   660
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   330
      Left            =   4920
      TabIndex        =   37
      Top             =   7845
      Width           =   840
   End
   Begin VB.Frame frmMoneda 
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
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   4965
      TabIndex        =   23
      Top             =   90
      Width           =   3360
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &Nacional"
         Height          =   345
         Index           =   0
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &Extranjera"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame fraAge 
      Caption         =   "Agencia"
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
      Height          =   1575
      Left            =   4965
      TabIndex        =   3
      Top             =   3225
      Visible         =   0   'False
      Width           =   3360
      Begin VB.ListBox lstAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   3105
      End
   End
   Begin MSComctlLib.TreeView tvOpe 
      Height          =   8085
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   14261
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglstFiguras"
      BorderStyle     =   1
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
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   180
      Top             =   4860
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
            Picture         =   "frmReportes.frx":0356
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":06A8
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":09FA
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportes.frx":0D4C
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFuentes 
      Caption         =   "Fuentes de Financiamiento"
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
      Height          =   1545
      Left            =   4965
      TabIndex        =   1
      Top             =   4815
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ListBox lstFuentes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   300
         Width           =   3105
      End
   End
   Begin VB.Frame fraTCambio 
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
      Height          =   690
      Left            =   4995
      TabIndex        =   14
      Top             =   1875
      Visible         =   0   'False
      Width           =   3330
      Begin VB.TextBox txtTipCambio 
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
         Height          =   345
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   3
         Left            =   4500
         TabIndex        =   16
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtTipCambio2 
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
         Height          =   345
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T.C. Cierre Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   19
         Top             =   705
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "T.C. del Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   255
         Width           =   1020
      End
   End
   Begin VB.Frame fraTC 
      Caption         =   "Tipo de Cambio"
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
      Height          =   1395
      Left            =   5010
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox TxtTipCamFijAnt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   360
         TabIndex        =   9
         Text            =   "0"
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox txtTipCamCompraSBS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   360
         TabIndex        =   8
         Text            =   "0"
         Top             =   945
         Width           =   1125
      End
      Begin VB.TextBox txtTipCamFij 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   1920
         TabIndex        =   7
         Text            =   "0"
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox txtTipCamVentaSBS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   325
         Left            =   1920
         TabIndex        =   6
         Text            =   "0"
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label lblTipcambiAnt 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Fijo.Ant"
         Height          =   195
         Left            =   510
         TabIndex        =   13
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Comp.SBS"
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblTipcamFij 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Fijo.Act"
         Height          =   195
         Left            =   2010
         TabIndex        =   11
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Venta.SBS"
         Height          =   195
         Left            =   1875
         TabIndex        =   10
         Top             =   735
         Width           =   1080
      End
   End
   Begin VB.Frame fraEuros 
      Caption         =   "Euros"
      Height          =   750
      Left            =   5040
      TabIndex        =   40
      Top             =   990
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox txtTCEuros 
         Height          =   285
         Left            =   1215
         TabIndex        =   41
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "T.C. Euros"
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
         Left            =   165
         TabIndex        =   42
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraPeriodo 
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
      ForeColor       =   &H8000000D&
      Height          =   1035
      Left            =   5010
      TabIndex        =   31
      Top             =   870
      Visible         =   0   'False
      Width           =   3360
      Begin VB.TextBox txtAnio 
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
         Height          =   330
         Left            =   1575
         MaxLength       =   4
         TabIndex        =   33
         Top             =   615
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
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
         ItemData        =   "frmReportes.frx":109E
         Left            =   870
         List            =   "frmReportes.frx":10C6
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   255
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   285
         TabIndex        =   35
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   285
         TabIndex        =   34
         Top             =   330
         Width           =   390
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha"
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
      Left            =   5010
      TabIndex        =   20
      Top             =   870
      Visible         =   0   'False
      Width           =   1695
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   390
         TabIndex        =   21
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   135
      End
   End
   Begin VB.Frame FraTasaMinEncaje 
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
      Height          =   690
      Left            =   5040
      TabIndex        =   47
      Top             =   4080
      Visible         =   0   'False
      Width           =   3330
      Begin VB.TextBox Text2 
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
         Height          =   345
         Left            =   1500
         MaxLength       =   16
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   690
         Width           =   1185
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "A&justado"
         Height          =   255
         Index           =   2
         Left            =   4500
         TabIndex        =   49
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtTasaMinEncaje 
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
         Height          =   345
         Left            =   1860
         MaxLength       =   16
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Minimo Encaje"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   52
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "T.C. Cierre Mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   51
         Top             =   705
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "SUCAVE"
      Height          =   330
      Left            =   5880
      TabIndex        =   39
      Top             =   7845
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   5895
      TabIndex        =   36
      Top             =   7845
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblNroUIT 
      Caption         =   "Nº UITs"
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
      Left            =   5100
      TabIndex        =   45
      Top             =   2680
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lExpand  As Boolean
Dim lExpandO As Boolean
Dim sArea  As String
Dim sTipoRepo As Integer
Dim Progress As clsProgressBar
Dim psCtaCont As String
Dim Arrays2() As String
Dim nLin As Long
Dim lbLibroOpen As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim objPista As COMManejador.Pista 'ARLO20170217

'EJVG20121002 ***
Private Type TConcentraFondos
    CodPersona As String
    Nombre  As String
    CtaIFCodCtaCorriente As String
    CtaIFDescCtaCorriente As String
    SaldoCtaCorriente As Currency
    CtaIFCodCtaAhorro As String
    CtaIFDescCtaAhorro As String
    SaldoCtaAhorro As Currency
    SaldoTotalInversion As Currency
    SaldoTotalDPFOver As Currency
End Type
Dim MatCtasBcosMN() As TConcentraFondos
Dim MatCtasBcosME() As TConcentraFondos
Dim MatCtasCMACsMN() As TConcentraFondos
Dim MatCtasCMACsME() As TConcentraFondos
Dim MatCtasCRACsMN() As TConcentraFondos
Dim MatCtasCRACsME() As TConcentraFondos
'END EJVG *******
Dim fsOtrasIFisRptConcentraFondos As String 'EJVG20130927
Dim lnPosFilaMNTotal As Integer, lnPosFilaMETotal As Integer 'EJVG20131230

'FRHU 20140104 RQ13658
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
    bSaldoA As Boolean
    bSaldoD As Boolean
End Type
Private Type TColumna
    fila As Integer
    Primero As Currency ' Columna 1
    Segundo As Currency ' Columna 2
    DebeAumento As Currency ' Columna 3
    HaberDisminucion As Currency ' Columna 4
    AjusteDebe As Currency ' Columna 5
    AjusteHaber As Currency ' Columna 6
    OperacionDebe As Currency ' Columna 7
    OperacionHaber As Currency ' Columna 8
End Type
'FIN FRHU 20140104

Public Sub inicio(sObj As String, Optional plExpandO As Boolean = False)
    sArea = sObj
    lExpandO = plExpandO
    Me.Show 0, frmMdiMain
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer
ValidaDatos = False
   If fraFechaRango.Visible Then
      If Not ValFecha(txtFechaDel) Then
         txtFechaDel.SetFocus: Exit Function
      End If
      If Not ValFecha(txtFechaAl) Then
         txtFechaAl.SetFocus: Exit Function
      End If
   End If
   If fraFecha.Visible Then
      If Not ValFecha(txtFecha) Then
         txtFecha.SetFocus: Exit Function
      End If
   End If
   If fraPeriodo.Visible Then
      If nVal(txtAnio) = 0 Then
         MsgBox "Ingrese Año para generar Reporte...", vbInformation, "¡Aviso!"
         txtAnio.SetFocus
         Exit Function
      End If
      If cboMes.ListIndex = -1 Then
        MsgBox "Selecciones Mes para generar Reporte...", vbInformation, "¡Aviso!"
        cboMes.SetFocus
        Exit Function
      End If
   End If
   If Not tvOpe.SelectedItem.Child Is Nothing Then
        MsgBox "Seleccione Reporte de último Nivel", vbInformation, "¡Aviso!"
        tvOpe.SetFocus
        Exit Function
   End If
   
   ValidaDatos = True
    
   If fraAge.Visible = True Then
        ValidaDatos = False
        For i = 0 To lstAge.ListCount - 1
            If lstAge.Selected(i) Then
                ValidaDatos = True
            End If
        Next
        If ValidaDatos = False Then
            MsgBox "Ud. debe seleccionar al menos una agencia", vbInformation, "Aviso"
            lstAge.SetFocus
            Exit Function
        End If
          
'        lstAge.Selected(lstAge.ListCount - 1) = False
          
    End If
   
    If ValidaDatos = False Then
        Exit Function
    End If
    
    '*** PEAC 20080919
    If Me.txtTipCambio.Visible Then
        If CDbl(Me.txtTipCambio.Text) <= 0 Then
            MsgBox "Ud. debe ingresar el tipo de cambio", vbInformation, "Aviso"
            Me.txtTipCambio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    '*** FIN PEAC
    
    If fraFuentes.Visible = True Then
        ValidaDatos = False
        For i = 0 To lstFuentes.ListCount - 1
            If lstFuentes.Selected(i) Then
                ValidaDatos = True
            End If
        Next
        If ValidaDatos = False Then
            MsgBox "Ud. debe seleccionar al menos una fuente", vbInformation, "Aviso"
            lstFuentes.SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub CboMes_Click()
If Me.fraTCambio.Visible Then
    txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fraTCambio.Visible Then
            txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
        End If
        txtAnio.SetFocus
    End If
End Sub

Private Sub cmdFTP_Click()
Dim Opt As Integer

    Select Case Mid(gsOpeCod, 1, 6)
        Case RepCGEncBCRObligacion
            frmAnxEncajeBCR.GeneraAnx01MN txtFechaAl
        
        Case RepCGEncBCRObligacionME
            frmAnxEncajeBCR.GeneraAnx01ME txtFechaAl
        
        Case RepCGEncBCRCredDeposi
'            Opt = MsgBox("Desea Configurar el Anexo para su exportacion", vbInformation + vbYesNo, "AVISO")
'            If Opt = vbNo Then
                frmAnxEncajeBCR.GeneraAnx02MN txtFechaAl
'            Else
'                FrmAnxencajeBCRAnexo2.lblFecha = txtFechaAl
'                FrmAnxencajeBCRAnexo2.OptSoles.value = True
'                FrmAnxencajeBCRAnexo2.Show
'            End If
            
        
        Case RepCGEncBCRCredDeposiME
'            Opt = MsgBox("Desea Configurar el Anexo para su exportacion", vbInformation + vbYesNo, "AVISO")
'            If Opt = vbNo Then
                frmAnxEncajeBCR.GeneraAnx02ME txtFechaAl
'            Else
'                FrmAnxencajeBCRAnexo2.lblFecha = txtFechaAl
'                FrmAnxencajeBCRAnexo2.OptDolares.value = True
'                FrmAnxencajeBCRAnexo2.Show
'            End If
            
        
        Case RepCGEncBCRCredRecibi
            frmAnxEncajeBCR.GeneraAnx03MN txtFechaAl

        Case RepCGEncBCRCredRecibiME
            frmAnxEncajeBCR.GeneraAnx03ME txtFechaAl
        
        Case RepCGEncBCRObligaExon
            frmAnxEncajeBCR.GeneraAnx04MN txtFechaAl
        
        Case RepCGEncBCRObligaExonME
            frmAnxEncajeBCR.GeneraAnx04ME txtFechaAl
            
        Case "761206", "762206" 'PASI 20140305 TI-ERS102-2013
            frmAnxEncajeBCR.GeneraTxt01 Mid(gsOpeCod, 3, 1), txtFechaAl
        Case "761207", "762207" 'PASI 20140305 TI-ERS102-2013
            frmAnxEncajeBCR.GeneraTxt02 Mid(gsOpeCod, 3, 1), txtFechaAl
        Case "761208", "762208" 'PASI 20140305 TI-ERS102-2013
            frmAnxEncajeBCR.GeneraTxt03 Mid(gsOpeCod, 3, 1), txtFechaAl
        Case "761209", "762209" 'PASI 20140305 TI-ERS102-2013
            frmAnxEncajeBCR.GeneraTxt04 Mid(gsOpeCod, 3, 1), txtFechaAl
        
    End Select
End Sub

Private Sub cmdGenerar_Click()
Dim lsMoneda As String
Dim ldFecha  As Date
Dim sFuente As String
Dim i As Integer
Dim oEstadistica As NEstadistica


'On Error GoTo cmdGenerarErr
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    If Not ValidaDatos Then Exit Sub
    
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    lsMoneda = IIf(optMoneda(0).value, "1", "2")
    If Me.fraPeriodo.Visible Then
        ldFecha = CDate(DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))) - 1
    End If
    If lsMoneda = "1" Then
        gsSimbolo = gcMN
    Else
        gsSimbolo = gcME
    End If
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
    
    'gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    
    Select Case Mid(gsOpeCod, 1, 4)
        Case Mid(gContRepBaseFormula, 1, 4)
            'If Not (gsOpeCod = gContRepBaseNotasEstadoSitFinan Or gsOpeCod = gContRepBaseNotasEstadoResultado Or gsOpeCod = gContRepBaseFormaB2) Then
            If Not (gsOpeCod = gContRepBaseNotasEstadoSitFinan Or gsOpeCod = gContRepBaseNotasEstadoResultado Or gsOpeCod = gContRepBaseFormaB2 Or gsOpeCod = gContRepEstadoFlujoEfectivo Or gsOpeCod = gContRepHojaTrabajoFlujoEfectivo Or gsOpeCod = gContRepBaseNotasComplementariaInfoAnual Or gsOpeCod = gContRepEstadoSitFinanEEFF1 Or gsOpeCod = gContRepEstadoSitFinanEEFF2 Or gsOpeCod = gContRepEstadoSitFinanEEFF3) Then 'FRHU 20140104 Agregue gContRepEstadoFlujoEfectivo,gContRepHojaTrabajoFlujoEfectivo
            frmRepBaseFormula.inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
            End If
    End Select
    
    Select Case Mid(gsOpeCod, 1, 6)
            
        '***************** CARTAS FIANZA ********************************
        Case OpeCGCartaFianzaRepIngreso, OpeCGCartaFianzaRepIngresoME
            ImprimeCartasFianza txtFechaDel, txtFechaAl, True
        Case OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME
            ImprimeCartasFianza txtFechaDel, txtFechaAl, False
        
        '********************* REPORTES DE CAJA GENERAL **************************
        Case OpeCGRepFlujoDiarioResMN, OpeCGRepFlujoDiarioResME
            ResumenFlujoDiario Me.txtFecha
        Case OpeCGRepFlujoDiarioDetMN, OpeCGRepFlujoDiarioDetME
            DetalleFlujoDiario txtFechaDel, txtFechaAl
        
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            frmCajaGenRepFlujos.Show 0, Me
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
            frmAdeudRepDet.inicio False
        Case OpeCGRepRepBancosResumenPFMN, OpeCGRepRepBancosResumenPFME
             frmPFReportes.Show 1
       
        Case OpeCGRepRepBancosConcentFdos
             ImprimeConcentracionFondos txtFecha.Text, Val(txtTipCambio.Text)
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
            frmCajaGenRepFlujos.Show 0, Me
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
            frmAdeudRepDet.inicio False
        
        Case OpeCGRepRepCMACSResumenPFMN, OpeCGRepRepCMACSResumenPFME
            frmPFReportes.Show 1
        'ALPA 20090316 **************************************************
        Case OpeCGRepProyecNroCliente
            Call ProyeccionesNroCliente
        '***************************************************************
        'Orden de Pago
        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
             ReporteOrdenesPago txtFechaDel, txtFechaAl, gsOpeCod
 
        'A rendir Cuenta
        Case OpeCGRepArendirLibroAuxMN, OpeCGRepArendirLibroAuxME
             ReporteArendirCuentaLibro txtFechaDel, txtFechaAl, gsOpeCod
             
        Case OpeCGRepArendirViaticoLibroAuxMN, OpeCGRepArendirViaticoLibroAuxME
             ReporteArendirCuentaViaticosLibro txtFechaDel, txtFechaAl, gsOpeCod

        Case OpeCGRepArendirPendienteMN, OpeCGRepArendirPendienteME
             ReporteARendirCuentaPendientes gsOpeCod, CInt(lsMoneda), txtFecha
   
        Case OpeCGRepArendirViaticosMN, OpeCGRepArendirViaticosME
             frmRepViaticos.Show 1
             
        '*** PEAC 20101108
        Case 461046
            ReporteSustentacionARendirCuenta gsOpeCod, CInt(lsMoneda), txtFechaDel, txtFechaAl
             
        'Cheques
        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
            ResumenCheques txtFecha, Mid(gsOpeCod, 1, 6), True
        'Remesas Cheques
        Case OpeCGRepChequesEnvMN, OpeCGRepChequesEnvME
            ResumenChequesRem txtFechaDel, txtFechaAl, gsOpeCod
            
        Case OpeCGRepChequesAnulMN, OpeCGRepChequesAnulME
            ResumenChequesAnul txtFechaDel, txtFechaAl, gsOpeCod
            
        Case OpeCGRepChequesCobMN, OpeCGRepChequesCobME
            ResumenChequesCob txtFechaDel, txtFechaAl, gsOpeCod
            
        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
            ResumenCheques txtFecha, Mid(gsOpeCod, 1, 6)

        Case OpeCGRepChqRecibidoCajaMN, OpeCGRepChqRecibidoCajaME
            ResumenChqRecibidos txtFecha, gsOpeCod
        Case OpeCGRepChqDepositadoMN, OpeCGRepChqDepositadoME
            ResumenChqDepositados txtFecha, gsOpeCod
        Case OpeCGRepChqPorCobrarConvMN, OpeCGRepChqPorCobrarConvME
            ReporteChequesPorCobrarConvenio Me.txtFecha, gsOpeCod
     
        'ADEUDADOS
        Case OpeCGAdeudRepGeneralMN, OpeCGAdeudRepGeneralME
            frmAdeudRepGen.inicio
        Case OpeCGAdeudRepDetalleMN, OpeCGADeudRepDetalleME
            frmAdeudRepDet.inicio True
        Case OpeCGAdeudRepSaldLinFinancDescalceMN
            ReporteSaldosLineaFinanciamientoDescalce gdFecSis
        Case OpeCGAdeudRepCortoPlazoMN, OpeCGAdeudRepCortoPlazoME
            frmAdeudRepGen.inicio True
        Case OpeCGAdeudRepxFecVenc, OpeCGAdeudRepxFecVencME
            frmAdeudRepVenc.inicio
            
        'Adeudados vinculados
        Case OpeCGAdeudRepVinculadosMN
             ReporteAdeudadosVinculados OpeCGAdeudRepVinculadosMN
             
        Case OpeCGAdeudRepVinculadosME
             If txtTCEuros.Text = "" Then
                MsgBox "Ingrese Tipo de Cambio Euros", vbInformation, "Aviso"
                txtTCEuros.SetFocus
                Exit Sub
             End If
             ReporteAdeudadosVinculados OpeCGAdeudRepVinculadosME, txtTCEuros.Text
             
                          
                    
        Case OpeCGRepPresuFlujoCaja, OpeCGRepPresuFlujoCajaME
            frmpresFlujoCaja.Show 1
        Case OpeCGRepPresuServDeuda, OpeCGRepPresuServDeudaME
            frmPresServDeuda.ImprimeServicioDeuda txtAnio, ldFecha, txtTipCambio
        Case OpeCGRepPresuFinancia, OpeCGRepPresuFinanciaME
            frmPresFinanExtInt.ImprimeReporteFinanciamientoIE txtAnio, ldFecha, txtTipCambio
        Case OpeCGRepOtrBilletesFalsosMN, OpeCGRepOtrBilletesFalsosME
            'frmReporteMonedaFalsa.lsMoneda = lsMoneda
            frmReporteMonedaFalsa.Show 1
        'ENCAJE
        Case OpeCGRepEncajeConsolSdoEnc, OpeCGRepEncajeConsolSdoEncME
            frmCajaGenReportes.ConsolidaSdoEnc gsOpeCod, txtFechaDel, txtFechaAl
        Case OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME
            frmSdoEncaje.CalculaSdoEnc lsMoneda, txtFechaDel, txtFechaAl

        Case OpeCGRepEncajeConsolPosLiq, OpeCGRepEncajeConsolPosLiqME
            frmCajaGenReportes.ConsolidaSdoEnc gsOpeCod, txtFechaDel, txtFechaAl, 2
        
        Case OpeCGRepEstadOpeUsuario, OpeCGRepEstadOpeUsuarioME
            frmRepOpeDiaUsuario.Show 1
            
        Case OpeEstEncajeSimulacionPlaEncajeMN, OpeEstEncajeSimulacionPlaEncajeME
            frmCGSimuladorPlanillaEncaje.Show 1
                       
        
        
        'Informe de Encaje al BCR
         Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME, _
              RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME, _
              RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME, _
              RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME, _
              RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
            frmRepEncajeBCR.ImprimeEncajeBCR gsOpeCod, txtFechaDel, txtFechaAl, CDbl(txtTipCambio.Text), CDbl(txtTipCambio2.Text)
        
        '*** Agregado por PASI 20140220 TI-ERS102-2013
        Case "761206", "761207", "761208", "761209", "762206", "762207", "762208", "762209"
            'frmGruRepEncajeBCR.ImprimeEncajeBCR gsOpeCod, txtFechaDel, txtFechaAl, CDbl(txtTipCambio.Text), CDbl(txtTipCambio2.Text)
            frmGruRepEncajeBCR.ImprimeEncajeBCR gsOpeCod, txtFechaDel, txtFechaAl, CDbl(txtTipCambio.Text), CDbl(txtTipCambio2.Text) 'PASIERS1382014 agrego txtTipCambio2
        '*** Fin PASI
        
        ' Saldos de Caja - Bancos y Agencias
        Case OpeCGRepSaldoBcos, OpeCGRepSaldoCajAge
            frmRptDiaLiquidez.inicio Mid(gsOpeCod, 1, 6)
        Case gOpeRptTripleJump, gOpeRptGtoFinanProd 'EJVG20111117
            frmRptAnioMesGral.Ini gsOpeCod, gsOpeDescHijo
        Case gOpeRptConcentracFondos 'EJVG20111123
            Call generarRptConcentracionFondos
        Case gOpRepCOFIDE  'MIOL20130626, RQ13329
            frmReporteCOFIDE.Show 1
        '************************* CNTABILIDAD **********************************
        Case gContLibroDiario
               frmContabDiario.Show 0, Me
        Case gContLibroMayor
               frmContabMayor.Show 0, Me
        Case gContLibroMayCta
               frmContabMayorDet.Show 0, Me
        Case gContLibInvBal
            frmLibroInventBalanc.Show 0, Me
        Case gContLibroCaja
            frmLibroCaja.Show 0, Me
            
        Case gConvFoncodes
              frmRptFoncodes.Show 0, Me

        Case gContRegCompraGastos
               frmRegCompraGastos.Show 0, Me
        Case 760201 'LIMA
               frmRegCompraGastos.PorDocumento True
               frmRegCompraGastos.Show 0, Me
        'EJVG20130426 ***
        Case gContRepBaseNotasEstadoSitFinan, gContRepBaseNotasEstadoResultado
            Call frmNIIFNotasEstado.inicio(gsOpeCod, gsOpeDescHijo)
        Case gContRepBaseFormaB2
            frmNIIFFormaB2.Show 1
        'END EJVG *******
        'FRHU 20140104 RQ13658
        Case gContRepEstadoFlujoEfectivo
            Call RepEstadoFlujoEfectivo(Me.txtAnio, Me.cboSemestre.ListIndex + 1)
        'FIN FRHU 20140104
        'FRHU 20140104 RQ13659
        Case gContRepHojaTrabajoFlujoEfectivo
            Call RepHojaTrabajoFlujoEfectivo(Me.txtAnio, Me.cboSemestre.ListIndex + 1)
        'FIN FRHU 20140104
        Case gContRepBaseNotasComplementariaInfoAnual
            frmNIIFNotasComplementaria.Show 1
        Case gContRepEstadoSitFinanEEFF1, gContRepEstadoSitFinanEEFF2, gContRepEstadoSitFinanEEFF3 'EJVG20140911
            Call frmNIIFBaseFormulasEEFFRep.inicio(gsOpeCod, gsOpeDescHijo)
        '************************* CONTABILIDAD HISTORICA**********************************
        
        Case gHistoContabDiario 'JACA 20111229
               frmHistoContabDiario.Show 0, Me
        Case gHistoContabMayor 'JACA 20111229
               frmHistoContabMayor.Show 0, Me
        Case gHistoLibroCaja 'JACA 20111229
            frmHistoLibroCaja.Show 0, Me
        Case gHistoContLibroMayCta 'ALPA 20111229
            frmHistoContabMayorDet.Show , Me

        '**************************************************
        '** GITU 20080923, Segun ACTA Nº 195-2008/TI-D
        Case 760202
               frmRepGastosAdmin.Show 0, Me
        '**************************************************
        Case gContRegVentas
                frmRegVenta.setnMoneda Val(lsMoneda) 'YIHU20152002-ERS181-2014
               frmRegVenta.Show 0, Me
        
        Case gContRepEstadIngGastos
               EstadisticaIngresoGasto cboMes.ListIndex + 1, txtAnio
        Case gContRepPlanillaPagoProv
               frmRepPagProv.Show 0, Me
        Case gContRepControlGastoProv
               frmProveeConsulMov.Show 1, Me
        Case gContRepPagoProvDetalleDoc
               frmRepProvCaja.Show 1, Me
        Case gContRepCompraVenta
            frmRepResCVenta.Show 0, Me

        Case gContRepEstadProv
            
            frmRepEstadProv.Show 1, Me

        'Otros Ajustes
        Case gContAjReclasiCartera
            frmAjusteReCartera.Show 0, Me
        Case gContAjReclasiGaranti
            frmAjusteGarantias.Show 0, Me
        Case gContAjInteresDevenga
            frmAjusteIntDevengado.inicio True
        Case gContAjInteresSuspens
            frmAjusteIntDevengado.inicio False
        
        Case 701228
            frmAjusteIntDevengado.inicio 11
        Case 702228
            frmAjusteIntDevengado.inicio 11
           
        Case 701241, 702241 'Ajuste de Garantias
            frmAjusteIntDevengado.inicio 19
        
        'ALPA 20131129****************************
        Case 701246, 702246
            frmAjusteIntDevengado.inicio 24
        '*****************************************
        'ALPA 20131202****************************
        Case 701247, 702247
            frmAjusteIntDevengado.inicio 25
        '*****************************************
        'ALPA 20131202****************************
        Case 701248, 702248
            frmAjusteIntDevengado.inicio 26
        '*****************************************
        'Riesgos
        Case gRiesgoCalfCarCred:
                    Call frmRiesgosReportes.inicio(gRiesgoCalfCarCred, gdFecSis)
                    frmRiesgosReportes.Show 1
        Case gRiesgoCalfAltoRiesgo:
                    Call frmRiesgosReportes.inicio(gRiesgoCalfAltoRiesgo, gdFecSis)
                    frmRiesgosReportes.Show 1
        Case gRiesgoConceCarCred:
                    Call frmRiesgosReportes.inicio(gRiesgoConceCarCred, gdFecSis)
                    frmRiesgosReportes.Show 1
        Case gRiesgoEstratDepPlazo:
                    Call frmRiesgosReportes.inicio(gRiesgoEstratDepPlazo, gdFecSis)
                    frmRiesgosReportes.Show 1
        Case gRiesgoPrincipClientesAhorros:
                    Call frmRiesgosReportes.inicio(gRiesgoPrincipClientesAhorros, gdFecSis)
                    frmRiesgosReportes.Show 1
        Case gRiesgoPrincipClientesCreditos:
                    Call frmRiesgosReportes.inicio(gRiesgoPrincipClientesCreditos, gdFecSis)
                    frmRiesgosReportes.Show 1
            
            
        '''captaciones
        Case gContRepCaptacCVMonExtr
            Consolida763201 gbBitCentral, Me.tvOpe.SelectedItem.Text, Me.lstAge, CDate(txtFechaDel.Text), CDate(txtFechaAl.Text)
        Case gContRepCaptacSituacCaptac
            Imprime763202 gbBitCentral, Me.tvOpe.SelectedItem.Text, Me.lstAge, CDate(txtFecha.Text)
        Case gContRepCaptacMovCV
            Imprime763203 Me.tvOpe.SelectedItem.Text, CDate(txtFechaDel.Text), TxtTipCamFijAnt, txtTipCamCompraSBS, txtTipCamVentaSBS, gbBitCentral
       
        Case gContRepCaptacIngPagos, gContRepCaptacCredDesem
            lstAge.Selected(lstAge.ListCount - 1) = False
            For i = 0 To lstFuentes.ListCount - 1
                If lstFuentes.Selected(i) = True Then
                    If Len(Trim(sFuente)) = 0 Then
                        sFuente = "'" & Right(lstFuentes.List(i), 1) & "'"
                    Else
                        sFuente = sFuente & ", '" & Right(lstFuentes.List(i), 1) & "'"
                    End If
                End If
            Next
            If Mid(gsOpeCod, 1, 6) = gContRepCaptacIngPagos Then
                Imprime763204 Me.tvOpe.SelectedItem.Text, Trim(txtFecha.Text), sFuente, Me.lstAge, gbBitCentral
            Else
                Imprime763205 Me.tvOpe.SelectedItem.Text, Trim(txtFechaDel.Text), Trim(txtFechaAl.Text), sFuente, lstAge, gbBitCentral
            End If
    
        'Movimientos Inusuales
        Case gContRepInusuales
             frmRepRiesgos.Show 1
             
        'Otros Reportes
        Case gProvContxCtasCont
             ProvContxCtasCont txtFechaDel, txtFechaAl
             
        Case gProvContxOpe
             ProvContxOperacion txtFechaDel, txtFechaAl
             
        Case gIntDevPFxFxAG
             InteresDevvengadoPFxFxAG txtFechaDel, txtFechaAl
             
        Case gIntDevCTSxFxAG
             InteresCTSxFxAG txtFechaDel, txtFechaAl
             
        Case gPlazoFijoRango
             PlazoFijoxRango Me.txtFecha, lsMoneda
             
        Case gRepInstPubliResgos
             RepInstPubliResgos txtFecha.Text
             
        Case gITFQuincena
             ITFQuincena txtFechaDel, txtFechaAl, 1
        'ALPA 20081007****************************************************************
        Case gITFQuincenaTipoCambioDiario
        '*****************************************************************************
            ITFQuincena txtFechaDel, txtFechaAl, 2
        'GITU 23-02-2009
        Case gGenArchivoDaot
            GeneraArchivoDAOT Me.txtAnio, txtNroUITs, txtTipCambio
        'End Gitu
        'MIOL 20120702, SEGUN RQ12122 ************************************************
        Case gGenPosCamDiaria
        frmRepBaseFormula.inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
        'END MIOL ********************************************************************
        'FRHU 20131205, SEGUN RQ13649 ************************************************
        Case gPropuestaDistribuUtilidad
        frmPropuestaDistriUtilidades.inicio Me.txtAnio, Me.cboSemestre.ListIndex + 1, Me.cboTipo.ListIndex + 1
        'FIN FRHU ********************************************************************
        'FRHU 20131209, SEGUN RQ13650 ************************************************
        Case gPropuestaCapitaliUtilidad
        ReporteAnexo02PropuestaCapitalizacionUtilidades Me.txtAnio, Me.cboSemestre.ListIndex + 1, Me.cboTipo.ListIndex + 1
        'FIN FRHU ********************************************************************
        'FRHU 20131219, SEGUN RQ13656 ************************************************
        Case gEstadoDeCambioPatrimonio
        ReporteEstadoCambioPatrimonio Me.txtAnio, Me.cboSemestre.ListIndex + 1
        'FIN FRHU ********************************************************************
        Case gPlazoFijoIntCash    'By Capi 04082008
             ImprimePlazoFijoIntCash lsMoneda
        
        'By Capi 08022008
        Case gCarteraDetxIntDSP
             ImprimeCarteraDetxIntDSP Val(txtTipCambio)
        Case gCarteraResxLineas
             ImprimeCarteraResxLineas Val(txtTipCambio)
        
        'End By
             
        'JEOM
        'Reportes Balance Contabilidad
        Case gPignoraticiosVencidos
             PignoraticiosVencidos
             
        Case gCarteraCreditos
             CarteraCreditos Me.txtFecha
        
        Case gCarteraInteres
             InteresCreditos Me.txtFecha
             
        Case gCreditosCastigados
             CreditosCastigados Me.txtFecha
        
        Case gInteresesDiferidos
             InteresesDiferidos Me.txtFecha
        'ALPA 20090309 descomentar gitu ********************************************
        Case gInteresesCastigados
            InteresCreditosCastigados Me.txtFecha
        '*********************************************************
        'JUEZ 20130116 ****************************
        Case gGastosCastigados
            GastosCreditosCastigados Me.txtFecha
        'END JUEZ *********************************
        
        Case gPignoraticiosVigentes
             PignoraticiosVigentes
        
        
        Case gCreditosCondonados
             CreditosCondonados Me.txtFecha
             
        Case gProvisionCreditos
             ProvisionCreditos
             
        Case gIntCreditosRefinanciados
             IntCreditosRefinanciados Me.txtFechaDel, Me.txtFechaAl
             
        Case gProvisionCartasFianzas
              ProvisionCartasFianzas
        
        Case gDetalleGarantias
             DetalleGarantias Me.txtFecha, txtTipCambio, lsMoneda
             
        'JEOM
        'Reportes Balance Planeamiento
        Case gCarteraCredPla
             CarteraCreditosPla Me.txtFecha
        
        Case gCredDesembolsosPla
             CredDesembolsosPla Me.txtFecha, Me.txtTipCambio
        
        Case gCarteraVencidaPla
             CarteraVencidaPla Me.txtFecha
        
        Case gCarteraRefinanciadaPla
             CarteraRefinanciadaPla Me.txtFecha
        
        Case gCarteraJudicialesPla
             CarteraJudicialesPla Me.txtFecha
             
        
        'By Capi 18122007 Para Planeamiento
        Case gCarteraRecupCapital
             SayCarteraRecupCapital Me.txtFecha
             
        '*** PEAC 20080915
         
        Case gNumCliCtasEnAhorrosCreditos
             SayNumCliCtasEnAhorrosCreditos Me.txtTipCambio
        Case gOpeRptInfoEstadColocBCRP 'EJVG20121113
            frmRptAnioMesGral.Ini gOpeRptInfoEstadColocBCRP, gsOpeDesc
        Case gCarteraDeAhorros 'FRHU20140121 RQ13826
            frmRptCarteraAhorros.Ini gCarteraDeAhorros, gsOpeDesc
        'JEOM
        'Reportes Balance Riesgos
        Case gCredVigentesRiesgos
             CreditosVigentes Me.txtFecha
             
        Case gCredRefinanciadosRiesgos
             CreditosRefinanciados Me.txtFecha

        Case gPlazoFijoRiesgos
             PlazoFijoRiesgos Me.txtFecha, lsMoneda
        'MIOL 20130209 RFC138-2012 ********
        Case gPPosCambRiesgos
             frmReportePosCambiaria.Show 1
        'END MIOL *************************
        
         'NAGL 20171005 ********************
        Case gVarCambBackTestRiesgos
             frmAnxVarCambYBackTesting.inicio
               
        Case gPosicionAfectaRiesgoCamb
             frmAnxCalculoRCambiario.inicio
             
        Case gRepSegTOSE
             frmReporteSegTOSE.inicio
        'END NAGL *************************
        
        'ANEXOS
        Case gContAnx02CredTpoGarantia 'Creditos Directos por Tipo de Garantia
            frmAnx02CreDirGarantia.GeneraAnx02CreditosTipoGarantia txtAnio, cboMes.ListIndex + 1, nVal(txtTipCambio), cboMes.Text, 1
        Case gContAnx03FujoCrediticio
            frmAnx02CreDirGarantia.GeneraAnx03FlujoCrediticioPorTipoCred gbBitCentral, txtAnio, cboMes.ListIndex + 1, nVal(txtTipCambio), cboMes.Text, 1
        
        Case 770036 ' BRGO
            generaAnexo4A_BienesAdjudRecup txtFecha
        Case 770037 ' BRGO
            generaAnexo4B_VentaBsAdjudRecup txtFecha
        
        Case gContAnx07N
            'frmAnexo7RiesgoInteres.Inicio True
            frmAnx7RiesgoTasaInteres.Show 1 'JACA 20110510
        Case gContAnx10DepColocaPer 'Depositos, Colocaciones y Persona por Oficinas
            frmAnx10DepColocPers.Show 0, frmReportes
        
        '''''''''''''''''''
         Case gContAnx09
             frmRepBaseFormula.inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
         
         Case gContAnx11MovDepsMonto
             
            'Anterior txt
            '------------
            ''''ImprimeAnx11 gbBitCentral, Me.tvOpe.SelectedItem.Text, CDate(txtFecha.Text), Me.lstAge, Val(TxtTipCamFijAnt.Text), Val(txtTipCamFij.Text), Val(txtTipCamCompraSBS.Text), Val(txtTipCamVentaSBS.Text)
            
            ImprimeAnx11xls gbBitCentral, Me.tvOpe.SelectedItem.Text, CDate(txtFecha.Text), Me.lstAge, Val(TxtTipCamFijAnt.Text), Val(txtTipCamFij.Text), Val(txtTipCamCompraSBS.Text), Val(txtTipCamVentaSBS.Text)
        
        'ANEXO 12-II
        Case "770120" 'JACA 20111017
            frmAnexo12_II.Show 1
            
        Case gContAnx13DepsSEscMonto
         
            'Anterior txt
            '------------
            ''''ImprimeAnx13 Me.tvOpe.SelectedItem.Text, Val(txtTipCamFij.Text), gbBitCentral
                
            ImprimeAnx13XLS Me.tvOpe.SelectedItem.Text, Val(Me.txtTipCambio.Text), gbBitCentral, txtFecha 'JAOR 20200725
            
        Case gContAnx13DepsSEscMonto_Nuevo

            'Nuevo
            '------------
            ImprimeAnx13XLS Me.tvOpe.SelectedItem.Text, Val(Me.txtTipCambio.Text), gbBitCentral, ldFecha

            
        Case gContAnx15A_Estad      'Informe Estadístico
            frmAnx15AEstadDia.ImprimeEstadisticaDiaria gsOpeCod, lsMoneda, txtFecha
        Case gContAnx15A_Efect      'Descomposición de Efectivo
            'frmAnx15AEfectivoCaja.ImprimeEfectivoCaja gsOpeCod, lsMoneda, txtFecha 'Comentado by NAGL 20180925
             frmAnx15AEfectivoCaja.ImprimeEfectivoCajaNew gsOpeCod, lsMoneda, txtFecha 'NAGL Según TIC1807210002
        Case gContAnx15A_Banco      'Consolidado Bancos
            frmAnx15AConsolBancos.ImprimeConsolidaBancos gsOpeCod, lsMoneda, txtFecha
        'ALPA20130930*************************
        Case gContAnx15A_Repor      'Anexo 15A
            'ALPA20130918
            'frmAnexo15.ImprimirAnexo15A gsOpeCod, lsMoneda, txtfecha, IIf(txtTipCambio.Text = "", "1", IIf(txtTipCambio.Text = "0", "1", CDbl(txtTipCambio.Text)))
            frmAnx15AReporteNew.ImprimeAnexo15A gsOpeCod, lsMoneda, txtFecha, IIf(txtTipCambio.Text = "", "1", IIf(txtTipCambio.Text = "0", "1", CDbl(txtTipCambio.Text)))
        Case gContAnx15B
            'Call ReporteAnexo15B(IIf(txtTipCambio.Text = "", "1", IIf(txtTipCambio.Text = "0", "1", CDbl(txtTipCambio.Text)))) comentado by NAGL
             frmAnx15BRatioCobertLiquidezNew.ImprimeAnexo15BRatioCoberLiqudz gsOpeCod, lsMoneda, txtFecha, IIf(txtTipCambio.Text = "", "1", IIf(txtTipCambio.Text = "0", "1", CDbl(txtTipCambio.Text))) 'NAGL 20170428
        Case gContAnx15C
            Call ReporteAnexo15C
        Case gContAnex15BMens
            frmAnx15B_PromedioMensRatioCobertLiqu.inicio 'NAGL ERS079-2017 20171226
            
        Case gContAnx16LiqVenc
            'frmAnexo16LiquidezVenc.Inicio False
            frmAnx16LiquidezPlazoVenc.inicio "770160", "Anexo 16: Cuadros de Liquidez por Plazos de Vencimiento"
        Case gContAnx16A
            'frmAnexo16ALiquidezVenc.Inicio False
            'Call ReporteAnexo16A(txtTipCambio.Text, txtFecha.Text) Comentado by NAGL
             frmFluctuacionesInvDPVAnexo16ANew.inicio gsOpeCod, txtFecha 'NAGL 20170428
        Case gContAnx16B
            'frmAnexo16LiquidezVenc.Inicio False
            frmAnx16LiquidezPlazoVenc.inicio "770162", "Anexo 16B: Simulación de Escenario de Estrés y Plan de Contingencia"
        Case gContAnx17A_FSD
            frmFondoSeguroDep.inicio txtFechaDel, txtFechaAl
        Case 770250
            Anexo6RFA Format(cboMes.ListIndex + 1, "00"), txtAnio, cboMes
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case gContAnx17CtasFuncionarios
            Imprime770175 ldFecha, gbBitCentral, Me.tvOpe.SelectedItem.Text
          Case gContAnx17ListadoFSD
             Imprime770130 gbBitCentral, Me.tvOpe.SelectedItem.Text, Val(txtTipCamFij.Text), txtFecha.Text
        Case gContAnx17ListadoGenCtas
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'REPORTES
        Case gRiesgoSBSA02A
             frmRepBaseFormula.inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo, "CAR", False
        Case gRiesgoSBSA02B
             frmRepBaseFormula.inicio Me.txtAnio, cboMes.ListIndex + 1, gsOpeCod, gsOpeDescHijo
        
        Case gRiesgoSBSA050
        
            'Anterior txt
            '------------
'            ImprimeRep05SBS cboMes.ListIndex + 1, txtAnio
        
            ImprimeRep05SBSXls cboMes.ListIndex + 1, txtAnio
        
        Case gContRep06Crediticio
             frmRep6Crediticio.ImprimeAnexo6Crediticio gsOpeCod, ldFecha, Me.lstAge
             
        Case gContRepSpreadFinanciero 'JUCS 10032017
            frmRepSpreadFinanciero.ImprimeSpreadFinanciero gsOpeCod, ldFecha, Me.lstAge 'JUCS 10032017
        
        Case gRepCreditosIncumplidos 'Reporte 14
             frmRep14CredIncump.ImprimeReporte14 gsOpeCod, ldFecha, Val(txtTipCambio.Text)
             
        
        Case gRiesgoSBSA190
            Call frmRiesgosGrupEconom.inicio(gRiesgoSBSA190, gdFecSis)
            frmRiesgosGrupEconom.Show 1
            
        Case gRiesgoSBSA191
            Call frmRiesgosGrupEconom.inicio(gRiesgoSBSA191, gdFecSis)
            frmRiesgosGrupEconom.Show 1
            
        Case gRiesgoSBSA200
            Call frmRiesgosGrupEconom.inicio(gRiesgoSBSA200, gdFecSis)
            frmRiesgosGrupEconom.Show 1

        Case gRiesgoSBSA201
            Call frmRiesgosGrupEconom.inicio(gRiesgoSBSA201, gdFecSis)
            frmRiesgosGrupEconom.Show 1
        
        Case gRiesgoSBSA210 'Modify GITU 20-03-2009
            'Call frmRiesgosGrupEconom.Inicio(gRiesgoSBSA210, gdFecSis)
            'frmRiesgosGrupEconom.Show 1
            'Call frmReporte21
            frmReporte21.Show 1
        Case 780211 'Add by Gitu 25-03-2009
            'generaReporte25 (Me.txtFecha)
            generaReporte25 cboMes.ListIndex + 1, txtAnio.Text '** PASI 20130914
        Case 780130
            frmReporte16.Show 1
             
        Case gPatrEfecAjxInfla
            'ALPA 20090725***********************************************************************
            'ImprimeRep3SBS_PatrimEfectAjustxInfl 760112, CInt(txtAnio.Text), cboMes.ListIndex + 1
            Dim nMes As Integer
            nMes = cboMes.ListIndex + 1
            ImprimeRep3SBS_PatrimEfectAjustxInfl 760112, txtAnio.Text, IIf(Len(Trim(nMes)) = 1, "0" & CStr(Trim(nMes)), CStr(Trim(nMes))), txtFecha.Text
            '************************************************************************************
        Case gRiesgoSBS2B1
            Call Reporte2B1RiesgodeMercado
'ALPA 20090930**************
        Case gRiesgoSBS2D
            Call Reporte2DBasilea
'****************************
        Case gContAnx18
            ImprimeAnx18SBS_InmMovEquip 760181, CInt(txtAnio.Text), cboMes.ListIndex + 1
        Case gContAnx24
            ImprimeAnx24_CTS gbBitCentral, cboMes.ListIndex + 1, Val(txtAnio.Text), Val(txtTipCambio.Text)
        Case 780220
            Set Progress = New clsProgressBar
            Progress.ShowForm Me
            Progress.Max = 100
            Progress.Progress 10, "Procesando Reporte de Estadistica Adelantada....."
            Set oEstadistica = New NEstadistica
            Call oEstadistica.MuestraEstadisticaAdelantada(gdFecSis, Me.txtTipCambio)
            Progress.Max = 100
            Progress.Progress 100, "Procesando Reporte de Estadistica Adelantada....."
            Set oEstadistica = Nothing
            Progress.CloseForm Me
            MsgBox "Reporte Generado Satisfactoriamente.....", vbInformation, "Aviso"
            
        'PEAC 20201125
        Case 780350 ' REPORTE 35
            'ImprimeAnx18SBS_InmMovEquip 760181, CInt(txtAnio.Text), cboMes.ListIndex + 1
            Reporte35 cboMes.ListIndex + 1, CInt(txtAnio.Text)
            
            
        'By Capi 18122007 Para Planeamiento
        Case gCarteraRecupCapital
             SayCarteraRecupCapital Me.txtFecha
        'ALPA 20090722********
'        Case gRiesgoSBSA212
'            Call Reporte15Tesoreria
        '*********************
        Case 763517
            'Call ReporteReparticionGastos
            Call ReporteReparticionGastosNew 'EJVG20130124
        'MAVM 20090810 **************
        Case gFinancFinanciamientoRecibido
            If psCtaCont <> "" Then
                Call ReporteSBS12Financ(ldFecha, Arrays2)
                psCtaCont = ""
                Erase Arrays2
            Else
                MsgBox "Debe Elegir las Ctas Contables de las EF", vbCritical
            End If
        'PASIERS024-2015************************************
        Case 780113
            frmRepSBS13.Show 1
        'end PASI
        
        Case 780114 'gFinancObligacionesExterior
            If psCtaCont <> "" Then
                Call ReporteSBS14Financ(ldFecha, Val(txtTipCambio.Text), Arrays2)
                psCtaCont = ""
                Erase Arrays2
            Else
                MsgBox "Debe Elegir las Ctas Contables de las EF", vbCritical
            End If
        '****************************
        Case 763516
            'Call ReporteResumenAhorros
            Call ReporteResumenAhorrosNew 'NAGL 20180907 ACTA 111 - 2018
        'ALPA 20100413*****************************************
        Case 763518
            Call ImprimirListadoDeCalificacionDeCartera
        Case 763519
            Call ImprimirDetalleDeCalificacionDeCartera
        '******************************************************
        'ALPA 20101229*****************************************
        Case gVerificacionDevenSusp
            Call ReporteVerificacionInteresDevenSusp
        '******************************************************
        Case gVerificacionDiferidos
            Call ReporteVerificacionInteresesDiferidos
        'ALPA 20110813*****************************************
        Case OpeCGRepAdeudadoCalendarioMN, OpeCGRepAdeudadoCalendarioME
            Call ReporteAdeudadosCalendario
        
        Case OpeCGRepAdeudadoCalendarioVigenteMN, OpeCGRepAdeudadoCalendarioVigenteME
            Call ReporteAdeudadosCalendarioVigente
        
        Case OpeCGRepAnalisisDeCtaMN, OpeCGRepAnalisisDeCtaME
            Call ReporteAdeudados
        '******************************************************
        'ALPA 20130708*****************************************
        Case gReporteComCartFi
            Call ReporteCartaFianza
        '*****************************************************
        'MIOL 20130814 ***************************************
        Case gGenRepProyEjec
            frmRepGastosProyEjec.Show 1
        '*****************************************************
        'ALPA 20141219****************************************
        Case 763525
            Call ReporteCuentasuInactivasRestringidas(CDate(txtFecha.Text))
        '*****************************************************
        'NAGL 202008*****************************************
        Case 763528
            Call GeneraReporteDiferidosAmpliados_NoAmpliados(CDate(txtFecha.Text))
        'NAGL 202008 Según ACTA N°063-2020*******************
        
        'PASI20160404*******
        Case gGenRepProvInd
            frmRepProvisionIndiv.Show 1
        'end PASI***********
        
        'REPORTES DE ACTIVOS FIJOS
        'FRHU 20131211 RQ13651
        Case gMovDepreAcumuladaIntangible
            RepMovDepreAcumuladaIntangible txtAnio, cboMes.ListIndex + 1
        Case gMovActivoFijoIntangible
            RepMovActivoFijoIntangible txtAnio, cboMes.ListIndex + 1
        Case gCtoIntangibleOtrosAcivosAmo
            RepCtoIntangibleOtrosActivos Me.txtAnio, Me.cboSemestre.ListIndex + 1
        Case gCtoInmuebleMaquinariaEquipo
            RepCtoInmuebleMaquinariayEquipos Me.txtAnio, Me.cboSemestre.ListIndex + 1
        Case gCtoDepreciacionDeActivoFijo
            RepCtoDepreciacionActivosFijos Me.txtAnio, Me.cboSemestre.ListIndex + 1
        'FIN FRHU 20131211
        'ALPA 2014040227
        Case gContRepGPSpot
            Call ReporteGananciaPerdidaSPOT(txtFechaDel.Text, txtFechaAl.Text)
        'YIHU RS015-2015 20150421**************************************
        Case gContRepAsientoxCuenta
            frmRepAsientoxCuenta.Show 1
        '**************************************************************
        'PASI20161005 ERS0532016
        Case "764100"
            frmDocAutorizado.Show 1
        '**************************************************************
    End Select
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reporte "
                Set objPista = Nothing
                '****
Exit Sub
cmdGenerarErr:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub RecibirArray(ByVal Arrays As Variant)
    Dim i As Integer
    For i = 0 To UBound(Arrays)
        If Arrays(i, 0) <> "" Then
            If UBound(Arrays) = 0 Then
                psCtaCont = Arrays(i, 0)
            Else
                If i = 0 Then
                    psCtaCont = Arrays(i, 0)
                Else
                    psCtaCont = psCtaCont & "," & Arrays(i, 0)
                End If
            End If
        End If
    Next i
    ReDim Arrays2(UBound(Arrays), 3)
    Arrays2 = Arrays
End Sub

'MAVM 20091008 *********
Public Sub ReporteSBS12Financ(ByVal sdFecha As Date, ByVal Arra As Variant)
    Dim rs As ADODB.Recordset
    Dim oCtaCont As New DCtaCont
    Set rs = New ADODB.Recordset
    
    Dim lsArchivo As String
    Dim lbExcel As Boolean
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Set rs = oCtaCont.DarReporteSBS12_FinanciamientoRecibido(sdFecha, psCtaCont)
    
    If rs.BOF Then
        Set rs = Nothing
        MsgBox "No Existen Datos", vbInformation, Me.Caption
        Exit Sub
    End If

    lsArchivo = App.path & "\SPOOLER\" & "RepSBS12_FinancRecib_" & Format(sdFecha, "mmyyyy") & ".XLS"

    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbExcel Then
        ExcelAddHoja "RepSBS12_FinancRecib", xlLibro, xlHoja1
        GeneraReporteSBS12_FinancRecib sdFecha, rs, Arra, xlHoja1
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        If lsArchivo <> "" Then
            CargaArchivo lsArchivo, App.path & "\SPOOLER\"
        End If
    End If
End Sub

Private Sub GeneraReporteSBS12_FinancRecib(ByVal pdFecha As Date, ByVal R As ADODB.Recordset, ByVal pArray As Variant, Optional xlHoja1 As Excel.Worksheet)
    Dim oBarra As clsProgressBar
    Dim i, j As Integer
    Dim lnFila As Integer
    Dim sCodPersona As String
    Dim sMoneda As String
    Dim rsAdeudados As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsPagare As ADODB.Recordset
    Dim oCta As DCtaCont
    
    Set oBarra = New clsProgressBar
    oBarra.ShowForm frmReportes
    oBarra.Max = 2
    oBarra.Progress 0, "REPORTE SBS 12: FINANC RECIB", "Cargando Datos", "", vbBlue
    
    lnFila = CabeceraReporte(pdFecha, "DESCRIPCION DE FINANCIAMIENTOS RECIBIDOS PARA APOYO A LA PEQUEÑA Y MICROEMPRESA*", "REPORTE 12", 9, xlHoja1)
    
    lnFila = lnFila - 13
    
    Set oCta = New DCtaCont
    Set rsAdeudados = oCta.DarEntidadesFinac
    
    Do While Not rsAdeudados.EOF
        For i = 0 To UBound(pArray)
            If rsAdeudados!vEntidad = pArray(i, 2) Then
                Do While Not R.EOF
                    Dim Total As Double
                    If pArray(i, 0) = R!cCtaContCod Then
                        Total = Total + R!Total90
                        sMoneda = IIf(IsNull(Mid(R!cCtaContCodOriginal, 3, 1)), "", Mid(R!cCtaContCodOriginal, 3, 1))
                        If sCodPersona = "" Then
                            sCodPersona = pArray(i, 3) 'DevolverCodPersona(pArray(i, 0), pArray(i, 1))
                        End If
                        Exit Do
                    End If
                    R.MoveNext
                Loop
            End If
                    
            If Total <> 0 And i = UBound(pArray) Then
                xlHoja1.Cells(lnFila + j, 3) = rsAdeudados!vEntidad
                xlHoja1.Cells(lnFila + j, 4) = sMoneda
                xlHoja1.Cells(lnFila + j, 5) = Format(Round(Format(Total, "##0,")), "####,000")
                
                Dim TotalI As Double
                Dim contador As Integer
                Dim iDias As Integer
                Dim nMayor As Integer
                Set rsTotal = oCta.DarTotales_ReporteSBS(sCodPersona, pdFecha)
                        Do While Not rsTotal.EOF
                            TotalI = TotalI + rsTotal!nCtaIFIntValor
                            If rsTotal!FF > nMayor Then
                                nMayor = rsTotal!FF
                                iDias = nMayor
                            End If
                            'iDias = iDias + rsTotal!FF
                            contador = contador + 1
                            rsTotal.MoveNext
                        Loop
                        
                If TotalI <> 0 Then
'                    xlHoja1.Cells(lnFila + J, 6) = Round(iDias / 360) & " a"
'                    xlHoja1.Range("F" & 12 & ":" & "F" & 18).NumberFormat = "@"
'
'                    xlHoja1.Cells(lnFila + J, 7) = Round(TotalI / contador, 2)
'                    xlHoja1.Range("G" & 12 & ":" & "G" & 18).NumberFormat = "#,##0.00"
                    xlHoja1.Range("F" & 12 & ":" & "F" & 19).NumberFormat = "@"
                    xlHoja1.Cells(lnFila + j, 6) = Round(iDias / 360) & " años"
                    
                    iDias = 0
                    nMayor = 0
                    
                    xlHoja1.Cells(lnFila + j, 7) = Round(TotalI / contador, 2)
                    xlHoja1.Range("G" & 12 & ":" & "G" & 19).NumberFormat = "#,##0.00"

                End If
                
                'Foncodes
                If rsAdeudados!vEntidad = "FONCODES" Then
                    xlHoja1.Cells(lnFila + j, 7) = 7
                    xlHoja1.Range("G" & 12 & ":" & "G" & 18).NumberFormat = "#,##0.00"
                End If
                
                xlHoja1.Cells(lnFila + j, 8) = IIf(rsAdeudados!vEntidad = "BID", "MICROEMPRESAS", "MICROCREDITOS")
                xlHoja1.Cells(lnFila + j, 9) = "PRESTAMO"
                Total = 0
                sMoneda = ""
                contador = 0
                sCodPersona = ""
                TotalI = 0
                j = j + 1
                R.MoveFirst
            End If
        Next i
        rsAdeudados.MoveNext
    Loop

    xlHoja1.Cells(31, 4) = "Gerente General"
    xlHoja1.Cells(31, 8) = "Funcionario Autorizado"
    xlHoja1.Range("D30:E30").MergeCells = True
    xlHoja1.Range("D31:E31").MergeCells = True
    xlHoja1.Range("D31:E31").HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(30, 4), xlHoja1.Cells(31, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(30, 8), xlHoja1.Cells(31, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
End Sub

Private Function DevolverCodPersona(ByVal sCtaContCod As String, ByVal sNombre As String) As String
    Dim oCta As DCtaCont
    Dim rsCodPersona As ADODB.Recordset
        
    Set oCta = New DCtaCont
    Set rsCodPersona = oCta.DarCodPersona_ReporteSBS(sCtaContCod, , 1)
    
    If rsCodPersona.RecordCount <> 0 Then
        DevolverCodPersona = rsCodPersona!cPersCod
    Else
        Set rsCodPersona = oCta.DarCodPersona_ReporteSBS(, IIf(InStr(1, sNombre, " ") = "0", Mid(sNombre, 1, Len(sNombre)), Mid(sNombre, 1, InStr(1, sNombre, " "))), 2)
        If rsCodPersona.RecordCount <> 0 Then
            DevolverCodPersona = rsCodPersona!cPersCod
        Else
            DevolverCodPersona = ""
        End If
    End If
End Function

Private Function CabeceraReporte(pdFecha As Date, psTitulo As String, psReporte As String, pnCols As Integer, Optional xlHoja1 As Excel.Worksheet) As Integer
Dim lnFila As Integer
Dim i As Integer

    xlHoja1.Range("A1:R100").Font.Size = 9

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(8, pnCols)).Font.Bold = True
    
    lnFila = 2
    xlHoja1.Cells(lnFila, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Cells(lnFila, 10) = "REPORTE 12"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A."
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 10) = "CODIGO: 109"

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = psTitulo
    
    xlHoja1.Cells(lnFila, 3).Font.Size = 11
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Cells.Font.Color = vbBlue
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 6) = "(En miles de " & StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016
    xlHoja1.Cells(lnFila, 6).Font.Size = 8
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 10) = Format(pdFecha, "mmm yyyy")
    xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(lnFila, 10)).Cells.Interior.Color = vbYellow
    
    lnFila = lnFila + 1
    
    lnFila = lnFila + 2
    
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 1, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 1, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 1, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 9, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 3), xlHoja1.Cells(lnFila + 9, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 9, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    lnFila = lnFila - 1
    xlHoja1.Cells(lnFila, 3) = "Entidad que Otorga el"
    xlHoja1.Cells(lnFila + 1, 3) = "Financiamiento"
    xlHoja1.Range("C9:C9").ColumnWidth = 20
    
    xlHoja1.Cells(lnFila, 4) = "Tipo de"
    xlHoja1.Cells(lnFila + 1, 4) = "Moneda"
    xlHoja1.Range("D9:D9").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila, 5) = "Monto"
    xlHoja1.Cells(lnFila + 1, 5) = "(Equivalente en"
    xlHoja1.Cells(lnFila + 2, 5) = StrConv(gcPEN_PLURAL, vbProperCase) & ")" 'marg ers044-2016
    xlHoja1.Range("E9:E9").ColumnWidth = 15
    
    xlHoja1.Cells(lnFila + 1, 6) = "Plazo"
    xlHoja1.Range("F9:F9").ColumnWidth = 15
    
    xlHoja1.Cells(lnFila + 1, 7) = "Tasa de"
    xlHoja1.Cells(lnFila + 2, 7) = "Interés % (1)"
    xlHoja1.Range("G9:G9").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila, 8) = "Financiamiento"
    xlHoja1.Cells(lnFila + 1, 8) = "otorgado para"
    xlHoja1.Cells(lnFila + 2, 8) = ":(2)"
    xlHoja1.Range("H9:H9").ColumnWidth = 20
    
    xlHoja1.Cells(lnFila + 1, 9) = "Observaciones"
    xlHoja1.Cells(lnFila + 2, 9) = ":(3)"
    xlHoja1.Range("I9:I9").ColumnWidth = 15
    
    lnFila = lnFila + 11
    xlHoja1.Cells(lnFila, 3) = "*Corresponde a aquellos fondos comprendidos para apoyo a la pequeña y microempresa (mediante contratos u otros mecanismos)"
    xlHoja1.Range("C21:C21").Font.Size = 8
    
    lnFila = lnFila + 2
    xlHoja1.Cells(lnFila, 3) = "(1)La tasa de interés será efectiva anual; incluye comisiones, seguros y otros. Esta tasa efectia sera referida al tipo de moneda del financiamiento"
    xlHoja1.Range("C22:C22").Font.Size = 8
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "(2)Deberá señalarse si el financiamiento ha sido otorgado para algún fin específico (indicándolo) o si es libre de disposición."
    xlHoja1.Range("C23:C23").Font.Size = 8
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "(3)Deberá especificar la modalidad de los fondos recibidos: Fideicomiso, Fondo Revolvente, Préstamo a Plazo, etc."
    xlHoja1.Range("C24:C24").Font.Size = 8
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "(4)Si la entidad que otorga el financiamiento recibió dichos fondos de otra entidad a su vez, deberá registrar aquí de que entidad lo recibió"
    xlHoja1.Range("C25:C25").Font.Size = 8
    
    CabeceraReporte = lnFila
End Function

Public Sub ReporteSBS14Financ(ByVal sdFecha As Date, ByVal nTipCambio As Double, ByVal Arra As Variant)
    Dim rs As ADODB.Recordset
    Dim oCtaCont As New DCtaCont
    Set rs = New ADODB.Recordset
    
    Dim lsArchivo As String
    Dim lbExcel As Boolean
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Set rs = oCtaCont.DarReporteSBS12_FinanciamientoRecibido(sdFecha, psCtaCont)
    
    If rs.BOF Then
        Set rs = Nothing
        MsgBox "No Existen Datos", vbInformation, Me.Caption
        Exit Sub
    End If

    lsArchivo = App.path & "\SPOOLER\" & "RepSBS14_ObligacExter_" & Format(sdFecha, "mmyyyy") & ".XLS"

    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbExcel Then
        ExcelAddHoja "RepSBS14_ObligacExter", xlLibro, xlHoja1
        GeneraReporteSBS14_ObligacExter sdFecha, nTipCambio, Arra, rs, xlHoja1 ', ,
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        If lsArchivo <> "" Then
            CargaArchivo lsArchivo, App.path & "\SPOOLER\"
        End If
    End If
End Sub

Private Sub GeneraReporteSBS14_ObligacExter(ByVal pdFecha As Date, ByVal nTipCambio As Double, ByVal pArray As Variant, ByVal R As ADODB.Recordset, Optional xlHoja1 As Excel.Worksheet)
    Dim oBarra As clsProgressBar
    Dim i, j, k As Integer
    Dim lnFila As Integer
    Dim sCodPersona As String
    Dim sMoneda As String
    Dim rsAdeudados As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsPagare As ADODB.Recordset
    Dim oCta As DCtaCont
    Dim SumaTotal As Double
    Dim TotalI As Double
    Dim contador As Integer
    
    'Totales
    Dim MontoAutorizado As Double
    Dim CapitalTrabajo As Double
    Dim Dias30 As Double
    Dim Dias90 As Double
    Dim Dias180 As Double
    Dim Dias270 As Double
    Dim Dias360 As Double
    Dim DiasMas360 As Double
    Dim DiasTotalA As Double
    
    Dim DiasTotalB As Double
    Dim DiasTotalAB As Double
    Dim TasaInteres As Double
    
    Set oBarra = New clsProgressBar
    oBarra.ShowForm frmReportes
    oBarra.Max = 2
    oBarra.Progress 0, "REPORTE SBS 14: OBLIGAC EXTERIOR", "Cargando Datos", "", vbBlue
    
    lnFila = CabeceraReporte14(pdFecha, nTipCambio, "OBLIGACIONES CON EL EXTERIOR", "REPORTE 14", 30, xlHoja1)
    
    lnFila = lnFila + 3
    
    Set oCta = New DCtaCont
    Set rsAdeudados = oCta.DarEntidadesFinac
    
    Do While Not rsAdeudados.EOF
        For i = 0 To UBound(pArray)
            If rsAdeudados!vEntidad = pArray(i, 2) Then
                Do While Not R.EOF
                    Dim Total As Double
                    If pArray(i, 0) = R!cCtaContCod Then
                        Total = Total + R!Total90
                        sMoneda = IIf(IsNull(Mid(R!cCtaContCodOriginal, 3, 1)), "", Mid(R!cCtaContCodOriginal, 3, 1))
                        If sCodPersona = "" Then
                            sCodPersona = DevolverCodPersona(pArray(i, 0), pArray(i, 1))
                        End If
                        Exit Do
                    End If
                    R.MoveNext
                Loop
            End If
            
            If Total <> 0 And i = UBound(pArray) Then
                Total = Total / nTipCambio
                xlHoja1.Cells(lnFila + j, 3) = rsAdeudados!vEntidad
                xlHoja1.Cells(lnFila + j, 4) = IIf(rsAdeudados!vEntidad = "BID", "BOIDCOB1XXX", "") 'Codigo
                xlHoja1.Cells(lnFila + j, 5) = "0"
                
                If rsAdeudados!vEntidad = "AECI" Then
                    xlHoja1.Cells(lnFila + j, 6) = "España"
                End If
                
                If rsAdeudados!vEntidad = "BID" Then
                    xlHoja1.Cells(lnFila + j, 6) = "EE.UU"
                End If
                
                If rsAdeudados!vEntidad = "SYMBIOTICS" Then
                    xlHoja1.Cells(lnFila + j, 6) = "Luxemburgo"
                End If
                
                xlHoja1.Cells(lnFila + j, 7) = IIf(sMoneda = "2", "US$", "N.S.")
                xlHoja1.Cells(lnFila + j, 8) = Format(Total, "####,000.00")
                xlHoja1.Cells(lnFila + j, 11) = Format(Total, "####,000.00")
                xlHoja1.Cells(lnFila + j, 30) = Format(Total, "####,000.00")
                
                MontoAutorizado = MontoAutorizado + Total
                CapitalTrabajo = CapitalTrabajo + Total
                
                If rsAdeudados!vEntidad = "BID" Then
                Dim iDias As Integer
                Dim rsCtaPendiente As ADODB.Recordset

                Set rsPagare = oCta.DarPagarexCodPersona(sCodPersona)
                Do While Not rsPagare.EOF
                    Set rsCtaPendiente = oCta.DarCuotasPendientes(sCodPersona, rsPagare!cCtaIfCod)
                        xlHoja1.Cells(lnFila + j, 12) = "0.00"
                        xlHoja1.Cells(lnFila + j, 13) = "0.00"
                        xlHoja1.Cells(lnFila + j, 14) = "0.00"
                        xlHoja1.Cells(lnFila + j, 15) = "0.00"
                        xlHoja1.Cells(lnFila + j, 16) = "0.00"
                        xlHoja1.Cells(lnFila + j, 17) = "0.00"
                        
                        Do While Not rsCtaPendiente.EOF
                            Dim Cant As Integer

                            If DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) >= 0 And DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) <= 30 Then
                                xlHoja1.Cells(lnFila + j, 12) = Format(Total / rsCtaPendiente.RecordCount, "####,000.00")
                                Dias30 = Dias30 + (Total / rsCtaPendiente.RecordCount)
                            End If

                            If DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) >= 31 And DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) <= 90 Then
                                xlHoja1.Cells(lnFila + j, 13) = Format(Total / rsCtaPendiente.RecordCount, "####,000.00")
                                Dias90 = Dias90 + (Total / rsCtaPendiente.RecordCount)
                                
                            End If

                            If DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) >= 91 And DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) <= 180 Then
                                xlHoja1.Cells(lnFila + j, 14) = Format(Total / rsCtaPendiente.RecordCount, "####,000.00")
                                Dias180 = Dias180 + (Total / rsCtaPendiente.RecordCount)
                            End If

                            If DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) >= 181 And DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) <= 270 Then
                                xlHoja1.Cells(lnFila + j, 15) = Format(Total / rsCtaPendiente.RecordCount, "####,000.00")
                                Dias270 = Dias270 + (Total / rsCtaPendiente.RecordCount)
                            End If

                            If DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) >= 271 And DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) <= 360 Then
                                xlHoja1.Cells(lnFila + j, 16) = Format(Total / rsCtaPendiente.RecordCount, "####,000.00")
                                Dias360 = Dias360 + (Total / rsCtaPendiente.RecordCount)
                            End If

                            If DateDiff("d", pdFecha, rsCtaPendiente!dVencimiento) > 360 Then
                                xlHoja1.Cells(lnFila + j, 17) = Format(Total / rsCtaPendiente.RecordCount, "####,000.00")
                                DiasMas360 = DiasMas360 + (Total / rsCtaPendiente.RecordCount)
                            End If
                            rsCtaPendiente.MoveNext
                        Loop
                    rsPagare.MoveNext
                Loop
                
                Else
                    xlHoja1.Cells(lnFila + j, 17) = Format(Total, "####,000.00")
                    DiasMas360 = DiasMas360 + Total
                End If
                
                xlHoja1.Cells(lnFila + j, 18) = Format(xlHoja1.Cells(lnFila + j, 12) + xlHoja1.Cells(lnFila + j, 13) + xlHoja1.Cells(lnFila + j, 14) + xlHoja1.Cells(lnFila + j, 15) + xlHoja1.Cells(lnFila + j, 16) + xlHoja1.Cells(lnFila + j, 17), "####,000.00")
                DiasTotalA = DiasTotalA + xlHoja1.Cells(lnFila + j, 18)

                Set rsTotal = oCta.DarTotales_ReporteSBS(sCodPersona, pdFecha)
                        Do While Not rsTotal.EOF
                            TotalI = TotalI + rsTotal!nCtaIFIntValor
                            TasaInteres = TasaInteres + rsTotal!nCtaIFIntValor
                            contador = contador + 1
                            rsTotal.MoveNext
                        Loop

                If TotalI <> 0 Then
                    xlHoja1.Cells(lnFila + j, 19) = Round(TotalI / contador, 2)
                    xlHoja1.Range("S" & 11 & ":" & "S" & 18).NumberFormat = "#,##0.00"
                End If
                
                Total = 0
                sMoneda = ""
                sCodPersona = ""
                TotalI = 0
                contador = 0
                j = j + 1
                R.MoveFirst
            End If
            
        Next i
        rsAdeudados.MoveNext
    Loop
    
    xlHoja1.Cells(17, 8) = Format(MontoAutorizado, "####,000.00")
    xlHoja1.Cells(17, 11) = Format(CapitalTrabajo, "####,000.00")
    xlHoja1.Cells(17, 30) = Format(CapitalTrabajo, "####,000.00")
    
    xlHoja1.Cells(17, 12) = Format(Dias30, "####,000.00")
    xlHoja1.Cells(17, 13) = Format(Dias90, "####,000.00")
    xlHoja1.Cells(17, 14) = Format(Dias180, "####,000.00")
    xlHoja1.Cells(17, 15) = Format(Dias270, "####,000.00")
    xlHoja1.Cells(17, 16) = Format(Dias360, "####,000.00")
    xlHoja1.Cells(17, 17) = Format(DiasMas360, "####,000.00")
    
    xlHoja1.Cells(17, 18) = Format(DiasTotalA, "####,000.00")
    'xlHoja1.Cells(17, 19) = Format(TasaInteres, "####,000.00")
    
    lnFila = lnFila + 10
    xlHoja1.Cells(lnFila, 3) = "* Se deben reportar las obligaciones con el exterior provenientes de lineas, préstamos y sobregiros, sin considerar los intereses y gastos por pagar devengados por dichas obligaciones."
    xlHoja1.Range("C21:C21").Font.Size = 8
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "1/Consignar el nombre de la entidad registrado en la base de datos S.W.I.F.T. En el caso de entidades no registradas en la base de sdatos S.W.I.F.T. Consignar el nombre completo de la entidad."
    xlHoja1.Range("C22:C22").Font.Size = 8
    
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "2/Señalar el código B.I.C. Registrado para la entidad en la base de datos S.W.F.T. En el caso de entidades que carecen de código B.I.C. por no estar registardas en la base de datos S.W.I.F.T.,   consignar el código otros. Este código sustituirá el código B.I.C. Hasta que la entidad se registre en la base de datos del S.W.I.F.T."
    xlHoja1.Range("C23:C23").Font.Size = 8

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "3/ Considere (F) si es entidad financiera, (O) si es organismo internacional y (N) en otro caso."
    xlHoja1.Range("C24:C24").Font.Size = 8

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "4/ El monto total de obligaciones directas al tipo de cambio al cierre del mes de reporte debe ser igual al total de adeudos y obligaciones financieras con el exterior del Balance de Comprobación (2404+2407+2405+2604+2605+2607)"
    xlHoja1.Range("C25:C25").Font.Size = 8

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 3) = "5/ Considere la tasa de Interés promedio de cada obligacion directa y no el spread sobre la LIBOR.Para el total,consignar el promedio ponderado por el total de obligaciones directas."
    xlHoja1.Range("C26:C26").Font.Size = 8

    xlHoja1.Cells(31, 4) = "Gerente General"
    xlHoja1.Cells(31, 8) = "Contador General"
    xlHoja1.Range("D30:E30").MergeCells = True
    xlHoja1.Range("D31:E31").MergeCells = True
    xlHoja1.Range("D31:E31").HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(30, 4), xlHoja1.Cells(31, 5)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(30, 8), xlHoja1.Cells(31, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
End Sub

Private Function CabeceraReporte14(pdFecha As Date, nTipCambio As Double, psTitulo As String, psReporte As String, pnCols As Integer, Optional xlHoja1 As Excel.Worksheet) As Integer
Dim lnFila As Integer
Dim i As Integer

    xlHoja1.Range("A1:R100").Font.Size = 9

    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(7, pnCols)).Font.Bold = True
    
    lnFila = 2
    xlHoja1.Cells(lnFila, 28) = "REPORTE 14"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A."
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 28) = "CODIGO: 109"
    xlHoja1.Cells(lnFila, 28) = "TC " & nTipCambio

    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 5) = psTitulo
    
    xlHoja1.Cells(lnFila, 5).Font.Size = 11
    
    xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(lnFila, 3)).Cells.Font.Color = vbBlue
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 6) = "Al " & Mid(Format(pdFecha, "Long Date"), InStr(1, Format(pdFecha, "Long Date"), ",") + 2, Len(Format(pdFecha, "Long Date")))
    xlHoja1.Cells(lnFila, 6).Font.Size = 8
    
    lnFila = lnFila + 1
    
    lnFila = lnFila + 2
    
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 1, pnCols)).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 1, pnCols)).VerticalAlignment = xlVAlignCenter
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 1, pnCols)).WrapText = True
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 9, pnCols)).BorderAround xlContinuous, xlMedium
    xlHoja1.Range(xlHoja1.Cells(lnFila + 1, 3), xlHoja1.Cells(lnFila + 9, pnCols)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnFila - 1, 3), xlHoja1.Cells(lnFila + 9, pnCols)).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    lnFila = lnFila - 1
    xlHoja1.Cells(lnFila, 3) = "Entidad "
    xlHoja1.Cells(lnFila + 1, 3) = "1/"
    xlHoja1.Range("C9:C9").ColumnWidth = 15
    
    xlHoja1.Cells(17, 3) = "TOTAL"
    xlHoja1.Cells(18, 3) = "TOTAL (N.S.)"

    xlHoja1.Cells(lnFila, 4) = "Codigo"
    xlHoja1.Cells(lnFila + 1, 4) = "2/"
    xlHoja1.Range("D9:D9").ColumnWidth = 13

    xlHoja1.Cells(lnFila, 5) = "Tipo de"
    xlHoja1.Cells(lnFila + 1, 5) = "Entidad"
    xlHoja1.Cells(lnFila + 2, 5) = "3/"
    xlHoja1.Range("E9:E9").ColumnWidth = 10

    xlHoja1.Cells(lnFila + 1, 6) = "Pais de"
    xlHoja1.Cells(lnFila + 2, 6) = "Origen"
    xlHoja1.Range("F9:F9").ColumnWidth = 10

    xlHoja1.Cells(lnFila + 1, 7) = "Moneda"
    xlHoja1.Range("G9:G9").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 1, 8) = "Monto Autorizado (equiv. En US$)"
    xlHoja1.Range("H9:H9").ColumnWidth = 15
    
    xlHoja1.Range("I9:K9").MergeCells = True
    
    xlHoja1.Cells(lnFila + 1, 9) = "Por Destino"
    xlHoja1.Range("I9:K9").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 1, 12) = "Por a vencer"
    xlHoja1.Range("L9:R9").ColumnWidth = 80
    
    xlHoja1.Range("L9:R9").MergeCells = True
    
    xlHoja1.Cells(lnFila + 2, 9) = "Exportacion"
    xlHoja1.Range("I10:I10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 10) = "Importacion"
    xlHoja1.Range("J10:J10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 11) = "Cap. de Trabajo"
    xlHoja1.Range("K10:K10").ColumnWidth = 15
    
    xlHoja1.Cells(lnFila + 2, 12) = "0 - 30 D"
    xlHoja1.Range("L10:L10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 13) = "31 - 90 D"
    xlHoja1.Range("M10:M10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 14) = "91 - 180 D"
    xlHoja1.Range("N10:N10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 15) = "181 - 270 D"
    xlHoja1.Range("O10:O10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 16) = "271 - 360 D"
    xlHoja1.Range("P10:P10").ColumnWidth = 10
    
    xlHoja1.Cells(lnFila + 2, 17) = "mas de 360 D"
    xlHoja1.Range("Q10:Q10").ColumnWidth = 13
    
    xlHoja1.Cells(lnFila + 2, 18) = "TOTAL (A) 4/"
    xlHoja1.Range("R10:R10").ColumnWidth = 20
    
    xlHoja1.Cells(lnFila + 1, 19) = "Tasa de Interes"
    xlHoja1.Cells(lnFila + 2, 19) = "Promedio " & gcPEN_SIMBOLO 'marg ers044-2016
    xlHoja1.Range("S10:S10").ColumnWidth = 15
    
    xlHoja1.Range("I8:S8").MergeCells = True
    xlHoja1.Cells(lnFila, 9) = "Obligaciones Directas (Equivalente en US$)"
       
    xlHoja1.Range("T9:V9").MergeCells = True
    
    xlHoja1.Cells(lnFila + 1, 20) = "Por Destino"
    
    xlHoja1.Cells(lnFila + 1, 20) = "Por a vencer"

    xlHoja1.Cells(lnFila + 2, 20) = "Exportacion"

    xlHoja1.Cells(lnFila + 2, 21) = "Importacion"

    xlHoja1.Cells(lnFila + 2, 22) = "Cap. de Trabajo"

    xlHoja1.Cells(lnFila + 2, 23) = "0 - 30 D"

    xlHoja1.Cells(lnFila + 2, 24) = "31 - 90 D"

    xlHoja1.Cells(lnFila + 2, 25) = "91 - 180 D"

    xlHoja1.Cells(lnFila + 2, 26) = "181 - 270 D"

    xlHoja1.Cells(lnFila + 2, 27) = "271 - 360 D"

    xlHoja1.Cells(lnFila + 2, 28) = "mas de 360 D"

    xlHoja1.Cells(lnFila + 2, 29) = "TOTAL (B)"

    xlHoja1.Cells(lnFila, 30) = "TOTAL OBLIGACIONES"
    xlHoja1.Cells(lnFila + 1, 30) = "CON EL EXTERIOR"
    xlHoja1.Cells(lnFila + 2, 30) = "(A+B)"
    
    CabeceraReporte14 = lnFila
End Function

'***********************

Private Sub cmdImprimir_Click()
    
    Dim i As Integer
    Dim lsCadena As String
    Dim j As Integer

    If Not ValidaDatos Then Exit Sub
        Imprime763201 gbBitCentral, Me.tvOpe.SelectedItem.Text, Me.lstAge, CDate(txtFechaDel.Text), CDate(txtFechaAl.Text)

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
 
Private Sub cmdArchivo_Click()
Dim lsMoneda As String
Dim ldFecha  As Date
Dim sFuente As String
Dim i As Integer

On Error GoTo cmdArchivoErr
    If tvOpe.Nodes.Count = 0 Then
        MsgBox "Lista de operaciones se encuentra vacia", vbInformation, "Aviso"
        Exit Sub
    End If
    If tvOpe.SelectedItem.Tag = "1" Then
        MsgBox "Operación seleccionada no valida...!", vbInformation, "Aviso"
        tvOpe.SetFocus
        Exit Sub
    End If
    If Not ValidaDatos Then Exit Sub
    
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    lsMoneda = IIf(optMoneda(0).value, "1", "2")
    If Me.fraPeriodo.Visible Then
        ldFecha = CDate(DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))) - 1
    End If
    If lsMoneda = "1" Then
        gsSimbolo = gcMN
    Else
        gsSimbolo = gcME
    End If
    If Left(tvOpe.SelectedItem.Key, 1) <> "P" Then
        gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60)) & ": " & Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescHijo = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
        gsOpeDescPadre = Trim(Mid(tvOpe.SelectedItem.Parent.Text, 9, 60))
    Else
      gsOpeDesc = Trim(Mid(tvOpe.SelectedItem.Text, 9, 70))
    End If
   
    Select Case Mid(gsOpeCod, 1, 6)
        Case 770036
            'Modificado PASIERS0142015
            'generaAnexo4A_BienesAdjudRecup_SUCAVE txtFecha
                GeneraAnexoSBS4A_SUCAVE txtFecha
            'end PASI
        Case 770037
            'Modificado PASIERS0142015
            'generaAnexo4B_VentaBsAdjudRecup_SUCAVE txtFecha
            GeneraAnexoSBS4B_SUCAVE txtFecha
            'end PASI
        Case gContAnx15A_Repor      'Anexo 15A
             
            'Modificado PASIERS0282015
            'frmAnx15AReporteNew.Genera15A_Sucave_1y2 lsMoneda, txtFecha
            GeneraAnexoSBS15A_SUCAVE lsMoneda, CDate(txtFecha.Text)
            'End PASI
        Case gContAnx15B
            'ALPA20131030
            'frmAnexo15BPosicionLiquidez.inicia 2, 0, Me
            'MODIFICADO PASIERS0282015
            'GeneratxtAnexo15B lsmoneda, txtFecha
            GeneraAnexoSBS15B_SUCAVE CDate(txtFecha.Text)
            'END PASI
        Case gContAnx15C
            'Modificado PASIERS0282015
            'GeneratxtAnexo15C lsMoneda, txtFecha
            GeneraAnexoSBS15C_SUCAVE CDate(txtFecha.Text)
            'end PASI
        'PASIERS0282015
         Case gContAnx16A
            GeneraAnexoSBS16A_SUCAVE CDate(txtFecha.Text)
        'END PASI********************************************
        Case gContAnx02CredTpoGarantia
            frmAnx02CreDirGarantia.GeneraAnx02CreditosTipoGarantia txtAnio, cboMes.ListIndex + 1, nVal(txtTipCambio), cboMes.Text, 2
        Case gContAnx03FujoCrediticio
            frmAnx02CreDirGarantia.GeneraAnx03FlujoCrediticioPorTipoCred gbBitCentral, txtAnio, cboMes.ListIndex + 1, nVal(txtTipCambio), cboMes.Text, 2
        Case gContRep06Crediticio
            frmRep6Crediticio.GeneraAnexo6CrediticioSUCAVE gsOpeCod, ldFecha
                
        Case RepCGEncBCRObligacion 'Anexo21 Soles
            Call FrmAnexo21.GeneraSUCAVEAnx21Soles(txtFechaAl)
        
        Case RepCGEncBCRObligacionME 'Anexo21 Dolares
            Call FrmAnexo21.GeneraSUCAVEAnx21Dolares(txtFechaAl)
        
        'PASIERS1332014
        Case gRepCreditosIncumplidos 'Reporte 14
             GeneraReporteSBS_CreditosIncumplidos_SUCAVE (ldFecha)
        Case gContAnx10DepColocaPer 'Anexo 10
                GeneraAnx10SBS_DepColocaPer DateAdd("D", -1, CDate(("01/" + Format(DatePart("M", gdFecSis), "00") + "/" + Format(DatePart("YYYY", gdFecSis), "0000"))))
        'end PASI
        'PASI20160127********
        Case gContAnx07N
            GeneraAnexoSBS7_SUCAVE (DateAdd("M", 1, CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio.Text)) - 1)
        'end PASI
            
    End Select
Exit Sub
cmdArchivoErr:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdValorizacion_Click()
frmValorizacionDiariaAnexo15ANew.GeneraValorizacionDiaria gsOpeCod, txtFecha
End Sub 'NAGL ERS079-2016 20170407

Private Sub Form_Activate()
    If tvOpe.Enabled And tvOpe.Visible Then
        tvOpe.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim sCod As String
    On Error GoTo ERROR
    CentraForm Me
    frmMdiMain.Enabled = False
    
    Dim oCambio As nTipoCambio
    Set oCambio = New nTipoCambio
    
    TxtTipCamFijAnt = Format(gnTipCambio, "#,##0.0000")
    txtTipCamFij = Format(oCambio.EmiteTipoCambio(DateAdd("m", 1, gdFecSis), TCFijoMes), "#,##0.0000")
    txtTipCambio = txtTipCamFij
    txtTipCamCompraSBS = "0.0000"
    txtTipCamVentaSBS = "0.0000"
    
    Dim oConst As New NConstSistemas
    If Not lExpandO Then
       sCod = oConst.LeeConstSistema(gConstSistContraerListaOpe)
       If sCod <> "" Then
         lExpand = IIf(UCase(Trim(sCod)) = "FALSE", False, True)
       End If
    Else
       lExpand = lExpandO
    End If
    
    LoadOpeUsu "2"
    LoadAgencia
    txtFecha = gdFecSis
    txtAnio = Year(gdFecSis)
    cboMes.ListIndex = Month(gdFecSis) - 1
    cboSemestre.ListIndex = 0 'FRHU 20131205
    cboTipo.ListIndex = 0 'FRHU 20131205
    txtFechaDel = CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000"))
    txtFechaAl = gdFecSis
    
    If gbBitCentral = True Then
       'txtFecha = oConst.LeeConstSistema(gConstSistCierreMesNegocio)
       txtFecha = gdFecSis
    Else
        Dim oCon As New DConecta
        Dim rsCierre As ADODB.Recordset
        oCon.AbreConexion
        'If oCon.AbreConexionRemota(gsCodAge, False, False) Then
            sCod = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont'))  OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
            Set rsCierre = oCon.CargaRecordSet(sCod)
            If Not rsCierre.EOF Then
                txtFecha = CDate(Trim(rsCierre!cValorVar))
            End If
        'End If
        'oCon.CierraConexion
        oCon.AbreConexion
    End If
    
    cboMes.ListIndex = Month(txtFecha) - 1
    txtAnio = Year(txtFecha)
    RSClose rsCierre
    
    CentraForm Me
    
    Exit Sub
ERROR:
    MsgBox TextErr(Err.Description), vbExclamation, Me.Caption
End Sub


Sub LoadAgencia()
    Dim sqlAge As String
    Dim rsAge As ADODB.Recordset
    Dim oCon As DConecta
    Dim gcCentralCom As String

    Set oCon = New DConecta

    oCon.AbreConexion
    gcCentralCom = "DBCOMUNES.."

    Set rsAge = New ADODB.Recordset
    rsAge.CursorLocation = adUseClient

    If gbBitCentral = True Then
        sqlAge = "Select cAgeDescripcion cNomtab, cAgeCod cValor From Agencias"
    Else
        sqlAge = "Select cAgeDescripcion cNomtab, '112' + cAgeCod cValor From Agencias"
    End If
    Set rsAge = oCon.CargaRecordSet(sqlAge)


    lstAge.Clear

    If Not RSVacio(rsAge) Then
        While Not rsAge.EOF
            lstAge.AddItem Trim(rsAge!cNomtab) & space(500) & Trim(rsAge!cvalor)
            rsAge.MoveNext
        Wend
        lstAge.AddItem "Consolidado" & space(500) & "CONSOL"
    End If

    rsAge.Close
    Set rsAge = Nothing

    'Fuentes de Financiamiento

    If Not gbBitCentral = True Then
        Set rsAge = New ADODB.Recordset
        rsAge.CursorLocation = adUseClient

        sqlAge = "select cvalor, cNomTab from dbcomunes..TablaCod where ccodtab like '22%'  and cvalor<>''"
        Set rsAge = oCon.CargaRecordSet(sqlAge)


        lstFuentes.Clear
        While Not rsAge.EOF
            lstFuentes.AddItem Trim(rsAge!cNomtab) & space(500) & Trim(rsAge!cvalor)
            rsAge.MoveNext
        Wend
        rsAge.Close
        Set rsAge = Nothing
    Else
        Set rsAge = New ADODB.Recordset
        rsAge.CursorLocation = adUseClient

        sqlAge = "select clineacred as cvalor, cdescripcion as cNomTab from coloclineacredito where  len(clineacred)=2  "
        
''''        sqlAge = "SELECT P.cPersCod as cValor, cPersNombre as cNomTab "
''''        sqlAge = sqlAge & " FROM InstitucionFinanc IFc INNER JOIN PERSONA P "
''''        sqlAge = sqlAge & " ON IFc.cPersCod=P.cPersCod "
''''        sqlAge = sqlAge & " where IFc.cIFTpo='05' Order by cPersNombre "
        
        Set rsAge = oCon.CargaRecordSet(sqlAge)
 
        lstFuentes.Clear
             While Not rsAge.EOF
                lstFuentes.AddItem Trim(rsAge!cNomtab) & space(500) & Trim(rsAge!cvalor)
                rsAge.MoveNext
            Wend
        rsAge.Close
        Set rsAge = Nothing
    End If
  
End Sub

Sub LoadOpeUsu(psMoneda As String)
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As MSComctlLib.Node

Set clsGen = New DGeneral
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, sArea, MatOperac, NroRegOpe, psMoneda)
Set clsGen = Nothing
tvOpe.Nodes.Clear
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    nodOpe.Expanded = lExpand
    rsUsu.MoveNext
Loop
RSClose rsUsu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMdiMain.Enabled = True
End Sub

Private Sub OptMoneda_Click(Index As Integer)
    Dim sDig As String
    Dim sCod As String
    Dim oConec As DConecta
    Set oConec = New DConecta
    
    On Error GoTo ERROR
    If optMoneda(0) Then
        sDig = "2"
    Else
        sDig = "1"
    End If
    oConec.AbreConexion
    LoadOpeUsu sDig
    oConec.CierraConexion
    tvOpe.SetFocus
    Exit Sub
ERROR:
    MsgBox TextErr(Err.Description), vbExclamation, Me.Caption
End Sub

Private Sub ActivaControles(Optional plFechaRango As Boolean = True, _
                           Optional plFechaAl As Boolean = False, _
                           Optional plFechaPeriodo As Boolean = False, _
                           Optional plTpoCambio As Boolean = False, _
                           Optional plMoneda As Boolean = True, _
                           Optional plTC As Boolean = False, _
                           Optional plAgencia As Boolean = False, _
                           Optional plFuentes As Boolean = False, _
                           Optional cmdSUCAVE As Boolean = False, _
                           Optional plUsuario As Boolean = False, _
                           Optional plNroUITs As Boolean = False, _
                           Optional plInst As Boolean = False, Optional plTasaMinEncaje As Boolean = False, _
                           Optional plValorEncajeAnterior As Boolean = False)
fraFechaRango.Visible = plFechaRango
fraFecha.Visible = plFechaAl
fraPeriodo.Visible = plFechaPeriodo
fraTCambio.Visible = plTpoCambio
fraTCambio.Height = 690
txtTipCambio2.Visible = False
'frValorEncaje.Visible = plValorEncajeAnterior
'Add Gitu 24-02-2009
txtNroUITs.Visible = plNroUITs
lblNroUIT.Visible = plNroUITs

If plNroUITs Then
    Label5.Caption = "T.C. Fijo"
    cboMes.Visible = False
    Label6.Visible = False
Else
    Label5.Caption = "T.C. Mes"
    cboMes.Visible = True
    Label6.Visible = True
End If
'End Gitu

If plFechaAl Then
    txtFecha.SetFocus
End If
If plFechaRango Then
    txtFechaDel.SetFocus
End If
fraPeriodo.Enabled = True
frmMoneda.Visible = plMoneda
If gsOpeCod = "770030" Then
    cmdArchivo.Visible = cmdSUCAVE
End If
'FRHU 20131205
'FRHU 20140105 Se agrego 760116 y 760118
If gsOpeCod = "763412" Or gsOpeCod = "763413" Or gsOpeCod = "763414" Or gsOpeCod = "760116" Or gsOpeCod = "760118" Then
    cboMes.Visible = False
    Label6.Visible = False
    Me.fraSemestre.Visible = True
Else
   'FRHU 20121216
    If gsOpeCod = "763903" Or gsOpeCod = "763904" Or gsOpeCod = "763905" Then
        cboMes.Visible = False
        Label6.Visible = False
        Me.fraSemestre.Visible = True
        Me.cboTipo.Visible = False
        Me.Label14.Visible = False
    Else
        Me.fraSemestre.Visible = False
        Me.cboTipo.Visible = False
        Me.Label14.Visible = False
    End If
    ''FIN FRHU
End If
''FIN FRHU

'EJVG20120903 ***
'If gsOpeCod = "770036" Or gsOpeCod = "770037" Or gsOpeCod = gContAnx15A_Repor Or gsOpeCod = gContAnx15B Or gsOpeCod = gContAnx15C Then
If gsOpeCod = "770036" Or gsOpeCod = "770037" Then
    cmdArchivo.Visible = cmdSUCAVE
End If

'END EJVG *******
fraTC.Visible = plTC
fraAge.Visible = plAgencia
fraFuentes.Visible = plFuentes
txtInst.Visible = plInst

Me.FraTasaMinEncaje.Visible = plTasaMinEncaje '*** PEAC 20100908

Blanquea
End Sub

Private Sub Blanquea()
Dim i As Integer

For i = 0 To lstAge.ListCount - 1
    lstAge.Selected(i) = False
Next
For i = 0 To lstFuentes.ListCount - 1
    lstFuentes.Selected(i) = False
Next

End Sub

Private Sub tvOpe_Click()
On Error Resume Next
tvOpe.SelectedItem.ForeColor = vbRed
End Sub
Private Sub tvOpe_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 42 Or KeyCode = 38 Or KeyCode = 40 Then
    tvOpe.SelectedItem.ForeColor = "&H80000008"
End If
End Sub

Private Sub tvOpe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    tvOpe.SelectedItem.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_NodeClick(ByVal Node As MSComctlLib.Node)
    tvOpe.SelectedItem.ForeColor = vbRed
    gsOpeCod = Mid(tvOpe.SelectedItem.Text, 1, 6)
    cmdGenerar.Caption = "&Generar"
    cmdImprimir.Visible = False
    fraEuros.Visible = False
    cmdArchivo.Visible = False
    cmdFTP.Visible = False
    
    Select Case Mid(gsOpeCod, 1, 4)
        Case Mid(gContRepBaseFormula, 1, 4)
            If gsOpeCod = gContRepBaseNotasEstadoSitFinan Or gsOpeCod = gContRepBaseNotasEstadoResultado Or gsOpeCod = gContRepBaseFormaB2 Then
            ActivaControles False, False, False
            Else
            ActivaControles False, False, True
            End If
    Case Else
    Select Case Mid(gsOpeCod, 1, 6)
        Case OpeCGRepFlujoDiarioResMN, OpeCGRepFlujoDiarioResME
            ActivaControles False, True
        Case OpeCGRepFlujoDiarioDetMN, OpeCGRepFlujoDiarioDetME
            ActivaControles True
        
        '***************** CARTAS FIANZA ********************************
        Case OpeCGCartaFianzaRepIngreso, OpeCGCartaFianzaRepIngresoME
            ActivaControles True
            txtFechaDel.SetFocus
            
        Case OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME
            ActivaControles True
            txtFechaDel.SetFocus
            
        '********************* REPORTES DE CAJA GENERAL **************************
        Case OpeCGRepRepBancosFlujoMensMN, OpeCGRepRepBancosFlujoMensME
            ActivaControles False
        Case OpeCGRepRepBancosFlujoPFMN, OpeCGRepRepBancosFlujoPFME
            ActivaControles False
        Case OpeCGRepRepBancosResumenPFMN, OpeCGRepRepBancosResumenPFME
            ActivaControles False
        Case OpeCGRepRepBancosConcentFdos
            ActivaControles False, True, , True, False
        Case OpeCGRepRepCMACSFlujoMensMN, OpeCGRepRepCMACSFlujoMensME
            ActivaControles False
        Case OpeCGRepRepCMACSFlujoPFMN, OpeCGRepRepCMACSFlujoPFME
            ActivaControles False
        Case OpeCGRepRepCMACSResumenPFMN, OpeCGRepRepCMACSResumenPFME
            ActivaControles False
        Case OpeCGRepRepOPGirMN, OpeCGRepRepOPGirME
            ActivaControles True
        'A rendir Cuenta
        Case OpeCGRepArendirLibroAuxMN, OpeCGRepArendirLibroAuxME
            ActivaControles True, False
        Case OpeCGRepArendirPendienteMN, OpeCGRepArendirPendienteME
            ActivaControles False, True
            
        Case OpeCGRepArendirViaticoLibroAuxMN, OpeCGRepArendirViaticoLibroAuxME
            ActivaControles True, False

        Case 461045 'Reporten Sustentacion Viaticos GITU
            frmReporteSustViaticos.Show 1
            
        '*** PEAC 20101108
        Case 461046 'Reporten Sustentacion a Rendir Cuenta
            ActivaControles True, False
            
        'Remesa de Cheques
        Case OpeCGRepChequesEnvMN, OpeCGRepChequesEnvME
            ActivaControles True
            
        Case OpeCGRepChequesAnulMN, OpeCGRepChequesAnulME
            ActivaControles True
        
        Case OpeCGRepRepChqRecDetMN, OpeCGRepRepChqRecDetME
            ActivaControles False, True
        Case OpeCGRepRepChqRecResMN, OpeCGRepRepChqRecResME
            ActivaControles False, True
        
        Case OpeCGRepRepChqValDetMN, OpeCGRepRepChqValDetME
            ActivaControles False
        Case OpeCGRepRepChqValResMN, OpeCGRepRepChqValResME
            ActivaControles False, True
        
        Case OpeCGRepRepChqValorizadosDetMN, OpeCGRepRepChqValorizadosDetME
            ActivaControles False
        Case OpeCGRepRepChqValorizadosResMN, OpeCGRepRepChqValorizadosResME
            ActivaControles False
        
        Case OpeCGRepRepChqAnulDetMN, OpeCGRepRepChqAnulDetME
            ActivaControles False
        Case OpeCGRepRepChqAnulResMN, OpeCGRepRepChqAnulResME
            ActivaControles False

        Case OpeCGRepRepChqObsDetMN, OpeCGRepRepChqObsDetME
            ActivaControles False
        Case OpeCGRepRepChqObsResMN, OpeCGRepRepChqObsResME
            ActivaControles False
        Case OpeCGRepChqRecibidoCajaMN, OpeCGRepChqRecibidoCajaME
            ActivaControles False, True
        Case OpeCGRepChqDepositadoMN, OpeCGRepChqDepositadoME
            ActivaControles False, True
        Case OpeCGRepChqPorCobrarConvMN, OpeCGRepChqPorCobrarConvME '*** PEAC 20100510
            ActivaControles False, True
        
        'Adeudados
        Case OpeCGAdeudRepSaldLinFinancDescalceMN
            ActivaControles False, False
         
        Case OpeCGAdeudRepVinculadosME
             ActivaControles False, False
             fraEuros.Visible = True
             txtTCEuros.SetFocus
             txtTCEuros.Text = gnTipoCambioEuro
 
             
        'Presupuesto
        Case OpeCGRepPresuFlujoCaja, OpeCGRepPresuFlujoCajaME
            ActivaControles False, False, False
        Case OpeCGRepPresuServDeuda, OpeCGRepPresuServDeudaME
            ActivaControles False, False, True, True
        Case OpeCGRepPresuFinancia, OpeCGRepPresuFinanciaME
            ActivaControles False, False, True, True
            
         'ENCAJE
        Case OpeCGRepEncajeConsolSdoEnc, OpeCGRepEncajeConsolSdoEncME, OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME, OpeCGRepEncajeConsolPosLiq, OpeCGRepEncajeConsolPosLiqME
            ActivaControles True
            txtFechaDel.SetFocus
        Case OpeCGRepEncajeAgencia, OpeCGRepEncajeAgenciaME
            ActivaControles True
            txtFechaDel.SetFocus
                               
       
        Case OpeEstEncajeSimulacionPlaEncajeMN, OpeEstEncajeSimulacionPlaEncajeME
'              frmCGSimuladorPlanillaEncaje.Show 1
            
            
        'Informe de Encaje al BCR
         Case RepCGEncBCRObligacion, RepCGEncBCRObligacionME, RepCGEncBCRCredDeposi, RepCGEncBCRCredDeposiME, RepCGEncBCRCredRecibi, RepCGEncBCRCredRecibiME, RepCGEncBCRObligaExon, RepCGEncBCRObligaExonME, RepCGEncBCRLinCredExt, RepCGEncBCRLinCredExtME
            If Mid(gsOpeCod, 1, 6) = RepCGEncBCRObligacion Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRCredDeposi Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRCredRecibi Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRObligaExon Or Mid(gsOpeCod, 1, 6) = RepCGEncBCRLinCredExt Then
                
                '*** PEAC 20100908
                If gsOpeCod = RepCGEncBCRObligacion Or gsOpeCod = RepCGEncBCRObligacionME Then
                    ActivaControles True, , , , , , , , , , , , False
                Else
                    ActivaControles True
                End If
                '***FIN PEAC
                
                cmdFTP.Visible = True
                If gsOpeCod = 761201 Then
                   cmdArchivo.Visible = True
                End If
            Else
                ActivaControles True, , , True
                fraTCambio.Height = 1095
                txtTipCambio2.Visible = True
                cmdFTP.Visible = True
                If gsOpeCod = 762201 Then
                   cmdArchivo.Visible = True
                End If
            End If
        '***Agregado por PASI 20140220 TI-ERS102-2013
        Case "761206", "761207", "761208", "761209", "762206", "762207", "762208", "762209"
            If Mid(gsOpeCod, 1, 6) = "761206" Or Mid(gsOpeCod, 1, 6) = "761207" Or Mid(gsOpeCod, 1, 6) = "761208" Or Mid(gsOpeCod, 1, 6) = "761209" Then
                If gsOpeCod = "761206" Then
                    ActivaControles True, , , , , , , , , , , , False
                     'cmdArchivo.Visible = True
                Else
                    ActivaControles True
                    'cmdArchivo.Visible = True
                End If
                cmdFTP.Visible = True
            Else
                ActivaControles True, , , True
                fraTCambio.Height = 1095 'PASIERS1382014
                txtTipCambio2.Visible = True 'PASIERS1382014
                cmdFTP.Visible = True
                'cmdArchivo.Visible = True
            End If
        '***Fin pasi
            txtFechaDel.SetFocus
         ' saldos de caja - bancos y agencias
        Case OpeCGRepSaldoBcos, OpeCGRepSaldoCajAge
               ActivaControles False, False, False, False, False, False, False, False, False
        Case gOpeRptTripleJump, gOpeRptGtoFinanProd 'EJVG20111117
               ActivaControles False, False, False, False, True
        Case gOpeRptConcentracFondos 'EJVG20121002
            ActivaControles False, True, False, False, True
        '************************* CONTABILIDAD **********************************
        Case gContLibroDiario
            ActivaControles False
        Case gContLibroMayor
            ActivaControles False
        Case gContLibroMayCta, gHistoContLibroMayCta
            ActivaControles False
        Case gConvFoncodes
            ActivaControles False, False, False, False, False, False, False, False, False
       
        Case gContRegCompraGastos
            ActivaControles False
        Case gContRegVentas
            ActivaControles False
        
        Case gContRepEstadIngGastos
            ActivaControles False, False, True

        Case gContRepCompraVenta
            ActivaControles True
            
        Case gContRepPlanillaPagoProv
            ActivaControles False
        Case gContRepControlGastoProv
            ActivaControles False
        
        Case gContRepCompraVenta
            ActivaControles False
            
        'FRHU 20140104 Reporte de Estado de Flujo de Efectivo*************************
        Case gContRepEstadoFlujoEfectivo ' RQ13658
            ActivaControles False, False, True, False, , , , , , , False
        'FIN FRHU 20131219************************************************************
        'FRHU 20140104 Reporte de Estado de Flujo de Efectivo*************************
        Case gContRepHojaTrabajoFlujoEfectivo ' RQ13659
            ActivaControles False, False, True, False, , , , , , , False
        'FIN FRHU 20131219************************************************************
        
        'Otros Ajustes
        Case gContAjReclasiCartera
            ActivaControles False
        Case gContAjReclasiGaranti
            ActivaControles False
        Case gContAjInteresDevenga
            ActivaControles False
        Case gContAjInteresSuspens
            ActivaControles False
            
                    
        'CAPTACIONES
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Case gContRepCaptacCVMonExtr
            ActivaControles True, False, False, False, False, True, True
            cmdGenerar.Caption = "&Consolida"
            cmdImprimir.Visible = True
        Case gContRepCaptacSituacCaptac
            ActivaControles False, True, False, False, False, True, False
        Case gContRepCaptacMovCV
            ActivaControles True, False, False, False, False, True, False
        Case gContRepCaptacIngPagos
            ActivaControles False, True, False, False, False, False, True, True
        Case gContRepCaptacCredDesem
            ActivaControles True, False, False, False, False, False, True, True
        'ALPA 20141219**************************
        Case 763525
            ActivaControles False, True, False, False, False, False, False, False
        '***************************************
        'NAGL 202008***************************
        Case 763528
            ActivaControles False, True, False, False, False, False, False, False
        'NAGL Según Acta N°063-2020************
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' OTROS REPORTES
       Case gProvContxCtasCont
            ActivaControles True, False, False, False, , , , , False
             
        Case gProvContxOpe
            ActivaControles True, False, False, False, , , , , False
             
        Case gIntDevPFxFxAG
            ActivaControles True, False, False, False, , , , , False
             
        Case gIntDevCTSxFxAG
            ActivaControles True, False, False, False, , , , , False
             
        Case gPlazoFijoRango
            ActivaControles False, True, False, False, True, , , , False
        Case gRepInstPubliResgos
            ActivaControles False, True, False, False, False, , , , False
        
        Case gPlazoFijoRango
            ActivaControles False, True, False, False, True, , , , False
        
        Case gITFQuincena
            ActivaControles True, False, False, False, False, , , , False
        'ALPA 20081007************************************************************
        Case gITFQuincenaTipoCambioDiario
            ActivaControles True, False, False, False, False, , , , False
        '**************************************************************************
        'GITU 23-02-2008
        Case gGenArchivoDaot
            ActivaControles False, False, True, True, , , , , , , True
        'End Gitu
        'By Capi 04012008
        Case gPlazoFijoIntCash
            ActivaControles False, False, False, False, True, , , , False
        'FRHU 20131205 ANEXO 01 Y 02 ************************************************
        Case gPropuestaDistribuUtilidad 'RQ13649
            ActivaControles False, False, True, False, , , , , , , False
        Case gPropuestaCapitaliUtilidad ' RQ13650
            ActivaControles False, False, True, False, , , , , , , False
        'FIN FRHU ********************************************************************
        'FRHU 20131219 Estado de Cambio Patrimonio************************************
        Case gEstadoDeCambioPatrimonio ' RQ13656
            ActivaControles False, False, True, False, , , , , , , False
        'FIN FRHU 20131219************************************************************
        'By Capi 08022008
        Case gCarteraDetxIntDSP
            ActivaControles False, False, False, True, False, , , , False
        Case gCarteraResxLineas
            ActivaControles False, False, False, True, False, , , , False
        'End by
       'JEOM Reportes Balance contabilidad
        Case gCarteraCreditos
             ActivaControles False, True, False, False, False, , , , False
        
        Case gCarteraInteres
             ActivaControles False, True, False, False, False, , , , False
             
        Case gCreditosCastigados
             ActivaControles False, True, False, False, False, , , , False
             
        Case gInteresesDiferidos
             ActivaControles False, True, False, False, False, , , , False
'
'        Case gPignoraticiosVigentes
'             ActivaControles True, False, False, False, False, , , , False
             
        Case gCreditosCondonados
             ActivaControles False, True, False, False, False, , , , False
        
        Case gIntCreditosRefinanciados
             ActivaControles True, False, False, False, False, , , , False
        
        Case gDetalleGarantias
             ActivaControles False, True, False, True, True, , , , False
        ''''''''''''''''''''''''''''''''''
        'JEOM Reportes Balance Planeamiento
        Case gCredDesembolsosPla
             ActivaControles False, True, False, True, False, , , , False
        
        Case gCarteraVencidaPla
             ActivaControles False, True, False, False, False, , , , False
        
        Case gCarteraRefinanciadaPla
             ActivaControles False, True, False, False, False, , , , False
     
        Case gCarteraJudicialesPla
             ActivaControles False, True, False, False, False, , , , False
             
        'By Capi 18122007 Para Planeamiento
        Case gCarteraRecupCapital
             ActivaControles False, True, False, False, False, , , , False
             
        '*** PEAC 20080915
             
        Case gNumCliCtasEnAhorrosCreditos
             ActivaControles False, False, False, True, False, , , , False
        Case gOpeRptInfoEstadColocBCRP 'EJVG20121113
            ActivaControles False, False, False, False, False, , , , False
        Case gCarteraDeAhorros 'FRHU20140121 RQ13826
            ActivaControles False, False, False, False, False, , , , False
                ''''''''''''''''''''''''''''''''''
        'JEOM Reportes Balance Riesgos
        Case gCredVigentesRiesgos
             ActivaControles False, True, False, False, False, , , , False
        
        Case gCredRefinanciadosRiesgos
             ActivaControles False, True, False, False, False, , , , False
             
        Case gPlazoFijoRiesgos
             ActivaControles False, True, False, False, True, , , , False
        'FRHU REPORTES DE ACTIVOS FIJOS
        Case gMovDepreAcumuladaIntangible
             ActivaControles False, False, True
        Case gMovActivoFijoIntangible
             ActivaControles False, False, True
        Case gCtoIntangibleOtrosAcivosAmo
             ActivaControles False, False, True, False, , , , , , , False
        Case gCtoInmuebleMaquinariaEquipo
             ActivaControles False, False, True, False, , , , , , , False
        Case gCtoDepreciacionDeActivoFijo
             ActivaControles False, False, True, False, , , , , , , False
        'ANEXOS
        Case gContAnx02CredTpoGarantia, gContAnx03FujoCrediticio
            
            If Mid(gsOpeCod, 1, 6) = gContAnx02CredTpoGarantia Then
                ActivaControles False, False, True, True, , , , , True
            Else
                ActivaControles False, False, True, True, , , , , True
            End If
            cmdValorizacion.Visible = False 'NAGL20170407
            
        Case 770036 'BRGO
            'ActivaControles False, True
            ActivaControles False, True, , , , , , , True 'EJVG20120901
        Case 770037 'BRGO
            'ActivaControles False, True
            ActivaControles False, True, , , , , , , True 'EJVG20120901
        Case gContAnx07
            ActivaControles False
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx09
            ActivaControles False, False, False
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx10DepColocaPer
            ActivaControles False
            cmdArchivo.Visible = True 'PASIERS1332014
            cmdValorizacion.Visible = False 'NAGL20170407
        '''''''''''''''''''''''''
        Case gContAnx11MovDepsMonto
            ActivaControles False, True, False, False, False, True, True
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx13DepsSEscMonto
            ActivaControles False, True, False, True, False, False 'JAOR 20202507
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx13DepsSEscMonto_Nuevo
            ActivaControles False, False, True, True, False, False
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx17ListadoFSD
            ActivaControles False, True, False, False, False, True
        Case gContAnx17ListadoGenCtas
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx17CtasFuncionarios
            ActivaControles False, False, True, False, False, False, False
            cmdValorizacion.Visible = False 'NAGL20170407
        '''''''''''''''''''''''''
            
'        Case gContAnx15A_Estad, gContAnx15A_Efect, gContAnx15A_Banco, gContAnx15A_Repor
'            ActivaControles False, True, , True, , , , , True
        Case gContAnx15A_Estad, gContAnx15A_Efect, gContAnx15A_Banco
            ActivaControles False, True, , , , , , , True
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx15A_Repor
            ActivaControles False, True, , True, , , , , True, , , , , True
            'PASIERS0282015***********************************
            cmdArchivo.Visible = True
            'END PASI*****************************************
            txtFecha.SetFocus
        'ALPA20130904************************************
            cmdValorizacion.Visible = True 'NAGL20170407
            
        Case gContAnx15B
            ActivaControles False, True, , True, , , , , True
            cmdArchivo.Visible = True 'PASIERS0282015***********************************
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx15C
            ActivaControles False, True, , , , , , , True
            cmdArchivo.Visible = True 'PASIERS0282015***********************************
        '************************************************
            cmdValorizacion.Visible = False 'NAGL20170407
            
        Case gContAnex15BMens
            ActivaControles False, False, , , , , , , False
            cmdValorizacion.Visible = False 'NAGL ERS079-2017 20180105
          
        Case gContAnx16LiqVenc
            ActivaControles False
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx16A
            ActivaControles False, True
            cmdArchivo.Visible = True 'PASIERS0282015********************
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx16B
            ActivaControles False
            cmdValorizacion.Visible = False 'NAGL20170407
        Case gContAnx17A_FSD
            ActivaControles True, , , False
            txtFechaDel.SetFocus
            cmdValorizacion.Visible = False 'NAGL20170407
            
        Case gContAnx18
            ActivaControles False, False, True, False, False
            cboMes.SetFocus
            cmdValorizacion.Visible = False 'NAGL20170407
        Case 770250
            ActivaControles False, False, True, False, False
            cboMes.SetFocus
            
        'Reportes SBS
        Case gRiesgoSBSA02A, gRiesgoSBSA02B
            ActivaControles False, False, True, False, False, False, False, False
        
        Case gRiesgoSBSA050
            ActivaControles False, False, True, False, False, False, False, False
        
        Case gContRep06Crediticio
            ActivaControles False, False, True, , , , True, , True
            
        Case gContRepSpreadFinanciero 'JUCS 10032017
            ActivaControles False, False, True, , , , True, , True 'JUCS 10032017
        
        Case gRepCreditosIncumplidos 'Reporte 14
            ActivaControles False, False, True, True, False, False, False, False, True
            cmdArchivo.Visible = True 'PASIERS1332014
        
        Case gPatrEfecAjxInfla
            ActivaControles False, False, True, False
            cboMes.SetFocus
        'ALPA 20090727****************************************
        Case gRiesgoSBS2B1
            ActivaControles False, False, True, False
            cboMes.SetFocus
 'ALPA 20090930**************
        Case gRiesgoSBS2D
            ActivaControles False, False, True, False
            cboMes.SetFocus
        '****************************************************
        Case gContAnx24
            ActivaControles False, False, True, True, False
            
        Case 780220
            ActivaControles False, False, False, True, False, False, False, False, False
        
        Case 780350 '*** PEAC 20201125
            ActivaControles False, False, True, False
        
        Case 780211 'Add By Gitu 26-03-2009
            'ActivaControles False, True, False, False, True, , , , False
            ActivaControles False, False, True, False '** PASI 20130913
'        Case gRiesgoSBSA212 'ALPA 20090720*******************************
'            ActivaControles False, True, False, False, True, , , , False
            '************************************************************
        
        'MAVM 20090810****************
        Case gFinancFinanciamientoRecibido
            ActivaControles False, False, True, False, False, False, False, False, , , , True
            'psCtaCont = ""
        Erase Arrays2
        
        Case 780114
            ActivaControles False, False, True, True, False, False, False, False, , , , True
            'psCtaCont = ""
        Erase Arrays2
               
        '*****************************
        'ALPA 20090929**************************************
        Case 763516
            ActivaControles False, True, False, True, False, False, False, False, , , , False
        '***************************************************
        'ALPA 20090929**************************************
        Case 763517
            ActivaControles False, True, False, False, False, False, False, False, , , , False
        '***************************************************
        'ALPA 20101228**************************************
        Case gVerificacionDevenSusp
            ActivaControles False, True, False, False, False, False, False, False, , , , False
        '***************************************************
        Case gVerificacionDiferidos
            ActivaControles False, True, False, True, False, False, False, False, , , , False
        Case OpeCGRepAdeudadoCalendarioMN, OpeCGRepAdeudadoCalendarioME
            ActivaControles False, False, False
        Case OpeCGRepAdeudadoCalendarioVigenteMN, OpeCGRepAdeudadoCalendarioVigenteME
            ActivaControles False, True, False
        Case OpeCGRepAnalisisDeCtaMN, OpeCGRepAnalisisDeCtaME
            ActivaControles False, True, False
        Case gReporteComCartFi
            ActivaControles False, True, False
        'ALPA20140228**************************************************
        Case gContRepGPSpot
            ActivaControles True
        '**************************************************************
        'PASI20160127 **********
        Case gContAnx07N
            fraPeriodo.Visible = True
            cmdArchivo.Visible = True
        'end PASI
        Case Else
            ActivaControles False, False, False
   End Select
   End Select

End Sub

Private Sub tvOpe_Collapse(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H80000008"
End Sub

Private Sub tvOpe_Expand(ByVal Node As MSComctlLib.Node)
    Node.ForeColor = "&H8000000D"
End Sub

Private Sub tvOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdGenerar_Click
    End If
End Sub

Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
Dim oCambio As nTipoCambio
Dim sFecha  As Date
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If fraTCambio.Visible Then
        sFecha = "01/" & IIf(Len(Trim(cboMes.ListIndex + 1)) = 1, "0" & Trim(Str(cboMes.ListIndex + 1)), Trim(cboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text)
        Set oCambio = New nTipoCambio
        If Len(Trim(cboMes.Text)) > 0 And Val(txtAnio.Text) > 1900 Then
            txtTipCambio.Text = Format(oCambio.EmiteTipoCambio(sFecha, TCFijoDia), "#,##0.0000")
        End If
        txtTipCambio.SetFocus
    Else
        cmdGenerar.SetFocus
    End If
End If
End Sub
 

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = True Then
       If fraTCambio.Visible = True Then
        txtTipCambio.SetFocus
       Else
        cmdGenerar.SetFocus
    End If
    End If
End If
End Sub

Private Sub txtFecha_LostFocus()
Dim oCambio As nTipoCambio
Dim sFecha As String
If fraTCambio.Visible = True Then
    'sFecha = DateAdd("m", 1, "01/" & IIf(Len(Trim(cboMes.ListIndex + 1)) = 1, "0" & Trim(Str(cboMes.ListIndex + 1)), Trim(cboMes.ListIndex + 1)) & "/" & Trim(txtAnio.Text))
    'sFecha = DateAdd("d", -1, sFecha)
    If Not IsDate(txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    Set oCambio = New nTipoCambio
    txtTipCambio.Text = Format(oCambio.EmiteTipoCambio(Trim(txtFecha.Text), TCFijoDia), "#,##0.0000")
ElseIf fraTC.Visible = True Then
    'txtTipCamFij = Format(gnTipCambio, "#,##0.0000")
    'TxtTipCamFijAnt = Format(oCambio.EmiteTipoCambio(DateAdd("m", -1, gdFecSis), TCFijoMes), "#,##0.0000")
    If Not IsDate(txtFecha) Then
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    Set oCambio = New nTipoCambio
    txtTipCamFij = Format(oCambio.EmiteTipoCambio(txtFecha, TCFijoMes), "#,##0.0000")
    TxtTipCamFijAnt = Format(oCambio.EmiteTipoCambio(DateAdd("m", -1, txtFecha), TCFijoMes), "#,##0.0000")
End If
End Sub

Private Sub txtFechaAl_Validate(Cancel As Boolean)
Dim oTC As New nTipoCambio
If IsDate(txtFechaAl.Text) Then
Else
    txtFechaAl = Format(gdFecSis, "dd/MM/YYYY")
End If
    txtTipCambio = Format(oTC.EmiteTipoCambio(txtFechaAl, TCFijoMes), "#,##0.00###")
    txtTipCambio2 = Format(oTC.EmiteTipoCambio(CDate(txtFechaAl) + 1, TCFijoMes), "#,##0.00###")
      
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaDel) = True Then
       txtFechaAl.SetFocus
    End If
End If
End Sub

Private Sub txtFechaAl_GotFocus()
    fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFechaAl) = True Then
       cmdGenerar.SetFocus
    End If
End If
End Sub

Private Sub txtInst_Click()
    'If gsOpeCod = gFinancFinanciamientoRecibido Then
        frmListarCuentas.Show
    'Else
        'frmListarIF.Show
    'End If
End Sub

Private Sub txtTCEuros_KeyPress(KeyAscii As Integer)
 KeyAscii = NumerosDecimales(txtTCEuros, KeyAscii, 14, 6)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

Private Sub txtTipCambio_GotFocus()
fEnfoque txtTipCambio
End Sub

Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii, 14, 5)
If KeyAscii = 13 Then
    If Not txtTipCambio2.Visible Then
        cmdGenerar.SetFocus
    Else
        txtTipCambio2.SetFocus
    End If
End If
End Sub

Private Sub txtTipCambio2_GotFocus()
    fEnfoque txtTipCambio2
End Sub

Private Sub txtTipCambio2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTipCambio2, KeyAscii, , 3)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

'ALPA 20090316********************************************
Public Sub ProyeccionesNroCliente()
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
    Dim rs      As ADODB.Recordset
    Dim rsCta   As ADODB.Recordset
    Dim oClie As New DPersonas
    Dim oSdo  As New NCtasaldo
    Set oSdo = Nothing
    'Set oImp = New NContImprimir
    Dim nDebe  As Currency, nHaber  As Currency
    Dim nDebeD As Currency, nHaberH As Currency
    Dim nHaberD As Currency
    Dim nSaldo  As Currency, nSaldoIni As Currency
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ProyeccionesNroCliente"
    lsNomHoja = "Hoja"
    'nLin = 0
    lsArchivo1 = "\spooler\ProyNroCli" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    'nLin = gnLinPage
    nSaltoContador = 12
    'P = 0
    
    Set rsCta = oClie.RecuperaProyeccionesClientes

    Set oClie = Nothing
    
    'sFecha = Mid(rsCta!cMovNro, 1, 8)
   
    Do While Not rsCta.EOF
        DoEvents
        
        If rsCta!cPersSexo = "M" Then
            xlHoja1.Cells(5, 2) = rsCta!nCPersH2006
            xlHoja1.Cells(5, 3) = rsCta!nCPersH2007
            xlHoja1.Cells(5, 4) = rsCta!nCPersH2008
            xlHoja1.Cells(5, 5) = rsCta!nCPersH2009
        End If
        If rsCta!cPersSexo = "F" Then
            xlHoja1.Cells(6, 2) = rsCta!nCPersH2006
            xlHoja1.Cells(6, 3) = rsCta!nCPersH2007
            xlHoja1.Cells(6, 4) = rsCta!nCPersH2008
            xlHoja1.Cells(6, 5) = rsCta!nCPersH2009
        End If
        If rsCta!cPersSexo = "R" Then
            xlHoja1.Cells(8, 2) = rsCta!nCPersH2006
            xlHoja1.Cells(8, 3) = rsCta!nCPersH2007
            xlHoja1.Cells(8, 4) = rsCta!nCPersH2008
            xlHoja1.Cells(8, 5) = rsCta!nCPersH2009
        End If
        If rsCta!cPersSexo = "U" Then
            xlHoja1.Cells(9, 2) = rsCta!nCPersH2006
            xlHoja1.Cells(9, 3) = rsCta!nCPersH2007
            xlHoja1.Cells(9, 4) = rsCta!nCPersH2008
            xlHoja1.Cells(9, 5) = rsCta!nCPersH2009
        End If
        rsCta.MoveNext
        If rsCta.EOF Then
           Exit Do
        End If
    Loop
    
    'rs.Close: Set rs = Nothing
    rsCta.Close: Set rsCta = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    'Set oImp = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub
''ALPA 20090316********************************************
'Public Sub Reporte15Tesoreria()
'    Dim fs As Scripting.FileSystemObject
'    Dim lbExisteHoja As Boolean
'    Dim lsArchivo1 As String
'    Dim lsNomHoja  As String
'    Dim lsNombreAgencia As String
'    Dim lsCodAgencia As String
'    Dim lsMes As String
'    Dim lnContador As Integer
'    Dim lsArchivo As String
'    Dim xlsAplicacion As Excel.Application
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim rs      As ADODB.Recordset
'    Dim rsCta   As ADODB.Recordset
'    Dim oClie As New DPersonas
'    Dim oSdo  As New NCtasaldo
'    Set oSdo = Nothing
'    'Set oImp = New NContImprimir
'    'Dim nSaldo  As Currency, nSaldoIni As Currency
'    Dim nSaldo24S  As Currency, nSaldo24D As Currency
'    Dim nSaldo21S  As Currency, nSaldo21D As Currency
'    Dim sTexto As String
'    Dim sDocFecha As String
'    Dim nSaltoContador As Integer
'    Dim sFecha As String
'    Dim sMov As String
'    Dim sDoc As String
'    Dim N As Integer
'    Dim pnLinPage As Integer
'On Error GoTo GeneraExcelErr
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'    lsArchivo = "Reporte_15_Tesoreria"
'    lsNomHoja = "IndicedeLiquidez"
'    'nLin = 0
'    lsArchivo1 = "\spooler\Repor15Tes" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    'nLin = gnLinPage
'    nSaltoContador = 12
'    'P = 0
'
'    Set rsCta = oClie.Reporte15Tesoreria(txtFecha.Text)
'
'    Set oClie = Nothing
'
'    'sFecha = Mid(rsCta!cMovNro, 1, 8)
'
'    Do While Not rsCta.EOF
'        DoEvents
'
'        If rsCta!nOrden = 4 Then
'            xlHoja1.Cells(19, 3) = rsCta!nSaldos
'            xlHoja1.Cells(19, 4) = rsCta!nSaldoD
'        End If
'        If rsCta!nOrden = 26 Or rsCta!nOrden = 31 Or rsCta!nOrden = 32 Or rsCta!nOrden = 33 Then
'            nSaldo24S = nSaldo24S + rsCta!nSaldos
'            nSaldo24D = nSaldo24D + rsCta!nSaldoD
'        End If
'        If rsCta!nOrden = 24 Then
'            xlHoja1.Cells(22, 3) = rsCta!nSaldos
'            xlHoja1.Cells(22, 4) = rsCta!nSaldoD
'        End If
'        If rsCta!nOrden = 21 Or rsCta!nOrden = 20 Then
'            nSaldo21S = nSaldo21S + rsCta!nSaldos
'            nSaldo21D = nSaldo21D + rsCta!nSaldoD
'        End If
'
'        If rsCta!nOrden = 19 Then
'            nSaldo21S = nSaldo21S - rsCta!nSaldos
'            nSaldo21D = nSaldo21D - rsCta!nSaldoD
'        End If
'
'        rsCta.MoveNext
'        If rsCta.EOF Then
'           Exit Do
'        End If
'    Loop
'     xlHoja1.Cells(21, 3) = nSaldo24S
'     xlHoja1.Cells(21, 4) = nSaldo24D
'     xlHoja1.Cells(20, 3) = nSaldo21S
'     xlHoja1.Cells(20, 4) = nSaldo21D
'    rsCta.Close: Set rsCta = Nothing
'
'    xlHoja1.SaveAs App.path & lsArchivo1
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'    'Set oImp = Nothing
'Exit Sub
'GeneraExcelErr:
'    MsgBox Err.Description, vbInformation, "Aviso"
'    Exit Sub
'End Sub
'ALPA 20090316********************************************
Public Sub Reporte2B1RiesgodeMercado()
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
    Dim rs      As ADODB.Recordset
    Dim rsCta   As ADODB.Recordset
    Dim rsFactor   As ADODB.Recordset
    Dim oFact As New DCtaCont
    Dim oSdo  As New NCtasaldo
    Set oSdo = Nothing
    Dim nSaldo24S  As Currency, nSaldo24D As Currency
    Dim nSaldo21S  As Currency, nSaldo21D As Currency
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteDeMercado2b1"
    lsNomHoja = "Reporte"
    'nLin = 0
    'FactorAjusteRiesgoOperac
    lsArchivo1 = "\spooler\Repor2B1" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    'nLin = gnLinPage
    nSaltoContador = 12
    'P = 0
     
    nMes = cboMes.ListIndex + 1
    Dim sPorceAjust As String
    Set rsFactor = oFact.FactorAjusteRiesgoOperac(txtAnio.Text, IIf(Len(Trim(nMes)) = 1, "0" & CStr(Trim(nMes)), CStr(Trim(nMes))), 2)
    xlHoja1.Range("A7") = "AL " & Format(gdFecSis, "DD") & " DE  " & UCase(Format(gdFecSis, "MMMM")) & " DEL  " & Format(gdFecSis, "YYYY")
    Do While Not rsFactor.EOF
        DoEvents
        xlHoja1.Cells(43, 2) = rsFactor!nFacAjuste
        sPorceAjust = CStr(20 - IIf(IsNull(rsFactor!nFacRequerimiento), 0, rsFactor!nFacRequerimiento)) 'IIf(IIf(IsNull(rsFactor!nFacRequerimiento), 0, rsFactor!nFacRequerimiento) = 0, 0, 1 / rsFactor!nFacRequerimiento)
        xlHoja1.Cells(42, 2) = sPorceAjust
        rsFactor.MoveNext
        
        If rsFactor.EOF Then
           Exit Do
        End If
    Loop
    Set oFact = Nothing
    Set rsCta = CargaDatosPatrimonioEfec(txtAnio.Text, IIf(Len(Trim(nMes)) = 1, "0" & CStr(Trim(nMes)), CStr(Trim(nMes))), 2)
    
    Do While Not rsCta.EOF
        DoEvents
        
        If rsCta!cCtaContCod = "1" Then
            xlHoja1.Cells(17, 2) = rsCta!nSaldoFinImporte
        End If
        If rsCta!cCtaContCod = "2" Then
            xlHoja1.Cells(17, 3) = rsCta!nSaldoFinImporte
        End If
        rsCta.MoveNext
        
        If rsCta.EOF Then
           Exit Do
        End If
    Loop
    rsCta.Close: Set rsCta = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    'Set oImp = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub
Public Sub Reporte2DBasilea()
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
    Dim rs      As ADODB.Recordset
    Dim rsCta   As ADODB.Recordset
    Dim rsFactor   As ADODB.Recordset
    Dim rsRieOpe   As ADODB.Recordset
    Dim oFact As New DCtaCont
    Dim oSdo  As New NCtasaldo
    Set oSdo = Nothing
    Dim nSaldo24S  As Currency, nSaldo24D As Currency
    Dim nSaldo21S  As Currency, nSaldo21D As Currency
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nFactorAjuste As Currency
    Dim nFactorReque As Currency
    Dim nAnio1, nAnio2, nAnio3, nRiesMerc As Currency
    Dim nFactAjusRM As Currency
    Dim nFactRequRM As Currency
    'ALPA 20120627***************************
    Dim lnAnio_ As Integer
    Dim lnMes_ As String
    Dim lnDia_ As String
    Dim ldFecha As Date
    Dim lsFecha As String
    Dim objCta As DCtaCont
    Dim oRsCta As ADODB.Recordset
    Dim lnMontoExposicionAjustadas As Currency
    '****************************************
On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Reporte2D"
    lsNomHoja = "reporte"
    'nLin = 0
    'FactorAjusteRiesgoOperac
    lsArchivo1 = "\spooler\Repor2B1" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    'nLin = gnLinPage
    nSaltoContador = 12
    'P = 0
     
    nMes = cboMes.ListIndex + 1
    Set rsFactor = oFact.FactorAjusteRiesgoOperac(txtAnio.Text, IIf(Len(Trim(nMes)) = 1, "0" & CStr(Trim(nMes)), CStr(Trim(nMes))), 2)
    
    
    
    Do While Not rsFactor.EOF
        DoEvents
        xlHoja1.Cells(53, 3) = rsFactor!nFacAjuste
        xlHoja1.Cells(54, 3) = IIf(IIf(IsNull(rsFactor!nFacRequerimiento), 0, rsFactor!nFacRequerimiento) = 0, 0, rsFactor!nFacRequerimiento)
        nFactorAjuste = rsFactor!nFacAjuste
        nFactorReque = IIf(IIf(IsNull(rsFactor!nFacRequerimiento), 0, rsFactor!nFacRequerimiento) = 0, 0, rsFactor!nFacRequerimiento)
        rsFactor.MoveNext
        
        If rsFactor.EOF Then
           Exit Do
        End If
    Loop
    Set oFact = Nothing
    lnMes_ = IIf(Len(CStr(nMes)) = 1, "0" & CStr(nMes), CStr(nMes))
    lnAnio_ = CInt(txtAnio.Text)
    If lnMes_ = "01" Or lnMes_ = "03" Or lnMes_ = "05" Or lnMes_ = "07" Or lnMes_ = "08" Or lnMes_ = "10" Or lnMes_ = "12" Then
        lnDia_ = "31"
    ElseIf lnMes_ = "02" Then
        If (lnAnio_ Mod 4) = 0 Then
            lnDia_ = "29"
        Else
            lnDia_ = "28"
        End If
    Else
        lnDia_ = "30"
    End If
    ldFecha = CDate(lnAnio_ & "/" & lnMes_ & "/" & lnDia_)
    lsFecha = CStr(lnAnio_ & "" & lnMes_ & "" & lnDia_)
    'xlHoja1.Range("B6") = "AL " & Format(ldFecha, "DD") & " DE  " & UCase(Format(ldFecha, "MMMM")) & " DEL  " & Format(ldFecha, "YYYY")

    Set objCta = New DCtaCont
    Set oRsCta = New ADODB.Recordset
    Set oRsCta = objCta.CargarExposicionesAjustadas2A1(ldFecha)
    If Not RSVacio(oRsCta) Then
        lnMontoExposicionAjustadas = oRsCta!nValor
    End If
    Set oRsCta = Nothing
    Set objCta = Nothing
    
    Set rsRieOpe = oFact.ReporteRiesgoCambiario(txtAnio.Text, IIf(Len(Trim(nMes)) = 1, "0" & CStr(Trim(nMes)), CStr(Trim(nMes))))
    
    Do While Not rsRieOpe.EOF
        DoEvents
        nAnio1 = (rsRieOpe!Saldo_Anualizado1_51 + rsRieOpe!Saldo_Anualizado1_52_57) - (rsRieOpe!Saldo_Anualizado1_41 + rsRieOpe!Saldo_Anualizado1_42_49)
               
        nFactAjusRM = rsRieOpe!nFacAjuste
        nFactRequRM = (20 - rsRieOpe!nFacPonRequer)
        
        nAnio2 = (rsRieOpe!Saldo_Anualizado2_51 + rsRieOpe!Saldo_Anualizado2_52_57) - (rsRieOpe!Saldo_Anualizado2_41 + rsRieOpe!Saldo_Anualizado2_42_49)
        
        nAnio3 = (rsRieOpe!Saldo_Anualizado3_51 + rsRieOpe!Saldo_Anualizado3_52_57) - (rsRieOpe!Saldo_Anualizado3_41 + rsRieOpe!Saldo_Anualizado3_42_49)
                
        rsRieOpe.MoveNext
        
        If rsRieOpe.EOF Then
           Exit Do
        End If
    Loop
    Set oFact = Nothing
    nRiesMerc = ((nAnio1 + nAnio2 + nAnio3) * 15 / 300) * nFactAjusRM * nFactRequRM
    Set rsCta = CargaDatosPatrimonioEfec(txtAnio.Text, IIf(Len(Trim(nMes)) = 1, "0" & CStr(Trim(nMes)), CStr(Trim(nMes))), 2)
    
    Do While Not rsCta.EOF
        DoEvents
        
        If rsCta!cCtaContCod = "1" Then
            nSaldo12 = nSaldo12 + rsCta!nSaldoFinImporte
        End If
        If rsCta!cCtaContCod = "2" Then
            nSaldo12 = nSaldo12 - rsCta!nSaldoFinImporte
        End If
        rsCta.MoveNext
        
        If rsCta.EOF Then
           Exit Do
        End If
    Loop
    
    xlHoja1.Range("B5") = "AL " & Format(gdFecSis, "DD") & " DE  " & UCase(Format(gdFecSis, "MMMM")) & " DEL  " & Format(gdFecSis, "YYYY")
    Dim rsPatrEfec As ADODB.Recordset
    Set rsPatrEfec = New ADODB.Recordset
    Set rsPatrEfec = CargaReporte3Patrimonio2D(lsFecha)
    Dim nPatr1, nPatr2_1, nPatr2_2 As Currency
    Do While Not rsPatrEfec.EOF
        nPatr1 = rsPatrEfec!nValor1
        nPatr2_1 = rsPatrEfec!nValor2
        nPatr2_2 = nPatr1 + nPatr2_1
        
        rsPatrEfec.MoveNext
        If rsPatrEfec.EOF Then
           Exit Do
        End If
    Loop
'    xlHoja1.Cells(85, 5) = nSaldo85
   
   RSClose rsPatrEfec
   Set rsPatrEfec = Nothing
    Dim nSumPatriTem As Currency
    xlHoja1.Cells(14, 3) = lnMontoExposicionAjustadas
    nSaldo12 = ((IIf(nSaldo12 < 0, nSaldo12 * -1, nSaldo12) * 1) / 10) * nFactorAjuste * (20 - nFactorReque)
    xlHoja1.Cells(23, 3) = Format(Round(nSaldo12 / 1, 2), "######.##")
    xlHoja1.Cells(34, 3) = Round(nRiesMerc / 1, 2)
    nPatr1 = nPatr1 - (xlHoja1.Cells(37, 4) + xlHoja1.Cells(29, 4))
    xlHoja1.Cells(44, 4) = Format(Round(nPatr1 / 1, 2), "######.##")
    xlHoja1.Cells(48, 4) = Format(Round(nPatr2_1 / 1, 2), "######.##")
    xlHoja1.Cells(53, 4) = Format(Round(nPatr2_2 / 1, 2), "######.##")
    
    rsCta.Close: Set rsCta = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    'Set oImp = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub
Public Sub ReporteResumenAhorros()
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
    
    Dim rsAhorro   As ADODB.Recordset
    Dim rsAhorroCtaCbles   As ADODB.Recordset
    Dim rsListaRestringidas  As ADODB.Recordset
    
    Dim oAhorro As New DCapMovimientos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Integer
    Dim sCtaAhorro As String
On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ResumenAhorroPlaMoneConsol"
    'Primera Hoja ******************************************************
    lsNomHoja = "Ahorros" 'CPMN
    '*******************************************************************
    lsArchivo1 = "\spooler\RepResAhorro" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    nSaltoContador = 10
     
    nMes = cboMes.ListIndex + 1
    Set rsAhorro = oAhorro.ReporteResumenAhorro(Format(txtFecha.Text, "YYYY/MM/DD"), CDbl(txtTipCambio.Text))
    xlHoja1.Cells(4, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    Do While Not rsAhorro.EOF
        DoEvents
        If rsAhorro!nPlazo = 1 And rsAhorro!cMoneda = "1" Then
            xlHoja1.Cells(nSaltoContador, 1) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador, 2) = rsAhorro!nSaldoAhorro
            xlHoja1.Cells(nSaltoContador, 3) = rsAhorro!nSaldoAhorroIF
            xlHoja1.Cells(nSaltoContador, 4) = rsAhorro!nSaldoAhorroEmSisFi
            xlHoja1.Cells(nSaltoContador, 5) = rsAhorro!nSaldoPF + rsAhorro!nSaldoPFIF
            xlHoja1.Cells(nSaltoContador, 6) = rsAhorro!nSaldoPFRestr
            xlHoja1.Cells(nSaltoContador, 7) = rsAhorro!nSaldoPFIF
            xlHoja1.Cells(nSaltoContador, 8) = rsAhorro!nIntDev
            xlHoja1.Cells(nSaltoContador, 9) = rsAhorro!nSaldoPlazFiEmSisFi
            xlHoja1.Cells(nSaltoContador, 10) = rsAhorro!nSaldoCTS
            nSaltoContador = nSaltoContador + 1
        End If
        rsAhorro.MoveNext
        nContTotal = nContTotal + 1
        If rsAhorro.EOF Then
           Exit Do
        End If
    Loop
'     xlHoja1.Cells(nSaltoContador, 1) = "Total"
'     xlHoja1.Range("B" & nSaltoContador).Formula = "=sum(B10:B" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("C" & nSaltoContador).Formula = "=sum(C10:C" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("D" & nSaltoContador).Formula = "=sum(D10:D" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("E" & nSaltoContador).Formula = "=sum(E10:E" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("F" & nSaltoContador).Formula = "=sum(F10:F" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("G" & nSaltoContador).Formula = "=sum(G10:G" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("H" & nSaltoContador).Formula = "=sum(H10:H" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("I" & nSaltoContador).Formula = "=sum(I10:I" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("J" & nSaltoContador).Formula = "=sum(J10:J" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 10)).Borders.LineStyle = 1
'     'Segunda Hoja ******************************************************
'     lsNomHoja = "LPMN"
'     '*******************************************************************
    If nContTotal > 0 Then
    rsAhorro.MoveFirst
    End If
    
    nSaltoContador = 10
    nMes = cboMes.ListIndex + 1
    xlHoja1.Cells(4, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    Do While Not rsAhorro.EOF
        DoEvents
        If rsAhorro!nPlazo = 2 And rsAhorro!cMoneda = "1" Then
            xlHoja1.Cells(nSaltoContador, 12) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador, 13) = rsAhorro!nSaldoAhorro + rsAhorro!nSaldoAhorroIF
            xlHoja1.Cells(nSaltoContador, 14) = rsAhorro!nSaldoPF
            xlHoja1.Cells(nSaltoContador, 15) = rsAhorro!nSaldoPFRestr
            xlHoja1.Cells(nSaltoContador, 16) = rsAhorro!nSaldoPFIF
            xlHoja1.Cells(nSaltoContador, 17) = rsAhorro!nIntDev
            xlHoja1.Cells(nSaltoContador, 18) = rsAhorro!nSaldoCTS
            nSaltoContador = nSaltoContador + 1
        End If
        rsAhorro.MoveNext
        
        If rsAhorro.EOF Then
           Exit Do
        End If
    Loop
'     xlHoja1.Cells(nSaltoContador, 1) = "Total"
'     xlHoja1.Range("B" & nSaltoContador).Formula = "=sum(B10:B" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("C" & nSaltoContador).Formula = "=sum(C10:C" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("D" & nSaltoContador).Formula = "=sum(D10:D" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("E" & nSaltoContador).Formula = "=sum(E10:E" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("F" & nSaltoContador).Formula = "=sum(F10:F" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("G" & nSaltoContador).Formula = "=sum(G10:G" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Borders.LineStyle = 1
'    'Tercera Hoja ******************************************************
'     lsNomHoja = "CPME"
'    '******************************************************************

    If nContTotal > 0 Then
    rsAhorro.MoveFirst
    End If
    
    nSaltoContador = 32
     
    nMes = cboMes.ListIndex + 1
    'Set rsAhorro = oAhorro.ReporteResumenAhorro(Format(txtFecha.Text, "YYYY/MM/DD"))
    xlHoja1.Cells(4, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    Do While Not rsAhorro.EOF
        DoEvents
        If rsAhorro!nPlazo = 1 And rsAhorro!cMoneda = "2" Then
            xlHoja1.Cells(nSaltoContador, 1) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador, 2) = rsAhorro!nSaldoAhorro
            xlHoja1.Cells(nSaltoContador, 3) = rsAhorro!nSaldoAhorroIF
            xlHoja1.Cells(nSaltoContador, 4) = rsAhorro!nSaldoAhorroEmSisFi
            xlHoja1.Cells(nSaltoContador, 5) = rsAhorro!nSaldoPF + rsAhorro!nSaldoPFIF
            xlHoja1.Cells(nSaltoContador, 6) = rsAhorro!nSaldoPFRestr
            xlHoja1.Cells(nSaltoContador, 7) = rsAhorro!nSaldoPFIF
            xlHoja1.Cells(nSaltoContador, 8) = rsAhorro!nIntDev
            xlHoja1.Cells(nSaltoContador, 9) = rsAhorro!nSaldoPlazFiEmSisFi
            xlHoja1.Cells(nSaltoContador, 10) = rsAhorro!nSaldoCTS
            nSaltoContador = nSaltoContador + 1
        End If
        rsAhorro.MoveNext
        
        If rsAhorro.EOF Then
           Exit Do
        End If
    Loop
'     xlHoja1.Cells(nSaltoContador, 1) = "Total"
'     xlHoja1.Range("B" & nSaltoContador).Formula = "=sum(B10:B" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("C" & nSaltoContador).Formula = "=sum(C10:C" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("D" & nSaltoContador).Formula = "=sum(D10:D" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("E" & nSaltoContador).Formula = "=sum(E10:E" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("F" & nSaltoContador).Formula = "=sum(F10:F" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("G" & nSaltoContador).Formula = "=sum(G10:G" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("H" & nSaltoContador).Formula = "=sum(H10:H" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("I" & nSaltoContador).Formula = "=sum(I10:I" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("J" & nSaltoContador).Formula = "=sum(J10:J" & (nSaltoContador - 1) & ")"
'
'     xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 10)).Borders.LineStyle = 1
'    'Cuarta Hoja ******************************************************
'    lsNomHoja = "LPME"
'    '******************************************************************
'     For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
    If nContTotal > 0 Then
    rsAhorro.MoveFirst
    End If
    
    nSaltoContador = 32
     
    nMes = cboMes.ListIndex + 1
    'Set rsAhorro = oAhorro.ReporteResumenAhorro(Format(txtFecha.Text, "YYYY/MM/DD"))
    xlHoja1.Cells(4, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    Do While Not rsAhorro.EOF
        DoEvents
        If rsAhorro!nPlazo = 2 And rsAhorro!cMoneda = "2" Then
           xlHoja1.Cells(nSaltoContador, 12) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador, 13) = rsAhorro!nSaldoAhorro + rsAhorro!nSaldoAhorroIF
            xlHoja1.Cells(nSaltoContador, 14) = rsAhorro!nSaldoPF
            xlHoja1.Cells(nSaltoContador, 15) = rsAhorro!nSaldoPFRestr
            xlHoja1.Cells(nSaltoContador, 16) = rsAhorro!nSaldoPFIF
            xlHoja1.Cells(nSaltoContador, 17) = rsAhorro!nIntDev
            xlHoja1.Cells(nSaltoContador, 18) = rsAhorro!nSaldoCTS
            nSaltoContador = nSaltoContador + 1
        End If
        rsAhorro.MoveNext
        
        If rsAhorro.EOF Then
           Exit Do
        End If
    Loop
'     xlHoja1.Cells(nSaltoContador, 1) = "Total"
'     xlHoja1.Range("B" & nSaltoContador).Formula = "=sum(B10:B" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("C" & nSaltoContador).Formula = "=sum(C10:C" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("D" & nSaltoContador).Formula = "=sum(D10:D" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("E" & nSaltoContador).Formula = "=sum(E10:E" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("F" & nSaltoContador).Formula = "=sum(F10:F" & (nSaltoContador - 1) & ")"
'     xlHoja1.Range("G" & nSaltoContador).Formula = "=sum(G10:G" & (nSaltoContador - 1) & ")"
     
'     xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Borders.LineStyle = 1
     
    Set oAhorro = Nothing
        
    rsAhorro.Close: Set rsAhorro = Nothing
    
    Set rsAhorroCtaCbles = New ADODB.Recordset
        
    Set rsAhorroCtaCbles = oAhorro.ReporteResumenAhorroCtasCbles(Format(txtFecha.Text, "YYYY/MM/DD"))
    Do While Not rsAhorroCtaCbles.EOF
            If rsAhorroCtaCbles!cCtaContCod = "2101" Then
                'Modificado PASI20140519 TIC1405190012
                'xlHoja1.Cells(59, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                xlHoja1.Cells(60, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                'end Pasi
            End If
            If rsAhorroCtaCbles!cCtaContCod = "2104" Then
                'Modificado PASI20140519 TIC1405190012
                'xlHoja1.Cells(62, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                xlHoja1.Cells(63, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                'end Pasi
            End If
            If rsAhorroCtaCbles!cCtaContCod = "2105" Then
                'Modificado PASI20140519 TIC1405190012
                'xlHoja1.Cells(63, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                xlHoja1.Cells(64, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                'end Pasi
            End If
            If rsAhorroCtaCbles!cCtaContCod = "2106" Then
                'Modificado PASI20140519 TIC1405190012
                'xlHoja1.Cells(64, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                xlHoja1.Cells(65, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                'end Pasi
            End If
            If rsAhorroCtaCbles!cCtaContCod = "2308" Then
                'Modificado PASI20140519 TIC1405190012
                'xlHoja1.Cells(71, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                xlHoja1.Cells(72, 3) = rsAhorroCtaCbles!nSaldoFinImporte
                'end Pasi
            End If
        'ALPA 20110914**************
            'Modificado PASI20140519 TIC1405190012
            'xlHoja1.Cells(47, 1) = txtTipCambio.Text
            xlHoja1.Cells(48, 1) = txtTipCambio.Text
            'end Pasi
        '***************************
        rsAhorroCtaCbles.MoveNext
        nContTotal = nContTotal + 1
        If rsAhorroCtaCbles.EOF Then
           Exit Do
        End If
    Loop
    
    'Set rsAhorroCtaCbles = Nothing
    'Segunda Hoja ******************************************************
    lsNomHoja = "restringidos"
    '******************************************************************
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
    rsAhorroCtaCbles.Close: Set rsAhorroCtaCbles = Nothing
    xlHoja1.Cells(2, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    xlHoja1.Cells(2, 8) = gdFecSis
    Set rsListaRestringidas = New ADODB.Recordset
    Set rsListaRestringidas = oAhorro.ReporteListaCuentasRestringidas(Format(txtFecha.Text, "YYYY/MM/DD"))
    nSaltoContador = 6
    Do While Not rsListaRestringidas.EOF
        xlHoja1.Cells(nSaltoContador, 1) = rsListaRestringidas!cPersNombre
        xlHoja1.Cells(nSaltoContador, 2) = rsListaRestringidas!cAhorro
        If sCtaAhorro = rsListaRestringidas!cAhorro Then
            xlHoja1.Cells(nSaltoContador, 3) = 0
        Else
            xlHoja1.Cells(nSaltoContador, 3) = rsListaRestringidas!nSaldoCapital
        End If
        xlHoja1.Cells(nSaltoContador, 4) = rsListaRestringidas!nCobertura
        xlHoja1.Cells(nSaltoContador, 5) = rsListaRestringidas!cCredito
        xlHoja1.Cells(nSaltoContador, 6) = rsListaRestringidas!nMontoCredito
        xlHoja1.Cells(nSaltoContador, 7) = rsListaRestringidas!cMoneda
        xlHoja1.Cells(nSaltoContador, 8) = rsListaRestringidas!cAgeDescripcion
        xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders.LineStyle = 1
        sCtaAhorro = rsListaRestringidas!cAhorro
        nSaltoContador = nSaltoContador + 1
        rsListaRestringidas.MoveNext
        If rsListaRestringidas.EOF Then
           Exit Do
        End If
    Loop
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Public Sub ReporteResumenAhorrosNew()
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
    
    Dim rsAhorro   As New ADODB.Recordset
    Dim rsAhorroCtaCbles   As New ADODB.Recordset
    Dim rsListaRestringidas  As New ADODB.Recordset
    
    Dim oAhorro As New DCapMovimientos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Integer
    Dim sCtaAhorro As String
    
    '***NAGL Según TIC1806120011
    Dim nSaltoContador2 As Integer
    Dim nCantCL, nCantCLParam As Integer
    Dim X1 As Integer, X2 As Integer
    Dim lnFila As Integer
    Dim TpoCamb As String
    Dim CapCortAh As String, SaldEmpCortAh As String
    Dim CapCortPF As String, CapRestCortPF As String, SaldEmpCortPF As String, CTSCort As String
    Dim CapLargPF As String, CapRestLargPF As String, IntDevLarg As String
    Dim nC2101 As Double, nC2104 As Double, nC2105 As Double, nC2106 As Double, nC2308 As Double 'Sección Cuenta Corto Plazo
    Dim nL2101 As Double, nL2104 As Double, nL2105 As Double, nL2106 As Double, nL2308 As Double 'Sección Cuenta Largo Plazo
    '***END NAGL 20180614***
On Error GoTo GeneraExcelErr

  Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ResumenAhorroPlaMoneConsol"
    'Primera Hoja ******************************************************
    lsNomHoja = "Ahorros" 'CPMN
    '*******************************************************************
    lsArchivo1 = "\spooler\RepResAhorro" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    nC2101 = 0
    nC2104 = 0
    nC2105 = 0
    nC2106 = 0
    nC2308 = 0
    nL2101 = 0
    nL2104 = 0
    nL2105 = 0
    nL2106 = 0
    nL2308 = 0
    
    nSaltoContador = 10
    lnFila = nSaltoContador
    X1 = 0
    X2 = 0
    nMes = cboMes.ListIndex + 1
    Set rsAhorro = oAhorro.ReporteResumenAhorro(Format(txtFecha.Text, "YYYY/MM/DD"), CDbl(txtTipCambio.Text))
    nCantCL = (rsAhorro.RecordCount) / 2
    nCantCLParam = nCantCL
    xlHoja1.Cells(4, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    Do While Not rsAhorro.EOF
        If rsAhorro!cMoneda = "1" Then
            'RESUMEN Corto Plazo MN
            xlHoja1.Cells(nSaltoContador, 1) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador, 2) = rsAhorro!nSaldoAhorroCP
            CapCortAh = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 2), xlHoja1.Cells(nSaltoContador, 2)).Address(False, False)
            
            xlHoja1.Cells(nSaltoContador, 3) = rsAhorro!nSaldoAhorroIFCP
            xlHoja1.Cells(nSaltoContador, 4) = rsAhorro!nSaldoAhorroEmSisFiCP
            SaldEmpCortAh = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 4), xlHoja1.Cells(nSaltoContador, 4)).Address(False, False)
            xlHoja1.Cells(nSaltoContador, 5) = rsAhorro!nSaldoPFCP + rsAhorro!nSaldoPFIFCP
            CapCortPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 5), xlHoja1.Cells(nSaltoContador, 5)).Address(False, False)
            xlHoja1.Cells(nSaltoContador, 6) = rsAhorro!nSaldoPFRestrCP
            CapRestCortPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 6), xlHoja1.Cells(nSaltoContador, 6)).Address(False, False)
             
            xlHoja1.Cells(nSaltoContador, 7) = rsAhorro!nSaldoPFIFCP
            xlHoja1.Cells(nSaltoContador, 8) = rsAhorro!nIntDevCP
            xlHoja1.Cells(nSaltoContador, 9) = rsAhorro!nSaldoPlazFiEmSisFiCP
            SaldEmpCortPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 9), xlHoja1.Cells(nSaltoContador, 9)).Address(False, False)
            xlHoja1.Cells(nSaltoContador, 10) = rsAhorro!nSaldoCTSCP
            CTSCort = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 10), xlHoja1.Cells(nSaltoContador, 10)).Address(False, False)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 11), xlHoja1.Cells(nSaltoContador, 11)).Interior.ColorIndex = 31
            
            'RESUMEN Largo Plazo MN
            xlHoja1.Cells(nSaltoContador, 12) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador, 13) = rsAhorro!nSaldoAhorroLP + rsAhorro!nSaldoAhorroIFLP

            xlHoja1.Cells(nSaltoContador, 14) = rsAhorro!nSaldoPFLP
            CapLargPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 14), xlHoja1.Cells(nSaltoContador, 14)).Address(False, False)
            xlHoja1.Cells(nSaltoContador, 15) = rsAhorro!nSaldoPFRestrLP
            CapRestLargPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 15), xlHoja1.Cells(nSaltoContador, 15)).Address(False, False)

            xlHoja1.Cells(nSaltoContador, 16) = rsAhorro!nSaldoPFIFLP
            xlHoja1.Cells(nSaltoContador, 17) = rsAhorro!nIntDevLP
            IntDevLarg = xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 17), xlHoja1.Cells(nSaltoContador, 17)).Address(False, False)
            xlHoja1.Cells(nSaltoContador, 18) = rsAhorro!nSaldoCTSLP
            xlHoja1.Cells(nSaltoContador, 19).Formula = "=" & "Sum" & "(" & CapCortAh & "," & SaldEmpCortAh & "," & CapCortPF & "," & CapRestCortPF & "," & SaldEmpCortPF & "," & CTSCort & "," & CapLargPF & "," & CapRestLargPF & "," & IntDevLarg & ")"
            xlHoja1.Cells(nSaltoContador, 20).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 19), xlHoja1.Cells(nSaltoContador, 19)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(lnFila + nCantCL + X1 + 8, 20), xlHoja1.Cells(lnFila + nCantCL + X1 + 8, 20)).Address(False, False)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 20), xlHoja1.Cells(nSaltoContador, 20)).Font.Bold = True
            nSaltoContador = nSaltoContador + 1
            X1 = X1 + 1
        Else
            If nCantCL = nCantCLParam Then
                nSaltoContador2 = lnFila + nCantCL + 3
                'RESUMEN CORTO PLAZO ME
                xlHoja1.Cells(nSaltoContador2, 1) = "RESUMEN Corto Plazo ME"
                nSaltoContador2 = nSaltoContador2 + 2
                
                xlHoja1.Cells(nSaltoContador2, 1) = "Agencias"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 1), xlHoja1.Cells(nSaltoContador2 + 2, 1)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 1), xlHoja1.Cells(nSaltoContador2 + 2, 1)).VerticalAlignment = xlCenter
                
                xlHoja1.Cells(nSaltoContador2, 2) = "Ahorros"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 2), xlHoja1.Cells(nSaltoContador2, 4)).Merge True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 2), xlHoja1.Cells(nSaltoContador2, 4)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 2) = "Capital Clientes"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 2), xlHoja1.Cells(nSaltoContador2 + 2, 2)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 2), xlHoja1.Cells(nSaltoContador2 + 2, 2)).VerticalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 3) = "Capital Instit. Financieras"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 3), xlHoja1.Cells(nSaltoContador2 + 2, 3)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 3), xlHoja1.Cells(nSaltoContador2 + 2, 3)).VerticalAlignment = xlJustify
                xlHoja1.Cells(nSaltoContador2 + 1, 4) = "Saldo Con Otras Emp. Sist Finan"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 4), xlHoja1.Cells(nSaltoContador2 + 2, 4)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 4), xlHoja1.Cells(nSaltoContador2 + 2, 4)).VerticalAlignment = xlJustify
                
                xlHoja1.Cells(nSaltoContador2, 5) = "Plazo Fijo"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 5), xlHoja1.Cells(nSaltoContador2, 9)).Merge True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 5), xlHoja1.Cells(nSaltoContador2, 9)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 5) = "Capital Clientes"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 5), xlHoja1.Cells(nSaltoContador2 + 2, 5)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 5), xlHoja1.Cells(nSaltoContador2 + 2, 5)).VerticalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 6) = "Capital Restringido"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 6), xlHoja1.Cells(nSaltoContador2 + 2, 6)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 6), xlHoja1.Cells(nSaltoContador2 + 2, 6)).VerticalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 7) = "Capital Instit.Financieras"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 7), xlHoja1.Cells(nSaltoContador2 + 2, 7)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 7), xlHoja1.Cells(nSaltoContador2 + 2, 7)).VerticalAlignment = xlJustify
                xlHoja1.Cells(nSaltoContador2 + 1, 8) = "Intereses Devengados"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 8), xlHoja1.Cells(nSaltoContador2 + 2, 8)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 8), xlHoja1.Cells(nSaltoContador2 + 2, 8)).VerticalAlignment = xlJustify
                xlHoja1.Cells(nSaltoContador2 + 1, 9) = "Saldo Con Otras Emp. Sist Finan"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 9), xlHoja1.Cells(nSaltoContador2 + 2, 9)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 9), xlHoja1.Cells(nSaltoContador2 + 2, 9)).VerticalAlignment = xlJustify
                
                xlHoja1.Cells(nSaltoContador2 + 1, 10) = "CTS"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 10), xlHoja1.Cells(nSaltoContador2 + 2, 10)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 10), xlHoja1.Cells(nSaltoContador2 + 2, 10)).VerticalAlignment = xlJustify
                
                'RESUMEN LARGO PLAZO ME
                xlHoja1.Cells(nSaltoContador2 - 2, 12) = "RESUMEN Largo Plazo ME"
                
                xlHoja1.Cells(nSaltoContador2, 12) = "Agencias"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 12), xlHoja1.Cells(nSaltoContador2 + 2, 12)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 12), xlHoja1.Cells(nSaltoContador2 + 2, 12)).VerticalAlignment = xlCenter
                
                xlHoja1.Cells(nSaltoContador2, 13) = "Ahorros"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 13), xlHoja1.Cells(nSaltoContador2 + 2, 13)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 13), xlHoja1.Cells(nSaltoContador2 + 2, 13)).VerticalAlignment = xlCenter
                
                xlHoja1.Cells(nSaltoContador2, 14) = "Plazo Fijo"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 14), xlHoja1.Cells(nSaltoContador2, 17)).Merge True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 14), xlHoja1.Cells(nSaltoContador2, 17)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 14) = "Capital Clientes"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 14), xlHoja1.Cells(nSaltoContador2 + 2, 14)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 14), xlHoja1.Cells(nSaltoContador2 + 2, 14)).VerticalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 15) = "Capital Restringido"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 15), xlHoja1.Cells(nSaltoContador2 + 2, 15)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 15), xlHoja1.Cells(nSaltoContador2 + 2, 15)).VerticalAlignment = xlCenter
                xlHoja1.Cells(nSaltoContador2 + 1, 16) = "Capital Instit. Financieras"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 16), xlHoja1.Cells(nSaltoContador2 + 2, 16)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 16), xlHoja1.Cells(nSaltoContador2 + 2, 16)).VerticalAlignment = xlJustify
                xlHoja1.Cells(nSaltoContador2 + 1, 17) = "Intereses Devengados"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 17), xlHoja1.Cells(nSaltoContador2 + 2, 17)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 17), xlHoja1.Cells(nSaltoContador2 + 2, 17)).VerticalAlignment = xlJustify
                
                xlHoja1.Cells(nSaltoContador2 + 1, 18) = "CTS"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 18), xlHoja1.Cells(nSaltoContador2 + 2, 18)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 18), xlHoja1.Cells(nSaltoContador2 + 2, 18)).VerticalAlignment = xlCenter
                
                xlHoja1.Cells(nSaltoContador2 + 1, 19) = "TOTAL ME"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 19), xlHoja1.Cells(nSaltoContador2 + 2, 19)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 19), xlHoja1.Cells(nSaltoContador2 + 2, 19)).VerticalAlignment = xlCenter
                
                xlHoja1.Cells(nSaltoContador2 + 1, 20) = "TOTAL ME AL T/C"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 20), xlHoja1.Cells(nSaltoContador2 + 2, 20)).MergeCells = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 20), xlHoja1.Cells(nSaltoContador2 + 2, 20)).VerticalAlignment = xlCenter
                
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 2, 1), xlHoja1.Cells(nSaltoContador2 + 2, 20)).Font.Name = "Arial"
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 2, 1), xlHoja1.Cells(nSaltoContador2 + 2, 20)).Font.Size = 10
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 2, 1), xlHoja1.Cells(nSaltoContador2 + 2, 20)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 2, 1), xlHoja1.Cells(nSaltoContador2 - 2, 20)).HorizontalAlignment = xlLeft
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 1), xlHoja1.Cells(nSaltoContador2 + 2, 20)).HorizontalAlignment = xlCenter
                
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 4, 11), xlHoja1.Cells(nSaltoContador2 + 2, 11)).Interior.ColorIndex = 31
                ExcelCuadro xlHoja1, 1, nSaltoContador2, 10, CCur(nSaltoContador2 + 2)
                ExcelCuadro xlHoja1, 1, nSaltoContador2 + 1, 10, CCur(nSaltoContador2 + 1)
                ExcelCuadro xlHoja1, 12, nSaltoContador2, 18, CCur(nSaltoContador2 + 2)
                ExcelCuadro xlHoja1, 12, nSaltoContador2 + 1, 20, CCur(nSaltoContador2 + 2)
                
                nCantCLParam = nCantCLParam + 1
                nSaltoContador2 = nSaltoContador - 1 + nCantCL + 8
                
                xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1) = Format(txtTipCambio, "#,##0.000")
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1), xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1)).Interior.ColorIndex = 27
                TpoCamb = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1), xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1)).Address(False, False)
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1), xlHoja1.Cells(nSaltoContador2 + nCantCL + 1, 1)).Font.Bold = True
            End If
            
            'RESUMEN Corto Plazo MN
            xlHoja1.Cells(nSaltoContador2, 1) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador2, 2) = rsAhorro!nSaldoAhorroCP
            CapCortAh = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 2), xlHoja1.Cells(nSaltoContador2, 2)).Address(False, False)
            
            xlHoja1.Cells(nSaltoContador2, 3) = rsAhorro!nSaldoAhorroIFCP
            xlHoja1.Cells(nSaltoContador2, 4) = rsAhorro!nSaldoAhorroEmSisFiCP
            SaldEmpCortAh = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 4), xlHoja1.Cells(nSaltoContador2, 4)).Address(False, False)
            xlHoja1.Cells(nSaltoContador2, 5) = rsAhorro!nSaldoPFCP + rsAhorro!nSaldoPFIFCP
            CapCortPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 5), xlHoja1.Cells(nSaltoContador2, 5)).Address(False, False)
            xlHoja1.Cells(nSaltoContador2, 6) = rsAhorro!nSaldoPFRestrCP
            CapRestCortPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 6), xlHoja1.Cells(nSaltoContador2, 6)).Address(False, False)
             
            xlHoja1.Cells(nSaltoContador2, 7) = rsAhorro!nSaldoPFIFCP
            xlHoja1.Cells(nSaltoContador2, 8) = rsAhorro!nIntDevCP
            xlHoja1.Cells(nSaltoContador2, 9) = rsAhorro!nSaldoPlazFiEmSisFiCP
            SaldEmpCortPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 9), xlHoja1.Cells(nSaltoContador2, 9)).Address(False, False)
            xlHoja1.Cells(nSaltoContador2, 10) = rsAhorro!nSaldoCTSCP
            CTSCort = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 10), xlHoja1.Cells(nSaltoContador2, 10)).Address(False, False)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 11), xlHoja1.Cells(nSaltoContador2, 11)).Interior.ColorIndex = 31
            
            'RESUMEN Largo Plazo MN
            xlHoja1.Cells(nSaltoContador2, 12) = rsAhorro!cAgeDescripcion
            xlHoja1.Cells(nSaltoContador2, 13) = rsAhorro!nSaldoAhorroLP + rsAhorro!nSaldoAhorroIFLP

            xlHoja1.Cells(nSaltoContador2, 14) = rsAhorro!nSaldoPFLP
            CapLargPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 14), xlHoja1.Cells(nSaltoContador2, 14)).Address(False, False)
            xlHoja1.Cells(nSaltoContador2, 15) = rsAhorro!nSaldoPFRestrLP
            CapRestLargPF = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 15), xlHoja1.Cells(nSaltoContador2, 15)).Address(False, False)

            xlHoja1.Cells(nSaltoContador2, 16) = rsAhorro!nSaldoPFIFLP
            xlHoja1.Cells(nSaltoContador2, 17) = rsAhorro!nIntDevLP
            IntDevLarg = xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 17), xlHoja1.Cells(nSaltoContador2, 17)).Address(False, False)
            xlHoja1.Cells(nSaltoContador2, 18) = rsAhorro!nSaldoCTSLP
            xlHoja1.Cells(nSaltoContador2, 19).Formula = "=" & "Sum" & "(" & CapCortAh & "," & SaldEmpCortAh & "," & CapCortPF & "," & CapRestCortPF & "," & SaldEmpCortPF & "," & CTSCort & "," & CapLargPF & "," & CapRestLargPF & "," & IntDevLarg & ")"
            xlHoja1.Cells(nSaltoContador2, 20).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 19), xlHoja1.Cells(nSaltoContador2, 19)).Address(False, False) & "*" & TpoCamb
            nSaltoContador2 = nSaltoContador2 + 1
            X2 = X2 + 1
        End If
        
        If (X1 = nCantCL) Or (X2 = nCantCL) Then
            If (X2 = nCantCL) Then
                nSaltoContador = nSaltoContador2
                lnFila = lnFila + nCantCL + 8
            End If
            
            ExcelCuadro xlHoja1, 1, lnFila, 1, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 2, lnFila, 2, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 3, lnFila, 3, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 4, lnFila, 4, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 5, lnFila, 5, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 6, lnFila, 6, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 7, lnFila, 7, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 8, lnFila, 8, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 9, lnFila, 9, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 10, lnFila, 10, CCur(nSaltoContador - 1)
            
            xlHoja1.Cells(nSaltoContador, 1) = "Total"
            xlHoja1.Cells(nSaltoContador, 2).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(nSaltoContador - 1, 2)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 3).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 3), xlHoja1.Cells(nSaltoContador - 1, 3)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 4).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 4), xlHoja1.Cells(nSaltoContador - 1, 4)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(nSaltoContador - 1, 5)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 6).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(nSaltoContador - 1, 6)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 7).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(nSaltoContador - 1, 7)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 8).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(nSaltoContador - 1, 8)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 9).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 9), xlHoja1.Cells(nSaltoContador - 1, 9)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 10).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 10), xlHoja1.Cells(nSaltoContador - 1, 10)).Address(False, False) & ")"
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 11), xlHoja1.Cells(nSaltoContador, 11)).Interior.ColorIndex = 31
            ExcelCuadro xlHoja1, 1, nSaltoContador, 10, CCur(nSaltoContador)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 10)).Font.Color = vbBlue
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 10)).Font.Bold = True
            
            ExcelCuadro xlHoja1, 12, lnFila, 12, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 13, lnFila, 13, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 14, lnFila, 14, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 15, lnFila, 15, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 16, lnFila, 16, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 17, lnFila, 17, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 18, lnFila, 18, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 19, lnFila, 19, CCur(nSaltoContador - 1)
            ExcelCuadro xlHoja1, 20, lnFila, 20, CCur(nSaltoContador - 1)
            
            xlHoja1.Cells(nSaltoContador, 12) = "Total"
            xlHoja1.Cells(nSaltoContador, 13).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 13), xlHoja1.Cells(nSaltoContador - 1, 13)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 14).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 14), xlHoja1.Cells(nSaltoContador - 1, 14)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 15).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 15), xlHoja1.Cells(nSaltoContador - 1, 15)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 16).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 16), xlHoja1.Cells(nSaltoContador - 1, 16)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 17).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 17), xlHoja1.Cells(nSaltoContador - 1, 17)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 18).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 18), xlHoja1.Cells(nSaltoContador - 1, 18)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 19).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 19), xlHoja1.Cells(nSaltoContador - 1, 19)).Address(False, False) & ")"
            xlHoja1.Cells(nSaltoContador, 20).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(lnFila, 20), xlHoja1.Cells(nSaltoContador - 1, 20)).Address(False, False) & ")"
            ExcelCuadro xlHoja1, 12, nSaltoContador, 20, CCur(nSaltoContador)
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 12), xlHoja1.Cells(nSaltoContador, 20)).Font.Color = vbBlue
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 12), xlHoja1.Cells(nSaltoContador, 20)).Font.Bold = True
            
            If (X1 = nCantCL) Then
                X1 = X1 + 1
            ElseIf (X2 = nCantCL) Then
            
               'Con el Tipo de Cambio
               xlHoja1.Cells(nSaltoContador + 1, 2).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 2), xlHoja1.Cells(nSaltoContador, 2)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 3), xlHoja1.Cells(nSaltoContador, 3)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 4).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 4), xlHoja1.Cells(nSaltoContador, 4)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 5).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 5), xlHoja1.Cells(nSaltoContador, 5)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 6).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 6), xlHoja1.Cells(nSaltoContador, 6)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 7).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 7), xlHoja1.Cells(nSaltoContador, 7)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 8).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 8), xlHoja1.Cells(nSaltoContador, 8)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 9).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 9), xlHoja1.Cells(nSaltoContador, 9)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 10).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 10), xlHoja1.Cells(nSaltoContador, 10)).Address(False, False) & "*" & TpoCamb

               xlHoja1.Cells(nSaltoContador + 1, 13).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 13), xlHoja1.Cells(nSaltoContador, 13)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 14).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 14), xlHoja1.Cells(nSaltoContador, 14)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 15).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 15), xlHoja1.Cells(nSaltoContador, 15)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 16).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 16), xlHoja1.Cells(nSaltoContador, 16)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 17).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 17), xlHoja1.Cells(nSaltoContador, 17)).Address(False, False) & "*" & TpoCamb
               xlHoja1.Cells(nSaltoContador + 1, 18).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 18), xlHoja1.Cells(nSaltoContador, 18)).Address(False, False) & "*" & TpoCamb
               
               'Consolidadas
                xlHoja1.Cells(nSaltoContador + 3, 2).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 2), xlHoja1.Cells(nSaltoContador + 1, 2)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 2), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 2)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 2) = "Cta 2102"
                xlHoja1.Cells(nSaltoContador + 3, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 3), xlHoja1.Cells(nSaltoContador + 1, 3)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 3), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 3)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 3, 4).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 4), xlHoja1.Cells(nSaltoContador + 1, 4)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 4), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 4)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 4) = "Cta 230201"
                xlHoja1.Cells(nSaltoContador + 3, 5).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 5), xlHoja1.Cells(nSaltoContador + 1, 5)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 5), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 5)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 5) = "Cta 2103"
                xlHoja1.Cells(nSaltoContador + 3, 6).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 6), xlHoja1.Cells(nSaltoContador + 1, 6)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 6), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 6)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 6) = "Cta 2107"
                xlHoja1.Cells(nSaltoContador + 3, 7).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 7), xlHoja1.Cells(nSaltoContador + 1, 7)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 7), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 7)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 3, 8).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 8), xlHoja1.Cells(nSaltoContador + 1, 8)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 8), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 8)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 8) = "Cta 2108"
                xlHoja1.Cells(nSaltoContador + 3, 9).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 9), xlHoja1.Cells(nSaltoContador + 1, 9)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 9), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 9)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 9) = "Cta 230301"
                xlHoja1.Cells(nSaltoContador + 3, 10).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 10), xlHoja1.Cells(nSaltoContador + 1, 10)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 10), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 10)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 10) = "Cta 2103"
                xlHoja1.Cells(nSaltoContador + 3, 13).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 13), xlHoja1.Cells(nSaltoContador + 1, 13)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 13), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 13)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 3, 14).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 14), xlHoja1.Cells(nSaltoContador + 1, 14)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 14), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 14)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 14) = "Cta 2103"
                xlHoja1.Cells(nSaltoContador + 3, 15).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 15), xlHoja1.Cells(nSaltoContador + 1, 15)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 15), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 15)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 15) = "Cta 2107"
                xlHoja1.Cells(nSaltoContador + 3, 16).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 16), xlHoja1.Cells(nSaltoContador + 1, 16)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 16), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 16)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 3, 17).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 17), xlHoja1.Cells(nSaltoContador + 1, 17)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 17), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 17)).Address(False, False)
                xlHoja1.Cells(nSaltoContador + 4, 17) = "Cta 2108"
                xlHoja1.Cells(nSaltoContador + 3, 18).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 1, 18), xlHoja1.Cells(nSaltoContador + 1, 18)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador - nCantCL - 8, 18), xlHoja1.Cells(nSaltoContador - nCantCL - 8, 18)).Address(False, False)
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 3, 2), xlHoja1.Cells(nSaltoContador + 3, 10)).Interior.ColorIndex = 31
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 3, 12), xlHoja1.Cells(nSaltoContador + 3, 18)).Interior.ColorIndex = 31
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 3, 2), xlHoja1.Cells(nSaltoContador + 3, 18)).Font.Size = 9
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 3, 2), xlHoja1.Cells(nSaltoContador + 3, 18)).Font.Color = vbWhite
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 3, 2), xlHoja1.Cells(nSaltoContador + 4, 18)).HorizontalAlignment = xlCenter
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador + 3, 2), xlHoja1.Cells(nSaltoContador + 4, 18)).Font.Bold = True
                
            End If
            
        End If
       rsAhorro.MoveNext
    Loop
    
    
    Set rsAhorroCtaCbles = oAhorro.ReporteResumenAhorroCtasCbles(Format(txtFecha.Text, "YYYY/MM/DD"))
    Do While Not rsAhorroCtaCbles.EOF
        If rsAhorroCtaCbles!cCtaContCod = "2101" Then
            'xlHoja1.Cells(60, 3) = rsAhorroCtaCbles!nSaldoFinImporte
            nC2101 = Format(rsAhorroCtaCbles!nSaldoFinImporteCP, "#,##0.00")
            nL2101 = Format(rsAhorroCtaCbles!nSaldoFinImporteLP, "#,##0.00")
        End If
        If rsAhorroCtaCbles!cCtaContCod = "2104" Then
            'xlHoja1.Cells(63, 3) = rsAhorroCtaCbles!nSaldoFinImporte
            nC2104 = Format(rsAhorroCtaCbles!nSaldoFinImporteCP, "#,##0.00")
            nL2104 = Format(rsAhorroCtaCbles!nSaldoFinImporteLP, "#,##0.00")
        End If
        If rsAhorroCtaCbles!cCtaContCod = "2105" Then
            'xlHoja1.Cells(64, 3) = rsAhorroCtaCbles!nSaldoFinImporte
            nC2105 = Format(rsAhorroCtaCbles!nSaldoFinImporteCP, "#,##0.00")
            nL2105 = Format(rsAhorroCtaCbles!nSaldoFinImporteLP, "#,##0.00")
        End If
        If rsAhorroCtaCbles!cCtaContCod = "2106" Then
            'xlHoja1.Cells(65, 3) = rsAhorroCtaCbles!nSaldoFinImporte
            nC2106 = Format(rsAhorroCtaCbles!nSaldoFinImporteCP, "#,##0.00")
            nL2106 = Format(rsAhorroCtaCbles!nSaldoFinImporteLP, "#,##0.00")
        End If
        If rsAhorroCtaCbles!cCtaContCod = "2308" Then
            'xlHoja1.Cells(72, 3) = rsAhorroCtaCbles!nSaldoFinImporte
            nC2308 = Format(rsAhorroCtaCbles!nSaldoFinImporteCP, "#,##0.00")
            nL2308 = Format(rsAhorroCtaCbles!nSaldoFinImporteLP, "#,##0.00")
        End If
        rsAhorroCtaCbles.MoveNext
    Loop
    
    'RESUMEN SEGÚN CUENTAS CONTABLES
    nSaltoContador2 = nSaltoContador2 + 8
    xlHoja1.Cells(nSaltoContador2, 2) = "RESUMEN  según cuentas del Balance de comprobación"
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 2), xlHoja1.Cells(nSaltoContador2, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2, 2), xlHoja1.Cells(nSaltoContador2, 5)).HorizontalAlignment = xlCenter
    nSaltoContador2 = nSaltoContador2 + 2
    
    xlHoja1.Cells(nSaltoContador2, 2) = "CTA CTB."
    xlHoja1.Cells(nSaltoContador2, 3) = "Corto Plazo"
    xlHoja1.Cells(nSaltoContador2, 4) = "Largo Plazo"
    xlHoja1.Cells(nSaltoContador2, 5) = "TOTALES"
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 2, 2), xlHoja1.Cells(nSaltoContador2, 5)).Font.Bold = True
    ExcelCuadro xlHoja1, 2, nSaltoContador2, 5, CCur(nSaltoContador2)
    
    xlHoja1.Cells(nSaltoContador2 + 1, 2) = "Ahorros y obligac. Púb."
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 3), xlHoja1.Cells(nSaltoContador2 + 1, 5)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 3), xlHoja1.Cells(nSaltoContador2 + 1, 5)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 1, 5, CCur(nSaltoContador2 + 1)
    
    nSaltoContador2 = nSaltoContador2 + 1

    xlHoja1.Cells(nSaltoContador2 + 1, 1) = "Datos del Bal. Comprob"
    xlHoja1.Cells(nSaltoContador2 + 1, 2) = "C. 2101"
    xlHoja1.Cells(nSaltoContador2 + 1, 3) = nC2101
    xlHoja1.Cells(nSaltoContador2 + 1, 4) = nL2101
    xlHoja1.Cells(nSaltoContador2 + 1, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 3), xlHoja1.Cells(nSaltoContador2 + 1, 4)).Address(False, False) & ")"
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 1), xlHoja1.Cells(nSaltoContador2 + 1, 5)).Font.Color = vbBlue
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 1, 5, CCur(nSaltoContador2 + 1)
    
    xlHoja1.Cells(nSaltoContador2 + 2, 1) = "Cuadro Corto y largo plazo"
    xlHoja1.Cells(nSaltoContador2 + 2, 2) = "C. 2102"
    xlHoja1.Cells(nSaltoContador2 + 2, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 2), xlHoja1.Cells(nSaltoContador2 - 8, 2)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 2, 4) = 0
    xlHoja1.Cells(nSaltoContador2 + 2, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 2, 3), xlHoja1.Cells(nSaltoContador2 + 2, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 2, 5, CCur(nSaltoContador2 + 2)
    
    xlHoja1.Cells(nSaltoContador2 + 3, 1) = "Cuadro Corto y largo plazo"
    xlHoja1.Cells(nSaltoContador2 + 3, 2) = "C. 2103"
    xlHoja1.Cells(nSaltoContador2 + 3, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 5), xlHoja1.Cells(nSaltoContador2 - 8, 5)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 10), xlHoja1.Cells(nSaltoContador2 - 8, 10)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 3, 4) = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 14), xlHoja1.Cells(nSaltoContador2 - 8, 14)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 3, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 3, 3), xlHoja1.Cells(nSaltoContador2 + 3, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 3, 5, CCur(nSaltoContador2 + 3)
    xlHoja1.Cells(nSaltoContador2 + 4, 1) = "Datos del Bal. Comprob"
    xlHoja1.Cells(nSaltoContador2 + 4, 2) = "C. 2104"
    xlHoja1.Cells(nSaltoContador2 + 4, 3) = nC2104
    xlHoja1.Cells(nSaltoContador2 + 4, 4) = nL2104
    xlHoja1.Cells(nSaltoContador2 + 4, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 4, 3), xlHoja1.Cells(nSaltoContador2 + 4, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 4, 5, CCur(nSaltoContador2 + 4)
    xlHoja1.Cells(nSaltoContador2 + 5, 1) = "Datos del Bal. Comprob"
    xlHoja1.Cells(nSaltoContador2 + 5, 2) = "C. 2105"
    xlHoja1.Cells(nSaltoContador2 + 5, 3) = nC2105
    xlHoja1.Cells(nSaltoContador2 + 5, 4) = nL2105
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 5, 5, CCur(nSaltoContador2 + 5)
    xlHoja1.Cells(nSaltoContador2 + 6, 1) = "Datos del Bal. Comprob"
    xlHoja1.Cells(nSaltoContador2 + 6, 2) = "C. 2106"
    xlHoja1.Cells(nSaltoContador2 + 6, 3) = nC2106
    xlHoja1.Cells(nSaltoContador2 + 6, 4) = nL2106
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 6, 5, CCur(nSaltoContador2 + 6)
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 4, 1), xlHoja1.Cells(nSaltoContador2 + 6, 5)).Font.Color = vbBlue
    
    xlHoja1.Cells(nSaltoContador2 + 7, 1) = "Cuadro Corto y largo plazo"
    xlHoja1.Cells(nSaltoContador2 + 7, 2) = "C. 2107"
    xlHoja1.Cells(nSaltoContador2 + 7, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 6), xlHoja1.Cells(nSaltoContador2 - 8, 6)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 7, 4) = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 15), xlHoja1.Cells(nSaltoContador2 - 8, 15)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 7, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 7, 3), xlHoja1.Cells(nSaltoContador2 + 7, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 7, 5, CCur(nSaltoContador2 + 7)
    xlHoja1.Cells(nSaltoContador2 + 8, 1) = "Cuadro Corto y largo plazo"
    xlHoja1.Cells(nSaltoContador2 + 8, 2) = "C. 2108"
    xlHoja1.Cells(nSaltoContador2 + 8, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 8), xlHoja1.Cells(nSaltoContador2 - 8, 8)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 8, 4) = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 17), xlHoja1.Cells(nSaltoContador2 - 8, 17)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 8, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 8, 3), xlHoja1.Cells(nSaltoContador2 + 8, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 8, 5, CCur(nSaltoContador2 + 8)
    ExcelCuadro xlHoja1, 2, nSaltoContador2 - 1, 2, CCur(nSaltoContador2 + 8) 'Borde para las cuentas
    
    xlHoja1.Cells(nSaltoContador2 + 9, 2) = "TOTAL"
    xlHoja1.Cells(nSaltoContador2 + 9, 3).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 3), xlHoja1.Cells(nSaltoContador2 + 8, 3)).Address(False, False) & ")"
    xlHoja1.Cells(nSaltoContador2 + 9, 4).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 4), xlHoja1.Cells(nSaltoContador2 + 8, 4)).Address(False, False) & ")"
    xlHoja1.Cells(nSaltoContador2 + 9, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 1, 5), xlHoja1.Cells(nSaltoContador2 + 8, 5)).Address(False, False) & ")"
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 9, 2), xlHoja1.Cells(nSaltoContador2 + 9, 5)).Interior.ColorIndex = 27
    ExcelCuadro xlHoja1, 2, nSaltoContador2 + 9, 5, CCur(nSaltoContador2 + 9)

    xlHoja1.Cells(nSaltoContador2 + 10, 2) = "Saldo Con Otras Emp. Sist Finan"
    xlHoja1.Cells(nSaltoContador2 + 10, 3) = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 4), xlHoja1.Cells(nSaltoContador2 - 8, 4)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 10, 4) = 0
    xlHoja1.Cells(nSaltoContador2 + 10, 5) = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 10, 3), xlHoja1.Cells(nSaltoContador2 + 10, 4)).Address(False, False) & ")"
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 10, 2), xlHoja1.Cells(nSaltoContador2 + 11, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 10, 2), xlHoja1.Cells(nSaltoContador2 + 11, 2)).VerticalAlignment = xlJustify
    ExcelCuadro xlHoja1, 2, nSaltoContador2 + 10, 2, CCur(nSaltoContador2 + 11)
    
    'xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 10, 3), xlHoja1.Cells(nSaltoContador2 + 10, 5)).Merge True
    'xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 10, 3), xlHoja1.Cells(nSaltoContador2 + 10, 5)).HorizontalAlignment = xlCenter
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 10, 5, CCur(nSaltoContador2 + 10)
    
    xlHoja1.Cells(nSaltoContador2 + 11, 1) = "Cuadro Corto y largo plazo"
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 11, 5, CCur(nSaltoContador2 + 11)
    
    xlHoja1.Cells(nSaltoContador2 + 12, 1) = "Cuadro Corto y largo plazo"
    xlHoja1.Cells(nSaltoContador2 + 12, 2) = "C. 230301"
    xlHoja1.Cells(nSaltoContador2 + 12, 3) = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 8, 9), xlHoja1.Cells(nSaltoContador2 - 8, 9)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 12, 4) = 0
    xlHoja1.Cells(nSaltoContador2 + 12, 5) = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 12, 3), xlHoja1.Cells(nSaltoContador2 + 12, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 12, 5, CCur(nSaltoContador2 + 12)
    
    xlHoja1.Cells(nSaltoContador2 + 13, 1) = "Datos del Bal. Comprob"
    xlHoja1.Cells(nSaltoContador2 + 13, 2) = "C. 2308"
    xlHoja1.Cells(nSaltoContador2 + 13, 3) = nC2308
    xlHoja1.Cells(nSaltoContador2 + 13, 4) = nL2308
    xlHoja1.Cells(nSaltoContador2 + 13, 5) = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 13, 3), xlHoja1.Cells(nSaltoContador2 + 13, 4)).Address(False, False) & ")"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 13, 5, CCur(nSaltoContador2 + 13)
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 13, 1), xlHoja1.Cells(nSaltoContador2 + 13, 5)).Font.Color = vbBlue
    ExcelCuadro xlHoja1, 2, nSaltoContador2 + 12, 2, CCur(nSaltoContador2 + 13)
    
    xlHoja1.Cells(nSaltoContador2 + 14, 2) = "TOTAL"
    xlHoja1.Cells(nSaltoContador2 + 14, 3).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 12, 3), xlHoja1.Cells(nSaltoContador2 + 13, 3)).Address(False, False) & ")"
    xlHoja1.Cells(nSaltoContador2 + 14, 4).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 12, 4), xlHoja1.Cells(nSaltoContador2 + 13, 4)).Address(False, False) & ")"
    xlHoja1.Cells(nSaltoContador2 + 14, 5).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 12, 5), xlHoja1.Cells(nSaltoContador2 + 13, 5)).Address(False, False) & ")"
    
    xlHoja1.Cells(nSaltoContador2 + 16, 3).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 9, 3), xlHoja1.Cells(nSaltoContador2 + 9, 3)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 14, 3), xlHoja1.Cells(nSaltoContador2 + 14, 3)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 16, 4).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 9, 4), xlHoja1.Cells(nSaltoContador2 + 9, 4)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 14, 4), xlHoja1.Cells(nSaltoContador2 + 14, 4)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 16, 5).Formula = "=" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 9, 5), xlHoja1.Cells(nSaltoContador2 + 9, 5)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 14, 5), xlHoja1.Cells(nSaltoContador2 + 14, 5)).Address(False, False)
    xlHoja1.Cells(nSaltoContador2 + 16, 6) = "Total cuadro Corto y largo plazo"
    
    ExcelCuadro xlHoja1, 3, nSaltoContador2 + 16, 5, CCur(nSaltoContador2 + 16)
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 14, 2), xlHoja1.Cells(nSaltoContador2 + 14, 5)).Interior.ColorIndex = 27
    ExcelCuadro xlHoja1, 2, nSaltoContador2 + 14, 5, CCur(nSaltoContador2 + 14)
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 3, 1), xlHoja1.Cells(nSaltoContador2 + 16, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 3, 1), xlHoja1.Cells(nSaltoContador2 + 16, 2)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 9, 1), xlHoja1.Cells(nSaltoContador2 + 9, 5)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 + 16, 1), xlHoja1.Cells(nSaltoContador2 + 16, 5)).Font.Bold = True
    
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 3, 1), xlHoja1.Cells(nSaltoContador2 + 16, 6)).Font.Name = "Arial"
    xlHoja1.Range(xlHoja1.Cells(nSaltoContador2 - 3, 1), xlHoja1.Cells(nSaltoContador2 + 16, 6)).Font.Size = 11
    
    'Segunda Hoja ******************************************************
    lsNomHoja = "restringidos"
    '******************************************************************
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
    'rsAhorroCtaCbles.Close: Set rsAhorroCtaCbles = Nothing
    xlHoja1.Cells(2, 1) = "AL " & Format(txtFecha.Text, "DD") & " DE  " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    xlHoja1.Cells(2, 8) = gdFecSis
    Set rsListaRestringidas = New ADODB.Recordset
    Set rsListaRestringidas = oAhorro.ReporteListaCuentasRestringidas(Format(txtFecha.Text, "YYYY/MM/DD"))
    nSaltoContador = 6
    Do While Not rsListaRestringidas.EOF
        xlHoja1.Cells(nSaltoContador, 1) = rsListaRestringidas!cPersNombre
        xlHoja1.Cells(nSaltoContador, 2) = rsListaRestringidas!cAhorro
        If sCtaAhorro = rsListaRestringidas!cAhorro Then
            xlHoja1.Cells(nSaltoContador, 3) = 0
        Else
            xlHoja1.Cells(nSaltoContador, 3) = rsListaRestringidas!nSaldoCapital
        End If
        xlHoja1.Cells(nSaltoContador, 4) = rsListaRestringidas!nCobertura
        xlHoja1.Cells(nSaltoContador, 5) = rsListaRestringidas!cCredito
        xlHoja1.Cells(nSaltoContador, 6) = rsListaRestringidas!nMontoCredito
        xlHoja1.Cells(nSaltoContador, 7) = rsListaRestringidas!cMoneda
        xlHoja1.Cells(nSaltoContador, 8) = rsListaRestringidas!cAgeDescripcion
        xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 8)).Borders.LineStyle = 1
        sCtaAhorro = rsListaRestringidas!cAhorro
        nSaltoContador = nSaltoContador + 1
        rsListaRestringidas.MoveNext
        If rsListaRestringidas.EOF Then
           Exit Do
        End If
    Loop
   
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub 'NAGL 20180907 ACTA N°111-2018

'*********************************************************
Public Sub ReporteReparticionGastos()
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
    Dim rs      As ADODB.Recordset
    
    
    Dim oFact As New DCtaCont
    Dim oSdo  As New NCtasaldo
    Set oSdo = Nothing
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteDistribucionGastos"
    lsNomHoja = "DGastos"
    'nLin = 0
    'FactorAjusteRiesgoOperac
    lsArchivo1 = "\spooler\RDGastos" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    
    nSaltoContador = 7
    
     
    xlHoja1.Range("B3") = Format(CDate(txtFecha.Text), "MM")
    Set rs = CargaDatosReparticionGastos(Format(txtFecha.Text, "YYYYMMDD"))
    
    Do While Not rs.EOF
            DoEvents
            xlHoja1.Cells(nSaltoContador, 1) = rs!cUser
            xlHoja1.Cells(nSaltoContador, 2) = rs!dFechaDia
            xlHoja1.Cells(nSaltoContador, 3) = rs!cCtaContCod
            xlHoja1.Cells(nSaltoContador, 4) = rs!cDocNro
            xlHoja1.Cells(nSaltoContador, 5) = rs!cPersNombre
            xlHoja1.Cells(nSaltoContador, 6) = rs!cMovDesc
            xlHoja1.Cells(nSaltoContador, 7) = rs!AG01 + rs!AG02 + rs!AG03 + rs!AG04 + rs!AG06 + rs!AG07 + rs!AG09 + rs!AG10 + rs!AG12 + rs!AG13 + rs!AG24 + rs!AG25 + rs!AG31
            xlHoja1.Cells(nSaltoContador, 9) = rs!AG02 + rs!AG03 + rs!AG04 + rs!AG06 + rs!AG07 + rs!AG09 + rs!AG10 + rs!AG12 + rs!AG13 + rs!AG24 + rs!AG25 + rs!AG31
            xlHoja1.Cells(nSaltoContador, 10) = rs!AG01
            xlHoja1.Cells(nSaltoContador, 11) = rs!AG02
            xlHoja1.Cells(nSaltoContador, 12) = rs!AG03
            xlHoja1.Cells(nSaltoContador, 13) = rs!AG04
            xlHoja1.Cells(nSaltoContador, 14) = rs!AG06
            xlHoja1.Cells(nSaltoContador, 15) = rs!AG07
            xlHoja1.Cells(nSaltoContador, 16) = rs!AG09
            xlHoja1.Cells(nSaltoContador, 17) = rs!AG10
            xlHoja1.Cells(nSaltoContador, 18) = rs!AG12
            xlHoja1.Cells(nSaltoContador, 19) = rs!AG13
            xlHoja1.Cells(nSaltoContador, 20) = rs!AG24
            xlHoja1.Cells(nSaltoContador, 21) = rs!AG25
            xlHoja1.Cells(nSaltoContador, 22) = rs!AG31
            xlHoja1.Cells(nSaltoContador, 23) = rs!AG01 + rs!AG02 + rs!AG03 + rs!AG04 + rs!AG06 + rs!AG07 + rs!AG09 + rs!AG10 + rs!AG12 + rs!AG13 + rs!AG24 + rs!AG25 + rs!AG31
        nSaltoContador = nSaltoContador + 1
        rs.MoveNext
        If rs.EOF Then
           Exit Do
        End If
    Loop
    rs.Close: Set rs = Nothing
    xlHoja1.Range("G" & nSaltoContador).Formula = "=sum(G7:G" & nSaltoContador - 1 & ")"
    xlHoja1.Range("H" & nSaltoContador).Formula = "=sum(H7:H" & nSaltoContador - 1 & ")"
    xlHoja1.Range("I" & nSaltoContador).Formula = "=sum(I7:I" & nSaltoContador - 1 & ")"
    xlHoja1.Range("J" & nSaltoContador).Formula = "=sum(J7:J" & nSaltoContador - 1 & ")"
    xlHoja1.Range("K" & nSaltoContador).Formula = "=sum(K7:K" & nSaltoContador - 1 & ")"
    xlHoja1.Range("L" & nSaltoContador).Formula = "=sum(L7:L" & nSaltoContador - 1 & ")"
    xlHoja1.Range("M" & nSaltoContador).Formula = "=sum(M7:M" & nSaltoContador - 1 & ")"
    xlHoja1.Range("N" & nSaltoContador).Formula = "=sum(N7:N" & nSaltoContador - 1 & ")"
    xlHoja1.Range("O" & nSaltoContador).Formula = "=sum(O7:O" & nSaltoContador - 1 & ")"
    xlHoja1.Range("P" & nSaltoContador).Formula = "=sum(P7:P" & nSaltoContador - 1 & ")"
    xlHoja1.Range("Q" & nSaltoContador).Formula = "=sum(Q7:Q" & nSaltoContador - 1 & ")"
    xlHoja1.Range("R" & nSaltoContador).Formula = "=sum(R7:R" & nSaltoContador - 1 & ")"
    xlHoja1.Range("S" & nSaltoContador).Formula = "=sum(S7:S" & nSaltoContador - 1 & ")"
    xlHoja1.Range("T" & nSaltoContador).Formula = "=sum(T7:T" & nSaltoContador - 1 & ")"
    xlHoja1.Range("U" & nSaltoContador).Formula = "=sum(U7:U" & nSaltoContador - 1 & ")"
    xlHoja1.Range("V" & nSaltoContador).Formula = "=sum(V7:V" & nSaltoContador - 1 & ")"
    xlHoja1.Range("W" & nSaltoContador).Formula = "=sum(W7:W" & nSaltoContador - 1 & ")"
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    'Set oImp = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub
'EJVG20130124 ***
Public Sub ReporteReparticionGastosNew()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lsArchivo As String
    Dim rs As ADODB.Recordset
    Dim i As Long, j As Long
    Dim lnNroColumnas As Long, lnNroFilas As Long
    Dim lsLetraColumnaExcel As String

    On Error GoTo ErrReporteReparticionGastosNew

    lsArchivo = "\spooler\RDGastos" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    Set rs = CargaDatosReparticionGastosNew(Format(txtFecha.Text, "YYYYMMDD"))
    
    lnNroColumnas = rs.Fields.Count
    lnNroFilas = rs.RecordCount

    If lnNroFilas = 0 Then
        MsgBox "No existe información para generar a esta fecha", vbInformation, "Aviso"
        Exit Sub
    End If

    Set xlsLibro = xlsAplicacion.Workbooks.Add
    Set xlsHoja = xlsLibro.Worksheets.Add

    For i = 0 To lnNroColumnas - 1
        xlsHoja.Cells(3, i + 1) = "'" & rs.Fields(i).Name
    Next i
    
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Range("A4").CopyFromRecordset rs
    
    xlsHoja.Cells(1, CInt(lnNroColumnas / 2)) = "DISTRIBUCION DE GASTOS POR AGENCIAS"
    xlsHoja.Cells(1, CInt(lnNroColumnas / 2)).Font.Bold = True
    xlsHoja.Cells(1, CInt(lnNroColumnas / 2)).Font.Size = 14
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, lnNroColumnas)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, lnNroColumnas)).Interior.Color = RGB(217, 217, 217)
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3, lnNroColumnas)).Font.Bold = True
    
    xlsHoja.Range(xlsHoja.Cells(3, 1), xlsHoja.Cells(3 + lnNroFilas, lnNroColumnas)).Borders.Weight = xlThin
    xlsHoja.Range(xlsHoja.Cells(3, 7), xlsHoja.Cells(4 + lnNroFilas, lnNroColumnas)).NumberFormat = "#,##0.00"

    For i = 7 To lnNroColumnas
        lsLetraColumnaExcel = xlsHoja.Cells(1, i).Address
        lsLetraColumnaExcel = Mid(lsLetraColumnaExcel, 2, InStr(2, lsLetraColumnaExcel, "$") - 2)
        xlsHoja.Range(lsLetraColumnaExcel & (4 + lnNroFilas)) = "=SUM(" & lsLetraColumnaExcel & "4:" & lsLetraColumnaExcel & (3 + lnNroFilas) & ")"
    Next

    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
Exit Sub
ErrReporteReparticionGastosNew:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub
'END EJVG *******
'ALPA 20100425***************************************************
Private Sub ImprimirListadoDeCalificacionDeCartera()
Dim oImpresion As NContImprimir
Dim sCad As String
Set oImpresion = New NContImprimir

sCad = oImpresion.ImprimeListadodeCalificacion(gdFecSis, 76)
EnviaPrevio sCad, "RESUMEN DE CARTERA POR LINEAS (CALIFICACION)", gnLinPage, False
End Sub

Private Sub ImprimirDetalleDeCalificacionDeCartera()
Dim oImpresion As NContImprimir
Dim sCad As String
Set oImpresion = New NContImprimir

sCad = oImpresion.ImprimeDetalledeCalificacion(gdFecSis, 60)
EnviaPrevio sCad, "RESUMEN DE CARTERA POR LINEAS (CALIFICACION)", gnLinPage, False
End Sub
'******************************************************************

'*** PEAC 20101108
Public Sub ReporteSustentacionARendirCuenta(ByVal psOpeCod As String, ByVal pnMoneda As Moneda, ByVal pdFechaDel As Date, ByVal pdFechaAl As Date)

Dim lsUsuario As String
Dim lsAgencia As String
Dim lsArea As String
Dim ldFechaIni As String
Dim ldFechaFin As String
Dim nU As Integer

ldFechaIni = pdFechaDel
ldFechaFin = pdFechaAl

If ldFechaIni > ldFechaFin Then
    MsgBox "Fecha final debe ser mayor", vbOKOnly, "Error"
    Exit Sub
End If

lsArchivo = App.path & "\SPOOLER\SustAREndirCta_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    Exit Sub
End If

GeneraReporteSustentacionARendirCuenta ldFechaIni, ldFechaFin

ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
CargaArchivo lsArchivo, App.path & "\SPOOLER\"
End Sub

'*** PEAC 20101108
Function GeneraReporteSustentacionARendirCuenta(pdFechaIni As String, pdFechaFin As String) As Integer
    Dim rsMay As New ADODB.Recordset
    Dim oOperacion As DOperacion
    Set oOperacion = New DOperacion
    Dim N As Integer
    Dim lsCtaArendir As String
    Dim lsCtaPendiente As String
    Dim lsUsuario As String
    Dim lsFechaER As String
    Dim lsFechaRen As String
    Dim lnImporte As Double
    Dim lsArea    As String
    Dim lsLugar   As String
    Dim lsPeriodo As String
    Dim lsMotivo  As String
    Dim lnTotDev As Double
    Dim lnTotRend As Double
    Dim lnTotal As Double
   
    Dim sImpre      As String
    Dim sImpreAge   As String
    Dim lsRepTitulo As String
    Dim lsHoja      As String
    
    Dim nPosIni As Long
    Dim nPosFin As Long
    Dim lsNomMesP As String
    Dim lsNomMesLL As String
    Dim w As Long
    Dim lsAgencia As String
    Dim oARend As New NARendir
    
    Set rsMay = oARend.GetSustentacionArendirCuenta(pdFechaIni, pdFechaFin)

    nLin = 1
    lsHoja = "SustCuentasARendir"
    ExcelAddHoja lsHoja, xlLibro, xlHoja1
        
    Do While Not rsMay.EOF
          Dim rsDocRend As ADODB.Recordset
          Dim lnMovNro As Long
          
          lsUsuario = rsMay!Usuario
          lsFechaER = rsMay!Fecha
          lnImporte = CStr(rsMay!mto_otorg)
          lsArea = rsMay!AREA
          lsAgencia = rsMay!Agencia
          lsFechaRen = rsMay!FechaRend
          lsPeriodo = "Del " + CStr(pdFechaIni) + " Al " + CStr(pdFechaFin)
          
          lsMotivo = rsMay!Descripcion
          lnMovNro = rsMay!nMovNroRef
          
          Set rsDocRend = oARend.GetDocSustentariosArendirCuenta(lnMovNro)
          
          If rsDocRend.RecordCount <> 0 Then
                                                      
          ImprimeSustARendirCuentaExcel lsUsuario, lsFechaER, lnImporte, lsArea, _
                                             lsAgencia, lsPeriodo, lsMotivo, pdFechaIni, _
                                             pdFechaFin, nLin, lsFechaRen
                                             
          xlHoja1.Range("A" & 0 + nLin & ":H" & 0 + nLin).Font.Bold = True
          nLin = nLin + 1
                    
          Dim nVal As Double
          Dim nItem As Integer
          Dim lsFechaSust As String
          Dim lsDocAbr As String
          Dim lsDocNro As String
          Dim lsProvSust As String
          Dim lsDetalleSust As String
          Dim lnImporteSust As Double

          nItem = 1
          nVal = 0
          Do While Not rsDocRend.EOF
             lsFechaSust = CDate(rsDocRend!dDocFecha)
             lsDocAbr = rsDocRend!cDocAbrev
             lsDocNro = rsDocRend!cDocNro
             lsProvSust = rsDocRend!cPersNombre
             lsDetalleSust = rsDocRend!Detalle
             lnImporteSust = rsDocRend!nDocImporte

             ImprimeDetalleArendirCuenta nLin, nItem, lsFechaSust, lsDocAbr, lsDocNro, lsProvSust, lsDetalleSust, lnImporteSust
             nVal = nVal + lnImporteSust
             nItem = nItem + 1
             rsDocRend.MoveNext
             If rsDocRend.EOF Then
                Exit Do
             End If
             
          Loop
          
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).LineStyle = xlContinuous
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).Weight = xlMedium
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).NumberFormat = "#,##0.00"
          xlHoja1.Cells(nLin, 8) = nVal
          lnTotRend = lnTotRend + nVal
          nLin = nLin + 1
          xlHoja1.Cells(nLin, 6) = "Devolucion a Caja"
          xlHoja1.Cells(nLin, 8) = rsMay!mto_otorg - nVal
          lnTotDev = lnTotDev + (rsMay!mto_otorg - nVal)
          nLin = nLin + 1
          xlHoja1.Cells(nLin, 8) = rsMay!mto_otorg
          lnTotal = lnTotal + rsMay!mto_otorg
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).LineStyle = xlContinuous
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).Borders(xlEdgeTop).Weight = xlMedium
          xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).NumberFormat = "#,##0.00"
          nLin = nLin + 2
          End If
          
          rsMay.MoveNext
          If rsMay.EOF Then
             Exit Do
          End If
       Loop
       nLin = nLin + 1
       xlHoja1.Cells(nLin, 6) = "Total de Rendiciones"
       xlHoja1.Cells(nLin, 8) = lnTotRend
       nLin = nLin + 1
       xlHoja1.Cells(nLin, 6) = "Total de Devoluciones a Caja"
       xlHoja1.Cells(nLin, 8) = lnTotDev
       nLin = nLin + 1
       xlHoja1.Cells(nLin, 6) = "TOTAL"
       xlHoja1.Cells(nLin, 8) = lnTotal
    
    RSClose rsMay
    
    Exit Function
ErrImprime:
     MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
       If lbLibroOpen Then
          xlLibro.Close
          xlAplicacion.Quit
       End If
       Set xlAplicacion = Nothing
       Set xlLibro = Nothing
       Set xlHoja1 = Nothing

End Function


'*** PEAC 20101108
Private Sub ImprimeSustARendirCuentaExcel(psUsuario As String, psFechaER As String, pnImporte As Double, psArea As String, _
                                               psLugarViaje As String, psPeriodo As String, psMotivo As String, _
                                               pdFecha As String, pdFecha2 As String, lnLin As Long, psFechaRen As String)
    
    xlHoja1.Range("A1:S1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 8  'Item
    xlHoja1.Range("B1:B1").ColumnWidth = 12 'Fecha
    xlHoja1.Range("C1:C1").ColumnWidth = 15 'Documento
    xlHoja1.Range("D1:D1").ColumnWidth = 8  'Serie
    xlHoja1.Range("E1:E1").ColumnWidth = 12 'Numero
    xlHoja1.Range("F1:F1").ColumnWidth = 60 'Proveedor
    xlHoja1.Range("G1:G1").ColumnWidth = 60 'Detalle
    xlHoja1.Range("H1:H1").ColumnWidth = 12 'Importe
        
    xlHoja1.Cells(nLin, 2) = "Sustentación de a Rendir Cuenta"
    xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    xlHoja1.Range("A" & nLin & ":I" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 2) = "Del " & pdFecha & " Al " & pdFecha2
    xlHoja1.Range("A" & nLin & ":I" & nLin).Font.Bold = True

    xlHoja1.Range("A" & nLin & ":I" & nLin).Merge True
    xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    
    nLin = nLin + 2
    
    xlHoja1.Cells(nLin, 1) = "Usuario"
    xlHoja1.Cells(nLin, 3) = psUsuario
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).Font.Bold = True
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Fecha"
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).NumberFormat = "dd/mm/yyyy;@"
    xlHoja1.Cells(nLin, 3) = psFechaER
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Importe Otorgado"
    xlHoja1.Cells(nLin, 3) = pnImporte
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).NumberFormat = "#,###0.00"
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Fecha de Rendición"
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).NumberFormat = "dd/mm/yyyy;@"
    xlHoja1.Cells(nLin, 3) = Trim(psFechaRen)
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 1) = "Area Funcional"
    xlHoja1.Cells(nLin, 3) = psArea
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "Agencia"
    xlHoja1.Cells(nLin, 3) = psLugarViaje
    xlHoja1.Range("A" & 0 + nLin & ":B" & 0 + nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "Motivo"
    xlHoja1.Cells(nLin, 3) = psMotivo
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft
    
    nLin = nLin + 2
    xlHoja1.Cells(nLin, 1) = "Item"
    xlHoja1.Cells(nLin, 2) = "Fecha"
    xlHoja1.Cells(nLin, 3) = "Documento"
    xlHoja1.Cells(nLin, 4) = "Serie"
    xlHoja1.Cells(nLin, 5) = "Número"
    xlHoja1.Cells(nLin, 6) = "Proveedor"
    xlHoja1.Cells(nLin, 7) = "Detalle"
    xlHoja1.Cells(nLin, 8) = "Importe"
       
    xlHoja1.Range("A" & nLin & ":I" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":I" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":H" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":H" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":H" & nLin).Borders(xlInsideVertical).Color = vbBlack

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
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub

'*** PEAC 20101108
Private Sub ImprimeDetalleArendirCuenta(pnlin As Long, pnItem As Integer, psFecha As String, psDocumento As String, psDocNro As String, _
                                          psProveedor As String, psDetalle As String, pnImporte As Double)
    Dim Item As Integer
    nLin = pnlin
    Item = pnItem
    xlHoja1.Cells(nLin, 1) = Item
    
    xlHoja1.Cells(nLin, 2) = psFecha
    
    xlHoja1.Cells(nLin, 3) = psDocumento
    
    If InStr(1, psDocNro, "-") <> 0 Then
        xlHoja1.Cells(nLin, 4) = "'" & Format(Left(psDocNro, InStr(1, psDocNro, "-") - 1), "000")
        xlHoja1.Cells(nLin, 5) = "'" & Format(Mid(psDocNro, InStr(1, psDocNro, "-") + 1), "000000000")
    Else
        xlHoja1.Cells(nLin, 5) = "'" & psDocNro
    End If
    
    xlHoja1.Cells(nLin, 6) = psProveedor
    xlHoja1.Cells(nLin, 7) = psDetalle
    xlHoja1.Cells(nLin, 8) = pnImporte
    xlHoja1.Range("H" & nLin + 0 & ":H" & nLin + 0).NumberFormat = "#,##0.00"
    nLin = nLin + 1
End Sub
'ALPA 20101228**********************************************************************
Public Sub ReporteVerificacionInteresDevenSusp()
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
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "VerificacionIntDevSuspenso"
    'Primera Hoja ******************************************************
    lsNomHoja = "Intereses"
    '*******************************************************************
    lsArchivo1 = "\spooler\RepIntDevSus" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    nSaltoContador = 8
     
    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.RecuperaDatosReporteVerificacionIntDevSuspen(txtFecha.Text)
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    xlHoja1.Cells(5, 2) = Format(txtFecha.Text, "DD") & " DE " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
    '        DoEvents
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = rsCreditos!cAgeCod
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cAgeDescripcion
                xlHoja1.Cells(nSaltoContador, 3) = rsCreditos!cMoneda
                xlHoja1.Cells(nSaltoContador, 4) = rsCreditos!cCalGen
                xlHoja1.Cells(nSaltoContador, 5) = rsCreditos!cCalGenDescr
                xlHoja1.Cells(nSaltoContador, 6) = rsCreditos!cPersCod
                xlHoja1.Cells(nSaltoContador, 7) = rsCreditos!cPersNombre
                xlHoja1.Cells(nSaltoContador, 8) = rsCreditos!cRefinan
                xlHoja1.Cells(nSaltoContador, 9) = rsCreditos!cTpoCredCod
                xlHoja1.Cells(nSaltoContador, 10) = rsCreditos!cConsDescripcion
                xlHoja1.Cells(nSaltoContador, 11) = rsCreditos!cCtaCod
                xlHoja1.Cells(nSaltoContador, 12) = rsCreditos!nIntSusp
                xlHoja1.Cells(nSaltoContador, 13) = rsCreditos!nIntDev
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
'GeneraExcelErr:
'    MsgBox Err.Description, vbInformation, "Aviso"
'    Exit Sub
End Sub

'***********************************************************************************
'*********************************************
'ALPA20130905*********************************
'***********************************************************************************
Public Sub ReporteAnexo15B(ByVal pnTipoCambio As Currency)
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    'Dim lsNombreAgencia As String
    'Dim lsCodAgencia As String
    'Dim lsMes As String
    'Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim TituloProgress As String 'NAGL 20170407
    Dim MensajeProgress As String 'NAGL 20170407
    Dim oBarra As clsProgressBar 'NAGL 20170407
    Dim nprogress As Integer 'NAGL 20170407
    
    'Dim rsRep15B As ADODB.Recordset
    'Dim oRep15B As New DbalanceCont
    'Dim nTotalAcredores20 As Currency
    'Dim nTotalAcredores10 As Currency
    'Dim nTotalAcredoresTo As Currency
    Dim oDbalanceCont As DbalanceCont
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim pdFechaFinDeMes As Date
    Dim pdFechaFinDeMesMA As Date
    Dim dFechaAnte As Date
    Dim ldFechaPro As Date
    Dim pdFecha As Date
    Dim nDia As Integer
    Dim oCambio As nTipoCambio
    Dim lnTipoCambioFC As Currency
    Dim lnTipoCambioProceso As Currency
    'Dim lnTipoCambioBalance As Currency
    Dim lnTipoCambioBalanceAnterior As Currency
    Dim nTipoCambioAn As Currency
    Dim loRs As ADODB.Recordset
    Dim lnSubastasMN As Currency
    Dim lnSubastasME As Currency
    Dim lnOtrosDepMND1y2_C0 As Currency
    Dim lnOtrosDepMND4aM_C0 As Currency
    Dim lnOtrosDepMND1y2_C1 As Currency
    Dim lnOtrosDepMED1y2_C0 As Currency
    Dim lnOtrosDepMED4aM_C0 As Currency
    Dim lnOtrosDepMED1y2_C1 As Currency
    Dim lnSubastasMND1y2_C1 As Currency
    Dim lnSubastasMND1y2_C0 As Currency
    Dim lnSubastasMND4aM_C0 As Currency
    Dim lnSubastasMED1y2_C1 As Currency
    Dim lnSubastasMED1y2_C0 As Currency
    Dim lnSubastasMED4aM_C0 As Currency
    
    Dim lnSubastasMED3o3_C0 As Currency
    Dim lnSubastasMND3o3_C0 As Currency
    Dim lnSubastasMED3o3_C1 As Currency
    Dim lnSubastasMND3o3_C1 As Currency
    
    Dim lnOtrosDepMND3o3_C0 As Currency
    Dim lnOtrosDepMED3o3_C0 As Currency
    Dim lnOtrosDepMND3o3_C1 As Currency
    Dim lnOtrosDepMED3o3_C1 As Currency
    
    Dim lnSubastasMND4aM_C1 As Currency
    Dim lnSubastasMED4aM_C1 As Currency
    
    Dim nTotalObligSugEncajMN As Currency
    Dim nTotalTasaBaseEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency
    Dim nTotalTasaBaseEncajME  As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME  As Currency
    Dim nTotalTasaBaseEncajMNDiario As Currency
    Dim nTotalTasaBaseEncajMEDiario  As Currency
    Dim nTotalTasaBaseEncajMN_DADiario As Currency
    Dim nTotalTasaBaseEncajME_DADiario As Currency
    Dim nTotalTasaBaseEncajMN_DA  As Currency
    Dim nTotalObligSugEncajMN_DA As Currency
    Dim nTotalTasaBaseEncajME_DA As Currency
    Dim nTotalObligSugEncajME_DA As Currency
    Dim ix As Integer, rx As Integer
    Dim nSubValor1 As Currency
    Dim nSubValor2 As Currency
    
On Error GoTo GeneraExcelErr

    Set oBarra = New clsProgressBar
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    nprogress = 0
    oBarra.Progress nprogress, "ANEXO 15B: Ratio de Cobertura de Liquidez", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "ANEXO 15B: Ratio de Cobertura de Liquidez"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    'NAGL20170407
    
    pdFecha = Format(txtFecha.Text, "YYYY/MM/DD")
    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New DbalanceCont
    'If nDia >= 15 Then
        'dFechaAnte = DateAdd("d", -(nDia - 1), pdFecha)
    'Else
        'dFechaAnte = DateAdd("d", -(nDia - 1), DateAdd("m", -1, pdFecha))
    'End If
    Set oCambio = New nTipoCambio
    
    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, pdFecha), TCFijoDia), "#,##0.0000")
    End If
    nTipoCambioAn = pnTipoCambio
    
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ANEXO_15B"
    'Primera Hoja ******************************************************
    lsNomHoja = "Anx15B"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_15B_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    'lsArchivo1 = "\spooler\Anx15B_New_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & ".xls" NAGL 20170407
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    xlHoja1.Cells(4, 2) = "AL " & Format(txtFecha.Text, "YYYY/MM/DD")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "1", 0, 1, "100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "2", 0, 1, "100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "3", 0, 1, "100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(1, pdFecha, "4", 0, 1, "100", "B1")
    
    'NAGL ERS079-2016 20170407 (Saldo en Caja)
    nSaldoDiario1 = ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, 1) 'oDbalanceCont.ObtenerCtaContSaldoDiario("1111", pdFecha)
    nSaldoDiario2 = ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, 2) 'oDbalanceCont.ObtenerCtaContSaldoDiario("1121", pdFecha)
    xlHoja1.Cells(10, 3) = nSaldoDiario1
    xlHoja1.Cells(10, 4) = nSaldoDiario2
    
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "1", nSaldoDiario1, 1, "200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "2", nSaldoDiario2, 1, "200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "3", xlHoja1.Cells(10, 6), 1, "200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(2, pdFecha, "4", xlHoja1.Cells(10, 7), 1, "200", "B1")
    
    xlHoja1.Range(xlHoja1.Cells(10, 3), xlHoja1.Cells(10, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
    
    'Inicio TOSE
        nTotalObligSugEncajMN = 0
        nTotalTasaBaseEncajMN = 0
        nTotalObligSugEncajME = 0
        nTotalTasaBaseEncajME = 0
        lnTotalObligacionesAlDiaMN = 0
        lnTotalObligacionesAlDiaME = 0
        nTotalTasaBaseEncajMNDiario = 0
        nTotalTasaBaseEncajMEDiario = 0
        nTotalTasaBaseEncajMN_DADiario = 0
        nTotalTasaBaseEncajME_DADiario = 0
        nTotalTasaBaseEncajMN_DA = 0
        nTotalTasaBaseEncajMN = 0
        nTotalTasaBaseEncajME_DA = 0
        nTotalTasaBaseEncajME = 0
        ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
        ldFechaPro = DateAdd("d", -Day(ldFechaPro), ldFechaPro)
        
        For ix = 1 To Day(DateAdd("d", -Day(pdFecha), pdFecha))
            ldFechaPro = DateAdd("d", 1, ldFechaPro)
                If Month(ldFechaPro) = Month(DateAdd("d", 1, ldFechaPro)) Then
                    lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(ldFechaPro, TCFijoDia), "#,##0.0000")
                Else
                    lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(DateAdd("d", -1, ldFechaPro), TCFijoDia), "#,##0.0000")
                End If
                
                'SOLES
                nTotalObligSugEncajMN_DA = oDbalanceCont.SaldoCtas(1, "761201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
                nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "232") 'Ahorros
                nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234")
                nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") 'Depositos a plazo fijo
                'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234"))
                lnTotalObligacionesAlDiaMN = lnTotalObligacionesAlDiaMN + nTotalObligSugEncajMN_DA '*************NAGL ERS079-2016 20170407
                
                'DOLARES
                nTotalObligSugEncajME_DA = oDbalanceCont.SaldoCtas(1, "762201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
                nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "232") 'Ahorros
                nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234")
                nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") 'Depositos a plazo fijo
                'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234"))
                lnTotalObligacionesAlDiaME = lnTotalObligacionesAlDiaME + nTotalObligSugEncajME_DA '*************NAGL ERS079-2016 20170407
    Next ix
    
    If nDia >= 15 Then
         lnTipoCambioBalanceAnterior = Format(oCambio.EmiteTipoCambio(DateAdd("d", -1, pdFecha), TCFijoDia), "#,##0.0000")
    Else
         lnTipoCambioBalanceAnterior = Format(oCambio.EmiteTipoCambio(pdFechaFinDeMesMA, TCFijoDia), "#,##0.0000")
    End If 'NAGL ERS079-2016 20170407
    
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue
    
    nSaldoDiario1 = oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1112", pdFecha, 1) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("111802", pdFecha, "1", 0)
    nSaldoDiario2 = oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1122", pdFecha, 1) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("112802", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior, 2)
    nSaldoDiario1 = IIf(nSaldoDiario1 > 0, nSaldoDiario1, 0)
    nSaldoDiario2 = IIf(nSaldoDiario2 > 0, nSaldoDiario2, 0)
    xlHoja1.Cells(11, 3) = nSaldoDiario1
    xlHoja1.Cells(11, 4) = nSaldoDiario2
    
    '*********NAGL ERS079-2016 20170407 Ajuste por Encaje Exigible
    nSaldoDiario1 = (lnTotalObligacionesAlDiaMN / Day(DateAdd("d", -Day(pdFecha), pdFecha))) * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100) * 0.25 * -1
    nSaldoDiario2 = (lnTotalObligacionesAlDiaME / Day(DateAdd("d", -Day(pdFecha), pdFecha))) * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("33") / 100) * -1
    
    xlHoja1.Cells(12, 3) = nSaldoDiario1
    xlHoja1.Cells(12, 4) = nSaldoDiario2
    
    '***************NAGL ERS079-2016 20170407
    
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "1", nSaldoDiario1, 1, "300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "2", nSaldoDiario2, 1, "300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "3", xlHoja1.Cells(11, 6), 1, "300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(3, pdFecha, "4", xlHoja1.Cells(11, 7), 1, "300", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "1", nSaldoDiario1, 1, "301", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "2", nSaldoDiario2, 1, "301", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "3", xlHoja1.Cells(11, 6), 1, "301", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(300, pdFecha, "4", xlHoja1.Cells(11, 7), 1, "301", "B1")
    
    '***************NAGL ERS079-2016 20170407 VALORES REPRESENTATIVOS
    nSaldoDiario1 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "C[BD]")
    nSaldoDiario2 = oDbalanceCont.ObtenerSumaValorRepresentativos(pdFecha, "LT")
    
    xlHoja1.Cells(14, 3) = nSaldoDiario1
    xlHoja1.Cells(15, 3) = nSaldoDiario2
   
    
    '***************NAGL ERS079-2016 20170407
    xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(11, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(12, 3), xlHoja1.Cells(12, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(14, 3), xlHoja1.Cells(14, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(15, 3), xlHoja1.Cells(15, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue
    
    lnOtrosDepMND1y2_C0 = 0
    lnOtrosDepMND1y2_C1 = 0
    lnOtrosDepMED1y2_C0 = 0
    lnOtrosDepMED4aM_C0 = 0
    lnOtrosDepMND4aM_C0 = 0
    lnOtrosDepMED1y2_C1 = 0
    lnSubastasMND1y2_C1 = 0
    lnSubastasMND1y2_C0 = 0
    lnSubastasMND4aM_C0 = 0

    lnSubastasMED3o3_C0 = 0
    lnSubastasMND3o3_C0 = 0
    lnSubastasMED3o3_C1 = 0
    lnSubastasMND3o3_C1 = 0
    
    lnOtrosDepMND3o3_C0 = 0
    lnOtrosDepMED3o3_C0 = 0
    lnOtrosDepMND3o3_C1 = 0
    lnOtrosDepMED3o3_C1 = 0
    
    lnSubastasMED4aM_C1 = 0
    lnSubastasMND4aM_C1 = 0

    
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado15B("1", pdFecha, pnTipoCambio, 30)
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnSubastasMND1y2_C1 = lnSubastasMND1y2_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnSubastasMND3o3_C1 = lnSubastasMND3o3_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnSubastasMND1y2_C0 = lnSubastasMND1y2_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnSubastasMND3o3_C0 = lnSubastasMND3o3_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                lnSubastasMND4aM_C0 = lnSubastasMND4aM_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
      loRs.MoveNext
    Loop
    End If
    Set loRs = Nothing
    
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado15B("2", pdFecha, pnTipoCambio, 30)
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnSubastasMED1y2_C1 = lnSubastasMED1y2_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnSubastasMED3o3_C1 = lnSubastasMED3o3_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnSubastasMED1y2_C0 = lnSubastasMED1y2_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnSubastasMED3o3_C0 = lnSubastasMED3o3_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                lnSubastasMED4aM_C0 = lnSubastasMED4aM_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If

      loRs.MoveNext
    Loop
    End If
    Set loRs = Nothing
  
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15B(20, pdFecha, "1", pnTipoCambio, 1, 30)
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnOtrosDepMND1y2_C1 = lnOtrosDepMND1y2_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnOtrosDepMND3o3_C1 = lnOtrosDepMND3o3_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnOtrosDepMND1y2_C0 = lnOtrosDepMND1y2_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnOtrosDepMND3o3_C0 = lnOtrosDepMND3o3_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                lnOtrosDepMND4aM_C0 = lnOtrosDepMND4aM_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        loRs.MoveNext
    Loop
    End If
    Set loRs = Nothing
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15B(20, pdFecha, "2", pnTipoCambio, 1, 30)
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnOtrosDepMED1y2_C1 = lnOtrosDepMED1y2_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnOtrosDepMED3o3_C1 = lnOtrosDepMED3o3_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                lnOtrosDepMED1y2_C0 = lnOtrosDepMED1y2_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                lnOtrosDepMED3o3_C0 = lnOtrosDepMED3o3_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                lnOtrosDepMED4aM_C0 = lnOtrosDepMED4aM_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        loRs.MoveNext
    Loop
    End If
    
    Set loRs = Nothing
    Set loRs = oDbalanceCont.ObtenerFondeoEncaje(pdFecha, pnTipoCambio, 30)
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(36, 3) = CCur(xlHoja1.Cells(36, 3)) + loRs!nSaldCntMN '
                    xlHoja1.Cells(36, 4) = CCur(xlHoja1.Cells(36, 4)) + loRs!nSaldCntME '
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(38, 3) = CCur(xlHoja1.Cells(38, 3)) + loRs!nSaldCntMN '
                    xlHoja1.Cells(38, 4) = CCur(xlHoja1.Cells(38, 4)) + loRs!nSaldCntME '
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(40, 3) = CCur(xlHoja1.Cells(40, 3)) + loRs!nSaldCntMN '
                    xlHoja1.Cells(40, 4) = CCur(xlHoja1.Cells(40, 4)) + loRs!nSaldCntME '
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(37, 3) = CCur(xlHoja1.Cells(37, 3)) + loRs!nSaldCntMN '
                    xlHoja1.Cells(37, 4) = CCur(xlHoja1.Cells(37, 4)) + loRs!nSaldCntME '
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(38, 3) = CCur(xlHoja1.Cells(38, 3)) + loRs!nSaldCntMN '
                    xlHoja1.Cells(38, 4) = CCur(xlHoja1.Cells(38, 4)) + loRs!nSaldCntME '
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(40, 3) = CCur(xlHoja1.Cells(40, 3)) + loRs!nSaldCntMN '
                    xlHoja1.Cells(40, 4) = CCur(xlHoja1.Cells(40, 4)) + loRs!nSaldCntME '
            End If
        End If
        loRs.MoveNext
    Loop
    End If
    xlHoja1.Cells(36, 3) = CCur(xlHoja1.Cells(36, 3)) - lnOtrosDepMND1y2_C1 - lnSubastasMND1y2_C1 '
    xlHoja1.Cells(36, 4) = CCur(xlHoja1.Cells(36, 4)) - lnOtrosDepMED1y2_C1 - lnSubastasMED1y2_C1 '
    xlHoja1.Cells(37, 3) = CCur(xlHoja1.Cells(37, 3)) - lnOtrosDepMND1y2_C0 - lnSubastasMND1y2_C0 '
    xlHoja1.Cells(37, 4) = CCur(xlHoja1.Cells(37, 4)) - lnOtrosDepMED1y2_C0 - lnSubastasMED1y2_C0
    xlHoja1.Cells(38, 3) = CCur(xlHoja1.Cells(38, 3)) - lnOtrosDepMND3o3_C0 - lnSubastasMND3o3_C0 - lnSubastasMND3o3_C1 - lnOtrosDepMND3o3_C1
    xlHoja1.Cells(38, 4) = CCur(xlHoja1.Cells(38, 4)) - lnOtrosDepMED3o3_C0 - lnSubastasMED3o3_C0 - lnSubastasMED3o3_C1 - lnOtrosDepMED3o3_C1
    
    xlHoja1.Cells(36, 3) = IIf(CCur(xlHoja1.Cells(36, 3)) < 0, 0, CCur(xlHoja1.Cells(36, 3)))
    xlHoja1.Cells(37, 3) = IIf(CCur(xlHoja1.Cells(37, 3)) < 0, 0, CCur(xlHoja1.Cells(37, 3)))
    xlHoja1.Cells(38, 3) = IIf(CCur(xlHoja1.Cells(38, 3)) < 0, 0, CCur(xlHoja1.Cells(38, 3)))
    
    xlHoja1.Cells(36, 4) = IIf(CCur(xlHoja1.Cells(36, 4)) < 0, 0, CCur(xlHoja1.Cells(36, 4)))
    xlHoja1.Cells(37, 4) = IIf(CCur(xlHoja1.Cells(37, 4)) < 0, 0, CCur(xlHoja1.Cells(37, 4)))
    xlHoja1.Cells(38, 4) = IIf(CCur(xlHoja1.Cells(38, 4)) < 0, 0, CCur(xlHoja1.Cells(38, 4)))
    
    xlHoja1.Range(xlHoja1.Cells(36, 3), xlHoja1.Cells(38, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    xlHoja1.Cells(40, 3) = CCur(xlHoja1.Cells(40, 3)) + (lnSubastasMND1y2_C1 + lnSubastasMND1y2_C0 + lnSubastasMND3o3_C0) + lnOtrosDepMND1y2_C1 + lnOtrosDepMND1y2_C0 + lnOtrosDepMND3o3_C0 + (lnSubastasMND3o3_C1 + lnOtrosDepMND3o3_C1)
    xlHoja1.Cells(40, 4) = CCur(xlHoja1.Cells(40, 4)) + (lnSubastasMED1y2_C1 + lnSubastasMED1y2_C0 + lnSubastasMED3o3_C0) + lnOtrosDepMED1y2_C1 + lnOtrosDepMED1y2_C0 + lnOtrosDepMED3o3_C0 + (lnSubastasMED3o3_C1 + lnOtrosDepMED3o3_C1)
    
    xlHoja1.Cells(40, 3) = IIf(CCur(xlHoja1.Cells(40, 3)) < 0, 0, CCur(xlHoja1.Cells(40, 3)))
    xlHoja1.Cells(40, 4) = IIf(CCur(xlHoja1.Cells(40, 4)) < 0, 0, CCur(xlHoja1.Cells(40, 4)))
    xlHoja1.Range(xlHoja1.Cells(40, 3), xlHoja1.Cells(40, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue
    
    '**********NAGL ERS079-2016 20170407
    nSubValor1 = oDbalanceCont.ObtenerSaldoDiarioRestringido(pdFecha, "1", 30)
    'nSaldoDiario1 = IIf(nSubValor1 < 0, 0, nSubValor1)
    nSaldoDiario1 = nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25170301", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + IIf(nSubValor1 < 0, 0, nSubValor1)
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25170302", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + IIf(nSubValor1 < 0, 0, nSubValor1)
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25170303", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + IIf(nSubValor1 < 0, 0, nSubValor1)
    
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251704", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251705", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251706", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2116", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("211701", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2118", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    nSubValor1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2518", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + nSubValor1
    
    '************NAGL ERS079-2016 20170407
    nSubValor2 = 0
    nSaldoDiario2 = 0
    nSubValor2 = oDbalanceCont.ObtenerSaldoDiarioRestringido(pdFecha, "2", 30)
    'nSaldoDiario2 = IIf(nSubValor2 < 0, 0, nSubValor2)
    nSaldoDiario2 = nSubValor2
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25270301", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + IIf(nSubValor2 < 0, 0, nSubValor2), 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25270302", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + IIf(nSubValor2 < 0, 0, nSubValor2), 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25270303", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + IIf(nSubValor2 < 0, 0, nSubValor2), 2)
    
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252704", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252705", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252706", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2126", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("212701", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2128", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    nSubValor2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2528", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior
    nSaldoDiario2 = Round(nSaldoDiario2 + nSubValor2, 2)
    '************* NAGL ERS079-2016 20170407
    
    '************NAGL ERS079-2016 20170407 Otras obligaciones con el público y con instituciones recaudadoras de tributos 30 días
    xlHoja1.Cells(41, 3) = nSaldoDiario1
    xlHoja1.Cells(41, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(41, 3), xlHoja1.Cells(41, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    '***********NAGL ERS079-2016 20170407 Depósitos de empresas del sistema financiero y OFI
    'nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("23", pdFecha, "1")
    'nSaldoDiario2 = Round(oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("23", pdFecha, "2") / pnTipoCambio, 2)
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2312", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2313", pdFecha, "1") + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2318", pdFecha, "1", 0)
    nSaldoDiario2 = Round(oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2322", pdFecha, "2") / pnTipoCambio, 2) + Round(oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("2323", pdFecha, "2") / pnTipoCambio, 2) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2328", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior, 2)
    xlHoja1.Cells(43, 3) = nSaldoDiario1
    xlHoja1.Cells(43, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(43, 3), xlHoja1.Cells(43, 4)).NumberFormat = "#,##0.00;-#,##0.00" '******NAGL
    
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 1, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 1, 1)
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 1, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 1, 1)
    xlHoja1.Cells(44, 3) = nSaldoDiario1
    xlHoja1.Cells(44, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(44, 3), xlHoja1.Cells(44, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    nSaldoDiario1 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 1)
    'nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "1", 3, 1)
    nSaldoDiario2 = oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 1)
    'nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 0) + oDbalanceCont.ObtenerSaldoAdeudadoMyPCP(pdFecha, "2", 3, 1)
    xlHoja1.Cells(45, 3) = nSaldoDiario1
    xlHoja1.Cells(45, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(45, 3), xlHoja1.Cells(45, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue
    
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2514190201", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141903", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141905", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141906", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141911", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141912", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141913", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141914", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25141915", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251501", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251502", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25150301", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25150401", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251505", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251506", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251509", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251601", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25160201", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25160202", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25160203", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251702", pdFecha, "1", 0)
    'nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25170301", pdFecha, "1", 0)
    'nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25170302", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517030301", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517030302", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517030901", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517030902", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("251704", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517050101", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517050102", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517050103", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2517050201", pdFecha, "1", 0)
    
    nSaldoDiario2 = 0
    nSaldoDiario2 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2524190201", pdFecha, "1", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241903", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241905", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241906", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241911", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241912", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241913", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241914", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25241915", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252501", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252502", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25250301", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25250401", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252505", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252506", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252509", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252601", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25260201", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25260202", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25260203", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252702", pdFecha, "2", nTipoCambioAn)
    'nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25270301", pdFecha, "2", nTipoCambioAn)
    'nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("25270302", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527030301", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527030302", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527030901", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527030902", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("252704", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527050101", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527050102", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527050103", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = nSaldoDiario2 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("2527050201", pdFecha, "2", nTipoCambioAn)
    nSaldoDiario2 = Round(nSaldoDiario2 / lnTipoCambioBalanceAnterior, 2) '*******NAGL ERS079-2016 20170407
    
    xlHoja1.Cells(48, 3) = nSaldoDiario1
    xlHoja1.Cells(48, 4) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(48, 3), xlHoja1.Cells(48, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("1116", pdFecha, "1")
    nSaldoDiario2 = Round(oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("1126", pdFecha, "2") / pnTipoCambio, 2)
    nSaldoDiario1 = IIf(nSaldoDiario1 < 0, 0, nSaldoDiario1)
    nSaldoDiario2 = IIf(nSaldoDiario2 < 0, 0, nSaldoDiario2)
    xlHoja1.Cells(25, 3) = nSaldoDiario1
    xlHoja1.Cells(25, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(25, 3), xlHoja1.Cells(25, 4)).NumberFormat = "#,##0.00;-#,##0.00"
        
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "1", nSaldoDiario1, 1, "1200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "2", nSaldoDiario2, 1, "1200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "3", xlHoja1.Cells(25, 6), 1, "1200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(12, pdFecha, "4", xlHoja1.Cells(25, 7), 1, "1200", "B1")
    'ALPA 20140506**************************************************************************************************
'    nSaldoDiario1 = oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("1113", pdFecha, "1") + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("111803", pdFecha, "1", 0)
'    nSaldoDiario2 = Round((oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("1123", pdFecha, "2") / pnTipoCambio), 2) + Round((oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("112803", pdFecha, "2", nTipoCambioAn) / nTipoCambioAn), 2)
    'nSaldoDiario1 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondosxPlazo(pdFecha, 1, "01,02,04", "", 30) + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondosxPlazo(pdFecha, 1, "03", "1090100012521", 30)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("111803", pdFecha, "1", 0)
    'nSaldoDiario2 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondosxPlazo(pdFecha, 2, "01,02,04", "", 30) + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondosxPlazo(pdFecha, 2, "03", "1090100012521", 30)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("112803", pdFecha, "2", 0) / nTipoCambioAn, 2)
    
    '**************NAGL ERS079-2016 20170407 Tasa de Encaje

    nSaldoDiario1 = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100)
    nSaldoDiario2 = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("33") / 100)
    xlHoja1.Cells(10, 10) = nSaldoDiario1
    xlHoja1.Cells(10, 11) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(10, 10), xlHoja1.Cells(10, 11)).NumberFormat = "#,##0.00;-#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(13, 3), xlHoja1.Cells(13, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue
    
    'NAGL ERS079-2016 20170407 *************************************** FONDOS DISPONIBLES EN EL SISTEMA FINANCIERO NACIONAL
    nSaldoDiario1 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 1, "01,02,04", "") + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 1, "03", "1090100012521") - oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1112", pdFecha, 1)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("111803", pdFecha, "1", 0)
    nSaldoDiario2 = (oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 2, "01,02,04", "") + oDbalanceCont.obtenerSumaConsolidadoCtasxConcentracionFondos(pdFecha, 2, "03", "1090100012521") - oDbalanceCont.ObtenerDispobiblesenSFN("1090100822183", "1122", pdFecha, 1)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("112803", pdFecha, "2", 0) / lnTipoCambioBalanceAnterior, 2)
    
    xlHoja1.Cells(26, 3) = nSaldoDiario1
    xlHoja1.Cells(26, 4) = nSaldoDiario2
    xlHoja1.Range(xlHoja1.Cells(26, 3), xlHoja1.Cells(26, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "1", nSaldoDiario1, 1, "1300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "2", nSaldoDiario2, 1, "1300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "3", xlHoja1.Cells(26, 6), 1, "1300", "B1") 'NAGL ERS079-2016 20170407 antes xlHoja1.Cells(25, 6)
    Call oDbalanceCont.InsertaDetallaReporte15A(13, pdFecha, "4", xlHoja1.Cells(26, 7), 1, "1300", "B1") 'NAGL ERS079-2016 20170407 antes xlHoja1.Cells(25, 7)

    nSaldoDiario1 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, "1", "1") ' -oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141102", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141103", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141104", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141109", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141112", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141113", pdFecha, "1") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141111", pdFecha, "1")
    'nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141802", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141803", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141804", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141809", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141812", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141813", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("141811", pdFecha, "1", 0)
    nSaldoDiario2 = oDbalanceCont.ObtenerColocacSaldosDiarioTramos(pdFecha, "1", "2") 'Round((oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("142102", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141203", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141204", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141209", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141212", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141213", pdFecha, "2") + oDbalanceCont.ObtenerCtaSaldoDiarioxMoneda("141211", pdFecha, "2")) / pnTipoCambio, 2)
    'nSaldoDiario2 = nSaldoDiario2 + Round((oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142802", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142803", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142804", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142809", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142812", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142813", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("142811", pdFecha, "2", 0)) / nTipoCambioAn, 2)
    xlHoja1.Cells(29, 3) = nSaldoDiario1
    xlHoja1.Cells(29, 4) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(29, 3), xlHoja1.Cells(29, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "1", nSaldoDiario1, 1, "1600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "2", nSaldoDiario2, 1, "1600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "3", xlHoja1.Cells(29, 6), 1, "1600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(16, pdFecha, "4", xlHoja1.Cells(29, 7), 1, "1600", "B1")
    
    'NAGL ERS079-2016 20170407 (CUENTAS POR COBRAR - OTROS)********************************************************
    'nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("151701", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("151702", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170901", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170902", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170903", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170905", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170907", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170908", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170909", pdFecha, "1", 0)
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("151701", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("151702", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170901", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170905", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170907", pdFecha, "1", 0)
    nSaldoDiario1 = nSaldoDiario1 + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15170908", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15171902", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15171903", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15171909", pdFecha, "1", 0) + oDbalanceCont.ObtenerSumaOperacionesReportesVcto30(pdFecha, "1")
    nSaldoDiario2 = Round((oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("152701", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("152702", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15270901", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15270905", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15270907", pdFecha, "2", 0)) / lnTipoCambioBalanceAnterior, 2)
    nSaldoDiario2 = nSaldoDiario2 + Round(((oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15270908", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15271902", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15271903", pdFecha, "2", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiario("15271909", pdFecha, "2", 0)) / lnTipoCambioBalanceAnterior) + oDbalanceCont.ObtenerSumaOperacionesReportesVcto30(pdFecha, "2"), 2)
    
    xlHoja1.Cells(31, 3) = nSaldoDiario1
    xlHoja1.Cells(31, 4) = nSaldoDiario2
    
    xlHoja1.Range(xlHoja1.Cells(31, 3), xlHoja1.Cells(31, 4)).NumberFormat = "#,##0.00;-#,##0.00"
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "1", nSaldoDiario1, 1, "1800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "2", nSaldoDiario2, 1, "1800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "3", xlHoja1.Cells(31, 6), 1, "1800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(18, pdFecha, "4", xlHoja1.Cells(31, 7), 1, "1800", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "1", xlHoja1.Cells(13, 3), 1, "400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "2", xlHoja1.Cells(13, 4), 1, "400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "3", xlHoja1.Cells(13, 6), 1, "400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(4, pdFecha, "4", xlHoja1.Cells(13, 7), 1, "400", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "1", 0, 1, "500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "2", 0, 1, "500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "3", 0, 1, "500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(5, pdFecha, "4", 0, 1, "500", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "1", 0, 1, "600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "2", 0, 1, "600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "3", 0, 1, "600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(6, pdFecha, "4", 0, 1, "600", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "1", 0, 1, "700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "2", 0, 1, "700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "3", 0, 1, "700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(7, pdFecha, "4", 0, 1, "700", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "1", 0, 1, "800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "2", 0, 1, "800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "3", 0, 1, "800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(8, pdFecha, "4", 0, 1, "800", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "1", 0, 1, "900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "2", 0, 1, "900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "3", 0, 1, "900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(9, pdFecha, "4", 0, 1, "900", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "1", xlHoja1.Cells(20, 3), 1, "1000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "2", xlHoja1.Cells(20, 4), 1, "1000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "3", xlHoja1.Cells(20, 6), 1, "1000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(10, pdFecha, "4", xlHoja1.Cells(20, 7), 1, "1000", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "1", 0, 1, "1100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "2", 0, 1, "1100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "3", 0, 1, "1100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(11, pdFecha, "4", 0, 1, "1100", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "1", 0, 1, "1400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "2", 0, 1, "1400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "3", 0, 1, "1400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(14, pdFecha, "4", 0, 1, "1400", "B1")

    
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "1", 0, 1, "1500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "2", 0, 1, "1500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "3", 0, 1, "1500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(15, pdFecha, "4", 0, 1, "1500", "B1")


    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "1", 0, 1, "1700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "2", 0, 1, "1700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "3", 0, 1, "1700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(17, pdFecha, "4", 0, 1, "1700", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "1", 0, 1, "1900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "2", 0, 1, "1900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "3", 0, 1, "1900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(19, pdFecha, "4", 0, 1, "1900", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "1", xlHoja1.Cells(34, 3), 1, "2000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "2", xlHoja1.Cells(34, 4), 1, "2000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "3", xlHoja1.Cells(34, 6), 1, "2000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(20, pdFecha, "4", xlHoja1.Cells(34, 7), 1, "2000", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "1", 0, 1, "2100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "2", 0, 1, "2100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "3", 0, 1, "2100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(21, pdFecha, "4", 0, 1, "2100", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "1", xlHoja1.Cells(36, 3), 1, "2200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "2", xlHoja1.Cells(36, 4), 1, "2200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "3", xlHoja1.Cells(36, 6), 1, "2200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(22, pdFecha, "4", xlHoja1.Cells(36, 7), 1, "2200", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "1", xlHoja1.Cells(37, 3), 1, "2300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "2", xlHoja1.Cells(37, 4), 1, "2300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "3", xlHoja1.Cells(37, 6), 1, "2300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(23, pdFecha, "4", xlHoja1.Cells(37, 7), 1, "2300", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "1", xlHoja1.Cells(38, 3), 1, "2400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "2", xlHoja1.Cells(38, 4), 1, "2400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "3", xlHoja1.Cells(38, 6), 1, "2400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(24, pdFecha, "4", xlHoja1.Cells(38, 7), 1, "2400", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "1", xlHoja1.Cells(39, 3), 1, "2500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "2", xlHoja1.Cells(39, 4), 1, "2500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "3", xlHoja1.Cells(39, 6), 1, "2500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(25, pdFecha, "4", xlHoja1.Cells(39, 7), 1, "2500", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "1", xlHoja1.Cells(40, 3), 1, "2600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "2", xlHoja1.Cells(40, 4), 1, "2600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "3", xlHoja1.Cells(40, 6), 1, "2600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(26, pdFecha, "4", xlHoja1.Cells(40, 7), 1, "2600", "B1")

    'Call oDbalanceCont.InsertaDetallaReporte15A(27, pdFecha, "1", xlHoja1.Cells(39, 3), 1, "2700", "B1")
    'Call oDbalanceCont.InsertaDetallaReporte15A(27, pdFecha, "2", xlHoja1.Cells(39, 4), 1, "2700", "B1")
    'Call oDbalanceCont.InsertaDetallaReporte15A(27, pdFecha, "3", xlHoja1.Cells(39, 6), 1, "2700", "B1")
    'Call oDbalanceCont.InsertaDetallaReporte15A(27, pdFecha, "4", xlHoja1.Cells(39, 7), 1, "2700", "B1")

     Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "1", xlHoja1.Cells(41, 3), 1, "2800", "B1")
     Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "2", xlHoja1.Cells(41, 4), 1, "2800", "B1")
     Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "3", xlHoja1.Cells(41, 6), 1, "2800", "B1")
     Call oDbalanceCont.InsertaDetallaReporte15A(28, pdFecha, "4", xlHoja1.Cells(41, 7), 1, "2800", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "1", xlHoja1.Cells(42, 3), 1, "2900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "2", xlHoja1.Cells(42, 4), 1, "2900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "3", xlHoja1.Cells(42, 6), 1, "2900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(29, pdFecha, "4", xlHoja1.Cells(42, 7), 1, "2900", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "1", xlHoja1.Cells(43, 3), 1, "3000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "2", xlHoja1.Cells(43, 4), 1, "3000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "3", xlHoja1.Cells(43, 6), 1, "3000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(30, pdFecha, "4", xlHoja1.Cells(43, 7), 1, "3000", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "1", xlHoja1.Cells(44, 3), 1, "3100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "2", xlHoja1.Cells(44, 4), 1, "3100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "3", xlHoja1.Cells(44, 6), 1, "3100", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(31, pdFecha, "4", xlHoja1.Cells(44, 7), 1, "3100", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "1", xlHoja1.Cells(45, 3), 1, "3200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "2", xlHoja1.Cells(45, 4), 1, "3200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "3", xlHoja1.Cells(45, 6), 1, "3200", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(32, pdFecha, "4", xlHoja1.Cells(45, 7), 1, "3200", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "1", xlHoja1.Cells(46, 3), 1, "3300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "2", xlHoja1.Cells(46, 4), 1, "3300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "3", xlHoja1.Cells(46, 6), 1, "3300", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(33, pdFecha, "4", xlHoja1.Cells(46, 7), 1, "3300", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "1", xlHoja1.Cells(47, 3), 1, "3400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "2", xlHoja1.Cells(47, 4), 1, "3400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "3", xlHoja1.Cells(47, 6), 1, "3400", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(34, pdFecha, "4", xlHoja1.Cells(47, 7), 1, "3400", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "1", xlHoja1.Cells(48, 3), 1, "3500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "2", xlHoja1.Cells(48, 4), 1, "3500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "3", xlHoja1.Cells(48, 6), 1, "3500", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(35, pdFecha, "4", xlHoja1.Cells(48, 7), 1, "3500", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "1", xlHoja1.Cells(50, 3), 1, "3600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "2", xlHoja1.Cells(50, 4), 1, "3600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "3", xlHoja1.Cells(50, 6), 1, "3600", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(36, pdFecha, "4", xlHoja1.Cells(50, 7), 1, "3600", "B1")
    
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "1", xlHoja1.Cells(51, 3), 1, "3700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "2", xlHoja1.Cells(51, 4), 1, "3700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "3", xlHoja1.Cells(51, 6), 1, "3700", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(37, pdFecha, "4", xlHoja1.Cells(51, 7), 1, "3700", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "1", xlHoja1.Cells(52, 3), 1, "3800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "2", xlHoja1.Cells(52, 4), 1, "3800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "3", xlHoja1.Cells(52, 6), 1, "3800", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(38, pdFecha, "4", xlHoja1.Cells(52, 7), 1, "3800", "B1")

    
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "1", xlHoja1.Cells(53, 3), 1, "3900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "2", xlHoja1.Cells(53, 4), 1, "3900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "3", xlHoja1.Cells(53, 6), 1, "3900", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(39, pdFecha, "4", xlHoja1.Cells(53, 7), 1, "3900", "B1")

    Call oDbalanceCont.InsertaDetallaReporte15A(40, pdFecha, "1", xlHoja1.Cells(54, 3), 1, "4000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(40, pdFecha, "2", xlHoja1.Cells(54, 4), 1, "4000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(40, pdFecha, "3", xlHoja1.Cells(54, 6), 1, "4000", "B1")
    Call oDbalanceCont.InsertaDetallaReporte15A(40, pdFecha, "4", xlHoja1.Cells(54, 7), 1, "4000", "B1")
    
    oBarra.Progress 10, "ANEXO 15B: Ratio de Cobertura de Liquidez", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
        xlHoja1.SaveAs App.path & lsArchivo1
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Public Sub ReporteAnexo15C()
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
    
    Dim rsRep15B As ADODB.Recordset
    Dim oRep15B As New DbalanceCont
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContadorX As Integer
    Dim nSaltoContadorY As Integer
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Integer
    Dim sCtaAhorro As String
    Dim pdFecha As Date
    Dim X As Integer
    Dim lilineasCol As Integer 'NAGL 20170904
    Dim liLineas As Integer 'NAGL 20170904
    Dim lnPosicion As Integer 'NAGL 20170904
    Dim oDbalanceCont As New DbalanceCont 'NAGL 20170904
    Dim CantDiasTotal As Integer, CantDias As Integer 'NAGL 20170904
    Dim lsTOSE() As String
    Dim nValorAnexo As Currency     'JIPR20200824
    Dim pdFechaFinDeMes As Date     'JIPR20200824
    Dim pdFechaFinDeMesMA As Date   'JIPR20200824
    Dim nDia As Integer             'JIPR20200824
    
    On Error GoTo GeneraExcelErr

    Set Progress = New clsProgressBar
    Progress.ShowForm frmReportes
    Progress.Max = Day(txtFecha.Text) * 4 'Cambiado by NAGL 20170904 - Se paso de 3 a 4
    Progress.Progress 0, "Anexo 15-C: Posición Mensual de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ANEXO_15C"
    'Primera Hoja ******************************************************
    lsNomHoja = "Anx15C" '"Anexo15" 'CPMN
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_15C_" & gsCodUser & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    ReDim lsTOSE(2)
    nSaltoContadorX = 13
    nSaltoContadorY = 4
     
    nMes = cboMes.ListIndex + 1
    CantDias = 1
    lilineasCol = 30
    
    xlHoja1.Cells(9, 3) = "Fecha: " & Format(txtFecha.Text, "YYYY") & "/" & UCase(Format(txtFecha.Text, "MMMM"))
    pdFecha = DateAdd("d", -Day(txtFecha.Text), txtFecha.Text)
    CantDiasTotal = Day(txtFecha.Text) 'NAGL 20170914

    For X = 1 To Day(txtFecha.Text)
    Progress.Progress X, "Anexo 15-C: Posición Mensual de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
    pdFecha = DateAdd("d", 1, pdFecha)
    nSaltoContadorX = 13
    'lilineas = 13
    Set rsRep15B = oRep15B.ObtenerListaReporte15B(Format(pdFecha, "YYYY/MM/DD"), "1", 0, "A1", Format(pdFecha, "YYYY/MM/DD"))
        Do While Not rsRep15B.EOF
                nSaltoContadorY = 2 + X
                xlHoja1.Cells(nSaltoContadorX, nSaltoContadorY) = rsRep15B!nSaldo
                xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, nSaltoContadorY), xlHoja1.Cells(nSaltoContadorX, nSaltoContadorY)).NumberFormat = "#,##0.00;-#,##0.00"
                'Saltos cambiado by NAGL 20190621
                If nSaltoContadorX = 23 Then
                    nSaltoContadorX = 26
                ElseIf nSaltoContadorX = 35 Then
                    nSaltoContadorX = 38
                ElseIf nSaltoContadorX = 46 Then
                    nSaltoContadorX = 56
                Else
                    nSaltoContadorX = nSaltoContadorX + 1
                End If
                '**************************NAGL 20170914 - ERS002-2017******************
                If CantDias >= 28 Then
                   If CantDiasTotal <> 28 Then
                        xlHoja1.Cells(11, lilineasCol) = CantDias
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol), xlHoja1.Cells(11, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                        ExcelCuadro xlHoja1, lilineasCol, 11, CCur(lilineasCol), 11
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol), xlHoja1.Cells(11, lilineasCol)).VerticalAlignment = xlTop
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol), xlHoja1.Cells(11, lilineasCol)).HorizontalAlignment = xlCenter
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol), xlHoja1.Cells(11, lilineasCol)).Font.Bold = True
                        xlHoja1.Cells(24, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(13, lilineasCol), xlHoja1.Cells(23, lilineasCol)).Address(False, False) & ")"
                        xlHoja1.Cells(36, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(26, lilineasCol), xlHoja1.Cells(35, lilineasCol)).Address(False, False) & ")"
                        xlHoja1.Cells(47, lilineasCol).Formula = "=" & "If" & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(13, lilineasCol), xlHoja1.Cells(23, lilineasCol)).Address(False, False) & ")" & "=" & "0" & "," & "0" & "," & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(45, lilineasCol), xlHoja1.Cells(46, lilineasCol)).Address(False, False) & ")" & "/" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(13, lilineasCol), xlHoja1.Cells(23, lilineasCol)).Address(False, False) & ")" & ")" & "*" & "100" & ")"
                        xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                        xlHoja1.Range(xlHoja1.Cells(36, lilineasCol), xlHoja1.Cells(36, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                        xlHoja1.Range(xlHoja1.Cells(47, lilineasCol), xlHoja1.Cells(47, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                        ExcelCuadro xlHoja1, lilineasCol, 24, CCur(lilineasCol), 24
                        ExcelCuadro xlHoja1, lilineasCol, 36, CCur(lilineasCol), 36
                        ExcelCuadro xlHoja1, lilineasCol, 47, CCur(lilineasCol), 47
                        If nSaltoContadorX <> 56 Then
                            If nSaltoContadorX = 14 Then
                                ExcelCuadro xlHoja1, lilineasCol, nSaltoContadorX - 1, CCur(lilineasCol), nSaltoContadorX - 1
                            Else
                                ExcelCuadro xlHoja1, lilineasCol, nSaltoContadorX, CCur(lilineasCol), nSaltoContadorX
                            End If
                            If nSaltoContadorX = 38 Or nSaltoContadorX = 41 Or nSaltoContadorX = 44 Then
                               xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, lilineasCol), xlHoja1.Cells(nSaltoContadorX, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                            End If
                        End If
                   End If
                   If CantDias = CantDiasTotal Then
                        xlHoja1.Cells(11, lilineasCol + 1) = "Promedio Mensual"
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol + 1), xlHoja1.Cells(11, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol + 1), xlHoja1.Cells(11, lilineasCol + 1)).VerticalAlignment = xlTop
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol + 1), xlHoja1.Cells(11, lilineasCol + 1)).HorizontalAlignment = xlCenter
                        xlHoja1.Range(xlHoja1.Cells(11, lilineasCol + 1), xlHoja1.Cells(11, lilineasCol + 1)).Font.Bold = True
                        ExcelCuadro xlHoja1, lilineasCol + 1, 11, CCur(lilineasCol + 1), 11
                        If nSaltoContadorX <> 56 Then
                            If nSaltoContadorX = 14 Then
                                xlHoja1.Cells(nSaltoContadorX - 1, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX - 1, 3), xlHoja1.Cells(nSaltoContadorX - 1, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                                ExcelCuadro xlHoja1, lilineasCol + 1, nSaltoContadorX - 1, CCur(lilineasCol + 1), nSaltoContadorX - 1
                            ElseIf nSaltoContadorX = 46 Then
                                ExcelCuadro xlHoja1, lilineasCol + 1, nSaltoContadorX + 1, CCur(lilineasCol + 1), nSaltoContadorX + 1
                            End If
                            
                            If nSaltoContadorX = 38 Then
                               xlHoja1.Cells(38, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol + 1), xlHoja1.Cells(24, lilineasCol + 1)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(36, lilineasCol + 1), xlHoja1.Cells(36, lilineasCol + 1)).Address(False, False) & "*100"
                            ElseIf nSaltoContadorX = 41 Or nSaltoContadorX = 44 Then
                               xlHoja1.Cells(41, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(39, lilineasCol + 1), xlHoja1.Cells(39, lilineasCol + 1)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(40, lilineasCol + 1), xlHoja1.Cells(40, lilineasCol + 1)).Address(False, False) & "*100"
                               xlHoja1.Cells(44, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(41, lilineasCol + 1), xlHoja1.Cells(41, lilineasCol + 1)).Address(False, False)
                            Else
                               xlHoja1.Cells(nSaltoContadorX, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, 3), xlHoja1.Cells(nSaltoContadorX, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                            End If 'NAGL 20191016 Según Anx03_ERS006-2019
                    
                            xlHoja1.Cells(24, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(24, 3), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                            xlHoja1.Cells(36, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(36, 3), xlHoja1.Cells(36, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                            xlHoja1.Cells(47, lilineasCol + 1).Formula = "=" & "If" & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(13, lilineasCol + 1), xlHoja1.Cells(23, lilineasCol + 1)).Address(False, False) & ")" & "=" & "0" & "," & "0" & "," & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(45, lilineasCol + 1), xlHoja1.Cells(46, lilineasCol + 1)).Address(False, False) & ")" & "/" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(13, lilineasCol + 1), xlHoja1.Cells(23, lilineasCol + 1)).Address(False, False) & ")" & ")" & "*" & "100" & ")"
                            xlHoja1.Range(xlHoja1.Cells(24, lilineasCol + 1), xlHoja1.Cells(24, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                            xlHoja1.Range(xlHoja1.Cells(36, lilineasCol + 1), xlHoja1.Cells(36, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                            xlHoja1.Range(xlHoja1.Cells(47, lilineasCol + 1), xlHoja1.Cells(47, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                            ExcelCuadro xlHoja1, lilineasCol + 1, 24, CCur(lilineasCol + 1), 24
                            ExcelCuadro xlHoja1, lilineasCol + 1, 36, CCur(lilineasCol + 1), 36
                        
                            ExcelCuadro xlHoja1, lilineasCol + 1, nSaltoContadorX, CCur(lilineasCol + 1), nSaltoContadorX
                            If nSaltoContadorX = 38 Or nSaltoContadorX = 41 Or nSaltoContadorX = 44 Then
                               xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, lilineasCol + 1), xlHoja1.Cells(nSaltoContadorX, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                            End If
                         End If
                    End If
                    xlHoja1.Range(xlHoja1.Cells(13, lilineasCol), xlHoja1.Cells(50, lilineasCol + 1)).NumberFormat = "#,###0.00"
                    xlHoja1.Range(xlHoja1.Cells(12, 3), xlHoja1.Cells(12, lilineasCol + 1)).Merge True
                    ExcelCuadro xlHoja1, 3, 12, CCur(lilineasCol + 1), 12
                    xlHoja1.Range(xlHoja1.Cells(25, 3), xlHoja1.Cells(25, lilineasCol + 1)).Merge True
                    ExcelCuadro xlHoja1, lilineasCol + 1, 25, CCur(lilineasCol + 1), 25
                End If '******************FIN NAGL 20170914*****************************
            rsRep15B.MoveNext
            If rsRep15B.EOF Then
               Exit Do
            End If
        Loop
        CantDias = CantDias + 1
        If CantDias > 28 Then
           lilineasCol = lilineasCol + 1
        End If
    Next X
    
    'ME
    nSaltoContadorX = 56
    nSaltoContadorY = 3
    CantDias = 1
    lilineasCol = 30
    pdFecha = DateAdd("d", -Day(txtFecha.Text), txtFecha.Text)
    
     For X = 1 To Day(txtFecha.Text)
        Progress.Progress Day(txtFecha.Text) + X, "Anexo 15-C: Ratio de Cobertura de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
        pdFecha = DateAdd("d", 1, pdFecha)
        Set rsRep15B = oRep15B.ObtenerListaReporte15B(Format(pdFecha, "YYYY/MM/DD"), "2", 0, "A1", pdFecha)
        nSaltoContadorX = 56
        Do While Not rsRep15B.EOF
                nSaltoContadorY = 2 + X
                xlHoja1.Cells(nSaltoContadorX, nSaltoContadorY) = rsRep15B!nSaldo
                xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, nSaltoContadorY), xlHoja1.Cells(nSaltoContadorX, nSaltoContadorY)).NumberFormat = "#,##0.00;-#,##0.00"
                'Saltos cambiado by NAGL 20190621
                If nSaltoContadorX = 68 Then
                    nSaltoContadorX = 71
                ElseIf nSaltoContadorX = 80 Then
                    nSaltoContadorX = 84
                Else
                    nSaltoContadorX = nSaltoContadorX + 1
                End If
                '**************************NAGL 20170914 - ERS002-2017******************
                If CantDias >= 28 Then
                    xlHoja1.Cells(54, lilineasCol) = CantDias
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol), xlHoja1.Cells(54, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                    ExcelCuadro xlHoja1, lilineasCol, 54, CCur(lilineasCol), 54
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol), xlHoja1.Cells(54, lilineasCol)).VerticalAlignment = xlTop
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol), xlHoja1.Cells(54, lilineasCol)).HorizontalAlignment = xlCenter
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol), xlHoja1.Cells(54, lilineasCol)).Font.Bold = True
                    xlHoja1.Cells(69, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(56, lilineasCol), xlHoja1.Cells(68, lilineasCol)).Address(False, False) & ")"
                    xlHoja1.Cells(82, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(71, lilineasCol), xlHoja1.Cells(80, lilineasCol)).Address(False, False) & ")"
                    xlHoja1.Cells(94, lilineasCol).Formula = "=" & "If" & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(56, lilineasCol), xlHoja1.Cells(68, lilineasCol)).Address(False, False) & ")" & "=" & "0" & "," & "0" & "," & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(91, lilineasCol), xlHoja1.Cells(93, lilineasCol)).Address(False, False) & ")" & "/" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(56, lilineasCol), xlHoja1.Cells(68, lilineasCol)).Address(False, False) & ")" & ")" & "*" & "100" & ")"
                    xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                    xlHoja1.Range(xlHoja1.Cells(82, lilineasCol), xlHoja1.Cells(82, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                    xlHoja1.Range(xlHoja1.Cells(94, lilineasCol), xlHoja1.Cells(94, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                    ExcelCuadro xlHoja1, lilineasCol, 69, CCur(lilineasCol), 69
                    ExcelCuadro xlHoja1, lilineasCol, 82, CCur(lilineasCol), 82
                    ExcelCuadro xlHoja1, lilineasCol, 94, CCur(lilineasCol), 94
                    ExcelCuadro xlHoja1, lilineasCol, 81, CCur(lilineasCol), 81
                    If nSaltoContadorX <> 94 Then
                        If nSaltoContadorX = 57 Then
                            ExcelCuadro xlHoja1, lilineasCol, nSaltoContadorX - 1, CCur(lilineasCol), nSaltoContadorX - 1
                        Else
                            ExcelCuadro xlHoja1, lilineasCol, nSaltoContadorX, CCur(lilineasCol), nSaltoContadorX
                        End If
                        If nSaltoContadorX = 84 Or nSaltoContadorX = 87 Or nSaltoContadorX = 90 Then
                           xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, lilineasCol), xlHoja1.Cells(nSaltoContadorX, lilineasCol)).Interior.Color = RGB(153, 153, 255)
                        End If
                    End If
                  End If
                  If CantDias = CantDiasTotal Then
                    xlHoja1.Cells(54, lilineasCol + 1) = "Promedio Mensual"
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol + 1), xlHoja1.Cells(54, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol + 1), xlHoja1.Cells(54, lilineasCol + 1)).VerticalAlignment = xlTop
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol + 1), xlHoja1.Cells(54, lilineasCol + 1)).HorizontalAlignment = xlCenter
                    xlHoja1.Range(xlHoja1.Cells(54, lilineasCol + 1), xlHoja1.Cells(54, lilineasCol + 1)).Font.Bold = True
                    ExcelCuadro xlHoja1, lilineasCol + 1, 54, CCur(lilineasCol + 1), 54
                    If nSaltoContadorX <> 94 Then
                        If nSaltoContadorX = 57 Then
                            xlHoja1.Cells(nSaltoContadorX - 1, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX - 1, 3), xlHoja1.Cells(nSaltoContadorX - 1, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                            ExcelCuadro xlHoja1, lilineasCol + 1, nSaltoContadorX - 1, CCur(lilineasCol + 1), nSaltoContadorX - 1
                        ElseIf nSaltoContadorX = 93 Then
                            ExcelCuadro xlHoja1, lilineasCol + 1, nSaltoContadorX + 1, CCur(lilineasCol + 1), nSaltoContadorX + 1
                            'xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX + 2, lilineasCol + 1), xlHoja1.Cells(nSaltoContadorX + 2, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                        End If
                        
                        If nSaltoContadorX = 84 Then
                               xlHoja1.Cells(84, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol + 1), xlHoja1.Cells(69, lilineasCol + 1)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(82, lilineasCol + 1), xlHoja1.Cells(82, lilineasCol + 1)).Address(False, False) & "*100"
                        ElseIf nSaltoContadorX = 87 Or nSaltoContadorX = 90 Then
                               xlHoja1.Cells(87, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(85, lilineasCol + 1), xlHoja1.Cells(85, lilineasCol + 1)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(86, lilineasCol + 1), xlHoja1.Cells(86, lilineasCol + 1)).Address(False, False) & "*100"
                               xlHoja1.Cells(90, lilineasCol + 1).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(87, lilineasCol + 1), xlHoja1.Cells(87, lilineasCol + 1)).Address(False, False)
                        Else
                               xlHoja1.Cells(nSaltoContadorX, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, 3), xlHoja1.Cells(nSaltoContadorX, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                        End If 'NAGL 20191016 Según Anx03_ERS006-2019
                        
                        xlHoja1.Cells(69, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(69, 3), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                        xlHoja1.Cells(82, lilineasCol + 1).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(82, 3), xlHoja1.Cells(82, lilineasCol)).Address(False, False) & ")" & "," & "2" & ")"
                        xlHoja1.Cells(94, lilineasCol + 1).Formula = "=" & "If" & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(56, lilineasCol + 1), xlHoja1.Cells(68, lilineasCol + 1)).Address(False, False) & ")" & "=" & "0" & "," & "0" & "," & "(" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(91, lilineasCol + 1), xlHoja1.Cells(93, lilineasCol + 1)).Address(False, False) & ")" & "/" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(56, lilineasCol + 1), xlHoja1.Cells(68, lilineasCol + 1)).Address(False, False) & ")" & ")" & "*" & "100" & ")"
                        xlHoja1.Range(xlHoja1.Cells(69, lilineasCol + 1), xlHoja1.Cells(69, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                        xlHoja1.Range(xlHoja1.Cells(82, lilineasCol + 1), xlHoja1.Cells(82, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                        xlHoja1.Range(xlHoja1.Cells(94, lilineasCol + 1), xlHoja1.Cells(94, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                        ExcelCuadro xlHoja1, lilineasCol + 1, 69, CCur(lilineasCol + 1), 69
                        ExcelCuadro xlHoja1, lilineasCol + 1, 82, CCur(lilineasCol + 1), 82
                        ExcelCuadro xlHoja1, lilineasCol + 1, 81, CCur(lilineasCol + 1), 81
                    
                        ExcelCuadro xlHoja1, lilineasCol + 1, nSaltoContadorX, CCur(lilineasCol + 1), nSaltoContadorX
                        If nSaltoContadorX = 84 Or nSaltoContadorX = 87 Or nSaltoContadorX = 90 Then
                           xlHoja1.Range(xlHoja1.Cells(nSaltoContadorX, lilineasCol + 1), xlHoja1.Cells(nSaltoContadorX, lilineasCol + 1)).Interior.Color = RGB(153, 153, 255)
                        End If
                     End If
                    xlHoja1.Range(xlHoja1.Cells(56, lilineasCol), xlHoja1.Cells(94, lilineasCol + 1)).NumberFormat = "#,###0.00"
                    xlHoja1.Range(xlHoja1.Cells(55, 3), xlHoja1.Cells(55, lilineasCol + 1)).Merge True
                    ExcelCuadro xlHoja1, lilineasCol + 1, 55, CCur(lilineasCol + 1), 55
                    xlHoja1.Range(xlHoja1.Cells(70, 3), xlHoja1.Cells(70, lilineasCol + 1)).Merge True
                    ExcelCuadro xlHoja1, lilineasCol + 1, 70, CCur(lilineasCol + 1), 70
                End If '******************FIN NAGL 20170914*****************************
                
            rsRep15B.MoveNext
            If rsRep15B.EOF Then
               Exit Do
            End If
        Loop
        CantDias = CantDias + 1
        If CantDias > 28 Then
           lilineasCol = lilineasCol + 1
        End If
    Next X
    
    '***************************************NAGL 20170904 ERS 002-2017************************************'
    '*******SECCIÓN ENCAJE EXIGIBLE********'
    lilineasCol = 3
    lnPosicion = 1
    xlHoja1.Cells(99, 2) = "Encaje Exigible MN"
    xlHoja1.Cells(100, 2) = "Encaje Exigible ME"
    xlHoja1.Cells(101, 2) = "Encaje Exigible / Activos Liquidos MN"
    xlHoja1.Cells(102, 2) = "Encaje Exigible / Activos Liquidos ME"
    xlHoja1.Range(xlHoja1.Cells(99, 2), xlHoja1.Cells(102, 2)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(99, 2), xlHoja1.Cells(102, 2)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(99, 2), xlHoja1.Cells(102, 2)).Font.Bold = True
    pdFecha = DateAdd("d", -Day(txtFecha.Text), txtFecha.Text)
    
    For X = 1 To Day(txtFecha.Text)
        Progress.Progress Day(txtFecha.Text) * 2 + X, "Anexo 15-C: Ratio de Cobertura de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
        pdFecha = DateAdd("d", 1, pdFecha)
        xlHoja1.Cells(98, lilineasCol) = Format(lnPosicion, "#,##0")
        xlHoja1.Cells(99, lilineasCol) = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "2400")
        xlHoja1.Cells(99, lilineasCol).NumberFormat = "#,##0.00"
        xlHoja1.Cells(100, lilineasCol) = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "2400")
        xlHoja1.Cells(100, lilineasCol).NumberFormat = "#,##0.00"
        nValorAnexo = (oDbalanceCont.ObtenerValorAnexo(100, 101)) 'JIPR20200824
        xlHoja1.Cells(101, lilineasCol).Formula = "=" & "+" & (xlHoja1.Range(xlHoja1.Cells(99, lilineasCol), xlHoja1.Cells(99, lilineasCol)).Address(False, False)) & "*" & (nValorAnexo) & "/" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & "*" & "100"
        'xlHoja1.Cells(101, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(99, lilineasCol), xlHoja1.Cells(99, lilineasCol)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & "*" & "100" 'JIPR20200824
        xlHoja1.Cells(101, lilineasCol).NumberFormat = "#,##0.00"
        nValorAnexo = (oDbalanceCont.ObtenerValorAnexo(100, 102)) 'JIPR20200824
        xlHoja1.Cells(102, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(100, lilineasCol), xlHoja1.Cells(100, lilineasCol)).Address(False, False) & "*" & nValorAnexo & "/" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & "*" & "100"
        'xlHoja1.Cells(102, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(100, lilineasCol), xlHoja1.Cells(100, lilineasCol)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & "*" & "100" JIPR20200824
        xlHoja1.Cells(102, lilineasCol).NumberFormat = "#,##0.00"
        lilineasCol = lilineasCol + 1
        lnPosicion = lnPosicion + 1
    Next X
    
    ExcelCuadro xlHoja1, 3, 98, lilineasCol, 98
    ExcelCuadro xlHoja1, 2, 99, lilineasCol, 99
    ExcelCuadro xlHoja1, 2, 100, lilineasCol, 100
    ExcelCuadro xlHoja1, 2, 101, lilineasCol, 101
    ExcelCuadro xlHoja1, 2, 102, lilineasCol, 102
    xlHoja1.Range(xlHoja1.Cells(98, 3), xlHoja1.Cells(98, lilineasCol)).Font.Bold = True
    
    xlHoja1.Cells(98, lilineasCol) = "Promedio"
    xlHoja1.Range(xlHoja1.Cells(98, 3), xlHoja1.Cells(98, lilineasCol)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(98, 3), xlHoja1.Cells(98, lilineasCol)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(99, lilineasCol).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(99, 3), xlHoja1.Cells(99, lilineasCol - 1)).Address(False, False) & ")" & "," & "2" & ")"
    xlHoja1.Cells(100, lilineasCol).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(100, 3), xlHoja1.Cells(100, lilineasCol - 1)).Address(False, False) & ")" & "," & "2" & ")"
    nValorAnexo = (oDbalanceCont.ObtenerValorAnexo(100, 101)) 'ANPS20210510
    xlHoja1.Cells(101, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(99, lilineasCol), xlHoja1.Cells(99, lilineasCol)).Address(False, False) & "*" & (nValorAnexo) & "/" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & "*" & "100" 'ANPS20210510
    nValorAnexo = (oDbalanceCont.ObtenerValorAnexo(100, 102)) 'ANPS20210510
    xlHoja1.Cells(102, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(100, lilineasCol), xlHoja1.Cells(100, lilineasCol)).Address(False, False) & "*" & (nValorAnexo) & "/" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & "*" & "100" 'ANPS20210510

    xlHoja1.Range(xlHoja1.Cells(99, lilineasCol), xlHoja1.Cells(102, lilineasCol)).NumberFormat = "#,###0.00"
'***************************************************************FIN NAGL 20170904****************************************'

    '***************************************NAGL 20190621 Según Anx02_ERS006-2019************************************'
    '************SECCIÓN TOSE**********'
    lilineasCol = 3
    lnPosicion = 1
    xlHoja1.Cells(105, 2) = "T.O.S.E. MN"
    xlHoja1.Cells(106, 2) = "T.O.S.E. ME"
    xlHoja1.Cells(107, 2) = "Activos Líquidos / TOSE expresado en Soles"
    xlHoja1.Cells(108, 2) = "T.C."
    xlHoja1.Range(xlHoja1.Cells(105, 2), xlHoja1.Cells(107, 2)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(105, 2), xlHoja1.Cells(108, 2)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(105, 2), xlHoja1.Cells(107, 2)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(108, 2), xlHoja1.Cells(108, 3)).Font.Bold = True
    pdFecha = DateAdd("d", -Day(txtFecha.Text), txtFecha.Text)
    
    'Para el Tipo de Cambio
    xlHoja1.Cells(108, 3) = Format(oDbalanceCont.ObtenerTipoCambioCierreNew(txtFecha.Text), "#,##0.000")
    xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Interior.ColorIndex = 6
    xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Font.Color = vbRed
    
    For X = 1 To Day(txtFecha.Text)
        Progress.Progress Day(txtFecha.Text) * 3 + X, "Anexo 15-C: Ratio de Cobertura de Liquidez", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
        pdFecha = DateAdd("d", 1, pdFecha)
        xlHoja1.Cells(104, lilineasCol) = Format(lnPosicion, "#,##0")
        xlHoja1.Cells(105, lilineasCol) = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "1", "A1", "1900")
        xlHoja1.Cells(105, lilineasCol).NumberFormat = "#,##0.00"
        lsTOSE(1) = xlHoja1.Range(xlHoja1.Cells(105, lilineasCol), xlHoja1.Cells(105, lilineasCol)).Address(False, False)
        
        xlHoja1.Cells(106, lilineasCol) = oDbalanceCont.ObtenerActivosLiquidosReporte15A(pdFecha, "2", "A1", "1900")
        xlHoja1.Cells(106, lilineasCol).NumberFormat = "#,##0.00"
        lsTOSE(2) = xlHoja1.Range(xlHoja1.Cells(106, lilineasCol), xlHoja1.Cells(106, lilineasCol)).Address(False, False)
        
        xlHoja1.Cells(107, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & "*" & xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Address(False, False) & ")/" & "Sum" & "(" & lsTOSE(1) & "+" & lsTOSE(2) & "*" & xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Address(False, False) & ")*100"
        xlHoja1.Cells(107, lilineasCol).NumberFormat = "#,##0.00"
        lilineasCol = lilineasCol + 1
        lnPosicion = lnPosicion + 1
    Next X
    
    ExcelCuadro xlHoja1, 3, 104, lilineasCol, 104
    ExcelCuadro xlHoja1, 2, 105, lilineasCol, 105
    ExcelCuadro xlHoja1, 2, 106, lilineasCol, 106
    ExcelCuadro xlHoja1, 2, 107, lilineasCol, 107
    xlHoja1.Range(xlHoja1.Cells(104, 3), xlHoja1.Cells(104, lilineasCol)).Font.Bold = True
    
    xlHoja1.Cells(104, lilineasCol) = "Promedio"
    xlHoja1.Range(xlHoja1.Cells(104, 3), xlHoja1.Cells(104, lilineasCol)).Interior.Color = RGB(153, 153, 255)
    xlHoja1.Range(xlHoja1.Cells(104, 3), xlHoja1.Cells(104, lilineasCol)).HorizontalAlignment = xlCenter
    
    lsTOSE(1) = xlHoja1.Range(xlHoja1.Cells(105, lilineasCol), xlHoja1.Cells(105, lilineasCol)).Address(False, False)
    lsTOSE(2) = xlHoja1.Range(xlHoja1.Cells(106, lilineasCol), xlHoja1.Cells(106, lilineasCol)).Address(False, False)
    xlHoja1.Cells(105, lilineasCol).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(105, 3), xlHoja1.Cells(105, lilineasCol - 1)).Address(False, False) & ")" & "," & "2" & ")"
    xlHoja1.Cells(106, lilineasCol).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(106, 3), xlHoja1.Cells(106, lilineasCol - 1)).Address(False, False) & ")" & "," & "2" & ")"
    xlHoja1.Cells(107, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & "*" & xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Address(False, False) & ")/" & "Sum" & "(" & lsTOSE(1) & "+" & lsTOSE(2) & "*" & xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Address(False, False) & ")*100"
    xlHoja1.Range(xlHoja1.Cells(105, lilineasCol), xlHoja1.Cells(107, lilineasCol)).NumberFormat = "#,###0.00"
    '***************************************************************END NAGL 20190621****************************************'
    
    
   '************SECCIÓN ACTIVOS JIPR20200824- ANPS**********'
    lilineasCol = 3
    lnPosicion = 1
    xlHoja1.Cells(113, 2) = "Activos Líquidos "
    xlHoja1.Cells(114, 2) = "Activos Totales (cuenta 1 de Balance)"
    xlHoja1.Cells(115, 2) = "Activos Líquidos / Activos Totales expresado en Soles"
    xlHoja1.Range(xlHoja1.Cells(112, 2), xlHoja1.Cells(115, 2)).HorizontalAlignment = xlRight
     xlHoja1.Range(xlHoja1.Cells(113, 2), xlHoja1.Cells(115, 2)).Interior.Color = RGB(153, 153, 255)
    pdFecha = DateAdd("d", -Day(txtFecha.Text), txtFecha.Text)

    For X = 1 To Day(txtFecha.Text)
        Progress.Progress Day(txtFecha.Text) * 4 + X, "Anexo 15-C: Activos Líquidos", "CONFIGURANDO HOJA DE CALCULO", "", vbBlue
        pdFecha = DateAdd("d", 1, pdFecha)
        xlHoja1.Cells(112, lilineasCol) = Format(lnPosicion, "#,##0")
        xlHoja1.Cells(113, lilineasCol).Formula = "=" & "Sum" & "(" & xlHoja1.Range(xlHoja1.Cells(24, lilineasCol), xlHoja1.Cells(24, lilineasCol)).Address(False, False) & "+" & xlHoja1.Range(xlHoja1.Cells(69, lilineasCol), xlHoja1.Cells(69, lilineasCol)).Address(False, False) & "*" & xlHoja1.Range(xlHoja1.Cells(108, 3), xlHoja1.Cells(108, 3)).Address(False, False)
        
         pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
        pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
        nDia = Day(gdFecSis)
        Set oDbalanceCont = New dBalanceCont
        If nDia >= 15 Then
            xlHoja1.Cells(114, lilineasCol) = Format(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1", txtFecha.Text, "0", 0), "#,##0.000")
        Else
           xlHoja1.Cells(114, lilineasCol) = Format(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1", pdFechaFinDeMesMA, "0", 0), "#,##0.000")
        End If
        
        xlHoja1.Cells(115, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(113, lilineasCol), xlHoja1.Cells(113, lilineasCol)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(114, lilineasCol), xlHoja1.Cells(114, lilineasCol)).Address(False, False) & "*" & "100"
        xlHoja1.Cells(113, lilineasCol).NumberFormat = "#,##0.00"
        xlHoja1.Cells(114, lilineasCol).NumberFormat = "#,##0.00"
        xlHoja1.Cells(115, lilineasCol).NumberFormat = "#,##0.00"

        lilineasCol = lilineasCol + 1
        lnPosicion = lnPosicion + 1
    Next X
    
    ExcelCuadro xlHoja1, 3, 112, lilineasCol, 112
    ExcelCuadro xlHoja1, 2, 113, lilineasCol, 113
    ExcelCuadro xlHoja1, 2, 114, lilineasCol, 114
    ExcelCuadro xlHoja1, 2, 115, lilineasCol, 115
    xlHoja1.Range(xlHoja1.Cells(113, 2), xlHoja1.Cells(113, 2)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(112, 3), xlHoja1.Cells(112, lilineasCol)).Interior.Color = RGB(153, 153, 255)
    
    
    xlHoja1.Cells(112, lilineasCol) = "Promedio"
    
    xlHoja1.Cells(113, lilineasCol).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(113, 3), xlHoja1.Cells(113, lilineasCol - 1)).Address(False, False) & ")" & "," & "2" & ")"
    xlHoja1.Cells(114, lilineasCol).Formula = "=" & "Round" & "(" & "Average" & "(" & xlHoja1.Range(xlHoja1.Cells(114, 3), xlHoja1.Cells(114, lilineasCol - 1)).Address(False, False) & ")" & "," & "2" & ")"
    xlHoja1.Cells(115, lilineasCol).Formula = "=" & "+" & xlHoja1.Range(xlHoja1.Cells(113, lilineasCol), xlHoja1.Cells(113, lilineasCol)).Address(False, False) & "/" & xlHoja1.Range(xlHoja1.Cells(114, lilineasCol), xlHoja1.Cells(114, lilineasCol)).Address(False, False) & "*" & "100"
    xlHoja1.Range(xlHoja1.Cells(113, lilineasCol), xlHoja1.Cells(115, lilineasCol)).NumberFormat = "#,###0.00"
    
      '************SECCIÓN ACTIVOS JIPR20200824**********'
      
    Progress.CloseForm Me 'ALPA20131031
    xlHoja1.Range("C13:AH13").EntireColumn.AutoFit
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub '********Adecuaciones by NAGL 20190621 Según Anx002-ERS006-2019

''ALPA 20101228**********************************************************************
Public Sub ReporteVerificacionInteresesDiferidos()
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
    Dim rsCreditos As ADODB.Recordset
    Dim oCreditos As New DCreditos
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteValidacionIntDiferidos"
    'Primera Hoja ******************************************************
    lsNomHoja = "CredRefinanciados"
    '*******************************************************************
    lsArchivo1 = "\spooler\RepIntDiferidos" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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

    nSaltoContador = 7

    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.RecuperaDatosParaValidarInteresesDiferidos(txtFecha.Text, "1", CDbl(txtTipCambio.Text))
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    xlHoja1.Cells(2, 7) = Format(txtFecha.Text, "DD") & " DE " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
    '        DoEvents
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 16)).Borders.LineStyle = 1 'FRHU20131014
                xlHoja1.Cells(nSaltoContador, 1) = rsCreditos!cPersCod
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cPersNombre
                xlHoja1.Cells(nSaltoContador, 3) = rsCreditos!cCtaCod
                xlHoja1.Cells(nSaltoContador, 4) = rsCreditos!cMoneda
                xlHoja1.Cells(nSaltoContador, 5) = rsCreditos!CtaCnt
                xlHoja1.Cells(nSaltoContador, 6) = rsCreditos!nIntDiferido
                xlHoja1.Cells(nSaltoContador, 7) = rsCreditos!nSaldoxDiferir
                xlHoja1.Cells(nSaltoContador, 8) = rsCreditos!nCapitalPagado
                xlHoja1.Cells(nSaltoContador, 9) = rsCreditos!nCapitalCreCanxRef
                xlHoja1.Cells(nSaltoContador, 10) = rsCreditos!nInteresCreCanxRef
                xlHoja1.Cells(nSaltoContador, 11) = rsCreditos!nCuotas
                xlHoja1.Cells(nSaltoContador, 12) = rsCreditos!nNroCuota
                xlHoja1.Cells(nSaltoContador, 13) = rsCreditos!nCuotasPendientes
                '** FRHU 20131014
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 14), xlHoja1.Cells(nSaltoContador, 15)).NumberFormat = "dd/mm/yyyy"
                xlHoja1.Cells(nSaltoContador, 14) = rsCreditos!dVenc
                xlHoja1.Cells(nSaltoContador, 15) = rsCreditos!dVigencia
                '** END FRHU
                'ALPA 20151023********************************************
                xlHoja1.Cells(nSaltoContador, 16) = IIf(rsCreditos!nTipoDiferido = 1, "REFINANCIADO", "RECLASIFICADO")
                '*********************************************************
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing

    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
'GeneraExcelErr:
'    MsgBox Err.Description, vbInformation, "Aviso"
'    Exit Sub
End Sub

Public Sub ReporteAdeudadosCalendario()
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
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AdeudCalendario"
    'Primera Hoja ******************************************************
    lsNomHoja = "AdeudCalendario"
    '*******************************************************************
    lsArchivo1 = "\spooler\AdeudCalendario" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    nSaltoContador = 10
     
    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.RecuperaDatosReporteAdeudadosCalendario
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    'xlHoja1.Cells(5, 2) = Format(txtFecha.Text, "DD") & " DE " & UCase(Format(txtFecha.Text, "MMMM")) & " DEL  " & Format(txtFecha.Text, "YYYY")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
    '        DoEvents
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 11)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 2) = "'" & Format(rsCreditos!dVencimiento, "YYYY/MM/DD")
                xlHoja1.Cells(nSaltoContador, 3) = rsCreditos!cAdeudado
                xlHoja1.Cells(nSaltoContador, 4) = rsCreditos!cCtaIFDesc
                xlHoja1.Cells(nSaltoContador, 5) = CDbl(rsCreditos!nCAPITALMN)
                xlHoja1.Cells(nSaltoContador, 6) = CDbl(rsCreditos!nInteresMN)
                xlHoja1.Cells(nSaltoContador, 7) = CDbl(rsCreditos!nCAPITALMN + rsCreditos!nInteresMN)
                xlHoja1.Cells(nSaltoContador, 8) = CDbl(Format(rsCreditos!nCapitalME, "###,###,###0.00"))
                xlHoja1.Cells(nSaltoContador, 9) = CDbl(Format(rsCreditos!nInteresME, "###,###,###0.00"))
                xlHoja1.Cells(nSaltoContador, 10) = CDbl(rsCreditos!nInteresME + rsCreditos!nInteresME)
                
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub

Public Sub ReporteAdeudadosCalendarioVigente()
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
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim dFechaCP As Date
    Dim lsCelda As String
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AdeudCalendarioVigente"
    'Primera Hoja ******************************************************
    lsNomHoja = "AdeudCalendarioVigente"
    '*******************************************************************
    lsArchivo1 = "\spooler\AdeudCalendarioVigente" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
     
    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.RecuperaDatosReporteAdeudadosCalendarioVigente(txtFecha.Text, Mid(gsOpeCod, 3, 1))
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    dFechaCP = DateAdd("D", 360, txtFecha.Text)
    xlHoja1.Cells(1, 3) = "'" & Format(txtFecha.Text, "YYYY/MM/DD")
    xlHoja1.Cells(2, 3) = "'" & Format(dFechaCP, "YYYY/MM/DD")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
    '        DoEvents
                If nSaltoContador = 6 Then
                    xlHoja1.Range("G4").Formula = "'" & Format(rsCreditos!dVencimiento, "YYYY/MM/DD")
                End If
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 9)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = Format(rsCreditos!dVencimiento, "YYYY/MM/DD")
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cAdeudado
                xlHoja1.Cells(nSaltoContador, 3) = Format(rsCreditos!dCtaIFAper, "YYYY/MM/DD")
                xlHoja1.Cells(nSaltoContador, 4) = CDbl(rsCreditos!nCapital)
                xlHoja1.Cells(nSaltoContador, 5) = CDbl(rsCreditos!nFactorC)
                xlHoja1.Cells(nSaltoContador, 6) = CDbl(rsCreditos!nInteres)
                xlHoja1.Cells(nSaltoContador, 7) = CDbl(Format(rsCreditos!nFactorI, "###,###,###0.00"))
                xlHoja1.Cells(nSaltoContador, 8) = CDbl(Format(rsCreditos!nCapitalI, "###,###,###0.00"))
                xlHoja1.Cells(nSaltoContador, 9) = CDbl(rsCreditos!nApagar)
                If DateDiff("d", dFechaCP, rsCreditos!dVencimiento) <= 0 Then
                lsCelda = CStr(nSaltoContador)
                End If
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
        xlHoja1.Range("E1").Formula = "=sum(H6:H" & lsCelda & ")"
    If nSaltoContador > CInt(lsCelda) Then
        xlHoja1.Range("E2").Formula = "=sum(H" & (CInt(lsCelda) + 1) & ":H" & (nSaltoContador - 1) & ")"
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub

Public Sub ReporteAdeudados()
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
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim dFechaCP As Date
    Dim lsCelda As String
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AnalisisCtaAdeudadosACortoyLP"
    'Primera Hoja ******************************************************
    lsNomHoja = "AnalisisCtaAdeudadosACortoyLP"
    '*******************************************************************
    lsArchivo1 = "\spooler\AnalisisCtaAdeudadosACortoyLP" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
     
    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.RecuperaDatosReporteAnalisisDeCtaCyLPlazo(txtFecha.Text, Mid(gsOpeCod, 3, 1))
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
               
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 6)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = rsCreditos!cPersNombre
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cCtaIFDesc
                xlHoja1.Cells(nSaltoContador, 3) = CDbl(rsCreditos!nCapitalCP)
                xlHoja1.Cells(nSaltoContador, 4) = CDbl(rsCreditos!nInteresCP)
                xlHoja1.Cells(nSaltoContador, 5) = CDbl(rsCreditos!nInteresLP)
                xlHoja1.Cells(nSaltoContador, 6) = CDbl(rsCreditos!nCapital - rsCreditos!nCapitalCP)
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub
'EJVG20121002 ***
Public Sub generarRptConcentracionFondos()
    Dim oCtaIf As New NCajaCtaIF
    Dim xlsAplicacion As New Excel.Application
    Dim rsCta As New ADODB.Recordset
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim ldFecha As Date
    Dim i As Long
    Dim oConstSis As New DConstSistemas
    Dim lnTpoCambio As Currency
    
    '********NAGL 20181002*********'
    Dim TituloProgress As String
    Dim MensajeProgress As String
    Dim oBarra As New clsProgressBar
    Dim nprogress As Integer
    '*********END NAGL*************'
    
    On Error GoTo ErrRprConcentracion
    
    '***********NAGL 20181002*****************'
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    nprogress = 0
    oBarra.Progress nprogress, "Reporte de Concentración de Fondos", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "Reporte de Concentración de Fondos"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    '*************END NAGL********************'
    
    lsArchivo = "\spooler\RptConcentracionFondos" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    ldFecha = CDate(txtFecha.Text)
    lnTpoCambio = LeeTpoCambio(Format(ldFecha, "dd/mm/yyyy"))
    fsOtrasIFisRptConcentraFondos = oConstSis.LeeConstSistema(444) 'EJVG20130927
    
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    Call LlenaMatrizConcentraFondos(MatCtasBcosMN, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 1, "01,02", ""))
    Call LlenaMatrizConcentraFondos(MatCtasBcosME, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 2, "01,02", ""))
    
    'ANDE 20170822
    '    Call LlenaMatrizConcentraFondos(MatCtasCMACsMN, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 1, "03", "1090100012521"))
    '    Call LlenaMatrizConcentraFondos(MatCtasCMACsME, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 2, "03", "1090100012521"))
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    Call LlenaMatrizConcentraFondos(MatCtasCMACsMN, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 1, "03,04", "1090100012521"))
    Call LlenaMatrizConcentraFondos(MatCtasCMACsME, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 2, "03,04", "1090100012521"))
    'end ANDE
    'Call LlenaMatrizConcentraFondos(MatCtasCRACsMN, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 1, "04", ""))
    oBarra.Progress 3, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    Call LlenaMatrizConcentraFondos(MatCtasCRACsMN, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 1, fsOtrasIFisRptConcentraFondos, "")) 'EJVG20130927
    'Call LlenaMatrizConcentraFondos(MatCtasCRACsME, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 2, "04", ""))
    Call LlenaMatrizConcentraFondos(MatCtasCRACsME, oCtaIf.obtenerConsolidadoCtasxConcentracionFondos(ldFecha, 2, fsOtrasIFisRptConcentraFondos, "")) 'EJVG20130927

    Set xlsLibro = xlsAplicacion.Workbooks.Add
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    'HOJA SOLES
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = StrConv(gcPEN_PLURAL, vbUpperCase) 'marg ers044-2016
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    Call generaHojaMonedaRptConcentacion(gMonedaNacional, xlsHoja, ldFecha)
    
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    'HOJA DOLARES
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "DOLARES"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    Call generaHojaMonedaRptConcentacion(gMonedaExtranjera, xlsHoja, ldFecha)
    
     oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    'HOJA CONCENTRACION
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "CONCENTRACION"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 10
    xlsHoja.Cells.Font.Bold = True
    Call generaHojaConcentracionRptConcentacion(xlsHoja, ldFecha)
    
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    'HOJA CRUCE CAJAS
    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "CRUCE CAJA"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 10
    xlsHoja.Cells.Font.Bold = True
    Call generaHojaCruceCajasRptConcentracion(xlsHoja, ldFecha)
    
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue 'NAGL 20181001
    'EJVG20131230 ***
    xlsAplicacion.Range("SOLES!B" & lnPosFilaMNTotal + 1) = "SALDOS CONSOLIDADO EN BANCOS, CMACS y OTRAS IFIs: "
    xlsAplicacion.Range("SOLES!D" & lnPosFilaMNTotal + 1) = "TOTAL MN Y ME:"
    xlsAplicacion.Range("SOLES!E" & lnPosFilaMNTotal + 1).NumberFormat = "#,##0.00"
    xlsAplicacion.Range("SOLES!E" & lnPosFilaMNTotal + 1) = "=SOLES!E" & lnPosFilaMNTotal & "+DOLARES!E" & lnPosFilaMETotal & "*" & lnTpoCambio
    xlsAplicacion.Range("SOLES!E" & lnPosFilaMNTotal + 1).Interior.Color = RGB(255, 255, 0)
    xlsAplicacion.Range("SOLES!B" & lnPosFilaMNTotal + 1 & ":" & "SOLES!E" & lnPosFilaMNTotal + 1).Borders.Weight = xlMedium
    'END EJVG *******

    For Each xlHoja1 In xlsLibro.Worksheets
        If UCase(xlHoja1.Name) = "HOJA1" Or UCase(xlHoja1.Name) = "HOJA2" Or UCase(xlHoja1.Name) = "HOJA3" Then
            xlHoja1.Delete
        End If
    Next
    
    '*******NAGL 20181001*********
    oBarra.Progress 10, "Reporte de Concentración de Fondos", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing
    '*********END NAGL*************
    
    MsgBox "Se ha generado satisfactoriamente el reporte de Concentración de Fondos", vbInformation, "Aviso"
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True

    Set oCtaIf = Nothing
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Set oConstSis = Nothing
    Exit Sub
ErrRprConcentracion:
    MsgBox Err.Description, vbCritical, "Aviso"
    Exit Sub
End Sub
Private Sub LlenaMatrizConcentraFondos(ByRef pMat() As TConcentraFondos, ByVal rs As ADODB.Recordset)
    ReDim pMat(rs.RecordCount)
    Do While Not rs.EOF
        pMat(rs.Bookmark - 1).CodPersona = rs!cPersCod
        pMat(rs.Bookmark - 1).Nombre = rs!cPersNombre
        pMat(rs.Bookmark - 1).CtaIFCodCtaCorriente = rs!cCtaIFCodCtaCorriente
        pMat(rs.Bookmark - 1).CtaIFDescCtaCorriente = rs!cCtaIFDescCtaCorriente
        pMat(rs.Bookmark - 1).SaldoCtaCorriente = rs!nSaldoCtaCorriente
        pMat(rs.Bookmark - 1).CtaIFCodCtaAhorro = rs!cCtaIFCodCtaAhorro
        pMat(rs.Bookmark - 1).CtaIFDescCtaAhorro = rs!cCtaIFDescCtaAhorro
        pMat(rs.Bookmark - 1).SaldoCtaAhorro = rs!nSaldoCtaAhorro
        pMat(rs.Bookmark - 1).SaldoTotalInversion = rs!nSaldoInversion
        pMat(rs.Bookmark - 1).SaldoTotalDPFOver = rs!nSaldoDPFOver
        rs.MoveNext
    Loop
End Sub
Public Sub generaHojaCruceCajasRptConcentracion(ByRef xlsHoja As Worksheet, ByVal pdFecha As Date)
    Dim oInstFinan As New DInstFinanc
    Dim oCtaIf As New NCajaCtaIF
    Dim MatIFi() As TConcentraFondos
    Dim rsInstFinan As New ADODB.Recordset, rsCtasAhoEnCaja As New ADODB.Recordset
    Dim lnPosActual As Integer, lnPosAnterior As Integer
    Dim lnTotAhoCteCMACMaynasEnOtrasIfis As Currency, lnTotInversionesCMACMaynasEnOtrasIfis As Currency
    Dim lnTotAhoCteCMACsEnCMACMaynas As Currency, lnTotInversionesCMACsEnCMACMaynas As Currency
    Dim iMat As Integer, iMoneda As Integer, i As Integer
    
    xlsHoja.Columns("A:Z").ColumnWidth = 16
    xlsHoja.Range("A:Z").EntireColumn.Font.Bold = True
    xlsHoja.Range("A:Z").NumberFormat = "#,##0.00"
    xlsHoja.Columns("A:A").ColumnWidth = 4
    xlsHoja.Columns("F:F").ColumnWidth = 0.5
    xlsHoja.Cells(2, 2) = "SALDOS DISPONIBLE AL " & Format(pdFecha, "dd/mm/yyyy")
    xlsHoja.Range("B2").Font.Size = 16
    xlsHoja.Cells(2, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(2, 11) = "T/C.:"
    xlsHoja.Cells(2, 11).HorizontalAlignment = xlRight
    'xlsHoja.Cells(2, 12) = Format(TipoCambioCierre(Year(pdFecha), Month(pdFecha), True), "##,##0.000")
    xlsHoja.Cells(2, 12) = Format(LeeTpoCambio(Format(pdFecha, "dd/mm/yyyy")), "##,##0.000") 'EJVG20131230
    xlsHoja.Cells(2, 12).Font.Size = 12
    xlsHoja.Cells(2, 12).Font.Color = RGB(255, 0, 0)
    xlsHoja.Cells(2, 12).Interior.Color = RGB(255, 255, 0)
    xlsHoja.Cells(2, 12).Borders.Weight = xlMedium
    
    lnPosActual = 4
    lnPosAnterior = lnPosActual
    For iMoneda = 1 To 2
        xlsHoja.Cells(lnPosActual, 2) = "LO QUE CMAC-M TIENE EN OTRAS CMACs"
        xlsHoja.Cells(lnPosActual, 7) = "LO QUE OTRAS CMACs TIENEN EN LA CMAC-M"
        xlsHoja.Range("B" & lnPosActual & ":Z" & lnPosActual).Font.Italic = True
        xlsHoja.Range("B" & lnPosActual & ":Z" & lnPosActual).Font.Color = RGB(0, 0, 255)
        xlsHoja.Cells(lnPosActual + 1, 2) = "CAJAS"
        xlsHoja.Cells(lnPosActual + 2, 2) = "MUNICIPALES"
        xlsHoja.Cells(lnPosActual + 1, 3) = "AHORRO CTE"
        xlsHoja.Cells(lnPosActual + 2, 3) = "Saldo en " & IIf(iMoneda = 1, gcPEN_SIMBOLO, "US$") 'MARG ERS044-2016
        xlsHoja.Cells(lnPosActual + 1, 4) = "SALDOS"
        xlsHoja.Cells(lnPosActual + 2, 4) = IIf(iMoneda = 1, "MN", "ME")
        xlsHoja.Cells(lnPosActual + 1, 5) = "TOTAL"
        xlsHoja.Range("E" & lnPosActual + 1 & ":E" & lnPosActual + 2).MergeCells = True
        xlsHoja.Cells(lnPosActual + 1, 7) = "AHORRO CTE"
        xlsHoja.Cells(lnPosActual + 2, 7) = IIf(iMoneda = 1, "MN", "ME")
        xlsHoja.Cells(lnPosActual + 1, 8) = "CUENTAS"
        xlsHoja.Cells(lnPosActual + 2, 8) = "A P/F"
        xlsHoja.Cells(lnPosActual + 1, 9) = "TOTAL"
        xlsHoja.Range("I" & lnPosActual + 1 & ":I" & lnPosActual + 2).MergeCells = True
        xlsHoja.Cells(lnPosActual + 1, 10) = "SALDOS"
        xlsHoja.Cells(lnPosActual + 2, 10) = "AHORRO CTE"
        xlsHoja.Cells(lnPosActual + 1, 11) = "SALDOS"
        xlsHoja.Cells(lnPosActual + 2, 11) = "PLAZO FIJO"
        xlsHoja.Cells(lnPosActual + 1, 12) = "Saldos en " & IIf(iMoneda = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
        xlsHoja.Cells(lnPosActual + 2, 12) = "favor o en contra"
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).HorizontalAlignment = xlCenter
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).VerticalAlignment = xlCenter
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).Borders(xlEdgeBottom).Weight = xlThin
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).Borders(xlEdgeBottom).LineStyle = xlDouble
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).Borders(xlEdgeTop).Weight = xlThin
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).Borders(xlEdgeLeft).Weight = xlThin
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).Borders(xlEdgeRight).Weight = xlThin
        xlsHoja.Range("B" & lnPosActual + 1 & ":L" & lnPosActual + 2).Borders(xlInsideVertical).Weight = xlThin

        lnPosActual = lnPosActual + 3
        lnPosAnterior = lnPosActual

        Set rsInstFinan = oInstFinan.obtenerInstitucionesFinancieras("03,04", "1090100012521") 'Agregado Param 04 by NAGL 20171116
        Do While Not rsInstFinan.EOF
            lnTotAhoCteCMACMaynasEnOtrasIfis = 0
            lnTotInversionesCMACMaynasEnOtrasIfis = 0

            If iMoneda = 1 Then
                MatIFi = MatCtasCMACsMN
            Else
                MatIFi = MatCtasCMACsME
            End If

            For iMat = 0 To UBound(MatIFi) - 1
                If MatIFi(iMat).CodPersona = rsInstFinan!cPersCod Then
                    lnTotAhoCteCMACMaynasEnOtrasIfis = lnTotAhoCteCMACMaynasEnOtrasIfis + MatIFi(iMat).SaldoCtaCorriente
                    lnTotInversionesCMACMaynasEnOtrasIfis = lnTotInversionesCMACMaynasEnOtrasIfis + MatIFi(iMat).SaldoTotalInversion + MatIFi(iMat).SaldoTotalDPFOver
                End If
            Next
            'Ahorro Cte y DPF de las CMACs en la CMACMaynas
            lnTotAhoCteCMACsEnCMACMaynas = 0
            lnTotInversionesCMACsEnCMACMaynas = 0
            Set rsCtasAhoEnCaja = oCtaIf.obtenerTotalAhorroEnCMACxFecha(pdFecha, iMoneda, rsInstFinan!cPersCod)
            If Not RSVacio(rsCtasAhoEnCaja) Then
                lnTotAhoCteCMACsEnCMACMaynas = rsCtasAhoEnCaja!nSaldoCtaCte
                lnTotInversionesCMACsEnCMACMaynas = rsCtasAhoEnCaja!nSaldoDPF
            End If
            '***
            If lnTotAhoCteCMACMaynasEnOtrasIfis <> 0 Or lnTotInversionesCMACMaynasEnOtrasIfis <> 0 Or lnTotAhoCteCMACsEnCMACMaynas <> 0 Or lnTotInversionesCMACsEnCMACMaynas <> 0 Then
               'If lnTotInversionesCMACMaynasEnOtrasIfis <> 0 Then 'NAGL 20171116 'Comentado by NAGL 20190502
                If lnTotInversionesCMACsEnCMACMaynas > 0 Then 'Agregado by NAGL 20190502 Según INC1904010002
                    '*****NAGL 20190502 Según INC1904010002****'
                    lnTotAhoCteCMACMaynasEnOtrasIfis = 0
                    lnTotInversionesCMACMaynasEnOtrasIfis = 0
                    lnTotAhoCteCMACsEnCMACMaynas = 0
                    '**************END NAGL*******************
                    xlsHoja.Range("B" & lnPosActual) = rsInstFinan!cPersNombre
                    xlsHoja.Range("C" & lnPosActual) = Format(lnTotAhoCteCMACMaynasEnOtrasIfis, gsFormatoNumeroView)
                    xlsHoja.Range("D" & lnPosActual) = Format(lnTotInversionesCMACMaynasEnOtrasIfis, gsFormatoNumeroView)
                    xlsHoja.Range("E" & lnPosActual) = "=C" & lnPosActual & "+D" & lnPosActual
                    xlsHoja.Range("G" & lnPosActual) = Format(lnTotAhoCteCMACsEnCMACMaynas, gsFormatoNumeroView)
                    xlsHoja.Range("H" & lnPosActual) = Format(lnTotInversionesCMACsEnCMACMaynas, gsFormatoNumeroView)
                    xlsHoja.Range("I" & lnPosActual) = "=G" & lnPosActual & "+H" & lnPosActual
                    xlsHoja.Range("J" & lnPosActual) = "=C" & lnPosActual & "-G" & lnPosActual
                    xlsHoja.Range("K" & lnPosActual) = "=D" & lnPosActual & "-H" & lnPosActual
                    xlsHoja.Range("L" & lnPosActual) = "=E" & lnPosActual & "-I" & lnPosActual
                    lnPosActual = lnPosActual + 1
                End If
            End If
            rsInstFinan.MoveNext
        Loop
        If lnPosActual = lnPosAnterior Then lnPosActual = lnPosActual + 1
        xlsHoja.Range("B" & lnPosActual) = "TOTAL " & IIf(iMoneda = 1, StrConv(gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
        xlsHoja.Range("C" & lnPosActual) = "=SUM(C" & lnPosAnterior & ":C" & lnPosActual - 1 & ")"
        xlsHoja.Range("D" & lnPosActual) = "=SUM(D" & lnPosAnterior & ":D" & lnPosActual - 1 & ")"
        xlsHoja.Range("E" & lnPosActual) = "=SUM(E" & lnPosAnterior & ":E" & lnPosActual - 1 & ")"
        xlsHoja.Range("G" & lnPosActual) = "=SUM(G" & lnPosAnterior & ":G" & lnPosActual - 1 & ")"
        xlsHoja.Range("H" & lnPosActual) = "=SUM(H" & lnPosAnterior & ":H" & lnPosActual - 1 & ")"
        xlsHoja.Range("I" & lnPosActual) = "=SUM(I" & lnPosAnterior & ":I" & lnPosActual - 1 & ")"
        xlsHoja.Range("J" & lnPosActual) = "=SUM(J" & lnPosAnterior & ":J" & lnPosActual - 1 & ")"
        xlsHoja.Range("K" & lnPosActual) = "=SUM(K" & lnPosAnterior & ":K" & lnPosActual - 1 & ")"
        xlsHoja.Range("L" & lnPosActual) = "=SUM(L" & lnPosAnterior & ":L" & lnPosActual - 1 & ")"
    
        xlsHoja.Range("B" & lnPosAnterior & ":L" & lnPosActual - 1).Borders(xlEdgeBottom).Weight = xlThin
        xlsHoja.Range("B" & lnPosAnterior & ":L" & lnPosActual - 1).Borders(xlEdgeTop).Weight = xlThin
        xlsHoja.Range("B" & lnPosAnterior & ":L" & lnPosActual - 1).Borders(xlEdgeLeft).Weight = xlThin
        xlsHoja.Range("B" & lnPosAnterior & ":L" & lnPosActual - 1).Borders(xlEdgeRight).Weight = xlThin
        xlsHoja.Range("B" & lnPosAnterior & ":L" & lnPosActual - 1).Borders(xlInsideVertical).Weight = xlThin
        xlsHoja.Range("B" & lnPosActual & ":L" & lnPosActual).Borders.Weight = xlThin
        
        xlsHoja.Range("E" & lnPosAnterior - 2 & ":E" & lnPosActual).Interior.Color = IIf(iMoneda = 1, RGB(255, 255, 153), RGB(204, 255, 204))
        xlsHoja.Range("I" & lnPosAnterior - 2 & ":I" & lnPosActual).Interior.Color = IIf(iMoneda = 1, RGB(255, 255, 153), RGB(204, 255, 204))
        xlsHoja.Range("B" & lnPosActual & ":I" & lnPosActual).Interior.Color = IIf(iMoneda = 1, RGB(255, 255, 153), RGB(204, 255, 204))
        xlsHoja.Range("F" & lnPosAnterior - 2 & ":F" & lnPosActual).Interior.Color = RGB(128, 128, 128)
        
        lnPosActual = lnPosActual + 3
        lnPosAnterior = lnPosActual
    Next

    xlsHoja.Range("B" & lnPosActual) = "TOTAL RESUMEN"
    xlsHoja.Range("C" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",C4:" & "C" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",C4:" & "C" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("D" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",D4:" & "D" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",D4:" & "D" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("E" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",E4:" & "E" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",E4:" & "E" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("G" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",G4:" & "G" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",G4:" & "G" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("H" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",H4:" & "H" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",H4:" & "H" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("I" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",I4:" & "I" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",I4:" & "I" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("J" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",J4:" & "J" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",J4:" & "J" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("K" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",K4:" & "K" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",K4:" & "K" & lnPosActual - 1 & ")*L2"
    xlsHoja.Range("L" & lnPosActual) = "=SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL SOLES"",L4:" & "L" & lnPosActual - 1 & ")+SUMIF(B4:B" & lnPosActual - 1 & ",""TOTAL DOLARES"",L4:" & "L" & lnPosActual - 1 & ")*L2"
    
    xlsHoja.Range("B" & lnPosActual & ":L" & lnPosActual).Borders.Weight = xlThin
    xlsHoja.Range("C" & lnPosActual & ":I" & lnPosActual).Interior.Color = IIf(iMoneda = 1, RGB(255, 255, 153), RGB(204, 255, 204))
    xlsHoja.Range("F" & lnPosActual & ":F" & lnPosActual).Interior.Color = RGB(128, 128, 128)
    xlsHoja.Range("L" & lnPosActual & ":L" & lnPosActual).Interior.Color = RGB(0, 255, 255)
    xlsHoja.Range("B" & lnPosActual + 1) = "(EXPRESADO EN " & StrConv(gcPEN_PLURAL, vbUpperCase) & ")" 'marg ers044-2016
    xlsHoja.Range("B" & lnPosActual + 1).Font.Italic = True
    xlsHoja.Range("B" & lnPosActual + 1).Font.Size = 8
End Sub
Public Sub generaHojaConcentracionRptConcentacion(ByRef xlsHoja As Worksheet, ByVal pdFecha As Date)
    Dim oInstFinan As New DInstFinanc
    Dim rsInstFinan As New ADODB.Recordset
    Dim lnTotalIFiMN As Currency, lnTotalIFiME As Currency
    Dim i As Integer, lnPosActual As Integer, lnPosAnterior As Integer, lnPosPrimBanco As Integer, lnPosUltBanco As Integer, lnPosPrimCMAC As Integer, lnPosUltCMAC As Integer, lnPosPrimCRAC As Integer, lnPosUltCRAC As Integer
    Dim lnSubSoles As String, lnSubDolares As String, lnSubDolaresConver As String, lnTotal As String
    Dim iMat As Integer

    xlsHoja.Columns("A:A").ColumnWidth = 2
    xlsHoja.Columns("B:B").ColumnWidth = 25
    xlsHoja.Columns("C:C").ColumnWidth = 15
    xlsHoja.Columns("D:D").ColumnWidth = 15
    xlsHoja.Columns("E:E").ColumnWidth = 15
    xlsHoja.Columns("F:F").ColumnWidth = 15
    xlsHoja.Columns("G:G").ColumnWidth = 15
    xlsHoja.Columns("H:H").ColumnWidth = 15
    xlsHoja.Columns("I:I").ColumnWidth = 15
    xlsHoja.Columns("J:J").ColumnWidth = 15
    
    
    xlsHoja.Cells(2, 2) = "CAJA MAYNAS S.A"
    xlsHoja.Cells(2, 10).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
    xlsHoja.Cells(2, 10) = pdFecha 'Format(pdFecha, "dd/mm/yyyy")
    xlsHoja.Cells(2, 10).Borders.Weight = xlMedium
    xlsHoja.Cells(2, 10).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(2, 10).HorizontalAlignment = xlCenter
    
    xlsHoja.Cells(3, 9) = "TIPO CAMBIO SBS"
    'xlsHoja.Cells(3, 10) = Format(TipoCambioCierre(Year(pdFecha), Month(pdFecha), True), "##,##0.000")
    xlsHoja.Cells(3, 10) = Format(LeeTpoCambio(Format(pdFecha, "dd/mm/yyyy")), "##,##0.000") 'EJVG20131230
    xlsHoja.Range("I3", "J4").Borders.Weight = xlMedium
    xlsHoja.Cells(2, 2).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    
    xlsHoja.Cells(3, 2) = "AREA DE TESORERIA"
    xlsHoja.Cells(4, 2) = "CONTROL DE CONCENTRACION DE FONDOS DE BANCOS"
    xlsHoja.Cells(4, 2).Font.Size = 13
    
    'BANCOS Y FINANCIERAS
    xlsHoja.Range("B4", "J4").MergeCells = True
    xlsHoja.Range("B5", "J5").MergeCells = True
    xlsHoja.Range("B4", "J4").HorizontalAlignment = xlCenter
    
    xlsHoja.Range("B4", "J4").Borders.Weight = xlMedium
    xlsHoja.Range("B5", "J5").Borders.Weight = xlMedium
    xlsHoja.Range("B4", "J4").Interior.Color = RGB(204, 255, 255)
    
    xlsHoja.Cells(6, 2) = "BANCOS"
    xlsHoja.Range("B6", "B8").MergeCells = True
    xlsHoja.Cells(6, 3) = StrConv(gcPEN_PLURAL, vbUpperCase) 'marg ers044-2016
    xlsHoja.Range("C6", "C8").MergeCells = True
    xlsHoja.Cells(6, 4) = "DOLARES"
    xlsHoja.Range("D6", "E7").MergeCells = True
    xlsHoja.Cells(8, 4) = "US$"
    xlsHoja.Cells(8, 5) = gcPEN_SIMBOLO 'MARG ERS044-2016
    
    xlsHoja.Cells(6, 6) = "TOTAL"
    xlsHoja.Range("F6", "F7").MergeCells = True
    xlsHoja.Cells(8, 6) = gcPEN_SIMBOLO 'marg ers044-2016
    
    xlsHoja.Range("B6", "J8").Borders.Weight = xlMedium
    xlsHoja.Range("B6", "J8").HorizontalAlignment = xlCenter
    xlsHoja.Range("B6", "J8").VerticalAlignment = xlCenter
    xlsHoja.Range("B6", "J8").Interior.Color = RGB(192, 192, 192)
    
    xlsHoja.Range("G6").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("G7").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("H6").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("H7").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("I6").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("I7").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("J6").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    xlsHoja.Range("J7").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    
    xlsHoja.Range("G6") = "% CONCENTRACION"
    xlsHoja.Range("G7") = "EN FUNCION AL"
    xlsHoja.Range("G8") = "PATRIMONIO EFECTIVO"
    xlsHoja.Range("H7") = "ACEPTABLE(+)"
    xlsHoja.Range("H8") = "EXCESO(-)"
    xlsHoja.Range("I6") = "% CONCENTRACION"
    xlsHoja.Range("I7") = "EN FUNCION AL"
    xlsHoja.Range("I8") = "FONDOS TOTALES"
    xlsHoja.Range("J7") = "ACEPTABLE(+)"
    xlsHoja.Range("J8") = "EXCESO(-)"
    
    xlsHoja.Range("G6", "J8").Font.Color = RGB(0, 0, 255)
    xlsHoja.Range("G6", "J8").Font.Size = 7
    xlsHoja.Range("G6", "J8").Font.Name = "Arial Narrow"
    xlsHoja.Range("G6", "J8").Font.Bold = False
    
    xlsHoja.Range("H8").Font.Color = RGB(255, 0, 0)
    xlsHoja.Range("J8").Font.Color = RGB(255, 0, 0)
    lnPosActual = 9
    lnPosPrimBanco = lnPosActual
    
    Set rsInstFinan = oInstFinan.obtenerInstitucionesFinancieras("01,02", "1090100822183")
    For i = 0 To rsInstFinan.RecordCount - 1
        lnTotalIFiMN = 0
        lnTotalIFiME = 0

        For iMat = 0 To UBound(MatCtasBcosMN) - 1
            If MatCtasBcosMN(iMat).CodPersona = rsInstFinan!cPersCod Then
                lnTotalIFiMN = lnTotalIFiMN + MatCtasBcosMN(iMat).SaldoCtaCorriente + MatCtasBcosMN(iMat).SaldoCtaAhorro + MatCtasBcosMN(iMat).SaldoTotalInversion + MatCtasBcosMN(iMat).SaldoTotalDPFOver
            End If
        Next
        For iMat = 0 To UBound(MatCtasBcosME) - 1
            If MatCtasBcosME(iMat).CodPersona = rsInstFinan!cPersCod Then
                lnTotalIFiME = lnTotalIFiME + MatCtasBcosME(iMat).SaldoCtaCorriente + MatCtasBcosME(iMat).SaldoCtaAhorro + MatCtasBcosME(iMat).SaldoTotalInversion + MatCtasBcosME(iMat).SaldoTotalDPFOver
            End If
        Next
        If lnTotalIFiMN <> 0 Or lnTotalIFiME <> 0 Then
            xlsHoja.Range("B" & lnPosActual) = rsInstFinan!cPersNombre
            xlsHoja.Range("C" & lnPosActual) = Format(lnTotalIFiMN, gsFormatoNumeroView)
            xlsHoja.Range("D" & lnPosActual) = Format(lnTotalIFiME, gsFormatoNumeroView)
            xlsHoja.Range("E" & lnPosActual) = "=D" & lnPosActual & "*$J$3"
            xlsHoja.Range("F" & lnPosActual) = "=C" & lnPosActual & "+E" & lnPosActual
            lnPosActual = lnPosActual + 1
        End If
        rsInstFinan.MoveNext
    Next
    
    lnPosUltBanco = lnPosActual - 1
    If lnPosActual <> 9 Then
        xlsHoja.Range("B" & lnPosActual) = "SALDOS"
        xlsHoja.Range("C" & lnPosActual) = "=SUM(C9:C" & (lnPosActual - 1) & ")"
        xlsHoja.Range("D" & lnPosActual) = "=SUM(D9:D" & (lnPosActual - 1) & ")"
        xlsHoja.Range("E" & lnPosActual) = "=SUM(E9:E" & (lnPosActual - 1) & ")"
        xlsHoja.Range("F" & lnPosActual) = "=SUM(F9:F" & (lnPosActual - 1) & ")"
    Else
        xlsHoja.Range("B" & lnPosActual) = "SALDOS"
        xlsHoja.Range("C" & lnPosActual) = 0
        xlsHoja.Range("D" & lnPosActual) = 0
        xlsHoja.Range("E" & lnPosActual) = 0
        xlsHoja.Range("F" & lnPosActual) = 0
    End If

    xlsHoja.Range("F" & lnPosActual).Interior.Color = RGB(255, 255, 0)
    lnSubSoles = "=C" & lnPosActual
    lnSubDolares = "=D" & lnPosActual
    lnSubDolaresConver = "=E" & lnPosActual
    lnTotal = "=F" & lnPosActual
    
    xlsHoja.Range("B9", "J" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("B9", "J" & lnPosActual).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    xlsHoja.Range("B" & lnPosActual, "J" & lnPosActual).Borders.Weight = xlMedium
    
    xlsHoja.Range("B9", "F" & lnPosActual).NumberFormat = "#,##0.00"
    
    'CMACs
    lnPosActual = lnPosActual + 2
    
    xlsHoja.Range("B6", "J8").Copy
    xlsHoja.Range("B" & lnPosActual).PasteSpecial
    xlsHoja.Range("B" & lnPosActual) = "CMACs"
    
    lnPosActual = lnPosActual + 3
    lnPosAnterior = lnPosActual
    lnPosPrimCMAC = lnPosActual
    
    Set rsInstFinan = oInstFinan.obtenerInstitucionesFinancieras("03,04") 'Agregado Param 04 by NAGL 20171116
    For i = 0 To rsInstFinan.RecordCount - 1
        lnTotalIFiMN = 0
        lnTotalIFiME = 0

        For iMat = 0 To UBound(MatCtasCMACsMN) - 1
            If MatCtasCMACsMN(iMat).CodPersona = rsInstFinan!cPersCod Then
                lnTotalIFiMN = lnTotalIFiMN + MatCtasCMACsMN(iMat).SaldoCtaCorriente + MatCtasCMACsMN(iMat).SaldoCtaAhorro + MatCtasCMACsMN(iMat).SaldoTotalInversion + MatCtasCMACsMN(iMat).SaldoTotalDPFOver
            End If
        Next
        For iMat = 0 To UBound(MatCtasCMACsME) - 1
            If MatCtasCMACsME(iMat).CodPersona = rsInstFinan!cPersCod Then
                lnTotalIFiME = lnTotalIFiME + MatCtasCMACsME(iMat).SaldoCtaCorriente + MatCtasCMACsME(iMat).SaldoCtaAhorro + MatCtasCMACsME(iMat).SaldoTotalInversion + MatCtasCMACsME(iMat).SaldoTotalDPFOver
            End If
        Next
        If lnTotalIFiMN <> 0 Or lnTotalIFiME <> 0 Then
            xlsHoja.Range("B" & lnPosActual) = rsInstFinan!cPersNombre
            xlsHoja.Range("C" & lnPosActual) = Format(lnTotalIFiMN, gsFormatoNumeroView)
            xlsHoja.Range("D" & lnPosActual) = Format(lnTotalIFiME, gsFormatoNumeroView)
            xlsHoja.Range("E" & lnPosActual) = "=D" & lnPosActual & "*$J$3"
            xlsHoja.Range("F" & lnPosActual) = "=C" & lnPosActual & "+E" & lnPosActual
            lnPosActual = lnPosActual + 1
        End If
        rsInstFinan.MoveNext
    Next
    lnPosUltCMAC = lnPosActual - 1
    
    If lnPosAnterior <> lnPosActual Then
        xlsHoja.Range("B" & lnPosActual) = "SALDOS"
        xlsHoja.Range("C" & lnPosActual) = "=SUM(C" & lnPosAnterior & ":C" & (lnPosActual - 1) & ")"
        xlsHoja.Range("D" & lnPosActual) = "=SUM(D" & lnPosAnterior & ":D" & (lnPosActual - 1) & ")"
        xlsHoja.Range("E" & lnPosActual) = "=SUM(E" & lnPosAnterior & ":E" & (lnPosActual - 1) & ")"
        xlsHoja.Range("F" & lnPosActual) = "=SUM(F" & lnPosAnterior & ":F" & (lnPosActual - 1) & ")"
    Else
        xlsHoja.Range("B" & lnPosActual) = "SALDOS"
        xlsHoja.Range("C" & lnPosActual) = 0
        xlsHoja.Range("D" & lnPosActual) = 0
        xlsHoja.Range("E" & lnPosActual) = 0
        xlsHoja.Range("F" & lnPosActual) = 0
    End If
    xlsHoja.Range("F" & lnPosActual).Interior.Color = RGB(255, 255, 0)
    lnSubSoles = lnSubSoles & "+C" & lnPosActual
    lnSubDolares = lnSubDolares & "+D" & lnPosActual
    lnSubDolaresConver = lnSubDolaresConver & "+E" & lnPosActual
    lnTotal = lnTotal & "+F" & lnPosActual
    
    xlsHoja.Range("B" & lnPosAnterior, "J" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosAnterior, "J" & lnPosActual).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    xlsHoja.Range("B" & lnPosActual, "J" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("B9", "F" & lnPosActual).NumberFormat = "#,##0.00"
    
    'Otras IFis
    lnPosActual = lnPosActual + 2
    xlsHoja.Range("B6", "J8").Copy
    xlsHoja.Range("B" & lnPosActual).PasteSpecial
    'xlsHoja.Range("B" & lnPosActual) = "CRACs"
    xlsHoja.Range("B" & lnPosActual) = "Otras IFIs" 'EJVG20130927
    
    lnPosActual = lnPosActual + 3
    lnPosAnterior = lnPosActual
    lnPosPrimCRAC = lnPosActual
    
    'Set rsInstFinan = oInstFinan.obtenerInstitucionesFinancieras("04")
    Set rsInstFinan = oInstFinan.obtenerInstitucionesFinancieras(fsOtrasIFisRptConcentraFondos) 'EJVG20130927
    For i = 0 To rsInstFinan.RecordCount - 1
        lnTotalIFiMN = 0
        lnTotalIFiME = 0

        For iMat = 0 To UBound(MatCtasCRACsMN) - 1
            If MatCtasCRACsMN(iMat).CodPersona = rsInstFinan!cPersCod Then
                lnTotalIFiMN = lnTotalIFiMN + MatCtasCRACsMN(iMat).SaldoCtaCorriente + MatCtasCRACsMN(iMat).SaldoCtaAhorro + MatCtasCRACsMN(iMat).SaldoTotalInversion + MatCtasCRACsMN(iMat).SaldoTotalDPFOver
            End If
        Next
        For iMat = 0 To UBound(MatCtasCRACsME) - 1
            If MatCtasCRACsME(iMat).CodPersona = rsInstFinan!cPersCod Then
                lnTotalIFiME = lnTotalIFiME + MatCtasCRACsME(iMat).SaldoCtaCorriente + MatCtasCRACsME(iMat).SaldoCtaAhorro + MatCtasCRACsME(iMat).SaldoTotalInversion + MatCtasCRACsME(iMat).SaldoTotalDPFOver
            End If
        Next
        If lnTotalIFiMN <> 0 Or lnTotalIFiME <> 0 Then
            xlsHoja.Range("B" & lnPosActual) = rsInstFinan!cPersNombre
            xlsHoja.Range("C" & lnPosActual) = Format(lnTotalIFiMN, gsFormatoNumeroView)
            xlsHoja.Range("D" & lnPosActual) = Format(lnTotalIFiME, gsFormatoNumeroView)
            xlsHoja.Range("E" & lnPosActual) = "=D" & lnPosActual & "*$J$3"
            xlsHoja.Range("F" & lnPosActual) = "=C" & lnPosActual & "+E" & lnPosActual
            lnPosActual = lnPosActual + 1
        End If
        rsInstFinan.MoveNext
    Next
    lnPosUltCRAC = lnPosActual - 1
    
    If lnPosAnterior <> lnPosActual Then
        xlsHoja.Range("B" & lnPosActual) = "SALDOS"
        xlsHoja.Range("C" & lnPosActual) = "=SUM(C" & lnPosAnterior & ":C" & (lnPosActual - 1) & ")"
        xlsHoja.Range("D" & lnPosActual) = "=SUM(D" & lnPosAnterior & ":D" & (lnPosActual - 1) & ")"
        xlsHoja.Range("E" & lnPosActual) = "=SUM(E" & lnPosAnterior & ":E" & (lnPosActual - 1) & ")"
        xlsHoja.Range("F" & lnPosActual) = "=SUM(F" & lnPosAnterior & ":F" & (lnPosActual - 1) & ")"
    Else
        xlsHoja.Range("B" & lnPosActual) = "SALDOS"
        xlsHoja.Range("C" & lnPosActual) = 0
        xlsHoja.Range("D" & lnPosActual) = 0
        xlsHoja.Range("E" & lnPosActual) = 0
        xlsHoja.Range("F" & lnPosActual) = 0
    End If
    xlsHoja.Range("F" & lnPosActual).Interior.Color = RGB(255, 255, 0)
    lnSubSoles = lnSubSoles & "+C" & lnPosActual
    lnSubDolares = lnSubDolares & "+D" & lnPosActual
    lnSubDolaresConver = lnSubDolaresConver & "+E" & lnPosActual
    lnTotal = lnTotal & "+F" & lnPosActual
    
    xlsHoja.Range("B" & lnPosAnterior, "J" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosAnterior, "J" & lnPosActual).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    xlsHoja.Range("B" & lnPosActual, "J" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("B9", "F" & lnPosActual).NumberFormat = "#,##0.00"

    'FONDOS
    lnPosActual = lnPosActual + 2
    xlsHoja.Range("B" & lnPosActual) = "FONDOS TOTALES"
    xlsHoja.Range("B" & lnPosActual).Interior.Color = RGB(0, 255, 0)
    xlsHoja.Range("B" & lnPosActual).Font.Color = RGB(0, 0, 255)
    xlsHoja.Range("C" & lnPosActual) = lnSubSoles
    xlsHoja.Range("D" & lnPosActual) = lnSubDolares
    xlsHoja.Range("E" & lnPosActual) = lnSubDolaresConver
    xlsHoja.Range("C" & lnPosActual, "E" & lnPosActual).Interior.Color = RGB(255, 255, 153)
    xlsHoja.Range("F" & lnPosActual) = lnTotal
    xlsHoja.Range("F" & lnPosActual).Interior.Color = RGB(0, 0, 255)
    xlsHoja.Range("F" & lnPosActual).Font.Color = RGB(255, 255, 255)
    
    lnPosActual = lnPosActual + 2
    xlsHoja.Range("B" & lnPosActual) = "PATRIMONIO"
    xlsHoja.Range("B" & lnPosActual + 1) = "EFECTIVO"
    xlsHoja.Range("B" & lnPosActual + 2) = "DEL MES"
    
    'Patrimonio Efectivo *************************************
    Dim lnMesAnterior As Date
    Dim lnMontoPatrimEfect As Currency
    lnMesAnterior = DateAdd("M", -1, pdFecha)
    Dim rsCtaCont As New ADODB.Recordset
    Set rsCtaCont = CargaDatosPatrimonioEfec(Year(lnMesAnterior), Format(Month(lnMesAnterior), "00"), 0)
    lnMontoPatrimEfect = GeneraReporte780030(Year(lnMesAnterior), Month(lnMesAnterior), rsCtaCont, False)
    '*********************************************************
    xlsHoja.Range("C" & lnPosActual + 1) = lnMontoPatrimEfect
    xlsHoja.Range("C" & lnPosActual + 1).NumberFormat = "#,##0.00"
    
    xlsHoja.Range("B" & lnPosActual + 3) = "CONCENTRACION"
    xlsHoja.Range("B" & lnPosActual + 4) = "30% DEL"
    xlsHoja.Range("B" & lnPosActual + 5) = "PATRIMONIO (por ifi)"
    xlsHoja.Range("B" & lnPosActual, "B" & lnPosActual + 2).Interior.Color = RGB(204, 255, 255)
    
    xlsHoja.Range("C" & lnPosActual + 4) = "=C" & (lnPosActual + 1) & "*30%"
    xlsHoja.Range("C" & lnPosActual + 4).NumberFormat = "#,##0.00"
    
    xlsHoja.Range("B" & lnPosActual + 6) = "CONCENTRACION"
    xlsHoja.Range("B" & lnPosActual + 7) = "50% DE LOS"
    xlsHoja.Range("B" & lnPosActual + 8) = "FONDOS TOTALES (por ifi)"
    
    xlsHoja.Range("C" & lnPosActual + 7) = "=F" & (lnPosActual - 2) & "*50%"
    xlsHoja.Range("C" & lnPosActual + 7).NumberFormat = "#,##0.00"
    
    xlsHoja.Range("B" & lnPosActual - 2, "J" & lnPosActual + 12).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual - 2, "J" & lnPosActual + 12).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    xlsHoja.Range("B" & lnPosActual - 2, "J" & lnPosActual + 12).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    
    xlsHoja.Range("B" & lnPosActual - 2, "F" & lnPosActual - 2).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual + 6, "C" & lnPosActual + 8).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual + 6, "C" & lnPosActual + 8).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    
    xlsHoja.Range("B" & lnPosActual, "C" & lnPosActual + 2).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual, "C" & lnPosActual + 2).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    xlsHoja.Range("B" & lnPosActual + 3, "C" & lnPosActual + 5).Borders.Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual + 3, "C" & lnPosActual + 5).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    
    xlsHoja.Range("B" & lnPosActual, "B" & lnPosActual + 8).HorizontalAlignment = xlCenter
    
    For i = lnPosPrimBanco To lnPosUltBanco
        xlsHoja.Range("G" & i).NumberFormat = "0.00%"
        xlsHoja.Range("G" & i) = "=F" & i & "/$C$" & (lnPosActual + 1)
        xlsHoja.Range("H" & i) = "=F" & i & "-$C$" & (lnPosActual + 4)
        xlsHoja.Range("I" & i).NumberFormat = "0.00%"
        xlsHoja.Range("I" & i) = "=F" & i & "/$F$" & (lnPosActual - 2)
        xlsHoja.Range("J" & i) = "=F" & i & "-$C$" & (lnPosActual + 7)
    Next
    For i = lnPosPrimCMAC To lnPosUltCMAC
        xlsHoja.Range("G" & i).NumberFormat = "0.00%"
        xlsHoja.Range("G" & i) = "=F" & i & "/$C$" & (lnPosActual + 1)
        xlsHoja.Range("H" & i) = "=F" & i & "-$C$" & (lnPosActual + 4)
        xlsHoja.Range("I" & i).NumberFormat = "0.00%"
        xlsHoja.Range("I" & i) = "=F" & i & "/$F$" & (lnPosActual - 2)
        xlsHoja.Range("J" & i) = "=F" & i & "-$C$" & (lnPosActual + 7)
    Next
    For i = lnPosPrimCRAC To lnPosUltCRAC
        xlsHoja.Range("G" & i).NumberFormat = "0.00%"
        xlsHoja.Range("G" & i) = "=F" & i & "/$C$" & (lnPosActual + 1)
        xlsHoja.Range("H" & i) = "=F" & i & "-$C$" & (lnPosActual + 4)
        xlsHoja.Range("I" & i).NumberFormat = "0.00%"
        xlsHoja.Range("I" & i) = "=F" & i & "/$F$" & (lnPosActual - 2)
        xlsHoja.Range("J" & i) = "=F" & i & "-$C$" & (lnPosActual + 7)
    Next
    Set oInstFinan = Nothing
    Set rsInstFinan = Nothing
End Sub
Public Sub generaHojaMonedaRptConcentacion(ByVal pnTpoMoneda As Moneda, ByRef xlsHoja As Worksheet, ByVal pdFecha As Date)
    Dim oCtaIf As New NCajaCtaIF
    '********NAGL 20181001****************
    Dim DAnexRiesg As New DAnexoRiesgos
    Dim oEnc As New NEncajeBCR
    Dim rsEncDiario As New ADODB.Recordset
    '********END NAGL*********************
    Dim rsDetalleCtas As New ADODB.Recordset
    Dim MatIFi() As TConcentraFondos
    Dim lsCodPersAnt As String
    
    Dim lsNomHoja As String, lsArchivo As String, lsFormulaTotal As String
    Dim ldFechaRep, ldFechaTC As Date
    Dim lnPosActual As Integer, lnPosAnterior As Integer
    Dim lnAcumDPOvernight As Currency, lnAcumInversiones As Currency
    Dim i As Integer
    Dim iMat As Integer
    
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
        
    xlsHoja.Columns("A:A").ColumnWidth = 2
    
    xlsHoja.Cells(3, 3) = "INFORME ESTADISTICO:  SALDOS EN INSTITUCIONES FINANCIERAS"
    xlsHoja.Cells(3, 3).Font.Size = 13
    xlsHoja.Range("B3", "I3").MergeCells = True
    xlsHoja.Cells(3, 3).HorizontalAlignment = 3
    
    xlsHoja.Cells(4, 3) = "(EN  MONEDA  " & IIf(pnTpoMoneda = gMonedaNacional, "NACIONAL", "EXTRANJERA") & ")"
    xlsHoja.Cells(4, 3).Font.Size = 11
    xlsHoja.Range("B4", "I4").MergeCells = True
    xlsHoja.Cells(4, 3).HorizontalAlignment = 3
    
    xlsHoja.Columns("B:B").ColumnWidth = 30
    xlsHoja.Columns("C:C").ColumnWidth = 25
    xlsHoja.Columns("D:D").ColumnWidth = 13
    xlsHoja.Columns("E:E").ColumnWidth = 25
    xlsHoja.Columns("F:F").ColumnWidth = 13
    xlsHoja.Columns("G:G").ColumnWidth = 13
    xlsHoja.Columns("H:H").ColumnWidth = 13
    xlsHoja.Columns("I:I").ColumnWidth = 13
    
    xlsHoja.Cells(5, 2) = "AREA: TESORERIA"

    xlsHoja.Range("G5", "H5").MergeCells = True
    
    xlsHoja.Cells(5, 7) = "FECHA DE REPORTE"
    xlsHoja.Cells(5, 7).HorizontalAlignment = 3
    xlsHoja.Cells(5, 7).Font.Size = 10
    
    xlsHoja.Cells(5, 9).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
    xlsHoja.Cells(5, 9) = pdFecha 'Format(pdFecha, gsFormatoFechaView)
    xlsHoja.Cells(5, 9).HorizontalAlignment = 3
    xlsHoja.Cells(5, 9).Font.Size = 10
    xlsHoja.Cells(5, 9).Interior.Color = RGB(255, 204, 153)
    
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders.LineStyle = xlContinuous
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeTop).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeBottom).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeLeft).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders(xlEdgeRight).Weight = xlMedium 'Ancho del borde de Linea
    xlsHoja.Range(xlsHoja.Cells(5, 7), xlsHoja.Cells(5, 9)).Borders.ColorIndex = xlAutomatic 'Color del Borde de Linea
    
    'A. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN BANCOS
    xlsHoja.Cells(6, 2) = "A. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN BANCOS Y FINANCIERAS"
    xlsHoja.Cells(6, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(6, 2).Font.Size = 10
   
    ldFechaRep = pdFecha
    If FechaEsFinMes(ldFechaRep) = True Then
        ldFechaTC = ldFechaRep
    Else
        ldFechaTC = DateAdd("D", -Day(ldFechaRep), ldFechaRep)
    End If
    
    xlsHoja.Cells(6, 7) = "TIPO DE CAMBIO FIJO:"
    'xlsHoja.Cells(6, 9) = Format(TipoCambioCierre(Year(pdFecha), Month(pdFecha), True), "##,##0.000")
    xlsHoja.Cells(6, 9) = Format(LeeTpoCambio(Format(pdFecha, "dd/mm/yyyy")), "##,##0.000")
    xlsHoja.Range(xlsHoja.Cells(6, 7), xlsHoja.Cells(6, 9)).Borders.Weight = xlMedium
    
    xlsHoja.Cells(7, 2) = "BANCOS Y FINANCIERAS"
    xlsHoja.Cells(7, 3) = "CUENTAS CORRIENTES"
    xlsHoja.Cells(7, 3).HorizontalAlignment = 3
    xlsHoja.Cells(7, 5) = "CUENTAS DE AHORROS"
    xlsHoja.Cells(7, 5).HorizontalAlignment = 3
    xlsHoja.Cells(7, 7) = "CERTIF. BANK.;"
    xlsHoja.Cells(7, 8) = "PLAZOS FIJOS -"
    xlsHoja.Cells(7, 9) = "TOTAL"
    
    xlsHoja.Cells(8, 3) = "Nro Cta."
    xlsHoja.Cells(8, 3).HorizontalAlignment = 3
    xlsHoja.Cells(8, 4) = "Saldo en " & IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'MARG ERS044-2016
    xlsHoja.Cells(8, 4).HorizontalAlignment = 3
    xlsHoja.Cells(8, 5) = "Nro Cta."
    xlsHoja.Cells(8, 5).HorizontalAlignment = 3
    xlsHoja.Cells(8, 6) = "Saldo en " & IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'MARG ERS044-2016
    xlsHoja.Cells(8, 6).HorizontalAlignment = 3
    xlsHoja.Cells(8, 7) = "FONDOS MUTUOS"
    xlsHoja.Cells(8, 8) = "OVERNIGHT"
    
    xlsHoja.Range("B7", "B8").MergeCells = True
    xlsHoja.Cells(7, 2).HorizontalAlignment = 3
    xlsHoja.Cells(7, 2).VerticalAlignment = 2
    
    xlsHoja.Range("C7", "D7").MergeCells = True
    xlsHoja.Range("E7", "F7").MergeCells = True
    
    xlsHoja.Range("I7", "I8").MergeCells = True
    xlsHoja.Cells(7, 9).HorizontalAlignment = 3
    xlsHoja.Cells(7, 9).VerticalAlignment = 2
    
    xlsHoja.Range("I7", "I8").MergeCells = True
    xlsHoja.Cells(7, 10).HorizontalAlignment = 3
    xlsHoja.Cells(7, 10).VerticalAlignment = 2
    
    xlsHoja.Range("B7", "I7").Borders.LineStyle = xlContinuous
    xlsHoja.Range("B7", "I7").Borders.Weight = xlMedium
    xlsHoja.Range("B7", "I7").Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("B8", "I8").Borders.LineStyle = xlContinuous
    xlsHoja.Range("B8", "I8").Borders.Weight = xlMedium
    xlsHoja.Range("B8", "I8").Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("A1", "I8").Font.Bold = True
    xlsHoja.Range("B7", "I8").Interior.Color = RGB(255, 255, 153)
    
    '***
    If pnTpoMoneda = gMonedaNacional Then
        MatIFi = MatCtasBcosMN
    Else
        MatIFi = MatCtasBcosME
    End If
    
    lsCodPersAnt = ""
    lnPosActual = 9
    
    For iMat = 0 To UBound(MatIFi) - 1
        If lsCodPersAnt = MatIFi(iMat).CodPersona Then
            xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 2), xlsHoja.Cells(lnPosActual, 2)).MergeCells = True
            xlsHoja.Cells(lnPosActual - 1, 2).HorizontalAlignment = xlLeft
            xlsHoja.Cells(lnPosActual - 1, 2).VerticalAlignment = xlCenter
        End If
        xlsHoja.Cells(lnPosActual, 2) = MatIFi(iMat).Nombre
        xlsHoja.Cells(lnPosActual, 3) = MatIFi(iMat).CtaIFDescCtaCorriente
        xlsHoja.Cells(lnPosActual, 4) = Format(MatIFi(iMat).SaldoCtaCorriente, gsFormatoNumeroView)
        xlsHoja.Cells(lnPosActual, 5) = MatIFi(iMat).CtaIFDescCtaAhorro
        xlsHoja.Cells(lnPosActual, 6) = Format(MatIFi(iMat).SaldoCtaAhorro, gsFormatoNumeroView)
        xlsHoja.Cells(lnPosActual, 7) = Format(MatIFi(iMat).SaldoTotalInversion, gsFormatoNumeroView)
        xlsHoja.Cells(lnPosActual, 8) = Format(MatIFi(iMat).SaldoTotalDPFOver, gsFormatoNumeroView)
        lsCodPersAnt = MatIFi(iMat).CodPersona
        lnPosActual = lnPosActual + 1
    Next
    '***
    For i = 9 To lnPosActual - 1
        If xlsHoja.Cells(i, 3) = "" Then
            xlsHoja.Cells(i, 3).Interior.Color = RGB(192, 192, 192)
        End If
        If xlsHoja.Cells(i, 5) = "" Then
            xlsHoja.Cells(i, 5).Interior.Color = RGB(192, 192, 192)
        End If
        xlsHoja.Cells(i, 9).Formula = "= D" & i & "+ F" & i & "+ G" & i & "+H" & i
    Next
    
    xlsHoja.Range("I9", "I" & (lnPosActual - 1)).Borders.LineStyle = xlContinuous
    xlsHoja.Range("I9", "I" & (lnPosActual - 1)).Borders.Weight = xlThin
    xlsHoja.Range("I9", "I" & (lnPosActual - 1)).Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("B" & lnPosActual, "C" & lnPosActual).MergeCells = True
    xlsHoja.Range("B" & lnPosActual).value = "SUB TOTAL"
    xlsHoja.Range("D" & lnPosActual, "I" & lnPosActual).NumberFormat = "#,##0.00" 'Formato de Celda General, Numero, Fecha
    xlsHoja.Range("D" & lnPosActual).Formula = "=SUM(D9:D" & (lnPosActual - 1) & ")"
    xlsHoja.Range("F" & lnPosActual).Formula = "=SUM(F9:F" & (lnPosActual - 1) & ")"
    xlsHoja.Range("G" & lnPosActual).Formula = "=SUM(G9:G" & (lnPosActual - 1) & ")"
    xlsHoja.Range("H" & lnPosActual).Formula = "=SUM(H9:H" & (lnPosActual - 1) & ")"
    xlsHoja.Range("I" & lnPosActual).Formula = "=SUM(I9:I" & (lnPosActual - 1) & ")"
    
    lsFormulaTotal = "=I" & lnPosActual
    
    xlsHoja.Range("B9", "I" & lnPosActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("B9", "I" & lnPosActual).Borders.Weight = xlThin
    xlsHoja.Range("B9", "I" & lnPosActual).Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("B9", "B" & lnPosActual).Borders(xlEdgeLeft).Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual, "I" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("I9", "I" & lnPosActual).Borders(xlEdgeRight).Weight = xlMedium

    'A1. INVERSIONES DE LA CAJA MAYNAS EN BANCOS Y FINANCIERAS
    lnPosActual = lnPosActual + 2
    xlsHoja.Cells(lnPosActual, 2) = "A1. INVERSIONES DE LA CAJA MAYNAS EN BANCOS Y FINANCIERAS"
    xlsHoja.Cells(lnPosActual, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(lnPosActual, 2).Font.Size = 10
    xlsHoja.Cells(lnPosActual, 2).Font.Bold = True
    
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    xlsHoja.Cells(lnPosActual, 2) = "BANCO"
    xlsHoja.Cells(lnPosActual, 3) = "DEPOSITO"
    xlsHoja.Cells(lnPosActual, 4) = "TEA"
    xlsHoja.Cells(lnPosActual, 5) = IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'MARG ERS044-2016
    xlsHoja.Cells(lnPosActual, 6) = "APERTURA"
    xlsHoja.Cells(lnPosActual, 7) = "VENCIMIENTO"
    xlsHoja.Cells(lnPosActual, 8) = "OBSERVACION"
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).Interior.Color = RGB(255, 204, 153)
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).Font.Bold = True
    
    Set rsDetalleCtas = oCtaIf.obtenerDPFyOvernyInversion(ldFechaRep, pnTpoMoneda, "01,02", "")
    For i = 0 To rsDetalleCtas.RecordCount - 1
        lnPosActual = lnPosActual + 1
        xlsHoja.Cells(lnPosActual, 2) = rsDetalleCtas!cPersNombre
        xlsHoja.Cells(lnPosActual, 3) = rsDetalleCtas!cDeposito
        xlsHoja.Cells(lnPosActual, 4).NumberFormat = "0.00%"
        xlsHoja.Cells(lnPosActual, 4) = (rsDetalleCtas!TEA / 100)
        xlsHoja.Cells(lnPosActual, 5).NumberFormat = "#,##0.00"
        xlsHoja.Cells(lnPosActual, 5) = Format(rsDetalleCtas!nSaldo, gsFormatoNumeroView)
        xlsHoja.Range(xlsHoja.Cells(lnPosActual, 6), xlsHoja.Cells(lnPosActual, 7)).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
        xlsHoja.Cells(lnPosActual, 6) = rsDetalleCtas!dCtaIFAper 'Format(rsDetalleCtas!dCtaIFAper, gsFormatoFechaView)
        xlsHoja.Cells(lnPosActual, 7) = rsDetalleCtas!dCtaIFVenc 'Format(rsDetalleCtas!dCtaIFVenc, gsFormatoFechaView)
        rsDetalleCtas.MoveNext
    Next
    xlsHoja.Range("B" & lnPosAnterior, "H" & lnPosActual).Borders.Weight = xlThin
    
    'B. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN OTRAS CAJAS MUNICIPALES
    lnPosActual = lnPosActual + 2
    xlsHoja.Cells(lnPosActual, 2) = "B. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN OTRAS CAJAS MUNICIPALES Y CRAC'S" 'ANDE 20170824 modificación del titulo
    xlsHoja.Cells(lnPosActual, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(lnPosActual, 2).Font.Size = 10
    xlsHoja.Cells(lnPosActual, 2).Font.Bold = True
    
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    xlsHoja.Cells(lnPosActual, 2) = "CAJAS MUNICIPALES"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 2)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 3) = "CUENTAS AHORRO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 3), xlsHoja.Cells(lnPosActual, 4)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 3).HorizontalAlignment = xlCenter
    xlsHoja.Cells(lnPosActual + 1, 3) = "Nro. Cta"
    xlsHoja.Cells(lnPosActual + 1, 4) = "Saldo en " & IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'MARG ERS044-2016
    xlsHoja.Cells(lnPosActual, 5) = "CUENTAS A PLAZO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 5), xlsHoja.Cells(lnPosActual + 1, 5)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 6) = "T.E.A. AHORRO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 6), xlsHoja.Cells(lnPosActual + 1, 6)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 7) = "TOTAL"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 7), xlsHoja.Cells(lnPosActual + 1, 7)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 8) = "PLAZO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 8), xlsHoja.Cells(lnPosActual + 1, 8)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 9) = "VENCIMIENTO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 9), xlsHoja.Cells(lnPosActual + 1, 9)).MergeCells = True
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 9)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 9)).VerticalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 9)).Font.Bold = True
    
    xlsHoja.Range(xlsHoja.Cells(lnPosAnterior, 2), xlsHoja.Cells(lnPosActual + 1, 9)).Borders.Weight = xlMedium
    
    lnPosActual = lnPosActual + 2
    lnPosAnterior = lnPosActual
    
    '***
    If pnTpoMoneda = gMonedaNacional Then
        MatIFi = MatCtasCMACsMN
    Else
        MatIFi = MatCtasCMACsME
    End If
    lsCodPersAnt = ""
    
    For iMat = 0 To UBound(MatIFi) - 1
         If lsCodPersAnt = MatIFi(iMat).CodPersona Then
             xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 2), xlsHoja.Cells(lnPosActual, 2)).MergeCells = True
             xlsHoja.Cells(lnPosActual - 1, 2).HorizontalAlignment = xlLeft
             xlsHoja.Cells(lnPosActual - 1, 2).VerticalAlignment = xlCenter
         End If
         xlsHoja.Cells(lnPosActual, 2) = MatIFi(iMat).Nombre
         xlsHoja.Cells(lnPosActual, 3) = MatIFi(iMat).CtaIFDescCtaAhorro
         xlsHoja.Cells(lnPosActual, 4) = Format(MatIFi(iMat).SaldoCtaAhorro, gsFormatoNumeroView)
         xlsHoja.Cells(lnPosActual, 5) = Format(MatIFi(iMat).SaldoTotalInversion + MatIFi(iMat).SaldoTotalDPFOver, gsFormatoNumeroView)
         lsCodPersAnt = MatIFi(iMat).CodPersona
         lnPosActual = lnPosActual + 1
    Next
    '***
    For i = lnPosAnterior To lnPosActual - 1
        If xlsHoja.Cells(i, 3) = "" Then
            xlsHoja.Cells(i, 3).Interior.Color = RGB(192, 192, 192)
        End If
        If xlsHoja.Cells(i, 6) = "" Then
            xlsHoja.Cells(i, 6).Interior.Color = RGB(192, 192, 192)
        End If
        xlsHoja.Cells(i, 7).Formula = "= D" & i & "+ E" & i
    Next
    
    xlsHoja.Range("B" & lnPosActual, "C" & lnPosActual).MergeCells = True
    xlsHoja.Range("B" & lnPosActual).value = "TOTALES"
    If lnPosAnterior <> lnPosActual Then
        xlsHoja.Range("D" & lnPosActual).Formula = "=SUM(D" & lnPosAnterior & ":D" & (lnPosActual - 1) & ")"
        xlsHoja.Range("E" & lnPosActual).Formula = "=SUM(E" & lnPosAnterior & ":E" & (lnPosActual - 1) & ")"
        xlsHoja.Range("G" & lnPosActual).Formula = "=SUM(G" & lnPosAnterior & ":G" & (lnPosActual - 1) & ")"
    Else
        xlsHoja.Range("D" & lnPosActual).value = 0
        xlsHoja.Range("E" & lnPosActual).value = 0
        xlsHoja.Range("G" & lnPosActual).value = 0
    End If
    xlsHoja.Range("D" & lnPosActual, "G" & lnPosActual).NumberFormat = "#,##0.00"
    
    lsFormulaTotal = lsFormulaTotal & "+G" & lnPosActual
    
    xlsHoja.Range("B" & lnPosAnterior, "I" & lnPosActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("B" & lnPosAnterior, "I" & lnPosActual).Borders.Weight = xlThin
    xlsHoja.Range("B" & lnPosAnterior, "I" & lnPosActual).Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("B" & lnPosAnterior, "B" & lnPosActual).Borders(xlEdgeLeft).Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual, "I" & lnPosActual).Borders.Weight = xlMedium
    xlsHoja.Range("I" & lnPosAnterior, "I" & lnPosActual).Borders(xlEdgeRight).Weight = xlMedium
        
    'B1. INVERSIONES DE CAJA MAYNAS EN CMACs
    lnPosActual = lnPosActual + 2
    xlsHoja.Cells(lnPosActual, 2) = "'B1. INVERSIONES DE CAJA MAYNAS EN CMACs"
    xlsHoja.Cells(lnPosActual, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(lnPosActual, 2).Font.Size = 10
    xlsHoja.Cells(lnPosActual, 2).Font.Bold = True
    
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    xlsHoja.Cells(lnPosActual, 2) = "BANCO"
    xlsHoja.Cells(lnPosActual, 3) = "DEPOSITO"
    xlsHoja.Cells(lnPosActual, 4) = "TEA"
    xlsHoja.Cells(lnPosActual, 5) = IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'MARG ERS044-2016
    xlsHoja.Cells(lnPosActual, 6) = "APERTURA"
    xlsHoja.Cells(lnPosActual, 7) = "VENCIMIENTO"
    xlsHoja.Cells(lnPosActual, 8) = "OBSERVACION"
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).Interior.Color = RGB(255, 204, 153)
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).Font.Bold = True
    
    Set rsDetalleCtas = oCtaIf.obtenerDPFyOvernyInversion(ldFechaRep, pnTpoMoneda, "03,04", "") 'ANDE 20170822 agregado del valor 04 al parametro
    For i = 0 To rsDetalleCtas.RecordCount - 1
        lnPosActual = lnPosActual + 1
        xlsHoja.Cells(lnPosActual, 2) = rsDetalleCtas!cPersNombre
        xlsHoja.Cells(lnPosActual, 3) = rsDetalleCtas!cDeposito
        xlsHoja.Cells(lnPosActual, 4).NumberFormat = "0.00%"
        xlsHoja.Cells(lnPosActual, 4) = (rsDetalleCtas!TEA / 100)
        xlsHoja.Cells(lnPosActual, 5).NumberFormat = "#,##0.00"
        xlsHoja.Cells(lnPosActual, 5) = Format(rsDetalleCtas!nSaldo, gsFormatoNumeroView)
        xlsHoja.Range(xlsHoja.Cells(lnPosActual, 6), xlsHoja.Cells(lnPosActual, 7)).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
        xlsHoja.Cells(lnPosActual, 6) = rsDetalleCtas!dCtaIFAper 'Format(rsDetalleCtas!dCtaIFAper, gsFormatoFechaView)
        xlsHoja.Cells(lnPosActual, 7) = rsDetalleCtas!dCtaIFVenc 'Format(rsDetalleCtas!dCtaIFVenc, gsFormatoFechaView)
        rsDetalleCtas.MoveNext
    Next

    xlsHoja.Range("B" & lnPosAnterior, "H" & lnPosActual).Borders.Weight = xlThin
    '******************************
    'C. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN OTRAS IFIs
    lnPosActual = lnPosActual + 2
    'xlsHoja.Cells(lnPosActual, 2) = "C. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN CAJA RURALES"
    xlsHoja.Cells(lnPosActual, 2) = "C. CONSOLIDADO DE DEPOSITOS DE CAJA MAYNAS EN OTRAS IFIs" 'EJVG20130927
    xlsHoja.Cells(lnPosActual, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(lnPosActual, 2).Font.Size = 10
    xlsHoja.Cells(lnPosActual, 2).Font.Bold = True
    
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    'xlsHoja.Cells(lnPosActual, 2) = "CAJAS RURALES"
    xlsHoja.Cells(lnPosActual, 2) = "OTRAS IFIs" 'EJVG20130927
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 2)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 3) = "CUENTAS AHORRO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 3), xlsHoja.Cells(lnPosActual, 4)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 3).HorizontalAlignment = xlCenter
    xlsHoja.Cells(lnPosActual + 1, 3) = "Nro. Cta"
    xlsHoja.Cells(lnPosActual + 1, 4) = "Saldo en " & IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'marg ers044-2016
    xlsHoja.Cells(lnPosActual, 5) = "CUENTAS A PLAZO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 5), xlsHoja.Cells(lnPosActual + 1, 5)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 6) = "T.E.A. AHORRO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 6), xlsHoja.Cells(lnPosActual + 1, 6)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 7) = "TOTAL"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 7), xlsHoja.Cells(lnPosActual + 1, 7)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 8) = "PLAZO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 8), xlsHoja.Cells(lnPosActual + 1, 8)).MergeCells = True
    xlsHoja.Cells(lnPosActual, 9) = "VENCIMIENTO"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 9), xlsHoja.Cells(lnPosActual + 1, 9)).MergeCells = True
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 9)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 9)).VerticalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual + 1, 9)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(lnPosAnterior, 2), xlsHoja.Cells(lnPosActual + 1, 9)).Borders.Weight = xlMedium
    
    lnPosActual = lnPosActual + 2
    lnPosAnterior = lnPosActual

    If pnTpoMoneda = gMonedaNacional Then
        MatIFi = MatCtasCRACsMN
    Else
        MatIFi = MatCtasCRACsME
    End If
    lsCodPersAnt = ""
    
    For iMat = 0 To UBound(MatIFi) - 1
         If lsCodPersAnt = MatIFi(iMat).CodPersona Then
             xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 2), xlsHoja.Cells(lnPosActual, 2)).MergeCells = True
             xlsHoja.Cells(lnPosActual - 1, 2).HorizontalAlignment = xlLeft
             xlsHoja.Cells(lnPosActual - 1, 2).VerticalAlignment = xlCenter
         End If
         xlsHoja.Cells(lnPosActual, 2) = MatIFi(iMat).Nombre
         xlsHoja.Cells(lnPosActual, 3) = MatIFi(iMat).CtaIFDescCtaAhorro
         xlsHoja.Cells(lnPosActual, 4) = Format(MatIFi(iMat).SaldoCtaAhorro, gsFormatoNumeroView)
         xlsHoja.Cells(lnPosActual, 5) = Format(MatIFi(iMat).SaldoTotalInversion + MatIFi(iMat).SaldoTotalDPFOver, gsFormatoNumeroView)
         lsCodPersAnt = MatIFi(iMat).CodPersona
         lnPosActual = lnPosActual + 1
    Next

    For i = lnPosAnterior To lnPosActual - 1
        If xlsHoja.Cells(i, 3) = "" Then
            xlsHoja.Cells(i, 3).Interior.Color = RGB(192, 192, 192)
        End If
        If xlsHoja.Cells(i, 6) = "" Then
            xlsHoja.Cells(i, 6).Interior.Color = RGB(192, 192, 192)
        End If
        xlsHoja.Cells(i, 7).Formula = "= D" & i & "+ E" & i
    Next
    
    xlsHoja.Range("B" & lnPosActual, "C" & lnPosActual).MergeCells = True
    xlsHoja.Range("B" & lnPosActual).value = "TOTALES"
    If lnPosAnterior <> lnPosActual Then
        xlsHoja.Range("D" & lnPosActual).Formula = "=SUM(D" & lnPosAnterior & ":D" & lnPosAnterior & ")"
        xlsHoja.Range("E" & lnPosActual).Formula = "=SUM(E" & lnPosAnterior & ":E" & (lnPosActual - 1) & ")"
        xlsHoja.Range("G" & lnPosActual).Formula = "=SUM(G" & lnPosAnterior & ":G" & (lnPosActual - 1) & ")"
    Else
        xlsHoja.Range("D" & lnPosActual).value = 0
        xlsHoja.Range("E" & lnPosActual).value = 0
        xlsHoja.Range("G" & lnPosActual).value = 0
    End If
    xlsHoja.Range("D" & lnPosActual, "G" & lnPosActual).NumberFormat = "#,##0.00"
    
    lsFormulaTotal = lsFormulaTotal & "+G" & lnPosActual
    
    xlsHoja.Range("B" & lnPosAnterior, "I" & lnPosActual).Borders.LineStyle = xlContinuous
    xlsHoja.Range("B" & lnPosAnterior, "I" & lnPosActual).Borders.Weight = xlThin
    xlsHoja.Range("B" & lnPosAnterior, "I" & lnPosActual).Borders.ColorIndex = xlAutomatic
    
    xlsHoja.Range("B" & lnPosAnterior, "B" & lnPosActual).Borders(xlEdgeLeft).Weight = xlMedium
    xlsHoja.Range("B" & lnPosActual, "I" & lnPosActual).Borders.Weight = xlMedium
    
    'C1. INVERSIONES DE CAJA MAYNAS EN CRACs
    lnPosActual = lnPosActual + 2
    'xlsHoja.Cells(lnPosActual, 2) = "C1. INVERSIONES DE CAJA MAYNAS EN CRACs"
    xlsHoja.Cells(lnPosActual, 2) = "C1. INVERSIONES DE CAJA MAYNAS EN OTRAS IFIs" 'EJVG20130927
    xlsHoja.Cells(lnPosActual, 2).Font.Color = RGB(0, 0, 255)
    xlsHoja.Cells(lnPosActual, 2).Font.Size = 10
    xlsHoja.Cells(lnPosActual, 2).Font.Bold = True
    
    lnPosActual = lnPosActual + 1
    lnPosAnterior = lnPosActual
    xlsHoja.Cells(lnPosActual, 2) = "BANCO"
    xlsHoja.Cells(lnPosActual, 3) = "DEPOSITO"
    xlsHoja.Cells(lnPosActual, 4) = "TEA"
    xlsHoja.Cells(lnPosActual, 5) = IIf(pnTpoMoneda = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'marg ers044-2016
    xlsHoja.Cells(lnPosActual, 6) = "APERTURA"
    xlsHoja.Cells(lnPosActual, 7) = "VENCIMIENTO"
    xlsHoja.Cells(lnPosActual, 8) = "OBSERVACION"
    
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).Interior.Color = RGB(255, 204, 153)
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 2), xlsHoja.Cells(lnPosActual, 8)).Font.Bold = True
    
    'Set rsDetalleCtas = oCtaIf.obtenerDPFyOvernyInversion(ldFechaRep, pnTpoMoneda, "04", "")
    Set rsDetalleCtas = oCtaIf.obtenerDPFyOvernyInversion(ldFechaRep, pnTpoMoneda, fsOtrasIFisRptConcentraFondos, "") 'EJVG20130927
    For i = 0 To rsDetalleCtas.RecordCount - 1
        lnPosActual = lnPosActual + 1
        xlsHoja.Cells(lnPosActual, 2) = rsDetalleCtas!cPersNombre
        xlsHoja.Cells(lnPosActual, 3) = rsDetalleCtas!cDeposito
        xlsHoja.Cells(lnPosActual, 4).NumberFormat = "0.00%"
        xlsHoja.Cells(lnPosActual, 4) = (rsDetalleCtas!TEA / 100)
        xlsHoja.Cells(lnPosActual, 5).NumberFormat = "#,##0.00"
        xlsHoja.Cells(lnPosActual, 5) = Format(rsDetalleCtas!nSaldo, gsFormatoNumeroView)
        xlsHoja.Range(xlsHoja.Cells(lnPosActual, 6), xlsHoja.Cells(lnPosActual, 7)).NumberFormat = "dd/mm/yyyy" 'EJVG20130927
        xlsHoja.Cells(lnPosActual, 6) = rsDetalleCtas!dCtaIFAper 'Format(rsDetalleCtas!dCtaIFAper, gsFormatoFechaView)
        xlsHoja.Cells(lnPosActual, 7) = rsDetalleCtas!dCtaIFVenc 'Format(rsDetalleCtas!dCtaIFVenc, gsFormatoFechaView)
        rsDetalleCtas.MoveNext
    Next
    xlsHoja.Range("B" & lnPosAnterior, "H" & lnPosActual).Borders.Weight = xlThin
    '**********************
    'SALDOS CONSOLIDADOS
    lnPosActual = lnPosActual + 3
    'xlsHoja.Cells(lnPosActual, 2) = "SALDOS CONSOLIDADO EN BANCOS, CMACS y CRACS: "
    xlsHoja.Cells(lnPosActual, 2) = "SALDOS CONSOLIDADO EN BANCOS, CMACS y OTRAS IFIs: " 'EJVG20130927
    xlsHoja.Range("B" & lnPosActual, "C" & lnPosActual).MergeCells = True
    xlsHoja.Cells(lnPosActual, 4) = "TOTAL MN:"
    xlsHoja.Cells(lnPosActual, 5).Formula = lsFormulaTotal
    xlsHoja.Cells(lnPosActual, 5).Interior.Color = IIf(pnTpoMoneda = gMonedaNacional, RGB(255, 255, 0), RGB(0, 255, 0))
    xlsHoja.Range("B" & lnPosActual, "E" & lnPosActual).Borders.Weight = xlMedium
    
    'EJVG20131230 ***
    If pnTpoMoneda = gMonedaNacional Then
        lnPosFilaMNTotal = lnPosActual
    Else
        lnPosFilaMETotal = lnPosActual
    End If
    'xlsHoja.Cells(lnPosActual, 7) = "EFECTIVO DEL DIA"
    'xlsHoja.Cells(lnPosActual, 8) = SaldoCajasObligExoneradas(pdFecha, pnTpoMoneda)'Comentado by NAGL 20181001
    
    '*********************Agregado by NAGL 20181001 TIC1807210002******************************************
    Set oCtaIf = Nothing
    xlsHoja.Cells(lnPosActual, 7) = "BOVEDA EFECTIVO"
    xlsHoja.Cells(lnPosActual, 8) = oCtaIf.ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, pnTpoMoneda)
    
    xlsHoja.Cells(lnPosActual + 1, 7) = "DINERO EN TRÁNSITO"
    xlsHoja.Cells(lnPosActual + 1, 8) = DAnexRiesg.ObtieneSaldoEfectTransitoTotal(pdFecha, CStr(pnTpoMoneda))
    Set rsEncDiario = oEnc.ObtenerParamEncajeDiarioxCod("04", CStr(pnTpoMoneda), Format(pdFecha, "yyyymmdd")) 'NAGL Agregó pdFecha 20181015
    Do While Not rsEncDiario.EOF
        xlsHoja.Cells(lnPosActual + 2, 7) = "CAJA CHICA"
        xlsHoja.Cells(lnPosActual + 2, 8) = Format(rsEncDiario!nValor, "#,###0.00")
        rsEncDiario.MoveNext
    Loop
    xlsHoja.Cells(lnPosActual + 3, 7) = "EFECTIVO DEL DIA"
    xlsHoja.Cells(lnPosActual + 3, 8).Formula = "=" & "Sum" & "(" & xlsHoja.Range(xlsHoja.Cells(lnPosActual, 8), xlsHoja.Cells(lnPosActual + 2, 8)).Address(False, False) & ")"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 8), xlsHoja.Cells(lnPosActual + 3, 8)).NumberFormat = "#,##0.00"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 7), xlsHoja.Cells(lnPosActual + 3, 8)).Borders.Weight = xlMedium
    '*************************************END NAGL*********************************************************
    
    lnPosActual = lnPosActual + 4 '1 NAGL 20181001
    xlsHoja.Cells(lnPosActual, 7) = "EFECTIVO DISPONIBLE"
    xlsHoja.Cells(lnPosActual, 8) = SaldoEfectivoDisponible(pnTpoMoneda)
    xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 8), xlsHoja.Cells(lnPosActual, 8)).NumberFormat = "#,##0.00"
    xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 7), xlsHoja.Cells(lnPosActual, 8)).Borders.Weight = xlMedium
    xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 7), xlsHoja.Cells(lnPosActual, 8)).Interior.Color = RGB(153, 153, 255) 'NAGL20181001
    'END EJVG *******
    xlsHoja.Range(xlsHoja.Cells(lnPosActual - 1, 7), xlsHoja.Cells(lnPosActual, 8)).Font.Bold = True 'NAGL 20181001
    xlsHoja.Range(xlsHoja.Cells(lnPosActual, 7), xlsHoja.Cells(lnPosActual, 7)).EntireColumn.AutoFit 'NAGL 20181001
    
    Set oCtaIf = Nothing
    Set rsDetalleCtas = Nothing
End Sub
'END EJVG *******
'ALPA 20130708********************************
Public Sub ReporteCartaFianza()
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
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
    Dim dFechaCP As Date
    Dim lsCelda As String
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteCartaFianza"
    'Primera Hoja ******************************************************
    lsNomHoja = "CFianza"
    '*******************************************************************
    lsArchivo1 = "\spooler\" & lsArchivo & "_" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    xlHoja1.Cells(3, 1) = "COMISION AL " & Format(txtFecha.Text, "dd/mm/yyyy")
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 19)).Font.Color = RGB(255, 255, 255)
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 19)).Interior.Color = RGB(166, 166, 166)
    
    
    nSaltoContador = 6
    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.ReporteValidacionComisionCF(txtFecha.Text, Mid(gsOpeCod, 3, 1))
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 19)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = rsCreditos!cCtaCod
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cPersNombre
                xlHoja1.Cells(nSaltoContador, 3) = rsCreditos!cAgeDescripcion
                xlHoja1.Cells(nSaltoContador, 4) = rsCreditos!cMoneda
                xlHoja1.Cells(nSaltoContador, 5) = Format(rsCreditos!dFecVig, "YYYY/MM/DD")
                xlHoja1.Cells(nSaltoContador, 6) = Format(rsCreditos!dFechaPagoComision, "YYYY/MM/DD")
                xlHoja1.Cells(nSaltoContador, 7) = Format(rsCreditos!dVencimiento, "YYYY/MM/DD")
                xlHoja1.Cells(nSaltoContador, 8) = Format(rsCreditos!nComision, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 9) = rsCreditos!nPeriodo
                xlHoja1.Cells(nSaltoContador, 10) = Format(rsCreditos!nValorxDia, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 11) = Format(rsCreditos!nDiasTranscurridos, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 12) = Format(rsCreditos!nAcumuladoxComision, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 13) = rsCreditos!nDiasMes
                xlHoja1.Cells(nSaltoContador, 14) = Format(rsCreditos!nComisionActual, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 15) = Format(rsCreditos!nDevengadoTotal, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 16) = Format(((rsCreditos!nPeriodo - rsCreditos!nDiasTranscurridos - rsCreditos!nDiasMes) * rsCreditos!nValorxDia), "###,###,###.#000")
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 16), xlHoja1.Cells(nSaltoContador, 16)).Interior.Color = RGB(255, 192, 0)
                xlHoja1.Cells(nSaltoContador, 17) = Format(rsCreditos!nMontoComisionRestante, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 18) = Format(rsCreditos!nDiasFaltantesAnterior, "###,###,###.#000")
                xlHoja1.Cells(nSaltoContador, 19) = rsCreditos!cCFEstado
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
End Sub
'*********************************************
'ALPA20131204*********************************
Public Sub GeneratxtAnexo15B(ByVal psMoneda As String, ByVal pdFecha As String)

Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bexiste As Boolean
Dim bencontrado As Boolean
Dim fs As New Scripting.FileSystemObject
Dim oRS As ADODB.Recordset
Dim oDBalCont As DbalanceCont
Set oDBalCont = New DbalanceCont
On Error GoTo ErrBegin15B
  
Set oRS = oDBalCont.ObtenerDatos15paraTxt(pdFecha, "B1")

bexiste = False
If Not (oRS.BOF Or oRS.EOF) Then
    bexiste = True
Else
    Exit Sub
End If



    'Anexo 01 del 15B
    '================
    

    
    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".115"
    


    Open psArchivoAGrabar For Output As #1
    
    
    Print #1, "01150400" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012" '0& "000000000000000"
    
    Do While Not oRS.EOF
        Print #1, oRS!cCodigoText & LlenaCerosSUCAVE(Val(oRS!nSaldoMN)) & LlenaCerosSUCAVE(Val(oRS!nSaldoME)) & LlenaCerosSUCAVE(Val(oRS!nSaldoMN2)) & LlenaCerosSUCAVE(Val(oRS!nSaldoME2))
        oRS.MoveNext
    Loop

    Close #1
    
     
    MsgBox "Archivos SUCAVE Anexo 15A:  (Anx 01-I y Anx 02-II)" & Chr(13) & "generados satisfactoriamente en: " & Chr(13) & Chr(13) & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
     
    Exit Sub

ErrBegin15B:
  ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja, True
    
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
   
End Sub
Public Sub GeneratxtAnexo15C(ByVal psMoneda As String, ByVal pdFecha As String)

Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim bexiste As Boolean
Dim bencontrado As Boolean
Dim fs As New Scripting.FileSystemObject
Dim oRS As ADODB.Recordset
Dim oDBalCont As DbalanceCont
Set oDBalCont = New DbalanceCont
Dim nContador As Integer
Dim cMoneda As String
Dim sCadena As String
Dim c15 As Currency
Dim c20 As Currency
Dim c35 As Currency
Dim c45 As Currency
Dim c47 As Currency
Dim c50 As Currency
Dim c55 As Currency
Dim c65 As Currency
Dim c70 As Currency
Dim c80 As Currency
Dim c90 As Currency
Dim c100 As Currency
Dim c113 As Currency
Dim c115 As Currency
Dim c118 As Currency
Dim c120 As Currency
Dim c130 As Currency
Dim c140 As Currency
Dim c150 As Currency
Dim c151 As Currency
Dim c152 As Currency
Dim c160 As Currency
Dim c170 As Currency
Dim c180 As Currency
Dim c190 As Currency
Dim c200 As Currency
Dim c210 As Currency
Dim c220 As Currency
Dim c230 As Currency
Dim c240 As Currency
Dim c250 As Currency
Dim c260 As Currency
Dim c270 As Currency
Dim cDia As String
On Error GoTo ErrBegin15B
  
Set oRS = oDBalCont.ObtenerDatos15paraTxt(pdFecha, "C1")

bexiste = False
If Not (oRS.BOF Or oRS.EOF) Then
    bexiste = True
Else
    Exit Sub
End If



    'Anexo 01 del 15C
    '================
    
    
    psArchivoAGrabar = App.path & "\SPOOLER\01" & Format(pdFecha, "YYMMdd") & ".115"
    


    Open psArchivoAGrabar For Output As #1
    
    nContador = 100
    Print #1, "01150300" & gsCodCMAC & Format(pdFecha, "YYYYMMDD") & "012" '0& "000000000000000"
    
    Do While Not oRS.EOF
        If oRS!cMoneda = "1" Then
           If cMoneda <> oRS!cMoneda Then
                nContador = 100
                Print #1, "100000" & String(15 * 32, "0")
                nContador = nContador + 1
                c15 = 0
                c20 = 0
                c35 = 0
                c45 = 0
                c47 = 0
                c50 = 0
                c55 = 0
                c65 = 0
                c70 = 0
                c80 = 0
                c90 = 0
                c100 = 0
                c113 = 0
                c115 = 0
                c118 = 0
                c120 = 0
                c130 = 0
                c140 = 0
                c150 = 0
                c151 = 0
                c152 = 0
                c160 = 0
                c170 = 0
                c180 = 0
                c190 = 0
                c200 = 0
                c210 = 0
                c220 = 0
                c230 = 0
                c240 = 0
                c250 = 0
                c260 = 0
                c270 = 0
           End If
        Else
           If cMoneda <> oRS!cMoneda Then
                nContador = 200
                sCadena = CStr(nContador) & "000" & LlenaCerosSUCAVESinVacios(Val(c15)) & LlenaCerosSUCAVESinVacios(Val(c20)) & LlenaCerosSUCAVESinVacios(Val(c35)) & LlenaCerosSUCAVESinVacios(Val(c45))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c47)) & LlenaCerosSUCAVESinVacios(Val(c50)) & LlenaCerosSUCAVESinVacios(Val(c55)) & LlenaCerosSUCAVESinVacios(Val(c65))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c70)) & LlenaCerosSUCAVESinVacios(Val(c80)) & LlenaCerosSUCAVESinVacios(Val(c90)) & LlenaCerosSUCAVESinVacios(Val(c100))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c113)) & LlenaCerosSUCAVESinVacios(Val(c115)) & LlenaCerosSUCAVESinVacios(Val(c118)) & LlenaCerosSUCAVESinVacios(Val(c120))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c130)) & LlenaCerosSUCAVESinVacios(Val(c140)) & LlenaCerosSUCAVESinVacios(Val(c150)) & LlenaCerosSUCAVESinVacios(Val(c151))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c152)) & LlenaCerosSUCAVESinVacios(Val(c160)) & LlenaCerosSUCAVESinVacios(Val(c170)) & LlenaCerosSUCAVESinVacios(Val(c180))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c190)) & LlenaCerosSUCAVESinVacios(Val(c200)) & LlenaCerosSUCAVESinVacios(Val(c210)) & LlenaCerosSUCAVESinVacios(Val(c230))
                sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c240)) & LlenaCerosSUCAVESinVacios(Val(c250)) & LlenaCerosSUCAVESinVacios(Val(c260)) & LlenaCerosSUCAVESinVacios(Val(c270))
                Print #1, sCadena
                nContador = 300
                Print #1, "300000" & String(15 * 32, "0")
                nContador = nContador + 1
                c15 = 0
                c20 = 0
                c35 = 0
                c45 = 0
                c47 = 0
                c50 = 0
                c55 = 0
                c65 = 0
                c70 = 0
                c80 = 0
                c90 = 0
                c100 = 0
                c113 = 0
                c115 = 0
                c118 = 0
                c120 = 0
                c130 = 0
                c140 = 0
                c150 = 0
                c151 = 0
                c152 = 0
                c160 = 0
                c170 = 0
                c180 = 0
                c190 = 0
                c200 = 0
                c210 = 0
                c220 = 0
                c230 = 0
                c240 = 0
                c250 = 0
                c260 = 0
                c270 = 0
           End If
        End If
        cDia = Right("00" & Day(oRS!dfecha), 3)
        cMoneda = oRS!cMoneda
        sCadena = CStr(nContador) & cDia & LlenaCerosSUCAVESinVacios(Val(oRS![15])) & LlenaCerosSUCAVESinVacios(Val(oRS![20])) & LlenaCerosSUCAVESinVacios(Val(oRS![35])) & LlenaCerosSUCAVESinVacios(Val(oRS![45]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![47])) & LlenaCerosSUCAVESinVacios(Val(oRS![50])) & LlenaCerosSUCAVESinVacios(Val(oRS![55])) & LlenaCerosSUCAVESinVacios(Val(oRS![65]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![70])) & LlenaCerosSUCAVESinVacios(Val(oRS![80])) & LlenaCerosSUCAVESinVacios(Val(oRS![90])) & LlenaCerosSUCAVESinVacios(Val(oRS![100]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![113])) & LlenaCerosSUCAVESinVacios(Val(oRS![115])) & LlenaCerosSUCAVESinVacios(Val(oRS![118])) & LlenaCerosSUCAVESinVacios(Val(oRS![120]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![130])) & LlenaCerosSUCAVESinVacios(Val(oRS![140])) & LlenaCerosSUCAVESinVacios(Val(oRS![150])) & LlenaCerosSUCAVESinVacios(Val(oRS![151]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![152])) & LlenaCerosSUCAVESinVacios(Val(oRS![160])) & LlenaCerosSUCAVESinVacios(Val(oRS![170])) & LlenaCerosSUCAVESinVacios(Val(oRS![180]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![190])) & LlenaCerosSUCAVESinVacios(Val(oRS![200])) & LlenaCerosSUCAVESinVacios(Val(oRS![210])) & LlenaCerosSUCAVESinVacios(Val(oRS![230]))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(oRS![240])) & LlenaCerosSUCAVESinVacios(Val(oRS![250])) & LlenaCerosSUCAVESinVacios(Val(oRS![260])) & LlenaCerosSUCAVESinVacios(Val(oRS![270]))
        Print #1, sCadena
        
        c15 = c15 + oRS![15]
        c20 = c20 + oRS![20]
        c35 = c35 + oRS![35]
        c45 = c45 + oRS![45]
        c47 = c47 + oRS![47]
        c50 = c50 + oRS![50]
        c55 = c55 + oRS![55]
        c65 = c65 + oRS![65]
        c70 = c70 + oRS![70]
        c80 = c80 + oRS![80]
        c90 = c90 + oRS![90]
        c100 = c100 + oRS![100]
        c113 = c113 + oRS![113]
        c115 = c115 + oRS![115]
        c118 = c118 + oRS![118]
        c120 = c120 + oRS![120]
        c130 = c130 + oRS![130]
        c140 = c140 + oRS![140]
        c150 = c150 + oRS![150]
        c151 = c151 + oRS![151]
        c152 = c152 + oRS![152]
        c160 = c160 + oRS![160]
        c170 = c170 + oRS![170]
        c180 = c180 + oRS![180]
        c190 = c190 + oRS![190]
        c200 = c200 + oRS![200]
        c210 = c210 + oRS![210]
        c220 = c220 + oRS![220]
        c230 = c230 + oRS![230]
        c240 = c240 + oRS![240]
        c250 = c250 + oRS![250]
        c260 = c260 + oRS![260]
        c270 = c270 + oRS![270]
        oRS.MoveNext
    Loop
        nContador = 400
        sCadena = CStr(nContador) & "000" & LlenaCerosSUCAVESinVacios(Val(c15)) & LlenaCerosSUCAVESinVacios(Val(c20)) & LlenaCerosSUCAVESinVacios(Val(c35)) & LlenaCerosSUCAVESinVacios(Val(c45))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c47)) & LlenaCerosSUCAVESinVacios(Val(c50)) & LlenaCerosSUCAVESinVacios(Val(c55)) & LlenaCerosSUCAVESinVacios(Val(c65))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c70)) & LlenaCerosSUCAVESinVacios(Val(c80)) & LlenaCerosSUCAVESinVacios(Val(c90)) & LlenaCerosSUCAVESinVacios(Val(c100))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c113)) & LlenaCerosSUCAVESinVacios(Val(c115)) & LlenaCerosSUCAVESinVacios(Val(c118)) & LlenaCerosSUCAVESinVacios(Val(c120))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c130)) & LlenaCerosSUCAVESinVacios(Val(c140)) & LlenaCerosSUCAVESinVacios(Val(c150)) & LlenaCerosSUCAVESinVacios(Val(c151))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c152)) & LlenaCerosSUCAVESinVacios(Val(c160)) & LlenaCerosSUCAVESinVacios(Val(c170)) & LlenaCerosSUCAVESinVacios(Val(c180))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c190)) & LlenaCerosSUCAVESinVacios(Val(c200)) & LlenaCerosSUCAVESinVacios(Val(c210)) & LlenaCerosSUCAVESinVacios(Val(c230))
        sCadena = sCadena & LlenaCerosSUCAVESinVacios(Val(c240)) & LlenaCerosSUCAVESinVacios(Val(c250)) & LlenaCerosSUCAVESinVacios(Val(c260)) & LlenaCerosSUCAVESinVacios(Val(c270))
        Print #1, sCadena
    Close #1
    
     
    MsgBox "Archivos SUCAVE Anexo 15A:  (Anx 01-I y Anx 02-II)" & Chr(13) & "generados satisfactoriamente en: " & Chr(13) & Chr(13) & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
     
    Exit Sub

ErrBegin15B:
  ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja, True
    
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
   
End Sub
'*********************************************
'Public Sub ReporteAnexo16A(ByVal pnTipoCambio As Currency, ByVal pdFecha As Date)
Public Sub ReporteAnexo16A(ByVal pdFecha As Date) 'NAGL 20170425
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
    Dim TituloProgress As String 'NAGL20170407
    Dim MensajeProgress As String 'NAGL20170407
    Dim oBarra As clsProgressBar 'NAGL20170407
    Dim nprogress As Integer 'NAGL20170407
    
    'Dim rsRep15B As ADODB.Recordset
    'Dim oRep15B As New DbalanceCont
    Dim oDbalanceCont As DbalanceCont
    Dim oDAnx As DAnexoRiesgos  'NAGL 20190518
    Dim nSaldoDiario1 As Currency
    Dim nSaldoDiario2 As Currency
    Dim pdFechaFinDeMes As Date
    Dim pdFechaFinDeMesMA As Date
    Dim dFechaAnte As Date
    Dim ldFechaPro As Date
    'Dim pdFecha As Date
    Dim nDia As Integer
    Dim oCambio As nTipoCambio
    Dim lnTipoCambioFC As Currency
    Dim lnTipoCambioProceso As Currency
    Dim nTipoCambioAn As Currency
    Dim loRs As ADODB.Recordset
    Dim oEst As New NEstadisticas 'NAGL
    Dim oPatrimonio As New DAnexoRiesgos 'NAGL 20170425
    
    Dim nTotalObligSugEncajMN As Currency
    'Dim nTotalTasaBaseEncajMN As Currency
    Dim nTotalObligSugEncajME As Currency
    'Dim nTotalTasaBaseEncajME  As Currency
    Dim lnTotalObligacionesAlDiaMN As Currency
    Dim lnTotalObligacionesAlDiaME  As Currency
    Dim lnTotalObligacionesAlDiaMN_MA As Currency
    Dim lnTotalObligacionesAlDiaME_MA  As Currency
    Dim lnTotalObligacionesAlDiaMN_DA As Currency
    Dim lnTotalObligacionesAlDiaME_DA As Currency
    Dim nTotalTasaBaseEncajMN_DA  As Currency
    Dim nTotalObligSugEncajMN_DA As Currency
    Dim nTotalTasaBaseEncajME_DA As Currency
    Dim nTotalObligSugEncajME_DA As Currency
    Dim oRs1 As ADODB.Recordset
    Dim ix As Integer
    Dim lnTotalObligacionesAlDiaPromedioMN As Currency
    Dim lnTotalObligacionesAlDiaPromedioME  As Currency
    Dim lnTotalObligacionesBase As Currency 'NAGL
    Dim lnTotalObligacionesxDiasDeclarar As Currency 'NAGL
    Dim lnTotalObligacionesAlDiaPromedioMN_MA As Currency
    Dim lnTotalObligacionesAlDiaPromedioME_MA  As Currency
    Dim nPromedTMMN As Double
    Dim nPromedTMME As Double
    Dim nTasaExigibleME As Currency 'NAGL
    Dim nTasaBaseMaginalMN As Currency
    Dim nTasaBaseMaginalME As Currency
    Dim nExigibleMaginalMN As Currency
    Dim nExigibleMaginalME As Currency
    Dim lnEncajeExigibleRGMN As Currency
    Dim lnEncajeExigibleRGME As Currency
    
    Dim lnOtrosDepMND1y2_C0 As Currency
    Dim lnOtrosDepMND4aM_C0 As Currency
    Dim lnOtrosDepMND1y2_C1 As Currency
    Dim lnOtrosDepMED1y2_C0 As Currency
    Dim lnOtrosDepMED4aM_C0 As Currency
    Dim lnOtrosDepMED1y2_C1 As Currency
    Dim lnSubastasMND1y2_C1 As Currency
    Dim lnSubastasMND1y2_C0 As Currency
    Dim lnSubastasMND4aM_C0 As Currency
    Dim lnSubastasMED1y2_C1 As Currency
    Dim lnSubastasMED1y2_C0 As Currency
    Dim lnSubastasMED4aM_C0 As Currency
    
    Dim lnSubastasMED3o3_C0 As Currency
    Dim lnSubastasMND3o3_C0 As Currency
    Dim lnSubastasMED3o3_C1 As Currency
    Dim lnSubastasMND3o3_C1 As Currency
    
    Dim lnOtrosDepMND3o3_C0 As Currency
    Dim lnOtrosDepMED3o3_C0 As Currency
    Dim lnOtrosDepMND3o3_C1 As Currency
    Dim lnOtrosDepMED3o3_C1 As Currency
    
    Dim lnSubastasMND4aM_C1 As Currency
    Dim lnSubastasMED4aM_C1 As Currency
    
    Dim nTotalAcredores20 As Currency
    Dim nTotalAcredores10 As Currency
    Dim nTotalAcredoresTo As Currency
    
    Dim nTotalDepositantes20 As Currency
    Dim nTotalDepositantes10 As Currency
    Dim nTotalDepositantesTo As Currency
    Dim nTotalAcredores201, nTotalAcredores101, nTotalAcredores202, nTotalAcredores102 As Currency
    Dim nTotalDepositantes201, nTotalDepositantes101, nTotalDepositantes202, nTotalDepositantes102 As Currency
    
On Error GoTo GeneraExcelErr
    'NAGL
    Set oBarra = New clsProgressBar
    oBarra.ShowForm frmReportes
    oBarra.Max = 10
    nprogress = 0
    oBarra.Progress nprogress, "ANEXO 16A: Cuadro de Liquidez por plazos de Vencimiento", "GENERANDO EL ARCHIVO", "", vbBlue
    TituloProgress = "ANEXO 16A: Cuadro de Liquidez por plazos de Vencimiento"
    MensajeProgress = "GENERANDO EL ARCHIVO"
    'NAGL
    
    'pdFecha = Format(txtfecha.Text, "YYYY/MM/DD")
    pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, pdFecha)), DateAdd("m", 1, pdFecha))
    pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    nDia = Day(pdFecha)
    Set oDbalanceCont = New DbalanceCont
    Set oDAnx = New DAnexoRiesgos
    If nDia >= 15 Then
        dFechaAnte = DateAdd("d", -(nDia - 1), pdFecha)
    Else
        dFechaAnte = DateAdd("d", -(nDia - 1), DateAdd("m", -1, pdFecha))
    End If
    Set oCambio = New nTipoCambio
    
    If Month(pdFecha) = Month(DateAdd("d", 1, pdFecha)) Then
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(pdFecha, TCFijoDia), "#,##0.0000")
    Else
        lnTipoCambioFC = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, pdFecha), TCFijoDia), "#,##0.0000")
    End If
    nTipoCambioAn = lnTipoCambioFC
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Anexo16_SBS"
    'Primera Hoja ******************************************************
    lsNomHoja = "Anx16AMN"
    '*******************************************************************
    lsArchivo1 = "\spooler\ANEXO_16A_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    xlHoja1.Cells(4, 2) = "AL " & Format(pdFecha, "DD") & " DE " & UCase(Format(pdFecha, "MMMM")) & " DEL  " & Format(pdFecha, "YYYY")
    oBarra.Progress 1, TituloProgress, MensajeProgress, "", vbBlue
    '****************************OBLIGACIONES SUJETAS A ENCAJE*******************************


        lnTotalObligacionesAlDiaMN = 0 '
        lnTotalObligacionesAlDiaME = 0 '
        nTotalObligSugEncajMN = 0
        nTotalObligSugEncajME = 0
        lnTotalObligacionesAlDiaMN_DA = 0 '
        lnTotalObligacionesAlDiaME_DA = 0 '
        nTotalObligSugEncajMN_DA = 0
        nTotalObligSugEncajME_DA = 0

        ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
        'ldFechaPro = DateAdd("d", -Day(ldFechaPro), ldFechaPro)

       For ix = 1 To Day(pdFecha)
            ldFechaPro = DateAdd("d", 1, ldFechaPro)
            If Month(ldFechaPro) = Month(DateAdd("d", 1, ldFechaPro)) Then
                lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(ldFechaPro, TCFijoDia), "#,##0.0000")
            Else
                lnTipoCambioProceso = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, ldFechaPro), TCFijoDia), "#,##0.0000")
            End If 'NAGL ERS 079-2016 20170407
            
            '************************************************************************************************************************
            'nTotalObligSugEncajMN_DA = oDbalanceCont.SaldoCtas(1, "761201", ldFechaPro, DateAdd("d", -Day(pdFecha), pdFecha), lnTipoCambioProceso, lnTipoCambioProceso)
            'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoCtas(4, "761201", ldFechaPro, DateAdd("d", -Day(pdFecha), pdFecha), lnTipoCambioProceso, lnTipoCambioProceso)
            'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoCtas(7, "761201", ldFechaPro, DateAdd("d", -Day(pdFecha), pdFecha), lnTipoCambioProceso, lnTipoCambioProceso)
            'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - oEst.GetEstadSaldoCheques(ldFechaPro, "Plazo", "1", gbBitCentral)
            'nTotalObligSugEncajMN = nTotalObligSugEncajMN + nTotalObligSugEncajMN_DA
            
            'SOLES
            nTotalObligSugEncajMN_DA = oDbalanceCont.SaldoCtas(1, "761201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
            nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "232") 'Ahorros
            nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234")
            nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1, "233") 'Depositos a plazo fijo
            'nTotalObligSugEncajMN_DA = nTotalObligSugEncajMN_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 1, "234"))
            lnTotalObligacionesAlDiaMN = lnTotalObligacionesAlDiaMN + nTotalObligSugEncajMN_DA '*************NAGL ERS079-2016 20170407
   
            'DOLARES
            'nTotalObligSugEncajME_DA = oDbalanceCont.SaldoCtas(1, "762201", ldFechaPro, DateAdd("d", -Day(pdFecha), pdFecha), lnTipoCambioProceso, lnTipoCambioProceso)
            'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoCtas(4, "762201", ldFechaPro, DateAdd("d", -Day(pdFecha), pdFecha), lnTipoCambioProceso, lnTipoCambioProceso)
            'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoCtas(7, "762201", ldFechaPro, DateAdd("d", -Day(pdFecha), pdFecha), lnTipoCambioProceso, lnTipoCambioProceso)
            'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - oEst.GetEstadSaldoCheques(ldFechaPro, "Plazo", "2", gbBitCentral)
            'nTotalObligSugEncajME = nTotalObligSugEncajME + nTotalObligSugEncajME_DA
        
            nTotalObligSugEncajME_DA = oDbalanceCont.SaldoCtas(1, "762201", ldFechaPro, pdFechaFinDeMes, lnTipoCambioProceso, lnTipoCambioProceso) 'Obligaciones Inmediatas
            nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "232") 'Ahorros
            nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234")
            nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - oDbalanceCont.SaldoCajasCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") - oDbalanceCont.SaldoCracsAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2, "233") 'Depositos a plazo fijo
            'nTotalObligSugEncajME_DA = nTotalObligSugEncajME_DA - (oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "232") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "233") + oDbalanceCont.SaldoChequeAhoPlaFijCTS(Format(ldFechaPro, "yyyymmdd"), 2, "234"))
            lnTotalObligacionesAlDiaME = lnTotalObligacionesAlDiaME + nTotalObligSugEncajME_DA '*************NAGL ERS079-2016 20170407
        
        Next ix
        
        'lnTotalObligacionesAlDiaPromedioMN_MA = (lnTotalObligacionesAlDiaMN / Day(DateAdd("d", -Day(pdFecha), pdFecha)))
        'lnTotalObligacionesAlDiaPromedioME_MA = (lnTotalObligacionesAlDiaME / Day(DateAdd("d", -Day(pdFecha), pdFecha)))
        'lnTotalObligacionesAlDiaMN_MA = lnTotalObligacionesAlDiaMN + (lnTotalObligacionesAlDiaMN / Day(DateAdd("d", -Day(pdFecha), pdFecha)))
        'lnTotalObligacionesAlDiaME_MA = lnTotalObligacionesAlDiaME + (lnTotalObligacionesAlDiaME / Day(DateAdd("d", -Day(pdFecha), pdFecha)))
        
        'lnTotalObligacionesAlDiaMN_MA = lnTotalObligacionesAlDiaMN
        'lnTotalObligacionesAlDiaME_MA = lnTotalObligacionesAlDiaME
        'lnTotalObligacionesAlDiaMN = 0 '
        'lnTotalObligacionesAlDiaME = 0 '
        'nTotalTasaBaseEncajMN_DA = 0 '
        'nTotalTasaBaseEncajME_DA = 0 '
        
    'Fin TOSE
    
    'lnTotalObligacionesAlDiaMN = oDbalanceCont.ObtenerParamEncDiarioxCodigo("07")
    'lnTotalObligacionesAlDiaME = oDbalanceCont.ObtenerParamEncDiarioxCodigo("06")
    'lnTotalObligacionesAlDiaPromedioMN = (lnTotalObligacionesAlDiaMN / oDbalanceCont.ObtenerParamEncDiarioxCodigo("10"))
    'lnTotalObligacionesAlDiaPromedioME = (lnTotalObligacionesAlDiaME / oDbalanceCont.ObtenerParamEncDiarioxCodigo("10"))
    
    'lnTotalObligacionesAlDiaMN = lnTotalObligacionesAlDiaPromedioMN * Day(ldFechaPro)
    'lnTotalObligacionesAlDiaME = lnTotalObligacionesAlDiaPromedioME * Day(ldFechaPro)

    'nTasaBaseMaginalMN = oDbalanceCont.ObtenerParamEncDiarioxCodigo("27")
    'nTasaBaseMaginalME = oDbalanceCont.ObtenerParamEncDiarioxCodigo("28")

    'nPromedTMMN = lnTotalObligacionesAlDiaMN * CCur(nTasaBaseMaginalMN)
    'nPromedTMME = lnTotalObligacionesAlDiaME * Round(CDbl((oDbalanceCont.ObtenerParamEncDiarioxCodigo("08") / oDbalanceCont.ObtenerParamEncDiarioxCodigo("06"))), 6)

    'nExigibleMaginalMN = (lnTotalObligacionesAlDiaMN_MA - lnTotalObligacionesAlDiaMN) * CCur(nTasaBaseMaginalMN)
    'nExigibleMaginalME = nPromedTMME + IIf(lnTotalObligacionesAlDiaME_MA, (lnTotalObligacionesAlDiaME_MA - oDbalanceCont.ObtenerParamEncDiarioxCodigo("06")) * oDbalanceCont.ObtenerParamEncDiarioxCodigo("03") / 100, 0)

    
    'lnEncajeExigibleRGMN = nExigibleMaginalMN * CCur(nTasaBaseMaginalMN) + lnTotalObligacionesAlDiaMN 'CCur((oDbalanceCont.ObtenerParamEncDiarioxCodigo("09") / oDbalanceCont.ObtenerParamEncDiarioxCodigo("07")))
    'lnEncajeExigibleRGME = nExigibleMaginalME * CCur(nTasaBaseMaginalME) + lnTotalObligacionesAlDiaME 'CCur((oDbalanceCont.ObtenerParamEncDiarioxCodigo("08") / oDbalanceCont.ObtenerParamEncDiarioxCodigo("06")))
    'lnEncajeExigibleRGMN = Round(lnTotalObligacionesAlDiaMN_MA, 2)
    'If (nExigibleMaginalME / lnTotalObligacionesAlDiaME_MA) * 100 > 40 Then
    'lnEncajeExigibleRGME = Round(lnTotalObligacionesAlDiaME_MA * oDbalanceCont.ObtenerParamEncDiarioxCodigo("31") / 100, 2)
    'Else
    'lnEncajeExigibleRGME = nExigibleMaginalME
    'End If Comentado por NAGL
    '******************************************************************************
    
    lnTotalObligacionesBase = oDbalanceCont.ObtenerParamEncDiarioxCodigo("06") 'Tose Base Mes de Referencia ME
    nExigibleMaginalME = oDbalanceCont.ObtenerParamEncDiarioxCodigo("08")
    lnTotalObligacionesAlDiaPromedioME = Round(lnTotalObligacionesBase / oDbalanceCont.ObtenerParamEncDiarioxCodigo("10"), 2)
    nTasaExigibleME = Round(nExigibleMaginalME / oDbalanceCont.ObtenerParamEncDiarioxCodigo("10"), 2)
    lnTotalObligacionesxDiasDeclarar = lnTotalObligacionesAlDiaPromedioME * Day(pdFecha)
    nPromedTMME = Round((nTasaExigibleME / lnTotalObligacionesAlDiaPromedioME), 6)
    
    nTasaBaseMaginalME = (oDbalanceCont.ObtenerParamEncDiarioxCodigo("03") / 100)
    
    If lnTotalObligacionesAlDiaME > lnTotalObligacionesxDiasDeclarar Then
    lnEncajeExigibleRGME = (lnTotalObligacionesxDiasDeclarar * nPromedTMME) + (lnTotalObligacionesAlDiaME - lnTotalObligacionesxDiasDeclarar) * nTasaBaseMaginalME
    Else
    lnEncajeExigibleRGME = Round(lnTotalObligacionesAlDiaME * nPromedTMME, 2)
    End If
    
    lnEncajeExigibleRGMN = lnTotalObligacionesAlDiaMN * (oDbalanceCont.ObtenerParamEncDiarioxCodigo("32") / 100)
    '**********************************NAGL ERS 079-2016 20170407
    oBarra.Progress 2, TituloProgress, MensajeProgress, "", vbBlue

    '****************************OBLIGACIONES X CTA ***********************
     cargarPasivosObligxCtaAhorrosANX6 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC
     cargarPasivosObligxCtaAhorrosANX6 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC
     cargarPasivosObligxCtaCTSANX6 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC
     cargarPasivosObligxCtaCTSANX6 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC
     cargarDepositosInmovilizadosANX16 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarDepositosInmovilizadosANX16 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarObligacionesVistaANX16 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarObligacionesVistaANX16 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
     cargarPlazoFijoRangosPersoneriaRangoAnexo6 xlHoja1.Application, pdFecha, "1", Round(lnEncajeExigibleRGMN, 2), Round(lnEncajeExigibleRGME, 2), lnTipoCambioFC
     cargarFondeoAhPFCTSxProducto xlHoja1.Application, pdFecha, lnTipoCambioFC '***NAGL ERS006-2019 20190518
    '******************************************************************************
    oBarra.Progress 3, TituloProgress, MensajeProgress, "", vbBlue
    
    'Patrimonio Efectivo  NAGL 20170425
    xlHoja1.Cells(3, 18) = Format(lnTipoCambioFC, "#,##0.0000")
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oPatrimonio.CargaPatrimonioEfectivo(pdFecha)
    xlHoja1.Cells(1, 18) = Format(oRs1!dFechaPatrimonio, "dd/mm/yyyy")
    xlHoja1.Cells(2, 18) = Format(oRs1!nPatrimonioMN, "#,##0.00") 'NAGL 20170621
    xlHoja1.Cells(4, 18) = Format(oRs1!nPatrimonioME, "#,##0.00") 'NAGL 20170621
    
    Set oRs1 = New ADODB.Recordset
    'Disponible
    Set oRs1 = oDbalanceCont.ObtenerOverNightTramosResidual("1", pdFecha, "2")
    
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(9, 2 + oRs1!cRango) = oRs1!nSaldo
    oRs1.MoveNext
    Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(9, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '*******************************************************************************************
    
    'xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(9, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '*******************************************************************************************
    oBarra.Progress 4, TituloProgress, MensajeProgress, "", vbBlue
    
    Set oRs1 = New ADODB.Recordset
    'Inversiones a Vencimiento
    Set oRs1 = oDbalanceCont.ObtenerInversionesAVencimientoResidual(pdFecha, 1)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(12, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    
    'xlHoja1.Cells(11, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131401", pdFecha, "1", 0) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131402", pdFecha, "1", 0)
    
    'Inversiones Disponibles para la venta
    Set oRs1 = oDbalanceCont.ObtenerInversionesVentaTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(11, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    xlHoja1.Cells(11, 3) = CCur(xlHoja1.Cells(11, 3)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131402", pdFecha, "1", 0) 'CDBCRP
    xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(11, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '*************************NAGL ERS 079-2016 20170407
    
    Set oRs1 = New ADODB.Recordset
    'Cuentas por Cobrar - Operaciones de Reporte
    Set oRs1 = oDbalanceCont.ObtenerCuentasxCobrarTramosResidual("1", pdFecha)
    oBarra.Progress 5, TituloProgress, MensajeProgress, "", vbBlue
    
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(45, 2 + oRs1!cRango) = oRs1!nSaldo
    oRs1.MoveNext
    Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(45, 3), xlHoja1.Cells(45, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '**********NAGL ERS 079-2016 20170407

    'Fondeo 1 PF
    'Call PintaFondeo(xlHoja1, pdFecha, lnTipoCambioFC, "1") 'Comentado by NAGL 20190514 Considerado en el Método cargarFondeoAhPFCTSxProducto ERS006-2019
    'Fin fondeo Soles
    
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerRestringidosxTramosxProducto("1", pdFecha, 0, "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(52, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    oBarra.Progress 6, TituloProgress, MensajeProgress, "", vbBlue
    
    Set oRs1 = oDbalanceCont.ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto(pdFecha, "1", "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(54, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "1", "1")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(55, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    xlHoja1.Cells(55, 3) = CCur(xlHoja1.Cells(55, 3)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("241802", pdFecha, "1", 0) 'Activado by NAGL 20190626 Primera Banda

    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "1", "0")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(56, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    xlHoja1.Cells(56, 3) = CCur(xlHoja1.Cells(56, 3)) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("241807", pdFecha, "1", 0) 'Activado by NAGL 20190626 Primera Banda
    
    'Inversiones a valor Razonable con cambios en resultados - supuesto
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1311", pdFecha, "1", 0)
    xlHoja1.Cells(65, 3) = nSaldoDiario1 '****NAGL ERS 079-2016 20170407
    
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2312", pdFecha, "1", 0)
    xlHoja1.Cells(87, 3) = nSaldoDiario1  'Depósitos Ifis/OFIs según Supuesto
    
    oBarra.Progress 7, TituloProgress, MensajeProgress, "", vbBlue
    
    cargarDatosBalanceANX16 xlHoja1.Application, 1, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407

    '***LOS SALDOS DE CREDITOS, INTERESES DEVENGADOS E INTERESES FUTUROS SE ENCUENTRAN EN ESTA SECCIÓN TANTO EN MN COMO EN ME Según Anx03_ERS006-2019
    lsNomHoja = "AnxRendInt"
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
    
    '******************BEGIN TRASLADO BY NAGL 20191030 Según Anx03_ERS006-2019*******************'
    'SALDOS DE CRÉDITOS VIGENTES
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 5 Then 'EmpreSist
                xlHoja1.Cells(12, 3) = oRs1![1M]
                xlHoja1.Cells(12, 5) = oRs1![2M]
                xlHoja1.Cells(12, 7) = oRs1![3M]
                xlHoja1.Cells(12, 9) = oRs1![4M]
                xlHoja1.Cells(12, 11) = oRs1![5M]
                xlHoja1.Cells(12, 13) = oRs1![6M]
                xlHoja1.Cells(12, 15) = oRs1![7-9M]
                xlHoja1.Cells(12, 17) = oRs1![10-12M]
                xlHoja1.Cells(12, 19) = oRs1![1-2A]
                xlHoja1.Cells(12, 21) = oRs1![2-5A]
                xlHoja1.Cells(12, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 6 Then 'Corporativos
                xlHoja1.Cells(13, 3) = oRs1![1M]
                xlHoja1.Cells(13, 5) = oRs1![2M]
                xlHoja1.Cells(13, 7) = oRs1![3M]
                xlHoja1.Cells(13, 9) = oRs1![4M]
                xlHoja1.Cells(13, 11) = oRs1![5M]
                xlHoja1.Cells(13, 13) = oRs1![6M]
                xlHoja1.Cells(13, 15) = oRs1![7-9M]
                xlHoja1.Cells(13, 17) = oRs1![10-12M]
                xlHoja1.Cells(13, 19) = oRs1![1-2A]
                xlHoja1.Cells(13, 21) = oRs1![2-5A]
                xlHoja1.Cells(13, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 2 Then 'Grandes Empr.
                xlHoja1.Cells(14, 3) = oRs1![1M]
                xlHoja1.Cells(14, 5) = oRs1![2M]
                xlHoja1.Cells(14, 7) = oRs1![3M]
                xlHoja1.Cells(14, 9) = oRs1![4M]
                xlHoja1.Cells(14, 11) = oRs1![5M]
                xlHoja1.Cells(14, 13) = oRs1![6M]
                xlHoja1.Cells(14, 15) = oRs1![7-9M]
                xlHoja1.Cells(14, 17) = oRs1![10-12M]
                xlHoja1.Cells(14, 19) = oRs1![1-2A]
                xlHoja1.Cells(14, 21) = oRs1![2-5A]
                xlHoja1.Cells(14, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 3 Then 'Medianas Empr
                xlHoja1.Cells(15, 3) = oRs1![1M]
                xlHoja1.Cells(15, 5) = oRs1![2M]
                xlHoja1.Cells(15, 7) = oRs1![3M]
                xlHoja1.Cells(15, 9) = oRs1![4M]
                xlHoja1.Cells(15, 11) = oRs1![5M]
                xlHoja1.Cells(15, 13) = oRs1![6M]
                xlHoja1.Cells(15, 15) = oRs1![7-9M]
                xlHoja1.Cells(15, 17) = oRs1![10-12M]
                xlHoja1.Cells(15, 19) = oRs1![1-2A]
                xlHoja1.Cells(15, 21) = oRs1![2-5A]
                xlHoja1.Cells(15, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 4 Then 'Pequ Empr
                xlHoja1.Cells(33, 3) = oRs1![1M]
                xlHoja1.Cells(33, 5) = oRs1![2M]
                xlHoja1.Cells(33, 7) = oRs1![3M]
                xlHoja1.Cells(33, 9) = oRs1![4M]
                xlHoja1.Cells(33, 11) = oRs1![5M]
                xlHoja1.Cells(33, 13) = oRs1![6M]
                xlHoja1.Cells(33, 15) = oRs1![7-9M]
                xlHoja1.Cells(33, 17) = oRs1![10-12M]
                xlHoja1.Cells(33, 19) = oRs1![1-2A]
                xlHoja1.Cells(33, 21) = oRs1![2-5A]
                xlHoja1.Cells(33, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 5 Then 'MicroEmpresas
                xlHoja1.Cells(34, 3) = oRs1![1M]
                xlHoja1.Cells(34, 5) = oRs1![2M]
                xlHoja1.Cells(34, 7) = oRs1![3M]
                xlHoja1.Cells(34, 9) = oRs1![4M]
                xlHoja1.Cells(34, 11) = oRs1![5M]
                xlHoja1.Cells(34, 13) = oRs1![6M]
                xlHoja1.Cells(34, 15) = oRs1![7-9M]
                xlHoja1.Cells(34, 17) = oRs1![10-12M]
                xlHoja1.Cells(34, 19) = oRs1![1-2A]
                xlHoja1.Cells(34, 21) = oRs1![2-5A]
                xlHoja1.Cells(34, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 7 Then 'Consumos
                xlHoja1.Cells(51, 3) = oRs1![1M]
                xlHoja1.Cells(51, 5) = oRs1![2M]
                xlHoja1.Cells(51, 7) = oRs1![3M]
                xlHoja1.Cells(51, 9) = oRs1![4M]
                xlHoja1.Cells(51, 11) = oRs1![5M]
                xlHoja1.Cells(51, 13) = oRs1![6M]
                xlHoja1.Cells(51, 15) = oRs1![7-9M]
                xlHoja1.Cells(51, 17) = oRs1![10-12M]
                xlHoja1.Cells(51, 19) = oRs1![1-2A]
                xlHoja1.Cells(51, 21) = oRs1![2-5A]
                xlHoja1.Cells(51, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 8 Then 'Hip.Vigentes
                xlHoja1.Cells(45, 3) = oRs1![1M]
                xlHoja1.Cells(45, 5) = oRs1![2M]
                xlHoja1.Cells(45, 7) = oRs1![3M]
                xlHoja1.Cells(45, 9) = oRs1![4M]
                xlHoja1.Cells(45, 11) = oRs1![5M]
                xlHoja1.Cells(45, 13) = oRs1![6M]
                xlHoja1.Cells(45, 15) = oRs1![7-9M]
                xlHoja1.Cells(45, 17) = oRs1![10-12M]
                xlHoja1.Cells(45, 19) = oRs1![1-2A]
                xlHoja1.Cells(45, 21) = oRs1![2-5A]
                xlHoja1.Cells(45, 23) = oRs1![m5A]
        End If
    oRs1.MoveNext
    Loop
    End If
    
    Set oRs1 = New ADODB.Recordset
    
    Set oRs1 = oDbalanceCont.ObtenerCreditosTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 5 Then 'EmpreSist
                xlHoja1.Cells(12, 4) = oRs1![1M]
                xlHoja1.Cells(12, 6) = oRs1![2M]
                xlHoja1.Cells(12, 8) = oRs1![3M]
                xlHoja1.Cells(12, 10) = oRs1![4M]
                xlHoja1.Cells(12, 12) = oRs1![5M]
                xlHoja1.Cells(12, 14) = oRs1![6M]
                xlHoja1.Cells(12, 16) = oRs1![7-9M]
                xlHoja1.Cells(12, 18) = oRs1![10-12M]
                xlHoja1.Cells(12, 20) = oRs1![1-2A]
                xlHoja1.Cells(12, 22) = oRs1![2-5A]
                xlHoja1.Cells(12, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 6 Then 'Corporativos
                xlHoja1.Cells(13, 4) = oRs1![1M]
                xlHoja1.Cells(13, 6) = oRs1![2M]
                xlHoja1.Cells(13, 8) = oRs1![3M]
                xlHoja1.Cells(13, 10) = oRs1![4M]
                xlHoja1.Cells(13, 12) = oRs1![5M]
                xlHoja1.Cells(13, 14) = oRs1![6M]
                xlHoja1.Cells(13, 16) = oRs1![7-9M]
                xlHoja1.Cells(13, 18) = oRs1![10-12M]
                xlHoja1.Cells(13, 20) = oRs1![1-2A]
                xlHoja1.Cells(13, 22) = oRs1![2-5A]
                xlHoja1.Cells(13, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 2 Then 'Grandes Empr.
                xlHoja1.Cells(14, 4) = oRs1![1M]
                xlHoja1.Cells(14, 6) = oRs1![2M]
                xlHoja1.Cells(14, 8) = oRs1![3M]
                xlHoja1.Cells(14, 10) = oRs1![4M]
                xlHoja1.Cells(14, 12) = oRs1![5M]
                xlHoja1.Cells(14, 14) = oRs1![6M]
                xlHoja1.Cells(14, 16) = oRs1![7-9M]
                xlHoja1.Cells(14, 18) = oRs1![10-12M]
                xlHoja1.Cells(14, 20) = oRs1![1-2A]
                xlHoja1.Cells(14, 22) = oRs1![2-5A]
                xlHoja1.Cells(14, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 3 Then 'Medianas Empr
                xlHoja1.Cells(15, 4) = oRs1![1M]
                xlHoja1.Cells(15, 6) = oRs1![2M]
                xlHoja1.Cells(15, 8) = oRs1![3M]
                xlHoja1.Cells(15, 10) = oRs1![4M]
                xlHoja1.Cells(15, 12) = oRs1![5M]
                xlHoja1.Cells(15, 14) = oRs1![6M]
                xlHoja1.Cells(15, 16) = oRs1![7-9M]
                xlHoja1.Cells(15, 18) = oRs1![10-12M]
                xlHoja1.Cells(15, 20) = oRs1![1-2A]
                xlHoja1.Cells(15, 22) = oRs1![2-5A]
                xlHoja1.Cells(15, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 4 Then 'Pequ Empr
                xlHoja1.Cells(33, 4) = oRs1![1M]
                xlHoja1.Cells(33, 6) = oRs1![2M]
                xlHoja1.Cells(33, 8) = oRs1![3M]
                xlHoja1.Cells(33, 10) = oRs1![4M]
                xlHoja1.Cells(33, 12) = oRs1![5M]
                xlHoja1.Cells(33, 14) = oRs1![6M]
                xlHoja1.Cells(33, 16) = oRs1![7-9M]
                xlHoja1.Cells(33, 18) = oRs1![10-12M]
                xlHoja1.Cells(33, 20) = oRs1![1-2A]
                xlHoja1.Cells(33, 22) = oRs1![2-5A]
                xlHoja1.Cells(33, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 5 Then 'MicroEmpresas
                xlHoja1.Cells(34, 4) = oRs1![1M]
                xlHoja1.Cells(34, 6) = oRs1![2M]
                xlHoja1.Cells(34, 8) = oRs1![3M]
                xlHoja1.Cells(34, 10) = oRs1![4M]
                xlHoja1.Cells(34, 12) = oRs1![5M]
                xlHoja1.Cells(34, 14) = oRs1![6M]
                xlHoja1.Cells(34, 16) = oRs1![7-9M]
                xlHoja1.Cells(34, 18) = oRs1![10-12M]
                xlHoja1.Cells(34, 20) = oRs1![1-2A]
                xlHoja1.Cells(34, 22) = oRs1![2-5A]
                xlHoja1.Cells(34, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 7 Then 'Consumos
                xlHoja1.Cells(51, 4) = oRs1![1M]
                xlHoja1.Cells(51, 6) = oRs1![2M]
                xlHoja1.Cells(51, 8) = oRs1![3M]
                xlHoja1.Cells(51, 10) = oRs1![4M]
                xlHoja1.Cells(51, 12) = oRs1![5M]
                xlHoja1.Cells(51, 14) = oRs1![6M]
                xlHoja1.Cells(51, 16) = oRs1![7-9M]
                xlHoja1.Cells(51, 18) = oRs1![10-12M]
                xlHoja1.Cells(51, 20) = oRs1![1-2A]
                xlHoja1.Cells(51, 22) = oRs1![2-5A]
                xlHoja1.Cells(51, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 8 Then 'Hip.Vigentes
                xlHoja1.Cells(45, 4) = oRs1![1M]
                xlHoja1.Cells(45, 6) = oRs1![2M]
                xlHoja1.Cells(45, 8) = oRs1![3M]
                xlHoja1.Cells(45, 10) = oRs1![4M]
                xlHoja1.Cells(45, 12) = oRs1![5M]
                xlHoja1.Cells(45, 14) = oRs1![6M]
                xlHoja1.Cells(45, 16) = oRs1![7-9M]
                xlHoja1.Cells(45, 18) = oRs1![10-12M]
                xlHoja1.Cells(45, 20) = oRs1![1-2A]
                xlHoja1.Cells(45, 22) = oRs1![2-5A]
                xlHoja1.Cells(45, 24) = oRs1![m5A]
        End If
    oRs1.MoveNext
    Loop
    End If
    
    'SALDOS DE CRÉDITOS REFINANCIADOS
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosRTramosResidual("1", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 6 Then
                xlHoja1.Cells(25, 3) = oRs1![1M]
                xlHoja1.Cells(25, 5) = oRs1![2M]
                xlHoja1.Cells(25, 7) = oRs1![3M]
                xlHoja1.Cells(25, 9) = oRs1![4M]
                xlHoja1.Cells(25, 11) = oRs1![5M]
                xlHoja1.Cells(25, 13) = oRs1![6M]
                xlHoja1.Cells(25, 15) = oRs1![7-9M]
                xlHoja1.Cells(25, 17) = oRs1![10-12M]
                xlHoja1.Cells(25, 19) = oRs1![1-2A]
                xlHoja1.Cells(25, 21) = oRs1![2-5A]
                xlHoja1.Cells(25, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 2 Then
                xlHoja1.Cells(26, 3) = oRs1![1M]
                xlHoja1.Cells(26, 5) = oRs1![2M]
                xlHoja1.Cells(26, 7) = oRs1![3M]
                xlHoja1.Cells(26, 9) = oRs1![4M]
                xlHoja1.Cells(26, 11) = oRs1![5M]
                xlHoja1.Cells(26, 13) = oRs1![6M]
                xlHoja1.Cells(26, 15) = oRs1![7-9M]
                xlHoja1.Cells(26, 17) = oRs1![10-12M]
                xlHoja1.Cells(26, 19) = oRs1![1-2A]
                xlHoja1.Cells(26, 21) = oRs1![2-5A]
                xlHoja1.Cells(26, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 3 Then
                xlHoja1.Cells(27, 3) = oRs1![1M]
                xlHoja1.Cells(27, 5) = oRs1![2M]
                xlHoja1.Cells(27, 7) = oRs1![3M]
                xlHoja1.Cells(27, 9) = oRs1![4M]
                xlHoja1.Cells(27, 11) = oRs1![5M]
                xlHoja1.Cells(27, 13) = oRs1![6M]
                xlHoja1.Cells(27, 15) = oRs1![7-9M]
                xlHoja1.Cells(27, 17) = oRs1![10-12M]
                xlHoja1.Cells(27, 19) = oRs1![1-2A]
                xlHoja1.Cells(27, 21) = oRs1![2-5A]
                xlHoja1.Cells(27, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 4 Then
                xlHoja1.Cells(40, 3) = oRs1![1M]
                xlHoja1.Cells(40, 5) = oRs1![2M]
                xlHoja1.Cells(40, 7) = oRs1![3M]
                xlHoja1.Cells(40, 9) = oRs1![4M]
                xlHoja1.Cells(40, 11) = oRs1![5M]
                xlHoja1.Cells(40, 13) = oRs1![6M]
                xlHoja1.Cells(40, 15) = oRs1![7-9M]
                xlHoja1.Cells(40, 17) = oRs1![10-12M]
                xlHoja1.Cells(40, 19) = oRs1![1-2A]
                xlHoja1.Cells(40, 21) = oRs1![2-5A]
                xlHoja1.Cells(40, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 5 Then
                xlHoja1.Cells(41, 3) = oRs1![1M]
                xlHoja1.Cells(41, 5) = oRs1![2M]
                xlHoja1.Cells(41, 7) = oRs1![3M]
                xlHoja1.Cells(41, 9) = oRs1![4M]
                xlHoja1.Cells(41, 11) = oRs1![5M]
                xlHoja1.Cells(41, 13) = oRs1![6M]
                xlHoja1.Cells(41, 15) = oRs1![7-9M]
                xlHoja1.Cells(41, 17) = oRs1![10-12M]
                xlHoja1.Cells(41, 19) = oRs1![1-2A]
                xlHoja1.Cells(41, 21) = oRs1![2-5A]
                xlHoja1.Cells(41, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 7 Then
                xlHoja1.Cells(52, 3) = oRs1![1M]
                xlHoja1.Cells(52, 5) = oRs1![2M]
                xlHoja1.Cells(52, 7) = oRs1![3M]
                xlHoja1.Cells(52, 9) = oRs1![4M]
                xlHoja1.Cells(52, 11) = oRs1![5M]
                xlHoja1.Cells(52, 13) = oRs1![6M]
                xlHoja1.Cells(52, 15) = oRs1![7-9M]
                xlHoja1.Cells(52, 17) = oRs1![10-12M]
                xlHoja1.Cells(52, 19) = oRs1![1-2A]
                xlHoja1.Cells(52, 21) = oRs1![2-5A]
                xlHoja1.Cells(52, 23) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 8 Then
                xlHoja1.Cells(46, 3) = oRs1![1M]
                xlHoja1.Cells(46, 5) = oRs1![2M]
                xlHoja1.Cells(46, 7) = oRs1![3M]
                xlHoja1.Cells(46, 9) = oRs1![4M]
                xlHoja1.Cells(46, 11) = oRs1![5M]
                xlHoja1.Cells(46, 13) = oRs1![6M]
                xlHoja1.Cells(46, 15) = oRs1![7-9M]
                xlHoja1.Cells(46, 17) = oRs1![10-12M]
                xlHoja1.Cells(46, 19) = oRs1![1-2A]
                xlHoja1.Cells(46, 21) = oRs1![2-5A]
                xlHoja1.Cells(46, 23) = oRs1![m5A]
        End If
    oRs1.MoveNext
    Loop
    End If
    
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerCreditosRTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 6 Then
                xlHoja1.Cells(25, 4) = oRs1![1M]
                xlHoja1.Cells(25, 6) = oRs1![2M]
                xlHoja1.Cells(25, 8) = oRs1![3M]
                xlHoja1.Cells(25, 10) = oRs1![4M]
                xlHoja1.Cells(25, 12) = oRs1![5M]
                xlHoja1.Cells(25, 14) = oRs1![6M]
                xlHoja1.Cells(25, 16) = oRs1![7-9M]
                xlHoja1.Cells(25, 18) = oRs1![10-12M]
                xlHoja1.Cells(25, 20) = oRs1![1-2A]
                xlHoja1.Cells(25, 22) = oRs1![2-5A]
                xlHoja1.Cells(25, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 2 Then
                xlHoja1.Cells(26, 4) = oRs1![1M]
                xlHoja1.Cells(26, 6) = oRs1![2M]
                xlHoja1.Cells(26, 8) = oRs1![3M]
                xlHoja1.Cells(26, 10) = oRs1![4M]
                xlHoja1.Cells(26, 12) = oRs1![5M]
                xlHoja1.Cells(26, 14) = oRs1![6M]
                xlHoja1.Cells(26, 16) = oRs1![7-9M]
                xlHoja1.Cells(26, 18) = oRs1![10-12M]
                xlHoja1.Cells(26, 20) = oRs1![1-2A]
                xlHoja1.Cells(26, 22) = oRs1![2-5A]
                xlHoja1.Cells(26, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 3 Then
                xlHoja1.Cells(27, 4) = oRs1![1M]
                xlHoja1.Cells(27, 6) = oRs1![2M]
                xlHoja1.Cells(27, 8) = oRs1![3M]
                xlHoja1.Cells(27, 10) = oRs1![4M]
                xlHoja1.Cells(27, 12) = oRs1![5M]
                xlHoja1.Cells(27, 14) = oRs1![6M]
                xlHoja1.Cells(27, 16) = oRs1![7-9M]
                xlHoja1.Cells(27, 18) = oRs1![10-12M]
                xlHoja1.Cells(27, 20) = oRs1![1-2A]
                xlHoja1.Cells(27, 22) = oRs1![2-5A]
                xlHoja1.Cells(27, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 4 Then
                xlHoja1.Cells(40, 4) = oRs1![1M]
                xlHoja1.Cells(40, 6) = oRs1![2M]
                xlHoja1.Cells(40, 8) = oRs1![3M]
                xlHoja1.Cells(40, 10) = oRs1![4M]
                xlHoja1.Cells(40, 12) = oRs1![5M]
                xlHoja1.Cells(40, 14) = oRs1![6M]
                xlHoja1.Cells(40, 16) = oRs1![7-9M]
                xlHoja1.Cells(40, 18) = oRs1![10-12M]
                xlHoja1.Cells(40, 20) = oRs1![1-2A]
                xlHoja1.Cells(40, 22) = oRs1![2-5A]
                xlHoja1.Cells(40, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 5 Then
                xlHoja1.Cells(41, 4) = oRs1![1M]
                xlHoja1.Cells(41, 6) = oRs1![2M]
                xlHoja1.Cells(41, 8) = oRs1![3M]
                xlHoja1.Cells(41, 10) = oRs1![4M]
                xlHoja1.Cells(41, 12) = oRs1![5M]
                xlHoja1.Cells(41, 14) = oRs1![6M]
                xlHoja1.Cells(41, 16) = oRs1![7-9M]
                xlHoja1.Cells(41, 18) = oRs1![10-12M]
                xlHoja1.Cells(41, 20) = oRs1![1-2A]
                xlHoja1.Cells(41, 22) = oRs1![2-5A]
                xlHoja1.Cells(41, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 7 Then
                xlHoja1.Cells(52, 4) = oRs1![1M]
                xlHoja1.Cells(52, 6) = oRs1![2M]
                xlHoja1.Cells(52, 8) = oRs1![3M]
                xlHoja1.Cells(52, 10) = oRs1![4M]
                xlHoja1.Cells(52, 12) = oRs1![5M]
                xlHoja1.Cells(52, 14) = oRs1![6M]
                xlHoja1.Cells(52, 16) = oRs1![7-9M]
                xlHoja1.Cells(52, 18) = oRs1![10-12M]
                xlHoja1.Cells(52, 20) = oRs1![1-2A]
                xlHoja1.Cells(52, 22) = oRs1![2-5A]
                xlHoja1.Cells(52, 24) = oRs1![m5A]
        End If
        If oRs1!cTpoCredCod = 8 Then
                xlHoja1.Cells(46, 4) = oRs1![1M]
                xlHoja1.Cells(46, 6) = oRs1![2M]
                xlHoja1.Cells(46, 8) = oRs1![3M]
                xlHoja1.Cells(46, 10) = oRs1![4M]
                xlHoja1.Cells(46, 12) = oRs1![5M]
                xlHoja1.Cells(46, 14) = oRs1![6M]
                xlHoja1.Cells(46, 16) = oRs1![7-9M]
                xlHoja1.Cells(46, 18) = oRs1![10-12M]
                xlHoja1.Cells(46, 20) = oRs1![1-2A]
                xlHoja1.Cells(46, 22) = oRs1![2-5A]
                xlHoja1.Cells(46, 24) = oRs1![m5A]
        End If
    oRs1.MoveNext
    Loop
    End If
    
    '***INTERESES DEVENGADOS DESCOMENTAR URGNET
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerInteresDevengadoResidual(pdFecha, 1)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTipoCredito = 105 Then 'SistFin
                xlHoja1.Cells(16, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 106 Then 'Corporativos
                xlHoja1.Cells(17, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 200 Then 'Grand.Empr
                xlHoja1.Cells(18, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 300 Then 'Medi.Empr
                xlHoja1.Cells(19, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 500 Then 'MicroEmpr
                xlHoja1.Cells(35, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 400 Then 'Peq.Empr
                xlHoja1.Cells(36, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 700 Then 'Consu
                xlHoja1.Cells(53, 3) = oRs1!nSaldo
        End If
        If oRs1!cTipoCredito = 800 Then 'Hip
                xlHoja1.Cells(47, 3) = oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
    End If

    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerInteresDevengadoResidual(pdFecha, "2")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTipoCredito = 105 Then
                xlHoja1.Cells(16, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 106 Then
                xlHoja1.Cells(17, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 200 Then
                xlHoja1.Cells(18, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 300 Then
                xlHoja1.Cells(19, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 500 Then
                xlHoja1.Cells(35, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 400 Then
                xlHoja1.Cells(36, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 700 Then
                xlHoja1.Cells(53, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        If oRs1!cTipoCredito = 800 Then
                xlHoja1.Cells(47, 4) = Round(oRs1!nSaldo / lnTipoCambioFC, 2)
        End If
        oRs1.MoveNext
    Loop
    End If
    '******************END TRASLADO BY NAGL 20191030 Según Anx03_ERS006-2019******************'

    '***INTERESES FUTUROS VIGENTES BY NAGL 20191030 Según Anx03_ERS006-2019****************'
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerInteresFuturoResidual(pdFecha, CStr(lnTipoCambioFC), "VIG")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 5 Then 'SistFin
                xlHoja1.Cells(20, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 6 Then 'Corporativos
                xlHoja1.Cells(21, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 2 Then 'Grand.Empr
                xlHoja1.Cells(22, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 3 Then 'Medi.Empr
                xlHoja1.Cells(23, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 5 Then 'MicroEmpr
                xlHoja1.Cells(37, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 4 Then 'Peq.Empr
                xlHoja1.Cells(38, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 7 Then 'Consu
                xlHoja1.Cells(54, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 8 Then 'Hip
                xlHoja1.Cells(48, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
    End If
    '***INTERESES FUTUROS REFINANCIADO BY NAGL 20191030 Según Anx03_ERS006-2019****************'
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDbalanceCont.ObtenerInteresFuturoResidual(pdFecha, CStr(lnTipoCambioFC), "REF")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!cTpoCredCod = 1 And oRs1!nTpoInstCorp = 6 Then 'Corporativos
                xlHoja1.Cells(28, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 2 Then 'Grand.Empr
                xlHoja1.Cells(29, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 3 Then 'Medi.Empr
                xlHoja1.Cells(30, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 5 Then 'MicroEmpr
                xlHoja1.Cells(43, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 4 Then 'Peq.Empr
                xlHoja1.Cells(42, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 7 Then 'Consu
                xlHoja1.Cells(55, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        If oRs1!cTpoCredCod = 8 Then 'Hip
                xlHoja1.Cells(49, oRs1!cRango + (oRs1!cRango + IIf(oRs1!nMoneda = "1", 1, 2))) = oRs1!nSaldo
        End If
        oRs1.MoveNext
    Loop
    End If
    '******************************************************************************************'
    
    lsNomHoja = "Disponible"
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
    
    'InicioAhorro
    'Call PintaFondeoAhorro(xlHoja1, pdFecha, lnTipoCambioFC, 1)
    'InicioCTS
    'Call PintaFondeoCTS(xlHoja1, pdFecha, lnTipoCambioFC, 1)
    '**********Comentado by NAGL 20190514**********
    
    'InicioObligacionesVista
    Call PintaFondeoObligacionesVista(xlHoja1, pdFecha, lnTipoCambioFC, 1) 'NAGL ERS 079-2016 20170407
    
    xlHoja1.Cells(110, 2) = Format(lnTipoCambioFC, "#,##0.0000")
    
    lsNomHoja = "Anx16AME"
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
    'InicioAhorro
    'Disponible
    Set oRs1 = oDbalanceCont.ObtenerOverNightTramosResidual("2", pdFecha, "2")
    
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(9, 2 + oRs1!cRango) = oRs1!nSaldo
    oRs1.MoveNext
    Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(9, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    
    '*******************************************************************************************
    Set oRs1 = New ADODB.Recordset
    'Inversiones a Vencimiento
    Set oRs1 = oDbalanceCont.ObtenerInversionesAVencimientoResidual(pdFecha, 2)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(12, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    
    'xlHoja1.Cells(11, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132401", pdFecha, "1", lnTipoCambioFC) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132402", pdFecha, "1", 0)
   
   'Inversiones Disponibles para la venta
    Set oRs1 = oDbalanceCont.ObtenerInversionesVentaTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(11, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    xlHoja1.Cells(11, 3) = CCur(xlHoja1.Cells(11, 3)) + (oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132402", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC) 'CDBCRP
    xlHoja1.Range(xlHoja1.Cells(11, 3), xlHoja1.Cells(11, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '**********************NAGL ERS 079-2016 20170407
    
    oBarra.Progress 8, TituloProgress, MensajeProgress, "", vbBlue
    
    Set oRs1 = New ADODB.Recordset
    'Cuentas por Cobrar - Operaciones de Reporte
    Set oRs1 = oDbalanceCont.ObtenerCuentasxCobrarTramosResidual("2", pdFecha)
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(45, 2 + oRs1!cRango) = oRs1!nSaldo
    oRs1.MoveNext
    Loop
    End If
    xlHoja1.Range(xlHoja1.Cells(45, 3), xlHoja1.Cells(45, 13)).NumberFormat = "#,##0.00;-#,##0.00"
    '**********NAGL ERS 079-2016 20170407
    
    'Fondeo PF
    'Call PintaFondeo(xlHoja1, pdFecha, lnTipoCambioFC, "2") 'Comentado by NAGL 20190514
    'Fin Fondeo PF
    
    Set oRs1 = oDbalanceCont.ObtenerRestringidosxTramosxProducto("2", pdFecha, 0, "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(52, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    
    Set oRs1 = oDbalanceCont.ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto(pdFecha, "2", "233")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(54, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "2", "1")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(55, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    xlHoja1.Cells(55, 3) = CCur(xlHoja1.Cells(55, 3)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242802", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2) 'Activado by NAGL 20190626 Primera Banda
 
    Set oRs1 = oDbalanceCont.ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(pdFecha, "2", "0")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        xlHoja1.Cells(56, 2 + oRs1!cRango) = oRs1!nSaldo
        oRs1.MoveNext
    Loop
    End If
    xlHoja1.Cells(56, 3) = CCur(xlHoja1.Cells(56, 3)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242807", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2) 'Activado by NAGL 20190626 Primera Banda
    
    'xlHoja1.Cells(56, 3) = CCur(xlHoja1.Cells(56, 3)) + Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242802", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
    'xlHoja1.Cells(57, 3) = CCur(xlHoja1.Cells(57, 3)) + Round((oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242806", pdFecha, "2", lnTipoCambioFC) + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("24280501", pdFecha, "2", lnTipoCambioFC)) / lnTipoCambioFC, 2)
    'xlHoja1.Cells(56, 3) = CCur(xlHoja1.Cells(56, 3)) + (oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("24", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC)
    'Comentado by NAGL 20190514
    
    'Inversiones a valor Razonable con cambios en resultados - supuesto
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1321", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
    xlHoja1.Cells(65, 3) = nSaldoDiario1 'NAGL ERS 079-2016 20170407
    
    'Comentado by NAGL 20190515
    'nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2121", pdFecha, "2", lnTipoCambioFC)
    'xlHoja1.Cells(78, 3) = Round(nSaldoDiario1 / lnTipoCambioFC, 2)
    
    'Depósitos de empresas del sistema financiero y OFI (15)
    nSaldoDiario1 = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2322", pdFecha, "2", lnTipoCambioFC)
    xlHoja1.Cells(87, 3) = Round(nSaldoDiario1 / lnTipoCambioFC, 2)
    
    cargarDatosBalanceANX16 xlHoja1.Application, 2, pdFecha, lnTipoCambioFC '***NAGL ERS 079-2016 20170407
    
    lsNomHoja = "DisponibleDolares"
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
    
    '************************************************************
    'InicioAhorro
    'Call PintaFondeoAhorro(xlHoja1, pdFecha, lnTipoCambioFC, "2")
    'InicioCTS
    'Call PintaFondeoCTS(xlHoja1, pdFecha, lnTipoCambioFC, "2")
    '*************Comentado by NAGL 20190514**********************
    
    'InicioVista
    Call PintaFondeoObligacionesVista(xlHoja1, pdFecha, lnTipoCambioFC, 2) '***NAGL ERS 079-2016 20170407
    
    lsNomHoja = "Anx16AInd"
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
    
    'InicioAhorro
    'stp_sel_Indicadores16Acredores
    nTotalAcredoresTo = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2102", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210303", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210305", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2107", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2302", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2303", pdFecha, "0", lnTipoCambioFC)
    
    'JIPR20200824
    'nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2406", pdFecha, "0", lnTipoCambioFC)
    'nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2600", pdFecha, "0", lnTipoCambioFC)
    
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("24", pdFecha, "0", lnTipoCambioFC)
    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("26", pdFecha, "0", lnTipoCambioFC)
    
'    nTotalAcredoresTo = nTotalAcredoresTo - oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2408", pdFecha, "0", lnTipoCambioFC) '*NAGL
'    nTotalAcredoresTo = nTotalAcredoresTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2606", pdFecha, "0", lnTipoCambioFC) '*NAGL
'    nTotalAcredoresTo = nTotalAcredoresTo - oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2608", pdFecha, "0", lnTipoCambioFC) '*NAGL
    
    nTotalDepositantesTo = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2102", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210303", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("210305", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2107", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2302", pdFecha, "0", lnTipoCambioFC)
    nTotalDepositantesTo = nTotalDepositantesTo + oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2303", pdFecha, "0", lnTipoCambioFC)
    
    '**************************************ERS 079-2016 20170407
    
    'DEL REPORTE 1 - BCRP
    xlHoja1.Cells(4, 13) = Round(lnEncajeExigibleRGMN, 2)
    xlHoja1.Cells(5, 13) = Day(pdFecha)
    xlHoja1.Cells(4, 14) = Round(lnEncajeExigibleRGME, 2)
    
    'DE BALANCE CONSOLIDADO
    xlHoja1.Cells(9, 13) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("24", pdFecha, "0", lnTipoCambioFC)
    xlHoja1.Cells(10, 13) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1", pdFecha, "0", lnTipoCambioFC)
    
    'De Activos Liquidos Anexo 15C
    
    Dim CajasFFMN As Currency, CajasFFME As Currency
    Dim FondosBCRPMN As Currency, FondosBCRPME As Currency
    Dim FondosSFNMN As Currency, FondosSFNME As Currency
    Dim FondosNActNMN As Currency, FondosNActNME As Currency
    Dim ValorBCRPMN As Currency
    Dim ValorBCRPMN2 As Currency
    Dim ValorGCMN As Currency
    Dim SumaTotalCMN As Currency, SumaTotalCME As Currency
    
   
    FondosBCRPMN = 0
    FondosSFNMN = 0
    FondosNActNMN = 0
    ValorBCRPMN = 0
    ValorBCRPMN2 = 0
    CajasFFMN = 0
    FondosBCRPME = 0
    FondosSFNME = 0
    FondosNActNME = 0
    
    ldFechaPro = DateAdd("d", -Day(pdFecha), pdFecha)
    For ix = 1 To Day(pdFecha)
    ldFechaPro = DateAdd("d", 1, ldFechaPro)
    
    CajasFFMN = CajasFFMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "300")
    FondosBCRPMN = FondosBCRPMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "425")
    FondosSFNMN = FondosSFNMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "450")
    FondosNActNMN = FondosNActNMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "600")
    ValorBCRPMN = ValorBCRPMN + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "725")
    ValorBCRPMN2 = ValorBCRPMN2 + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "1", "A1", "750")
    
    'ME
    CajasFFME = CajasFFME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "300")
    FondosBCRPME = FondosBCRPME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "425")
    FondosSFNME = FondosSFNME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "450")
    FondosNActNME = FondosNActNME + oDbalanceCont.ObtenerActivosLiquidosReporte15A(ldFechaPro, "2", "A1", "600")
    Next ix
    
    SumaTotalCMN = CajasFFMN + FondosBCRPMN + FondosSFNMN + FondosNActNMN + ValorBCRPMN + ValorBCRPMN2
    SumaTotalCME = CajasFFME + FondosBCRPME + FondosSFNME + FondosNActNME
    
    xlHoja1.Cells(5, 6) = Round(SumaTotalCMN / Day(pdFecha), 2)
    xlHoja1.Cells(6, 6) = Round(SumaTotalCME / Day(pdFecha), 2)
    xlHoja1.Cells(7, 6) = Format(lnTipoCambioFC, "#,##0.0000")
    '**************************************ERS 079-2016 20170407
     
    Dim nSaldoFondeoTotal As Currency
    '********Comentado by NAGL 20190516
    'nSaldoDiario1 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazo("1", pdFecha, 0, 30) + oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazo("2", pdFecha, lnTipoCambioFC, 30)
    'nSaldoDiario2 = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado("1", pdFecha, 0) + oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado("2", pdFecha, lnTipoCambioFC)
    'xlHoja1.Cells(29, 2) = nSaldoDiario1
    'xlHoja1.Cells(29, 3) = nSaldoDiario2
        
    nSaldoDiario1 = oDbalanceCont.ObtenerAdeudadosExterior(pdFecha, "1", "0") + (oDbalanceCont.ObtenerAdeudadosExterior(pdFecha, "2", "0") / lnTipoCambioFC) '**NAGL
    xlHoja1.Cells(10, 3) = nSaldoDiario1 / oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2", pdFecha, "0", lnTipoCambioFC)
    'NAGL ERS 079-2016 20170407
    
    'nSaldoFondeoTotal = oDbalanceCont.ObtenerIndicadores16FondeoCubierto(pdFecha, lnTipoCambioFC)
    'xlHoja1.Cells(11, 3) = Round(nSaldoFondeoTotal / nTotalDepositantesTo, 2)
    'Comentado by NAGL 20190516
     
    oBarra.Progress 9, TituloProgress, MensajeProgress, "", vbBlue
    
    '**********NAGL Agregó esta sección ERS006-2019******************'
    Set oRs1 = New ADODB.Recordset
    Set oRs1 = oDAnx.ObtieneFondeoAnx16ANew(pdFecha, lnTipoCambioFC, "16AInd")
    If Not (oRs1.BOF Or oRs1.EOF) Then
    Do While Not oRs1.EOF
        If oRs1!nTipoCobertura = 1 Then
                xlHoja1.Cells(11, 3) = Round(oRs1!nTotal / nTotalDepositantesTo, 2)
                xlHoja1.Cells(33, 2) = Round(oRs1!nVenc30Dias, 2)
                xlHoja1.Cells(33, 3) = Round(oRs1!nTotal, 2)
        ElseIf oRs1!nTipoCobertura = 0 Then
                xlHoja1.Cells(34, 2) = Round(oRs1!nVenc30Dias, 2)
                xlHoja1.Cells(34, 3) = Round(oRs1!nTotal, 2)
        ElseIf oRs1!nTipoCobertura = 3 Then
                xlHoja1.Cells(29, 2) = Round(oRs1!nVenc30Dias, 2)
                xlHoja1.Cells(29, 3) = Round(oRs1!nTotal, 2)
        End If
        oRs1.MoveNext
    Loop
    End If
    '***********************20190516*********************************'
    
    '****CUADRO DE VALIDACIÓN DE SALDOS
    xlHoja1.Cells(30, 3) = nTotalDepositantesTo
    xlHoja1.Cells(31, 3) = nTotalAcredoresTo
    '**************Comentado by NAGL 20190514*********
    'xlHoja1.Cells(33, 2) = oDbalanceCont.ObtenerFondeoFSD(pdFecha, lnTipoCambioFC, 30, 1)
    'xlHoja1.Cells(33, 3) = nSaldoFondeoTotal
    'xlHoja1.Cells(34, 2) = oDbalanceCont.ObtenerFondeoFSD(pdFecha, lnTipoCambioFC, 30, 0)
    'xlHoja1.Cells(34, 3) = oDbalanceCont.ObtenerFondeoFSD(pdFecha, lnTipoCambioFC, 30000, 0)
    '*************END NAGL 20190514********************
    xlHoja1.Cells(42, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2", pdFecha, "0", lnTipoCambioFC)
    '**********NAGL ERS 079-2016 20170407
    cargarDetalleAcreedoresDepositantes xlHoja1.Application, 0, pdFecha, lnTipoCambioFC 'NAGL ERS 079-2016 20170407
    cargarDetalleAcreedoresDepositantes xlHoja1.Application, 1, pdFecha, lnTipoCambioFC 'NAGL ERS 079-2016 20170407
    
    '**************************ERS 079-2016 20170407
    lsNomHoja = "Anx16BReg"
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
        
        'Inversiones a valor razonable con cambios en resultados
        xlHoja1.Cells(71, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("13110705", pdFecha, "1", 0)
        xlHoja1.Cells(71, 4) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("13210705", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
        
        'CDBCRP
        xlHoja1.Cells(75, 3) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("131402", pdFecha, "1", 0)
        xlHoja1.Cells(75, 4) = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("132402", pdFecha, "2", lnTipoCambioFC) / lnTipoCambioFC
        
        
        Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("1", pdFecha, "1314010101")
            Do While Not oRs1.EOF
            If (oRs1!cRango < 9) Then
                xlHoja1.Cells(74, 1 + oRs1!cRango * 2) = oRs1!nSaldo
            Else
                xlHoja1.Cells(74, 19) = CCur(xlHoja1.Cells(74, 19)) + oRs1!nSaldo
            End If
            oRs1.MoveNext
            Loop
                      
       Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("1", pdFecha, "131407190[12]")
            Do While Not oRs1.EOF
            If (oRs1!cRango < 9) Then
                xlHoja1.Cells(76, 1 + oRs1!cRango * 2) = oRs1!nSaldo
            Else
                xlHoja1.Cells(76, 19) = CCur(xlHoja1.Cells(76, 19)) + oRs1!nSaldo
            End If
            oRs1.MoveNext
            Loop
            
       Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("1", pdFecha, "13140507")
            Do While Not oRs1.EOF
            If (oRs1!cRango < 9) Then
                xlHoja1.Cells(77, 1 + oRs1!cRango * 2) = oRs1!nSaldo
            Else
                xlHoja1.Cells(77, 19) = CCur(xlHoja1.Cells(77, 19)) + oRs1!nSaldo
            End If
            oRs1.MoveNext
            Loop
                
          Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("2", pdFecha, "1324010101")
            Do While Not oRs1.EOF
            If (oRs1!cRango < 9) Then
                xlHoja1.Cells(74, 2 + oRs1!cRango * 2) = oRs1!nSaldo
            Else
                xlHoja1.Cells(74, 20) = CCur(xlHoja1.Cells(74, 20)) + oRs1!nSaldo
            End If
            oRs1.MoveNext
            Loop

       Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("2", pdFecha, "132407190[12]")
            Do While Not oRs1.EOF
            If (oRs1!cRango < 9) Then
                xlHoja1.Cells(76, 2 + oRs1!cRango * 2) = oRs1!nSaldo
            Else
                xlHoja1.Cells(76, 20) = CCur(xlHoja1.Cells(76, 20)) + oRs1!nSaldo
            End If
            oRs1.MoveNext
            Loop
            
        Set oRs1 = oDbalanceCont.ObtenerInversionesVenta16B("2", pdFecha, "13240507")
            Do While Not oRs1.EOF
            If (oRs1!cRango < 9) Then
                xlHoja1.Cells(77, 2 + oRs1!cRango * 2) = oRs1!nSaldo
            Else
                xlHoja1.Cells(77, 20) = CCur(xlHoja1.Cells(77, 20)) + oRs1!nSaldo
            End If
            oRs1.MoveNext
            Loop
        '********************************************NAGL ERS 079-2016 20170407
    
    oBarra.Progress 10, "ANEXO 16A: Cuadro de Liquidez por plazos de Vencimiento", "Generación Terminada", "", vbBlue
    oBarra.CloseForm frmReportes
    Set oBarra = Nothing 'NAGL20170407

    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub cargarDetalleAcreedoresDepositantes(ByVal pobj_Excel As Excel.Application, ByVal pTipo As Integer, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oDBalance As DbalanceCont '***NAGL
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Set prs = New ADODB.Recordset
    Set oDBalance = New DbalanceCont
    
        Set prs = oDBalance.ObtenerDetalleAcreedoresDepositantes(psFecha, lnTipoCambioFC, 20, pTipo)
        If Not prs.EOF Or prs.BOF Then
            If pTipo = 0 Then
                nFilas = 46
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("Anx16AInd!A" & nFilas)
                    pcelda.value = IIf(pTipo = 0, prs(1), "")
                    Set pcelda = pobj_Excel.Range("Anx16AInd!B" & nFilas)
                    pcelda.value = IIf(pTipo = 0, prs(2), 0)
                    Set pcelda = pobj_Excel.Range("Anx16AInd!C" & nFilas)
                    pcelda.value = IIf(pTipo = 0, prs(3), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
                ElseIf pTipo = 1 Then
                nFilas = 69
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("Anx16AInd!A" & nFilas)
                    pcelda.value = IIf(pTipo = 1, prs(1), "")
                    Set pcelda = pobj_Excel.Range("Anx16AInd!B" & nFilas)
                    pcelda.value = IIf(pTipo = 1, prs(2), 0)
                    Set pcelda = pobj_Excel.Range("Anx16AInd!C" & nFilas)
                    pcelda.value = IIf(pTipo = 1, prs(3), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            End If
        End If
        Set prs = Nothing
End Sub '***NAGL ERS 079-2016 20170407*****'

Private Sub cargarPasivosObligxCtaAhorrosANX6(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim oDbalanceCont As New DbalanceCont '***NAGL
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Cant = 0
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
        
        Set prs = oCtaIf.GetObligxCtaAhorrosSBSanx6A(Format(psFecha, "yyyymmdd"), cMoneda, "232")
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("Ahorro2102!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("Ahorro2102!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("Ahorro2102!A274")
                pobj_Excel.Range("Ahorro2102!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("Ahorro2102!B274")
                pobj_Excel.Range("Ahorro2102!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2112", psFecha, "1", lnTipoCambioFC)
                End If '***NAGL ERS 079-2016 20170407***'
                
            ElseIf cMoneda = 2 Then
                Cant = 0
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("Ahorro2102!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("Ahorro2102!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                    Cant = Cant + 1
                Loop
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("Ahorro2102!I274")
                pobj_Excel.Range("Ahorro2102!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("Ahorro2102!J274")
                pobj_Excel.Range("Ahorro2102!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2122", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If '***NAGL ERS 079-2016 20170407***'
            End If
        End If
        Set prs = Nothing
End Sub

Private Sub cargarPasivosObligxCtaCTSANX6(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim oDbalanceCont As New DbalanceCont '***NAGL
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Cant = 0
        Set prs = oCtaIf.GetObligxCtaAhorrosSBSanx6A(Format(psFecha, "yyyymmdd"), cMoneda, "234")
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("CTS210305!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTS210305!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("CTS210305!A274")
                pobj_Excel.Range("CTS210305!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("CTS210305!B274")
                pobj_Excel.Range("CTS210305!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211305", psFecha, "1", lnTipoCambioFC)
                End If '***NAGL ERS 079-2016 20170407***'
                 
            End If
            If cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("CTS210305!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTS210305!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("CTS210305!I274")
                pobj_Excel.Range("CTS210305!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("CTS210305!J274")
                pobj_Excel.Range("CTS210305!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212305", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If '***NAGL ERS 079-2016 20170407***'
            End If

        End If
        Set prs = Nothing
End Sub

Private Sub cargarDepositosInmovilizadosANX16(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As New NCajaCtaIF
    Dim oDbalanceCont As New DbalanceCont
    Dim prs As New ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    
    Cant = 0
        Set prs = oCtaIf.GetDepositosInmovilizadosSBSanx16A(Format(psFecha, "yyyymmdd"), cMoneda, lnTipoCambioFC)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("DepInmov210701!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("DepInmov210701!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("DepInmov210701!A274")
                pobj_Excel.Range("DepInmov210701!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("DepInmov210701!B274")
                pobj_Excel.Range("DepInmov210701!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211701", psFecha, "1", lnTipoCambioFC)
                End If
                 
            End If
            If cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("DepInmov210701!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("DepInmov210701!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("DepInmov210701!I274")
                pobj_Excel.Range("DepInmov210701!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("DepInmov210701!J274")
                pobj_Excel.Range("DepInmov210701!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212701", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If
            End If

        End If
        Set prs = Nothing
End Sub '************NAGL ERS 079-2016 20170407*********'

Private Sub cargarObligacionesVistaANX16(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal lnTipoCambioFC As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim oDbalanceCont As New DbalanceCont
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Cant = 0
        Set prs = oCtaIf.GetObligacionesVistaSBSanx16A(Format(psFecha, "yyyymmdd"), cMoneda, lnTipoCambioFC)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("OligVista2101!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("OligVista2101!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("OligVista2101!A274")
                pobj_Excel.Range("OligVista2101!A274").value = psFecha
                Set pcelda = pobj_Excel.Range("OligVista2101!B274")
                pobj_Excel.Range("OligVista2101!B274").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2111", psFecha, "1", lnTipoCambioFC)
                End If
                 
            End If
            If cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("OligVista2101!I" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("OligVista2101!J" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    Cant = Cant + 1
                    prs.MoveNext
                Loop
                
                If (Cant = 272) Then
                Set pcelda = pobj_Excel.Range("OligVista2101!I274")
                pobj_Excel.Range("OligVista2101!I274").value = psFecha
                Set pcelda = pobj_Excel.Range("OligVista2101!J274")
                pobj_Excel.Range("OligVista2101!J274").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2121", psFecha, "2", lnTipoCambioFC) / lnTipoCambioFC, 2)
                End If
            End If

        End If
        Set prs = Nothing
End Sub '************NAGL ERS 079-2016 20170407********'

Private Sub cargarDatosBalanceANX16(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String, psFecha As Date, ByVal nTipoCambio As Currency)
    Dim pcelda As Excel.Range
    Dim oDbalanceCont As New DbalanceCont
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim Cant As Integer
    Set prs = New ADODB.Recordset
    If (cMoneda = 1) Then
        Set pcelda = pobj_Excel.Range("Anx16AMN!X3")
        pobj_Excel.Range("Anx16AMN!X3").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1111", psFecha, "1", 0) 'Total Caja
        Set pcelda = pobj_Excel.Range("Anx16AMN!X4")
        pobj_Excel.Range("Anx16AMN!X4").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1112", psFecha, "1", 0) 'Total BCRP
        Set pcelda = pobj_Excel.Range("Anx16AMN!V4")
        pobj_Excel.Range("Anx16AMN!V4").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("11120[45]01", psFecha, "1", 0) 'Con vencimiento
        
        Set pcelda = pobj_Excel.Range("Anx16AMN!X5")
        pobj_Excel.Range("Anx16AMN!X5").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1113", psFecha, "1", 0) 'Total Banco y otras Empresas SFN
        
        pobj_Excel.Range("Anx16AMN!V5").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("11130___03", psFecha, "1", 0) 'Con vencimiento
        Set pcelda = pobj_Excel.Range("Anx16AMN!V5")
        
        pobj_Excel.Range("Anx16AMN!X6").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1115", psFecha, "1", 0) 'Total Canje.
        Set pcelda = pobj_Excel.Range("Anx16AMN!X6") 'NAGL 20170621
        pobj_Excel.Range("Anx16AMN!X7").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1116", psFecha, "1", 0) 'Total Efectos Cobro Inm.
        Set pcelda = pobj_Excel.Range("Anx16AMN!X7")
        pobj_Excel.Range("Anx16AMN!X8").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("111701", psFecha, "1", 0) 'Disponible Restring.'NAGL 20190522 Se considerará 111701
        Set pcelda = pobj_Excel.Range("Anx16AMN!X8")
        
        pobj_Excel.Range("Anx16AMN!X9").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1118", psFecha, "1", 0) 'Rend.Devengados
        Set pcelda = pobj_Excel.Range("Anx16AMN!X9")
        
        pobj_Excel.Range("Anx16AMN!V13").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("11", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!V13")
        
        pobj_Excel.Range("Anx16AMN!Q11").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1314", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q11")
        pobj_Excel.Range("Anx16AMN!Q12").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1315", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q12")
        
        pobj_Excel.Range("Anx16AMN!Q15").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141109", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q15")
        pobj_Excel.Range("Anx16AMN!Q16").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141110", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q16")
        pobj_Excel.Range("Anx16AMN!Q17").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141111", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q17")
        pobj_Excel.Range("Anx16AMN!Q18").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141112", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q18")
        pobj_Excel.Range("Anx16AMN!Q19").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141809", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q19")
        pobj_Excel.Range("Anx16AMN!Q20").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141810", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q20")
        pobj_Excel.Range("Anx16AMN!Q21").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141811", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q21")
        pobj_Excel.Range("Anx16AMN!Q22").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141812", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q22")
        
        pobj_Excel.Range("Anx16AMN!Q24").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141410", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q24")
        pobj_Excel.Range("Anx16AMN!Q25").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141411", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q25")
        pobj_Excel.Range("Anx16AMN!Q26").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141412", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q26")
        
        pobj_Excel.Range("Anx16AMN!Q29").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141113", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q29")
        pobj_Excel.Range("Anx16AMN!Q30").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141102", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q30")
        pobj_Excel.Range("Anx16AMN!Q31").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141802", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q31")
        pobj_Excel.Range("Anx16AMN!Q32").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141813", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q32")
        
        pobj_Excel.Range("Anx16AMN!Q34").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141413", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q34")
        pobj_Excel.Range("Anx16AMN!Q35").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141402", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q35")
        
        pobj_Excel.Range("Anx16AMN!Q37").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141104", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q37")
        pobj_Excel.Range("Anx16AMN!Q38").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141404", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q38")
        pobj_Excel.Range("Anx16AMN!Q39").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141804", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q39")
        
        pobj_Excel.Range("Anx16AMN!Q41").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141103", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q41")
        pobj_Excel.Range("Anx16AMN!Q42").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141403", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q42")
        pobj_Excel.Range("Anx16AMN!Q43").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("141803", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q43")
        
        'pobj_Excel.Range("Anx16AMN!Q45").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151711", psFecha, "1", 0)
        'Set pcelda = pobj_Excel.Range("Anx16AMN!Q45") Comentado by NAGL 20190515
        pobj_Excel.Range("Anx16AMN!Q49").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211303", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q49")
        pobj_Excel.Range("Anx16AMN!Q50").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211704", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q50")
        pobj_Excel.Range("Anx16AMN!Q51").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2118", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q51")
        
        pobj_Excel.Range("Anx16AMN!Q78").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2111", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q78")
        pobj_Excel.Range("Anx16AMN!Q80").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2112", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q80")
        pobj_Excel.Range("Anx16AMN!Q83").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211305", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q83")
        pobj_Excel.Range("Anx16AMN!Q87").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2312", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!Q87")
        
        'NAGL 20190515
        pobj_Excel.Range("Anx16AMN!T45").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2313", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T45")
        pobj_Excel.Range("Anx16AMN!T46").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2318", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T46")
        
        'NAGL ERS 006-2019 20190515
        'Otras Obligaciones con Inst.Recaudadoras
        pobj_Excel.Range("Anx16AMN!U47").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170301", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!U47")
        pobj_Excel.Range("Anx16AMN!U48").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170302", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!U48")
        pobj_Excel.Range("Anx16AMN!U49").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170303", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!U49")
        pobj_Excel.Range("Anx16AMN!T50").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170309", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T50")
        pobj_Excel.Range("Anx16AMN!T51").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251704", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T51")
        pobj_Excel.Range("Anx16AMN!T52").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170501", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T52")
        pobj_Excel.Range("Anx16AMN!T53").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2517050201", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T53")
        pobj_Excel.Range("Anx16AMN!T54").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251706", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T54")
        pobj_Excel.Range("Anx16AMN!T55").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2116", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T55")
        pobj_Excel.Range("Anx16AMN!T56").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211704", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T56")
        pobj_Excel.Range("Anx16AMN!T57").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("211701", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T57")
        pobj_Excel.Range("Anx16AMN!T58").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2118", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T58")
        pobj_Excel.Range("Anx16AMN!T59").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2518", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T59")
        
        'NAGL ERS006-2019 20190515
        'ADEUDADOS DEL PAIS
        pobj_Excel.Range("Anx16AMN!W46").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2411", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W46")
        pobj_Excel.Range("Anx16AMN!W47").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2412", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W47")
        pobj_Excel.Range("Anx16AMN!W48").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2413", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W48")
        pobj_Excel.Range("Anx16AMN!W49").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2416", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W49")
        pobj_Excel.Range("Anx16AMN!W50").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("241802", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W50")
        pobj_Excel.Range("Anx16AMN!W51").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2419", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W51")
        pobj_Excel.Range("Anx16AMN!W52").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2612", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W52")
        pobj_Excel.Range("Anx16AMN!W53").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2613", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W53")
        pobj_Excel.Range("Anx16AMN!W54").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2616", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W54")
        pobj_Excel.Range("Anx16AMN!W55").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2618", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W55")
        pobj_Excel.Range("Anx16AMN!W56").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2619", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W56")
        
        'NAGL ERS006-2019 20190515
        'ADEUDADOS DEL EXTERIOR
        pobj_Excel.Range("Anx16AMN!W59").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2414", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W59")
        pobj_Excel.Range("Anx16AMN!W60").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2415", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W60")
        pobj_Excel.Range("Anx16AMN!W61").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2417", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W61")
        pobj_Excel.Range("Anx16AMN!W62").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("241807", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W62")
        pobj_Excel.Range("Anx16AMN!W63").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2419", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W63")
        pobj_Excel.Range("Anx16AMN!W64").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2614", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W64")
        pobj_Excel.Range("Anx16AMN!W65").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2615", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W65")
        pobj_Excel.Range("Anx16AMN!W66").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2617", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W66")
        pobj_Excel.Range("Anx16AMN!W67").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2618", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W67")
        pobj_Excel.Range("Anx16AMN!W68").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2619", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!W68")
        
        'NAGL ERS006-2019 CAMBIOS LA CUENTAS POR COBRAR Y PAGAR
        'SECCIÓN CUENTAS POR COBRAR
        pobj_Excel.Range("Anx16AMN!T61").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151401", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T61")
        pobj_Excel.Range("Anx16AMN!T62").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151402", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T62")
        pobj_Excel.Range("Anx16AMN!T63").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151509", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T63")
        pobj_Excel.Range("Anx16AMN!T64").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151701", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T64")
        pobj_Excel.Range("Anx16AMN!T65").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151702", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T65")
        pobj_Excel.Range("Anx16AMN!T66").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15171905", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T66")
        pobj_Excel.Range("Anx16AMN!T67").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15171909", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T67")
        pobj_Excel.Range("Anx16AMN!T68").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15171910", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T68")
        pobj_Excel.Range("Anx16AMN!T69").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1512", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T69")
        pobj_Excel.Range("Anx16AMN!T70").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151405", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T70")
        pobj_Excel.Range("Anx16AMN!T71").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151711", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T71")
         pobj_Excel.Range("Anx16AMN!T72").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15150101", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T72")
        pobj_Excel.Range("Anx16AMN!T73").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15150102", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T73")
        pobj_Excel.Range("Anx16AMN!T74").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15150103", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T74")
        pobj_Excel.Range("Anx16AMN!T75").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15171903", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T75")
        pobj_Excel.Range("Anx16AMN!T76").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151703", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T76")
        pobj_Excel.Range("Anx16AMN!T77").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15171904", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T77")
        pobj_Excel.Range("Anx16AMN!T78").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15171902", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T78")
         pobj_Excel.Range("Anx16AMN!T79").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1519071901", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T79")
        pobj_Excel.Range("Anx16AMN!T80").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1519071902", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T80")
        pobj_Excel.Range("Anx16AMN!T81").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("151711", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T81")
        pobj_Excel.Range("Anx16AMN!T82").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1518071101", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T82")
        
        'SECCIÓN CUENTAS POR PAGAR
        pobj_Excel.Range("Anx16AMN!T83").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2514190201", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T83")
        pobj_Excel.Range("Anx16AMN!T84").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141903", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T84")
        pobj_Excel.Range("Anx16AMN!T85").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141904", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T85")
         pobj_Excel.Range("Anx16AMN!T86").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141905", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T86")
        pobj_Excel.Range("Anx16AMN!T87").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141906", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T87")
        pobj_Excel.Range("Anx16AMN!T88").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141907", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T88")
        pobj_Excel.Range("Anx16AMN!T89").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141910", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T89")
        pobj_Excel.Range("Anx16AMN!T90").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141912", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T90")
        pobj_Excel.Range("Anx16AMN!T91").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141913", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T91")
        pobj_Excel.Range("Anx16AMN!T92").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141914", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T92")
        pobj_Excel.Range("Anx16AMN!T93").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251501", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T93")
        pobj_Excel.Range("Anx16AMN!T94").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251502", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T94")
        pobj_Excel.Range("Anx16AMN!T95").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25150301", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T95")
        pobj_Excel.Range("Anx16AMN!T96").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25150401", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T96")
        pobj_Excel.Range("Anx16AMN!T97").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251505", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T97")
        pobj_Excel.Range("Anx16AMN!T98").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251506", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T98")
        pobj_Excel.Range("Anx16AMN!T99").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251509", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T99")
        pobj_Excel.Range("Anx16AMN!T100").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251601", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T100")
        pobj_Excel.Range("Anx16AMN!T101").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251602", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T101")
        pobj_Excel.Range("Anx16AMN!T102").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251702", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T102")
        pobj_Excel.Range("Anx16AMN!T103").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2517050205", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T103")
        pobj_Excel.Range("Anx16AMN!T104").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2512", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T104")
        pobj_Excel.Range("Anx16AMN!T105").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251402", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T105")
        pobj_Excel.Range("Anx16AMN!T106").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251701", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T106")
        pobj_Excel.Range("Anx16AMN!T107").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25150402", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T107")
        pobj_Excel.Range("Anx16AMN!T108").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2517050202", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T108")
        pobj_Excel.Range("Anx16AMN!T109").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2517050203", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T109")
        pobj_Excel.Range("Anx16AMN!T110").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25141915", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T110")
        pobj_Excel.Range("Anx16AMN!T111").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251411", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!T111")
        
        'Cuenta 15 NAGL 20190515
        pobj_Excel.Range("Anx16AMN!R73").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Anx16AMN!R73")
        
    ElseIf (cMoneda = 2) Then
        Set pcelda = pobj_Excel.Range("Anx16AME!X3")
        pobj_Excel.Range("Anx16AME!X3").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1121", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'Total Caja
        Set pcelda = pobj_Excel.Range("Anx16AME!X4")
        pobj_Excel.Range("Anx16AME!X4").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1122", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'Total BCRP
        Set pcelda = pobj_Excel.Range("Anx16AME!V4")
        pobj_Excel.Range("Anx16AME!V4").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("11220[45]01", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'Con vencimiento
        
        Set pcelda = pobj_Excel.Range("Anx16AME!X5")
        pobj_Excel.Range("Anx16AME!X5").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1123", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'Total Banco y otras Empresas SFN
        
        pobj_Excel.Range("Anx16AME!V5").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("11230___03", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'Con vencimiento
        Set pcelda = pobj_Excel.Range("Anx16AME!V5")
        
        pobj_Excel.Range("Anx16AME!X6").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1125", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!X6") 'NAGL 20170621
        pobj_Excel.Range("Anx16AME!X7").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1126", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!X7")
        pobj_Excel.Range("Anx16AME!X8").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("112701", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'NAGL 20190522 Se considerará 112701
        Set pcelda = pobj_Excel.Range("Anx16AME!X8")
        
        pobj_Excel.Range("Anx16AME!X9").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1128", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'Rend.Devengados
        Set pcelda = pobj_Excel.Range("Anx16AME!X9")
        
        pobj_Excel.Range("Anx16AME!V1").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("11", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!V1")
        
        
        pobj_Excel.Range("Anx16AME!Q11").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1324", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q11")
        pobj_Excel.Range("Anx16AME!Q12").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1325", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q12")
        
        pobj_Excel.Range("Anx16AME!Q15").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142109", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q15")
        pobj_Excel.Range("Anx16AME!Q16").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142110", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q16")
        pobj_Excel.Range("Anx16AME!Q17").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142111", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q17")
        pobj_Excel.Range("Anx16AME!Q18").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142112", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q18")
        pobj_Excel.Range("Anx16AME!Q19").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142809", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q19")
        pobj_Excel.Range("Anx16AME!Q20").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142810", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q20")
        pobj_Excel.Range("Anx16AME!Q21").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142811", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q21")
        pobj_Excel.Range("Anx16AME!Q22").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142812", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q22")
        
        pobj_Excel.Range("Anx16AME!Q24").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142410", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q24")
        pobj_Excel.Range("Anx16AME!Q25").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142411", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q25")
        pobj_Excel.Range("Anx16AME!Q26").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142412", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q26")
        
        pobj_Excel.Range("Anx16AME!Q29").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142113", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q29")
        pobj_Excel.Range("Anx16AME!Q30").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142102", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q30")
        pobj_Excel.Range("Anx16AME!Q31").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142802", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q31")
        pobj_Excel.Range("Anx16AME!Q32").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142813", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q32")
        
        pobj_Excel.Range("Anx16AME!Q34").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142413", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q34")
        pobj_Excel.Range("Anx16AME!Q35").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142402", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q35")
        
        pobj_Excel.Range("Anx16AME!Q37").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142104", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q37")
        pobj_Excel.Range("Anx16AME!Q38").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142404", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q38")
        pobj_Excel.Range("Anx16AME!Q39").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142804", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q39")
        
        pobj_Excel.Range("Anx16AME!Q41").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142103", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q41")
        pobj_Excel.Range("Anx16AME!Q42").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142403", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q42")
        pobj_Excel.Range("Anx16AME!Q43").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("142803", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q43")
        
        'pobj_Excel.Range("Anx16AME!Q45").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152711", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        'Set pcelda = pobj_Excel.Range("Anx16AME!Q45")'Comentado by NAGL 20190515
        pobj_Excel.Range("Anx16AME!Q49").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212303", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q49")
        pobj_Excel.Range("Anx16AME!Q50").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212704", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q50")
        pobj_Excel.Range("Anx16AME!Q51").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2128", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q51")
        
        pobj_Excel.Range("Anx16AME!Q65").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1321", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q65")
        
        pobj_Excel.Range("Anx16AME!Q78").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2121", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q78")
        pobj_Excel.Range("Anx16AME!Q80").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2122", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q80")
        pobj_Excel.Range("Anx16AME!Q83").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212305", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q83")
        pobj_Excel.Range("Anx16AME!Q87").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2322", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!Q87")
        
        'NAGL 20190515
        pobj_Excel.Range("Anx16AME!T45").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2323", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T45")
        pobj_Excel.Range("Anx16AME!T46").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2328", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T46")
        
        'NAGL ERS 006-2019 20190515
        'Otras Obligaciones con Inst.Recaudadoras
        pobj_Excel.Range("Anx16AME!U47").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270301", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!U47")
        pobj_Excel.Range("Anx16AME!U48").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270302", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!U48")
        pobj_Excel.Range("Anx16AME!U49").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270303", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!U49")
        pobj_Excel.Range("Anx16AME!T50").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270309", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T50")
        pobj_Excel.Range("Anx16AME!T51").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252704", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T51")
        pobj_Excel.Range("Anx16AME!T52").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270501", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T52")
        pobj_Excel.Range("Anx16AME!T53").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2527050201", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T53")
        pobj_Excel.Range("Anx16AME!T54").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252706", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T54")
        pobj_Excel.Range("Anx16AME!T55").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2126", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T55")
        pobj_Excel.Range("Anx16AME!T56").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212704", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T56")
        pobj_Excel.Range("Anx16AME!T57").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("212701", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T57")
        pobj_Excel.Range("Anx16AME!T58").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2128", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T58")
        pobj_Excel.Range("Anx16AME!T59").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2528", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T59")
        
        'NAGL ERS006-2019 20190515
        'ADEUDADOS DEL PAIS
        pobj_Excel.Range("Anx16AME!W46").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2421", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W46")
        pobj_Excel.Range("Anx16AME!W47").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2422", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W47")
        pobj_Excel.Range("Anx16AME!W48").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2423", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W48")
        pobj_Excel.Range("Anx16AME!W49").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2426", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W49")
        pobj_Excel.Range("Anx16AME!W50").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242802", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W50")
        pobj_Excel.Range("Anx16AME!W51").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2429", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W51")
        pobj_Excel.Range("Anx16AME!W52").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2622", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W52")
        pobj_Excel.Range("Anx16AME!W53").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2623", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W53")
        pobj_Excel.Range("Anx16AME!W54").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2626", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W54")
        pobj_Excel.Range("Anx16AME!W55").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2628", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W55")
        pobj_Excel.Range("Anx16AME!W56").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2629", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W56")
        
        'NAGL ERS006-2019 20190515
        'ADEUDADOS DEL EXTERIOR
        pobj_Excel.Range("Anx16AME!W59").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2424", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W59")
        pobj_Excel.Range("Anx16AME!W60").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2425", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W60")
        pobj_Excel.Range("Anx16AME!W61").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2427", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W61")
        pobj_Excel.Range("Anx16AME!W62").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("242807", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W62")
        pobj_Excel.Range("Anx16AME!W63").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2429", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W63")
        pobj_Excel.Range("Anx16AME!W64").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2624", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W64")
        pobj_Excel.Range("Anx16AME!W65").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2625", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W65")
        pobj_Excel.Range("Anx16AME!W66").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2627", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W66")
        pobj_Excel.Range("Anx16AME!W67").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2628", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W67")
        pobj_Excel.Range("Anx16AME!W68").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2629", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!W68")
        
        'NAGL ERS006-2019 CAMBIOS LA CUENTAS POR COBRAR Y PAGAR
        'SECCIÓN CUENTAS POR COBRAR
        pobj_Excel.Range("Anx16AME!T61").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152401", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T61")
        pobj_Excel.Range("Anx16AME!T62").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152402", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T62")
        pobj_Excel.Range("Anx16AME!T63").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152509", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T63")
        pobj_Excel.Range("Anx16AME!T64").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152701", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T64")
        pobj_Excel.Range("Anx16AME!T65").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152702", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T65")
        pobj_Excel.Range("Anx16AME!T66").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15271905", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T66")
        pobj_Excel.Range("Anx16AME!T67").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15271909", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T67")
        pobj_Excel.Range("Anx16AME!T68").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15271910", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T68")
        pobj_Excel.Range("Anx16AME!T69").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1522", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T69")
        pobj_Excel.Range("Anx16AME!T70").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152405", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T70")
        pobj_Excel.Range("Anx16AME!T71").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152711", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T71")
         pobj_Excel.Range("Anx16AME!T72").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15250101", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T72")
        pobj_Excel.Range("Anx16AME!T73").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15250102", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T73")
        pobj_Excel.Range("Anx16AME!T74").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15250103", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T74")
        pobj_Excel.Range("Anx16AME!T75").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15271903", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T75")
        pobj_Excel.Range("Anx16AME!T76").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152703", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T76")
        pobj_Excel.Range("Anx16AME!T77").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15271904", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T77")
        pobj_Excel.Range("Anx16AME!T78").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15271902", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T78")
         pobj_Excel.Range("Anx16AME!T79").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1529071901", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T79")
        pobj_Excel.Range("Anx16AME!T80").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1529071902", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T80")
        pobj_Excel.Range("Anx16AME!T81").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("152711", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T81")
        pobj_Excel.Range("Anx16AME!T82").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1528071101", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T82")
        
        'SECCIÓN CUENTAS POR PAGAR
        pobj_Excel.Range("Anx16AME!T83").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2524190201", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T83")
        pobj_Excel.Range("Anx16AME!T84").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241903", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T84")
        pobj_Excel.Range("Anx16AME!T85").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241904", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T85")
         pobj_Excel.Range("Anx16AME!T86").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241905", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T86")
        pobj_Excel.Range("Anx16AME!T87").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241906", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T87")
        pobj_Excel.Range("Anx16AME!T88").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241907", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T88")
        pobj_Excel.Range("Anx16AME!T89").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241910", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T89")
        pobj_Excel.Range("Anx16AME!T90").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241912", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T90")
        pobj_Excel.Range("Anx16AME!T91").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241913", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T91")
        pobj_Excel.Range("Anx16AME!T92").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241914", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T92")
        pobj_Excel.Range("Anx16AME!T93").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252501", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T93")
        pobj_Excel.Range("Anx16AME!T94").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252502", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T94")
        pobj_Excel.Range("Anx16AME!T95").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25250301", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T95")
        pobj_Excel.Range("Anx16AME!T96").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25250401", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T96")
        pobj_Excel.Range("Anx16AME!T97").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252505", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T97")
        pobj_Excel.Range("Anx16AME!T98").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252506", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T98")
        pobj_Excel.Range("Anx16AME!T99").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252509", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T99")
        pobj_Excel.Range("Anx16AME!T100").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252601", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T100")
        pobj_Excel.Range("Anx16AME!T101").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252602", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T101")
        pobj_Excel.Range("Anx16AME!T102").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252702", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T102")
        pobj_Excel.Range("Anx16AME!T103").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2527050205", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T103")
        pobj_Excel.Range("Anx16AME!T104").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2522", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T104")
        pobj_Excel.Range("Anx16AME!T105").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252402", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T105")
        pobj_Excel.Range("Anx16AME!T106").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252701", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T106")
        pobj_Excel.Range("Anx16AME!T107").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25250402", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T107")
        pobj_Excel.Range("Anx16AME!T108").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2527050202", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T108")
        pobj_Excel.Range("Anx16AME!T109").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2527050203", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T109")
        pobj_Excel.Range("Anx16AME!T110").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25241915", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T110")
        pobj_Excel.Range("Anx16AME!T111").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252411", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!T111")
        
        'Cuenta 15 NAGL 20190515
        pobj_Excel.Range("Anx16AME!R73").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("15", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Anx16AME!R73")
    End If
End Sub 'NAGL ERS 079-2016 20170407

Private Sub cargarFondeoAhPFCTSxProducto(ByVal pobj_Excel As Excel.Application, pdFecha As Date, nTipoCambio As Currency)
Dim pcelda As Excel.Range
Dim oDAnx As New DAnexoRiesgos
Dim rs As New ADODB.Recordset
Dim nColumnas As Integer

Set rs = oDAnx.ObtieneFondeoAnx16ANew(pdFecha, nTipoCambio)
nColumnas = 66
If Not (rs.BOF And rs.EOF) Then
    Do While Not rs.EOF
        If rs!cTipoProd = "233" Then
            If rs!cTipo = "FE" Then
                Set pcelda = pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 49)
                pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 49).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 49)
                pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 49).value = Format(rs!nSaldCntME, "#,##0.00")
                
            ElseIf rs!cTipo = "FME" Then
                Set pcelda = pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 50)
                pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 50).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 50)
                pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 50).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FGA" Then
                Set pcelda = pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 51)
                pobj_Excel.Range("Anx16AMN!" & Chr(nColumnas + CInt(rs!cRango)) & 51).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 51)
                pobj_Excel.Range("Anx16AME!" & Chr(nColumnas + CInt(rs!cRango)) & 51).value = Format(rs!nSaldCntME, "#,##0.00")
            End If
        ElseIf rs!cTipoProd = "232" Then
            If rs!cTipo = "FE" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 71)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 71).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 71)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 71).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FME" Then
               Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 72)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 72).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 72)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 72).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FGA" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 73)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 73).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 73)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 73).value = Format(rs!nSaldCntME, "#,##0.00")
            End If
        ElseIf rs!cTipoProd = "234" Then
            If rs!cTipo = "FE" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 82)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 82).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 82)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 82).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FME" Then
               Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 83)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 83).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 83)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 83).value = Format(rs!nSaldCntME, "#,##0.00")
            
            ElseIf rs!cTipo = "FGA" Then
                Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 84)
                pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(rs!cRango)) & 84).value = Format(rs!nSaldCntMN, "#,##0.00")
                Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 84)
                pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(rs!cRango)) & 84).value = Format(rs!nSaldCntME, "#,##0.00")
            End If
        End If
        rs.MoveNext
    Loop
End If
End Sub 'NAGL ERS006-2019 20190514

Private Sub cargarPlazoFijoRangosPersoneriaRangoAnexo6(ByVal pobj_Excel As Excel.Application, psFecha As Date, psFiltro As String, nPromedioEncajeMAMN As Currency, nPromedioEncajeMAME As Currency, nTipoCambio As Currency)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Dim nColumnas As Integer
    Dim oDbalanceCont As DbalanceCont
    Dim nRango As Integer 'ALPA20140211
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Set oDbalanceCont = New DbalanceCont
    Dim ix As Integer
    Dim lnTipoCambioFCMA As Currency
    Dim lnToTalCajaFondosMN, lnToTalCajaFondosME As Currency
    Dim pdFecha As Date
    Dim oEst As New NEstadisticas 'NAGL
    Dim SaldoCajaME As Currency  'NAGL
    Dim oCambio As New nTipoCambio 'NAGL
    'Dim ldFechaAnt As String 'NAGL
    Dim lnTotalBCRPMN As Currency 'NAGL
    Dim lnTotalBCRPME As Currency 'NAGL
    
         'Inicio
        Dim pdFechaFinDeMes As Date
        Dim pdFechaFinDeMesMA As Date
        Dim nSaldoCajaDiarioMesAnteriorME As Currency
        Dim nSaldoCajaDiarioMesAnteriorMN As Currency
        Dim lnToTalOMN As Currency
        Dim lnToTalOME As Currency
        Dim ldFechaPro As Date
        Dim lnToTalTotalCajaFondosMN As Currency
        Dim lnToTalTotalCajaFondosME As Currency
        Dim lnTotalSaldoBCRPAnexoDiarioMN As Currency
        Dim lnTotalSaldoBCRPAnexoDiarioME As Currency
        
        nSaldoCajaDiarioMesAnteriorMN = 0
        pdFechaFinDeMes = DateAdd("d", -Day(DateAdd("m", 1, psFecha)), DateAdd("m", 1, psFecha))
        pdFechaFinDeMesMA = DateAdd("d", -Day(pdFechaFinDeMes), pdFechaFinDeMes)
    
        ldFechaPro = DateAdd("d", -Day(psFecha), psFecha)
        ldFechaPro = DateAdd("d", -Day(ldFechaPro), ldFechaPro)
        nSaldoCajaDiarioMesAnteriorME = 0
        
         lnTipoCambioFCMA = 0
        For ix = 1 To Day(pdFechaFinDeMesMA)
            ldFechaPro = DateAdd("d", 1, ldFechaPro)
            nSaldoCajaDiarioMesAnteriorMN = nSaldoCajaDiarioMesAnteriorMN + (oDbalanceCont.SaldoCtas(32, "761201", ldFechaPro, pdFechaFinDeMesMA, lnTipoCambioFCMA, lnTipoCambioFCMA) / Day(pdFechaFinDeMesMA))
        Next ix
        
        lnToTalCajaFondosMN = 0
        lnToTalCajaFondosME = 0
        lnTotalSaldoBCRPAnexoDiarioMN = 0
        lnTotalSaldoBCRPAnexoDiarioME = 0
        lnToTalTotalCajaFondosMN = 0
        lnToTalTotalCajaFondosME = 0
        lnTotalBCRPMN = 0 '***NAGL
        lnTotalBCRPME = 0 '****NAGL
        lnToTalOMN = 0
        lnToTalOME = 0
        
        ldFechaPro = DateAdd("d", -Day(psFecha), psFecha) '***NAGL
        'ldFechaAnt = ldFechaPro
        
       'CAJA  - DEPOSITOS EN EL BCRP
        For ix = 1 To Day(psFecha)
            ldFechaPro = DateAdd("d", 1, ldFechaPro)
            If Month(ldFechaPro) = Month(DateAdd("d", 1, ldFechaPro)) Then
                   lnTipoCambioFCMA = Format(oCambio.EmiteTipoCambio(ldFechaPro, TCFijoDia), "#,##0.0000")
            Else
                   lnTipoCambioFCMA = Format(oCambio.EmiteTipoCambio(DateAdd("d", 1, ldFechaPro), TCFijoDia), "#,##0.0000")
            End If
            
            'CAJA ANTERIOR MN
            'SaldoCajAnt = Round(oEst.GetCajaAnterior(ldFechaAnt, "761201", "32"), 2)
            lnToTalTotalCajaFondosMN = lnToTalTotalCajaFondosMN + Round(nSaldoCajaDiarioMesAnteriorMN, 2)
            
            'CAJA ME
            'SaldoCajaME = oDbalanceCont.SaldoCajasObligExoneradas(Format(ldFechaPro, "yyyymmdd"), 2)
            SaldoCajaME = oDbalanceCont.ObtenerCtaContSaldoDiario("1121", ldFechaPro) + oDbalanceCont.ObtenerCtaContSaldoDiario("112701", ldFechaPro)
            lnToTalTotalCajaFondosME = lnToTalTotalCajaFondosME + SaldoCajaME
           
            'DEPOSITOS BCRP
            lnTotalSaldoBCRPAnexoDiarioMN = oDbalanceCont.SaldoBCRPAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 1)
            lnTotalBCRPMN = lnTotalBCRPMN + lnTotalSaldoBCRPAnexoDiarioMN
            
            lnTotalSaldoBCRPAnexoDiarioME = oDbalanceCont.SaldoBCRPAnexoDiario(Format(ldFechaPro, "yyyymmdd"), 2)
            lnTotalBCRPME = lnTotalBCRPME + lnTotalSaldoBCRPAnexoDiarioME
            
        Next ix
        
        lnToTalOMN = lnToTalTotalCajaFondosMN + lnTotalBCRPMN 'lnTotalSaldoBCRPAnexoDiarioMN
        lnToTalOME = lnToTalTotalCajaFondosME + lnTotalBCRPME 'lnTotalSaldoBCRPAnexoDiarioME
        'Fin

        'SOLES*********************************************************************
        Set pcelda = pobj_Excel.Range("Disponible!B52")
        'pobj_Excel.Range("Disponible!B51").value = lnToTalOMN - nPromedioEncajeMAMN  '*****NAGL ERS 079-2016 20170407 CAMBIO AL REPORTE BCRP1
        pobj_Excel.Range("Disponible!B52").value = Round(nPromedioEncajeMAMN / Day(psFecha), 2) '***NAGL ERS 079-2016 20170407
        
        Set pcelda = pobj_Excel.Range("Disponible!B51") 'Superavit
        'pobj_Excel.Range("Disponible!B50").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("111201", psFecha, "1", 0)
        pobj_Excel.Range("Disponible!B51").value = Round((lnToTalOMN - nPromedioEncajeMAMN) / Day(psFecha), 2)
        pobj_Excel.Range("Disponible!B51:Disponible!B52").NumberFormat = "#,##0.00;-#,##0.00"
        
        Set pcelda = pobj_Excel.Range("Disponible!B55")
        pobj_Excel.Range("Disponible!B55").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1117", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Disponible!B56")
        pobj_Excel.Range("Disponible!B56").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("1113010_0[12]", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Disponible!B57")
        pobj_Excel.Range("Disponible!B57").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("111201", psFecha, "1", 0)
        
        Set pcelda = pobj_Excel.Range("Disponible!B58")
        pobj_Excel.Range("Disponible!B58").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1111", psFecha, "1", 0)
        
        Set pcelda = pobj_Excel.Range("Disponible!B59")
        pobj_Excel.Range("Disponible!B59").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1115", psFecha, "1", 0) 'NAGL 20170621
        
        Set pcelda = pobj_Excel.Range("Disponible!B60")
        pobj_Excel.Range("Disponible!B60").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1116", psFecha, "1", 0)
        
        Set pcelda = pobj_Excel.Range("Disponible!B61")
        pobj_Excel.Range("Disponible!B61").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1118", psFecha, "1", 0)
        
        pobj_Excel.Range("Disponible!B55:Disponible!B61").NumberFormat = "#,##0.00;-#,##0.00"
        'FIN SOLES
        
        'DOLARES*********************************************************************
        
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B52") 'Encaje Exigible ME
        pobj_Excel.Range("DisponibleDolares!B52").value = Round(nPromedioEncajeMAME / Day(psFecha), 2) '***NAGL ERS 079-2016 20170407
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B51") 'Superavit
        pobj_Excel.Range("DisponibleDolares!B51").value = Round((lnToTalOME - nPromedioEncajeMAME) / Day(psFecha), 2) 'NAGL ERS 079-2016 20170407
        
        'Set pcelda = pobj_Excel.Range("DisponibleDolares!B50")
        'pobj_Excel.Range("DisponibleDolares!B50").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("112201", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        pobj_Excel.Range("DisponibleDolares!B51:DisponibleDolares!B52").NumberFormat = "#,##0.00;-#,##0.00"
        
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B55")
        pobj_Excel.Range("DisponibleDolares!B55").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1127", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B56")
        pobj_Excel.Range("DisponibleDolares!B56").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("1123010_0[12]", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B57")
        pobj_Excel.Range("DisponibleDolares!B57").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioxLike("112201", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B58")
        pobj_Excel.Range("DisponibleDolares!B58").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1121", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B59")
        pobj_Excel.Range("DisponibleDolares!B59").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1125", psFecha, "2", nTipoCambio) / nTipoCambio, 2) 'NAGL 20170621
        
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B60")
        pobj_Excel.Range("DisponibleDolares!B60").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1126", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        
        Set pcelda = pobj_Excel.Range("DisponibleDolares!B61")
        pobj_Excel.Range("DisponibleDolares!B61").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1128", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        
        pobj_Excel.Range("DisponibleDolares!B55:DisponibleDolares!B61").NumberFormat = "#,##0.00;-#,##0.00"
        'FIN DOLARES
        
        Set prs = oCtaIf.GetPlazoFijoRangosPersoneriaRangoAnexo6(Format(psFecha, "yyyy/mm/dd"), psFiltro)
        nColumnas = 66
        If Not prs.EOF Or prs.BOF Then
            Do While Not prs.EOF
                If prs!cTipo = 1 Then
                    Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 38)
                    pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value + prs!SALDO_MN
                    
                    Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 38)
                    pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value + prs!SALDO_ME
                End If
                If prs!cTipo = 2 Then
                    Set pcelda = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 38)
                    pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value = pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value - prs!SALDO_MN
                    pobj_Excel.Range("Disponible!" & Chr(nColumnas + CInt(prs!cRango)) & 39).value = prs!SALDO_MN
                    
                    Set pcelda = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 38)
                    pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value = pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 38).value - prs!SALDO_ME
                    pobj_Excel.Range("DisponibleDolares!" & Chr(nColumnas + CInt(prs!cRango)) & 39).value = prs!SALDO_ME
                End If
                prs.MoveNext
            Loop
        End If
        Set prs = Nothing
        
        '****NAGL 20190515 ERS006-2019 NUEVO FORMATO
        Set pcelda = pobj_Excel.Range("Creditos!C51")
        pobj_Excel.Range("Creditos!C51").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251703", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Creditos!C52")
        pobj_Excel.Range("Creditos!C52").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251704", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Creditos!C53")
        pobj_Excel.Range("Creditos!C53").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251705", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Creditos!C54")
        pobj_Excel.Range("Creditos!C54").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2116", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Creditos!C55")
        pobj_Excel.Range("Creditos!C55").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2312", psFecha, "1", 0)
        Set pcelda = pobj_Excel.Range("Creditos!C56")
        pobj_Excel.Range("Creditos!C56").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2313", psFecha, "1", 0)
        
        Set pcelda = pobj_Excel.Range("Creditos!E51")
        pobj_Excel.Range("Creditos!E51").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252703", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Creditos!E52")
        pobj_Excel.Range("Creditos!E52").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252704", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Creditos!E53")
        pobj_Excel.Range("Creditos!E53").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252705", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Creditos!E54")
        pobj_Excel.Range("Creditos!E54").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2126", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Creditos!E55")
        pobj_Excel.Range("Creditos!E55").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2322", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        Set pcelda = pobj_Excel.Range("Creditos!E56")
        pobj_Excel.Range("Creditos!E56").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("2323", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        
        pobj_Excel.Range("Creditos!C51:Creditos!C56").NumberFormat = "#,##0.00;-#,##0.00"
        pobj_Excel.Range("Creditos!E51:Creditos!E56").NumberFormat = "#,##0.00;-#,##0.00"
        '***************************************************
        
        'Comentado by NAGL 20190515, Proveniente del Cálculo de Balance,(En su Moneda Original)
        '20141028********************Errores detectados por KARU
        'Set pcelda = pobj_Excel.Range("Creditos!H63")
        'pobj_Excel.Range("Creditos!H63").value = IIf(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170301", psFecha, "1", 0) < 0, 0, oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170301", psFecha, "1", 0))
        'Set pcelda = pobj_Excel.Range("Creditos!H64")
        'pobj_Excel.Range("Creditos!H64").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170302", psFecha, "1", 0)
        'Set pcelda = pobj_Excel.Range("Creditos!H65")
        'pobj_Excel.Range("Creditos!H65").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25170303", psFecha, "1", 0)
        'Set pcelda = pobj_Excel.Range("Creditos!H66")
        'pobj_Excel.Range("Creditos!H66").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251704", psFecha, "1", 0)
        'Set pcelda = pobj_Excel.Range("Creditos!H67")
        'pobj_Excel.Range("Creditos!H67").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("251705", psFecha, "1", 0)
        
        'Set pcelda = pobj_Excel.Range("Creditos!J63")
        'pobj_Excel.Range("Creditos!J63").value = IIf(Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270301", psFecha, "2", nTipoCambio) / nTipoCambio, 2) < 0, 0, Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270301", psFecha, "2", nTipoCambio) / nTipoCambio, 2))
        'Set pcelda = pobj_Excel.Range("Creditos!J64")
        'pobj_Excel.Range("Creditos!J64").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270302", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        'Set pcelda = pobj_Excel.Range("Creditos!J65")
        'pobj_Excel.Range("Creditos!J65").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("25270303", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        'Set pcelda = pobj_Excel.Range("Creditos!J66")
        'pobj_Excel.Range("Creditos!J66").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252704", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        'Set pcelda = pobj_Excel.Range("Creditos!J67")
        'pobj_Excel.Range("Creditos!J67").value = Round(oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("252705", psFecha, "2", nTipoCambio) / nTipoCambio, 2)
        
        Set pcelda = pobj_Excel.Range("Creditos!A54")
        If (Month(psFecha) >= 1 And Month(psFecha) <= 4) Then
            nRango = (5 - CInt(Month(psFecha)))
        ElseIf Month(psFecha) = 11 Then
            nRango = 6
        ElseIf Month(psFecha) = 12 Then
            nRango = 5
        ElseIf (Month(psFecha) >= 5 And Month(psFecha) <= 10) Then
            nRango = (11 - CInt(Month(psFecha)))
        End If 'NAGL 20170519
        '1M  2M  3M  4M  5M  6M  7-9 M   10-12 M

        pobj_Excel.Range("Creditos!A54").value = nRango
        
        'pobj_Excel.Range("Creditos!E64:Creditos!E67").NumberFormat = "#,##0.00;-#,##0.00" 'Comentado by NAGL 20190515
        'pobj_Excel.Range("Creditos!B104:Creditos!B104").NumberFormat = "#,##0.00;-#,##0.00" 'Comentado by NAGL 20190515
        
        pdFechaFinDeMesMA = DateAdd("d", -Day(psFecha), psFecha)
        Set pcelda = pobj_Excel.Range("Creditos!I36")
        pobj_Excel.Range("Creditos!I36").value = pdFechaFinDeMesMA

        Set pcelda = pobj_Excel.Range("Creditos!I38")
        pobj_Excel.Range("Creditos!I38").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1404", pdFechaFinDeMesMA, "0", 0)
        Set pcelda = pobj_Excel.Range("Creditos!I39")
        pobj_Excel.Range("Creditos!I39").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1405", pdFechaFinDeMesMA, "0", 0)
        Set pcelda = pobj_Excel.Range("Creditos!I40")
        pobj_Excel.Range("Creditos!I40").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1406", pdFechaFinDeMesMA, "0", 0)
        
        'NAGL 20190516********
        Set pcelda = pobj_Excel.Range("Creditos!I42")
        pobj_Excel.Range("Creditos!I42").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("14", pdFechaFinDeMesMA, "0", 0)
        Set pcelda = pobj_Excel.Range("Creditos!I43")
        pobj_Excel.Range("Creditos!I43").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1408", pdFechaFinDeMesMA, "0", 0)
        Set pcelda = pobj_Excel.Range("Creditos!I44")
        pobj_Excel.Range("Creditos!I44").value = oDbalanceCont.ObtenerCtaContSaldoBalanceDiarioFindeMes("1409", pdFechaFinDeMesMA, "0", 0)
        pobj_Excel.Range("Creditos!I38:Creditos!I44").NumberFormat = "#,##0.00;-#,##0.00"
        '***********************
End Sub
'EJVG20131230 ***
Private Function SaldoCajasObligExoneradas(ByVal pdFecha As String, ByVal pnMoneda As Moneda) As Currency
    Dim oNCaja As New NCajaCtaIF
    Dim oEnc As New NEncajeBCR
    Dim rsEncDiario As New ADODB.Recordset
    
    Set rsEncDiario = oEnc.ObtenerParamEncajeDiarioxCod("04")
    SaldoCajasObligExoneradas = ObtieneEfectivoDiaxRptConcentracionGastos(pdFecha, pnMoneda) + IIf(pnMoneda = gMonedaNacional, rsEncDiario!nValor, 0)
    
    Set rsEncDiario = Nothing
    Set oEnc = Nothing
    Set oNCaja = Nothing
End Function
Private Function SaldoEfectivoDisponible(ByVal pnMoneda As Moneda) As Currency
    Dim MatIFi() As TConcentraFondos
    Dim i As Integer
    Dim lnMonto As Currency
    
    If pnMoneda = gMonedaNacional Then
        MatIFi = MatCtasBcosMN
    Else
        MatIFi = MatCtasBcosME
    End If
    For i = 0 To UBound(MatIFi) - 1
        If MatIFi(i).CodPersona <> "1090100822183" Then
            lnMonto = lnMonto + MatIFi(i).SaldoCtaAhorro + MatIFi(i).SaldoCtaCorriente
        End If
    Next
    SaldoEfectivoDisponible = lnMonto
End Function
'END EJVG *******
'FRHU 20131209 RQ13650
Public Sub ReporteAnexo02PropuestaCapitalizacionUtilidades(ByVal pnAnio As Integer, ByVal pnSemestre As Integer, ByVal pnTipo As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lnTipo As Integer
    lnAnio = pnAnio
    lnSemestre = pnSemestre
    lnTipo = pnTipo
    
    If ValidarParametroUtilidad(lnAnio) = False Then
        MsgBox "Primero debe registrar los parametros de configuracion del Anexo 01, referente al año " & lnAnio
        Exit Sub
    End If
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Anexo02PropuestaCapitalizacionUtilidades"
    'Primera Hoja
    lsNomHoja = "Anexo02"
    '************
    lsArchivo1 = "\spooler\RepAnexo02ProCapiUtilidad" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    Set rsCapitali = oDbalanceCont.recuperarReporteAnexo02(CStr(lnAnio), lnSemestre, lnTipo)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B4", "L4").HorizontalAlignment = xlCenter
    xlHoja1.Range("B5", "L5").HorizontalAlignment = xlCenter
    xlHoja1.Range("L10", "L10").HorizontalAlignment = xlRight
    xlHoja1.Range("L17", "L17").HorizontalAlignment = xlRight
    xlHoja1.Range("B4", "L4").MergeCells = True
    xlHoja1.Range("B5", "L5").MergeCells = True
    If lnTipo = 1 Then
    xlHoja1.Cells(4, 2) = "PROPUESTA DE CAPITALIZACIÓN DE UTILIDADES " & CStr(lnAnio) & " (EN " & StrConv(gcPEN_PLURAL, vbUpperCase) & ")" 'marg ers044-2016
    Else
    xlHoja1.Cells(4, 2) = "PROPUESTA DE CAPITALIZACIÓN DE UTILIDADES " & CStr(lnAnio) & " (EN MILES DE " & StrConv(gcPEN_PLURAL, vbUpperCase) & ")" 'marg ers044-2016
    End If
    xlHoja1.Cells(5, 2) = "Y DE LA RESERVA ESPECIAL " & CStr(lnAnio - 1)
    
    xlHoja1.Range("B4", "L4").Font.Bold = True
    xlHoja1.Range("B4", "L4").Font.Size = 11
    xlHoja1.Range("B5", "L5").Font.Size = 11
    'xlHoja1.Cells(2, 7) = Format(txtfecha.Text, "DD") & " DE " & UCase(Format(txtfecha.Text, "MMMM")) & " DEL  " & Format(txtfecha.Text, "YYYY")
    If nPase = 1 Then
        If Not rsCapitali.BOF And Not rsCapitali.EOF Then
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 15)).Borders.LineStyle = 1
                xlHoja1.Cells(10, 12) = gcPEN_SIMBOLO & " " & Format(rsCapitali!NuevCapiSociPaga, "#,###") 'marg ers044-2016
                xlHoja1.Cells(12, 12) = Format(rsCapitali!CapiPagaAnte, "#,###")
                xlHoja1.Cells(13, 12) = Format(rsCapitali!CapiReseEspe, "#,###")
                xlHoja1.Cells(14, 12) = Format(rsCapitali!UtiComproCapi, "#,###")
                xlHoja1.Cells(15, 12) = Format(rsCapitali!UtiDividenMPM, "#,###")
                xlHoja1.Cells(17, 12) = gcPEN_SIMBOLO & " " & Format(rsCapitali!NuevFondoReserva, "#,###")  'marg ers044-2016
                xlHoja1.Cells(19, 12) = Format(rsCapitali!FondReseLegaAnte, "#,###")
                xlHoja1.Cells(20, 12) = Format(rsCapitali!ReservaLegal, "#,###")
                xlHoja1.Cells(21, 12) = Format(rsCapitali!ReseLegalEspe, "#,###")
                
                xlHoja1.Cells(12, 3) = "a) Capital Pagado Anterior (Capital al " & Format(rsCapitali!Fecha, "dd/mm/yyyy") & ")"
                xlHoja1.Cells(13, 3) = "b) Capitalización Reserva Especial Año " & CStr(lnAnio - 1) & " (Reserva Especial Art. 4° (D.S. N° 157-90-EF)) "
                xlHoja1.Cells(14, 4) = rsCapitali!nValorCapi & "% de la Utilidad de Libre Disposición, Resolución de Alcaldia N° 153-2003-A-MPM (Abril 2003)"
                xlHoja1.Cells(15, 4) = rsCapitali!nValorMuni & "% de la Utilidad de Libre Disposición, Resolucion SBS"
                
                xlHoja1.Cells(19, 3) = "a) Fondo de Reserva Legal Anterior (Art. 67° Ley de Banco 26702) " & Format(CDate(rsCapitali!Fecha), "dd/mm/yyyy")
                xlHoja1.Cells(20, 3) = "b) Incremento Fondo de Reserva Legal (" & rsCapitali!nValorNeta & "%) por Aplicación de Utilidades Año " & CStr(lnAnio)
                xlHoja1.Cells(21, 3) = "c) Reserva Especial (" & rsCapitali!nValorReal & "% Art.4° DS N° 157-90-EF) - Año " & CStr(lnAnio)
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 14), xlHoja1.Cells(nSaltoContador, 15)).NumberFormat = "dd/mm/yyyy"
'                xlHoja1.Cells(15, 13) = rsUtilidad!nValorNeta & "%"
'                xlHoja1.Cells(19, 13) = rsUtilidad!nValorReal & "%"
'                xlHoja1.Cells(23, 13) = rsUtilidad!nValorCapi & "%"
'                xlHoja1.Cells(24, 13) = rsUtilidad!nValorMuni & "%"
        End If
    End If
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
'
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
Private Function ValidarParametroUtilidad(ByVal pnAnio As Integer) As Boolean
Dim i As Integer
Dim bValidar As Boolean
Dim panAnio As Integer
panAnio = pnAnio
Dim oDbalanceCont As DbalanceCont
Set oDbalanceCont = New DbalanceCont
Dim rsUtilidad As ADODB.Recordset
Set rsUtilidad = New ADODB.Recordset
Set rsUtilidad = oDbalanceCont.recuperarConfParametroUtilidad(panAnio)

If Not rsUtilidad.BOF And Not rsUtilidad.EOF Then
    i = 1
    Do While Not rsUtilidad.EOF
'        Me.FEParametroUtilidad.AdicionaFila
'        Me.FEParametroUtilidad.TextMatrix(i, 1) = rsUtilidad!cParamUtilidad
'        Me.FEParametroUtilidad.TextMatrix(i, 2) = rsUtilidad!nValor
'        Me.FEParametroUtilidad.TextMatrix(i, 3) = rsUtilidad!nAnio
         If rsUtilidad!nAnio = 0 Then
         'MsgBox "Primero debe registrar los parametros de configuracion del Anexo 01"
         ValidarParametroUtilidad = False
         Exit Function
         End If
'        Me.FEParametroUtilidad.TextMatrix(i, 4) = rsUtilidad!nParamVar
'        rsUtilidad.MoveNext
'        i = i + 1
         ValidarParametroUtilidad = True
         Exit Function
    Loop
End If

End Function
'FIN FRHU 20131209
'FRHU 20131219 RQ13656 - Estado de Cambio en el Patrimonio
Public Sub ReporteEstadoCambioPatrimonio(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lnSe As Integer
    Dim lcSemestre As String
    Dim lnTipo As Integer
    Dim i As Integer
    
    lnAnio = pnAnio
    lnSe = pnSemestre
    
    If pnSemestre = 1 Then
    lcSemestre = "Junio"
    Else
    lcSemestre = "Diciembre"
    End If
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "EstadosDeCambiosEnElPatrimonio"
    'Primera Hoja
    lsNomHoja = "Hoja1"
    '************
    lsArchivo1 = "\spooler\EstadosDeCambiosEnElPatrimonio" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    Set rsCapitali = oDbalanceCont.recuperarEstadoDeCambioPatrimonio(CStr(lnAnio), lnSe)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    i = 6
    
    If nPase = 1 Then
        Do While Not rsCapitali.EOF
                i = i + 1
                xlHoja1.Cells(i, 3) = rsCapitali!cNomAF
                xlHoja1.Cells(i, 4) = Format(rsCapitali!nCapitalSocial, "#,###.00")
                xlHoja1.Cells(i, 5) = Format(rsCapitali!nCapitalAdicio, "#,###.00")
                xlHoja1.Cells(i, 6) = Format(rsCapitali!nReservaObliga, "#,###.00")
                xlHoja1.Cells(i, 7) = Format(rsCapitali!nReservaVolunt, "#,###.00")
                xlHoja1.Cells(i, 8) = Format(rsCapitali!nResultadoAcum, "#,###.00")
                xlHoja1.Cells(i, 9) = Format(rsCapitali!nResultadoEjer, "#,###.00")
                xlHoja1.Cells(i, 10) = Format(rsCapitali!nTotalAjustePa, "#,###.00")
                xlHoja1.Cells(i, 11) = Format(rsCapitali!nTotalPatrimon, "#,###.00")
                
                'xlHoja1.Range("B" & Trim(Str(i)) & ":" & "L" & Trim(Str(i))).Borders.LineStyle = 1
                rsCapitali.MoveNext
                
        Loop
    End If
    
    xlHoja1.Range("C" & Trim(Str(i)) & ":" & "K" & Trim(Str(i))).Font.Bold = True '43
    xlHoja1.Range("C" & Trim(Str(i - 18)) & ":" & "K" & Trim(Str(i - 18))).Font.Bold = True '25 -> 43-18
    xlHoja1.Range("C" & Trim(Str(i - 36)) & ":" & "K" & Trim(Str(i - 36))).Font.Bold = True '7 -> 43- 36
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
'FIN FRHU 20131219

'REPORTE DE ACTIVOS FIJOS

'FRHU 20131211 RQ13651 - Movimiento de la depreciacion acumulada e intangible
Public Sub RepMovDepreAcumuladaIntangible(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lcSe As String
    Dim lcSemestre As String
    Dim lnTipo As Integer
    Dim i As Integer
    
    lnAnio = pnAnio
    lcSe = pnSemestre
    
    If Len(CStr(pnSemestre)) = 1 Then
    lcSe = "0" & CStr(pnSemestre)
    Else
    lcSe = CStr(pnSemestre)
    End If
    
    lcSemestre = Me.cboMes.Text
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AFMovDepreciacionAcumuladaIntangible"
    'Primera Hoja
    lsNomHoja = "Hoja1"
    '************
    lsArchivo1 = "\spooler\AFMovDepreciacionAcumuladaIntangible" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    Set rsCapitali = oDbalanceCont.RepAFDepreciacionAcumuladaIntangible(CStr(lnAnio), lcSe)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B5", "J5").MergeCells = True
    xlHoja1.Range("B5", "J5").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5", "J5").Font.Bold = True
    xlHoja1.Cells(5, 2) = "MOVIMIENTO DE LA DEPRECIACION ACUMULADA E INTANGIBLES " & UCase(lcSemestre) & " AÑO " & CStr(lnAnio) & " en " & StrConv(gcPEN_PLURAL, vbProperCase) & " (" & gcPEN_SIMBOLO & ") 1/ " 'marg ers044-2016
    
    i = 10
    
    If nPase = 1 Then
        Do While Not rsCapitali.EOF
                i = i + 1
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
                'xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 15)).Borders.LineStyle = 1
                xlHoja1.Cells(i, 2) = rsCapitali!cNomAF
                xlHoja1.Cells(i, 3) = rsCapitali!cOrden
                xlHoja1.Cells(i, 4) = Format(rsCapitali!nSaldoIni, "#,###")
                xlHoja1.Cells(i, 5) = Format(rsCapitali!nDepre, "#,###")
                xlHoja1.Cells(i, 6) = Format(rsCapitali!OtrasAdiciones, "#,###")
                xlHoja1.Cells(i, 7) = Format(rsCapitali!IncreRevaVolu, "#,###")
                xlHoja1.Cells(i, 8) = Format(rsCapitali!Ventas, "#,###")
                xlHoja1.Cells(i, 9) = Format(rsCapitali!nRetiro, "#,###")
                xlHoja1.Cells(i, 10) = Format(rsCapitali!nSaldoFin, "#,###")
                
                xlHoja1.Range("B" & Trim(Str(i)) & ":" & "J" & Trim(Str(i))).Borders.LineStyle = 1
                rsCapitali.MoveNext
                
        Loop
    End If
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
'FIN FRHU 20131211

'FRHU 20131214 RQ13652 - Movimiento de Activo Fijo e Intangible
Public Sub RepMovActivoFijoIntangible(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lcSe As String
    Dim lcSemestre As String
    Dim lnTipo As Integer
    Dim i As Integer
    
    lnAnio = pnAnio
    lcSe = pnSemestre
    
    If Len(CStr(pnSemestre)) = 1 Then
    lcSe = "0" & CStr(pnSemestre)
    Else
    lcSe = CStr(pnSemestre)
    End If
    
    lcSemestre = Me.cboMes.Text
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AFMovActivoFijoIntangible"
    'Primera Hoja
    lsNomHoja = "Hoja1"
    '************
    lsArchivo1 = "\spooler\AFMovActivoFijoIntangible" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    Set rsCapitali = oDbalanceCont.RepAFActivoFijoIntangible(CStr(lnAnio), lcSe)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B5", "N5").MergeCells = True
    xlHoja1.Range("B5", "N5").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5", "N5").Font.Bold = True
    xlHoja1.Cells(5, 2) = "MOVIMIENTO DEL ACTIVO FIJO E INTANGIBLES " & UCase(lcSemestre) & " AÑO " & CStr(lnAnio) & " en " & StrConv(gcPEN_PLURAL, vbProperCase) & " (" & gcPEN_SIMBOLO & ") 1/ " 'marg ers044-2016
    
    i = 10
    
    If nPase = 1 Then
        Do While Not rsCapitali.EOF
                i = i + 1
                xlHoja1.Cells(i, 4) = rsCapitali!cOrden
                xlHoja1.Cells(i, 5) = Format(rsCapitali!nSaldoIni, "#,###")
                xlHoja1.Cells(i, 8) = Format(rsCapitali!nCompra, "#,###")
                xlHoja1.Cells(i, 12) = Format(rsCapitali!nRetiro, "#,###")
                xlHoja1.Cells(i, 14) = Format(rsCapitali!nSaldoFin, "#,###")
                'xlHoja1.Range("B" & Trim(Str(i)) & ":" & "J" & Trim(Str(i))).Borders.LineStyle = 1
                rsCapitali.MoveNext
                
        Loop
    End If
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
'FIN FRHU 20131214 RQ13652

'FRHU 20131216 RQ13653 - Control de Intangibles y otros Activos Amortizables
Public Sub RepCtoIntangibleOtrosActivos(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lnSe As Integer
    Dim lcSemestre As String
    Dim lnTipo As Integer
    Dim i As Integer
    
    lnAnio = pnAnio
    lnSe = pnSemestre
    
    If pnSemestre = 1 Then
    lcSemestre = "Junio"
    Else
    lcSemestre = "Diciembre"
    End If
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AFCtoIntangibleOtrosActivosAmortizables"
    'Primera Hoja
    lsNomHoja = "Hoja1"
    '************
    lsArchivo1 = "\spooler\AFCtoIntangibleOtrosActivosAmortizables" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    Set rsCapitali = oDbalanceCont.RepAFControlIntangible(CStr(lnAnio), lnSe)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B2", "L2").MergeCells = True
    xlHoja1.Range("B2", "L2").HorizontalAlignment = xlCenter
    xlHoja1.Range("B2", "L2").Font.Bold = True
    xlHoja1.Cells(2, 2) = "Activos Intangibles " & lcSemestre & "  " & CStr(lnAnio)
    
    i = 3
    
    If nPase = 1 Then
        Do While Not rsCapitali.EOF
                i = i + 1
                xlHoja1.Cells(i, 2) = rsCapitali!cNomAF
                xlHoja1.Cells(i, 3) = Format(rsCapitali!nSaldoIni, "#,###.00")
                xlHoja1.Cells(i, 4) = Format(rsCapitali!nAdicionCompra, "#,###.00")
                xlHoja1.Cells(i, 5) = Format(rsCapitali!nTransferencia, "#,###.00")
                xlHoja1.Cells(i, 6) = Format(rsCapitali!nRetiro, "#,###.00")
                xlHoja1.Cells(i, 7) = Format(rsCapitali!nReclacifica, "#,###.00")
                xlHoja1.Cells(i, 8) = Format(rsCapitali!nDesvaloriza, "#,###.00")
                xlHoja1.Cells(i, 9) = Format(rsCapitali!nAjustes, "#,###.00")
                xlHoja1.Cells(i, 10) = Format(rsCapitali!nSaldoFin, "#,###.00")
                xlHoja1.Cells(i, 11) = Format(rsCapitali!nAmortiza, "#,###.00")
                xlHoja1.Cells(i, 12) = Format(rsCapitali!nValorNeto, "#,###.00")
                xlHoja1.Range("B" & Trim(Str(i)) & ":" & "L" & Trim(Str(i))).Borders.LineStyle = 1
                rsCapitali.MoveNext
                
        Loop
    End If
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
'FRHU 20131216 RQ13653

'FRHU 20131217 RQ13654 - Control de Inmuebles, Maquinaria y Equipos
Public Sub RepCtoInmuebleMaquinariayEquipos(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lnSe As Integer
    Dim lcSemestre As String
    Dim lnTipo As Integer
    Dim i As Integer
    
    lnAnio = pnAnio
    lnSe = pnSemestre
    
    If pnSemestre = 1 Then
    lcSemestre = "Junio"
    Else
    lcSemestre = "Diciembre"
    End If
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AFCtoInmueblesMaquinariayEquipo"
    'Primera Hoja
    lsNomHoja = "Hoja1"
    '************
    lsArchivo1 = "\spooler\AFCtoInmueblesMaquinariayEquipo" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    Set rsCapitali = oDbalanceCont.RepAFControlInmuebleMaquinaria(CStr(lnAnio), lnSe)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B3", "M3").MergeCells = True
    xlHoja1.Range("B3", "M3").HorizontalAlignment = xlCenter
    xlHoja1.Range("B3", "M3").Font.Bold = True
    xlHoja1.Cells(3, 2) = "Activos Fijos " & lcSemestre & "  " & CStr(lnAnio)
    
    i = 4
    
    If nPase = 1 Then
        Do While Not rsCapitali.EOF
                i = i + 1
                xlHoja1.Cells(i, 2) = rsCapitali!cNomAF
                xlHoja1.Cells(i, 3) = Format(rsCapitali!nSaldoIni, "#,###.00")
                xlHoja1.Cells(i, 4) = Format(rsCapitali!nAdicionCosto, "#,###.00")
                xlHoja1.Cells(i, 5) = Format(rsCapitali!nTransferencia, "#,###.00")
                xlHoja1.Cells(i, 6) = Format(rsCapitali!nRetiro, "#,###.00")
                xlHoja1.Cells(i, 7) = Format(rsCapitali!nRevaluacion, "#,###.00")
                xlHoja1.Cells(i, 8) = Format(rsCapitali!nDesvaloriza, "#,###.00")
                xlHoja1.Cells(i, 9) = Format(rsCapitali!nReclasifica, "#,###.00")
                xlHoja1.Cells(i, 10) = Format(rsCapitali!nAjuste, "#,###.00")
                xlHoja1.Cells(i, 11) = Format(rsCapitali!nSaldoFin, "#,###.00")
                xlHoja1.Cells(i, 12) = Format(rsCapitali!nDepreciacion, "#,###.00")
                xlHoja1.Cells(i, 13) = Format(rsCapitali!nValorNeto, "#,###.00")
                xlHoja1.Range("B" & Trim(Str(i)) & ":" & "M" & Trim(Str(i))).Borders.LineStyle = 1
                rsCapitali.MoveNext
                
        Loop
    End If
    
    xlHoja1.Range("B" & Trim(Str(i)) & ":" & "M" & Trim(Str(i))).Font.Bold = True
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
'FRHU 20131217 RQ13654

'FRHU 20131218 RQ13655 - Control de Depreciacion de Activos Fijos
Public Sub RepCtoDepreciacionActivosFijos(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim lsArchivo As String
    Dim rsCapitali As ADODB.Recordset
    Dim oDbalanceCont As New DbalanceCont
    Dim nPase As Integer
    Dim lbExisteHoja As Boolean
    
    Dim lnAnio As Integer
    Dim lnSemestre As Integer
    Dim lnSe As Integer
    Dim lcSemestre As String
    Dim lnTipo As Integer
    Dim i As Integer
    
    lnAnio = pnAnio
    lnSe = pnSemestre
    
    If pnSemestre = 1 Then
    lcSemestre = "Junio"
    Else
    lcSemestre = "Diciembre"
    End If
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "AFCtoDepreciacionActivosFijos"
    'Primera Hoja
    lsNomHoja = "Hoja1"
    '************
    lsArchivo1 = "\spooler\AFCtoDepreciacionActivosFijos" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
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
    
    Set rsCapitali = oDbalanceCont.RepAFControlDepreciacion(CStr(lnAnio), lnSe)
    nPase = 1
    If (rsCapitali Is Nothing) Then
        nPase = 0
    End If
    
    xlHoja1.Range("B3", "L3").MergeCells = True
    xlHoja1.Range("B3", "L3").HorizontalAlignment = xlCenter
    xlHoja1.Range("B3", "L3").Font.Bold = True
    xlHoja1.Cells(3, 2) = "Depreciacón Activos Fijos " & lcSemestre & "  " & CStr(lnAnio)
    
    i = 4
    
    If nPase = 1 Then
        Do While Not rsCapitali.EOF
                i = i + 1
                xlHoja1.Cells(i, 2) = rsCapitali!cNomAF
                xlHoja1.Cells(i, 3) = Format(rsCapitali!nSaldoIni, "#,###.00")
                xlHoja1.Cells(i, 4) = Format(rsCapitali!nCargoRes, "#,###.00")
                xlHoja1.Cells(i, 5) = Format(rsCapitali!nCargoCos, "#,###.00")
                xlHoja1.Cells(i, 6) = Format(rsCapitali!nTransferencia, "#,###.00")
                xlHoja1.Cells(i, 7) = Format(rsCapitali!nRetiro, "#,###.00")
                xlHoja1.Cells(i, 8) = Format(rsCapitali!nRevaluacion, "#,###.00")
                xlHoja1.Cells(i, 9) = Format(rsCapitali!nDesvalorizacion, "#,###.00")
                xlHoja1.Cells(i, 10) = Format(rsCapitali!nReclasificacion, "#,###.00")
                xlHoja1.Cells(i, 11) = Format(rsCapitali!nAjuste, "#,###.00")
                xlHoja1.Cells(i, 12) = Format(rsCapitali!nSaldoFin, "#,###.00")
                xlHoja1.Range("B" & Trim(Str(i)) & ":" & "L" & Trim(Str(i))).Borders.LineStyle = 1
                rsCapitali.MoveNext
                
        Loop
    End If
    
    xlHoja1.Range("B" & Trim(Str(i)) & ":" & "M" & Trim(Str(i))).Font.Bold = True
    
    Set oDbalanceCont = Nothing
    If nPase = 1 Then
        rsCapitali.Close
    End If
    Set rsCapitali = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

    Exit Sub

End Sub
'FRHU 20131218 RQ13655
'ALPA20140118
Private Sub PintaFondeo(ByRef xlHoja1 As Excel.Worksheet, ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal pnMoneda As Integer)
    Dim loRs As ADODB.Recordset
    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Set loRs = New ADODB.Recordset
    
Set loRs = oDbalanceCont.ObtenerFondeoEncajexTramosxProducto(pdFecha, pnTipoCambio, "233")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(49, loRs!cRango + 2) = CCur(xlHoja1.Cells(49, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
                    'xlHoja1.Cells(33, 4) = CCur(xlHoja1.Cells(33, 4)) + loRs!nSaldCntME
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
                    'xlHoja1.Cells(35, 4) = CCur(xlHoja1.Cells(35, 4)) + loRs!nSaldCntME
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
                    'xlHoja1.Cells(37, 4) = CCur(xlHoja1.Cells(37, 4)) + loRs!nSaldCntME
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
                    'xlHoja1.Cells(34, 4) = CCur(xlHoja1.Cells(34, 4)) + loRs!nSaldCntME
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
                    'xlHoja1.Cells(35, 4) = CCur(xlHoja1.Cells(35, 4)) + loRs!nSaldCntME
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
                    'xlHoja1.Cells(37, 4) = CCur(xlHoja1.Cells(37, 4)) + loRs!nSaldCntME
            End If
        End If
        loRs.MoveNext
    Loop
    End If

    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProducto(pnMoneda, pdFecha, pnTipoCambio, "233")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                'lnSubastasMND1y2_C1 = lnSubastasMND1y2_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(49, loRs!cRango + 2) = CCur(xlHoja1.Cells(49, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                'lnSubastasMND3o3_C1 = lnSubastasMND3o3_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                'lnSubastasMND1y2_C0 = lnSubastasMND1y2_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                'lnSubastasMND3o3_C0 = lnSubastasMND3o3_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                'lnSubastasMND4aM_C0 = lnSubastasMND4aM_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
      loRs.MoveNext
    Loop
    End If
    
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto(100, pdFecha, pnMoneda, pnTipoCambio, 1, "233")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                'lnOtrosDepMND1y2_C1 = lnOtrosDepMND1y2_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(49, loRs!cRango + 2) = CCur(xlHoja1.Cells(49, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                'lnOtrosDepMND3o3_C1 = lnOtrosDepMND3o3_C1 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                'lnOtrosDepMND1y2_C0 = lnOtrosDepMND1y2_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                'lnOtrosDepMND3o3_C0 = lnOtrosDepMND3o3_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(50, loRs!cRango + 2) = CCur(xlHoja1.Cells(50, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(51, loRs!cRango + 2) = CCur(xlHoja1.Cells(51, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                'lnOtrosDepMND4aM_C0 = lnOtrosDepMND4aM_C0 + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        loRs.MoveNext
    Loop
    End If
End Sub
Private Sub PintaFondeoAhorro(ByRef xlHoja1 As Excel.Worksheet, ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal pnMoneda As Integer)
      Dim loRs As ADODB.Recordset
    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Set loRs = New ADODB.Recordset

    Set loRs = oDbalanceCont.ObtenerFondeoEncajexTramosxProducto(pdFecha, pnTipoCambio, "232")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(71, loRs!cRango + 2) = CCur(xlHoja1.Cells(71, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
        End If
        loRs.MoveNext
    Loop
    End If

    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProducto(pnMoneda, pdFecha, pnTipoCambio, "232")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(71, loRs!cRango + 2) = CCur(xlHoja1.Cells(71, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                 xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                
            End If
        End If
      loRs.MoveNext
    Loop
    End If
    
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto(100, pdFecha, pnMoneda, pnTipoCambio, 1, "232")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(71, loRs!cRango + 2) = CCur(xlHoja1.Cells(71, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                 xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                 xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                 xlHoja1.Cells(72, loRs!cRango + 2) = CCur(xlHoja1.Cells(72, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(73, loRs!cRango + 2) = CCur(xlHoja1.Cells(73, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                
            End If
        End If
        loRs.MoveNext
    Loop
    End If
End Sub 'NAGL 20170407

Private Sub PintaFondeoCTS(ByRef xlHoja1 As Excel.Worksheet, ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal pnMoneda As Integer)
    Dim loRs As ADODB.Recordset
    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Set loRs = New ADODB.Recordset
  Set loRs = oDbalanceCont.ObtenerFondeoEncajexTramosxProducto(pdFecha, pnTipoCambio, "234")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(82, loRs!cRango + 2) = CCur(xlHoja1.Cells(82, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
        End If
        loRs.MoveNext
    Loop
    End If

    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProducto(pnMoneda, pdFecha, pnTipoCambio, "234")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(82, loRs!cRango + 2) = CCur(xlHoja1.Cells(82, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                 xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                
            End If
        End If
      loRs.MoveNext
    Loop
    End If
    
    Set loRs = oDbalanceCont.ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto(100, pdFecha, pnMoneda, pnTipoCambio, 1, "234")
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                xlHoja1.Cells(82, loRs!cRango + 2) = CCur(xlHoja1.Cells(82, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                 xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                 xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "3" Then
                 xlHoja1.Cells(83, loRs!cRango + 2) = CCur(xlHoja1.Cells(83, loRs!cRango + 2)) - IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
                 xlHoja1.Cells(84, loRs!cRango + 2) = CCur(xlHoja1.Cells(84, loRs!cRango + 2)) + IIf(IsNull(loRs!nSaldo), 0, loRs!nSaldo)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                
            End If
        End If
        loRs.MoveNext
    Loop
    End If
End Sub
'ALPA 20101228**********************************************************************
Private Sub PintaFondeoObligacionesVista(ByRef xlHoja1 As Excel.Worksheet, ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal pnMoneda As Integer)
    Dim loRs As ADODB.Recordset
    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Set loRs = New ADODB.Recordset
    Set loRs = oDbalanceCont.ObtenerFondeoObligVista(pdFecha, pnTipoCambio) 'Antes ObtenerFondeoGirosxPagar 20190518
    If Not (loRs.BOF Or loRs.EOF) Then
    Do While Not loRs.EOF
        If loRs!nTipoCobertura = 1 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(98, 3) = CCur(xlHoja1.Cells(98, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(98, 3) = CCur(xlHoja1.Cells(98, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(100, 3) = CCur(xlHoja1.Cells(100, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If

        End If
        If loRs!nTipoCobertura = 0 Then
            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
                    xlHoja1.Cells(99, 3) = CCur(xlHoja1.Cells(99, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "3" Then
                    xlHoja1.Cells(99, 3) = CCur(xlHoja1.Cells(99, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
                    xlHoja1.Cells(100, 3) = CCur(xlHoja1.Cells(100, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
            End If
        End If
        loRs.MoveNext
    Loop
    End If
    
'    Set loRs = oDbalanceCont.ObtenerFondeoRetencionesJudiciales(pdFecha, pnTipoCambio)
'    If Not (loRs.BOF Or loRs.EOF) Then
'    Do While Not loRs.EOF
'        If loRs!nTipoCobertura = 1 Then
'            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
'                    xlHoja1.Cells(98, 3) = CCur(xlHoja1.Cells(98, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
'            End If
'            If loRs!nPersoneria = "3" Then
'                    xlHoja1.Cells(98, 3) = CCur(xlHoja1.Cells(98, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
'            End If
'            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
'                    xlHoja1.Cells(100, 3) = CCur(xlHoja1.Cells(100, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
'            End If
'
'        End If
'        If loRs!nTipoCobertura = 0 Then
'            If loRs!nPersoneria = "1" Or loRs!nPersoneria = "2" Then
'                    xlHoja1.Cells(99, 3) = CCur(xlHoja1.Cells(99, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
'            End If
'            If loRs!nPersoneria = "3" Then
'                    xlHoja1.Cells(99, 3) = CCur(xlHoja1.Cells(99, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
'            End If
'            If loRs!nPersoneria = "4" Or loRs!nPersoneria = "5" Or loRs!nPersoneria = "6" Or loRs!nPersoneria = "7" Or loRs!nPersoneria = "8" Or loRs!nPersoneria = "9" Then
'                    xlHoja1.Cells(100, 3) = CCur(xlHoja1.Cells(100, 3)) + IIf(pnMoneda = 1, loRs!nSaldCntMN, loRs!nSaldCntME)
'            End If
'        End If
'        loRs.MoveNext
'    Loop
'    End If 'Comentado by NAGL 20190518
End Sub

Public Sub ReporteGananciaPerdidaSPOT(pdFechaInicio As Date, pdFechaFin As Date)
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
    
    Dim rsCompraVenta As ADODB.Recordset
    Set rsCompraVenta = New ADODB.Recordset
    
    Dim oCompraVenta As NCompraVenta
    Set oCompraVenta = New NCompraVenta
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim N As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
On Error GoTo GeneraExcelGPErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ReporteGananciaPerdidaSPOT"
    'Primera Hoja ******************************************************
    lsNomHoja = "GananciaPerdida"
    '*******************************************************************
    lsArchivo1 = "\spooler\ReporteGananciaPerdidaSPOT" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
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
    
    nSaltoContador = 8
    nContTotal = 0
    Set rsCompraVenta = oCompraVenta.ObtenerReporteGanaciaPerdidaSpot(Format(pdFechaInicio, "YYYYMMDD"), Format(pdFechaFin, "YYYYMMDD"))
    nPase = 1
    If (rsCompraVenta Is Nothing) Then
        nPase = 0
    End If
    xlHoja1.Cells(4, 5) = "Reporte Al " & Format(pdFechaFin, "DD") & " DE " & UCase(Format(pdFechaFin, "MMMM")) & " DEL  " & Format(pdFechaFin, "YYYY")
    If nPase = 1 Then
        Do While Not rsCompraVenta.EOF
    '        DoEvents
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = rsCompraVenta!dFechaMov
                xlHoja1.Cells(nSaltoContador, 2) = rsCompraVenta!cAgeDescripcion
                xlHoja1.Cells(nSaltoContador, 3) = rsCompraVenta!cMovDesc
                xlHoja1.Cells(nSaltoContador, 4) = rsCompraVenta!cUser
                xlHoja1.Cells(nSaltoContador, 5) = rsCompraVenta!nMontoCompra
                xlHoja1.Cells(nSaltoContador, 6) = rsCompraVenta!nMontoVenta
                xlHoja1.Cells(nSaltoContador, 7) = rsCompraVenta!nTCSpot
                xlHoja1.Cells(nSaltoContador, 8) = rsCompraVenta!nTCFM
                xlHoja1.Cells(nSaltoContador, 9) = rsCompraVenta!nTCFD
                xlHoja1.Cells(nSaltoContador, 10) = rsCompraVenta!nGananciMen
                xlHoja1.Cells(nSaltoContador, 11) = rsCompraVenta!nPerdidaMen
                xlHoja1.Cells(nSaltoContador, 12) = rsCompraVenta!nGananciDia
                xlHoja1.Cells(nSaltoContador, 13) = rsCompraVenta!nPerdidaDia
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 5), xlHoja1.Cells(nSaltoContador, 13)).NumberFormat = "#,##0.00;-#,##0.00"
                nSaltoContador = nSaltoContador + 1
            rsCompraVenta.MoveNext
            nContTotal = nContTotal + 1
            If rsCompraVenta.EOF Then
               Exit Do
            End If
        Loop
    End If
    If nContTotal > 0 Then
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Font.Bold = True
            xlHoja1.Range("A" & nSaltoContador, "D" & nSaltoContador).MergeCells = True
            xlHoja1.Cells(nSaltoContador, 1) = "Totales"
            
            xlHoja1.Cells(nSaltoContador, 5).Formula = "=SUM(E8:E" & (nSaltoContador - 1) & ")"
            xlHoja1.Cells(nSaltoContador, 6).Formula = "=SUM(F8:F" & (nSaltoContador - 1) & ")"
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 5), xlHoja1.Cells(nSaltoContador, 6)).NumberFormat = "#,##0.00;-#,##0.00"
            
            xlHoja1.Cells(nSaltoContador, 10).Formula = "=SUM(J8:J" & (nSaltoContador - 1) & ")"
            xlHoja1.Cells(nSaltoContador, 11).Formula = "=SUM(K8:K" & (nSaltoContador - 1) & ")"
            xlHoja1.Cells(nSaltoContador, 12).Formula = "=SUM(L8:L" & (nSaltoContador - 1) & ")"
            xlHoja1.Cells(nSaltoContador, 13).Formula = "=SUM(M8:M" & (nSaltoContador - 1) & ")"
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 10), xlHoja1.Cells(nSaltoContador, 13)).NumberFormat = "#,##0.00;-#,##0.00"
            xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
    End If
    Set oCompraVenta = Nothing
    If nPase = 1 Then
        rsCompraVenta.Close
    End If
    Set rsCompraVenta = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

Exit Sub
GeneraExcelGPErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub
'FRHU 20140104 RQ13658
Public Sub RepEstadoFlujoEfectivo(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim ix As Integer
    Dim ixAc As Integer
    Dim nPos As Integer
    Dim rsFlujo As New ADODB.Recordset
    Dim rsFilas As New ADODB.Recordset
    Dim oRep As New DRepFormula
    Dim oRepA As DRepFormula
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim fs As New Scripting.FileSystemObject
    Dim nPorcenPrincipal As Integer
    Dim rsNotasDet As New ADODB.Recordset
    Dim iCab As Long, iDet As Long
    Dim lsPath As String, lsArchivo As String
    Dim lbAbierto As Boolean
    Dim lnFilaActual As Integer, lnColumnaActual As Integer
    Dim lnNivelMax As Integer, lnUltimaColumna As Integer
    Dim lsNombreMesEvalua As String, lsNombreMesCompara As String
    Dim lnAnioEvalua As Integer, lnAnioCompara As Integer
    Dim lsFormula1Evalua As String, lsFormula1Compara As String
    Dim lsFormula2Evalua As String, lsFormula2Compara As String
    Dim lsFormula3Evalua As String, lsFormula3Compara As String
    Dim lsFormula4Evalua As String, lsFormula4Compara As String
    Dim lsFormula5Evalua As String, lsFormula5Compara As String
    Dim lnMontoFormula1Evalua As Currency, lnMontoFormula1Compara As Currency
    Dim lnMontoFormula2Evalua As Currency, lnMontoFormula2Compara As Currency
    Dim lnMontoFormula3Evalua As Currency, lnMontoFormula3Compara As Currency
    Dim lnMontoFormula4Evalua As Currency, lnMontoFormula4Compara As Currency
    Dim lnMontoFormula5Evalua As Currency, lnMontoFormula5Compara As Currency
    Dim lnMontoConsolidadoAnu As Currency
    Dim lnMontoConsolidadoMen As Currency
    Dim lnMontoConsolidadoAnuAc As Currency
    Dim nCantidadTrabajadores As Integer
    Dim nReNeAc As Currency
    Dim nReNeAcAA As Currency
    Dim nPromedio12Meses1 As Currency
    Dim nPromedio12Meses2 As Currency
    Dim nPromedio12Meses3 As Currency
    Dim lsFormula1 As String
    Dim lsFormula2 As String
    Dim lsFormula3 As String
    Dim lsDepartamento As String
    Dim nSumaFiCol As Currency
    Dim oBala As DbalanceCont
    Dim oRS As ADODB.Recordset
    Dim ldFechaProceso As Date
    Dim oBalInser As DbalanceCont
    Dim oRsInsert As ADODB.Recordset
    Dim lnMontoResultadoMes As Currency
    Dim nSumaDistCreditosDirectos As Currency
    Dim lnSumaActivoMes1 As Currency
    Dim lnSumaActivoMes2 As Currency
    Dim oMov As New DMov
    Dim sMovNro As String
    Dim dfecha As Date
    Dim cMES As String
    Dim cMesDos As String
    Dim cValorFilas As String
    Dim dFechaEvaluaP As Date
    Dim dFechaEvaluaS As Date
    Dim lsFormulaP As String
    Dim lsFormulaS As String
    '************
    Dim MatDatos() As TColumna
    ReDim MatDatos(0)
    Dim lsTmp As String
    Dim lcrSaldoP As Currency
    Dim lcrSaldoTotalP As Currency
    Dim lcrSaldoS As Currency
    Dim lcrSaldoTotalS As Currency
    Dim fila As Integer
    '************
    lsPath = App.path & "\FormatoCarta\EstadoDeFlujoDeEfectivo.xlsx"
    lsArchivo = "EstadoDeFlujoDeEfectivo_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & ".xlsx"
            
     'valida formato carta
    If Len(Dir(lsPath)) = 0 Then
        MsgBox "No se pudo encontrar el archivo: " & lsPath & "," & Chr(10) & "comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    'verifica formato carta abierto
    If fs.FileExists(lsPath) Then
        lbAbierto = True
        Do While lbAbierto
            If ArchivoEstaAbierto(lsPath) Then
                lbAbierto = True
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(lsPath) + " para continuar", vbRetryCancel, "Aviso") = vbCancel Then
                    Exit Sub
                End If
            Else
                lbAbierto = False
            End If
        Loop
    End If
    
    Set xlsLibro = xlsAplicacion.Workbooks.Open(lsPath)
    Set xlsHoja = xlsLibro.ActiveSheet
        
    If pnSemestre = 1 Then
        dfecha = CDate("30/06/" & CStr(pnAnio))
        xlsHoja.Cells(9, 3) = "JUN." & " " & CStr(pnAnio)
        cMES = "Al 30 de Junio de"
    Else
        dfecha = CDate("31/12/" & CStr(pnAnio))
        xlsHoja.Cells(9, 3) = "DIC." & " " & CStr(pnAnio)
        cMES = "Al 31 de Diciembre de"
    End If
        
    dFechaEvaluaP = dfecha
    dFechaEvaluaS = DateAdd("M", -6, dfecha)
        
    If Month(dFechaEvaluaS) = 6 Then
        xlsHoja.Cells(9, 4) = "JUN." & " " & CStr(Year(dFechaEvaluaS))
        cMesDos = "Al 30 de Junio de"
    Else
        xlsHoja.Cells(9, 4) = "DIC." & " " & CStr(Year(dFechaEvaluaS))
        cMesDos = "Al 31 de Diciembre de"
    End If
         
    xlsHoja.Range("B6", "D6").MergeCells = True
    xlsHoja.Range("B6", "D6").HorizontalAlignment = xlCenter
    xlsHoja.Cells(6, 2) = cMES & " " & CStr(Year(dFechaEvaluaP)) & " y " & cMesDos & " " & CStr(Year(dFechaEvaluaS))

    Set rsFlujo = oRep.ObtenerEstadoFlujoEfectivo()
    
    lnFilaActual = 10
    lnColumnaActual = 2
    'Comienza el Recorrido
    If Not RSVacio(rsFlujo) Then
        For iCab = 1 To rsFlujo.RecordCount
            ReDim Preserve MatDatos(iCab)
            Select Case rsFlujo!nNivel
                Case 1
                 xlsHoja.Cells(lnFilaActual, 2) = rsFlujo!cDescripcion
                Case 2
                 xlsHoja.Cells(lnFilaActual, 2) = "     " & rsFlujo!cDescripcion
                Case 3
                 xlsHoja.Cells(lnFilaActual, 2) = "          " & rsFlujo!cDescripcion
                End Select
            If rsFlujo!nTipo = 2 Then
                lsTmp = ""
                lcrSaldoTotalP = 0
                lcrSaldoTotalS = 0
                cValorFilas = rsFlujo!cvalor
                'Set rsFilas = oRep.ObtenerEstadoFlujoEfectivo()
                For fila = 1 To Len(cValorFilas)
                    If (Mid(Trim(cValorFilas), fila, 1) >= "0" And Mid(Trim(cValorFilas), fila, 1) <= "9") Then
                        lsTmp = lsTmp + Mid(Trim(cValorFilas), fila, 1)
                    Else
                        lcrSaldoP = MatDatos(CInt(lsTmp)).Primero
                        lcrSaldoTotalP = lcrSaldoTotalP + lcrSaldoP
                        lcrSaldoS = MatDatos(CInt(lsTmp)).Segundo
                        lcrSaldoTotalS = lcrSaldoTotalS + lcrSaldoS
                        lsTmp = ""
                    End If
                'rsFilas.MoveNext
                Next fila
                xlsHoja.Cells(lnFilaActual, 3) = lcrSaldoTotalP
                xlsHoja.Cells(lnFilaActual, 4) = lcrSaldoTotalS
                MatDatos(iCab).Primero = xlsHoja.Cells(lnFilaActual, 3)
                MatDatos(iCab).Segundo = xlsHoja.Cells(lnFilaActual, 4)
            Else
                If CStr(rsFlujo!cvalor) <> "" Then
                    If Year(dFechaEvaluaP) <= 2012 And Year(dFechaEvaluaS) <= 2012 Then
                        lsFormulaP = Mid(rsFlujo!cvalor, 1, InStr(rsFlujo!cvalor, "/") - 1) ' Antes del 2013
                        xlsHoja.Cells(lnFilaActual, 3) = ObtenerResultadoFormula(dFechaEvaluaP, lsFormulaP, 0, "")
                        xlsHoja.Cells(lnFilaActual, 4) = ObtenerResultadoFormula(dFechaEvaluaS, lsFormulaP, 0, "")
                    ElseIf Year(dFechaEvaluaP) >= 2013 And Year(dFechaEvaluaS) >= 2013 Then
                        lsFormulaS = Mid(rsFlujo!cvalor, InStr(rsFlujo!cvalor, "/") + 1, Len(rsFlujo!cvalor)) ' Despues del 2013
                        xlsHoja.Cells(lnFilaActual, 3) = ObtenerResultadoFormula(dFechaEvaluaP, lsFormulaS, 0, "")
                        xlsHoja.Cells(lnFilaActual, 4) = ObtenerResultadoFormula(dFechaEvaluaS, lsFormulaS, 0, "")
                    Else
                        lsFormulaP = Mid(rsFlujo!cvalor, 1, InStr(rsFlujo!cvalor, "/") - 1) ' Antes del 2013
                        lsFormulaS = Mid(rsFlujo!cvalor, InStr(rsFlujo!cvalor, "/") + 1, Len(rsFlujo!cvalor)) ' Despues del 2013
                        xlsHoja.Cells(lnFilaActual, 3) = ObtenerResultadoFormula(dFechaEvaluaP, lsFormulaS, 0, "")
                        xlsHoja.Cells(lnFilaActual, 4) = ObtenerResultadoFormula(dFechaEvaluaS, lsFormulaP, 0, "")
                    End If
                    MatDatos(iCab).Primero = xlsHoja.Cells(lnFilaActual, 3)
                    MatDatos(iCab).Segundo = xlsHoja.Cells(lnFilaActual, 4)
                Else
                    MatDatos(iCab).Primero = 0
                    MatDatos(iCab).Segundo = 0
                End If
            End If
            lnFilaActual = lnFilaActual + 1
            rsFlujo.MoveNext
        Next iCab
        xlsHoja.SaveAs App.path & "\Spooler\" & lsArchivo
        MsgBox "Reporte se generó satisfactoriamente en " & App.path & "\Spooler\" & lsArchivo, vbInformation, "Aviso"
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
    Else
        MsgBox "No existe la configuración respectiva para generar el presente Reporte", vbInformation, "Aviso"
    End If
    Set rsFlujo = Nothing
    Set xlsHoja = Nothing
    Set xlsLibro = Nothing
    Set xlsAplicacion = Nothing
End Sub
Private Function ObtenerResultadoFormula(ByVal pdFecha As Date, ByVal psFormula As String, ByVal pnMoneda As Integer, Optional psAgencia As String = "") As Currency
    Dim oBal As New DbalanceCont
    Dim oNBal As New NBalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    Dim sTempAD As String
    Dim nPosicion As Integer
    Dim signo As String
    Dim LsSigno As String
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0
    lsTmp = ""
    lsFormula = Replace(lsFormula, "M", pnMoneda)
    sTempAD = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                
                MatDatos(nCtaCont).CuentaContable = lsTmp
                
                If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
                    MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
                Else
                    'If Trim(psAgencia) = "" Then
                        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), 1)
                        'MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, "0", "1", 0, True)
                    'Else
                        'MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
                    'End If
                End If
                
                If nCtaCont > 1 Then
                    If Mid(Trim(lsFormula), i, 1) = ")" Then
                        nPosicion = 0
                    Else
                        nPosicion = i
                    End If
                End If
                    If sTempAD = "" Then
                        If nCtaCont = 1 Then
                            If ((i - Len(Trim(lsTmp))) - 3) > 1 Then
                                sTempAD = Mid(Trim(lsFormula), (i - Len(Trim(lsTmp))) - 3, 2)
                            Else
                                sTempAD = ""
                            End If
                        Else
                            sTempAD = "" 'Mid(Trim(lsFormula), (i - Len(MatDatos(nCtaCont).CuentaContable)) - 3, 2)
                        End If
                    End If
                
                If sTempAD = "SA" Or sTempAD = "SD" Then
                    MatDatos(nCtaCont).CuentaContable = DepuraSaldoAD(MatDatos(nCtaCont).CuentaContable)
                    If sTempAD = "SA" Then
                        MatDatos(nCtaCont).bSaldoA = True
                        MatDatos(nCtaCont).bSaldoD = False
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = True
                    End If
                    Else
                        MatDatos(nCtaCont).bSaldoA = False
                        MatDatos(nCtaCont).bSaldoD = False
                End If
            End If
            If nPosicion = 0 Then
               sTempAD = ""
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp
        'MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
        If MatDatos(nCtaCont).CuentaContable = "100" Or MatDatos(nCtaCont).CuentaContable = "1000" Then
            MatDatos(nCtaCont).Saldo = MatDatos(nCtaCont).CuentaContable
        Else
            'If Trim(psAgencia) = "" Then
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), 1)
                'MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
            'Else
            '    MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensualxAgencia(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True, psAgencia)
            'End If
        End If
    End If
    'Genero la formula en cadena
    lsTmp = ""
    lsCadFormula = ""
    Dim nEncontrado As Integer
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    nEncontrado = 0
                    If MatDatos(j).CuentaContable = lsTmp Then
                            
                            If MatDatos(j).bSaldoA = True Or MatDatos(j).bSaldoD = True Then
                                MatDatos(j).Saldo = oNBal.CalculaSaldoBECuentaAD(MatDatos(j).CuentaContable, pnMoneda, MatDatos(j).bSaldoA, CStr(pnMoneda), Trim(psAgencia), Format(pdFecha, "YYYY"), Format(pdFecha, "MM"))
                                nEncontrado = 1
                            End If
                                If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                                    
                                    If Right(Trim(lsCadFormula), 1) = "-" Or Right(Trim(lsCadFormula), 1) = "+" Then
                                        If Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "-"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo >= 0 Then
                                            LsSigno = "+"
                                        ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo < 0 Then
                                            LsSigno = "-"
                                        End If
                                    Else
                                        LsSigno = ""
                                    End If
                                    If LsSigno = "" Then
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                                    Else
                                        lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & LsSigno & Format(Abs(MatDatos(j).Saldo), "#0.00")
                                    End If
                                    nEncontrado = 1
                                Else
                                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                                    nEncontrado = 1
                                End If
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            If nEncontrado = 1 Or (Mid(Trim(lsFormula), i, 1) = "S" Or Mid(Trim(lsFormula), i, 1) = "A" Or Mid(Trim(lsFormula), i, 1) = "D") Then
            lsCadFormula = lsCadFormula & Mid(Trim(lsFormula), i, 1)
            Else
            lsCadFormula = lsCadFormula & "" & Mid(Trim(lsFormula), i, 1)
            End If
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp Then
               'lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
               If Left(Format(MatDatos(j).Saldo, "#0.00"), 1) = "-" And (Right(lsCadFormula, 1) = "-" Or Right(lsCadFormula, 1) = "+") Then
                    If Right(Trim(lsCadFormula), 1) = "-" Or Right(Trim(lsCadFormula), 1) = "+" Then
                       If Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo >= 0 Then
                            LsSigno = "-"
                       ElseIf Right(Trim(lsCadFormula), 1) = "-" And MatDatos(j).Saldo < 0 Then
                            LsSigno = "+"
                       ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo >= 0 Then
                            LsSigno = "+"
                       ElseIf Right(Trim(lsCadFormula), 1) = "+" And MatDatos(j).Saldo < 0 Then
                            LsSigno = "-"
                       End If
                    Else
                         LsSigno = ""
                    End If
                       If LsSigno = "" Then
                            lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                       Else
                           lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & LsSigno & Format(Abs(MatDatos(j).Saldo), "#0.00")
                       End If
                       nEncontrado = 1
                    'lsCadFormula = Left(lsCadFormula, Len(lsCadFormula) - 1) & Format(MatDatos(j).Saldo, "#0.00")
                    'nEncontrado = 1
                Else
                    lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                    nEncontrado = 1
                End If
               Exit For
           End If
        Next j
    End If
    lsCadFormula = Replace(Replace(lsCadFormula, "SA", ""), "SD", "")
    ObtenerResultadoFormula = oFormula.ExprANum(lsCadFormula)
    Set oBal = Nothing
    Set oFormula = Nothing
End Function
Private Function DepuraSaldoAD(ByVal sCta As String) As String
Dim i As Integer
Dim Cad As String
    Cad = ""
    For i = 1 To Len(sCta)
        If Mid(sCta, i, 1) >= "0" And Mid(sCta, i, 1) <= "9" Then
            Cad = Cad + Mid(sCta, i, 1)
        End If
    Next i
    DepuraSaldoAD = Cad
End Function
Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function
'FIN FRHU 20140104
'FRHU 20140115 RQ13659
Public Sub RepHojaTrabajoFlujoEfectivo(ByVal pnAnio As Integer, ByVal pnSemestre As Integer)
    Dim ix As Integer
    Dim ixAc As Integer
    Dim nPos As Integer
    Dim rsFlujo As New ADODB.Recordset
    Dim rsFilas As New ADODB.Recordset
    Dim oRep As New DRepFormula
    Dim oRepA As DRepFormula
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim fs As New Scripting.FileSystemObject
    Dim nPorcenPrincipal As Integer
    Dim rsNotasDet As New ADODB.Recordset
    Dim iCab As Long, iDet As Long
    Dim lsPath As String, lsArchivo As String
    Dim lbAbierto As Boolean
    Dim lnFilaActual As Integer, lnColumnaActual As Integer
    Dim lnNivelMax As Integer, lnUltimaColumna As Integer
    Dim lsNombreMesEvalua As String, lsNombreMesCompara As String
    Dim lnAnioEvalua As Integer, lnAnioCompara As Integer
    Dim lsFormula1Evalua As String, lsFormula1Compara As String
    Dim lsFormula2Evalua As String, lsFormula2Compara As String
    Dim lsFormula3Evalua As String, lsFormula3Compara As String
    Dim lsFormula4Evalua As String, lsFormula4Compara As String
    Dim lsFormula5Evalua As String, lsFormula5Compara As String
    Dim lnMontoFormula1Evalua As Currency, lnMontoFormula1Compara As Currency
    Dim lnMontoFormula2Evalua As Currency, lnMontoFormula2Compara As Currency
    Dim lnMontoFormula3Evalua As Currency, lnMontoFormula3Compara As Currency
    Dim lnMontoFormula4Evalua As Currency, lnMontoFormula4Compara As Currency
    Dim lnMontoFormula5Evalua As Currency, lnMontoFormula5Compara As Currency
    Dim lnMontoConsolidadoAnu As Currency
    Dim lnMontoConsolidadoMen As Currency
    Dim lnMontoConsolidadoAnuAc As Currency
    Dim nCantidadTrabajadores As Integer
    Dim nReNeAc As Currency
    Dim nReNeAcAA As Currency
    Dim nPromedio12Meses1 As Currency
    Dim nPromedio12Meses2 As Currency
    Dim nPromedio12Meses3 As Currency
    Dim lsFormula1 As String
    Dim lsFormula2 As String
    Dim lsFormula3 As String
    Dim lsDepartamento As String
    Dim nSumaFiCol As Currency
    Dim oBala As DbalanceCont
    Dim oRS As ADODB.Recordset
    Dim ldFechaProceso As Date
    Dim oBalInser As DbalanceCont
    Dim oRsInsert As ADODB.Recordset
    Dim lnMontoResultadoMes As Currency
    Dim nSumaDistCreditosDirectos As Currency
    Dim lnSumaActivoMes1 As Currency
    Dim lnSumaActivoMes2 As Currency
    Dim oMov As New DMov
    Dim sMovNro As String
    Dim dfecha As Date
    Dim cMES As String
    Dim cMesDos As String
    Dim cValorFilas As String
    Dim dFechaEvaluaP As Date
    Dim dFechaEvaluaS As Date
    Dim lsFormulaP As String
    Dim lsFormulaS As String
    '************
    Dim MatDatos() As TColumna
    ReDim MatDatos(0)
    Dim lsTmp As String
    Dim lcrSaldo As Currency
    Dim lcrSaldoTotal As Currency
    Dim fila As Integer
    '************
    '************ Hoja de Trabajo
    Dim rsHoja As New ADODB.Recordset
    Dim nValor As Integer
    Dim j As Integer
    Dim lcSigno As String
    Dim Fiexcel As Integer
    '*************
    lsPath = App.path & "\FormatoCarta\HojaDeTrabajoFlujoEfectivo.xlsx"
    lsArchivo = "HojaDeTrabajoFlujoEfectivo_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & ".xls"
            
     'valida formato carta
    If Len(Dir(lsPath)) = 0 Then
        MsgBox "No se pudo encontrar el archivo: " & lsPath & "," & Chr(10) & "comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    'verifica formato carta abierto
    If fs.FileExists(lsPath) Then
        lbAbierto = True
        Do While lbAbierto
            If ArchivoEstaAbierto(lsPath) Then
                lbAbierto = True
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(lsPath) + " para continuar", vbRetryCancel, "Aviso") = vbCancel Then
                    Exit Sub
                End If
            Else
                lbAbierto = False
            End If
        Loop
    End If
    
    Set xlsLibro = xlsAplicacion.Workbooks.Open(lsPath)
    Set xlsHoja = xlsLibro.ActiveSheet
        
    If pnSemestre = 1 Then
        dfecha = CDate("30/06/" & CStr(pnAnio))
        xlsHoja.Cells(6, 3) = CStr(dfecha)
        'cMES = "Al 30 de Junio de"
    Else
        dfecha = CDate("31/12/" & CStr(pnAnio))
        xlsHoja.Cells(6, 3) = CStr(dfecha)
        'cMES = "Al 31 de Diciembre de"
    End If
        
    dFechaEvaluaP = dfecha
    dFechaEvaluaS = DateAdd("M", -6, dfecha)
        
    If Month(dFechaEvaluaS) = 6 Then
        xlsHoja.Cells(6, 4) = "30/06/" & " " & CStr(Year(dFechaEvaluaS))
        'cMesDos = "Al 30 de Junio de"
    Else
        xlsHoja.Cells(6, 4) = "31/12" & " " & CStr(Year(dFechaEvaluaS))
        'cMesDos = "Al 31 de Diciembre de"
    End If
         
    xlsHoja.Range("B3", "N3").MergeCells = True
    xlsHoja.Range("B3", "N3").HorizontalAlignment = xlLeft
    xlsHoja.Cells(3, 2) = "HOJA DE TRABAJO FLUJO EFECTIVO AL" & "  " & CStr(dfecha)
    
    '************************** ACTIVO - PASIVO - AJUSTES
    nValor = 1 'Activo
    lnFilaActual = 9 'Fila donde va iniciar el recorrido
    lnColumnaActual = 2
    Do While nValor <= 3
        Set rsHoja = oRep.RecuperaHojaTrabajoFE(nValor)
        Select Case nValor
            Case 1: xlsHoja.Cells(lnFilaActual - 1, 2) = "ACTIVO"
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual - 1)) & ":" & "B" & Trim(Str(lnFilaActual - 1))).Font.Bold = True 'Negrita
                    xlsHoja.Range("C" & Trim(Str(lnFilaActual - 1)) & ":" & "P" & Trim(Str(lnFilaActual - 1))).MergeCells = 1 'Combinar Celdas
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual - 1)) & ":" & "P" & Trim(Str(lnFilaActual - 1))).Borders.LineStyle = 1 'Linea
            Case 2: xlsHoja.Cells(lnFilaActual + 2, 2) = "PASIVO"
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 2)) & ":" & "B" & Trim(Str(lnFilaActual + 2))).Font.Bold = True 'Negrita
                    xlsHoja.Range("C" & Trim(Str(lnFilaActual + 2)) & ":" & "P" & Trim(Str(lnFilaActual + 2))).MergeCells = 1 'Combinar Celdas
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 2)) & ":" & "P" & Trim(Str(lnFilaActual + 2))).Borders.LineStyle = 1 'Linea
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual - 1)) & ":" & "P" & Trim(Str(lnFilaActual - 1))).Borders(xlEdgeBottom).LineStyle = 1  'Linea de Abajo
                    'Campos Vacios
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual)) & ":" & "B" & Trim(Str(lnFilaActual))).Borders(xlEdgeLeft).LineStyle = 1 'Linea Izquierda
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual)) & ":" & "B" & Trim(Str(lnFilaActual))).Borders(xlEdgeRight).LineStyle = 1 'Linea Derecha
                    xlsHoja.Range("P" & Trim(Str(lnFilaActual)) & ":" & "P" & Trim(Str(lnFilaActual))).Borders(xlEdgeRight).LineStyle = 1  'Linea Derecha
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 1)) & ":" & "B" & Trim(Str(lnFilaActual + 1))).Borders(xlEdgeLeft).LineStyle = 1 'Linea Izquierda
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 1)) & ":" & "B" & Trim(Str(lnFilaActual + 1))).Borders(xlEdgeRight).LineStyle = 1 'Linea Derecha
                    xlsHoja.Range("P" & Trim(Str(lnFilaActual + 1)) & ":" & "P" & Trim(Str(lnFilaActual + 1))).Borders(xlEdgeRight).LineStyle = 1  'Linea Derecha
                    lnFilaActual = lnFilaActual + 3 ' fila donde inicia el 2 recorrido
            Case 3: xlsHoja.Cells(lnFilaActual + 2, 2) = "AJUSTES"
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 2)) & ":" & "B" & Trim(Str(lnFilaActual + 2))).Font.Bold = True 'Negrita
                    xlsHoja.Range("C" & Trim(Str(lnFilaActual + 2)) & ":" & "P" & Trim(Str(lnFilaActual + 2))).MergeCells = 1 'Combinar Celdas
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 2)) & ":" & "P" & Trim(Str(lnFilaActual + 2))).Borders.LineStyle = 1 'Linea
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual - 1)) & ":" & "P" & Trim(Str(lnFilaActual - 1))).Borders(xlEdgeBottom).LineStyle = 1  'Linea de Abajo
                    'Campos Vacios
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual)) & ":" & "B" & Trim(Str(lnFilaActual))).Borders(xlEdgeLeft).LineStyle = 1 'Linea Izquierda
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual)) & ":" & "B" & Trim(Str(lnFilaActual))).Borders(xlEdgeRight).LineStyle = 1 'Linea Derecha
                    xlsHoja.Range("P" & Trim(Str(lnFilaActual)) & ":" & "P" & Trim(Str(lnFilaActual))).Borders(xlEdgeRight).LineStyle = 1  'Linea Derecha
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 1)) & ":" & "B" & Trim(Str(lnFilaActual + 1))).Borders(xlEdgeLeft).LineStyle = 1 'Linea Izquierda
                    xlsHoja.Range("B" & Trim(Str(lnFilaActual + 1)) & ":" & "B" & Trim(Str(lnFilaActual + 1))).Borders(xlEdgeRight).LineStyle = 1 'Linea Derecha
                    xlsHoja.Range("P" & Trim(Str(lnFilaActual + 1)) & ":" & "P" & Trim(Str(lnFilaActual + 1))).Borders(xlEdgeRight).LineStyle = 1  'Linea Derecha
                    lnFilaActual = lnFilaActual + 3 ' fila donde inicia el 3 recorrido
        End Select
        If Not RSVacio(rsHoja) Then 'Si hay Informacion
            For iCab = 1 To rsHoja.RecordCount
                xlsHoja.Range("B" & Trim(Str(lnFilaActual)) & ":" & "B" & Trim(Str(lnFilaActual))).Borders(xlEdgeLeft).LineStyle = 1 'Linea Izquierda
                xlsHoja.Range("B" & Trim(Str(lnFilaActual)) & ":" & "B" & Trim(Str(lnFilaActual))).Borders(xlEdgeRight).LineStyle = 1 'Linea Derecha
                xlsHoja.Range("P" & Trim(Str(lnFilaActual)) & ":" & "P" & Trim(Str(lnFilaActual))).Borders(xlEdgeRight).LineStyle = 1  'Linea Derecha
                ReDim Preserve MatDatos(iCab)
                If nValor = 1 Or nValor = 2 Then
                    xlsHoja.Cells(lnFilaActual, 2) = rsHoja!cDescripcion
                Else
                    Select Case rsHoja!nNivel
                        Case 1: xlsHoja.Cells(lnFilaActual, 2) = rsHoja!cDescripcion
                        Case 2: xlsHoja.Cells(lnFilaActual, 2) = "          " & rsHoja!cDescripcion
                        Case 3: xlsHoja.Cells(lnFilaActual, 2) = "                    " & rsHoja!cDescripcion
                    End Select
                End If
                
                If CStr(rsHoja!cForMen2013) <> "" Then
                    If Year(dFechaEvaluaP) <= 2012 And Year(dFechaEvaluaS) <= 2012 Then
                        'lsFormulaP = Mid(rsFlujo!cValor, 1, InStr(rsFlujo!cValor, "/") - 1) ' Antes del 2013
                        xlsHoja.Cells(lnFilaActual, 3) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cForMen2013, 0, ""), "#,###.00")
                        xlsHoja.Cells(lnFilaActual, 4) = Format(ObtenerResultadoFormula(dFechaEvaluaS, rsHoja!cForMen2013, 0, ""), "#,###.00")
                    ElseIf Year(dFechaEvaluaP) >= 2013 And Year(dFechaEvaluaS) >= 2013 Then
                        'lsFormulaS = Mid(rsFlujo!cValor, InStr(rsFlujo!cValor, "/") + 1, Len(rsFlujo!cValor)) ' Despues del 2013
                        xlsHoja.Cells(lnFilaActual, 3) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cForMay2013, 0, ""), "#,###.00")
                        xlsHoja.Cells(lnFilaActual, 4) = Format(ObtenerResultadoFormula(dFechaEvaluaS, rsHoja!cForMay2013, 0, ""), "#,###.00")
                    Else
                        xlsHoja.Cells(lnFilaActual, 3) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cForMay2013, 0, ""), "#,###.00")
                        xlsHoja.Cells(lnFilaActual, 4) = Format(ObtenerResultadoFormula(dFechaEvaluaS, rsHoja!cForMen2013, 0, ""), "#,###.00")
                    End If
                    MatDatos(iCab).Primero = xlsHoja.Cells(lnFilaActual, 3)
                    MatDatos(iCab).Segundo = xlsHoja.Cells(lnFilaActual, 4)
                Else
                        MatDatos(iCab).Primero = Format(0, "#.00")
                        MatDatos(iCab).Segundo = Format(0, "#.00")
                End If
                
                If nValor = 1 Then
                    If MatDatos(iCab).Primero > MatDatos(iCab).Segundo Then
                        MatDatos(iCab).DebeAumento = MatDatos(iCab).Primero - MatDatos(iCab).Segundo
                        xlsHoja.Cells(lnFilaActual, 5) = Format(MatDatos(iCab).DebeAumento, "#,###.00")
                        xlsHoja.Cells(lnFilaActual, 6) = Format(0, "#.00")
                    Else
                        MatDatos(iCab).HaberDisminucion = MatDatos(iCab).Segundo - MatDatos(iCab).Primero
                        xlsHoja.Cells(lnFilaActual, 5) = Format(0, "#.00")
                        xlsHoja.Cells(lnFilaActual, 6) = Format(MatDatos(iCab).HaberDisminucion, "#,###.00")
                    End If
                Else
                    If MatDatos(iCab).Primero < MatDatos(iCab).Segundo Then
                        MatDatos(iCab).DebeAumento = MatDatos(iCab).Segundo - MatDatos(iCab).Primero
                        xlsHoja.Cells(lnFilaActual, 6) = Format(0, "#.00")
                        xlsHoja.Cells(lnFilaActual, 5) = Format(MatDatos(iCab).DebeAumento, "#,###.00")
                    Else
                        MatDatos(iCab).HaberDisminucion = MatDatos(iCab).Primero - MatDatos(iCab).Segundo
                        xlsHoja.Cells(lnFilaActual, 6) = Format(MatDatos(iCab).HaberDisminucion, "#,###.00")
                        xlsHoja.Cells(lnFilaActual, 5) = Format(0, "#.00")
                    End If
                End If
                
                If rsHoja!bMov = True Then
                xlsHoja.Cells(lnFilaActual, 7) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cAjusDebe, 0, ""), "#,###.00")
                xlsHoja.Cells(lnFilaActual, 9) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cAjusHaber, 0, ""), "#,###.00")
                MatDatos(iCab).AjusteDebe = xlsHoja.Cells(lnFilaActual, 7)
                MatDatos(iCab).AjusteHaber = xlsHoja.Cells(lnFilaActual, 9)
                Else
                xlsHoja.Cells(lnFilaActual, 7) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cAjusDebe, 0, ""), "#,###.00")
                xlsHoja.Cells(lnFilaActual, 9) = Format(ObtenerResultadoFormula(dFechaEvaluaP, rsHoja!cAjusHaber, 0, ""), "#,###.00")
                MatDatos(iCab).AjusteDebe = xlsHoja.Cells(lnFilaActual, 7)
                MatDatos(iCab).AjusteHaber = xlsHoja.Cells(lnFilaActual, 9)
                End If
                
                If 3 = 3 Then 'Para ver como se van hacer con las columnas (Sumatoria de las columnas)
                    lsTmp = ""
                    'lcrSaldoTotalP = 0
                    lcrSaldo = 0
                    For j = 1 To 6 'Toma el valor de las ultimas 6 columnas
                        Select Case j
                        Case 1: cValorFilas = rsHoja!cOperDebe
                        Case 2: cValorFilas = rsHoja!cOperHaber
                        Case 3: cValorFilas = rsHoja!cInveDebe
                        Case 4: cValorFilas = rsHoja!cInveHaber
                        Case 5: cValorFilas = rsHoja!cFinaDebe
                        Case 6: cValorFilas = rsHoja!cFinaHaber
                        End Select
                        
                        If cValorFilas <> "" Then ' Inicia el recorrido de columna por columna
                            lcSigno = ""
                            For fila = 1 To Len(cValorFilas)
                                If (Mid(Trim(cValorFilas), fila, 1) >= "0" And Mid(Trim(cValorFilas), fila, 1) <= "9") Then
                                    lsTmp = lsTmp + Mid(Trim(cValorFilas), fila, 1)
                                Else
                                    If lsTmp <> "" Then 'Inicia cuando es diferente a "C"
                                        If Len(lsTmp) >= 2 Then 'Es una cuenta contable
                                            If fila = 2 Then
                                                lcrSaldo = lcrSaldo + Format(ObtenerResultadoFormula(dFechaEvaluaP, lsTmp, 0, ""), "#,###.00")
                                            Else
                                                If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + Format(ObtenerResultadoFormula(dFechaEvaluaP, lsTmp, 0, ""), "#,###.00")
                                                Else
                                                    lcrSaldo = lcrSaldo - Format(ObtenerResultadoFormula(dFechaEvaluaP, lsTmp, 0, ""), "#,###.00")
                                                End If
                                            End If
                                        Else
                                            Select Case CInt(lsTmp)
                                            Case 1
                                            Case 2
                                            Case 3
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).DebeAumento
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).DebeAumento
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).DebeAumento
                                                    End If
                                                End If
                                            Case 4
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).HaberDisminucion
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).HaberDisminucion
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).HaberDisminucion
                                                    End If
                                                End If
                                            Case 5
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteDebe
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteDebe
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).AjusteDebe
                                                    End If
                                                End If
                                            Case 6
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteHaber
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteHaber
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).AjusteHaber
                                                    End If
                                                End If
                                            End Select
                                        End If
                                                       
                                        If Mid(Trim(cValorFilas), fila, 1) = "-" Then
                                            lcSigno = "-"
                                        End If
                                        If Mid(Trim(cValorFilas), fila, 1) = "+" Then
                                            lcSigno = "+"
                                        End If
                                    End If 'Final cuando es diferente a "C"
                                    lsTmp = ""
                                End If
                            Next fila
                            'FRHU20140313 - Obtiene el ultimo valor de la cadena
                            If lsTmp <> "" Then 'Inicia cuando es diferente a "C"
                                        If Len(lsTmp) >= 2 Then 'Es una cuenta contable
                                            If fila = 2 Then
                                                lcrSaldo = lcrSaldo + Format(ObtenerResultadoFormula(dFechaEvaluaP, lsTmp, 0, ""), "#,###.00")
                                            Else
                                                If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + Format(ObtenerResultadoFormula(dFechaEvaluaP, lsTmp, 0, ""), "#,###.00")
                                                Else
                                                    lcrSaldo = lcrSaldo - Format(ObtenerResultadoFormula(dFechaEvaluaP, lsTmp, 0, ""), "#,###.00")
                                                End If
                                            End If
                                        Else
                                            Select Case CInt(lsTmp)
                                            Case 1
                                            Case 2
                                            Case 3
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).DebeAumento
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).DebeAumento
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).DebeAumento
                                                    End If
                                                End If
                                            Case 4
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).HaberDisminucion
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).HaberDisminucion
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).HaberDisminucion
                                                    End If
                                                End If
                                            Case 5
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteDebe
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteDebe
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).AjusteDebe
                                                    End If
                                                End If
                                            Case 6
                                                If fila = 2 Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteHaber
                                                Else
                                                    If lcSigno = "+" Then
                                                    lcrSaldo = lcrSaldo + MatDatos(iCab).AjusteHaber
                                                    Else
                                                    lcrSaldo = lcrSaldo - MatDatos(iCab).AjusteHaber
                                                    End If
                                                End If
                                            End Select
                                        End If
                                                       
                                        If Mid(Trim(cValorFilas), fila, 1) = "-" Then
                                            lcSigno = "-"
                                        End If
                                        If Mid(Trim(cValorFilas), fila, 1) = "+" Then
                                            lcSigno = "+"
                                        End If
                            End If 'Final cuando es diferente a "C"
                            'FIN FRHU20140313
                            Select Case j
                            Case 1
                            cValorFilas = rsHoja!cOperDebe
                            xlsHoja.Cells(lnFilaActual, 11) = lcrSaldo
                            MatDatos(iCab).OperacionDebe = xlsHoja.Cells(lnFilaActual, 11)
                            Case 2
                            cValorFilas = rsHoja!cOperHaber
                            xlsHoja.Cells(lnFilaActual, 12) = lcrSaldo
                            MatDatos(iCab).OperacionHaber = xlsHoja.Cells(lnFilaActual, 11)
                            Case 3
                            cValorFilas = rsHoja!cInveDebe
                            xlsHoja.Cells(lnFilaActual, 13) = lcrSaldo
                            Case 4
                            cValorFilas = rsHoja!cInveHaber
                            xlsHoja.Cells(lnFilaActual, 14) = lcrSaldo
                            Case 5
                            cValorFilas = rsHoja!cFinaDebe
                            xlsHoja.Cells(lnFilaActual, 15) = lcrSaldo
                            Case 6
                            cValorFilas = rsHoja!cFinaHaber
                            xlsHoja.Cells(lnFilaActual, 16) = lcrSaldo
                            End Select
                        End If  ' Termina el recorrido columna por columna
                    lsTmp = "" 'FRHU 20140313
                    lcSigno = "" 'FRHU 20140313
                    lcrSaldo = 0 'FRHU 20140313
                    Next j 'Fin Toma el valor de las ultimas 6 columnas
                End If
                lnFilaActual = lnFilaActual + 1
                rsHoja.MoveNext
            Next iCab
        End If  '******************* Fin Si hay Informacion
        nValor = nValor + 1
        Set rsHoja = Nothing
    Loop
    '************************** FIN ACTIVO - PASIVO - AJUSTES
    xlsHoja.Range("B" & Trim(Str(lnFilaActual - 1)) & ":" & "P" & Trim(Str(lnFilaActual - 1))).Borders(xlEdgeBottom).LineStyle = 1  'Linea de Abajo
    xlsHoja.SaveAs App.path & "\Spooler\" & lsArchivo
    MsgBox "Reporte se generó satisfactoriamente en " & App.path & "\Spooler\" & lsArchivo, vbInformation, "Aviso"
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set rsHoja = Nothing
    Set xlsHoja = Nothing
    Set xlsLibro = Nothing
    Set xlsAplicacion = Nothing
End Sub
'FIN FRHU 20140115 RQ13659
'ALPA 20141219**********************************************************************
'Public Sub ReporteCuentasuInactivasRestringidas(pdFecha As Date)
'    Dim fs As Scripting.FileSystemObject
'    Dim lbExisteHoja As Boolean
'    Dim lsArchivo1 As String
'    Dim lsNomHoja  As String
'    Dim lsNombreAgencia As String
'    Dim lsCodAgencia As String
'    Dim lsMes As String
'    Dim lnContador As Long
'    Dim lsArchivo As String
'    Dim xlsAplicacion As Excel.Application
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'
'    Dim rsBalance As ADODB.Recordset
'    Set rsBalance = New ADODB.Recordset
'
'    Dim oBalance As DbalanceCont
'    Set oBalance = New DbalanceCont
'
'    Dim sTexto As String
'    Dim sDocFecha As String
'    Dim nSaltoContador As Long
'    Dim sFecha As String
'    Dim sMov As String
'    Dim sDoc As String
'    Dim n As Integer
'    Dim pnLinPage As Integer
'    Dim nMES As Integer
'    Dim nSaldo12 As Currency
'    Dim nContTotal As Long
'    Dim nPase As Integer
'On Error GoTo ReporteCuentasuInactivasRestringidasErr
'
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'    lsArchivo = "ReporteCuentasInactivasRestringidas"
'    'Primera Hoja ******************************************************
'    lsNomHoja = "CuentasInactivasRestringidas"
'    '*******************************************************************
'    lsArchivo1 = "\spooler\ReporteCuentasInactivasRestringidas" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    nSaltoContador = 8
'    nContTotal = 0
'    Set rsBalance = oBalance.ObtenerCuentasInactivasRestringidas(CDate(Format(pdFecha, "YYYY/MM/DD")))
'    nPase = 1
'    If (rsBalance Is Nothing) Then
'        nPase = 0
'    End If
'    xlHoja1.Cells(4, 5) = "Reporte Al " & Format(pdFecha, "DD") & " DE " & UCase(Format(pdFecha, "MMMM")) & " DEL  " & Format(pdFecha, "YYYY")
'    If nPase = 1 Then
'        Do While Not rsBalance.EOF
'    '        DoEvents
'                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 13)).Borders.LineStyle = 1
'                xlHoja1.Cells(nSaltoContador, 1) = nSaltoContador - 7
'                xlHoja1.Cells(nSaltoContador, 2) = rsBalance!Agencia
'                xlHoja1.Cells(nSaltoContador, 3) = rsBalance!NombreCliente
'                xlHoja1.Cells(nSaltoContador, 4) = rsBalance!TipoCliente
'                xlHoja1.Cells(nSaltoContador, 5) = rsBalance!NroCuenta
'                xlHoja1.Cells(nSaltoContador, 6) = rsBalance!Moneda
'                xlHoja1.Cells(nSaltoContador, 7) = rsBalance!Producto
'                xlHoja1.Cells(nSaltoContador, 8) = rsBalance!SubProducto
'                xlHoja1.Cells(nSaltoContador, 9) = Format(rsBalance!dFecha, "YYYY/MM/DD")
'                xlHoja1.Cells(nSaltoContador, 10) = rsBalance!CapitalRecla
'                xlHoja1.Cells(nSaltoContador, 11) = rsBalance!InteresRecla
'                xlHoja1.Cells(nSaltoContador, 12) = rsBalance!Total
'                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 10), xlHoja1.Cells(nSaltoContador, 12)).NumberFormat = "#,##0.00;-#,##0.00"
'                nSaltoContador = nSaltoContador + 1
'            rsBalance.MoveNext
'            nContTotal = nContTotal + 1
'            If rsBalance.EOF Then
'               Exit Do
'            End If
'        Loop
'    End If
'    Set oBalance = Nothing
'    If nPase = 1 Then
'        rsBalance.Close
'    End If
'    Set rsBalance = Nothing
'
'    xlHoja1.SaveAs App.path & lsArchivo1
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'Exit Sub
'ReporteCuentasuInactivasRestringidasErr:
'    MsgBox Err.Description, vbInformation, "Aviso"
'    Exit Sub
'End Sub
''************************************************************

Private Sub ReporteCuentasuInactivasRestringidas(pdFecha As Date)
    Dim oBalance As DbalanceCont
    Dim R As ADODB.Recordset
    Dim lMatCabecera As Variant
    Dim lsMensaje As String
    Dim lsNombreArchivo As String

    lsNombreArchivo = "ReporteInactivas"
    
    ReDim lMatCabecera(12, 0)
    
    Set oBalance = New DbalanceCont
    Set R = oBalance.ObtenerCuentasInactivasRestringidas(CDate(Format(pdFecha, "YYYY/MM/DD")))
    Set oBalance = Nothing
    If Not R Is Nothing Then
        Call GeneraReporteEnArchivoExcelInicio(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Cuentas Inactivas Restringidas", "", lsNombreArchivo, lMatCabecera, R, 2, , , True, True)
    Else
        MsgBox lsMensaje, vbInformation, "AVISO"
    End If
End Sub

Private Sub GeneraReporteDiferidosAmpliados_NoAmpliados(pdFecha As Date)
Dim fs As Scripting.FileSystemObject
Dim lbExisteHoja As Boolean
Dim lsArchivo1 As String
Dim lsNomHoja  As String
Dim lsArchivo As String
Dim xlsAplicacion As Excel.Application
Dim xlsLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim oBalance As New DbalanceCont
Dim rs As New ADODB.Recordset
Dim i As Integer, nRep As Integer
Dim psTipo As String, psDescrip As String

Dim oBarraDet As clsProgressBar
Dim TituloProgress As String
Dim MensajeProgress As String
Set oBarraDet = New clsProgressBar
oBarraDet.ShowForm Me
oBarraDet.Max = 5

psDescrip = "Reporte de Intereses Diferidos en Créditos Ampliados y No Ampliados"
oBarraDet.Progress 0, psDescrip, "GENERANDO EL ARCHIVO", "", vbBlue
TituloProgress = psDescrip
MensajeProgress = "GENERANDO EL ARCHIVO"

Set fs = New Scripting.FileSystemObject
Set xlsAplicacion = New Excel.Application
lsArchivo = "ReporteInteresesDiferidosAmp_NoAmp"

lsArchivo1 = "\spooler\ReporteInteresesDiferidosAmp_NoAmp_" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
Else
    MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Por favor solicitar el formato correspondiente", vbInformation, "Advertencia"
    Exit Sub
End If
i = 1
nRep = 2
oBarraDet.Progress 2, TituloProgress, MensajeProgress, "", vbBlue
Do While i <= nRep
    If i = 1 Then
        lsNomHoja = "InteresesDiferidosConAmpliados"
        psTipo = "Amp"
    Else
        lsNomHoja = "InteresDiferidosSinAmpliados"
        psTipo = "No_Amp"
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
    
    Set rs = oBalance.ObtieneReporteIntDifAmp_NoAmp(pdFecha, psTipo)
    oBarraDet.Progress 2 + i, TituloProgress, MensajeProgress, "", vbBlue
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 2)).CopyFromRecordset rs
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Font.Color = vbBlack
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Font.Name = "Calibri"
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Font.Size = 11
    
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(4, 8), xlHoja1.Cells(rs.RecordCount + 3, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(4, 7), xlHoja1.Cells(rs.RecordCount + 3, 7)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(4, 13), xlHoja1.Cells(rs.RecordCount + 3, 13)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(4, 14), xlHoja1.Cells(rs.RecordCount + 3, 14)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(4, 17), xlHoja1.Cells(rs.RecordCount + 3, 19)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(4, 21), xlHoja1.Cells(rs.RecordCount + 3, 23)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(4, 25), xlHoja1.Cells(rs.RecordCount + 3, 29)).HorizontalAlignment = xlRight
    xlsAplicacion.DisplayAlerts = False
    xlsAplicacion.Selection.Replace What:="NULL", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False
    
    xlHoja1.Range(xlHoja1.Cells(4, 7), xlHoja1.Cells(rs.RecordCount + 3, 7)).Style = "Comma"
    xlHoja1.Range(xlHoja1.Cells(4, 7), xlHoja1.Cells(rs.RecordCount + 3, 7)).NumberFormat = "#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(4, 13), xlHoja1.Cells(rs.RecordCount + 3, 13)).Style = "Comma"
    xlHoja1.Range(xlHoja1.Cells(4, 13), xlHoja1.Cells(rs.RecordCount + 3, 13)).NumberFormat = "#,##0.00"

    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(rs.RecordCount + 3, 29)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Set rs = Nothing
    i = i + 1
Loop
    oBarraDet.Progress 5, TituloProgress, MensajeProgress, "", vbBlue
    oBarraDet.CloseForm Me
    Set oBarraDet = Nothing
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub 'NAGL 202008 Según ACTA N°063-2020
