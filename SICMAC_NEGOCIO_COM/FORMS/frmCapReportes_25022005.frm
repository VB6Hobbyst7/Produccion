VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPORTES DE CAPTACIONES"
   ClientHeight    =   7935
   ClientLeft      =   2700
   ClientTop       =   1590
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapReportes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   270
      Left            =   75
      TabIndex        =   41
      Top             =   7575
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "&Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6870
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Frame Frapersoneria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Personería"
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
      Height          =   2355
      Left            =   6840
      TabIndex        =   37
      Top             =   1800
      Visible         =   0   'False
      Width           =   2115
      Begin VB.ListBox lstpersoneria 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         ItemData        =   "frmCapReportes.frx":030A
         Left            =   105
         List            =   "frmCapReportes.frx":0329
         TabIndex        =   38
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame fracmacs 
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   6885
      TabIndex        =   34
      Top             =   6150
      Visible         =   0   'False
      Width           =   1515
      Begin VB.CheckBox chkLlamadas 
         Caption         =   "Llamadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   36
         Top             =   285
         Width           =   1170
      End
      Begin VB.CheckBox chkRecepcion 
         Caption         =   "Recepcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   35
         Top             =   555
         Width           =   1080
      End
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Ordenar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   6900
      TabIndex        =   30
      Top             =   6150
      Width           =   1890
      Begin VB.CheckBox chkTotal 
         Caption         =   "Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   495
         Width           =   765
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Transacción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   32
         Top             =   270
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   31
         Top             =   510
         Width           =   1080
      End
   End
   Begin VB.Frame frafechacheques 
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
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   6900
      TabIndex        =   27
      Top             =   6165
      Width           =   1845
      Begin VB.OptionButton Option1 
         Caption         =   "de Valorizacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   510
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "de Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.Frame fraCheque 
      Caption         =   "Estado Actual Cheq"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1965
      Left            =   6930
      TabIndex        =   25
      Top             =   4155
      Visible         =   0   'False
      Width           =   1845
      Begin VB.ListBox lstcheques 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "frmCapReportes.frx":03A7
         Left            =   180
         List            =   "frmCapReportes.frx":03BD
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   315
         Width           =   1560
      End
   End
   Begin VB.Frame fraUser 
      Height          =   675
      Left            =   6990
      TabIndex        =   22
      Top             =   3465
      Width           =   1770
      Begin VB.CheckBox Check1 
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   30
         TabIndex        =   24
         Top             =   255
         Width           =   690
      End
      Begin SICMACT.TxtBuscar TxtBuscarUser 
         Height          =   345
         Left            =   720
         TabIndex        =   23
         Top             =   195
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         sTitulo         =   ""
         EnabledText     =   0   'False
         ForeColor       =   12582912
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6825
      Left            =   45
      TabIndex        =   20
      Top             =   675
      Width           =   6705
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   4215
         Top             =   4470
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
               Picture         =   "frmCapReportes.frx":0422
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":0774
               Key             =   "Bebe"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":0AC6
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCapReportes.frx":0E18
               Key             =   "Hijito"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeRep 
         Height          =   6525
         Left            =   120
         TabIndex        =   21
         Top             =   165
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   11509
         _Version        =   393217
         HideSelection   =   0   'False
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
      Begin VB.OLE OleExcel 
         Class           =   "Excel.Sheet.8"
         Height          =   870
         Left            =   4680
         OleObjectBlob   =   "frmCapReportes.frx":116A
         TabIndex        =   40
         Top             =   5820
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame fraAgencias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   660
      Left            =   75
      TabIndex        =   15
      Top             =   0
      Width           =   8715
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   165
         TabIndex        =   17
         Top             =   300
         Width           =   930
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2475
         TabIndex        =   18
         Top             =   240
         Width           =   6045
      End
   End
   Begin VB.Frame fraTipoCambio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   7020
      TabIndex        =   13
      Top             =   2805
      Visible         =   0   'False
      Width           =   1725
      Begin SICMACT.EditMoney EditMoney3 
         Height          =   285
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame fraMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Montos"
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
      Height          =   975
      Left            =   7020
      TabIndex        =   8
      Top             =   1770
      Visible         =   0   'False
      Width           =   1725
      Begin SICMACT.EditMoney txtMonto 
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtMontoF 
         Height          =   285
         Left            =   480
         TabIndex        =   12
         Top             =   585
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   278
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   630
         Width           =   180
      End
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
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
      Height          =   975
      Left            =   7035
      TabIndex        =   3
      Top             =   750
      Width           =   1725
      Begin MSMask.MaskEdBox txtFechaF 
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Top             =   570
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   480
         TabIndex        =   5
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAl 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   630
         Width           =   180
      End
      Begin VB.Label lblDel 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   285
         Width           =   285
      End
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   2085
      Top             =   6405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkCondensado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Condensado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7965
      TabIndex        =   1
      Top             =   7500
      Width           =   1035
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6870
      TabIndex        =   0
      Top             =   7500
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   360
      Left            =   6735
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   635
      _Version        =   393217
      TextRTF         =   $"frmCapReportes.frx":2982
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   -45
      Top             =   7170
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmCapReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim Flag As Boolean
Dim flag1 As Boolean
Dim Char12 As Boolean
Dim lbPrtCom As Boolean
Dim NumA As String
Dim SalA As String
Dim NumC As String
Dim NumP As String
Dim SalP As String
Dim SalC As String
Dim TB As Currency
Dim TBD As Currency
Dim ca As Currency
Dim CAE As Currency
Dim RegCmacS As Integer
Dim vBuffer As String
Dim lscadena As String
Dim Progreso As clsProgressBar
'Dim WithEvents lsRep As nCaptaReportes

Dim oGen As DGeneral

Private Sub Check1_Click()
If Check1.value = 1 Then
    TxtBuscarUser.Enabled = True
Else
    TxtBuscarUser.Enabled = False
End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = 1 Then
    
        TxtAgencia.Text = ""
        TxtAgencia_EmiteDatos
        
        fraUser.Enabled = False
        Check1.Enabled = False
        Check1.value = 0
        TxtBuscarUser.Text = ""
        
    Else
        fraUser.Enabled = True
        Check1.Enabled = True
        Check1.value = 0
        TxtBuscarUser.Text = ""
        
    End If
End Sub

Private Sub cmdExportarExcel_Click()
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nFila As Long, i As Long

Dim orep As nCaptaReportes, RSTEMP As ADODB.Recordset
Dim lsRep As String

lsRep = Mid(TreeRep.SelectedItem.Key, 2, 7)

If MsgBox("Este reporte puede demorar unos minutos..." & vbCrLf & "¿Desea procesar la información ?", vbOKOnly + vbQuestion, "AVISO") = vbNo Then
    Exit Sub
End If
prgBar.value = 0
prgBar.Visible = True

Set orep = New nCaptaReportes
'(lsRep, Me.TxtFecha, Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, Val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria)
'  lscadena = lscadena & GetRepCapNMejoresClientes(CLng(pnMontoIni), pnTipoCambio, psCodAge, psEmpresa, psNomAge, pdFecSis)
Select Case lsRep
       Case "280708"
            'Set RSTEMP = orep.GetDataCapNMejoresClientes(Me.txtMonto.value, Val(EditMoney3.Text), Me.TxtAgencia.Text, gsNomCmac, gsNomAge, gdFecSis, 1)
            
       Case "280709"
            'Set RSTEMP = orep.GetDataCapNMejoresClientes(Me.txtMonto.value, Val(EditMoney3.Text), Me.TxtAgencia.Text, gsNomCmac, gsNomAge, gdFecSis, 2)
            
End Select


If RSTEMP.EOF Then
    MsgBox "No se encontro información para este reporte", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
End If

Dim lsArchivoN As String, lbLibroOpen As Boolean




   lsArchivoN = App.path & "\Spooler\Rep" & lsRep & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
   
   OleExcel.Class = "ExcelWorkSheet"
   lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
   If lbLibroOpen Then
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
            
            
                       
            
            nFila = 1
            
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
            xlHoja1.Range("F2:H2").MergeCells = True
            xlHoja1.Cells(nFila, 6) = Format(gdFecSis, "Long Date")
             
             prgBar.value = 2
                 
             
            nFila = 3
            xlHoja1.Cells(nFila, 1) = "REPORTE DE LOS " & CStr(Me.txtMonto.Text) & " MEJORES CLIENTES DE LA CAJA " & IIf(lsRep = "280708", " CON I.F. ", " SIN I.F.")
             
            
            xlHoja1.Range("A1:M5").Font.Bold = True
            
            xlHoja1.Range("A3:M3").MergeCells = True
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
            xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter
 
            'xlHoja1.Range("A5:H5").AutoFilter
            
            nFila = 5
            
            xlHoja1.Cells(nFila, 1) = "ITEM"
            xlHoja1.Cells(nFila, 2) = "CODIGO"
            xlHoja1.Cells(nFila, 3) = "NOMBRE"
            xlHoja1.Cells(nFila, 4) = "DIRECCION"
            xlHoja1.Cells(nFila, 5) = "SALDO"
            xlHoja1.Cells(nFila, 6) = "FONO"
            xlHoja1.Cells(nFila, 7) = "FEC. NAC."
            xlHoja1.Cells(nFila, 8) = "ZONA"
            
            i = 0
            While Not RSTEMP.EOF
                nFila = nFila + 1
                
                prgBar.value = ((i) / RSTEMP.RecordCount) * 100
                
                i = i + 1

                
                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                xlHoja1.Cells(nFila, 2) = RSTEMP!cCodPers
                xlHoja1.Cells(nFila, 3) = RSTEMP!cNomPers
                xlHoja1.Cells(nFila, 4) = RSTEMP!cDirPers
                xlHoja1.Cells(nFila, 5) = Format(RSTEMP!nSaldo, "#,##0.00")
                xlHoja1.Cells(nFila, 6) = RSTEMP!cTelPers & ""
                xlHoja1.Cells(nFila, 7) = Format(RSTEMP!dFecNac, gsFormatoFechaView)
                xlHoja1.Cells(nFila, 8) = RSTEMP!Zona
                                            
                RSTEMP.MoveNext
                
            Wend
            
           ' xlHoja1.Columns.AutoFit
            
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
            
            prgBar.value = 100
            
   End If
   
   Set RSTEMP = Nothing
   
   prgBar.Visible = False
    
End Sub
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
  MsgBox Err.Description, vbInformation, "Aviso"
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
   MsgBox Err.Description, vbInformation, "Aviso"
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



Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsprevio
    Set oPrevio = New Previo.clsprevio
    Dim lscadena As String
    Dim lsRep As String
    Dim lsEstadosCheques As String
    Dim lsOptionsCheques As String
    Dim lscmacllamada As String
    Dim lscmacrecepcion As String
    Dim lsOrden As String
    Dim lscheck As String
    Dim i As Integer
    Dim orep As nCaptaReportes
    Set orep = New nCaptaReportes
    

    
    
    If fraFecha.Visible = True And txtFecha.Visible = True Then
        If Not IsDate(txtFecha) Then
            MsgBox "Fecha no valida", vbInformation, "Aviso"
            Me.txtFecha.SetFocus
            Exit Sub
        End If
    End If
    
    If fraFecha.Visible = True And txtFechaF.Visible = True Then
        If Not IsDate(txtFechaF) Then
            MsgBox "Fecha no valida", vbInformation, "Aviso"
            Me.txtFechaF.SetFocus
            Exit Sub
        End If
    End If
    
    
    If fraFecha.Visible = True And txtFecha.Visible = True And txtFechaF.Visible = True Then
        If CDate(txtFechaF) < CDate(txtFecha) Then
            MsgBox "La fecha de finalizacion debe ser mayor a la fecha de inicio.", vbInformation, "Aviso"
            Me.txtFechaF.SetFocus
            Exit Sub
        End If
    End If
    
    
       
    If TxtAgencia.Visible = True Then
        If TxtAgencia.Text = "" And Me.chkTodos.value = 0 Then
            MsgBox "Agencia No valida", vbInformation, "Aviso"
            Me.TxtAgencia.SetFocus
            Exit Sub
        End If
    End If
    
    If Mid(TreeRep.SelectedItem.Key, 2, 7) = "280710" Then
    
        Call PrintExtractos(Trim(TxtAgencia.Text), CDate(txtFecha), CDate(txtFechaF))
        
        Exit Sub
        
    End If
    
    
    
    If fraUser.Visible = True Then
        If Check1.value = 1 Then
            If Len(Trim(TxtBuscarUser.Text)) > 0 Then
            Else
                MsgBox "Seleccione un usuario", vbInformation, "Aviso"
                Me.TxtBuscarUser.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If fraTipoCambio.Visible = True Then
        If Val(EditMoney3.Text) = 0 Then
            MsgBox "Ingrese Tipo de Cambio válido", vbExclamation, "Aviso"
            If fraFecha.Visible = True Then
                GetTipCambio (txtFecha.Text)
            Else
                GetTipCambio (gdFecSis)
            End If
            EditMoney3.Text = gnTipCambio
            EditMoney3.SetFocus
        End If
    End If
    
    If fracmacs.Visible = True Then
        If chkLlamadas.value = 0 And chkRecepcion.value = 0 Then
            MsgBox "Seleccione una opcion de llamada/recepcion", vbExclamation, "Aviso"
            chkLlamadas.SetFocus
            Exit Sub
        End If
    End If
    
    If fraCheque.Visible = True Then
        For i = 0 To lstcheques.ListCount - 1
            If lstcheques.Selected(i) = True Then
                If Len(Trim(lsEstadosCheques)) > 0 Then
                    lsEstadosCheques = lsEstadosCheques & ", " & Right(lstcheques.List(i), 1) & ""
                Else
                    lsEstadosCheques = "" & Right(lstcheques.List(i), 1) & ""
                End If
            End If
        Next
        
        If Option1(0).value = True Then
            lsOptionsCheques = "1"
        ElseIf Option1(1).value = True Then
            lsOptionsCheques = "2"
        End If
    Else
        lsEstadosCheques = ""
        lsOptionsCheques = ""
    End If
    
    If fraOrden.Visible = True Then
        If Option2(0).value = True Then
            lsOrden = "1"
            lscheck = "0"
        Else
            lsOrden = "2"
            If chkTotal.value = 1 Then
                lscheck = "1"
            Else
                lscheck = "0"
            End If
        End If
    End If
    
    
    If fracmacs.Visible = True Then
        If chkLlamadas.value = 1 Then
            lscmacllamada = "1"
        Else
            lscmacllamada = "0"
        End If
        If chkRecepcion.value = 1 Then
            lscmacrecepcion = "1"
        Else
            lscmacrecepcion = "0"
        End If
    End If
    
    rtfCartas.FileName = App.path & cPlantillaCartaRenPF
    
    Dim pspersoneria As Integer
    If lstpersoneria.ListCount > 0 And lstpersoneria.Visible Then
        If lstpersoneria.ListIndex = -1 Then
            MsgBox "Seleccione el tipo de personería para este reporte", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        pspersoneria = lstpersoneria.ItemData(lstpersoneria.ListIndex)
    End If
    lsRep = Mid(TreeRep.SelectedItem.Key, 2, 7)
    
    lscadena = orep.Reporte(lsRep, IIf(lsRep = "280216", gdFecSis, Me.txtFecha), Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, Val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria)
    
    If lsRep = "280231" Then
        'oPrevio.Show lscadena, Caption, True, 66
        Exit Sub
    End If
    
    If chkCondensado.value = 1 Then
        oPrevio.Show lscadena, Caption, True, 66
    Else
        oPrevio.Show lscadena, Caption, False, 66
    End If
    
End Sub

Private Sub PrintExtractos(ByVal SCodAge As String, ByVal Fechaini As Date, ByVal FechaFin As Date)
    Dim orep As nCaptaReportes
    Set orep = New nCaptaReportes
    
    
    'Call orep.ImpExtractosBatch(SCodAge, Fechaini, FechaFin, sLpt)
    
    Set orep = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim lsCodCab As String
    Dim lsCodCab1 As String
    Dim oCons As DConstante
    Set oCons = New DConstante
    Set oGen = New DGeneral
    LlenaArbol
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    Set Progreso = New clsProgressBar
    
    txtFecha = Format(gdFecSis, gsFormatoFechaView)
    txtFechaF = Format(gdFecSis, gsFormatoFechaView)
    
    lstpersoneria.ItemData(0) = gPersonaNat
    lstpersoneria.ItemData(1) = gPersonaJurSFL
    lstpersoneria.ItemData(2) = gPersonaJurCFL
    lstpersoneria.ItemData(3) = gPersonaJurCFLCMAC
    lstpersoneria.ItemData(4) = gPersonaJurCFLCRAC
    lstpersoneria.ItemData(5) = gPersonaJurCFLFONCODES
    lstpersoneria.ItemData(6) = gPersonaJurCFLCooperativa
    lstpersoneria.ItemData(7) = gPersonaJurCFLEdpyme
    lstpersoneria.ItemData(8) = 0
    
    
    Me.TxtAgencia.rs = oCons.GetAgencias(, , True)
    Usuario.Inicio gsCodUser
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



Private Sub LlenaArbol()
Dim sqlv As String
'Dim PObjConec as DConecta
Dim rsUsu As New ADODB.Recordset
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String

Dim PObjConec As New DConecta

    'Set PObjConec = New DConecta

'    sqlv = " Select cOpeCod Codigo, UPPER(cCapRepDescripcion) + ' [' + Rtrim(ltrim(Str(nCapRepDiaCalculo))) + '-' + rtrim(ltrim(str(nCapRepRangoMonto))) + '-' + rtrim(ltrim(Str(bCapRepTipoCambio))) + ']' Descripcion, nOpeNiv Nivel From OpeTpo OP" _
'         & " Inner Join CaptaReportes CAPR On CAPR.cCapRepCod = OP.cOpeCod " _
'         & " WHERE copecod like '28%' " _
'         & " order by cOpecod "

    sqlv = "exec  Cap_ReportesCaptaciones"
    
    PObjConec.AbreConexion
     
    Set rsUsu = PObjConec.CargaRecordSet(sqlv)
    
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("Codigo")
        sOperacion = sOpeCod & " " & UCase(rsUsu("Descripcion"))
        Select Case rsUsu("Nivel")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = TreeRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = TreeRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
    
    PObjConec.CierraConexion
Set PObjConec = Nothing
End Sub





Private Sub Option2_Click(Index As Integer)
chkTotal.Enabled = IIf(Index = 0, False, True)
chkTotal.value = 0
End Sub

Private Sub treeRep_Click()
    Dim i As Integer
    Dim lsOpc As String
        
        Limpia
        
        Me.Caption = "REPORTES DE CAPTACIONES " & Mid(TreeRep.SelectedItem.Text, 8, Len(TreeRep.SelectedItem.Text) - 14)
      
         
        fraMonto.Caption = "Montos"
        Frapersoneria.Visible = False
        
        'Nuevos
        Select Case Mid(TreeRep.SelectedItem, 1, 6)
        
        '''''''''''''''''INICIO'''''''''''''''''
        
        Case gCapRepDiaCapEstadAho
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapEstadPF
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapEstadCTS
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapCtasMov
            HabilitaControles True, True, True, True, False, False, False, False, False
        Case gCapRepDiaCapSaldTpoCta
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepDiaCapEstratCta
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepDiaCapPFVenc
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapInact
            HabilitaControles False, False, False, False, False, False, False, False, False
        Case 280232
            HabilitaControles True, False, True, True, False, False, False, False, False
            
        Case gCapRepDiaCapConsInact
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapApert
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCapCanc
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case gCapRepDiaCartaRenPF
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case gCapRepDiaServGirosApert
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case gCapRepDiaServGirosCanc
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case gCapRepDiaServConvCob
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoPN
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoPJSFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoPJCFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoCMAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoCRAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoFoncodes
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoCooperativa
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaAhoEdipyme
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280632"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280633"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280634"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280635"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280636"
                Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case "280637"
                Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case "280639"
                HabilitaControles True, False, False, False, False, False, False, False, False
                
            
        Case gCapRepMensCapSaldCtaPFPN
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFPJSFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFPJCFL
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCMAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCRAC
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFFoncodes
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCooperativa
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFEdpyme
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFPNPJSFL
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFNascaPalpa
                 Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaPFCaneteMala
                 Frapersoneria.Visible = True
                HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapPF
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case "280631"
                Frapersoneria.Visible = True
                HabilitaControles True, False, False, False, False, False, False, False, False
                
                
        Case gCapRepMensCapSaldCtaCTS
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSConvenio
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSExternos
            HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSNascaPalpa
            HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensCapSaldCtaCTSCaneteMala
            HabilitaControles False, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapCTS
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapCTSConvenio
                HabilitaControles True, False, False, False, False, False, False, False, False
        Case gCapRepMensBloqCapCTSExternos
                HabilitaControles True, False, False, False, False, False, False, False, False
                
        
        Case gCapRepMensCapListGralCtas
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case "V03"
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case "280701"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280702"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280703"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280704"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280705"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280706"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280707"
            HabilitaControles True, False, True, True, True, True, True, True, True
        Case "280213"
            HabilitaControles True, True, True, True, False, False, False, False, False, , True 'ojo
        Case "280214"
            HabilitaControles True, True, True, True, False, False, False, False, False, , True 'ojo
        Case "280215"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280216"
            'HabilitaControles True, False,  True, True, False, False, False, False, False, False
            HabilitaControles True, False, False, False, False, False, False, False, False, False
        Case "280217"
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case "280708"
            fraMonto.Caption = "Nro. Clientes"
            cmdExportarExcel.Visible = True
            HabilitaControles True, False, False, False, False, True, True, True, False
        Case "280709"
            fraMonto.Caption = "Nro. Clientes"
            cmdExportarExcel.Visible = True
            HabilitaControles True, False, False, False, False, True, True, True, False
        Case "280710"
            HabilitaControles True, False, True, True, True, False, False, False, False
                       
            
        Case "280224"
            HabilitaControles False, False, True, True, True, False, False, False, False
        Case "280225"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280227"
            HabilitaControles True, False, True, True, False, False, False, False, False
        Case "280228"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280219"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280220"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280230" 'cheques recibidos
            HabilitaControles True, False, True, True, True, False, False, False, False, True
        Case "280231"
            HabilitaControles True, False, True, True, True, False, False, False, False
            
            
        Case "280223"
            HabilitaControles True, False, True, True, True, False, False, False, False
        Case "280222"
            HabilitaControles True, False, True, True, False, False, False, False, False, , , True
        
        '''''''''''''FIN'''''''''''''''''''''
        'Case 280213, 280214
        '    HabilitaControles True
        Case Else
            HabilitaControles False, False, False, False, False, False, False, False, False
        End Select
        End Sub

Private Sub HabilitaControles(ByVal pfraAgencia As Boolean, ByVal pUser As Boolean, ByVal pfraFecha As Boolean, pFecha As Boolean, pfechaf As Boolean, _
                              ByVal pfraMonto As Boolean, ByVal ptipocambio As Boolean, ByVal pMonto As Boolean, _
                              ByVal pMontof As Boolean, Optional pFraCheque As Boolean = False, Optional pFraOrden As Boolean = False, Optional pFraCmacs As Boolean = False)
    
    fraCheque.Visible = pFraCheque
    frafechacheques.Visible = pFraCheque
    
    fraOrden.Visible = pFraOrden
    
    fraAgencias.Visible = pfraAgencia
    
    fraUser.Visible = pUser
    fraFecha.Visible = pFecha
    
    fraMonto.Visible = pfraMonto
    txtMonto.Visible = pMonto
    txtMontoF.Visible = pMontof
    Label3.Visible = pMontof
    
    fraFecha.Visible = pfraFecha
    txtFecha.Visible = pFecha
    txtFechaF.Visible = pfechaf
    lblAl.Visible = pfechaf
    
    fraTipoCambio.Visible = ptipocambio
    
    fracmacs.Visible = pFraCmacs
    
End Sub


Private Sub Limpia()
    txtFecha.Text = "__/__/____"
    txtFechaF.Text = "__/__/____"
    txtMonto.Text = ""
    txtMontoF.Text = ""
    chkTodos.value = 0
    TxtAgencia.Text = ""
    lblAgencia.Caption = ""
    TxtBuscarUser.Text = ""
    chkLlamadas.value = 0
    chkRecepcion.value = 0
    cmdExportarExcel.Visible = False
    
End Sub

Private Sub treeRep_Collapse(ByVal Node As MSComctlLib.Node)
'    If Right(Node.Key, 2) = "00" Then
'        Node.ExpandedImage = 2
'    Else
'        Node.ExpandedImage = 1
'    End If
End Sub
Private Sub treeRep_Expand(ByVal Node As MSComctlLib.Node)
'    If Right(Node.Key, 2) = "00" Then
'        Node.ExpandedImage = 2
'    Else
'        Node.ExpandedImage = 1
'    End If
End Sub

Private Sub TreeRep_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    treeRep_Click
End If
End Sub

Private Sub treeRep_NodeCheck(ByVal Node As MSComctlLib.Node)
'    Dim i As Integer
'    TreeRep.SelectedItem = Node
'    Select Case Len(Node.Key)
'       Case Node.Key = "P"
'         If Node.Checked = True Then
'              For i = 1 To TreeRep.Nodes.Count
'                  TreeRep.Nodes(i).Checked = True
'              Next
'         Else
'              For i = 1 To TreeRep.Nodes.Count
'                  TreeRep.Nodes(i).Checked = False
'                 Next
'         End If
'       Case 7 And Right(Node.Key, 2) = "00"
'           If Node.Checked = True Then
'              For i = 1 To TreeRep.Nodes.Count
'                If Mid(TreeRep.Nodes(i).Key, 2, 4) = Mid(Node.Key, 2, 4) Then
'                      TreeRep.Nodes(i).Checked = True
'                      TreeRep.Nodes(i).Image = "Hoja"
'                      TreeRep.Nodes(i).ForeColor = vbBlue
'                End If
'              Next
'              TreeRep.SelectedItem.Image = "Close"
'           Else
'              For i = 1 To TreeRep.Nodes.Count
'                If Mid(TreeRep.Nodes(i).Key, 2, 4) = Mid(Node.Key, 2, 4) Then
'                      TreeRep.Nodes(i).Checked = False
'                      TreeRep.Nodes(i).Image = "Hoja1"
'                     TreeRep.Nodes(i).ForeColor = vbBlack
'                 End If
'              Next
'              TreeRep.SelectedItem.Image = "Open"
'           End If
'       Case Else
'
'           If Node.Checked = True Then
'               Node.Image = "Hoja"
'               Node.ForeColor = vbBlue
'           Else
'               Node.Image = "Hoja1"
'               Node.ForeColor = vbBlack
'           End If
'    End Select
End Sub

Private Sub TreeRep_NodeClick(ByVal Node As MSComctlLib.Node)
    treeRep_Click
End Sub

'Private Sub TxtAgencia_EmiteDatos()
'    Me.lblAgencia.Caption = TxtAgencia.psDescripcion
'End Sub

Private Sub TxtAgencia_EmiteDatos()

Me.lblAgencia.Caption = TxtAgencia.psDescripcion
If chkTodos.value = 0 Then
    If TxtAgencia <> "" And lblAgencia <> "" Then
        TxtBuscarUser = ""
        TxtBuscarUser.psRaiz = "USUARIOS " & TxtAgencia.psDescripcion
        TxtBuscarUser.Enabled = True
        TxtBuscarUser.rs = oGen.GetUserAreaAgencia(Usuario.cAreaCodAct, TxtAgencia)
    End If
Else
    TxtBuscarUser.Text = ""
    TxtBuscarUser.Enabled = False
End If
End Sub



Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtFechaF.Enabled And txtFechaF.Visible Then txtFechaF.SetFocus
    End If
End Sub

Private Sub txtFechaF_GotFocus()
fEnfoque txtFechaF
End Sub

Private Sub txtFechaF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TreeRep.SetFocus
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMontoF.Enabled And txtMontoF.Visible Then txtMontoF.SetFocus
    End If
End Sub
