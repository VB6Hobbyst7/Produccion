VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAdeudProv1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  PROVISION DE ADEUDADOS"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmAdeudProv1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB 
      Height          =   180
      Left            =   1575
      TabIndex        =   35
      Top             =   7875
      Visible         =   0   'False
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   34
      Top             =   7815
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "                                                                           "
            TextSave        =   "                                                                           "
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   8985
      TabIndex        =   32
      Top             =   6960
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8985
      TabIndex        =   31
      Top             =   7350
      Width           =   1155
   End
   Begin VB.Frame fraCabecera 
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   75
      TabIndex        =   7
      Top             =   60
      Width           =   10245
      Begin VB.Frame Frame3 
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   600
         Left            =   4485
         TabIndex        =   25
         Top             =   15
         Width           =   4230
         Begin VB.OptionButton optBuscar 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   2
            Left            =   3255
            TabIndex        =   4
            Top             =   225
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton optBuscar 
            Caption         =   "Adeudado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   2130
            TabIndex        =   3
            Top             =   225
            Width           =   1080
         End
         Begin VB.OptionButton optBuscar 
            Caption         =   "Institución Financiera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   225
            Width           =   1845
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
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
         Left            =   9120
         TabIndex        =   8
         Top             =   1005
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         Caption         =   "Movimiento"
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
         Height          =   600
         Left            =   -15
         TabIndex        =   24
         Top             =   15
         Width           =   4350
         Begin VB.TextBox txtTasaVac 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3195
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   840
            TabIndex        =   0
            Top             =   210
            Width           =   1290
            _ExtentX        =   2275
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tasa VAC :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2340
            TabIndex        =   27
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha : "
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
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   270
            Width           =   720
         End
      End
      Begin VB.Frame fraopciones 
         Caption         =   "Institución Financiera"
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
         Height          =   645
         Left            =   0
         TabIndex        =   20
         Top             =   690
         Visible         =   0   'False
         Width           =   9030
         Begin Sicmact.TxtBuscar txtCodObjeto 
            Height          =   345
            Left            =   1065
            TabIndex        =   6
            Top             =   240
            Width           =   2625
            _ExtentX        =   4630
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
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Objeto :"
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
            Height          =   210
            Left            =   105
            TabIndex        =   23
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3750
            TabIndex        =   22
            Top             =   315
            Width           =   930
         End
         Begin VB.Label lblObjDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4695
            TabIndex        =   21
            Top             =   255
            Width           =   3570
         End
      End
      Begin VB.Frame FraGenerales 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   660
         Left            =   30
         TabIndex        =   13
         Top             =   675
         Visible         =   0   'False
         Width           =   9030
         Begin VB.ComboBox cboEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7665
            Style           =   1  'Simple Combo
            TabIndex        =   15
            Text            =   "cboEstado"
            Top             =   285
            Width           =   1290
         End
         Begin VB.TextBox txtNroCtaIF 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   4680
            TabIndex        =   14
            Top             =   285
            Width           =   2310
         End
         Begin Sicmact.TxtBuscar txtBuscarCtaIF 
            Height          =   315
            Left            =   1065
            TabIndex        =   5
            Top             =   285
            Width           =   2685
            _ExtentX        =   4736
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
            sTitulo         =   ""
         End
         Begin VB.Label lblDescIF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "sss"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   1605
            TabIndex        =   19
            Top             =   0
            Width           =   315
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta IF:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "N° Cuenta :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   3750
            TabIndex        =   17
            Top             =   315
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   7020
            TabIndex        =   16
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   555
         Left            =   9495
         Picture         =   "frmAdeudProv1.frx":08CA
         Stretch         =   -1  'True
         Top             =   60
         Width           =   660
      End
   End
   Begin VB.Frame fraDetalle 
      Enabled         =   0   'False
      Height          =   6345
      Left            =   60
      TabIndex        =   28
      Top             =   1425
      Width           =   10275
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   8925
         TabIndex        =   12
         Top             =   5145
         Width           =   1155
      End
      Begin VB.TextBox txtGlosa 
         Height          =   945
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   5280
         Width           =   8685
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " &Todos"
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
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   255
         Width           =   915
      End
      Begin MSComctlLib.ListView lstCabecera 
         Height          =   4515
         Left            =   60
         TabIndex        =   10
         Top             =   555
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   7964
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   29
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Entidad Financiera"
            Object.Width           =   3510
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Adeudado"
            Object.Width           =   2736
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Capital"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Int. Prov."
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ult. Pago"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Periodo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "tasaInt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Dias"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Interes * Vac Actual"
            Object.Width           =   847
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Interes Real"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "SK Real"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "EntidaCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "EntidadCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Vencimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Concesional"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "cCodLinCred"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "cDesLinCred"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Sk x Vac Ant"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "SK LP Real"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "SK LP x VAC Ant"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "SK LP x VAC Act"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "dUltActsaldos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "cTpoCuota"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Comision"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "SaldoProvisionMes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   1275
         X2              =   7785
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PROVISION ADEUDADOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7905
         TabIndex        =   33
         Top             =   255
         Width           =   2250
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   5055
         Width           =   465
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   435
      Left            =   -30
      TabIndex        =   30
      Top             =   7845
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   767
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmAdeudProv1.frx":10A0
   End
End
Attribute VB_Name = "frmAdeudProv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbTrans  As Boolean
Dim lnTasaVac As Double
Dim lMN As Boolean
Dim cMoneda As String
Dim sErrorPro As String
Dim objPista As COMManejador.Pista 'ARLO20170217

'Dim oPrg As clsProgressBar

Private Function ValidaProvisionar() As Boolean
    
Dim I As Integer

Dim oAdeud As New NCajaAdeudados
Dim oCta As New DCtaCont

Dim lsPersCod As String
Dim lsPersNombre As String
Dim lsIFTpo As String
Dim lsCtaIFCod As String
Dim cMonedaPago As String
Dim lscCodLinCred As String

Dim lsCtaDebe As String
Dim lsCtaHaber As String

Dim lsCuentaCorto As String
Dim lsCuentaLargo As String

Dim lsCuentaVac As String

Dim lsMensaje As String
Dim lsMensaje1 As String

Dim nCantidad As Long

Dim gsMovNro As String

'ALPA 20110809
Dim lsCuentaComsionAdelantadaD As String
Dim lsCuentaComsionAdelantadaH As String

    If Len(Trim(Me.txtGlosa.Text)) = 0 Then
        ValidaProvisionar = False
        MsgBox "Ingrese una glosa", vbInformation, "Aviso"
        Exit Function
    End If

    nCantidad = 0
    For I = 1 To lstCabecera.ListItems.Count
        If lstCabecera.ListItems(I).Checked = True Then
            nCantidad = nCantidad + 1
        End If
    Next
    
    If nCantidad = 0 Then
        ValidaProvisionar = False
        MsgBox "Ud. debe seleccionar al menos un registro para provisionar", vbInformation, "Aviso"
        Exit Function
    End If

    'Mes Cerrado
    gsMovNro = GeneraMovNroActualiza(CDate(Me.txtFecha.Text), gsCodUser, gsCodCMAC, gsCodAge)
            
''''    If Not PermiteModificarAsiento(gsMovNro, False) Then
''''        ValidaProvisionar = False
''''        MsgBox "Ud. está intentando provisionar un mes cerrado", vbInformation, "Aviso"
''''        Exit Function
''''    End If
    'Fin Mes Cerrado

    For I = 1 To lstCabecera.ListItems.Count
    
        If lstCabecera.ListItems(I).Checked = True Then
                             
            'Inicializamos para las estadisticas
            
            lsPersCod = Mid(lstCabecera.ListItems(I).SubItems(14), 4, 13)
            lsPersNombre = lstCabecera.ListItems(I).SubItems(2)
            lsIFTpo = Mid(lstCabecera.ListItems(I).SubItems(14), 1, 2)
            lsCtaIFCod = Mid(lstCabecera.ListItems(I).SubItems(14), 18, 10)
            cMonedaPago = lstCabecera.ListItems(I).SubItems(13)
            
            lscCodLinCred = oAdeud.GetSubCtaLineaCredito(lstCabecera.ListItems(I).SubItems(17))
              
            'lsCtaDebe = oAdeud.GetOpeCta(gsOpeCod, "D", "0", lsPersCod, lsIfTpo) & lscCodLinCred
            lsCtaDebe = oAdeud.GetOpeCta(gsOpeCod, "D", "1", lsPersCod, lsIFTpo) & lscCodLinCred 'SE QUITA PARA E CONFIANZA
            lsCtaHaber = oAdeud.GetOpeCta(gsOpeCod, "H", "1", lsPersCod, lsIFTpo)
            'ALPA 20110809
            lsCuentaComsionAdelantadaD = GetOpeCta(gsOpeCod, "D", "B", lsPersCod, lsIFTpo)
            lsCuentaComsionAdelantadaH = GetOpeCta(gsOpeCod, "H", "B", lsPersCod, lsIFTpo)

            'CP LP
            lsCuentaCorto = oAdeud.GetOpeCta(gsOpeCod, "D", "5", lsPersCod, lsIFTpo)
            lsCuentaLargo = oAdeud.GetOpeCta(gsOpeCod, "H", "5", lsPersCod, lsIFTpo)
            'lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 50)
             
            If cMoneda = "1" And cMonedaPago = "2" Then
                lsCuentaVac = oAdeud.GetOpeCta(gsOpeCod, "V", "2", lsPersCod, lsIFTpo) & lscCodLinCred
            End If
            
            sErrorPro = ""
            
            If Trim(lsCtaDebe) = "" Then
                sErrorPro = "Cuenta Debe No está registrada (D:(Provisión de Interes)) "
            End If
            
            If Trim(lsCtaHaber) = "" And Trim(sErrorPro) = "" Then
                sErrorPro = "Cuenta Haber No está registrada (H:(Provisión de Interes)) "
            End If

            If Trim(lsCuentaComsionAdelantadaD) = "" And Trim(sErrorPro) = "" Then
                sErrorPro = "Cuenta lsCuentaComsionAdelantada No está registrada (D:(Comision Adelantada)) "
            End If

            If Trim(lsCuentaComsionAdelantadaH) = "" And Trim(sErrorPro) = "" Then
                sErrorPro = "Cuenta lsCuentaComsionAdelantada No está registrada (H:(Comision Adelantada)) "
            End If

            If Trim(lsCuentaCorto) = "" And Trim(sErrorPro) = "" Then
                sErrorPro = "Cuenta lsCuentaCorto No está registrada (D:(Plazo)) "
            End If
            
            If Trim(lsCuentaLargo) = "" And Trim(sErrorPro) = "" Then
                sErrorPro = "Cuenta Largo No está registrada (H:(Plazo)) "
            End If

            
            lsMensaje1 = ""
            Dim pnPivot As Integer
            pnPivot = 0
            'Comision Adelantada D y H
            If Len(Trim(lsCuentaComsionAdelantadaD)) = 0 Then
                lsMensaje1 = lsMensaje1 & "CuentaDebe No Definida " & Chr(10)
                MsgBox "Cta Comision D no está definida"
                pnPivot = 1
            End If
            'Comision Adelantada D y H
            If Len(Trim(lsCuentaComsionAdelantadaH)) = 0 Then
                lsMensaje1 = lsMensaje1 & "CuentaDebe No Definida " & Chr(10)
                'MsgBox "Cta Comision H no está definida"
                pnPivot = 1
            End If
           
            'Cuenta Debe
            If Len(Trim(lsCtaDebe)) = 0 Then
                lsMensaje1 = lsMensaje1 & "CuentaDebe No Definida " & Chr(10)
                'MsgBox "Cta debe no está definida"
                pnPivot = 1
            End If
            
            lsMensaje1 = oCta.VerificaExisteCuenta(lsCtaDebe, True)
            If lsMensaje1 = "" Then
            Else
                lsMensaje1 = lsMensaje1 & "Cuenta Debe No Existe" & Chr(10)
                'MsgBox "Cta debe no está definida"
                pnPivot = 1
            End If
            
            lsMensaje1 = oCta.VerificaExisteCuenta(lsCtaDebe, True)
            If lsMensaje1 <> "" Then
                lsMensaje1 = lsMensaje1 & "Cuenta Debe no es Ultimo Nivel" & Chr(10)
                pnPivot = 1
            End If
            
            'Cuenta Haber
            If Len(Trim(lsCtaHaber)) = 0 Then
                lsMensaje1 = lsMensaje1 & " CuentaHaber No Definida " & Chr(10)
                pnPivot = 1
            End If
            
            lsMensaje1 = oCta.VerificaExisteCuenta(lsCtaHaber, False)
            If lsMensaje1 = "" Then
            Else
                lsMensaje1 = lsMensaje1 & "Cuenta Haber No Existe" & Chr(10)
                pnPivot = 1
            End If
            
            lsMensaje1 = oCta.VerificaExisteCuenta(lsCtaHaber, True)
            If lsMensaje1 <> "" Then
                lsMensaje1 = lsMensaje1 & "Cuenta Haber no es Ultimo Nivel" & Chr(10)
                pnPivot = 1
            End If
            
            'Solo entra cuando es pago de adeudado mas no cuando se provisiona GITU 05-08-2008
            If Not (gsOpeCod = "401805" Or gsOpeCod = "402805") Then
                'Cuenta CP
                If Len(Trim(lsCuentaCorto)) = 0 Then
                    lsMensaje1 = lsMensaje1 & " CuentaCP No Definida" & Chr(10)
                    pnPivot = 1
                End If
            
                lsMensaje1 = oCta.VerificaExisteCuenta(lsCuentaCorto, False)
                If lsMensaje1 = "" Then
                Else
                    lsMensaje1 = lsMensaje1 & "Cuenta CP No Existe" & Chr(10)
                    pnPivot = 1
                End If
            
                lsMensaje1 = oCta.VerificaExisteCuenta(lsCuentaCorto, True)
                If lsMensaje1 <> "" Then
                    lsMensaje1 = lsMensaje1 & "Cuenta CP no es Ultimo Nivel" & Chr(10)
                    pnPivot = 1
                End If
            
                'Cuenta LP
                If Len(Trim(lsCuentaLargo)) = 0 Then
                    lsMensaje1 = lsMensaje1 & " CuentaLP No Definida" & Chr(10)
                    pnPivot = 1
                End If
            
                lsMensaje1 = oCta.VerificaExisteCuenta(lsCuentaLargo, False)
                If lsMensaje1 = "" Then
                Else
                    lsMensaje1 = lsMensaje1 & "Cuenta LP No Existe" & Chr(10)
                    pnPivot = 1
                End If
            
                lsMensaje1 = oCta.VerificaExisteCuenta(lsCuentaLargo, True)
                If lsMensaje1 <> "" Then
                    lsMensaje1 = lsMensaje1 & "Cuenta LP no es Ultimo Nivel" & Chr(10)
                    pnPivot = 1
                End If
            End If
            'End Gitu
            If cMoneda = "1" And cMonedaPago = "2" Then
                If Len(Trim(lsCuentaVac)) = 0 Then
                    lsMensaje1 = lsMensaje & " CuentaVAC No Definida"
                    pnPivot = 1
                End If
                
                lsMensaje1 = oCta.VerificaExisteCuenta(lsCuentaVac, True)
                
                If lsMensaje1 = "" Then
                Else
                    lsMensaje1 = lsMensaje1 & "Cuenta VAC No Existe" & Chr(10)
                    pnPivot = 1
                End If
                
                lsMensaje1 = oCta.VerificaExisteCuenta(lsCuentaVac, True)
                If lsMensaje1 <> "" Then
                    lsMensaje1 = lsMensaje1 & "Cuenta VAC no es Ultimo Nivel" & Chr(10)
                    pnPivot = 1
                End If
            End If
            
            If Len(Trim(lsMensaje1)) > 0 Then
                lsMensaje = Trim(lsPersNombre) & " [" & lsPersCod & "] - " & lsCtaIFCod & " : " & Chr(10) & lsMensaje1 & Chr(10) & Chr(10)
                lstCabecera.ListItems(I).Checked = False
            End If
                  
        End If
    Next
    
    If Len(Trim(lsMensaje)) = 0 And pnPivot = 0 Then
        ValidaProvisionar = True
    Else
        ValidaProvisionar = False
        MsgBox "Los siguientes registros no se pueden provisionar" & Chr(10) & Chr(10) & lsMensaje, vbExclamation, "Aviso"
    End If

End Function
 

Private Sub chkTodos_Click()
Dim I As Integer
    If lstCabecera.ListItems.Count = 0 Then
    Else
        For I = 1 To lstCabecera.ListItems.Count
            If lstCabecera.ListItems(I).Text <> "" Then
                lstCabecera.ListItems(I).Checked = IIf(chkTodos.value = 1, True, False)
            End If
        Next
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim lnImporte  As Currency
    Dim lsPersCod  As String
    Dim lsIFTpo    As String
    Dim lsCtaIFCod As String
    Dim I As Integer, j As Integer
    Dim lsMsgErr   As String
    Dim lnFilaActual    As Integer
    Dim lsListaAsientos As String
    Dim oAdeu           As New NCajaAdeudados
    Dim nRetorno        As Integer
    Dim lsImpre         As String
    
    If chkTodos.value = 1 Then
        If MsgBox("Esta marcado la opción todos, esta seguro de proceder?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    
    On Error GoTo ErrorAceptar
    If Not ValidaProvisionar Then
        'ALPA 20110831*****************
        If Trim(sErrorPro) <> "" Then
            MsgBox sErrorPro
        End If
        '******************************

        Exit Sub
    End If
    
    
    If MsgBox(" ¿ Desea Grabar la Operación ? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        Exit Sub
    End If
    
   
    lnFilaActual = lstCabecera.SelectedItem.Index
    lsListaAsientos = ""
    Me.cmdAceptar.Enabled = False
    nRetorno = oAdeu.GrabaProvisionAdeudado(gsOpeCod, cMoneda, Me.txtFecha.Text, Val(txtTasaVac.Text), lstCabecera, Me.txtGlosa.Text, Right(gsCodAge, 2), gsCodUser, lsListaAsientos)
    
    If nRetorno = 0 Then
        If Len(Trim(lsListaAsientos)) > 0 Then
            lsImpre = ImprimeAsientosContables(lsListaAsientos, PB, Status, "")
            EnviaPrevio lsImpre, "ASIENTOS DE PROVISION DE PAGO DE ADEUDADOS", gnLinPage, False
        End If

        MsgBox "Provisión Efectuada Satisfactoriamente", vbInformation, "Aviso"
        
        lstCabecera.ListItems.Remove lnFilaActual

    ElseIf nRetorno = 2 Then
        MsgBox "Se efectuaron solo algunas grabaciones de Provisión de Adeudados", vbInformation, "Aviso"
        Me.cmdAceptar.Enabled = True
    ElseIf nRetorno = 1 Then
        MsgBox "No se pudo efectuar ninguna grabación de Provisión de Adeudados", vbInformation, "Aviso"
        Me.cmdAceptar.Enabled = True
    End If
    
    lstCabecera.ListItems.Clear
  
    txtGlosa = ""
        
    MsgBox "Provisión Efectuada satisfactoriamente", vbInformation, "Aviso"
     
    cmdAceptar.Enabled = False
    cmdCancelar.SetFocus
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
    
    Exit Sub
ErrorAceptar:
    lsMsgErr = TextErr(Err.Description)
    MsgBox lsMsgErr, vbInformation, "Aviso"
End Sub

 
Private Sub cmdCancelar_Click()
    cmdAceptar.Enabled = False
    cmdProcesar.Enabled = True
    Limpiar
    BlanqueaTodo
    fraCabecera.Enabled = True
    fraDetalle.Enabled = False
    cmdProcesar.SetFocus
End Sub

Private Sub cmdProcesar_Click()
    On Error GoTo ErrorAceptar
    
    If ValFecha(txtFecha) = False Then Exit Sub
    If Valida = False Then Exit Sub
    Me.MousePointer = 11
    Me.cmdProcesar.Enabled = False
    CargaDatos txtFecha
    Me.cmdProcesar.Enabled = True
    
    If lstCabecera.ListItems.Count > 0 Then
        lstCabecera.ListItems(1).Selected = True
        fraCabecera.Enabled = False
        fraDetalle.Enabled = True
        cmdAceptar.Enabled = True
        cmdProcesar.Enabled = False
        chkTodos.value = 1
    Else
        cmdAceptar.Enabled = False
        chkTodos.value = 0
    End If
    
    
    Me.MousePointer = 0
    Exit Sub
ErrorAceptar:
    Me.MousePointer = 0
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Function Valida() As Boolean
    Valida = False
    If optBuscar(0).value = True Then
        If txtCodObjeto = "" Then
            Valida = False
            MsgBox "Seleccione un tipo de Institución Financiera", vbExclamation, "Aviso"
            Exit Function
        End If
    ElseIf optBuscar(1).value = True Then
        If txtBuscarCtaIF = "" Then
            Valida = False
            MsgBox "Seleccione un Pagaré", vbExclamation, "Aviso"
            Exit Function
        End If
    End If

Valida = True

End Function

Private Sub cmdSalir_Click()
'Dim i As Integer
'i = 1
'Do While True
'
'    MsgBox lstCabecera.ListItems(i).Text
'    i = i + 1
'
'    If i > lstCabecera.ListItems.Count Then
'        Exit Do
'    End If
'Loop
    Unload Me
End Sub

Private Sub Form_Load()
'    CentraForm Me
'    Me.Caption = "  " & gsOpeDesc
'    Me.txtFecha = gdFecSis
'
'    Dim oAdeud As New DCaja_Adeudados
'    lnTasaVac = oAdeud.CargaIndiceVAC(gdFecSis)
'    txtTasaVac = lnTasaVac
'
'    'Agregado para Busquedas
'
'    Dim oOpe As DOperacion
'    Dim oGen As DGeneral
'    Set oOpe = New DOperacion
'    Set oGen = New DGeneral
'    txtBuscarCtaIF.rs = oOpe.GetRsOpeObj("40" & Mid(gsOpeCod, 3, 1) & "832", "0", , "' and ci.cCtaIFCod LIKE '__" & Mid(gsOpeCod, 3, 1) & "%' and ci.cCtaIFEstado = '" & gEstadoCtaIFActiva)
'    CargaCombo cboEstado, oGen.GetConstante(gCGEstadoCtaIF)
'
'    Dim oIF As New DCajaCtasIF
'    Me.txtCodObjeto.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), "__05" & Mid(gsOpeCod, 3, 1) & "%", 1)
'    Set oIF = Nothing
' edpyme - verifica q exista valor vac
    CentraForm Me
    Me.Caption = "  " & gsOpeDesc
    Me.txtFecha = gdFecSis
    
    Dim oAdeud As New DCaja_Adeudados
    lnTasaVac = oAdeud.CargaIndiceVAC(gdFecSis)
        
    If lnTasaVac = 0 Then
       MsgBox "Tasa VAC no ha sido definida", vbInformation, "Aviso"
       txtTasaVac = 0
    Else
        txtTasaVac = lnTasaVac
    End If
    
    'Agregado para Busquedas
    
    Dim oOpe As DOperacion
    Dim oGen As DGeneral
    Set oOpe = New DOperacion
    Set oGen = New DGeneral
    txtBuscarCtaIF.rs = oOpe.GetRsOpeObj("40" & Mid(gsOpeCod, 3, 1) & "832", "0", , "' and ci.cCtaIFCod LIKE '__" & Mid(gsOpeCod, 3, 1) & "%' and ci.cCtaIFEstado = '" & gEstadoCtaIFActiva)
    CargaCombo cboEstado, oGen.GetConstante(gCGEstadoCtaIF)
    
    Dim oIF As New DCajaCtasIF
    Me.txtCodObjeto.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), "__05" & Mid(gsOpeCod, 3, 1) & "%", 1)
    Set oIF = Nothing
    cMoneda = Mid(gsOpeCod, 3, 1)
    Call TamanoColumnas

End Sub

Private Sub CargaDatos(ldFecha As Date)
'    Dim rs As ADODB.Recordset
'    Dim N As Integer
'    Dim lnMontoTotal As Currency
'    Dim lnInteres As Currency
'    Dim lnTotal As Integer, I As Integer
'    Dim sCadena As String
'
'    Dim L As ListItem
'
'    lstCabecera.ListItems.Clear
'
'    If optBuscar(0).value = True Then
'        sCadena = " AND ci.CPERSCOD='" & Mid(txtCodObjeto, 4, 13) & "'"
'    ElseIf optBuscar(1).value = True Then
'        sCadena = " AND ci.CPERSCOD='" & Mid(txtBuscarCtaIF, 4, 13) & "' and ci.cCtaIFCod='" & Right(txtBuscarCtaIF, 7) & "'"
'    End If
'
'    lnTasaVac = txtTasaVac
'    If lnTasaVac = 0 Then
'        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada" & Chr(13) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'            txtTasaVac.SetFocus
'            Exit Sub
'        End If
'    End If
'    Dim oAdeud As New DCaja_Adeudados
'    Dim oIF As New NCajaAdeudados
'
'    Set rs = oAdeud.GetAdeudadosProvision(txtFecha, Mid(gsOpeCod, 3, 1), sCadena)
'
'    If rs.BOF Then
'        lnTotal = 0
'    Else
'        lnTotal = rs.RecordCount
'        PB.Visible = True
'        PB.Min = 0
'        PB.Max = rs.RecordCount
'        PB.value = 0
'    End If
'
'    I = 0
'
'    Do While Not rs.EOF
'
'        PB.value = PB.value + 1
'
'        I = I + 1
'
'        Set L = lstCabecera.ListItems.Add(, , I)
'        L.SubItems(2) = Trim(rs!cPersNombre)
'        L.SubItems(3) = Trim(rs!cCtaIFDesc)
'
'        L.SubItems(5) = Trim(rs!nInteresPagado)  ' Interes acumulado pagado por cuota
'        L.SubItems(6) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
'        L.SubItems(7) = Trim(rs!nNroCuota)    ' numero de cuota pendiente
'        L.SubItems(8) = Trim(rs!nPeriodo)
'        L.SubItems(9) = Trim(rs!nInteres)
'        L.SubItems(10) = Trim(rs!nDiasUltPAgo + IIf(gsCodCMAC = "102", 1, 0))
'        If Val(Left(rs!cIFTpo, 2)) = gTpoIFFuenteFinanciamiento Then
'            lnMontoTotal = rs!nSaldoCap - rs!nSaldoConcesion
'        Else
'            lnMontoTotal = rs!nSaldoCap + rs!nInteresPagado
'        End If
'        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
'            L.SubItems(4) = Format(rs!nSaldoCap * lnTasaVac, "#,#0.00")   'Saldo * la tasa vac
'        Else
'            If Not gsCodCMAC = "102" Then
'                L.SubItems(4) = Format(rs!nSaldoCap, "#,#0.00")
'            Else
'                L.SubItems(4) = Format(lnMontoTotal, "#,#0.00")
'            End If
'        End If
'        lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo + IIf(gsCodCMAC = "102", 1, 0), rs!nPeriodo, rs!nInteres, lnMontoTotal)
'        L.SubItems(12) = Format(lnInteres, "#0.00")
'        If lnInteres > 0 Then
'           L.SubItems(1) = "1"
'           L.Checked = True
'           L.ListSubItems.Item(2).ForeColor = vbBlue
'           L.ListSubItems.Item(3).ForeColor = vbBlue
'        Else
'            L.Checked = False
'            L.ListSubItems.Item(2).ForeColor = vbRed
'            L.ListSubItems.Item(3).ForeColor = vbRed
'        End If
'        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
'            If lnTasaVac > 0 Then
'                lnInteres = Format(lnInteres * lnTasaVac, "#0.00")
'            Else
'                lnInteres = Format(lnInteres, "#0.00")
'            End If
'        End If
'        L.SubItems(11) = Format(lnInteres, "#,#0.00")  'interes al cambio en soles si el pago es en dolares
'        L.SubItems(13) = Trim(rs!cMonedaPago)
'        L.SubItems(14) = rs!cIFTpo & "." & Trim(rs!cPersCod & "." & rs!cCtaIFCod)
'        L.SubItems(15) = rs!nSaldoCap
'        L.SubItems(16) = rs!dVencimiento
'        L.SubItems(17) = rs!nSaldoConcesion
'        rs.MoveNext
'    Loop
'    RSClose rs
'
'    PB.Visible = False
'    PB.value = 0

'edpyme - por la verificacion que exista vac
  Dim rs As ADODB.Recordset
    Dim n As Integer
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency
    Dim lnTotal As Integer, I As Integer
    Dim sCadena As String
    Dim oAdeud As New DCaja_Adeudados
    Dim oIF As New NCajaAdeudados
    Dim L As ListItem
    Dim lnSaldoTotal As Currency
    Dim oCaja As New NCajaAdeudados
    Dim lnTasaInteresReal As Currency
    Dim lnTasaVacTempo As Double
    Dim nPlazo As Integer
    lstCabecera.ListItems.Clear
    
    If optBuscar(0).value = True Then
        sCadena = " AND CIA.CPERSCOD='" & Mid(txtCodObjeto, 4, 13) & "'"
    ElseIf optBuscar(1).value = True Then
        sCadena = " AND CIA.CPERSCOD='" & Mid(txtBuscarCtaIF, 4, 13) & "' and CIA.cCtaIFCod='" & Right(txtBuscarCtaIF, 7) & "'"
    End If
  
    lnTasaVac = oAdeud.CargaIndiceVAC(ldFecha)
    If lnTasaVac = 0 Then
        If MsgBox("Tasa VAC no ha sido definida " & Chr(13) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            txtTasaVac.SetFocus
            Exit Sub
        End If
    End If
    txtTasaVac = lnTasaVac
    

    
    Set rs = oAdeud.GetAdeudadosProvision(gsOpeCod, txtFecha, Mid(gsOpeCod, 3, 1), sCadena)
    
    ' ==================== LLENAMOS MATRIZ ====================
     
    I = 0
    lstCabecera.ListItems.Clear
    PB.value = PB.Min
    If rs.BOF Then
    Else
        Do While Not rs.EOF
            
            'PB.value = PB.value + 1
            I = I + 1
            Status.Panels(1).Text = "Proceso " & Format(I * 100 / PB.Max, gsFormatoNumeroView) & "%"
            'Status.Panels(1).Text = "Proceso " & Format(PB.value * 100 / PB.Max, gsFormatoNumeroView) & "%"
            
    
            
            'I = I + 1
            
            Set L = lstCabecera.ListItems.Add(, , I)
            
            L.SubItems(2) = Trim(rs!cPersNombre)
            L.SubItems(3) = Trim(rs!cCtaIFDesc)
            L.SubItems(5) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
            L.SubItems(6) = Trim(rs!nNroCuota)
            L.SubItems(7) = Trim(rs!nCtaIFIntPeriodo)
            L.SubItems(8) = Trim(rs!nCtaIFIntValor)
            L.SubItems(9) = rs!nDiasUltPAgo
            nPlazo = rs!nCtaIFPlazo
            L.SubItems(23) = Format(rs!dCuotaUltModSaldos, "dd/MM/YYYY")
            L.SubItems(24) = rs!cTpoCuota
            If rs!nComisionMonto > 0 And rs!nDiasUltPAgo > 0 Then
            L.SubItems(25) = Round(SacarComisionPorDias(rs!nDiasUltPAgo, rs!nComisionMonto / rs!nCtaIFCuotas, nPlazo), 2)
            Else
            L.SubItems(25) = 0
            End If
            L.SubItems(27) = rs!nSaldoMes
            'If Val(Left(rs!cIFTpo, 2)) = gTpoIFFuenteFinanciamiento Then
                'SI ES COFIDE
                'Se asume que no se provisionara lo no concesional porque no se pagará
            '   lnSaldoTotal = rs!nSaldoCap - rs!nSaldoConcesion
            'Else
            lnSaldoTotal = rs!nSaldoCap '+ rs!nInteresPagado
            'End If
           
            'Intereses
            'If rs!nCtaIFIntPeriodo = 30 Then
            '    nTasaMensual = rs!nCtaIFIntValor
            'ElseIf rs!nCtaIFIntPeriodo = 360 Then
            '    nTasaMensual = oAdeud.GetInteresAnualAMensual(rs!nCtaIFIntValor)
            'End If
            
            'lnTasaInteresReal = Format(oAdeud.InteresReal(nTasaMensual, rs!nDiasUltPAgo), "0.000000")
            
            lnTasaInteresReal = Format(oCaja.InteresReal1(rs!nCtaIFIntValor, rs!nDiasUltPAgo, rs!nCtaIFIntPeriodo), "0.000000")
            
            'Interes REAL Sin Vac
            L.SubItems(11) = Format(lnSaldoTotal * lnTasaInteresReal, "0.00")
            'L.SubItems(11) = Format(oCaja.InteresReal1(rs!nCtaIFIntValor, rs!nDiasUltPAgo, rs!nCtaIFIntPeriodo), "0.000000")
            'Saldo de Capital REAL Sin Vac
            L.SubItems(12) = lnSaldoTotal
            
            'Saldo de Capital LP REAL Sin Vac
            L.SubItems(20) = Format(rs!nSaldoCapLP, "0.00")
            
            'VAC SI HUBIERA
            
            'Saldo de Capital
            If rs!cMonedaPago = "2" And Mid(rs!cCtaIfCod, 3, 1) = "1" Then
            
                'Interes x Vac Actual
                L.SubItems(10) = Format(lnSaldoTotal * lnTasaInteresReal * lnTasaVac, "0.00")
                
                'Saldo de Capital x VAC Actual
                L.SubItems(4) = Format(lnSaldoTotal * lnTasaVac, "0.00")
                
                'Saldo de Capital LP x VAC Actual
                L.SubItems(22) = Format(rs!nSaldoCapLP * lnTasaVac, "0.00")
                
                'VAC Anterior
                lnTasaVacTempo = oAdeud.CargaIndiceVAC(rs!dCuotaUltModSaldos)
                
                'Saldo de Capital x VAC Anterior
                L.SubItems(19) = Format(lnSaldoTotal * lnTasaVacTempo, "0.00")
                
                'Saldo de Capital LP x VAC Anterior
                L.SubItems(21) = Format(rs!nSaldoCapLP * lnTasaVacTempo, "0.00")
             
            Else
                
                'Interes
                L.SubItems(10) = Format(lnSaldoTotal * lnTasaInteresReal, "0.00")
                
                'Saldo de Capital
                L.SubItems(4) = Format(lnSaldoTotal, "0.00")
                
                'Saldo de Capital LP
                L.SubItems(22) = Format(rs!nSaldoCapLP, "0.00")
                
                'Saldo de Capital
                L.SubItems(19) = Format(lnSaldoTotal, "0.00")
                 
                'Saldo de Capital LP
                L.SubItems(21) = Format(rs!nSaldoCapLP, "0.00")
                
            End If
            
            L.SubItems(13) = Trim(rs!cMonedaPago)
            L.SubItems(14) = rs!cIFTpo & "." & Trim(rs!cPersCod & "." & rs!cCtaIfCod)
            L.SubItems(15) = Format(rs!dVencimiento, "dd/MM/YYYY")
            L.SubItems(16) = 0 'Habia Saldo Concesional
            L.SubItems(17) = rs!cCodLinCred
            L.SubItems(18) = rs!cDesLinCred
            
'            If lnTasaInteresReal > 0 Then
'               L.SubItems(1) = "1"
'               L.Checked = True
'               L.ListSubItems.Item(2).ForeColor = vbBlue
'               L.ListSubItems.Item(3).ForeColor = vbBlue
'            Else
'                L.Checked = False
'                L.ListSubItems.Item(2).ForeColor = vbRed
'                L.ListSubItems.Item(3).ForeColor = vbRed
'            End If
            
            rs.MoveNext
        Loop
        RSClose rs
    
'        cmdProvisionar.Enabled = True
    
    End If
    
    Set rs = Nothing
    Set oAdeud = Nothing
    
    
''    If rs.BOF Then
''        lnTotal = 0
''    Else
''        lnTotal = rs.RecordCount
''        PB.Visible = True
''        PB.Min = 0
''        PB.Max = rs.RecordCount
''        PB.value = 0
''    End If
''
''    i = 0
''
''    Do While Not rs.EOF
''
''        PB.value = PB.value + 1
''
''        i = i + 1
''
''''        Set L = lstCabecera.ListItems.Add(, , I)
''''        L.SubItems(2) = Trim(rs!cPersNombre)
''''        L.SubItems(3) = Trim(rs!cCtaIFDesc)
''''
''''        L.SubItems(5) = Trim(rs!nInteresPagado)  ' Interes acumulado pagado por cuota
''''        L.SubItems(6) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
''''        L.SubItems(7) = Trim(rs!nNroCuota)    ' numero de cuota pendiente
''''        L.SubItems(8) = Trim(rs!nPeriodo)
''''        L.SubItems(9) = Trim(rs!nInteres)
''''        L.SubItems(10) = Trim(rs!nDiasUltPAgo)
''''        If Val(Left(rs!cIFTpo, 2)) = gTpoIFFuenteFinanciamiento Then
''''            lnMontoTotal = rs!nSaldoCap - rs!nSaldoConcesion
''''        Else
''''            lnMontoTotal = rs!nSaldoCap + rs!nInteresPagado
''''        End If
''''        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
''''            L.SubItems(4) = Format(rs!nSaldoCap * lnTasaVac, "#,#0.00")   'Saldo * la tasa vac
''''        Else
''''            If Not gsCodCMAC = "102" Then
''''                L.SubItems(4) = Format(rs!nSaldoCap, "#,#0.00")
''''            Else
''''                L.SubItems(4) = Format(lnMontoTotal, "#,#0.00")
''''            End If
''''        End If
''''        lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo + IIf(gsCodCMAC = "102", 1, 0), rs!nPeriodo, rs!nInteres, lnMontoTotal)
''''        L.SubItems(12) = Format(lnInteres, "#0.00")
''''        If lnInteres > 0 Then
''''           L.SubItems(1) = "1"
''''           L.Checked = True
''''           L.ListSubItems.Item(2).ForeColor = vbBlue
''''           L.ListSubItems.Item(3).ForeColor = vbBlue
''''        Else
''''            L.Checked = False
''''            L.ListSubItems.Item(2).ForeColor = vbRed
''''            L.ListSubItems.Item(3).ForeColor = vbRed
''''        End If
''''        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
''''            If lnTasaVac > 0 Then
''''                lnInteres = Format(lnInteres * lnTasaVac, "#0.00")
''''            Else
''''                lnInteres = Format(lnInteres, "#0.00")
''''            End If
''''        End If
''''        L.SubItems(11) = Format(lnInteres, "#,#0.00")  'interes al cambio en soles si el pago es en dolares
''''        L.SubItems(13) = Trim(rs!cMonedaPago)
''''        L.SubItems(14) = rs!cIFTpo & "." & Trim(rs!cPersCod & "." & rs!cCtaIFCod)
''''        L.SubItems(15) = rs!nSaldoCap
''''        L.SubItems(16) = rs!dVencimiento
''''        L.SubItems(17) = rs!nSaldoConcesion
''
''            i = i + 1
''
''            Set L = lstCabecera.ListItems.Add(, , i)
''
''            L.SubItems(2) = Trim(rs!cPersNombre)
''            L.SubItems(3) = Trim(rs!cCtaIFDesc)
''            L.SubItems(5) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
''            L.SubItems(6) = Trim(rs!nNroCuota)
''            L.SubItems(7) = Trim(rs!nCtaIFIntPeriodo)
''            L.SubItems(8) = Trim(rs!nCtaIFIntValor)
''            L.SubItems(9) = rs!nDiasUltPAgo
''            L.SubItems(23) = Format(rs!dCuotaUltPago, "dd/MM/YYYY")
''            L.SubItems(24) = rs!cTpoCuota
''
''            lnSaldoTotal = rs!nSaldoCap '+ rs!nInteresPagado
''            'End If
''
''            'Intereses
''            'If rs!nCtaIFIntPeriodo = 30 Then
''            '    nTasaMensual = rs!nCtaIFIntValor
''            'ElseIf rs!nCtaIFIntPeriodo = 360 Then
''            '    nTasaMensual = oAdeud.GetInteresAnualAMensual(rs!nCtaIFIntValor)
''            'End If
''
''            'lnTasaInteresReal = Format(oAdeud.InteresReal(nTasaMensual, rs!nDiasUltPAgo), "0.000000")
''
''            lnTasaInteresReal = Format(oCaja.InteresReal1(rs!nCtaIFIntValor, rs!nDiasUltPAgo, rs!nCtaIFIntPeriodo), "0.000000")
''
''            'Interes REAL Sin Vac
''            L.SubItems(11) = Format(lnSaldoTotal * lnTasaInteresReal, "0.00")
''
''            'Saldo de Capital REAL Sin Vac
''            L.SubItems(12) = lnSaldoTotal
''
''            'Saldo de Capital LP REAL Sin Vac
''            L.SubItems(20) = Format(rs!nSaldoCapLP, "0.00")
''
''            'VAC SI HUBIERA
''
''            'Saldo de Capital
''            If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
''
''                'Interes x Vac Actual
''                L.SubItems(10) = Format(lnSaldoTotal * lnTasaInteresReal * lnTasaVac, "0.00")
''
''                'Saldo de Capital x VAC Actual
''                L.SubItems(4) = Format(lnSaldoTotal * lnTasaVac, "0.00")
''
''                'Saldo de Capital LP x VAC Actual
''                L.SubItems(22) = Format(rs!nSaldoCapLP * lnTasaVac, "0.00")
''
''                'VAC Anterior
''                lnTasaVacTempo = oAdeud.CargaIndiceVAC(rs!dCuotaUltPago)
''
''                'Saldo de Capital x VAC Anterior
''                L.SubItems(19) = Format(lnSaldoTotal * lnTasaVacTempo, "0.00")
''
''                'Saldo de Capital LP x VAC Anterior
''                L.SubItems(21) = Format(rs!nSaldoCapLP * lnTasaVacTempo, "0.00")
''
''            Else
''
''                'Interes
''                L.SubItems(10) = Format(lnSaldoTotal * lnTasaInteresReal, "0.00")
''
''                'Saldo de Capital
''                L.SubItems(4) = Format(lnSaldoTotal, "0.00")
''
''                'Saldo de Capital LP
''                L.SubItems(22) = Format(rs!nSaldoCapLP, "0.00")
''
''                'Saldo de Capital
''                L.SubItems(19) = Format(lnSaldoTotal, "0.00")
''
''                'Saldo de Capital LP
''                L.SubItems(21) = Format(rs!nSaldoCapLP, "0.00")
''
''            End If
''
''            L.SubItems(13) = Trim(rs!cMonedaPago)
''            L.SubItems(14) = rs!cIFTpo & "." & Trim(rs!cPersCod & "." & rs!cCtaIFCod)
''            L.SubItems(15) = Format(rs!dVencimiento, "dd/MM/YYYY")
''            L.SubItems(16) = 0 'Habia Saldo Concesional
''            L.SubItems(17) = rs!cCodLinCred
''            L.SubItems(18) = rs!cDesLinCred
''
''
''        rs.MoveNext
''    Loop
''    RSClose rs
    
    PB.Visible = False
    PB.value = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
    Set frmAdeudProv1 = Nothing
End Sub

Private Sub lstCabecera_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lnPeriodo As Long
    Dim lnTasaInt As Currency
    Dim lnMontoTotal  As Currency
    Dim lnInteres As Currency
    Dim lnDias    As Long
    Dim oAdeud As New DCaja_Adeudados
      
    Item.Selected = True
      
'    If Item.SubItems(1) = "" Then
'        Item.Checked = False
'        Exit Sub
'    End If
    
    If Val(Item.SubItems(9)) < 0 Then
        MsgBox "No puede provisionar días negativos", vbInformation, "Aviso"
        Item.Checked = False
    ElseIf Val(Item.SubItems(9)) = 0 Then
        MsgBox "No puede provisionar cero días", vbInformation, "Aviso"
        Item.Checked = False
    End If
         
         
    
'    If Item.SubItems(11) <= 0 Then
'        MsgBox "Monto no válido", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    If Val(Item.SubItems(10)) > 0 Then
'
'        If Val(Left(Item.SubItems(14), 2)) = gTpoIFFuenteFinanciamiento Then
'            lnMontoTotal = CCur(Item.SubItems(15)) - CCur(Item.SubItems(17))
'            '10500
'        Else
'            lnMontoTotal = CCur(Item.SubItems(15)) + CCur(Item.SubItems(5))
'        End If
'        lnTasaInt = CCur(Item.SubItems(9))
'        '4.13
'        lnPeriodo = Val(Item.SubItems(8))
'        '360
'        lnDias = Item.SubItems(10)
'        '4
'        lnInteres = oAdeud.CalculaInteres(lnDias, lnPeriodo, lnTasaInt, lnMontoTotal)
'        '4.7226
'        Item.SubItems(12) = Format(lnInteres, "#0.00")
'        '4.72
'        If Item.SubItems(13) = "2" And Mid(Item.SubItems(14), 20, 1) = "1" Then
'            Item.SubItems(11) = Format(lnInteres * lnTasaVac, "#,#0.00")
'        Else
'            Item.SubItems(11) = Format(lnInteres, "#,#0.00")
'        End If
'
'    End If
'
'    Set oAdeud = Nothing

End Sub

Private Sub optBuscar_Click(Index As Integer)
fraopciones.Visible = IIf(Index = 0, True, False)
FraGenerales.Visible = IIf(Index = 1, True, False)
Limpiar
BlanqueaTodo
End Sub


Private Sub BlanqueaTodo()
Dim I As Integer
    
    lstCabecera.ListItems.Clear
    txtBuscarCtaIF.Text = ""
    txtCodObjeto.Text = ""
  
End Sub



Private Sub optBuscar_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        txtCodObjeto.SetFocus
    ElseIf Index = 1 Then
        txtBuscarCtaIF.SetFocus
    Else
        cmdProcesar.SetFocus
    End If
End If

End Sub

Private Sub txtBuscarCtaIF_EmiteDatos()
If txtBuscarCtaIF <> "" Then
    Set frmAdeudCal = Nothing
    CargaDatosCuentas Mid(txtBuscarCtaIF, 4, 13), Mid(txtBuscarCtaIF, 1, 2), Mid(txtBuscarCtaIF, 18, 10)
    cmdProcesar.SetFocus
End If
End Sub

Sub Limpiar()
txtNroCtaIF = ""
lblDescIF = ""
txtNroCtaIF = ""
lblDescIF = ""
cboEstado.ListIndex = -1

txtCodObjeto = ""
lblObjDesc = ""
End Sub

Sub CargaDatosCuentas(psPersCod As String, pnIfTpo As CGTipoIF, psCtaIFCod As String)
Dim rs As ADODB.Recordset
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF

Set rs = New ADODB.Recordset
Limpiar
 
txtNroCtaIF = Trim(txtBuscarCtaIF.psDescripcion)
lblDescIF = oCtaIf.NombreIF(psPersCod)
 
Set rs = oCtaIf.GetDatosCtaIf(psPersCod, pnIfTpo, psCtaIFCod)
If Not rs.EOF And Not rs.EOF Then
    cboEstado = rs!cEstadoCons & Space(50) & rs!cCtaIFEstado
End If
RSClose rs
Set oCtaIf = Nothing

End Sub
Private Sub txtCodObjeto_EmiteDatos()
If txtCodObjeto <> "" Then
    lblObjDesc = txtCodObjeto.psDescripcion
    cmdProcesar.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    Dim oAdeud As New DCaja_Adeudados
    If KeyAscii = 13 Then
        If ValFecha(txtFecha) = True Then
            lnTasaVac = oAdeud.CargaIndiceVAC(txtFecha)
            txtTasaVac = Format(lnTasaVac, "#,###.00####")
            txtTasaVac.SetFocus
        End If
    End If
    Set oAdeud = Nothing
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtTasaVac_GotFocus()
    fEnfoque txtTasaVac
End Sub

Private Sub txtTasaVac_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTasaVac, KeyAscii, 15, 8)
    If KeyAscii = 13 Then
        If Me.optBuscar(0).value = True Then
            Me.optBuscar(0).SetFocus
        ElseIf Me.optBuscar(1).value = True Then
            Me.optBuscar(1).SetFocus
        ElseIf Me.optBuscar(2).value = True Then
            Me.optBuscar(2).SetFocus
        End If
    End If
End Sub



Private Sub TamanoColumnas()
    
    lstCabecera.ColumnHeaders(1).Width = 650
    lstCabecera.ColumnHeaders(2).Width = 200
    
    lstCabecera.ColumnHeaders(3).Width = 1500
    
    lstCabecera.ColumnHeaders(4).Text = "Pagaré"
    lstCabecera.ColumnHeaders(4).Width = 1900
    lstCabecera.ColumnHeaders(5).Width = 0 '1000 skxvac actual
    
    lstCabecera.ColumnHeaders(6).Text = "F.Ult.Pago"
    lstCabecera.ColumnHeaders(6).Width = 1100
    
    lstCabecera.ColumnHeaders(7).Text = "Cuota"
    lstCabecera.ColumnHeaders(7).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(7).Width = 650
    
    lstCabecera.ColumnHeaders(8).Text = "Per"
    lstCabecera.ColumnHeaders(8).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(8).Width = 650
    
    lstCabecera.ColumnHeaders(9).Text = "Int"
    lstCabecera.ColumnHeaders(9).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(9).Width = 600
    
    lstCabecera.ColumnHeaders(10).Text = "Dias"
    lstCabecera.ColumnHeaders(10).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(10).Width = 600
    
    lstCabecera.ColumnHeaders(11).Text = "Interes"
    lstCabecera.ColumnHeaders(11).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(11).Width = 1000
    
    lstCabecera.ColumnHeaders(12).Width = 1000 'Interes Real
    lstCabecera.ColumnHeaders(13).Width = 1000 'Saldo de Cap Real
    lstCabecera.ColumnHeaders(14).Width = 0 'Moneda
    lstCabecera.ColumnHeaders(15).Width = 0 '1000 Entidad Cod
    
    lstCabecera.ColumnHeaders(16).Text = "F. Vcto."
    lstCabecera.ColumnHeaders(16).Width = 1100
    
    lstCabecera.ColumnHeaders(17).Width = 0 'Concesional
    
    lstCabecera.ColumnHeaders(18).Text = "L.Cred"
    lstCabecera.ColumnHeaders(18).Width = 700
    
    lstCabecera.ColumnHeaders(19).Text = "L.Cred"
    lstCabecera.ColumnHeaders(19).Width = 2000
    
    lstCabecera.ColumnHeaders(20).Width = 0 '1000 SK x Vac
    lstCabecera.ColumnHeaders(21).Width = 0 ' 1000 KCLP Real
    lstCabecera.ColumnHeaders(22).Width = 0 ' 1000 KCLP xvac Ant
    lstCabecera.ColumnHeaders(23).Width = 0 ' 1000 KCLP xvac Act
    
    lstCabecera.ColumnHeaders(24).Text = "F.Ult.Act"
    lstCabecera.ColumnHeaders(24).Width = 1100
    
    lstCabecera.ColumnHeaders(25).Width = 0 '1000 cTpoCuota
    
    
End Sub

Private Sub txtTasaVac_LostFocus()
    If txtTasaVac = "" Then txtTasaVac = 0
End Sub

