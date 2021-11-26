VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmACGAdeuExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   Icon            =   "frmACGAdeuExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmACGAdeuExtorno.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
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
      Height          =   405
      Left            =   7500
      TabIndex        =   9
      Top             =   5250
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8880
      TabIndex        =   10
      Top             =   5250
      Width           =   1335
   End
   Begin VB.Frame fraCabecera 
      Height          =   975
      Left            =   60
      TabIndex        =   11
      Top             =   105
      Width           =   10125
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8760
         TabIndex        =   6
         Top             =   330
         Width           =   1245
      End
      Begin MSMask.MaskEdBox mskFechaMovimiento 
         Height          =   330
         Left            =   210
         TabIndex        =   0
         Top             =   405
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
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
         Height          =   795
         Left            =   3000
         TabIndex        =   18
         Top             =   120
         Width           =   4005
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
            Height          =   450
            Index           =   2
            Left            =   2670
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   270
            Value           =   -1  'True
            Width           =   1170
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
            Height          =   450
            Index           =   1
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   270
            Width           =   1170
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
            Height          =   450
            Index           =   0
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   270
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "========= Datos a Filtrar ========="
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   19
            Top             =   30
            Width           =   3720
         End
      End
      Begin VB.Frame fraIF 
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
         Left            =   105
         TabIndex        =   15
         Top             =   855
         Visible         =   0   'False
         Width           =   9915
         Begin Sicmact.TxtBuscar txtCodObjeto 
            Height          =   345
            Left            =   1680
            TabIndex        =   4
            Top             =   210
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
         Begin VB.Label lblDescIF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4320
            TabIndex        =   17
            Top             =   210
            Width           =   5520
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   240
            Width           =   1530
         End
      End
      Begin VB.Frame fraIFAdeud 
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
         Height          =   645
         Left            =   105
         TabIndex        =   12
         Top             =   855
         Visible         =   0   'False
         Width           =   9915
         Begin Sicmact.TxtBuscar txtCodObjetoIF 
            Height          =   345
            Left            =   840
            TabIndex        =   5
            Top             =   210
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
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblDescIFAdeud 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3480
            TabIndex        =   13
            Top             =   210
            Width           =   6360
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   1545
      End
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   195
      Left            =   1800
      TabIndex        =   26
      Top             =   5790
      Visible         =   0   'False
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   5700
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4065
      Left            =   60
      TabIndex        =   21
      Top             =   1140
      Width           =   10125
      Begin VB.TextBox txtMovDesc 
         Height          =   465
         Left            =   960
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3450
         Width           =   9030
      End
      Begin MSComctlLib.ListView lstCabecera 
         Height          =   2760
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   4868
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
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
         NumItems        =   42
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "cMovNro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "nMovNro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "cMovDesc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "cOpeCod"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "cOpeDesc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "cPersCod"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "cIFTpo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cCtaIFCod"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "cPersNombre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "cCtaIFDesc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "cTpoCuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "nNroCuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "nTipoEstadistica"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "cConsDescripcion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "nCapitalMovido"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "nInteres1Movido"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "nInteres2Movido"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "nComisionMovida"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "nSaldoAnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "nSaldoNew"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "nSaldoLPAnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "nSaldoLPNew"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "bVac"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "nCapitalMovidoReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "nInteres1MovidoReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "nInteres2MovidoReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "nComisionMovidaReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "nSaldoAntReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "nSaldoNewReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "nSaldoLPAntReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "nSaldoLPNewReal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "dFecUltActAnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "dFecUltActNew"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "dFecUltPagoAnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "dFecUltPagoNew"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "cCodLinCred"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "nNroCuota6"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "lsMovNro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "cEstadoCuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "dCancelacion"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         TabIndex        =   25
         Top             =   3045
         Width           =   315
      End
      Begin VB.Label lblCantidad 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   24
         Top             =   3015
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "** Para ver calendario de cuotas Presionar Doble Click"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6030
         TabIndex        =   23
         Top             =   3015
         Width           =   3900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   3540
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmACGAdeuExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipoCtaIf As CGTipoCtaIF
Dim gcOpeCod As String
Dim gcOpeCodBuscar As String
Dim cMoneda As String
Dim lnTasaVac As Double

Public Function Inicio(pcOpeCod As String)
gcOpeCod = pcOpeCod
cMoneda = Mid(pcOpeCod, 3, 1)

'
If gcOpeCod = "401711" Then 'Registro
    gcOpeCodBuscar = "401701"
ElseIf gcOpeCod = "401712" Then 'Confirmacion
    gcOpeCodBuscar = "401702"
ElseIf gcOpeCod = "401713" Then 'Provision
    gcOpeCodBuscar = "401705"
ElseIf gcOpeCod = "401714" Then 'Pago
    gcOpeCodBuscar = "401706"
End If

If gcOpeCod = "402711" Then 'Registro
    gcOpeCodBuscar = "402701"
ElseIf gcOpeCod = "402712" Then 'Confirmacion
    gcOpeCodBuscar = "402702"
ElseIf gcOpeCod = "402713" Then 'Provision
    gcOpeCodBuscar = "402705"
ElseIf gcOpeCod = "402714" Then 'Pago
    gcOpeCodBuscar = "402706"
End If
'

Me.Show 1
End Function
  
Private Sub cmdbuscar_Click()

    cmdExtornar.Enabled = False

    If ValidaBuscar Then
        Buscar
    End If
        
End Sub

Private Function ValidaBuscar() As Boolean

Dim oAdeud As New DACGAdeudados
Dim lnTasaVac As Double

    ValidaBuscar = True

    If optBuscar(0).value Then 'Institucion
        If Len(txtCodObjeto.Text) = 16 Then
             
        Else
            ValidaBuscar = False
            MsgBox "No hay institucion seleccionada", vbInformation, "Aviso"
            Exit Function
        End If
    ElseIf optBuscar(1).value = True Then 'Adeudado
        If Len(txtCodObjetoIF.Text) = 24 Then
             
        Else
            ValidaBuscar = False
            MsgBox "No hay Adeudado seleccionado", vbInformation, "Aviso"
            Exit Function
        End If
    ElseIf optBuscar(2).value = True Then 'Todo
    
    End If
 
    If ValFecha(mskFechaMovimiento) = True Then
    Else
        ValidaBuscar = False
        MsgBox "Ingrese una fecha correcta", vbInformation, "Aviso"
        Exit Function
    End If
     
End Function


Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()
    Dim oAdeudado As New DACGAdeudados
    Dim lnFilaActual As Long
    
    Dim oFun As New NContFunciones
    Dim lsMovNroExt As String
    Dim lbEliminaMov As Boolean
    
    'Consideraciones
    ' 1.    No Extornar meses cerrados contablemente
    
    If ValidaExtornar = True Then
    
        If MsgBox("Desea Extornar Registro?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            
            'Contable
            lnFilaActual = lstCabecera.SelectedItem.Index
            lsMovNroExt = lstCabecera.ListItems(lnFilaActual).SubItems(39)
            lbEliminaMov = oFun.PermiteModificarAsiento(lsMovNroExt, False)
   
            Set oFun = Nothing
            
            If Not lbEliminaMov Then
               If MsgBox("Fecha de Extorno corresponde a un mes ya Cerrado, ¿ Desea Extornar Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then Exit Sub
            End If
   
            If Left(lsMovNroExt, 6) <> Format(gdFecSis, "yyyymm") And lbEliminaMov Then
               If DateDiff("m", DateAdd("m", 1, CDate(Mid(lsMovNroExt, 7, 2) & "/" & Mid(lsMovNroExt, 5, 2) & "/" & Left(lsMovNroExt, 4))), gdFecSis) = 0 Then
                  If Day(gdFecSis) < 7 Then
                     lbEliminaMov = True
                  Else
                     lbEliminaMov = False
                  End If
               Else
                  lbEliminaMov = False
               End If
            End If
   
            'If chkGeneracionAsientos.value = 1 Then
            '   lbEliminaMov = False
            'End If
            
            'Fin Contable
            
            If oAdeudado.CGGrabaExtornoAdeudado(gcOpeCod, gcOpeCodBuscar, cMoneda, gdFecSis, lstCabecera, lnFilaActual, txtMovDesc.Text, Right(gsCodAge, 2), gsCodUser, lbEliminaMov) = 0 Then
                
                MsgBox "Extorno Efectuado Satisfactoriamente", vbInformation, "Aviso"
                
                lstCabecera.ListItems.Remove lnFilaActual
                lblCantidad.Caption = Val(lblCantidad.Caption) - 1
                
                If Val(lblCantidad.Caption) = 0 Then
                    cmdExtornar.Enabled = False
                Else
                    cmdExtornar.Enabled = True
                End If
                 
            Else
                MsgBox "No se pudo efectuar grabación de Provision de Adeudados", vbInformation, "Aviso"
            End If
        
        End If
        
    End If
End Sub

Private Function ValidaExtornar() As Boolean

ValidaExtornar = False

    If lstCabecera.ListItems.Count = 0 Then
        MsgBox "No hay elementos que extornar", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(txtMovDesc.Text)) = 0 Then
        txtMovDesc.SetFocus
        MsgBox "Ingrese una glosa", vbInformation, "Aviso"
        Exit Function
    End If
    
ValidaExtornar = True
        
End Function

 
Private Sub TamanoColumnas()
 
Dim I As Integer
    
    lstCabecera.ColumnHeaders(1).Width = 400
    lstCabecera.ColumnHeaders(2).Width = 200
    
    lstCabecera.ColumnHeaders(3).Text = "Movimiento"
    lstCabecera.ColumnHeaders(3).Width = 1500
     
    lstCabecera.ColumnHeaders(4).Width = 0
    
    lstCabecera.ColumnHeaders(5).Text = "Glosa"
    lstCabecera.ColumnHeaders(5).Width = 1500
    
    lstCabecera.ColumnHeaders(6).Width = 0
    lstCabecera.ColumnHeaders(7).Width = 0
    lstCabecera.ColumnHeaders(8).Width = 0
    lstCabecera.ColumnHeaders(9).Width = 0
    lstCabecera.ColumnHeaders(10).Width = 0
    
    lstCabecera.ColumnHeaders(11).Text = "Entidad"
    lstCabecera.ColumnHeaders(11).Width = 1500
    
    lstCabecera.ColumnHeaders(12).Text = "Pagaré"
    lstCabecera.ColumnHeaders(12).Width = 1500
    
    lstCabecera.ColumnHeaders(13).Width = 0
    
    lstCabecera.ColumnHeaders(14).Text = "Cuota"
    lstCabecera.ColumnHeaders(14).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(14).Width = 650
     
    lstCabecera.ColumnHeaders(15).Width = 0
    lstCabecera.ColumnHeaders(16).Width = 0
    
    lstCabecera.ColumnHeaders(17).Text = "Capital"
    lstCabecera.ColumnHeaders(17).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(17).Width = 1000
    
    lstCabecera.ColumnHeaders(18).Text = "Int.Prov"
    lstCabecera.ColumnHeaders(18).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(18).Width = 1000
    
    lstCabecera.ColumnHeaders(19).Text = "Int.Adic"
    lstCabecera.ColumnHeaders(19).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(19).Width = 1000
    
    lstCabecera.ColumnHeaders(20).Text = "Comisión"
    lstCabecera.ColumnHeaders(20).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(20).Width = 1000
 
    For I = 21 To 42
        lstCabecera.ColumnHeaders(I).Width = 0
    Next
       
End Sub

Private Sub Form_Load()
         
    Dim oOpe As New DOperacion
    
    lnTipoCtaIf = gTpoCtaIFCtaAdeud
    
    CentraForm frmACGAdeuExtorno
    
    If gcOpeCod = "401711" Or gcOpeCod = "402711" Then
        Me.Caption = "EXTORNO DE REGISTRO DE PAGARES DE ADEUDADOS M" & IIf(cMoneda = "1", "N", "E")
    ElseIf gcOpeCod = "401712" Or gcOpeCod = "402712" Then
        Me.Caption = "EXTORNO DE CONFIRMACION DE PAGARES DE ADEUDADOS M" & IIf(cMoneda = "1", "N", "E")
    ElseIf gcOpeCod = "401713" Or gcOpeCod = "402713" Then
        Me.Caption = "EXTORNO DE PROVISIÓN DE ADEUDADOS M" & IIf(cMoneda = "1", "N", "E")
    ElseIf gcOpeCod = "401714" Or gcOpeCod = "402714" Then
        Me.Caption = "EXTORNO DE PAGO DE ADEUDADOS M" & IIf(cMoneda = "1", "N", "E")
    End If
    
    'Instituciones Financieras
    txtCodObjeto.psRaiz = "Instituciones Financieras"
    txtCodObjeto.rs = oOpe.GetOpeObj(gcOpeCod, "1")
      
    'Adeudados
    txtCodObjetoIF.psRaiz = "Adeudados"
    txtCodObjetoIF.rs = oOpe.GetOpeObj(gcOpeCod, "2")

    mskFechaMovimiento.Text = Format(gdFecSis, "dd/MM/YYYY")
    
    CambiaTamaño True
    
    TamanoColumnas
     
End Sub
 
Private Sub Buscar()
    
Dim oAdeud As New DACGAdeudados
Dim rs As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim sCadena As String
Dim bCancelacion As Boolean
Dim dCancelacion As Date

Dim L As ListItem

Dim lnTotal As Long
Dim I As Long

    lstCabecera.ListItems.Clear
    
    If optBuscar(0).value = True Then
        sCadena = " AND CIF.cPersCod='" & Mid(txtCodObjeto, 4, 13) & "'"
    ElseIf optBuscar(1).value = True Then
        sCadena = " AND CIF.cPersCod='" & Mid(txtCodObjetoIF, 4, 13) & "' and CIF.cCtaIFCod='" & Right(txtCodObjetoIF, 7) & "'"
    End If
     
    Set rs = oAdeud.GetUltimosMovimientosExtorno(gcOpeCodBuscar, sCadena, Format(mskFechaMovimiento.Text, "YYYYMMdd"))
    
    If rs.BOF Then
        lnTotal = 0
    Else
        lnTotal = rs.RecordCount
        PB.Visible = True
        PB.Min = 0
        PB.Max = rs.RecordCount
        PB.value = 0
    End If
     
    lblCantidad.Caption = lnTotal
     
    ' ==================== LLENAMOS MATRIZ ====================
     
    I = 0
    lstCabecera.ListItems.Clear
    
    If rs.BOF Then
        MsgBox " Datos no encontrados ", vbInformation, "Aviso"
        Me.mskFechaMovimiento.SetFocus
        Exit Sub
    Else
        Do While Not rs.EOF
            
            PB.value = PB.value + 1
            Status.Panels(1).Text = "Proceso " & Format(PB.value * 100 / PB.Max, gsFormatoNumeroView) & "%"
    
            'Cancelacion
            If Format(rs!dCancelacion, "YYYYMMdd") <> "19000101" Then
                bCancelacion = True
                dCancelacion = rs!dCancelacion
            Else
                bCancelacion = False
                dCancelacion = "01/01/1900"
            End If
            
            I = I + 1
            
            Set L = lstCabecera.ListItems.Add(, , I)
            L.SmallIcon = 1
            L.SubItems(2) = Trim(rs!cMovNro)
            L.SubItems(3) = rs!nMovNro
            L.SubItems(4) = Trim(rs!cMovDesc)
            L.SubItems(5) = Trim(rs!cOpeCod)
            L.SubItems(6) = Trim(rs!cOpeDesc)
            L.SubItems(7) = Trim(rs!cPerscod)
            L.SubItems(8) = Trim(rs!cIFTpo)
            L.SubItems(9) = Trim(rs!cCtaIFCod)
            L.SubItems(10) = Trim(rs!cPersNombre)
            L.SubItems(11) = Trim(rs!cCtaIFDesc)
            L.SubItems(12) = Trim(rs!cTpoCuota)
            L.SubItems(13) = rs!nNroCuota
            L.SubItems(14) = rs!nTipoEstadistica
            L.SubItems(15) = Trim(rs!cConsDescripcion)
            L.SubItems(16) = Format(rs!nCapitalMovido, "0.00")
            L.SubItems(17) = Format(rs!nInteres1Movido, "0.00")
            L.SubItems(18) = Format(rs!nInteres2Movido, "0.00")
            L.SubItems(19) = Format(rs!nComisionMovida, "0.00")
            L.SubItems(20) = Format(rs!nSaldoAnt, "0.00")
            L.SubItems(21) = Format(rs!nSaldoNew, "0.00")
            L.SubItems(22) = Format(rs!nSaldoLPAnt, "0.00")
            L.SubItems(23) = Format(rs!nSaldoLPNew, "0.00")
            L.SubItems(24) = rs!bVac
            L.SubItems(25) = Format(rs!nCapitalMovidoREAL, "0.00")
            L.SubItems(26) = Format(rs!nInteres1MovidoREAL, "0.00")
            L.SubItems(27) = Format(rs!nInteres2MovidoREAL, "0.00")
            L.SubItems(28) = Format(rs!nComisionMovidaREAL, "0.00")
            L.SubItems(29) = Format(rs!nSaldoAntREAL, "0.00")
            L.SubItems(30) = Format(rs!nSaldoNewREAL, "0.00")
            L.SubItems(31) = Format(rs!nSaldoLPAntREAL, "0.00")
            L.SubItems(32) = Format(rs!nSaldoLPNewREAL, "0.00")
            L.SubItems(33) = Format(rs!dFecUltActAnt, "dd/MM/YYYY")
            L.SubItems(34) = Format(rs!dFecUltActNew, "dd/MM/YYYY")
            L.SubItems(35) = Format(rs!dFecUltPagoAnt, "dd/MM/YYYY")
            L.SubItems(36) = Format(rs!dFecUltPagoNew, "dd/MM/YYYY")
            L.SubItems(37) = Trim(rs!cCodLinCred)
            
            'Solo en Extorno de Pagos y Provisiones Relacionar con Calendario
            If gcOpeCodBuscar = "401705" Or gcOpeCodBuscar = "402705" Or gcOpeCodBuscar = "401706" Or gcOpeCodBuscar = "402706" Then

                'Cuota 6
                Set rsTemp = oAdeud.GetDatosCuota6(rs!cPerscod, rs!cIFTpo, rs!cCtaIFCod, Format(rs!dVencimiento, "MM/dd/YYYY"), "", , bCancelacion)
                If rsTemp.BOF Then
                    L.SubItems(38) = ""
                Else
                    L.SubItems(38) = rsTemp!nNroCuota
                End If
                rsTemp.Close
                
            End If
            
            L.SubItems(39) = Trim(rs!cMovNro) 'del nMovNro
            L.SubItems(40) = Trim(rs!cEstAnt)
            L.SubItems(41) = Format(rs!dCancelacion, "dd/MM/YYYY")
            
            rs.MoveNext
        Loop
        
        rs.Close
        cmdExtornar.Enabled = True
    End If
    
    Set rs = Nothing
    Set oAdeud = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmACGAdeuExtorno = Nothing
End Sub

Private Sub lstCabecera_DblClick()
Dim lscPersCod As String
Dim lscIFTpo As String
Dim lscCtaIFCod As String

If lstCabecera.ListItems.Count > 0 Then
    lscPersCod = Trim(lstCabecera.SelectedItem.SubItems(7))
    lscIFTpo = Trim(lstCabecera.SelectedItem.SubItems(8))
    lscCtaIFCod = Trim(lstCabecera.SelectedItem.SubItems(9))
    frmACGAdeuCalRap.CargarDatos lscPersCod, lscIFTpo, lscCtaIFCod
End If
     
End Sub

Private Sub mskFechaMovimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optBuscar(2).SetFocus
    End If
End Sub

Private Sub optBuscar_Click(Index As Integer)
    If Index = 2 Then
        CambiaTamaño True
    Else
        CambiaTamaño False
    End If

    fraIF.Visible = IIf(Index = 0, True, False)
    fraIFAdeud.Visible = IIf(Index = 1, True, False)
    
    txtCodObjeto.Text = ""
    txtCodObjetoIF.Text = ""
    lblDescIF.Caption = ""
    lblDescIFAdeud.Caption = ""
    
    cmdExtornar.Enabled = False
    
End Sub

Private Sub optBuscar_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 0 Then
            txtCodObjeto.SetFocus
        ElseIf Index = 1 Then
            txtCodObjetoIF.SetFocus
        Else
            cmdBuscar.SetFocus
        End If
    End If

End Sub

Private Sub txtCodObjeto_EmiteDatos()

If txtCodObjeto <> "" Then
    lblDescIF.Caption = txtCodObjeto.psDescripcion
    cmdBuscar.SetFocus
Else
    lblDescIF.Caption = ""
End If

End Sub

Private Sub txtCodObjetoIF_EmiteDatos()

Dim lscPersCod As String
Dim lscIFTpo As String
Dim lscCtaIFCod As String

If txtCodObjetoIF <> "" Then
    lblDescIFAdeud.Caption = txtCodObjetoIF.psDescripcion
    cmdBuscar.SetFocus
End If

End Sub

Private Sub txtMovDesc_GotFocus()
   fEnfoque txtMovDesc
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        cmdExtornar.SetFocus
    End If
End Sub

Private Sub CambiaTamaño(pbTipo As Boolean)
    If pbTipo = True Then
        Me.fraCabecera.Height = 975
        Me.Frame1.Top = 1140
        Me.cmdExtornar.Top = 5250
        Me.cmdCerrar.Top = 5250
        Me.PB.Top = 5790
        Me.Height = 6390
    Else
        Me.fraCabecera.Height = 1605
        Me.Frame1.Top = 1710
        Me.cmdExtornar.Top = 5820
        Me.cmdCerrar.Top = 5820
        Me.PB.Top = 6360
        Me.Height = 7005
    End If
End Sub
