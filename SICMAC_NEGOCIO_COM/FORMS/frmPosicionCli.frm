VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPosicionCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posición del Cliente"
   ClientHeight    =   7560
   ClientLeft      =   960
   ClientTop       =   1860
   ClientWidth     =   10200
   Icon            =   "frmPosicionCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar PbCargaPosicion 
      Height          =   255
      Left            =   2280
      TabIndex        =   42
      Top             =   7320
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdEstadoCuenta 
      Caption         =   "&Estado Cta. Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CommandButton CmdHistorial 
      Caption         =   "Historial Creditos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      TabIndex        =   35
      Top             =   6720
      Width           =   1845
   End
   Begin VB.Frame fraDetalle 
      Height          =   5475
      Left            =   105
      TabIndex        =   14
      Top             =   1185
      Width           =   9915
      Begin RichTextLib.RichTextBox rtfComentario 
         Height          =   855
         Left            =   1185
         TabIndex        =   16
         Top             =   4500
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmPosicionCli.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   8775
         Picture         =   "frmPosicionCli.frx":0386
         ScaleHeight     =   240
         ScaleWidth      =   705
         TabIndex        =   15
         Top             =   270
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar Analistas"
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
         Left            =   120
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   1965
      End
      Begin TabDlg.SSTab TabPosicion 
         Height          =   3960
         Left            =   105
         TabIndex        =   19
         Top             =   495
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   6985
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   617
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Créditos   "
         TabPicture(0)   =   "frmPosicionCli.frx":E51C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblClientePreferncial"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblSegRiesgo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblSegExperianExt"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LstCreditos"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   " &Ahorros   "
         TabPicture(1)   =   "frmPosicionCli.frx":E538
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label6"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lblSolesAho"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lblDolaresAho"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lstAhorros"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "&Pignoraticio   "
         TabPicture(2)   =   "frmPosicionCli.frx":E554
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label7"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label8"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "lblSolesPig"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "lblDolaresPig"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "lblSegPredExt"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "lstPrendario"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).ControlCount=   7
         TabCaption(3)   =   " Créditos Judiciales   "
         TabPicture(3)   =   "frmPosicionCli.frx":E570
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lstJudicial"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Carta Fianza"
         TabPicture(4)   =   "frmPosicionCli.frx":E58C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "LstCartaFianza"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Sist. Financiero"
         TabPicture(5)   =   "frmPosicionCli.frx":E5A8
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "lstFinanciero"
         Tab(5).ControlCount=   1
         Begin MSComctlLib.ListView LstCreditos 
            Height          =   2790
            Left            =   135
            TabIndex        =   20
            Top             =   495
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   4921
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   26
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha Desembolso"
               Object.Width           =   2170
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Nro. Crédito"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Agencia"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Sub Tipo Crédito"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Estado"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Participación"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Analista"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Analista Inicial"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Nota"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Monto "
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Text            =   "Saldo Cap."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Cod. Ant 1"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Cod. Ant 2"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Moneda"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   15
               Text            =   "Reprogramado"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "FechaCancelacion"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   17
               Text            =   "Fec.Solicitud"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Text            =   "Dias Atraso Acum."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   19
               Text            =   "Calificación"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   20
               Text            =   "Condicion"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   21
               Text            =   "Tipo de Producto"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   22
               Text            =   "Refinanciado a vigente"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   23
               Text            =   "Fecha Cambio Ref-Vig"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   24
               Text            =   "Campana"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   25
               Text            =   "Tipo Crédito"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstAhorros 
            Height          =   2790
            Left            =   -74910
            TabIndex        =   21
            Top             =   495
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   4921
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Producto"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Agencia"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Nro. Cuenta"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Nro. Cta Antigua"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Estado"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Participación"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "SaldoCont"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "SaldoDisp"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Bloqueo Parcial"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Motivo de Bloqueo"
               Object.Width           =   7231
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Moneda"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstPrendario 
            Height          =   2775
            Left            =   -74910
            TabIndex        =   22
            Top             =   495
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   2381
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Agencia"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Nro. Cuenta"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Estado"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Participación"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Nº Renovación"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Monto Prestado"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Fecha Vencimiento"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "CodAgencia"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Saldo Capital"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Tasacion"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Moneda"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstJudicial 
            Height          =   2775
            Left            =   -74925
            TabIndex        =   23
            Top             =   495
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha"
               Object.Width           =   2170
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Nro. Crédito"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Agencia"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Tipo Crédito"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Estado"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Participación"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Saldo Cap."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Reprogramado"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Fecha Castigado"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Ult Fecha Pago"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Tipo Producto"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView LstCartaFianza 
            Height          =   2790
            Left            =   -74940
            TabIndex        =   24
            Top             =   495
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   4921
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Emitida"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Nro. Crédito"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Agencia"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Estado"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Participacion"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Analista"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Monto"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Moneda"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Honrada"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView lstFinanciero 
            Height          =   2790
            Left            =   -74940
            TabIndex        =   36
            Top             =   495
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   4921
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   2469
            EndProperty
         End
         Begin VB.Label lblSegExperianExt 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            TabIndex        =   41
            Top             =   3360
            Width           =   3255
         End
         Begin VB.Label lblSegPredExt 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -74880
            TabIndex        =   40
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label lblSegRiesgo 
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   3360
            Width           =   2775
         End
         Begin VB.Label lblClientePreferncial 
            Caption         =   "CLIENTE PREFERENCIAL"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6120
            TabIndex        =   38
            Top             =   3360
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label lblDolaresPig 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Height          =   285
            Left            =   -67665
            TabIndex        =   34
            Top             =   3600
            Width           =   2145
         End
         Begin VB.Label lblSolesPig 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   -67665
            TabIndex        =   33
            Top             =   3280
            Width           =   2145
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "DOLARES"
            Height          =   195
            Left            =   -68460
            TabIndex        =   32
            Top             =   3600
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "SOLES"
            Height          =   195
            Left            =   -68460
            TabIndex        =   31
            Top             =   3360
            Width           =   525
         End
         Begin VB.Label lblDolaresAho 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Height          =   285
            Left            =   -67680
            TabIndex        =   30
            Top             =   3375
            Width           =   2145
         End
         Begin VB.Label lblSolesAho 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   -70815
            TabIndex        =   29
            Top             =   3375
            Width           =   2145
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "DOLARES"
            Height          =   195
            Left            =   -68475
            TabIndex        =   28
            Top             =   3465
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL PIGNORATICIO"
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
            Left            =   -70680
            TabIndex        =   27
            Top             =   3480
            Width           =   2010
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "TOTAL AHORROS"
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
            Left            =   -73185
            TabIndex        =   26
            Top             =   3465
            Width           =   1590
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "SOLES"
            Height          =   195
            Left            =   -71445
            TabIndex        =   25
            Top             =   3465
            Width           =   525
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios :"
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
         Left            =   105
         TabIndex        =   17
         Top             =   4485
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdSaldosConsol 
      Caption         =   "&Saldos Consolidados"
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
      Height          =   390
      Left            =   5730
      TabIndex        =   13
      Top             =   6675
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Height          =   390
      Left            =   7860
      TabIndex        =   8
      Top             =   6675
      Width           =   1020
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8940
      TabIndex        =   1
      Top             =   6675
      Width           =   1020
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   480
      Left            =   810
      TabIndex        =   10
      Top             =   4815
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   847
      _Version        =   393217
      RightMargin     =   35000
      TextRTF         =   $"frmPosicionCli.frx":E5C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar barraestado 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   9
      Top             =   7290
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10583
            Picture         =   "frmPosicionCli.frx":E644
            Text            =   "Seleccione el Cliente "
            TextSave        =   "Seleccione el Cliente "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8467
            MinWidth        =   8467
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   8820
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
         Height          =   345
         Left            =   7530
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label LblPersCod 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   870
         TabIndex        =   11
         Top             =   262
         Width           =   1755
      End
      Begin VB.Label lblDocJur 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label lblDocNat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1395
         TabIndex        =   6
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2760
         TabIndex        =   5
         Top             =   262
         Width           =   4620
      End
      Begin VB.Label lblDocJuridico 
         AutoSize        =   -1  'True
         Caption         =   "RUC :"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   660
         Width           =   435
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad :"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmPosicionCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLima As Integer
Dim nSaldosAho As Integer
Dim fbPersNatural As Boolean
Dim sCalifiFinal As String
Dim lsPersTDoc As String 'ALPA 20100922
Dim bPermisoCargo As Boolean 'AMDO 20130726 TI-ERS086-2013

'Private Function GeneraImpresion() As String
'Dim i As Integer
'Dim ContLineas As Integer
'Dim R As ADODB.Recordset
'Dim oCred As COMDCredito.DCOMCredito
'
'    GeneraImpresion = ""
'    GeneraImpresion = oImpresora.gPrnSaltoLinea
'
'    GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'
'    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CREDITOS : " & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Nota", 8) & ImpreFormat("Monto", 10) & ImpreFormat("Saldo Cap", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    ContLineas = 12
'    If lstCreditos.ListItems.Count > 0 Then
'
'        For i = 1 To lstCreditos.ListItems.Count
'            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstCreditos.ListItems(i).Text, 5)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(1), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(4), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstCreditos.ListItems(i).SubItems(3)), "AGENCIA", ""), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(2), 20)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(5), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(6), 9)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(7), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(8), 3)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstCreditos.ListItems(i).SubItems(9)), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstCreditos.ListItems(i).SubItems(10)), 10)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstCreditos.ListItems(i).SubItems(13), 8) & oImpresora.gPrnSaltoLinea
'            ContLineas = ContLineas + 1
'
'            If ContLineas > 56 Then
'                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CREDITOS : " & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Nota", 8) & ImpreFormat("Monto", 10) & ImpreFormat("Saldo Cap", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                ContLineas = 13
'            End If
'
'        Next i
'
'        '--------------GARANTIAS----------------------
'        Set R = New ADODB.Recordset
'        Set oCred = New COMDCredito.DCOMCredito
'        GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'        GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE GARANTIAS : " & oImpresora.gPrnSaltoLinea
'        GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'        GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("MONTO REALIZ.", 13) & ImpreFormat("MONTO DISPON.", 13) & oImpresora.gPrnSaltoLinea
'        GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 5
'            Set R = oCred.RecuperaGarantiasCreditoConsol(lblPersCod.Caption, gdFecSis)
'            Do While Not R.EOF
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(R!nRealizacion, 8, , True)
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(R!nPorGravar, 8, , True)
'                ContLineas = ContLineas + 1
'                If ContLineas > 56 Then
'                    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                    GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE GARANTIAS : " & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("MONTO REALIZ.", 13) & ImpreFormat("MONTO DISPON.", 13) & oImpresora.gPrnSaltoLinea
'                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                    ContLineas = 13
'                End If
'                R.MoveNext
'            Loop
'        Set oCred = Nothing
'        Set R = Nothing
'
'    Else
'        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas de Creditos" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 2
'    End If
'
'    '********** A H O R R O S **********************
'    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE AHORROS : " & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Producto", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10)
'    If nSaldosAho = 1 Then
'        GeneraImpresion = GeneraImpresion & ImpreFormat("Saldo Cont", 10) & Space(4) & ImpreFormat("Saldo Disp", 10)
'    End If
'    GeneraImpresion = GeneraImpresion & ImpreFormat("Motiv. Bloqueo", 15) & Space(4) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'
'    ContLineas = ContLineas + 6
'    If lstAhorros.ListItems.Count > 0 Then
'        For i = 1 To lstAhorros.ListItems.Count
'            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstAhorros.ListItems(i).Text, 5)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(1), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(2), 10)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstAhorros.ListItems(i).SubItems(3)), "AGENCIA", ""), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(4), 20)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(6), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(7), 12)
'            If nSaldosAho = 1 Then
'                GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(8), 12, 2)
'                GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(9), 12, 2)
'            End If
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(10), 17)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(i).SubItems(11), 10) & oImpresora.gPrnSaltoLinea
'            ContLineas = ContLineas + 1
'            If ContLineas > 56 Then
'                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE AHORROS : " & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Producto", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10)
'                If nSaldosAho = 1 Then
'                    GeneraImpresion = GeneraImpresion & ImpreFormat("SaldoCont", 10) & Space(4) & ImpreFormat("Saldo Dispon", 8)
'                End If
'                GeneraImpresion = GeneraImpresion & ImpreFormat("Motiv. Bloqueo", 15) & Space(4) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                ContLineas = 13
'            End If
'        Next i
'
'        GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'        GeneraImpresion = GeneraImpresion & Space(5) & "TOTAL AHORROS:   SOLES=" & Trim(lblSolesAho.Caption) & Space(10 - Len(lblSolesAho.Caption)) & " DOLARES=" & Trim(lblDolaresAho.Caption) & oImpresora.gPrnSaltoLinea
'
'    Else
'        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas de Ahorros" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 2
'    End If
'
'    'Para Pignoraticio
'    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE PIGNORATICIO : " & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("No Renov.", 10) & ImpreFormat("Monto", 6) & ImpreFormat("Fecha Venc", 12) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Tasación", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    ContLineas = ContLineas + 6
'    If lstPrendario.ListItems.Count > 0 Then
'        For i = 1 To lstPrendario.ListItems.Count
'            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstPrendario.ListItems(i).Text, 5)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(1), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstPrendario.ListItems(i).SubItems(2)), "AGENCIA", ""), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(3), 20)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(4), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(5), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(6), 3)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstPrendario.ListItems(i).SubItems(7)), 10)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(8), 11)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstPrendario.ListItems(i).SubItems(10)), 10)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstPrendario.ListItems(i).SubItems(11)), 10)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(i).SubItems(12), 11) & oImpresora.gPrnSaltoLinea
'
'            ContLineas = ContLineas + 1
'            If ContLineas > 56 Then
'                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE PIGNORATICIO : " & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 10) & ImpreFormat("Estado", 10)
'                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("No Renov.", 10) & ImpreFormat("Monto", 8) & ImpreFormat("Fecha Venc", 10) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Tasación", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                ContLineas = 13
'            End If
'        Next i
'        GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'        GeneraImpresion = GeneraImpresion & Space(5) & "TOTAL PIGNORATICIO:   SOLES=" & Trim(lblSolesPig.Caption) & Space(10 - Len(lblSolesPig.Caption)) & " DOLARES=" & Trim(lblDolaresPig.Caption) & oImpresora.gPrnSaltoLinea
'    Else
'        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas de Pignoraticio" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 2
'    End If
'
'    'Judicial
'    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE JUDICIAL : " & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 13)
'    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Fecha Cast.", 10) & ImpreFormat("Fecha Ult. Pago", 10) & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    ContLineas = ContLineas + 4
'
'    If lstJudicial.ListItems.Count > 0 Then
'        For i = 1 To lstJudicial.ListItems.Count
'            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstJudicial.ListItems(i).Text, 5)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(1), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(4), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstJudicial.ListItems(i).SubItems(3)), "AGENCIA", ""), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(2), 20)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(5), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(6), 9)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstJudicial.ListItems(i).SubItems(7)), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(8), 11)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(i).SubItems(9), 10) & oImpresora.gPrnSaltoLinea
'            ContLineas = ContLineas + 1
'            If ContLineas > 56 Then
'                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE JUDICIAL : " & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 10) & ImpreFormat("Estado", 10)
'                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Fecha Cast.", 10) & ImpreFormat("Fecha Ult. Pago", 10) & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                ContLineas = 13
'            End If
'        Next i
'    Else
'        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas en Judicial" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 2
'    End If
'
'    '***************** CARTA FIANZA ************************
'    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CARTA FIANZA : " & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Monto", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    ContLineas = ContLineas + 6
'    If LstCartaFianza.ListItems.Count > 0 Then
'
'        For i = 1 To LstCartaFianza.ListItems.Count
'            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(LstCartaFianza.ListItems(i).Text, 5)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(i).SubItems(1), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(LstCartaFianza.ListItems(i).SubItems(3)), "AGENCIA", ""), 8)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(i).SubItems(2), 20)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(i).SubItems(4), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(i).SubItems(5), 12)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(i).SubItems(6), 3)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(LstCartaFianza.ListItems(i).SubItems(7)), 10)
'            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(i).SubItems(8), 8) & oImpresora.gPrnSaltoLinea
'            ContLineas = ContLineas + 1
'
'            If ContLineas > 56 Then
'                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CARTA FIANZA : " & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
'                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Monto", 15) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                ContLineas = 13
'            End If
'
'        Next i
'    Else
'        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cartas Fianza" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 2
'    End If
'
'    '***************** COMENTARIOS ************************
'    Dim sCadena As String, sLinea As String
'    Dim nPos As Integer
'    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE COMENTARIO : " & oImpresora.gPrnSaltoLinea
'    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'    ContLineas = ContLineas + 5
'    sCadena = Trim(rtfComentario.Text)
'    If sCadena <> "" Then
'        Do
'            nPos = InStr(1, sCadena, oImpresora.gPrnSaltoLinea, vbTextCompare)
'            If nPos <> 0 Then
'                sLinea = Mid(sCadena, 1, nPos)
'                sCadena = Mid(sCadena, nPos + 1, Len(sCadena) - nPos)
'            Else
'                sLinea = sCadena
'                sCadena = ""
'            End If
'            GeneraImpresion = GeneraImpresion & Space(5) & sLinea
'
'            ContLineas = ContLineas + 1
'
'            If ContLineas > 56 Then
'                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & lblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocnat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE COMENTARIOS : " & oImpresora.gPrnSaltoLinea
'                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
'                ContLineas = 11
'            End If
'
'        Loop Until sCadena = ""
'    Else
'        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Comentarios" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        ContLineas = ContLineas + 2
'    End If
'End Function

Private Function GeneraImpresion() As String
Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim MatCreditos As Variant
Dim MatAhorros As Variant
Dim MatPrendario As Variant
Dim MatJudicial As Variant
Dim MatCartaFianza As Variant
Dim bPreferencial As Boolean 'FRHU RQ13790 20140106

Dim i As Integer

Dim MatSistFinanc As Variant
Dim J As Integer

'INICIO ORCR-20140913*********
'ReDim MatCreditos(LstCreditos.ListItems.Count, 14)
'ReDim MatCreditos(LstCreditos.ListItems.count, 15)'Comento Segmentacion Riesgo por cliente JOEP20201113
ReDim MatCreditos(lstCreditos.ListItems.count, 17) 'Add Segmentacion Riesgo por cliente JOEP20201113

For i = 1 To lstCreditos.ListItems.count
    MatCreditos(i, 1) = lstCreditos.ListItems(i).Text
    MatCreditos(i, 2) = lstCreditos.ListItems(i).SubItems(1)
    MatCreditos(i, 3) = lstCreditos.ListItems(i).SubItems(4)
    MatCreditos(i, 4) = lstCreditos.ListItems(i).SubItems(3)
    MatCreditos(i, 5) = lstCreditos.ListItems(i).SubItems(2)
    MatCreditos(i, 6) = lstCreditos.ListItems(i).SubItems(5)
    MatCreditos(i, 7) = lstCreditos.ListItems(i).SubItems(6)
    MatCreditos(i, 8) = lstCreditos.ListItems(i).SubItems(7)
    
    'MAVM 20090914 ***
    MatCreditos(i, 9) = lstCreditos.ListItems(i).SubItems(8)
    'MAVM 20090914 ***
    
    MatCreditos(i, 10) = lstCreditos.ListItems(i).SubItems(9)
    MatCreditos(i, 11) = lstCreditos.ListItems(i).SubItems(10)
    MatCreditos(i, 12) = lstCreditos.ListItems(i).SubItems(11)
    MatCreditos(i, 13) = lstCreditos.ListItems(i).SubItems(14)
    '--------------------------------
    'MatCreditos(i, 14) = LstCreditos.ListItems(i).SubItems(18)
    MatCreditos(i, 14) = lstCreditos.ListItems(i).SubItems(19)
    MatCreditos(i, 15) = lstCreditos.ListItems(i).SubItems(15)
    
    'MatCreditos(i, 16) = LstCreditos.ListItems(i).SubItems(26) 'Add Segmentacion Riesgo por cliente JOEP20201113
    MatCreditos(i, 16) = lblSegRiesgo.Caption   'Add Segmentacion Riesgo por cliente JOEP20201113
'FIN ORCR-20140913************
Next i

ReDim MatAhorros(lstAhorros.ListItems.count, 12)
For i = 1 To lstAhorros.ListItems.count
    MatAhorros(i, 1) = lstAhorros.ListItems(i).Text
    MatAhorros(i, 2) = lstAhorros.ListItems(i).SubItems(1)
    MatAhorros(i, 3) = lstAhorros.ListItems(i).SubItems(2)
    MatAhorros(i, 4) = lstAhorros.ListItems(i).SubItems(3)
    MatAhorros(i, 5) = lstAhorros.ListItems(i).SubItems(4)
    MatAhorros(i, 6) = lstAhorros.ListItems(i).SubItems(6)
    MatAhorros(i, 7) = lstAhorros.ListItems(i).SubItems(7)
    MatAhorros(i, 8) = lstAhorros.ListItems(i).SubItems(8)
    MatAhorros(i, 9) = lstAhorros.ListItems(i).SubItems(9)
    MatAhorros(i, 10) = lstAhorros.ListItems(i).SubItems(10)
    MatAhorros(i, 11) = lstAhorros.ListItems(i).SubItems(11)
    MatAhorros(i, 12) = lstAhorros.ListItems(i).SubItems(12)
Next i

ReDim MatPrendario(lstPrendario.ListItems.count, 12)
For i = 1 To lstPrendario.ListItems.count
    MatPrendario(i, 1) = lstPrendario.ListItems(i).Text
    MatPrendario(i, 2) = lstPrendario.ListItems(i).SubItems(1)
    MatPrendario(i, 3) = lstPrendario.ListItems(i).SubItems(2)
    MatPrendario(i, 4) = lstPrendario.ListItems(i).SubItems(3)
    MatPrendario(i, 5) = lstPrendario.ListItems(i).SubItems(4)
    MatPrendario(i, 6) = lstPrendario.ListItems(i).SubItems(5)
    MatPrendario(i, 7) = lstPrendario.ListItems(i).SubItems(6)
    MatPrendario(i, 8) = lstPrendario.ListItems(i).SubItems(7)
    MatPrendario(i, 9) = lstPrendario.ListItems(i).SubItems(8)
    MatPrendario(i, 10) = lstPrendario.ListItems(i).SubItems(10)
    MatPrendario(i, 11) = lstPrendario.ListItems(i).SubItems(11)
    MatPrendario(i, 12) = lstPrendario.ListItems(i).SubItems(12)
Next i

'INICIO ORCR-20140913*********
'ReDim MatJudicial(lstJudicial.ListItems.Count, 10)
ReDim MatJudicial(lstJudicial.ListItems.count, 11)
For i = 1 To lstJudicial.ListItems.count
    MatJudicial(i, 1) = lstJudicial.ListItems(i).Text
    MatJudicial(i, 2) = lstJudicial.ListItems(i).SubItems(1)
    MatJudicial(i, 3) = lstJudicial.ListItems(i).SubItems(4)
    MatJudicial(i, 4) = lstJudicial.ListItems(i).SubItems(3)
    MatJudicial(i, 5) = lstJudicial.ListItems(i).SubItems(2)
    MatJudicial(i, 6) = lstJudicial.ListItems(i).SubItems(5)
    MatJudicial(i, 7) = lstJudicial.ListItems(i).SubItems(6)
    MatJudicial(i, 8) = lstJudicial.ListItems(i).SubItems(7)
    MatJudicial(i, 9) = lstJudicial.ListItems(i).SubItems(8)
    MatJudicial(i, 10) = lstJudicial.ListItems(i).SubItems(9)
    '--------------------------------
    MatJudicial(i, 11) = lstJudicial.ListItems(i).SubItems(10)
'FIN ORCR-20140913************
Next i

ReDim MatCartaFianza(LstCartaFianza.ListItems.count, 9)
For i = 1 To LstCartaFianza.ListItems.count
    MatCartaFianza(i, 1) = LstCartaFianza.ListItems(i).Text
    MatCartaFianza(i, 2) = LstCartaFianza.ListItems(i).SubItems(1)
    MatCartaFianza(i, 3) = LstCartaFianza.ListItems(i).SubItems(3)
    MatCartaFianza(i, 4) = LstCartaFianza.ListItems(i).SubItems(2)
    MatCartaFianza(i, 5) = LstCartaFianza.ListItems(i).SubItems(4)
    MatCartaFianza(i, 6) = LstCartaFianza.ListItems(i).SubItems(5)
    MatCartaFianza(i, 7) = LstCartaFianza.ListItems(i).SubItems(6)
    MatCartaFianza(i, 8) = LstCartaFianza.ListItems(i).SubItems(7)
    MatCartaFianza(i, 9) = LstCartaFianza.ListItems(i).SubItems(8)
Next i

'ARCV 25-10-2006
ReDim MatSistFinanc(lstFinanciero.ListItems.count, 5)

For i = 0 To lstFinanciero.ListItems.count - 1
    For J = 0 To 4
        If J = 0 Then
            MatSistFinanc(i, J) = lstFinanciero.ListItems(i + 1).Text
        Else
            MatSistFinanc(i, J) = lstFinanciero.ListItems(i + 1).ListSubItems(J).Text
        End If
    Next J
Next i

'MatSistFinanc(0, 0) = lstFinanciero.ListItems(1).Text
'MatSistFinanc(0, 1) = lstFinanciero.ListItems(1).ListSubItems(1).Text
'MatSistFinanc(0, 2) = ""
'MatSistFinanc(0, 3) = ""
'MatSistFinanc(0, 4) = ""
'
'For i = 1 To 8
'    MatSistFinanc(i, 0) = lstFinanciero.ListItems(i + 1).Text
'Next i
'
'For i = 1 To 2
'    For j = 2 To 5
'        MatSistFinanc(i, j - 1) = lstFinanciero.ListItems(i + 1).ListSubItems(j - 1).Text
'    Next j
'Next i
'
'MatSistFinanc(3, 1) = lstFinanciero.ListItems(4).ListSubItems(1).Text
'MatSistFinanc(3, 2) = ""
'MatSistFinanc(3, 3) = ""
'MatSistFinanc(3, 4) = ""
'MatSistFinanc(3, 5) = ""
'MatSistFinanc(4, 1) = ""
'MatSistFinanc(4, 2) = ""
'MatSistFinanc(4, 3) = ""
'MatSistFinanc(4, 4) = ""
'MatSistFinanc(4, 5) = ""
'
'For i = 5 To 6
'    For j = 2 To 4
'        MatSistFinanc(i, j - 1) = lstFinanciero.ListItems(i + 1).ListSubItems(j - 1).Text
'    Next j
'Next i
'
'MatSistFinanc(7, 1) = ""
'MatSistFinanc(7, 2) = ""
'MatSistFinanc(7, 3) = ""
'MatSistFinanc(7, 4) = ""
'
'For i = 1 To 3
'    MatSistFinanc(8, i) = lstFinanciero.ListItems(9).ListSubItems(i).Text
'    If lstFinanciero.ListItems.Count = 10 Then
'        MatSistFinanc(9, i) = lstFinanciero.ListItems(10).ListSubItems(i).Text
'    End If
'Next i
'
'For i = 5 To 8
'    MatSistFinanc(i, 4) = ""
'Next i
'
'If lstFinanciero.ListItems.Count = 10 Then
'    MatSistFinanc(9, 4) = ""
'End If

'------------------------
Set oCredDoc = New COMNCredito.NCOMCredDoc
If Me.lblClientePreferncial.Visible = True Then
bPreferencial = True
Else
bPreferencial = False
End If
'ALPA***20080819*********************************************************************************************************************************************************************************************
'Se agrego la varible sCalifiFinal
'FRHU RQ13790 20140106 - Se agrego el parametro bPreferencial
GeneraImpresion = oCredDoc.GeneraImpresionPosicionCliente(bPreferencial, gsNomCmac, gdFecSis, gsNomAge, LblPersCod.Caption, lblNomPers.Caption, lblDocNat.Caption, lblDocJur.Caption, lstCreditos.ListItems.count, _
                                            MatCreditos, lstAhorros.ListItems.count, MatAhorros, lblSolesAho.Caption, lblDolaresAho.Caption, lstPrendario.ListItems.count, MatPrendario, lblSolesPig.Caption, _
                                            lblDolaresPig.Caption, lstJudicial.ListItems.count, MatJudicial, LstCartaFianza.ListItems.count, MatCartaFianza, rtfComentario.Text, nSaldosAho, MatSistFinanc, gsCodUser)
'End ALPA***/************************************************************************************************************************************************************************************************
Set oCredDoc = Nothing
End Function

Private Sub BuscaJudicial(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
Dim L As ListItem

    lstJudicial.ListItems.Clear
    
    If Not (pRs.EOF And pRs.BOF) Then
        lstJudicial.Enabled = True
    End If
    
    Do While Not pRs.EOF
        'INICIO ORCR20140812***
        Set L = lstJudicial.ListItems.Add(, , pRs.Bookmark)
        
        L.SubItems(1) = Format(pRs!dIngRecup, "dd/mm/yyyy") 'fecha Vigencia
        L.SubItems(2) = pRs!cCtaCod  'Cta
        L.SubItems(3) = pRs!cAgeDescripcion 'Agencia
        
        L.SubItems(4) = pRs!cTipoCredDescrip 'Tipo de Credito
        
        L.SubItems(5) = pRs!cEstado 'Estado
        L.SubItems(6) = pRs!cParticip 'Participacion
        L.SubItems(7) = Format(pRs!nSaldo, "#0.00") ' Saldo Cap
        
        L.SubItems(8) = pRs!cReprogracion ' Reprogramado
        
        If IsNull(pRs!dFecCast) Then
            L.SubItems(9) = "" 'Fecha Venc
        Else
            L.SubItems(9) = Format(pRs!dFecCast, "dd/mm/yyyy") 'Fecha Venc
        End If
        
        If IsNull(pRs!cMovnro) Then
            L.SubItems(10) = ""
        Else
            L.SubItems(10) = Mid(pRs!cMovnro, 7, 2) & "/" & Mid(pRs!cMovnro, 5, 2) & "/" & Mid(pRs!cMovnro, 1, 4) 'Ultimo Movimiento
        End If
        
        L.SubItems(11) = pRs!cTipoProdDescrip 'Tipo de Credito
        
        'FIN ORCR20140812***
        pRs.MoveNext
    Loop
    
    Exit Sub


End Sub

Private Sub BuscaPignoraticio(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
'Dim oCreditos As COMDCredito.DCOMCreditos
'Dim R As ADODB.Recordset
Dim L As ListItem
Dim SumaSol As Double, SumaDol As Double

'JOEP20210511 Segmentacion Prendario Externo
Dim objSegPred As COMDCredito.DCOMCreditos
Dim rsSegPred As ADODB.Recordset
Set objSegPred = New COMDCredito.DCOMCreditos
lblSegPredExt.Caption = "--"
'JOEP20210511 Segmentacion Prendario Externo

    SumaSol = 0
    SumaDol = 0
    
    lstPrendario.ListItems.Clear
'    Set oCreditos = New COMDCredito.DCOMCreditos
    
'    If nLima = 1 Then           'Trujillo
'        Set R = oCreditos.DatosPosicionClientePigno(psPersCod)
'    Else 'Lima
'        Set R = oCreditos.DatosPosicionClientePignoLima(psPersCod)
'    End If
    
    
    
    If Not (pRs.EOF And pRs.BOF) Then
        lstPrendario.Enabled = True
    
    'JOEP20210511 Segmentacion Prendario Externo
        Set rsSegPred = objSegPred.DataSegPrendarioExterno(psPersCod)
        If Not (rsSegPred.EOF And rsSegPred.BOF) Then
            lblSegPredExt.Font = "6dp"
            lblSegPredExt.Caption = rsSegPred!cTitulo & Chr(13) & rsSegPred!cSegmento
        End If
        
        Set objSegPred = Nothing
        RSClose rsSegPred
    'JOEP20210511 Segmentacion Prendario Externo
        
        Do While Not pRs.EOF
            
            Set L = lstPrendario.ListItems.Add(, , pRs.Bookmark)
            
            L.SubItems(1) = Format(pRs!dVigencia, "dd/mm/yyyy") 'fecha Vigencia
            L.SubItems(2) = pRs!cAgeDescripcion 'Agencia
            L.SubItems(3) = pRs!cCtaCod 'Cuenta
            L.SubItems(4) = pRs!cEstado 'estado
            L.SubItems(5) = pRs!cParticip 'Participacion
            L.SubItems(6) = Trim(str(pRs!nNroRenov)) 'Nro Renovacion
            L.SubItems(7) = Format(pRs!nMontoCol, "#0.00") ' Prestamo
            L.SubItems(8) = Format(pRs!dvenc, "dd/mm/yyyy") 'Fecha Venc
            L.SubItems(9) = pRs!cAgeCod 'Cod Agencia
            If Not (pRs!nPrdEstado = gColPEstRemat Or pRs!nPrdEstado = gColPEstSubas Or pRs!nPrdEstado = gColPEstAdjud Or pRs!nPrdEstado = gColPEstChafa Or pRs!nPrdEstado = "2113") Then
                L.SubItems(10) = Format(IIf(IsNull(pRs!nSaldo), 0, pRs!nSaldo), "#0.00") 'Saldo Capital
            Else
                L.SubItems(10) = Format("0.00", "#0.00")    'Saldo Capital
            End If
            
            L.SubItems(11) = Format(IIf(IsNull(pRs!nTasacion), 0, pRs!nTasacion), "#0.00")
            L.SubItems(12) = pRs!sMoneda
            
            
            SumaSol = SumaSol + pRs!nSaldo
            
            
            If pRs!nPrdEstado = gColPEstDesem Or pRs!nPrdEstado = gColPEstRenov Then
                L.Bold = True
                L.ForeColor = vbBlue
                L.ListSubItems(1).Bold = True
                L.ListSubItems(2).Bold = True
                L.ListSubItems(3).Bold = True
                L.ListSubItems(4).Bold = True
                L.ListSubItems(5).Bold = True
                L.ListSubItems(6).Bold = True
                L.ListSubItems(7).Bold = True
                L.ListSubItems(8).Bold = True
                L.ListSubItems(9).Bold = True
                L.ListSubItems(10).Bold = True
                L.ListSubItems(1).ForeColor = vbBlue
                L.ListSubItems(2).ForeColor = vbBlue
                L.ListSubItems(3).ForeColor = vbBlue
                L.ListSubItems(4).ForeColor = vbBlue
                L.ListSubItems(5).ForeColor = vbBlue
                L.ListSubItems(6).ForeColor = vbBlue
                L.ListSubItems(7).ForeColor = vbBlue
                L.ListSubItems(8).ForeColor = vbBlue
                L.ListSubItems(9).ForeColor = vbBlue
                L.ListSubItems(10).ForeColor = vbBlue
            Else
                L.Bold = False
                L.ForeColor = vbBlack
                L.ListSubItems(1).Bold = False
                L.ListSubItems(2).Bold = False
                L.ListSubItems(3).Bold = False
                L.ListSubItems(4).Bold = False
                L.ListSubItems(5).Bold = False
                L.ListSubItems(6).Bold = False
                L.ListSubItems(7).Bold = False
                L.ListSubItems(8).Bold = False
                L.ListSubItems(9).Bold = False
                L.ListSubItems(10).Bold = False
                L.ListSubItems(1).ForeColor = vbBlack
                L.ListSubItems(2).ForeColor = vbBlack
                L.ListSubItems(3).ForeColor = vbBlack
                L.ListSubItems(4).ForeColor = vbBlack
                L.ListSubItems(5).ForeColor = vbBlack
                L.ListSubItems(6).ForeColor = vbBlack
                L.ListSubItems(7).ForeColor = vbBlack
                L.ListSubItems(8).ForeColor = vbBlack
                L.ListSubItems(9).ForeColor = vbBlack
                L.ListSubItems(10).ForeColor = vbBlack
            End If
            
            pRs.MoveNext
        Loop
    End If
    'R.Close
    'Set R = Nothing
    'Set oCreditos = Nothing
    lblSolesPig.Caption = Format(SumaSol, "#0.00")
    
    Exit Sub

End Sub

Private Sub BuscaAhorros(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
    'Dim oCreditos As COMDCredito.DCOMCreditos
    'Dim R As ADODB.Recordset
    Dim L As ListItem
    Dim SumaSol As Double, SumaDol As Double

    SumaSol = 0
    SumaDol = 0
    On Error GoTo ERRORBuscaCreditos
    lstAhorros.ListItems.Clear
    'Set oCreditos = New COMDCredito.DCOMCreditos
    'Set R = oCreditos.DatosPosicionClienteAhorro(psPersCod)
    
    If Not (pRs.EOF And pRs.BOF) Then
        lstAhorros.Enabled = True
    End If
    
    Do While Not pRs.EOF
        
        Set L = lstAhorros.ListItems.Add(, , pRs.Bookmark)
        
        L.SubItems(1) = Format(pRs!dApertura, "dd/mm/yyyy")
        L.SubItems(2) = pRs!cTipoAho
        L.SubItems(3) = pRs!cAgeDescripcion
        L.SubItems(4) = pRs!cCtaCod
        L.SubItems(5) = IIf(IsNull(pRs!CCTACODANT), "", pRs!CCTACODANT)
        L.SubItems(6) = pRs!cEstado
        L.SubItems(7) = pRs!cParticip
        If nSaldosAho = 1 Then       'Maynas GRVA
            L.SubItems(8) = pRs!nSaldoCont
            L.SubItems(9) = pRs!nSaldoDisp - pRs!nBloqueoParcial
            L.SubItems(10) = pRs!nBloqueoParcial
        End If
        
        If pRs!sMoneda = "SOLES" Then
            SumaSol = SumaSol + pRs!nSaldoDisp
        Else
            SumaDol = SumaDol + pRs!nSaldoDisp
        End If
        
        
        L.SubItems(11) = IIf(IsNull(pRs!cMotivo), "", pRs!cMotivo)
        L.SubItems(12) = pRs!sMoneda 'Moneda
        
        If UCase(pRs!cEstado) = "ACTIVA" Then
            L.Bold = True
            L.ForeColor = vbBlue
            L.ListSubItems(1).Bold = True
            L.ListSubItems(2).Bold = True
            L.ListSubItems(3).Bold = True
            L.ListSubItems(4).Bold = True
            L.ListSubItems(5).Bold = True
            L.ListSubItems(6).Bold = True
            L.ListSubItems(7).Bold = True
            L.ListSubItems(8).Bold = True
            L.ListSubItems(9).Bold = True
            L.ListSubItems(10).Bold = True
            L.ListSubItems(11).Bold = True
            L.ListSubItems(12).Bold = True
            L.ListSubItems(1).ForeColor = vbBlue
            L.ListSubItems(2).ForeColor = vbBlue
            L.ListSubItems(3).ForeColor = vbBlue
            L.ListSubItems(4).ForeColor = vbBlue
            L.ListSubItems(5).ForeColor = vbBlue
            L.ListSubItems(6).ForeColor = vbBlue
            L.ListSubItems(7).ForeColor = vbBlue
            L.ListSubItems(8).ForeColor = vbBlue
            L.ListSubItems(9).ForeColor = vbBlue
            L.ListSubItems(10).ForeColor = vbBlue
            L.ListSubItems(11).ForeColor = vbBlue
            L.ListSubItems(12).ForeColor = vbBlue
        Else
            L.Bold = False
            L.ForeColor = vbBlack
            L.Bold = True
            L.ForeColor = vbBlue
            L.ListSubItems(1).Bold = False
            L.ListSubItems(2).Bold = False
            L.ListSubItems(3).Bold = False
            L.ListSubItems(4).Bold = False
            L.ListSubItems(5).Bold = False
            L.ListSubItems(6).Bold = False
            L.ListSubItems(7).Bold = False
            L.ListSubItems(8).Bold = False
            L.ListSubItems(9).Bold = False
            L.ListSubItems(10).Bold = False
            L.ListSubItems(11).Bold = False
            L.ListSubItems(12).Bold = False
            L.ListSubItems(1).ForeColor = vbBlack
            L.ListSubItems(2).ForeColor = vbBlack
            L.ListSubItems(3).ForeColor = vbBlack
            L.ListSubItems(4).ForeColor = vbBlack
            L.ListSubItems(5).ForeColor = vbBlack
            L.ListSubItems(6).ForeColor = vbBlack
            L.ListSubItems(7).ForeColor = vbBlack
            L.ListSubItems(8).ForeColor = vbBlack
            L.ListSubItems(9).ForeColor = vbBlack
            L.ListSubItems(10).ForeColor = vbBlack
            L.ListSubItems(11).ForeColor = vbBlack
            L.ListSubItems(12).ForeColor = vbBlack
        End If
        
        pRs.MoveNext
    Loop
    
    lblSolesAho.Caption = Format(SumaSol, "#0.00")
    
    lblDolaresAho.Caption = Format(SumaDol, "#0.00")
    
    'Maynas GRVA
    If nSaldosAho = 1 Then
        lstAhorros.ColumnHeaders(9).Width = 1300
        lstAhorros.ColumnHeaders(10).Width = 1300
        lstAhorros.ColumnHeaders(9).Text = "Saldo contable"
        lstAhorros.ColumnHeaders(10).Text = "Saldo disponible"
        
        'Add By GITU 2012-09-05
        lstAhorros.ColumnHeaders(11).Width = 1300
        lstAhorros.ColumnHeaders(11).Text = "Bloqueo Parcial"
        'End GITU
        
        Label4.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        lblSolesAho.Visible = True
        lblDolaresAho.Visible = True
    Else
        lstAhorros.ColumnHeaders(9).Width = 0
        lstAhorros.ColumnHeaders(10).Width = 0
        lstAhorros.ColumnHeaders(11).Width = 0
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        lblSolesAho.Visible = False
        lblDolaresAho.Visible = False
    End If
    
    'R.Close
    'Set R = Nothing
    'Set oCreditos = Nothing
    Exit Sub
    
ERRORBuscaCreditos:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub BuscaCreditos(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
Dim L As ListItem

Dim nEstado As Integer
sCalifiFinal = ""
Dim ldFecCalifCredCancelados As Date 'Fecha para comparar calificacion de creditos cancelados
ldFecCalifCredCancelados = CDate("1900/01/01") 'SEGÚN aCTA 77-2008/TI-D

'JOEP20210511 Segmentacion Prendario Externo
Dim objSegExperian As COMDCredito.DCOMCreditos
Dim rsSegExperian As ADODB.Recordset
Set objSegExperian = New COMDCredito.DCOMCreditos
lblSegExperianExt.Caption = "--"
'JOEP20210511 Segmentacion Prendario Externo

On Error GoTo ERRORBuscaCreditos
    
    lstCreditos.ListItems.Clear
    
    If Not (pRs.EOF And pRs.BOF) Then
        lstCreditos.Enabled = True
    End If
    
    'JOEP20210511 Segmentacion Prendario Externo
    Set rsSegExperian = objSegExperian.DataSegExperianExterno(psPersCod)
    If Not (rsSegExperian.EOF And rsSegExperian.BOF) Then
        lblSegExperianExt.Font = "6dp"
        lblSegExperianExt.Caption = rsSegExperian!cTitulo & Chr(13) & rsSegExperian!cSegmento
    End If
    
    Set objSegExperian = Nothing
    RSClose rsSegExperian
    'JOEP20210511 Segmentacion Prendario Externo
    
    Do While Not pRs.EOF
        'INICIO ORCR-20140913*********
        Set L = lstCreditos.ListItems.Add(, , pRs.Bookmark)
        
        If IsNull(pRs!dDesembolso) Then
            L.SubItems(1) = ""
        Else
            L.SubItems(1) = Format(pRs!dDesembolso, "dd/mm/yyyy")
        End If
        
        L.SubItems(2) = pRs!cCtaCod 'Credito
        L.SubItems(3) = pRs!cAgeDescripcion 'Agencia
        L.SubItems(4) = pRs!cTipoCredDescrip 'Tipo de Credito
        
        L.SubItems(5) = IIf(IsNull(pRs!cEstadoDesc), "", pRs!cEstadoDesc) 'Estado
        L.SubItems(6) = pRs!cRelacionDesc 'Participacion
        L.SubItems(7) = IIf(IsNull(pRs!cPersAnalista), "", pRs!cPersAnalista) 'Analista
        L.SubItems(8) = IIf(IsNull(pRs!AnalistaInicial), "", pRs!AnalistaInicial)
        L.SubItems(9) = IIf(IsNull(pRs!nAnalistaNota), "", pRs!nAnalistaNota) 'Nota
        L.SubItems(10) = Format(pRs!nPrestamo, "0.00")
        L.SubItems(11) = Format(pRs!nSaldo, "#0.00") 'Saldo Capital
        L.SubItems(12) = "" 'Codigo Antiguo 1
        L.SubItems(13) = "" 'Codigo Antiguo 2
        L.SubItems(14) = pRs!sMoneda 'Moneda
        
        L.SubItems(15) = pRs!cReprogracion 'Reprogramado
        
        If IsNull(pRs!dCancelado) Then
            L.SubItems(16) = ""
        Else
            L.SubItems(16) = Format(pRs!dCancelado, "dd/mm/yyyy") 'Fec Cancelacion
        End If
        
        L.SubItems(17) = Format(pRs!dSolicitado, "dd/mm/yyyy") 'fecha de Solicitud
        L.SubItems(18) = pRs!nDiasAtrasoAcum 'prs!cRFA  'CUSCO
        
        If Not IsNull(pRs!dCancelado) And Format(pRs!dCancelado, "YYYY/MM/DD") >= ldFecCalifCredCancelados Or Format(pRs!dSolicitado, "YYYY/MM/DD") >= ldFecCalifCredCancelados Then
            If pRs!cRelacionDesc = "TITULAR" And pRs!nNotVigenteCa = 0 Then
                L.SubItems(19) = pRs!cCalif
            End If
        Else
            L.SubItems(19) = ""
        End If
        
        L.SubItems(20) = pRs!cCondicion
        L.SubItems(21) = pRs!cTipoProdDescrip 'Tipo de Producto
        L.SubItems(22) = pRs!cRefVig 'Refinanciado Vigente
        L.SubItems(23) = pRs!dRefVig 'Fecha Refinanciado Vigente
        
    'JOEP20180904 CP
        L.SubItems(24) = pRs!Campana
        L.SubItems(25) = pRs!TipoCredito
    'JOEP20180904 CP
    'Segmentacion Riesgo por cliente JOEP20201113
        'L.SubItems(26) = pRs!cNivelRiesgoCredito
        lblSegRiesgo.Caption = ""
        lblSegRiesgo.Caption = pRs!cNivelRiesgoCliente
    'Segmentacion Riesgo por cliente JOEP20201113
    
        If pRs!nPrdEstado = gColocEstRefMor Or pRs!nPrdEstado = gColocEstRefVenc _
            Or pRs!nPrdEstado = gColocEstRefNorm Or pRs!nPrdEstado = gColocEstVigVenc _
            Or pRs!nPrdEstado = gColocEstVigMor Or pRs!nPrdEstado = gColocEstVigNorm Then
            L.Bold = True
            L.ForeColor = vbBlue
            L.ListSubItems(1).Bold = True
            L.ListSubItems(2).Bold = True
            L.ListSubItems(3).Bold = True
            L.ListSubItems(4).Bold = True
            L.ListSubItems(5).Bold = True
            L.ListSubItems(6).Bold = True
            L.ListSubItems(7).Bold = True
            L.ListSubItems(8).Bold = True
            L.ListSubItems(9).Bold = True
            L.ListSubItems(10).Bold = True
            L.ListSubItems(11).Bold = True
            L.ListSubItems(12).Bold = True
            L.ListSubItems(13).Bold = True
            L.ListSubItems(14).Bold = True
            L.ListSubItems(15).Bold = True
            L.ListSubItems(16).Bold = True
            L.ListSubItems(17).Bold = True
            L.ListSubItems(18).Bold = True
            L.ListSubItems(19).Bold = True
            L.ListSubItems(20).Bold = True
            'ALPA 20150730
            L.ListSubItems(21).Bold = True
            L.ListSubItems(22).Bold = True
            L.ListSubItems(23).Bold = True
            '*************
        'JOEP20180904 CP
            L.ListSubItems(24).Bold = True
            L.ListSubItems(25).Bold = True
        'JOEP20180904 CP
        
        'Segmentacion Riesgo por cliente JOEP20201113
            'L.ListSubItems(26).Bold = True
        'Segmentacion Riesgo por cliente JOEP20201113

            L.ListSubItems(1).ForeColor = vbBlue
            L.ListSubItems(2).ForeColor = vbBlue
            L.ListSubItems(3).ForeColor = vbBlue
            L.ListSubItems(4).ForeColor = vbBlue
            L.ListSubItems(5).ForeColor = vbBlue
            L.ListSubItems(6).ForeColor = vbBlue
            L.ListSubItems(7).ForeColor = vbBlue
            L.ListSubItems(8).ForeColor = vbBlue
            L.ListSubItems(9).ForeColor = vbBlue
            L.ListSubItems(10).ForeColor = vbBlue
            L.ListSubItems(11).ForeColor = vbBlue
            L.ListSubItems(12).ForeColor = vbBlue
            L.ListSubItems(13).ForeColor = vbBlue
            L.ListSubItems(14).ForeColor = vbBlue
            L.ListSubItems(15).ForeColor = vbBlue
            L.ListSubItems(16).ForeColor = vbBlue
            L.ListSubItems(17).ForeColor = vbBlue
            L.ListSubItems(18).ForeColor = vbBlue
            L.ListSubItems(19).ForeColor = vbBlue
            L.ListSubItems(20).ForeColor = vbBlue
            'ALPA 20150730
            L.ListSubItems(21).ForeColor = vbBlue
            L.ListSubItems(22).ForeColor = vbBlue
            L.ListSubItems(23).ForeColor = vbBlue
            '*************
           
         'JOEP20180904 CP
            L.ListSubItems(24).ForeColor = vbBlue
            L.ListSubItems(25).ForeColor = vbBlue
         'JOEP20180904 CP
            
        'Segmentacion Riesgo por cliente JOEP20201113
            'L.ListSubItems(26).ForeColor = vbBlue
        'Segmentacion Riesgo por cliente JOEP20201113
            
            nEstado = 1
        Else
            L.ForeColor = vbBlack
            L.Bold = False
            L.ListSubItems(1).Bold = False
            L.ListSubItems(2).Bold = False
            L.ListSubItems(3).Bold = False
            L.ListSubItems(4).Bold = False
            L.ListSubItems(5).Bold = False
            L.ListSubItems(6).Bold = False
            L.ListSubItems(7).Bold = False
            L.ListSubItems(8).Bold = False
            L.ListSubItems(9).Bold = False
            L.ListSubItems(10).Bold = False
            L.ListSubItems(11).Bold = False
            L.ListSubItems(12).Bold = False
            L.ListSubItems(13).Bold = False
            L.ListSubItems(14).Bold = False
            L.ListSubItems(15).Bold = False
            L.ListSubItems(16).Bold = False
            L.ListSubItems(17).Bold = False
            L.ListSubItems(18).Bold = False
            L.ListSubItems(19).Bold = False
            L.ListSubItems(20).Bold = False
            
            L.ListSubItems(1).ForeColor = vbBlack
            L.ListSubItems(2).ForeColor = vbBlack
            L.ListSubItems(3).ForeColor = vbBlack
            L.ListSubItems(4).ForeColor = vbBlack
            L.ListSubItems(5).ForeColor = vbBlack
            L.ListSubItems(6).ForeColor = vbBlack
            L.ListSubItems(7).ForeColor = vbBlack
            L.ListSubItems(8).ForeColor = vbBlack
            L.ListSubItems(9).ForeColor = vbBlack
            L.ListSubItems(10).ForeColor = vbBlack
            L.ListSubItems(11).ForeColor = vbBlack
            L.ListSubItems(12).ForeColor = vbBlack
            L.ListSubItems(13).ForeColor = vbBlack
            L.ListSubItems(14).ForeColor = vbBlack
            L.ListSubItems(15).ForeColor = vbBlack
            L.ListSubItems(16).ForeColor = vbBlack
            L.ListSubItems(17).ForeColor = vbBlack
            L.ListSubItems(18).ForeColor = vbBlack
            L.ListSubItems(19).ForeColor = vbBlack
            L.ListSubItems(20).ForeColor = vbBlack
            
            nEstado = 0
        End If
        'FIN ORCR-20140913************
        pRs.MoveNext
    Loop
    
    Exit Sub
    
ERRORBuscaCreditos:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub BuscaCF(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
'Dim oCreditos As COMDCredito.DCOMCreditos
'Dim R As ADODB.Recordset
Dim L As ListItem
On Error GoTo ERRORBuscaCF
    LstCartaFianza.ListItems.Clear
    'Set oCreditos = New COMDCredito.DCOMCreditos
    'Set R = oCreditos.DatosPosicionClienteCF(psPersCod)
    
    If Not (pRs.EOF And pRs.BOF) Then
        LstCartaFianza.Enabled = True
    End If
    
    Do While Not pRs.EOF
        Set L = LstCartaFianza.ListItems.Add(, , pRs.Bookmark)
        
        L.SubItems(1) = Format(pRs!dEmitido, "dd/mm/yyyy") 'fecha de Solicitud
        L.SubItems(2) = pRs!cCtaCod 'Credito
        L.SubItems(3) = pRs!cAgeDescripcion 'Agencia
        L.SubItems(4) = IIf(IsNull(pRs!cEstadoDesc), "", pRs!cEstadoDesc) 'Estado
        L.SubItems(5) = pRs!cRelacionDesc 'Participacion
        L.SubItems(6) = IIf(IsNull(pRs!cPersAnalista), "", pRs!cPersAnalista) 'Analista
        L.SubItems(7) = Format(pRs!nPrestamo, "0.00")
        L.SubItems(8) = pRs!sMoneda 'Moneda
        L.SubItems(9) = Format(pRs!dCancelado, "dd/mm/yyyy") 'fecha de Honrado
        
        If pRs!nPrdEstado = gColocEstVigNorm Then
            L.Bold = True
            L.ForeColor = vbBlue
            L.ListSubItems(1).Bold = True
            L.ListSubItems(2).Bold = True
            L.ListSubItems(3).Bold = True
            L.ListSubItems(4).Bold = True
            L.ListSubItems(5).Bold = True
            L.ListSubItems(6).Bold = True
            L.ListSubItems(7).Bold = True
            L.ListSubItems(8).Bold = True
            L.ListSubItems(9).Bold = True
            L.ListSubItems(1).ForeColor = vbBlue
            L.ListSubItems(2).ForeColor = vbBlue
            L.ListSubItems(3).ForeColor = vbBlue
            L.ListSubItems(4).ForeColor = vbBlue
            L.ListSubItems(5).ForeColor = vbBlue
            L.ListSubItems(6).ForeColor = vbBlue
            L.ListSubItems(7).ForeColor = vbBlue
            L.ListSubItems(8).ForeColor = vbBlue
            L.ListSubItems(9).ForeColor = vbBlue
        Else
            L.ForeColor = vbBlack
            L.Bold = False
            L.ListSubItems(1).Bold = False
            L.ListSubItems(2).Bold = False
            L.ListSubItems(3).Bold = False
            L.ListSubItems(4).Bold = False
            L.ListSubItems(5).Bold = False
            L.ListSubItems(6).Bold = False
            L.ListSubItems(7).Bold = False
            L.ListSubItems(8).Bold = False
            L.ListSubItems(9).Bold = False
            L.ListSubItems(1).ForeColor = vbBlack
            L.ListSubItems(2).ForeColor = vbBlack
            L.ListSubItems(3).ForeColor = vbBlack
            L.ListSubItems(4).ForeColor = vbBlack
            L.ListSubItems(5).ForeColor = vbBlack
            L.ListSubItems(6).ForeColor = vbBlack
            L.ListSubItems(7).ForeColor = vbBlack
            L.ListSubItems(8).ForeColor = vbBlack
            L.ListSubItems(9).ForeColor = vbBlack
        End If
        
        pRs.MoveNext
    Loop
    'R.Close
    'Set R = Nothing
    'Set oCreditos = Nothing
    Exit Sub
    
ERRORBuscaCF:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub BuscaComentarios(ByVal sPersona As String, ByVal pRs As ADODB.Recordset)
Dim oCreditos As COMDCredito.DCOMCreditos 'ANPS
Dim R As ADODB.Recordset 'ANPS
Dim sComentario As String, sFecha As String, sUsuario As String
Dim sCadena As String
Dim i As Integer

    

    rtfComentario.Text = ""
    Set oCreditos = New COMDCredito.DCOMCreditos
    Set R = oCreditos.DatosPosicionClienteComentarios(sPersona) 'ANPS
    If Not (pRs.EOF And pRs.BOF) Then
        rtfComentario.Enabled = True
    Else
        rtfComentario.Enabled = False
    End If
    sCadena = ""
    i = 0
    Do While Not pRs.EOF
    If pRs("cMovNro") = "0" And pRs("cPersCod") = "0" Then 'ANPS
        sComentario = Trim(pRs("cComentario")) 'ANPS
        sCadena = sComentario 'ANPS
      '  rtfComentario.BackColor = vbRed   'ANPS
        rtfComentario.Font.Bold = True 'ANPS
       rtfComentario.Font.Size = 9        '12.5 'ANPS
    Else
        sComentario = Trim(pRs("cComentario"))
        sFecha = Mid(pRs("cMovNro"), 7, 2) & "/" & Mid(pRs("cMovNro"), 5, 2) & "/" & Mid(pRs("cMovNro"), 1, 4)
        sFecha = sFecha & " " & Mid(pRs("cMovNro"), 9, 2) & ":" & Mid(pRs("cMovNro"), 11, 2) & ":" & Mid(pRs("cMovNro"), 13, 2)
        sUsuario = Right(pRs("cMovNro"), 4)
        If sCadena <> "" Then sCadena = sCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        i = i + 1
        sCadena = sCadena & "COMENTARIO " & Format$(i, "00") & " : " & oImpresora.gPrnSaltoLinea
        sCadena = sCadena & sComentario & oImpresora.gPrnSaltoLinea
        sCadena = sCadena & "USUARIO : " & sUsuario & "  FECHA : " & sFecha
    End If
        pRs.MoveNext
    
    Loop
    If sCadena <> "" Then
        rtfComentario.Text = sCadena
    End If
End Sub

Private Sub cmdBuscar_Click()

Dim oPersona As COMDPersona.UCOMPersona
Dim sPersCod As String
'FRHU 20121202
'Dim oCliPre As COMNCredito.NCOMCredito     'COMENTADO POR ARLO 20170722
'Set oCliPre = New COMNCredito.NCOMCredito  'COMENTADO POR ARLO 20170722
Dim bValidar As Boolean
'FIN 20121202
    lstCreditos.Enabled = False
    lstAhorros.Enabled = False
    LstCartaFianza.Enabled = False
    lstJudicial.Enabled = False
    lstPrendario.Enabled = False
    
    Set oPersona = frmBuscaPersona.Inicio
    Screen.MousePointer = 11 'CTI6-20210503-ERS032-2019. Agregó
    If Not oPersona Is Nothing Then
        LblPersCod.Caption = oPersona.sPersCod
        lblNomPers.Caption = oPersona.sPersNombre
        lblDocNat.Caption = Trim(oPersona.sPersIdnroDNI)
        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
        lsPersTDoc = "1"
        '**DAOR 20080410 *****************************************
        If oPersona.sPersPersoneria = "1" Then
            fbPersNatural = True
            'madm 20100707
            If Trim(oPersona.sPersIdnroDNI) = "" Then
                If Not Trim(oPersona.sPersIdnroOtro) = "" Then
                    lblDocNat.Caption = Trim(oPersona.sPersIdnroOtro)
                    lsPersTDoc = Trim(oPersona.sPersTipoDoc) 'aqUI
                End If
            End If
            'end madm
        Else
            fbPersNatural = False
            lsPersTDoc = "3"
        End If
        '*********************************************************
    Else
        Exit Sub
    End If
    sPersCod = oPersona.sPersCod
    'FRHU 20121202
    'bValidar = oCliPre.ValidarClientePreferencial(sPersCod) 'COMETADO POR ARLO 20170722
    bValidar = False 'AGREGADO POR ARLO 20170722
    If bValidar Then
        Me.lblClientePreferncial.Visible = True
    Else
        Me.lblClientePreferncial.Visible = False
    End If
    'FIN FRHU
    Set oPersona = Nothing
        
    If sPersCod <> "" Then 'MAVM 20120605 Reportad Por EIRE
        Call BuscarPosicionCliente(sPersCod, lsPersTDoc)
    End If
    
    If sPersCod <> "" Then
        cmdSaldosConsol.Enabled = True
        cmdImprimir.Enabled = True
    End If
    Screen.MousePointer = 0 'CTI6-20210503-ERS032-2019. Agregó
End Sub

Private Sub BuscarPosicionCliente(ByVal psPersCod As String, Optional psPersTDoc As String = "1")
    Dim oCreds As COMDCredito.DCOMCreditos
    Dim rsCred As ADODB.Recordset
    Dim rsAho As ADODB.Recordset
    Dim rsPig As ADODB.Recordset
    Dim rsJud As ADODB.Recordset
    Dim rsCF As ADODB.Recordset
    Dim rsCom As ADODB.Recordset

    'Se Agrego para la Calificacion RCC
    Dim rsCalSBS As ADODB.Recordset
    Dim rsEndSBS As ADODB.Recordset
    Dim rsCalCMAC As ADODB.Recordset
    Dim rsDeuEnt As ADODB.Recordset 'FRHU20140221 RQ14016
    Dim bExitoBusqueda As Boolean
    Dim dFechaRep As Date
    Dim lsPersDoc As String '**DAOR 20080410
    Dim lsPersTDoc As String 'ALPA 20100922
    Set oCreds = New COMDCredito.DCOMCreditos
    
    'CTI6-20210503-ERS032-2019.
    Dim nCantTab As Integer
    nCantTab = TabPosicion.Tabs
    PbCargaPosicion.Min = 0
    PbCargaPosicion.Max = nCantTab
    PbCargaPosicion.value = 0
    PbCargaPosicion.Visible = True
    'Fin CTI6-20210503
    
    '**Modificado por DAOR 20080410 ***************************************************
    'bExitoBusqueda = oCreds.BuscarPosicionCliente(psPersCod, IIf(Check1.value = 1, True, False), nLima, _
    '                                IIf(lblDocJur.Caption = "", True, False), _
    '                                IIf(lblDocJur.Caption <= "", Trim(lblDocNat.Caption), Trim(lblDocJur.Caption)), _
    '                                rsCred, rsAho, rsPig, rsJud, rsCF, rsCom, dFechaRep, rsCalSBS, rsEndSBS, rsCalCMAC)
    '
    If fbPersNatural Then
        lsPersDoc = IIf(lblDocNat.Caption = "", Trim(lblDocJur.Caption), Trim(lblDocNat.Caption))
        
    Else
        lsPersDoc = IIf(lblDocJur.Caption <= "", Trim(lblDocNat.Caption), Trim(lblDocJur.Caption))
    End If
    bExitoBusqueda = oCreds.BuscarPosicionCliente(psPersCod, IIf(Check1.value = 1, True, False), nLima, _
                                    fbPersNatural, lsPersDoc, rsCred, rsAho, rsPig, rsJud, rsCF, rsCom, dFechaRep, rsCalSBS, rsEndSBS, rsCalCMAC, gdFecSis, 1, psPersTDoc, _
                                    rsDeuEnt) 'FRHU 20140221 RQ14016 se agrego rsDeuEnt
    '************************************************************************************
    
    Set oCreds = Nothing
    
    PbCargaPosicion.value = PbCargaPosicion.value + 1 'CTI6-20210503-ERS032-2019
    Call BuscaCreditos(psPersCod, rsCred)
    
    PbCargaPosicion.value = PbCargaPosicion.value + 1 'CTI6-20210503-ERS032-2019
    Call BuscaAhorros(psPersCod, rsAho)
    
    PbCargaPosicion.value = PbCargaPosicion.value + 1 'CTI6-20210503-ERS032-2019
    Call BuscaPignoraticio(psPersCod, rsPig)
    
    PbCargaPosicion.value = PbCargaPosicion.value + 1 'CTI6-20210503-ERS032-2019
    Call BuscaJudicial(psPersCod, rsJud)
    
    PbCargaPosicion.value = PbCargaPosicion.value + 1 'CTI6-20210503-ERS032-2019
    Call BuscaCF(psPersCod, rsCF)
    Call BuscaComentarios(psPersCod, rsCom)
    If bExitoBusqueda Then
        PbCargaPosicion.value = PbCargaPosicion.value + 1 'CTI6-20210503-ERS032-2019
        Call BuscaCalificacionRCC(dFechaRep, rsCalSBS, rsEndSBS, rsCalCMAC, rsDeuEnt) 'FRHU20140221 se agrego rsDeuEnt
    End If
    PbCargaPosicion.Visible = False
End Sub

Private Sub BuscaCalificacionRCC(ByVal pdFechaRep As Date, _
                                ByVal prsCalSBS As ADODB.Recordset, _
                                ByVal prsEndSBS As ADODB.Recordset, _
                                ByVal prsCalCMAC As ADODB.Recordset, _
                                Optional ByVal prsDeuEnt As ADODB.Recordset) 'FRHU20140221 se agrego prsDeuEnt

Dim Item As ListItem
Dim fil As Integer
Dim lnCorreFinanciero As Long

    lnCorreFinanciero = 0
    lstFinanciero.ListItems.Clear

    lnCorreFinanciero = lnCorreFinanciero + 1
    Set Item = Me.lstFinanciero.ListItems.Add(, , "Calificacion SBS-RCC ")
    Item.SubItems(1) = Format(pdFechaRep, "dd/mm/yyyy")
    Item.SubItems(2) = ""
    Item.SubItems(3) = ""
    Item.SubItems(4) = ""
    
    lnCorreFinanciero = lnCorreFinanciero + 1
    Set Item = Me.lstFinanciero.ListItems.Add(, , "Normal")
    Item.SubItems(1) = "Potencial"
    Item.SubItems(2) = "Deficiente"
    Item.SubItems(3) = "Dudoso"
    Item.SubItems(4) = "Perdida"
     
    lstFinanciero.ListItems(1).ForeColor = vbRed
    lstFinanciero.ListItems(1).Bold = True
    lstFinanciero.ListItems(2).ListSubItems(2).ForeColor = vbRed
     
    lstFinanciero.ListItems(2).ForeColor = vbBlue
    lstFinanciero.ListItems(2).Bold = True
    For fil = 1 To 4
        lstFinanciero.ListItems(2).ListSubItems(fil).ForeColor = vbBlue
        lstFinanciero.ListItems(2).ListSubItems(fil).Bold = True
    Next
    'Calificacion SBS
    If Len(Trim(lblDocNat.Caption)) = 0 And Len(Trim(lblDocJur.Caption)) = 0 Then
        MsgBox "Cliente no registra documento, Favor Actualizar Datos ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If prsCalSBS.BOF And prsCalSBS.EOF Then
        lnCorreFinanciero = lnCorreFinanciero + 1
        Set Item = Me.lstFinanciero.ListItems.Add(, , "No Registrado")
        Item.SubItems(1) = "No Registrado"
        Item.SubItems(2) = "No Registrado"
        Item.SubItems(3) = "No Registrado"
        Item.SubItems(4) = "No Registrado"
    Else
        Do While Not prsCalSBS.EOF
            lnCorreFinanciero = lnCorreFinanciero + 1
            Set Item = Me.lstFinanciero.ListItems.Add(, , prsCalSBS!nNormal & "%")
            Item.SubItems(1) = prsCalSBS!nPotencial & "%"
            Item.SubItems(2) = prsCalSBS!nDeficiente & "%"
            Item.SubItems(3) = prsCalSBS!nDudoso & "%"
            Item.SubItems(4) = prsCalSBS!nPerdido & "%"
            
            lnCorreFinanciero = lnCorreFinanciero + 1
            Set Item = Me.lstFinanciero.ListItems.Add(, , "NRO ENTIDADES : ")
            Item.SubItems(1) = prsCalSBS!Can_Ents
            Item.SubItems(2) = ""
            Item.SubItems(3) = ""
            Item.SubItems(4) = ""
            prsCalSBS.MoveNext
        Loop
        'FRHU 20140221 RQ14016
        lnCorreFinanciero = lnCorreFinanciero + 1
        Set Item = Me.lstFinanciero.ListItems.Add(, , "Entidad")
        Item.SubItems(1) = "Moneda"
        Item.SubItems(2) = "Saldo MN"
        Item.SubItems(3) = "Calif. Entidad"
        Item.SubItems(4) = "%"
        
        lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbBlue
        lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
        For fil = 1 To 4
            lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).ForeColor = vbBlue
            lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).Bold = True
        Next
        Do While Not prsDeuEnt.EOF
            lnCorreFinanciero = lnCorreFinanciero + 1
            Set Item = Me.lstFinanciero.ListItems.Add(, , prsDeuEnt!Entidad)
            Item.SubItems(1) = prsDeuEnt!Moneda
            Item.SubItems(2) = prsDeuEnt!Saldo
            Item.SubItems(3) = prsDeuEnt!Clasificacion
            Item.SubItems(4) = prsDeuEnt!Porcentaje & "%"
            prsDeuEnt.MoveNext
        Loop
        prsDeuEnt.Close
        Set prsDeuEnt = Nothing
        'FIN FRHU 20140221 RQ14016
        'Endeudamiento SBS
        lnCorreFinanciero = lnCorreFinanciero + 1
        Set Item = Me.lstFinanciero.ListItems.Add(, , "Endeudamiento")
        lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbRed
        lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
        Item.SubItems(1) = ""
        Item.SubItems(2) = ""
        Item.SubItems(3) = ""
        Item.SubItems(4) = ""

        lnCorreFinanciero = lnCorreFinanciero + 1
        Set Item = Me.lstFinanciero.ListItems.Add(, , "Directa Soles")
        Item.SubItems(1) = "Directa Dolar"
        Item.SubItems(2) = "Indirecta Soles"
        Item.SubItems(3) = "Indirecta Dolar"
        Item.SubItems(4) = ""
        
        lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbBlue
        lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
        For fil = 1 To 3
            lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).ForeColor = vbBlue
            lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).Bold = True
        Next

        lnCorreFinanciero = lnCorreFinanciero + 1
        Set Item = Me.lstFinanciero.ListItems.Add(, , prsEndSBS!DDirSoles)
        Item.SubItems(1) = Format(prsEndSBS!DDirDolar, "#,#00.00")
        Item.SubItems(2) = Format(prsEndSBS!dIndSoles, "#,#00.00")
        Item.SubItems(3) = Format(prsEndSBS!dIndDolar, "#,#00.00")
        Item.SubItems(4) = ""
        prsEndSBS.Close
    End If
    prsCalSBS.Close
    Set prsCalSBS = Nothing
    
        
    'Calificacion CMAC
          
    lnCorreFinanciero = lnCorreFinanciero + 1
    Set Item = Me.lstFinanciero.ListItems.Add(, , "Calificacion CMAC - Riesgos")
    Item.SubItems(1) = ""
    Item.SubItems(2) = ""
    Item.SubItems(3) = ""
    Item.SubItems(4) = ""
    
    lnCorreFinanciero = lnCorreFinanciero + 1
    Set Item = Me.lstFinanciero.ListItems.Add(, , "Fecha")
    
    Item.SubItems(1) = "Calif. Final"
    Item.SubItems(2) = "Calif. Riesgos"
    Item.SubItems(3) = "Calif. S.Financ"
    Item.SubItems(4) = ""
         
    lstFinanciero.ListItems(lnCorreFinanciero - 1).ForeColor = vbRed
    lstFinanciero.ListItems(lnCorreFinanciero - 1).Bold = True
     
    lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbBlue
    lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
    For fil = 1 To 3
        lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).ForeColor = vbBlue
        lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).Bold = True
    Next
          
    If Not (prsCalCMAC Is Nothing) Then
        Do While Not prsCalCMAC.EOF
            Set Item = Me.lstFinanciero.ListItems.Add(, , Format(prsCalCMAC!dfecha, "dd/MM/YYYY"))
            Item.SubItems(1) = prsCalCMAC!nCalFinal
            Item.SubItems(2) = prsCalCMAC!nCalRiesgos
            Item.SubItems(3) = prsCalCMAC!nCalSistFinan
            Item.SubItems(4) = ""
            DoEvents
            prsCalCMAC.MoveNext
        Loop
        prsCalCMAC.Close
    Else
        lnCorreFinanciero = lnCorreFinanciero + 1
        Set Item = Me.lstFinanciero.ListItems.Add(, , "No Registrado")
        Item.SubItems(1) = "No Registrado"
        Item.SubItems(2) = "No Registrado"
        Item.SubItems(3) = "No Registrado"
        Item.SubItems(4) = ""
    End If
    Set prsCalCMAC = Nothing

End Sub
'**DAOR 20070512
Private Sub cmdEstadoCuenta_Click()
    If Me.lstCreditos.SelectedItem.SubItems(2) <> "" Then
        Call ImprimeEstadoCuentaCredito(Me.lstCreditos.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub cmdHistorial_Click()
    Dim oNCredDoc As COMNCredito.NCOMCredDoc
    Dim sCadImp As String
    '***Modificado por ELRO el 20111117, según Acta 316-2011/TI-D
    Dim oPrev As clsprevio
    'Dim oPrev As PrevioCredito.clsPrevioCredito
    '***Fin Modificado por ELRO**********************************
    
    Set oNCredDoc = New COMNCredito.NCOMCredDoc
        sCadImp = oNCredDoc.ImpreRepor_HistorialCliente(LblPersCod.Caption, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
    Set oNCredDoc = Nothing
    
    If Len(sCadImp) = 0 Then
        MsgBox "No se encontraron datos del reporte", vbInformation, "AVISO"
    Else
        '***Modificado por ELRO el 20111117, según Acta 316-2011/TI-D
        Set oPrev = New clsprevio
        'Set oPrev = New PrevioCredito.clsPrevioCredito
        '***Fin Modificado por ELRO**********************************
        oPrev.Show sCadImp, "LISTA DE CREDITOS DEL CLIENTE", True
        Set oPrev = Nothing
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim oPrev As previo.clsprevio

    If LblPersCod <> "" Then
        MsgBox "Este documento no representa un Historial de Pagos; es un documento interno", vbInformation, "Aviso" 'JGPA20181011 ACTA 034-2018
        Set oPrev = New previo.clsprevio
        oPrev.Show GeneraImpresion, "Posicion de Cliente", True, , gImpresora
        Set oPrev = Nothing
    End If
    
End Sub

Private Sub cmdSaldosConsol_Click()
    
    If LblPersCod <> "" Then
'        frmConsolSaldos.ConsolidaDatos LblPersCod, lblNomPers, lblDocNat, lblDocJur
    End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'Dim oDatos As COMDConstSistema.DCOMGeneral
Dim oCred As COMDCredito.DCOMCreditos
Dim oAho As COMDCaptaGenerales.DCOMCaptaGenerales
    
    TabPosicion.Tab = 0 'JOEP20210511 Segmentacion Prendario Externo
    
    Set oCred = New COMDCredito.DCOMCreditos
    Set oAho = New COMDCaptaGenerales.DCOMCaptaGenerales

    Call oCred.CargarValoresPosicionCliente(nLima, nSaldosAho)
    
'   nSaldosAho = oAho.GetVisualizaSaldoPosicion(gsCodCargo) 'AMDO TI-ERS086-2013  20130726
    Set oCred = Nothing

    CentraForm Me
    
    'Modificado AMDO TI-ERS086-2013 20130726
    '    If nSaldosAho = 1 Then
    '        cmdSaldosConsol.Visible = True
    '    End If
    nSaldosAho = 0
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    'bPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gPerPosCliente)
    bPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gPerPosCliente, gsCodPersUser) 'RIRO20141027 ERS159
    
    If bPermisoCargo Then
        cmdSaldosConsol.Visible = True
        nSaldosAho = 1
    End If
    'END AMDO************

    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    sCalifiFinal = ""
End Sub

Private Sub lstAhorros_DblClick()
    If Me.lstAhorros.SelectedItem.SubItems(4) <> "" Then
        'WIOR 20130517 ************************
    Dim oNCapta As COMNCaptaGenerales.NCOMCaptaMovimiento
            
        Set oNCapta = New COMNCaptaGenerales.NCOMCaptaMovimiento
            
        If Mid(Trim(lstAhorros.SelectedItem.SubItems(4)), 6, 3) = "233" Then
            If oNCapta.EsDepositoGarantia(Trim(lstAhorros.SelectedItem.SubItems(4))) Then
                MsgBox "La Cuenta fue Aperturada como Depósito en Garantía, para garantizar una Carta Fianza.", vbInformation, "Aviso"
            End If
        End If
        'WIOR FIN ************************
        
        'Add by GITU 2013-05-28 Valida si es del area de Prevencion de Lavado de Activo
        Dim lsNomProd As String
        lsNomProd = IIf(Mid(Me.lstAhorros.SelectedItem.SubItems(4), 6, 3) = 232, "Ahorros", (IIf(Mid(Me.lstAhorros.SelectedItem.SubItems(4), 6, 3) = 233, "Plazo Fijo", "CTS")))
        frmCapMantenimiento.Caption = lsNomProd
        If oNCapta.EsOficialOAuxiliardeCumplimiento(gsCodCargo) Then
            Call frmCapMantenimiento.MuestraPosicionCliente(Me.lstAhorros.SelectedItem.SubItems(4), True, True)
        Else
            Call frmCapMantenimiento.MuestraPosicionCliente(Me.lstAhorros.SelectedItem.SubItems(4), False, True)
        End If
        'End GITU
    End If
End Sub

Private Sub LstCartaFianza_DblClick()
    If LstCartaFianza.SelectedItem.SubItems(2) <> "" Then
         'margERS41--
        Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
        Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
        If oCaja.PermitirVisualizarEstadoExpediente(gsCodUser, gsCodAge, gsCodArea, gsCodCargo) Then 'checkear si no ingresa
            If Me.LstCartaFianza.SelectedItem.SubItems(4) = "APROBADO" Then
                Dim Estado As Integer
                Call get_EstadoExpediente(Me.LstCartaFianza.SelectedItem.SubItems(2), Estado)
            End If
        End If
        'end marg
        
        Call frmCFHistorial.CargaCFHistorial(LstCartaFianza.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub LstCreditos_DblClick()
    If Me.lstCreditos.SelectedItem.SubItems(2) <> "" Then
        'marg-ERS041--
        Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
        Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral 'ARLO 20170926 ERS060-2017
        Dim oCreditos As COMDCredito.DCOMCreditos       'ARLO 20170926 ERS060-2017
        Set oCreditos = New COMDCredito.DCOMCreditos
        If Not oCreditos.VerificaClienteCampania(Me.lstCreditos.SelectedItem.SubItems(2)) Then 'ARLO 20170926 ERS060-2017
            If oCaja.PermitirVisualizarEstadoExpediente(gsCodUser, gsCodAge, gsCodArea, gsCodCargo) Then 'checkear si no ingresa
                If Me.lstCreditos.SelectedItem.SubItems(5) = "APROBADO" Then
                    Dim Estado As Integer
                    Call get_EstadoExpediente(Me.lstCreditos.SelectedItem.SubItems(2), Estado)
                End If
            End If
            Set oCreditos = Nothing
        End If
        '</marg
        
        Call frmCredConsulta.ConsultaCliente(Me.lstCreditos.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub lstJudicial_DblClick()
    If lstJudicial.SelectedItem.SubItems(2) <> "" Then
        Call frmColRecRConsulta.MuestraPosicionCliente(lstJudicial.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub lstPrendario_DblClick()

    If lstPrendario.SelectedItem.SubItems(3) <> "" Then
        
        Call frmColPMantPrestamoPig.BuscaContrato(lstPrendario.SelectedItem.SubItems(3), 1, LblPersCod)
        frmColPMantPrestamoPig.AxCodCta.NroCuenta = lstPrendario.SelectedItem.SubItems(3)
        frmColPMantPrestamoPig.Show 1
        'Call frmColPContratosxCliente
    End If
End Sub

'**DAOR 20070512, Procedimiento que imprime el estado de cuenta de crédito
Public Sub ImprimeEstadoCuentaCredito(ByVal psCtaCod As String)
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim nTasaCompAnual As Double
        
    Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\EstadoCuentaCredito.doc")
        
    With oWord.Selection.Find
        .Text = "<<cFecha>>"
        .Replacement.Text = Format(gdFecSis, "dd/mm/yyyy")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosParaEstadoCuentaCredito(psCtaCod)
    Set oDCred = Nothing
        
    If Not (R.EOF And R.BOF) Then
        With oWord.Selection.Find
            .Text = "<<cNomCli>>"
            .Replacement.Text = R!cnomcli
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<cDirCli>>"
            .Replacement.Text = R!cDirCli
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<cCtaCod>>"
            .Replacement.Text = R!cCtaCod
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<cNomAge>>"
            .Replacement.Text = R!vAgencia
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<cMoneda>>"
            .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = 1, "NUEVOS SOLES", "DOLARES")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<nMonDes>>"
            .Replacement.Text = Format(R!nMontoCol, "#0.00")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        nTasaCompAnual = Format(((1 + R!nTasaInteres / 100) ^ (360 / 30) - 1) * 100, "#.00")
        With oWord.Selection.Find
            .Text = "<<nTEA>>"
            .Replacement.Text = nTasaCompAnual
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<nCEA>>"
            .Replacement.Text = R!nTasCosEfeAnu
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nSaldo>>"
           .Replacement.Text = Format(R!nSaldo, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
                     
        With oWord.Selection.Find
           .Text = "<<dFecVenPag>>"
           .Replacement.Text = IIf(DateDiff("d", R!dVencPag, "1900-01-01") = 0, "", Format(R!dVencPag, "dd/mm/yyyy"))
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nCapPag>>"
           .Replacement.Text = Format(R!nCapPag, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nIntPag>>"
           .Replacement.Text = Format(R!nIntPag, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nMorPag>>"
           .Replacement.Text = Format(R!nMorPag, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nSegDesPag>>"
           .Replacement.Text = Format(R!nSegDesPag, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nSegBiePag>>"
           .Replacement.Text = Format(0, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nComPorPag>>"
           .Replacement.Text = Format(0, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nItfPag>>"
           .Replacement.Text = Format(R!nItfPag, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<dFecPagPag>>"
           .Replacement.Text = IIf(DateDiff("d", R!dVencPag, "1900-01-01") = 0, "", Format(R!dPago, "dd/mm/yyyy"))
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
           .Text = "<<dFecVenAPag>>"
           .Replacement.Text = Format(R!dVencAPag, "dd/mm/yyyy")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
           .Text = "<<nMontoAPag>>"
           .Replacement.Text = Format(R!nMontoAPag, "#0.00")
           .Forward = True
           .Wrap = wdFindContinue
           .Format = False
           .Execute Replace:=wdReplaceAll
        End With
    End If
    
    oDoc.SaveAs (App.Path & "\FormatoCarta\EstadoCuentaCredito_" & psCtaCod & ".doc")
    oDoc.Close
    Set oDoc = Nothing
    
    Set oWord = CreateObject("Word.Application")
        oWord.Visible = True
    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\EstadoCuentaCredito_" & psCtaCod & ".doc")
    Set oDoc = Nothing
    Set oWord = Nothing
    
End Sub
'***Agregado por ELRO el 20121112, según OYP-RFC115-2012
Public Sub iniciarFormulario(ByVal psPersCod As String)
Dim oDCOMPersonas As New COMDPersona.DCOMPersonas
Dim rsPersona As New ADODB.Recordset
Dim oPersona As New COMDPersona.UCOMPersona
Dim oPersona2 As New COMDPersona.UCOMPersona
Dim sPersCod As String


Set rsPersona = oDCOMPersonas.BuscaCliente(psPersCod, BusquedaCodigo)
Set oDCOMPersonas = Nothing

oPersona2.CargaDatos rsPersona!cPersCod, _
                     rsPersona!cPersNombre, _
                     Format(IIf(IsNull(rsPersona!dPersNacCreac), gdFecSis, rsPersona!dPersNacCreac), "dd/mm/yyyy"), _
                     IIf(IsNull(rsPersona!cPersDireccDomicilio), "", rsPersona!cPersDireccDomicilio), _
                     IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono), rsPersona!nPersPersoneria, _
                     IIf(IsNull(rsPersona!cPersIDnroDNI), "", rsPersona!cPersIDnroDNI), _
                     IIf(IsNull(rsPersona!cPersIDnroRUC), "", rsPersona!cPersIDnroRUC), _
                     IIf(IsNull(rsPersona!cPersIDnro), "", rsPersona!cPersIDnro), _
                     IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo), _
                     IIf(IsNull(rsPersona!cActiGiro1), "", rsPersona!cActiGiro1), _
                     IIf(IsNull(rsPersona!nTipoId), "1", rsPersona!nTipoId)

    lstCreditos.Enabled = False
    lstAhorros.Enabled = False
    LstCartaFianza.Enabled = False
    lstJudicial.Enabled = False
    lstPrendario.Enabled = False
    
    Set oPersona = oPersona2
    If Not oPersona Is Nothing Then
        LblPersCod.Caption = oPersona.sPersCod
        lblNomPers.Caption = oPersona.sPersNombre
        lblDocNat.Caption = Trim(oPersona.sPersIdnroDNI)
        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
        lsPersTDoc = "1"
        If oPersona.sPersPersoneria = "1" Then
            fbPersNatural = True
            If Trim(oPersona.sPersIdnroDNI) = "" Then
                If Not Trim(oPersona.sPersIdnroOtro) = "" Then
                    lblDocNat.Caption = Trim(oPersona.sPersIdnroOtro)
                    lsPersTDoc = Trim(oPersona.sPersTipoDoc)
                End If
            End If
        Else
            fbPersNatural = False
            lsPersTDoc = "3"
        End If
    Else
        Exit Sub
    End If
    sPersCod = oPersona.sPersCod
    Set oPersona = Nothing
        
    If sPersCod <> "" Then
        Call BuscarPosicionCliente(sPersCod, lsPersTDoc)
    End If
    
    If sPersCod <> "" Then
        cmdSaldosConsol.Enabled = True
        cmdImprimir.Enabled = True
    End If
    
    cmdBuscar.Enabled = False
    TabPosicion.Tab = 1
    Show 1
End Sub
'***Fin Agregado por ELRO el 20121112*******************

'marg ERS041 07-06-2016------
Private Sub get_EstadoExpediente(ByVal psCta As String, ByRef Estado As Integer)

Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
Dim c As New ADODB.Recordset
Dim d As New ADODB.Recordset
Dim ubicacion As String
Dim observacion As String
'Dim estado As Integer
Dim count As Integer

    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set c = oCF.get_CredAdmControlDesembolso(psCta)
    
    Set d = oCF.get_ControlCreditosObsAdmCred(psCta)
        If Not d.BOF And Not d.EOF Then
        For count = 1 To d.RecordCount
            If (d!nRegulariza = 0) Then
                 observacion = observacion & d!cDescripcion & vbCrLf
            End If
            d.MoveNext
        Next
    End If
    
    If Not c.BOF And Not c.EOF Then
       If (IsNull(c!dIngreso) And IsNull(c!dUltSalidaObs) And IsNull(c!dUltIngresoObs) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente aun se encuentra en Comité de Créditos y está pendiente de revisión por la Administración de Créditos"
           Estado = 0
       End If
       If (Not (IsNull(c!dIngreso)) And IsNull(c!dUltSalidaObs) And IsNull(c!dUltIngresoObs) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente ingresó al area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & IIf(observacion <> "", " y tiene las siguientes observaciones:", " para su respectiva observación")
           Estado = 1
       End If
       If (Not (IsNull(c!dIngreso)) And Not (IsNull(c!dUltSalidaObs)) And IsNull(c!dUltIngresoObs) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente salió del area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & " por las siguientes observaciónes:"
           Estado = 2
       End If
        If (Not (IsNull(c!dIngreso)) And (IsNull(c!dUltSalidaObs)) And Not (IsNull(c!dUltIngresoObs)) And IsNull(c!dSalida)) Then
           ubicacion = "El Expediente reigresó al area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & " para su respectivo levantamiento de observaciones:"
           Estado = 3
       End If
       If (Not (IsNull(c!dSalida))) Then
           ubicacion = "El Expediente salió del area de Administración de Créditos el " & Format(c!dIngreso, "dd/mm/yyyy") & " para su respectivo desembolso"
           Estado = 4
       End If
    Else
        ubicacion = "El Expediente aun se encuentra en Comité de Créditos y está pendiente de revisión por la Administración de Créditos"
        Estado = 0
    End If
    
    
'    If (Estado <> 4) Then
        MsgBox ubicacion & vbCrLf & observacion, vbInformation, "Aviso"
'    End If
End Sub
'</marg-------------------
