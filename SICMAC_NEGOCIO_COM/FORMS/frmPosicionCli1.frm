VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPosicionCli1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posición del Cliente"
   ClientHeight    =   7380
   ClientLeft      =   960
   ClientTop       =   1860
   ClientWidth     =   10080
   Icon            =   "frmPosicionCli1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
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
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmPosicionCli1.frx":030A
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
         Picture         =   "frmPosicionCli1.frx":0387
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
         Width           =   1965
      End
      Begin TabDlg.SSTab TabPosicion 
         Height          =   3960
         Left            =   105
         TabIndex        =   19
         Top             =   480
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   6985
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
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
         TabPicture(0)   =   "frmPosicionCli1.frx":E51D
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LstCreditos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   " &Ahorros   "
         TabPicture(1)   =   "frmPosicionCli1.frx":E539
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstAhorros"
         Tab(1).Control(1)=   "lblDolaresAho"
         Tab(1).Control(2)=   "lblSolesAho"
         Tab(1).Control(3)=   "Label6"
         Tab(1).Control(4)=   "Label4"
         Tab(1).Control(5)=   "Label5"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "&Pignoraticio   "
         TabPicture(2)   =   "frmPosicionCli1.frx":E555
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstPrendario"
         Tab(2).Control(1)=   "lblDolaresPig"
         Tab(2).Control(2)=   "lblSolesPig"
         Tab(2).Control(3)=   "Label8"
         Tab(2).Control(4)=   "Label7"
         Tab(2).Control(5)=   "Label3"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   " Créditos Judiciales   "
         TabPicture(3)   =   "frmPosicionCli1.frx":E571
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lstJudicial"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Carta Fianza"
         TabPicture(4)   =   "frmPosicionCli1.frx":E58D
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "LstCartaFianza"
         Tab(4).ControlCount=   1
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
            NumItems        =   17
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nro."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Fecha Desembolso"
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
               Text            =   "Tipo Crédito"
               Object.Width           =   6174
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
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Nota"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "Monto "
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Saldo Cap."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Cod. Ant 1"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Cod. Ant 2"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Moneda"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "FechaCancelacion"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   15
               Text            =   "Fec.Solicitud"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "RFA"
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
            NumItems        =   12
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
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "SaldoDisp"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Motivo de Bloque"
               Object.Width           =   7231
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
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
            NumItems        =   10
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
               Object.Width           =   6174
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
               SubItemIndex    =   8
               Text            =   "Fecha Castigado"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Ult Fecha Pago"
               Object.Width           =   2540
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
            Top             =   3360
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
            Left            =   -70725
            TabIndex        =   33
            Top             =   3360
            Width           =   2145
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "DOLARES"
            Height          =   195
            Left            =   -68460
            TabIndex        =   32
            Top             =   3450
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "SOLES"
            Height          =   195
            Left            =   -71355
            TabIndex        =   31
            Top             =   3450
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
            Left            =   -73470
            TabIndex        =   27
            Top             =   3450
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
      Enabled         =   -1  'True
      RightMargin     =   35000
      TextRTF         =   $"frmPosicionCli1.frx":E5A9
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
      Top             =   7110
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10583
            Picture         =   "frmPosicionCli1.frx":E62A
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
         Left            =   4440
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
         Caption         =   "Doc. Jurídico :"
         Height          =   195
         Left            =   3045
         TabIndex        =   4
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Natural :"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   660
         Width           =   990
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
Attribute VB_Name = "frmPosicionCli1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLima As Integer
Dim nSaldosAho As Integer


Private Function GeneraImpresion() As String
Dim I As Integer
Dim ContLineas As Integer
Dim R As ADODB.Recordset
Dim oCred As DCredito

    GeneraImpresion = ""
    GeneraImpresion = oImpresora.gPrnSaltoLinea
    
    GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    
    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CREDITOS : " & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Nota", 8) & ImpreFormat("Monto", 10) & ImpreFormat("Saldo Cap", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    ContLineas = 12
    If LstCreditos.ListItems.Count > 0 Then
        
        For I = 1 To LstCreditos.ListItems.Count
            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(LstCreditos.ListItems(I).Text, 5)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(1), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(4), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(LstCreditos.ListItems(I).SubItems(3)), "AGENCIA", ""), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(2), 20)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(5), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(6), 9)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(7), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(8), 3)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(LstCreditos.ListItems(I).SubItems(9)), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(LstCreditos.ListItems(I).SubItems(10)), 10)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCreditos.ListItems(I).SubItems(13), 8) & oImpresora.gPrnSaltoLinea
            ContLineas = ContLineas + 1
            
            If ContLineas > 56 Then
                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CREDITOS : " & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Nota", 8) & ImpreFormat("Monto", 10) & ImpreFormat("Saldo Cap", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                ContLineas = 13
            End If
            
        Next I
        
        '--------------GARANTIAS----------------------
        Set R = New ADODB.Recordset
        Set oCred = New DCredito
        GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
        GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE GARANTIAS : " & oImpresora.gPrnSaltoLinea
        GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
        GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("MONTO REALIZ.", 13) & ImpreFormat("MONTO DISPON.", 13) & oImpresora.gPrnSaltoLinea
        GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 5
            Set R = oCred.RecuperaGarantiasCreditoConsol(LblPersCod.Caption, gdFecSis)
            Do While Not R.EOF
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(R!nRealizacion, 8, , True)
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(R!nPorGravar, 8, , True)
                ContLineas = ContLineas + 1
                If ContLineas > 56 Then
                    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                    GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE GARANTIAS : " & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("MONTO REALIZ.", 13) & ImpreFormat("MONTO DISPON.", 13) & oImpresora.gPrnSaltoLinea
                    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                    ContLineas = 13
                End If
                R.MoveNext
            Loop
        Set oCred = Nothing
        Set R = Nothing
        
    Else
        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas de Creditos" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 2
    End If
    
    '********** A H O R R O S **********************
    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE AHORROS : " & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Producto", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10)
    If nSaldosAho = 1 Then
        GeneraImpresion = GeneraImpresion & ImpreFormat("Saldo Cont", 10) & Space(4) & ImpreFormat("Saldo Disp", 10)
    End If
    GeneraImpresion = GeneraImpresion & ImpreFormat("Motiv. Bloqueo", 15) & Space(4) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    
    ContLineas = ContLineas + 6
    If lstAhorros.ListItems.Count > 0 Then
        For I = 1 To lstAhorros.ListItems.Count
            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstAhorros.ListItems(I).Text, 5)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(1), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(2), 10)
            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstAhorros.ListItems(I).SubItems(3)), "AGENCIA", ""), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(4), 20)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(6), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(7), 12)
            If nSaldosAho = 1 Then
                GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(8), 12, 2)
                GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(9), 12, 2)
            End If
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(10), 17)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstAhorros.ListItems(I).SubItems(11), 10) & oImpresora.gPrnSaltoLinea
            ContLineas = ContLineas + 1
            If ContLineas > 56 Then
                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE AHORROS : " & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Producto", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10)
                If nSaldosAho = 1 Then
                    GeneraImpresion = GeneraImpresion & ImpreFormat("SaldoCont", 10) & Space(4) & ImpreFormat("Saldo Dispon", 8)
                End If
                GeneraImpresion = GeneraImpresion & ImpreFormat("Motiv. Bloqueo", 15) & Space(4) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                ContLineas = 13
            End If
        Next I
        
        GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
        GeneraImpresion = GeneraImpresion & Space(5) & "TOTAL AHORROS:   SOLES=" & Trim(lblSolesAho.Caption) & Space(10 - Len(lblSolesAho.Caption)) & " DOLARES=" & Trim(lblDolaresAho.Caption) & oImpresora.gPrnSaltoLinea
        
    Else
        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas de Ahorros" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 2
    End If
    
    'Para Pignoraticio
    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE PIGNORATICIO : " & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("No Renov.", 10) & ImpreFormat("Monto", 6) & ImpreFormat("Fecha Venc", 12) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Tasación", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    ContLineas = ContLineas + 6
    If lstPrendario.ListItems.Count > 0 Then
        For I = 1 To lstPrendario.ListItems.Count
            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstPrendario.ListItems(I).Text, 5)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(1), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstPrendario.ListItems(I).SubItems(2)), "AGENCIA", ""), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(3), 20)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(4), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(5), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(6), 3)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstPrendario.ListItems(I).SubItems(7)), 10)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(8), 11)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstPrendario.ListItems(I).SubItems(10)), 10)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstPrendario.ListItems(I).SubItems(11)), 10)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstPrendario.ListItems(I).SubItems(12), 11) & oImpresora.gPrnSaltoLinea
            
            ContLineas = ContLineas + 1
            If ContLineas > 56 Then
                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE PIGNORATICIO : " & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 10) & ImpreFormat("Estado", 10)
                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("No Renov.", 10) & ImpreFormat("Monto", 8) & ImpreFormat("Fecha Venc", 10) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Tasación", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                ContLineas = 13
            End If
        Next I
        GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
        GeneraImpresion = GeneraImpresion & Space(5) & "TOTAL PIGNORATICIO:   SOLES=" & Trim(lblSolesPig.Caption) & Space(10 - Len(lblSolesPig.Caption)) & " DOLARES=" & Trim(lblDolaresPig.Caption) & oImpresora.gPrnSaltoLinea
    Else
        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas de Pignoraticio" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 2
    End If
    
    'Judicial
    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE JUDICIAL : " & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 13)
    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Fecha Cast.", 10) & ImpreFormat("Fecha Ult. Pago", 10) & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    ContLineas = ContLineas + 4
    
    If lstJudicial.ListItems.Count > 0 Then
        For I = 1 To lstJudicial.ListItems.Count
            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(lstJudicial.ListItems(I).Text, 5)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(1), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(4), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(lstJudicial.ListItems(I).SubItems(3)), "AGENCIA", ""), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(2), 20)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(5), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(6), 9)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(lstJudicial.ListItems(I).SubItems(7)), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(8), 11)
            GeneraImpresion = GeneraImpresion & ImpreFormat(lstJudicial.ListItems(I).SubItems(9), 10) & oImpresora.gPrnSaltoLinea
            ContLineas = ContLineas + 1
            If ContLineas > 56 Then
                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE JUDICIAL : " & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Tipo", 10) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 10) & ImpreFormat("Estado", 10)
                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Saldo Cap.", 10) & ImpreFormat("Fecha Cast.", 10) & ImpreFormat("Fecha Ult. Pago", 10) & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                ContLineas = 13
            End If
        Next I
    Else
        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cuentas en Judicial" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 2
    End If
    
    '***************** CARTA FIANZA ************************
    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CARTA FIANZA : " & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
    GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Monto", 10) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    ContLineas = ContLineas + 6
    If LstCartaFianza.ListItems.Count > 0 Then
        
        For I = 1 To LstCartaFianza.ListItems.Count
            GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat(LstCartaFianza.ListItems(I).Text, 5)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(I).SubItems(1), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(Replace(UCase(LstCartaFianza.ListItems(I).SubItems(3)), "AGENCIA", ""), 8)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(I).SubItems(2), 20)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(I).SubItems(4), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(I).SubItems(5), 12)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(I).SubItems(6), 3)
            GeneraImpresion = GeneraImpresion & ImpreFormat(CDbl(LstCartaFianza.ListItems(I).SubItems(7)), 10)
            GeneraImpresion = GeneraImpresion & ImpreFormat(LstCartaFianza.ListItems(I).SubItems(8), 8) & oImpresora.gPrnSaltoLinea
            ContLineas = ContLineas + 1
            
            If ContLineas > 56 Then
                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE CARTA FIANZA : " & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & ImpreFormat("NRO", 5) & ImpreFormat("Fecha", 12) & ImpreFormat("Agencia", 10) & ImpreFormat("Cuenta", 20) & ImpreFormat("Estado", 10)
                GeneraImpresion = GeneraImpresion & ImpreFormat("Particip", 10) & ImpreFormat("Analista", 10) & ImpreFormat("Monto", 15) & ImpreFormat("Moneda", 10) & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                ContLineas = 13
            End If
            
        Next I
    Else
        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Cartas Fianza" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 2
    End If
    
    '***************** COMENTARIOS ************************
    Dim sCadena As String, sLinea As String
    Dim nPos As Integer
    GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE COMENTARIO : " & oImpresora.gPrnSaltoLinea
    GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
    ContLineas = ContLineas + 5
    sCadena = Trim(rtfComentario.Text)
    If sCadena <> "" Then
        Do
            nPos = InStr(1, sCadena, oImpresora.gPrnSaltoLinea, vbTextCompare)
            If nPos <> 0 Then
                sLinea = Mid(sCadena, 1, nPos)
                sCadena = Mid(sCadena, nPos + 1, Len(sCadena) - nPos)
            Else
                sLinea = sCadena
                sCadena = ""
            End If
            GeneraImpresion = GeneraImpresion & Space(5) & sLinea
            
            ContLineas = ContLineas + 1
            
            If ContLineas > 56 Then
                GeneraImpresion = GeneraImpresion & oImpresora.gPrnSaltoPagina
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomCmac & Space(90) & "Fecha : " & Format(FechaHora(gdFecSis), "dd/mm/yyyy hh:mm:ss") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & gsNomAge & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(60) & "REPORTE DE POSICION DE CLIENTE" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Cliente : " & LblPersCod.Caption & Space(2) & lblNomPers.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "Documentos :   DNI :" & lblDocNat.Caption & Space(2) & "RUC :" & lblDocJur.Caption & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & "SECCION DE COMENTARIOS : " & oImpresora.gPrnSaltoLinea
                GeneraImpresion = GeneraImpresion & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
                ContLineas = 11
            End If
                
        Loop Until sCadena = ""
    Else
        GeneraImpresion = GeneraImpresion & Space(5) & "Cliente No Posee Comentarios" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        ContLineas = ContLineas + 2
    End If
End Function

Private Sub BuscaJudicial(ByVal psPersCod As String)
Dim oCreditos As DCreditos
Dim R As ADODB.Recordset
Dim L As ListItem

    lstJudicial.ListItems.Clear
    Set oCreditos = New DCreditos
    Set R = oCreditos.DatosPosicionClienteJudicial(psPersCod)
    
    If Not (R.EOF And R.BOF) Then
        lstJudicial.Enabled = True
    End If
    
    Do While Not R.EOF
        
        Set L = lstJudicial.ListItems.Add(, , R.Bookmark)
        
        L.SubItems(1) = Format(R!dIngRecup, "dd/mm/yyyy") 'fecha Vigencia
        L.SubItems(2) = R!cCtaCod  'Cta
        L.SubItems(3) = R!cAgeDescripcion 'Agencia
        L.SubItems(4) = R!cTipoCred 'Tipo de Credito
        L.SubItems(5) = R!cEstado 'Estado
        L.SubItems(6) = R!cParticip 'Participacion
        L.SubItems(7) = Format(R!nSaldo, "#0.00") ' Saldo Cap
        If IsNull(R!dFecCast) Then
            L.SubItems(8) = "" 'Fecha Venc
        Else
            L.SubItems(8) = Format(R!dFecCast, "dd/mm/yyyy") 'Fecha Venc
        End If
        If IsNull(R!cmovnro) Then
            L.SubItems(9) = ""
        Else
            L.SubItems(9) = Mid(R!cmovnro, 7, 2) & "/" & Mid(R!cmovnro, 5, 2) & "/" & Mid(R!cmovnro, 1, 4) 'Ultimo Movimiento
        End If
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCreditos = Nothing
    Exit Sub


End Sub

Private Sub BuscaPignoraticio(ByVal psPersCod As String)
Dim oCreditos As DCreditos
Dim R As ADODB.Recordset
Dim L As ListItem
Dim SumaSol As Double, SumaDol As Double

    SumaSol = 0
    SumaDol = 0
    
    lstPrendario.ListItems.Clear
    Set oCreditos = New DCreditos
    
    If nLima = 1 Then           'Trujillo
        Set R = oCreditos.DatosPosicionClientePigno(psPersCod)
    Else 'Lima
        Set R = oCreditos.DatosPosicionClientePignoLima(psPersCod)
    End If
    
    If Not (R.EOF And R.BOF) Then
        lstPrendario.Enabled = True
        Do While Not R.EOF
            
            Set L = lstPrendario.ListItems.Add(, , R.Bookmark)
            
            L.SubItems(1) = Format(R!dVigencia, "dd/mm/yyyy") 'fecha Vigencia
            L.SubItems(2) = R!cAgeDescripcion 'Agencia
            L.SubItems(3) = R!cCtaCod 'Cuenta
            L.SubItems(4) = R!cEstado 'estado
            L.SubItems(5) = R!cParticip 'Participacion
            L.SubItems(6) = Trim(Str(R!nNroRenov)) 'Nro Renovacion
            L.SubItems(7) = Format(R!nMontoCol, "#0.00") ' Prestamo
            L.SubItems(8) = Format(R!dvenc, "dd/mm/yyyy") 'Fecha Venc
            L.SubItems(9) = R!cAgeCod 'Cod Agencia
            If Not (R!nPrdEstado = gColPEstRemat Or R!nPrdEstado = gColPEstSubas Or R!nPrdEstado = gColPEstAdjud Or R!nPrdEstado = gColPEstChafa Or R!nPrdEstado = "2113") Then
                L.SubItems(10) = Format(IIf(IsNull(R!nSaldo), 0, R!nSaldo), "#0.00") 'Saldo Capital
            Else
                L.SubItems(10) = Format("0.00", "#0.00")    'Saldo Capital
            End If
            
            L.SubItems(11) = Format(IIf(IsNull(R!nTasacion), 0, R!nTasacion), "#0.00")
            L.SubItems(12) = R!sMoneda
            
            
            SumaSol = SumaSol + R!nSaldo
            
            
            If R!nPrdEstado = gColPEstDesem Or R!nPrdEstado = gColPEstRenov Then
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
            
            R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
    Set oCreditos = Nothing
    lblSolesPig.Caption = Format(SumaSol, "#0.00")
    
    Exit Sub

End Sub

Private Sub BuscaAhorros(ByVal psPersCod As String)
Dim oCreditos As DCreditos
Dim R As ADODB.Recordset
Dim L As ListItem
Dim SumaSol As Double, SumaDol As Double

SumaSol = 0
SumaDol = 0
On Error GoTo ERRORBuscaCreditos
    lstAhorros.ListItems.Clear
    Set oCreditos = New DCreditos
    Set R = oCreditos.DatosPosicionClienteAhorro(psPersCod)
    
    If Not (R.EOF And R.BOF) Then
        lstAhorros.Enabled = True
    End If
    
    Do While Not R.EOF
        
        Set L = lstAhorros.ListItems.Add(, , R.Bookmark)
        
        L.SubItems(1) = Format(R!dApertura, "dd/mm/yyyy")
        L.SubItems(2) = R!cTipoAho
        L.SubItems(3) = R!cAgeDescripcion
        L.SubItems(4) = R!cCtaCod
        L.SubItems(5) = IIf(IsNull(R!CCTACODANT), "", R!CCTACODANT)
        L.SubItems(6) = R!cEstado
        L.SubItems(7) = R!cParticip
        If nSaldosAho = 1 Then       'Lima
            L.SubItems(8) = R!nsaldocont
            L.SubItems(9) = R!nSaldoDisp
        End If
        
        If R!sMoneda = "SOLES" Then
            SumaSol = SumaSol + R!nSaldoDisp
        Else
            SumaDol = SumaDol + R!nSaldoDisp
        End If
        
        
        L.SubItems(10) = IIf(IsNull(R!cMotivo), "", R!cMotivo)
        L.SubItems(11) = R!sMoneda 'Moneda
        
        If UCase(R!cEstado) = "ACTIVA" Then
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
        End If
        
        R.MoveNext
    Loop
    
    lblSolesAho.Caption = Format(SumaSol, "#0.00")
    
    lblDolaresAho.Caption = Format(SumaDol, "#0.00")
    
    R.Close
    Set R = Nothing
    Set oCreditos = Nothing
    Exit Sub
    
ERRORBuscaCreditos:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub
Private Sub BuscaCreditos(ByVal psPersCod As String)
Dim oCreditos As DCreditos
Dim R As ADODB.Recordset
Dim L As ListItem

On Error GoTo ERRORBuscaCreditos
    LstCreditos.ListItems.Clear
    Set oCreditos = New DCreditos
    Set R = oCreditos.DatosPosicionCliente(psPersCod, IIf(Check1.value = 1, True, False))
    
    If Not (R.EOF And R.BOF) Then
        LstCreditos.Enabled = True
    End If
    
    Do While Not R.EOF
        Set L = LstCreditos.ListItems.Add(, , R.Bookmark)
        If IsNull(R!dAprobacion) Then
            L.SubItems(1) = ""
        Else
            L.SubItems(1) = Format(R!dAprobacion, "dd/mm/yyyy")
        End If
        
        'L.SubItems(1) = Format(R!dsolicitado, "dd/mm/yyyy") 'fecha de Solicitud
        L.SubItems(2) = R!cCtaCod 'Credito
        L.SubItems(3) = R!cAgeDescripcion 'Agencia
        L.SubItems(4) = R!cTipoCred 'Tipo de Credito (MES, COM, CON)
        L.SubItems(5) = IIf(IsNull(R!cEstadoDesc), "", R!cEstadoDesc) 'Estado
        L.SubItems(6) = R!cRelacionDesc 'Participacion
        L.SubItems(7) = IIf(IsNull(R!cPersAnalista), "", R!cPersAnalista) 'Analista
        L.SubItems(8) = IIf(IsNull(R!nAnalistaNota), "", R!nAnalistaNota) 'Nota
        L.SubItems(9) = Format(R!nPrestamo, "0.00")
        L.SubItems(10) = Format(R!nSaldo, "#0.00") 'Saldo Capital
        L.SubItems(11) = IIf(IsNull(R!cCodAnt1), "", R!cCodAnt1) 'Codigo Antiguo 1
        L.SubItems(12) = IIf(IsNull(R!cCodAnt2), "", R!cCodAnt2) 'Codigo Antiguo 2
        L.SubItems(13) = R!sMoneda 'Moneda
        If IsNull(R!dCancelado) Then
            L.SubItems(14) = ""
        Else
            L.SubItems(14) = Format(R!dCancelado, "dd/mm/yyyy") 'Fec Cancelacion
        End If
        L.SubItems(15) = Format(R!dsolicitado, "dd/mm/yyyy") 'fecha de Solicitud
'        If IsNull(R!dAprobacion) Then
'            L.SubItems(15) = ""
'        Else
'            L.SubItems(15) = Format(R!dAprobacion, "dd/mm/yyyy")
'        End If
        L.SubItems(16) = R!cRFA
        
        If R!nPrdEstado = gColocEstRefMor Or R!nPrdEstado = gColocEstRefVenc _
            Or R!nPrdEstado = gColocEstRefNorm Or R!nPrdEstado = gColocEstVigVenc _
            Or R!nPrdEstado = gColocEstVigMor Or R!nPrdEstado = gColocEstVigNorm Then
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
            L.ListSubItems(16).Bold = True
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
        End If
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCreditos = Nothing
    Exit Sub
    
ERRORBuscaCreditos:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub BuscaCF(ByVal psPersCod As String)
Dim oCreditos As DCreditos
Dim R As ADODB.Recordset
Dim L As ListItem

On Error GoTo ERRORBuscaCF
    LstCartaFianza.ListItems.Clear
    Set oCreditos = New DCreditos
    Set R = oCreditos.DatosPosicionClienteCF(psPersCod)
    
    If Not (R.EOF And R.BOF) Then
        LstCartaFianza.Enabled = True
    End If
    
    Do While Not R.EOF
        Set L = LstCartaFianza.ListItems.Add(, , R.Bookmark)
        
        L.SubItems(1) = Format(R!dEmitido, "dd/mm/yyyy") 'fecha de Solicitud
        L.SubItems(2) = R!cCtaCod 'Credito
        L.SubItems(3) = R!cAgeDescripcion 'Agencia
        L.SubItems(4) = IIf(IsNull(R!cEstadoDesc), "", R!cEstadoDesc) 'Estado
        L.SubItems(5) = R!cRelacionDesc 'Participacion
        L.SubItems(6) = IIf(IsNull(R!cPersAnalista), "", R!cPersAnalista) 'Analista
        L.SubItems(7) = Format(R!nPrestamo, "0.00")
        L.SubItems(8) = R!sMoneda 'Moneda
        L.SubItems(9) = Format(R!dCancelado, "dd/mm/yyyy") 'fecha de Honrado
        
        If R!nPrdEstado = gColocEstVigNorm Then
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
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oCreditos = Nothing
    Exit Sub
    
ERRORBuscaCF:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub BuscaComentarios(ByVal sPersona As String)
Dim oCreditos As DCreditos
Dim R As ADODB.Recordset
Dim sComentario As String, sFecha As String, sUsuario As String
Dim sCadena As String
Dim I As Integer

    rtfComentario.Text = ""
    Set oCreditos = New DCreditos
    Set R = oCreditos.DatosPosicionClienteComentarios(sPersona)
    If Not (R.EOF And R.BOF) Then
        rtfComentario.Enabled = True
    Else
        rtfComentario.Enabled = False
    End If
    sCadena = ""
    I = 0
    Do While Not R.EOF
        sComentario = Trim(R("cComentario"))
        sFecha = Mid(R("cMovNro"), 7, 2) & "/" & Mid(R("cMovNro"), 5, 2) & "/" & Mid(R("cMovNro"), 1, 4)
        sFecha = sFecha & " " & Mid(R("cMovNro"), 9, 2) & ":" & Mid(R("cMovNro"), 11, 2) & ":" & Mid(R("cMovNro"), 13, 2)
        sUsuario = Right(R("cMovNro"), 4)
        If sCadena <> "" Then sCadena = sCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        I = I + 1
        sCadena = sCadena & "COMENTARIO " & Format$(I, "00") & " : " & oImpresora.gPrnSaltoLinea
        sCadena = sCadena & sComentario & oImpresora.gPrnSaltoLinea
        sCadena = sCadena & "USUARIO : " & sUsuario & "  FECHA : " & sFecha
        R.MoveNext
    Loop
    If sCadena <> "" Then
        rtfComentario.Text = sCadena
    End If
End Sub

Private Sub CmdBuscar_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim sPersCod As String

    LstCreditos.Enabled = False
    lstAhorros.Enabled = False
    LstCartaFianza.Enabled = False
    lstJudicial.Enabled = False
    lstPrendario.Enabled = False
    
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        LblPersCod.Caption = oPersona.sPersCod
        lblNomPers.Caption = oPersona.sPersNombre
        lblDocNat.Caption = Trim(oPersona.sPersIdnroDNI)
        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
    Else
        Exit Sub
    End If
    sPersCod = oPersona.sPersCod
    Set oPersona = Nothing
    Call BuscaCreditos(sPersCod)
    Call BuscaAhorros(sPersCod)
    Call BuscaPignoraticio(sPersCod)
    Call BuscaJudicial(sPersCod)
    Call BuscaCF(sPersCod)
    Call BuscaComentarios(sPersCod)
    
    If sPersCod <> "" Then
        cmdSaldosConsol.Enabled = True
        cmdImprimir.Enabled = True
    End If
    
End Sub

Private Sub CmdHistorial_Click()
    Dim oNCredDoc As NCredDoc
    Dim sCadImp As String
    Dim oPrev As clsPrevio
    
    Set oNCredDoc = New NCredDoc
        sCadImp = oNCredDoc.ImpreRepor_HistorialCliente(LblPersCod.Caption, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
    Set oNCredDoc = Nothing
    
    If Len(sCadImp) = 0 Then
        MsgBox "No se encontraron datos del reporte", vbInformation, "AVISO"
    Else
        Set oPrev = New clsPrevio
        oPrev.Show sCadImp, "LISTA DE CREDITOS DEL CLIENTE", True
        Set oPrev = Nothing
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim oPrev As Previo.clsPrevio

    If LblPersCod <> "" Then
        Set oPrev = New Previo.clsPrevio
        oPrev.Show GeneraImpresion, "Posicion de Cliente", True
        Set oPrev = Nothing
    End If
    
End Sub

Private Sub cmdSaldosConsol_Click()
    
    If LblPersCod <> "" Then
        frmConsolSaldos.ConsolidaDatos LblPersCod, lblNomPers, lblDocNat, lblDocJur
    End If
    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oDatos As DGeneral

    CentraForm Me
    Set oDatos = New DGeneral
    nLima = oDatos.LeeConstSistema(103)
    nSaldosAho = oDatos.LeeConstSistema(106)
    Set oDatos = Nothing
    
    If nSaldosAho = 1 Then
        cmdSaldosConsol.Visible = True
    End If
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub lstAhorros_DblClick()
    If Me.lstAhorros.SelectedItem.SubItems(4) <> "" Then
        Call frmCapMantenimiento.MuestraPosicionCliente(Me.lstAhorros.SelectedItem.SubItems(4))
    End If
End Sub

Private Sub LstCartaFianza_DblClick()
    If LstCartaFianza.SelectedItem.SubItems(2) <> "" Then
        Call frmCFHistorial.CargaCFHistorial(LstCartaFianza.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub LstCreditos_DblClick()
    If Me.LstCreditos.SelectedItem.SubItems(2) <> "" Then
        Call frmCredConsulta.ConsultaCliente(Me.LstCreditos.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub lstJudicial_DblClick()
    If lstJudicial.SelectedItem.SubItems(2) <> "" Then
        Call frmColRecRConsulta.MuestraPosicionCliente(lstJudicial.SelectedItem.SubItems(2))
    End If
End Sub

Private Sub lstPrendario_DblClick()

    If lstPrendario.SelectedItem.SubItems(3) <> "" Then
        Call frmPigConsulta.MuestraContratoPosicion(lstPrendario.SelectedItem.SubItems(3), LblPersCod)
        'Call frmColPContratosxCliente
    End If
End Sub

