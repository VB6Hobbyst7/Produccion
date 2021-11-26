VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdmCredRegControlCreditos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Control de Créditos"
   ClientHeight    =   11265
   ClientLeft      =   2670
   ClientTop       =   1875
   ClientWidth     =   12105
   Icon            =   "frmAdmCredRegControlCreditos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11265
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SST1 
      Height          =   11055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   19500
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Registrar Control de Créditos"
      TabPicture(0)   =   "frmAdmCredRegControlCreditos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkRevisaDesembControl"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Regularizar Observaciones"
      TabPicture(1)   =   "frmAdmCredRegControlCreditos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "chkRevisaDesembObs"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Mantenimiento"
      TabPicture(2)   =   "frmAdmCredRegControlCreditos.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ActXCodCta2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SSTab2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdBuscaMant"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame9"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CheckBox chkRevisaDesembObs 
         Caption         =   "Revisado para Desembolso"
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
         Left            =   -74880
         TabIndex        =   133
         Top             =   6840
         Width           =   2655
      End
      Begin VB.CheckBox chkRevisaDesembControl 
         Caption         =   "Revisado para Desembolso"
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
         Height          =   195
         Left            =   120
         TabIndex        =   132
         Top             =   10440
         Width           =   2655
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   8400
         TabIndex        =   128
         Top             =   10320
         Width           =   3375
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   2400
            TabIndex        =   131
            Top             =   160
            Width           =   900
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar"
            Height          =   375
            Left            =   1440
            TabIndex        =   130
            Top             =   160
            Width           =   900
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   120
            TabIndex        =   129
            Top             =   160
            Width           =   900
         End
      End
      Begin VB.Frame Frame9 
         Height          =   855
         Left            =   -66960
         TabIndex        =   66
         Top             =   6240
         Width           =   3615
         Begin VB.CommandButton cmdCancelarMant 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   1080
            TabIndex        =   69
            Top             =   240
            Width           =   900
         End
         Begin VB.CommandButton cmdGrabarMant 
            Caption         =   "&Grabar"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   900
         End
         Begin VB.CommandButton CmdSalirMant 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   2520
            TabIndex        =   67
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "&Lista Creditos"
         Height          =   1320
         Left            =   -67560
         TabIndex        =   58
         Top             =   600
         Width           =   3195
         Begin VB.ListBox LstObsMant 
            Height          =   840
            ItemData        =   "frmAdmCredRegControlCreditos.frx":035E
            Left            =   120
            List            =   "frmAdmCredRegControlCreditos.frx":0360
            TabIndex        =   59
            Top             =   240
            Width           =   2940
         End
      End
      Begin VB.CommandButton CmdBuscaMant 
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
         Left            =   -70800
         TabIndex        =   57
         Top             =   600
         Width           =   900
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   43
         Top             =   480
         Width           =   9855
         Begin VB.CheckBox chkPostDesembolsoBusca 
            Caption         =   "Post Desembolso"
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
            Left            =   240
            TabIndex        =   101
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton cmdBuscaObs 
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
            Left            =   4320
            TabIndex        =   47
            Top             =   360
            Width           =   900
         End
         Begin VB.Frame Frame7 
            Caption         =   "&Lista Creditos"
            Height          =   1320
            Left            =   6120
            TabIndex        =   45
            Top             =   120
            Width           =   3555
            Begin VB.ListBox LstObs 
               Height          =   840
               ItemData        =   "frmAdmCredRegControlCreditos.frx":0362
               Left            =   75
               List            =   "frmAdmCredRegControlCreditos.frx":0369
               TabIndex        =   46
               Top             =   225
               Width           =   3300
            End
         End
         Begin SICMACT.ActXCodCta ActXCodCta1 
            Height          =   540
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   3915
            _extentx        =   6906
            _extenty        =   953
            texto           =   "Credito :"
            enabledcmac     =   -1
            enabledcta      =   -1
            enabledprod     =   -1
            enabledage      =   -1
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   41
         Top             =   2040
         Width           =   9855
         Begin SICMACT.FlexEdit FlexEdit1 
            Height          =   4125
            Left            =   360
            TabIndex        =   42
            Top             =   240
            Width           =   8850
            _extentx        =   15610
            _extenty        =   7276
            cols0           =   4
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Observacion-cCtaCod-OK"
            encabezadosanchos=   "400-7000-0-500"
            font            =   "frmAdmCredRegControlCreditos.frx":0375
            font            =   "frmAdmCredRegControlCreditos.frx":03A1
            font            =   "frmAdmCredRegControlCreditos.frx":03CD
            font            =   "frmAdmCredRegControlCreditos.frx":03F9
            font            =   "frmAdmCredRegControlCreditos.frx":0425
            fontfixed       =   "frmAdmCredRegControlCreditos.frx":0451
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-3"
            listacontroles  =   "0-0-0-4"
            encabezadosalineacion=   "L-L-C-C"
            formatosedit    =   "0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   -64920
         TabIndex        =   37
         Top             =   2040
         Width           =   1215
         Begin VB.CommandButton cmdCancelReg 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   900
         End
         Begin VB.CommandButton cmdGrabarReg 
            Caption         =   "&Grabar"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   900
         End
         Begin VB.CommandButton cmdSalirReg 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Credito"
         Height          =   6330
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   11655
         Begin VB.CheckBox chkPostDesembolso 
            Caption         =   "Post Desembolso"
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
            Left            =   3960
            TabIndex        =   83
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtcompra 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   70
            Top             =   5880
            Visible         =   0   'False
            Width           =   7335
         End
         Begin VB.CheckBox ChkConstitucion 
            Alignment       =   1  'Right Justify
            Caption         =   "Constitucion Garantia :"
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
            Left            =   9120
            TabIndex        =   50
            Top             =   3050
            Width           =   2295
         End
         Begin VB.Frame FraListaCred 
            Caption         =   "&Lista Creditos"
            Height          =   840
            Left            =   8040
            TabIndex        =   7
            Top             =   120
            Width           =   3555
            Begin VB.ListBox LstCred 
               Height          =   450
               ItemData        =   "frmAdmCredRegControlCreditos.frx":047F
               Left            =   75
               List            =   "frmAdmCredRegControlCreditos.frx":0481
               TabIndex        =   8
               Top             =   225
               Width           =   3300
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
            Left            =   6240
            TabIndex        =   6
            Top             =   240
            Width           =   900
         End
         Begin VB.CheckBox chkCompraDeuda 
            Alignment       =   1  'Right Justify
            Caption         =   "Compra de Deuda :"
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
            Left            =   240
            TabIndex        =   5
            Top             =   5880
            Width           =   2295
         End
         Begin VB.CheckBox chkClientePref 
            Caption         =   "Cliente Preferencial"
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
            Left            =   3960
            TabIndex        =   4
            Top             =   120
            Width           =   2175
         End
         Begin SICMACT.ActXCodCta ActxCta 
            Height          =   420
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   3675
            _extentx        =   6482
            _extenty        =   741
            texto           =   "Credito :"
            enabledcmac     =   -1
            enabledcta      =   -1
            enabledprod     =   -1
            enabledage      =   -1
         End
         Begin MSMask.MaskEdBox txtFechaRevision 
            Height          =   300
            Left            =   6720
            TabIndex        =   10
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ListView ListaRelacion 
            Height          =   750
            Left            =   960
            TabIndex        =   71
            Top             =   3795
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   1323
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nombre de Cliente"
               Object.Width           =   7231
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Relación"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Cliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Valor. Rel."
               Object.Width           =   0
            EndProperty
         End
         Begin SICMACT.FlexEdit FEGarantCred 
            Height          =   1215
            Left            =   960
            TabIndex        =   76
            Top             =   4560
            Width           =   10485
            _extentx        =   18494
            _extenty        =   2143
            cols0           =   12
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "-Garantia-Gravament-Comercial-Realizacion-Disponible-Titular-Nro Docum-TipoDoc-cNumGarant-Legal-Poliza"
            encabezadosanchos=   "300-3800-1200-1200-1200-1200-3500-1200-0-1500-1800-1500"
            font            =   "frmAdmCredRegControlCreditos.frx":0483
            font            =   "frmAdmCredRegControlCreditos.frx":04AB
            font            =   "frmAdmCredRegControlCreditos.frx":04D3
            font            =   "frmAdmCredRegControlCreditos.frx":04FB
            font            =   "frmAdmCredRegControlCreditos.frx":0523
            fontfixed       =   "frmAdmCredRegControlCreditos.frx":054B
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-R-R-R-R-L-L-L-C-L-L"
            formatosedit    =   "0-0-2-2-2-2-0-0-0-0-0-0"
            lbbuscaduplicadotext=   -1
            colwidth0       =   300
            rowheight0      =   300
         End
         Begin MSComctlLib.ListView ListaRelacion1 
            Height          =   750
            Left            =   7080
            TabIndex        =   81
            Top             =   3795
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   1323
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Credito"
               Object.Width           =   3527
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Monto"
               Object.Width           =   2187
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Cuota"
               Object.Width           =   1236
            EndProperty
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "RelacionRef :"
            Height          =   195
            Left            =   7200
            TabIndex        =   82
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Garantias :"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   4800
            Width           =   765
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Relacion :"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   3960
            Width           =   720
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   3465
            Width           =   525
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   77
            Top             =   3420
            Width           =   6015
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   15
            Left            =   10080
            TabIndex        =   75
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "F. Pago :"
            Height          =   195
            Left            =   9000
            TabIndex        =   74
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   14
            Left            =   960
            TabIndex        =   73
            Top             =   3015
            Width           =   7815
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Convenio :"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   3100
            Width           =   765
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   13
            Left            =   7320
            TabIndex        =   49
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Frecuencia :"
            Height          =   195
            Left            =   6240
            TabIndex        =   48
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   4
            Left            =   960
            TabIndex        =   36
            Top             =   1590
            Width           =   5175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Datos del Crédito"
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
            TabIndex        =   34
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1290
            Width           =   585
         End
         Begin VB.Label lblTrib 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Juridico :"
            Height          =   195
            Left            =   9000
            TabIndex        =   32
            Top             =   1320
            Width           =   1020
         End
         Begin VB.Label lblNat 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Natural :"
            Height          =   195
            Left            =   6240
            TabIndex        =   31
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   3
            Left            =   10080
            TabIndex        =   30
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   2
            Left            =   7320
            TabIndex        =   29
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   5
            Left            =   7320
            TabIndex        =   27
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Analista :"
            Height          =   195
            Left            =   6240
            TabIndex        =   26
            Top             =   1680
            Width           =   645
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   6
            Left            =   3840
            TabIndex        =   25
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad :"
            Height          =   195
            Left            =   2880
            TabIndex        =   24
            Top             =   1320
            Width           =   825
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            BorderWidth     =   2
            X1              =   120
            X2              =   11520
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Actividad :"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2205
            Width           =   750
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Index           =   7
            Left            =   960
            TabIndex        =   22
            Top             =   1980
            Width           =   5175
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   6240
            TabIndex        =   21
            Top             =   2760
            Width           =   540
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   8
            Left            =   7320
            TabIndex        =   20
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   9000
            TabIndex        =   19
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   9
            Left            =   10080
            TabIndex        =   18
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Exposición :"
            Height          =   195
            Left            =   6240
            TabIndex        =   17
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   10
            Left            =   7320
            TabIndex        =   16
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   2760
            Width           =   585
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   11
            Left            =   960
            TabIndex        =   14
            Top             =   2640
            Width           =   5175
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas :"
            Height          =   195
            Left            =   9000
            TabIndex        =   13
            Top             =   2040
            Width           =   585
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Index           =   12
            Left            =   10080
            TabIndex        =   12
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Revisión :"
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
            Left            =   4920
            TabIndex        =   11
            Top             =   720
            Width           =   1725
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   6720
         Width           =   11655
         Begin TabDlg.SSTab w 
            Height          =   3285
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5794
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Observaciones"
            TabPicture(0)   =   "frmAdmCredRegControlCreditos.frx":0571
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "frameObservaciones"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Exoneraciones"
            TabPicture(1)   =   "frmAdmCredRegControlCreditos.frx":058D
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "frameExoneraciones"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Autorizaciones"
            TabPicture(2)   =   "frmAdmCredRegControlCreditos.frx":05A9
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FrameAuto"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Post Desembolso"
            TabPicture(3)   =   "frmAdmCredRegControlCreditos.frx":05C5
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "framePostDesembolso"
            Tab(3).ControlCount=   1
            Begin VB.Frame FrameAuto 
               Height          =   2655
               Left            =   -74880
               TabIndex        =   110
               Top             =   480
               Width           =   11175
               Begin VB.CommandButton cmdQuitarAutorizacion 
                  Caption         =   "Quitar"
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   127
                  Top             =   2193
                  Width           =   975
               End
               Begin VB.CommandButton cmdAgregarAutorizacion 
                  Caption         =   "Agregar"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   117
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.ComboBox cboNivAutorizacion 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   114
                  Top             =   645
                  Width           =   3735
               End
               Begin VB.ComboBox cboAutorizacion 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   111
                  Top             =   240
                  Width           =   3735
               End
               Begin SICMACT.FlexEdit FEAutorizaciones 
                  Height          =   1935
                  Left            =   5160
                  TabIndex        =   113
                  Top             =   240
                  Width           =   5895
                  _extentx        =   10398
                  _extenty        =   3413
                  cols0           =   6
                  highlight       =   1
                  allowuserresizing=   3
                  rowsizingmode   =   1
                  encabezadosnombres=   "#-Autorización-Nivel de Autorización-Apoderado 1-Apoderado 2-Apoderado 3"
                  encabezadosanchos=   "300-2000-2000-2000-2000-2000"
                  font            =   "frmAdmCredRegControlCreditos.frx":05E1
                  font            =   "frmAdmCredRegControlCreditos.frx":060D
                  font            =   "frmAdmCredRegControlCreditos.frx":0639
                  font            =   "frmAdmCredRegControlCreditos.frx":0665
                  font            =   "frmAdmCredRegControlCreditos.frx":0691
                  fontfixed       =   "frmAdmCredRegControlCreditos.frx":06BD
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  lbultimainstancia=   -1
                  columnasaeditar =   "X-X-X-3-X-X"
                  listacontroles  =   "0-0-0-3-0-0"
                  encabezadosalineacion=   "C-L-L-L-L-L"
                  formatosedit    =   "0-0-0-0-0-0"
                  textarray0      =   "#"
                  lbeditarflex    =   -1
                  colwidth0       =   300
                  rowheight0      =   300
                  forecolorfixed  =   -2147483630
               End
               Begin SICMACT.FlexEdit FEApoderadoAuto 
                  Height          =   1095
                  Left            =   120
                  TabIndex        =   116
                  Top             =   1080
                  Width           =   4935
                  _extentx        =   8705
                  _extenty        =   1931
                  cols0           =   3
                  highlight       =   1
                  allowuserresizing=   3
                  rowsizingmode   =   1
                  encabezadosnombres=   "Nº-Nº Apoderado-Descripción"
                  encabezadosanchos=   "0-1500-3350"
                  font            =   "frmAdmCredRegControlCreditos.frx":06EB
                  font            =   "frmAdmCredRegControlCreditos.frx":0717
                  font            =   "frmAdmCredRegControlCreditos.frx":0743
                  font            =   "frmAdmCredRegControlCreditos.frx":076F
                  font            =   "frmAdmCredRegControlCreditos.frx":079B
                  fontfixed       =   "frmAdmCredRegControlCreditos.frx":07C7
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  lbultimainstancia=   -1
                  columnasaeditar =   "X-X-2"
                  listacontroles  =   "0-0-3"
                  encabezadosalineacion=   "C-C-L"
                  formatosedit    =   "0-0-0"
                  textarray0      =   "Nº"
                  lbeditarflex    =   -1
                  rowheight0      =   300
                  forecolorfixed  =   -2147483630
               End
               Begin VB.Label Label24 
                  Caption         =   "Nivel de Autorización:"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   115
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label25 
                  Caption         =   "Autorización:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   112
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.Frame framePostDesembolso 
               Height          =   2775
               Left            =   -74760
               TabIndex        =   91
               Top             =   360
               Width           =   10815
               Begin VB.TextBox txtdesembolso 
                  Height          =   285
                  Left            =   2040
                  MaxLength       =   100
                  TabIndex        =   96
                  Top             =   720
                  Width           =   6675
               End
               Begin VB.CommandButton CmdAgregaPost 
                  Caption         =   "Agrega"
                  Height          =   375
                  Left            =   9360
                  TabIndex        =   95
                  Top             =   960
                  Width           =   1020
               End
               Begin VB.CommandButton cmdborrarPost 
                  Caption         =   "Borrar"
                  Height          =   375
                  Left            =   9360
                  TabIndex        =   94
                  Top             =   1440
                  Width           =   1020
               End
               Begin VB.ComboBox cboagencia 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  ItemData        =   "frmAdmCredRegControlCreditos.frx":07F5
                  Left            =   2040
                  List            =   "frmAdmCredRegControlCreditos.frx":07F7
                  Style           =   2  'Dropdown List
                  TabIndex        =   93
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.ComboBox cboRF 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   5880
                  Style           =   2  'Dropdown List
                  TabIndex        =   92
                  Top             =   240
                  Width           =   2775
               End
               Begin SICMACT.FlexEdit flexPost 
                  Height          =   1365
                  Left            =   720
                  TabIndex        =   97
                  Top             =   1200
                  Width           =   8010
                  _extentx        =   14129
                  _extenty        =   2408
                  cols0           =   4
                  highlight       =   1
                  allowuserresizing=   3
                  rowsizingmode   =   1
                  encabezadosnombres=   "#-Observacion-User-Agencia"
                  encabezadosanchos=   "400-5200-1000-800"
                  font            =   "frmAdmCredRegControlCreditos.frx":07F9
                  font            =   "frmAdmCredRegControlCreditos.frx":0825
                  font            =   "frmAdmCredRegControlCreditos.frx":0851
                  font            =   "frmAdmCredRegControlCreditos.frx":087D
                  font            =   "frmAdmCredRegControlCreditos.frx":08A9
                  fontfixed       =   "frmAdmCredRegControlCreditos.frx":08D5
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  lbultimainstancia=   -1
                  columnasaeditar =   "X-X-X-X"
                  listacontroles  =   "0-0-0-0"
                  encabezadosalineacion=   "L-L-C-C"
                  formatosedit    =   "0-0-0-0"
                  textarray0      =   "#"
                  lbeditarflex    =   -1
                  colwidth0       =   405
                  rowheight0      =   300
                  forecolorfixed  =   -2147483630
               End
               Begin VB.Label Label32 
                  AutoSize        =   -1  'True
                  Caption         =   "Observaciones :"
                  Height          =   195
                  Left            =   720
                  TabIndex        =   100
                  ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
                  Top             =   720
                  Width           =   1155
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Agencia :"
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
                  Left            =   720
                  TabIndex        =   99
                  Top             =   240
                  Width           =   750
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "User :"
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
                  Left            =   5280
                  TabIndex        =   98
                  Top             =   240
                  Width           =   480
               End
            End
            Begin VB.Frame frameExoneraciones 
               Height          =   2775
               Left            =   -74880
               TabIndex        =   90
               Top             =   360
               Width           =   11175
               Begin VB.CommandButton cmdQuitaExo 
                  Caption         =   "Quitar"
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   109
                  Top             =   2320
                  Width           =   975
               End
               Begin VB.CommandButton cmdAgregarExo 
                  Caption         =   "Agregar"
                  Height          =   375
                  Left            =   360
                  TabIndex        =   108
                  Top             =   2320
                  Width           =   975
               End
               Begin SICMACT.FlexEdit FEExoneraciones 
                  Height          =   1935
                  Left            =   5160
                  TabIndex        =   107
                  Top             =   360
                  Width           =   5895
                  _extentx        =   10398
                  _extenty        =   3413
                  cols0           =   6
                  highlight       =   1
                  allowuserresizing=   3
                  rowsizingmode   =   1
                  encabezadosnombres=   "#-Exoneración-Nivel de Exoneración-Apoderado 1-Apoderado 2-Apoderado 3"
                  encabezadosanchos=   "300-2000-2000-2000-2000-2000"
                  font            =   "frmAdmCredRegControlCreditos.frx":0903
                  font            =   "frmAdmCredRegControlCreditos.frx":092F
                  font            =   "frmAdmCredRegControlCreditos.frx":095B
                  font            =   "frmAdmCredRegControlCreditos.frx":0987
                  font            =   "frmAdmCredRegControlCreditos.frx":09B3
                  fontfixed       =   "frmAdmCredRegControlCreditos.frx":09DF
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  lbultimainstancia=   -1
                  columnasaeditar =   "X-X-X-X-X-X"
                  listacontroles  =   "0-0-0-0-0-0"
                  encabezadosalineacion=   "C-L-L-L-L-L"
                  formatosedit    =   "0-0-0-0-0-0"
                  textarray0      =   "#"
                  lbeditarflex    =   -1
                  colwidth0       =   300
                  rowheight0      =   300
                  forecolorfixed  =   -2147483630
               End
               Begin SICMACT.FlexEdit FEApoderados 
                  Height          =   1095
                  Left            =   120
                  TabIndex        =   106
                  Top             =   1200
                  Width           =   4935
                  _extentx        =   8705
                  _extenty        =   1931
                  cols0           =   3
                  highlight       =   1
                  allowuserresizing=   3
                  rowsizingmode   =   1
                  encabezadosnombres=   "Nº-Nº Apoderado-Descripción"
                  encabezadosanchos=   "0-1500-3350"
                  font            =   "frmAdmCredRegControlCreditos.frx":0A0D
                  font            =   "frmAdmCredRegControlCreditos.frx":0A39
                  font            =   "frmAdmCredRegControlCreditos.frx":0A65
                  font            =   "frmAdmCredRegControlCreditos.frx":0A91
                  font            =   "frmAdmCredRegControlCreditos.frx":0ABD
                  fontfixed       =   "frmAdmCredRegControlCreditos.frx":0AE9
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  lbultimainstancia=   -1
                  columnasaeditar =   "X-X-2"
                  listacontroles  =   "0-0-3"
                  encabezadosalineacion=   "C-C-L"
                  formatosedit    =   "0-0-0"
                  textarray0      =   "Nº"
                  lbeditarflex    =   -1
                  rowheight0      =   300
                  forecolorfixed  =   -2147483630
               End
               Begin VB.ComboBox cboNivelExonera 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   105
                  Top             =   840
                  Width           =   3735
               End
               Begin VB.ComboBox cboExoneraciones 
                  Height          =   315
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   103
                  Top             =   360
                  Width           =   3735
               End
               Begin VB.Label Label16 
                  Caption         =   "Nivel de Exoneración:"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   104
                  Top             =   800
                  Width           =   975
               End
               Begin VB.Label Label11 
                  Caption         =   "Exoneración:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  Top             =   405
                  Width           =   1215
               End
            End
            Begin VB.Frame frameObservaciones 
               Height          =   2775
               Left            =   120
               TabIndex        =   84
               Top             =   360
               Width           =   11175
               Begin VB.CommandButton cmdBorraObs 
                  Caption         =   "Borrar"
                  Height          =   375
                  Left            =   9960
                  TabIndex        =   89
                  Top             =   1440
                  Width           =   900
               End
               Begin VB.CommandButton cmdAgregaObs 
                  Caption         =   "Agrega"
                  Height          =   375
                  Left            =   9960
                  TabIndex        =   88
                  Top             =   480
                  Width           =   900
               End
               Begin VB.TextBox txtObservaciones 
                  Height          =   285
                  Left            =   1560
                  MaxLength       =   100
                  TabIndex        =   85
                  Top             =   240
                  Width           =   6670
               End
               Begin SICMACT.FlexEdit FlexObs 
                  Height          =   1845
                  Left            =   240
                  TabIndex        =   87
                  Top             =   720
                  Width           =   8850
                  _extentx        =   15610
                  _extenty        =   3254
                  cols0           =   3
                  highlight       =   1
                  allowuserresizing=   3
                  rowsizingmode   =   1
                  encabezadosnombres=   "#-Observacion-cCtaCod"
                  encabezadosanchos=   "400-7000-0"
                  font            =   "frmAdmCredRegControlCreditos.frx":0B17
                  font            =   "frmAdmCredRegControlCreditos.frx":0B43
                  font            =   "frmAdmCredRegControlCreditos.frx":0B6F
                  font            =   "frmAdmCredRegControlCreditos.frx":0B9B
                  font            =   "frmAdmCredRegControlCreditos.frx":0BC7
                  fontfixed       =   "frmAdmCredRegControlCreditos.frx":0BF3
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  backcolorcontrol=   -2147483643
                  lbultimainstancia=   -1
                  columnasaeditar =   "X-X-X"
                  listacontroles  =   "0-0-0"
                  encabezadosalineacion=   "L-L-C"
                  formatosedit    =   "0-0-0"
                  textarray0      =   "#"
                  lbeditarflex    =   -1
                  colwidth0       =   405
                  rowheight0      =   300
                  forecolorfixed  =   -2147483630
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Observaciones :"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   86
                  ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
                  Top             =   240
                  Width           =   1155
               End
            End
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4125
         Left            =   -74760
         TabIndex        =   51
         Top             =   2040
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7276
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Observaciones"
         TabPicture(0)   =   "frmAdmCredRegControlCreditos.frx":0C21
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label23"
         Tab(0).Control(1)=   "FlexObsMant"
         Tab(0).Control(2)=   "txtObservacionesMant"
         Tab(0).Control(3)=   "cmdAgregaObsMant"
         Tab(0).Control(4)=   "cmdBorraObsMant"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Exoneraciones"
         TabPicture(1)   =   "frmAdmCredRegControlCreditos.frx":0C3D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame10"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Autorizaciones"
         TabPicture(2)   =   "frmAdmCredRegControlCreditos.frx":0C59
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame11"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame11 
            Height          =   2655
            Left            =   120
            TabIndex        =   134
            Top             =   480
            Width           =   11175
            Begin VB.ComboBox cboAutorizacionMant 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   240
               Width           =   3735
            End
            Begin VB.ComboBox cboNivelAutMant 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   645
               Width           =   3735
            End
            Begin VB.CommandButton cmdAgregaAutMant 
               Caption         =   "Agregar"
               Height          =   375
               Left            =   120
               TabIndex        =   136
               Top             =   2160
               Width           =   975
            End
            Begin VB.CommandButton cmdQuitarAutMant 
               Caption         =   "Quitar"
               Height          =   375
               Left            =   1200
               TabIndex        =   135
               Top             =   2193
               Width           =   975
            End
            Begin SICMACT.FlexEdit FEAutorizacionesMant 
               Height          =   1935
               Left            =   5160
               TabIndex        =   139
               Top             =   240
               Width           =   5895
               _extentx        =   10398
               _extenty        =   3413
               cols0           =   6
               highlight       =   1
               allowuserresizing=   3
               rowsizingmode   =   1
               encabezadosnombres=   "#-Autorización-Nivel de Autorización-Apoderado 1-Apoderado 2-Apoderado 3"
               encabezadosanchos=   "300-2000-2000-2000-2000-2000"
               font            =   "frmAdmCredRegControlCreditos.frx":0C75
               font            =   "frmAdmCredRegControlCreditos.frx":0CA1
               font            =   "frmAdmCredRegControlCreditos.frx":0CCD
               font            =   "frmAdmCredRegControlCreditos.frx":0CF9
               font            =   "frmAdmCredRegControlCreditos.frx":0D25
               fontfixed       =   "frmAdmCredRegControlCreditos.frx":0D51
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               lbultimainstancia=   -1
               columnasaeditar =   "X-X-X-3-X-X"
               listacontroles  =   "0-0-0-3-0-0"
               encabezadosalineacion=   "C-L-L-L-L-L"
               formatosedit    =   "0-0-0-0-0-0"
               textarray0      =   "#"
               lbeditarflex    =   -1
               colwidth0       =   300
               rowheight0      =   300
               forecolorfixed  =   -2147483630
            End
            Begin SICMACT.FlexEdit FEApoderadoAutMant 
               Height          =   1095
               Left            =   120
               TabIndex        =   140
               Top             =   1080
               Width           =   4935
               _extentx        =   8705
               _extenty        =   1931
               cols0           =   3
               highlight       =   1
               allowuserresizing=   3
               rowsizingmode   =   1
               encabezadosnombres=   "Nº-Nº Apoderado-Descripción"
               encabezadosanchos=   "0-1500-3350"
               font            =   "frmAdmCredRegControlCreditos.frx":0D7F
               font            =   "frmAdmCredRegControlCreditos.frx":0DAB
               font            =   "frmAdmCredRegControlCreditos.frx":0DD7
               font            =   "frmAdmCredRegControlCreditos.frx":0E03
               font            =   "frmAdmCredRegControlCreditos.frx":0E2F
               fontfixed       =   "frmAdmCredRegControlCreditos.frx":0E5B
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               lbultimainstancia=   -1
               columnasaeditar =   "X-X-2"
               listacontroles  =   "0-0-3"
               encabezadosalineacion=   "C-C-L"
               formatosedit    =   "0-0-0"
               textarray0      =   "Nº"
               lbeditarflex    =   -1
               rowheight0      =   300
               forecolorfixed  =   -2147483630
            End
            Begin VB.Label Label22 
               Caption         =   "Autorización:"
               Height          =   255
               Left            =   120
               TabIndex        =   142
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label21 
               Caption         =   "Nivel de Autorización:"
               Height          =   495
               Left            =   120
               TabIndex        =   141
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.Frame Frame10 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   118
            Top             =   360
            Width           =   11175
            Begin VB.ComboBox cboExoneraMat 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   360
               Width           =   3735
            End
            Begin VB.ComboBox cboNivExoneraMat 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   123
               Top             =   840
               Width           =   3735
            End
            Begin VB.CommandButton cmdAgregarExoneraMat 
               Caption         =   "Agregar"
               Height          =   375
               Left            =   120
               TabIndex        =   120
               Top             =   2320
               Width           =   975
            End
            Begin VB.CommandButton cmdQuitarExoMant 
               Caption         =   "Quitar"
               Height          =   375
               Left            =   1200
               TabIndex        =   119
               Top             =   2320
               Width           =   975
            End
            Begin SICMACT.FlexEdit FEExoneraMat 
               Height          =   1935
               Left            =   5160
               TabIndex        =   121
               Top             =   360
               Width           =   5895
               _extentx        =   10398
               _extenty        =   3413
               cols0           =   6
               highlight       =   1
               allowuserresizing=   3
               rowsizingmode   =   1
               encabezadosnombres=   "#-Exoneración-Nivel de Exoneración-Apoderado 1-Apoderado 2-Apoderado 3"
               encabezadosanchos=   "300-2000-2000-2000-2000-2000"
               font            =   "frmAdmCredRegControlCreditos.frx":0E89
               font            =   "frmAdmCredRegControlCreditos.frx":0EB5
               font            =   "frmAdmCredRegControlCreditos.frx":0EE1
               font            =   "frmAdmCredRegControlCreditos.frx":0F0D
               font            =   "frmAdmCredRegControlCreditos.frx":0F39
               fontfixed       =   "frmAdmCredRegControlCreditos.frx":0F65
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               lbultimainstancia=   -1
               columnasaeditar =   "X-X-X-X-X-X"
               listacontroles  =   "0-0-0-0-0-0"
               encabezadosalineacion=   "C-L-L-L-L-L"
               formatosedit    =   "0-0-0-0-0-0"
               textarray0      =   "#"
               lbeditarflex    =   -1
               colwidth0       =   300
               rowheight0      =   300
               forecolorfixed  =   -2147483630
            End
            Begin SICMACT.FlexEdit FEApoderadoExoneraMat 
               Height          =   1095
               Left            =   120
               TabIndex        =   122
               Top             =   1200
               Width           =   4935
               _extentx        =   8705
               _extenty        =   1931
               cols0           =   3
               highlight       =   1
               allowuserresizing=   3
               rowsizingmode   =   1
               encabezadosnombres=   "Nº-Nº Apoderado-Descripción"
               encabezadosanchos=   "0-1500-3350"
               font            =   "frmAdmCredRegControlCreditos.frx":0F93
               font            =   "frmAdmCredRegControlCreditos.frx":0FBF
               font            =   "frmAdmCredRegControlCreditos.frx":0FEB
               font            =   "frmAdmCredRegControlCreditos.frx":1017
               font            =   "frmAdmCredRegControlCreditos.frx":1043
               fontfixed       =   "frmAdmCredRegControlCreditos.frx":106F
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               lbultimainstancia=   -1
               columnasaeditar =   "X-X-2"
               listacontroles  =   "0-0-3"
               encabezadosalineacion=   "C-C-L"
               formatosedit    =   "0-0-0"
               textarray0      =   "Nº"
               lbeditarflex    =   -1
               rowheight0      =   300
               forecolorfixed  =   -2147483630
            End
            Begin VB.Label Label20 
               Caption         =   "Exoneración:"
               Height          =   255
               Left            =   120
               TabIndex        =   126
               Top             =   405
               Width           =   1215
            End
            Begin VB.Label Label17 
               Caption         =   "Nivel de Exoneración:"
               Height          =   495
               Left            =   120
               TabIndex        =   125
               Top             =   800
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdBorraObsMant 
            Caption         =   "Borrar"
            Height          =   375
            Left            =   -66240
            TabIndex        =   63
            Top             =   1800
            Width           =   1020
         End
         Begin VB.CommandButton cmdAgregaObsMant 
            Caption         =   "Agrega"
            Height          =   375
            Left            =   -66240
            TabIndex        =   62
            Top             =   1320
            Width           =   1020
         End
         Begin VB.TextBox txtObservacionesMant 
            Height          =   285
            Left            =   -73560
            MaxLength       =   100
            TabIndex        =   61
            Top             =   720
            Width           =   6670
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Borrar"
            Height          =   375
            Left            =   -66840
            TabIndex        =   54
            Top             =   1560
            Width           =   660
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Agrega"
            Height          =   375
            Left            =   -66840
            TabIndex        =   53
            Top             =   600
            Width           =   660
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   -73680
            MaxLength       =   100
            TabIndex        =   52
            Top             =   600
            Width           =   6670
         End
         Begin SICMACT.FlexEdit FlexEdit2 
            Height          =   2085
            Left            =   -74880
            TabIndex        =   55
            Top             =   960
            Width           =   7890
            _extentx        =   13917
            _extenty        =   3678
            cols0           =   3
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Observacion-cCtaCod"
            encabezadosanchos=   "400-7000-0"
            font            =   "frmAdmCredRegControlCreditos.frx":109D
            font            =   "frmAdmCredRegControlCreditos.frx":10C9
            font            =   "frmAdmCredRegControlCreditos.frx":10F5
            font            =   "frmAdmCredRegControlCreditos.frx":1121
            font            =   "frmAdmCredRegControlCreditos.frx":114D
            fontfixed       =   "frmAdmCredRegControlCreditos.frx":1179
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X"
            listacontroles  =   "0-0-0"
            encabezadosalineacion=   "L-L-C"
            formatosedit    =   "0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit FlexObsMant 
            Height          =   2085
            Left            =   -74880
            TabIndex        =   64
            Top             =   1440
            Width           =   8130
            _extentx        =   14340
            _extenty        =   3678
            cols0           =   3
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Observacion-cCtaCod"
            encabezadosanchos=   "400-7000-0"
            font            =   "frmAdmCredRegControlCreditos.frx":11A7
            font            =   "frmAdmCredRegControlCreditos.frx":11D3
            font            =   "frmAdmCredRegControlCreditos.frx":11FF
            font            =   "frmAdmCredRegControlCreditos.frx":122B
            font            =   "frmAdmCredRegControlCreditos.frx":1257
            fontfixed       =   "frmAdmCredRegControlCreditos.frx":1283
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X"
            listacontroles  =   "0-0-0"
            encabezadosalineacion=   "L-L-C"
            formatosedit    =   "0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones :"
            Height          =   195
            Left            =   -74760
            TabIndex        =   65
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones :"
            Height          =   195
            Left            =   -74880
            TabIndex        =   56
            ToolTipText     =   "(días de atraso promedio últimas 6 cuotas)"
            Top             =   600
            Width           =   1155
         End
      End
      Begin SICMACT.ActXCodCta ActXCodCta2 
         Height          =   420
         Left            =   -74640
         TabIndex        =   60
         Top             =   600
         Width           =   3675
         _extentx        =   6482
         _extenty        =   741
         texto           =   "Credito :"
         enabledcmac     =   -1
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
      End
   End
End
Attribute VB_Name = "frmAdmCredRegControlCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim nMiVivienda As Integer
'Dim nPrestamo As Currency
'Dim nCalendDinamico As Integer
'Dim bCuotaCom As Integer
'Dim lcDNI As String, lcRUC As String
'Dim lnEstCred As Long
'Dim lcExonera As String
'Dim lcAuto As String
'Dim lcExoneraMant As String
'Dim i As Integer
'Dim j As Integer
'Dim bEncontrado As Boolean
'Dim nCuentaExo As Integer
'Dim nCuentaAuto As Integer
'Dim nCuentaExoMant As Integer
'Dim nCuentaObs As Integer
'Dim nCuentaObsMant As Integer
'Dim nCuentaPost As Integer
'Dim MatExo() As String
'Dim MatObs() As String
'Dim lnBuscaUsu As Integer
'Dim fbPostDesembolso As Boolean 'WIOR 20120408
'Dim nCargoNivApr As String 'RECO20140328 ERS174-2013 ANEXO 01
'Dim nCanFirmas As Integer 'RECO20140328 ERS174-2013 ANEXO 01
'Dim nCanFirmasAutMant As Integer 'RECO20141009
'Dim nCanFirmasExo As Integer 'RECO20140328 ERS174-2013 ANEXO 01
'Dim nCanFirmasExoMant As Integer 'RECO20140328 ERS174-2013 ANEXO 01
'Dim sPersCargoExo As String 'RECO20140328 ERS174-2013 ANEXO 01
'Dim sPersCargoAuto As String 'RECO20140328 ERS174-2013 ANEXO 01
'Dim sPersCargoExoMant As String 'RECO20140328 ERS174-2013 ANEXO 01
'
'
'Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'
'Dim oCred As COMDCredito.DCOMCreditos
'Dim rsCred As ADODB.Recordset
'Dim rsComun As ADODB.Recordset
'Dim prsRelac As ADODB.Recordset
'
'Dim bCargado As Boolean
'Dim sEstado As String
'Dim oCredito As COMDCredito.DCOMCredito
'Dim bRegistrado As Boolean 'WIOR 20120405
'Dim bPostDesem As Boolean 'WIOR 20120620
'
'
'    'WIOR 20120619-INICIO
'    Call LimpiaFlex(FlexObs)
'    'Call LimpiaFlex(FlexExo)
'    'Call LimpiaFlex(FlexAuto)
'    Call LimpiaFlex(flexPost)
'    'FIN
'
'    'Verifica que el num de cred no haya sido registrado
'    Set oCred = New COMDCredito.DCOMCreditos
'    bRegistrado = oCred.BuscaRegistroControlAdmCred(ActxCta.NroCuenta) 'WIOR 20120405
'    Set oCred = New COMDCredito.DCOMCreditos
'    bCargado = oCred.CargaDatosControlCreditosAdmCred(psCtaCod, rsCred, rsComun, sEstado, lnEstCred)
'    bPostDesem = oCred.ObtieneCredCFPostDesembolso(psCtaCod) 'WIOR 20120620
'
'    If bRegistrado And lnEstCred = 2002 Then 'WIOR 20120405
'        MsgBox "El control de este crédito ya fue registrado.", vbOKOnly + vbExclamation, "Aviso"
'        Exit Function
'    ElseIf bRegistrado And Mid(psCtaCod, 6, 3) = "514" Then 'WIOR 20120616
'        MsgBox "El control de Carta Fianza ya fue registrado.", vbOKOnly + vbExclamation, "Aviso"
'        Exit Function
'    End If
'
'    'WIOR 20120405
'    If bRegistrado = False And lnEstCred = 2020 And Me.chkPostDesembolso.value = 1 Then
'        MsgBox "Este crédito no fue registrado previo desembolso, si desea registrar el control de este crédito, marque el Check como Preferencial.", vbOKOnly + vbInformation, "Atención"
'        If MsgBox("Desea registrar este credito vigente, NO Preferencial?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
'            w.TabVisible(0) = True
'            w.TabVisible(1) = True
'            w.TabVisible(2) = True
'            fbPostDesembolso = False
'        Else
'            Exit Function
'        End If
'    End If
'    'WIOR - FIN
'
'
'    Set oCred = Nothing
'    If bCargado Then
'
'        If lnEstCred = 2020 And Me.chkClientePref.value = 0 And Me.chkPostDesembolso.value = 0 Then
'            MsgBox "Este crédito está en estado Vigente, si desea registrar el control de este crédito, marque el Check como Cliente Preferencial o Post Desembolso.", vbOKOnly + vbInformation, "Atención"
'            bCargado = False
'            Exit Function
'        End If
'
'        'WIOR 20120405
'        If lnEstCred = 2020 And bRegistrado And Me.chkClientePref.value = 1 Then
'            MsgBox "Este crédito no es preferencial, si desea registrar el control de este crédito, marque el Check como Post Desembolso.", vbOKOnly + vbInformation, "Atención"
'            bCargado = False
'            Exit Function
'        End If
'        If lnEstCred = 2002 And (Me.chkClientePref.value = 1 Or Me.chkPostDesembolso.value = 1) Then
'            MsgBox "Este crédito aún no fue desembolsado, desmarque el Check como Cliente Preferencial o Post Desembolso.", vbOKOnly + vbInformation, "Atención"
'            bCargado = False
'            Exit Function
'        End If
'        'WIOR - FIN
'
'        'WIOR 20120620 *********************************************
'         If bRegistrado And bPostDesem Then
'            MsgBox "El control Post Desembolso de este crédito ya fue registrado.", vbOKOnly + vbExclamation, "Aviso"
'            Exit Function
'         End If
'        'WIOR FIN **************************************************
'
'        lblcodigo(4).Caption = sEstado
'        nPrestamo = IIf(IsNull(rsComun!nMontoCol), rsComun!nMontoSol, rsComun!nMontoCol)
'        If nPrestamo = 0 Then
'            nPrestamo = rsComun!nMontoSol
'        End If
'        nMiVivienda = IIf(IsNull(rsComun!bMiVivienda), 0, rsComun!bMiVivienda)
'        bCuotaCom = IIf(IsNull(rsComun!bCuotaCom), 0, rsComun!bCuotaCom)
'        nCalendDinamico = IIf(IsNull(rsComun!nCalendDinamico), 0, rsComun!nCalendDinamico)
'        lblcodigo(0) = rsComun!cPersCod
''        lblcodigo(1) = PstaNombre(rsComun!cTitular)
'        lblcodigo(2) = IIf(IsNull(rsComun!Dni), "", rsComun!Dni)
'        lblcodigo(3) = IIf(IsNull(rsComun!Ruc), "", rsComun!Ruc)
'
'        lcDNI = Trim(IIf(IsNull(rsComun!Dni), "", rsComun!Dni))
'        lcRUC = Trim(IIf(IsNull(rsComun!Ruc), "", rsComun!Ruc))
'
'        lblcodigo(5) = IIf(IsNull(rsComun!cCodAna), "", rsComun!cCodAna)
'        lblcodigo(6) = IIf(IsNull(rsComun!cCondicion), "", rsComun!cCondicion)
'        lblcodigo(7) = IIf(IsNull(rsComun!CIIU), "", rsComun!CIIU)
'        lblcodigo(8) = IIf(IsNull(rsComun!Ptmo_prop), "", rsComun!Ptmo_prop)
'        lblcodigo(9) = IIf(IsNull(rsComun!cmoneda), "", rsComun!cmoneda)
'        lblcodigo(10) = IIf(IsNull(rsComun!ExposiCred), "", rsComun!ExposiCred)
'        lblcodigo(11) = IIf(IsNull(rsComun!cDestinoDescripcion), "", rsComun!cDestinoDescripcion)
'        lblcodigo(12) = IIf(IsNull(rsComun!nCuotas), "", rsComun!nCuotas)
'        lblcodigo(13) = IIf(IsNull(rsComun!cPlazoCE), "", rsComun!cPlazoCE)
'        lblcodigo(14) = IIf(IsNull(rsComun!Convenio), "", rsComun!Convenio)
'        lblcodigo(15) = IIf(IsNull(rsComun!dFechaPago), "", rsComun!dFechaPago)
'
'        'MADM 20120327 - Ref
'        lblcodigo(1) = IIf(IsNull(rsComun!MotivoRefinan), "", rsComun!MotivoRefinan)
'
'        If Trim(lblcodigo(1).Caption) <> "" Then
'            Call CargaPersonasRelacCredRefinan(psCtaCod)
'        End If
'        'END MADM
'
'        Me.txtFechaRevision.Text = CDate(gdFecSis)
'        Me.txtFechaRevision.SetFocus
'
'        Call CargarFlexGarantia(psCtaCod)
'        Call CargaPersonasRelacCred(psCtaCod, prsRelac)
'
'        'WIOR 20120517************************
'        Me.chkClientePref.Enabled = False
'        Me.chkPostDesembolso.Enabled = False
'        'WIOR FIN*****************************
'
'        'RECO20140401 ERS174-2013 ANEXO 01*******
'        frameExoneraciones.Enabled = True
'        FrameAuto.Enabled = True
'        'RECO FIN********************************
'
'        'JUEZ 20140725 ***********************
'        If chkPostDesembolso.value = 1 Then
'            chkRevisaDesembControl.value = 1
'            chkRevisaDesembControl.Enabled = False
'        Else
'            chkRevisaDesembControl.Enabled = True
'            chkRevisaDesembControl.value = 0
'        End If
'        'END JUEZ ****************************
'    Else
'        MsgBox "No se encontro datos de este crédito, por favo revise el correcto número del credito o el estado debe ser ''APROBADO''.", vbOKOnly + vbExclamation, "Atención"
'    End If
'
'    CargaDatos = bCargado
'
'End Function
'
'Private Sub ActXCodCta1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'
''        'Verifica que el num de cred no haya sido registrado
''        Set oCred = New COMDCredito.DCOMCreditos
''        If oCred.BuscaRegistroControlAdmCred(ActxCta.NroCuenta) Then
''            If MsgBox("El control de este crédito ya fue registrado, desea Regularizar las observaciones?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
''
''            Else
''                ValidaDatos = False
''                Exit Function
''            End If
''        End If
'
'        Call CargaObs(Me.ActXCodCta1.NroCuenta)
'        CmdCancelarMant_Click
'    End If
'End Sub
'
'Private Sub ActXCodCta2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Dim oCred As COMDCredito.DCOMCreditos
'        Dim oCons As COMDConstantes.DCOMConstantes
'        Dim rs As ADODB.Recordset
'        Dim rs1 As ADODB.Recordset
'        Dim L As ListItem
'
'        Set oCons = New COMDConstantes.DCOMConstantes
'        Set rs = oCons.RecuperaConstantes(9005)
'        Set oCons = Nothing
'        'Call Llenar_Combo_con_Recordset(rs, cboExoneracionesMant)
'
'        Set oCred = New COMDCredito.DCOMCreditos
'        Set rs1 = oCred.ObtieneQuienExoneraAdmCred
'        Set oCred = Nothing
'        'Call CargaQuienExoneraMant(rs1)
'
'        'Call CargaUserExoneraMant(Right(Trim(Me.cboQuienExoneMant.Text), 6))
'        Call CargaObsMant(Me.ActXCodCta2.NroCuenta)
'        'Call CargaExo(Me.ActXCodCta2.NroCuenta)    'RECO20140404 ERS174 ANEX 01
'        Call CargarExoneracionesMant 'RECO20141009
'        Call CargarAutorizacionesMant 'RECO20141009
'        cmdCancelReg_Click
'    End If
'End Sub
'
'Private Sub ActxCta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        lblcodigo(0).Caption = ""
'        lblcodigo(1).Caption = ""
'        lblcodigo(2).Caption = ""
'        lblcodigo(3).Caption = ""
'        lblcodigo(4).Caption = ""
'
'        lblcodigo(5).Caption = ""
'        lblcodigo(6).Caption = ""
'        lblcodigo(7).Caption = ""
'        lblcodigo(8).Caption = ""
'        lblcodigo(9).Caption = ""
'        lblcodigo(10).Caption = ""
'        lblcodigo(11).Caption = ""
'        lblcodigo(12).Caption = ""
'        lblcodigo(13).Caption = ""
'        lblcodigo(14).Caption = ""
'        lblcodigo(15).Caption = ""
'        ''RECO20141013 *************************
'        'If gsCodAge = "01" Then
'        '    CargaDatos ActxCta.NroCuenta
'        'ElseIf ActxCta.Age = gsCodAge Then
'        '    CargaDatos ActxCta.NroCuenta
'        'Else
'        '    MsgBox "No se puede realizar la operación, debido a que el crédito no pertenece a su agencia.", vbInformation, "Alerta SICMACM"
'        '    Call LimpiarControles
'        'End If
''        If CargaDatos(ActxCta.NroCuenta) Then
''    '            cmdImprimir.Enabled = True
''        Else
''    '            cmdImprimir.Enabled = False
''        End If
'        'RECO FIN******************************
'        If ValidaAgenciaAutorizada(ActxCta.Age) Then
'            CargaDatos ActxCta.NroCuenta
'        Else
'            MsgBox "No tiene los permisos para la revisión de crédito de esta agencia. Por favor comuníquese con la Jefatura del área de Administración de Créditos. ", vbInformation, "Aviso"
'        End If
'        'RECO FIN 20150313
'    End If
'End Sub
'
'Private Sub cboAgencia_Click()
'    Dim oCredn As COMNCredito.NCOMCredito
'    Dim rsR3 As ADODB.Recordset
'
'    Set oCredn = New COMNCredito.NCOMCredito
'    Set rsR3 = oCredn.obtenerAgenciaUserOpe(IIf(Right(Me.cboagencia.Text, 4) = "", CInt(gsCodAge), Right(Me.cboagencia.Text, 4)))
'    Set oCredn = Nothing
'
'    cboRF.Clear
'    Do While Not rsR3.EOF
'        Me.cboRF.AddItem Trim(rsR3!cConsDescripcion) & Space(100) & 1 'Trim(str(rsR3!nConsValor))
'        rsR3.MoveNext
'    Loop
'    rsR3.Close
'
''    Call Llenar_Combo_con_Recordset(rsR3, cboagencia)
'End Sub
'
''Private Sub cboAutorizaciones_Click()
''  If Me.cboAutorizaciones.ListIndex <> -1 Then
''        If CInt(Right(Me.cboAutorizaciones.Text, 3)) = 4 Then
''            Me.txtautoesp.Enabled = True
''            Me.txtautoesp.Visible = True
''        Else
''            Me.txtautoesp.Visible = False
''            Me.txtautoesp.Enabled = False
''        End If
''    End If
''End Sub
''RECO COMENT
''madm 20110308
''Private Sub cboExoneraciones_Click()
''    If Me.cboExoneraciones.ListIndex <> -1 Then
''        If CInt(Right(Me.cboExoneraciones.Text, 3)) = 4 Then
''            Me.cboExoneraDet.Enabled = True
''            Me.cboExoneraDet.Visible = True
''            LlenarCboCARDET
''        ElseIf CInt(Right(Me.cboExoneraciones.Text, 4)) = 12 Then
''            Me.txtExoneraDet.Visible = True
''            Me.txtExoneraDet.Enabled = True
''        Else
''            Me.cboExoneraDet.Enabled = False
''            Me.cboExoneraDet.Visible = False
''            Me.txtExoneraDet.Visible = False
''            Me.txtExoneraDet.Enabled = False
''        End If
''    End If
''End Sub
''RECO COMEN FIN
''RECO20140328 ERS174-2013 ANEXO 01**************************
'Private Sub cboAutorizacion_Click()
'    Call LimpiarFE(Me.FEApoderadoAuto)
'    Call CargarNivelAutorizacion(cboNivAutorizacion, FEApoderadoAuto, cboAutorizacion, 1)
'    Call CargarCargoNivExoAuto(Me.cboNivAutorizacion, Me.FEApoderadoAuto, Me.cboAutorizacion, 2)
'End Sub
'
'Private Sub cboAutorizacionMant_Click()
'    Call LimpiarFE(Me.FEApoderadoAutMant)
'    Call CargarNivelAutorizacion(cboNivelAutMant, FEApoderadoAutMant, cboAutorizacionMant, 2)
'    Call CargarCargoNivExoAuto(Me.cboNivelAutMant, Me.FEApoderadoAutMant, Me.cboAutorizacionMant, 2)
'End Sub
'
'Private Sub cboExoneraciones_Click()
'    Call LimpiarFE(Me.FEApoderados)
'    Call CargarNivelExoneracion(Me.cboNivelExonera, Me.FEApoderados, Me.cboExoneraciones, 1)
'    Call CargarCargoNivExoAuto(Me.cboNivelExonera, Me.FEApoderados, Me.cboExoneraciones, 1)
'End Sub
'
'Private Sub cboExoneraMat_Click()
'    Call LimpiarFE(Me.FEApoderadoExoneraMat)
'    Call CargarNivelExoneracion(Me.cboNivExoneraMat, Me.FEApoderadoExoneraMat, Me.cboExoneraMat, 2)
'    Call CargarCargoNivExoAuto(Me.cboExoneraMat, Me.FEApoderadoExoneraMat, Me.cboExoneraMat, 3)
'End Sub
''RECO FIN***************************************************
'
''Private Sub cboExoneracionesMant_Click()
''    If Me.cboExoneracionesMant.ListIndex <> -1 Then
''        If CInt(Right(Me.cboExoneracionesMant.Text, 4)) = 4 Then
''            Me.cboExoneraDetMant.Enabled = True
''            Me.cboExoneraDetMant.Visible = True
''            LlenarCboCARDET1
''        ElseIf CInt(Right(Me.cboExoneracionesMant.Text, 4)) = 12 Then
''            Me.txtExoneraDetMant.Visible = True
''            Me.txtExoneraDetMant.Enabled = True
''            Me.txtExoneraDetMant.Text = ""
''        Else
''            Me.cboExoneraDetMant.Enabled = False
''            Me.cboExoneraDetMant.Visible = False
''            Me.txtExoneraDetMant.Visible = False
''            Me.txtExoneraDetMant.Enabled = False
''        End If
''    End If
''End Sub
'''end madm
'
''Private Sub cboExoneraciones_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{Tab}", True
''    End If
''End Sub
''RECO20140328 ERS174-2013 ANEXO 01***************
'Private Sub cboNivAutorizacion_Click()
'    Call CargarCargoNivExoAuto(Me.cboNivAutorizacion, Me.FEApoderadoAuto, Me.cboAutorizacion, 2)
'End Sub
''RECO20141009************************************
'Private Sub cboNivelAutMant_Click()
'    Call CargarCargoNivExoAuto(Me.cboNivelAutMant, Me.FEApoderadoAutMant, Me.cboAutorizacionMant, 4)
'End Sub
''RECO FIN****************************************
'Private Sub cboNivelExonera_Click()
'    Call CargarCargoNivExoAuto(Me.cboNivelExonera, Me.FEApoderados, Me.cboExoneraciones, 1)
'End Sub
'Private Sub cboNivExoneraMat_Click()
'    Call CargarCargoNivExoAuto(Me.cboNivExoneraMat, Me.FEApoderadoExoneraMat, Me.cboExoneraMat, 3)
'End Sub
''RECO FIN****************************************
'
''Private Sub cboQuienAuto_Click()
'' 'MADM 20110614
''            If cboQuienAuto.ListIndex <> -1 Then
''                If CInt(Right(Trim(Me.cboQuienAuto.Text), 6)) = 8 Then
''                        cboUsuAuto2.Enabled = True
''                        cboUsuAuto3.Enabled = True
''                        cboUsuarioAuto.Enabled = True
''                ElseIf CInt(Right(Trim(Me.cboQuienAuto.Text), 6)) = 7 Then
''                        cboUsuarioAuto.Enabled = True
''                        cboUsuAuto2.Enabled = True
''                        cboUsuAuto3.Enabled = False
''                Else
''                        cboUsuarioAuto.Enabled = True
''                        cboUsuExo2.Enabled = False
''                        cboUsuExo3.Enabled = False
''                End If
''                cboAutorizaciones.Enabled = True
''            End If
''        'end madm
''End Sub
'
''Private Sub cboQuienExone_Click()
'''    If cboQuienExone.ListIndex = -1 Then
'''            MsgBox "Debe Escoger quien exonera.", vbInformation, "Aviso"
'''            Exit Sub
'''    End If
'''    Call CargaUserExonera(Right(Trim(Me.cboQuienExone.Text), 6))
''
''    'MADM 20110614
''            If cboQuienExone.ListIndex <> -1 Then
''                If CInt(Right(Trim(Me.cboQuienExone.Text), 6)) = 8 Then
''                        cboUsuExo2.Enabled = True
''                        cboUsuExo3.Enabled = True
''                        cboUsuarioExone.Enabled = True
''                ElseIf CInt(Right(Trim(Me.cboQuienExone.Text), 6)) = 7 Then
''                        cboUsuarioExone.Enabled = True
''                        cboUsuExo2.Enabled = True
''                        cboUsuExo3.Enabled = False
''                Else
''                        cboUsuarioExone.Enabled = True
''                        cboUsuExo2.Enabled = False
''                        cboUsuExo3.Enabled = False
''                End If
''                cboExoneraciones.Enabled = True
''            End If
''        'end madm
''End Sub
'
''Private Sub cboQuienExone_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{Tab}", True
''    End If
''End Sub
'
''Private Sub cboQuienExoneMant_Click()
'' 'MADM 20110614
''            If cboQuienExoneMant.ListIndex <> -1 Then
''                If CInt(Right(Trim(Me.cboQuienExoneMant.Text), 6)) = 8 Then
''                        cboUsuExo2Mant.Enabled = True
''                        cboUsuExo3Mant.Enabled = True
''                        cboUsuarioExoneMant.Enabled = True
''                ElseIf CInt(Right(Trim(Me.cboQuienExoneMant.Text), 6)) = 7 Then
''                        cboUsuarioExoneMant.Enabled = True
''                        cboUsuExo2Mant.Enabled = True
''                        cboUsuExo3Mant.Enabled = False
''                Else
''                        cboUsuarioExoneMant.Enabled = True
''                        cboUsuExo2Mant.Enabled = False
''                        cboUsuExo3Mant.Enabled = False
''                End If
''                cboExoneracionesMant.Enabled = True
''            End If
''        'end madm
''End Sub
'
''Private Sub cboUsuarioExone_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{Tab}", True
''    End If
''End Sub
''RECO20140404 ERS174-2013 ANEXO 01******************************************************************
''Private Sub cboUsuExo2_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{Tab}", True
''    End If
''End Sub
''Private Sub cboUsuExo2_LostFocus()
''    If Left(cboUsuarioExone.Text, 4) = "" Or Left(cboUsuarioExone.Text, 4) = "0000" Then
''        MsgBox "Debe seleccionar un usuario anterior."
''        cboUsuarioExone.SetFocus
''    End If
''End Sub
''Private Sub cboUsuExo3_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{Tab}", True
''    End If
''End Sub
''Private Sub cboUsuExo3_LostFocus()
''    If Left(cboUsuExo2.Text, 4) = "" Or Left(cboUsuExo2.Text, 4) = "0000" Then
''        MsgBox "Debe seleccionar un usuario anterior."
''        cboUsuExo2.SetFocus
''    End If
''End Sub
''RECO FIN*****************************************************************************************
'
''WIOR 20140405
'Private Sub chkClientePref_Click()
'If Me.chkClientePref.value = 1 Then
'    Me.chkPostDesembolso.value = 0
'    w.TabVisible(3) = True
'Else
'    w.TabVisible(3) = False
'End If
'End Sub
''WIOR - FIN
'
'Private Sub chkCompraDeuda_Click()
'    If chkCompraDeuda.value = 1 Then
'        txtcompra.Visible = True
'        txtcompra.Enabled = True
'        txtcompra.SetFocus
'        txtcompra = ""
'    Else
'        txtcompra = ""
'        txtcompra.Enabled = False
'        txtcompra.Visible = False
'    End If
'End Sub
''WIOR 20140405
'Private Sub chkPostDesembolso_Click()
'If Me.chkPostDesembolso.value = 1 Then
'    Me.chkClientePref.value = 0
'    w.TabVisible(3) = True
'    w.TabVisible(0) = False
'    w.TabVisible(1) = False
'    w.TabVisible(2) = False
'    fbPostDesembolso = True
'Else
'    w.TabVisible(0) = True
'    w.TabVisible(1) = True
'    w.TabVisible(2) = True
'    w.TabVisible(3) = False
'    fbPostDesembolso = False
'End If
'End Sub
''WIOR - FIN
''RECO20140404 ERS174-2013 ANEXO 01***********************************************
''Private Sub cmdAgregaAuto_Click()
''
''    If Me.cboAutorizaciones.ListIndex = -1 Or Me.cboAutorizaciones.ListIndex = -1 Then
''        MsgBox "Seleccione una Autorizacion o verifique quien autoriza.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
''
''    If Me.cboUsuarioAuto.ListIndex = -1 Or Left(Me.cboUsuarioAuto.Text, 4) = "0000" Then
''        MsgBox "Seleccione al menos un usuario.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '*************verifica que los usuarios no se repitan
''    lnBuscaUsu = IIf(Left(Me.cboUsuarioAuto.Text, 4) = Left(Me.cboUsuAuto2, 4), 1, IIf(Left(Me.cboUsuarioAuto.Text, 4) = Left(Me.cboUsuAuto3.Text, 4), 1, IIf(Left(Me.cboUsuAuto2.Text, 4) = Left(Me.cboUsuAuto3, 4) And Len(Left(Me.cboUsuAuto2.Text, 4)) = 4 And Left(Me.cboUsuAuto2.Text, 4) <> "0000", 1, 0)))
''    If lnBuscaUsu = 1 Then
''        MsgBox "Los usuarios escogidos no deben repetirse.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '***************---------------------------------------------------
'    '*********verifica vacios intermedios
''    If Left(Me.cboUsuAuto3, 4) = "0000" Or cboUsuAuto3.ListIndex = -1 Then
''
''    Else
''        If Left(cboUsuAuto2, 4) = "0000" Or cboUsuAuto2.ListIndex = -1 Then
''            MsgBox "Debe seleccionar el segundo usuario.", vbOKOnly + vbInformation, "Atención"
''            Exit Sub
''        End If
''    End If
'    '*******************---------------------------------------------------
'    ' para que no grabe mas de una vez las exoneraciones
''    bEncontrado = False
''    For i = 1 To FlexAuto.Rows - 1
''        If FlexAuto.TextMatrix(i, 3) <> "" Then
''            If txtautoesp.Visible = False Then
''                If CInt(FlexAuto.TextMatrix(i, 7)) = CInt(Right(Trim(Me.cboAutorizaciones.Text), 3)) Then
''                    bEncontrado = True
''                    Exit For
''                End If
''            End If
''        End If
''    Next i
''    If bEncontrado Then
''        MsgBox "Este dato ya fue ingresado.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '------------------------------------------------
''    lcAuto = Mid(cboAutorizaciones.Text, 1, Len(cboAutorizaciones.Text) - 3)
''    With FlexAuto
''        .AdicionaFila
''        .TextMatrix(FlexAuto.row, 1) = Trim(lcAuto)
''        .TextMatrix(FlexAuto.row, 2) = Trim(cboQuienAuto.Text)
''        .TextMatrix(FlexAuto.row, 3) = Left(Me.cboUsuarioAuto.Text, 4)
''        .TextMatrix(FlexAuto.row, 4) = IIf(Left(Me.cboUsuAuto2.Text, 4) <> "0000", Left(Me.cboUsuAuto2.Text, 4), "")
''        .TextMatrix(FlexAuto.row, 5) = IIf(Left(Me.cboUsuAuto3.Text, 4) <> "0000", Left(Me.cboUsuAuto3.Text, 4), "")
''        .TextMatrix(FlexAuto.row, 6) = Trim(ActxCta.NroCuenta)
''        .TextMatrix(FlexAuto.row, 7) = CInt(Right(Trim(Me.cboAutorizaciones.Text), 3))
''        .TextMatrix(FlexAuto.row, 8) = Trim(Right(Me.cboQuienAuto, 6))
''
''        If Me.txtautoesp.Visible Then
''            .TextMatrix(FlexAuto.row, 9) = Trim(txtautoesp.Text)
''        Else
''            .TextMatrix(FlexAuto.row, 9) = ""
''        End If
''
''    End With
''
''    cboAutorizaciones.ListIndex = -1
''    Me.cboQuienAuto.ListIndex = -1
''    Me.cboUsuarioAuto.ListIndex = -1
''    Me.cboUsuAuto2.ListIndex = -1
''    Me.cboUsuAuto3.ListIndex = -1
''    cboAutorizaciones.SetFocus
''
''     If Me.txtautoesp.Visible Then
''        txtautoesp.Text = ""
''    End If
''    'madm 20110308
''    If Me.txtAutorizacionesDet.Visible Then
''        Me.txtAutorizacionesDet.Text = ""
''        Me.txtAutorizacionesDet.Enabled = False
''        Me.txtAutorizacionesDet.Visible = False
''    End If
''    'end madm
''End Sub
''Private Sub cmdAgregaExo_Click()
''
''    If Me.cboExoneraciones.ListIndex = -1 Or Me.cboQuienExone.ListIndex = -1 Then
''        MsgBox "Seleccione una exoneracion o verifique quien exonera.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
''
''    If Me.cboUsuarioExone.ListIndex = -1 Or Left(Me.cboUsuarioExone.Text, 4) = "0000" Then
''        MsgBox "Seleccione al menos un usuario.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '*************verifica que los usuarios no se repitan
''    lnBuscaUsu = IIf(Left(Me.cboUsuarioExone.Text, 4) = Left(Me.cboUsuExo2, 4), 1, IIf(Left(Me.cboUsuarioExone.Text, 4) = Left(Me.cboUsuExo3.Text, 4), 1, IIf(Left(Me.cboUsuExo2.Text, 4) = Left(Me.cboUsuExo3, 4) And Len(Left(Me.cboUsuExo2.Text, 4)) = 4 And Left(Me.cboUsuExo2.Text, 4) <> "0000", 1, 0)))
''    If lnBuscaUsu = 1 Then
''        MsgBox "Los usuarios escogidos no deben repetirse.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '***************---------------------------------------------------
'    '*********verifica vacios intermedios
''    If Left(Me.cboUsuExo3, 4) = "0000" Or cboUsuExo3.ListIndex = -1 Then
''
''    Else
''        If Left(cboUsuExo2, 4) = "0000" Or cboUsuExo2.ListIndex = -1 Then
''            MsgBox "Debe seleccionar el segundo usuario.", vbOKOnly + vbInformation, "Atención"
''            Exit Sub
''        End If
''    End If
'    '*******************---------------------------------------------------
'    ' para que no grabe mas de una vez las exoneraciones
''    bEncontrado = False
''    For i = 1 To FlexExo.Rows - 1
''        If FlexExo.TextMatrix(i, 3) <> "" Then
''            If CInt(FlexExo.TextMatrix(i, 7)) = CInt(Right(Trim(Me.cboExoneraciones.Text), 3)) Then
''                bEncontrado = True
''                Exit For
''            End If
''        End If
''    Next i
''    If bEncontrado Then
''        MsgBox "Este dato ya fue ingresado.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '------------------------------------------------
''    lcExonera = Mid(cboExoneraciones.Text, 1, Len(cboExoneraciones.Text) - 3)
''    With FlexExo
''        .AdicionaFila
''        .TextMatrix(FlexExo.row, 1) = Trim(lcExonera)
''        .TextMatrix(FlexExo.row, 2) = Trim(cboQuienExone.Text)
''        .TextMatrix(FlexExo.row, 3) = Left(Me.cboUsuarioExone.Text, 4)
''        .TextMatrix(FlexExo.row, 4) = IIf(Left(Me.cboUsuExo2.Text, 4) <> "0000", Left(Me.cboUsuExo2.Text, 4), "")
''        .TextMatrix(FlexExo.row, 5) = IIf(Left(Me.cboUsuExo3.Text, 4) <> "0000", Left(Me.cboUsuExo3.Text, 4), "")
''        .TextMatrix(FlexExo.row, 6) = Trim(ActxCta.NroCuenta)
''        .TextMatrix(FlexExo.row, 7) = CInt(Right(Trim(Me.cboExoneraciones.Text), 3))
''        .TextMatrix(FlexExo.row, 8) = Trim(Right(Me.cboQuienExone, 6))
''        '.TextMatrix(FlexExo.Row, 9) = IIf(Me.cboExoneraDet.Visible, CInt(Right(Me.cboExoneraDet, 6)), "")
''        If Me.cboExoneraDet.Visible And Me.cboExoneraDet.ListIndex <> -1 Then
''            .TextMatrix(FlexExo.row, 9) = CInt(Right(Me.cboExoneraDet.Text, 4))
''        Else
''            .TextMatrix(FlexExo.row, 9) = "0"
''        End If
''
''        .TextMatrix(FlexExo.row, 10) = IIf(Me.txtExoneraDet.Visible, Trim(Me.txtExoneraDet.Text), "")
''    End With
''
''    cboExoneraciones.ListIndex = -1
''    Me.cboQuienExone.ListIndex = -1
''    Me.cboUsuarioExone.ListIndex = -1
''    Me.cboUsuExo2.ListIndex = -1
''    Me.cboUsuExo3.ListIndex = -1
''    cboExoneraciones.SetFocus
'    'madm 20110308
''    If Me.cboExoneraDet.Visible Then
''        Me.cboExoneraDet.ListIndex = -1
''        Me.cboExoneraDet.Enabled = False
''        Me.cboExoneraDet.Visible = False
''    End If
''    If Me.txtExoneraDet.Visible Then
''        Me.txtExoneraDet.Text = ""
''        Me.txtExoneraDet.Enabled = False
''        Me.txtExoneraDet.Visible = False
''    End If
'    'end madm
''End Sub
''Private Sub cmdAgregaExoMant_Click()
'' If Me.cboExoneracionesMant.ListIndex = -1 Or Me.cboQuienExoneMant.ListIndex = -1 Then
''        MsgBox "Seleccione una exoneracion o verifique quien exonera.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
''
''    If Me.cboUsuarioExoneMant.ListIndex = -1 Or Left(Me.cboUsuarioExoneMant.Text, 4) = "0000" Then
''        MsgBox "Seleccione al menos un usuario.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '*************verifica que los usuarios no se repitan
''    lnBuscaUsu = IIf(Left(Me.cboUsuarioExoneMant.Text, 4) = Left(Me.cboUsuExo2Mant, 4), 1, IIf(Left(Me.cboUsuarioExoneMant.Text, 4) = Left(Me.cboUsuExo3Mant.Text, 4), 1, IIf(Left(Me.cboUsuExo2Mant.Text, 4) = Left(Me.cboUsuExo3Mant, 4) And Len(Left(Me.cboUsuExo2Mant.Text, 4)) = 4 And Left(Me.cboUsuExo2Mant.Text, 4) <> "0000", 1, 0)))
''    If lnBuscaUsu = 1 Then
''        MsgBox "Los usuarios escogidos no deben repetirse.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '***************---------------------------------------------------
'    '*********verifica vacios intermedios
''    If Left(Me.cboUsuExo3Mant, 4) = "0000" Or cboUsuExo3Mant.ListIndex = -1 Then
''
''    Else
''        If Left(cboUsuExo2Mant, 4) = "0000" Or cboUsuExo2Mant.ListIndex = -1 Then
''            MsgBox "Debe seleccionar el segundo usuario.", vbOKOnly + vbInformation, "Atención"
''            Exit Sub
''        End If
''    End If
'    '*******************---------------------------------------------------
'    ' para que no grabe mas de una vez las exoneraciones
''    bEncontrado = False
''    For i = 1 To FlexExoMant.Rows - 1
''        If FlexExoMant.TextMatrix(i, 3) <> "" Then
''            If CInt(FlexExoMant.TextMatrix(i, 7)) = CInt(Right(Trim(Me.cboExoneracionesMant.Text), 3)) Then
''                bEncontrado = True
''                Exit For
''            End If
''        End If
''    Next i
''    If bEncontrado Then
''        MsgBox "Este dato ya fue ingresado.", vbOKOnly + vbInformation, "Atención"
''        Exit Sub
''    End If
'    '------------------------------------------------
''    lcExoneraMant = Mid(cboExoneracionesMant.Text, 1, Len(cboExoneracionesMant.Text) - 3)
''    With FlexExoMant
''        .AdicionaFila
''        .TextMatrix(FlexExoMant.row, 1) = Trim(lcExoneraMant)
''        .TextMatrix(FlexExoMant.row, 2) = Trim(cboQuienExoneMant.Text)
''        .TextMatrix(FlexExoMant.row, 3) = Left(Me.cboUsuarioExoneMant.Text, 4)
''        .TextMatrix(FlexExoMant.row, 4) = IIf(Left(Me.cboUsuExo2Mant.Text, 4) <> "0000", Left(Me.cboUsuExo2Mant.Text, 4), "")
''        .TextMatrix(FlexExoMant.row, 5) = IIf(Left(Me.cboUsuExo3Mant.Text, 4) <> "0000", Left(Me.cboUsuExo3Mant.Text, 4), "")
''        .TextMatrix(FlexExoMant.row, 6) = CInt(Right(Trim(Me.cboExoneracionesMant.Text), 4))
''        .TextMatrix(FlexExoMant.row, 7) = Trim(Right(Me.cboQuienExoneMant, 4))
''        If Me.cboExoneraDetMant.Visible And Me.cboExoneraDetMant.ListIndex <> -1 Then
''            .TextMatrix(FlexExoMant.row, 8) = CInt(Right(Me.cboExoneraDetMant.Text, 4))
''        Else
''            .TextMatrix(FlexExoMant.row, 8) = "0"
''        End If
''        .TextMatrix(FlexExoMant.row, 9) = IIf(Me.txtExoneraDetMant.Visible, Trim(Me.txtExoneraDetMant.Text), "")
''    End With
''
''    cboExoneracionesMant.ListIndex = -1
''    Me.cboQuienExoneMant.ListIndex = -1
''    Me.cboUsuarioExoneMant.ListIndex = -1
''    Me.cboUsuExo2Mant.ListIndex = -1
''    Me.cboUsuExo3Mant.ListIndex = -1
''    cboExoneracionesMant.SetFocus
'    'madm 20110308
''    If Me.cboExoneraDetMant.Visible Then
''        Me.cboExoneraDetMant.ListIndex = -1
''        Me.cboExoneraDetMant.Enabled = False
''        Me.cboExoneraDetMant.Visible = False
''    End If
''    If Me.txtExoneraDetMant.Visible Then
''        Me.txtExoneraDetMant.Text = ""
''        Me.txtExoneraDetMant.Enabled = False
''        Me.txtExoneraDetMant.Visible = False
''    End If
''End Sub
''RECO FIN*********************************************************************************************************
'
''Private Sub chkExoneracion_Click()
''    If chkExoneracion.value = 1 Then
''        cboExoneraciones.Enabled = True
''        cboQuienExone.Enabled = True
''        cboUsuarioExone.Enabled = True
''    Else
''        cboExoneraciones.Enabled = False
''        cboQuienExone.Enabled = False
''        cboUsuarioExone.Enabled = False
''    End If
''End Sub
'
'Private Sub cmdAgregaObs_Click()
'
'    If Len(txtObservaciones.Text) = 0 Then Exit Sub
'
'    If FlexObs.Rows - 1 > 0 Then
'        For i = 1 To FlexObs.Rows - 1
'            If Trim(FlexObs.TextMatrix(i, 1)) = Trim(txtObservaciones.Text) Then
'                MsgBox "Esta Observacion ya fue registrada.", vbOKOnly + vbInformation, "Atención"
'                Exit Sub
'            End If
'        Next i
'    End If
'
'    With FlexObs
'        .AdicionaFila
'        .TextMatrix(FlexObs.row, 1) = Trim(txtObservaciones.Text)
'        .TextMatrix(FlexObs.row, 2) = Trim(ActxCta.NroCuenta)
'    End With
'
'    txtObservaciones.Text = ""
'    txtObservaciones.SetFocus
'
'End Sub
'
'Private Sub cmdAgregaObsMant_Click()
'  If Len(txtObservacionesMant.Text) = 0 Then Exit Sub
'
'    If FlexObs.Rows - 1 > 0 Then
'        For i = 1 To FlexObsMant.Rows - 1
'            If Trim(FlexObsMant.TextMatrix(i, 1)) = Trim(txtObservacionesMant.Text) Then
'                MsgBox "Esta Observacion ya fue registrada.", vbOKOnly + vbInformation, "Atención"
'                Exit Sub
'            End If
'        Next i
'    End If
'
'    With FlexObsMant
'        .AdicionaFila
'        .TextMatrix(FlexObsMant.row, 1) = Trim(UCase(txtObservacionesMant.Text))
'        .TextMatrix(FlexObsMant.row, 2) = Trim(Me.ActXCodCta2.NroCuenta)
'    End With
'
'    txtObservacionesMant.Text = ""
'    txtObservacionesMant.SetFocus
'End Sub
'
'Private Sub CmdAgregaPost_Click()
'If Len(txtdesembolso.Text) = 0 Then Exit Sub
'
'    If flexPost.Rows - 1 > 0 Then
'        For i = 1 To flexPost.Rows - 1
'            If Trim(flexPost.TextMatrix(i, 1)) = Trim(txtdesembolso.Text) Then
'                MsgBox "Esta Observacion del Post Desembolso ya fue registrada.", vbOKOnly + vbInformation, "Atención"
'                Exit Sub
'            End If
'        Next i
'    End If
'
'    With flexPost
'        .AdicionaFila
'        .TextMatrix(flexPost.row, 1) = Trim(Me.txtdesembolso.Text)
'        .TextMatrix(flexPost.row, 2) = Trim(Left(cboRF.Text, 4))
'        .TextMatrix(flexPost.row, 3) = Trim(Right(cboagencia.Text, 2))
'    End With
'
'    cboRF.ListIndex = -1
'    cboagencia.ListIndex = -1
'    txtdesembolso.Text = ""
'    txtdesembolso.SetFocus
'End Sub
'
''RECO20140328 ERS174-2013 ANEXO 01*****************
'Private Sub cmdAgregarExo_Click()
'    'FormateaFlex FEExoneraciones
'    If ValidaDatosFE(Me.FEApoderados) = False Then
'        MsgBox "No se pueden agregar datos vacios, por favor complete la información", vbCritical, "Aviso"
'        Exit Sub
'    End If
'    Dim nIndex As Integer
'        FEExoneraciones.AdicionaFila
'        FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 1) = Me.cboExoneraciones.Text
'        FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 2) = Me.cboNivelExonera.Text
'    For nIndex = 1 To nCanFirmasExo
'        FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 2 + nIndex) = Me.FEApoderados.TextMatrix(nIndex, 2)
'        'sPersCargoExo = sPersCargoExo & Trim(Mid(Me.FEApoderados.TextMatrix(nIndex, 2), Len(Me.FEApoderados.TextMatrix(nIndex, 2)) - 13, Len(Me.FEApoderados.TextMatrix(nIndex, 2)))) & ","
'    Next
'    'FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 4) = sPersCargoExo
'    sPersCargoExo = ""
'End Sub
''RECO FIN******************************************
'
''RECO20140202 ERS174-2013 ANEXO 01**********************************************
''Private Sub cmdBorraAuto_Click()
'' If MsgBox("¿Está seguro de borrar esta linea?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
''    Call FlexAuto.EliminaFila(FlexAuto.row)
''End Sub
''Private Sub cmdBorraExo_Click()
''    If MsgBox("¿Está seguro de borrar esta linea?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
''    Call FlexExo.EliminaFila(FlexExo.row)
''End Sub
''Private Sub cmdBorraExoMant_Click()
''    If MsgBox("¿Está seguro de borrar esta linea?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
''    Call FlexExoMant.EliminaFila(FlexExo.row)
''End Sub
''RECO FIN************************************************************************
'
''RECO20140328 ERS174-2013 ANEXO 01*****************
'Private Sub cmdAgregarExoneraMat_Click()
'    If ValidaDatosFE(Me.FEApoderadoExoneraMat) = False Then
'        MsgBox "No se pueden agregar datos vacios, por favor complete la información", vbCritical, "Aviso"
'        Exit Sub
'    End If
'    'FormateaFlex Me.FEExoneraMat
'    Dim nIndex As Integer
'    FEExoneraMat.AdicionaFila
'    FEExoneraMat.TextMatrix(FEExoneraMat.Rows - 1, 1) = Me.cboExoneraMat.Text
'    FEExoneraMat.TextMatrix(FEExoneraMat.Rows - 1, 2) = Me.cboNivExoneraMat.Text
'    For nIndex = 1 To nCanFirmasExoMant
'        FEExoneraMat.TextMatrix(FEExoneraMat.Rows - 1, 2 + nIndex) = Me.FEApoderadoExoneraMat.TextMatrix(nIndex, 2)
'        'sPersCargoExoMant = sPersCargoExoMant & Trim(Mid(Me.FEApoderadoExoneraMat.TextMatrix(nIndex, 2), Len(Me.FEApoderadoExoneraMat.TextMatrix(nIndex, 2)) - 13, Len(Me.FEApoderadoExoneraMat.TextMatrix(nIndex, 2)))) & ","
'    Next
'    'FEExoneraMat.TextMatrix(FEExoneraMat.Rows - 1, 4) = sPersCargoExoMant
'End Sub
'Private Sub cmdAgregarAutorizacion_Click()
'
'    If ValidaDatosFE(Me.FEApoderadoAuto) = False Then
'        MsgBox "No se pueden agregar datos vacios, por favor complete la información", vbCritical, "Aviso"
'        Exit Sub
'    End If
'
'    'FormateaFlex Me.FEAutorizaciones
'    Dim nIndex As Integer
'    FEAutorizaciones.AdicionaFila
'    FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 1) = Me.cboAutorizacion.Text
'    FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 2) = Me.cboNivAutorizacion.Text
'    For nIndex = 1 To nCanFirmas
'        FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 2 + nIndex) = Me.FEApoderadoAuto.TextMatrix(nIndex, 2)
'        'sPersCargoAuto = sPersCargoAuto & Trim(Mid(Me.FEApoderadoAuto.TextMatrix(nIndex, 2), Len(Me.FEApoderadoAuto.TextMatrix(nIndex, 2)) - 13, Len(Me.FEApoderadoAuto.TextMatrix(nIndex, 2)))) & ","
'    Next
'    'FEAutorizaciones.TextMatrix(FEAutorizaciones.Rows - 1, 4) = sPersCargoAuto
'End Sub
''RECO FIN******************************************
''RECO20141009**************************************
'Private Sub cmdAgregaAutMant_Click()
'    If ValidaDatosFE(Me.FEApoderadoAutMant) = False Then
'        MsgBox "No se pueden agregar datos vacios, por favor complete la información", vbCritical, "Aviso"
'        Exit Sub
'    End If
'
'    'FormateaFlex Me.FEAutorizacionesMant
'    Dim nIndex As Integer
'    FEAutorizacionesMant.AdicionaFila
'    FEAutorizacionesMant.TextMatrix(FEAutorizacionesMant.Rows - 1, 1) = Me.cboAutorizacionMant.Text
'    FEAutorizacionesMant.TextMatrix(FEAutorizacionesMant.Rows - 1, 2) = Me.cboNivelAutMant.Text
'    For nIndex = 1 To nCanFirmasAutMant
'        FEAutorizacionesMant.TextMatrix(FEAutorizacionesMant.Rows - 1, 2 + nIndex) = Me.FEApoderadoAutMant.TextMatrix(nIndex, 2)
'    Next
'End Sub
''RECO FIN******************************************
'Private Sub cmdBorraObs_Click()
'    If MsgBox("¿Está seguro de borrar esta linea?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
'    Call FlexObs.EliminaFila(FlexObs.row)
'End Sub
'
'Private Sub cmdBorraObsMant_Click()
'    If MsgBox("¿Está seguro de borrar esta linea?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
'    Call FlexObsMant.EliminaFila(FlexObsMant.row)
'End Sub
'
'Private Sub cmdborrarPost_Click()
' If MsgBox("¿Está seguro de borrar esta linea?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
'    Call flexPost.EliminaFila(flexPost.row)
'End Sub
'
'Private Sub CmdBuscaMant_Click()
'Dim oCredito As COMDCredito.DCOMCreditos
'Dim R As ADODB.Recordset
'Dim R1 As ADODB.Recordset
'Dim oPers As COMDPersona.UCOMPersona
'
'    LstObsMant.Clear
'    Set oPers = frmBuscaPersona.Inicio()
'    If Not oPers Is Nothing Then
'        Set oCredito = New COMDCredito.DCOMCreditos
'
'            Set R = oCredito.RecuperaCuentasExoObsParaAdmCred(oPers.sPersCod)
'
'        Do While Not R.EOF
'            LstObsMant.AddItem R!cCtaCod
'            R.MoveNext
'        Loop
'        R.Close
'        Set R = Nothing
'        Set oCredito = Nothing
'    End If
'    If LstObsMant.ListCount = 0 Then
'        MsgBox "El Cliente No Tiene Creditos con Observaciones.", vbInformation, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdBuscaObs_Click()
'Dim oCredito As COMDCredito.DCOMCreditos
'Dim R As ADODB.Recordset
'Dim oPers As COMDPersona.UCOMPersona
'
'    LstObs.Clear
'    Set oPers = frmBuscaPersona.Inicio()
'    If Not oPers Is Nothing Then
'        Set oCredito = New COMDCredito.DCOMCreditos
'        Set R = oCredito.RecuperaCuentasObsParaAdmCred(oPers.sPersCod)
'        Do While Not R.EOF
'            LstObs.AddItem R!cCtaCod
'            R.MoveNext
'        Loop
'        R.Close
'        Set R = Nothing
'        Set oCredito = Nothing
'    End If
'    If LstObs.ListCount = 0 Then
'        MsgBox "El Cliente No Tiene Creditos con Observaciones.", vbInformation, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdBuscar_Click()
'Dim oCredito As COMDCredito.DCOMCreditos
'Dim R As ADODB.Recordset
'Dim oPers As COMDPersona.UCOMPersona
'
'    LstCred.Clear
'    Set oPers = frmBuscaPersona.Inicio()
'    If Not oPers Is Nothing Then
'        Set oCredito = New COMDCredito.DCOMCreditos
'        Set R = oCredito.RecuperaCuentasParaAdmCred(oPers.sPersCod, 1) '1=CONTROL DE CREDI
'        Do While Not R.EOF
'            LstCred.AddItem R!cCtaCod
'            R.MoveNext
'        Loop
'        R.Close
'        Set R = Nothing
'        Set oCredito = Nothing
'    End If
'    If LstCred.ListCount = 0 Then
'        MsgBox "El Cliente No Tiene Creditos en estado Aprobado.", vbInformation, "Aviso"
'    End If
'End Sub
'
'Private Sub cmdCancelar_Click()
'     LimpiarControles
'End Sub
'
'Private Sub CmdCancelarMant_Click()
' LimpiarControlesMant
'End Sub
'
'Private Sub cmdCancelReg_Click()
'    LimpiarControlesObs
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oCred As COMDCredito.DCOMCreditos
'Dim vCodCta As String
'Dim lsMovNro As String
'Dim lnCodExone As Integer
'Dim lnCodAuto As Integer
'    If validaDatos = False Then
'        Exit Sub
'    End If
'    vCodCta = ActxCta.NroCuenta
'
'
'
'    If MsgBox("¿Desea Registrar el Control de este Crédito?.", vbInformation + vbYesNo, "Atención") = vbYes Then
'        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        Set oCred = New COMDCredito.DCOMCreditos
''        If nCuentaExo > 0 Then
''            Call oCred.RegistraControlAdmCred(vCodCta, Format(CDate(Me.txtFechaRevision), "yyyymmdd"), Trim(Me.txtObservaciones), lnCodExone, Trim(Right(Me.cboQuienExone, 6)), Left(Me.cboUsuarioExone, 4), lsMovNro, Me.chkCompraDeuda.value, FlexObs.GetRsNew, FlexExo.GetRsNew, Me.chkClientePref.value, Me.ChkConstitucion.value, IIf(txtcompra.Visible = True, Trim(txtcompra.Text), ""), FlexAuto.GetRsNew)
''        Else
''            Call oCred.RegistraControlAdmCred(vCodCta, Format(CDate(Me.txtFechaRevision), "yyyymmdd"), Trim(Me.txtObservaciones), , , ,                                                                      lsMovNro, Me.chkCompraDeuda.value, FlexObs.GetRsNew,                , Me.chkClientePref.value, Me.ChkConstitucion.value, IIf(txtcompra.Visible = True, Trim(txtcompra.Text), ""), FlexAuto.GetRsNew, nCuentaPost, flexPost.GetRsNew)
''        End If
'
'
'         'RECO20140204 ERS174-2013 ANEXO 01
'         Call oCred.RegistraControlAdmCred(vCodCta, Format(CDate(Me.txtFechaRevision), "yyyymmdd"), Trim(Me.txtObservaciones), , , , lsMovNro, Me.chkCompraDeuda.value, FlexObs.GetRsNew, LlenarRsExonera(FEExoneraciones), Me.chkClientePref.value, Me.ChkConstitucion.value, IIf(txtcompra.Visible = True, Trim(txtcompra.Text), ""), LlenarRsAuto(FEAutorizaciones), nCuentaPost, flexPost.GetRsNew, fbPostDesembolso, chkRevisaDesembControl.value)
'         'JUEZ 20140725 Se agregó chkRevisaDesembControl.value
'         'WIOR 20120505
'         'Call oCred.RegistraControlAdmCred(vCodCta, Format(CDate(Me.txtFechaRevision), "yyyymmdd"), Trim(Me.txtObservaciones), lnCodExone, Trim(Right(Me.cboQuienExone, 6)), Left(Me.cboUsuarioExone, 4), lsMovNro, Me.chkCompraDeuda.value, FlexObs.GetRsNew, FlexExo.GetRsNew, Me.chkClientePref.value, Me.ChkConstitucion.value, IIf(txtcompra.Visible = True, Trim(txtcompra.Text), ""), FlexAuto.GetRsNew, nCuentaPost, flexPost.GetRsNew, fbPostDesembolso)
'
'        Set oCred = Nothing
'
'        MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly, "Atención"
'        LimpiarControles
'        'WIOR 20120508-INICIO
'        Call LimpiaFlex(FlexObs)
'        'Call LimpiaFlex(FlexExo)
'        'Call LimpiaFlex(FlexAuto)
'        Call LimpiaFlex(flexPost)
'        'FIN
'    End If
'End Sub
'
'Private Sub cmdGrabarMant_Click()
'Dim oCred1 As COMDCredito.DCOMCreditos
'Dim vCodCta As String
'Dim lsMovNro As String
'Dim lnCodExoneMant As Integer
'    If ValidaDatosMant = False Then
'        Exit Sub
'    End If
'
'    vCodCta = ActXCodCta2.NroCuenta
'
'    If MsgBox("¿Desea Actualizar el Control de este Crédito?.", vbInformation + vbYesNo, "Atención") = vbYes Then
'
'        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'        Set oCred1 = New COMDCredito.DCOMCreditos
'
'        Call oCred1.RegistraControlAdmCredMant(vCodCta, lsMovNro, FlexObsMant.GetRsNew, LlenarRsExoneraMat(FEExoneraMat), LlenarRsAuto(FEAutorizacionesMant))
'
'        Set oCred1 = Nothing
'
'        MsgBox "Los datos se actualizaron satisfactoriamente.", vbOKOnly, "Atención"
'
'        LimpiarControlesMant
'    End If
'End Sub
'
'Private Sub cmdGrabarReg_Click()
'Dim oCred As COMDCredito.DCOMCreditos
'Dim vCodCta As String
'Dim lsMovNro As String
'Dim lnCodExone As Integer
'
'    If ValidaDatosObs = False Then
'        Exit Sub
'    End If
'
'    'ActXCodCta1.NroCuenta
'
'    vCodCta = ActXCodCta1.NroCuenta
'
'    If MsgBox("¿Está seguro de Regularizar estas observaciones?.", vbInformation + vbYesNo, "Atención") = vbYes Then
'
'        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'        Set oCred = New COMDCredito.DCOMCreditos
'            '********RECO20131112 ERS133****************
'            'Call oCred.RegularizaObsAdmCred(vCodCta, lsMovNro, Me.FlexEdit1.GetRsNew)
'            If chkPostDesembolsoBusca.value = 1 Then
'                Call oCred.RegularizaObsPosAdmCred(vCodCta, lsMovNro, Me.FlexEdit1.GetRsNew)
'            Else
'                'Call oCred.RegularizaObsAdmCred(vCodCta, lsMovNro, Me.FlexEdit1.GetRsNew)
'                Call oCred.RegularizaObsAdmCred(vCodCta, lsMovNro, Me.FlexEdit1.GetRsNew, chkRevisaDesembObs.value) 'JUEZ 20140725
'            End If
'            '***************END RECO*********************
'        Set oCred = Nothing
'
'        MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly, "Atención"
'
'        LimpiarControlesObs
'
'    End If
'End Sub
''RECO20140804****************************************
'Private Sub cmdQuitaExo_Click()
'    FEExoneraciones.EliminaFila (FEExoneraciones.row)
'End Sub
'
''RECO FIN********************************************
''RECO20141009****************************
'Private Sub cmdQuitarExoMant_Click()
'    FEExoneraMat.EliminaFila (FEExoneraMat.row)
'End Sub
'
'Private Sub cmdQuitarAutMant_Click()
'    FEAutorizacionesMant.EliminaFila (FEAutorizacionesMant.row)
'End Sub
''RECO FIN *******************************
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub CmdSalirMant_Click()
'    Unload Me
'End Sub
''********aki
'Private Sub cmdSalirReg_Click()
'    Unload Me
'End Sub
'
'
''RECO20140308 ERS174-2013 ANEXO 01**************************************
''Private Sub FEAutorizaciones_Click()
'    'Call FEExoneraciones_OnCellChange(FEExoneraciones.row, FEExoneraciones.Col)
''End Sub
'Private Sub FEExoneraciones_OnCellChange(pnRow As Long, pnCol As Long)
'    Dim loDRHCargo As COMDPersona.UCOMAcceso
'    Dim loDatos  As ADODB.Recordset
'
'    Set loDRHCargo = New COMDPersona.UCOMAcceso
'    Set loDatos = New ADODB.Recordset
'
'    Set loDatos = loDRHCargo.RecuperaDatosPersRRHH(Mid(Me.FEExoneraciones.TextMatrix(Me.FEExoneraciones.row, 4), 1, Len(Me.FEExoneraciones.TextMatrix(Me.FEExoneraciones.row, 4)) - 1))
'    FEExoneraciones.CargaCombo loDatos
'    Set loDatos = Nothing
'End Sub
'Private Sub FEAutorizaciones_OnCellChange(pnRow As Long, pnCol As Long)
'    Dim loDRHCargo As COMDPersona.UCOMAcceso
'    Dim loDatos  As ADODB.Recordset
'
'    Set loDRHCargo = New COMDPersona.UCOMAcceso
'    Set loDatos = New ADODB.Recordset
'
'    Set loDatos = loDRHCargo.RecuperaDatosPersRRHH(Me.FEAutorizaciones.TextMatrix(Me.FEAutorizaciones.row, 4))
'    Me.FEAutorizaciones.CargaCombo loDatos
'    Set loDatos = Nothing
'End Sub
'Private Sub FEExoneraMat_OnCellChange(pnRow As Long, pnCol As Long)
'    Dim loDRHCargo As COMDPersona.UCOMAcceso
'    Dim loDatos  As ADODB.Recordset
'
'    Set loDRHCargo = New COMDPersona.UCOMAcceso
'    Set loDatos = New ADODB.Recordset
'
'    Set loDatos = loDRHCargo.RecuperaDatosPersRRHH(Me.FEExoneraMat.TextMatrix(Me.FEExoneraMat.row, 4))
'    Me.FEExoneraMat.CargaCombo loDatos
'    Set loDatos = Nothing
'End Sub
''RECO FIN***************************************************************
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
'        Dim sCuenta As String
'        sCuenta = frmValTarCodAnt.inicia(gColPYMEEmp, False)
'
'        If sCuenta <> "" Then
'            ActxCta.NroCuenta = sCuenta
'            ActxCta.SetFocusCuenta
'        End If
'
'    End If
'End Sub
'
''RECO20140328 ERS174-2013 ANEXO 01 COMENTADO*************************
''madm 20110308
''Sub LlenarCboCARDET()
''    Dim oCred As COMDCredito.DCOMCreditos
''    Dim oCons As COMDConstantes.DCOMConstantes
''    Dim rs As ADODB.Recordset
''    Dim rs1 As ADODB.Recordset
''    Dim L As ListItem
''
''    Set oCons = New COMDConstantes.DCOMConstantes
''        Set rs = oCons.RecuperaConstantes(9012)
''    Set oCons = Nothing
''    Call Llenar_Combo_con_Recordset(rs, cboExoneraDet)
''End Sub
''
''Sub LlenarCboCARDET1()
''    Dim oCred As COMDCredito.DCOMCreditos
''    Dim oCons As COMDConstantes.DCOMConstantes
''    Dim rs As ADODB.Recordset
''    Dim rs1 As ADODB.Recordset
''    Dim L As ListItem
''
''    Set oCons = New COMDConstantes.DCOMConstantes
''        Set rs = oCons.RecuperaConstantes(9012)
''    Set oCons = Nothing
''    Call Llenar_Combo_con_Recordset(rs, Me.cboExoneraDetMant)
''End Sub
''RECO FIN***********************************************************
'Private Sub Form_Load()
'
'    CentraForm Me
'
'    Dim oCred As COMDCredito.DCOMCreditos
'    Dim oCons As COMDConstantes.DCOMConstantes
'    Dim rs As ADODB.Recordset
'    Dim rs1 As ADODB.Recordset
'    Dim Rs2 As ADODB.Recordset
'    Dim lrAgenc As ADODB.Recordset
'    Dim L As ListItem
'    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
'
'    Set oCons = New COMDConstantes.DCOMConstantes
'    Set rs = oCons.RecuperaConstantes(9005)
'    'Call Llenar_Combo_con_Recordset(rs, cboExoneraciones)  'RECO20140328 ERS174-2013 ANEXO 01 COMENTADO
'
'    Set oCred = New COMDCredito.DCOMCreditos
'        Set rs1 = oCred.ObtieneQuienExoneraAdmCred
'    Set oCred = Nothing
'    'Call CargaQuienExonera(rs1)  'RECO20140328 ERS174-2013 ANEXO 01 COMENTADO
'
'    'MADM 20120321
'    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
'    Set lrAgenc = loCargaAg.dObtieneAgencias(True)
'    Call llenar_cbo_agencia(lrAgenc, cboagencia)
'
'    Set oCons = New COMDConstantes.DCOMConstantes
'    Set Rs2 = oCons.RecuperaConstantes(9013)
'    Set oCons = Nothing
'    'Call Llenar_Combo_con_Recordset(Rs2, cboAutorizaciones) 'RECO20140328 ERS174-2013 ANEXO 01 COMENTADO
'
'    Set oCred = New COMDCredito.DCOMCreditos
'    Set rs1 = oCred.ObtieneQuienExoneraAdmCred
'    Set oCred = Nothing
'    'Call CargaQuienAutoriza(rs1) 'RECO20140328 ERS174-2013 ANEXO 01 COMENTADO
'
'    'Call CargaUserAutoriza(Right(Trim(Me.cboQuienAuto.Text), 6))  'RECO20140328 ERS174-2013 ANEXO 01 COMENTADO
'    'END MADM
'
'    'Call CargaUserExonera(Right(Trim(Me.cboQuienExone.Text), 6)) 'RECO20140328 ERS174-2013 ANEXO 01 COMENTADO
'
'    ActxCta.NroCuenta = ""
'    ActxCta.CMAC = gsCodCMAC
'    ActxCta.Age = gsCodAge
'
'    ActXCodCta1.NroCuenta = ""
'    ActXCodCta1.CMAC = gsCodCMAC
'    ActXCodCta1.Age = gsCodAge
'
'    ActXCodCta2.NroCuenta = ""
'    ActXCodCta2.CMAC = gsCodCMAC
'    ActXCodCta2.Age = gsCodAge
'
'    LstCred.Clear
'    LstObs.Clear
'    ListaRelacion.ListItems.Clear
'    ListaRelacion1.ListItems.Clear
'
'    'WIOR 20120405
'    If Me.chkClientePref.value = 0 And Me.chkPostDesembolso.value = 0 Then
'        w.TabVisible(3) = False
'    End If
'    fbPostDesembolso = False
'    'WIOR-FIN
'    'RECO20140326 ERS174-2013 ANEXO-01************************************************
'    Call CargarExoneraciones(Me.cboExoneraciones)
'    Call CargarExoneraciones(Me.cboExoneraMat)
'    Call CargarAutorizaciones(Me.cboAutorizacion)
'    Call CargarAutorizaciones(Me.cboAutorizacionMant) 'RECO20141009
'    Call CargarNivelExoneracion(Me.cboNivelExonera, Me.FEApoderados, Me.cboExoneraciones, 1)
'    Call CargarNivelExoneracion(Me.cboNivExoneraMat, Me.FEApoderadoExoneraMat, Me.cboExoneraMat, 2)
'    Call CargarNivelAutorizacion(Me.cboNivAutorizacion, Me.FEApoderadoAuto, Me.cboAutorizacion, 1)
'    Call CargarNivelAutorizacion(Me.cboNivelAutMant, Me.FEApoderadoAutMant, Me.cboAutorizacionMant, 1)    'RECO20141009
'    'RECO FIN*************************************************************************
'End Sub
'
'Private Sub LstCred_Click()
'        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
'            ActxCta.NroCuenta = LstCred.Text
'            ActxCta.SetFocusCuenta
'        End If
'End Sub
'
'Private Sub LstObs_Click()
'    If LstObs.ListCount > 0 And LstObs.ListIndex <> -1 Then
'        ActXCodCta1.NroCuenta = LstObs.Text
'        ActXCodCta1.SetFocusCuenta
'    End If
'End Sub
'
'Private Sub LstObsMant_Click()
'  If LstObsMant.ListCount > 0 And LstObsMant.ListIndex <> -1 Then
'        'ActXCodCta2.NroCuenta = LstObs.Text
'        ActXCodCta2.NroCuenta = LstObsMant.Text
'        ActXCodCta2.SetFocusCuenta
'    End If
'End Sub
'
'Private Sub txtFechaRevision_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{Tab}", True
'    End If
'End Sub
'
'Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
'     KeyAscii = Letras(KeyAscii)
'     If KeyAscii = 13 Then
'        SendKeys "{Tab}", True
'    End If
'End Sub
'
'Function validaDatos() As Boolean
'
'    Dim oCred As COMDCredito.DCOMCreditos
'
'    If Len(ActxCta.NroCuenta) < 18 Then
'        MsgBox "Ingrese un crédito.", vbInformation, "Aviso"
'        validaDatos = False
'        Exit Function
'    End If
'
'    If Len(Trim(lblcodigo(0))) = 0 Then
'        MsgBox "Ingrese un crédito.", vbInformation, "Aviso"
'        validaDatos = False
'        Exit Function
'    End If
'
''    'Verifica que el num de cred no haya sido registrado
''    Set oCred = New COMDCredito.DCOMCreditos
''    If oCred.BuscaRegistroControlAdmCred(ActxCta.NroCuenta) Then
''        If MsgBox("El control de este crédito ya fue registrado, desea Regularizar las observaciones?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
''
''        Else
''            ValidaDatos = False
''            Exit Function
''        End If
''    End If
'
'    'valida fecha revision
'    If ValidaFecha(txtFechaRevision.Text) <> "" Then
'        MsgBox "No se registro fecha de Revisión", vbInformation, "Aviso"
'        validaDatos = False
'        txtFechaRevision.SetFocus
'        Exit Function
'    End If
'
''    '----------verifica observacion
''    nCuentaObs = 0
''    For i = 1 To FlexEdit1.Rows - 1
''        If Len(Me.FlexEdit1.TextMatrix(i, 2)) > 0 Then
''            nCuentaObs = nCuentaObs + 1
''        End If
''    Next i
''    If nCuentaObs = 0 Then
''        MsgBox "Este crédito no tiene observaciones.", vbInformation, "Aviso"
''        ValidaDatos = False
''        txtObservaciones.SetFocus
''        Exit Function
''    End If
'
'    nCuentaExo = 0
'    'For i = 1 To FlexExo.Rows - 1
'    '    If Len(Me.FlexExo.TextMatrix(i, 2)) > 0 Then
'    '        nCuentaExo = nCuentaExo + 1
'    '    End If
'    'Next i
'
'    'nCuentaAuto = 0
'    'For i = 1 To FlexAuto.Rows - 1
'    'If Len(Me.FlexAuto.TextMatrix(i, 2)) > 0 Then
'    '    nCuentaAuto = nCuentaAuto + 1
'    'End If
'    'Next i
'
'    nCuentaPost = 0
'    For i = 1 To flexPost.Rows - 1
'    If Len(Me.flexPost.TextMatrix(i, 2)) > 0 Then
'        nCuentaPost = nCuentaPost + 1
'    End If
'    Next i
'
'    'JUEZ 20140725 ******************************************
'    If FlexObs.GetRsNew Is Nothing And chkPostDesembolso.value = 0 And chkRevisaDesembControl.value = 0 Then
'        MsgBox "Crédito sin observaciones, se procederá a marcar su revisión", vbInformation, "Aviso"
'        chkRevisaDesembControl.value = 1
'    End If
'    'END JUEZ ***********************************************
'
'    validaDatos = True
'End Function
'
'Sub LimpiarControles()
'    Dim i As Integer
'
'    ActxCta.Enabled = True
'    ActxCta.NroCuenta = fgIniciaAxCuentaCF
'    txtFechaRevision = "__/__/____"
'    txtObservaciones = ""
'    'cboExoneraciones.ListIndex = -1
'    'cboQuienExone.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'
'    'cboUsuarioExone.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'    'cboUsuExo2.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'    'cboUsuExo3.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'
'    Me.chkCompraDeuda.value = 0
'    Me.chkClientePref.value = 0
'    Me.ChkConstitucion.value = 0
'
'    txtcompra = ""
'
'    LstCred.Clear
'
'    FlexObs.Clear
'    FlexObs.Rows = 2
'    FlexObs.FormaCabecera
'    FlexObs.FormateaColumnas
'    FlexObs.TextMatrix(1, 0) = "1"
'
'    'FlexExo.Clear 'RECO20140404 ERS174 ANEX 01
'    'FlexExo.Rows = 2'RECO20140404 ERS174 ANEX 01
'    'FlexExo.FormaCabecera'RECO20140404 ERS174 ANEX 01
'    'FlexExo.FormateaColumnas'RECO20140404 ERS174 ANEX 01
'    'FlexExo.TextMatrix(1, 0) = "1"'RECO20140404 ERS174 ANEX 01
'
'    FlexEdit1.Clear
'    FlexEdit1.Rows = 2
'    FlexEdit1.FormaCabecera
'    FlexEdit1.FormateaColumnas
'    FlexEdit1.TextMatrix(1, 0) = "1"
'
'    'MADM 20120319
'    'FlexAuto.Clear'RECO20140404 ERS174 ANEX 01
'    'FlexAuto.Rows = 2'RECO20140404 ERS174 ANEX 01
'    'FlexAuto.FormaCabecera'RECO20140404 ERS174 ANEX 01
'    'FlexAuto.FormateaColumnas'RECO20140404 ERS174 ANEX 01
'    'FlexAuto.TextMatrix(1, 0) = "1"'RECO20140404 ERS174 ANEX 01
'
'    flexPost.Clear
'    flexPost.Rows = 2
'    flexPost.FormaCabecera
'    flexPost.FormateaColumnas
'    flexPost.TextMatrix(1, 0) = "1"
'
'    FEGarantCred.Clear
'    FEGarantCred.Rows = 2
'    FEGarantCred.FormaCabecera
'    FEGarantCred.FormateaColumnas
'    FEGarantCred.TextMatrix(1, 0) = "1"
'
'    ListaRelacion.ListItems.Clear
'    ListaRelacion1.ListItems.Clear
'
'    'cboUsuarioAuto.ListIndex = -1'RECO20140404 ERS174 ANEX 01
'    'cboUsuAuto2.ListIndex = -1'RECO20140404 ERS174 ANEX 01
'    'cboUsuAuto3.ListIndex = -1'RECO20140404 ERS174 ANEX 01
'
'    For i = 0 To 15
'        Me.lblcodigo(i).Caption = ""
'    Next
'
'    'Me.cboAutorizaciones.ListIndex = -1'RECO20140404 ERS174 ANEX 01
'    'cboQuienAuto.ListIndex = -1'RECO20140404 ERS174 ANEX 01
'
'    'cboUsuarioAuto.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'    'cboUsuAuto2.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'    'cboUsuAuto3.ListIndex = -1 'RECO20140404 ERS174 ANEX 01
'
'    'END MADM
'
'    'If Me.cboExoneraDet.Visible Then'RECO20140404 ERS174 ANEX 01
'        'Me.cboExoneraDet.ListIndex = -1'RECO20140404 ERS174 ANEX 01
'        'Me.cboExoneraDet.Enabled = False'RECO20140404 ERS174 ANEX 01
'        'Me.cboExoneraDet.Visible = False'RECO20140404 ERS174 ANEX 01
'    'End If'RECO20140404 ERS174 ANEX 01
'    'If Me.txtExoneraDet.Visible Then'RECO20140404 ERS174 ANEX 01
'        'Me.txtExoneraDet.Text = ""'RECO20140404 ERS174 ANEX 01
'        'Me.txtExoneraDet.Enabled = False'RECO20140404 ERS174 ANEX 01
'        'Me.txtExoneraDet.Visible = False'RECO20140404 ERS174 ANEX 01
'    'End If'RECO20140404 ERS174 ANEX 01
'
'    'WIOR 20120517************************
'    Me.chkClientePref.Enabled = True
'    Me.chkClientePref.value = 0
'    Me.chkPostDesembolso.Enabled = True
'    Me.chkPostDesembolso.value = 0
'    'WIOR FIN*****************************
'    'RECO20140407 ERS174-2014 ANEXO 01*************************************
'    Me.FEExoneraciones.Clear
'    Me.FEExoneraciones.Rows = 2
'    Me.FEExoneraciones.FormaCabecera
'    Me.FEExoneraciones.FormateaColumnas
'
'    Me.FEApoderados.Clear
'    Me.FEApoderados.Rows = 2
'    Me.FEApoderados.FormaCabecera
'    Me.FEApoderados.FormateaColumnas
'
'    Me.FEAutorizaciones.Clear
'    Me.FEAutorizaciones.Rows = 2
'    Me.FEAutorizaciones.FormaCabecera
'    Me.FEAutorizaciones.FormateaColumnas
'
'    Me.FEApoderadoAuto.Clear
'    Me.FEApoderadoAuto.Rows = 2
'    Me.FEApoderadoAuto.FormaCabecera
'    Me.FEApoderadoAuto.FormateaColumnas
'
'    Me.FEExoneraMat.Clear
'    Me.FEExoneraMat.Rows = 2
'    Me.FEExoneraMat.FormaCabecera
'    Me.FEExoneraMat.FormateaColumnas
'
'    Me.FEApoderadoExoneraMat.Clear
'    Me.FEApoderadoExoneraMat.Rows = 2
'    Me.FEApoderadoExoneraMat.FormaCabecera
'    Me.FEApoderadoExoneraMat.FormateaColumnas
'    'RECO FIN**************************************************************
'
'    Me.chkRevisaDesembControl.value = 0 'JUEZ 20140725
'    Me.chkRevisaDesembControl.Enabled = False 'JUEZ 20140725
'End Sub
'
''RECO20140407 ERS174-2014 ANEXO 01*************************************
''Sub CargaQuienExonera(ByVal pRs As ADODB.Recordset)
''
''    Dim sDes As String
''    Dim nCodigo As Integer
''    On Error GoTo ErrHandler
''
''        Do Until pRs.EOF
''            nCodigo = pRs!cRHCargoCod
''            sDes = pRs!cRHCargoDescripcion
''
''            cboQuienExone.AddItem Left(pRs!cRHCargoDescripcion, 50) + Space(50) + pRs!cRHCargoCod
''
''            pRs.MoveNext
''        Loop
''    Exit Sub
''ErrHandler:
''    MsgBox "Error al cargar Quien Exonera", vbInformation, "AVISO"
''End Sub
''Sub CargaQuienAutoriza(ByVal pRs As ADODB.Recordset)
''
''    Dim sDes As String
''    Dim nCodigo As Integer
''    On Error GoTo ErrHandler
''
''        Do Until pRs.EOF
''            nCodigo = pRs!cRHCargoCod
''            sDes = pRs!cRHCargoDescripcion
''
''            'cboQuienAuto.AddItem Left(pRs!cRHCargoDescripcion, 50) + Space(50) + pRs!cRHCargoCod 'RECO20140328 ERS174-2013 ANEXO 01
''            pRs.MoveNext
''        Loop
''    Exit Sub
''ErrHandler:
''    MsgBox "Error al cargar Quien Exonera", vbInformation, "AVISO"
''End Sub
''Sub CargaUserExonera(ByVal psQuienExo As String)
''
''    Dim oGarantia As COMDCredito.DCOMCreditos
''    Dim sDes As String
''    Dim nCodigo As Integer
''    On Error GoTo ErrHandler
''    Dim R As Recordset
''
''        Set oGarantia = New COMDCredito.DCOMCreditos
''           Set R = oGarantia.ObtieneUserExoneraAdmCred(psQuienExo)
''        Set oGarantia = Nothing
''
''        cboUsuarioExone.Clear
''        Me.cboUsuExo2.Clear
''        Me.cboUsuExo3.Clear
''        Do While Not R.EOF
''            cboUsuarioExone.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            cboUsuExo2.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            cboUsuExo3.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            R.MoveNext
''        Loop
''        R.Close
''        Set R = Nothing
''
''    Exit Sub
''ErrHandler:
''    MsgBox "Error al cargar Quien Exonera", vbInformation, "AVISO"
''End Sub
''Sub CargaUserAutoriza(ByVal psQuienExo As String)
''
''    Dim oGarantia As COMDCredito.DCOMCreditos
''    Dim sDes As String
''    Dim nCodigo As Integer
''    On Error GoTo ErrHandler
''    Dim R As Recordset
''
''        Set oGarantia = New COMDCredito.DCOMCreditos
''           Set R = oGarantia.ObtieneUserExoneraAdmCred(psQuienExo)
''        Set oGarantia = Nothing
''
''        cboUsuarioExone.Clear
''        Me.cboUsuExo2.Clear
''        Me.cboUsuExo3.Clear
''        Do While Not R.EOF
''            cboUsuarioAuto.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            cboUsuAuto2.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            cboUsuAuto3.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            R.MoveNext
''        Loop
''        R.Close
''        Set R = Nothing
''
''    Exit Sub
''ErrHandler:
''    MsgBox "Error al cargar Quien Exonera", vbInformation, "AVISO"
''End Sub
''RECO CONENT FIN
''*** PEAC 20101229
'Sub CargaObs(ByVal psCtaCod As String)
'
'Dim oCred As COMDCredito.DCOMCreditos
'Dim rsCred As ADODB.Recordset
'Dim rsComun As ADODB.Recordset
'Dim bCargado As Boolean
'Dim sEstado As String
'Dim bRevisaDesemb As Boolean 'JUEZ 20140725
'
'    Set oCred = New COMDCredito.DCOMCreditos
'    '***********RECO20131112 ERS133*********************
'    'Set rsCred = oCred.BuscaObsAdmCred(psCtaCod)
'    If chkPostDesembolsoBusca.value = 1 Then
'        Set rsCred = oCred.BuscaObsPosAdmCred(psCtaCod)
'    Else
'        Set rsCred = oCred.BuscaObsAdmCred(psCtaCod)
'    End If
'    '***********END RECO********************************
'    Set oCred = Nothing
'
'    If Not (rsCred.EOF And rsCred.BOF) Then
'    'JUEZ 20140725 *****************************************************
'    bRevisaDesemb = IIf(IsNull(rsCred!bRevisaDesemb), False, rsCred!bRevisaDesemb)
'    If chkPostDesembolsoBusca.value = 1 Then
'        chkRevisaDesembObs.value = 1
'    Else
'        chkRevisaDesembObs.value = IIf(bRevisaDesemb, 1, 0)
'    End If
'    'END JUEZ **********************************************************
'    FlexEdit1.Clear
'    FlexEdit1.Rows = 2
'    FlexEdit1.FormaCabecera
'    FlexEdit1.FormateaColumnas
'    FlexEdit1.TextMatrix(1, 0) = "1"
'
'        Do While Not rsCred.EOF
'
'            FlexEdit1.AdicionaFila
'            FlexEdit1.TextMatrix(FlexEdit1.row, 1) = Trim(rsCred!cDescripcion)
'            FlexEdit1.TextMatrix(FlexEdit1.row, 2) = Trim(rsCred!cCtaCod)
'            FlexEdit1.TextMatrix(FlexEdit1.row, 3) = Trim(rsCred!nRegulariza)
'
'            rsCred.MoveNext
'        Loop
'        rsCred.Close
'        Set rsCred = Nothing
'
'    Else
'        MsgBox "No se encontro Observaciones de este crédito.", vbOKOnly + vbExclamation, "Atención"
'        Set rsCred = Nothing
'
'        LimpiarControlesObs
'
'    End If
'
'End Sub
''madm 20110309
''RECO20140407 ERS174-2014 ANEXO 01*************************************
''Sub CargaExo(ByVal psCtaCod As String)
''
''Dim oCred As COMDCredito.DCOMCreditos
''Dim rsCred As ADODB.Recordset
''Dim rsComun As ADODB.Recordset
''Dim bCargado As Boolean
''Dim sEstado As String
''
''    Set oCred = New COMDCredito.DCOMCreditos
''    Set rsCred = oCred.BuscaExoAdmCred(psCtaCod)
''    Set oCred = Nothing
''
''    If Not (rsCred.EOF And rsCred.BOF) Then
''            FlexExoMant.Clear
''            FlexExoMant.Rows = 2
''            FlexExoMant.FormaCabecera
''            FlexExoMant.FormateaColumnas
''            FlexExoMant.TextMatrix(1, 0) = "1"
''        Do While Not rsCred.EOF
''
''            FlexExoMant.AdicionaFila
''            FlexExoMant.TextMatrix(FlexExoMant.row, 1) = Trim(rsCred!cDesExonera)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 2) = Trim(rsCred!cCodQuienExok)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 3) = Trim(rsCred!cUsu1)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 4) = Trim(rsCred!cUsu2)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 5) = Trim(rsCred!cUsu3)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 6) = Trim(rsCred!nCodExonera)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 7) = Trim(rsCred!cCodQuienExo)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 8) = Trim(rsCred!nTipoCAR)
''            FlexExoMant.TextMatrix(FlexExoMant.row, 9) = Trim(rsCred!CDescripcionOtro)
''            rsCred.MoveNext
''        Loop
''        rsCred.Close
''        Set rsCred = Nothing
''
'''select nCodExonera,cCodQuienExo,nTipoCAR,CDescripcionOtro
'''K.cConsDescripcion cDesExonera,K1.cConsDescripcion cCodQuienExok
''
''    Else
''        LimpiaFlex FlexExoMant 'WIOR 20130121
''        MsgBox "No se encontro Exoneraciones de este crédito.", vbOKOnly + vbExclamation, "Atención"
''        Set rsCred = Nothing
'''        LimpiarControlesExo
''    End If
''
''End Sub
'
'Sub CargaObsMant(ByVal psCtaCod As String)
'
'Dim oCred As COMDCredito.DCOMCreditos
'Dim rsCred As ADODB.Recordset
'Dim rsComun As ADODB.Recordset
'Dim bCargado As Boolean
'Dim sEstado As String
'
'    Set oCred = New COMDCredito.DCOMCreditos
'    Set rsCred = oCred.BuscaObsAdmCred(psCtaCod)
'    Set oCred = Nothing
'
'    If Not (rsCred.EOF And rsCred.BOF) Then
'    FlexObsMant.Clear
'    FlexObsMant.Rows = 2
'    FlexObsMant.FormaCabecera
'    FlexObsMant.FormateaColumnas
'    FlexObsMant.TextMatrix(1, 0) = "1"
'        Do While Not rsCred.EOF
'            FlexObsMant.AdicionaFila
'            FlexObsMant.TextMatrix(FlexObsMant.row, 1) = Trim(rsCred!cDescripcion)
'            FlexObsMant.TextMatrix(FlexObsMant.row, 2) = Trim(rsCred!cCtaCod)
'
'            rsCred.MoveNext
'        Loop
'        rsCred.Close
'        Set rsCred = Nothing
'
'    Else
'        LimpiaFlex FlexObsMant 'WIOR 20130121
'        MsgBox "No se encontro Observaciones de este crédito.", vbOKOnly + vbExclamation, "Atención"
'        Set rsCred = Nothing
''        LimpiarControlesObsMant
'    End If
'End Sub
''RECO20140407 ERS174-2014 ANEXO 01*************************************
''Sub CargaQuienExoneraMant(ByVal pRs As ADODB.Recordset)
''
''    Dim sDes As String
''    Dim nCodigo As Integer
''    On Error GoTo ErrHandler
''
''        Do Until pRs.EOF
''            nCodigo = pRs!cRHCargoCod
''            sDes = pRs!cRHCargoDescripcion
''
''            'cboQuienExoneMant.AddItem Left(pRs!cRHCargoDescripcion, 50) + Space(50) + pRs!cRHCargoCod
''
''            pRs.MoveNext
''        Loop
''    Exit Sub
''ErrHandler:
''    MsgBox "Error al cargar Quien Exonera", vbInformation, "AVISO"
''End Sub
''Sub CargaUserExoneraMant(ByVal psQuienExo As String)
''
''    Dim oGarantia As COMDCredito.DCOMCreditos
''    Dim sDes As String
''    Dim nCodigo As Integer
''    On Error GoTo ErrHandler
''    Dim R As Recordset
''
''        Set oGarantia = New COMDCredito.DCOMCreditos
''           Set R = oGarantia.ObtieneUserExoneraAdmCred(psQuienExo)
''        Set oGarantia = Nothing
''
''        cboUsuarioExoneMant.Clear
''        Me.cboUsuExo2Mant.Clear
''        Me.cboUsuExo3Mant.Clear
''        Do While Not R.EOF
''            cboUsuarioExoneMant.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            cboUsuExo2Mant.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            cboUsuExo3Mant.AddItem UCase(R!cUser) & Space(2) & R!cPersNombre
''            R.MoveNext
''        Loop
''        R.Close
''        Set R = Nothing
''
''    Exit Sub
''ErrHandler:
''    MsgBox "Error al cargar Quien Exonera", vbInformation, "AVISO"
''End Sub
'
'Sub LimpiarControlesMant()
'    Dim i As Integer
'
'    ActXCodCta2.Enabled = True
'    ActXCodCta2.NroCuenta = fgIniciaAxCuentaCF
'    'cboExoneracionesMant.ListIndex = -1
'    'cboQuienExoneMant.ListIndex = -1
'
'    'cboUsuarioExoneMant.ListIndex = -1
'    'cboUsuExo2Mant.ListIndex = -1
'    'cboUsuExo3Mant.ListIndex = -1
'
'    LstObsMant.Clear
'
'    FlexObsMant.Clear
'    FlexObsMant.Rows = 2
'    FlexObsMant.FormaCabecera
'    FlexObsMant.FormateaColumnas
'    FlexObsMant.TextMatrix(1, 0) = "1"
'
'    'FlexExoMant.Clear
'    'FlexExoMant.Rows = 2
'    'FlexExoMant.FormaCabecera
'    'FlexExoMant.FormateaColumnas
'    'FlexExoMant.TextMatrix(1, 0) = "1"
'
'    'If Me.cboExoneraDetMant.Visible Then
'    '    Me.cboExoneraDetMant.ListIndex = -1
'    '    Me.cboExoneraDetMant.Enabled = False
'    '    Me.cboExoneraDetMant.Visible = False
'    'End If
'
'    'If Me.txtExoneraDetMant.Visible Then
'    '    Me.txtExoneraDetMant.Text = ""
'    '    Me.txtExoneraDetMant.Enabled = False
'    '    Me.txtExoneraDetMant.Visible = False
'    'End If
'    FEExoneraMat.Clear 'RECO20140804
'    FormateaFlex FEExoneraMat 'RECO20140804
'    FEApoderadoExoneraMat.Clear 'RECO20140804
'    FormateaFlex FEApoderadoExoneraMat 'RECO20140804
'    FEApoderadoAutMant.Clear 'RECO20141009
'    FEAutorizacionesMant.Clear 'RECO20141009
'    FormateaFlex FEApoderadoAutMant 'RECO20141009
'    FormateaFlex FEAutorizacionesMant 'RECO20141009
'End Sub
''end madm
'
''*** PEAC 20101229
'Function ValidaDatosObs() As Boolean
'
'Dim oCred As COMDCredito.DCOMCreditos
'
'    If Len(ActXCodCta1.NroCuenta) < 18 Then
'        MsgBox "Ingrese un crédito.", vbInformation, "Aviso"
'        ValidaDatosObs = False
'        Exit Function
'    End If
'
'    '----------verifica observacion
'    nCuentaObs = 0
'    For i = 1 To FlexEdit1.Rows - 1
'        If Len(Me.FlexEdit1.TextMatrix(i, 2)) > 0 Then
'            nCuentaObs = nCuentaObs + 1
'        End If
'    Next i
'    If nCuentaObs = 0 Then
'        MsgBox "Este crédito no tiene observaciones.", vbInformation, "Aviso"
'        ValidaDatosObs = False
'        Exit Function
'    End If
'
'    ValidaDatosObs = True
'End Function
'
''*** PEAC 20101229
'Sub LimpiarControlesObs()
'    Dim i As Integer
'
'    ActXCodCta1.Enabled = True
'    ActXCodCta1.NroCuenta = fgIniciaAxCuentaCF
'
'    LstObs.Clear
'
'    FlexEdit1.Clear
'    FlexEdit1.Rows = 2
'    FlexEdit1.FormaCabecera
'    FlexEdit1.FormateaColumnas
'    FlexEdit1.TextMatrix(1, 0) = "1"
'
'    chkPostDesembolsoBusca.value = 0
'
'    chkRevisaDesembObs.value = 0 'JUEZ 20140725
'End Sub
''madm 20110309
'Sub LimpiarControlesExo()
'    Dim i As Integer
'
'    ActXCodCta2.Enabled = True
'    ActXCodCta2.NroCuenta = fgIniciaAxCuentaCF
'
'    LstObsMant.Clear
'
''    FlexExoMant.Clear RECO20140328 ERS174-2013 ANEXO 01
''    FlexExoMant.Rows = 2 RECO20140328 ERS174-2013 ANEXO 01
''    FlexExoMant.FormaCabecera RECO20140328 ERS174-2013 ANEXO 01
''    FlexExoMant.FormateaColumnas RECO20140328 ERS174-2013 ANEXO 01
''    FlexExoMant.TextMatrix(1, 0) = "1" RECO20140328 ERS174-2013 ANEXO 01
'
'
'     sPersCargoExo = "" 'RECO20140328 ERS174-2013 ANEXO 01
'     sPersCargoAuto = "" 'RECO20140328 ERS174-2013 ANEXO 01
'     sPersCargoExoMant = "" 'RECO20140328 ERS174-2013 ANEXO 01
'End Sub
'Function ValidaDatosMant() As Boolean
'
'Dim oCred As COMDCredito.DCOMCreditos
'
'    If Len(ActXCodCta2.NroCuenta) < 18 Then
'        MsgBox "Ingrese un crédito.", vbInformation, "Aviso"
'        ValidaDatosMant = False
'        Exit Function
'    End If
'
'    nCuentaExoMant = 0
'    'For i = 1 To FlexExoMant.Rows - 1
'    '    If Len(Me.FlexExoMant.TextMatrix(i, 2)) > 0 Then
'    '        nCuentaExoMant = nCuentaExoMant + 1
'    '    End If
'    'Next i
'
'    ValidaDatosMant = True
'End Function
''MADM 20110309
'Sub LimpiarControlesObsMant()
'    Dim i As Integer
'
'    ActXCodCta2.Enabled = True
'    ActXCodCta2.NroCuenta = fgIniciaAxCuentaCF
'
'    LstObsMant.Clear
'
'    FlexObsMant.Clear
'    FlexObsMant.Rows = 2
'    FlexObsMant.FormaCabecera
'    FlexObsMant.FormateaColumnas
'    FlexObsMant.TextMatrix(1, 0) = "1"
'
'End Sub
'
'Sub CargarFlexGarantia(ByVal psCtaCod As String)
'Dim RGar As ADODB.Recordset
'Dim oDGarantia As COMDCredito.DCOMGarantia 'MADM 20110505
'Dim rsGarantReal As ADODB.Recordset 'MADM 20110505
'Dim oNCredito As COMNCredito.NCOMCredito
'Set oNCredito = New COMNCredito.NCOMCredito
'
'Set RGar = oNCredito.obtenerGarantxCredito(psCtaCod)
'
' Call LimpiaFlex(FEGarantCred)
'            'Set RGar = oCredito.RecuperaGarantiasCredito(ActxCta.NroCuenta)
'            Do While Not RGar.EOF
'                    FEGarantCred.AdicionaFila
'                    FEGarantCred.RowHeight(RGar.Bookmark) = 280
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 1) = RGar!cTpoGarDescripcion
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 2) = Format(RGar!nGravado, "#,#0.00")
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 3) = Format(RGar!nTasacion, "#,#0.00")
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 4) = Format(RGar!nRealizacion, "#,#0.00")
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 5) = Format(RGar!nPorGravar, "#,#0.00")
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 6) = Trim(RGar!cPersNombre)
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 7) = Trim(RGar!cNroDoc)
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 8) = Trim(RGar!cTpoDoc)
'                    FEGarantCred.TextMatrix(RGar.Bookmark, 9) = Trim(RGar!cNumGarant)
'
'                    'MADM 20110506 * Num Garantia
'                     Set oDGarantia = New COMDCredito.DCOMGarantia
'                     Set rsGarantReal = oDGarantia.RecuperaGarantiaRealMaxAprobacion(ActxCta.NroCuenta, RGar!cNumGarant)
'                     Set oDGarantia = Nothing
'
'                    If Not (rsGarantReal.EOF Or rsGarantReal.BOF) Then
'                        'FEGarantCred.TextMatrix(rsGarantReal.Bookmark, 10) = Trim(rsGarantReal!nApruebaLegal)
'                        If (rsGarantReal!cNumGarant <> "") And ((DateDiff("d", rsGarantReal!dCertifGravamen, gdFecSis) > 365) Or (DateDiff("d", rsGarantReal!dTasacion, gdFecSis) > 730)) Then
'                             Select Case rsGarantReal!nVerificaLegal
'                                Case 1
'                                    FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Pendiente"
'                                Case 2
'                                    FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Aprobado"
'                                Case 3
'                                    FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Desaprobado"
'                                Case 4
'                                    FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Pendiente por Regularizar"
'                                Case 0
'                                    FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Pendiente"
'                            End Select
'                            FEGarantCred.TextMatrix(RGar.Bookmark, 11) = Trim(IIf(Len(Trim(rsGarantReal!cNumPoliza)) > 1, IIf(Len(Trim(rsGarantReal!nEstadoPolizaNew)) > 1, rsGarantReal!nEstadoPolizaNew, "No registrado"), "No tiene"))
'                        Else
'                            FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "Conforme"
'                            'FEGarantCred.TextMatrix(Rgar.Bookmark, 11) = "No tiene"
'                            FEGarantCred.TextMatrix(RGar.Bookmark, 11) = Trim(IIf(Len(Trim(rsGarantReal!cNumPoliza)) > 1, IIf(Len(Trim(rsGarantReal!nEstadoPolizaNew)) > 1, rsGarantReal!nEstadoPolizaNew, "No registrado"), "No tiene"))
'                        End If
'                        'FEGarantCred.TextMatrix(Rgar.Bookmark, 11) = Trim(rsGarantReal!nEstadoPolizaNew)
'                    Else
'                        FEGarantCred.TextMatrix(RGar.Bookmark, 10) = "No aplica"
'                        FEGarantCred.TextMatrix(RGar.Bookmark, 11) = "No tiene"
'                    End If
'
'                    'END MADM
'                    RGar.MoveNext
'            Loop
'End Sub
'
'Private Sub CargaPersonasRelacCred(ByVal psCtaCod As String, ByVal prsRelac As ADODB.Recordset)
'Dim s As ListItem
'Dim oRelPersCred As UCredRelac_Cli
'    Set oRelPersCred = New UCredRelac_Cli
'
'    On Error GoTo ErrorCargaPersonasRelacCred
'    ListaRelacion.ListItems.Clear
'    Call oRelPersCred.CargaRelacPersCred(psCtaCod, prsRelac)
'    oRelPersCred.IniciarMatriz
'    Do While Not oRelPersCred.EOF
'        Set s = ListaRelacion.ListItems.Add(, , oRelPersCred.ObtenerNombre)
'        s.SubItems(1) = oRelPersCred.ObtenerRelac
'        s.SubItems(2) = oRelPersCred.ObtenerCodigo
'        s.SubItems(3) = oRelPersCred.ObtenerValorRelac
'        oRelPersCred.siguiente
'    Loop
'
'    Exit Sub
'
'ErrorCargaPersonasRelacCred:
'        MsgBox err.Description, vbCritical, "Aviso"
'End Sub
'
'Private Sub CargaPersonasRelacCredRefinan(ByVal psCtaCod As String)
'Dim s As ListItem
'Dim oCred As COMDCredito.DCOMCreditos
'Set oCred = New COMDCredito.DCOMCreditos
'Dim rsRef As ADODB.Recordset
'
'    On Error GoTo ErrorRecuperaAdmCredRefinan
'    ListaRelacion1.ListItems.Clear
'    Set rsRef = oCred.RecuperaAdmCredRefinan(psCtaCod)
'    Set oCred = Nothing
'
'    If Not rsRef.EOF Then
'        Do While Not rsRef.EOF
'            Set s = ListaRelacion1.ListItems.Add(, , rsRef!cCtaCodRef)
'            s.SubItems(1) = rsRef!nMontoRef
'            s.SubItems(2) = rsRef!nCuotas
'            rsRef.MoveNext
'        Loop
'   rsRef.Close
'   Set rsRef = Nothing
'   End If
'   Exit Sub
'ErrorRecuperaAdmCredRefinan:
'        MsgBox err.Description, vbCritical, "Aviso"
'End Sub
'
'Sub llenar_cbo_agencia(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
'pcboObjeto.Clear
'Do While Not pRs.EOF
'    pcboObjeto.AddItem Trim(pRs!cAgeDescripcion) & Space(100) & Trim(str(pRs!cAgeCod))
'    pRs.MoveNext
'Loop
'pRs.Close
'End Sub
'
''RECO20140314 ERS-174 2013 ANEXO-01******************************************************
'Public Sub CargarExoneraciones(ByVal cboControl As ComboBox)
'    Dim loExonera As New COMDCredito.DCOMNivelAprobacion
'    Dim lrDatos As ADODB.Recordset
'
'    Set loExonera = New COMDCredito.DCOMNivelAprobacion
'    Set lrDatos = New ADODB.Recordset
'    Set lrDatos = loExonera.RecuperaTpoExoneraciones
'    If Not (lrDatos.BOF And lrDatos.EOF) Then
'        Dim nIndex As Integer
'        For nIndex = 1 To lrDatos.RecordCount
'            cboControl.AddItem Trim(lrDatos!cExoneraDesc) & Space(100) & Trim((lrDatos!cExoneraCod))
'            lrDatos.MoveNext
'        Next
'    End If
'    'cboControl.ListIndex = 0
'    Set lrDatos = Nothing
'End Sub
'Public Sub CargarAutorizaciones(ByVal cboControl As ComboBox)
'    Dim loExonera As New COMDCredito.DCOMNivelAprobacion
'    Dim lrDatos As ADODB.Recordset
'
'    Set loExonera = New COMDCredito.DCOMNivelAprobacion
'    Set lrDatos = New ADODB.Recordset
'    Set lrDatos = loExonera.RecuperaTpoAutorizaciones
'
'    If Not (lrDatos.BOF And lrDatos.EOF) Then
'        Dim nIndex As Integer
'        For nIndex = 1 To lrDatos.RecordCount
'            cboControl.AddItem Trim(lrDatos!cExoneraDesc) & Space(100) & Trim((lrDatos!cExoneraCod))
'            lrDatos.MoveNext
'        Next
'    End If
'    'cboAutorizacion.ListIndex = 0
'    Set lrDatos = Nothing
'End Sub
'Public Sub CargarNivelExoneracion(ByVal cboControlNiv As ComboBox, ByVal FEControl As FlexEdit, ByVal cboControlExo As ComboBox, ByVal nTpo As Integer)
'    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
'    Dim lrDatos As ADODB.Recordset
'    Dim lnCanFirmas As Integer
'
'    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
'    Set lrDatos = New ADODB.Recordset
'
'    Set lrDatos = oDNiv.RecuperaNivelesExoneracion(Trim(Right(cboControlExo.Text, 8)))
'
'    If Not (lrDatos.EOF And lrDatos.BOF) Then
'        Dim nIndex As Integer
'        cboControlNiv.Clear
'        If nTpo = 1 Then
'            nCanFirmasExo = lrDatos!nNumCantFirmas
'            lnCanFirmas = nCanFirmasExo
'        Else
'            nCanFirmasExoMant = lrDatos!nNumCantFirmas
'            lnCanFirmas = nCanFirmasExoMant
'        End If
'
'        For nIndex = 1 To lrDatos.RecordCount
'            cboControlNiv.AddItem Trim(lrDatos!cNivAprDesc) & Space(100) & Trim((lrDatos!cNivAprCod))
'            lrDatos.MoveNext
'        Next
'
'        For nIndex = 1 To lnCanFirmas
'            FEControl.AdicionaFila
'            FEControl.TextMatrix(nIndex, 1) = "Apoderado " & nIndex
'        Next
'    End If
'    'cboControlNiv.ListIndex = 0
'End Sub
'Public Sub CargarNivelAutorizacion(ByVal cboControlNiv As ComboBox, ByVal FEControl As FlexEdit, ByVal cboControlAut As ComboBox, ByVal nTpo As Integer)
'    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
'    Dim lrDatos As ADODB.Recordset
'    Dim lnCanFirmas As Integer
'
'    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
'    Set lrDatos = New ADODB.Recordset
'
'    Set lrDatos = oDNiv.RecuperaNivelesExoneracion(Trim(Right(cboControlAut.Text, 8)))
'
'    If Not (lrDatos.EOF And lrDatos.BOF) Then
'        Dim nIndex As Integer
'        cboControlNiv.Clear
'
'        If nTpo = 1 Then
'            nCanFirmas = lrDatos!nNumCantFirmas
'            lnCanFirmas = nCanFirmas
'        Else
'            nCanFirmasAutMant = lrDatos!nNumCantFirmas
'            lnCanFirmas = nCanFirmasAutMant
'        End If
'
'        'nCanFirmas = lrDatos!nNumCantFirmas
'        For nIndex = 1 To lrDatos.RecordCount
'            cboControlNiv.AddItem Trim(lrDatos!cNivAprDesc) & Space(100) & Trim((lrDatos!cNivAprCod))
'            lrDatos.MoveNext
'        Next
'
'        For nIndex = 1 To lnCanFirmas
'            FEControl.AdicionaFila
'            FEControl.TextMatrix(nIndex, 1) = "Apoderado " & nIndex
'        Next
'    End If
'    Set lrDatos = Nothing
'    'cboNivAutorizacion.ListIndex = 0
'End Sub
'Public Sub CargarCargoNivExoAuto(ByVal cboControl As ComboBox, ByVal FEControl As FlexEdit, ByVal cboControlExo As ComboBox, ByVal nTpoOpe As Integer)
'    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
'    Dim lrCargo As ADODB.Recordset
'    Dim lrDatos As ADODB.Recordset
'    Dim lrDRHCargo As COMDPersona.UCOMAcceso
'    Dim lrDatosRRHH  As ADODB.Recordset
'    Dim lsCadCargo As String
'    Dim lnFirmas As Integer
'
'    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
'    Set lrCargo = New ADODB.Recordset
'    Set lrDatos = New ADODB.Recordset
'    Set lrDRHCargo = New COMDPersona.UCOMAcceso
'    Set lrDatosRRHH = New ADODB.Recordset
'    Set lrCargo = oDNiv.RecuperaNivelesAprValores(Trim(Right(cboControl.Text, 8)))
'    Set lrDatos = oDNiv.RecuperaCargosNivelesExoneracion(Trim(Right(cboControlExo.Text, 8)), Trim(Right(cboControl.Text, 8)))
'
'    If Not (lrCargo.EOF And lrCargo.BOF) Then
'        Dim i As Integer
'        For i = 1 To lrCargo.RecordCount
'            lsCadCargo = lsCadCargo & lrCargo!cValorCod & ","
'            lrCargo.MoveNext
'        Next
'    End If
'
'    If Not (lrDatos.EOF And lrDatos.BOF) Then
'        If nTpoOpe = 1 Then
'            nCanFirmasExo = lrDatos!nNumCantFirmas
'            lnFirmas = nCanFirmasExo
'        ElseIf nTpoOpe = 2 Then
'            nCanFirmas = lrDatos!nNumCantFirmas
'            lnFirmas = nCanFirmas
'        ElseIf nTpoOpe = 3 Then
'            nCanFirmasExoMant = lrDatos!nNumCantFirmas
'            lnFirmas = nCanFirmasExoMant
'        Else
'            nCanFirmasAutMant = lrDatos!nNumCantFirmas
'            lnFirmas = nCanFirmasAutMant
'        End If
'    End If
'    FEControl.Clear
'    FormateaFlex FEControl
'
'    Dim nIndex As Integer
'
'    For nIndex = 1 To lnFirmas
'            FEControl.AdicionaFila
'            FEControl.TextMatrix(nIndex, 1) = "Apoderado " & nIndex
'    Next
'
'    Set lrDatosRRHH = lrDRHCargo.RecuperaPersCargo(lsCadCargo)
'    FEControl.CargaCombo lrDatosRRHH
'    Set lrDatosRRHH = Nothing
'    Set lrCargo = Nothing
'End Sub
'
'Public Function LlenarRsExonera(ByVal FEControl As FlexEdit) As ADODB.Recordset
'    Dim rsExoneraciones As New ADODB.Recordset
'    Dim nIndex As Integer
'    Set rsExoneraciones = New ADODB.Recordset
'
'    If FEControl.Rows >= 2 Then
'        If FEControl.TextMatrix(nIndex, 1) = "" Then
'                Exit Function
'        End If
'    rsExoneraciones.CursorType = adOpenStatic
'    rsExoneraciones.Fields.Append "cCtaCod", adVarChar, 18, adFldIsNullable
'    rsExoneraciones.Fields.Append "nCodExonera", adVarChar, 7, adFldIsNullable
'    rsExoneraciones.Fields.Append "cCodQuienExo", adVarChar, 7, adFldIsNullable
'    rsExoneraciones.Fields.Append "Apoder1", adVarChar, 4, adFldIsNullable
'    rsExoneraciones.Fields.Append "Apoder2", adVarChar, 4, adFldIsNullable
'    rsExoneraciones.Fields.Append "Apoder3", adVarChar, 4, adFldIsNullable
'    rsExoneraciones.Fields.Append "nDesExoneraCAR", adInteger, 4, adFldIsNullable
'    rsExoneraciones.Fields.Append "cDesExoneraOtros", adVarChar, 100, adFldIsNullable
'    rsExoneraciones.Open
'
'
'        For nIndex = 1 To FEControl.Rows - 1
'            'FEExoneraciones.TextMatrix(FEExoneraciones.Rows - 1, 4) = sPersCargoExo
'            rsExoneraciones.AddNew
'            rsExoneraciones.Fields("cCtaCod") = ActxCta.NroCuenta
'            rsExoneraciones.Fields("nCodExonera") = Right(FEControl.TextMatrix(nIndex, 1), 7)
'            rsExoneraciones.Fields("cCodQuienExo") = Right(FEControl.TextMatrix(nIndex, 2), 7)
'            rsExoneraciones.Fields("Apoder1") = Right(FEControl.TextMatrix(nIndex, 3), 4)
'            rsExoneraciones.Fields("Apoder2") = Right(FEControl.TextMatrix(nIndex, 4), 4)
'            rsExoneraciones.Fields("Apoder3") = Right(FEControl.TextMatrix(nIndex, 5), 4)
'            rsExoneraciones.Fields("nDesExoneraCAR") = 0
'            rsExoneraciones.Fields("cDesExoneraOtros") = ""
'            rsExoneraciones.Update
'            rsExoneraciones.MoveFirst
'        Next
'    End If
'    Set LlenarRsExonera = rsExoneraciones
'    'rsExoneraciones.Close
'End Function
'Public Function LlenarRsAuto(ByVal FEControl As FlexEdit) As ADODB.Recordset
'    Dim rsAutorizaciones As New ADODB.Recordset
'    Dim nIndex As Integer
'    Set rsAutorizaciones = New ADODB.Recordset
'
'    If FEControl.Rows >= 2 Then
'        If FEControl.TextMatrix(nIndex, 1) = "" Then
'                Exit Function
'        End If
'        rsAutorizaciones.CursorType = adOpenStatic
'        rsAutorizaciones.Fields.Append "cCtaCod", adVarChar, 18, adFldIsNullable
'        rsAutorizaciones.Fields.Append "nCodAutoriza", adVarChar, 7, adFldIsNullable
'        rsAutorizaciones.Fields.Append "cCodQuienAuto", adVarChar, 7, adFldIsNullable
'        rsAutorizaciones.Fields.Append "Apoder1", adVarChar, 4, adFldIsNullable
'        rsAutorizaciones.Fields.Append "Apoder2", adVarChar, 4, adFldIsNullable
'        rsAutorizaciones.Fields.Append "Apoder3", adVarChar, 4, adFldIsNullable
'        rsAutorizaciones.Fields.Append "AutoOtro", adVarChar, 100, adFldIsNullable
'        rsAutorizaciones.Open
'        For nIndex = 1 To FEControl.Rows - 1
'            rsAutorizaciones.AddNew
'            rsAutorizaciones.Fields("cCtaCod") = ActxCta.NroCuenta
'            rsAutorizaciones.Fields("nCodAutoriza") = Right(FEControl.TextMatrix(nIndex, 1), 7)
'            rsAutorizaciones.Fields("cCodQuienAuto") = Right(FEControl.TextMatrix(nIndex, 2), 7)
'            rsAutorizaciones.Fields("Apoder1") = Right(FEControl.TextMatrix(nIndex, 3), 4)
'            rsAutorizaciones.Fields("Apoder2") = Right(FEControl.TextMatrix(nIndex, 4), 4)
'            rsAutorizaciones.Fields("Apoder3") = Right(FEControl.TextMatrix(nIndex, 5), 4)
'            rsAutorizaciones.Fields("AutoOtro") = ""
'            rsAutorizaciones.Update
'            rsAutorizaciones.MoveFirst
'        Next
'    End If
'    Set LlenarRsAuto = rsAutorizaciones
'End Function
'Public Function LlenarRsExoneraMat(ByVal FEControl As FlexEdit) As ADODB.Recordset
'    Dim rsExoneracionesMat As New ADODB.Recordset
'     Dim nIndex As Integer
'    Set rsExoneracionesMat = New ADODB.Recordset
'
'    If FEControl.Rows >= 2 Then
'        If FEControl.TextMatrix(nIndex, 1) = "" Then
'            Exit Function
'        End If
'
'    rsExoneracionesMat.CursorType = adOpenStatic
'    'rsExoneracionesMat.Fields.Append "cCtaCod", adVarChar, 18, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "nCodExonera", adVarChar, 7, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "cCodQuienExo", adVarChar, 7, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "Apoder1", adVarChar, 4, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "Apoder2", adVarChar, 4, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "Apoder3", adVarChar, 4, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "nTipoExoneraCAR", adInteger, 4, adFldIsNullable
'    rsExoneracionesMat.Fields.Append "cDesExoneraOtros", adVarChar, 100, adFldIsNullable
'    rsExoneracionesMat.Open
'
'        For nIndex = 1 To FEControl.Rows - 1
'            rsExoneracionesMat.AddNew
'            'rsExoneracionesMat.Fields("psCtaCod") = ActxCta.NroCuenta
'            rsExoneracionesMat.Fields("nCodExonera") = Right(FEControl.TextMatrix(nIndex, 1), 7)
'            rsExoneracionesMat.Fields("cCodQuienExo") = Right(FEControl.TextMatrix(nIndex, 2), 7)
'            rsExoneracionesMat.Fields("Apoder1") = Right(FEControl.TextMatrix(nIndex, 3), 4)
'            rsExoneracionesMat.Fields("Apoder2") = Right(FEControl.TextMatrix(nIndex, 4), 4)
'            rsExoneracionesMat.Fields("Apoder3") = Right(FEControl.TextMatrix(nIndex, 5), 4)
'            rsExoneracionesMat.Fields("nTipoExoneraCAR") = 0
'            rsExoneracionesMat.Fields("cDesExoneraOtros") = ""
'            rsExoneracionesMat.Update
'            rsExoneracionesMat.MoveFirst
'        Next
'    End If
'
'    Set LlenarRsExoneraMat = rsExoneracionesMat
'    'rsExoneracionesMat.Close
'End Function
'
'Public Function ValidaDatosFE(ByVal FEControl As FlexEdit) As Boolean
'    Dim nIndex As Integer
'
'    For nIndex = 1 To FEControl.Rows - 1
'        If FEControl.TextMatrix(nIndex, 2) = "" Then
'            ValidaDatosFE = False
'            Exit Function
'        End If
'    Next
'    ValidaDatosFE = True
'End Function
'Public Sub LimpiarFE(ByVal FEControl As FlexEdit)
'    FEControl.Clear
'    FormateaFlex FEControl
'End Sub
''RECO FIN********************************************************************************
''RECO20141009***********************************************************
'Public Sub CargarExoneracionesMant()
'    Dim oCred As New COMDCredito.DCOMCreditos
'    Dim rs As New ADODB.Recordset
'    Dim x As Integer
'
'    Set rs = oCred.DevulveExoneracionesCredControlAdm(ActXCodCta2.NroCuenta)
'    FEExoneraMat.Clear
'    FormateaFlex FEExoneraMat
'
'    If Not (rs.EOF And rs.BOF) Then
'        For x = 1 To rs.RecordCount
'            FEExoneraMat.AdicionaFila
'            FEExoneraMat.TextMatrix(x, 1) = rs!cExoneraCod
'            FEExoneraMat.TextMatrix(x, 2) = rs!cNivAprCod
'            FEExoneraMat.TextMatrix(x, 3) = rs!Apoderado1
'            FEExoneraMat.TextMatrix(x, 4) = rs!Apoderado2
'            FEExoneraMat.TextMatrix(x, 5) = rs!Apoderado3
'            rs.MoveNext
'        Next
'    End If
'End Sub
'Public Sub CargarAutorizacionesMant()
'    Dim oCred As New COMDCredito.DCOMCreditos
'    Dim rs As New ADODB.Recordset
'    Dim x As Integer
'
'    Set rs = oCred.DevulveAutorizacionesCredControlAdm(ActXCodCta2.NroCuenta)
'    FEAutorizacionesMant.Clear
'    FormateaFlex FEAutorizacionesMant
'
'    If Not (rs.EOF And rs.BOF) Then
'        For x = 1 To rs.RecordCount
'            FEAutorizacionesMant.AdicionaFila
'            FEAutorizacionesMant.TextMatrix(x, 1) = rs!cAutorizaCod
'            FEAutorizacionesMant.TextMatrix(x, 2) = rs!cNivAprCod
'            FEAutorizacionesMant.TextMatrix(x, 3) = rs!Apoderado1
'            FEAutorizacionesMant.TextMatrix(x, 4) = rs!Apoderado2
'            FEAutorizacionesMant.TextMatrix(x, 5) = rs!Apoderado3
'            rs.MoveNext
'        Next
'    End If
'End Sub
''RECO FIN***************************************************************
''RECO20150316***********************************************************
'Public Function ValidaAgenciaAutorizada(ByVal psCtaCodAge As String) As Boolean
'    Dim objCred As New COMDCredito.DCOMCreditos
'    Dim rs As New ADODB.Recordset
'    Dim lsCadAge As String
'    Dim i As Integer, j As Integer
'
'    Set rs = objCred.ObtieneCredAdmAgeCodAutirza(gsCodAge)
'    If Not (rs.EOF And rs.BOF) Then
'        Dim lsAgeCod As String
'        lsCadAge = rs!cAgeCadAuto
'        For j = 1 To Len(lsCadAge)
'            If Mid(lsCadAge, j, 1) <> "," Then
'                lsAgeCod = lsAgeCod & Mid(lsCadAge, j, 1)
'            Else
'                If psCtaCodAge = lsAgeCod Then
'                    ValidaAgenciaAutorizada = True
'                    Exit Function
'                End If
'                lsAgeCod = ""
'            End If
'        Next
'    End If
'End Function
''RECO FIN***************************************************************


