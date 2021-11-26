VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCajaGenMantCtasHC 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frmCajaGenMantCtasHC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   5010
      TabIndex        =   74
      Top             =   5460
      Width           =   1260
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6330
      TabIndex        =   67
      Top             =   5460
      Width           =   1260
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
      Height          =   1515
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   7620
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1065
         Width           =   2280
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   5010
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1110
         Width           =   1410
      End
      Begin VB.TextBox txtNroCtaIF 
         Height          =   315
         Left            =   4995
         TabIndex        =   1
         Top             =   285
         Width           =   2520
      End
      Begin Sicmact.TxtBuscar txtBuscarCtaIF 
         Height          =   315
         Left            =   1020
         TabIndex        =   4
         Top             =   285
         Width           =   2820
         _ExtentX        =   4974
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
      Begin VB.Label Label1 
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
         Left            =   195
         TabIndex        =   12
         Top             =   315
         Width           =   810
      End
      Begin VB.Label lblDescIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1020
         TabIndex        =   11
         Top             =   675
         Width           =   2790
      End
      Begin VB.Label Label2 
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
         Left            =   3990
         TabIndex        =   10
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta :"
         Height          =   210
         Left            =   3990
         TabIndex        =   9
         Top             =   765
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         Height          =   210
         Left            =   195
         TabIndex        =   8
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lblDescCtaTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5010
         TabIndex        =   7
         Top             =   720
         Width           =   2490
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Estado :"
         Height          =   210
         Left            =   195
         TabIndex        =   6
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   3990
         TabIndex        =   5
         Top             =   1140
         Width           =   1020
      End
   End
   Begin TabDlg.SSTab tabcuentas 
      Height          =   3690
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   6509
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   617
      TabCaption(0)   =   "Interés"
      TabPicture(0)   =   "frmCajaGenMantCtasHC.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdDelInt"
      Tab(0).Control(1)=   "cmdAddInt"
      Tab(0).Control(2)=   "fraInteres"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Plazo Fijo"
      TabPicture(1)   =   "frmCajaGenMantCtasHC.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PlazoFijo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Adeudados"
      TabPicture(2)   =   "frmCajaGenMantCtasHC.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkMalPagador"
      Tab(2).Control(1)=   "fraAdeudados"
      Tab(2).Control(2)=   "fraPlazos"
      Tab(2).Control(3)=   "cmdCalendario"
      Tab(2).Control(4)=   "txtLinCredCod"
      Tab(2).Control(5)=   "Label28"
      Tab(2).Control(6)=   "lblLinCredDesc"
      Tab(2).Control(7)=   "Label27"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Plazo Fijo Restringido"
      TabPicture(3)   =   "frmCajaGenMantCtasHC.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fgPlazoInt"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdHabilitar"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdRestringir"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdAgregar"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdEliminar"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdCancelar"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.CheckBox chkMalPagador 
         Caption         =   "Check1"
         Height          =   240
         Left            =   -73800
         TabIndex        =   75
         Top             =   3375
         Width           =   210
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3000
         Width           =   1080
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2985
         Width           =   1080
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3000
         Width           =   1110
      End
      Begin VB.CommandButton cmdRestringir 
         Caption         =   "&Restringir"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3000
         Width           =   1110
      End
      Begin VB.CommandButton cmdHabilitar 
         Caption         =   "&Habilitar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Frame PlazoFijo 
         Caption         =   "Plazo Fijo"
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
         Height          =   1890
         Left            =   -74820
         TabIndex        =   50
         Top             =   975
         Width           =   7185
         Begin VB.TextBox txtInteres 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   368
            Width           =   1410
         End
         Begin Spinner.uSpinner spnPlazo 
            Height          =   300
            Left            =   1230
            TabIndex        =   52
            Top             =   375
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            Max             =   1000
            MaxLength       =   4
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
         Begin MSMask.MaskEdBox txtFechaAper 
            Height          =   300
            Left            =   1230
            TabIndex        =   53
            Top             =   885
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaInt 
            Height          =   300
            Left            =   3840
            TabIndex        =   54
            Top             =   885
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaCap 
            Height          =   300
            Left            =   3840
            TabIndex        =   55
            Top             =   1365
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaVenc 
            Height          =   300
            Left            =   1230
            TabIndex        =   56
            Top             =   1365
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Días"
            Height          =   210
            Left            =   2070
            TabIndex        =   63
            Top             =   420
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Plazo a :"
            Height          =   210
            Left            =   225
            TabIndex        =   62
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Interes Reg."
            Height          =   210
            Left            =   2640
            TabIndex        =   61
            Top             =   420
            Width           =   870
         End
         Begin VB.Label lblFechaCap 
            AutoSize        =   -1  'True
            Caption         =   "Capitalización :"
            Height          =   210
            Left            =   2640
            TabIndex        =   60
            Top             =   1410
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Interes Al:"
            Height          =   210
            Left            =   2640
            TabIndex        =   59
            Top             =   930
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   210
            Left            =   210
            TabIndex        =   58
            Top             =   930
            Width           =   735
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento :"
            Height          =   210
            Left            =   180
            TabIndex        =   57
            Top             =   1410
            Width           =   975
         End
      End
      Begin VB.Frame fraAdeudados 
         Height          =   2220
         Left            =   -72210
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox txtPerGracia 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   750
            MaxLength       =   3
            TabIndex        =   34
            Text            =   "0"
            Top             =   975
            Width           =   585
         End
         Begin VB.TextBox txtComisionInicial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3450
            TabIndex        =   33
            Text            =   "0"
            Top             =   1005
            Width           =   750
         End
         Begin VB.TextBox txtCuotaCap 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H80000006&
            Height          =   315
            Left            =   3450
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   32
            Text            =   "0"
            Top             =   1815
            Width           =   750
         End
         Begin VB.TextBox txtTramo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3450
            TabIndex        =   31
            Text            =   "0"
            Top             =   570
            Width           =   750
         End
         Begin VB.CheckBox chkInterno 
            Caption         =   "Plaza Interna"
            Height          =   210
            Left            =   180
            TabIndex        =   30
            Top             =   645
            Width           =   1575
         End
         Begin VB.TextBox txtComisionMonto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2970
            TabIndex        =   29
            Text            =   "0"
            Top             =   1365
            Width           =   1230
         End
         Begin MSMask.MaskEdBox txtFechaUltpago 
            Height          =   330
            Left            =   975
            TabIndex        =   35
            Top             =   1380
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtContrato 
            Height          =   300
            Left            =   975
            TabIndex        =   36
            Top             =   1815
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Sicmact.EditMoney txtPorcentAfect 
            Height          =   315
            Left            =   3450
            TabIndex        =   37
            Top             =   210
            Width           =   750
            _ExtentX        =   1138
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Ult.Pago :"
            Height          =   195
            Left            =   165
            TabIndex        =   49
            Top             =   1455
            Width           =   705
         End
         Begin VB.Label Label13 
            Caption         =   "Dias"
            Height          =   195
            Left            =   1365
            TabIndex        =   48
            Top             =   1035
            Width           =   375
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   4080
            TabIndex        =   47
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cuota Pago K:"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   2370
            TabIndex        =   46
            Top             =   1845
            Width           =   1035
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Comision Inicial :"
            Height          =   210
            Left            =   2220
            TabIndex        =   45
            Top             =   1095
            Width           =   1170
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Gracia :"
            Height          =   195
            Left            =   150
            TabIndex        =   44
            Top             =   1035
            Width           =   555
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   4245
            TabIndex        =   43
            Top             =   705
            Width           =   150
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Tramo Concesional :"
            Height          =   210
            Left            =   1920
            TabIndex        =   42
            Top             =   645
            Width           =   1470
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Contrato :"
            Height          =   210
            Left            =   150
            TabIndex        =   41
            Top             =   1860
            Width           =   705
         End
         Begin VB.Label lblComision 
            Alignment       =   2  'Center
            Caption         =   "Monto"
            Height          =   285
            Left            =   2460
            TabIndex        =   40
            Top             =   1455
            Width           =   465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   1920
            X2              =   4530
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   2010
            X2              =   4620
            Y1              =   1755
            Y2              =   1755
         End
         Begin VB.Label lblPorcentAfect 
            Caption         =   "Porcentaje  de  Afectacion  : "
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
            Left            =   225
            TabIndex        =   39
            Top             =   255
            Width           =   3120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   4275
            TabIndex        =   38
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame fraPlazos 
         Height          =   2220
         Left            =   -74820
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   2580
         Begin VB.TextBox txtPlazo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1020
            TabIndex        =   23
            Text            =   "0"
            Top             =   1440
            Width           =   765
         End
         Begin VB.ComboBox cboTpoCuota 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCajaGenMantCtasHC.frx":037A
            Left            =   1020
            List            =   "frmCajaGenMantCtasHC.frx":0384
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1020
            Width           =   1410
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   1020
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   240
            Width           =   1410
         End
         Begin Spinner.uSpinner spnCuotas 
            Height          =   315
            Left            =   1020
            TabIndex        =   22
            Top             =   630
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "N° Cuotas :"
            Height          =   195
            Left            =   165
            TabIndex        =   27
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Plazo                             Dias "
            Height          =   195
            Left            =   165
            TabIndex        =   26
            Top             =   1530
            Width           =   2055
         End
         Begin VB.Label Label24 
            Caption         =   "Tip.Cuota :"
            Height          =   255
            Left            =   150
            TabIndex        =   25
            Top             =   1140
            Width           =   1245
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Capital :"
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
            Left            =   150
            TabIndex        =   24
            Top             =   300
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdCalendario 
         Caption         =   "Calen&dario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68910
         TabIndex        =   18
         Top             =   3000
         Width           =   1260
      End
      Begin VB.CommandButton cmdDelInt 
         Caption         =   "&Modificar"
         Height          =   360
         Left            =   -68865
         TabIndex        =   17
         Top             =   2970
         Width           =   1260
      End
      Begin VB.CommandButton cmdAddInt 
         Caption         =   "&Agregar"
         Height          =   360
         Left            =   -70155
         TabIndex        =   16
         Top             =   2970
         Width           =   1260
      End
      Begin VB.Frame fraInteres 
         Height          =   2055
         Left            =   -74820
         TabIndex        =   14
         Top             =   840
         Width           =   7215
         Begin Sicmact.FlexEdit fgInteres 
            Height          =   1830
            Left            =   60
            TabIndex        =   15
            Top             =   150
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3228
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Registro-Interes-Período-cCtaIfCod-cPersCod-cIfTpo-lbNuevo"
            EncabezadosAnchos=   "500-2000-2000-1800-0-0-0-0"
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-X-X-X-X"
            ListaControles  =   "0-2-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-R-R-C-C-C-C"
            FormatosEdit    =   "0-0-2-3-3-3-3-3"
            CantDecimales   =   4
            TextArray0      =   "N°"
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin Sicmact.TxtBuscar txtLinCredCod 
         Height          =   315
         Left            =   -73800
         TabIndex        =   64
         Top             =   3000
         Width           =   1545
         _ExtentX        =   2725
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
      Begin Sicmact.FlexEdit fgPlazoInt 
         Height          =   2055
         Left            =   120
         TabIndex        =   68
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3625
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Cuenta-Descripcion-Restringido-Capital-Total-CtaCont-CtaContRe-valor"
         EncabezadosAnchos=   "500-2000-0-1500-1500-1500-1200-1200-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-X-X-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-R-R-L-L-C-R"
         FormatosEdit    =   "0-0-1-3-2-1-0-0-3"
         CantDecimales   =   4
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label28 
         Caption         =   "Mal Pagador"
         Height          =   255
         Left            =   -74790
         TabIndex        =   76
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label lblLinCredDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -72195
         TabIndex        =   66
         Top             =   3015
         Width           =   3180
      End
      Begin VB.Label Label27 
         Caption         =   "Linea Crédito"
         Height          =   225
         Left            =   -74820
         TabIndex        =   65
         Top             =   3060
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCajaGenMantCtasHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim rsAdeud As ADODB.Recordset
Dim lnMonedaPago As Moneda
Dim lnCapitalCal As Currency

Dim lsCodPersG As String
Dim lsIFTpoG As String
Dim lsCtaIFCodG As String

Dim gNroFilas As Integer ' numero de filas recuperadas de Plazo Fijo Restringido
'Dim gFilaActual As Integer

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtSaldo.Enabled Then
        txtSaldo.SetFocus
    Else
        tabcuentas.Tab = 0
        fgInteres.SetFocus
    End If
End If
End Sub

Private Sub CmdAceptar_Click()
Dim lsMovNro As String
Dim nMalPg As Integer
Dim oCont As NContFunciones

Set oCont = New NContFunciones
If txtFechaInt = "__/__/____" Then
    txtFechaInt = txtFechaAper
End If
If txtFechaUltpago = "__/__/____" Then
   txtFechaUltpago = txtFechaAper
End If
If txtFechaCap = "__/__/____" Then
   txtFechaCap = txtFechaAper
End If

If txtFechaVenc = "__/__/____" Then
   txtFechaVenc = CDate(txtFechaCap) + spnPlazo.Valor
End If

If chkMalPagador.value = 1 Then
   nMalPg = 1
Else
   nMalPg = 0
End If

If Valida = False Then Exit Sub

If MsgBox("Desea Actualizar la Información registrada??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oCtaIf.ActualizaCtas Mid(txtBuscarCtaIF, 4, 13), Mid(txtBuscarCtaIF, 1, 2), Mid(txtBuscarCtaIF, 18, 10), _
                         txtNroCtaIF, txtFechaAper, txtFechaCap, txtFechaInt, spnPlazo.Valor, txtInteres, _
                         Val(Right(cboEstado, 3)), lsMovNro, fgInteres.GetRsNew, nVal(txtCapital), Me.spnCuotas.Valor, nVal(Me.txtPerGracia), nVal(Me.txtComisionInicial), nVal(txtComisionMonto), Val(Me.chkInterno.value), nVal(Me.txtCuotaCap), Me.txtFechaVenc, Me.txtFechaUltpago, nVal(Right(cboTpoCuota, 2)), CCur(Me.txtTramo), rsAdeud, lnMonedaPago, nVal(frmAdeudCal.txtComision), Me.txtLinCredCod, nVal(frmAdeudCal.txtFechaCuota), nVal(frmAdeudCal.txtCapital) _
                         , Me.txtPorcentAfect.value, nMalPg
    txtBuscarCtaIF.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
    MsgBox "Datos Actualizados satisfactoriamente", vbInformation, "¡Aviso!"
    
    Unload frmAdeudCal
    Set rsAdeud = Nothing
End If
Set oCont = Nothing
End Sub

Function Valida() As Boolean
Valida = True
If Trim(Len(txtNroCtaIF)) = 0 Then
    MsgBox "Nro de Cuenta no Ingresado", vbInformation, "Aviso"
    Valida = False
    txtNroCtaIF.SetFocus
    Exit Function
End If
If cboEstado = "" Then
    MsgBox "Estado no válido", vbInformation, "Aviso"
    Valida = False
    cboEstado.SetFocus
    Exit Function
End If
If fgInteres.TextMatrix(1, 0) = "" Then
    If MsgBox("Lista de Interés se encuentra vacia. Desea proseguir??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        Valida = False
        cboEstado.SetFocus
        Exit Function
    End If
End If
If tabcuentas.TabVisible(1) Then
    If spnPlazo.Valor = 0 And (Val(Mid(txtBuscarCtaIF, 18, 2)) = gTpoCtaIFCtaPF Or Val(Mid(txtBuscarCtaIF, 18, 2)) = gTpoCtaIFCtaAdeud) Then
        MsgBox "Plazo de Cuenta no válido", vbInformation, "Aviso"
        Valida = False
        spnPlazo.SetFocus
        Exit Function
    End If
    If txtInteres = "" Then
        MsgBox "Monto de Interes no Ingresado o no válido", vbInformation, "Aviso"
        Valida = False
        txtInteres.SetFocus
        Exit Function
    End If
    If ValFecha(txtFechaAper) = False Then
        Valida = False
        Exit Function
    End If
    If ValFecha(txtFechaInt) = False Then
        Valida = False
        Exit Function
    End If
    If ValFecha(txtFechaCap) = False Then
        Valida = False
        Exit Function
    End If
    If CDate(txtFechaAper) > CDate(txtFechaInt) Then
        MsgBox "Fecha de Apertura no puede ser mayor que fecha de Interes"
        txtFechaInt.SetFocus
        Valida = False
        Exit Function
    End If
    If CDate(txtFechaAper) > CDate(txtFechaCap) And txtFechaCap.Visible Then
        MsgBox "Fecha de Apertura no puede ser mayor que fecha de Capitalización"
        If txtFechaCap.Visible Then
            txtFechaCap.SetFocus
        End If
        Valida = False
        Exit Function
    End If
End If
If tabcuentas.TabVisible(2) Then
   
End If

End Function
Private Sub cmdAddInt_Click()
fgInteres.SoloFila = True
fgInteres.lbEditarFlex = True
fgInteres.AdicionaFila
fgInteres.TextMatrix(fgInteres.Row, 1) = gdFecSis
fgInteres.TextMatrix(fgInteres.Row, 3) = 360
SendKeys "{ENTER}"
fgInteres.SetFocus
End Sub

Private Sub cmdAgregar_Click()
    fgPlazoInt.AdicionaFila
   
    cmdRestringir.Enabled = False
    cmdHabilitar.Enabled = False
    
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = False
    cmdCancelar.Enabled = True
        
    fgPlazoInt.Col = 2
    Call fgPlazoInt_RowColChange
    fgPlazoInt.Col = 1
End Sub

Private Sub cmdCalendario_Click()

If txtBuscarCtaIF <> "" And Val(txtCapital) > 0 Then
    frmAdeudCal.Inicio True, Trim(txtBuscarCtaIF), _
                       Trim(lblDescIF) & " " & Me.txtNroCtaIF, _
                        txtCapital, txtContrato, nVal(fgInteres.TextMatrix(fgInteres.Rows - 1, 2)), _
                       True _
                       
                       
    If frmAdeudCal.OK Then
        Set rsAdeud = frmAdeudCal.fgCronograma.GetRsNew(1)
        spnCuotas.Valor = frmAdeudCal.spnCuotas.Valor
        If frmAdeudCal.optTpoCuota(0) Then
            cboTpoCuota.ListIndex = 0
        Else
            cboTpoCuota.ListIndex = 1
        End If
        lnCapitalCal = frmAdeudCal.txtCapital
        txtPlazo = frmAdeudCal.txtPlazoCuotas
        spnPlazo.Valor = txtPlazo
        chkInterno = frmAdeudCal.chkInterno
        txtPerGracia = frmAdeudCal.SpnGracia.Valor
        txtTramo = frmAdeudCal.txtTramo
        txtCuotaCap = frmAdeudCal.txtCuotaPagoK
        If Mid(txtBuscarCtaIF, 20, 1) = "1" Then
            lnMonedaPago = IIf(frmAdeudCal.chkVac = vbChecked, gMonedaExtranjera, gMonedaNacional)
        Else
            lnMonedaPago = Mid(txtBuscarCtaIF, 20, 1)
        End If
    Else
        Set rsAdeud = Nothing
    End If
    'cmdAceptar.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
If fgPlazoInt.Row <= gNroFilas Then
   Exit Sub
End If
fgPlazoInt.EliminaFila (fgPlazoInt.Row)
fgPlazoInt_RowColChange
'gFilaActual = fgPlazoInt.Rows
If fgPlazoInt.TextMatrix(1, 1) = "" Then
   cmdRestringir.Enabled = False
   cmdHabilitar.Enabled = False
   
End If
cmdAgregar.Enabled = True
cmdEliminar.Enabled = True
cmdCancelar.Enabled = False
End Sub

Private Sub cmdDelInt_Click()
If fgInteres.TextMatrix(1, 0) <> "" Then
    If fgInteres.Row = fgInteres.Rows - 1 Then
        If Val(fgInteres.TextMatrix(fgInteres.Row, 7)) = 1 Then
           If MsgBox("¿ Seguro que desea Modificar datos Interes ya registrado ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
              Exit Sub
           End If
        End If
    Else
        MsgBox "No se puede modificar datos de interés porque existen datos con fechas posteriores", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    fgInteres.SoloFila = True
    fgInteres.lbEditarFlex = True
    fgInteres.TextMatrix(fgInteres.Row, 7) = "2"
    SendKeys "{ENTER}"
    fgInteres.SetFocus
End If
End Sub

'''Private Sub EliminaCuenta(sCod As String, nItem As Integer)
'''If fgDetalle.TextMatrix(1, 0) <> "" Then
'''    EliminaFgObj Val(fgDetalle.TextMatrix(fgDetalle.Row, 0))
'''    fgDetalle.EliminaFila fgDetalle.Row, False
'''End If
'''If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
'''   RefrescaFgObj Val(fgDetalle.TextMatrix(fgDetalle.Row, 0))
'''End If
'''End Sub

Private Sub cmdEliminar_Click()
Dim I          As Integer
Dim oCtaIf     As NCajaCtaIF
Dim oMov       As New DMov
Dim gsMovNro   As String

Set oCtaIf = New NCajaCtaIF
Set oMov = New DMov

If fgPlazoInt.TextMatrix(1, 2) = "" Then
    MsgBox "No existe elemento para Eliminar", vbInformation, "Aviso"
    Me.fgPlazoInt.EliminaFila (Me.fgPlazoInt.Row)
    Exit Sub
End If

If Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 3) <> "0" Then
   MsgBox " No se puede eliminar esta Garantía debido a que aun existe capital en restringido " & vbCrLf & "               realize la Habilitacion del Capital y luego Elimine   ", vbInformation, "Aviso"
   Exit Sub
End If
   gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
   If MsgBox(" ¿ Esta seguro de eliminar esta garantia ? ", vbQuestion + vbYesNo) = vbYes Then
       If oCtaIf.fbEliminarGarantia(lsCodPersG, lsIFTpoG, lsCtaIFCodG, Left(Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 1), 13), Mid(Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 1), 15, 2), Right(Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 1), 7), gsMovNro) = False Then
          MsgBox " Garantia no fue  eliminada ", vbInformation, "Aviso"
          Exit Sub
       Else
          Me.fgPlazoInt.EliminaFila (Me.fgPlazoInt.Row)
          'gfilactual = fgPlazoInt.Rows - 1
          Me.cmdRestringir.Enabled = False
          Call MuestraAdeudoGarantias(lsCodPersG, lsIFTpoG, lsCtaIFCodG)
       End If
   End If
'    Me.fgPlazoInt.EliminaFila (Me.fgPlazoInt.Row)
'    Me.cmdRestringir.Enabled = False
'    Call MuestraAdeudoGarantias(lsCodPersG, lsIFTpoG, lsCtaIFCodG)

'fgPlazoInt.EliminaFila (fgPlazoInt.Row)
If fgPlazoInt.TextMatrix(1, 1) = "" Then
    cmdRestringir.Enabled = False
    cmdHabilitar.Enabled = False
End If
End Sub

Private Sub cmdHabilitar_Click()
Dim oCtaIf As NCajaCtaIF
Dim oMov As DMov
Dim oCta   As DCtaCont

Dim gsMovNro As String
Dim nCta As String
Dim nCtaRe As String
Dim nImporte As Currency
Dim nOpeCod As Long
Dim I As Integer
Dim nItem As Long
Dim sPersCod As String
Dim sIFTpo As String
Dim sCtaIFCod As String
Dim lsMensaje As String


Set oCta = New DCtaCont
Set oCtaIf = New NCajaCtaIF
Set oMov = New DMov

If fgPlazoInt.TextMatrix(1, 2) = "" Then
    MsgBox "Debe Ingresar una Cuenta", vbInformation, "Aviso"
    Exit Sub
End If
If fgPlazoInt.TextMatrix(fgPlazoInt.Row, 3) = 0 Then
    MsgBox "Cuenta no se puede Habilitar", vbInformation, "Aviso"
    Exit Sub
Else
    
    nCta = fgPlazoInt.TextMatrix(fgPlazoInt.Row, 6)  'cCtaCod
    nCtaRe = Left(nCta, 3) & "70901" & Mid(nCta, 7, 2) & "03" 'cCtaCodRestringida
        
    lsMensaje = oCta.VerificaExisteCuenta(nCtaRe, True)
        
    If lsMensaje <> "" Then
       MsgBox lsMensaje, vbInformation, "Aviso"
       fgPlazoInt.EliminaFila (fgPlazoInt.Row)
       Exit Sub
    End If
End If

 If MsgBox(" ¿ Esta Seguro de Realizar la operacion ? ", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If

    'nCta = fgPlazoInt.TextMatrix(fgPlazoInt.Row, 6)  'cCtaCod
    'nCtaRe = Left(nCta, 3) & "7" & Mid(nCta, 5, Len(nCta)) 'cCtaCodRestringida
    nImporte = fgPlazoInt.TextMatrix(fgPlazoInt.Row, 3) 'nMovImporte
                       
    sPersCod = Mid(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1), 1, 13)
    sIFTpo = Mid(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1), 15, 2)
    sCtaIFCod = Mid(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1), 18, 7)
                           
    If Mid(nCta, 3, 1) = 1 Then
       nOpeCod = gOpeCGOpeMantCtasBancosReMN 'cOpeCod
    Else
       nOpeCod = gOpeCGOpeMantCtasBancosReME 'cOpeCod
    End If
            
    gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
         If oCtaIf.Habilitar(gsMovNro, nOpeCod, nCta, nImporte, sPersCod, sIFTpo, sCtaIFCod, nCtaRe, lsCodPersG, lsIFTpoG, lsCtaIFCodG) = False Then
            MsgBox " El registro de Habilitacion no se efectuó ", vbInformation
         Else
            MsgBox "Cuenta Habilitada", vbInformation, "Aviso"
            Call MuestraAdeudoGarantias(lsCodPersG, lsIFTpoG, lsCtaIFCodG)
'                'ImprimeAsientoContable gsMovNro, , , , False, False, "RESTRINGIR", Left(fgPlazoInt.TextMatrix(i, 1), 12), fgPlazoInt.TextMatrix(i, 6), , , , 1, , "17", , "", ""
        End If
Set oMov = Nothing
Set oCtaIf = Nothing
Set oCta = Nothing
End Sub

Private Sub cmdRestringir_Click()
Dim oCtaIf As NCajaCtaIF
Dim oMov As DMov
Dim oCta   As DCtaCont

Dim gsMovNro As String
Dim nCta As String
Dim nCtaRe As String
Dim nImporte As Currency
Dim nOpeCod As Long
Dim I As Integer
Dim nItem As Long
Dim sPersCod As String
Dim sIFTpo As String
Dim sCtaIFCod As String
Dim lsMensaje As String


Set oCta = New DCtaCont
Set oCtaIf = New NCajaCtaIF
Set oMov = New DMov

If fgPlazoInt.TextMatrix(1, 2) = "" Then
    MsgBox "Debe Ingresar una Cuenta", vbInformation, "Aviso"
    Exit Sub
End If

'If Val(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1)) = "" Then Exit Sub

If fgPlazoInt.TextMatrix(fgPlazoInt.Row, 4) = 0 Then
    MsgBox "Cuenta no se puede Restrigir", vbInformation, "Aviso"
    Exit Sub
Else
    
    nCta = fgPlazoInt.TextMatrix(fgPlazoInt.Row, 6)  'cCtaCod
    nCtaRe = Left(nCta, 3) & "70901" & Mid(nCta, 7, 2) & "03" 'cCtaCodRestringida
        
    lsMensaje = oCta.VerificaExisteCuenta(nCtaRe, True)
        
    If lsMensaje <> "" Then
       MsgBox lsMensaje, vbInformation, "Aviso"
       fgPlazoInt.EliminaFila (fgPlazoInt.Row)
       cmdAgregar.Enabled = True
       cmdEliminar.Enabled = True
       cmdCancelar.Enabled = False
       Exit Sub
    End If
End If

 If MsgBox(" ¿ Esta Seguro de Realizar la operacion ? ", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If

    nImporte = fgPlazoInt.TextMatrix(fgPlazoInt.Row, 4) 'nMovImporte
                       
    sPersCod = Mid(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1), 1, 13)
    sIFTpo = Mid(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1), 15, 2)
    sCtaIFCod = Mid(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 1), 18, 7)
                           
        If Mid(nCta, 3, 1) = 1 Then
           nOpeCod = gOpeCGOpeMantCtasBancosReMN 'cOpeCod
        Else
           nOpeCod = gOpeCGOpeMantCtasBancosReME 'cOpeCod
        End If
            
    gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
         If oCtaIf.Restringir(gsMovNro, nOpeCod, nCta, nImporte, sPersCod, sIFTpo, sCtaIFCod, nCtaRe, lsCodPersG, lsIFTpoG, lsCtaIFCodG) = False Then
            MsgBox " El registro del Restringido no se efectuó ", vbInformation
         Else
            MsgBox "Cuenta Restringida", vbInformation, "Aviso"
            Call MuestraAdeudoGarantias(lsCodPersG, lsIFTpoG, lsCtaIFCodG)
'                'ImprimeAsientoContable gsMovNro, , , , False, False, "RESTRINGIR", Left(fgPlazoInt.TextMatrix(i, 1), 12), fgPlazoInt.TextMatrix(i, 6), , , , 1, , "17", , "", ""
        End If
Set oMov = Nothing
Set oCtaIf = Nothing
Set oCta = Nothing
End Sub
Private Sub MuestraAdeudoGarantias(psPersCod As String, psIFTpo As String, psCtaIFCod As String)
    Dim res As New ADODB.Recordset
    'Dim oGar As New DACGAdeuGarantias
    Dim oCtaIf As NCajaCtaIF
    
    Set res = New ADODB.Recordset
    
    Me.fgPlazoInt.Clear
    Me.fgPlazoInt.Rows = 2
    Me.fgPlazoInt.FormaCabecera
    'gFilaActual = 0
    Set oCtaIf = New NCajaCtaIF
    Set res = oCtaIf.foACGBuscarGarantia(psPersCod, psIFTpo, psCtaIFCod, gdFecSis)
    If RSVacio(res) Then
        Me.fgPlazoInt.Rows = 2
    Else
        While Not res.EOF
            Me.fgPlazoInt.AdicionaFila
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 1) = res!Codigo
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 2) = res!Descripcion
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 3) = res!sr
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 4) = res!sk
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 5) = res!sr + res!sk
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 6) = res!ctak
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 7) = res!Ctar
            Me.fgPlazoInt.TextMatrix(Me.fgPlazoInt.Row, 8) = 1
            res.MoveNext
        Wend
        Me.fgPlazoInt.lbEditarFlex = True
        Call fgPlazoInt_RowColChange
    End If
    gNroFilas = res.RecordCount
    If res.RecordCount >= 1 Then
        cmdRestringir.Enabled = True
        cmdHabilitar.Enabled = True
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCancelar.Enabled = False
    Else
        cmdAgregar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCancelar.Enabled = True
    End If
    
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgPlazoInt_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim I As Integer
Set oCtaIf = New NCajaCtaIF
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset

Set rs1 = oCtaIf.GetCtaIFCont(gdFecSis, psDataCod)
If Not rs1.EOF Then
    If fgPlazoInt.TextMatrix(pnRow, 1) = "" Then
       MsgBox "Cuenta ya Asignada", vbInformation, "Aviso"
       fgPlazoInt.Rows = fgPlazoInt.Rows - 1
       cmdAgregar.Enabled = True
       cmdEliminar.Enabled = True
       cmdCancelar.Enabled = False
       cmdSalir.Enabled = True
       Exit Sub
    End If
    
    fgPlazoInt.TextMatrix(pnRow, 3) = rs1!Restringido
    fgPlazoInt.TextMatrix(pnRow, 4) = rs1!Capital
    fgPlazoInt.TextMatrix(pnRow, 5) = rs1!Total
    fgPlazoInt.TextMatrix(pnRow, 6) = rs1!cCtaContCod
    
    cmdRestringir.Enabled = True
    cmdHabilitar.Enabled = True

End If

Set oCtaIf = Nothing
End Sub

Private Sub fgPlazoInt_RowColChange()
Dim I As Integer
Dim nFilaAct As Integer

With fgPlazoInt
    For I = 1 To .Rows - 1
        If Val(.TextMatrix(.Row, 8)) = 0 Then
            nFilaAct = .Row
            Exit For
        End If
    Next
End With
If nFilaAct = 0 Then
    fgPlazoInt.lbEditarFlex = False
Else
    fgPlazoInt.lbEditarFlex = True
    If Val(fgPlazoInt.TextMatrix(fgPlazoInt.Row, 8)) = 1 Then
        fgPlazoInt.Row = nFilaAct
    End If
End If
   
End Sub

Private Sub Form_Load()
Set oCtaIf = New NCajaCtaIF
Dim oGen As DGeneral

Set oGen = New DGeneral
Set oOpe = New DOperacion
txtBuscarCtaIF.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")
Me.Caption = "  " & gsOpeDesc
CargaCombo cboEstado, oGen.GetConstante(gCGEstadoCtaIF)
CargaCombo cboTpoCuota, oGen.GetConstante(gCGAdeudCalTpoCuota)
CentraForm Me
tabcuentas.TabVisible(1) = False
tabcuentas.TabVisible(2) = False
fraInteres.Enabled = False
fgPlazoInt.rsTextBuscar = oCtaIf.GetCtaIFContArb(gdFecSis)
End Sub


Private Sub txtComisionInicial_GotFocus()
  fEnfoque txtComisionInicial
End Sub

Private Sub txtComisionInicial_KeyPress(KeyAscii As Integer)
If nVal(txtCapital) = 0 Then
    MsgBox "Primero ingresar Monto de Prestamo", vbInformation, "¡Aviso!"
    txtCapital.SetFocus
    Exit Sub
End If
KeyAscii = NumerosDecimales(txtComisionInicial, KeyAscii, 10, 4)
If KeyAscii = 13 Then
   txtComisionMonto = Format(Round(nVal(txtCapital) * txtComisionInicial / 100, 2), gsFormatoNumeroView)
   txtComisionInicial = Format(txtComisionInicial, gsFormatoNumeroView)
   txtComisionMonto.SetFocus
End If
End Sub

Private Sub txtComisionMonto_GotFocus()
fEnfoque txtComisionMonto
End Sub

Private Sub txtComisionMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtComisionMonto, KeyAscii, 14, 4)
If KeyAscii = 13 Then
   txtCuotaCap.SetFocus
End If
End Sub

Private Sub txtLinCredCod_EmiteDatos()
lblLinCredDesc = txtLinCredCod.psDescripcion
If lblLinCredDesc <> "" Then
    cmdCalendario.SetFocus
End If
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtSaldo, KeyAscii)
If KeyAscii = 13 Then
    cmdAddInt.SetFocus
End If
End Sub

Private Sub txtcapital_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCapital, KeyAscii)
If KeyAscii = 13 Then
    If spnCuotas.Enabled Then
        spnCuotas.SetFocus
    Else
        cmdCalendario.SetFocus
    End If
End If
End Sub

Private Sub txtNroCtaIF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCapital.Enabled And txtCapital.Visible Then
        txtCapital.SetFocus
    ElseIf Me.txtSaldo.Enabled Then
        txtSaldo.SetFocus
    'Else
        'cmdAceptar.SetFocus
    End If
End If
End Sub

Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtInteres.SetFocus
End If
End Sub

Private Sub txtBuscarCtaIF_EmiteDatos()
On Error Resume Next
If txtBuscarCtaIF <> "" Then
    Set frmAdeudCal = Nothing
    Set rsAdeud = Nothing
    fraInteres.Enabled = False
    
    Dim oAdeud As New NCajaAdeudados
    txtLinCredCod = ""
    lblLinCredDesc = ""
    txtLinCredCod.rs = oAdeud.GetLineaCredito(Mid(gsOpeCod, 3, 1), Mid(txtBuscarCtaIF, 4, 13))
    Set oAdeud = Nothing
    
    CargaDatosCuentas Mid(txtBuscarCtaIF, 4, 13), Mid(txtBuscarCtaIF, 1, 2), Mid(txtBuscarCtaIF, 18, 10)
    txtNroCtaIF.SetFocus
End If
End Sub
Sub CargaDatosCuentas(psPersCod As String, pnIfTpo As CGTipoIF, psCtaIFCod As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Limpiar
fraInteres.Enabled = True
tabcuentas.TabVisible(1) = True

    lsCodPersG = psPersCod
    lsIFTpoG = pnIfTpo
    lsCtaIFCodG = psCtaIFCod

'lblDescCtaTipo = oCtaIf.NombreIF(psPersCod)
txtNroCtaIF = Trim(txtBuscarCtaIF.psDescripcion)
lblDescIF = oCtaIf.NombreIF(psPersCod)
lblDescCtaTipo = oCtaIf.EmiteTipoCuentaIF(psCtaIFCod)
If Mid(txtBuscarCtaIF, 19, 1) = gTpoCtaIFCtaAdeud Then
    tabcuentas.TabCaption(1) = "Datos de Pagaré"
    tabcuentas.TabVisible(2) = True
    PlazoFijo.Caption = ""
Else
    tabcuentas.TabCaption(1) = "Plazo Fijo"
    PlazoFijo.Caption = "Plazo Fijo"
End If
Set rs = oCtaIf.GetRsInteresCtasIF(psPersCod, pnIfTpo, psCtaIFCod)
If Not rs.EOF And Not rs.BOF Then
    Set fgInteres.Recordset = rs
End If
rs.Close
Set rs = oCtaIf.GetDatosCtaIf(psPersCod, pnIfTpo, psCtaIFCod)
If Not rs.EOF And Not rs.EOF Then
    txtSaldo = Format(Abs(rs!Saldo), "#,#0.00")
    cboEstado = rs!cEstadoCons & Space(50) & rs!cCtaIFEstado
    txtFechaAper = IIf(IsNull(rs!dCtaIFAper), "__/__/____", Format(rs!dCtaIFAper, gsFormatoFechaView))
    txtFechaCap = IIf(IsNull(rs!dCtaIfCap), "__/__/____", Format(rs!dCtaIfCap, gsFormatoFechaView))
    txtFechaInt = IIf(IsNull(rs!dCtaIfInt), "__/__/____", Format(rs!dCtaIfInt, gsFormatoFechaView))
    txtFechaVenc = IIf(IsNull(rs!dCtaIFVenc), "__/__/____", Format(rs!dCtaIFVenc, gsFormatoFechaView))
    If Not IsNull(rs!nPorcentAfect) Then
        Me.txtPorcentAfect.value = rs!nPorcentAfect
    End If
    spnPlazo.Valor = rs!nCtaIFPlazo
    txtInteres = Format(rs!nInteres, gsFormatoNumeroView)

    If Mid(txtBuscarCtaIF, 19, 1) = gTpoCtaIFCtaPF Then
        tabcuentas.TabVisible(2) = False
    End If
    If Mid(txtBuscarCtaIF, 19, 1) = gTpoCtaIFCtaAdeud Then
        tabcuentas.TabVisible(2) = True
        fraPlazos.Visible = True
        fraAdeudados.Visible = True
        txtCapital = Format(Abs(rs!nMontoPrestado), gsFormatoNumeroView)
        txtSaldo = Format(Abs(rs!nSaldoCap), gsFormatoNumeroView)
        If Not IsNull(rs!nTpoCuota) Then
           cboTpoCuota = rs!cTpoCtaDesc & Space(50) & rs!nTpoCuota
        Else
           cboTpoCuota.ListIndex = 0
        End If
        txtPlazo = rs!nCtaIFPlazo
        
        If rs!bMalPg = True Then
            chkMalPagador.value = 1
        Else
            chkMalPagador.value = 0
        End If
        
        If Not IsNull(rs!nCtaIFCuotas) Then
           spnCuotas.Valor = rs!nCtaIFCuotas
        End If
        chkInterno.value = IIf(rs!cPlaza = 1, vbChecked, vbUnchecked)
        txtPerGracia = IIf(IsNull(rs!nPeriodoGracia), 0, rs!nPeriodoGracia)
        txtFechaUltpago = IIf(IsNull(rs!dCuotaUltPago), "__/__/____", rs!dCuotaUltPago)
        txtComisionInicial = IIf(IsNull(rs!nComisionInicial), 0, rs!nComisionInicial)
        txtComisionMonto = IIf(IsNull(rs!nComisionMonto), 0, rs!nComisionMonto)
        
        txtFechaCap.Visible = False
        lblFechaCap.Visible = False
        txtFechaInt = IIf(IsNull(rs!dCuotaUltPago), txtFechaInt, Format(rs!dCuotaUltPago, gsFormatoFechaView))
        txtFechaVenc = IIf(IsNull(rs!dVencimiento), txtFechaVenc, Format(rs!dVencimiento, gsFormatoFechaView))
        
        txtTramo = IIf(IsNull(rs!nTramoConcesion), 0, rs!nTramoConcesion)
        txtCuotaCap = IIf(IsNull(rs!nCuotaPagoCap), 0, rs!nCuotaPagoCap)
        txtContrato = IIf(IsNull(rs!dCtaIFAper), "__/__/____", Format(rs!dCtaIFAper, gsFormatoFechaView))
        Me.txtLinCredCod = IIf(IsNull(rs!cCodLinCred), "", rs!cCodLinCred)
        Me.lblLinCredDesc = txtLinCredCod.psDescripcion
    End If
End If
Call MuestraAdeudoGarantias(lsCodPersG, lsIFTpoG, lsCtaIFCodG)
'Set rs = oCtaIf.GetRsInteresCtasIF(psPersCod, pnIfTpo, psCtaIFCod)
'If Not rs.EOF And Not rs.BOF Then
'    Set fgPlazoInt.Recordset = rs
'End If

RSClose rs
End Sub
Private Sub txtFechaAper_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFechaInt.SetFocus
End If
End Sub
Private Sub txtFechaAper_Validate(Cancel As Boolean)
If ValFecha(txtFechaAper) = False Then
    Cancel = True
End If
End Sub
Private Sub txtFechaCap_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub
Private Sub txtFechaCap_Validate(Cancel As Boolean)
If ValFecha(txtFechaCap) = False Then
    Cancel = True
End If
End Sub
Private Sub txtFechaInt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtFechaCap.SetFocus
End If
End Sub
Private Sub txtFechaInt_Validate(Cancel As Boolean)
If ValFecha(txtFechaInt) = False Then
    Cancel = True
End If
End Sub

Private Sub txtInteres_GotFocus()
fEnfoque txtInteres

End Sub

Private Sub txtInteres_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtInteres, KeyAscii)
If KeyAscii = 13 Then
    If txtFechaAper.Enabled Then txtFechaAper.SetFocus
    If txtFechaInt.Enabled Then txtFechaInt.SetFocus
End If
End Sub
Sub Limpiar()
tabcuentas.TabVisible(1) = False
tabcuentas.TabVisible(2) = False
txtNroCtaIF = ""
lblDescIF = ""
fraInteres.Enabled = False
lblDescCtaTipo = ""
txtNroCtaIF = ""
lblDescIF = ""
lblDescCtaTipo = ""
fgInteres.Clear
fgInteres.FormaCabecera
fgInteres.Rows = 2
spnPlazo.Valor = 0
txtFechaAper = "__/__/____"
txtFechaCap = "__/__/____"
txtFechaInt = "__/__/____"
cboEstado.ListIndex = -1
txtSaldo = "0.00"
txtCapital = "0.00"
txtInteres = "0.00"
Me.txtPorcentAfect.value = 0
End Sub

'Private Sub CalculaTotal(Optional lCalcImpuestos As Boolean = True)
'    Dim N As Integer, m As Integer
'    Dim nSTot As Currency
'    Dim nITot As Currency, nImp As Currency
'    Dim nTot  As Currency
'    Dim lnTotalFIla As Currency
'    Dim lnI As Integer
'
'    nSTot = 0: nTot = 0
'    If fgImp.TextMatrix(1, 1) = "" Then
'       lCalcImpuestos = False
'    End If
'    For m = 1 To fgImp.Rows - 1
'       nITot = 0
'       For N = 1 To fgPlazoInt.Rows - 1
'          If fgImp.TextMatrix(m, 2) = "." Then
'             If lCalcImpuestos Then
'                If fgImp.TextMatrix(m, 11) = "0" Then
'                    nImp = Round(Val(Format(fgPlazoInt.TextMatrix(N, 3), gsFormatoNumeroDato)) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
'                    fgPlazoInt.TextMatrix(N, m + 6) = Format(nImp, gsFormatoNumeroView)
'                Else
'                    lnTotalFIla = Val(Format(fgPlazoInt.TextMatrix(N, 3), gsFormatoNumeroDato))
'
'                    For lnI = 7 To Me.fgPlazoInt.Cols - 1
'                        If fgImp.TextMatrix(lnI - 7 + 1, 11) = "0" And fgPlazoInt.TextMatrix(N, lnI) <> "" Then
'                        lnTotalFIla = lnTotalFIla + Val(Format(fgPlazoInt.TextMatrix(N, lnI), gsFormatoNumeroDato))
'                        End If
'                    Next lnI
'
'                    nImp = Round(Val(lnTotalFIla) * Val(Format(fgImp.TextMatrix(m, 4), gsFormatoNumeroDato)) / 100, 2)
'                    fgPlazoInt.TextMatrix(N, m + 6) = Format(nImp, gsFormatoNumeroView)
'                End If
'             Else
'                nImp = fgPlazoInt.TextMatrix(N, m + 6)
'             End If
'             nITot = nITot + nImp
'          Else
'             If lCalcImpuestos Then fgPlazoInt.TextMatrix(N, m + 6) = "0.00"
'          End If
'       Next
'       fgImp.TextMatrix(m, 5) = Format(nITot, gsFormatoNumeroView)
'       nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 8) = "D", 1, -1)
'    Next
'    For N = 1 To fgPlazoInt.Rows - 1
'       nSTot = nSTot + Val(Format(fgPlazoInt.TextMatrix(N, 3), gsFormatoNumeroDato))
'    Next
'    txtSTotal = Format(nSTot, gsFormatoNumeroView)
'    txtTotal = Format(nSTot + nTot, gsFormatoNumeroView)
'End Sub
'

