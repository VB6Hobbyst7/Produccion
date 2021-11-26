VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdeudOperaciones1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  ADEUDADOS: OPERACIONES"
   ClientHeight    =   9315
   ClientLeft      =   1305
   ClientTop       =   615
   ClientWidth     =   10620
   Icon            =   "frmAdeudOperaciones1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar PB 
      Height          =   180
      Left            =   1620
      TabIndex        =   64
      Top             =   9120
      Visible         =   0   'False
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   63
      Top             =   9075
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15699
            MinWidth        =   15699
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCabecera 
      BorderStyle     =   0  'None
      Height          =   1350
      Left            =   60
      TabIndex        =   29
      Top             =   15
      Width           =   10260
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
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   9030
         Begin Sicmact.TxtBuscar txtCodObjeto 
            Height          =   345
            Left            =   1065
            TabIndex        =   5
            Top             =   240
            Width           =   2625
            _extentx        =   4630
            _extenty        =   609
            appearance      =   1
            font            =   "frmAdeudOperaciones1.frx":08CA
            stitulo         =   ""
            enabledtext     =   0   'False
         End
         Begin VB.Label lblObjDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4695
            TabIndex        =   37
            Top             =   255
            Width           =   3570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3750
            TabIndex        =   36
            Top             =   315
            Width           =   930
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
            TabIndex        =   35
            Top             =   285
            Width           =   630
         End
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
         Left            =   15
         TabIndex        =   31
         Top             =   15
         Width           =   3960
         Begin VB.TextBox txtOpeCod 
            Alignment       =   2  'Center
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
            Left            =   930
            TabIndex        =   0
            Top             =   195
            Width           =   900
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   315
            Left            =   2535
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
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
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1995
            TabIndex        =   33
            Top             =   225
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   225
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdCalcular 
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
         Left            =   9060
         TabIndex        =   8
         Top             =   1065
         Width           =   1125
      End
      Begin VB.CheckBox chkCancelacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelación"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   9045
         TabIndex        =   7
         Top             =   795
         Width           =   1185
      End
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
         Left            =   4050
         TabIndex        =   30
         Top             =   15
         Width           =   4965
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
            Left            =   270
            TabIndex        =   2
            Top             =   225
            Width           =   1845
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
            Width           =   1095
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
         Left            =   0
         TabIndex        =   38
         Top             =   690
         Visible         =   0   'False
         Width           =   9030
         Begin VB.TextBox txtNroCtaIF 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4680
            TabIndex        =   40
            Top             =   285
            Width           =   2310
         End
         Begin VB.ComboBox cboEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7665
            Style           =   1  'Simple Combo
            TabIndex        =   39
            Text            =   "cboEstado"
            Top             =   285
            Width           =   1290
         End
         Begin Sicmact.TxtBuscar txtBuscarCtaIF 
            Height          =   315
            Left            =   1065
            TabIndex        =   6
            Top             =   285
            Width           =   2685
            _extentx        =   4736
            _extenty        =   556
            appearance      =   1
            font            =   "frmAdeudOperaciones1.frx":08F6
            stitulo         =   ""
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   7020
            TabIndex        =   44
            Top             =   330
            Width           =   585
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
            TabIndex        =   43
            Top             =   315
            Width           =   885
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
            TabIndex        =   42
            Top             =   315
            Width           =   810
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
            TabIndex        =   41
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Left            =   9120
         Picture         =   "frmAdeudOperaciones1.frx":0922
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   9000
      TabIndex        =   27
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   9000
      TabIndex        =   26
      Top             =   7200
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9000
      TabIndex        =   28
      Top             =   8160
      Width           =   1155
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Frame5"
      Enabled         =   0   'False
      Height          =   7455
      Left            =   120
      TabIndex        =   45
      Top             =   1560
      Width           =   10245
      Begin VB.CheckBox chkConcesion 
         Appearance      =   0  'Flat
         Caption         =   "Pagar Cuota Concesionada"
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
         Height          =   255
         Left            =   6960
         TabIndex        =   86
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "&Glosa"
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
         Left            =   600
         TabIndex        =   84
         Top             =   3600
         Width           =   9420
         Begin VB.TextBox txtMovDesc 
            Height          =   555
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   0
            Width           =   9420
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   5880
         TabIndex        =   82
         Top             =   4080
         Width           =   1335
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Pagando"
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
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.CheckBox chkCambiar 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox ChkAntes 
         Caption         =   "Cancelación Adelantada"
         Height          =   435
         Left            =   8880
         TabIndex        =   80
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtCanTotal 
         Height          =   285
         Left            =   8820
         TabIndex        =   79
         Top             =   4920
         Visible         =   0   'False
         Width           =   1300
      End
      Begin VB.Frame fraEditar 
         Height          =   735
         Left            =   5640
         TabIndex        =   71
         Top             =   2760
         Width           =   4455
         Begin VB.TextBox txtEdiicionInter 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            TabIndex        =   77
            Text            =   "0.00"
            Top             =   315
            Width           =   1365
         End
         Begin VB.CommandButton cmdE_Editar 
            Caption         =   "&Editar"
            Height          =   255
            Left            =   3300
            TabIndex        =   74
            Top             =   405
            Width           =   1050
         End
         Begin VB.TextBox txtEdiicionCom 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1785
            TabIndex        =   73
            Text            =   "0.00"
            Top             =   330
            Width           =   1365
         End
         Begin VB.CommandButton cmdE_Cancelar 
            Caption         =   "C&ancelar"
            Height          =   240
            Left            =   3315
            TabIndex        =   72
            Top             =   150
            Width           =   1050
         End
         Begin VB.CommandButton cmdE_Aceptar 
            Caption         =   "A&ceptar"
            Height          =   255
            Left            =   3300
            TabIndex        =   75
            Top             =   405
            Width           =   1050
         End
         Begin VB.Label Label21 
            Caption         =   "Intereres :"
            Height          =   240
            Left            =   165
            TabIndex        =   78
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label15 
            Caption         =   "Comision :"
            Height          =   225
            Left            =   1800
            TabIndex        =   76
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.TextBox txtComision 
         BackColor       =   &H00E1FFFC&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5280
         TabIndex        =   69
         Top             =   2340
         Width           =   960
      End
      Begin VB.TextBox txtInteres 
         BackColor       =   &H00E1FFFC&
         Enabled         =   0   'False
         Height          =   300
         Left            =   3120
         TabIndex        =   67
         Top             =   2340
         Width           =   1230
      End
      Begin VB.TextBox txtMonto 
         BackColor       =   &H00E1FFFC&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   65
         Top             =   2340
         Width           =   1230
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   7320
         TabIndex        =   62
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   4320
         Width           =   1440
      End
      Begin VB.TextBox lblTotal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   8940
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2355
         Width           =   1260
      End
      Begin MSComctlLib.ListView lstCabecera 
         Height          =   1680
         Left            =   45
         TabIndex        =   9
         Top             =   600
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   2963
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
         NumItems        =   53
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Entidad"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Adeudado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cuot"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SKBase"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CapitalSC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Capital"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "IntProvBase"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Int. Prov."
            Object.Width           =   1730
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "IntCalSC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Int. Cal"
            Object.Width           =   1729
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "IntTotalSC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Int. Total"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "FechaUltPago"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Periodo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "Comision"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Dias"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "MonedaPag"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Objeto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Vencimiento"
            Object.Width           =   1924
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "TasaInt"
            Object.Width           =   1289
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "nSaldoCapLP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "cCodLinCred"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "Ajuste VAC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "cPersCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "cIFTpo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "cCtaIFCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "SaldoProviMes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Text            =   "nCapitalConce"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   49
            Text            =   "nInteresConce"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   50
            Text            =   "nSaldoCapConce"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   51
            Text            =   "nProvisionConce"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   52
            Text            =   "nComisionConce"
            Object.Width           =   2540
         EndProperty
      End
      Begin TabDlg.SSTab TabDoc 
         Height          =   2775
         Left            =   90
         TabIndex        =   11
         Top             =   4560
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Transferencia"
         TabPicture(0)   =   "frmAdeudOperaciones1.frx":10F8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDocTrans"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraTransferencia"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkDocOrigen"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Efectivo"
         TabPicture(1)   =   "frmAdeudOperaciones1.frx":1114
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Shape2"
         Tab(1).Control(1)=   "Label12"
         Tab(1).Control(2)=   "txtBilleteImporte"
         Tab(1).Control(3)=   "cmdEfectivo"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Cheque Recibido"
         TabPicture(2)   =   "frmAdeudOperaciones1.frx":1130
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Shape5"
         Tab(2).Control(1)=   "Label2"
         Tab(2).Control(2)=   "fgChqRecibido"
         Tab(2).Control(3)=   "txtChqRecImporte"
         Tab(2).Control(4)=   "cmdChqRecibido"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Otros "
         TabPicture(3)   =   "frmAdeudOperaciones1.frx":114C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Shape3"
         Tab(3).Control(1)=   "Label18"
         Tab(3).Control(2)=   "fgObj"
         Tab(3).Control(3)=   "fgOtros"
         Tab(3).Control(4)=   "cmdEliminarCta"
         Tab(3).Control(5)=   "cmdAgregarCta"
         Tab(3).Control(6)=   "txtTotalOtrasCtas"
         Tab(3).ControlCount=   7
         Begin VB.CheckBox chkDocOrigen 
            Caption         =   " Documento"
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
            Left            =   450
            TabIndex        =   47
            Top             =   2010
            Width           =   1365
         End
         Begin VB.CommandButton cmdChqRecibido 
            Caption         =   "&Cheques "
            Height          =   315
            Left            =   -74790
            TabIndex        =   19
            Top             =   2400
            Width           =   1395
         End
         Begin VB.TextBox txtChqRecImporte 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H80000012&
            Height          =   285
            Left            =   -68070
            TabIndex        =   20
            Tag             =   "0"
            Text            =   "0.00"
            Top             =   2385
            Width           =   1680
         End
         Begin VB.TextBox txtTotalOtrasCtas 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H80000012&
            Height          =   285
            Left            =   -69270
            TabIndex        =   22
            Tag             =   "3"
            Top             =   1425
            Width           =   1440
         End
         Begin VB.CommandButton cmdAgregarCta 
            Caption         =   "A&gregar"
            Height          =   360
            Left            =   -67500
            TabIndex        =   24
            Top             =   450
            Width           =   1020
         End
         Begin VB.CommandButton cmdEliminarCta 
            Caption         =   "&Eliminar"
            Height          =   360
            Left            =   -67500
            TabIndex        =   25
            Top             =   840
            Width           =   1020
         End
         Begin VB.CommandButton cmdEfectivo 
            Caption         =   "Descomposición de Efectivo"
            Height          =   405
            Left            =   -69540
            TabIndex        =   16
            Top             =   1860
            Width           =   3045
         End
         Begin VB.Frame fraTransferencia 
            Caption         =   "Entidad Financiera"
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
            Left            =   180
            TabIndex        =   49
            Top             =   420
            Width           =   8295
            Begin Sicmact.EditMoney txtBancoImporte 
               Height          =   255
               Left            =   6285
               TabIndex        =   13
               Top             =   1185
               Width           =   1785
               _extentx        =   2937
               _extenty        =   450
               font            =   "frmAdeudOperaciones1.frx":1168
               text            =   "0.00"
               enabled         =   -1
               borderstyle     =   0
            End
            Begin Sicmact.TxtBuscar txtBuscaEntidad 
               Height          =   360
               Left            =   1095
               TabIndex        =   12
               Top             =   300
               Width           =   2580
               _extentx        =   4551
               _extenty        =   635
               appearance      =   1
               appearance      =   1
               font            =   "frmAdeudOperaciones1.frx":1194
               appearance      =   1
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Importe"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   5190
               TabIndex        =   53
               Top             =   1200
               Width           =   615
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H8000000C&
               Height          =   315
               Left            =   4980
               Top             =   1155
               Width           =   3105
            End
            Begin VB.Label lblDesCtaIfTransf 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   1095
               TabIndex        =   52
               Top             =   720
               Width           =   6990
            End
            Begin VB.Label lblDescIfTransf 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   3735
               TabIndex        =   51
               Top             =   300
               Width           =   4350
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta N° :"
               Height          =   210
               Left            =   180
               TabIndex        =   50
               Top             =   360
               Width           =   810
            End
         End
         Begin VB.Frame fraDocTrans 
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
            Height          =   675
            Left            =   150
            TabIndex        =   48
            Top             =   2040
            Width           =   4560
            Begin VB.OptionButton optDoc 
               Caption         =   "Carta"
               Height          =   345
               Index           =   1
               Left            =   2310
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   240
               Width           =   2055
            End
            Begin VB.OptionButton optDoc 
               Caption         =   "Cheque"
               Height          =   345
               Index           =   0
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   2055
            End
         End
         Begin Sicmact.FlexEdit fgChqRecibido 
            Height          =   1725
            Left            =   -74880
            TabIndex        =   18
            Top             =   540
            Width           =   8415
            _extentx        =   14843
            _extenty        =   3149
            cols0           =   12
            encabezadosnombres=   "#-Opc-Banco-NroCheque-Fecha-Importe-Cuenta-cAreaCod-cAgeCod-nMovNro-cPersCod-cIFTpo"
            encabezadosanchos=   "0-420-3200-1800-1200-1500-0-0-0-0-0-0"
            font            =   "frmAdeudOperaciones1.frx":11B8
            font            =   "frmAdeudOperaciones1.frx":11DC
            font            =   "frmAdeudOperaciones1.frx":1200
            font            =   "frmAdeudOperaciones1.frx":1224
            font            =   "frmAdeudOperaciones1.frx":1248
            fontfixed       =   "frmAdeudOperaciones1.frx":126C
            columnasaeditar =   "X-1-X-X-X-X-X-X-X-X-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-4-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-L-C-R-L-C-C-C-C-C"
            formatosedit    =   "0-0-0-0-0-2-0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin Sicmact.FlexEdit fgOtros 
            Height          =   945
            Left            =   -74850
            TabIndex        =   21
            Top             =   450
            Width           =   7305
            _extentx        =   12779
            _extenty        =   1667
            cols0           =   4
            highlight       =   1
            encabezadosnombres=   "#-Cuenta-Descripcion-Importe"
            encabezadosanchos=   "300-1800-3500-1300"
            font            =   "frmAdeudOperaciones1.frx":129A
            font            =   "frmAdeudOperaciones1.frx":12BE
            font            =   "frmAdeudOperaciones1.frx":12E2
            font            =   "frmAdeudOperaciones1.frx":1306
            font            =   "frmAdeudOperaciones1.frx":132A
            fontfixed       =   "frmAdeudOperaciones1.frx":134E
            columnasaeditar =   "X-1-X-3"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-0"
            encabezadosalineacion=   "C-L-L-R"
            formatosedit    =   "0-0-0-2"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   300
            rowheight0      =   300
         End
         Begin Sicmact.FlexEdit fgObj 
            Height          =   945
            Left            =   -74850
            TabIndex        =   23
            Top             =   1770
            Width           =   6195
            _extentx        =   10927
            _extenty        =   1667
            cols0           =   8
            highlight       =   2
            allowuserresizing=   1
            encabezadosnombres=   "#-Ord-Código-Descripción-CtaCont-SubCta-ObjPadre-ItemCtaCont"
            encabezadosanchos=   "350-400-1200-3000-0-900-0-0"
            font            =   "frmAdeudOperaciones1.frx":137C
            font            =   "frmAdeudOperaciones1.frx":13A8
            font            =   "frmAdeudOperaciones1.frx":13D4
            font            =   "frmAdeudOperaciones1.frx":1400
            font            =   "frmAdeudOperaciones1.frx":142C
            fontfixed       =   "frmAdeudOperaciones1.frx":1458
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-X-X-X-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-L-C-C-C-C"
            formatosedit    =   "0-0-3-0-0-0-0-0"
            textarray0      =   "#"
            lbbuscaduplicadotext=   -1
            colwidth0       =   345
            rowheight0      =   300
         End
         Begin VB.TextBox txtBilleteImporte 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H80000012&
            Height          =   285
            Left            =   -68190
            TabIndex        =   17
            Tag             =   "0"
            Top             =   2340
            Width           =   1680
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -69330
            TabIndex        =   54
            Top             =   2385
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -69210
            TabIndex        =   56
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label Label18 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -70680
            TabIndex        =   55
            Top             =   1455
            Width           =   615
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   315
            Left            =   -69525
            Top             =   2325
            Width           =   3045
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   315
            Left            =   -69420
            Top             =   2370
            Width           =   3045
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   315
            Left            =   -70890
            Top             =   1410
            Width           =   3105
         End
      End
      Begin MSComctlLib.ListView lstDetalle 
         Height          =   645
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   1138
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuenta"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Objeto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Monto VAC"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Comisión:"
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
         Height          =   195
         Left            =   4440
         TabIndex        =   70
         Top             =   2370
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Interes:"
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
         Height          =   195
         Left            =   2400
         TabIndex        =   68
         Top             =   2370
         Width           =   660
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
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
         Height          =   195
         Left            =   480
         TabIndex        =   66
         Top             =   2370
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Indice VAC :"
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
         Left            =   6240
         TabIndex        =   61
         Top             =   2370
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         Height          =   195
         Left            =   8280
         TabIndex        =   59
         Top             =   2355
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas de CMACT"
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
         Left            =   120
         TabIndex        =   58
         Top             =   -15
         Width           =   1605
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
         Left            =   120
         TabIndex        =   57
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label lblTasaVAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7200
         TabIndex        =   60
         Top             =   2340
         Width           =   1080
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   75
      X2              =   10245
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frmAdeudOperaciones1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMN As Boolean
Dim lsCtaContDebe() As String
Dim lsCtaContHaber() As String
Dim aObj() As String
Dim lbCargar As Boolean
Dim lsGridDH As String
Dim lsPosCtaBusqueda As String
Dim lnTasaVac As Double
Dim lsCtaConcesional As String
Dim oAdeud As DCaja_Adeudados
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim lsCtaOrdenD As String
Dim lsCtaOrdenH As String

'Efectivo
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

'Documento de Transferencia
Dim lsDocumento As String
Dim lnTpoDoc As TpoDoc
Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim ldFechaDoc  As Date
Dim cMoneda As String

'Variable para Refrescar Cuentas a Utilizar
Dim lsIFTpo As String
Dim lnCambiarMonto As Integer
Dim lnContador As Integer 'ALPA20130618*********************************
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub TamanoColumnas()
    'Exit Sub
    lstCabecera.ColumnHeaders(1).Width = 0
    lstCabecera.ColumnHeaders(2).Width = 0
    
    
    lstCabecera.ColumnHeaders(3).Width = 1500 'Entidad
    lstCabecera.ColumnHeaders(4).Width = 0
    
    lstCabecera.ColumnHeaders(5).Text = "Pagaré"
    lstCabecera.ColumnHeaders(5).Width = 1500 'Pagaré
    
    lstCabecera.ColumnHeaders(6).Text = "F. Vcto."
    lstCabecera.ColumnHeaders(6).Width = 1100
    
    lstCabecera.ColumnHeaders(7).Width = 0
    
    lstCabecera.ColumnHeaders(8).Text = "Cuota"
    lstCabecera.ColumnHeaders(8).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(8).Width = 650
    
    lstCabecera.ColumnHeaders(9).Text = "Per"
    lstCabecera.ColumnHeaders(9).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(9).Width = 650
    
    lstCabecera.ColumnHeaders(10).Text = "Int"
    lstCabecera.ColumnHeaders(10).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(10).Width = 600

    lstCabecera.ColumnHeaders(11).Text = "F.Ult.Pago"
    lstCabecera.ColumnHeaders(11).Width = 1100
    
    lstCabecera.ColumnHeaders(12).Text = "F.Ult.Act"
    lstCabecera.ColumnHeaders(12).Width = 1100
     
    lstCabecera.ColumnHeaders(13).Width = 0
    lstCabecera.ColumnHeaders(14).Width = 0
    lstCabecera.ColumnHeaders(15).Width = 0
    lstCabecera.ColumnHeaders(16).Width = 0
    lstCabecera.ColumnHeaders(17).Width = 0
    lstCabecera.ColumnHeaders(18).Width = 0
    
    lstCabecera.ColumnHeaders(19).Text = "Int.Prov."
    lstCabecera.ColumnHeaders(19).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(19).Width = 1000
    
    lstCabecera.ColumnHeaders(20).Width = 0
    
    lstCabecera.ColumnHeaders(21).Text = "Capital"
    lstCabecera.ColumnHeaders(21).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(21).Width = 1000
    
    lstCabecera.ColumnHeaders(22).Width = 0
    
    lstCabecera.ColumnHeaders(23).Text = "Int.Cal."
    lstCabecera.ColumnHeaders(23).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(23).Width = 1000
    
    lstCabecera.ColumnHeaders(24).Width = 0
    
    lstCabecera.ColumnHeaders(25).Text = "Interés"
    lstCabecera.ColumnHeaders(25).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(25).Width = 1000
    
    lstCabecera.ColumnHeaders(26).Width = 0
    
    lstCabecera.ColumnHeaders(27).Text = "Comisión"
    lstCabecera.ColumnHeaders(27).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(27).Width = 1000
    
    lstCabecera.ColumnHeaders(28).Width = 0
    
    lstCabecera.ColumnHeaders(29).Text = "Total"
    lstCabecera.ColumnHeaders(29).Alignment = lvwColumnRight
    lstCabecera.ColumnHeaders(29).Width = 1000
    
    lstCabecera.ColumnHeaders(30).Width = 0
    lstCabecera.ColumnHeaders(31).Width = 0
    
    lstCabecera.ColumnHeaders(32).Text = "L.Cred"
    lstCabecera.ColumnHeaders(32).Width = 0 '700
    
    lstCabecera.ColumnHeaders(33).Text = "L.Cred"
    lstCabecera.ColumnHeaders(33).Width = 0 '2000

    lstCabecera.ColumnHeaders(34).Width = 0

    lstCabecera.ColumnHeaders(35).Width = 0 '1500
    lstCabecera.ColumnHeaders(36).Width = 0
    lstCabecera.ColumnHeaders(37).Width = 0
    lstCabecera.ColumnHeaders(38).Width = 0
    lstCabecera.ColumnHeaders(39).Width = 0
    lstCabecera.ColumnHeaders(40).Width = 0
    
    lstCabecera.ColumnHeaders(43).Width = 1000
     
End Sub


  
Private Sub CalculaTotalRetiros()
    txtImporte = Format(nVal(txtBancoImporte) + nVal(txtBilleteImporte) + nVal(txtChqRecImporte) + nVal(txtTotalOtrasCtas), gsFormatoNumeroView)
End Sub

Private Sub ChkAntes_Click()
    If Me.ChkAntes.value = 1 Then
        Me.txtCanTotal.Visible = True
    Else
        Me.txtCanTotal.Visible = False
    End If
End Sub

Private Sub chkCambiar_Click()
    If chkCambiar.value = 1 Then
        txtMonto.Enabled = True
        txtInteres.Enabled = True
        txtComision.Enabled = True
        lblTotal.Enabled = True
        lnCambiarMonto = 1
    Else
        txtMonto.Enabled = False
        txtInteres.Enabled = False
        txtComision.Enabled = False
        lblTotal.Enabled = False
        lnCambiarMonto = 0
    End If
End Sub

Private Sub chkCancelacion_Click()
    If chkCancelacion.value = 1 Then
        lstCabecera.ColumnHeaders(3).Width = 0
    Else
        lstCabecera.ColumnHeaders(3).Width = 500
    End If
    Me.cmdCalcular.SetFocus
End Sub

Private Sub chkDocOrigen_Click()
    If chkDocOrigen.value = Checked Then
        fraDocTrans.Enabled = True
    Else
        fraDocTrans.Enabled = False
    End If
End Sub

Private Function ValidaDatos() As Boolean
Dim lbMontoDet As Boolean
Dim I As Integer
ValidaDatos = False
    If Val(lblTotal.Text) = 0 Then
        MsgBox "No se seleccionó Adeudado a pagar...", vbInformation, "¡AViso1"
        Exit Function
    End If
    If txtImporte <> lblTotal Then
        MsgBox "Monto a pagar no coincide con Cuota de deuda", vbInformation, "¡Aviso!"
        Exit Function
    End If
    If lstCabecera.ListItems.Count = 0 Then
        MsgBox "No existen Cuentas de para realizar la Operación", vbInformation, "Aviso"
        Me.cmdSalir.SetFocus
        Exit Function
    End If
    If Val(lblTotal) <= 0 Then
        MsgBox "El Monto de operación no Válido", vbInformation, "Aviso"
        lstCabecera.SetFocus
        Exit Function
    End If
    
    If nVal(txtBancoImporte) > 0 And txtBuscaEntidad = "" Then
        MsgBox "Cuenta de Banco no Válida", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
    If txtMovDesc = "" Then
        MsgBox "Ingrese Descripción de Operación !!!", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        Exit Function
    End If
    lbMontoDet = False
ValidaDatos = True
End Function

''''Private Sub cmdAceptar_Click()
''''Dim oDocPago As clsDocPago
''''Dim lsCuentaAho As String
''''
''''Dim lsMovNro As String
''''Dim oCon     As NContFunciones
''''Dim oCaja As nCajaGeneral
''''Dim rsAdeud  As ADODB.Recordset
''''Dim i As Integer
''''On Error GoTo AceptarErr
''''If Not ValidaDatos() Then
''''    Exit Sub
''''End If
''''
''''Set oDocPago = New clsDocPago
''''Set oCon = New NContFunciones
''''Set oCaja = New nCajaGeneral
''''
''''If MsgBox(" Desea Grabar Operación ? ", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
''''    lsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
''''
''''    oCaja.GrabaPagoCuotaAdeudados lsMovNro, gsOpeCod, CDate(txtFecha.Text), txtMovDesc, nVal(txtImporte), _
''''            lstCabecera.SelectedItem.SubItems(18), lstCabecera.SelectedItem.SubItems(3), rsBill, rsMon, _
''''            gdFecSis, txtBuscaEntidad, txtBancoImporte, _
''''            lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, _
''''            fgChqRecibido.GetRsNew, fgOtros.GetRsNew, fgObj.GetRsNew, GetRsNewDeListView(lstDetalle, "0-0-0-2-0-0-2"), _
''''            chkCancelacion.value = vbChecked, Format(lblTasaVAC, "#,##0.00###"), lsCtaConcesional, lstCabecera.SelectedItem.SubItems(21), lsCtaOrdenD, lsCtaOrdenH, lstCabecera.SelectedItem.SubItems(22)
''''
''''            lstCabecera.SelectedItem.ListSubItems(1).ForeColor = &H808000
''''            lstCabecera.SelectedItem.ListSubItems(2).ForeColor = &H808000
''''
''''            For i = 1 To lstDetalle.ListItems.Count
''''                lstDetalle.ListItems(i).ForeColor = &H808000
''''                lstDetalle.ListItems(i).ListSubItems(2).ForeColor = &H808000
''''            Next
''''
''''    ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento, True, False
''''
''''    Set oCaja = Nothing
''''    Set oCon = Nothing
''''    Set oDocPago = Nothing
''''
''''    MsgBox "Pago Efectuado satisfactoriamente", vbInformation, "Aviso"
''''
''''    cmdAceptar.Enabled = False
''''    cmdCancelar.SetFocus
''''
''''End If
''''
''''Exit Sub
''''AceptarErr:
''''    MsgBox "Error N° [" & Err.Number & "] " & TextErr(Err.Description), vbInformation, "Aviso"
''''End Sub

Private Sub cmdAceptar_Click()
    
    Dim oAdeudado As New NCajaAdeudados
    Dim lsListaAsientos As String
    Dim lsImpre As String
    
    Dim nRetorno As Long
    Dim bPagoLote As Integer
    Dim bCancelacionadelan As Boolean
    
    Dim oContImp As NContImprimir
    Set oContImp = New NContImprimir
    
    Dim sErrorAdeudado As String
    'Consideraciones
    ' 1.    Se puede Pagar unicamente cuotas que aun no han sido pagadas.
    '       Es decir Cuotas Registradas y Provisionadas
    
'    If ValidaPagar = True Then
    
        If MsgBox("Desea Pagar Registros?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            
            lsListaAsientos = ""
            Me.cmdAceptar.Enabled = False
                    
            'Cancelación Adelantada GITU
            If Me.ChkAntes.value = 1 Then
                bCancelacionadelan = True
            End If
                
                
            'Pago Lote
            
            Me.MousePointer = 11
            sErrorAdeudado = ""
                    
            'nRetorno = oAdeudado.GrabaPagoAdeudado(gsOpeCod, cmoneda, txtFecha.Text, Val(Me.lblTasaVAC.Caption), lstCabecera, txtMovDesc.Text, Right(gsCodAge, 2), gsCodUser, lsListaAsientos, lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, Trim(txtBuscaEntidad.Text), IIf(chkCancelacion.value = 1, True, False), bCancelacionadelan, Val(Me.txtCanTotal.Text), sErrorAdeudado)
            'ALPA20130618*****************************************************************
            'nRetorno = oAdeudado.GrabaPagoAdeudado(gsOpeCod, cmoneda, txtFecha.Text, Val(Me.lblTasaVAC.Caption), lstCabecera, txtMovDesc.Text, Right(gsCodAge, 2), gsCodUser, lsListaAsientos, lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, Trim(txtBuscaEntidad.Text), IIf(chkCancelacion.value = 1, True, False), bCancelacionadelan, Val(Me.txtCanTotal.Text), sErrorAdeudado, CCur(txtMonto.Text), CCur(txtInteres.Text), CCur(txtComision.Text), CCur(lblTotal.Text), lnCambiarMonto)
            nRetorno = oAdeudado.GrabaPagoAdeudado(gsOpeCod, cMoneda, txtFecha.Text, Val(Me.lblTasaVAC.Caption), lstCabecera, txtMovDesc.Text, Right(gsCodAge, 2), gsCodUser, lsListaAsientos, lnTpoDoc, lsNroDoc, ldFechaDoc, lsNroVoucher, Trim(txtBuscaEntidad.Text), IIf(chkCancelacion.value = 1, True, False), bCancelacionadelan, Val(Me.txtCanTotal.Text), sErrorAdeudado, CCur(txtMonto.Text), CCur(txtInteres.Text), CCur(txtComision.Text), CCur(lblTotal.Text), lnCambiarMonto, chkConcesion.value)
            '*****************************************************************************
            
            If Trim(sErrorAdeudado) <> "" Then
                MsgBox sErrorAdeudado
                Exit Sub
            End If

            'MsgBox Time
            Me.MousePointer = 0
            If nRetorno = 0 Then
                
                If chkCancelacion.value = 1 Then
                    MsgBox "Cancelación Efectuada Satisfactoriamente", vbInformation, "Aviso"
                Else
                    MsgBox "Pago Efectuado Satisfactoriamente", vbInformation, "Aviso"
                        'ARLO20170217
                        Set objPista = New COMManejador.Pista
                        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Pago Efecutado "
                        Set objPista = Nothing
                        '****
                End If
    
                If Len(Trim(lsListaAsientos)) > 0 Then
                    lsImpre = ImprimeAsientosContables(lsListaAsientos, PB, Status, "")
                    EnviaPrevio lsImpre, "ASIENTOS DE PAGO DE ADEUDADOS", gnLinPage, False
                End If
                

'                Blanquea
                cmdAceptar.Enabled = False
            ElseIf nRetorno = 2 Then
                If Len(Trim(lsListaAsientos)) > 0 Then
                    lsImpre = ImprimeAsientosContables(lsListaAsientos, PB, Status, "")
                    EnviaPrevio lsImpre, "ASIENTOS DE PAGO DE ADEUDADOS", gnLinPage, False
                End If
                
                MsgBox "Se efectuaron solo algunas grabaciones de Pago de Adeudados", vbInformation, "Aviso"
                cmdAceptar.Enabled = True
            ElseIf nRetorno = 1 Then
                MsgBox "No se pudo efectuar ninguna grabación de Pago de Adeudados", vbInformation, "Aviso"
                cmdAceptar.Enabled = True
            End If
        
        End If
        
'    End If
    
End Sub



Private Sub BlanqueaTodo()
Dim I As Integer
     

    lstCabecera.ListItems.Clear
    lstDetalle.ListItems.Clear

    txtBuscarCtaIF.Text = ""
    txtCodObjeto.Text = ""
        
    txtBuscaEntidad = ""
    txtBancoImporte = "0.00"
    txtBilleteImporte = "0.00"
    txtTotalOtrasCtas = "0.00"
    
    txtMovDesc = ""
    Set rsBill = Nothing
    Set rsMon = Nothing
    lblDescIfTransf = ""
    lblDesCtaIfTransf = ""
    
    lsDocumento = ""
    lnTpoDoc = -1
    lsNroDoc = ""
    lsNroVoucher = ""
    
    LimpiaControles
    fgOtros.Clear
    fgOtros.Rows = 2
    fgOtros.FormaCabecera
    
    fgObj.Clear
    fgObj.Rows = 2
    fgObj.FormaCabecera
'    txtDias.Text = ""
    'lblTasaVAC.Caption = ""
    lblTotal.Text = ""
    
    Set frmCajaGenEfectivo = Nothing
End Sub

Private Sub LimpiaControles(Optional lsTipo As Integer = 0)
    txtMovDesc = ""
    lblTotal = "0.00"
    txtBilleteImporte = "0.00"
    CargaCuentasGrid ""
    
End Sub


Private Sub cmdAgregarCta_Click()
Dim oOpe As New DOperacion
fgOtros.AdicionaFila
fgOtros.rsTextBuscar = oOpe.EmiteOpeCtasNivel(gsOpeCod, , "4")
fgOtros.SetFocus
Set oOpe = Nothing
End Sub

Private Sub cmdCalcular_Click()
    lsIFTpo = ""
    lnContador = 0
    If ValFecha(Me.txtFecha) = False Then Exit Sub
    If Valida = False Then Exit Sub
    Me.cmdCalcular.Enabled = False
'    CargaBancos CDate(txtFecha)
    MostrarPago
    Me.cmdCalcular.Enabled = True
    'fgInteres.SetFocus
    If lstCabecera.ListItems.Count > 0 Then
        lstCabecera.ListItems(1).Selected = True
'        Total2 False
        fraCabecera.Enabled = False
        fraDetalle.Enabled = True
        cmdAceptar.Enabled = True
        cmdCalcular.Enabled = False
    Else
        cmdAceptar.Enabled = False
    End If
    
End Sub

Private Sub cmdChqRecibido_Click()
Dim oDocRec As New NDocRec
Dim rs As ADODB.Recordset
Dim nRow As Integer
Set rs = New ADODB.Recordset
Set rs = oDocRec.GetChequesNoDepositados(Mid(gsOpeCod, 3, 1))
fgChqRecibido.Rows = 2
fgChqRecibido.Clear
fgChqRecibido.FormaCabecera
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        fgChqRecibido.AdicionaFila
        nRow = fgChqRecibido.row
        fgChqRecibido.TextMatrix(nRow, 2) = rs!banco
        fgChqRecibido.TextMatrix(nRow, 3) = rs!cNroDoc
        fgChqRecibido.TextMatrix(nRow, 4) = rs!Fecha
        fgChqRecibido.TextMatrix(nRow, 5) = rs!nMonto
        fgChqRecibido.TextMatrix(nRow, 5) = rs!Objeto
        fgChqRecibido.TextMatrix(nRow, 6) = rs!cAreaCod
        fgChqRecibido.TextMatrix(nRow, 7) = rs!cAgeCod
        fgChqRecibido.TextMatrix(nRow, 8) = rs!nMovNro
        fgChqRecibido.TextMatrix(nRow, 9) = rs!cPersCod
        fgChqRecibido.TextMatrix(nRow, 10) = rs!cIFTpo
        rs.MoveNext
    Loop
End If
RSClose rs
Set oDocRec = Nothing
End Sub

Private Sub cmdefectivo_Click()
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, 0, Mid(gsOpeCod, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
         Set rsBill = frmCajaGenEfectivo.rsBilletes
         Set rsMon = frmCajaGenEfectivo.rsMonedas
        txtBilleteImporte.Text = Format(frmCajaGenEfectivo.Total, gsFormatoNumeroView)
    Else
        Unload frmCajaGenEfectivo
        Set frmCajaGenEfectivo = Nothing
        txtBilleteImporte.Text = "0.00"
        RSClose rsBill
        RSClose rsMon
        Exit Sub
    End If
    CalculaTotalRetiros
    If rsBill Is Nothing And rsMon Is Nothing Then
        MsgBox "No se Ingreso de Billetaje", vbInformation, "Aviso"
        Exit Sub
    End If
End Sub

Private Sub cmdEliminarCta_Click()
If fgOtros.TextMatrix(fgOtros.row, 0) <> "" Then
   EliminaCuenta fgOtros.TextMatrix(fgOtros.row, 1), fgOtros.TextMatrix(fgOtros.row, 0)
   txtTotalOtrasCtas.Text = Format(fgOtros.SumaRow(3), gsFormatoNumeroView)
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    cmdAceptar.Enabled = False
    cmdCalcular.Enabled = True
    Limpiar
    BlanqueaTodo
    fraCabecera.Enabled = True
    fraDetalle.Enabled = False
    cmdCalcular.SetFocus
End Sub



'Private Sub Command1_Click()
'    Dim I As Integer
'    Dim J As Integer
'
'    For I = 1 To lstCabecera.ListItems.Count
'        For J = 1 To 22
'            If J = 1 Then
'                If fgInteres.TextMatrix(I, J) = lstCabecera.ListItems(I).SubItems(J) Then
'                Else
'                    MsgBox " " & I & ", " & J
'                End If
'            Else
'                If fgInteres.TextMatrix(I, J) = lstCabecera.ListItems(I).SubItems(J) Then
'                Else
'                    MsgBox " " & I & ", " & J
'                End If
'            End If
'        Next
'    Next
'
'End Sub
'
'Private Sub Command2_Click()
'    Dim I As Integer
'    Dim J As Integer
'    Dim K As Integer
'    Dim i1 As Integer
'    Dim i2 As Integer
'    Dim j1 As Integer
'
'    For K = 1 To fgInteres.Rows - 1
'        fgInteres.row = K
'        lstCabecera.ListItems(K).Selected = True
'        Total False
'        Total2 False
'
'        'Valido las cabeceras
'        For J = 1 To 22
'            If J = 1 Then
'                If fgInteres.TextMatrix(K, J) = lstCabecera.ListItems(K).SubItems(J) Then
'                Else
'                    MsgBox " " & K & ", " & J
'                End If
'            Else
'                If fgInteres.TextMatrix(K, J) = lstCabecera.ListItems(K).SubItems(J) Then
'                Else
'                    MsgBox " " & K & ", " & J
'                End If
'            End If
'        Next
'        'Fin validacion
'
'        For I = 1 To lstDetalle.ListItems.Count
'            For J = 1 To 6
'                If J = 1 Then
'                    If fgDetalle.TextMatrix(I, J) = lstDetalle.ListItems(I).SubItems(J) Then
'                    Else
'                        MsgBox " " & I & ", " & J
'                    End If
'                Else
'                    If fgDetalle.TextMatrix(I, J) = lstDetalle.ListItems(I).SubItems(J) Then
'                    Else
'                        MsgBox " " & I & ", " & J
'                    End If
'                End If
'            Next
'        Next
'
'    Next
'End Sub

  
 
Private Sub fgOtros_OnCellChange(pnRow As Long, pnCol As Long)
txtTotalOtrasCtas = Format(fgOtros.SumaRow(3), gsFormatoNumeroView)
CalculaTotalRetiros
End Sub

Private Sub fgOtros_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oCta As New DGeneral
If psDataCod <> "" Then
    AsignaCtaObj psDataCod
    fgOtros.col = 2
End If
Set oCta = Nothing
End Sub

Private Sub fgOtros_RowColChange()
RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
End Sub

Private Sub Form_Activate()
    If lbCargar = False Then
        Unload Me
    End If
End Sub

Public Sub Inicio(psGridDH As String, psPosCtaBusqueda As String)
    lsGridDH = psGridDH
    lsPosCtaBusqueda = psPosCtaBusqueda
    Me.Show 1
End Sub
Private Sub Form_Load()
'edpyme - valida q exista valor vac
 Dim n As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim I As Integer, j As Integer
  
    CentraForm Me
    TabDoc.Tab = 0
    Me.Caption = "  " & gsOpeDesc
    gsSimbolo = gcMN
    If Mid(gsOpeCod, 3, 1) = "2" Then
        gsSimbolo = gcME
    End If
    cMoneda = Mid(gsOpeCod, 3, 1)

    lbCargar = True
    Set oOpe = New DOperacion
    Set oAdeud = New DCaja_Adeudados
    Set oCtaIf = New NCajaCtaIF


    lnTasaVac = oAdeud.CargaIndiceVAC(gdFecSis)
    If lnTasaVac = 0 Then
       MsgBox "Tasa VAC no ha sido definida ", vbInformation, "Aviso"
       lblTasaVAC = 0
    Else
       lblTasaVAC = lnTasaVac
    End If
    
    LimpiaControles

    txtOpeCod = gsOpeCod
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
    Set rs = oOpe.CargaOpeCta(gsOpeCod, "H", "6")
    If Not rs.EOF Then
        lsCtaConcesional = rs!cCtaContCod
    End If
    RSClose rs
    
    lsCtaOrdenD = oOpe.EmiteOpeCta(gsOpeCod, "D", 7)
    lsCtaOrdenH = oOpe.EmiteOpeCta(gsOpeCod, "H", 7)

    'Agregado para Busquedas
     
    Dim oGen As DGeneral
    'Set oOpe = New DOperacion
    Set oGen = New DGeneral
    txtBuscarCtaIF.rs = oOpe.GetRsOpeObj("40" & Mid(gsOpeCod, 3, 1) & "832", "0", , "' and ci.cCtaIFCod LIKE '__" & Mid(gsOpeCod, 3, 1) & "%' and ci.cCtaIFEstado = '" & gEstadoCtaIFActiva)
    CargaCombo cboEstado, oGen.GetConstante(gCGEstadoCtaIF)
    
    Dim oIF As New DCajaCtasIF
    Me.txtCodObjeto.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), "__05" & Mid(gsOpeCod, 3, 1) & "%", 1)
    Set oIF = Nothing

    TamanoColumnas
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oAdeud = Nothing
    Set frmCajaGenEfectivo = Nothing
    Set frmAdeudOperaciones1 = Nothing
    
End Sub

Private Sub lstCabecera_Click()

If Me.lstCabecera.ListItems.Count > 0 Then
    If Me.lstCabecera.SelectedItem.SubItems(2) <> "" Then
       Me.txtMonto.Text = Me.lstCabecera.SelectedItem.SubItems(20)
       Me.txtInteres.Text = Me.lstCabecera.SelectedItem.SubItems(22)
       Me.txtComision.Text = CCur(Me.lstCabecera.SelectedItem.SubItems(26)) + CCur(Me.lstCabecera.SelectedItem.SubItems(52)) 'CON VAC
       If nVal(Me.lstCabecera.SelectedItem.SubItems(48)) <> 0# Then
       
            chkConcesion.Visible = 1 'ALPA20130618*******************
       End If
       'Me.txtComision.Text = Me.lstCabecera.SelectedItem.SubItems(27) SIN VAC
       CalculaSuma
       lnContador = lnContador + 1
    End If
    If Me.lstCabecera.ListItems.Count > 0 Then
        If nVal(Me.lstCabecera.SelectedItem.SubItems(48)) <> 0# Then
            chkConcesion.Visible = 1
            chkConcesion.value = 1
        Else
            chkConcesion.Visible = 0
            chkConcesion.value = 0
        End If
    End If
End If
End Sub
 
Public Function CalculaSuma()
Dim I As Integer
Dim nCantidad As Long
Dim nTotal As Double
nCantidad = 0
nTotal = 0

'For i = 1 To lstCabecera.ListItems.Count
    'If lstCabecera.ListItems(i).Checked = True Then
        
        'nTotal = nTotal + Val(lstCabecera.ListItems(i).SubItems(28))
        nTotal = Val(lstCabecera.ListItems.Item(lstCabecera.SelectedItem.Index).SubItems(28))
        
    '    nCantidad = nCantidad + 1
    'End If
'Next

For I = 1 To lstCabecera.ListItems.Count
    lstCabecera.ListItems(I).Checked = False
Next
lstCabecera.ListItems.Item(lstCabecera.SelectedItem.Index).Checked = True

lblTotal.Text = Format(nTotal, "0.00")
Me.txtBancoImporte = Format(nTotal, "0.00")
End Function

Private Sub lstCabecera_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Me.lstCabecera.ListItems.Count > 0 Then
    If Me.lstCabecera.SelectedItem.SubItems(2) <> "" Then
       Me.txtMonto.Text = Me.lstCabecera.SelectedItem.SubItems(20)
       Me.txtInteres.Text = Me.lstCabecera.SelectedItem.SubItems(22)
       'Me.txtComision.Text = Me.lstCabecera.SelectedItem.SubItems(27)
       Me.txtComision.Text = Me.lstCabecera.SelectedItem.SubItems(26)
       CalculaSuma
    End If
End If
End Sub

Private Sub lstCabecera_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    
    If KeyCode >= 37 And KeyCode <= 40 Then
         
        If KeyCode = 38 Or KeyCode = 40 Then
            If Me.lstCabecera.ListItems.Count > 0 Then
                I = lstCabecera.SelectedItem.Index
                If I > 0 Then
                    If KeyCode = 38 Then
                        If I > 1 Then
                            I = I - 1
                        End If
                    ElseIf KeyCode = 40 Then
                        If Me.lstCabecera.SelectedItem.Index <> Me.lstCabecera.ListItems.Count Then
                            I = I + 1
                        End If
                    End If
                    If Me.lstCabecera.SelectedItem.Text <> "" Then
'                        Call Total2(False, i)
                    End If
                End If
            End If
        ElseIf KeyCode = 37 Or KeyCode = 39 Then
            If Me.lstCabecera.ListItems.Count > 0 Then
                If Me.lstCabecera.SelectedItem.Text <> "" Then
'                    Call Total2(False)
                End If
            End If
        End If
    Else
        If KeyCode = 13 Then
            'lstDetalle.SetFocus
        Else
            KeyCode = 0
        End If
    End If
End Sub

Private Sub lstCabecera_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Me.lstCabecera.ListItems.Count > 0 Then
            If Me.lstCabecera.SelectedItem.Text <> "" Then
                'pili comento esto
                'Call Total2(False)
                'lstDetalle.SetFocus
            End If
        End If
    Else
        KeyAscii = 0
    End If
    
End Sub

Private Sub LstDetalle_DblClick()
     
    If lstDetalle.ListItems.Count > 0 Then
        If Mid(lstDetalle.SelectedItem.SubItems(1), 1, 4) = "2418" Or Mid(lstDetalle.SelectedItem.SubItems(1), 1, 4) = "2428" Then
            
        Else
            txtMonto.Text = lstDetalle.SelectedItem.SubItems(3)
            'pili
            'fraEdicion.Visible = True
            fraDetalle.Enabled = False
            cmdAceptar.Enabled = False
            cmdCancelar.Enabled = False
            txtMonto.SetFocus
        End If
    End If
End Sub

Private Sub lstDetalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lstDetalle.ListItems.Count > 0 Then
            If Mid(lstDetalle.SelectedItem.SubItems(1), 1, 4) = "2418" Or Mid(lstDetalle.SelectedItem.SubItems(1), 1, 4) = "2428" Then
                
            Else
                txtMonto.Text = lstDetalle.SelectedItem.SubItems(3)
                'pili
                'fraEdicion.Visible = True
                fraDetalle.Enabled = False
                cmdAceptar.Enabled = False
                cmdCancelar.Enabled = False
                txtMonto.SetFocus
            End If
        End If
    End If
End Sub

Private Sub OptDoc_Click(Index As Integer)
Dim oDocPago As clsDocPago
    Set oDocPago = New clsDocPago
    If optDoc(0).value Then
        oDocPago.InicioCheque "", True, Mid(txtBuscaEntidad, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, _
                     CCur(txtImporte), gdFecSis, gsNomCmacRUC, txtBuscaEntidad, lblDescIfTransf, _
                     lblDesCtaIfTransf, "", True, , , , Mid(txtBuscaEntidad, 1, 2), Mid(txtBuscaEntidad, 4, 13), Mid(txtBuscaEntidad, 18, 10)
        If oDocPago.vbOk Then
            lsDocumento = oDocPago.vsFormaDoc
            lnTpoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            ldFechaDoc = oDocPago.vdFechaDoc
            txtMovDesc = oDocPago.vsGlosa
            optDoc(0).value = True
        Else
            optDoc(0).value = False
            Exit Sub
        End If
    Else
        oDocPago.InicioCarta "", "", gsOpeCod, gsOpeDesc, txtMovDesc, "", CCur(txtImporte), _
                     gdFecSis, lblDescIfTransf, lblDesCtaIfTransf, gsNomCmac, "", ""
        If oDocPago.vbOk Then
            lsDocumento = oDocPago.vsFormaDoc
            lnTpoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            ldFechaDoc = oDocPago.vdFechaDoc
            txtMovDesc = oDocPago.vsGlosa
            optDoc(1).value = True
        Else
            optDoc(1).value = False
            Exit Sub
        End If
    End If
    Set oDocPago = Nothing
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

End Sub


Private Sub txtBancoImporte_GotFocus()
fEnfoque txtBancoImporte
End Sub

Private Sub txtBancoImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBancoImporte = Format(txtBancoImporte, gsFormatoNumeroView)
    CalculaTotalRetiros
    chkDocOrigen.SetFocus
End If
End Sub

Private Sub txtBancoImporte_LostFocus()
    txtBancoImporte = Format(txtBancoImporte, gsFormatoNumeroView)
    CalculaTotalRetiros
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
lblDesCtaIfTransf = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
If txtBancoImporte.Visible Then
    txtBancoImporte.SetFocus
End If
End Sub

 
Private Sub txtCanTotal_LostFocus()
    Me.txtBancoImporte.Text = Val(Me.txtCanTotal.Text)
    Me.txtInteres.Text = Round(Val(Me.txtCanTotal.Text) - Val(Me.txtMonto.Text), 2)
End Sub

Private Sub txtEdiicionInter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtEdiicionCom.SetFocus
    End If
End Sub

Private Sub txtEdiicionCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       cmdE_Aceptar.SetFocus
    End If
End Sub

'pili
'Private Sub txtDias_GotFocus()
'    fEnfoque txtDias
'End Sub

'Private Sub txtDias_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosEnteros(KeyAscii)
'    If KeyAscii = 13 Then
'        CalculoInteres Int(txtDias)
'        If Me.lstCabecera.ListItems.Count > 0 Then
'            If Me.lstCabecera.SelectedItem.Text <> "" Then
'                Call Total2(False, , False)
'            End If
'        End If
'        lstCabecera.SetFocus
'    End If
'End Sub

'orden giuliana
'Private Sub txtComision_KeyPress(KeyAscii As Integer)
'    'Lo que se va a grabar
'    lstCabecera.ListItems(nFila).SubItems(22) = Format(nInteresGrabar, "0.00")
'    lstCabecera.ListItems(nFila).SubItems(26) = Format(nComisionGrabar, "0.00")
'
'    'Calendario
'    If cMonedaPago = "2" And Mid(cCtaIfCod, 3, 1) = "1" Then 'VAC
'        lstCabecera.ListItems(nFila).SubItems(23) = Format(nInteresGrabar / nVac, "0.00")
'        lstCabecera.ListItems(nFila).SubItems(27) = Format(nComisionGrabar / nVac, "0.00")
'    Else
'        lstCabecera.ListItems(nFila).SubItems(23) = Format(nInteresGrabar, "0.00")
'        lstCabecera.ListItems(nFila).SubItems(27) = Format(nComisionGrabar, "0.00")
'    End If
'
'    'Diferencias Reales
'
'    'Interes Real a Pagar (Diferencia)
'    lstCabecera.ListItems(nFila).SubItems(25) = Format(Val(.lstCabecera.ListItems(nFila).SubItems(23)) - Val(.lstCabecera.ListItems(nFila).SubItems(19)), "0.00")
'
'    If cMonedaPago = "2" And Mid(cCtaIfCod, 3, 1) = "1" Then 'VAC
'        'Interes a Pagar (Diferencia)
'        'InteresxVac - Interes Prov
'        lstCabecera.ListItems(nFila).SubItems(24) = Format((Val(.lstCabecera.ListItems(nFila).SubItems(23)) * nVac) - Val(.lstCabecera.ListItems(nFila).SubItems(18)), "0.00")
'    Else
'        'Interes a Pagar (Diferencia)
'        'Interes - Interes Prov Real
'        lstCabecera.ListItems(nFila).SubItems(24) = Format(Val(.lstCabecera.ListItems(nFila).SubItems(23)) - Val(.lstCabecera.ListItems(nFila).SubItems(19)), "0.00")
'    End If
'
'    'Total
'    lstCabecera.ListItems(nFila).SubItems(29) = Format(Val(.lstCabecera.ListItems(nFila).SubItems(21)) + Val(.lstCabecera.ListItems(nFila).SubItems(23)) + Val(.lstCabecera.ListItems(nFila).SubItems(27)), "0.00")
'    'Total con/sin VAC
'    lstCabecera.ListItems(nFila).SubItems(28) = Format(Val(.lstCabecera.ListItems(nFila).SubItems(20)) + Val(.lstCabecera.ListItems(nFila).SubItems(22)) + Val(.lstCabecera.ListItems(nFila).SubItems(26)), "0.00")
'
'    CalculaSuma
'End Sub
'

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(txtFecha) = False Then
            txtFecha.SetFocus
        Else
            lnTasaVac = oAdeud.CargaIndiceVAC(txtFecha)
            lblTasaVAC = lnTasaVac
            If Me.optBuscar(0).value = True Then
                Me.optBuscar(0).SetFocus
            ElseIf Me.optBuscar(1).value = True Then
                Me.optBuscar(1).SetFocus
            ElseIf Me.optBuscar(2).value = True Then
                Me.optBuscar(2).SetFocus
            End If
        End If
    End If
End Sub
 
Private Sub txtFecha_LostFocus()
    txtFecha_KeyPress 13
End Sub

Private Sub txtInteres_LostFocus()
    Me.lblTotal.Text = Val(Me.txtMonto.Text) + Val(Me.txtInteres.Text) + Val(Me.txtComision.Text)
    Me.txtBancoImporte.Text = Val(Me.txtMonto.Text) + Val(Me.txtInteres.Text) + Val(Me.txtComision.Text)
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case TabDoc.Tab
            Case 0:
                txtBuscaEntidad.SetFocus
            Case 1
                cmdEfectivo.SetFocus
            Case 2
                cmdChqRecibido.SetFocus
            Case 3
                cmdAgregarCta.SetFocus
        End Select
    End If
End Sub

'Private Sub CargaBancos(ldFecha As Date)
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'    Dim N As Integer
'    Dim lnMontoTotal As Currency
'    Dim lnInteres As Currency
'    Dim lnTotal As Integer, i As Integer
'    Dim lnCapital As Currency
'    Dim rsCMP As New ADODB.Recordset
'
'    Dim sCadena As String
'
'    Dim L As ListItem
'
'    If optBuscar(0).value = True Then
'        sCadena = " AND cia.CPERSCOD='" & Mid(txtCodObjeto, 4, 13) & "'"
'    ElseIf optBuscar(1).value = True Then
'        sCadena = " AND cia.CPERSCOD='" & Mid(txtBuscarCtaIF, 4, 13) & "' and cia.cCtaIFCod='" & Right(txtBuscarCtaIF, 7) & "'"
'    End If
'
'    lnTasaVac = oAdeud.CargaIndiceVAC(ldFecha)
'    If lnTasaVac = 0 Then
'        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada", vbQuestion, "Aviso") = vbNo Then
'            Exit Sub
'        End If
'    End If
'    lblTasaVAC = lnTasaVac
'
'
'    'CARGAMOS LOS ADEUDADOS PENDIENTES
'    Set rs = oAdeud.GetAdeudadosProvision(ldFecha, Mid(gsOpeCod, 3, 1), sCadena)
'
'    If rs.BOF Then
'    Else
'        PB.Visible = True
'        PB.Min = 0
'        PB.Max = rs.RecordCount
'        PB.value = 0
'    End If
'
'    lnTotal = rs.RecordCount
'    i = 0
'
'    lstCabecera.ListItems.Clear
'    Do While Not rs.EOF
'
'        PB.value = PB.value + 1
'
'        i = i + 1
'
'        Set L = lstCabecera.ListItems.Add(, , i)
'        L.SubItems(1) = Trim(rs!cPersNombre)    'entidad
'        L.SubItems(2) = Trim(rs!cCtaIFDesc)    'cuenta
'        L.SubItems(3) = Trim(rs!nNroCuota)    ' numero de cuota pendiente
'        'se oculta *
'        L.SubItems(4) = Format(rs!nSaldoCap, "#,#0.00") 'Saldocapital
'
'        If CInt(rs!bMalPg) = 1 Then
'            Set rsCMP = oAdeud.GetCuotaMalPagador(rs!cPersCod, rs!cIFTpo, rs!cCtaIfCod)
'
'            If Not RSVacio(rsCMP) Then
'                lnCapital = IIf(Me.chkCancelacion.value = 1, rs!nSaldoCap, rs!nCapitalCuota + CDbl(rsCMP!nCapital / 6))
'                'se oculta *
'                L.SubItems(5) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Base
'                'Se muestra
'                If rs!cMonedaPago = "2" And Mid(rs!cCtaIfCod, 3, 1) = "1" Then
'                    L.SubItems(6) = Format(lnCapital * lnTasaVac, "#,#0.00") ' Saldo * la tasa vac
'                Else
'                    L.SubItems(6) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Normal
'                End If
'                'se oculta *
'
'                L.SubItems(7) = Format(rs!nInterespagado + CDbl(rsCMP!nInteres / 6), "#,#0.00") ' Interes acumulado pagado por cuota
'                L.SubItems(8) = Format(rs!nInterespagado + CDbl(rsCMP!nInteres / 6), "#,#0.00")
'
'                If Val(rs!cIFTpo) = gTpoIFFuenteFinanciamiento Then
'                    lnMontoTotal = rs!nSaldoCap + CDbl(rsCMP!nCapital / 6) - rs!nSaldoConcesion
'                Else
'                    lnMontoTotal = rs!nSaldoCap + CDbl(rsCMP!nCapital / 6) + rs!nInterespagado
'                End If
'                lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo, rs!nPeriodo, rs!nInteres, lnMontoTotal)
'
'                'se oculta
'                L.SubItems(9) = Format(lnInteres, "#,#0.00")
'                'se muestra
'                If rs!cMonedaPago = "2" And Mid(rs!cCtaIfCod, 3, 1) = "1" Then
'                    lnInteres = lnInteres * lnTasaVac
'                End If
'                L.SubItems(10) = Format(lnInteres, "#,#0.00")
'                L.SubItems(15) = rs!nComision + CDbl(rsCMP!nComision / 6)
'            Else
'                MsgBox " Este Adeudo no posee cuotas concesionales," & vbCrLf & " verifique Cuota Mal Pagador ", vbInformation, "Aviso"
'            End If
'        Else
'            lnCapital = IIf(Me.chkCancelacion.value = 1, rs!nSaldoCap, rs!nCapitalCuota)
'            'se oculta *
'            L.SubItems(5) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Base
'            'Se muestra
'            If rs!cMonedaPago = "2" And Mid(rs!cCtaIfCod, 3, 1) = "1" Then
'                L.SubItems(6) = Format(lnCapital * lnTasaVac, "#,#0.00") ' Saldo * la tasa vac
'            Else
'                L.SubItems(6) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Normal
'            End If
'            'se oculta *
'
'            L.SubItems(7) = Format(rs!nInterespagado, "#,#0.00")  ' Interes acumulado pagado por cuota
'            L.SubItems(8) = Format(rs!nInterespagado, "#,#0.00")
'
'            If Val(rs!cIFTpo) = gTpoIFFuenteFinanciamiento Then
'                lnMontoTotal = rs!nSaldoCap - rs!nSaldoConcesion
'            Else
'                lnMontoTotal = rs!nSaldoCap + rs!nInterespagado
'            End If
'            lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo, rs!nPeriodo, rs!nInteres, lnMontoTotal)
'
'            'se oculta
'            L.SubItems(9) = Format(lnInteres, "#,#0.00")
'            'se muestra
'            If rs!cMonedaPago = "2" And Mid(rs!cCtaIfCod, 3, 1) = "1" Then
'                lnInteres = lnInteres * lnTasaVac
'            End If
'            L.SubItems(10) = Format(lnInteres, "#,#0.00")
'            L.SubItems(15) = rs!nComision
'        End If
'        'se oculta *
'
'
'        L.SubItems(11) = Format(rs!nInterespagado + lnInteres, "#,#0.00")
'        L.SubItems(12) = Format((rs!nInterespagado + lnInteres), "#0.00")
'
'        L.SubItems(13) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
'
'        L.ListSubItems.Item(13).ForeColor = vbRed
'
'        L.SubItems(14) = Trim(rs!nPeriodo)
'        L.SubItems(20) = Trim(rs!nInteres)
'        L.SubItems(16) = Trim(rs!nDiasUltPAgo)
'        L.SubItems(17) = Trim(rs!cMonedaPago)
'        L.SubItems(18) = Trim(rs!cIFTpo & "." & rs!cPersCod & "." & rs!cCtaIfCod)
'        L.SubItems(19) = rs!dVencimiento
'
'        L.SubItems(21) = rs!nSaldoCapLP
'        L.SubItems(22) = rs!cCodLinCred
'        rs.MoveNext
'    Loop
'
'    PB.Visible = False
'    PB.value = 0
'
'    rs.Close
'    Set rs = Nothing
'End Sub

'Private Function Total2(lbTotal As Boolean, Optional NuevoI As Integer = 0, Optional pVerificar As Boolean = True) As Currency
'    Dim i       As Integer
'    Dim lnTotal As Currency
'    Dim oAdeu   As New DCaja_Adeudados
'    Dim lnFilaActual As Integer
'
'    If NuevoI > 0 Then
'        lnFilaActual = NuevoI
'    Else
'        lnFilaActual = lstCabecera.SelectedItem.Index
'    End If
'
'    If pVerificar = True Then
'    'If Not lsIFTpo = Left(lstCabecera.ListItems(lnFilaActual).SubItems(18), 2) Or lsIFTpo = "" Then
'    '    lsIFTpo = Left(lstCabecera.ListItems(lnFilaActual).SubItems(18), 2)
'        CargaCuentasGrid lstCabecera.ListItems(lnFilaActual).SubItems(3)
'    End If
'    'End If
'
'    If lbTotal = False Then
'        For i = 1 To Me.lstDetalle.ListItems.Count
'            lstDetalle.ListItems(i).SubItems(5) = Trim(lstCabecera.ListItems(lnFilaActual).SubItems(18))
'            Select Case lstDetalle.ListItems(i).SubItems(4)
'                Case "0"
'                    lstDetalle.ListItems(i).SubItems(3) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(6), "#0.00")   'capital que se muestra
'                    lstDetalle.ListItems(i).SubItems(6) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(5), "#0.00")  'capital que se oculta
'                Case "1"
'                    lstDetalle.ListItems(i).SubItems(3) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(10), "#0.00")
'                    lstDetalle.ListItems(i).SubItems(6) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(9), "#0.00")
'                Case "2"
'                    lstDetalle.ListItems(i).SubItems(3) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(8), "#0.00")
'                    lstDetalle.ListItems(i).SubItems(6) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(7), "#0.00")
'                Case "3"  'Comision
'                    lstDetalle.ListItems(i).SubItems(3) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(15), "#0.00")
'                    lstDetalle.ListItems(i).SubItems(6) = Format(lstCabecera.ListItems(lnFilaActual).SubItems(15), "#0.00")
'            End Select
'        Next
''        txtDias = lstCabecera.ListItems(lnFilaActual).SubItems(39)
'        Me.lblTasaVAC.Caption = oAdeu.CargaIndiceVAC(Me.txtFecha)
'    End If
'    lnTotal = 0
'    For i = 1 To Me.lstDetalle.ListItems.Count
'        lnTotal = lnTotal + CCur(IIf(lstDetalle.ListItems(i).SubItems(3) = "", "0", lstDetalle.ListItems(i).SubItems(3)))
'    Next
'    If lstCabecera.ListItems(lnFilaActual).SubItems(17) = "2" And Mid(lstCabecera.ListItems(lnFilaActual).SubItems(18), 20, 1) = "1" Then
'       lblTasaVAC = lnTasaVac
'    End If
'
'End Function
 
Private Sub CargaCuentasGrid(psIFCod As String)
Dim rs As ADODB.Recordset
Dim oOpe As New DOperacion
Dim I As Integer
Dim L As ListItem

If Not psIFCod = "" Then
    I = 0
    
    Set rs = oOpe.CargaOpeCtaIF(gsOpeCod, psIFCod, "D")
    If Not RSVacio(rs) Then
         
        lstDetalle.ListItems.Clear
        Do While Not rs.EOF
            I = I + 1
            Set L = lstDetalle.ListItems.Add(, , I)
            L.SubItems(1) = Trim(rs!cCtaContCod)
            L.SubItems(2) = Trim(rs!cCtaContDesc)
            L.SubItems(4) = Trim(rs!cOpeCtaOrden)
            
            
            If Mid(rs!cCtaContCod, 1, 4) = "2418" Or Mid(rs!cCtaContCod, 1, 4) = "2428" Then
                L.ListSubItems.Item(1).ForeColor = vbRed
            Else
                L.ListSubItems.Item(1).ForeColor = vbBlack
            End If
            
            rs.MoveNext
        Loop
                 
    Else
        lbCargar = False
        MsgBox "No se han definido Cuentas Contables para Operación", vbInformation, "Aviso"
    End If
End If
 

RSClose rs
Set oOpe = Nothing

End Sub

Private Sub CalculoInteres(ByVal lnDias As Long)
    Dim lnPeriodo As Long
    Dim lnTasaInt As Currency
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency

    If lstCabecera.ListItems.Count > 0 Then
        lstCabecera.SelectedItem.SubItems(16) = lnDias
        If Val(Mid(lstCabecera.SelectedItem.SubItems(18), 1, 2)) = gTpoIFFuenteFinanciamiento Then
            lnMontoTotal = CCur(lstCabecera.SelectedItem.SubItems(4))
        Else
            lnMontoTotal = CCur(lstCabecera.SelectedItem.SubItems(4)) + CCur(lstCabecera.SelectedItem.SubItems(7))
        End If
        lnTasaInt = CCur(lstCabecera.SelectedItem.SubItems(20))
        lnPeriodo = Val(lstCabecera.SelectedItem.SubItems(14))
        lnInteres = oAdeud.CalculaInteres(lnDias, lnPeriodo, lnTasaInt, lnMontoTotal)
    
        lstCabecera.SelectedItem.SubItems(9) = Format(lnInteres, "#0.00")
        If lstCabecera.SelectedItem.SubItems(17) = "2" And Mid(lstCabecera.SelectedItem.SubItems(18), 9, 1) = "1" Then
            lstCabecera.SelectedItem.SubItems(10) = Format(lnInteres * lnTasaVac, "#0.00")
        Else
            lstCabecera.SelectedItem.SubItems(10) = Format(lnInteres, "#0.00")
        End If
    End If
End Sub

Public Sub AsignaCtaObj(ByVal psCtaContCod As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim oContFunct As NContFunciones

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set oContFunct = New NContFunciones

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case Val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                lsRaiz = "Unidades Organizacionales"
                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                lsRaiz = "Cuentas de Entidades Financieras"
                Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
            Case ObjDescomEfectivo
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
            Case ObjPersona
                Set rs = Nothing
            Case Else
                lsRaiz = "Varios"
                Set rs = GetObjetos(Val(rs1!cObjetoCod))
        End Select
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.Show rs, "", lsRaiz
                        If oDescObj.lbOk Then
                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.row, 0), rs1!nCtaObjOrden, oDescObj.gsSelecCod, _
                                        oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                        Else
                            fgOtros.EliminaFila fgOtros.row, False
                            Exit Do
                        End If
                    Else
                        AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.row, 0), rs1!nCtaObjOrden, rs1!cObjetoCod, _
                                        rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    AdicionaObj psCtaContCod, fgOtros.TextMatrix(fgOtros.row, 0), rs1!nCtaObjOrden, _
                                    UP.sPersCod, UP.sPersNombre, _
                                    lsFiltro, rs1!cObjetoCod
                End If
                Set frmBuscaPersona = Nothing
            End If
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
Set oContFunct = Nothing
End Sub

Private Sub AdicionaObj(sCodCta As String, nFila As Integer, _
                        psOrden As String, psObjetoCod As String, psObjDescripcion As String, _
                        psSubCta As String, psObjPadre As String)
Dim nItem As Integer
    fgObj.AdicionaFila
    nItem = fgObj.row
    fgObj.TextMatrix(nItem, 0) = nFila
    fgObj.TextMatrix(nItem, 1) = psOrden
    fgObj.TextMatrix(nItem, 2) = psObjetoCod
    fgObj.TextMatrix(nItem, 3) = psObjDescripcion
    fgObj.TextMatrix(nItem, 4) = sCodCta
    fgObj.TextMatrix(nItem, 5) = psSubCta
    fgObj.TextMatrix(nItem, 6) = psObjPadre
    fgObj.TextMatrix(nItem, 7) = nFila
    
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
If fgOtros.TextMatrix(1, 0) <> "" Then
    EliminaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
    fgOtros.EliminaFila fgOtros.row
End If
If Len(fgOtros.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj Val(fgOtros.TextMatrix(fgOtros.row, 0))
End If
End Sub

Private Sub RefrescaFgObj(nItem As Integer)
Dim K  As Integer
For K = 1 To fgObj.Rows - 1
    If Len(fgObj.TextMatrix(K, 1)) Then
       If fgObj.TextMatrix(K, 0) = nItem Then
          fgObj.RowHeight(K) = 285
       Else
          fgObj.RowHeight(K) = 0
       End If
    End If
Next
End Sub


Private Sub EliminaFgObj(nItem As Integer)
Dim K  As Integer, m As Integer
K = 1
Do While K < fgObj.Rows
   If Len(fgObj.TextMatrix(K, 0)) > 0 Then
      If Val(fgObj.TextMatrix(K, 0)) = nItem Then
         fgObj.EliminaFila K, False
      Else
        If CCur(fgObj.TextMatrix(K, 0)) > nItem Then
           fgObj.TextMatrix(K, 0) = CCur(fgObj.TextMatrix(K, 0)) - 1
        End If
         K = K + 1
      End If
   Else
      K = K + 1
   End If
Loop
End Sub


Private Sub optBuscar_Click(Index As Integer)
fraopciones.Visible = IIf(Index = 0, True, False)
FraGenerales.Visible = IIf(Index = 1, True, False)
Limpiar
BlanqueaTodo
End Sub

Private Sub optBuscar_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        txtCodObjeto.SetFocus
    ElseIf Index = 1 Then
        txtBuscarCtaIF.SetFocus
    Else
        cmdCalcular.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarCtaIF_EmiteDatos()
If txtBuscarCtaIF <> "" Then
    Set frmAdeudCal = Nothing
     
    CargaDatosCuentas Mid(txtBuscarCtaIF, 4, 13), Mid(txtBuscarCtaIF, 1, 2), Mid(txtBuscarCtaIF, 18, 10)
    cmdCalcular.SetFocus
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
txtImporte.Text = "0.00"
txtBancoImporte.Text = "0.00"
txtBilleteImporte.Text = "0.00"
txtChqRecImporte.Text = "0.00"
txtTotalOtrasCtas.Text = "0.00"
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
    cboEstado = rs!cEstadoCons & space(50) & rs!cCtaIFEstado
End If
RSClose rs
Set oCtaIf = Nothing

End Sub
Private Sub txtCodObjeto_EmiteDatos()
If txtCodObjeto <> "" Then
    lblObjDesc = txtCodObjeto.psDescripcion
End If
cmdCalcular.SetFocus
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

'''''''''''''''''
 
Private Sub txtmonto_GotFocus()
    Call fEnfoque(txtMonto)
End Sub
'orden giuliana
'Private Sub txtMonto_KeyPress(KeyAscii As Integer)
'
'
'    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 12, 2, False)
'    If KeyAscii = 13 Then
'
'        lstDetalle.SelectedItem.SubItems(3) = Format(Me.txtMonto.Text, "#0.00")
'
'        If lstCabecera.SelectedItem.SubItems(17) = "2" And Mid(lstCabecera.SelectedItem.SubItems(18), 20, 1) = "1" Then
'            If nVal(lblTasaVAC) > 0 Then
'               lstDetalle.SelectedItem.SubItems(6) = Format(nVal(lstDetalle.SelectedItem.SubItems(3)) / Format(lblTasaVAC, "#,#0.00###"), "#0.00")
'            End If
'        Else
'            lstDetalle.SelectedItem.SubItems(6) = Format(nVal(lstDetalle.SelectedItem.SubItems(3)), "#0.00")
'        End If
'
'        Total2 True, , False
'
''        fraEdicion.Visible = False
'        fraDetalle.Enabled = True
'        cmdAceptar.Enabled = True
'        cmdCancelar.Enabled = True
'
'        Me.txtMonto.Text = Me.lstCabecera.SelectedItem.SubItems(20)
'        Me.txtComision.Text = Me.lstCabecera.SelectedItem.SubItems(27)
'        Me.txtInteres.Text = Me.lstCabecera.SelectedItem.SubItems(25)
'
'        txtMonto.Text = ""
'
'        lstDetalle.SetFocus
'    End If
'End Sub
'
Private Sub txtMonto_LostFocus()
    txtMonto.Text = Format(txtMonto.Text, "#0.00")
End Sub


''''Private Sub MostrarPago()
''''    Dim i As Integer
''''    Dim J As Integer
''''    Dim sCadena As String
''''    Dim nContador As Integer
''''
''''    Dim oAdeud As New NCajaAdeudados '  DACGAdeudados
''''    Dim rs As ADODB.Recordset
''''    Dim rsTemp As ADODB.Recordset
''''    Dim L As ListItem
''''
''''    Dim nTasaMensual As Double
''''    Dim lnTotal As Integer
''''    Dim lnSaldoTotal As Double
''''    Dim lnInteres As Double
''''
''''    Dim lnTasaVacTempo As Double
''''    Dim lsPersCod As String
''''    Dim lsCtaIFCod As String
''''
''''    ' ==================== ENCABEZADOS ====================
''''
''''    ' 2 => Entidad Financiera
''''
''''    ' ==================== INICIALIZAMOS VALORES ====================
''''
''''    lstCabecera.ListItems.Clear
''''
''''    If optBuscar(0).value = True Then
''''        sCadena = " AND CIF.cPersCod='" & Mid(Me.txtCodObjeto.Text, 4, 13) & "'"
''''        lsPersCod = Mid(Me.txtBuscaEntidad.Text, 4, 13)
''''    ElseIf optBuscar(1).value = True Then
''''        sCadena = " AND CIF.cPersCod='" & Mid(Me.txtBuscarCtaIF.Text, 4, 13) & "' and CIF.cCtaIFCod='" & Right(Me.txtBuscarCtaIF.Text, 7) & "'"
''''        lsPersCod = Mid(Me.txtBuscarCtaIF.Text, 4, 13)
''''        lsCtaIFCod = Right(Me.txtBuscarCtaIF.Text, 7)
''''    End If
''''
''''    Dim oAdeud1 As New DCaja_Adeudados
''''    lnTasaVac = oAdeud1.CargaIndiceVAC(Me.txtFecha.Text)
''''
''''    If lnTasaVac = 0 Then
''''        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada" & Chr(13) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
''''            Exit Sub
''''        End If
''''    End If
''''
''''    'Set rs = oAdeud.GetDatosAdeudadoPago(CDate(mskFechaMovimiento.Text), Val(cMoneda), sCadena, IIf(chkCancelacion.value = 1 Or chkPrePago.value = 1, True, False), IIf(chkRestringir.value = 1, True, False))
''''    Set rs = oAdeud1.GetAdeudadosProvision(Me.txtFecha.Text, Mid(gsOpeCod, 3, 1), sCadena)
''''    If rs.BOF Then
''''        cmdAceptar.Enabled = False
''''        lnTotal = 0
''''    Else
''''        cmdAceptar.Enabled = True
''''        lnTotal = rs.RecordCount
''''        PB.Visible = True
''''        PB.Min = 0
''''        PB.Max = rs.RecordCount
''''        PB.value = 0
''''    End If
''''
''''    'lblCantidad.Caption = lnTotal
''''
''''    ' ==================== LLENAMOS MATRIZ ====================
''''
''''    i = 0
''''    lstCabecera.ListItems.Clear
''''
''''    If rs.BOF Then
''''        MsgBox "No hay datos que mostrar", vbInformation, "Aviso"
''''    Else
''''        Do While Not rs.EOF
''''
''''            PB.value = PB.value + 1
''''            Status.Panels(1).Text = "Proceso " & Format(PB.value * 100 / PB.Max, gsFormatoNumeroView) & "%"
''''
''''            i = i + 1
''''
''''            Set L = lstCabecera.ListItems.Add(, , "")
''''
''''            L.SubItems(2) = Trim(rs!cPersNombre)
''''            L.SubItems(3) = rs!cIFTpo & "." & rs!cPersCod & "." & rs!cCtaIFCod
''''            L.SubItems(4) = rs!cCtaIFDesc
''''            L.SubItems(5) = Format(rs!dVencimiento, "dd/MM/YYYY")
''''            L.SubItems(6) = rs!cTpoCuota
''''            L.SubItems(7) = rs!nNroCuota
''''
''''            L.SubItems(8) = Trim(rs!nCtaIFIntPeriodo)
''''            L.SubItems(9) = Trim(rs!nCtaIFIntValor)
''''
''''            L.SubItems(10) = Format(rs!dCuotaUltPago, "dd/MM/YYYY")
''''            L.SubItems(11) = Format(rs!dCuotaUltPago, "dd/MM/YYYY")
''''
''''            L.SubItems(30) = Trim(rs!cMonedaPago)
''''            L.SubItems(31) = Trim(rs!cCodLinCred)
''''            L.SubItems(32) = Trim(rs!cdeslincred)
''''
''''            'Cuota 6
''''            Set rsTemp = oAdeud.GetDatosCuota6(rs!cPersCod, rs!cIFTpo, rs!cCtaIFCod, rs!dVencimiento, "'0', '2'", IIf(chkCancelacion.value = 1, True, False))
''''            If rsTemp.BOF Then
''''                L.SubItems(33) = ""
''''            Else
''''                L.SubItems(33) = rsTemp!nNroCuota
''''            End If
''''            rsTemp.Close
''''
''''            L.SubItems(39) = Trim(rs!cEstadoCuota)
''''
''''            'SK Real
''''            L.SubItems(12) = Format(rs!nSaldoCap, "0.00")
''''            'SK LP Real
''''            L.SubItems(15) = Format(rs!nSaldoCapLP, "0.00")
''''
''''            'Interes Provisionado Real
''''            L.SubItems(19) = Format(rs!nInteresProvisionadoReal, "0.00")
''''
''''            'Capital Calendario Pagar Real
''''            L.SubItems(21) = Format(rs!nCapital, "0.00")
''''
''''            'Interes Calendario Pagar Real
''''            L.SubItems(23) = Format(rs!nInteres, "0.00")
''''
''''            'Interes Real a Pagar (Diferencia)
''''            L.SubItems(25) = Format(rs!nInteres - rs!nInteresProvisionadoReal, "0.00")
''''
''''            'Comision Calendario Pagar Real
''''            L.SubItems(27) = Format(rs!nComision, "0.00")
''''
''''            'Total Cuota Calendario Pagar Real
''''            L.SubItems(29) = Format(rs!nTotalCuota, "0.00")
''''
''''            'VAC SI HUBIERA
''''
''''            If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
''''
''''                'lnTasaVac
''''
''''                'SK x Vac Actual
''''                L.SubItems(14) = Format(rs!nSaldoCap * lnTasaVac, "0.00")
''''
''''                'SK LP x Vac Actual
''''                L.SubItems(17) = Format(rs!nSaldoCapLP * lnTasaVac, "0.00")
''''
''''                'Capital Calendario Pagar
''''                L.SubItems(20) = Format(rs!nCapital * lnTasaVac, "0.00")
''''
''''                'Interes Calendario Pagar
''''                L.SubItems(22) = Format(rs!nInteres * lnTasaVac, "0.00")
''''
''''                'Comision Calendario Pagar
''''                L.SubItems(26) = Format(rs!nComision * lnTasaVac, "0.00")
''''
''''                'Total Cuota Calendario Pagar
''''                L.SubItems(28) = Format(rs!nTotalCuota * lnTasaVac, "0.00")
''''
''''                ' ==================== VAC ANTERIOR ====================
''''
''''                'VAC Anterior
''''                lnTasaVacTempo = oAdeud1.CargaIndiceVAC(rs!dCuotaUltPago)
''''
''''                'SK x Vac Anterior
''''                L.SubItems(13) = Format(rs!nSaldoCap * lnTasaVacTempo, "0.00")
''''
''''                'SK LP x Vac Anterior
''''                L.SubItems(16) = Format(rs!nSaldoCapLP * lnTasaVacTempo, "0.00")
''''
''''                'Interes Provisionado (El grabado ya esta por vac)
''''                L.SubItems(18) = Format(rs!nInterespagado, "0.00")
''''
''''
''''                'Interes a Pagar (Diferencia)
''''                L.SubItems(24) = Format((rs!nInteres * lnTasaVac) - rs!nInterespagado, "0.00")
''''
''''
''''            Else
''''
''''                'SK
''''                L.SubItems(14) = Format(rs!nSaldoCap, "0.00")
''''
''''                'SK LP
''''                L.SubItems(17) = Format(rs!nSaldoCapLP, "0.00")
''''
''''                'Capital Calendario Pagar
''''                L.SubItems(20) = Format(rs!nCapital, "0.00")
''''
''''                'Interes Calendario Pagar
''''                L.SubItems(22) = Format(rs!nInteres, "0.00")
''''
''''                'Comision Calendario Pagar
''''                L.SubItems(26) = Format(rs!nComision, "0.00")
''''
''''                'Total Cuota Calendario Pagar
''''                L.SubItems(28) = Format(rs!nTotalCuota, "0.00")
''''
''''                ' ==================== VAC ANTERIOR ====================
''''
''''                'SK
''''                L.SubItems(13) = Format(rs!nSaldoCap, "0.00")
''''
''''                'SK LP
''''                L.SubItems(16) = Format(rs!nSaldoCapLP, "0.00")
''''
''''                'Interes Provisionado
''''                L.SubItems(18) = Format(rs!nInteresProvisionadoReal, "0.00")
''''
''''
''''                'Interes a Pagar (Diferencia)
''''                L.SubItems(24) = Format(rs!nInteres - rs!nInteresProvisionadoReal, "0.00")
''''
''''            End If
''''            L.SubItems(39) = Format(rs!nDiasUltPAgo, "000")
''''            rs.MoveNext
''''        Loop
''''        RSClose rs
''''
''''        cmdAceptar.Enabled = True
''''
''''    End If
''''
''''    Set rs = Nothing
''''    Set oAdeud = Nothing
''''
''''    Status.Panels(1).Text = "Proceso " & Format(PB.Max * 100 / PB.Max, gsFormatoNumeroView) & "%"
''''
''''    PB.Visible = False
''''    PB.value = 0
''''
'''''    txtTasaVac.Enabled = False
''''
''''End Sub

Private Sub MostrarPago()
    Dim I As Integer
    Dim j As Integer
    Dim sCadena As String
    Dim sNumCadena As Integer
    
    Dim oAdeud As New NCajaAdeudados
    Dim oAdeud1 As New DCaja_Adeudados
    
    Dim rs As ADODB.Recordset
    Dim RSTEMP As ADODB.Recordset
    Dim L As ListItem
    
    Dim nTasaMensual As Double
    Dim lnTotal As Integer
    Dim lnSaldoTotal As Double
    Dim lnInteres As Double
    
    Dim lnTasaVacTempo As Double
    Dim lsPersCod As String
    Dim lsCtaIFCod As String
    Dim nContador As Integer
    Dim lnSaldoCapitalTotal As Currency
    Dim lnCapitalTotalCuota As Currency
    Dim lnInteresTotalCuota As Currency
    ' ==================== ENCABEZADOS ====================
    
    ' 2 => Entidad Financiera
     
    ' ==================== INICIALIZAMOS VALORES ====================
    
    lstCabecera.ListItems.Clear
    
    If optBuscar(0).value = True Then
        sCadena = " AND CIF.cPersCod='" & Mid(txtCodObjeto, 4, 13) & "'"
        lsPersCod = Mid(txtCodObjeto, 4, 13)
        lsCtaIFCod = ""
        sNumCadena = 1 'PEAC 20210722
    ElseIf optBuscar(1).value = True Then
        sCadena = " AND CIF.cPersCod='" & Mid(Me.txtBuscarCtaIF, 4, 13) & "' and CIF.cCtaIFCod='" & Right(txtBuscarCtaIF, 7) & "'"
        lsPersCod = Mid(txtBuscarCtaIF, 4, 13)
        lsCtaIFCod = Right(txtBuscarCtaIF, 7)
        sNumCadena = 2 'PEAC 20210722

'    ElseIf optBuscar(2).value = True Then
'        If GetWherePagoLote(lblRutaLote.Caption, sCadena) = False Then
'            Exit Sub
'        End If
    End If
     
    'PEAC 20210722
    If sNumCadena = 2 Then
        sCadena = Trim(Str(sNumCadena)) & lsPersCod & lsCtaIFCod '1 13 7
    ElseIf sNumCadena = 1 Then
        sCadena = Trim(Str(sNumCadena)) & lsPersCod '1 13
    End If
     
    lnTasaVac = Format(Me.lblTasaVAC, "0.000000")
    
'    If lnTasaVac = 0 Then
'        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada" & Chr(13) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'            txtTasaVac.SetFocus
'            Exit Sub
'        End If
'    End If
    
    Set rs = oAdeud.GetDatosAdeudadoPago(CDate(Me.txtFecha.Text), Val(cMoneda), sCadena, IIf(chkCancelacion.value = 1, True, False), False)
    
    If rs.BOF Then
        Me.cmdAceptar.Enabled = False
        lnTotal = 0
    Else
        Me.cmdAceptar.Enabled = True
        lnTotal = rs.RecordCount
        PB.Visible = True
        PB.Min = 0
        PB.Max = rs.RecordCount
        PB.value = 0
    End If
     
    'lblCantidad.Caption = lnTotal
     
    ' ==================== LLENAMOS MATRIZ ====================
     
    I = 0
    lstCabecera.ListItems.Clear
    
    If rs.BOF Then
        MsgBox "No hay datos que mostrar", vbInformation, "Aviso"
    Else
        Do While Not rs.EOF
            
            PB.value = PB.value + 1
            Status.Panels(1).Text = "Proceso " & Format(PB.value * 100 / PB.Max, gsFormatoNumeroView) & "%"
            
            I = I + 1
            
            Set L = lstCabecera.ListItems.Add(, , "")
            'ALPA20130618*********************************************
            lnSaldoCapitalTotal = rs!nSaldoCap + rs!nSaldoCapConce
            lnCapitalTotalCuota = rs!nCapital + rs!nCapitalConce
            lnInteresTotalCuota = rs!nInteres + rs!nInteresConce
            '*********************************************************
'            L.SmallIcon = 1
            L.SubItems(2) = Trim(rs!cPersNombre)
            L.SubItems(3) = rs!cIFTpo & "." & rs!cPersCod & "." & rs!cCtaIfCod
            L.SubItems(4) = rs!cCtaIFDesc
            L.SubItems(5) = Format(rs!dVencimiento, "dd/MM/YYYY")
            L.SubItems(6) = rs!cTpoCuota
            L.SubItems(7) = rs!nNroCuota
            
            L.SubItems(8) = Trim(rs!nCtaIFIntPeriodo)
            L.SubItems(9) = Trim(rs!nCtaIFIntValor)
            
            L.SubItems(10) = Format(rs!dCuotaUltPago, "dd/MM/YYYY")
            L.SubItems(11) = Format(rs!dCuotaUltModSaldos, "dd/MM/YYYY")
             
            L.SubItems(30) = Trim(rs!cMonedaPago)
            L.SubItems(31) = Trim(rs!cCodLinCred)
            L.SubItems(32) = Trim(rs!cDesLinCred)
            L.SubItems(41) = Trim(IIf(IsNull(rs!iVacAper), 0, rs!iVacAper))
            
            'Cuota 6
            Set RSTEMP = oAdeud.GetDatosCuota6(rs!cPersCod, rs!cIFTpo, rs!cCtaIfCod, rs!dVencimiento, "'0', '2'", IIf(chkCancelacion.value = 1, True, False))
            If RSTEMP.BOF Then
                L.SubItems(33) = ""
            Else
                L.SubItems(33) = RSTEMP!nNroCuota
            End If
            RSTEMP.Close
               
            L.SubItems(39) = Trim(rs!cEstadoCuota)
               
            'SK Real
            L.SubItems(12) = Format(lnSaldoCapitalTotal, "0.00") 'Format(rs!nSaldoCap, "0.00")
            'SK LP Real
            L.SubItems(15) = Format(rs!nSaldoCapLP, "0.00")
             
            'Interes Provisionado Real
            L.SubItems(19) = Format(rs!nInteresProvisionadoReal, "0.00")
            
            'Capital Calendario Pagar Real
            L.SubItems(21) = Format(lnCapitalTotalCuota, "0.00") 'Format(rs!nCapital, "0.00")
            
            'Interes Calendario Pagar Real
            L.SubItems(23) = Format(lnInteresTotalCuota, "0.00") 'Format(rs!nInteres, "0.00")
            
            'Interes Real a Pagar (Diferencia)
            L.SubItems(25) = Format(lnInteresTotalCuota - rs!nInteresProvisionadoReal, "0.00") 'Format(rs!nInteres - rs!nInteresProvisionadoReal, "0.00")
            
            'Comision Calendario Pagar Real
            L.SubItems(27) = Format(rs!nComision, "0.00")
            'Comision Calendario Concesionado Pagar Real
            L.SubItems(52) = Format(rs!nComisionConce, "0.00")
            'Total Cuota Calendario Pagar Real
            L.SubItems(29) = Format(rs!ntotalcuota, "0.00")
            
            'VAC SI HUBIERA
     
            If rs!cMonedaPago = "2" And Mid(rs!cCtaIfCod, 3, 1) = "1" Then
                'lnTasaVac
                
                'SK x Vac Actual
                L.SubItems(14) = Format(lnSaldoCapitalTotal * lnTasaVac, "0.00") 'Format(rs!nSaldoCap * lnTasaVac, "0.00")
                
                'SK LP x Vac Actual
                L.SubItems(17) = Format(rs!nSaldoCapLP * lnTasaVac, "0.00")
                 
                'Capital Calendario Pagar
                L.SubItems(20) = Format(lnCapitalTotalCuota, "0.00") 'Format(rs!nCapital * lnTasaVac, "0.00")
                'L.SubItems(20) = Format(rs!nCapital * rs!iVacAper, "0.00")
                
                'Interes Calendario Pagar
                L.SubItems(22) = Format(lnInteresTotalCuota * lnTasaVac, "0.00") 'Format(rs!nInteres * lnTasaVac, "0.00")
                
                'Comision Calendario Pagar
                L.SubItems(26) = Format(rs!nComision * lnTasaVac, "0.00")
                
                'Total Cuota Calendario Pagar
                'L.SubItems(28) = Val(lstCabecera.ListItems(I).SubItems(20)) + Val(lstCabecera.ListItems(I).SubItems(27)) + Val(lstCabecera.ListItems(I).SubItems(22))
                L.SubItems(28) = Val(lstCabecera.ListItems(I).SubItems(20)) + Val(lstCabecera.ListItems(I).SubItems(26)) + Val(lstCabecera.ListItems(I).SubItems(22)) + Val(lstCabecera.ListItems(I).SubItems(52))
                
                'L.SubItems(28) = Format(rs!ntotalcuota * lnTasaVac, "0.00")
    
                ' ==================== VAC ANTERIOR ====================
                
                'VAC Anterior
                lnTasaVacTempo = oAdeud1.CargaIndiceVAC(rs!dCuotaUltModSaldos)
    
                'SK x Vac Anterior
                L.SubItems(13) = Format(lnSaldoCapitalTotal * lnTasaVacTempo, "0.00") 'Format(rs!nSaldoCap * lnTasaVacTempo, "0.00")
                
                'SK LP x Vac Anterior
                L.SubItems(16) = Format(rs!nSaldoCapLP * lnTasaVacTempo, "0.00")
                
                'Interes Provisionado (El grabado ya esta por vac)
                L.SubItems(18) = Format(rs!nInteresProvisionado, "0.00")
                
                
                'Interes a Pagar (Diferencia)
                L.SubItems(24) = Format((lnInteresTotalCuota * lnTasaVac) - rs!nInteresProvisionado, "0.00") 'Format((rs!nInteres * lnTasaVac) - rs!nInteresProvisionado, "0.00")
                
                L.SubItems(42) = Format(lnCapitalTotalCuota * (lnTasaVac - rs!iVacAper), "0.00") 'Format(rs!nCapital * (lnTasaVac - rs!iVacAper), "0.00")
                
                'L.SubItems(28) = Val(lstCabecera.ListItems(i).SubItems(20)) + Val(lstCabecera.ListItems(i).SubItems(27)) + Val(lstCabecera.ListItems(i).SubItems(22)) + Val(lstCabecera.ListItems(i).SubItems(42))
                L.SubItems(48) = Format(rs!nCapitalConce * lnTasaVac, "0.00")
                L.SubItems(49) = Format(rs!nInteresConce * lnTasaVac, "0.00")
                L.SubItems(50) = Format(rs!nSaldoCapConce * lnTasaVac, "0.00")
                L.SubItems(51) = Format(rs!nProvisionConce * lnTasaVac, "0.00")
            Else
    
                'SK
                L.SubItems(14) = Format(lnSaldoCapitalTotal, "0.00") 'Format(rs!nSaldoCap, "0.00")
    
                'SK LP
                L.SubItems(17) = Format(rs!nSaldoCapLP, "0.00")
    
                'Capital Calendario Pagar
                L.SubItems(20) = Format(lnCapitalTotalCuota, "0.00") 'Format(rs!nCapital, "0.00")
                
                'Interes Calendario Pagar
                L.SubItems(22) = Format(lnInteresTotalCuota, "0.00") 'Format(rs!nInteres, "0.00")
                
                'Comision Calendario Pagar
                L.SubItems(26) = Format(rs!nComision, "0.00")
                
                'Total Cuota Calendario Pagar
                'L.SubItems(28) = Format(rs!ntotalcuota, "0.00")
                L.SubItems(28) = Val(lstCabecera.ListItems(I).SubItems(21)) + Val(lstCabecera.ListItems(I).SubItems(23)) + Val(lstCabecera.ListItems(I).SubItems(27)) + Val(lstCabecera.ListItems(I).SubItems(52))
        
                ' ==================== VAC ANTERIOR ====================
    
                'SK
                L.SubItems(13) = Format(lnSaldoCapitalTotal, "0.00")  'Format(rs!nSaldoCap, "0.00")
    
                'SK LP
                L.SubItems(16) = Format(rs!nSaldoCapLP, "0.00")
    
                'Interes Provisionado
                L.SubItems(18) = Format(rs!nInteresProvisionadoReal, "0.00")
                
                
                'Interes a Pagar (Diferencia)
                L.SubItems(24) = Format(lnInteresTotalCuota - rs!nInteresProvisionadoReal, "0.00") 'Format(rs!nInteres - rs!nInteresProvisionadoReal, "0.00")
                
                L.SubItems(48) = Format(rs!nCapitalConce, "0.00")
                L.SubItems(49) = Format(rs!nInteresConce, "0.00")
                L.SubItems(50) = Format(rs!nSaldoCapConce, "0.00")
                L.SubItems(51) = Format(rs!nProvisionConce, "0.00")
    
            End If
            
''            'Si se hubiera cargado Pago en Lote
''            If optBuscar(2).value = True Then
''                For i = 1 To nContador
''                    If InStr(1, rs!cCtaIFDesc, nAdeudo(i)) > 0 And rs!nMontoPrestado = nDesembolso(i) Then
''                        L.SubItems(34) = nAdeudo(i)
''                        L.SubItems(35) = Format(nCapital(i), "0.00")
''                        L.SubItems(36) = Format(nInteres(i), "0.00")
''                        L.SubItems(37) = Format(nComision(i), "0.00")
''                        L.SubItems(38) = Format(nCapital(i) + nInteres(i) + nComision(i), "0.00")
''
''                        L.SubItems(40) = Format(nCorrelativo(i), "0000")
''
''                        Exit For
''                    End If
''                Next
''
''                If Val(L.SubItems(35)) <> Val(L.SubItems(20)) Then 'Capital Diferente
''                    For j = 1 To 38
''                        L.ListSubItems.Item(j).ForeColor = vbRed
''                        L.ListSubItems.Item(j).ForeColor = vbRed
''                    Next
''                End If
''            End If
            L.SubItems(44) = rs!cPersCod
            L.SubItems(45) = rs!cIFTpo
            L.SubItems(46) = rs!cCtaIfCod
            L.SubItems(47) = rs!nSaldoMes
            rs.MoveNext
        Loop
        RSClose rs
    
        Me.cmdAceptar.Enabled = True
    
    End If
    
    Set rs = Nothing
    Set oAdeud = Nothing
    
    Status.Panels(1).Text = "Proceso " & Format(PB.Max * 100 / PB.Max, gsFormatoNumeroView) & "%"
    
    PB.Visible = False
    PB.value = 0
    
End Sub

Private Sub cmdE_Aceptar_Click()
    ActivaEdit False
    
    lstCabecera.SelectedItem.SubItems(22) = Format(Me.txtEdiicionInter.Text, gsFormatoNumeroDato)
    lstCabecera.SelectedItem.SubItems(24) = Format(Me.txtEdiicionInter.Text, gsFormatoNumeroDato)
    lstCabecera.SelectedItem.SubItems(26) = Format(Me.txtEdiicionCom.Text, gsFormatoNumeroDato)
    lstCabecera.SelectedItem.SubItems(27) = Format(Me.txtEdiicionCom.Text, gsFormatoNumeroDato)

    Me.txtInteres.Text = Format(Me.txtEdiicionInter.Text, gsFormatoNumeroView)
    Me.txtComision.Text = Format(Me.txtEdiicionCom.Text, gsFormatoNumeroView)
    
    'PEAC 20210722
    If Trim(lstCabecera.SelectedItem.SubItems(22)) = "" Then
        lstCabecera.SelectedItem.SubItems(22) = "0.00"
    ElseIf Trim(lstCabecera.SelectedItem.SubItems(26)) = "" Then
        lstCabecera.SelectedItem.SubItems(26) = "0.00"
    ElseIf Trim(lstCabecera.SelectedItem.SubItems(20)) = "" Then
        lstCabecera.SelectedItem.SubItems(20) = "0.00"
    ElseIf Trim(lstCabecera.SelectedItem.SubItems(42)) = "" Then
        lstCabecera.SelectedItem.SubItems(42) = "0.00"
    End If
    
    Me.lblTotal.Text = Format(CCur(lstCabecera.SelectedItem.SubItems(22)) + CCur(lstCabecera.SelectedItem.SubItems(26)) + CCur(lstCabecera.SelectedItem.SubItems(20)) + CCur(lstCabecera.SelectedItem.SubItems(42)), gsFormatoNumeroView)
    lstCabecera.SelectedItem.SubItems(28) = Format(CCur(lstCabecera.SelectedItem.SubItems(22)) + CCur(lstCabecera.SelectedItem.SubItems(26)) + CCur(lstCabecera.SelectedItem.SubItems(20)) + CCur(lstCabecera.SelectedItem.SubItems(42)), gsFormatoNumeroDato)
    txtImporte.Text = Me.lblTotal.Text
    txtBancoImporte.Text = Me.lblTotal.Text
End Sub

Private Sub cmdE_Cancelar_Click()
    ActivaEdit False
End Sub

Private Sub cmdE_Editar_Click()
    ActivaEdit True
End Sub

Private Sub ActivaEdit(pbValor As Boolean)
    Me.txtEdiicionCom.Enabled = pbValor
    Me.txtEdiicionInter.Enabled = pbValor
    Me.cmdE_Aceptar.Visible = pbValor
    Me.cmdE_Cancelar.Visible = pbValor
    Me.cmdE_Editar.Visible = Not pbValor
    lstCabecera.Enabled = Not pbValor
End Sub



