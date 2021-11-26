VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmLogContRegistro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Registro de Contratos"
   ClientHeight    =   7965
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   8415
   Icon            =   "frmLogContRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   120
      TabIndex        =   67
      Top             =   7320
      Width           =   8160
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Cancelar"
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
         Left            =   1320
         TabIndex        =   65
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdsalir 
         Cancel          =   -1  'True
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
         Height          =   345
         Left            =   6840
         TabIndex        =   66
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
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
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTContratos 
      Height          =   7140
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12594
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Datos del Contrato"
      TabPicture(0)   =   "frmLogContRegistro.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTpoMoneda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFecFin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFecIni"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "spnPlazo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboMoneda"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtMonto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraContrato"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraProv"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fraArea"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Cronograma"
      TabPicture(1)   =   "frmLogContRegistro.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCronograma"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraGarantia"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Items del Contrato"
      TabPicture(2)   =   "frmLogContRegistro.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraItemContrato"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraItemContrato 
         Caption         =   "Items Relacionados al contrato"
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
         Height          =   6615
         Left            =   -74880
         TabIndex        =   60
         Top             =   360
         Width           =   7935
         Begin VB.CommandButton cmdQuitarItemCont 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   960
            TabIndex        =   63
            Top             =   6120
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarItemCont 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   6120
            Width           =   855
         End
         Begin Sicmact.FlexEdit feOrden 
            Height          =   3255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   7695
            _extentx        =   13573
            _extenty        =   5741
            cols0           =   8
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Ag.Des.-Objeto-Descripcion-Solic.-P.Unitario-SubTotal-CtaContCod"
            encabezadosanchos=   "0-800-900-3000-700-1100-1100-0"
            font            =   "frmLogContRegistro.frx":035E
            font            =   "frmLogContRegistro.frx":0386
            font            =   "frmLogContRegistro.frx":03AE
            font            =   "frmLogContRegistro.frx":03D6
            font            =   "frmLogContRegistro.frx":03FE
            fontfixed       =   "frmLogContRegistro.frx":0426
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            columnasaeditar =   "X-1-2-X-X-X-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-1-1-0-0-0-0-0"
            encabezadosalineacion=   "C-C-L-L-R-R-R-L"
            formatosedit    =   "0-0-0-0-3-2-2-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            appearance      =   0
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit feObj 
            Height          =   615
            Left            =   9240
            TabIndex        =   70
            Top             =   360
            Width           =   4455
            _extentx        =   7858
            _extenty        =   1085
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Id-Objeto Orden-CtaContCod-CtaContDesc-Filtro-CodObjeto"
            encabezadosanchos=   "0-400-800-800-800-800-800"
            font            =   "frmLogContRegistro.frx":044C
            font            =   "frmLogContRegistro.frx":0474
            font            =   "frmLogContRegistro.frx":049C
            font            =   "frmLogContRegistro.frx":04C4
            font            =   "frmLogContRegistro.frx":04EC
            fontfixed       =   "frmLogContRegistro.frx":0514
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            columnasaeditar =   "X-X-X-X-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-L-C-C-C"
            formatosedit    =   "0-0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            appearance      =   0
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame fraCronograma 
         Caption         =   "Cronograma"
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
         Height          =   4365
         Left            =   -74880
         TabIndex        =   48
         Top             =   2040
         Width           =   7800
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "&Quitar"
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
            Left            =   120
            TabIndex        =   59
            Top             =   3840
            Width           =   1005
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
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
            Left            =   6600
            TabIndex        =   57
            Top             =   840
            Width           =   1005
         End
         Begin VB.ComboBox cboMonedaCro 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   840
            Width           =   1380
         End
         Begin VB.TextBox txtMontoCro 
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
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   5040
            MaxLength       =   15
            TabIndex        =   56
            Top             =   840
            Width           =   1380
         End
         Begin Sicmact.FlexEdit feCronograma 
            Height          =   2475
            Left            =   120
            TabIndex        =   58
            Top             =   1320
            Width           =   5640
            _extentx        =   9948
            _extenty        =   4366
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Nº Pago-Fecha de Pago-Moneda-Monto"
            encabezadosanchos=   "500-1000-1200-1000-1200"
            font            =   "frmLogContRegistro.frx":053A
            font            =   "frmLogContRegistro.frx":0562
            font            =   "frmLogContRegistro.frx":058A
            font            =   "frmLogContRegistro.frx":05B2
            font            =   "frmLogContRegistro.frx":05DA
            fontfixed       =   "frmLogContRegistro.frx":0602
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   7
            columnasaeditar =   "X-X-X-X-X"
            textstylefixed  =   3
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-C-C-C-C"
            formatosedit    =   "0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            appearance      =   0
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin MSComCtl2.DTPicker txtFechaPago 
            Height          =   315
            Left            =   1320
            TabIndex        =   52
            Top             =   840
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   128909313
            CurrentDate     =   37156
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pago"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   600
            TabIndex        =   49
            Top             =   360
            Width           =   600
         End
         Begin VB.Label lblNPago 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   50
            Tag             =   "txtcodigo"
            Top             =   360
            Width           =   525
         End
         Begin VB.Label lblTpoMonedaCro 
            AutoSize        =   -1  'True
            Caption         =   "S/"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   55
            Top             =   840
            Width           =   180
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cuota:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2880
            TabIndex        =   53
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Pago:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1140
         End
      End
      Begin VB.Frame fraGarantia 
         Caption         =   "Garantía"
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
         Height          =   1605
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   7800
         Begin VB.ComboBox cboTipoGarantia 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   240
            Width           =   5700
         End
         Begin VB.ComboBox cboMonedaGar 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   720
            Width           =   1380
         End
         Begin VB.TextBox txtMontoGar 
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
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   47
            Top             =   1200
            Width           =   1305
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   600
            TabIndex        =   41
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Monto "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   43
            Top             =   720
            Width           =   585
         End
         Begin VB.Label lblTpoMonedaGar 
            AutoSize        =   -1  'True
            Caption         =   "S/"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   46
            Top             =   1200
            Width           =   180
         End
      End
      Begin VB.Frame fraArea 
         Caption         =   "Area Usuaria"
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
         Height          =   1260
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   7920
         Begin VB.TextBox txtAgeDesc2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Top             =   720
            Width           =   5700
         End
         Begin VB.TextBox txtAgeDesc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   240
            Width           =   5700
         End
         Begin Sicmact.TxtBuscar txtArea 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1770
            _extentx        =   3122
            _extenty        =   556
            appearance      =   0
            appearance      =   0
            font            =   "frmLogContRegistro.frx":0628
            appearance      =   0
         End
         Begin Sicmact.TxtBuscar txtArea2 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1770
            _extentx        =   3122
            _extenty        =   556
            appearance      =   0
            appearance      =   0
            font            =   "frmLogContRegistro.frx":0654
            appearance      =   0
         End
      End
      Begin VB.Frame fraProv 
         Caption         =   "Proveedor "
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7920
         Begin VB.TextBox txtProvNom 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            TabIndex        =   3
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   5700
         End
         Begin Sicmact.TxtBuscar txtPersona 
            Height          =   315
            Left            =   135
            TabIndex        =   2
            Top             =   240
            Width           =   1740
            _extentx        =   3069
            _extenty        =   556
            appearance      =   0
            appearance      =   0
            font            =   "frmLogContRegistro.frx":0680
            appearance      =   0
            tipobusqueda    =   3
            stitulo         =   ""
         End
      End
      Begin VB.Frame fraContrato 
         Caption         =   "Contrato"
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
         Height          =   3795
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   7920
         Begin VB.ComboBox cboTipoContratacion 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   3360
            Width           =   1980
         End
         Begin VB.ComboBox cboobjContrato 
            Height          =   315
            Left            =   1440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   2880
            Width           =   2220
         End
         Begin VB.CommandButton cmdBuscarArchivo 
            Caption         =   "E&xaminar"
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
            Left            =   6480
            TabIndex        =   33
            ToolTipText     =   "Buscar Credito"
            Top             =   2380
            Width           =   1215
         End
         Begin VB.TextBox txtGlosa 
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
            Height          =   810
            Left            =   1440
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   1320
            Width           =   5700
         End
         Begin VB.TextBox txtNContrato 
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
            Height          =   330
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   26
            Top             =   840
            Width           =   1860
         End
         Begin VB.ComboBox cboTipoContrato 
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   360
            Width           =   2100
         End
         Begin VB.ComboBox cboTipoPago 
            Height          =   315
            Left            =   1440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   1860
         End
         Begin MSComCtl2.DTPicker txtFechaFirma 
            Height          =   315
            Left            =   5040
            TabIndex        =   28
            Top             =   840
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   128909313
            CurrentDate     =   37156
         End
         Begin Sicmact.TxtBuscar txtObjeto 
            Height          =   315
            Left            =   4680
            TabIndex        =   37
            Top             =   2880
            Width           =   2580
            _extentx        =   4551
            _extenty        =   556
            appearance      =   0
            appearance      =   0
            font            =   "frmLogContRegistro.frx":06AC
            appearance      =   0
            stitulo         =   ""
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Contratación:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   3360
            Width           =   1530
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Objeto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   36
            Top             =   2900
            Width           =   510
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Objeto Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   2895
            Width           =   1155
         End
         Begin VB.Label lblNombreArchivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1440
            TabIndex        =   32
            Tag             =   "txtnombre"
            Top             =   2400
            Width           =   4935
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Contrato Digital"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   2400
            Width           =   1080
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Firma"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   27
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Contrato"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3720
            TabIndex        =   23
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Pago"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.TextBox txtMonto 
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2760
         Width           =   1380
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2760
         Width           =   1380
      End
      Begin Spinner.uSpinner spnPlazo 
         Height          =   330
         Left            =   6480
         TabIndex        =   19
         Top             =   2760
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         Max             =   360
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
      Begin MSComCtl2.DTPicker txtFecIni 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   128909313
         CurrentDate     =   37156
      End
      Begin MSComCtl2.DTPicker txtFecFin 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   2760
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   128909313
         CurrentDate     =   37156
      End
      Begin VB.Label lblTpoMoneda 
         AutoSize        =   -1  'True
         Caption         =   "S/"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   2820
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   13
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto Contrato"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   15
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1410
         TabIndex        =   11
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   3240
         TabIndex        =   68
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Contractual"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuotas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   18
         Top             =   2520
         Width           =   495
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   375
      Left            =   7680
      TabIndex        =   69
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Filtro          =   ""
      Altura          =   0
   End
End
Attribute VB_Name = "frmLogContRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsPathFile As String
Dim fsRuta As String
Dim fsNomFile As String
Dim lsCodPers As String
Dim fMatCronograma() As Variant
Dim I, J As Integer
Dim pbActivaArchivo As Boolean
Dim psRutaContrato As String
Dim fsSubCta As String
Dim rs As ADODB.Recordset 'PASI20140722 TI-ERS077-2014
Dim sObjCod    As String, sObjDesc As String, sObjUnid As String 'PASI20140722 TI-ERS077-2014
Dim sCtaCod    As String, sCtaDesc As String 'PASI20140722 TI-ERS077-2014
Dim gsOpeCod As String 'PASI20140930 ERS0772014
Dim fRsAgencia As New ADODB.Recordset
Dim fRsServicio As New ADODB.Recordset
Dim fRsCompra As New ADODB.Recordset
Dim fnTpoCambio As Currency
Dim fntpodocorigen As Integer 'PASI20140110 ERS0772014
Dim fsCtaContCodProv As String
Dim fbEstCuota As Boolean
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Function ValidaDatos() As Boolean
Dim olog As DLogGeneral
Set olog = New DLogGeneral
Dim lnMontoTotCronograma As Currency

If Trim(Me.txtPersona.Text) = "" Then
    MsgBox "Ingrese Proveedor", vbInformation, "Aviso"
    Call VerPestana(0)
    txtPersona.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(Me.txtArea.Text) = "" Then
    MsgBox "Ingrese Area Usuaria", vbInformation, "Aviso"
    Call VerPestana(0)
    txtArea.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(cboMoneda.Text) = "" Then
    MsgBox "Seleccione la Moneda del Contrato", vbInformation, "Aviso"
    Call VerPestana(0)
    cboMoneda.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(Me.txtMonto.Text) = "" Or Trim(Me.txtMonto.Text) = "0.00" Then
    MsgBox "Ingrese Monto de Contrato", vbInformation, "Aviso"
    Call VerPestana(0)
    txtMonto.SetFocus
    ValidaDatos = False
    Exit Function
End If
    
If fraCronograma.Enabled = True Then 'PASI201408522 TI-ERS077-2014
    If Trim(Me.spnPlazo.Valor) = "0" Then
        MsgBox "Ingrese el Numero de Cuotas", vbInformation, "Aviso"
        Call VerPestana(0)
        spnPlazo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
End If

If Trim(cboTipoPago.Text) = "" Then
    MsgBox "Seleccione el Tipo de Pago", vbInformation, "Aviso"
    Call VerPestana(0)
    cboTipoPago.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(cboTipoContrato.Text) = "" Then
    MsgBox "Seleccione el Tipo de Contrato", vbInformation, "Aviso"
    Call VerPestana(0)
    cboTipoContrato.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(Me.txtNContrato.Text) = "" Then
    MsgBox "Ingrese Nº de Contrato", vbInformation, "Aviso"
    Call VerPestana(0)
    txtNContrato.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(Me.txtGlosa.Text) = "" Then
    MsgBox "Ingrese Glosa", vbInformation, "Aviso"
    Call VerPestana(0)
    txtGlosa.SetFocus
    ValidaDatos = False
    Exit Function
End If

'PASI20140723 TI-ERs077-2014
If Trim(cboTipoContrato.Text) <> "" Then
    Select Case CInt(Right(cboTipoContrato.Text, 4))
        Case 4, 5
            If Trim(cboobjContrato.Text) = "" Then
                MsgBox "Seleccione el Objeto de Contrato", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
            If Len(Trim(txtObjeto.Text)) = 0 Then
                MsgBox "Ud. debe de seleccionar el Objeto", vbInformation, "Aviso"
                'Call VerPestana(0) PASIERS0772014
                txtObjeto.SetFocus
                ValidaDatos = False
            Exit Function
            End If
    End Select
End If
'end PASI

If fraGarantia.Enabled = True Then 'PASI20140723 TI-ERs077-2014
    If Trim(Me.cboTipoGarantia.Text) = "" Then
        MsgBox "Ingrese el Tipo de Garantía", vbInformation, "Aviso"
        Call VerPestana(1)
        cboTipoGarantia.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Trim(Me.cboMonedaGar.Text) = "" Then
        MsgBox "Ingrese la Moneda de la Garantía", vbInformation, "Aviso"
        Call VerPestana(1)
        cboMonedaGar.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If Trim(Right(Trim(Me.cboTipoGarantia.Text), 4)) <> "4" Then 'WIOR 20130108
        If Trim(Me.txtMontoGar.Text) = "" Or Trim(Me.txtMontoGar.Text) = "0.00" Then
            MsgBox "Ingrese Monto de la Garantía", vbInformation, "Aviso"
            Call VerPestana(1)
            txtMontoGar.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    'WIOR 20130108 ****************************
    Else
        If Trim(Me.txtMontoGar.Text) = "" Then
            MsgBox "Ingrese Monto de la Garantía", vbInformation, "Aviso"
            Call VerPestana(1)
            txtMontoGar.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
End If
'WIOR FIN ********************************

If fraCronograma.Enabled = True Then
    If CInt(Me.spnPlazo.Valor) > (CInt(Me.lblNPago.Caption) - 1) Then
        MsgBox "Aun no ha registrado todas la Cuotas en el Cronograma", vbInformation, "Aviso"
        Call VerPestana(1)
        Me.cmdAgregar.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    Dim oTipoCam  As NTipoCambio
    Dim nTC As Double
     Dim nUIT As Double
    Set oTipoCam = New NTipoCambio
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, 1)
    nUIT = olog.ObtenerUITActual
    
    If fraGarantia.Enabled = True Then 'PASIERS0772014
        If CDbl(30 * nUIT) < (CDbl(Me.txtMonto.Text) * CDbl((IIf(Trim(Right(Me.cboMoneda.Text, 2)) = "1", 1, nTC)))) Then
            If Trim(Right(Me.cboTipoGarantia.Text, 2)) <> 1 Then
                MsgBox "Debe elegir como tipo de garantía FIEL CUMPLIMIENTO ", vbInformation, "Aviso"
                Call VerPestana(1)
                cboTipoGarantia.ListIndex = IndiceListaCombo(cboTipoGarantia, 1)
                Me.cboTipoGarantia.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
    End If
        lnMontoTotCronograma = Me.feCronograma.SumaRow(4)
        lnMontoTotCronograma = Round(lnMontoTotCronograma, 2) 'PASI20150107
        If (Trim(Right(cboTipoContrato.Text, 2)) = LogTipoContrato.ContratoServicio And Trim(Right(cboTipoPago.Text, 2)) = 1) Or Trim(Right(cboTipoContrato.Text, 2)) <> LogTipoContrato.ContratoServicio Then
             If Round(CDbl(txtMonto.Text), 2) <> lnMontoTotCronograma Then
                MsgBox "El monto total del cronograma no coincide con el total ingresado, verifique", vbInformation, "Aviso"
                'Call VerPestana(0)
                txtMonto.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
End If

If fraItemContrato.Enabled = True Then 'PASI20140723 TI-ERS077-2014
   If feOrden.TextMatrix(1, 1) = "" Then
        MsgBox "No se ha ingresado ningun Item de contrato", vbInformation, "Aviso"
        feOrden.SetFocus
        ValidaDatos = False
        Exit Function
   End If
   Dim I As Integer
   Dim nMonto As Double
   Dim nCant As Integer
   Dim nPreUnit As Double
   'PASI20150107
   Dim nMontoCont As Double
   Dim nMontoCuotas As Double
   'end PASI
   nMonto = 0
    For I = 1 To feOrden.Rows - 1
        nMonto = nMonto + (feOrden.TextMatrix(I, 6))
    Next I
    nMonto = Round(nMonto, 2) 'PASI20150107
    If CInt(Trim(Right(cboTipoContrato.Text, 4))) = LogTipoContrato.ContratoServicio Then
        If Trim(Right(cboTipoPago.Text, 2)) = 1 Then
            'Modificado PASI20150107
            'If (CDbl(txtMonto.Text) / CDbl(spnPlazo.valor)) <> nMonto Then 'Comentado PASI20150107
            nMontoCont = Round(CDbl(txtMonto.Text), 2)
            nMontoCuotas = Round(nMontoCont / CDbl(spnPlazo.Valor), 2)
            If nMontoCuotas <> nMonto Then 'PASI20150107
                MsgBox "El monto total de los Bienes/Servicios del Contrato no coincide con el Monto de Pago en las Cuotas del Cronograma, verifique", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
            End If
'        ElseIf Round(CDbl(txtMonto.Text), 2) <> nMonto Then 'Comentado PASI20150108
'            MsgBox "El monto total de los Bienes/Servicios del Contrato no coincide con el total ingresado, verifique", vbInformation, "Aviso"
'            'Call VerPestana(0)
'            'txtMonto.SetFocus
'            ValidaDatos = False
'            Exit Function
        End If
    Else
        If Round(CDbl(txtMonto.Text), 2) <> nMonto Then
            MsgBox "El monto total de los Bienes/Servicios del Contrato no coincide con el total ingresado, verifique", vbInformation, "Aviso"
            'Call VerPestana(0)
            'txtMonto.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
End If
'end PASI
'EJVG20131009 ***
If olog.ExisteNContrato(Trim(Me.txtNContrato.Text)) = 0 Then
    MsgBox "Nº de Contrato ya Existe", vbInformation, "Aviso"
    Call VerPestana(0)
    Me.txtNContrato.SetFocus
    ValidaDatos = False
    Exit Function
End If

'If cboTipoContratacion.ListIndex = -1 Then
'    MsgBox "Ud. debe de seleccionar el Tipo de Contratación", vbInformation, "Aviso"
'    Call VerPestana(0)
'    cboTipoContratacion.SetFocus
'    ValidaDatos = False
'    Exit Function
'End If
'If Len(Trim(txtObjeto.Text)) = 0 Then
'    MsgBox "Ud. debe de seleccionar el Objeto", vbInformation, "Aviso"
'    Call VerPestana(0)
'    txtObjeto.SetFocus
'    ValidaDatos = False
'    Exit Function
'End If
Set olog = Nothing
'END EJVG *******
ValidaDatos = True
End Function

Sub VerPestana(ByVal I As Integer)
If I = 0 Then
    Me.SSTContratos.TabVisible(I) = False
    Me.SSTContratos.TabVisible(I + 1) = False
    Me.SSTContratos.TabVisible(I) = True
    Me.SSTContratos.TabVisible(I + 1) = True
ElseIf I = 1 Then
    Me.SSTContratos.TabVisible(I - 1) = False
    Me.SSTContratos.TabVisible(I) = False
    Me.SSTContratos.TabVisible(I) = True
    Me.SSTContratos.TabVisible(I - 1) = True
End If
End Sub

Private Function ValidaCronograma() As Boolean
If Me.cboMonedaCro.Text = "" Then
    MsgBox "Seleccionar moneda para la cuota.", vbInformation, "Aviso"
    ValidaCronograma = False
    Exit Function
End If
If Trim(Right(cboTipoContrato.Text, 2)) = LogTipoContrato.ContratoServicio Then
    If Trim(Right(cboTipoPago.Text, 2)) <> 2 Then
        If Trim(Me.txtMontoCro.Text) = "" Or Trim(Me.txtMontoCro.Text) = "0.00" Then
            MsgBox "Ingrese el monto de la cuota.", vbInformation, "Aviso"
            ValidaCronograma = False
            Exit Function
        End If
    End If
Else
    If Trim(Me.txtMontoCro.Text) = "" Or Trim(Me.txtMontoCro.Text) = "0.00" Then
        MsgBox "Ingrese el monto de la cuota.", vbInformation, "Aviso"
        ValidaCronograma = False
        Exit Function
    End If
End If
If CInt(Me.lblNPago.Caption) > CInt(Me.spnPlazo.Valor) Then
    MsgBox "Solo tiene plazo de " & Trim(Me.spnPlazo.Valor) & " cuota" & IIf(Trim(Me.spnPlazo.Valor) = "1", "", "s") & ".", vbInformation, "Aviso"
    ValidaCronograma = False
    Exit Function
End If

ValidaCronograma = True
End Function
Private Sub cboMoneda_Click()
    cboTipoContratacion.ListIndex = -1
    txtObjeto.Text = ""
    If cboMoneda.Text <> "" Then
        If CInt(Right(cboMoneda.Text, 2)) = gMonedaNacional Then
            txtMonto.BackColor = vbWhite
            '''Me.lblTpoMoneda.Caption = "S/." 'MARG ERS044-2016
            Me.lblTpoMoneda.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
            gcOpeCod = "501215" 'PASI20140722 ERS0772014
        Else
            txtMonto.BackColor = RGB(200, 255, 200)
            Me.lblTpoMoneda.Caption = "$"
            gcOpeCod = "502215"  'PASI20140722 ERS0772014
        End If
        'cboMonedaGar.ListIndex = IndiceListaCombo(cboMonedaGar, CInt(Right(cboMoneda.Text, 2))) 'Comentado PASI20150107s
        'cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, CInt(Right(cboMoneda.Text, 2)))
        cboTipoContrato_Click
        Set fRsServicio = OrdenServicio() 'PASI20140110 ERS0772014
    End If
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      
    End If
End Sub
Private Sub cboMonedaCro_Click()
    If cboMonedaCro.Text <> "" Then
        If Trim(Right(cboTipoContrato.Text, 2)) = LogTipoContrato.ContratoServicio Then 'Contrato de Servicio con pago Variable
            If Trim(Right(cboTipoPago.Text, 2)) = 2 Then
                If CInt(Right(cboMonedaCro.Text, 2)) = gMonedaNacional Then
                    txtMontoCro.BackColor = vbWhite
                    '''Me.lblTpoMonedaCro.Caption = "S/." 'MARG ERS044-2016
                    Me.lblTpoMonedaCro.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
                Else
                    txtMontoCro.BackColor = RGB(200, 255, 200)
                    Me.lblTpoMonedaCro.Caption = "$"
                End If
                If Trim(cboMoneda.Text) <> "" Then
                    cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, CInt(Right(cboMoneda.Text, 2)))
                End If
                txtMontoCro.Enabled = False
                cmdAgregar.SetFocus
            Else
                    If CInt(Right(cboMonedaCro.Text, 2)) = gMonedaNacional Then
                        txtMontoCro.BackColor = vbWhite
                        '''Me.lblTpoMonedaCro.Caption = "S/." 'MARG ERS044-2016
                        Me.lblTpoMonedaCro.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
                    Else
                        txtMontoCro.BackColor = RGB(200, 255, 200)
                        Me.lblTpoMonedaCro.Caption = "$"
                    End If
                      If Trim(cboMoneda.Text) <> "" Then
                        cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, CInt(Right(cboMoneda.Text, 2)))
                    End If
                    'PASI20141227
                    txtMontoCro.Enabled = True
                    txtMontoCro.SetFocus
                    'end PASI
            End If
        ElseIf Trim(Right(cboTipoContrato.Text, 2)) <> "" Then  'NORMAL
            If CInt(Right(cboMonedaCro.Text, 2)) = gMonedaNacional Then
                txtMontoCro.BackColor = vbWhite
                '''Me.lblTpoMonedaCro.Caption = "S/." 'MARG ERS044-2016
                Me.lblTpoMonedaCro.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
            Else
                txtMontoCro.BackColor = RGB(200, 255, 200)
                Me.lblTpoMonedaCro.Caption = "$"
            End If
              If Trim(cboMoneda.Text) <> "" Then
                cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, CInt(Right(cboMoneda.Text, 2)))
            End If
             'PASI20141227
                    txtMontoCro.Enabled = True
                    txtMontoCro.SetFocus
             'end PASI
        End If
    End If
End Sub

Private Sub cboMonedaGar_Click()
If cboMonedaGar.Text <> "" Then
        If CInt(Right(cboMonedaGar.Text, 2)) = gMonedaNacional Then
            txtMontoGar.BackColor = vbWhite
            '''Me.lblTpoMonedaGar.Caption = "S/." 'MARG ERS044-2016
            Me.lblTpoMonedaGar.Caption = gcPEN_SIMBOLO 'MARG ERS044-2016
        Else
            txtMontoGar.BackColor = RGB(200, 255, 200)
            Me.lblTpoMonedaGar.Caption = "$"
        End If
        
        If Trim(cboMoneda.Text) <> "" Then
            cboMonedaGar.ListIndex = IndiceListaCombo(cboMonedaGar, CInt(Right(cboMoneda.Text, 2)))
        End If
    End If
End Sub
'PASI20140721 TI-ERS077-2014
Private Sub cboobjContrato_Click()
    If cboobjContrato.Text <> "" Then
        Select Case CInt(Right(cboTipoContrato.Text, 4))
            Case LogTipoContrato.ContratoServicio
                Select Case CInt(Right(cboobjContrato.Text, 4))
                    Case 1
                        fraGarantia.Enabled = False
                        fraCronograma.Enabled = True
                        'fraItemContrato.Enabled = False
                    Case 2
                        fraGarantia.Enabled = True
                        fraCronograma.Enabled = True
                        'fraItemContrato.Enabled = False
                End Select
            Case LogTipoContrato.ContratoArrendamiento
                Select Case CInt(Right(cboobjContrato.Text, 4))
                    Case 1
                        fraGarantia.Enabled = False
                        fraCronograma.Enabled = True
                        fraItemContrato.Enabled = False
                    Case 2
                        fraGarantia.Enabled = True
                        fraCronograma.Enabled = True
                        fraItemContrato.Enabled = False
                End Select
            Case LogTipoContrato.ContratoObra
                Select Case CInt(Right(cboobjContrato.Text, 4))
                    Case 1
                        fraGarantia.Enabled = False
                        fraCronograma.Enabled = False
                        fraItemContrato.Enabled = False
                    Case 2
                        fraGarantia.Enabled = True
                        fraCronograma.Enabled = False
                        fraItemContrato.Enabled = False
                End Select
            Case LogTipoContrato.ContratoSuministro
                Select Case CInt(Right(cboobjContrato.Text, 4))
                      Case 1
                        fraGarantia.Enabled = False
                        fraCronograma.Enabled = False
                        
                    Case 2
                        fraGarantia.Enabled = True
                        fraCronograma.Enabled = False
                        
                End Select
        End Select
    End If
End Sub
'end PASI
'EJVG20131009 ***
Private Sub cboTipoContratacion_Click()
    Dim oALmacen As New DLogAlmacen
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrCboTipoContratacion
    Screen.MousePointer = 11
    If cboTipoContratacion.Text <> "" Then
        If CInt(Right(cboTipoContratacion.Text, 4)) = 1 Then 'Bienes
            Set rs = oALmacen.GetBienesAlmacen(, "11','12','13")
        Else 'Servicios
            Set rs = OrdenServicio
        End If
    End If
    txtObjeto.rs = rs
    txtObjeto.Text = ""
    Set oALmacen = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCboTipoContratacion:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'PASI20140719 TI-ERS077-2014
Private Sub LimpiarCronograma()
    'GARANTIA
    cboTipoGarantia.ListIndex = IndiceListaCombo(cboTipoGarantia, 0)
    cboMonedaGar.ListIndex = IndiceListaCombo(cboMonedaGar, 0)
    Me.txtMontoGar.Text = ""

    'CRONOGRAMA
    Me.txtFechaPago.value = gdFecSis
    cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, 0)
    Me.txtMontoCro.Text = ""
    Call LimpiaFlex(Me.feCronograma)
    ReDim Preserve fMatCronograma(5, 1 To 1)
    lblNPago.Caption = 1 'PASI20150107
End Sub
Private Sub LimpiarItem()
    Call LimpiaFlex(Me.feOrden)
End Sub
Private Sub cboTipoContrato_Click()
    Dim oALmacen As New DLogAlmacen
    Dim rs As New ADODB.Recordset
    Dim lnTpoDoc As Integer
    
    On Error GoTo ErroCboTipoContrato
    Screen.MousePointer = 11
    spnPlazo.Enabled = True
    If cboTipoContrato.Text <> "" Then
        Select Case CInt(Right(cboTipoContrato.Text, 4))
            Case LogTipoContrato.ContratoServicio
                LimpiarCronograma
                LimpiarItem
                EstadoObjeto (False)
                Label18.Visible = True
                cboobjContrato.Visible = True
                ReiniciarEstadoFrame
                fraItemContrato.Enabled = True
                 'cancela_busqueda_actual
                If Trim(Right(cboTipoContrato.Text, 4)) <> "" Then 'PASI20140110 ERS0772014
                    lnTpoDoc = CInt(Trim(Right(cboTipoContrato.Text, 4)))
                End If
                fntpodocorigen = lnTpoDoc
                If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
                    fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                    feOrden.lbUltimaInstancia = True
                ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                    feOrden.lbUltimaInstancia = False
                End If
                Screen.MousePointer = 0
                txtNContrato.SetFocus
                Exit Sub
            Case LogTipoContrato.ContratoArrendamiento, LogTipoContrato.ContratoObra
                Set rs = OrdenServicio
                LimpiarCronograma
                LimpiarItem
                EstadoObjeto (True)
                ReiniciarEstadoFrame
                cboobjContrato_Click
                txtNContrato.SetFocus
                'If CInt(Me.spnPlazo.Valor) <> 0 Then
                '    EstadoTab 2
                'Else
                '    EstadoTab 1
                'End If
                If CInt(Right(cboTipoContrato.Text, 4)) = LogTipoContrato.ContratoObra Then
                    spnPlazo.Valor = 0
                    spnPlazo.Enabled = False
                End If
            Case LogTipoContrato.ContratoAdqBienes ', LogTipoContrato.ContratoSuministro
                LimpiarCronograma
                LimpiarItem
                EstadoObjeto (False)
                ReiniciarEstadoFrame
                fraItemContrato.Enabled = True
                Label18.Visible = True
                cboobjContrato.Visible = True
                'EstadoTab 3
                'cancela_busqueda_actual
                If Trim(Right(cboTipoContrato.Text, 4)) <> "" Then 'PASI20140110 ERS0772014
                    lnTpoDoc = CInt(Trim(Right(cboTipoContrato.Text, 4)))
                End If
                fntpodocorigen = lnTpoDoc
                If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
                    fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                    feOrden.lbUltimaInstancia = True
                ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                    feOrden.lbUltimaInstancia = False
                End If
                'PASIERS20141227
                    spnPlazo.Valor = 0
                    spnPlazo.Enabled = False
                'end PASI
                Screen.MousePointer = 0
                txtNContrato.SetFocus
                Exit Sub
            Case LogTipoContrato.ContratoSuministro
                LimpiarCronograma
                LimpiarItem
                EstadoObjeto (False)
                ReiniciarEstadoFrame
                fraItemContrato.Enabled = True
                Label18.Visible = True
                cboobjContrato.Visible = True
                'EstadoTab 3
                
                'cancela_busqueda_actual
                If Trim(Right(cboTipoContrato.Text, 4)) <> "" Then 'PASI20140110 ERS0772014
                    lnTpoDoc = CInt(Trim(Right(cboTipoContrato.Text, 4)))
                End If
                fntpodocorigen = lnTpoDoc
                If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
                    fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                    feOrden.lbUltimaInstancia = True
                ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                    feOrden.lbUltimaInstancia = False
                End If
                 'PASIERS20141227
                    spnPlazo.Valor = 0
                    spnPlazo.Enabled = False
                'end PASI
                Screen.MousePointer = 0
                txtNContrato.SetFocus
                Exit Sub
        End Select
    End If
    txtObjeto.rs = rs
    txtObjeto.Text = ""
   
    Set oALmacen = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErroCboTipoContrato:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub ReiniciarEstadoFrame()
    fraGarantia.Enabled = False
    fraCronograma.Enabled = False
    fraItemContrato.Enabled = False
End Sub
Private Sub EstadoObjeto(pbhabilita As Boolean)
    Label18.Visible = pbhabilita
    cboobjContrato.Visible = pbhabilita
    Label20.Visible = pbhabilita
    txtObjeto.Visible = pbhabilita
End Sub
Private Sub EstadoTab(pnTab As Integer)
'    Select Case pnTab
'        Case 1
'            Me.SSTContratos.TabEnabled(1) = False
'            Me.SSTContratos.TabVisible(2) = False
'        Case 2
'            Me.SSTContratos.TabEnabled(1) = True
'            Me.SSTContratos.TabVisible(2) = False
'        Case 3
'            Me.SSTContratos.TabEnabled(1) = False
'            Me.SSTContratos.TabVisible(2) = True
'    End Select
End Sub
Private Sub cboTipoGarantia_Click() 'PASI20140107
    If Trim(Right(cboTipoGarantia, 2)) <> "" Then
        If Trim(Right(cboTipoGarantia, 2)) = 4 Then
            cboMonedaGar.ListIndex = -1
            cboMonedaGar.Enabled = False
            txtMontoGar.Text = ""
            txtMontoGar.Enabled = False
        Else
            cboMonedaGar.Enabled = True
            txtMontoGar.Enabled = True
        End If
    End If
End Sub
Private Sub cboTipoPago_Click()
    cboTipoContrato.SetFocus
    LimpiarCronograma
    LimpiarItem
    EstadoObjeto (False)
    
    If Not spnPlazo.Valor <> 0 Then  'PASI20141227
        spnPlazo.Enabled = True
        spnPlazo.Valor = 0
    End If
    cboTipoContrato.ListIndex = -1
    
End Sub

'end PASI
'END EJVG *******
Private Sub cmdAgregar_Click()
If CInt(Me.spnPlazo.Valor) = 0 Then
    MsgBox "Nro de cuotas es 0, Favor de verificar", vbInformation, "Aviso"
Else
    If ValidaCronograma Then
        If MsgBox("Estas seguro de agregar registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            If Trim(Right(cboTipoContrato.Text, 2)) = LogTipoContrato.ContratoServicio Then
                If Trim(Right(cboTipoPago.Text, 2)) = 2 Then
                    ReDim Preserve fMatCronograma(5, 1 To CInt(Trim(lblNPago.Caption)))
                    fMatCronograma(1, CInt(lblNPago.Caption)) = Me.lblNPago.Caption
                    fMatCronograma(2, CInt(lblNPago.Caption)) = Format(txtFechaPago.value, "DD/MM/YYYY")
                    fMatCronograma(3, CInt(lblNPago.Caption)) = Trim(Right(Me.cboMonedaCro.Text, 4))
                    fMatCronograma(4, CInt(lblNPago.Caption)) = Trim(Left(Me.cboMonedaCro.Text, 20))
                    fMatCronograma(5, CInt(lblNPago.Caption)) = "-"
                    Call LeerMatriz(CInt(lblNPago.Caption))
                    Me.lblNPago.Caption = CInt(lblNPago.Caption) + 1
                    txtFechaPago.value = DateAdd("M", 1, CDate(txtFechaPago.value)) 'EJVG20131009
                Else
                    ReDim Preserve fMatCronograma(5, 1 To CInt(Trim(lblNPago.Caption)))
                    fMatCronograma(1, CInt(lblNPago.Caption)) = Me.lblNPago.Caption
                    fMatCronograma(2, CInt(lblNPago.Caption)) = Format(txtFechaPago.value, "DD/MM/YYYY")
                    fMatCronograma(3, CInt(lblNPago.Caption)) = Trim(Right(Me.cboMonedaCro.Text, 4))
                    fMatCronograma(4, CInt(lblNPago.Caption)) = Trim(Left(Me.cboMonedaCro.Text, 20))
                    fMatCronograma(5, CInt(lblNPago.Caption)) = Format(Me.txtMontoCro.Text, "##00.00")
                    Call LeerMatriz(CInt(lblNPago.Caption))
                    Me.lblNPago.Caption = CInt(lblNPago.Caption) + 1
                    txtFechaPago.value = DateAdd("M", 1, CDate(txtFechaPago.value)) 'EJVG20131009
                End If
            Else
                ReDim Preserve fMatCronograma(5, 1 To CInt(Trim(lblNPago.Caption)))
                fMatCronograma(1, CInt(lblNPago.Caption)) = Me.lblNPago.Caption
                fMatCronograma(2, CInt(lblNPago.Caption)) = Format(txtFechaPago.value, "DD/MM/YYYY")
                fMatCronograma(3, CInt(lblNPago.Caption)) = Trim(Right(Me.cboMonedaCro.Text, 4))
                fMatCronograma(4, CInt(lblNPago.Caption)) = Trim(Left(Me.cboMonedaCro.Text, 20))
                fMatCronograma(5, CInt(lblNPago.Caption)) = Format(Me.txtMontoCro.Text, "##00.00")
                Call LeerMatriz(CInt(lblNPago.Caption))
                Me.lblNPago.Caption = CInt(lblNPago.Caption) + 1
                txtFechaPago.value = DateAdd("M", 1, CDate(txtFechaPago.value)) 'EJVG20131009
            End If
        End If
        cmdAgregar.SetFocus
    End If
End If
'cboTipoContrato.SetFocus
End Sub

Private Sub LeerMatriz(ByVal tamano As Integer)
Dim I As Integer
Call LimpiaFlex(feCronograma)
For I = 0 To tamano - 1
    feCronograma.AdicionaFila
    feCronograma.TextMatrix(I + 1, 0) = I + 1
    feCronograma.TextMatrix(I + 1, 1) = fMatCronograma(1, I + 1)
    feCronograma.TextMatrix(I + 1, 2) = fMatCronograma(2, I + 1)
    feCronograma.TextMatrix(I + 1, 3) = fMatCronograma(4, I + 1)
    feCronograma.TextMatrix(I + 1, 4) = fMatCronograma(5, I + 1)
Next I
End Sub
Private Sub cmdBuscarArchivo_Click()
Dim I As Integer
CdlgFile.nHwd = Me.hwnd
CdlgFile.Filtro = "Contratos Digital (*.pdf)|*.pdf"
Me.CdlgFile.Altura = 300
CdlgFile.Show

fsPathFile = CdlgFile.Ruta
fsRuta = fsPathFile
        If fsPathFile <> Empty Then
            For I = Len(fsPathFile) - 1 To 1 Step -1
                    If Mid(fsPathFile, I, 1) = "\" Then
                        fsPathFile = Mid(CdlgFile.Ruta, 1, I)
                        fsNomFile = Mid(CdlgFile.Ruta, I + 1, Len(CdlgFile.Ruta) - I)
                        Exit For
                    End If
             Next I
          Screen.MousePointer = 11
          
            If pbActivaArchivo Then
                If Trim(txtNContrato.Text) = "" Then
                    MsgBox "Ingrese primero el Nº de Contrato"
                Else
                    lblNombreArchivo.Caption = UCase(Trim(Me.txtNContrato.Text)) & ".pdf"
                End If
            Else
                lblNombreArchivo.Caption = ""
            End If
        Else
           MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
           Exit Sub
        End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancela_Click()
    LimpiarDatos
End Sub

'Private Sub CmdGrabar_Click()
'On Error GoTo ErrorRegistrarContrato
'If ValidaDatos Then
'    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oLog As DLogGeneral
'        Set oLog = New DLogGeneral
'
'        If oLog.RegistrarContrato(Trim(Me.txtNContrato.Text), Trim(Me.txtPersona.Text), Trim(Me.txtArea.Text), Format(Me.txtFecIni.value, "DD/MM/YYYY"), _
'        Format(Me.txtFecFin.value, "DD/MM/YYYY"), CInt(Trim(Right(Me.cboMoneda.Text, 2))), CDbl(Me.txtMonto.Text), CInt(Trim(Me.spnPlazo.Valor)), _
'        CInt(Trim(Right(Me.cboTipoPago.Text, 2))), CInt(Trim(Right(Me.cboTipoContrato.Text, 2))), Format(Me.txtFechaFirma.value, "DD/MM/YYYY"), _
'        Trim(Me.txtGlosa.Text), Trim(Me.lblNombreArchivo.Caption), CInt(Trim(Right(Me.cboTipoGarantia.Text, 2))), CInt(Trim(Right(Me.cboMonedaGar.Text, 2))), _
'        CDbl(Trim(Me.txtMontoGar.Text)), 1) = 0 Then
'
'            If Trim(Me.lblNombreArchivo.Caption) <> "" Then
'                GrabarArchivo
'            End If
'
'            'REGISTRAR CRONOGRAMA
'            For I = 0 To (CInt(Me.lblNPago.Caption) - 2)
'                If oLog.RegistrarCronogramaContrato(Trim(Me.txtNContrato.Text), CInt(fMatCronograma(1, I + 1)), Format(fMatCronograma(2, I + 1), "DD/MM/YYYY"), _
'                 CInt(fMatCronograma(3, I + 1)), CDbl(fMatCronograma(5, I + 1)), 1, 0) = 1 Then
'                    MsgBox "No se registro el Nº  de Pago: " & fMatCronograma(1, I + 1), vbInformation, "Aviso"
'                    Exit Sub
'                End If
'            Next I
'
'            MsgBox "Contrato grabado Satisfactoriamente", vbInformation, "Aviso"
'            LimpiarDatos
'        Else
'            MsgBox "No se grabaron los datos de Contrato", vbInformation, "Aviso"
'        End If
'    End If
'End If
'Exit Sub
'ErrorRegistrarContrato:
'   MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
'End Sub
Private Sub cmdGrabar_Click()
    On Error GoTo ErrorRegistrarContrato
    Dim olog As DLogGeneral
'    Dim oDMov As DMov CPASI
'    Dim lsMovNro As String CPASI
'    Dim lnMovNro As Long CPASI
    Dim bTrans As Boolean
    Dim lnTpoContratacion As Integer
    Dim nContRef, nContItemRel As Integer 'PASI20140822 TI-ERS077-2014
    Dim Datoscontrato() As TContratoBS
    Dim Index As Integer, indexObj As Integer
    Dim lsSubCta As String
    Dim lnMovItem As Integer
    Dim iref As Integer
    Dim lnMonto As Currency
    
    nContRef = 0
    nContItemRel = 0
    If Not ValidaDatos Then Exit Sub
    'lnTpoContratacion = CInt(Trim(Right(cboTipoContratacion.Text, 4)))
    If MsgBox("Esta seguro de grabar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Set olog = New DLogGeneral
    'CONTRATO
    'Modificado PASI20140723 TI-ERS077-2014
'    oLog.RegistrarContrato_NEW Trim(txtNContrato.Text), Trim(txtPersona.Text), Trim(txtArea.Text), Format(txtFecIni.value, "DD/MM/YYYY"), _
'        Format(txtFecFin.value, "DD/MM/YYYY"), CInt(Trim(Right(cboMoneda.Text, 2))), CDbl(txtMonto.Text), CInt(Trim(spnPlazo.Valor)), _
'        CInt(Trim(Right(cboTipoPago.Text, 2))), CInt(Trim(Right(cboTipoContrato.Text, 2))), Format(txtFechaFirma.value, "DD/MM/YYYY"), _
'        Trim(txtGlosa.Text), Trim(lblNombreArchivo.Caption), CInt(Trim(Right(cboTipoGarantia.Text, 2))), CInt(Trim(Right(cboMonedaGar.Text, 2))), _
'        CDbl(Trim(txtMontoGar.Text)), 1, lnTpoContratacion, Trim(txtObjeto.Text) & IIf(lnTpoContratacion = 2, fsSubCta, "")

'        oLog.RegistraContratoProveedor Trim(txtNContrato.Text), Trim(txtPersona.Text), Trim(txtArea.Text), Trim(txtArea2.Text), Format(txtFecIni.value, "DD/MM/YYYY"), _ 'PASI BORRAR LUEGO
'        Format(txtFecFin.value, "DD/MM/YYYY"), CInt(Trim(Right(cboMoneda.Text, 2))), CDbl(txtMonto.Text), CInt(Trim(spnPlazo.Valor)), _
'        CInt(Trim(Right(cboTipoPago.Text, 2))), CInt(Trim(Right(cboTipoContrato.Text, 2))), Format(txtFechaFirma.value, "DD/MM/YYYY"), _
'        Trim(txtGlosa.Text), Trim(lblNombreArchivo.Caption), _
'        IIf(cboobjContrato.Visible = True, IIf(Val(Trim(Right(cboobjContrato.Text, 2))) = 1, 0, Val(Trim(Right(cboTipoGarantia.Text, 2)))), 0), _
'        IIf(cboobjContrato.Visible = True, IIf(Val(Trim(Right(cboobjContrato.Text, 2))) = 1, 0, Val(Trim(Right(cboMonedaGar.Text, 2)))), 0), _
'        IIf(cboobjContrato.Visible = True, IIf(Val(Trim(Right(cboobjContrato.Text, 2))) = 1, 0, Val(Trim(txtMontoGar.Text))), 0), 1, _
'        IIf(txtObjeto.Visible = True, Trim(txtObjeto.Text), "") & IIf(lnTpoContratacion = 1 Or lnTpoContratacion = 2 Or lnTpoContratacion = 3, fsSubCta, ""), _
'        IIf(cboobjContrato.Visible = True, Val(Right(cboobjContrato.Text, 2)), 0)
    
    If fraItemContrato.Enabled = True Then
        For Index = 1 To feOrden.Rows - 1 'ERS0772014
            ReDim Preserve Datoscontrato(Index)
            Datoscontrato(Index).sAgeCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 1))))
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes _
                Or fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 7))))
            ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
                lsSubCta = ""
                For indexObj = 1 To feObj.Rows - 1
                    If feObj.TextMatrix(indexObj, 1) = feOrden.TextMatrix(Index, 0) Then
                        lsSubCta = lsSubCta & feObj.TextMatrix(indexObj, 5)
                    End If
                Next
                Datoscontrato(Index).sCtaContCod = Trim(CStr(Trim(feOrden.TextMatrix(Index, 2)))) & lsSubCta
            End If
            Datoscontrato(Index).sObjeto = CStr(Trim(feOrden.TextMatrix(Index, 2)))
            Datoscontrato(Index).sDescripcion = CStr(Trim(feOrden.TextMatrix(Index, 3)))
            Datoscontrato(Index).nCantidad = Val(feOrden.TextMatrix(Index, 4))
            Datoscontrato(Index).nTotal = feOrden.TextMatrix(Index, 6)
        Next
    End If
    
    olog.dBeginTrans
    bTrans = True
    If Trim(lblNombreArchivo.Caption) <> "" Then
        GrabarArchivo
    End If
    nContRef = olog.RegistraContratoProveedor(Trim(txtNContrato.Text), _
                                 Trim(txtPersona.Text), _
                                 Trim(txtArea.Text), _
                                 Trim(txtArea2.Text), _
                                 Format(txtFecIni.value, "DD/MM/YYYY"), _
                                 Format(txtFecFin.value, "DD/MM/YYYY"), _
                                 CInt(Trim(Right(cboMoneda.Text, 2))), _
                                 CDbl(txtMonto.Text), _
                                 IIf(fraCronograma.Enabled = False, 0, CInt(Trim(spnPlazo.Valor))), _
                                 CInt(Trim(Right(cboTipoPago.Text, 2))), _
                                 IIf(cboobjContrato.Visible = True, Val(Right(cboobjContrato.Text, 2)), 0), _
                                 Format(txtFechaFirma.value, "DD/MM/YYYY"), _
                                 Trim(Replace(Replace(txtGlosa.Text, Chr(10), ""), Chr(13), "")), _
                                 Trim(lblNombreArchivo.Caption), _
                                 IIf(cboobjContrato.Visible = True, IIf(Val(Trim(Right(cboobjContrato.Text, 2))) = 1, 0, Val(Trim(Right(cboTipoGarantia.Text, 2)))), 0), _
                                 IIf(cboobjContrato.Visible = True, IIf(Val(Trim(Right(cboobjContrato.Text, 2))) = 1, 0, Val(Trim(Right(cboMonedaGar.Text, 2)))), 0), _
                                 IIf(cboobjContrato.Visible = True, IIf(Val(Trim(Right(cboobjContrato.Text, 2))) = 1, 0, Val(Trim(txtMontoGar.Text))), 0), _
                                 1, _
                                 CInt(Trim(Right(cboTipoContrato.Text, 2))), _
                                 IIf(txtObjeto.Visible = True, Trim(txtObjeto.Text), "") & IIf(lnTpoContratacion = 1 Or lnTpoContratacion = 2 Or lnTpoContratacion = 3, fsSubCta, ""))

        If fraCronograma.Enabled = True Then
              For I = 1 To UBound(fMatCronograma, 2)
                  olog.RegistrarCronogramaContrato_NEW Trim(txtNContrato.Text), CInt(fMatCronograma(1, I)), Format(fMatCronograma(2, I), "DD/MM/YYYY"), CInt(fMatCronograma(3, I)), CDbl(IIf(fMatCronograma(5, I) = "-", 0, fMatCronograma(5, I))), 1, 0, nContRef
                  If fraItemContrato.Enabled = True Then
                    If UBound(Datoscontrato) > 0 Then 'PASIERS0772014
                        If fntpodocorigen = LogTipoContrato.ContratoServicio Then
                            For lnMovItem = 1 To UBound(Datoscontrato)
                               olog.RegistrarContratoServicio Trim(txtNContrato.Text), nContRef, 0, CInt(fMatCronograma(1, I)), Datoscontrato(lnMovItem).sAgeCod, Datoscontrato(lnMovItem).sCtaContCod, Datoscontrato(lnMovItem).sDescripcion, lnMovItem, IIf(Trim(Right(cboTipoPago.Text, 2)) = 2, 0, Datoscontrato(lnMovItem).nTotal)
                            Next lnMovItem
                        End If
                End If
        End If
              Next I
        End If
        If fraItemContrato.Enabled = True Then
            If UBound(Datoscontrato) > 0 Then 'PASI20140210 ERS0772014
                If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
                    fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                        For lnMovItem = 1 To UBound(Datoscontrato)
                            olog.RegistrarContratoBienes Trim(txtNContrato.Text), nContRef, 0, Datoscontrato(lnMovItem).sAgeCod, Datoscontrato(lnMovItem).sCtaContCod, Datoscontrato(lnMovItem).sDescripcion, lnMovItem, Datoscontrato(lnMovItem).sObjeto, Datoscontrato(lnMovItem).nCantidad, Datoscontrato(lnMovItem).nTotal, 1, "", ""
    '                ElseIf fnTpoDocOrigen = LogTipoContrato.ContratoServicio Then
    '                        olog.RegistrarContratoServicio Trim(txtNContrato.Text), nContRef, Datoscontrato(lnMovItem).sAgeCod, Datoscontrato(lnMovItem).sCtaContCod, lnMovItem, Datoscontrato(lnMovItem).nCantidad, Datoscontrato(lnMovItem).nTotal
                        Next lnMovItem
                End If
            End If
        End If
        olog.RegistraSaldoContrato Trim(txtNContrato.Text), nContRef, CDbl(txtMonto.Text) 'PASI20140828 ERS0772014
    olog.dCommitTrans
    'end PASI
    Screen.MousePointer = 0
    MsgBox "Contrato grabado satisfactoriamente", vbInformation, "Aviso"
    Set olog = Nothing
        'ARLO 20160126 ***
        gsOpeCod = LogPistaRegistroContrato
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Grabo Contrato N° : " & txtNContrato.Text & " | Por un Monto : " & txtMonto.Text, txtPersona.Text, 3
        Set objPista = Nothing
        '***
    LimpiarDatos
    bTrans = False
    Exit Sub
ErrorRegistrarContrato:
    Screen.MousePointer = 0
    If bTrans Then
        olog.dRollbackTrans
        Set olog = Nothing
    End If
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"
End Sub
Sub LimpiarDatos()
'PROVEEDOR
Me.txtPersona.Text = ""
Me.txtProvNom.Text = ""

'AREA USUARIA
Me.txtArea.Text = ""
Me.txtAgeDesc.Text = ""
Me.txtArea2.Text = "" 'PASI20140816 TI-ERS077-2014
Me.txtAgeDesc2.Text = "" 'PASI20140816 TI-ERS077-2014

'PERIODO CONTRAACTUAL
Me.txtFecIni.value = gdFecSis
Me.txtFecFin.value = gdFecSis
cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, 0)
Me.txtMonto.Text = ""
Me.lblNPago.Caption = "1"
Me.spnPlazo.Enabled = True 'PASIERS20141227
Me.spnPlazo.Valor = "0"

'DATOS CONTRATO
cboTipoPago.ListIndex = IndiceListaCombo(cboTipoPago, 0)
cboTipoContrato.ListIndex = IndiceListaCombo(cboTipoContrato, 0)
Me.lblNombreArchivo.Caption = ""
Me.txtNContrato.Text = ""
Me.txtFecFin.value = gdFecSis
Me.txtGlosa.Text = ""
EstadoObjeto False 'PASI20140816 TI-ERS077-2014
'EJVG20131007 ***
cboTipoContratacion.ListIndex = IndiceListaCombo(cboTipoContratacion, 0)
txtObjeto.Text = ""
'END EJVG *******

'GARANTIA
cboTipoGarantia.ListIndex = IndiceListaCombo(cboTipoGarantia, 0)
cboMonedaGar.ListIndex = IndiceListaCombo(cboMonedaGar, 0)
Me.txtMontoGar.Text = ""

'CRONOGRAMA
Me.txtFechaPago.value = gdFecSis
cboMonedaCro.ListIndex = IndiceListaCombo(cboMonedaCro, 0)
Me.txtMontoCro.Text = ""
Call LimpiaFlex(Me.feCronograma)
ReDim Preserve fMatCronograma(5, 1 To 1)

'PASI20140724 TI-ERS077-2014
'Items del Contrato
Call LimpiaFlex(feOrden)
ReiniciarEstadoFrame
'END PASI
fbEstCuota = False 'PASI20150107
End Sub
Private Sub cmdQuitar_Click()
If CInt(lblNPago.Caption) > 1 Then
    If MsgBox("Estas seguro de quitar el ultimo registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Me.lblNPago.Caption = CInt(lblNPago.Caption) - 1
        If CInt(lblNPago.Caption) > 1 Then
            ReDim Preserve fMatCronograma(5, 1 To (CInt(lblNPago.Caption) - 1))
        End If
        Call LeerMatriz(CInt(lblNPago.Caption) - 1)
    End If
Else
    MsgBox "No hay datos a eliminar", vbInformation, "Aviso"
End If
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub feOrden_OnCellChange(pnRow As Long, pnCol As Long) 'PASI20140110 ERS0772014
    On Error GoTo ErrfeOrden_OnCellChange
    If feOrden.TextMatrix(1, 0) <> "" Then
        If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
            fntpodocorigen = LogTipoContrato.ContratoSuministro Or _
            fntpodocorigen = LogTipoContrato.ContratoServicio Then
            If pnCol = 4 Or pnCol = 5 Then
                feOrden.TextMatrix(pnRow, 6) = Format(Val(feOrden.TextMatrix(pnRow, 4)) * feOrden.TextMatrix(pnRow, 5), gsFormatoNumeroView)
            End If
        End If
    End If
    Exit Sub
ErrfeOrden_OnCellChange:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feOrden_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean) 'PASI20140110 ERS0772014
    If psDataCod <> "" Then
        If pnCol = 2 Then
            If fntpodocorigen = LogTipoContrato.ContratoServicio Then
                AsignaObjetosSerItem psDataCod
            End If
        End If
        If pnCol = 1 Or pnCol = 2 Then
            '*** Si esta vacio el campo de la cuenta contable y si ya eligió agencia y objeto
            If Len(Trim(feOrden.TextMatrix(pnRow, 1))) <> 0 And Len(Trim(feOrden.TextMatrix(pnRow, 2))) <> 0 Then
                feOrden.TextMatrix(pnRow, 7) = DameCtaCont(feOrden.TextMatrix(pnRow, 2), 0, Trim(feOrden.TextMatrix(pnRow, 1)))
            End If
            '***
        End If
    End If
End Sub
Private Sub feOrden_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean) 'PASI20140110 ERS0772014
    Dim sColumnas() As String
    sColumnas = Split(feOrden.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub feOrden_RowColChange()  'PASI20140110 ERS0772014
    If feOrden.col = 1 Then
        feOrden.rsTextBuscar = fRsAgencia
    ElseIf feOrden.col = 2 Then
        If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Or _
            fntpodocorigen = LogTipoContrato.ContratoSuministro Then
            feOrden.rsTextBuscar = fRsCompra
        ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
            feOrden.rsTextBuscar = fRsServicio
        End If
    End If
End Sub
Private Sub Form_Load()
Dim oConst As DConstantes
Dim oArea As DActualizaDatosArea
Dim oConstSist As NConstSistemas
Dim olog As DLogGeneral 'PASI20140718 TI-ERS077-2014
Set olog = New DLogGeneral

Set oConst = New DConstantes
Set oArea = New DActualizaDatosArea

txtArea.rs = oArea.GetAgenciasAreas
txtArea2.rs = oArea.GetAgenciasAreas 'PASI20140718 TI-ERS077-2014

CargaCombo oConst.GetConstante(gMoneda), Me.cboMoneda
CargaCombo oConst.GetConstante(gMoneda), Me.cboMonedaGar
CargaCombo oConst.GetConstante(gMoneda), Me.cboMonedaCro
CargaCombo oConst.GetConstante(gsLogContTipoPagoContratos), Me.cboTipoPago
CargaCombo olog.CargaTipoContrato(), Me.cboTipoContrato
CargaCombo oConst.GetConstante(gsLogContTipoGarantia), Me.cboTipoGarantia
CargaCombo oConst.GetConstante(gsLogContObjContrato), Me.cboobjContrato 'EJVG20131007 'Se cambio Tipo de Contrato por Objeto de Contrato PASI20140717 TI-ERS077-2014
'CargaCombo oConst.GetConstante(gsLogContTipoContratacion), Me.cboTipoContratacion 'Comentado por PASI20140717 TI-ERS077-2014

'Agregado por PASI20140718
    Label21.Visible = False
    cboTipoContratacion.Visible = False
    EstadoObjeto (False)
    EstadoTab 1
    fraGarantia.Enabled = False
    fraCronograma.Enabled = False
    fraItemContrato.Enabled = False
'end PASI
Me.txtFecIni.value = gdFecSis
Me.txtFecFin.value = gdFecSis
Me.txtFechaFirma.value = gdFecSis
Me.txtFechaPago = gdFecSis

If Trim(Mid(GetMaquinaUsuario, 1, 2)) = "01" Then
    Me.cmdBuscarArchivo.Enabled = True
    pbActivaArchivo = True
    
    'OBTENER RUTA DE CONTRATOS
    Set oConstSist = New NConstSistemas
    psRutaContrato = Trim(oConstSist.LeeConstSistema(gsLogContRutaContratos))
    
Else
    Me.cmdBuscarArchivo.Enabled = False
    pbActivaArchivo = False
    psRutaContrato = ""
End If
Me.lblNPago.Caption = 1
CargaVariables
fbEstCuota = False 'PASI20150107
End Sub
Private Sub CargaVariables()
    Dim oArea As New DActualizaDatosArea
    Dim oALmacen As New DLogAlmacen
    Dim rs As New ADODB.Recordset
    
    If gbBitTCPonderado Then
        fnTpoCambio = gnTipCambioPonderado
    Else
        fnTpoCambio = gnTipCambioC
    End If
    
    Set fRsAgencia = oArea.GetAgencias(, , True)
    Set fRsCompra = oALmacen.GetBienesAlmacen(, "11','12','13")
    'Set fRsServicio = OrdenServicio()

    Set rs = Nothing
    Set oALmacen = Nothing
    Set oArea = Nothing
End Sub
Sub GrabarArchivo()
If Trim(fsRuta) <> "" Then
Dim RutaFinal As String
RutaFinal = psRutaContrato
Dim a As New Scripting.FileSystemObject

If a.FolderExists(RutaFinal) = False Then
    a.CreateFolder (RutaFinal)
End If
Copiar fsRuta, RutaFinal & Trim(lblNombreArchivo.Caption)
Else
    MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
End If
End Sub

Private Sub Copiar(Archivo As String, Destino As String)
Dim a As New Scripting.FileSystemObject

If a.FileExists(Destino) = False Then
    a.CopyFile Archivo, Destino
Else
    MsgBox "Archivo ya existe", vbInformation, "Aviso"
End If
End Sub

Private Sub spnPlazo_Change()
If Not fbEstCuota Then Exit Sub 'PASI20150107
If CInt(lblNPago.Caption) > 1 Then
    If CInt(Me.spnPlazo.Valor) < (CInt(Me.lblNPago.Caption) - 1) Then
        If MsgBox("Esta Seguro de eliminar el ultimo registro del cronograma?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            'Me.lblNPago.Caption = CInt(Me.lblNPago.Caption) - 1 Comentado PASI20150107

            If CInt(Me.spnPlazo.Valor) > 0 Then
                ReDim Preserve fMatCronograma(5, 1 To CInt(Me.spnPlazo.Valor))
                Me.lblNPago.Caption = Me.spnPlazo.Valor + 1
            End If
            Call LeerMatriz(CInt(Me.spnPlazo.Valor))
        Else
            Me.spnPlazo.Valor = CInt(Me.spnPlazo.Valor) + 1
        End If
    End If
End If
fbEstCuota = False 'PASI20150107
End Sub
'PASI20140721 TI-ERS077-2014
Private Sub spnPlazo_DownClick()
    Dim nValor As Integer
    nValor = CInt(IIf(Right(cboTipoContrato.Text, 4) = "", 0, Right(cboTipoContrato.Text, 4)))
    If Me.spnPlazo.Valor < 1 And (nValor = 1 Or nValor = 2 Or nValor = 3 Or cboTipoContrato.ListIndex = -1) Then
        EstadoTab 1
    End If
End Sub
Private Sub spnPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fbEstCuota = True 'PASI20150107
        cboTipoPago.SetFocus
    End If
End Sub
Private Sub spnPlazo_UpClick()
    Dim nValor As Integer
    nValor = CInt(IIf(Right(cboTipoContrato.Text, 4) = "", 0, Right(cboTipoContrato.Text, 4)))
    If (nValor = 1 Or nValor = 2 Or nValor = 3 Or cboTipoContrato.ListIndex = -1) Then
        EstadoTab (2)
    End If
End Sub
'end PASI
Private Sub TxtArea_EmiteDatos()
    Me.txtAgeDesc.Text = txtArea.psDescripcion
    If txtArea.psDescripcion <> "" Then
        Me.txtArea2.SetFocus 'PASI20140718 TI-ERS077-2014
    End If
End Sub
Private Sub txtArea2_EmiteDatos() 'PASI20140718 TI-ERS077-2014
     Me.txtAgeDesc2.Text = txtArea2.psDescripcion
    If txtArea2.psDescripcion <> "" Then
        Me.txtFecIni.SetFocus
    End If
End Sub
Private Sub txtFecFin_Change()
If CDate(txtFecFin.value) < CDate(Me.txtFecIni.value) Then
    MsgBox "Fecha Final no puede ser menor a la Fecha Inicial.", vbInformation, "Aviso"
    txtFecFin.value = Me.txtFecIni.value
End If
End Sub
Private Sub txtFecIni_Change()
If CDate(txtFecFin.value) < CDate(Me.txtFecIni.value) Then
    MsgBox "Fecha Inicial no puede ser mayor a la Fecha Final.", vbInformation, "Aviso"
    Me.txtFecIni.value = txtFecFin.value
End If
End Sub
Private Sub txtMonto_GotFocus()
fEnfoque txtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        Me.spnPlazo.SetFocus
    End If
End Sub
Private Sub txtMonto_LostFocus()
If Trim(txtMonto.Text) = "" Then
        txtMonto.Text = "0.00"
    End If
    txtMonto.Text = Format(txtMonto.Text, "#0.00")
End Sub
Private Sub txtMontoCro_GotFocus()
fEnfoque txtMontoCro
End Sub
Private Sub txtMontoCro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoCro, KeyAscii, 10, 3)
 If KeyAscii = 13 Then
        Me.cmdAgregar.SetFocus
    End If
End Sub
Private Sub txtMontoCro_LostFocus()
If Trim(txtMontoCro.Text) = "" Then
        txtMontoCro.Text = "0.00"
    End If
    txtMontoCro.Text = Format(txtMontoCro.Text, "#0.00")
End Sub
Private Sub txtMontoGar_GotFocus()
fEnfoque txtMontoGar
End Sub
Private Sub txtMontoGar_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoGar, KeyAscii, 10, 3)
 If KeyAscii = 13 Then
    If fraCronograma.Enabled = True Then
        Me.txtFechaPago.SetFocus
    End If
End If
End Sub

Private Sub txtMontoGar_LostFocus()
If Trim(txtMontoGar.Text) = "" Then
        txtMontoGar.Text = "0.00"
    End If
    txtMontoGar.Text = Format(txtMontoGar.Text, "#0.00")
End Sub
Private Sub txtNContrato_Change()
If txtNContrato.SelStart > 0 Then
    I = Len(Mid(txtNContrato.Text, 1, txtNContrato.SelStart))
End If
txtNContrato.Text = UCase(txtNContrato.Text)
txtNContrato.SelStart = I

If pbActivaArchivo Then
    If Trim(Me.lblNombreArchivo.Caption) <> "" Then
        MsgBox "Debe cargar el archivo de contrato nuevamente.", vbInformation, "Aviso"
        fsRuta = ""
        Me.lblNombreArchivo.Caption = ""
    End If
End If
End Sub
Private Sub txtNContrato_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNContrato, KeyAscii, 20, 3)
If KeyAscii = 13 Then
    Me.txtGlosa.SetFocus
End If
End Sub
Private Sub txtNContrato_LostFocus()
Me.txtNContrato.Text = UCase(Trim(Me.txtNContrato.Text))
End Sub

Private Sub txtPersona_EmiteDatos()
Dim rs As ADODB.Recordset
On Error GoTo ErrorPersona
 
    Set rs = New ADODB.Recordset
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    Me.txtProvNom.Text = txtPersona.psDescripcion
    lsCodPers = txtPersona.psCodigoPersona
    If txtPersona.psDescripcion <> "" Then
        Set rs = oProv.GetProveedorAgeRetBuenCont(lsCodPers)
        If rs.EOF And rs.BOF Then
            'MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
            'Me.txtProvNom.Text = ""
            'Exit Sub
        Else
           ' Me.chkBuneCOnt.value = IIf(rs.Fields(1), 1, 0)
            'Me.chkRetencion.value = IIf(rs.Fields(0), 1, 0)
        End If
        Me.txtArea.SetFocus
    End If
    
    Exit Sub
ErrorPersona:
    MsgBox Err.Description, vbInformation, "Error"
End Sub
'EJVG20131009 ***
Private Function OrdenServicio() As ADODB.Recordset
    Dim oCon As New DConecta
    Dim sSqlO As String
    Dim lnMoneda As Integer
    If cboMoneda.Text <> "" Then
        lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
        oCon.AbreConexion
        sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
              & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
              & "WHERE b.cCtaContCod = a.cCtaContCod AND (a.cOpeCod='" & IIf(lnMoneda = 1, "501207", "502207") & "' AND (a.cOpeCtaDH='D'))"
        Set OrdenServicio = oCon.CargaRecordSet(sSqlO)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function
Private Sub txtObjeto_EmiteDatos()
    Dim lnTpoContratacion As Integer
    fsSubCta = ""
    If txtObjeto.Text <> "" Then
        If cboMoneda.Text <> "" Then
            'lnTpoContratacion = CInt(Trim(Right(cboTipoContratacion.Text, 2))) 'Comentado PASI20140719 TI-ERS077-2014
            lnTpoContratacion = CInt(Trim(Right(cboTipoContrato.Text, 2))) 'PASI20140719 TI-ERS077-2014
            If (lnTpoContratacion = 1 Or lnTpoContratacion = 2 Or lnTpoContratacion = 3) Then 'Se agrego 1 y 3
                fsSubCta = AsignaObjetosSer(txtObjeto.Text)
            End If
        End If
    End If
End Sub
Private Function AsignaObjetosSer(ByVal sCtaCod As String) As String
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As New DConecta
    Dim oCtaCont As New DCtaCont
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim oRHAreas As New DActualizaDatosArea
    Dim oCtaIf As New NCajaCtaIF
    Dim oEfect As New Defectivo
    Dim oDescObj As New ClassDescObjeto
    Dim oContFunct As New NContFunciones
    Dim lsRaiz As String, lsFiltro As String, sSQL As String
    Dim lsSubCta As String
        
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
    'EliminaObjeto feOrden.row
    lsSubCta = ""

    sSQL = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSQL)
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOK Then
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                'AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                                lsSubCta = lsSubCta & lsFiltro
                            Else
                                'EliminaObjeto feOrden.row
                                lsSubCta = ""
                                Exit Do
                            End If
                        Else
                            'AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                            lsSubCta = lsSubCta & lsFiltro
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If
    
    AsignaObjetosSer = lsSubCta

    Set rs = Nothing
    Set rs1 = Nothing
    Set oDescObj = Nothing
    Set oCon = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
    Set oContFunct = Nothing
    Set oContFunct = Nothing
    Exit Function
End Function
Private Function validaIngresoRegistros() As Boolean 'PASI20140110
    Dim I As Long, J As Long
    Dim col As Integer
    Dim Columnas() As String
    Dim lsColumnas As String
    
    lsColumnas = "1,2,6"
    Columnas = Split(lsColumnas, ",")
        
    validaIngresoRegistros = True
    If feOrden.TextMatrix(1, 0) <> "" Then
        For I = 1 To feOrden.Rows - 1
            For J = 1 To feOrden.Cols - 1
                For col = 0 To UBound(Columnas)
                    If Trim(Right((cboTipoContrato.Text), 4)) <> LogTipoContrato.ContratoSuministro Then
                        If J = Columnas(col) Then
                            If Len(Trim(feOrden.TextMatrix(I, J))) = 0 And feOrden.ColWidth(J) <> 0 Then
                                MsgBox "Ud. debe especificar el campo " & feOrden.TextMatrix(0, J), vbInformation, "Aviso"
                                validaIngresoRegistros = False
                                feOrden.TopRow = I
                                feOrden.row = I
                                feOrden.col = J
                                feOrden_RowColChange
                                Exit Function
                            End If
                        End If
                    End If
                Next
            Next
            If Trim(Right(cboTipoContrato.Text, 4)) = LogTipoContrato.ContratoServicio Then
                If Trim(Right(cboTipoPago.Text, 4)) <> 2 Then  'Variable
                        If IsNumeric(feOrden.TextMatrix(I, 6)) Then
                            If CCur(feOrden.TextMatrix(I, 6)) <= 0 Then
                                MsgBox "El Importe Total debe ser mayor a cero", vbInformation, "Aviso"
                                validaIngresoRegistros = False
                                feOrden.TopRow = I
                                feOrden.row = I
                                feOrden.col = 6
                                Exit Function
                            End If
                        Else
                            MsgBox "El Importe Total debe ser númerico", vbInformation, "Aviso"
                            validaIngresoRegistros = False
                            feOrden.TopRow = I
                            feOrden.row = I
                            feOrden.col = 6
                            Exit Function
                        End If
                End If
            Else
                 If IsNumeric(feOrden.TextMatrix(I, 6)) Then
                            If CCur(feOrden.TextMatrix(I, 6)) <= 0 Then
                                MsgBox "El Importe Total debe ser mayor a cero", vbInformation, "Aviso"
                                validaIngresoRegistros = False
                                feOrden.TopRow = I
                                feOrden.row = I
                                feOrden.col = 6
                                Exit Function
                            End If
                Else
                    MsgBox "El Importe Total debe ser númerico", vbInformation, "Aviso"
                    validaIngresoRegistros = False
                    feOrden.TopRow = I
                    feOrden.row = I
                    feOrden.col = 6
                    Exit Function
                End If
            End If
            If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Then   ' Or _
                'fntpodocorigen = LogTipoContrato.ContratoSuministro Then
                    If Len(Trim(feOrden.TextMatrix(I, 7))) = 0 Then
                        MsgBox "El Objeto " & feOrden.TextMatrix(I, 3) & Chr(10) & "no tiene configurado Plantilla Contable, consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
                        feOrden.TopRow = I
                        feOrden.row = I
                        feOrden.col = 2
                        validaIngresoRegistros = False
                        Exit Function
                    End If
            End If
        Next
    Else
        MsgBox "Ud. debe agregar los Bienes/Servicios a dar Conformidad", vbInformation, "Aviso"
        validaIngresoRegistros = False
    End If
End Function
'END EJVG *******
'Agregado por PASI20140718 TI-ERS077-2014
Private Sub cmdAgregarItemCont_Click()
If Not validaBusqueda Then Exit Sub
    If feOrden.TextMatrix(1, 0) <> "" Then
        If Not validaIngresoRegistros Then Exit Sub
    End If
    feOrden.AdicionaFila
    
    If fntpodocorigen = LogTipoContrato.ContratoAdqBienes Then 'Or _
       'fntpodocorigen = LogTipoContrato.ContratoSuministro
        feOrden.ColumnasAEditar = "X-1-2-3-4-5-X-X"
        feOrden.TextMatrix(feOrden.row, 4) = "0"
        feOrden.TextMatrix(feOrden.row, 5) = "0.00"
        feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    ElseIf fntpodocorigen = LogTipoContrato.ContratoSuministro Then
        feOrden.ColumnasAEditar = "X-X-2-3-4-5-X-X"
        feOrden.TextMatrix(feOrden.row, 1) = "-"
        feOrden.TextMatrix(feOrden.row, 4) = "0"
        feOrden.TextMatrix(feOrden.row, 5) = "0.00"
        feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    ElseIf fntpodocorigen = LogTipoContrato.ContratoServicio Then
        'Condicionar el tipo de pago
        If Trim(Right(cboTipoPago.Text, 4)) = 2 Then 'Variable
            'feOrden.ColumnasAEditar = "X-1-2-3-X-X-6-X" Comentado PASI20150108
            feOrden.TextMatrix(feOrden.row, 6) = "0.00"  'PASI20150108
            feOrden.ColumnasAEditar = "X-1-2-3-X-X-X-X" 'PASI20150108
        Else
            feOrden.ColumnasAEditar = "X-1-2-3-X-X-6-X"
        End If
    End If
    feOrden.TextMatrix(feOrden.row, 6) = "0.00"
    feOrden.col = 2
    feOrden.SetFocus
    feOrden_RowColChange
End Sub
Private Function validaBusqueda()
     validaBusqueda = True
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. primero debe de seleccionar el Tipo de Moneda", vbInformation, "Aviso"
        validaBusqueda = False
        cboMoneda.SetFocus
        Exit Function
    End If
    If Len(txtMonto.Text) = 0 Then
        MsgBox "Ud. primero debe de Ingresar el Monto del Contrato", vbInformation, "Aviso"
        validaBusqueda = False
        cboMoneda.SetFocus
        Exit Function
    End If
End Function
Public Function FormaSelect(psOpeCod As String, sObj As String, nNiv As Integer, psAgeCod As String) As String
    Dim sText As String
    sText = " SELECT b.cCtaContCod cObjetoCod, b.cCtaContDesc cObjetoDesc, e.cBSCod cObjCod," _
          & " upper(e.cBSDescripcion) as cObjDesc, 1 nObjetoNiv, CO.cConsDescripcion " _
          & " FROM  CtaCont b " _
          & " Inner JOIN CtaBS  c ON Replace(c.cCtaContCod,'AG','" & psAgeCod & "') = b.cCtaContCod And cOpeCod = '" & psOpeCod & "'" _
          & " Inner JOIN BienesServicios e ON e.cBSCod like c.cObjetoCod + '%'" _
          & " Inner Join Constante CO On nBSunidad = CO.nConsValor And nConsCod = '1019'" _
          & " "
    If nNiv > 0 Then
       sText = sText & "WHERE d.nObjetoNiv = " & nNiv & " "
    End If
    FormaSelect = sText & IIf(sObj <> "", " And e.cBSCod = '" & sObj & "' ", sObj) _
                & "ORDER BY e.cBSCod"
End Function
Private Sub cmdQuitarItemCont_Click()
    feOrden.EliminaFila feOrden.row
End Sub
Private Sub AsignaObjetosSerItem(ByVal sCtaCod As String)
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As New DConecta
    Dim oCtaCont As New DCtaCont
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim oRHAreas As New DActualizaDatosArea
    Dim oCtaIf As New NCajaCtaIF
    Dim oEfect As New Defectivo
    Dim oDescObj As New ClassDescObjeto
    Dim oContFunct As New NContFunciones
    Dim lsRaiz As String, lsFiltro As String, sSQL As String
        
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
    EliminaObjeto feOrden.row

    sSQL = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSQL)
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOK Then
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                            Else
                                EliminaObjeto feOrden.row
                                Exit Do
                            End If
                        Else
                            AdicionaObjeto feOrden.TextMatrix(feOrden.row, 0), IIf(IsNull(rs1!nCtaObjOrden), "", rs1!nCtaObjOrden), oDescObj.gsSelecCod, oDescObj.gsSelecDesc, lsFiltro, IIf(IsNull(rs1!cObjetoCod), "", rs1!cObjetoCod)
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If

    Set rs = Nothing
    Set rs1 = Nothing
    Set oDescObj = Nothing
    Set oCon = Nothing
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
    Set oContFunct = Nothing
    Set oContFunct = Nothing
    Exit Sub
End Sub
Private Function DameCtaCont(ByVal psObjeto As String, nNiv As Integer, psAgeCod As String) As String 'PASI20140110ERS0772014
    Dim oCon As New DConecta
    Dim oForm As New frmLogOCompra
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = oForm.FormaSelect(gcOpeCod, psObjeto, 0, psAgeCod)
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    If Not rs.EOF Then
        DameCtaCont = rs!cObjetoCod
    End If
    Set rs = Nothing
    Set oForm = Nothing
    Set oCon = Nothing
End Function
Private Sub AdicionaObjeto(ByVal pnItem As Integer, ByVal psCtaObjOrden As String, ByVal psCodigo As String, ByVal psDesc As String, ByVal psFiltro As String, ByVal psObjetoCod As String)
    feObj.AdicionaFila
    feObj.TextMatrix(feObj.row, 1) = pnItem
    feObj.TextMatrix(feObj.row, 2) = psCtaObjOrden
    feObj.TextMatrix(feObj.row, 3) = psCodigo
    feObj.TextMatrix(feObj.row, 4) = psDesc
    feObj.TextMatrix(feObj.row, 5) = psFiltro
    feObj.TextMatrix(feObj.row, 6) = psObjetoCod
End Sub
'end PASI
