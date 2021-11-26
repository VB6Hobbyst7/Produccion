VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersGarantiasHC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Garantias de Cliente"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "frmPersGarantiasHC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   660
      Left            =   15
      TabIndex        =   128
      Top             =   6330
      Width           =   9150
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
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
         Left            =   120
         TabIndex        =   135
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   60
         TabIndex        =   134
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   390
         Left            =   1215
         TabIndex        =   133
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton CmdLimpiar 
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
         Height          =   390
         Left            =   6600
         TabIndex        =   132
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
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
         Left            =   2355
         TabIndex        =   131
         Top             =   180
         Width           =   1125
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
         Left            =   7740
         TabIndex        =   130
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
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
         Left            =   1230
         TabIndex        =   129
         ToolTipText     =   "Salir(ALT+S)"
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.Frame FraBuscaPers 
      Height          =   1275
      Left            =   0
      TabIndex        =   124
      Top             =   0
      Width           =   9180
      Begin VB.CommandButton CmdBuscaPersona 
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
         Height          =   375
         Left            =   6945
         TabIndex        =   126
         ToolTipText     =   "Busca Documentos de Persona"
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   125
         ToolTipText     =   "Pulse este Boton para Mostrar los Datos de la Garantia"
         Top             =   675
         Width           =   1425
      End
      Begin MSComctlLib.ListView LstGaratias 
         Height          =   975
         Left            =   90
         TabIndex        =   127
         Top             =   165
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   1720
         View            =   3
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Garantia"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "codemi"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "nomemi"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "tipodoc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "cnumdoc"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSGarant 
      Height          =   4860
      Left            =   15
      TabIndex        =   0
      Top             =   1395
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   8573
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Relac. de la Garantia"
      TabPicture(0)   =   "frmPersGarantiasHC.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPrinc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraRelaGar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos de Garantia"
      TabPicture(1)   =   "frmPersGarantiasHC.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "framontos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraZonaCbo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Garantia Real"
      TabPicture(2)   =   "frmPersGarantiasHC.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraDatInm"
      Tab(2).Control(1)=   "fraDatVehic"
      Tab(2).Control(2)=   "FraGar"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Declaración Jurada"
      TabPicture(3)   =   "frmPersGarantiasHC.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Poliza"
      TabPicture(4)   =   "frmPersGarantiasHC.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      Begin VB.Frame FraRelaGar 
         Caption         =   "Representantes Garantia"
         Height          =   1800
         Left            =   105
         TabIndex        =   90
         Top             =   2970
         Width           =   8760
         Begin VB.CommandButton CmdCliCancelar 
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
            Height          =   300
            Left            =   7560
            TabIndex        =   95
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton CmdCliAceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6525
            TabIndex        =   94
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton CmdCliNuevo 
            Caption         =   "&Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6540
            TabIndex        =   92
            Top             =   1440
            Width           =   1005
         End
         Begin VB.CommandButton CmdCliEliminar 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7560
            TabIndex        =   91
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   1440
            Width           =   1005
         End
         Begin SICMACT.FlexEdit FERelPers 
            Height          =   1155
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   8520
            _extentx        =   15240
            _extenty        =   2037
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "-Codigo-Nombre-Relacion-Aux"
            encabezadosanchos=   "400-1500-5000-1450-0"
            font            =   "frmPersGarantiasHC.frx":0396
            fontfixed       =   "frmPersGarantiasHC.frx":03C2
            columnasaeditar =   "X-1-X-3-X"
            listacontroles  =   "0-1-0-3-0"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "C-C-L-L-C"
            formatosedit    =   "0-0-0-0-0"
            lbeditarflex    =   -1  'True
            lbultimainstancia=   -1  'True
            tipobusqueda    =   3
            colwidth0       =   405
            rowheight0      =   300
         End
      End
      Begin VB.Frame FraDatInm 
         Caption         =   "Datos del Inmueble"
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         Height          =   2250
         Left            =   -74655
         TabIndex        =   103
         Top             =   720
         Width           =   8475
         Begin VB.CommandButton CmdBuscaInmob 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7800
            TabIndex        =   111
            Top             =   285
            Width           =   390
         End
         Begin VB.TextBox TxtTelefono 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1905
            TabIndex        =   110
            Top             =   690
            Width           =   1395
         End
         Begin VB.ComboBox CboTipoInmueb 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4890
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   660
            Width           =   3390
         End
         Begin VB.TextBox TxtMontoHip 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   4890
            MaxLength       =   15
            TabIndex        =   108
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   1035
            Width           =   1245
         End
         Begin VB.TextBox TxtPrecioVenta 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   1905
            MaxLength       =   15
            TabIndex        =   107
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   1410
            Width           =   1245
         End
         Begin VB.TextBox TxtValorCConst 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   4905
            MaxLength       =   15
            TabIndex        =   106
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   1440
            Width           =   1245
         End
         Begin VB.TextBox TxtHipCuotaIni 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Height          =   315
            Left            =   1905
            MaxLength       =   15
            TabIndex        =   105
            Tag             =   "txtPrincipal"
            Text            =   "0.00"
            Top             =   1020
            Width           =   1260
         End
         Begin VB.ComboBox cboEstadoTasInm 
            Height          =   315
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   1800
            Width           =   3315
         End
         Begin MSMask.MaskEdBox txtFechaTasInm 
            Height          =   330
            Left            =   1905
            TabIndex        =   112
            Top             =   1785
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label LblInmobCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1905
            TabIndex        =   123
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor/Inmobiliaria :"
            Height          =   195
            Left            =   150
            TabIndex        =   122
            Top             =   315
            Width           =   1635
         End
         Begin VB.Label LblInmobNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3270
            TabIndex        =   121
            Top             =   300
            Width           =   4500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Telefono/Inmobiliaria :"
            Height          =   195
            Left            =   150
            TabIndex        =   120
            Top             =   705
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Inmueble :"
            Height          =   195
            Left            =   3390
            TabIndex        =   119
            Top             =   705
            Width           =   1320
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Cuota Inicial :"
            Height          =   195
            Left            =   135
            TabIndex        =   118
            Top             =   1095
            Width           =   960
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Monto Hipoteca :"
            Height          =   195
            Left            =   3390
            TabIndex        =   117
            Top             =   1110
            Width           =   1230
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Precio Venta :"
            Height          =   195
            Left            =   165
            TabIndex        =   116
            Top             =   1455
            Width           =   1005
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Valor Construccion :"
            Height          =   195
            Left            =   3390
            TabIndex        =   115
            Top             =   1485
            Width           =   1425
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Tasación    :"
            Height          =   195
            Left            =   180
            TabIndex        =   114
            Top             =   1860
            Width           =   1605
         End
         Begin VB.Label Label39 
            Caption         =   "Estado Tasación:"
            Height          =   315
            Left            =   3360
            TabIndex        =   113
            Top             =   1800
            Width           =   1575
         End
      End
      Begin VB.Frame fraDatVehic 
         Caption         =   "Datos del Vehículo"
         Height          =   1695
         Left            =   -74640
         TabIndex        =   96
         Top             =   1020
         Width           =   8475
         Begin VB.TextBox txtPlacaVehic 
            Height          =   315
            Left            =   1965
            TabIndex        =   98
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboEstadoTasVeh 
            Height          =   315
            Left            =   5100
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   945
            Width           =   3135
         End
         Begin MSMask.MaskEdBox txtFecTasVeh 
            Height          =   330
            Left            =   1965
            TabIndex        =   99
            Top             =   937
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label35 
            Caption         =   "Placa                        :"
            Height          =   255
            Left            =   300
            TabIndex        =   102
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Tasación    :"
            Height          =   195
            Left            =   240
            TabIndex        =   101
            Top             =   1005
            Width           =   1605
         End
         Begin VB.Label Label37 
            Caption         =   "Estado Tasación:"
            Height          =   255
            Left            =   3360
            TabIndex        =   100
            Top             =   975
            Width           =   1815
         End
      End
      Begin VB.Frame fraZonaCbo 
         Height          =   1725
         Left            =   -74895
         TabIndex        =   53
         Top             =   825
         Width           =   6060
         Begin VB.Frame frazona 
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   330
            TabIndex        =   54
            Top             =   210
            Width           =   5670
            Begin VB.ComboBox cmbPersUbiGeo 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   285
               Index           =   3
               Left            =   3210
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Urbanización"
               Top             =   450
               Width           =   1995
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   3210
               Style           =   2  'Dropdown List
               TabIndex        =   56
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Distrito"
               Top             =   90
               Width           =   1980
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   570
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Provincia"
               Top             =   450
               Width           =   1935
            End
            Begin VB.ComboBox cmbPersUbiGeo 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   570
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Tag             =   "cboPrincipal"
               ToolTipText     =   "Zona"
               Top             =   90
               Width           =   1920
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Zona :"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   165
               Left            =   2625
               TabIndex        =   67
               Top             =   510
               Width           =   390
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Distrito :"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   -15
               TabIndex        =   66
               Top             =   510
               Width           =   525
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Prov :"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   2625
               TabIndex        =   65
               Top             =   150
               Width           =   375
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Dpto :"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   15
               TabIndex        =   64
               Top             =   150
               Width           =   375
            End
         End
         Begin VB.TextBox txtDireccion 
            Height          =   555
            Left            =   900
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   1080
            Width           =   4995
         End
         Begin VB.Label Label1 
            Caption         =   " Zona :"
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
            Left            =   195
            TabIndex        =   69
            Top             =   15
            Width           =   555
         End
         Begin VB.Line Line2 
            X1              =   60
            X2              =   5940
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label34 
            Caption         =   "Dirección:"
            Height          =   195
            Left            =   60
            TabIndex        =   68
            Top             =   1260
            Width           =   795
         End
      End
      Begin VB.Frame framontos 
         Height          =   1740
         Left            =   -68775
         TabIndex        =   48
         Top             =   810
         Width           =   2655
         Begin VB.TextBox txtMontoRea 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   62
            Tag             =   "txtPrincipal"
            Top             =   810
            Width           =   1185
         End
         Begin VB.TextBox txtMontotas 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   61
            Tag             =   "txtPrincipal"
            Top             =   420
            Width           =   1185
         End
         Begin VB.TextBox txtMontoxGrav 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
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
            Left            =   1290
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   63
            Tag             =   "txtPrincipal"
            Top             =   1215
            Width           =   1185
         End
         Begin VB.Label lbltasa 
            AutoSize        =   -1  'True
            Caption         =   "Tasación     :"
            Height          =   195
            Left            =   135
            TabIndex        =   52
            ToolTipText     =   "Monto Tasación"
            Top             =   465
            Width           =   930
         End
         Begin VB.Label lblrealizacion 
            AutoSize        =   -1  'True
            Caption         =   "Realización  :"
            Height          =   195
            Left            =   135
            TabIndex        =   51
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lblMontoGrav 
            AutoSize        =   -1  'True
            Caption         =   "Disponible    :"
            Height          =   195
            Left            =   135
            TabIndex        =   50
            Top             =   1245
            Width           =   960
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   3
            Left            =   870
            TabIndex        =   49
            ToolTipText     =   "Monto Tasación"
            Top             =   15
            Width           =   630
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   -74895
         TabIndex        =   37
         Top             =   2625
         Width           =   8790
         Begin VB.CheckBox ChkGarReal 
            Caption         =   "Garantia Real"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   44
            Top             =   1560
            Width           =   1620
         End
         Begin VB.Frame FraTipoRea 
            Caption         =   "Tipo de Realizacion"
            Height          =   615
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   4320
            Begin VB.OptionButton OptTR 
               Caption         =   "De Lenta Realizacion"
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   90
               TabIndex        =   43
               Top             =   255
               Value           =   -1  'True
               Width           =   1950
            End
            Begin VB.OptionButton OptTR 
               Caption         =   "De Rapida Realizacion"
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2130
               TabIndex        =   42
               Top             =   270
               Width           =   1980
            End
         End
         Begin VB.TextBox txtcomentarios 
            Height          =   480
            Left            =   900
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   157
            Width           =   7725
         End
         Begin VB.Frame FraClase 
            Caption         =   "Clase de Garantia"
            Height          =   615
            Left            =   4530
            TabIndex        =   38
            Top             =   840
            Visible         =   0   'False
            Width           =   4185
            Begin VB.OptionButton OptCG 
               Caption         =   "Garantia Preferida"
               Height          =   240
               Index           =   1
               Left            =   2025
               TabIndex        =   40
               Top             =   255
               Width           =   1650
            End
            Begin VB.OptionButton OptCG 
               Caption         =   "Garantia No Preferida"
               Height          =   240
               Index           =   0
               Left            =   105
               TabIndex        =   39
               Top             =   255
               Value           =   -1  'True
               Width           =   1905
            End
         End
         Begin VB.ComboBox CboBanco 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   960
            Visible         =   0   'False
            Width           =   4170
         End
         Begin VB.Label Label4 
            Caption         =   "Banco"
            Height          =   255
            Left            =   4860
            TabIndex        =   47
            Top             =   885
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Comentario"
            Height          =   195
            Left            =   60
            TabIndex        =   46
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame FraGar 
         Caption         =   "De la Garantia"
         Enabled         =   0   'False
         Height          =   1725
         Left            =   -74655
         TabIndex        =   24
         Top             =   3015
         Width           =   8490
         Begin VB.CommandButton CmdBuscaTasa 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7860
            TabIndex        =   28
            Top             =   315
            Width           =   390
         End
         Begin VB.CommandButton CmdBuscaNot 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7860
            TabIndex        =   27
            Top             =   690
            Width           =   390
         End
         Begin VB.TextBox TxtRegNro 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   25
            Top             =   1320
            Width           =   1785
         End
         Begin MSMask.MaskEdBox TxtFechareg 
            Height          =   330
            Left            =   1965
            TabIndex        =   26
            Top             =   1305
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tasador                     :"
            Height          =   195
            Left            =   225
            TabIndex        =   36
            Top             =   345
            Width           =   1575
         End
         Begin VB.Label LblTasaPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3330
            TabIndex        =   35
            Top             =   330
            Width           =   4500
         End
         Begin VB.Label LblTasaPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1965
            TabIndex        =   34
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Notaria                       :"
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   720
            Width           =   1590
         End
         Begin VB.Label LblNotaPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3330
            TabIndex        =   32
            Top             =   705
            Width           =   4500
         End
         Begin VB.Label LblNotaPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1965
            TabIndex        =   31
            Top             =   705
            Width           =   1350
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Registro    :"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   1350
            Width           =   1530
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Registro :"
            Height          =   195
            Left            =   3345
            TabIndex        =   29
            Top             =   1350
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Detalle de Declaración Jurada"
         Height          =   3840
         Left            =   -74880
         TabIndex        =   16
         Top             =   900
         Width           =   8700
         Begin VB.CommandButton CmdDJAceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6480
            TabIndex        =   20
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton CmdDJCancelar 
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
            Height          =   300
            Left            =   7560
            TabIndex        =   19
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   3360
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton CmdDJNuevo 
            Caption         =   "&Nuevo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6480
            TabIndex        =   18
            Top             =   3360
            Width           =   1005
         End
         Begin VB.CommandButton CmdDJEliminar 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7560
            TabIndex        =   17
            ToolTipText     =   "Salir(ALT+S)"
            Top             =   3360
            Width           =   1005
         End
         Begin SICMACT.FlexEdit FEDeclaracionJur 
            Height          =   2955
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   8460
            _extentx        =   14923
            _extenty        =   5212
            cols0           =   7
            highlight       =   1
            allowuserresizing=   3
            encabezadosnombres=   "-Descripción-Cantidad-Valor Actual-Tipo Doc.-Nro Doc.-Aux"
            encabezadosanchos=   "400-5000-1450-1450-1450-1450-0"
            font            =   "frmPersGarantiasHC.frx":03F0
            font            =   "frmPersGarantiasHC.frx":041C
            font            =   "frmPersGarantiasHC.frx":0448
            font            =   "frmPersGarantiasHC.frx":0474
            font            =   "frmPersGarantiasHC.frx":04A0
            fontfixed       =   "frmPersGarantiasHC.frx":04CC
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-1-2-3-4-5-X"
            listacontroles  =   "0-0-0-0-3-0-0"
            encabezadosalineacion=   "C-L-R-R-L-L-C"
            formatosedit    =   "0-0-3-2-0-0-0"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            colwidth0       =   405
            rowheight0      =   300
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Total Declaración Jurada:"
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
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   3360
            Width           =   2220
         End
         Begin VB.Label LblTotDJ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   2520
            TabIndex        =   22
            Top             =   3360
            Width           =   1440
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3600
         Left            =   -74805
         TabIndex        =   1
         Top             =   750
         Width           =   8565
         Begin VB.CommandButton CmdBuscaSeg 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7725
            TabIndex        =   4
            Top             =   375
            Width           =   390
         End
         Begin VB.TextBox TxtNroPoliza 
            Height          =   285
            Left            =   1815
            TabIndex        =   3
            Top             =   810
            Width           =   1395
         End
         Begin VB.TextBox TxtMontoPol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1815
            TabIndex        =   2
            Top             =   1215
            Width           =   1395
         End
         Begin MSMask.MaskEdBox TxtFecVig 
            Height          =   330
            Left            =   5115
            TabIndex        =   5
            Top             =   825
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtFecCons 
            Height          =   330
            Left            =   1815
            TabIndex        =   6
            Top             =   1635
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtFecTas 
            Height          =   330
            Left            =   1815
            TabIndex        =   7
            Top             =   2085
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Aseguradora         :"
            Height          =   195
            Left            =   285
            TabIndex        =   15
            Top             =   405
            Width           =   1350
         End
         Begin VB.Label LblSegPersNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   3195
            TabIndex        =   14
            Top             =   390
            Width           =   4500
         End
         Begin VB.Label LblSegPersCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1830
            TabIndex        =   13
            Top             =   390
            Width           =   1350
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Nro Poliza             :"
            Height          =   195
            Left            =   285
            TabIndex        =   12
            Top             =   825
            Width           =   1350
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vigencia       :"
            Height          =   195
            Left            =   3480
            TabIndex        =   11
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Monto Poliza        :"
            Height          =   195
            Left            =   285
            TabIndex        =   10
            Top             =   1230
            Width           =   1320
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "F. Constitucion     :"
            Height          =   195
            Left            =   285
            TabIndex        =   9
            Top             =   1665
            Width           =   1320
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "F. Tasacion          :"
            Height          =   195
            Left            =   285
            TabIndex        =   8
            Top             =   2115
            Width           =   1335
         End
      End
      Begin VB.Frame fraPrinc 
         Height          =   2175
         Left            =   90
         TabIndex        =   70
         ToolTipText     =   "Datos del Cliente"
         Top             =   750
         Width           =   8775
         Begin VB.CommandButton CmdBuscaEmisor 
            Caption         =   "..."
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
            Height          =   300
            Left            =   7425
            TabIndex        =   77
            Top             =   255
            Width           =   390
         End
         Begin VB.TextBox txtNumDoc 
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
            Left            =   5730
            MaxLength       =   18
            TabIndex        =   76
            Tag             =   "txtPrincipal"
            Top             =   930
            Width           =   2970
         End
         Begin VB.ComboBox CmbDocGarant 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1515
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   945
            Width           =   3285
         End
         Begin VB.ComboBox CmbTipoGarant 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   570
            Width           =   3225
         End
         Begin VB.TextBox txtDescGarant 
            Height          =   330
            Left            =   1515
            MaxLength       =   60
            TabIndex        =   73
            Tag             =   "txtPrincipal"
            Top             =   1335
            Width           =   5715
         End
         Begin VB.ComboBox CboGarantia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   570
            Width           =   2490
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   1485
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo Moneda"
            Top             =   1740
            Width           =   2115
         End
         Begin VB.CheckBox ChkCF 
            Caption         =   "Carta Fianza"
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
            Height          =   375
            Left            =   5190
            TabIndex        =   78
            Top             =   1320
            Width           =   1425
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Emisor :"
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
            TabIndex        =   89
            Top             =   300
            Width           =   690
         End
         Begin VB.Label LblPersCodEmi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1515
            TabIndex        =   88
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label LblEmisor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   2895
            TabIndex        =   87
            Top             =   270
            Width           =   4500
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. :"
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
            Left            =   4980
            TabIndex        =   86
            Top             =   1005
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Garantía"
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
            TabIndex        =   85
            Top             =   1005
            Width           =   1230
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub Garantía"
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
            Left            =   4110
            TabIndex        =   84
            Top             =   630
            Width           =   1155
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   165
            TabIndex        =   83
            Top             =   1395
            Width           =   885
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Estado          :"
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
            Left            =   3690
            TabIndex        =   82
            Top             =   1800
            Width           =   1260
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   5070
            TabIndex        =   81
            Top             =   1740
            Width           =   2205
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Garantía"
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
            TabIndex        =   80
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   750
            TabIndex        =   79
            Top             =   1800
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmPersGarantiasHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmPersGarantias
'***     Descripcion:       Realiza el Mantenimiento y Registro de Nuevas Garantias
'***     Creado por:        NSSE
'***     Maquina:           07SIST_08
'***     Fecha-Tiempo:         08/06/2001 12:15:13 PM
'***     Ultima Modificacion: Creacion del Formulario
'*****************************************************************************************

Option Explicit
Private Enum TGarantiaTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum

Enum TGarantiaTipoInicio
    RegistroGarantia = 1
    MantenimientoGarantia = 2
    ConsultaGarant = 3
End Enum

Public pgcCtaCod As String

Dim Nivel1() As String
Dim ContNiv1 As Long
Dim Nivel2() As String
Dim ContNiv2 As Long
Dim Nivel3() As String
Dim ContNiv3 As Long
Dim Nivel4() As String
Dim ContNiv4 As Long
Dim bEstadoCargando As Boolean
Dim cmdEjecutar As Integer

Dim vTipoInicio As TGarantiaTipoInicio
Dim sNumgarant As String
Dim bCarga As Boolean
Dim bAsignadoACredito As Boolean

'Agregado por LMMD
Dim bCreditoCF As Boolean
Dim bValdiCCF As Boolean

Public Sub Inicio(ByVal pvTipoIni As TGarantiaTipoInicio)
    
    vTipoInicio = pvTipoIni
    If vTipoInicio = ConsultaGarant Then
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
        CmbDocGarant.Enabled = False
        txtNumDoc.Enabled = False
    End If
    
    If vTipoInicio = RegistroGarantia Then
        CmdNuevo.Enabled = True
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
    End If
    
    If vTipoInicio = MantenimientoGarantia Then
        CmdNuevo.Enabled = False
        ' AQUI Napo deberia inabilitar
        CmdEditar.Enabled = True
        CmdEliminar.Enabled = True
    End If
    
    Me.Show 1
End Sub

Private Function ValidaBuscar() As Boolean
    ValidaBuscar = True
    
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un tipo de Documento de Garantia", vbInformation, "Aviso"
        ValidaBuscar = False
        Exit Function
    End If
    
    If Trim(txtNumDoc.Text) = "" Then
        MsgBox "Ingrese el Numero de Documento", vbInformation, "Aviso"
        ValidaBuscar = False
        Exit Function
    End If
    
End Function

Private Function ValidaDatos() As Boolean
Dim i As Long
Dim Enc As Boolean
Dim oGarantia As COMNCredito.NCOMGarantia
Dim nValor As Double
Dim sCad As String
'Dim odGarantia As COMDCredito.DCOMGarantia
Dim bValidaCta As Boolean
Dim nValorPorcen As Double
Dim nValorCuota As Double

    ValidaDatos = True
    
    Set oGarantia = New COMNCredito.NCOMGarantia
    Call oGarantia.ValidaDatosGarantia(txtNumDoc.Text, bValidaCta, Trim(Right(CmbTipoGarant.Text, 10)), nValorCuota, nValorPorcen)
    Set oGarantia = Nothing
    'Valida cuando es una garantia de Plazo Fijo o CTS
    If Trim(Right(CmbTipoGarant, 3)) = "6" Then
        'Set odGarantia = New COMDCredito.DCOMGarantia
        'If odGarantia.ValidaPFCTS(txtNumDoc) = False Then
    
    'Se quito la Validacion
    '    If bValidaCta = False Then
    '        MsgBox "La cuenta de Plazo Fijo o Cts no es valida", vbInformation, "AVISO"
    '        ValidaDatos = False
    '        Exit Function
    '    End If
    '************************************
    End If
    If SSGarant.TabVisible(3) = True Then
        If FEDeclaracionJur.Rows = 1 Then
           MsgBox "Debe digitar la declaracion jurada", vbInformation, "AVISO"
           ValidaDatos = False
           Exit Function
        End If
    End If
    
    If SSGarant.TabVisible(3) = True Then
        If CDbl(LblTotDJ.Caption) <> CDbl(txtMontoRea) Then
            MsgBox "El monto de la realizacion no coincide con el monto de la declaracion jurada", vbInformation, "AVISO"
            ValidaDatos = False
            Exit Function
        End If
    End If
    
    'Verifica seleccion de Documento de Garantia
    If CmbDocGarant.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Documento de Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    ' verifica seleccion de super tipo de garantia
    If CboGarantia.ListIndex = -1 Then
        MsgBox "Seleccione tipo Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Ingreso de Numero de Documento de Garantia
    If Trim(txtNumDoc.Text) = "" Then
        MsgBox "Ingrese el Numero de Documento de la Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica seleccion de Tipo de Garantia
    If CmbTipoGarant.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica seleccion de Moneda
    If cmbMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la Moneda", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica la Zona
    If cmbPersUbiGeo(3).ListIndex = -1 Then
        MsgBox "Seleccione La Zona donde se Ubica la Garantia", vbInformation, "Aviso"
        SSGarant.Tab = 1
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Monto de Tasacion
    If Trim(txtMontotas.Text) = "" Or Trim(txtMontotas.Text) = "0.00" Then
        MsgBox "El Monto de Tasacion debe ser Mayor que Cero", vbInformation, "Aviso"
        SSGarant.Tab = 1
        txtMontotas.SetFocus
        ValidaDatos = False
        Exit Function
    End If

    'Verifica Monto de Realizacion
    If Trim(txtMontoRea.Text) = "" Or Trim(txtMontoRea.Text) = "0.00" Then
        MsgBox "El Monto de Realizacion debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    'Verifica Monto de Realizacion
    If Trim(txtMontoxGrav.Text) = "" Or Trim(txtMontoxGrav.Text) = "0.00" Then
        MsgBox "El Monto de Disponible debe ser Mayor que Cero", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
   Enc = False
   ' Verifica Existencia de Titular de la Garantia
   For i = 1 To FERelPers.Rows - 1
        If Trim(Right(FERelPers.TextMatrix(i, 3), 15)) <> "" Then
            If CInt(Trim(Right(FERelPers.TextMatrix(i, 3), 15))) = gPersRelGarantiaTitular Then
                Enc = True
                Exit For
            End If
        End If
   Next i
   If Not Enc Then
        MsgBox "Ingrese un Titular para la Garantia", vbInformation, "Aviso"
        ValidaDatos = False
        CmdCliNuevo.SetFocus
        Exit Function
   End If
   
    If Trim(Right(CmbDocGarant, 4)) = "15" Then
        If FEDeclaracionJur.Rows = 2 And FEDeclaracionJur.TextMatrix(1, 1) = "" Then
            MsgBox "Falta digitar la declaracion jurada", vbInformation, "AVISO"
            ValidaDatos = False
            Exit Function
        End If
    End If
   
    ' CMACICA_CSTS - 25/11/2003 -------------------------------------------------
    If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
        If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then
            If Trim(FEDeclaracionJur.TextMatrix(i, 1)) = "" Then
               MsgBox "Ingrese el Detalle de la Declaración Jurada", vbInformation, "Aviso"
               ValidaDatos = False
               SSGarant.Tab = 3
               CmdDJNuevo.SetFocus
               Exit Function
            End If
        End If
    End If
    '----------------------------------------------------------------------------

    ' CMACICA_CSTS - 05/12/2003 ----------------------------------------------------------------------------
    If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
        If CDbl(txtMontoRea.Text) > CDbl(LblTotDJ.Caption) Then
            MsgBox "El Monto de Realización no puede ser mayor al Total de la Declaración Jurada. ", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    End If
    ' ------------------------------------------------------------------------------------------------------
    
   'Validacion de Bancos
   'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Or CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaDepositosGarantia Then
   If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Then
        If CboBanco.ListIndex = -1 Then
            MsgBox "Seleccione un Banco", vbInformation, "Aviso"
            SSGarant.Tab = 1
            CboBanco.SetFocus
            ValidaDatos = False
            Exit Function
        End If
   End If
   
   'Valida Garantias Reales
   If ChkGarReal.value = 1 Then
        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
            If Trim(LblInmobCod.Caption) = "" Then
                MsgBox "Ingrese la Inmobiliaria o el Vendedor", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaInmob.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If CboTipoInmueb.ListIndex = -1 Then
                MsgBox "Seleccione el Tipo del Inmueble", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CboTipoInmueb.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(LblTasaPersCod.Caption) = "" Then
                MsgBox "Ingrese al Tasador", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaTasa.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(LblNotaPersCod.Caption) = "" Then
                MsgBox "Ingrese la Notaria", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaNot.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(LblSegPersCod.Caption) = "" Then
                MsgBox "Ingrese la Empresa Aseguradora", vbInformation, "Aviso"
                SSGarant.Tab = 2
                CmdBuscaSeg.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(TxtNroPoliza.Text) = "" Then
                MsgBox "Ingrese el numero de poliza", vbInformation, "Aviso"
                SSGarant.Tab = 4
                TxtNroPoliza.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(TxtFecVig.Text) = "__/__/____" Then
                MsgBox "Ingrese la fecha de Vigencia de la poliza", vbInformation, "Aviso"
                SSGarant.Tab = 4
                TxtFecVig.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(TxtMontoPol) = "" Then
                TxtMontoPol.Text = "0.00"
            End If
            
            If Trim(TxtMontoPol) = "0.00" Then
                MsgBox "Ingrese el Monto de la poliza", vbInformation, "Aviso"
                SSGarant.Tab = 4
                TxtMontoPol.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(TxtFecCons.Text) = "__/__/____" Then
                MsgBox "Ingrese la fecha de Constitucion de la poliza", vbInformation, "Aviso"
                SSGarant.Tab = 4
                TxtFecCons.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
            If Trim(TxtFecTas.Text) = "__/__/____" Then
                MsgBox "Ingrese la fecha de Tasacion de la poliza", vbInformation, "Aviso"
                SSGarant.Tab = 4
                TxtFecTas.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            
        End If
                        
        sCad = ValidaFecha(TxtFechareg.Text)
        If Trim(sCad) <> "" Then
            MsgBox sCad, vbInformation, "Aviso"
            TxtFechareg.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(TxtRegNro.Text) = "" Then
            MsgBox "Ingrese el Numero de Registro", vbInformation, "Aviso"
            TxtRegNro.SetFocus
            ValidaDatos = False
            Exit Function
        End If
   End If
     
   'Valida Que Monto Disponible No sea Mayor del 90%
   If ChkGarReal.value = 1 Then
        If Trim(Right(CmbTipoGarant.Text, 5)) <> "" Then
            If Trim(TxtPrecioVenta.Text) <> "" Then
                If CInt(Trim(Right(CmbTipoGarant.Text, 5))) = gPersGarantiaHipotecas Then
                    'Set oGarantia = New COMNCredito.NCOMGarantia
                    'nValor = oGarantia.PorcentajeGarantia(gPersGarantia & Trim(Right(CmbTipoGarant.Text, 10)))
                    nValor = nValorPorcen
                    'Set oGarantia = Nothing
                    
                    If CDbl(txtMontoxGrav.Text) > CDbl(Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")) Then
                        MsgBox "Monto Disponible No Puede Exeder al " & Format(nValor * 100, "#0.00") & "% del Precio de Venta", vbInformation, "Aviso"
                        txtMontoxGrav.Text = Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")
                        txtMontoxGrav.SetFocus
                        SSGarant.Tab = 1
                        ValidaDatos = False
                        Exit Function
                    End If
                    
                    'Valida Que la Cuota Inicial No se Menor al 10% del Precio de Venta
                    'Set oGarantia = New COMNCredito.NCOMGarantia
                    'nValor = oGarantia.PorcentajeGarantia("3052")
                    nValor = nValorCuota
                    'Set oGarantia = Nothing
                    
                    If CDbl(TxtHipCuotaIni.Text) < CDbl(Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")) Then
                        MsgBox "Monto de Cuota Inicial No Puede Ser Menor que  el " & Format(nValor * 100, "#0.00") & "% del Precio de Venta", vbInformation, "Aviso"
                        TxtHipCuotaIni.Text = Format(nValor * CDbl(TxtPrecioVenta.Text), "#0.00")
                        TxtHipCuotaIni.SetFocus
                        SSGarant.Tab = 2
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    ' valida que se haya digitado algun declaracion jurada
    
End Function
Private Sub HabilitaIngresoGarantReal(ByVal pbHabilita As Boolean)
    CmdBuscaInmob.Enabled = pbHabilita
    TxtTelefono.Enabled = pbHabilita
    CboTipoInmueb.Enabled = pbHabilita
    CmdBuscaTasa.Enabled = pbHabilita
    CmdBuscaNot.Enabled = pbHabilita
    CmdBuscaSeg.Enabled = pbHabilita
    TxtFechareg.Enabled = pbHabilita
    TxtRegNro.Enabled = pbHabilita
    FraDatInm.Enabled = pbHabilita
    fraDatVehic.Enabled = pbHabilita
    FraGar.Enabled = pbHabilita
End Sub
Private Sub HabilitaIngreso(ByVal pbHabilita As Boolean)
        FraClase.Enabled = pbHabilita
        FraTipoRea.Enabled = pbHabilita
        CboBanco.Enabled = pbHabilita
        ChkGarReal.Enabled = pbHabilita

        CmbDocGarant.Enabled = pbHabilita
        txtNumDoc.Enabled = pbHabilita
        CmdBuscar.Enabled = pbHabilita
        CmbTipoGarant.Enabled = pbHabilita
        CboGarantia.Enabled = pbHabilita
        cmbMoneda.Enabled = pbHabilita
        txtDescGarant.Enabled = pbHabilita
        cmbPersUbiGeo(0).Enabled = pbHabilita
        cmbPersUbiGeo(1).Enabled = pbHabilita
        cmbPersUbiGeo(2).Enabled = pbHabilita
        cmbPersUbiGeo(3).Enabled = pbHabilita
        txtMontotas.Enabled = pbHabilita
        txtMontoRea.Enabled = pbHabilita
        txtMontoxGrav.Enabled = pbHabilita
        txtcomentarios.Enabled = pbHabilita
        txtDireccion.Enabled = pbHabilita 'CUSCO
        FERelPers.lbEditarFlex = False
        '------------------------------------------
        FEDeclaracionJur.lbEditarFlex = False
        '------------------------------------------
        CmdCliNuevo.Enabled = pbHabilita
        CmdCliEliminar.Enabled = pbHabilita
        
        '---------------------------------
        CmdDJNuevo.Enabled = pbHabilita
        CmdDJEliminar.Enabled = pbHabilita
        '---------------------------------
        
        CmdNuevo.Enabled = Not pbHabilita
        CmdNuevo.Visible = Not pbHabilita
        CmdAceptar.Enabled = pbHabilita
        CmdAceptar.Visible = pbHabilita
        CmdEditar.Enabled = Not pbHabilita
        CmdEditar.Visible = Not pbHabilita
        CmdCancelar.Enabled = pbHabilita
        CmdCancelar.Visible = pbHabilita
        CmdEliminar.Enabled = Not pbHabilita
        CmdEliminar.Visible = Not pbHabilita
        cmdSalir.Enabled = Not pbHabilita
        CmdLimpiar.Enabled = Not pbHabilita
        CmdBuscar.Enabled = Not pbHabilita
        FraBuscaPers.Enabled = Not pbHabilita
        If vTipoInicio = MantenimientoGarantia Then
            CmdNuevo.Enabled = False
        End If
        framontos.Enabled = pbHabilita
        CmdBuscaEmisor.Enabled = pbHabilita
End Sub

Private Sub CargaUbicacionesGeograficas(ByVal prsUbic As ADODB.Recordset)
Dim i As Long
Dim nPos As Integer

    'Carga Niveles
    ContNiv1 = 0
    ContNiv2 = 0
    ContNiv3 = 0
    
    ContNiv4 = 0
    
'    Do While Not prsUbic.EOF
'        Select Case prsUbic!P
'            Case 1 ' Departamento
'                ContNiv1 = ContNiv1 + 1
'                ReDim Preserve Nivel1(ContNiv1)
'                Nivel1(ContNiv1 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'            Case 2 ' Provincia
'                ContNiv2 = ContNiv2 + 1
'                ReDim Preserve Nivel2(ContNiv2)
'                Nivel2(ContNiv2 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'            Case 3 'Distrito
'                ContNiv3 = ContNiv3 + 1
'                ReDim Preserve Nivel3(ContNiv3)
'                Nivel3(ContNiv3 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'            Case 4 'Zona
'                ContNiv4 = ContNiv4 + 1
'                ReDim Preserve Nivel4(ContNiv4)
'                Nivel4(ContNiv4 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
'        End Select
'        prsUbic.MoveNext
'    Loop
    
    If prsUbic.EOF Then Exit Sub
    
    Do While prsUbic!P = 1
        ContNiv1 = ContNiv1 + 1
        ReDim Preserve Nivel1(ContNiv1)
        Nivel1(ContNiv1 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
    Loop
        
    Do While prsUbic!P = 2
        ContNiv2 = ContNiv2 + 1
        ReDim Preserve Nivel2(ContNiv2)
        Nivel2(ContNiv2 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
    Loop
    
    Do While prsUbic!P = 3
        ContNiv3 = ContNiv3 + 1
        ReDim Preserve Nivel3(ContNiv3)
        Nivel3(ContNiv3 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
    Loop
    
    Do While prsUbic!P = 4
        ContNiv4 = ContNiv4 + 1
        ReDim Preserve Nivel4(ContNiv4)
        Nivel4(ContNiv4 - 1) = Trim(prsUbic!cUbiGeoDescripcion) & Space(50) & Trim(prsUbic!cUbiGeoCod)
        prsUbic.MoveNext
        If prsUbic.EOF Then Exit Do
    Loop
            
    'Carga el Nivel1 en el Control
    cmbPersUbiGeo(0).Clear
    For i = 0 To ContNiv1 - 1
        cmbPersUbiGeo(0).AddItem Nivel1(i)
        If Trim(Right(Nivel1(i), 12)) = "113000000000" Then
            nPos = i
        End If
    Next i
    cmbPersUbiGeo(0).ListIndex = nPos
    cmbPersUbiGeo(2).Clear
    cmbPersUbiGeo(3).Clear
    
End Sub

Private Sub LimpiaGarantiaReal()
    
    LblInmobCod.Caption = ""
    LblInmobNombre.Caption = ""
    TxtTelefono.Text = ""
    CboTipoInmueb.ListIndex = -1
    LblTasaPersCod.Caption = ""
    LblTasaPersNombre.Caption = ""
    LblNotaPersCod.Caption = ""
    LblNotaPersNombre.Caption = ""
    LblSegPersCod.Caption = ""
    LblSegPersNombre.Caption = ""
    TxtFechareg.Text = "__/__/____"
    TxtRegNro.Text = ""
    TxtHipCuotaIni.Text = "0.00"
    TxtMontoHip.Text = "0.00"
    TxtPrecioVenta.Text = "0.00"
    TxtValorCConst.Text = "0.00"
    TxtNroPoliza.Text = ""
    TxtFecVig.Text = "__/__/____"
    TxtMontoPol.Text = ""
    TxtFecCons.Text = "__/__/____"
    TxtFecTas.Text = "__/__/____"
    
End Sub

Private Sub LimpiaPantalla()
    bCarga = True
    LblEmisor.Tag = LblEmisor.Caption
    LblPersCodEmi.Tag = LblPersCodEmi.Caption
    Call LimpiaControles(Me)
    LblEmisor.Caption = LblEmisor.Tag
    LblPersCodEmi.Caption = LblPersCodEmi.Tag
    Call LimpiaFlex(FERelPers)
    '------------------------------------------
    Call LimpiaFlex(FEDeclaracionJur)
    '------------------------------------------
    Call InicializaCombos(Me)
    txtMontotas.BackColor = vbWhite
    txtMontotas.Text = "0.00"
    txtMontoRea.Text = "0.00"
    txtMontoRea.BackColor = vbWhite
    txtMontoxGrav.Text = "0.00"
    txtMontoxGrav.BackColor = vbWhite
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    LblEmisor.Caption = ""
    LblPersCodEmi.Caption = ""
    OptCG(1).value = True
    OptCG(0).value = True
    OptTR(0).value = True
    CboBanco.ListIndex = -1
    ChkGarReal.value = 0
    Call LimpiaGarantiaReal
    bCarga = False
    bAsignadoACredito = False
End Sub

Private Function CargaDatos(ByVal psNumGarant As String) As Boolean
Dim oGarantia As COMDCredito.DCOMGarantia
Dim nTempo As Integer
Dim nLevantada As Boolean

Dim rsGarantia As ADODB.Recordset
Dim rsRelGarantia As ADODB.Recordset
Dim rsGarantReal As ADODB.Recordset
Dim rsGarantDJ As ADODB.Recordset

    On Error GoTo ErrorCargaDatos
    
    Set oGarantia = New COMDCredito.DCOMGarantia
    bAsignadoACredito = False
    Call oGarantia.CargarDatosGarantia(psNumGarant, rsGarantia, rsRelGarantia, _
                                        rsGarantReal, rsGarantDJ, bAsignadoACredito)
    Set oGarantia = Nothing
    
    If rsGarantia!nEstado = 5 Then 'Si es levantada
        nLevantada = True
    Else
        nLevantada = False
    End If
    
    'bAsignadoACredito = oGarantia.PerteneceACredito(psNumGarant)
    'Set oGarantia = Nothing
    If rsGarantia.RecordCount = 0 Then
        CargaDatos = False
        Exit Function
    Else
        CargaDatos = True
    End If
        
    
    lblEstado.Caption = Trim(rsGarantia!cEstado)
    nTempo = IIf(IsNull(rsGarantia!nGarClase), 0, rsGarantia!nGarClase)
    OptCG(nTempo).value = True
    nTempo = IIf(IsNull(rsGarantia!nGarTpoRealiz), 0, rsGarantia!nGarTpoRealiz)
    OptTR(nTempo).value = True
    'PosicionSuperGarantias Trim(Str(R!nTpoGarantia))
    PosicionSuperGarantias rsGarantia!IdSupGarant
    CmbTipoGarant.ListIndex = IndiceListaCombo(CmbTipoGarant, Trim(Str(rsGarantia!nTpoGarantia)))

    cmbMoneda.ListIndex = IndiceListaCombo(cmbMoneda, Trim(Str(rsGarantia!nMoneda)))
    txtDescGarant.Text = IIf(IsNull(rsGarantia!cDescripcion), "", Trim(rsGarantia!cDescripcion))
    
    ChkGarReal.value = IIf(IsNull(rsGarantia!nGarantReal), 0, rsGarantia!nGarantReal)
    Call HabilitaIngresoGarantReal(False)
    CboBanco.ListIndex = IndiceListaCombo(CboBanco, IIf(IsNull(rsGarantia!cBancoPersCod), "", rsGarantia!cBancoPersCod))
    CboBanco.Enabled = False
    
    CmbDocGarant.ListIndex = IndiceListaCombo(CmbDocGarant, rsGarantia!cTpoDoc)
    txtNumDoc.Text = rsGarantia!cNroDoc
    
    'Carga Ubicacion Geografica
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "1" & Mid(rsGarantia!cZona, 2, 2) & String(9, "0"))
    cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "2" & Mid(rsGarantia!cZona, 2, 4) & String(7, "0"))
    cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "3" & Mid(rsGarantia!cZona, 2, 6) & String(5, "0"))
    cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & rsGarantia!cZona)
    
    If rsGarantia!nMoneda = gMonedaExtranjera Then
        txtMontotas.BackColor = RGB(200, 255, 200)
        txtMontoRea.BackColor = RGB(200, 255, 200)
        txtMontoxGrav.BackColor = RGB(200, 255, 200)
    Else
        txtMontotas.BackColor = vbWhite
        txtMontoRea.BackColor = vbWhite
        txtMontoxGrav.BackColor = vbWhite
    End If
    txtMontotas.Text = Format(rsGarantia!nTasacion, "#0.00")
    txtMontoRea.Text = Format(rsGarantia!nRealizacion, "#0.00")
    txtMontoxGrav.Text = Format(rsGarantia!nPorGravar - rsGarantia!nGravament, "#0.00")
    txtcomentarios.Text = Trim(IIf(IsNull(rsGarantia!cComentario), "", rsGarantia!cComentario))
    txtDireccion.Text = rsGarantia!cDireccion 'CUSCO
    'Personas Relacionadas con Garantias
'    Set oGarantia = New COMDCredito.DCOMGarantia
'    Set RRelPers = oGarantia.RecuperaRelacPersonaGarantia(psNumGarant)
'    Set oGarantia = Nothing
    Call LimpiaFlex(FERelPers)
    Do While Not rsRelGarantia.EOF
        FERelPers.AdicionaFila
        FERelPers.TextMatrix(rsRelGarantia.Bookmark, 1) = rsRelGarantia!cPersCod
        FERelPers.TextMatrix(rsRelGarantia.Bookmark, 2) = rsRelGarantia!cPersNombre
        FERelPers.TextMatrix(rsRelGarantia.Bookmark, 3) = rsRelGarantia!cRelacion
        rsRelGarantia.MoveNext
    Loop
    'RRelPers.Close
    'Set RRelPers = Nothing
    
    'R.Close
    'Set R = Nothing
    
    'Carga Garantias Reales
    If ChkGarReal.value = 1 Then
     '   Set oGarantia = New COMDCredito.DCOMGarantia
     '   Set R = oGarantia.RecuperaGarantiaReal(psNumGarant)
     '   Set oGarantia = Nothing
        If Not (rsGarantReal.EOF And rsGarantReal.BOF) Then
            LblInmobCod.Caption = IIf(IsNull(rsGarantReal!cPersCodVend), "", rsGarantReal!cPersCodVend)
            LblInmobNombre.Caption = IIf(IsNull(rsGarantReal!cVendPersNombre), "", rsGarantReal!cVendPersNombre)
            TxtTelefono.Text = IIf(IsNull(rsGarantReal!cPersVendTelef), "", rsGarantReal!cPersVendTelef)
            CboTipoInmueb.ListIndex = IndiceListaCombo(CboTipoInmueb, IIf(IsNull(rsGarantReal!nTipVivienda), 0, rsGarantReal!nTipVivienda))
            LblTasaPersCod.Caption = IIf(IsNull(rsGarantReal!cPersCodTasador), "", rsGarantReal!cPersCodTasador)
            LblTasaPersNombre.Caption = IIf(IsNull(rsGarantReal!cTasaPersNombre), "", rsGarantReal!cTasaPersNombre)
            LblNotaPersCod.Caption = IIf(IsNull(rsGarantReal!cPersNotaria), "", rsGarantReal!cPersNotaria)
            LblNotaPersNombre.Caption = IIf(IsNull(rsGarantReal!cNotaPersNombre), "", rsGarantReal!cNotaPersNombre)
            LblSegPersCod.Caption = IIf(IsNull(rsGarantReal!cPersCodSeguro), "", rsGarantReal!cPersCodSeguro)
            LblSegPersNombre.Caption = IIf(IsNull(rsGarantReal!cSegPersNombre), "", rsGarantReal!cSegPersNombre)
            TxtFechareg.Text = IIf(IsNull(rsGarantReal!dEscritura), "__/__/____", rsGarantReal!dEscritura)
            TxtRegNro.Text = IIf(IsNull(rsGarantReal!cRegistro), "", rsGarantReal!cRegistro)
            TxtHipCuotaIni.Text = Format(IIf(IsNull(rsGarantReal!nCuotaInicial), "0.00", rsGarantReal!nCuotaInicial), "#0.00")
            TxtMontoHip.Text = Format(IIf(IsNull(rsGarantReal!nMontoHipoteca), "0.00", rsGarantReal!nMontoHipoteca), "#0.00")
            TxtPrecioVenta.Text = Format(IIf(IsNull(rsGarantReal!nPrecioVenta), "0.00", rsGarantReal!nPrecioVenta), "#0.00")
            TxtValorCConst.Text = Format(IIf(IsNull(rsGarantReal!nValorConstruccion), "0.00", rsGarantReal!nValorConstruccion), "#0.00")
            TxtNroPoliza.Text = rsGarantReal!nNroPoliza
            TxtFecVig.Text = Format(rsGarantReal!dVigenciaPol, "dd/mm/yyyy")
            TxtMontoPol.Text = Format(rsGarantReal!nMontoPoliza, "#0.00")
            TxtFecCons.Text = Format(rsGarantReal!dConstitucion, "dd/mm/yyyy")
            TxtFecTas.Text = Format(rsGarantReal!dTasacion, "dd/mm/yyyy")
            '* CUSCO **
            Call HabilitarFramesGarantiaReal(CInt(Trim(Right(CmbTipoGarant, 10))))
            If fraDatVehic.Visible Then
                txtPlacaVehic.Text = rsGarantReal!cPlacaAuto
                cboEstadoTasVeh.ListIndex = IndiceListaCombo(cboEstadoTasInm, rsGarantReal!nEstadoTasacion)
                txtFecTasVeh.Text = IIf(rsGarantReal!dTasacionInmueble <> "01/01/1900", Format(rsGarantReal!dTasacionInmueble, "dd/mm/yyyy"), "__/__/____")
            Else
                cboEstadoTasInm.ListIndex = IndiceListaCombo(cboEstadoTasVeh, rsGarantReal!nEstadoTasacion)
                txtFechaTasInm.Text = IIf(rsGarantReal!dTasacionInmueble <> "01/01/1900", Format(rsGarantReal!dTasacionInmueble, "dd/mm/yyyy"), "__/__/____")
            End If
            ' *******
        End If
    End If
    
    
    LblTotDJ.Caption = "0.00"
    
    ' CMACICA_CSTS - 25/11/2003 -------------------------------------------------
    If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then
       ' Carga Detalle de Garantia DECLARACION JURADA
       'Set oGarantia = New COMDCredito.DCOMGarantia
       'Set RGarDetDJ = oGarantia.RecuperaGarantDeclaracionJur(psNumGarant)
       'Set oGarantia = Nothing
       Call LimpiaFlex(FEDeclaracionJur)
       Do While Not rsGarantDJ.EOF
          FEDeclaracionJur.AdicionaFila
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 1) = rsGarantDJ!cGarDjDescripcion
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 2) = rsGarantDJ!nGarDJCantidad
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 3) = rsGarantDJ!nGarDJPrecioUnit
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 4) = rsGarantDJ!cGarDJTpoDocDes
          FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 5) = rsGarantDJ!cGarDJNroDoc
          
          LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) + (FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 2) * FEDeclaracionJur.TextMatrix(rsGarantDJ.Bookmark, 3)), "#0.00"))
          
          rsGarantDJ.MoveNext
       Loop
       'RGarDetDJ.Close
       'Set RGarDetDJ = Nothing
    End If
    ' ---------------------------------------------------------------------------
    
    If bAsignadoACredito Or nLevantada Then
        framontos.Enabled = False
        CmdEliminar.Enabled = False
        CmdEditar.Enabled = True
        If nLevantada Then
            CmdEditar.Enabled = False
        End If
    Else
        framontos.Enabled = True
        CmdEliminar.Enabled = True
        CmdEditar.Enabled = True
    End If
    
    If Trim(Right(CmbDocGarant, 10)) = "15" Then
        SSGarant.TabVisible(3) = True
    End If
    Exit Function
    
ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub CargaBancos(ByVal prsBancos As ADODB.Recordset)
    
    CboBanco.Clear
    Do While Not prsBancos.EOF
        CboBanco.AddItem PstaNombre(prsBancos!cPersNombre) & Space(150) & prsBancos!cPersCod
        prsBancos.MoveNext
    Loop
End Sub

Private Sub CargaControles()
Dim oGarant As COMDCredito.DCOMGarantia
Dim rsBancos As ADODB.Recordset
Dim rsTInmue As ADODB.Recordset
Dim rsUbic As ADODB.Recordset
Dim rsTGaran As ADODB.Recordset
Dim rsMoneda As ADODB.Recordset
Dim rsRelac As ADODB.Recordset
Dim rsTDocum As ADODB.Recordset
Dim rsSuperG As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    
    'Cargar Objetos de los Controles
    Set oGarant = New COMDCredito.DCOMGarantia
    'avmm --- comentado 11-12-2006
    'Call oGarant.CargarObjetosControles(rsBancos, rsTInmue, rsUbic, rsTGaran, rsMoneda, rsRelac, rsTDocum, rsSuperG)
    Set oGarant = Nothing
    
    'Carga Bancos
    Call CargaBancos(rsBancos)
    
    'Cargar Ubicaciones Geograficas
    'Call CargaUbicacionesGeograficas(rsUbic)
    While Not rsUbic.EOF
        cmbPersUbiGeo(0).AddItem Trim(rsUbic!cUbiGeoDescripcion) & Space(50) & Trim(rsUbic!cUbiGeoCod)
        rsUbic.MoveNext
    Wend
    'Carga Tipos de Inmuebles
    Call Llenar_Combo_con_Recordset(rsTInmue, CboTipoInmueb)
    
    'Carga Tipos de Garantia
    Call CambiaTamañoCombo(CmbTipoGarant)
    Call Llenar_Combo_con_Recordset(rsTGaran, CmbTipoGarant)
    
    'Carga Monedas
    Call Llenar_Combo_con_Recordset(rsMoneda, cmbMoneda)
    
    'Carga Relacion de Personas con Garantia
    FERelPers.CargaCombo rsRelac
        
    ' CMACICA_CSTS - 25/11/2003 ----------------------------------------------------------
    'Carga Tipos de Documentos para el detalle de una Declaracion Jurada
    FEDeclaracionJur.CargaCombo rsTDocum


    Call CargarSuperGarantias(rsSuperG)
    
    Call CambiaTamañoCombo(CboGarantia)
    
    '-------------------------------------------------------------------------------------
        Exit Sub

ERRORCargaControles:
        MsgBox Err.Description, vbCritical, "Aviso"

End Sub

'Private Sub CargaControles()
'Dim R As ADODB.Recordset
'Dim oConstante As COMDConstantes.DCOMConstantes
'
'    On Error GoTo ERRORCargaControles
'
'    'Carga Bancos
'    Call CargaBancos
'    'Carga Tipos de Inmuebles
'    Call CargaComboConstante(gGarantTpoInmueb, CboTipoInmueb)
'
'    'Carga Ubicaciones Geograficas
'        Call CargaUbicacionesGeograficas
'    'Carga Tipos de Garantia
'        Call CambiaTamañoCombo(CmbTipoGarant)
'
'        Call CargaComboConstante(gPersGarantia, CmbTipoGarant)
'    'Carga Monedas
'        Call CargaComboConstante(gMoneda, cmbMoneda)
'
'    'Carga Relacion de Personas con Garantia
'        Set oConstante = New COMDConstantes.DCOMConstantes
'        FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelGarantia)
'        Set oConstante = Nothing
'
'
'    ' CMACICA_CSTS - 25/11/2003 ----------------------------------------------------------
'    'Carga Tipos de Documentos para el detalle de una Declaracion Jurada
'        Set oConstante = New COMDConstantes.DCOMConstantes
'        FEDeclaracionJur.CargaCombo oConstante.RecuperaConstantes(gColocPigTipoDocumento)
'        Set oConstante = Nothing
'
'    CargarSuperGarantias
'    Call CambiaTamañoCombo(CboGarantia)
'
'    '-------------------------------------------------------------------------------------
'        Exit Sub
'
'ERRORCargaControles:
'        MsgBox Err.Description, vbCritical, "Aviso"
'
'End Sub
Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TGarantiaTipoCombo)
Dim i As Long
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboProv
            cmbPersUbiGeo(1).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(1).Clear
                For i = 0 To ContNiv2 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel2(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(1).AddItem Nivel2(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel3(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel4(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
    End Select
End Sub

Private Sub CboGarantia_Click()
    If CboGarantia.ListIndex <> -1 Then
        Call ReLoadCmbTipoGarant(CboGarantia.ItemData(CboGarantia.ListIndex))
    End If
    '05-05-2005
    If CboGarantia.ListIndex = -1 Then Exit Sub
    Select Case CboGarantia.ItemData(CboGarantia.ListIndex)
        Case 1, 3
            OptTR(0).value = True
            OptTR(1).value = False
        Case 2
            OptTR(1).value = True
            OptTR(0).value = False
    End Select
    OptTR(0).Enabled = False
    OptTR(1).Enabled = False
    '******************************
End Sub

Private Sub CboGarantia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmbTipoGarant.SetFocus
    End If
End Sub

Private Sub CboTipoInmueb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoHip.SetFocus
    End If
End Sub

Private Sub ChkCF_Click()
    Dim oDCF As COMDCartaFianza.DCOMCartaFianza
    If ChkCF.value = 1 Then
        If MsgBox("Desea relacionar con Credito C.F", vbInformation + vbYesNo, "AVISO") = vbYes Then
            If Not IsLoadForm("Relacion de Credito Con Garantia") Then
                bCreditoCF = True
                FrmCredRelGarant.Caption = "Credito de la Carta Fianza"
                FrmCredRelGarant.Show vbModal
            End If
            
            'valida el credito
            If pgcCtaCod <> "" Then
                Set oDCF = New COMDCartaFianza.DCOMCartaFianza
                    If oDCF.ValidadCreditoCF(pgcCtaCod) = False Then
                        MsgBox "El Credito no corresponde a una" & vbCrLf & " Carta Fianza", vbInformation, "AVISO"
                        bValdiCCF = False
                    Else
                        bValdiCCF = True
                    End If
                Set oDCF = Nothing
            End If
        End If
    End If
End Sub

Private Sub ChkGarReal_Click()
    If ChkGarReal.value = 1 Then
        Call LimpiaGarantiaReal
        SSGarant.TabVisible(2) = True
        SSGarant.TabVisible(4) = True
        SSGarant.Tab = 2
        
        Call HabilitaIngresoGarantReal(True)
        'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
        '    FraDatInm.Enabled = True
        'Else
        '    FraDatInm.Enabled = False
        'End If
        If FraDatInm.Visible Then
            FraDatInm.Enabled = True
        Else
            fraDatVehic.Enabled = True
        End If
    Else
        SSGarant.Tab = 1
        SSGarant.TabVisible(2) = False
        SSGarant.TabVisible(4) = False
        Call HabilitaIngresoGarantReal(False)
    End If
    
End Sub

Private Sub CmbDocGarant_Click()
    If CmbDocGarant.Enabled = True Then
        If Len(CmbDocGarant) > 0 Then
            If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then  'Declaracion Jurada
                    Call LimpiaFlex(FEDeclaracionJur)
                    LblTotDJ.Caption = "0.00"
                    SSGarant.TabVisible(3) = True
                    'SSGarant.Tab = 3
                Else
                    Call LimpiaFlex(FEDeclaracionJur)
                    'SSGarant.Tab = 1
                    SSGarant.TabVisible(3) = False
                End If
        End If
    End If
End Sub

Private Sub CmbDocGarant_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtNumDoc.SetFocus 'cmbMoneda.SetFocus
     End If
End Sub

Private Sub cmbMoneda_Click()
    Call CmbMoneda_KeyPress(13)
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        'If txtNumDoc.Enabled Then
            'txtNumDoc.SetFocus
        'End If
        If CmdCliNuevo.Enabled And CmdCliNuevo.Visible Then
            CmdCliNuevo.SetFocus
        End If
     End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
'        Select Case Index
'            Case 0 'Combo Dpto
'                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboProv)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(2).Clear
'                    cmbPersUbiGeo(3).Clear
'                End If
'            Case 1 'Combo Provincia
'                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboDist)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(3).Clear
'                End If
'            Case 2 'Combo Distrito
'                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboZona)
'        End Select
Dim oUbic As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim i As Integer

If Index = 3 Then Exit Sub

Set oUbic = New COMDPersona.DCOMPersonas

Set rs = oUbic.CargarUbicacionesGeograficas(, Index + 2, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))

For i = Index + 1 To cmbPersUbiGeo.Count - 1
    cmbPersUbiGeo(i).Clear
Next

While Not rs.EOF
    cmbPersUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
    rs.MoveNext
Wend

Set oUbic = Nothing
End Sub


Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
        Select Case Index
            Case 0
                cmbPersUbiGeo(1).SetFocus
            Case 1
                cmbPersUbiGeo(2).SetFocus
            Case 2
                cmbPersUbiGeo(3).SetFocus
            Case 3
                txtDireccion.SetFocus 'txtMontotas.SetFocus
        End Select
     End If
End Sub

Private Sub CmbTipoGarant_Change()
    txtMontotas.Text = "0.00"
    txtMontoRea.Text = "0.00"
    txtMontoxGrav.Text = "0.00"
End Sub

Private Sub CmbTipoGarant_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
Dim nTipoRealizacion As Integer
Dim R As ADODB.Recordset
        
        txtMontotas.Text = "0.00"
        txtMontoRea.Text = "0.00"
        txtMontoxGrav.Text = "0.00"
    
        If CmbTipoGarant.ListIndex = -1 Then
            If Not bCarga Then
                MsgBox "Debe Escoger un Tipo de Garantia", vbInformation, "Aviso"
                Exit Sub
            Else
                Exit Sub
            End If
        End If
                
        
        Set oGarantia = New COMDCredito.DCOMGarantia
        Set R = oGarantia.RecuperaTiposDocumGarantias(CInt(Right(CmbTipoGarant.Text, 2)))
        Call oGarantia.RecuperaDatosTipoGarantias(CInt(Right(CmbTipoGarant.Text, 2)), R, nTipoRealizacion)
        Set oGarantia = Nothing
        
        '05-05-2005
        'Select Case nTipoRealizacion
        '    Case 0
        '        OptTR(0).value = False
        '        OptTR(1).value = False
        '    Case 300
        '        OptTR(0).value = True
        '    Case 400
        '        OptTR(1).value = True
        'End Select
        '*****************************
        
        CmbDocGarant.Clear
        Do While Not R.EOF
            CmbDocGarant.AddItem R!cDocDesc & Space(150) & R!nDocTpo
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Call CambiaTamañoCombo(CmbDocGarant, 300)
        
        'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Or CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaDepositosGarantia Then
        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Then
            CboBanco.Enabled = True
        Else
            CboBanco.Enabled = False
            CboBanco.ListIndex = -1
        End If
                
        'If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
        '    FraDatInm.Enabled = True
        'Else
        '    FraDatInm.Enabled = False
        'End If
        'If FraDatInm.Visible Then
        '    FraDatInm.Enabled = True
        'End If
        'If fraDatVehic.Visible Then
        '    fraDatVehic.Enabled = True
        'End If
        'CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------
        'DECLARACION JURADA
'        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
'            If CInt(Trim(Right(CmbDocGarant, 10))) = 15 Then  'Declaracion Jurada
'                Call LimpiaFlex(FEDeclaracionJur)
'                LblTotDJ.Caption = "0.00"
'                SSGarant.TabVisible(3) = True
'                'SSGarant.Tab = 3
'            Else
'                Call LimpiaFlex(FEDeclaracionJur)
'                'SSGarant.Tab = 1
'                SSGarant.TabVisible(3) = False
'            End If
'        End If
        '--------------------------------------------------------------------------------------------------------
        Call CambiaTamañoCombo(CmbTipoGarant, 300)
        Call HabilitarFramesGarantiaReal(CInt(Trim(Right(CmbTipoGarant, 10))))
End Sub

Private Sub HabilitarFramesGarantiaReal(ByVal pnSubTipoGarantia As Integer)

Dim oGarant As COMDCredito.DCOMGarantia
Dim bEsSubTipoGInmueble As Boolean
Set oGarant = New COMDCredito.DCOMGarantia

bEsSubTipoGInmueble = oGarant.EsSubTipoGarantiaInmueble(pnSubTipoGarantia)

If bEsSubTipoGInmueble Then
    FraDatInm.Visible = True
    fraDatVehic.Visible = False
Else
    FraDatInm.Visible = False
    fraDatVehic.Visible = True
End If

Set oGarant = Nothing
End Sub

Private Sub CmbTipoGarant_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        CmbDocGarant.SetFocus
     End If
End Sub

Private Sub CmdAceptar_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
Dim RelPers() As String
Dim GarDetDJ() As Variant
Dim oMantGarant As COMDCredito.DCOMGarantia
Dim lsNumGarant As String
Dim i As Long
Dim lrs As ADODB.Recordset
'* CUSCO *
Dim nEstadoTasacion As Integer
Dim dFechaTasacion As Date
'*********

    On Error GoTo ErrorCmdAceptar_Click
    If Not ValidaDatos Then
        Exit Sub
    End If
   
    If MsgBox("Se va a Grabar los Datos, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
   
    If FERelPers.Rows = 2 And FERelPers.TextMatrix(1, 1) = "" Then
        ReDim RelPers(0, 0)
    
    Else
        ReDim RelPers(FERelPers.Rows - 1, 4)
        For i = 1 To FERelPers.Rows - 1
            RelPers(i - 1, 0) = FERelPers.TextMatrix(i, 1) 'Codigo de Persona
            RelPers(i - 1, 1) = Trim(Right(CmbDocGarant.Text, 10)) 'Tipo de Doc de Garantia
            RelPers(i - 1, 2) = Trim(txtNumDoc.Text) 'Numero de Documento
            RelPers(i - 1, 3) = Right("00" & Trim(Right(FERelPers.TextMatrix(i, 3), 10)), 2) 'Relacion
        Next i
    End If
    
    ' CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------
    If FEDeclaracionJur.Rows = 2 And FEDeclaracionJur.TextMatrix(1, 1) = "" Then
        ReDim GarDetDJ(0, 0)
    Else
        ReDim GarDetDJ(FEDeclaracionJur.Rows - 1, 6)
        For i = 1 To FEDeclaracionJur.Rows - 1
            GarDetDJ(i - 1, 0) = FEDeclaracionJur.TextMatrix(i, 0) 'Item
            'GarDetDJ(i - 1, 1) = FEDeclaracionJur.TextMatrix(i, 1) 'Descripcion del Item
            GarDetDJ(i - 1, 1) = Replace(FEDeclaracionJur.TextMatrix(i, 1), "'", "''") 'Descripcion del Item
            GarDetDJ(i - 1, 2) = FEDeclaracionJur.TextMatrix(i, 2) 'Cantidad del Item
            GarDetDJ(i - 1, 3) = FEDeclaracionJur.TextMatrix(i, 3) 'Precio Unit. del Item
            GarDetDJ(i - 1, 4) = Trim(Right(FEDeclaracionJur.TextMatrix(i, 4), 4)) 'Tipo de Doc. del Item
            GarDetDJ(i - 1, 5) = FEDeclaracionJur.TextMatrix(i, 5) 'Nro. Doc. del Item
        Next i
    End If
    ' --------------------------------------------------------------------------------------------------------

    '** Nuevos Campos **
    If ChkGarReal.value = 1 Then
        If FraDatInm.Visible = True Then
            nEstadoTasacion = CInt(Trim(Right(cboEstadoTasInm.Text, 5)))
        Else
            nEstadoTasacion = CInt(Trim(Right(cboEstadoTasVeh.Text, 5)))
        End If
        If FraDatInm.Visible = True Then
            dFechaTasacion = CDate(txtFechaTasInm.Text)
        Else
            dFechaTasacion = CDate(txtFecTasVeh.Text)
        End If
    End If
    '*****
    
    If cmdEjecutar = 1 Then
        Set oGarantia = New COMDCredito.DCOMGarantia
            'Call oGarantia.NuevaGarantia(CStr(Trim(Right(CmbDocGarant.Text, 10))), _
                                         CStr(Trim(txtNumDoc.Text)), _
                                         CStr(Trim(Right(CmbTipoGarant.Text, 10))), _
                                         CboGarantia.ItemData(CboGarantia.ListIndex), CStr(Trim(Right(cmbMoneda.Text, 10))), _
                                         CStr(Trim(txtDescGarant.Text)), CStr(Trim(Right(cmbPersUbiGeo(3).Text, 15))), _
                                          CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), _
                                         CDbl(txtMontoxGrav.Text), CStr(Trim(txtcomentarios.Text)), _
                                         RelPers, CDate(gdFecSis), _
                                         CStr(LblPersCodEmi.Caption), CStr(IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13))), _
                                         CInt(ChkGarReal.value), CInt(IIf(OptCG(0).value, 0, 1)), _
                                          CInt(IIf(OptTR(0).value, 0, 1)), CStr(LblSegPersCod.Caption), _
                                         CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), _
                                         CStr(TxtRegNro.Text), _
                                          0, CStr(LblInmobCod.Caption), _
                                          CStr(TxtTelefono.Text), CStr(LblTasaPersCod.Caption), _
                                          CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), _
                                          "", CStr(LblNotaPersCod.Caption), GarDetDJ, CDbl(TxtHipCuotaIni.Text), CDbl(TxtMontoHip.Text), CDbl(TxtPrecioVenta.Text), CDbl(TxtValorCConst.Text), CStr(TxtNroPoliza.Text), IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons.Text), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecCons.Text), sNumgarant, lrs, gdFecSis, txtDireccion.Text, _
                                          txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, gsCodUser, gsCodCMAC, gsCodAge)
            MsgBox "La información NO se gurado, comunique al area de Informática.", vbInformation, "Aviso"
                
        Set oGarantia = Nothing
    Else
        Set oGarantia = New COMDCredito.DCOMGarantia
        Dim bVerificaDescobertura As Boolean
        
            'Call oGarantia.ActualizaGarantia(sNumgarant, Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text), _
                    Trim(Right(CmbTipoGarant.Text, 10)), CboGarantia.ItemData(CboGarantia.ListIndex), Trim(Right(cmbMoneda.Text, 10)), Trim(txtDescGarant.Text), _
                    Trim(Right(cmbPersUbiGeo(3).Text, 15)), CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), _
                    Trim(txtcomentarios.Text), RelPers, gdFecSis, LblPersCodEmi.Caption, IIf(CboBanco.ListIndex = -1, "", Right(CboBanco.Text, 13)), ChkGarReal.value, _
                    IIf(OptCG(0).value, 0, 1), IIf(OptTR(0).value, 0, 1), LblSegPersCod.Caption, CDate(IIf(TxtFechareg.Text = "__/__/____", "01/01/1950", TxtFechareg.Text)), TxtRegNro.Text, _
                    0, LblInmobCod.Caption, TxtTelefono.Text, LblTasaPersCod.Caption, CInt(IIf(Trim(Right(CboTipoInmueb.Text, 5)) = "", "0", Trim(Right(CboTipoInmueb.Text, 5)))), "", LblNotaPersCod.Caption, GarDetDJ, CDbl(TxtHipCuotaIni.Text), CDbl(TxtMontoHip.Text), CDbl(TxtPrecioVenta.Text), CDbl(TxtValorCConst.Text), TxtNroPoliza.Text, IIf(TxtFecVig.Text = "__/__/____", "01/01/1950", TxtFecVig.Text), Val(TxtMontoPol.Text), IIf(TxtFecCons.Text = "__/__/____", "01/01/1950", TxtFecCons), IIf(TxtFecTas.Text = "__/__/____", "01/01/1950", TxtFecTas), lrs, gdFecSis, txtDireccion.Text, _
                    txtPlacaVehic.Text, nEstadoTasacion, dFechaTasacion, gsProyectoActual, gsCodUser, gsCodCMAC, gsCodAge, bVerificaDescobertura)
            
            MsgBox "La información NO se gurado, comunique al area de Informática.", vbInformation, "Aviso"
            
            If bVerificaDescobertura Then
                MsgBox "El monto de Realización no puede ser menor a lo Gravado en Créditos", vbInformation, "Mensaje"
                Exit Sub
            End If
        Set oGarantia = Nothing
    End If
    If SSGarant.TabVisible(3) = True Then
        'Set oMantGarant = New COMDCredito.DCOMGarantia
        
        'If cmdEjecutar = 1 Then
           'cuando es una nueva garantia
           'lsNumGarant = oMantGarant.ObtenerMaxcNumGarant
           'Set oMantGarant = Nothing
           'Set lrs = oMantGarant.DJ(lsNumGarant, gdFecSis)
        'Else
            ' cuando es una actualizacion de la garantia
           'Set lrs = oMantGarant.DJ(sNumgarant, gdFecSis)
        'End If
        
        '07-05-2006
        'With DRDJ
        '    Set .DataSource = lrs
        '    .DataMember = ""
        '    '.Orientation = rptOrientPortrait
        '    .Inicio sNumgarant, gdFecSis
        '    .Refresh
        '    .Show vbModal
        'End With
        'Set oMantGarant = Nothing
        '************************
    End If
    cmdEjecutar = -1
    Call HabilitaIngreso(False)
    Call HabilitaIngresoGarantReal(False)
    
    Exit Sub


ErrorCmdAceptar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdBuscaEmisor_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblPersCodEmi.Caption = oPers.sPersCod
        LblEmisor.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdBuscaInmob_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblInmobCod.Caption = oPers.sPersCod
        LblInmobNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    
End Sub

Private Sub CmdBuscaNot_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblNotaPersCod.Caption = oPers.sPersCod
        LblNotaPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdBuscaPersona_Click()
    Call cmdCancelar_Click
    ObtieneDocumPersona
    If vTipoInicio = ConsultaGarant Then
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
    End If
End Sub

Private Sub ObtieneDocumPersona()
Dim oGaran As COMDCredito.DCOMGarantia
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
Dim L As ListItem
    
    LstGaratias.ListItems.Clear
    Set oPers = New COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    Set oGaran = New COMDCredito.DCOMGarantia
    
    If oPers Is Nothing Then
        Exit Sub
    End If
    Set R = oGaran.RecuperaGarantiasPersona(oPers.sPersCod, True)
    Set oGaran = Nothing
    If R.RecordCount > 0 Then
        Me.Caption = "Garantias de Cliente : " & oPers.sPersNombre
    End If
    LstGaratias.ListItems.Clear
    Set oPers = Nothing
    Do While Not R.EOF
        Set L = LstGaratias.ListItems.Add(, , IIf(IsNull(R!cDescripcion), "", R!cDescripcion))
        L.Bold = True
        If R!nMoneda = gMonedaExtranjera Then
            L.ForeColor = RGB(0, 125, 0)
        Else
            L.ForeColor = vbBlack
        End If
        L.SubItems(1) = Trim(R!cNumGarant)
        L.SubItems(2) = Trim(R!cPersCodEmisor)
        L.SubItems(3) = PstaNombre(R!cPersNombre)
        L.SubItems(4) = Trim(R!cTpoDoc)
        L.SubItems(5) = Trim(R!cNroDoc)
        
        R.MoveNext
    Loop
End Sub

Private Sub cmdBuscar_Click()
    bAsignadoACredito = False
    If Me.LstGaratias.ListItems.Count = 0 Then
        MsgBox "No Existe Garantia que Mostrar ", vbInformation, "Aviso"
        Exit Sub
    End If
    CmbDocGarant.Enabled = False
    txtNumDoc.Enabled = False
    Me.LblPersCodEmi.Caption = Me.LstGaratias.SelectedItem.SubItems(2)
    Me.LblEmisor.Caption = Me.LstGaratias.SelectedItem.SubItems(3)
    'Me.CmbDocGarant.ListIndex = IndiceListaCombo(CmbDocGarant, Me.LstGaratias.SelectedItem.SubItems(4))
    'Me.txtNumDoc.Text = Me.LstGaratias.SelectedItem.SubItems(5)
    
    Call CargaDatos(Trim(LstGaratias.SelectedItem.SubItems(1)))
    sNumgarant = Trim(Me.LstGaratias.SelectedItem.SubItems(1))
    
    If vTipoInicio = ConsultaGarant Then
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
    End If
End Sub

Private Sub CmdBuscaSeg_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblSegPersCod.Caption = oPers.sPersCod
        LblSegPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdBuscaTasa_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblTasaPersCod.Caption = oPers.sPersCod
        LblTasaPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
    If cmdEjecutar = 2 Then
        CargaDatos Trim(sNumgarant)
    Else
        If cmdEjecutar = 1 Then
            Call LimpiaPantalla
        End If
    End If
    Call HabilitaIngreso(False)
    If Me.ChkGarReal.value = 1 Then
        Call HabilitaIngresoGarantReal(False)
    End If
    
    Call LimpiaPantalla
    'CmbDocGarant.Enabled = True
    'txtNumDoc.Enabled = True
    'CmbDocGarant.SetFocus
    cmdEjecutar = -1
End Sub

Private Sub CmdCliAceptar_Click()
Dim i As Long
Dim oGarantia As COMNCredito.NCOMGarantia
Dim RelPers() As String

    For i = 1 To FERelPers.Rows - 2
        If Trim(FERelPers.TextMatrix(i, 1)) = Trim(FERelPers.TextMatrix(FERelPers.Rows - 1, 1)) Then
            MsgBox "Persona Ya Tiene Relacion de la Garantia", vbInformation, "Aviso"
            FERelPers.Row = FERelPers.Rows - 1
            FERelPers.Col = 1
            FERelPers.SetFocus
            Exit Sub
        End If
    Next i
    For i = 1 To FERelPers.Rows - 1
        If Len(Trim(FERelPers.TextMatrix(i, 1))) < 13 Then
            MsgBox "Codigo de Persona Incorrecto", vbInformation, "Aviso"
            FERelPers.Row = i
            FERelPers.Col = 1
            FERelPers.SetFocus
            Exit Sub
        End If
        If Len(Trim(FERelPers.TextMatrix(i, 3))) = 0 Then
            MsgBox "Relacion de Persona Con la Garantias es Incorrecto", vbInformation, "Aviso"
            FERelPers.Row = i
            FERelPers.Col = 3
            FERelPers.SetFocus
            Exit Sub
        End If
    Next i
    ReDim RelPers(FERelPers.Rows - 1)
    For i = 1 To FERelPers.Rows - 1
        RelPers(i - 1) = FERelPers.TextMatrix(i, 3)
    Next i
    Set oGarantia = New COMNCredito.NCOMGarantia
    If oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)) <> "" Then
        MsgBox oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)), vbInformation, "Aviso"
        Exit Sub
    End If
    Set oGarantia = Nothing
    
    FERelPers.lbEditarFlex = False
    CmdCliNuevo.Visible = True
    CmdCliEliminar.Visible = True
    CmdCliAceptar.Visible = False
    CmdCliCancelar.Visible = False
    CmdAceptar.Enabled = True
    CmdCancelar.Enabled = True
End Sub

Private Sub CmdCliCancelar_Click()
    Call FERelPers.EliminaFila(FERelPers.Row)
    FERelPers.lbEditarFlex = False
    CmdCliNuevo.Visible = True
    CmdCliEliminar.Visible = True
    CmdCliAceptar.Visible = False
    CmdCliCancelar.Visible = False
    CmdAceptar.Enabled = True
    CmdCancelar.Enabled = True
End Sub

Private Sub CmdCliEliminar_Click()
    If FERelPers.Row < 1 Then
        Exit Sub
    End If
    If MsgBox("Se va a Eliminar a la Persona " & FERelPers.TextMatrix(FERelPers.Row, 2) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If FERelPers.Row = 1 And FERelPers.Rows = 2 Then
            FERelPers.TextMatrix(1, 0) = ""
            FERelPers.TextMatrix(1, 1) = ""
            FERelPers.TextMatrix(1, 2) = ""
            FERelPers.TextMatrix(1, 3) = ""
        Else
            Call FERelPers.EliminaFila(FERelPers.Row)
        End If
    End If
End Sub

Private Sub CmdCliNuevo_Click()
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    CmdCliNuevo.Visible = False
    CmdCliEliminar.Visible = False
    CmdCliAceptar.Visible = True
    CmdCliCancelar.Visible = True
    CmdAceptar.Enabled = False
    CmdCancelar.Enabled = False
    FERelPers.SetFocus
End Sub

Private Sub CmdDJAceptar_Click()
Dim i As Long
Dim oGarantia As COMNCredito.NCOMGarantia
Dim RelPers() As String

'    For I = 1 To FEDeclaracionJur.Rows - 2
'        If Trim(FEDeclaracionJur.TextMatrix(I, 1)) = Trim(FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Rows - 1, 1)) Then
'            MsgBox "Persona Ya Tiene Relacion de la Garantia", vbInformation, "Aviso"
'            FEDeclaracionJur.Row = FEDeclaracionJur.Rows - 1
'            FEDeclaracionJur.Col = 1
'            FEDeclaracionJur.SetFocus
'            Exit Sub
'        End If
'    Next I
    
    LblTotDJ.Caption = "0.00"
    
    For i = 1 To FEDeclaracionJur.Rows - 1
        If Len(Trim(FEDeclaracionJur.TextMatrix(i, 1))) = 0 Then
            MsgBox "Falta Ingresar la Descripción del Item", vbInformation, "Aviso"
            FEDeclaracionJur.Row = i
            FEDeclaracionJur.Col = 1
            FEDeclaracionJur.SetFocus
            Exit Sub
        End If
        If FEDeclaracionJur.TextMatrix(i, 2) = 0 Then
            MsgBox "Falta Ingresar la Cantidad del item", vbInformation, "Aviso"
            FEDeclaracionJur.Row = i
            FEDeclaracionJur.Col = 2
            FEDeclaracionJur.SetFocus
            Exit Sub
        End If
        If FEDeclaracionJur.TextMatrix(i, 3) = 0 Then
            MsgBox "Falta Ingresar el Valor Actual del item", vbInformation, "Aviso"
            FEDeclaracionJur.Row = i
            FEDeclaracionJur.Col = 3
            FEDeclaracionJur.SetFocus
            Exit Sub
        End If
        
        LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) + (FEDeclaracionJur.TextMatrix(i, 2) * CCur(FEDeclaracionJur.TextMatrix(i, 3))), "#0.00"))
    
    Next i
'    ReDim RelPers(FEDeclaracionJur.Rows - 1)
'    For I = 1 To FEDeclaracionJur.Rows - 1
'        RelPers(I - 1) = FEDeclaracionJur.TextMatrix(I, 3)
'    Next I
'    Set oGarantia = New NGarantia
'    If oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)) <> "" Then
'        MsgBox oGarantia.ValidaDatos(RelPers, CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text)), vbInformation, "Aviso"
'        Exit Sub
'    End If
'    Set oGarantia = Nothing
    
    FEDeclaracionJur.lbEditarFlex = False
    CmdDJNuevo.Visible = True
    CmdDJEliminar.Visible = True
    CmdDJAceptar.Visible = False
    CmdDJCancelar.Visible = False
    CmdAceptar.Enabled = True
    CmdCancelar.Enabled = True
    
    
End Sub

Private Sub CmdDJCancelar_Click()
    Call FEDeclaracionJur.EliminaFila(FEDeclaracionJur.Row)
    FEDeclaracionJur.lbEditarFlex = False
    CmdDJNuevo.Visible = True
    CmdDJEliminar.Visible = True
    CmdDJAceptar.Visible = False
    CmdDJCancelar.Visible = False
    CmdAceptar.Enabled = True
    CmdCancelar.Enabled = True
End Sub

Private Sub CmdDJEliminar_Click()
If FEDeclaracionJur.Row < 1 Then
    Exit Sub
End If
If MsgBox("Se va a Eliminar al Item " & FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Row, 1) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    If FEDeclaracionJur.Row = 1 And FEDeclaracionJur.Rows = 2 Then
       LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) - (FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Row, 2) * FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Row, 3)), "#0.00"))
       
       FEDeclaracionJur.TextMatrix(1, 0) = ""
       FEDeclaracionJur.TextMatrix(1, 1) = ""
       FEDeclaracionJur.TextMatrix(1, 2) = 0
       FEDeclaracionJur.TextMatrix(1, 3) = 0
       FEDeclaracionJur.TextMatrix(1, 4) = ""
       FEDeclaracionJur.TextMatrix(1, 5) = ""
    Else
       LblTotDJ.Caption = CStr(Format(CDbl(LblTotDJ) - (FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Row, 2) * FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Row, 3)), "#0.00"))

       Call FEDeclaracionJur.EliminaFila(FEDeclaracionJur.Row)
    End If
End If
End Sub

Private Sub CmdDJNuevo_Click()
    FEDeclaracionJur.lbEditarFlex = True
    FEDeclaracionJur.AdicionaFila
    CmdDJNuevo.Visible = False
    CmdDJEliminar.Visible = False
    CmdDJAceptar.Visible = True
    CmdDJCancelar.Visible = True
    CmdAceptar.Enabled = False
    CmdCancelar.Enabled = False
    FEDeclaracionJur.SetFocus
End Sub

Private Sub cmdEditar_Click()
'Dim oGarantia As New COMDCredito.DCOMGarantia

    'If oGarantia.GarantiaEnUso(sNumgarant) Then
    '    MsgBox "Solo Puede Editar el Comentario, La Garantia ya esta en uso por un Credito", vbInformation, "Aviso"
    '    Set oGarantia = Nothing
    '    Exit Sub
    'End If
    'Set oGarantia = Nothing
    If LstGaratias.ListItems.Count = 0 Then
        MsgBox "Seleccione una Garantia", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    If bAsignadoACredito Then
        fraPrinc.Enabled = False
        FraRelaGar.Enabled = False
        fraZonaCbo.Enabled = True
        'Para llevar el Historico de los montos
        framontos.Enabled = True
        txtMontoxGrav.Enabled = True
        txtMontotas.Enabled = True
        txtMontoRea.Enabled = True
        '***************************
        'Frame2.Enabled = False
        FraTipoRea.Enabled = False
        txtcomentarios.Enabled = False
        
        
        cmbPersUbiGeo(0).Enabled = False
        cmbPersUbiGeo(1).Enabled = False
        cmbPersUbiGeo(2).Enabled = False
        cmbPersUbiGeo(3).Enabled = False
        txtDireccion.Enabled = False 'Nuevo campo Direccion
        txtcomentarios.Enabled = True
        FraDatInm.Enabled = False
        FraGar.Enabled = False
        
        
        CmdNuevo.Enabled = False
        CmdNuevo.Visible = False
        CmdAceptar.Enabled = True
        CmdAceptar.Visible = True
        CmdEditar.Enabled = False
        CmdEditar.Visible = False
        CmdCancelar.Enabled = True
        CmdCancelar.Visible = True
        CmdEliminar.Enabled = False
        CmdEliminar.Visible = False
        cmdSalir.Enabled = False
        CmdLimpiar.Enabled = False
        CmdBuscar.Enabled = False
        FraBuscaPers.Enabled = False
        '05-05
        ChkGarReal.Enabled = True
        '***********
    Else
        Call HabilitaIngreso(True)
        '05-05
        ChkGarReal.Enabled = True
        'If ChkGarReal.value = 1 Then
        '    Call HabilitaIngresoGarantReal(True)
        'End If
        '***********
        
        'Activa Controles segun tipo de garantia
        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaCartasFianza Or CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaDepositosGarantia Then
            CboBanco.Enabled = True
        Else
            CboBanco.Enabled = False
        End If
        
        If ChkGarReal.value = 1 Then
        '    If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaHipotecas Then
        '        FraDatInm.Enabled = True
        '    Else
        '        FraDatInm.Enabled = False
        '    End If
            If FraDatInm.Visible Then
                FraDatInm.Enabled = True
            Else
                fraDatVehic.Enabled = True
            End If
        End If
        
        'CMACICA_CSTS - 25/11/2003 ------------------------------------------------------------------------------
'        If CInt(Trim(Right(CmbTipoGarant, 10))) = gPersGarantiaOtrasGarantias Then
'            SSGarant.TabVisible(3) = True
'            'FEDeclaracionJur.Enabled = True
'        Else
'            'FEDeclaracionJur.Enabled = False
'            SSGarant.TabVisible(3) = False
'        End If
        '--------------------------------------------------------------------------------------------------------
        
        'CmbDocGarant.Enabled = False
        'txtNumDoc.Enabled = False
        If Trim(Right(CmbDocGarant, 10)) = "15" Then
            SSGarant.TabVisible(3) = True
        Else
            SSGarant.TabVisible(3) = False
        End If
        CmbTipoGarant.SetFocus
        cmdEjecutar = 2
    End If
End Sub

Private Sub cmdeliminar_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
    If MsgBox("Se va a Eliminar la Garantia, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oGarantia = New COMDCredito.DCOMGarantia
        'If oGarantia.GarantiaEnUso(sNumgarant) Then
        '    MsgBox "No se Puede Eliminar la Garantia porque ya esta en uso por un Credito", vbInformation, "Aviso"
        '    Set oGarantia = Nothing
        '    Exit Sub
        'End If
        Call oGarantia.EliminarGraantia(sNumgarant)
        Set oGarantia = Nothing
        Call LimpiaPantalla
        Call CmdLimpiar_Click
    End If
    cmdCancelar_Click
    CmdBuscar.Enabled = True
End Sub

Private Sub CmdLimpiar_Click()
    
    Call LimpiaPantalla
    HabilitaIngreso False
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    'CmbDocGarant.Enabled = True
    'txtNumDoc.Enabled = True
    'cmdBuscar.Enabled = True
    'CmbDocGarant.SetFocus
    
    If vTipoInicio = ConsultaGarant Then
        CmdNuevo.Enabled = False
        CmdEditar.Enabled = False
        CmdEliminar.Enabled = False
    End If
    
    'Agregado por LMMD CF
    bCreditoCF = False
    bValdiCCF = False
End Sub

Private Sub cmdNuevo_Click()
    Call HabilitaIngreso(True)
    Call LimpiaPantalla
    Call InicializaCombos(Me)
    cmdEjecutar = 1
    CmdBuscaEmisor.Enabled = True
    Call CmdBuscaEmisor_Click
    'CmbTipoGarant.SetFocus
    If CboGarantia.Enabled Then
        CboGarantia.SetFocus
    End If
    SSGarant.TabVisible(3) = False 'la ficha de la declaracion jurada
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FEDeclaracionJur_KeyPress(KeyAscii As Integer)
    Dim c As String
If KeyAscii = 13 Then
    If FEDeclaracionJur.Col = 5 Then
        If FEDeclaracionJur.TextMatrix(FEDeclaracionJur.Row, 5) <> "" Then
            CmdDJAceptar.SetFocus
        End If
    End If
End If


If FEDeclaracionJur.Col = 1 Then
    c = Chr(KeyAscii)
    c = UCase(c)
    KeyAscii = Asc(c)
End If
End Sub


Private Sub FEDeclaracionJur_OnCellChange(pnRow As Long, pnCol As Long)
    Dim c As String
    
    If FEDeclaracionJur.Col = 1 Then
        c = FEDeclaracionJur.TextMatrix(pnRow, pnCol)
        c = UCase(c)
        FEDeclaracionJur.TextMatrix(pnRow, pnCol) = c
    End If
End Sub

Private Sub FEDeclaracionJur_RowColChange()
Dim oConstante As COMDConstantes.DCOMConstantes

If FEDeclaracionJur.Col = 4 Then
    'Carga los tipos de documentos del item
    Set oConstante = New COMDConstantes.DCOMConstantes
    FEDeclaracionJur.CargaCombo oConstante.RecuperaConstantes(gColocPigTipoDocumento)
    Set oConstante = Nothing
End If

End Sub

Private Sub FERelPers_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FERelPers.Col = 3 Then
        If FERelPers.TextMatrix(FERelPers.Row, 3) <> "" Then
            CmdCliAceptar.SetFocus
        End If
    End If
End If

End Sub

Private Sub FERelPers_RowColChange()
Dim oConstante As COMDConstantes.DCOMConstantes

If FERelPers.Col = 3 Then
'Carga Relacion de Personas con Garantia
    Set oConstante = New COMDConstantes.DCOMConstantes
    FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelGarantia)
    Set oConstante = Nothing
End If

End Sub

Private Sub Form_Load()
    CentraForm Me
    SSGarant.Tab = 0
    SSGarant.TabVisible(2) = False
    SSGarant.TabVisible(3) = False
    SSGarant.TabVisible(4) = False
    bEstadoCargando = True
    Call CargaControles
    Call HabilitaIngreso(False)
    CmbDocGarant.Enabled = True
    txtNumDoc.Enabled = True
    CmdBuscar.Enabled = True
    bEstadoCargando = False
    cmdEjecutar = -1
    CmdEliminar.Enabled = False
    CmdEditar.Enabled = False
    CboGarantia.Enabled = False
    CmdNuevo.Enabled = True
    bCreditoCF = False
    bValdiCCF = False
End Sub

Private Sub OptCG_Click(Index As Integer)
'05-05
'    If OptCG(0).value = True Then
'        OptTR(0).Enabled = False
'        OptTR(1).Enabled = False
'        OptTR(0).value = True
'    Else
'        OptTR(0).Enabled = True
'        OptTR(1).Enabled = True
'        OptTR(0).value = True
'    End If
'***************
End Sub

Private Sub SSGarant_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If CmdCliAceptar.Visible And CmdCliAceptar.Enabled Then
            MsgBox "Pulse Aceptar para Registrar al Cliente", vbInformation, "Aviso"
            CmdCliAceptar.SetFocus
            SSGarant.Tab = 0
        End If
    End If
   
    If PreviousTab = 3 Then
        If CmdDJAceptar.Visible And CmdDJAceptar.Enabled Then
            MsgBox "Pulse Aceptar para Registrar el Item", vbInformation, "Aviso"
            CmdDJAceptar.SetFocus
            SSGarant.Tab = 3
        End If
    End If
    
    If SSGarant.Tab = 2 And cboEstadoTasInm.ListCount = 0 Then
        Dim oCons As COMDConstantes.DCOMConstantes
        Dim rs As ADODB.Recordset
        Dim rsTmp As ADODB.Recordset
        Set oCons = New COMDConstantes.DCOMConstantes
        Set rs = oCons.RecuperaConstantes(gCredGarantEstadoTasacion)
        Set rsTmp = rs.Clone
        Set oCons = Nothing
        Call Llenar_Combo_con_Recordset(rs, cboEstadoTasInm)
        Call Llenar_Combo_con_Recordset(rsTmp, cboEstadoTasVeh)
    End If
End Sub

Private Sub txtcomentarios_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtDescGarant_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
     If KeyAscii = 13 Then
'        If CmdCliNuevo.Enabled And CmdCliNuevo.Visible Then
'            CmdCliNuevo.SetFocus
'        End If
        cmbMoneda.SetFocus
     End If
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMontotas.SetFocus
End Sub

Private Sub TxtFecCons_GotFocus()
    fEnfoque TxtFecCons
End Sub

Private Sub TxtFecCons_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        TxtFecTas.SetFocus
    End If
End Sub

Private Sub TxtFechareg_GotFocus()
    fEnfoque TxtFechareg
End Sub

Private Sub TxtFechareg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRegNro.SetFocus
    End If
End Sub

Private Sub TxtFechareg_LostFocus()
Dim sCad As String
    If TxtFechareg.Text = "__/__/____" Then Exit Sub
    sCad = ValidaFecha(TxtFechareg.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFechareg.SetFocus
    End If
End Sub

Private Sub TxtFecTas_GotFocus()
    fEnfoque TxtFecTas
End Sub

Private Sub TxtFecVig_GotFocus()
    fEnfoque TxtFecVig
End Sub

Private Sub TxtFecVig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtMontoPol.SetFocus
    End If
End Sub

Private Sub TxtHipCuotaIni_GotFocus()
    fEnfoque TxtHipCuotaIni
End Sub

Private Sub TxtHipCuotaIni_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtHipCuotaIni, KeyAscii)
    If KeyAscii = 13 Then
        TxtMontoHip.SetFocus
    End If
End Sub

Private Sub TxtHipCuotaIni_LostFocus()
    If Trim(TxtHipCuotaIni.Text) = "" Then
        TxtHipCuotaIni.Text = "0.00"
    Else
        TxtHipCuotaIni.Text = Format(TxtHipCuotaIni.Text, "#0.00")
    End If
End Sub

Private Sub TxtMontoHip_GotFocus()
    fEnfoque TxtMontoHip
End Sub

Private Sub TxtMontoHip_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoHip, KeyAscii)
    If KeyAscii = 13 Then
        TxtPrecioVenta.SetFocus
    End If
End Sub

Private Sub TxtMontoHip_LostFocus()
    If Trim(TxtMontoHip.Text) = "" Then
        TxtMontoHip.Text = "0.00"
    Else
        TxtMontoHip.Text = Format(TxtMontoHip.Text, "#0.00")
    End If
End Sub

Private Sub TxtMontoPol_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtMontoPol, KeyAscii)
    If KeyAscii = 13 Then
        TxtFecCons.SetFocus
    End If
End Sub

Private Sub TxtMontoPol_LostFocus()
    If Trim(TxtMontoPol.Text) = "" Then
        TxtMontoPol.Text = "0.00"
    End If
    TxtMontoPol.Text = Format(TxtMontoPol.Text, "#0.00")
End Sub

Private Sub txtMontoRea_Change()
    'txtMontoxGrav.Text = "0.00"
End Sub

Private Sub txtMontoRea_GotFocus()
    fEnfoque txtMontoRea
End Sub

Private Sub txtMontoRea_KeyPress(KeyAscii As Integer)
Dim oGarantia As COMNCredito.NCOMGarantia
Dim oPersona As DPersona
Dim nValor As Double
Dim sCad As String
Dim oMantGarant As COMDCredito.DCOMGarantia
Dim oDCF As COMDCartaFianza.DCOMCartaFianza

     KeyAscii = NumerosDecimales(txtMontoRea, KeyAscii)
     If KeyAscii = 13 Then
        
'06-05-2005
        Set oGarantia = New COMNCredito.NCOMGarantia
'        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)  ', CLng(Trim(Right(CmbTipoGarant.Text, 10))))
'        If Not sCad = "" Then
'            MsgBox sCad, vbInformation, "Aviso"
'            txtMontoRea.SetFocus
'            Exit Sub
'        End If
'********************************
        ' se verifica si el cliente es preferencial para creditos automatico
        ' de ser asi se coge el 100 de la garantia
        
        'Set oPersona = New DPersona
        'If ObtenerTitularX <> "" Then
        '     If oPersona.ValidaPersonaPreferencial(ObtenerTitularX) = True Then
         '        txtMontoxGrav.Text = Format(CDbl(txtMontoRea.Text), "#0.00")
         
         
'         If MsgBox("¿Desea relacionar esta  garantia  con " & vbCrLf & _
'                   "algun credito automatico?", vbInformation + vbYesNo, "Informacion") = vbYes Then
'                    If Not IsLoadForm("Relacion de Credito Con Garantia") Then
'                          FrmCredRelGarant.Show vbModal
'                        If pgcCtaCod <> "" Then
'                           Set oMantGarant = New COMDCredito.DCOMGarantia
'                                If oMantGarant.VerificarCreditoAutomatico(pgcCtaCod) Then
'                                    txtMontoxGrav.Text = Format(CDbl(txtMontoRea.Text), "#0.00")
'                                Else
'                                    MsgBox "El credito no es automatico", vbInformation, "AVISO"
'                                    nValor = oGarantia.PorcentajeGarantia(gPersGarantia & Trim(Right(CmbTipoGarant.Text, 10)))
'                                    Set oGarantia = Nothing
'                                    If nValor > 1 Then
'                                        txtMontoxGrav.Text = Format(nValor, "#0.00")
'                                    Else
'                                        txtMontoxGrav.Text = Format(nValor * CDbl(txtMontoRea.Text), "#0.00")
'                                    End If
'                                End If
'                            Set oMantGarant = Nothing
'                        End If
'                     End If
'             Else
                    If bValdiCCF = True Then
                        Set oDCF = New COMDCartaFianza.DCOMCartaFianza
                        nValor = oDCF.ValorCoberturaGarantia
                        Set oDCF = Nothing
                    Else
                      nValor = oGarantia.PorcentajeGarantia(gPersGarantia & Trim(Right(CmbTipoGarant.Text, 10)))
                    End If
                     Set oGarantia = Nothing
                     If nValor > 1 Then
                         txtMontoxGrav.Text = Format(nValor, "#0.00")
                     Else
                         txtMontoxGrav.Text = Format(nValor * CDbl(txtMontoRea.Text), "#0.00")
                     End If
        '06-05-2005
        Set oGarantia = New COMNCredito.NCOMGarantia
        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)  ', CLng(Trim(Right(CmbTipoGarant.Text, 10))))
        If Not sCad = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            txtMontoRea.SetFocus
            Exit Sub
        End If
        '*******************************************
'            End If
      ' End If
       Set oPersona = Nothing
        If txtMontoxGrav.Enabled Then
            txtMontoxGrav.SetFocus
        Else
            txtMontoRea.Text = Format(txtMontoRea.Text, "#0.00")
        End If
     End If
End Sub
Function ValidacionCreditoAutomatico(ByVal psCtaCod As String) As Boolean

End Function

Function ObtenerTitularX() As String
    Dim i As Integer
    If FERelPers.Rows = 2 And FERelPers.TextMatrix(1, 1) = "" Then
       ObtenerTitularX = ""
    Else
        ReDim RelPers(FERelPers.Rows - 1, 4)
        For i = 1 To FERelPers.Rows - 1
            If Right("00" & Trim(Right(FERelPers.TextMatrix(i, 3), 10)), 2) = "01" Then
                ObtenerTitularX = FERelPers.TextMatrix(i, 1)   'Codigo de Persona
                Exit For
            End If
        Next i
    End If
End Function

Private Sub txtMontoRea_LostFocus()
    If Trim(txtMontoRea.Text) = "" Then
        txtMontoRea.Text = "0.00"
    Else
        txtMontoRea.Text = Format(txtMontoRea.Text, "#0.00")
    End If
    'Call txtMontoRea_KeyPress(13)
End Sub

Private Sub txtMontotas_GotFocus()
    fEnfoque txtMontotas
End Sub

Private Sub txtMontotas_KeyPress(KeyAscii As Integer)

     KeyAscii = NumerosDecimales(txtMontotas, KeyAscii)
     If KeyAscii = 13 Then
        txtMontoRea.SetFocus
     End If
End Sub

Private Sub txtMontotas_LostFocus()
    If Trim(txtMontotas.Text) = "" Then
        txtMontotas.Text = "0.00"
    Else
        txtMontotas.Text = Format(txtMontotas.Text, "#0.00")
    End If
End Sub

Private Sub txtMontoxGrav_GotFocus()
    fEnfoque txtMontoxGrav
End Sub

Private Sub txtMontoxGrav_KeyPress(KeyAscii As Integer)
Dim oGarantia As COMNCredito.NCOMGarantia
Dim sCad As String

     KeyAscii = NumerosDecimales(txtMontoxGrav, KeyAscii)
     If KeyAscii = 13 Then
        Set oGarantia = New COMNCredito.NCOMGarantia
        If Trim(txtMontoxGrav.Text) = "" Then
            txtMontoxGrav.Text = "0.00"
        End If
        sCad = oGarantia.ValidaDatos("", CDbl(txtMontotas.Text), CDbl(txtMontoRea.Text), CDbl(txtMontoxGrav.Text), True)
        If Not sCad = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        'txtcomentarios.SetFocus
        CmdAceptar.SetFocus
     End If
     
End Sub

Private Sub txtMontoxGrav_LostFocus()
    If Trim(txtMontoxGrav.Text) = "" Then
        txtMontoxGrav.Text = "0.00"
    Else
        txtMontoxGrav.Text = Format(txtMontoxGrav.Text, "#0.00")
    End If
End Sub


Private Sub TxtNroPoliza_GotFocus()
    fEnfoque TxtNroPoliza
End Sub

Private Sub TxtNroPoliza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtFecVig.SetFocus
    End If
End Sub

Private Sub txtNumDoc_GotFocus()
    fEnfoque txtNumDoc
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosEnteros(KeyAscii)
     If KeyAscii = 13 Then
        If Trim(Right(cmbMoneda.Text, 2)) = "2" Then
            txtMontotas.BackColor = RGB(200, 255, 200)
            txtMontoRea.BackColor = RGB(200, 255, 200)
            txtMontoxGrav.BackColor = RGB(200, 255, 200)
        Else
            txtMontotas.BackColor = vbWhite
            txtMontoRea.BackColor = vbWhite
            txtMontoxGrav.BackColor = vbWhite
        End If
        If txtDescGarant.Enabled Then
            txtDescGarant.SetFocus
        End If
     End If
End Sub

Private Sub TxtPrecioVenta_GotFocus()
    fEnfoque TxtPrecioVenta
End Sub

Private Sub TxtPrecioVenta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPrecioVenta, KeyAscii)
    If KeyAscii = 13 Then
        TxtValorCConst.SetFocus
    End If
End Sub

Private Sub TxtPrecioVenta_LostFocus()
Dim oGarantia As COMNCredito.NCOMGarantia
Dim sCad As String
Dim nPorc As Double

    If Trim(TxtPrecioVenta.Text) = "" Then
        TxtPrecioVenta.Text = "0.00"
    Else
        TxtPrecioVenta.Text = Format(TxtPrecioVenta.Text, "#0.00")
    End If
    
    If Trim(txtMontotas.Text) = "" Then
        txtMontotas.Text = "0.00"
    End If
    
    If Me.ChkGarReal.value = 1 Then
        Set oGarantia = New COMNCredito.NCOMGarantia
        If Trim(Right(CmbTipoGarant.Text, 5)) <> "" Then
            If CInt(Trim(Right(CmbTipoGarant.Text, 5))) = gPersGarantiaHipotecas Then
                nPorc = oGarantia.PorcentajeGarantia("3051")
                If Abs(((CDbl(txtMontotas.Text) - CDbl(TxtPrecioVenta.Text)) / CDbl(txtMontotas.Text))) > nPorc Then
                    MsgBox "La Diferencia entre el Monto de Tasacion y el Precio de Venta no debe ser mayor a " & Format(nPorc * 100, "#0.00") & "%", vbInformation, "Aviso"
                    TxtPrecioVenta.Text = Format(((100 - nPorc) / 100) * CDbl(txtMontotas.Text), "#0.00")
                    TxtPrecioVenta.SetFocus
                    Exit Sub
                End If
            End If
        End If
        Set oGarantia = Nothing
    End If
End Sub


Private Sub TxtRegNro_GotFocus()
    fEnfoque TxtRegNro
End Sub

Private Sub TxtRegNro_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtTelefono_GotFocus()
    fEnfoque TxtTelefono
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CboTipoInmueb.SetFocus
    End If
End Sub

Private Sub TxtValorCConst_GotFocus()
    fEnfoque TxtValorCConst
End Sub

Private Sub TxtValorCConst_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtValorCConst, KeyAscii)
    If KeyAscii = 13 Then
        CmdBuscaTasa.SetFocus
    End If
End Sub

Private Sub TxtValorCConst_LostFocus()
    If Trim(TxtValorCConst.Text) = "" Then
        TxtValorCConst.Text = "0.00"
    Else
        TxtValorCConst.Text = Format(TxtValorCConst.Text, "#0.00")
    End If
    
End Sub

Sub CargarSuperGarantias(ByVal prs As ADODB.Recordset)
'    Dim objDGarantias As COMDCredito.DCOMGarantia
'    Dim rs As ADODB.Recordset
    Dim sDes As String
    Dim nCodigo As Integer
    On Error GoTo ErrHandler
        'Set objDGarantias = New COMDCredito.DCOMGarantia
        'Set rs = objDGarantias.ListaSuperGarantias
        'Set objDGarantias = Nothing
        
        Do Until prs.EOF
            nCodigo = prs!nConsValor
            sDes = prs!cConsDescripcion
            
            CboGarantia.AddItem sDes
            CboGarantia.ItemData(CboGarantia.NewIndex) = nCodigo
            
            prs.MoveNext
        Loop
    Exit Sub
ErrHandler:
    'If Not objDGarantias Is Nothing Then Set objDGarantias = Nothing
    'If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error al cargar las garantias", vbInformation, "AVISO"
End Sub


'Sub ReconfigurarSubTipoGarant(ByVal pIdTipoGarant As Integer)
'    Dim i As Integer
'
'    For i = 0 To CmbTipoGarant.ListCount - 1
'        If Trim(Left(CmbTipoGarant.List(i), 3)) = pIdTipoGarant Then
'            CmbTipoGarant.RemoveItem (i)
'        End If
'    Next i
'End Sub

Sub ReLoadCmbTipoGarant(ByVal pnIdSuperGarant As Integer)
    Dim rs As ADODB.Recordset
    Dim objDGarantia As COMDCredito.DCOMGarantia
    On Error GoTo ErrHandler
        Set objDGarantia = New COMDCredito.DCOMGarantia
        Set rs = objDGarantia.CargarRelGarantia(pnIdSuperGarant)
        Set objDGarantia = Nothing
        If Not rs.EOF And Not rs.BOF Then
            CmbTipoGarant.Clear
        End If
                
        Select Case pnIdSuperGarant
            Case 1, 2
                OptCG(1).value = True
            Case 3
                OptCG(0).value = True
        End Select
        
        Do Until rs.EOF
            CmbTipoGarant.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim(Str(rs!nConsValor))
            rs.MoveNext
        Loop
        Set rs = Nothing
    Exit Sub
ErrHandler:
    If Not objDGarantia Is Nothing Then Set objDGarantia = Null
    If Not rs Is Nothing Then Set rs = Nothing
    MsgBox "Error al cargaer"
End Sub


Sub PosicionSuperGarantias(ByVal pintIndex As Integer)
'    Dim objDGarantia As DGarantia
    Dim i As Integer
    Dim nValorGarant As Integer
    On Error GoTo ErrHandler
'        Set objDGarantia = New DGarantia
'        nValorGarant = objDGarantia.ObtenerIdSuperGarantia(pintIndex)
'        Set objDGarantia = Nothing
        
        For i = 0 To CboGarantia.ListCount - 1
               If CboGarantia.ItemData(i) = pintIndex Then
                  CboGarantia.ListIndex = i
                  Exit For
               End If
        Next i
    Exit Sub
ErrHandler:
    'If Not objDGarantia Is Nothing Then Set objDGarantia = Nothing
    MsgBox "Error a cargar super garantia", vbInformation, "AVISO"
End Sub

Public Function IsLoadForm(ByVal FormCaption As String, Optional Active As Variant) As Boolean
    Dim rtn As Integer, i As Integer
    Dim Name As String
        
    rtn = False
    Name = LCase(FormCaption)
    Do Until i > Forms.Count - 1 Or rtn
        If LCase(Forms(i).Caption) = FormCaption Then

        rtn = True

End If
        i = i + 1
    Loop
    
    If rtn Then
        If Not IsMissing(Active) Then
            If Active Then
                Forms(i - 1).WindowState = vbNormal
            End If
        End If
    End If
    IsLoadForm = rtn
End Function



