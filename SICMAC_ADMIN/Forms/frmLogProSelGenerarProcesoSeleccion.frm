VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogProSelGenerarProcesoSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Proceso de Seleccion"
   ClientHeight    =   6285
   ClientLeft      =   225
   ClientTop       =   1935
   ClientWidth     =   11550
   Icon            =   "frmLogProSelGenerarProcesoSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   110.861
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   203.729
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin TabDlg.SSTab sstPro 
      Height          =   5115
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   8
      TabHeight       =   635
      TabCaption(0)   =   "  Proceso de Selección      "
      TabPicture(0)   =   "frmLogProSelGenerarProcesoSeleccion.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "    Etapas del Proceso       "
      TabPicture(1)   =   "frmLogProSelGenerarProcesoSeleccion.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSEta"
      Tab(1).Control(1)=   "cmdactualizarEtapas"
      Tab(1).Control(2)=   "cmdBases"
      Tab(1).Control(3)=   "cmdEtapas"
      Tab(1).Control(4)=   "cmdQuitarEtapa"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "    Comite Responsable       "
      TabPicture(2)   =   "frmLogProSelGenerarProcesoSeleccion.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameEvalLista"
      Tab(2).Control(1)=   "FrameEval"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "   Factores de Evaluacion    "
      TabPicture(3)   =   "frmLogProSelGenerarProcesoSeleccion.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtEditItem"
      Tab(3).Control(1)=   "cmdRangos"
      Tab(3).Control(2)=   "txtTotalPuntos"
      Tab(3).Control(3)=   "cmdActualizar"
      Tab(3).Control(4)=   "MSFlex"
      Tab(3).Control(5)=   "FrameValores"
      Tab(3).Control(6)=   "Label27"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "    Archivos CONSUCODE    "
      TabPicture(4)   =   "frmLogProSelGenerarProcesoSeleccion.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton cmdQuitarEtapa 
         Caption         =   "&Quitar Etapa"
         Height          =   375
         Left            =   -65280
         TabIndex        =   144
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4635
         Left            =   120
         TabIndex        =   96
         Top             =   420
         Width           =   11115
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cerrar"
            Height          =   354
            Left            =   9870
            TabIndex        =   143
            Top             =   4260
            Width           =   1215
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "&Grabar"
            Height          =   354
            Left            =   8580
            TabIndex        =   142
            Top             =   4260
            Width           =   1215
         End
         Begin VB.Frame Frame12 
            Height          =   1035
            Left            =   0
            TabIndex        =   135
            Top             =   1440
            Width           =   5835
            Begin VB.ComboBox cboTpoValorRef 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":0956
               Left            =   4020
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":0958
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   240
               Width           =   1710
            End
            Begin VB.ComboBox CboModalidad 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":095A
               Left            =   1680
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":095C
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   255
               Width           =   1650
            End
            Begin VB.ComboBox CboFinanciamiento 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":095E
               Left            =   1680
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":0960
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   600
               Width           =   4050
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Valor Ref"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   3360
               TabIndex        =   141
               Top             =   300
               Width           =   600
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Modalidad del Proceso"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   180
               TabIndex        =   140
               Top             =   300
               Width           =   1380
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Fuente Financiamiento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   180
               TabIndex        =   139
               Top             =   660
               Width           =   1410
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Publicación "
            Height          =   1395
            Left            =   5940
            TabIndex        =   126
            Top             =   0
            Width           =   5175
            Begin VB.ComboBox cboMedio 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":0962
               Left            =   2340
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":0969
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   960
               Width           =   2655
            End
            Begin MSMask.MaskEdBox MEBFechaPROMPYME 
               Height          =   315
               Left            =   3720
               TabIndex        =   129
               Top             =   600
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MEBcomunicacion 
               Height          =   315
               Left            =   3720
               TabIndex        =   130
               Top             =   240
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Otro medio donde se publico"
               Height          =   195
               Left            =   180
               TabIndex        =   133
               Top             =   1020
               Width           =   2025
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de la Publicacion en el Peruano"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   180
               TabIndex        =   132
               Top             =   660
               Width           =   2805
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de comunicacion a PROMPYME"
               Height          =   195
               Left            =   180
               TabIndex        =   131
               Top             =   330
               Width           =   2820
            End
         End
         Begin VB.CommandButton cmdEspTecnicas 
            Caption         =   "Especificaciones Tecnicas"
            Height          =   360
            Left            =   0
            TabIndex        =   124
            Top             =   4260
            Width           =   5835
         End
         Begin VB.Frame Frame7 
            Caption         =   "Evaluación Económica"
            Height          =   675
            Left            =   5940
            TabIndex        =   118
            Top             =   3520
            Width           =   3135
            Begin VB.TextBox txtminimo 
               Height          =   300
               Left            =   900
               MaxLength       =   2
               ScrollBars      =   2  'Vertical
               TabIndex        =   123
               Text            =   "70"
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtmaximo 
               Height          =   300
               Left            =   2340
               MaxLength       =   3
               ScrollBars      =   2  'Vertical
               TabIndex        =   125
               Text            =   "110"
               Top             =   240
               Width           =   435
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Mayor al            %    Menor al             %"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   122
               Top             =   300
               Width           =   2760
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Sistema de Adquisiciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   5940
            TabIndex        =   115
            Top             =   2535
            Width           =   1875
            Begin VB.OptionButton OptContrato 
               Caption         =   "Suma Alzada"
               Height          =   195
               Left            =   240
               TabIndex        =   117
               Top             =   300
               Width           =   1335
            End
            Begin VB.OptionButton OptContrato1 
               Caption         =   "Precio Unitario"
               Height          =   195
               Left            =   240
               TabIndex        =   116
               Top             =   560
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Modalidad Cont. / Adq."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   7920
            TabIndex        =   113
            Top             =   2535
            Width           =   1755
            Begin VB.OptionButton OptModalidad 
               Caption         =   "Total"
               Height          =   195
               Left            =   240
               TabIndex        =   119
               Top             =   300
               Width           =   975
            End
            Begin VB.OptionButton OptModalidad1 
               Caption         =   "Por Item"
               Height          =   195
               Left            =   240
               TabIndex        =   114
               Top             =   560
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Evaluación Técnica"
            Height          =   675
            Left            =   9180
            TabIndex        =   111
            Top             =   3520
            Width           =   1935
            Begin VB.TextBox txtPuntajeMinimo 
               Height          =   300
               Left            =   1380
               MaxLength       =   2
               ScrollBars      =   2  'Vertical
               TabIndex        =   127
               Text            =   "30"
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Puntaje Minimo"
               Height          =   195
               Left            =   170
               TabIndex        =   112
               Top             =   300
               Width           =   1080
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Tipo de Acto"
            Height          =   1035
            Left            =   5940
            TabIndex        =   106
            Top             =   1440
            Width           =   5175
            Begin VB.ComboBox CboBuenaPro 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":0976
               Left            =   1440
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":097D
               Style           =   2  'Dropdown List
               TabIndex        =   108
               Top             =   240
               Width           =   3555
            End
            Begin VB.ComboBox CboPrePro 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":098A
               Left            =   1440
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":0991
               Style           =   2  'Dropdown List
               TabIndex        =   107
               Top             =   600
               Width           =   3555
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Buena Pro"
               Height          =   195
               Left            =   120
               TabIndex        =   110
               Top             =   300
               Width           =   750
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Presentacion de Propuesta"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   120
               TabIndex        =   109
               Top             =   540
               Width           =   1215
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Venta de Bases"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   9770
            TabIndex        =   104
            Top             =   2535
            Width           =   1350
            Begin VB.TextBox txtCostoBases 
               Height          =   315
               Left            =   660
               MaxLength       =   7
               ScrollBars      =   2  'Vertical
               TabIndex        =   121
               Text            =   "0"
               Top             =   480
               Width           =   615
            End
            Begin VB.ComboBox cboMonedaCosto 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":099E
               Left            =   60
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09A0
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               Caption         =   "Costo"
               Height          =   195
               Left            =   120
               TabIndex        =   105
               Top             =   240
               Width           =   1125
            End
         End
         Begin VB.ComboBox cbound 
            BackColor       =   &H00FFF2E1&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   330
            ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09A2
            Left            =   3480
            List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09A4
            Style           =   2  'Dropdown List
            TabIndex        =   103
            Top             =   3240
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00FFF2E1&
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   3520
            MaxLength       =   7
            TabIndex        =   102
            Top             =   3320
            Visible         =   0   'False
            Width           =   560
         End
         Begin VB.Frame Frame11 
            Height          =   1395
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Width           =   5835
            Begin VB.TextBox txtSintesis 
               Height          =   675
               Left            =   780
               MaxLength       =   254
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   99
               Top             =   585
               Width           =   4935
            End
            Begin VB.ComboBox CboObj 
               Height          =   315
               ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09A6
               Left            =   780
               List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09A8
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   240
               Width           =   4940
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Sintesis"
               Height          =   195
               Left            =   160
               TabIndex        =   101
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Objeto"
               Height          =   195
               Left            =   160
               TabIndex        =   100
               Top             =   300
               Width           =   465
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItem 
            Height          =   1590
            Left            =   0
            TabIndex        =   134
            ToolTipText     =   "Para Cambiar las Unidades darle Doble Click"
            Top             =   2595
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   2805
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   13
            FixedCols       =   0
            ForeColorFixed  =   -2147483646
            BackColorSel    =   16773857
            ForeColorSel    =   -2147483635
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            GridColorUnpopulated=   -2147483633
            FocusRect       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   13
         End
      End
      Begin VB.CommandButton cmdEtapas 
         Caption         =   "Asignar &Fechas"
         Height          =   375
         Left            =   -65280
         TabIndex        =   94
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdBases 
         Caption         =   "&Bases"
         Height          =   375
         Left            =   -65280
         TabIndex        =   93
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdactualizarEtapas 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   -65280
         TabIndex        =   92
         Top             =   1860
         Width           =   1335
      End
      Begin VB.TextBox txtEditItem 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   -74400
         MaxLength       =   6
         TabIndex        =   77
         Top             =   2700
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRangos 
         Caption         =   "&Rangos"
         Height          =   375
         Left            =   -65220
         TabIndex        =   75
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox txtTotalPuntos 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   -65280
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Text            =   "0"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   -65220
         TabIndex        =   72
         Top             =   1680
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   4005
         Left            =   -74820
         TabIndex        =   74
         Top             =   840
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   7064
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Frame FrameValores 
         Caption         =   "Rangos de Valores"
         Height          =   4410
         Left            =   -74940
         TabIndex        =   62
         Top             =   600
         Visible         =   0   'False
         Width           =   11160
         Begin VB.TextBox txtvalorMaximo 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   7320
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Text            =   "0"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancelarValores 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   9720
            TabIndex        =   66
            Top             =   1100
            Width           =   1335
         End
         Begin VB.CommandButton cmdGuardarValores 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   9720
            TabIndex        =   65
            Top             =   660
            Width           =   1335
         End
         Begin VB.TextBox txtEditRangos 
            BackColor       =   &H80000001&
            ForeColor       =   &H80000018&
            Height          =   255
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   64
            Top             =   960
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.CommandButton cmdQuitarRango 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   480
            TabIndex        =   63
            Top             =   1440
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexValoresVer 
            Height          =   1455
            Left            =   1800
            TabIndex        =   68
            Top             =   360
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   2566
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   -2147483630
            Cols            =   4
            FixedCols       =   0
            ForeColorFixed  =   -2147483646
            BackColorSel    =   -2147483647
            ForeColorSel    =   -2147483624
            BackColorBkg    =   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            GridColorUnpopulated=   -2147483633
            FocusRect       =   0
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Puntaje Maximo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7680
            TabIndex        =   71
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label lblFactorEval 
            Alignment       =   2  'Center
            Caption         =   "Factor de Evaluacion:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   795
            Left            =   6720
            TabIndex        =   70
            Top             =   240
            Width           =   2805
         End
         Begin VB.Label lblmensaje 
            Alignment       =   2  'Center
            Caption         =   "Error Puntaje Total debe ser 100 ptos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   675
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Visible         =   0   'False
            Width           =   1605
         End
      End
      Begin VB.Frame Frame10 
         Height          =   3315
         Left            =   -74880
         TabIndex        =   46
         Top             =   480
         Width           =   11055
         Begin VB.CommandButton cmdGenerarArchivo 
            Caption         =   "&Generar Archivos"
            Height          =   360
            Left            =   7200
            TabIndex        =   54
            Top             =   1935
            Width           =   2535
         End
         Begin VB.TextBox txtDirVentaBases 
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2160
            TabIndex        =   53
            Top             =   420
            Width           =   5445
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2160
            TabIndex        =   52
            Top             =   780
            Width           =   2415
         End
         Begin VB.CommandButton cmdUbiGeo 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   7215
            TabIndex        =   51
            Top             =   810
            Width           =   315
         End
         Begin VB.TextBox txtUbigeo 
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   780
            Width           =   3150
         End
         Begin VB.ComboBox cboCIIU 
            Height          =   315
            ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09AA
            Left            =   2160
            List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09AC
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1140
            Width           =   8565
         End
         Begin VB.TextBox txtObs 
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2160
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   1500
            Width           =   8565
         End
         Begin VB.TextBox TxtVersion 
            Height          =   300
            Left            =   9720
            MaxLength       =   20
            TabIndex        =   47
            Text            =   "1.2.1"
            Top             =   420
            Width           =   975
         End
         Begin VB.TextBox txtUbigeoCod 
            Height          =   300
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Direccion Venta de Bases"
            Height          =   195
            Left            =   180
            TabIndex        =   61
            Top             =   480
            Width           =   1845
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Cuenta de Banco "
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Ubicacion Geografica"
            Height          =   195
            Left            =   4740
            TabIndex        =   59
            Top             =   840
            Width           =   1545
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Codigo CIIU"
            Height          =   195
            Left            =   180
            TabIndex        =   58
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   180
            TabIndex        =   57
            Top             =   1560
            Width           =   1065
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Version Inf. Enviada Cte."
            Height          =   195
            Left            =   7860
            TabIndex        =   56
            Top             =   480
            Width           =   1755
         End
      End
      Begin VB.CommandButton cmdCara 
         Caption         =   "Características por Item"
         Height          =   615
         Left            =   -65880
         TabIndex        =   26
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton cmdPostor 
         Caption         =   "Propuestas de Postores"
         Height          =   555
         Left            =   -65880
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdNuevoItem 
         Caption         =   "Agregar Item"
         Height          =   375
         Left            =   -65880
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdConsultas 
         Caption         =   "Observaciones y/o Consultas "
         Height          =   615
         Left            =   -65880
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdQuitarItem 
         Caption         =   "Quitar Item"
         Height          =   375
         Left            =   -65880
         TabIndex        =   22
         Top             =   2220
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   -73080
         TabIndex        =   21
         Top             =   540
         Width           =   7095
      End
      Begin VB.CommandButton cmdArchivo 
         Caption         =   "Archivo"
         Height          =   375
         Left            =   -65880
         TabIndex        =   20
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   -73140
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09AE
         Left            =   -73140
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09B0
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   -71800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   -70900
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "CMAC-T S.A."
         Top             =   840
         Width           =   1920
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09B2
         Left            =   -73140
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09B4
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09B6
         Left            =   -70920
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09BD
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09CA
         Left            =   -66480
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09D1
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09DE
         Left            =   -70920
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09E0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   6375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   -67440
         TabIndex        =   9
         Top             =   2520
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":09E2
         Left            =   -66480
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":09E9
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   -70920
         TabIndex        =   11
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   315
         Left            =   -66480
         TabIndex        =   27
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSEta 
         Height          =   4320
         Left            =   -74880
         TabIndex        =   95
         Top             =   600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7620
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   8
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   16056285
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin VB.Frame FrameEvalLista 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4185
         Left            =   -74880
         TabIndex        =   88
         Top             =   780
         Width           =   11115
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSComite 
            Height          =   4095
            Left            =   75
            TabIndex        =   91
            Top             =   45
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7223
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   5
            FixedCols       =   0
            ForeColorFixed  =   -2147483646
            BackColorSel    =   16773857
            ForeColorSel    =   -2147483635
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            GridColorUnpopulated=   -2147483633
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "&Quitar"
            Height          =   375
            Left            =   9780
            TabIndex        =   90
            Top             =   420
            Width           =   1335
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agr&egar"
            Height          =   375
            Left            =   9780
            TabIndex        =   89
            Top             =   15
            Width           =   1335
         End
      End
      Begin VB.Frame FrameEval 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1995
         Left            =   -74940
         TabIndex        =   78
         Top             =   840
         Visible         =   0   'False
         Width           =   11055
         Begin VB.CheckBox ChkSuplente 
            Caption         =   "Suplente"
            Height          =   255
            Left            =   9600
            TabIndex        =   84
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox CboCargos 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   840
            Width           =   3375
         End
         Begin VB.CommandButton cmdCancelarComite 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8400
            TabIndex        =   82
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdGrabarComite 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   7080
            TabIndex        =   81
            Top             =   1500
            Width           =   1275
         End
         Begin VB.CommandButton cmdPersona 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   2685
            TabIndex        =   80
            Top             =   510
            Width           =   315
         End
         Begin VB.TextBox txtPersona 
            Appearance      =   0  'Flat
            Height          =   290
            Left            =   3030
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   480
            Width           =   7600
         End
         Begin VB.TextBox txtPersCod 
            Height          =   300
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
            Height          =   195
            Left            =   720
            TabIndex        =   87
            Top             =   900
            Width           =   420
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Left            =   600
            TabIndex        =   86
            Top             =   540
            Width           =   540
         End
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Puntaje Total:"
         Height          =   195
         Left            =   -65040
         TabIndex        =   76
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de Bases"
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
         Left            =   -74820
         TabIndex        =   38
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Left            =   -74880
         TabIndex        =   37
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Proceso"
         Height          =   195
         Left            =   -74880
         TabIndex        =   36
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Objeto"
         Height          =   195
         Left            =   -74880
         TabIndex        =   35
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Acto de Presentacion de Propuesta"
         Height          =   315
         Left            =   -74880
         TabIndex        =   34
         Top             =   1680
         Width           =   4005
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Acto de Buena Pro"
         Height          =   315
         Left            =   -68520
         TabIndex        =   33
         Top             =   1560
         Width           =   2085
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de la Publicacion en el Peruano o fecha de aviso a PROMPYME"
         Height          =   435
         Left            =   -74880
         TabIndex        =   32
         Top             =   2400
         Width           =   3885
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Comunicacion a PROMPYME"
         Height          =   435
         Left            =   -68520
         TabIndex        =   31
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label2 
         Caption         =   "Fuente de Financiamiento"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   2040
         Width           =   3045
      End
      Begin VB.OLE OLEBases 
         Class           =   "Word.Document.8"
         Height          =   2055
         Left            =   -74760
         OleObjectBlob   =   "frmLogProSelGenerarProcesoSeleccion.frx":09F6
         SourceDoc       =   "C:\Plan Anual\A.D.S N 018 -2005-CMAC-T MERCHANDISING.doc"
         TabIndex        =   29
         Top             =   960
         Width           =   8775
      End
      Begin VB.Label Label1 
         Caption         =   "Otro Medio donde se Publico"
         Height          =   315
         Left            =   -68640
         TabIndex        =   28
         Top             =   1200
         Width           =   2085
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   11355
      Begin VB.ComboBox cboTpo 
         Height          =   315
         ItemData        =   "frmLogProSelGenerarProcesoSeleccion.frx":25A0E
         Left            =   6480
         List            =   "frmLogProSelGenerarProcesoSeleccion.frx":25A10
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   600
         Width           =   2130
      End
      Begin VB.TextBox txtMes 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdProSel 
         Caption         =   "&Procesos de Selección"
         Height          =   375
         Left            =   4920
         TabIndex        =   0
         Top             =   180
         Width           =   3075
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox TxtProceso 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TxtTipo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtTipoNro 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtTipoAnio 
         Height          =   315
         Left            =   3370
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdConsultarProceso 
         Caption         =   "Consultar Procesos de Selección"
         Height          =   375
         Left            =   8040
         TabIndex        =   43
         Top             =   180
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9360
         TabIndex        =   5
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Left            =   8760
         TabIndex        =   42
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Nº Proceso"
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
         Left            =   180
         TabIndex        =   41
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Selección"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   660
         Width           =   1335
      End
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   420
      Picture         =   "frmLogProSelGenerarProcesoSeleccion.frx":25A12
      Top             =   5460
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNN 
      Height          =   240
      Left            =   120
      Picture         =   "frmLogProSelGenerarProcesoSeleccion.frx":25D54
      Top             =   5460
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogProSelGenerarProcesoSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gnProSelNro  As Integer, gcBSGrupoCod As String, gnProSelTpoCod As Integer, _
    gnProSelSubTpo As Integer, gcBases As String, gnNroPlan As Integer, gnMonto As Currency

Dim nKeyAscii As Integer


Private Sub CboBuenaPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CboPrePro.SetFocus
End If
End Sub

Private Sub CboFinanciamiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   MEBcomunicacion.SetFocus
End If
End Sub


Private Sub cboMedio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CboBuenaPro.SetFocus
End If
End Sub

Private Sub CboModalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboTpoValorRef.SetFocus
End If
End Sub

Private Sub cboTpoValorRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CboFinanciamiento.SetFocus
End If
End Sub



Private Sub CboObj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtSintesis.SetFocus
End If
End Sub


'Private Sub CmdConsultarPlan_Click()
'    frmLogCnsProcesoSeleccion.Inicio 1
'End Sub

Private Sub cbound_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then cbound.Visible = False
End Sub

Private Sub cmdActualizar_Click()
On Error GoTo cmdActualizarErr
    Dim oCon As DConecta, sSQL As String, i As Integer, nCol As Integer, nRow As Integer
    Set oCon = New DConecta
    If gnProSelNro = 0 Then Exit Sub
    If lblmensaje.Visible Then
        MsgBox "Total de Puntos Incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Seguro que Desea Actualizar los Datos?", vbQuestion + vbYesNo) = vbYes Then
        If oCon.AbreConexion Then
            With MSFlex
                nCol = .Col
                nRow = .row
                .Col = 0
                i = 1
                Do While i < .Rows
                    .row = i
                    If .CellPicture = imgOK Then
                        sSQL = "update LogProSelEvalFactor set nVigente=1, nPuntaje =" & .TextMatrix(i, 3) & _
                        "where nProSelTpoCod=" & gnProSelTpoCod & " and nProSelSubTpo=" & gnProSelSubTpo & _
                        "and nProSelNro=" & gnProSelNro & " and nProSelItem = " & MSItem.TextMatrix(MSItem.row, 6) & " and nFactorNro=" & .TextMatrix(i, 0)
                    Else
                        sSQL = "update LogProSelEvalFactor set nVigente=0" & _
                        "where nProSelTpoCod=" & gnProSelTpoCod & " and nProSelSubTpo=" & gnProSelSubTpo & _
                        "and nProSelNro=" & gnProSelNro & " and nProSelItem = " & MSItem.TextMatrix(MSItem.row, 6) & " and nFactorNro=" & .TextMatrix(i, 0)
                    End If
                    oCon.Ejecutar sSQL
                    i = i + 1
                Loop
                .Col = nCol
                .row = nRow
                MsgBox "Factores de Evaluacion se Actualizaron Correctamente", vbInformation
            End With
            oCon.CierraConexion
        End If
    End If
    Exit Sub
cmdActualizarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdactualizarEtapas_Click()
    Dim oCon As DConecta, sSQL As String, i As Integer, nCol As Integer, nRow As Integer
    Set oCon = New DConecta
    If gnProSelNro = 0 Then Exit Sub
    If MsgBox("Seguro que Desea Actualizar los Datos?", vbQuestion + vbYesNo) = vbYes Then
        With MSEta
            nCol = .Col
            nRow = .row
            .Col = 1
            If oCon.AbreConexion Then
                Do While i < .Rows
                    .row = i
                    If .CellPicture = imgOK Then
                        sSQL = "update LogProSelEtapa set nEstado =1 where nProSelNro = " & gnProSelNro & " and nEtapaCod = " & .TextMatrix(i, 0)
                    Else
                        sSQL = "update LogProSelEtapa set nEstado =0 where nProSelNro = " & gnProSelNro & " and nEtapaCod = " & Val(.TextMatrix(i, 0))
                    End If
                    oCon.Ejecutar sSQL
                    i = i + 1
                Loop
                oCon.CierraConexion
            End If
            MsgBox "Lista de Etapas Actualizadas Correctamente", vbInformation
            ListaEtapas gnProSelNro
        End With
    End If
End Sub

Private Sub cmdBases_Click()
    Dim sEtapas As String, sPlazo As String, i As Integer
    Dim wApp As Word.Application, sArchivoBases As String
    Set wApp = New Word.Application
    If gnProSelNro = 0 Then Exit Sub
    If gcBases <> "" Then
        If MsgBox("Ya Existe un Archivo de Bases desea Crear Otro", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            wApp.Documents.Open gcBases
            wApp.Visible = True
            Exit Sub
        Else
            GoTo eBases
        End If
    End If
eBases:
    i = 1
    sEtapas = ""
    sPlazo = ""
    With MSEta
        Do While i < .Rows
           If Val(.TextMatrix(i, 7)) >= 1 Then
              If sEtapas = "" Then
                'sEtapas = .TextMatrix(i, 2) & Space(30) & .TextMatrix(i, 3) & Space(15) & .TextMatrix(i, 4) & vbCrLf
                sEtapas = .TextMatrix(i, 2) & vbCrLf
              Else
                'sEtapas = sEtapas & .TextMatrix(i, 2) & Space(30) & .TextMatrix(i, 3) & Space(15) & .TextMatrix(i, 4) & vbCrLf
                sEtapas = sEtapas & .TextMatrix(i, 2) & vbCrLf
              End If
            
              If sPlazo = "" Then
                 sPlazo = "Del " & IIf(Len(.TextMatrix(i, 3)) = 0, Space(10), .TextMatrix(i, 3)) & " Al " & IIf(Len(.TextMatrix(i, 4)) = 0, Space(10), .TextMatrix(i, 4)) & vbCrLf
              Else
                 sPlazo = sPlazo & "Del " & IIf(Len(.TextMatrix(i, 3)) = 0, Space(10), .TextMatrix(i, 3)) & " Al " & IIf(Len(.TextMatrix(i, 4)) = 0, Space(10), .TextMatrix(i, 4)) & vbCrLf
              End If
           End If
           i = i + 1
        Loop
        sPlazo = sPlazo & Format(i, "00")
    End With
    
    If gnProSelNro = 0 Then Exit Sub
    
    If gnProSelSubTpo <> 3 And gnProSelSubTpo <> 4 Then
        sArchivoBases = CStr(gnProSelTpoCod & gnProSelSubTpo & "_" & CboObj.Text)
    Else
        sArchivoBases = CStr(gnProSelTpoCod & gnProSelSubTpo & "_")
    End If
    
    gcBases = ImpBasesWORD(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), TxtProceso.Text & " " & TxtTipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
                TxtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtMes.Text, txtCostoBases.Text, _
                txtPuntajeMinimo.Text, CboObj.Text, cboMonedaCosto.Text)
    
'    Select Case sArchivoBases
'        Case "11_BIENES"
'            gcBases = ImpBasesWORD_11_Bienes(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, _
'                txtPuntajeMinimo.Text, CboObj.Text, cboMonedaCosto.Text)
'        Case "11_SERVICIOS"
'            gcBases = ImpBasesWORD_11_SERVICIOS(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, txtPuntajeMinimo.Text, CboObj.Text)
'        Case "12_OBRAS"
'            gcBases = ImpBasesWORD_12_OBRAS(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, _
'                txtPuntajeMinimo.Text, CboObj.Text, cboMonedaCosto.Text)
'        Case "21_BIENES"
'            gcBases = ImpBasesWORD_21_BIENES(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, txtPuntajeMinimo.Text, CboObj.Text)
'        Case "21_OBRAS"
'            gcBases = ImpBasesWORD_21_OBRAS(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, _
'                txtPuntajeMinimo.Text, CboObj.Text, cboMonedaCosto.Text)
'        Case "21_SERVICIOS"
'            gcBases = ImpBasesWORD_21_SERVICIOS(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, _
'                txtPuntajeMinimo.Text, CboObj.Text)
'        Case "31_"
'            gcBases = ImpBasesWORD_31_(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, txtPuntajeMinimo.Text, CboObj.Text)
'        Case "41_"
'            gcBases = ImpBasesWORD_41_(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
'                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
'                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, txtPuntajeMinimo.Text, CboObj.Text)
'    End Select
    'gcBases = ImpBasesWORD(txtSintesis.Text, TxtMonto.Text, TxtTipoAnio.Text, IIf(LblMoneda.Caption = "S/.", "Nuevos Soles", "Dolares"), _
                IIf(OptModalidad.value, "Suma Alzada", "Por Item"), txtProceso.Text & " " & txttipo & "-" & TxtTipoNro.Text & "-" & TxtTipoAnio.Text & "-" & "CMAC-T-S.A.", _
                txtProceso.Text, sEtapas, sPlazo, txtminimo.Text, txtmaximo.Text, gnProSelNro, gnProSelTpoCod, gnProSelSubTpo, txtmes.Text, txtCostoBases.Text, txtPuntajeMinimo.Text, CboObj.Text)
End Sub

Private Sub cmdCancelar_Click()
    txtMes.Text = "":    LblMoneda.Caption = "":    txtArchivo.Text = "":    TxtMonto.Text = ""
    txtPersCod.Text = "":    txtPersona.Text = "":    TxtProceso.Text = "":    TxtTipo.Text = ""
    TxtTipoAnio.Text = "":    TxtTipoNro.Text = "":    txtSintesis.Text = "":    CboBuenaPro.ListIndex = 0
    CboCargos.ListIndex = 0:    CboFinanciamiento.ListIndex = 0:    cboMedio.ListIndex = 0:    CboModalidad.ListIndex = 0
    CboObj.ListIndex = 0:    CboPrePro.ListIndex = 0:    MEBcomunicacion.Text = "__/__/____":    MEBFechaPROMPYME.Text = "__/__/____"
    txtmaximo = "110":    txtminimo = "70":    txtCostoBases.Text = "0":    txtPuntajeMinimo = "30":    FormaFlexEta
    FormaFlexItem
    FormaMSComite
    FormaFlexFactores
    sstPro.Tab = 0:    gnProSelNro = 0:    gnProSelSubTpo = 0:    gnProSelTpoCod = 0:    gcBSGrupoCod = "":    gcBases = ""
    OptContrato1.value = True:    OptModalidad1.value = True:    cmdGenerarArchivo.Enabled = False:    If cboTpo.ListCount > 0 Then cboTpo.ListIndex = 0
    txtUbigeo.Text = "": txtUbigeoCod.Text = "": txtDirVentaBases.Text = "": txtCuenta.Text = "": cmdGuardar.Enabled = True: txtObs.Text = ""
    If cboMonedaCosto.ListCount > 0 Then cboMonedaCosto.ListIndex = 0: cmdBases.Enabled = False
    Unload Me
End Sub

Private Sub cmdCancelarValores_Click()
    MSFlex.Enabled = True
    FrameValores.Visible = False
    sstPro.TabEnabled(0) = True
    sstPro.TabEnabled(1) = True
End Sub

Private Sub cmdEspTecnicas_Click()
Dim i As Integer
i = MSItem.row
If Len(MSItem.TextMatrix(i, 1)) <= 1 Then Exit Sub
If Val(MSItem.TextMatrix(i, 5)) = 0 And Val(MSItem.TextMatrix(i, 6)) = 0 Then Exit Sub
   frmLogProSelRegistroDatosItem.Inicio MSItem.TextMatrix(i, 5), MSItem.TextMatrix(i, 6), , , 1, , MSItem.TextMatrix(i, 1)
'Else
'   MsgBox "No se halla un Proceso/Item válido..." + Space(10), vbInformation
'End If
End Sub

Private Sub cmdEtapas_Click()
   If MSEta.Rows <= 2 Then Exit Sub
   With MSEta
    frmLogProSelRegistroDatosItem.Inicio gnProSelNro, MSEta.TextMatrix(MSEta.row, 0), IIf(.TextMatrix(.row, 3) = "", gdFecSis, .TextMatrix(.row, 3)), IIf(.TextMatrix(.row, 4) = "", gdFecSis, .TextMatrix(.row, 4)), 4, MSEta.TextMatrix(MSEta.row, 2)
    If frmLogProSelRegistroDatosItem.vpGrabado Then
        .TextMatrix(.row, 3) = frmLogProSelRegistroDatosItem.gdFechaI
        .TextMatrix(.row, 4) = frmLogProSelRegistroDatosItem.gdFechaF
        .TextMatrix(.row, 5) = frmLogProSelRegistroDatosItem.txtObs
    End If
   End With
'   ListaEtapas gnProSelNro
End Sub

Private Function CargarTipoProcesoCONSUCODE(ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pnTpo As Integer) As Integer
On Error GoTo CargarTipoProcesoCONSUCODEErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nConsValor from LogTipoProcesoCONSUCODE where nProSelTpoCod = " & pnProSelTpoCod & " and nProSelSubTpo=" & pnProSelSubTpo & " and nTipo =" & pnTpo
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            CargarTipoProcesoCONSUCODE = Rs(0)
        End If
        oCon.CierraConexion
    End If
    Exit Function
CargarTipoProcesoCONSUCODEErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Sub cmdGenerarArchivo_Click()
On Error GoTo cmdGenerarArchivoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, sCampos As String, _
    conexiondbf As New ADODB.Connection, sCamposValores As String, sCadena  As String, _
    cTipoCONSUCODE As String
    
    If Len(txtUbigeoCod.Text) = 0 Or Len(txtUbigeo.Text) = 0 Then
        MsgBox "Debe Seleccionar una Ubicacion Geografica", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select cCodCampo, cNomCampo, cTipoDato, nLongitud1, nLongitud2, nObligatorio  from convtablasdbf where cCodFormato=" & gsArchCONVOCA
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            If sCampos = "" Then
                Select Case Rs!cTipoDato
                    Case "N"
                        sCampos = Rs!cNomCampo & " N(" & Rs!nLongitud1 & "," & Rs!nLongitud2 & ")"
                    Case "C"
                        sCampos = Rs!cNomCampo & " C(" & Rs!nLongitud1 & ")"
                End Select
            Else
                Select Case Rs!cTipoDato
                    Case "N"
                        sCampos = sCampos & "," & Rs!cNomCampo & " N(" & Rs!nLongitud1 & "," & Rs!nLongitud2 & ")"
                    Case "C"
                        sCampos = sCampos & "," & Rs!cNomCampo & " C(" & Rs!nLongitud1 & ")"
                End Select
            End If
            Rs.MoveNext
        Loop
        oCon.CierraConexion
        
        sSQL = "CREATE TABLE " & App.path & "\spooler\CONVOCA(" & sCampos & ")"
                    
        conexiondbf.Open gsConnectionLOG_DBF
        conexiondbf.Execute sSQL
        conexiondbf.Close
                
        'MsgBox "Creacion de Tabla Completa CONVOCA", vbInformation
        
        Rs.MoveFirst
        sCampos = ""
        
        cTipoCONSUCODE = CargarTipoProcesoCONSUCODE(gnProSelTpoCod, gnProSelSubTpo, cboTpo.ItemData(cboTpo.ListIndex))
        
        Do While Not Rs.EOF
            If sCampos = "" Then
                sCampos = Rs!cNomCampo
            Else
                sCampos = sCampos & "," & Rs!cNomCampo
            End If
            
            Select Case Rs!cNomCampo
                Case "P_TIPO"
                    sCamposValores = sCamposValores & cTipoCONSUCODE
                Case "P_NUM"
                    sCamposValores = sCamposValores & "," & TxtTipoNro.Text
                Case "P_ANHO"
                    sCamposValores = sCamposValores & ",'" & TxtTipoAnio.Text & "'"
                Case "P_SIGLA"
                    sCamposValores = sCamposValores & ",'" & gsCmact & "'"
                Case "VERSION"
                    sCamposValores = sCamposValores & ",'" & TxtVersion.Text & "'"
                Case "PL_NUM"
                    sCamposValores = sCamposValores & "," & gnNroPlan
                Case "T_VALREF"
                    sCamposValores = sCamposValores & "," & cboTpoValorRef.ItemData(cboTpoValorRef.ListIndex)
                Case "T_MONEDA"
                    sCamposValores = sCamposValores & "," & IIf(LblMoneda.Caption = "$", 2, 1)
                Case "C_VALREF"
                    sCamposValores = sCamposValores & "," & CDbl(TxtMonto.Text)
                Case "T_OBJETO"
                    sCamposValores = sCamposValores & "," & CboObj.ItemData(CboObj.ListIndex)
                Case "C_SINTES"
                    sCamposValores = sCamposValores & ",'" & txtSintesis.Text & "'"
                Case "C_CIIU"
                    sCamposValores = sCamposValores & "," & "'" & Right(cboCIIU.Text, 4) & "'"
                Case "C_FUENTE"
                    sCamposValores = sCamposValores & ",'" & CboFinanciamiento.ItemData(CboFinanciamiento.ListIndex) & "'"
                Case "C_DIRVENTA"
                    sCamposValores = sCamposValores & "," & "'" & txtDirVentaBases.Text & "'"
                Case "C_COST"
                    sCamposValores = sCamposValores & "," & txtCostoBases.Text
                Case "F_VEN_INI"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnConvocatoria, True) & "'"
                Case "F_VEN_FIN"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnConvocatoria, False) & "'"
                Case "F_PC_INI"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnAbsolucionConsultas, True) & "'"
                Case "F_PC_FIN"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnAbsolucionConsultas, False) & "'"
                Case "F_AC"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnAtencionEvaluacionConsultas, True) & "'"
                Case "F_PO_INI"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnObservaciones, True) & "'"
                Case "F_PO_FIN"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnObservaciones, False) & "'"
                Case "F_IB"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnIntegracionBases, True) & "'"
                Case "F_PP"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnPresentacionPropuestas, True) & "'"
                Case "F_BP"
                    sCamposValores = sCamposValores & "," & "'" & FechaEtapa(cnOtorgamientoBP, True) & "'"
                Case "C_OBS"
                    sCamposValores = sCamposValores & "," & "'" & txtObs.Text & "'"
                Case "FILEBASE"
                    Dim nPos As Integer, sNomBases
                    nPos = InStr(1, gcBases, "BASE")
                    If nPos > 0 Then sNomBases = Mid(gcBases, nPos)
                    sCamposValores = sCamposValores & ",'" & sNomBases & "'"
                Case "C_INDELCT"
                    sCamposValores = sCamposValores & "," & gnConvocatoriaElectrónica
                Case "C_CTA"
                    sCamposValores = sCamposValores & "," & "'" & txtCuenta.Text & "'"
                Case "T_MON_BAS"
                    sCamposValores = sCamposValores & "," & cboMonedaCosto.ItemData(cboMonedaCosto.ListIndex)
                Case "T_ACT_PP"
                    sCamposValores = sCamposValores & "," & CboPrePro.ItemData(CboPrePro.ListIndex)
                Case "T_ACT_BP"
                    sCamposValores = sCamposValores & "," & CboBuenaPro.ItemData(CboBuenaPro.ListIndex)
                Case "F_PUB_CON"
                    sCamposValores = sCamposValores & ",'" & IIf(MEBFechaPROMPYME.Text = "__/__/____", "", MEBFechaPROMPYME.Text) & "'"
                Case "C_MEDPUB"
                    sCamposValores = sCamposValores & "," & cboMedio.ItemData(cboMedio.ListIndex)
            End Select
            Rs.MoveNext
        Loop
        
        sSQL = " INSERT INTO " & App.path & "\SPOOLER\TABLA(" & sCampos & " )" & _
               " VALUES(" & sCamposValores & ")"
        
        conexiondbf.Open gsConnectionLOG_DBF
        conexiondbf.Execute sSQL
        conexiondbf.Close
        Set conexiondbf = Nothing
        
'        sCadena = """" & App.path & "\transfer.exe""" & Space(2) & App.path & "\SPOOLER\" & "TABLA.DBF " & """" & App.path & "\SPOOLER\CONVOCA.dbf"
'        ChDir App.path & "\Spooler\"
'        Shell sCadena, vbHide
'        ChDir App.path
        MsgBox "Creacion del Archivo CONVOCA.DBF Terminada", vbInformation, "Aviso"
        
        sCampos = ""
        sSQL = "select cCodCampo, cNomCampo, cTipoDato, nLongitud1, nLongitud2, nObligatorio  from convtablasdbf where cCodFormato=" & gsArchITEM
        oCon.AbreConexion
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            If sCampos = "" Then
                Select Case Rs!cTipoDato
                    Case "N"
                        sCampos = Rs!cNomCampo & " N(" & Rs!nLongitud1 & "," & Rs!nLongitud2 & ")"
                    Case "C"
                        sCampos = Rs!cNomCampo & " C(" & Rs!nLongitud1 & ")"
                End Select
            Else
                Select Case Rs!cTipoDato
                    Case "N"
                        sCampos = sCampos & "," & Rs!cNomCampo & " N(" & Rs!nLongitud1 & "," & Rs!nLongitud2 & ")"
                    Case "C"
                        sCampos = sCampos & "," & Rs!cNomCampo & " C(" & Rs!nLongitud1 & ")"
                End Select
            End If
            Rs.MoveNext
        Loop
        oCon.CierraConexion
        
        sSQL = "CREATE TABLE " & App.path & "\spooler\ITEM (" & sCampos & ")"
                    
        conexiondbf.Open gsConnectionLOG_DBF
        conexiondbf.Execute sSQL
        conexiondbf.Close
        'MsgBox "Creacion de Tabla Completa ITEM", vbInformation
        
        sCamposValores = ""
        sCampos = ""
        
        Rs.MoveFirst
        Do While Not Rs.EOF
            If sCampos = "" Then
                sCampos = Rs!cNomCampo
            Else
                sCampos = sCampos & "," & Rs!cNomCampo
            End If
            Rs.MoveNext
        Loop
            
        conexiondbf.Open gsConnectionLOG_DBF
        
        Dim i As Integer
        
        i = 1
        With MSItem
        Do While i < .Rows
            If .TextMatrix(i, 0) <> "+" And .TextMatrix(i, 0) <> "-" Then
                sCamposValores = cTipoCONSUCODE & "," & TxtTipoNro.Text & ",'" & TxtTipoAnio.Text & "','" & _
                                gsCmact & "'," & i & "," & CDbl(.TextMatrix(i, 11)) & "," & CDbl(.TextMatrix(i, 4)) & ",'" & _
                                .TextMatrix(i, 3) & "','" & Mid(txtUbigeoCod.Text, 1, 2) & "','" & Mid(txtUbigeoCod.Text, 3, 2) & "','" & _
                                Mid(txtUbigeoCod.Text, 5, 2) & "','" & TxtVersion.Text & "','" & .TextMatrix(i, 10) & "'," & _
                                .TextMatrix(i, 9)
                sSQL = " INSERT INTO " & App.path & "\SPOOLER\TABLAITEM (" & sCampos & " )" & _
                   " VALUES(" & sCamposValores & ")"
                conexiondbf.Execute sSQL
            End If
            i = i + 1
        Loop
        End With
        conexiondbf.Close
        Set conexiondbf = Nothing
        
'        sCadena = """" & App.path & "\transfer.exe""" & Space(2) & App.path & "\SPOOLER\" & "TABLAITEM.DBF " & """" & App.path & "\SPOOLER\ITEM.dbf"
'        ChDir App.path & "\Spooler\"
'        Shell sCadena, vbHide
'        ChDir App.path
        MsgBox "Creacion del Archivo ITEM.DBF Terminada", vbInformation, "Aviso"
    End If
    Exit Sub
cmdGenerarArchivoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo cmdGuardar_ClickErr
    Dim nNro As Integer
    If Len(Trim(TxtTipoNro)) = 0 Then Exit Sub
'    If Not IsDate(MEBcomunicacion.Text) Or MEBcomunicacion.Text = "01/01/1900" Then Exit Sub
'    If Not IsDate(MEBFechaPROMPYME.Text) Or MEBFechaPROMPYME.Text = "01/01/1900" Then Exit Sub
    
    If lblmensaje.Visible Then
        MsgBox "El Total de Puntaje debe ser 100 Ptos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Not VerificaFechas Then
        MsgBox "Debe Ingresar Fechas Correspondientes a cada Etapa ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MSComite.TextMatrix(1, 4) = "" Then
        MsgBox "Debe Ingresar los Miembros del Comite Responsable", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Len(txtDirVentaBases.Text) = 0 And Len(txtUbigeoCod.Text) = 0 Then
        MsgBox "Debe Ingresar los Datos Necesarios para la Generacion de los Archivos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Seguro que Desea Modificar Los Datos del Proceso de Seleccion", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If Val(TxtTipoNro.Text) = 0 Then
        nNro = NroProceso(TxtTipoAnio.Text)
    Else
        nNro = Val(TxtTipoNro.Text)
    End If
    Dim oConn As New DConecta, sSQL As String
    sSQL = "update LogProcesoSeleccion " & _
            "set nModalidad= " & CboModalidad.ItemData(CboModalidad.ListIndex) & "," & _
            "cSintesis = '" & txtSintesis.Text & "'," & _
            "nNroProceso= " & nNro & "," & _
            "nPresentaPropuestaTpo= " & CboPrePro.ItemData(CboPrePro.ListIndex) & "," & _
            "nBuenaProTpo= " & CboBuenaPro.ItemData(CboBuenaPro.ListIndex) & ","
    
    If MEBFechaPROMPYME.Text <> "__/__/____" Then sSQL = sSQL & "dFechaPublicacion= '" & Format(MEBFechaPROMPYME.Text, "yyyymmdd") & "',"
    If MEBcomunicacion.Text <> "__/__/____" Then sSQL = sSQL & "dFechaPROMPYME = '" & Format(MEBcomunicacion.Text, "yyyymmdd") & "',"
    
    sSQL = sSQL & "nCostoBases=" & txtCostoBases.Text & ", "
    sSQL = sSQL & "nMonedaCostoBases=" & cboMonedaCosto.ItemData(cboMonedaCosto.ListIndex) & ", "
    sSQL = sSQL & "nRangoMenor= " & txtminimo.Text & ", "
    sSQL = sSQL & "cDirVentaBases= '" & txtDirVentaBases.Text & "', "
    sSQL = sSQL & "cNroCuenta= '" & txtCuenta.Text & "', "
    sSQL = sSQL & "cUbiGeo= '" & txtUbigeoCod.Text & "', "
    sSQL = sSQL & "cCIIU= '" & Right(cboCIIU.Text, 4) & "', "
    sSQL = sSQL & "cObservaciones= '" & txtObs.Text & "', "
    sSQL = sSQL & "cVersionCONSUCODE= '" & TxtVersion.Text & "', "
    sSQL = sSQL & "nPuntajeMinimoTecnico= " & txtPuntajeMinimo.Text & ", "
    sSQL = sSQL & "nFuenteFinanciemiento= " & CboFinanciamiento.ItemData(CboFinanciamiento.ListIndex) & ", " & _
            " nMedioPublicacion= " & cboMedio.ItemData(cboMedio.ListIndex) & ", " & _
            " nObjetoCod= " & CboObj.ItemData(CboObj.ListIndex) & ", " & _
            " nModalidadCompra= " & IIf(OptModalidad.value, 1, 0) & ", " & _
            " nContratoCompra= " & IIf(OptContrato.value, 1, 0) & ", " & _
            " nProSelMonto= " & CDbl(TxtMonto.Text) & _
            " Where nProSelNro = " & gnProSelNro
    If oConn.AbreConexion Then
        oConn.Ejecutar sSQL
        oConn.CierraConexion
        actualizaUnidades
        MsgBox "Proceso Modificado ... ", vbInformation
        TxtTipoNro.Text = nNro
        If txtCostoBases.Text = "0" Then txtCuenta.Enabled = False
        cmdGenerarArchivo.Enabled = True
        cmdBases.Enabled = True
        sstPro.Tab = 3
        oConn.CierraConexion
    End If
    
Exit Sub
cmdGuardar_ClickErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Function NroProceso(ByVal pnAnio As Integer) As Integer
On Error GoTo NroProcesoErr
    Dim oCon As DConecta, sSQL As String, Nro As Integer, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select Nro=isnull(max(nNroProceso),0) from LogProcesoSeleccion where nPlanAnualAnio=" & pnAnio
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            NroProceso = Rs!Nro + 1
        End If
        Set Rs = Nothing
        oCon.CierraConexion
    End If
    Exit Function
NroProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub cmdGuardarValores_Click()
On Error GoTo cmdGuardarValoresErr
    Dim oCon As DConecta, sSQL As String, i As Integer
    If Not ValidarRangosValorMax Then
        MsgBox "Los Puntajes no Deben Exeder el Maximo", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        With MSFlexValoresVer
            i = 1
            Do While i < .Rows
                sSQL = "update LogProSelEvalFactorRangos set nPuntaje = " & .TextMatrix(i, 3) & "," & _
                       " nRangoMin = " & .TextMatrix(i, 1) & "," & _
                       " nRangoMax = " & .TextMatrix(i, 2) & _
                       " where nFactorNro=" & MSFlex.TextMatrix(MSFlex.row, 0) & " and  nProSelNro = " & gnProSelNro & " and nRangoItem = " & .TextMatrix(i, 0)
                oCon.Ejecutar sSQL
                i = i + 1
            Loop
        End With
        oCon.CierraConexion
    End If
    MsgBox "Rangos Modificados", vbInformation
    cmdCancelarValores_Click
Exit Sub
cmdGuardarValoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona

Set X = frmBuscaPersona.Inicio(True)
If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersona.Text = X.sPersNombre
   txtPersCod = X.sPersCod
End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersona.Text = frmBuscaPersona.vpPersNom
'   txtPersCod = frmBuscaPersona.vpPersCod
'End If
End Sub

Private Sub cmdProSel_Click()
Inicializar
'cmdCancelar_Click
frmLogProSelCnsProcesoSeleccion.Inicio 1
With frmLogProSelCnsProcesoSeleccion
    If .gbBandera Then
        If Not VerificaEtapaCerrada(.gvnProSelNro, cnConvocatoria) Then
            gnProSelNro = .gvnProSelNro
            gcBSGrupoCod = .gvcBSGrupoCod
            gnProSelTpoCod = .gvnProSelTpoCod
            gnProSelSubTpo = .gvnProSelSubTpo
            gnNroPlan = .gvnNroPlan
            TxtProceso.Text = .gvcTipo
            TxtMonto.Text = Format(.gvnMonto, "###,###.00")
            gnMonto = .gvnMonto
            LblMoneda.Caption = .gvcMoneda
            TxtTipoAnio.Text = .gvnAnio
            txtMes.Text = .gvcMes
            gcBases = .gvcArchivoBases
            
            ListaEtapas gnProSelNro
            GeneraDetalleItem gnProSelNro, gcBSGrupoCod
            CargarCargosComite
            cargarComiteItemProceso gnProSelNro
            consultarProseso gnProSelNro
            MSItem.SetFocus
        Else
            MsgBox "Etapa Cerrada", vbInformation, "Aviso"
        End If
    End If
End With
End Sub

Private Sub CargarMonedaBases()
On Error GoTo CargarMonedaBasesErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nConsValor, cConsDescripcion from constante where nConsCod = 1011"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            cboMonedaCosto.AddItem IIf(Rs!cConsDescripcion = "SOLES", "S/.", "$"), cboMonedaCosto.ListCount
            cboMonedaCosto.ItemData(cboMonedaCosto.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    If cboMonedaCosto.ListCount > 0 Then cboMonedaCosto.ListIndex = 0
    Exit Sub
CargarMonedaBasesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdQuitarEtapa_Click()
Dim oConn As New DConecta
Dim k As Integer
Dim sSQL As String, nEtapaCod As Integer
Dim nProselNro As Integer

nProselNro = Val(TxtTipoNro.Text)
nEtapaCod = MSEta.TextMatrix(MSEta.row, 0)

If MsgBox("¿ Está seguro de quitar la etapa ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
      sSQL = " UPDATE LogProSelEtapa SET nEstado = 0 " & _
           " Where nProSelNro = " & Val(TxtTipoNro.Text) & "  and nEtapaCod = " & nEtapaCod & " " & _
           " "
      oConn.Ejecutar sSQL
      ListaEtapas nProselNro
   End If
End If
End Sub

Private Sub cmdQuitarRango_Click()
    On Error GoTo cmdQuitarRangoErr
    Dim oCon As DConecta, sSQL As String, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        With MSFlexValoresVer
            i = .row
                sSQL = "delete LogProSelEvalFactorRangos " & _
                       " where nFactorNro=" & MSFlex.TextMatrix(MSFlex.row, 0) & " and  nProSelNro = " & gnProSelNro & " and nRangoItem = " & .TextMatrix(i, 0)
                oCon.Ejecutar sSQL
        End With
        oCon.CierraConexion
    End If
    MsgBox "Rango Eliminado", vbInformation
    CargarFactoresVer MSFlex.TextMatrix(MSFlex.row, 0), gnProSelNro
Exit Sub
cmdQuitarRangoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub cmdRangos_Click()
    With MSFlex
        If gnProSelNro = 0 Then Exit Sub
        If Val(.TextMatrix(.row, 5)) = 2 Then
            txtvalorMaximo.Text = .TextMatrix(.row, 3)
            CargarFactoresVer .TextMatrix(.row, 0), gnProSelNro
            lblFactorEval.Caption = "Factor de Evaluacion: " & vbCrLf & .TextMatrix(.row, 1)
            MSFlex.Enabled = False
            sstPro.TabEnabled(0) = False
            sstPro.TabEnabled(1) = False
            FrameValores.Visible = True
        Else
            MSFlex.Enabled = True
            FrameValores.Visible = False
        End If
    End With
End Sub


Private Sub cmdUbiGeo_Click()
    With frmLogProSelSeleUbiGeo
        .FuenteConsucode True
        txtUbigeoCod.Text = .gvCodigo
        txtUbigeo.Text = .gvNoddo
    End With
End Sub

Private Sub Form_Load()
CentraForm Me
Inicializar
cargarFuentesFinanciamiento
CargarModalidad
CargarObjeto
CargarPrePro
CargarMedio
CargarCargosComite
CargarCboTpo
CargarCboTpoValorRef
CargraUnidades
CargarMonedaBases
'cmdCancelar_Click
cmdGenerarArchivo.Enabled = False
cmdBases.Enabled = False
If cboCIIU.ListCount = 0 Then CargarCIIU
sstPro.Tab = 0
End Sub

Private Sub CargraUnidades()
On Error GoTo CargraUnidadesErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nConsValor, cConsDescripcion from constante where nconscod = 9089 and nconscod <> nconsvalor"
        Set Rs = oCon.CargaRecordSet(sSQL)
        cbound.Clear
        Do While Not Rs.EOF
            cbound.AddItem Rs!cConsDescripcion, cbound.ListCount
            cbound.ItemData(cbound.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    If cbound.ListCount > 0 Then cbound.ListIndex = 0
    Exit Sub
CargraUnidadesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarCboTpoValorRef()
On Error GoTo CargarTpoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nConsValor, cConsDescripcion from Constante where nConsValor<>nConsCod and nConsCod=9087"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            cboTpoValorRef.AddItem Rs!cConsDescripcion, cboTpoValorRef.ListCount
            cboTpoValorRef.ItemData(cboTpoValorRef.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        If cboTpoValorRef.ListCount > 0 Then cboTpoValorRef.ListIndex = 0
        oCon.CierraConexion
    End If
    Exit Sub
CargarTpoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarCboTpo()
On Error GoTo CargarTpoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nConsValor, cConsDescripcion from Constante where nConsValor<>nConsCod and nConsCod=9088"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            cboTpo.AddItem Rs!cConsDescripcion, cboTpo.ListCount
            cboTpo.ItemData(cboTpo.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        If cboTpo.ListCount > 0 Then cboTpo.ListIndex = 0
        oCon.CierraConexion
    End If
    Exit Sub
CargarTpoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub Inicializar()
    gnProSelNro = 0
    gcBSGrupoCod = 0
End Sub

Private Sub CmdConsultarProceso_Click()
On Error GoTo msflex_clckErr
'    cmdCancelar_Click
    frmLogProSelCnsProcesoSeleccion.Inicio 2
    With frmLogProSelCnsProcesoSeleccion
        If .gbBandera Then
            gnProSelNro = .gvnProSelNro
            gcBSGrupoCod = .gvcBSGrupoCod
            TxtProceso.Text = .gvcTipo
            TxtMonto.Text = Format(.gvnMonto, "###,###.00")
            LblMoneda.Caption = .gvcMoneda
            txtSintesis.Text = .gvcDescripcion
            TxtTipoAnio.Text = .gvnAnio
            
            ListaEtapas gnProSelNro
            GeneraDetalleItem gnProSelNro, gcBSGrupoCod
            CargarCargosComite
            cargarComiteItemProceso gnProSelNro
            consultarProseso gnProSelNro
            MSItem.SetFocus
        End If
    End With
Exit Sub
msflex_clckErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'*********************************************************************************
'*********************************************************************************
'*********************************************************************************

Sub FormaFlexItem()
MSItem.Clear
MSItem.Rows = 2
MSItem.RowHeight(0) = 320
MSItem.RowHeight(1) = 8
MSItem.ColWidth(0) = 300: MSItem.ColAlignment(0) = 4
MSItem.ColWidth(1) = 0:   MSItem.ColAlignment(1) = 4:  MSItem.TextMatrix(0, 1) = " Codigo"
MSItem.ColWidth(2) = 400:   MSItem.ColAlignment(2) = 4:  MSItem.TextMatrix(0, 2) = "Item"
MSItem.ColWidth(3) = 3700:  MSItem.TextMatrix(0, 3) = " Descripción"
MSItem.ColWidth(4) = 700:  MSItem.TextMatrix(0, 4) = " Cantidad"
MSItem.ColWidth(5) = 0:  MSItem.TextMatrix(0, 5) = " nProSelNro"
MSItem.ColWidth(6) = 0:  MSItem.TextMatrix(0, 6) = " nProSelItem"
MSItem.ColWidth(7) = 0:  MSItem.TextMatrix(0, 7) = " nMonto"
MSItem.ColWidth(8) = 1000:  MSItem.TextMatrix(0, 8) = " Unidades"
MSItem.ColWidth(9) = 0:  MSItem.TextMatrix(0, 9) = " Cod Und"
MSItem.ColWidth(10) = 0:  MSItem.TextMatrix(0, 10) = " Cod Catalogo"
MSItem.ColWidth(11) = 1000:  MSItem.TextMatrix(0, 11) = " P. Uni."
MSItem.ColWidth(12) = 1000:  MSItem.TextMatrix(0, 12) = " Precio"
End Sub

Sub FormaFlexEta()
MSEta.Clear
MSEta.Rows = 2
MSEta.RowHeight(0) = 320
MSEta.RowHeight(1) = 10
MSEta.ColWidth(0) = 0
MSEta.ColWidth(1) = 600:   MSEta.TextMatrix(0, 1) = "Orden":       'MSEta.ColAlignment(1) = 4
MSEta.ColWidth(2) = 4300:  MSEta.TextMatrix(0, 2) = "Etapa"
'MSEta.ColWidth(3) = 2400
MSEta.ColWidth(3) = 900:   MSEta.TextMatrix(0, 3) = "Inicio":      MSEta.ColAlignment(3) = 4
MSEta.ColWidth(4) = 900:   MSEta.TextMatrix(0, 4) = "Término":     MSEta.ColAlignment(4) = 4
MSEta.ColWidth(5) = 2000:  MSEta.TextMatrix(0, 5) = "Observación"
MSEta.ColWidth(6) = 0:     MSEta.ColWidth(7) = 0
End Sub

Sub ListaEtapas(nProselNro As Integer)
'Sub ListaEtapas(nPlanNro As Integer, nPlanItem As Integer)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency, sSQL As String

nSuma = 0
FormaFlexEta

If oConn.AbreConexion Then

   'sSQL = "select e.*,e.nEtapaCod,t.cEtapa " & _
   '" From LogProcesoSeleccion p inner join LogProSelEtapa e on p.nProSelNro = e.nProSelNro " & _
   '"     inner join (select nConsValor as nEtapaCod, cConsDescripcion as cEtapa from Constante where nConsCod= " & gcEtapasProcesoSel & " and nConsCod<>nConsValor) t on t.nEtapaCod = e.nEtapaCod " & _
   ' Where P.nPlanAnualNro = " & nPlanNro & " And P.nPlanAnualItem = " & nPlanItem & " "
    
   sSQL = "select e.*, t.cDescripcion cEtapa " & _
            " From LogProSelEtapa e  " & _
            "     inner join LogEtapa t on e.nEtapaCod = t.nEtapaCod and t.nEstado = 1 " & _
            " Where e.nEstado >= 1 and e.nProSelNro = " & nProselNro & "  "
    
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
         i = i + 1
         InsRow MSEta, i
         MSEta.Col = 1
         MSEta.row = i
         If Rs!nEstado >= 1 Then
            Set MSEta.CellPicture = imgOK
         Else
            Set MSEta.CellPicture = imgNN
         End If
         MSEta.TextMatrix(i, 0) = Rs!nEtapaCod
         MSEta.TextMatrix(i, 1) = Rs!nOrden
         MSEta.TextMatrix(i, 2) = Rs!cEtapa
         'MSEta.TextMatrix(i, 3) = rs!cResponsable
         MSEta.TextMatrix(i, 3) = IIf(IsNull(Rs!dFechaInicio), "", Rs!dFechaInicio)
         MSEta.TextMatrix(i, 4) = IIf(IsNull(Rs!dFechaTermino), "", Rs!dFechaTermino)
         MSEta.TextMatrix(i, 5) = Rs!cObservacion
         MSEta.TextMatrix(i, 6) = Rs!nInterno
         MSEta.TextMatrix(i, 7) = Rs!nEstado
         Rs.MoveNext
      Loop
      MSEta.row = 1
      MSEta.ColSel = 7
   End If
End If
End Sub

Sub GeneraDetalleItem(vProSelNro As Integer, vBSGrupoCod As String)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSQL As String, sGrupo As String

sSQL = ""
nSuma = 0
FormaFlexItem

If oConn.AbreConexion Then
          
'    sSQL = "select v.nProSelNro, v.nProSelItem, cBSGrupoDescripcion=isnull(b.cBSGrupoDescripcion,'SIN GRUPO DEFINIDO'), cBSGrupoCod=ISNULL(b.cBSGrupoCod,'0000'),x.cBSCod," & _
            " y.cBSDescripcion, x.nCantidad, x.nMonto, ValorRef=v.nMonto, cUnidades=c.cConsDescripcion, " & _
            " nUnidades=c.nConsValor, y.cCatalogoCod " & _
            "from LogProSelItem v " & _
            "left outer join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
            "inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
            "inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
            "inner join constante c on x.nUnidades = c.nConsValor and c.nConsCod=9089" & _
            "where v.nProSelNro = " & vProSelNro & " order by v.nProSelItem, b.cBSGrupoDescripcion "

    sSQL = "select v.nProSelNro, v.nProSelItem, b.cBSGrupoDescripcion, b.cBSGrupoCod, x.cBSCod," & _
            " y.cBSDescripcion, x.nCantidad, x.nMonto, ValorRef=v.nMonto, cUnidades=c.cConsDescripcion, " & _
            " nUnidades=c.nConsValor, y.cCatalogoCod " & _
            "from LogProSelItem v " & _
            "inner join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
            "inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
            "inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
            "inner join constante c on x.nUnidades = c.nConsValor and c.nConsCod=9089" & _
            "where v.nProSelNro = " & vProSelNro & " order by v.nProSelItem, b.cBSGrupoDescripcion "

   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
        If sGrupo <> Rs!nProSelItem Then
         sGrupo = Rs!nProSelItem
         i = i + 1
         InsRow MSItem, i
         MSItem.Col = 0
         MSItem.row = i
         MSItem.RowHeight(i) = 300
         MSItem.CellFontSize = 10
         MSItem.CellFontBold = True
         MSItem.TextMatrix(i, 0) = "+" ' "-"
         MSItem.TextMatrix(i, 1) = Rs!nProSelItem
         MSItem.TextMatrix(i, 2) = Rs!nProSelItem
         MSItem.TextMatrix(i, 3) = Trim(Rs!cBSGrupoDescripcion)
         MSItem.TextMatrix(i, 4) = ""
         MSItem.TextMatrix(i, 5) = Rs!nProselNro
         MSItem.TextMatrix(i, 6) = Rs!nProSelItem
         MSItem.TextMatrix(i, 12) = FNumero(Rs!valorref)
         MSItem.TextMatrix(i, 8) = ""
         MSItem.TextMatrix(i, 9) = 0
         MSItem.TextMatrix(i, 10) = 0
        End If
        i = i + 1
        InsRow MSItem, i
        MSItem.RowHeight(i) = 0 ' 300
        MSItem.TextMatrix(i, 1) = Trim(Rs!cBSCod)
        MSItem.TextMatrix(i, 3) = Trim(Rs!cBSDescripcion)
        MSItem.TextMatrix(i, 4) = FNumero(Rs!nCantidad)
        MSItem.TextMatrix(i, 5) = Rs!nProselNro
        MSItem.TextMatrix(i, 6) = Rs!nProSelItem
        MSItem.TextMatrix(i, 11) = FNumero(Rs!nMonto)
        MSItem.TextMatrix(i, 8) = Trim(Rs!cUnidades)
        MSItem.TextMatrix(i, 9) = Rs!nunidades
        MSItem.TextMatrix(i, 10) = Trim(Rs!cCatalogoCod)
        MSItem.TextMatrix(i, 12) = FNumero(Val(Rs!nMonto) * Val(Rs!nCantidad))
        Rs.MoveNext
      Loop
      MSItem.row = 1
      MSItem.ColSel = 12
      oConn.CierraConexion
   End If
End If
End Sub

Private Sub CargarCargosComite()
On Error GoTo CargarCargosComiteErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select * from constante where nConsCod=9085 and nConsCod<>nConsValor"
        Set Rs = oCon.CargaRecordSet(sSQL)
        CboCargos.Clear
        Do While Not Rs.EOF
            CboCargos.AddItem Rs!cConsDescripcion, CboCargos.ListCount
            CboCargos.ItemData(CboCargos.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
            CboCargos.ListIndex = 0
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarCargosComiteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cargarComiteItemProceso(ByRef pnProSelNro As Integer)
On Error GoTo cargarComiteItemProcesoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
'        sSQL = "select distinct e.*, p.cPersNombre, c.cConsDescripcion, f.cDescripcion from LogProSelEvaluacionComite e " & _
'                "inner join persona p on p.cPersCod=e.cPersCod " & _
'                "inner join constante c on e.nCargo = c.nConsValor and nConsCod=9085 " & _
'                "inner join LogProSelEvaluacionFactor f on e.nFactorNro=f.nFactorNro " & _
'                "where nProSelNro=" & pnProSelNro & " and nProSelItem=" & pnProSelItem
        sSQL = "select s.*, c.cConsDescripcion,cPersNombre=replace(p.cPersNombre,'/',' ') from LogProSelComite s " & _
                "inner join constante c on s.nCargo=c.nConsValor and c.nConsCod=9085 " & _
                "inner join Persona p on s.cPersCod=p.cPersCod " & _
                "where nProSelNro= " & pnProSelNro & " order by bSuplente"
        Set Rs = oCon.CargaRecordSet(sSQL)
        FormaMSComite
        Do While Not Rs.EOF
            i = i + 1
            InsRow MSComite, i
            MSComite.TextMatrix(i, 0) = Rs!nProselNro
            MSComite.TextMatrix(i, 1) = Rs!cPersCod
            MSComite.TextMatrix(i, 2) = Rs!cConsDescripcion
            MSComite.TextMatrix(i, 3) = Rs!cPersNombre
            MSComite.TextMatrix(i, 4) = IIf(CBool(Rs!bSuplente), "SUPLENTE", "TITULAR")
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
cargarComiteItemProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub FormaMSComite()
    With MSComite
        .Clear
        .Rows = 2
        .RowHeight(0) = 360
        .RowHeight(1) = 8
        .ColWidth(0) = 0:     .TextMatrix(0, 0) = "Proceso"
        .ColWidth(1) = 0:    .TextMatrix(0, 1) = "cPersCod"
        .ColWidth(2) = 3000:    .TextMatrix(0, 2) = "Cargo"
        .ColWidth(3) = 3500:    .TextMatrix(0, 3) = "Nombre"
        .ColWidth(4) = 1500:    .TextMatrix(0, 4) = "Tipo"
        .WordWrap = True
    End With
End Sub

Private Sub cmdAgregar_Click()
'    MsgBox "Debe Ingresar 2 ó 6 ó 10 Miembros del Comite, entre Suplentes y Titulares", vbInformation
    If MSItem.Rows <= 2 Then Exit Sub
    FrameEvalLista.Visible = False
    frameEval.Visible = True
End Sub

Private Sub cmdQuitar_Click()
On Error GoTo cmdQuitarErr
    Dim oCon As DConecta, sSQL As String
    Set oCon = New DConecta
    If Val(MSComite.TextMatrix(MSComite.row, 0)) = 0 Then Exit Sub
    If MsgBox("Seguro que Desea Eliminar... ", vbQuestion + vbYesNo) = vbYes Then
        If oCon.AbreConexion Then
            sSQL = "delete LogProSelEtapaComite where nProSelNro=" & gnProSelNro & " and cPersCod='" & MSComite.TextMatrix(MSComite.row, 1) & "'"
            oCon.Ejecutar sSQL
            sSQL = "delete LogProSelComite where nProSelNro= " & MSComite.TextMatrix(MSComite.row, 0) & " and cPersCod=" & MSComite.TextMatrix(MSComite.row, 1)
            oCon.Ejecutar sSQL
            oCon.CierraConexion
            MsgBox "Miembro del Comite Eliminado", vbInformation
            cargarComiteItemProceso MSItem.TextMatrix(MSItem.row, 5) ', MSItem.TextMatrix(MSItem.Row, 1)
        End If
    End If
    Exit Sub
cmdQuitarErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdGrabarComite_Click()
On Error GoTo cmdGrabarComiteErr
    Dim oCon As DConecta, sSQL As String, pcPersCod As String, pnCargo As Integer, pbSuplente As Integer
    Set oCon = New DConecta
    pcPersCod = txtPersCod.Text
    pnCargo = CboCargos.ItemData(CboCargos.ListIndex)
    pbSuplente = ChkSuplente.value
    If oCon.AbreConexion Then
        sSQL = "insert into LogProSelComite(nProSelNro,cPersCod,nCargo,bSuplente) " & _
                "values(" & gnProSelNro & ",'" & pcPersCod & "'," & pnCargo & "," & pbSuplente & ")"
        oCon.Ejecutar sSQL
        If pbSuplente = 0 Then
            sSQL = "insert into LogProSelEtapaComite (nProSelNro,nEtapaCod,cPersCod)" & _
                    " Select " & gnProSelNro & ",nEtapaCod,'" & pcPersCod & "' from LogProSelEtapa where nProSelNro=" & gnProSelNro
            oCon.Ejecutar sSQL
        End If
        MsgBox "Comite Registrado Correctamente...", vbInformation
        oCon.CierraConexion
        cmdCancelarComite_Click
        cargarComiteItemProceso gnProSelNro
    End If
Exit Sub
cmdGrabarComiteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdCancelarComite_Click()
    FrameEvalLista.Visible = True
    frameEval.Visible = False
    txtPersCod.Text = ""
    txtPersona.Text = ""
    ChkSuplente.value = 0
    CboCargos.ListIndex = 0
    cargarComiteItemProceso gnProSelNro
End Sub

Private Sub cargarFuentesFinanciamiento()
    On Error GoTo cargarFuentesFinanciamientoErr
    Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset
    If oConn.AbreConexion Then
    sSQL = "select * from constante where nConsCod = '9046' and nConsValor<>'9046'"
    CboFinanciamiento.Clear
    Set Rs = oConn.CargaRecordSet(sSQL)
    Do While Not Rs.EOF
        With CboFinanciamiento
            .AddItem Rs!cConsDescripcion, .ListCount
            .ItemData(.ListCount - 1) = Rs!nConsValor
        End With
        Rs.MoveNext
    Loop
    CboFinanciamiento.ListIndex = 0
    oConn.CierraConexion
    End If
    Exit Sub
cargarFuentesFinanciamientoErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarModalidad()
On Error GoTo CargarModalidadErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    sSQL = "select * from Constante where nConsCod =9081 and nConsValor<>nConsCod"
    CboModalidad.Clear
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            CboModalidad.AddItem Rs!cConsDescripcion, CboModalidad.ListCount
            CboModalidad.ItemData(CboModalidad.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
Exit Sub
CargarModalidadErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarObjeto()
On Error GoTo CargarModalidadErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    sSQL = "select * from Constante where nConsCod =9048 and nConsValor<>nConsCod"
    CboObj.Clear
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            CboObj.AddItem Rs!cConsDescripcion, CboObj.ListCount
            CboObj.ItemData(CboObj.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
            CboObj.ListIndex = 0
        Loop
        oCon.CierraConexion
    End If
Exit Sub
CargarModalidadErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarPrePro()
On Error GoTo CargarModalidadErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    sSQL = "select * from Constante where nConsCod =9083 and nConsValor<>nConsCod"
    CboPrePro.Clear
    CboBuenaPro.Clear
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            CboPrePro.AddItem Rs!cConsDescripcion, CboPrePro.ListCount
            CboPrePro.ItemData(CboPrePro.ListCount - 1) = Rs!nConsValor
            CboBuenaPro.AddItem Rs!cConsDescripcion, CboBuenaPro.ListCount
            CboBuenaPro.ItemData(CboBuenaPro.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
Exit Sub
CargarModalidadErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarMedio()
    On Error GoTo CargarMedioErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select * from constante where nConsCod=9082 and nConsCod<>nConsValor"
        Set Rs = oCon.CargaRecordSet(sSQL)
        cboMedio.Clear
        Do While Not Rs.EOF
            cboMedio.AddItem Rs!cConsDescripcion, cboMedio.ListCount
            cboMedio.ItemData(cboMedio.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        cboMedio.ListIndex = 0
        oCon.CierraConexion
    End If
    Exit Sub
CargarMedioErr:
    MsgBox Err.Number & Err.Description
End Sub

Private Sub consultarProseso(pnProSelNro As Integer)
On Error GoTo consultarProsesoErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
'    sSQL = "select distinct p.*, r.cAbreviatura, t.cProSelTpoDescripcion as cProceso, g.cBSGrupoDescripcion, " & _
            "              p.nPlanAnualAnio,p.cSintesis, p.nPlanAnualMes,p.cBSGrupoCod " & _
            "  from LogProcesoSeleccion p " & _
            "       inner join LogProSelTpo t on p.nProSelTpoCod = t.nProSelTpoCod " & _
            "       inner join LogProSelTpoRangos r on p.nProSelTpoCod=r.nProSelTpoCod " & _
            "       left outer join BSGrupos g on g.cBSGrupoCod=p.cBSGrupoCod" & _
            " Where p.nProSelNro = " & pnProSelNro
    sSQL = "select distinct p.nProSelNro, p.nPlanAnualNro, p.nProSelTpoCod, p.nProSelSubTpo, p.dProSelFecha, p.nProSelMonto, p.bPresupuesto, p.nProSelEstado, p.nMoneda, p.cArchivoBases, p.nModalidad, p.nNroProceso, p.nPresentaPropuestaTpo, p.nBuenaProTpo, p.dFechaPublicacion, p.dFechaPROMPYME, p.nFuenteFinanciemiento, p.nObjetoCod, p.nMedioPublicacion, p.nPlanAnualMes, p.nPlanAnualAnio, p.nModalidadCompra, p.nContratoCompra, p.nPuntajeMinimoTecnico, p.nRangoMenor, p.nRangoMayor, p.cSintesis, p.nTipoCambio,p.nMonedaCostoBases, p.nCostoBases, r.cAbreviatura, t.cProSelTpoDescripcion as cProceso, cDirVentaBases, cNroCuenta, cUbiGeo, u.cUbigeoDescripcion, cCIIU, c.cCIIUDescripcion, cObservaciones, cVersionCONSUCODE " & _
            "  from LogProcesoSeleccion p " & _
            "       inner join LogProSelTpo t on p.nProSelTpoCod = t.nProSelTpoCod " & _
            "       inner join LogProSelTpoRangos r on p.nProSelTpoCod=r.nProSelTpoCod " & _
            "       left join LogProSelUbigeo u on p.cUbiGeo = u.cUbigeoCod " & _
            "       left join LogProSelCIIU c on p.cCIIU = c.cCIIUCod" & _
            " Where p.nProSelNro = " & pnProSelNro
            
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            txtSintesis.Text = Rs!cSintesis
            TxtTipo.Text = Rs!cAbreviatura
            TxtTipoNro.Text = pnProSelNro      'Rs!nNroProceso
            CboModalidad.ListIndex = EncuentraIndiceCbo(CboModalidad, Rs!nModalidad)
            CboBuenaPro.ListIndex = EncuentraIndiceCbo(CboBuenaPro, Rs!nBuenaProtpo)
            CboFinanciamiento.ListIndex = EncuentraIndiceCbo(CboFinanciamiento, Rs!nFuenteFinanciemiento)
            cboMedio.ListIndex = EncuentraIndiceCbo(cboMedio, Rs!nMediopublicacion)
            CboObj.ListIndex = IIf(Rs!nObjetoCod = 0, 0, Rs!nObjetoCod - 1)
            CboPrePro.ListIndex = EncuentraIndiceCbo(CboPrePro, Rs!nPresentaPropuestaTpo)
            CboBuenaPro.ListIndex = EncuentraIndiceCbo(CboBuenaPro, Rs!nBuenaProtpo)
            MEBcomunicacion.Text = IIf(IsNull(Rs!dFechaPublicacion), "__/__/____", Format(Rs!dFechaPublicacion, "dd/mm/yyyy"))
            MEBFechaPROMPYME.Text = IIf(IsNull(Rs!dFechaPROMPYME), "__/__/____", Format(Rs!dFechaPROMPYME, "dd/mm/yyyy"))
            txtmaximo.Text = Rs!nRangoMayor
            txtminimo.Text = Rs!nRangoMenor
            txtPuntajeMinimo.Text = Rs!nPuntajeMinimoTecnico
            txtCostoBases.Text = Rs!nCostoBases
'            TxtGrupo.Text = IIf(IsNull(Rs!cBSGrupoDescripcion), "", Rs!cBSGrupoDescripcion)
            OptContrato.value = CBool(Rs!nContratoCompra)
'            OptContrato1.value = Not CBool(rs!nContratoCompra)
            OptModalidad.value = CBool(Rs!nModalidadCompra)
            OptModalidad1.value = Not CBool(Rs!nModalidadCompra)
            cboMonedaCosto.ListIndex = EncuentraIndiceCbo(cboMonedaCosto, Rs!nMonedaCostoBases)
            txtDirVentaBases.Text = Rs!cDirVentaBases
            TxtVersion.Text = IIf(Rs!cVersionCONSUCODE = "", "1.2.1", Rs!cVersionCONSUCODE)
            txtCuenta.Text = Rs!cNroCuenta
            txtUbigeoCod = Rs!cUbiGeo
            txtUbigeo = IIf(IsNull(Rs!cUbigeoDescripcion), "", Rs!cUbigeoDescripcion)
            If Not IsNull(Rs!cCIIUDescripcion) Then cboCIIU.Text = Rs!cCIIUDescripcion & Space(200) & Rs!cCIIU
            txtObs.Text = Rs!cObservaciones
        End If
        oCon.CierraConexion
    End If
Exit Sub
consultarProsesoErr:
    MsgBox Err.Description & vbCrLf & Err.Description
End Sub

Private Function EncuentraIndiceCbo(ByVal cbo As ComboBox, ByVal valor As Integer) As Integer
On Error GoTo EncuentraIndiceCbo
    Dim i As Integer
    Do While i < cbo.ListCount
        If cbo.ItemData(i) = valor Then
            EncuentraIndiceCbo = i
            Exit Function
        End If
        i = i + 1
    Loop
    Exit Function
EncuentraIndiceCbo:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelGenerarProcesoSeleccion = Nothing
End Sub


Private Sub MEBcomunicacion_GotFocus()
SelTexto MEBcomunicacion
End Sub

Private Sub MEBcomunicacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   MEBFechaPROMPYME.SetFocus
End If
End Sub

Private Sub MEBFechaPROMPYME_GotFocus()
SelTexto MEBFechaPROMPYME
End Sub

Private Sub MEBFechaPROMPYME_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboMedio.SetFocus
End If
End Sub

Private Sub MSEta_DblClick()
    Dim nCol As Integer, nRow As Integer
    With MSEta
        nCol = .Col
        nRow = .row
        .Col = 1
        If .CellPicture = imgOK Then
            Set .CellPicture = imgNN
        Else
            Set .CellPicture = imgOK
'            cmdEtapas_Click
        End If
        .row = nRow
        .ColSel = 7
    End With
End Sub

Private Sub MSFlex_DblClick()
    Dim nCol As Integer
    With MSFlex
        nCol = .Col
        .Col = 0
        If .CellPicture = imgOK Then
            If .TextMatrix(.row, 4) = "Tecnica" Then Set .CellPicture = imgNN
        Else
            Set .CellPicture = imgOK
        End If
        .Col = nCol
        CalcularTotalPuntos
    End With
End Sub

Private Sub MSFlex_GotFocus()
If txtEditItem.Visible = False Then Exit Sub
MSFlex = txtEditItem
txtEditItem.Visible = False
CalcularTotalPuntos
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
On Error GoTo MSItemErr
    Dim i As Integer
    Select Case MSFlex.Col
        Case 3
            If IsNumeric(Chr(KeyAscii)) Then _
                EditaFlex MSFlex, txtEditItem, KeyAscii
    End Select
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSFlex_LeaveCell()
If txtEditItem.Visible = False Then Exit Sub
MSFlex = txtEditItem
txtEditItem.Visible = False
CalcularTotalPuntos
End Sub

Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
Select Case KeyAscii
    Case 0 To 32
         Edt = MSFlex
         Edt.SelStart = 1000
    Case Else
         Edt = Chr(KeyAscii)
         Edt.SelStart = 1
End Select
Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
         MSFlex.CellWidth, MSFlex.CellHeight
'Edt.Text = Chr(KeyAscii) ' & MSFlex
Edt.Visible = True
Edt.SetFocus
End Sub

Private Sub MSFlexValoresVer_GotFocus()
If txtEditRangos.Visible = False Then Exit Sub
If ValidarRangosValores(txtEditRangos.Text) Then
    MSFlexValoresVer = txtEditRangos
    txtEditRangos.Visible = False
Else
    txtEditRangos.Visible = False
End If
End Sub

Private Sub MSFlexValoresVer_KeyPress(KeyAscii As Integer)
On Error GoTo MSItemErr
    Dim i As Integer
    'Select Case MSFlexValoresVer.Col
    '    Case 3
            If IsNumeric(Chr(KeyAscii)) Then _
                EditaFlex MSFlexValoresVer, txtEditRangos, KeyAscii
    'End Select
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSFlexValoresVer_LeaveCell()
If txtEditRangos.Visible = False Then Exit Sub
If ValidarRangosValores(txtEditRangos.Text) Then
    MSFlexValoresVer = txtEditRangos
    txtEditRangos.Visible = False
Else
    txtEditRangos.Visible = False
End If
End Sub

Private Sub MSItem_Click()
    cbound.Visible = False
End Sub

Private Sub MSItem_LeaveCell()
    If txtEdit.Visible Then
        MSItem = FNumero(txtEdit)
        MSItem.TextMatrix(MSItem.row, 12) = FNumero(CDbl(MSItem.TextMatrix(MSItem.row, 4)) * CDbl(MSItem.TextMatrix(MSItem.row, 11)))
        CalculaValorRefxItem MSItem.TextMatrix(MSItem.row, 6)
        txtEdit.Visible = False
    End If
    If cbound.Visible Then
       MSItem.TextMatrix(MSItem.row, 8) = cbound.Text
       MSItem.TextMatrix(MSItem.row, 9) = cbound.ItemData(cbound.ListIndex)
       cbound.Visible = False
    End If
End Sub

Private Sub MSItem_SelChange()
    cbound.Visible = False
    CargarFactores gnProSelNro, Val(MSItem.TextMatrix(MSItem.row, 6))
End Sub

Private Sub txtCostoBases_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumDec(txtCostoBases, KeyAscii)
If nKeyAscii = 13 Then
   txtminimo.SetFocus
End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
     KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode MSItem, txtEdit, KeyCode, Shift
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
       KeyAscii = 0
    End If
    Select Case MSItem.Col
        Case 4, 11
            KeyAscii = DigNumDec(txtEdit, KeyAscii)
    End Select
End Sub

Private Sub txtEditItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
       KeyAscii = 0
    End If
Select Case MSItem.Col
    Case 3
        KeyAscii = DigNumEnt(KeyAscii)
End Select
End Sub

Private Sub txtEdititem_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlex, txtEditItem, KeyCode, Shift
End Sub

Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         Edt.Visible = False
         MSFlex.SetFocus
    Case 13
         MSFlex.SetFocus
    Case 37                     'Izquierda
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col > 1 Then
            MSFlex.Col = MSFlex.Col - 1
         End If
    Case 39                     'Derecha
         MSFlex.SetFocus
         DoEvents
         If MSFlex.Col < MSFlex.Cols - 1 Then
            MSFlex.Col = MSFlex.Col + 1
         End If
    Case 38
         MSFlex.SetFocus
         DoEvents
         If MSFlex.row > MSFlex.FixedRows + 1 Then
            MSFlex.row = MSFlex.row - 1
         End If
    Case 40
         MSFlex.SetFocus
         DoEvents
         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
         If MSFlex.row < MSFlex.Rows - 1 Then
            MSFlex.row = MSFlex.row + 1
         End If
End Select
End Sub

Private Sub MSItem_DblClick()
On Error GoTo MSItemErr
    Dim i As Integer, bTipo As Boolean
    With MSItem
        If Trim(.TextMatrix(.row, 0)) = "-" Then
           .TextMatrix(.row, 0) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
           .TextMatrix(.row, 0) = "-"
           i = .row + 1
           bTipo = False
        Else
            If .Col = 8 Then
                cbound.Visible = True
                cbound.Text = .TextMatrix(.row, .Col)
                cbound.Move .Left + .CellLeft - 30, .Top + .CellTop - 30, .CellWidth + 30
            End If
            Exit Sub
        End If
        Do While i < .Rows
            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 300
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation

End Sub

Private Sub MSItem_GotFocus()
    If txtEdit.Visible Then
        MSItem = FNumero(txtEdit)
        MSItem.TextMatrix(MSItem.row, 12) = FNumero(CDbl(MSItem.TextMatrix(MSItem.row, 4)) * CDbl(MSItem.TextMatrix(MSItem.row, 11)))
        CalculaValorRefxItem MSItem.TextMatrix(MSItem.row, 6)
        txtEdit.Visible = False
    End If
    If cbound.Visible Then
       MSItem.TextMatrix(MSItem.row, 8) = cbound.Text
       MSItem.TextMatrix(MSItem.row, 9) = cbound.ItemData(cbound.ListIndex)
       cbound.Visible = False
    End If
    If gnProSelNro = 0 Then Exit Sub
    CargarFactores gnProSelNro, Val(MSItem.TextMatrix(MSItem.row, 6))
End Sub

Private Sub MSItem_KeyPress(KeyAscii As Integer)
    On Error GoTo MSItemErr
    Dim i As Integer
    Select Case MSItem.Col
        Case 4, 11
            If IsNumeric(Chr(KeyAscii)) Then _
                EditaFlex MSItem, txtEdit, KeyAscii
    End Select
    
    If KeyAscii = 13 Then
        MSItem_DblClick
    ElseIf KeyAscii = 27 Then
        cbound.Visible = False
    End If
    
    Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub OptContrato_Click()
    If Not OptModalidad1.value Then
        OptContrato.value = True
    Else
'        OptContrato.value = False
        OptContrato1.value = True
    End If
End Sub

Private Sub OptModalidad1_Click()
    OptContrato1.value = True
End Sub

Private Sub sstPro_Click(PreviousTab As Integer)
    frameEval.Visible = False
    FrameEvalLista.Visible = True
End Sub

Private Sub txtEditRangos_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode MSFlexValoresVer, txtEditRangos, KeyCode, Shift
End Sub

Private Sub txtEditRangos_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then
       KeyAscii = 0
    End If
'Select Case MSFlexValoresVer.Col
'    Case 3
        KeyAscii = DigNumEnt(KeyAscii)
'End Select
End Sub

Private Sub txtmaximo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtmaximo.Text) < 70 Then
            txtmaximo.Text = 70
        ElseIf Val(txtmaximo.Text) > 110 Then
            txtmaximo.Text = 110
        End If
        txtPuntajeMinimo.SetFocus
    End If
End Sub

Private Sub txtminimo_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtminimo.Text) < 70 Then
            txtminimo.Text = 70
        ElseIf Val(txtminimo.Text > 110) Then
            txtminimo.Text = 110
        End If
        txtmaximo.SetFocus
    End If
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtPuntajeMinimo_KeyPress(KeyAscii As Integer)
nKeyAscii = KeyAscii
KeyAscii = DigNumEnt(KeyAscii)
If nKeyAscii = 13 Then
   cmdGuardar.SetFocus
End If
End Sub

Private Sub txtSintesis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtTipoNro_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub CargarFactoresEvaluacion(ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer)
On Error GoTo CargarFactoresEvaluacionErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    sSQL = "select cBSGrupoCod, x.nFactorNro, x.cFactorDescripcion, nPuntaje, x.cUnidades, x.nTipo " & _
         "from LogProSelEvalFactor f " & _
         "inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
         "Where nVigente = 1 And nProSelTpoCod = " & pnProSelTpoCod & " And nProSelSubTpo = " & pnProSelTpoCod & _
         "order by cBSGrupoCod, x.nFactornro "
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarFactoresEvaluacionErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub FormaFlexFactores()
With MSFlex
    .Clear
    .Rows = 2
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 800:     .TextMatrix(0, 0) = "Codigo":
    .ColWidth(1) = 4000:    .TextMatrix(0, 1) = "Factor"
    .ColWidth(2) = 1700:     .TextMatrix(0, 2) = "Unidades":
    .ColWidth(3) = 800:     .TextMatrix(0, 3) = "Puntos":       .ColAlignment(3) = 4
    .ColWidth(4) = 1200:     .TextMatrix(0, 4) = "Propuesta":
    .ColWidth(5) = 0:       .TextMatrix(0, 5) = "Formula"
End With
End Sub

Private Sub CargarFactores(ByVal pnProSelNro As Integer, ByVal pnProSelItem As Integer)
    On Error GoTo CargarFactoresErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String, i As Integer
    sSQL = "select f.nVigente, x.nFactorNro, x.cFactorDescripcion, x.cUnidades, x.nTipo, f.nPuntaje, nFormula from LogProSelEvalFactor f " & _
        "inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
        "where nProSelNro = " & pnProSelNro & " and nProSelItem = " & pnProSelItem
    
    Set oCon = New DConecta
    FormaFlexFactores
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Rs.EOF Then
            MsgBox "No Existen Factores de Evaluacion Registrados para el Proceso de Seleccion", vbInformation, "Aviso"
            Exit Sub
        End If
        Do While Not Rs.EOF
            With MSFlex
                i = i + 1
                InsRow MSFlex, i
                .TextMatrix(i, 0) = Rs!nFactorNro
                .TextMatrix(i, 1) = Rs!cFactorDescripcion
                .row = i
                .Col = 0
                .CellPictureAlignment = 1
                If Rs!nVigente Then
                    Set .CellPicture = imgOK
                Else
                    Set .CellPicture = imgNN
                End If
                .row = 1
'                If rs!nFormula = 3 Then
'                    .Row = i
'                    .Col = 2
'                    .CellPictureAlignment = 4
'                    Set .CellPicture = imgOK
'                    .Row = 1
'                End If
                .TextMatrix(i, 2) = Rs!cUnidades
                .TextMatrix(i, 3) = Rs!npuntaje
                .TextMatrix(i, 4) = IIf(Rs!nTipo = 0, "Tecnica", "Economica")
                .TextMatrix(i, 5) = Rs!nFormula
            End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    CalcularTotalPuntos
    Exit Sub
CargarFactoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CalcularTotalPuntos()
On Error GoTo CalcularTotalPuntosErr
    Dim i As Integer, nTotal As Integer, nRow As Integer, nCol As Integer
    With MSFlex
        nRow = .row
        nCol = .Col
        .Col = 0
        Do While i < .Rows
            .row = i
            If .CellPicture = imgOK Then
                nTotal = nTotal + Val(.TextMatrix(i, 3))
            End If
            i = i + 1
        Loop
        .row = nRow
        .Col = nCol
    End With
    If nTotal <> 100 And gnProSelNro <> 0 Then
        lblmensaje.Visible = True
        MsgBox "El Puntaje Total debe ser 100 ", vbInformation, "Aviso"
    Else
        lblmensaje.Visible = False
    End If
    txtTotalPuntos = nTotal
    Exit Sub
CalcularTotalPuntosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub CargarFactoresVer(pnFactorNro As Integer, pnProSelNro As Integer)
    On Error GoTo CargarFactoresErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    FormaFlexValoresVer
    sSQL = "select * from LogProSelEvalFactorRangos where nFactorNro=" & pnFactorNro & " and  nProSelNro = " & pnProSelNro
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            With MSFlexValoresVer
                i = i + 1
                InsRow MSFlexValoresVer, i
                .TextMatrix(i, 0) = Rs!nRangoItem
                .TextMatrix(i, 1) = Rs!nRangoMin
                .TextMatrix(i, 2) = Rs!nRangoMax
                .TextMatrix(i, 3) = Rs!npuntaje
            End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarFactoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub FormaFlexValoresVer()
With MSFlexValoresVer
    .Clear
    .Rows = 2
    .RowHeight(0) = 320
    .RowHeight(1) = 10
    .ColWidth(0) = 0:     .TextMatrix(0, 0) = "Item":        '.ColAlignment(1) = 4
    .ColWidth(1) = 1300:    .TextMatrix(0, 1) = "Minimo"
    .ColWidth(2) = 1300:     .TextMatrix(0, 2) = "Maximo":        '.ColAlignment(2) = 4
    .ColWidth(3) = 1300:     .TextMatrix(0, 3) = "Puntaje":        '.ColAlignment(2) = 4
End With
End Sub

Private Function ValidarRangosValorMax() As Boolean
    On Error GoTo ValidarRangosErr
    Dim npuntaje As Integer, i As Integer
    With MSFlexValoresVer
        i = 1
        ValidarRangosValorMax = True
        Do While i < .Rows
            If Val(.TextMatrix(i, 3)) > Val(txtvalorMaximo.Text) Then
                ValidarRangosValorMax = False
                Exit Function
            End If
            i = i + 1
        Loop
    End With
    Exit Function
ValidarRangosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Function ValidarRangosValores(ByVal valor As Integer) As Boolean
    On Error GoTo ValidarRangosErr
    Dim npuntaje As Integer, i As Integer
    With MSFlexValoresVer
        i = 1
        ValidarRangosValores = True
        Do While i < .Rows
            If Val(.TextMatrix(i, 3)) = Val(valor) Then
                ValidarRangosValores = False
                Exit Function
            End If
            i = i + 1
        Loop
    End With
    Exit Function
ValidarRangosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Function VerificaFechas() As Boolean
On Error GoTo VerificaFechasErr
    Dim i As Integer, dFechaAnterior As String
    With MSEta
        dFechaAnterior = .TextMatrix(1, 4)
        i = 2
        Do While i < .Rows
            If .TextMatrix(i, 3) = "" And .TextMatrix(i, 4) = "" Then
                VerificaFechas = False
                Exit Function
            'ElseIf dFechaAnterior < .TextMatrix(i, 4) Then
            '    VerificaFechas = False
            '    Exit Function
            End If
            i = i + 1
        Loop
        VerificaFechas = True
    End With
    Exit Function
VerificaFechasErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function FechaEtapa(ByVal pnEtapa As Integer, Optional ByVal pbInicio As Boolean = True) As String
On Error GoTo FechaEtapaErr
    Dim i As Integer
    With MSEta
        i = 1
        Do While i < .Rows
            If Val(.TextMatrix(i, 1)) = pnEtapa Then
                If pbInicio Then
                    FechaEtapa = .TextMatrix(i, 3)
                Else
                    FechaEtapa = .TextMatrix(i, 4)
                End If
            End If
            i = i + 1
        Loop
    End With
    Exit Function
FechaEtapaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Sub actualizaUnidades()
On Error GoTo actualizaUnidadesErr
    Dim oCon As DConecta, sSQL As String, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        i = 1
        Do While i < MSItem.Rows
            If MSItem.TextMatrix(i, 0) <> "+" And MSItem.TextMatrix(i, 0) <> "-" Then
                sSQL = "update LogProSelItemBS set nUnidades=" & MSItem.TextMatrix(i, 9) & "," & _
                        " nCantidad= " & CDbl(MSItem.TextMatrix(i, 4)) & ", " & _
                        " nMonto= " & CDbl(MSItem.TextMatrix(i, 11)) & _
                        " where nProSelNro=" & MSItem.TextMatrix(i, 5) & " and nProSelItem=" & MSItem.TextMatrix(i, 6) & _
                        " and cBSCod='" & MSItem.TextMatrix(i, 1) & "'"
                oCon.Ejecutar sSQL
            End If
            i = i + 1
        Loop
        
        i = 1
        Do While i < MSItem.Rows
            If MSItem.TextMatrix(i, 0) = "+" Or MSItem.TextMatrix(i, 0) = "-" Then
                sSQL = "update LogProSelItem set " & _
                        " nMonto= " & CDbl(MSItem.TextMatrix(i, 12)) & _
                        " where nProSelNro=" & MSItem.TextMatrix(i, 5) & " and nProSelItem=" & MSItem.TextMatrix(i, 6)
                oCon.Ejecutar sSQL
            End If
            i = i + 1
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
actualizaUnidadesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CalculaValorRefxItem(ByVal pnProSelItem As Integer)
On Error GoTo CalculaValorRefxItem
    Dim i As Integer, nValorRef As Currency, nMontoAnterior As Currency
    i = 1
    nMontoAnterior = gnMonto
    With MSItem
        Do While i < .Rows
            If .TextMatrix(i, 6) = pnProSelItem And (.TextMatrix(i, 0) <> "+" And .TextMatrix(i, 0) <> "-") Then
                nValorRef = nValorRef + CDbl(.TextMatrix(i, 12))
            End If
            i = i + 1
        Loop
    
        i = 1
        Do While i < .Rows
            If .TextMatrix(i, 6) = pnProSelItem And (.TextMatrix(i, 0) = "+" Or .TextMatrix(i, 0) = "-") Then
                .TextMatrix(i, 12) = FNumero(nValorRef)
            End If
            i = i + 1
        Loop
    
        i = 1
        nValorRef = 0
        Do While i < .Rows
            If .TextMatrix(i, 0) = "+" Or .TextMatrix(i, 0) = "-" Then
                nValorRef = nValorRef + .TextMatrix(i, 12)
            End If
            i = i + 1
        Loop
        TxtMonto.Text = FNumero(nValorRef)
        If ValidaRango(gnProSelTpoCod, gnProSelSubTpo, nValorRef, CboObj.ItemData(CboObj.ListIndex), IIf(LblMoneda.Caption = "$", 2, 1)) And _
            nValorRef >= nMontoAnterior * 0.75 And nValorRef <= nMontoAnterior * 1.25 Then
            'MsgBox "El Tipo de Proceso Correcto",vbInformation, "Aviso"
            cmdGuardar.Enabled = True
        Else
            MsgBox "El Tipo de Proceso debe Cambiar", vbInformation, "Aviso"
            cmdGuardar.Enabled = False
        End If
    End With
    Exit Sub
CalculaValorRefxItem:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub CargarCIIU()
    On Error GoTo CargarCIIUErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select cCIIUCod, cCIIUDescripcion from LogProSelCIIU " 'where cCIIUDescripcion like " & pcDescripcion & "'%'"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            cboCIIU.AddItem Rs!cCIIUDescripcion & Space(200) & Rs!cCIIUCod
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    cboCIIU.ListIndex = 0
    Exit Sub
CargarCIIUErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub TxtVersion_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
End Sub
