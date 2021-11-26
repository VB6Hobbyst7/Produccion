VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Credito"
   ClientHeight    =   6315
   ClientLeft      =   615
   ClientTop       =   1545
   ClientWidth     =   10920
   Icon            =   "frmCredConsulta.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame10 
      Height          =   2040
      Left            =   105
      TabIndex        =   137
      Top             =   30
      Width           =   10635
      Begin Sicmact.ActXCodCta ActxCta 
         Height          =   420
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   741
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin MSComctlLib.ListView listaClientes 
         Height          =   1650
         Left            =   5400
         TabIndex        =   138
         Top             =   225
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   2910
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre de Persona"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Relación"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "cCodCli"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado Actual :"
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
         Index           =   2
         Left            =   150
         TabIndex        =   142
         Top             =   1515
         Width           =   1485
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   1680
         TabIndex        =   141
         Top             =   1455
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Crédito :"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   140
         Top             =   780
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbltipoCredito 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   150
         TabIndex        =   139
         Top             =   1035
         Width           =   5040
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
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
      Left            =   9540
      TabIndex        =   3
      Top             =   5805
      Width           =   1215
   End
   Begin VB.CommandButton CmdNuevaCons 
      Caption         =   "&Nueva Consulta"
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
      Left            =   6450
      TabIndex        =   1
      Top             =   5805
      Width           =   1845
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Left            =   8325
      TabIndex        =   2
      Top             =   5805
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3570
      Left            =   45
      TabIndex        =   4
      Top             =   2190
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6297
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "frmCredConsulta.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Historial "
      TabPicture(1)   =   "frmCredConsulta.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Desembolsos"
      TabPicture(2)   =   "frmCredConsulta.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pagos &Realizados"
      TabPicture(3)   =   "frmCredConsulta.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Pagos &Pendientes"
      TabPicture(4)   =   "frmCredConsulta.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame6"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame7"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&Garantías"
      TabPicture(5)   =   "frmCredConsulta.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fragarantias"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Otros Datos"
      TabPicture(6)   =   "frmCredConsulta.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2(3)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   2670
         Left            =   -74790
         TabIndex        =   109
         Top             =   585
         Width           =   10365
         Begin VB.Label lblMontoSolicitado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   143
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Solicitado :"
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
            Index           =   17
            Left            =   465
            TabIndex        =   136
            Top             =   585
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Aprobado :"
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
            Index           =   18
            Left            =   450
            TabIndex        =   135
            Top             =   1575
            Width           =   945
         End
         Begin VB.Label lblMontoAprobado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   134
            Top             =   1545
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sugerido  :"
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
            Index           =   19
            Left            =   465
            TabIndex        =   133
            Top             =   1095
            Width           =   945
         End
         Begin VB.Label lblMontosugerido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   132
            Top             =   1065
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto "
            Height          =   195
            Index           =   20
            Left            =   3240
            TabIndex        =   131
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblcuotasSolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   130
            Top             =   570
            Width           =   675
         End
         Begin VB.Label lblcuotasAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   129
            Top             =   1545
            Width           =   675
         End
         Begin VB.Label lblcuotasugerida 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   128
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cuotas"
            Height          =   195
            Index           =   21
            Left            =   4200
            TabIndex        =   127
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblPlazoSolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   126
            Top             =   570
            Width           =   735
         End
         Begin VB.Label lblPlazoAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   125
            Top             =   1545
            Width           =   735
         End
         Begin VB.Label lblPlazoSugerido 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   124
            Top             =   1065
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
            Height          =   195
            Index           =   22
            Left            =   5055
            TabIndex        =   123
            Top             =   255
            Width           =   390
         End
         Begin VB.Label lblmontoCuotaAprobada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5655
            TabIndex        =   122
            Top             =   1545
            Width           =   1035
         End
         Begin VB.Label lblmontoCuotaSugerida 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5655
            TabIndex        =   121
            Top             =   1050
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            Height          =   195
            Index           =   23
            Left            =   5895
            TabIndex        =   120
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblfechsolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   119
            Top             =   570
            Width           =   1365
         End
         Begin VB.Label lblfechaAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   118
            Top             =   1545
            Width           =   1365
         End
         Begin VB.Label lblfechasugerida 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   117
            Top             =   1065
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
            Height          =   195
            Index           =   24
            Left            =   1860
            TabIndex        =   116
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblGraciaAprobada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6705
            TabIndex        =   115
            Top             =   1545
            Width           =   615
         End
         Begin VB.Label lblGraciasugerida 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6705
            TabIndex        =   114
            Top             =   1050
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Periodo Gracia"
            Height          =   390
            Index           =   25
            Left            =   6615
            TabIndex        =   113
            Top             =   240
            Width           =   765
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTipoGraciaApr 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   8415
            TabIndex        =   112
            Top             =   1545
            Width           =   1155
         End
         Begin VB.Label lblIntGraciaApr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7425
            TabIndex        =   111
            Top             =   1545
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Gracia Aprobada"
            Height          =   390
            Index           =   61
            Left            =   7965
            TabIndex        =   110
            Top             =   255
            Width           =   1155
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Cuotas Pendiente:"
         Height          =   3045
         Left            =   -74880
         TabIndex        =   96
         Top             =   450
         Width           =   6855
         Begin VB.Frame Frame9 
            Caption         =   "Cuotas Atrasadas "
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
            Height          =   795
            Left            =   90
            TabIndex        =   97
            Top             =   2175
            Width           =   6660
            Begin VB.Label lblCapitalCuoPend 
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
               Height          =   330
               Left            =   150
               TabIndex        =   107
               Top             =   375
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capital"
               Height          =   195
               Index           =   54
               Left            =   150
               TabIndex        =   106
               Top             =   180
               Width           =   480
            End
            Begin VB.Label lblGastoCuoPend 
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
               Height          =   330
               Left            =   2940
               TabIndex        =   105
               Top             =   375
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mora "
               Height          =   195
               Index           =   55
               Left            =   2055
               TabIndex        =   104
               Top             =   165
               Width           =   405
            End
            Begin VB.Label lblMoraCuoPend 
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
               Height          =   330
               Left            =   2025
               TabIndex        =   103
               Top             =   375
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gasto"
               Height          =   195
               Index           =   56
               Left            =   2970
               TabIndex        =   102
               Top             =   165
               Width           =   420
            End
            Begin VB.Label lblTotalCuoPend 
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
               ForeColor       =   &H8000000D&
               Height          =   330
               Left            =   4155
               TabIndex        =   101
               Top             =   375
               Width           =   1065
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total "
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
               Index           =   52
               Left            =   4155
               TabIndex        =   100
               Top             =   165
               Width           =   510
            End
            Begin VB.Label lblInteresCuoPend 
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
               Height          =   330
               Left            =   1095
               TabIndex        =   99
               Top             =   375
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Interes "
               Height          =   195
               Index           =   60
               Left            =   1110
               TabIndex        =   98
               Top             =   165
               Width           =   525
            End
         End
         Begin MSComctlLib.ListView lstCuotasPend 
            Height          =   1965
            Left            =   120
            TabIndex        =   108
            Top             =   210
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   3466
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
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cuota"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fec. Venc."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Capital"
               Object.Width           =   1941
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Interes"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Mora"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Atraso"
               Object.Width           =   1323
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Index           =   2
         Left            =   -74880
         TabIndex        =   85
         Top             =   450
         Width           =   10335
         Begin VB.Frame Frame5 
            Caption         =   "Total Pagado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2070
            Left            =   7905
            TabIndex        =   86
            Top             =   285
            Width           =   2010
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gasto :"
               Height          =   195
               Index           =   33
               Left            =   210
               TabIndex        =   94
               Top             =   1140
               Width           =   510
            End
            Begin VB.Label lblGastopagado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   795
               TabIndex        =   93
               Top             =   1110
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mora :"
               Height          =   195
               Index           =   32
               Left            =   225
               TabIndex        =   92
               Top             =   1470
               Width           =   450
            End
            Begin VB.Label lblIntMorPag 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   91
               Top             =   1425
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Interés :"
               Height          =   195
               Index           =   31
               Left            =   195
               TabIndex        =   90
               Top             =   825
               Width           =   570
            End
            Begin VB.Label lblintcompPag 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   89
               Top             =   780
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capital :"
               Height          =   195
               Index           =   30
               Left            =   195
               TabIndex        =   88
               Top             =   480
               Width           =   570
            End
            Begin VB.Label lblcapitalpagado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   87
               Top             =   450
               Width           =   960
            End
         End
         Begin MSComctlLib.ListView ListaPagos 
            Height          =   2370
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   4180
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Fecha Pago"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Nro Cuota"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Total Pagado"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Capital "
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Interés"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Mora"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Gastos "
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Atraso"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Saldo Cap"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2610
         Index           =   1
         Left            =   -74880
         TabIndex        =   77
         Top             =   450
         Width           =   10230
         Begin MSComctlLib.ListView listaDesembolsos 
            Height          =   2175
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   3836
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
               Text            =   "Fecha Desemb."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Nº Desemb."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Monto"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Gastos"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Estado"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Desembolso :"
            Height          =   210
            Index           =   50
            Left            =   7170
            TabIndex        =   84
            Top             =   1095
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbltotalDesembolso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   83
            Top             =   1044
            Width           =   1155
         End
         Begin VB.Label lblmontoDesembolsado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   82
            Top             =   672
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desembolsado :"
            Height          =   195
            Index           =   28
            Left            =   7170
            TabIndex        =   81
            Top             =   735
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbltipoDesembolso 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   80
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Desembolso :"
            Height          =   210
            Index           =   26
            Left            =   7170
            TabIndex        =   79
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2490
         Left            =   300
         TabIndex        =   55
         Top             =   600
         Width           =   10095
         Begin VB.Label lblfechavigencia 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   76
            Top             =   1170
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vigencia :"
            Height          =   195
            Index           =   15
            Left            =   6390
            TabIndex        =   75
            Top             =   1200
            Width           =   1425
         End
         Begin VB.Label lbltipocuota 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   74
            Top             =   825
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cuota :"
            Height          =   195
            Index           =   13
            Left            =   6375
            TabIndex        =   73
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label lbldestino 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   72
            Top             =   1875
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Destino del Crédito :"
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   71
            Top             =   1905
            Width           =   1425
         End
         Begin VB.Label lblcondicion 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   70
            Top             =   1545
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Crédito :"
            Height          =   195
            Index           =   11
            Left            =   105
            TabIndex        =   69
            Top             =   1560
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   10
            Left            =   8625
            TabIndex        =   68
            Top             =   555
            Width           =   120
         End
         Begin VB.Label lbltasainteres 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   67
            Top             =   510
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tasa de Interes :"
            Height          =   195
            Index           =   9
            Left            =   6375
            TabIndex        =   66
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label lblapoderado 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   65
            Top             =   1215
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado :"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   64
            Top             =   1230
            Width           =   1005
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblnota1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   63
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nota :"
            Height          =   195
            Index           =   7
            Left            =   6420
            TabIndex        =   62
            Top             =   225
            Width           =   435
         End
         Begin VB.Label lblanalista 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   61
            Top             =   885
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Analista :"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   60
            Top             =   915
            Width           =   645
         End
         Begin VB.Label lblLinea 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   59
            Top             =   555
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linea de Crédito :"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   58
            Top             =   600
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblfuente 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   57
            Top             =   225
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fuente de Ingreso :"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   56
            Top             =   255
            Width           =   1500
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tot Deuda Calend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   -67980
         TabIndex        =   44
         Top             =   450
         Width           =   1815
         Begin VB.Label lblSaldoKCalend 
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
            Height          =   315
            Left            =   810
            TabIndex        =   54
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo K"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   53
            Top             =   390
            Width           =   555
         End
         Begin VB.Label lblIntCompCalend 
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
            Height          =   315
            Left            =   810
            TabIndex        =   52
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Comp"
            Height          =   195
            Index           =   51
            Left            =   90
            TabIndex        =   51
            Top             =   810
            Width           =   630
         End
         Begin VB.Label lblGastoCalend 
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
            Height          =   315
            Left            =   810
            TabIndex        =   50
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto "
            Height          =   195
            Index           =   53
            Left            =   90
            TabIndex        =   49
            Top             =   1170
            Width           =   465
         End
         Begin VB.Label lblIntMorCalend 
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
            Height          =   315
            Left            =   810
            TabIndex        =   48
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Morat"
            Height          =   195
            Index           =   57
            Left            =   90
            TabIndex        =   47
            Top             =   1530
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total "
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
            Index           =   59
            Left            =   90
            TabIndex        =   46
            Top             =   2340
            Width           =   510
         End
         Begin VB.Label lblTotalCalend 
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
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   810
            TabIndex        =   45
            Top             =   2250
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2865
         Index           =   3
         Left            =   -74865
         TabIndex        =   20
         Top             =   555
         Width           =   10545
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   1290
            Left            =   5640
            TabIndex        =   29
            Top             =   240
            Width           =   1500
            Begin VB.CheckBox chkRefinanciado 
               Alignment       =   1  'Right Justify
               Caption         =   "Refinanciado?"
               Height          =   285
               Left            =   45
               TabIndex        =   32
               Top             =   525
               Width           =   1395
            End
            Begin VB.CheckBox chkProtesto 
               Alignment       =   1  'Right Justify
               Caption         =   "Protesto ?"
               Height          =   285
               Left            =   45
               TabIndex        =   31
               Top             =   210
               Width           =   1395
            End
            Begin VB.CheckBox chkCargoAuto 
               Alignment       =   1  'Right Justify
               Caption         =   "Cargo Automat."
               Height          =   285
               Left            =   45
               TabIndex        =   30
               Top             =   840
               Width           =   1395
            End
         End
         Begin VB.Frame fraDescPlan 
            Height          =   1005
            Left            =   5160
            TabIndex        =   23
            Top             =   1680
            Width           =   5280
            Begin VB.Label lblmodular 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1155
               TabIndex        =   28
               Top             =   585
               Width           =   2505
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cod Modular"
               Height          =   195
               Index           =   45
               Left            =   135
               TabIndex        =   27
               Top             =   630
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito Personal Descuentos Por Planillas"
               Height          =   195
               Index           =   44
               Left            =   120
               TabIndex        =   26
               Top             =   0
               Width           =   2955
            End
            Begin VB.Label lblinstitucion 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1155
               TabIndex        =   25
               Top             =   255
               Width           =   3960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Institución"
               Height          =   195
               Index           =   43
               Left            =   135
               TabIndex        =   24
               Top             =   300
               Width           =   720
            End
         End
         Begin VB.Frame fraCreditosRefinanciado 
            Caption         =   "Creditos Refinaciados"
            Height          =   1290
            Left            =   7320
            TabIndex        =   21
            Top             =   240
            Width           =   3045
            Begin MSComctlLib.ListView lstRefinanciados 
               Height          =   1080
               Left            =   60
               TabIndex        =   22
               Top             =   165
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   1905
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
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Cred. Anterior"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Capital "
                  Object.Width           =   1499
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Int Susp."
                  Object.Width           =   1499
               EndProperty
            End
         End
         Begin VB.Label lblmetodoLiquidacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4800
            TabIndex        =   43
            Top             =   1080
            Width           =   660
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Método Liquidación :"
            Height          =   195
            Index           =   47
            Left            =   3240
            TabIndex        =   42
            Top             =   1125
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Judicial"
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
            Index           =   42
            Left            =   120
            TabIndex        =   41
            Top             =   1455
            Width           =   660
         End
         Begin VB.Label lblfechajudicial 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   40
            Top             =   1725
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de  Ing. Judicial :"
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   39
            Top             =   1725
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Cancelación de Crédito"
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
            Index           =   40
            Left            =   120
            TabIndex        =   38
            Top             =   825
            Width           =   2250
         End
         Begin VB.Label lblfechaCancelacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   37
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Cancelación :"
            Height          =   195
            Index           =   39
            Left            =   120
            TabIndex        =   36
            Top             =   1095
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Rechazo de Crédito "
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
            Index           =   38
            Left            =   120
            TabIndex        =   35
            Top             =   165
            Width           =   1890
         End
         Begin VB.Label lblrechazo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1440
            TabIndex        =   34
            Top             =   405
            Width           =   4065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Rechazo :"
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1260
         End
      End
      Begin VB.Frame fragarantias 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   18
         Top             =   510
         Width           =   10335
         Begin MSComctlLib.ListView lstgarantias 
            Height          =   2475
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Width           =   9630
            _ExtentX        =   16986
            _ExtentY        =   4366
            View            =   3
            Arrange         =   2
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Tipo Garantía"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Documento"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Nº Doc."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Moneda"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Monto Garantia"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Tot Deuda Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   -66090
         TabIndex        =   5
         Top             =   450
         Width           =   1875
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total "
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
            Index           =   35
            Left            =   90
            TabIndex        =   17
            Top             =   2295
            Width           =   510
         End
         Begin VB.Label lblGastoFecha 
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
            Height          =   315
            Left            =   810
            TabIndex        =   16
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Morat"
            Height          =   195
            Index           =   34
            Left            =   90
            TabIndex        =   15
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label lblTotalFecha 
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
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   810
            TabIndex        =   14
            Top             =   2235
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   13
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Comp"
            Height          =   195
            Index           =   46
            Left            =   90
            TabIndex        =   12
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo K"
            Height          =   195
            Index           =   48
            Left            =   90
            TabIndex        =   11
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblSaldoKFecha 
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
            Height          =   315
            Left            =   810
            TabIndex        =   10
            Top             =   360
            Width           =   930
         End
         Begin VB.Label lblIntCompFecha 
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
            Height          =   315
            Left            =   810
            TabIndex        =   9
            Top             =   720
            Width           =   930
         End
         Begin VB.Label lblIntMorFecha 
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
            Height          =   315
            Left            =   810
            TabIndex        =   8
            Top             =   1440
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Penalidad"
            Height          =   195
            Index           =   63
            Left            =   90
            TabIndex        =   7
            Top             =   1770
            Width           =   705
         End
         Begin VB.Label lblPenalidadFecha 
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
            Height          =   315
            Left            =   810
            TabIndex        =   6
            Top             =   1770
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmCredConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ConsultaCliente(ByVal psCtaCod As String)
    ActxCta.NroCuenta = psCtaCod
    Call ActXCta_KeyPress(13)
    CmdNuevaCons.Enabled = False
    Me.Show 1
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
Dim oCredPersRela As UCredRelacion
Dim l As ListItem
Dim RDatos As ADODB.Recordset
Dim R As ADODB.Recordset
Dim oCred As DCredito
Dim oCalend As Dcalendario
Dim oMontoDesemb As Double
Dim nCapPag As Double
Dim nIntPag As Double
Dim nGasto As Double
Dim nMora As Double
Dim oNegCred As NCredito
Dim MatCalend As Variant
Dim oGarantia As DGarantia
Dim i As Integer

    On Error GoTo ErrorCargaDatos
    Screen.MousePointer = 11
    CargaDatos = True
    'Carga Relaciones de Credito
    Set oCredPersRela = New UCredRelacion
    Call oCredPersRela.CargaRelacPersCred(psCtaCod)
    oCredPersRela.IniciarMatriz
    listaClientes.ListItems.Clear
    Do While Not oCredPersRela.EOF
        Set l = listaClientes.ListItems.Add(, , oCredPersRela.ObtenerNombre)
        l.SubItems(1) = oCredPersRela.ObtenerRelac
        l.SubItems(2) = oCredPersRela.ObtenerCodigo
        oCredPersRela.siguiente
    Loop
    Set oCredPersRela = Nothing
    
    'Carga Datos del Credito
    Set oCred = New DCredito
    Set RDatos = oCred.RecuperaConsultaCred(psCtaCod)
    Set oCred = Nothing
    If RDatos.BOF And RDatos.EOF Then
        CargaDatos = False
        'Call CargaDatos(psCtaCod)
        RDatos.Close
        Set RDatos = Nothing
        Exit Function
    End If
    lbltipoCredito.Caption = Trim(RDatos!cTipoCredDescrip)
    lblEstado.Caption = Trim(RDatos!cEstActual)
    lblfuente.Caption = Trim(IIf(IsNull(RDatos!cFteIngreso), "", RDatos!cFteIngreso))
    lblLinea.Caption = Trim(RDatos!cLineaDesc)
    lblanalista.Caption = Trim(RDatos!cAnalista)
    lblapoderado.Caption = Trim(IIf(IsNull(RDatos!cApoderado), "", RDatos!cApoderado))
    lblcondicion.Caption = Trim(RDatos!cCondicion)
    lbldestino.Caption = Trim(RDatos!cDestino)
    lblnota1.Caption = IIf(IsNull(RDatos!nNota), "", RDatos!nNota)
    lbltasainteres.Caption = Format(IIf(IsNull(RDatos!nTasaInteres), 0, RDatos!nTasaInteres), "#0.00")
    lbltipocuota.Caption = Trim(IIf(IsNull(RDatos!cTipoCuota), "", RDatos!cTipoCuota))
    If IsNull(RDatos!dVigencia) Then
        lblfechavigencia.Caption = ""
    Else
        lblfechavigencia.Caption = Format(RDatos!dVigencia, "dd/mm/yyyy")
    End If
    
    'Ficha de Historial
    If IsNull(RDatos!dFecSol) Then
        lblfechsolicitud.Caption = ""
    Else
        lblfechsolicitud.Caption = Format(RDatos!dFecSol, "dd/mm/yyyy")
    End If
    lblMontoSolicitado.Caption = Format(IIf(IsNull(RDatos!nMontoSol), 0, RDatos!nMontoSol), "#0.00")
    lblcuotasSolicitud.Caption = IIf(IsNull(RDatos!nCuotasSol), "", RDatos!nCuotasSol)
    lblPlazoSolicitud.Caption = IIf(IsNull(RDatos!nPlazoSol), 0, RDatos!nPlazoSol)
    If IsNull(RDatos!dFecSug) Then
        lblfechasugerida.Caption = ""
    Else
        lblfechasugerida.Caption = Format(RDatos!dFecSug, "dd/mm/yyyy")
    End If
    lblMontosugerido.Caption = Format(IIf(IsNull(RDatos!nMontoSug), 0, RDatos!nMontoSug), "#0.00")
    lblcuotasugerida.Caption = Format(IIf(IsNull(RDatos!nCuotasSug), 0, RDatos!nCuotasSug), "#0.00")
    lblPlazoSugerido.Caption = IIf(IsNull(RDatos!nPlazoSug), 0, RDatos!nPlazoSug)
    lblmontoCuotaSugerida.Caption = Format(IIf(IsNull(RDatos!nCuotaSug), 0, RDatos!nCuotaSug), "#0.00")
    lblGraciasugerida.Caption = IIf(IsNull(RDatos!nPeriodoGraciaSug), 0, RDatos!nPeriodoGraciaSug)
    If IsNull(RDatos!dFecApr) Then
        lblfechaAprobado.Caption = ""
    Else
        lblfechaAprobado.Caption = Format(RDatos!dFecApr, "dd/mm/yyyy")
    End If
    lblMontoAprobado.Caption = Format(IIf(IsNull(RDatos!nMontoApr), 0, RDatos!nMontoApr), "#0.00")
    lblcuotasAprobado.Caption = IIf(IsNull(RDatos!nCuotasApr), 0, RDatos!nCuotasApr)
    lblPlazoAprobado.Caption = IIf(IsNull(RDatos!nplazoapr), 0, RDatos!nCuotasApr)
    lblmontoCuotaAprobada.Caption = Format(IIf(IsNull(RDatos!nCuotaApr), 0, RDatos!nCuotaApr), "#0.00")
    lblGraciaAprobada.Caption = IIf(IsNull(RDatos!nPeriodoGraciaApr), 0, RDatos!nPeriodoGraciaApr)
    lblIntGraciaApr.Caption = Format(IIf(IsNull(RDatos!nTasaGracia), 0, RDatos!nTasaGracia), "#0.00")
    lblTipoGraciaApr.Caption = IIf(IsNull(RDatos!cTipoGracia), "", RDatos!cTipoGracia)
    
    'Ficha Desembolsos Realizados
    listaDesembolsos.ListItems.Clear
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    oMontoDesemb = 0
    Do While Not R.EOF
        Set l = listaDesembolsos.ListItems.Add(, , Format(R!dVenc, "dd/mm/yyyy"))
        l.SubItems(1) = R!nCuota
        l.SubItems(2) = Format(R!nCapital, "#0.00")
        l.SubItems(3) = Format(R!nGasto, "#0.00")
        l.SubItems(4) = IIf(R!nColocCalendEstado = gColocCalendEstadoPendiente, "PENDIENTE", "CANCELADO")
        If R!nColocCalendEstado = gColocCalendEstadoPagado Then
            oMontoDesemb = oMontoDesemb + R!nCapital
        End If
        R.MoveNext
    Loop
    lbltipoDesembolso.Caption = IIf(IsNull(RDatos!cTipoDesemb), "", RDatos!cTipoDesemb)
    lblmontoDesembolsado.Caption = Format(oMontoDesemb, "#0.00")
    lbltotalDesembolso.Caption = Format(IIf(IsNull(RDatos!nMontoApr), 0, RDatos!nMontoApr), "#0.00")
    R.Close
    Set R = Nothing
    
    'Ficha Pagos Realizados
    Set oCalend = New Dcalendario
    Set R = oCalend.RecuperaCalendarioPagosRealizados(psCtaCod)
    Set oCalend = Nothing
    ListaPagos.ListItems.Clear
    nCapPag = 0
    nIntPag = 0
    nGasto = 0
    nMora = 0

    Do While Not R.EOF
        Set l = ListaPagos.ListItems.Add(, , Mid(R!dFecCanc, 7, 2) & "/" & Mid(R!dFecCanc, 5, 2) & "/" & Mid(R!dFecCanc, 1, 4))
        l.SubItems(1) = R!nCuota
        l.SubItems(2) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital) + _
                             IIf(IsNull(R!nIntComp), 0, R!nIntComp) + _
                             IIf(IsNull(R!nIntGracia), 0, R!nIntGracia) + _
                             IIf(IsNull(R!nIntMor), 0, R!nIntMor) + _
                             IIf(IsNull(R!nIntReprog), 0, R!nIntReprog) + _
                             IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso) + _
                             IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        l.SubItems(3) = Format(IIf(IsNull(R!nCapital), 0, R!nCapital), "#0.00")
        nCapPag = nCapPag + CDbl(l.SubItems(3))
        l.SubItems(4) = Format(IIf(IsNull(R!nIntComp), 0, R!nIntComp) + _
                               IIf(IsNull(R!nIntGracia), 0, R!nIntGracia) + _
                               IIf(IsNull(R!nIntReprog), 0, R!nIntReprog) + _
                               IIf(IsNull(R!nIntSuspenso), 0, R!nIntSuspenso), "#0.00")
        nIntPag = nIntPag + CDbl(l.SubItems(4))
        l.SubItems(5) = Format(IIf(IsNull(R!nIntMor), 0, R!nIntMor), "#0.00")
        nMora = nMora + CDbl(l.SubItems(5))
        l.SubItems(6) = Format(IIf(IsNull(R!nGasto), 0, R!nGasto), "#0.00")
        nGasto = nGasto + CDbl(l.SubItems(6))
        l.SubItems(7) = Format(DateDiff("d", R!dVenc, CDate(Mid(R!dFecCanc, 7, 2) & "/" & Mid(R!dFecCanc, 5, 2) & "/" & Mid(R!dFecCanc, 1, 4))), "#0.00")
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    lblcapitalpagado.Caption = Format(nCapPag, "#0.00")
    lblintcompPag.Caption = Format(nIntPag, "#0.00")
    lblIntMorPag.Caption = Format(nMora, "#0.00")
    lblGastopagado.Caption = Format(nGasto, "#0.00")
    
    'Ficha Pagos Pendientes
    Set oNegCred = New NCredito
    MatCalend = oNegCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
    lstCuotasPend.ListItems.Clear
    For i = 0 To UBound(MatCalend) - 1
        Set l = lstCuotasPend.ListItems.Add(, , MatCalend(i, 1))
        l.SubItems(1) = MatCalend(i, 0)
        l.SubItems(2) = MatCalend(i, 3)
        l.SubItems(3) = Format(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)), "#0.00")
        l.SubItems(4) = MatCalend(i, 6)
        l.SubItems(5) = DateDiff("d", CDate(MatCalend(i, 0)), gdFecSis)
    Next i
    
    lblCapitalCuoPend.Caption = Format(oNegCred.MatrizCapitalVencido(MatCalend, gdFecSis), "#0.00")
    lblInteresCuoPend.Caption = Format(oNegCred.MatrizIntCompVencido(MatCalend, gdFecSis), "#0.00")
    lblMoraCuoPend.Caption = Format(oNegCred.MatrizInteresMorFecha(psCtaCod, MatCalend), "#0.00")
    lblGastoCuoPend.Caption = Format(oNegCred.MatrizGastosFecha(psCtaCod, MatCalend), "#0.00")
    lblTotalCuoPend.Caption = Format(CDbl(lblCapitalCuoPend.Caption) + CDbl(lblInteresCuoPend.Caption) + CDbl(lblMoraCuoPend.Caption) + CDbl(lblGastoCuoPend.Caption), "#0.00")
    
    lblSaldoKCalend.Caption = oNegCred.MatrizCapitalCalendario(MatCalend)
    lblIntCompCalend.Caption = Format(oNegCred.MatrizIntCompCalendario(MatCalend) + _
                            oNegCred.MatrizIntGraciaCalendario(MatCalend) + _
                            oNegCred.MatrizIntReprogCalendario(MatCalend) + _
                            oNegCred.MatrizIntSuspensoCalendario(MatCalend), "#0.00")
    lblGastoCalend.Caption = Format(oNegCred.MatrizIntGastosCalendario(MatCalend), "#0.00")
    lblIntMorCalend.Caption = Format(oNegCred.MatrizIntMoratorioCalendario(MatCalend), "#0.00")
    lblTotalCalend.Caption = Format(CDbl(lblSaldoKCalend.Caption) + CDbl(lblIntCompCalend.Caption) + CDbl(lblGastoCalend.Caption) + CDbl(lblIntMorCalend.Caption), "#0.00")
    
    lblSaldoKFecha.Caption = Format(oNegCred.MatrizCapitalAFecha(psCtaCod, MatCalend), "#0.00")
    If UBound(MatCalend) > 0 Then
        lblIntCompFecha.Caption = Format(oNegCred.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, gdFecSis), "#0.00")
        lblGastoFecha.Caption = Format(oNegCred.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis), "#0.00")
        lblIntMorFecha.Caption = Format(oNegCred.MatrizInteresMorFecha(psCtaCod, MatCalend), "#0.00")
        lblPenalidadFecha.Caption = Format(oNegCred.CalculaGastoPenalidadCancelacion(CDbl(lblSaldoKFecha.Caption), CInt(Mid(psCtaCod, 9, 1))), "#0.00")
        lblTotalFecha.Caption = Format(CDbl(lblSaldoKFecha.Caption) + CDbl(lblIntCompFecha.Caption) + CDbl(lblGastoFecha.Caption) + CDbl(lblIntMorFecha.Caption) + CDbl(lblPenalidadFecha.Caption), "#0.00")
    End If
        
    
    
    'Ficha de Garantias
    lstgarantias.ListItems.Clear
    Set oGarantia = New DGarantia
    Set R = oGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oGarantia = Nothing
    Do While Not R.EOF
        Set l = lstgarantias.ListItems.Add(, , Trim(R!cTpoGarantia))
        l.SubItems(1) = Trim(R!cDescripcion)
        l.SubItems(2) = Trim(R!cDocDesc)
        l.SubItems(3) = Trim(R!cNroDoc)
        l.SubItems(4) = Trim(R!cMoneda)
        l.SubItems(5) = Format(R!nGravado, "#0.00")
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    'Ficha de Otros Datos
    If IsNull(RDatos!cMotivoRech) Then
        lblrechazo.Caption = ""
    Else
        lblrechazo.Caption = Trim(RDatos!cMotivoRech)
    End If
    If IsNull(RDatos!dFecCancel) Then
        lblfechaCancelacion.Caption = ""
    Else
        lblfechaCancelacion.Caption = Format(RDatos!dFecCancel, "dd/mm/yyyy")
    End If
    lblmetodoLiquidacion.Caption = Trim(IIf(IsNull(RDatos!cMetLiquidacion), "", RDatos!cMetLiquidacion))
    If IsNull(RDatos!dFecJud) Then
        lblfechajudicial.Caption = ""
    Else
        lblfechajudicial.Caption = Format(RDatos!dFecJud, "dd/mm/yyyy")
    End If
    
    If IsNull(RDatos!cProtesto) Then
        chkProtesto.value = 0
    Else
        If Trim(RDatos!cProtesto) = "1" Then
            chkProtesto.value = 1
        Else
            chkProtesto.value = 0
        End If
    End If
    If IsNull(RDatos!nEstRefin) Then
        chkRefinanciado.value = 0
    Else
        If RDatos!nEstRefin = 1 Then
            chkRefinanciado.value = 1
        Else
            chkRefinanciado.value = 0
        End If
    End If
        
    If IsNull(RDatos!bCargoAuto) Then
        chkCargoAuto.value = 0
    Else
        If RDatos!bCargoAuto = True Then
            chkCargoAuto.value = 1
        Else
            chkCargoAuto.value = 0
        End If
    End If
    
    lstRefinanciados.ListItems.Clear
    Set oCred = New DCredito
    Set R = oCred.RecuperaCreditosRefinanciados(psCtaCod)
    Set oCred = Nothing
    Do While Not R.EOF
        Set l = lstRefinanciados.ListItems.Add(, , R!cCtaCodRef)
        l.SubItems(1) = Format(IIf(IsNull(R!nCapitalRef), 0, R!nCapitalRef), "#0.00")
        l.SubItems(2) = Format(IIf(IsNull(R!nInteresRef), 0, R!nInteresRef), "#0.00")
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        lblinstitucion.Caption = PstaNombre(RDatos!cConvenio)
        lblmodular.Caption = Trim(RDatos!cCodModular)
    End If
    RDatos.Close
    Set RDatos = Nothing
    Set oNegCred = Nothing
    Screen.MousePointer = 0
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"

End Function

Private Sub ActXCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            ActxCta.NroCuenta = ""
            ActxCta.Enabled = True
            LimpiaControles Me, True
            MsgBox "No se encontro el Credito", vbInformation, "Aviso"
        Else
            ActxCta.Enabled = False
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim oCredDoc As NCredDoc
Dim Prev As Previo.clsPrevio
Dim MatRelacCred() As String
Dim MatDesembolsos() As String
Dim MatCuotasPend() As String
Dim MatHistorial(3, 7) As String
Dim MatDeudaVenc(5) As String
Dim MatDeudaAFecha(5) As String

Dim i As Integer
    On Error GoTo ErrorCmdImprimir_Click
    
    'Pago a la fecha
    MatDeudaAFecha(0) = lblSaldoKFecha.Caption
    MatDeudaAFecha(1) = lblIntCompFecha.Caption
    MatDeudaAFecha(2) = lblIntMorFecha.Caption
    MatDeudaAFecha(3) = lblGastoFecha.Caption
    MatDeudaAFecha(4) = lblTotalFecha.Caption
    
    'Cuotas Vencidas
    MatDeudaVenc(0) = lblCapitalCuoPend.Caption
    MatDeudaVenc(1) = lblInteresCuoPend.Caption
    MatDeudaVenc(2) = lblMoraCuoPend.Caption
    MatDeudaVenc(3) = lblGastoCuoPend.Caption
    MatDeudaVenc(4) = lblTotalCuoPend.Caption
    
    
    'Carga Cuotas Pendientes
    ReDim MatCuotasPend(lstCuotasPend.ListItems.Count, 6)
    For i = 0 To lstCuotasPend.ListItems.Count - 1
        MatCuotasPend(i, 0) = lstCuotasPend.ListItems(i + 1).Text
        MatCuotasPend(i, 1) = lstCuotasPend.ListItems(i + 1).SubItems(1)
        MatCuotasPend(i, 2) = lstCuotasPend.ListItems(i + 1).SubItems(2)
        MatCuotasPend(i, 3) = lstCuotasPend.ListItems(i + 1).SubItems(3)
        MatCuotasPend(i, 4) = lstCuotasPend.ListItems(i + 1).SubItems(4)
        MatCuotasPend(i, 5) = lstCuotasPend.ListItems(i + 1).SubItems(5)
    Next i
    
    'Carga Desembolsos
    ReDim MatDesembolsos(listaDesembolsos.ListItems.Count, 4)
    For i = 0 To listaDesembolsos.ListItems.Count - 1
        MatDesembolsos(i, 0) = listaDesembolsos.ListItems(i + 1).Text
        MatDesembolsos(i, 1) = listaDesembolsos.ListItems(i + 1).SubItems(2)
        MatDesembolsos(i, 2) = listaDesembolsos.ListItems(i + 1).SubItems(3)
        MatDesembolsos(i, 3) = listaDesembolsos.ListItems(i + 1).SubItems(4)
    Next i
    
    'Carga Relaciones de Clientes en Matriz
    ReDim MatRelacCred(listaClientes.ListItems.Count, 2)
    For i = 0 To listaClientes.ListItems.Count - 1
        MatRelacCred(i, 0) = PstaNombre(listaClientes.ListItems(i + 1).Text)
        MatRelacCred(i, 1) = PstaNombre(listaClientes.ListItems(i + 1).SubItems(1))
    Next i
    
    'Carga Relaciones de Clientes en Matriz
    MatHistorial(0, 0) = "SOLICITUD"
    MatHistorial(0, 1) = lblfechsolicitud.Caption
    MatHistorial(0, 2) = lblMontoSolicitado.Caption
    MatHistorial(0, 3) = lblcuotasSolicitud.Caption
    MatHistorial(0, 4) = lblPlazoSolicitud.Caption
    MatHistorial(0, 5) = ""
    MatHistorial(0, 6) = ""
    
    MatHistorial(0, 0) = "SUGERENCIA"
    MatHistorial(0, 1) = lblfechasugerida.Caption
    MatHistorial(0, 2) = lblMontosugerido.Caption
    MatHistorial(0, 3) = lblcuotasugerida.Caption
    MatHistorial(0, 4) = lblPlazoSugerido.Caption
    MatHistorial(0, 5) = lblmontoCuotaSugerida.Caption
    MatHistorial(0, 6) = lblGraciasugerida.Caption
    
    MatHistorial(0, 0) = "APROBACION"
    MatHistorial(0, 1) = lblfechaAprobado.Caption
    MatHistorial(0, 2) = lblMontoAprobado.Caption
    MatHistorial(0, 3) = lblcuotasAprobado.Caption
    MatHistorial(0, 4) = lblPlazoAprobado.Caption
    MatHistorial(0, 5) = lblmontoCuotaAprobada.Caption
    MatHistorial(0, 6) = lblGraciaAprobada.Caption
    
    If MatDeudaAFecha(1) = "" Then MatDeudaAFecha(1) = "0.00"
    If MatDeudaAFecha(2) = "" Then MatDeudaAFecha(2) = "0.00"
    If MatDeudaAFecha(3) = "" Then MatDeudaAFecha(3) = "0.00"
    If MatDeudaAFecha(4) = "" Then MatDeudaAFecha(4) = "0.00"
    
    
    Set oCredDoc = New NCredDoc
    Set Prev = New clsPrevio
    Prev.Show oCredDoc.ImprimeConsultaCredito(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, lbltipoCredito.Caption, lblEstado _
        , lblLinea.Caption, lblfuente.Caption, MatRelacCred, lblanalista.Caption, lblcondicion.Caption, lbldestino.Caption, lblapoderado.Caption _
        , lbltipocuota.Caption, MatHistorial, MatDesembolsos, MatCuotasPend, MatDeudaVenc, MatDeudaAFecha, IIf(chkProtesto.value = 1, "SI", "NO"), IIf(chkCargoAuto.value = 1, "SI", "NO"), IIf(chkRefinanciado.value = 1, "SI", "NO")), "", True
    Set Prev = Nothing
    Set oCredDoc = Nothing
    Exit Sub
ErrorCmdImprimir_Click:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub CmdNuevaCons_Click()
    ActxCta.Enabled = True
    LimpiaControles Me, True
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    lstCuotasPend.ListItems.Clear
    listaDesembolsos.ListItems.Clear
    ListaPagos.ListItems.Clear
    lstgarantias.ListItems.Clear
    listaClientes.ListItems.Clear
    chkProtesto.value = 0
    chkRefinanciado.value = 0
    chkCargoAuto.value = 0
    lstRefinanciados.ListItems.Clear
End Sub

Private Sub Form_Load()
    CentraSdi Me
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub
