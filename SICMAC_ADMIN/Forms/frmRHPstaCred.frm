VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRHPstaCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Historial Credito"
   ClientHeight    =   6135
   ClientLeft      =   480
   ClientTop       =   2130
   ClientWidth     =   11010
   Icon            =   "frmRHPstaCred.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCalMora 
      Caption         =   "&Calc. Mora"
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
      Left            =   7155
      TabIndex        =   99
      Top             =   5655
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   8370
      TabIndex        =   98
      Top             =   5655
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   345
      Left            =   90
      TabIndex        =   97
      Top             =   5760
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRHPstaCred.frx":030A
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
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   5940
      TabIndex        =   1
      Top             =   5655
      Width           =   1215
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
      Left            =   9585
      TabIndex        =   0
      Top             =   5655
      Width           =   1215
   End
   Begin VB.Frame fradatosgenerales 
      Height          =   2040
      Left            =   105
      TabIndex        =   2
      Top             =   -30
      Width           =   10755
      Begin Sicmact.ActxCtaCred lblCodCred 
         Height          =   555
         Left            =   165
         TabIndex        =   159
         Top             =   210
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   979
         enabledage      =   -1  'True
         enabledprod     =   -1  'True
         enabled         =   -1  'True
         Caption         =   "Cuenta N°"
      End
      Begin MSComctlLib.ListView listaClientes 
         Height          =   1650
         Left            =   5430
         TabIndex        =   3
         ToolTipText     =   "Presione doblick o [Enter] para editar al Cliente"
         Top             =   255
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
      Begin VB.Label lbltipoCredito 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1050
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Crédito :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Top             =   795
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         TabIndex        =   5
         Top             =   1455
         Width           =   2475
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
         Left            =   180
         TabIndex        =   4
         Top             =   1530
         Width           =   1485
         WordWrap        =   -1  'True
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3450
      Left            =   90
      TabIndex        =   8
      Top             =   2160
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6085
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
      TabPicture(0)   =   "frmRHPstaCred.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Historial "
      TabPicture(1)   =   "frmRHPstaCred.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Desembolsos Realizados"
      TabPicture(2)   =   "frmRHPstaCred.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pagos &Realizados"
      TabPicture(3)   =   "frmRHPstaCred.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Pagos &Pendientes"
      TabPicture(4)   =   "frmRHPstaCred.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame6"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&Garantías"
      TabPicture(5)   =   "frmRHPstaCred.frx":0416
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fragarantias"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Otros Datos"
      TabPicture(6)   =   "frmRHPstaCred.frx":0432
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2(3)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
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
         Height          =   2805
         Left            =   -66090
         TabIndex        =   126
         Top             =   450
         Width           =   1875
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Susp"
            Height          =   195
            Index           =   58
            Left            =   90
            TabIndex        =   138
            Top             =   1800
            Width           =   585
         End
         Begin VB.Label lblIntSuspFecha 
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
            TabIndex        =   137
            Top             =   1800
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
            TabIndex        =   136
            Top             =   1440
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
            TabIndex        =   135
            Top             =   720
            Width           =   930
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
            TabIndex        =   134
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo K"
            Height          =   195
            Index           =   48
            Left            =   90
            TabIndex        =   133
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Comp"
            Height          =   195
            Index           =   46
            Left            =   90
            TabIndex        =   132
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   131
            Top             =   1080
            Width           =   420
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
            TabIndex        =   130
            Top             =   2250
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Morat"
            Height          =   195
            Index           =   34
            Left            =   90
            TabIndex        =   129
            Top             =   1440
            Width           =   630
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
            TabIndex        =   128
            Top             =   1080
            Width           =   930
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
            Index           =   35
            Left            =   90
            TabIndex        =   127
            Top             =   2340
            Width           =   510
         End
      End
      Begin VB.Frame fragarantias 
         Height          =   2535
         Left            =   -74775
         TabIndex        =   122
         Top             =   510
         Width           =   10215
         Begin MSComctlLib.ListView lstgarantias 
            Height          =   2115
            Left            =   330
            TabIndex        =   123
            Top             =   240
            Width           =   9510
            _ExtentX        =   16775
            _ExtentY        =   3731
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nº Garantía"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tipo Garantía"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descripción"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Documento"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Nº Doc."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Moneda"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Monto Credito"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Monto Garantia"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Estado"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2625
         Index           =   3
         Left            =   -74865
         TabIndex        =   100
         Top             =   555
         Width           =   10545
         Begin VB.Frame fraCreditosRefinanciado 
            Caption         =   "Creditos Refinaciados"
            Height          =   1290
            Left            =   7455
            TabIndex        =   155
            Top             =   225
            Width           =   3045
            Begin MSComctlLib.ListView lstRefinanciados 
               Height          =   1080
               Left            =   60
               TabIndex        =   156
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
         Begin VB.Frame fraDescPlan 
            Height          =   1005
            Left            =   3495
            TabIndex        =   105
            Top             =   1560
            Width           =   5280
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Institución"
               Height          =   195
               Index           =   43
               Left            =   135
               TabIndex        =   110
               Top             =   300
               Width           =   720
            End
            Begin VB.Label lblinstitucion 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1155
               TabIndex        =   109
               Top             =   255
               Width           =   3960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito Personal Descuentos Por Planillas"
               Height          =   195
               Index           =   44
               Left            =   120
               TabIndex        =   108
               Top             =   0
               Width           =   2955
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cod Modular"
               Height          =   195
               Index           =   45
               Left            =   135
               TabIndex        =   107
               Top             =   630
               Width           =   900
            End
            Begin VB.Label lblmodular 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1155
               TabIndex        =   106
               Top             =   585
               Width           =   2505
            End
         End
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   1290
            Left            =   5910
            TabIndex        =   101
            Top             =   225
            Width           =   1500
            Begin VB.CheckBox chkCargoAuto 
               Alignment       =   1  'Right Justify
               Caption         =   "Cargo Automat."
               Height          =   285
               Left            =   45
               TabIndex        =   104
               Top             =   840
               Width           =   1395
            End
            Begin VB.CheckBox chkProtesto 
               Alignment       =   1  'Right Justify
               Caption         =   "Protesto ?"
               Height          =   285
               Left            =   45
               TabIndex        =   103
               Top             =   210
               Width           =   1395
            End
            Begin VB.CheckBox chkRefinanciado 
               Alignment       =   1  'Right Justify
               Caption         =   "Refinanciado?"
               Height          =   285
               Left            =   45
               TabIndex        =   102
               Top             =   525
               Width           =   1395
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Motivo de Rechazo :"
            Height          =   195
            Index           =   37
            Left            =   210
            TabIndex        =   121
            Top             =   435
            Width           =   1485
         End
         Begin VB.Label lblrechazo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1800
            TabIndex        =   120
            Top             =   405
            Width           =   4080
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
            Left            =   210
            TabIndex        =   119
            Top             =   165
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Cancelación :"
            Height          =   195
            Index           =   39
            Left            =   255
            TabIndex        =   118
            Top             =   1095
            Width           =   1695
         End
         Begin VB.Label lblfechaCancelacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2055
            TabIndex        =   117
            Top             =   1065
            Width           =   1200
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
            Left            =   210
            TabIndex        =   116
            Top             =   825
            Width           =   2250
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de  Ing. Judicial :"
            Height          =   195
            Index           =   41
            Left            =   270
            TabIndex        =   115
            Top             =   1725
            Width           =   1695
         End
         Begin VB.Label lblfechajudicial 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2055
            TabIndex        =   114
            Top             =   1725
            Width           =   1200
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
            Left            =   195
            TabIndex        =   113
            Top             =   1455
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Método Liquidación :"
            Height          =   195
            Index           =   47
            Left            =   3645
            TabIndex        =   112
            Top             =   1125
            Width           =   1485
         End
         Begin VB.Label lblmetodoLiquidacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5220
            TabIndex        =   111
            Top             =   1080
            Width           =   660
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tot Pend Calend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   -67980
         TabIndex        =   88
         Top             =   450
         Width           =   1815
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
            TabIndex        =   140
            Top             =   2250
            Width           =   915
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
            TabIndex        =   139
            Top             =   2340
            Width           =   510
         End
         Begin VB.Label lblIntSusPendCalend 
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
            TabIndex        =   125
            Top             =   1800
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Susp"
            Height          =   195
            Index           =   36
            Left            =   90
            TabIndex        =   124
            Top             =   1890
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Morat"
            Height          =   195
            Index           =   57
            Left            =   90
            TabIndex        =   96
            Top             =   1530
            Width           =   630
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
            TabIndex        =   95
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto "
            Height          =   195
            Index           =   53
            Left            =   90
            TabIndex        =   94
            Top             =   1170
            Width           =   465
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
            TabIndex        =   93
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Comp"
            Height          =   195
            Index           =   51
            Left            =   90
            TabIndex        =   92
            Top             =   810
            Width           =   630
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
            TabIndex        =   91
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo K"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   90
            Top             =   390
            Width           =   555
         End
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
            TabIndex        =   89
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2490
         Left            =   300
         TabIndex        =   52
         Top             =   495
         Width           =   10095
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fuente de Ingreso :"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   77
            Top             =   255
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblfuente 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   76
            Top             =   225
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linea de Crédito :"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   75
            Top             =   600
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLinea 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   74
            Top             =   555
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Analista :"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   73
            Top             =   915
            Width           =   645
         End
         Begin VB.Label lblanalista 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   72
            Top             =   885
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nota 1 :"
            Height          =   195
            Index           =   7
            Left            =   6420
            TabIndex        =   71
            Top             =   225
            Width           =   570
         End
         Begin VB.Label lblnota1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   70
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nota 2 :"
            Height          =   195
            Index           =   8
            Left            =   6405
            TabIndex        =   69
            Top             =   525
            Width           =   570
         End
         Begin VB.Label lblnota2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   68
            Top             =   495
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado :"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   67
            Top             =   1230
            Width           =   1005
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblapoderado 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   66
            Top             =   1215
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tasa de Interes :"
            Height          =   195
            Index           =   9
            Left            =   6375
            TabIndex        =   65
            Top             =   855
            Width           =   1200
         End
         Begin VB.Label lbltasainteres 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   64
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   10
            Left            =   8625
            TabIndex        =   63
            Top             =   855
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Crédito :"
            Height          =   195
            Index           =   11
            Left            =   105
            TabIndex        =   62
            Top             =   1560
            Width           =   1560
         End
         Begin VB.Label lblcondicion 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   61
            Top             =   1545
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Destino del Crédito :"
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   60
            Top             =   1905
            Width           =   1425
         End
         Begin VB.Label lbldestino 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   59
            Top             =   1875
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cuota :"
            Height          =   195
            Index           =   13
            Left            =   6375
            TabIndex        =   58
            Top             =   1155
            Width           =   1095
         End
         Begin VB.Label lbltipocuota 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   57
            Top             =   1125
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Periodo :"
            Height          =   195
            Index           =   14
            Left            =   6390
            TabIndex        =   56
            Top             =   1470
            Width           =   630
         End
         Begin VB.Label lblperiodo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   55
            Top             =   1440
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vigencia :"
            Height          =   195
            Index           =   15
            Left            =   6390
            TabIndex        =   54
            Top             =   1800
            Width           =   1425
         End
         Begin VB.Label lblfechavigencia 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7890
            TabIndex        =   53
            Top             =   1770
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2610
         Index           =   1
         Left            =   -74880
         TabIndex        =   38
         Top             =   450
         Width           =   10230
         Begin MSComctlLib.ListView listaDesembolsos 
            Height          =   1950
            Left            =   180
            TabIndex        =   39
            Top             =   375
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   3440
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
            NumItems        =   4
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
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Desembolso :"
            Height          =   210
            Index           =   26
            Left            =   7170
            TabIndex        =   51
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label lbltipoDesembolso 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   50
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prox. Desemb.:"
            Height          =   195
            Index           =   27
            Left            =   7170
            TabIndex        =   49
            Top             =   1470
            Width           =   1080
         End
         Begin VB.Label lblProximoDesembolso 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   48
            Top             =   1416
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desembolsado :"
            Height          =   195
            Index           =   28
            Left            =   7170
            TabIndex        =   47
            Top             =   735
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Prox. Desemb.:"
            Height          =   195
            Index           =   29
            Left            =   7170
            TabIndex        =   46
            Top             =   2190
            Width           =   1575
         End
         Begin VB.Label lblfechaproxdesem 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   45
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblmontoDesembolsado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   44
            Top             =   672
            Width           =   1155
         End
         Begin VB.Label lblMontoProxDesem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   43
            Top             =   1788
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto Prox. Desemb.:"
            Height          =   195
            Index           =   49
            Left            =   7170
            TabIndex        =   42
            Top             =   1830
            Width           =   1575
         End
         Begin VB.Label lbltotalDesembolso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   41
            Top             =   1044
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Desembolso :"
            Height          =   210
            Index           =   50
            Left            =   7170
            TabIndex        =   40
            Top             =   1095
            Width           =   1365
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Index           =   2
         Left            =   -74880
         TabIndex        =   36
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
            TabIndex        =   79
            Top             =   285
            Width           =   2010
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capital :"
               Height          =   195
               Index           =   30
               Left            =   195
               TabIndex        =   86
               Top             =   480
               Width           =   570
            End
            Begin VB.Label lblintcompPag 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   85
               Top             =   780
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Interés :"
               Height          =   195
               Index           =   31
               Left            =   195
               TabIndex        =   84
               Top             =   825
               Width           =   570
            End
            Begin VB.Label lblIntMorPag 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   795
               TabIndex        =   83
               Top             =   1425
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mora :"
               Height          =   195
               Index           =   32
               Left            =   225
               TabIndex        =   82
               Top             =   1470
               Width           =   450
            End
            Begin VB.Label lblGastopagado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   795
               TabIndex        =   81
               Top             =   1110
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gasto :"
               Height          =   195
               Index           =   33
               Left            =   210
               TabIndex        =   80
               Top             =   1140
               Width           =   510
            End
         End
         Begin MSComctlLib.ListView ListaPagos 
            Height          =   2370
            Left            =   180
            TabIndex        =   37
            Top             =   270
            Width           =   7470
            _ExtentX        =   13176
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
      Begin VB.Frame Frame7 
         Caption         =   "Cuotas Pendiente:"
         Height          =   2820
         Left            =   -74880
         TabIndex        =   35
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
            Height          =   840
            Left            =   0
            TabIndex        =   141
            Top             =   1980
            Width           =   6855
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Interes "
               Height          =   195
               Index           =   60
               Left            =   1110
               TabIndex        =   151
               Top             =   165
               Width           =   525
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
               TabIndex        =   150
               Top             =   375
               Width           =   915
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
               TabIndex        =   149
               Top             =   165
               Width           =   510
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
               TabIndex        =   148
               Top             =   375
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gasto"
               Height          =   195
               Index           =   56
               Left            =   2970
               TabIndex        =   147
               Top             =   165
               Width           =   420
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
               TabIndex        =   146
               Top             =   375
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mora "
               Height          =   195
               Index           =   55
               Left            =   2055
               TabIndex        =   145
               Top             =   165
               Width           =   405
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
               TabIndex        =   144
               Top             =   375
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capital"
               Height          =   195
               Index           =   54
               Left            =   150
               TabIndex        =   143
               Top             =   180
               Width           =   480
            End
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
               TabIndex        =   142
               Top             =   375
               Width           =   915
            End
         End
         Begin MSComctlLib.ListView lstCuotasPend 
            Height          =   1725
            Left            =   120
            TabIndex        =   78
            Top             =   210
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   3043
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
            NumItems        =   7
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
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "C/D"
               Object.Width           =   882
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2670
         Left            =   -74790
         TabIndex        =   9
         Top             =   585
         Width           =   10365
         Begin VB.TextBox txtComenta 
            Height          =   705
            Left            =   1530
            TabIndex        =   157
            Top             =   1755
            Width           =   5790
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comentario :"
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
            Index           =   62
            Left            =   465
            TabIndex        =   158
            Top             =   1785
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Gracia Aprobada"
            Height          =   390
            Index           =   61
            Left            =   7965
            TabIndex        =   154
            Top             =   255
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblIntGraciaApr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7425
            TabIndex        =   153
            Top             =   1245
            Width           =   945
         End
         Begin VB.Label lblTipoGraciaApr 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   8415
            TabIndex        =   152
            Top             =   1245
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Periodo Gracia"
            Height          =   390
            Index           =   25
            Left            =   6615
            TabIndex        =   34
            Top             =   240
            Width           =   765
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblGraciasugerida 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6705
            TabIndex        =   33
            Top             =   870
            Width           =   615
         End
         Begin VB.Label lblGraciaAprobada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6705
            TabIndex        =   32
            Top             =   1275
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
            Height          =   195
            Index           =   24
            Left            =   1860
            TabIndex        =   31
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblfechasugerida 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   30
            Top             =   885
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblfechaAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   29
            Top             =   1275
            Width           =   1365
         End
         Begin VB.Label lblfechsolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   28
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            Height          =   195
            Index           =   23
            Left            =   5895
            TabIndex        =   27
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblmontoCuotaSugerida 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5655
            TabIndex        =   26
            Top             =   870
            Width           =   1035
         End
         Begin VB.Label lblmontoCuotaAprobada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5655
            TabIndex        =   25
            Top             =   1275
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
            Height          =   195
            Index           =   22
            Left            =   5055
            TabIndex        =   24
            Top             =   255
            Width           =   390
         End
         Begin VB.Label lblPlazoSugerido 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   23
            Top             =   885
            Width           =   735
         End
         Begin VB.Label lblPlazoAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   22
            Top             =   1275
            Width           =   735
         End
         Begin VB.Label lblPlazoSolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   21
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cuotas"
            Height          =   195
            Index           =   21
            Left            =   4200
            TabIndex        =   20
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblcuotasugerida 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   19
            Top             =   885
            Width           =   675
         End
         Begin VB.Label lblcuotasAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   18
            Top             =   1275
            Width           =   675
         End
         Begin VB.Label lblcuotasSolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   17
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto "
            Height          =   195
            Index           =   20
            Left            =   3240
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblMontosugerido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   15
            Top             =   885
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
            TabIndex        =   14
            Top             =   915
            Width           =   945
         End
         Begin VB.Label lblMontoAprobado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   13
            Top             =   1275
            Width           =   1215
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
            Left            =   465
            TabIndex        =   12
            Top             =   1305
            Width           =   945
         End
         Begin VB.Label lblmontosolicitud 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   11
            Top             =   480
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
            TabIndex        =   10
            Top             =   495
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmRHPstaCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificado por ejrs el dia 12/12/2000
'para el cambio de base de datos consolidada
Option Explicit
Private Type TPlanPagos
    NumCuota As Integer
    FecVenc As Date
    Capital As Double
    CapPag As Double
    Interes As Double
    IntPag As Double
    IntGra As Double
    IntGraPag As Double
    MontoCuota As Double
    Gasto As Double
    Mora As Double
    MoraPag As Double
    Total As Double
    Estado As String * 1
    Modificado As Boolean
    DiasAtr As Integer
End Type

Private Type TGastos
    Codgasto As String
    NumCuota As Integer
    MontoGasto As Double
    Gastopag As Double
    Estado As String * 1
    Modificado As Boolean
End Type

Dim lnNumLineas As Long
Dim lsRefinanciado As String
Dim TotalInteresaFecha As Currency
Dim MatPagos() As TPlanPagos
Dim MatGastos() As TGastos
Dim ContCuotas As Integer
Dim ContGastos As Integer
Dim NSalcap As Double
Dim MontoPagado As Double
Dim estcred As String * 1
Dim cCalifica As String * 1
Dim NumTrans As Integer
Dim FH As String
Dim nDias As Integer
Dim CnxRem As Boolean
Dim MoraTotal As Double
Dim ValorCombo() As String
Dim ContCombo As Integer

Dim FecUltPago As Date
Dim TasaInt As Double
Dim Periodo As Integer
Dim FecPagPend As Date
Dim MatSaldos() As Double
Dim nSaldos As Integer
Dim Tipodesemb As String
Dim FecVig As Date
Dim VRefin As String
Dim VMetLiquid As String
Dim VLinCred As String
Dim dbConexion As New ADODB.Connection
Dim lsCodCred As String
Dim lsAgencia As String
Dim lsEstado As String
Dim lsEstadoCredito  As String
Dim lbIni As Boolean
Dim lnNroRepro As Integer
'Para Cuota Libre
Dim dFecUltPago As Date
Dim nTasaInt As Double
Dim nSaldoCap As Double
Dim vIntPend As Double
Dim vIntMorCal As Double

Private Sub EncabCalMora(CodCta As String)
    rtf.Text = rtf.Text & Space(5) & gsNomCmac & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & gsNomAge & Space(100) & "Fecha : " & Format(gdFecSis + Time, "dd/mm/yyyy hh:mm:ss AMPM") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & CentrarCadena("CALCULO DE MORA DEL CREDITO " & CodCta, 150) & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnSaltoLinea
    rtf.Text = rtf.Text & Space(5) & "Fecha Pago" + Space(2) + "# Cuota" + Space(2) + "F. Cuota" + Space(4) + "Atraso Acum/Apli." + Space(2) + "Monto en Mora" + Space(2) + "Cap. Cuota" + Space(2) + "Mora Aplic." + Space(2) + "A Mora" + Space(2) + "A Cuotas" + Space(2) + "Total Pagado"
End Sub

Private Sub CalculaMora(CodCta As String)
'Dim R As New ADODB.Recordset
'Dim sql As String
'Dim MPlan() As TPlanPagos
'Dim Cont As Integer
'Dim PosC As Integer
'Dim Monto As Double
'Dim VDAtr As Integer
'Dim PosCP As Integer
'Dim ACuota As Double
'Dim AMora As Double
'Dim mcuota As Double
'Dim Codlin As String
'Dim Mora As Double
'Dim DMora As Integer
'Dim IntKar As Double
'Dim CapKar As Double
'    Cont = 0
'    'Halla la Tasa Moratoria
'    sql = "Select cCodLinCred From Credito Where cCodCta ='" & CodCta & "'"
'    'R.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'    If Not R.BOF And Not R.EOF Then
'        Codlin = R!cCodLinCred
'    Else
'        MsgBox "Credito No Encontrado", vbInformation, "Aviso"
'        R.Close
'        Set R = Nothing
'        Exit Sub
'    End If
'    R.Close
'    sql = "Select nTasaMorat From " & gcCentralCom & "LineaCredito Where cCodLinCred='" & Codlin & "'"
'    R.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'        Mora = R!nTasaMorat
'    R.Close
'    'carga el plandespag
'    sql = "Select * from Plandespag Where cCodCta='" & CodCta & "' And cTipo='C' Order By cNroCuo"
'    R.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'        ReDim MPlan(Cont)
'        Do While Not R.EOF
'            Cont = Cont + 1
'            ReDim Preserve MPlan(Cont)
'            MPlan(Cont - 1).FecVenc = R!dFecVenc
'            MPlan(Cont - 1).Capital = R!nCapital
'            MPlan(Cont - 1).Interes = R!nInteres
'            MPlan(Cont - 1).Mora = 0
'            MPlan(Cont - 1).CapPag = 0
'            MPlan(Cont - 1).IntPag = 0
'            MPlan(Cont - 1).MoraPag = 0
'            MPlan(Cont - 1).Estado = "P"
'            MPlan(Cont - 1).NumCuota = R!cNroCuo
'            MPlan(Cont - 1).DiasAtr = 0
'            R.MoveNext
'        Loop
'    R.Close
'    'Carga el kardex
'    sql = "Select * From Kardex Where cCodCta='" & CodCta & "' And cCodOpe<>'" & gsDesEfect & "' And cCodOpe<>'" & gsDesCCN & "' And cCodOpe<>'" & gsDesCCE & "' And cCodOpe<>'" & gsDesCheque & "' And cCodOpe<>'" & gsDesRefEfect & "' Order By dFecTra"
'    R.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'    rtf.Text = ""
'    Call EncabCalMora(CodCta)
'    PosC = 0
'    PosCP = 0
'    Do While Not R.EOF
'        Monto = R!nMonTran - R!nOtrGas
'        Monto = Monto - R!nIntMor
'        rtf.Text = rtf.Text + oImpresora.gPrnSaltoLinea
'        'Fecha de Pago
'        rtf.Text = rtf.Text + Space(5) + ImpreFormat(Format(R!dFecTra, "dd/mm/yyyy"), 12, 0, False)
'        ACuota = 0
'        AMora = 0
'        IntKar = IIf(IsNull(R!nIntComp), 0, R!nIntComp)
'        CapKar = IIf(IsNull(R!nCapital), 0, R!nCapital)
'        PosC = PosCP
'        Do While Monto > 0
'            'Cuota que Pago
'            rtf.Text = rtf.Text + ImpreFormat(MPlan(PosC).NumCuota, 4, 0, False)
'            'Fecha que vencia la cuota
'            rtf.Text = rtf.Text + Space(5) + ImpreFormat(Format(MPlan(PosC).FecVenc, "dd/mm/yyyy"), 12, 0, False)
'            'rtf.Text = rtf.Text + ImpreFormat(Format(MPlan(PosC).FecVenc, "dd/mm/yyyy"), 12, 0, False)
'            VDAtr = R!dFecTra - MPlan(PosC).FecVenc
'            'Dias de Mora Acumulado
'            rtf.Text = rtf.Text + ImpreFormat(VDAtr, 10, 0, False)
'            'Dias de Mora real
'            rtf.Text = rtf.Text + ImpreFormat(VDAtr - MPlan(PosC).DiasAtr, 5, 0, False)
'            DMora = VDAtr - IIf(MPlan(PosC).DiasAtr < 0, 0, MPlan(PosC).DiasAtr)
'            MPlan(PosC).DiasAtr = IIf(VDAtr < 0, 0, VDAtr)
'            'Monto de la cuota (Monto en Mora)
'            If DMora > 0 Then
'                rtf.Text = rtf.Text + ImpreFormat((MPlan(PosC).Capital - MPlan(PosC).CapPag) + (MPlan(PosC).Interes - MPlan(PosC).IntPag), 14, 2, True)
'            Else
'                rtf.Text = rtf.Text + ImpreFormat(0, 14, 2, True)
'            End If
'            'Capital de la Cuota
'            If DMora > 0 Then
'                rtf.Text = rtf.Text + ImpreFormat(MPlan(PosC).Capital - MPlan(PosC).CapPag, 9, 2, True)
'            Else
'                rtf.Text = rtf.Text + ImpreFormat(0, 9, 2, True)
'            End If
'            'Recargos por Cuota
'            If DMora > 0 Then
'                MPlan(PosC).Mora = Mora * DMora * (MPlan(PosC).Capital - MPlan(PosC).CapPag) / 100
'            Else
'                MPlan(PosC).Mora = 0
'            End If
'            'Mora de la Cuota
'            rtf.Text = rtf.Text + ImpreFormat(CDbl(Format(MPlan(PosC).Mora, "#0.00")), 9, 2, True)
'            'Distribucion en PlanDespag
'            'Interes
'                If IntKar >= (MPlan(PosC).Interes - MPlan(PosC).IntPag) Then
'                    IntKar = IntKar - (MPlan(PosC).Interes - MPlan(PosC).IntPag)
'                    Monto = Monto - (MPlan(PosC).Interes - MPlan(PosC).IntPag)
'                    MPlan(PosC).IntPag = MPlan(PosC).IntPag + (MPlan(PosC).Interes - MPlan(PosC).IntPag)
'                    Monto = CDbl(Format(Monto, "#0.00"))
'                Else
'                    Monto = Monto - IntKar
'                    MPlan(PosC).IntPag = MPlan(PosC).IntPag + IntKar
'                    IntKar = 0
'                    Monto = CDbl(Format(Monto, "#0.00"))
'                End If
'            'Capital
'                If CapKar >= (MPlan(PosC).Capital - MPlan(PosC).CapPag) Then
'                    CapKar = CapKar - (MPlan(PosC).Capital - MPlan(PosC).CapPag)
'                    Monto = Monto - (MPlan(PosC).Capital - MPlan(PosC).CapPag)
'                    MPlan(PosC).CapPag = MPlan(PosC).CapPag + (MPlan(PosC).Capital - MPlan(PosC).CapPag)
'                    MPlan(PosC).Estado = "G"
'                    Monto = CDbl(Format(Monto, "#0.00"))
'                    PosCP = PosCP + 1
'                Else
'                    MPlan(PosC).CapPag = MPlan(PosC).CapPag + CapKar
'                    Monto = Monto - CapKar
'                    CapKar = 0
'                    Monto = CDbl(Format(Monto, "#0.00"))
'                End If
'                PosC = PosC + 1
'                If Monto > 0 Then
'                    rtf.Text = rtf.Text + oImpresora.gPrnSaltoLinea
'                    rtf.Text = rtf.Text + Space(17)
'                End If
'        Loop
'        'A mora
'        rtf.Text = rtf.Text + ImpreFormat(R!nIntMor, 6, 2, True)
'        'Acuotas
'        rtf.Text = rtf.Text + ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nIntComp), 0, R!nIntComp), 7, 2, True)
'        'Monto de pago
'        rtf.Text = rtf.Text + ImpreFormat(R!nMonTran, 10, 2, True)
'
'        R.MoveNext
'    Loop
'    R.Close
'    frmPrevio.Previo rtf, "REPORTE CALCULO MORA  DE CREDITO", True, 66
End Sub

Private Sub cmdAceptar_Click()
Unload Me
End Sub

Public Sub frmIni(psCodCta As String, psEstado As String)
    gcCentralPers = "dbPersona.dbo."
    gcCentralCom = "dbComunes.dbo."
    lsCodCred = psCodCta
    lsEstado = psEstado
    lblCodCred.Text = lsCodCred
    lblCodCred.Enabled = False
    cmdCancelar.Visible = False
    Me.KeyPreview = False
    Call Conexiones
    Screen.MousePointer = 0
    lbIni = True
    Me.Show 1
End Sub

Private Sub CmdCalMora_Click()
    Call CalculaMora(lblCodCred.Text)
End Sub

Private Sub cmdCancelar_Click()
Limpiar
lblCodCred.Enfoque 2
End Sub

Private Sub CmdImprimir_Click()
Dim oPrevio As clsPrevio
Set oPrevio = New clsPrevio
If Me.listaClientes.ListItems.Count > 0 Then
    rtf.Text = ""
    lnNumLineas = 0
    ImprimirPstaCred
    oPrevio.Show rtf.Text, "REPORTE GENERAL DE CREDITO", True, 66
Else
    MsgBox "Por Favor Ingrese una cuenta de Crédito Válida", vbInformation, "Aviso"
End If

End Sub
Private Sub ImprimirPstaCred()
Dim i As Integer
Dim Ancho As Long
Ancho = 65
Encabezado

rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "CLIENTES RELACIONADOS" + Space(95 - Len("DATOS GENERALES DEL CREDITO")) & "DATOS GENERALES DEL CREDITO" + oImpresora.gPrnBoldOFF + oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & String(Ancho, "-") & Space(9) & String(Ancho, "-") & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & "NOMBRE" & Space(43) & "RELACIONES" & Space(15) & "Linea de Credito     : " & Me.lblLinea + oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & String(Ancho, "-") & Space(9) & "Fuente de Ingreso    : " & lblfuente & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
For i = 1 To listaClientes.ListItems.Count
    If i = 1 Then
        rtf.Text = rtf.Text & Space(5) & CadDerecha(PstaNombre(listaClientes.ListItems(i), False), 50) & CadDerecha(listaClientes.ListItems(i).SubItems(1), 10)
        rtf.Text = rtf.Text & Space(14) & "Analista Responsable : " & lblanalista & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    End If
    If i = 2 Then
        rtf.Text = rtf.Text & Space(5) & CadDerecha(PstaNombre(listaClientes.ListItems(i), False), 50) & CadDerecha(listaClientes.ListItems(i).SubItems(1), 10)
        rtf.Text = rtf.Text & Space(14) & "Apoderado            : " & lblapoderado & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    End If
    If i = 3 Then
        rtf.Text = rtf.Text & Space(5) & CadDerecha(PstaNombre(listaClientes.ListItems(i), False), 50) & CadDerecha(listaClientes.ListItems(i).SubItems(1), 10)
        rtf.Text = rtf.Text & Space(14) & "Condición de Crédito : " & CadDerecha(lblcondicion, 15) & Space(7) & "Nota 1  : " & lblnota1 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    End If
    If i = 4 Then
        rtf.Text = rtf.Text & Space(5) & CadDerecha(PstaNombre(listaClientes.ListItems(i), False), 50) & CadDerecha(listaClientes.ListItems(i).SubItems(1), 10)
        rtf.Text = rtf.Text & Space(14) & "Destino del Crédito  : " & CadDerecha(lbldestino, 20) & Space(3) & "Nota 2 : " & lblnota2 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    End If
    If i = 5 Then
        rtf.Text = rtf.Text & Space(5) & CadDerecha(PstaNombre(listaClientes.ListItems(i), False), 50) & CadDerecha(listaClientes.ListItems(i).SubItems(1), 10)
        rtf.Text = rtf.Text & Space(14) & "Tipo de Cuota        : " & CadDerecha(Me.lbltipocuota, 20) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    End If

Next
If i > 5 Then
    rtf.Text = rtf.Text & Space(5) & String(150, "-") & oImpresora.gPrnSaltoLinea
End If
Select Case i
    Case 1
        rtf.Text = rtf.Text & Space(79) & "Analista Responsable : " & lblanalista & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Apoderado            : " & lblapoderado & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Condición de Crédito : " & CadDerecha(lblcondicion, 15) & Space(8) & "Nota 1 : " & lblnota1 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Destino del Crédito  : " & CadDerecha(lbldestino, 20) & Space(3) & "Nota 2 : " & lblnota2 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Tipo de Cuota        : " & CadDerecha(Me.lbltipocuota, 20) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    Case 2
        rtf.Text = rtf.Text & Space(79) & "Apoderado            : " & lblapoderado & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Condición de Crédito : " & CadDerecha(lblcondicion, 15) & Space(8) & "Nota 1 : " & lblnota1 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Destino del Crédito  : " & CadDerecha(lbldestino, 20) & Space(3) & "Nota 2 : " & lblnota2 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Tipo de Cuota        : " & CadDerecha(Me.lbltipocuota, 20) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    Case 3
        rtf.Text = rtf.Text & Space(79) & "Condición de Crédito : " & CadDerecha(lblcondicion, 15) & Space(8) & "Nota 1 : " & lblnota1 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Destino del Crédito  : " & CadDerecha(lbldestino, 20) & Space(3) & "Nota 2 : " & lblnota2 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Tipo de Cuota        : " & CadDerecha(Me.lbltipocuota, 20) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    Case 4
        rtf.Text = rtf.Text & Space(79) & "Destino del Crédito  : " & CadDerecha(lbldestino, 20) & Space(3) & "Nota 2 : " & lblnota2 & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        rtf.Text = rtf.Text & Space(79) & "Tipo de Cuota        : " & CadDerecha(Me.lbltipocuota, 20) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
    Case 5
        rtf.Text = rtf.Text & Space(79) & "Tipo de Cuota        : " & CadDerecha(Me.lbltipocuota, 20) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1

End Select

If Val(lblmontosolicitud) > 0 Then
    'rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    'lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "HISTORIAL DEL CREDITO" + oImpresora.gPrnBoldOFF + oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & Space(13) & " FECHA " & Space(10) & " MONTO " & Space(10) & "Nº CUOTAS" & Space(16) & "PLAZO" & Space(10) & "CUOTA" & Space(10) & "P.GRACIA" & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "SOLICITADO : " & lblfechsolicitud & Space(15 - Len(Format(lblmontosolicitud, "#,#0.00"))) & Format(lblmontosolicitud, "#,#0.00") & Space(15 - Len(Format(lblcuotasSolicitud, "#0"))) & Format(lblcuotasSolicitud, "#0") & Space(25 - Len(Format(lblPlazoSolicitud, "#0") & " Días")) & Format(lblPlazoSolicitud, "#0") & " Días" & oImpresora.gPrnSaltoLinea
    If Val(lblMontosugerido) > 0 Then
        rtf.Text = rtf.Text & Space(5) & "SUGERIDO   : " & Space(25 - Len(Format(lblMontosugerido, "#,#0.00"))) & Format(lblMontosugerido, "#,#0.00") & Space(15 - Len(Format(lblcuotasugerida, "#0"))) & Format(lblcuotasugerida, "#0") & Space(25 - Len(Format(lblPlazoSugerido, "#0") & " Días")) & Format(lblPlazoSugerido, "#0") & " Días" & Space(15 - Len(Format(lblmontoCuotaSugerida, "#,#0.00"))) & Format(lblmontoCuotaSugerida, "#,#0.00") & Space(10 - Len(Format(lblGraciasugerida, "#0"))) & Format(lblGraciasugerida, "#0") & oImpresora.gPrnSaltoLinea
    End If
    If Val(lblMontoAprobado) > 0 Then
        rtf.Text = rtf.Text & Space(5) & "APROBADO   : " & lblfechaAprobado & Space(15 - Len(Format(lblMontoAprobado, "#,#0.00"))) & Format(lblMontoAprobado, "#,#0.00") & Space(15 - Len(Format(lblcuotasAprobado, "#0"))) & Format(lblcuotasAprobado, "#0") & Space(25 - Len(Format(lblPlazoAprobado, "#0") & " Días")) & Format(lblPlazoAprobado, "#0") & " Días" & Space(15 - Len(Format(lblmontoCuotaAprobada, "#,#0.00"))) & Format(lblmontoCuotaAprobada, "#,#0.00") & Space(10 - Len(Format(lblGraciaAprobada, "#0"))) & Format(lblGraciaAprobada, "#0") & oImpresora.gPrnSaltoLinea
    End If
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
End If

If listaDesembolsos.ListItems.Count > 0 Then
    rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnBoldON & "DESEMBOLSOS REALIZADOS" & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "  FECHA      " & Space(17) & "MONTO" & Space(10) & "GASTOS" & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    For i = 1 To listaDesembolsos.ListItems.Count
        rtf.Text = rtf.Text & Space(5) & listaDesembolsos.ListItems(i) & Space(25 - Len(listaDesembolsos.ListItems(i).SubItems(2))) & Format(listaDesembolsos.ListItems(i).SubItems(2), "#,#0.00") & Space(15 - Len(listaDesembolsos.ListItems(i).SubItems(3))) & Format(listaDesembolsos.ListItems(i).SubItems(3), "#,#0.00") & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
       ' If lnNumLineas > 50 Then
       '     rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
       '     lnNumLineas = 0
       '     Encabezado
       '     rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "DESEMBOLSOS REALIZADOS" + oImpresora.gPrnBoldOFF  + oImpresora.gPrnSaltoLinea
       '     lnNumLineas = lnNumLineas + 1
       '     rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
       '     lnNumLineas = lnNumLineas + 1
       '     rtf.Text = rtf.Text & Space(5) & "  FECHA      " & Space(17) & "MONTO" & Space(10) & "GASTOS" & oImpresora.gPrnSaltoLinea
       '     lnNumLineas = lnNumLineas + 1
       '     rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
       '     lnNumLineas = lnNumLineas + 1
       ' End If
    Next i
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    If lnNumLineas > 50 Then
        rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
        lnNumLineas = 0
        Encabezado
    End If
End If
If ListaPagos.ListItems.Count > 0 Then
    rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnBoldON & "PAGOS REALIZADOS" & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "  FECHA" & Space(8) & "CUOTA" & Space(5) & "MONTO PAG" & Space(8) & "CAPITAL" & Space(7) & "INTERES" & Space(10) & "MORA" & Space(10) & "GASTOS" & Space(5) & "DIAS ATR." & Space(2) & "Saldo Cap." & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    For i = 1 To ListaPagos.ListItems.Count
        rtf.Text = rtf.Text & Space(5) & ListaPagos.ListItems(i) & Space(10 - Len(Format(ListaPagos.ListItems(i).SubItems(1), "#,#0.00"))) & Format(ListaPagos.ListItems(i).SubItems(1), "#0") & Space(15 - Len(Format(ListaPagos.ListItems(i).SubItems(2), "#,#0.00"))) & Format(ListaPagos.ListItems(i).SubItems(2), "#,#0.00") & Space(15 - Len(Format(ListaPagos.ListItems(i).SubItems(3), "#,#0.00"))) & Format(ListaPagos.ListItems(i).SubItems(3), "#,#0.00") _
                    & Space(15 - Len(Format(ListaPagos.ListItems(i).SubItems(4), "#,#0.00"))) & Format(ListaPagos.ListItems(i).SubItems(4), "#,#0.00") & Space(15 - Len(Format(ListaPagos.ListItems(i).SubItems(5), "#,#0.00"))) & Format(ListaPagos.ListItems(i).SubItems(5), "#,#0.00") & Space(15 - Len(Format(ListaPagos.ListItems(i).SubItems(6), "#,#0.00"))) & Format(ListaPagos.ListItems(i).SubItems(6), "#,#0.00") & Space(10 - Len(Format(ListaPagos.ListItems(i).SubItems(7), "#0"))) & Format(ListaPagos.ListItems(i).SubItems(7), "#0") & ImpreFormat(CCur(ListaPagos.ListItems(i).SubItems(8)), 12, , True) & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        If lnNumLineas > 54 Then
            rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
            lnNumLineas = 0
            Encabezado
            rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnBoldON & "PAGOS REALIZADOS" & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
            rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
            rtf.Text = rtf.Text & Space(5) & "  FECHA" & Space(8) & "CUOTA" & Space(5) & "MONTO PAG" & Space(8) & "CAPITAL" & Space(7) & "INTERES" & Space(10) & "MORA" & Space(10) & "GASTOS" & Space(5) & "DIAS ATR." & Space(2) & "Saldo Cap." & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
            rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
        End If
    Next i
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1

    rtf.Text = rtf.Text & Space(5) & "TOTAL PAGADO :" & Space(33 - Len(Format(lblcapitalpagado, "#,#0.00"))) & Format(lblcapitalpagado, "#,#0.00") & Space(15 - Len(Format(lblintcompPag, "#,#0.00"))) & Format(lblintcompPag, "#,#0.00") & Space(15 - Len(Format(lblIntMorPag, "#,#0.00"))) & Format(lblIntMorPag, "#,#0.00") & Space(15 - Len(Format(lblGastopagado, "#,#0.00"))) & Format(lblGastopagado, "#,#0.00") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    If lnNumLineas > 50 Then
        rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
        lnNumLineas = 0
        Encabezado
    End If
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
   ' If lnNumLineas > 50 Then
   '     rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
   '     lnNumLineas = 0
   '     Encabezado
   ' End If
End If
If lstCuotasPend.ListItems.Count > 0 Then
    rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "CUOTAS PENDIENTES" + oImpresora.gPrnBoldOFF + oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "  CUOTA" & Space(3) & "FECHA VENC." & Space(8) & "CAPITAL" & Space(9) & "INTERES" & Space(9) & "MORA" & Space(5) & "DIAS ATR." & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    For i = 1 To lstCuotasPend.ListItems.Count
        rtf.Text = rtf.Text & Space(10 - Len(Format(lstCuotasPend.ListItems(i), "#0"))) & Format(lstCuotasPend.ListItems(i), "#0") & Space(5) & lstCuotasPend.ListItems(i).SubItems(1) & Space(15 - Len(Format(lstCuotasPend.ListItems(i).SubItems(2), "#,#0.00"))) & Format(lstCuotasPend.ListItems(i).SubItems(2), "#,#0.00") & Space(15 - Len(Format(lstCuotasPend.ListItems(i).SubItems(3), "#,#0.00"))) & Format(lstCuotasPend.ListItems(i).SubItems(3), "#,#0.00") _
                    & Space(15 - Len(Format(lstCuotasPend.ListItems(i).SubItems(4), "#,#0.00"))) & Format(lstCuotasPend.ListItems(i).SubItems(4), "#,#0.00") & Space(10 - Len(Format(lstCuotasPend.ListItems(i).SubItems(5), "#0"))) & Format(lstCuotasPend.ListItems(i).SubItems(5), "#0") & oImpresora.gPrnSaltoLinea
        lnNumLineas = lnNumLineas + 1
        If lnNumLineas > 54 Then
            rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
            lnNumLineas = 0
            Encabezado
            rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "CUOTAS PENDIENTES" + oImpresora.gPrnBoldOFF + oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
            rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
            rtf.Text = rtf.Text & Space(5) & "  CUOTA" & Space(3) & "FECHA VENC." & Space(8) & "CAPITAL" & Space(9) & "INTERES" & Space(9) & "MORA" & Space(5) & "DIAS ATR." & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
            rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
            lnNumLineas = lnNumLineas + 1
        End If
    Next i
    rtf.Text = rtf.Text & Space(5) & String(140, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1

    rtf.Text = rtf.Text & Space(25) & ImpreFormat(CDbl(Me.lblSaldoKCalend), 12, 2, True)
    rtf.Text = rtf.Text & ImpreFormat(CDbl(Me.lblIntCompCalend), 12, 2, True) & ImpreFormat(CDbl(Me.lblIntMorCalend), 12, 2, True)
    rtf.Text = rtf.Text & ImpreFormat(CDbl(Me.lblGastoCalend), 12, 2, True)
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    If lnNumLineas > 50 Then
        rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
        lnNumLineas = 0
        Encabezado
    End If
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    'If lnNumLineas > 50 Then
    '    rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
    '    lnNumLineas = 0
    '    Encabezado
    'End If
    If lnNumLineas >= 44 And lnNumLineas <= 52 Then
        rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
        lnNumLineas = 0
        Encabezado
    End If
    rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON & "CUOTAS ATRASADAS" + Space(65) + "PAGO AL " & Format(gdFecSis, "dd/mm/yyyy") + oImpresora.gPrnBoldOFF + oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(70, "-") + Space(10) & String(60, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "CAPITAL         :" & ImpreFormat(CDbl(Me.lblCapitalCuoPend), 15, 2, True)
    rtf.Text = rtf.Text & Space(45) & "CAPITAL         : " & ImpreFormat(CDbl(Me.lblSaldoKFecha), 15, 2, True) & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "INTERES         :" & ImpreFormat(CDbl(Me.lblInteresCuoPend), 15, 2, True)
    rtf.Text = rtf.Text & Space(45) & "INTERES         :" & ImpreFormat(CDbl(Me.lblIntCompFecha), 15, 2, True) & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "MORA            :" & ImpreFormat(CDbl(Me.lblMoraCuoPend), 15, 2, True)
    rtf.Text = rtf.Text & Space(45) & "MORA            :" & ImpreFormat(CDbl(Me.lblIntMorFecha), 15, 2, True) & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & "GASTOS          :" & ImpreFormat(CDbl(Me.lblGastoCuoPend), 15, 2, True)
    rtf.Text = rtf.Text & Space(45) & "GASTOS          :" & ImpreFormat(CDbl(Me.lblGastoFecha), 15, 2, True) & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(85) & "INT SUSPENSO    :" & ImpreFormat(CDbl(Me.lblIntSuspFecha), 15, 2, True) & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) & String(70, "-") + Space(10) & String(60, "-") & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "TOTAL           :" & ImpreFormat(CDbl(Me.lblTotalCuoPend), 15, 2, True)
    rtf.Text = rtf.Text & Space(45) & "TOTAL           :" & ImpreFormat(CDbl(Me.lblTotalFecha), 15, 2, True) & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    rtf.Text = rtf.Text & oImpresora.gPrnSaltoLinea
    lnNumLineas = lnNumLineas + 1
    If lnNumLineas > 50 Then
        rtf.Text = rtf.Text & oImpresora.gPrnSaltoPagina
        lnNumLineas = 0
        Encabezado
    End If
    
End If
rtf.Text = rtf.Text & Space(5) + oImpresora.gPrnBoldON + "OTROS DATOS" + oImpresora.gPrnBoldOFF + oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(10) & "PROTESTO          : " & IIf(chkProtesto.value = 1, "Si", "No") & oImpresora.gPrnSaltoLinea
rtf.Text = rtf.Text & Space(10) & "CARGO AUTOMATICO  : " & IIf(chkCargoAuto.value = 1, "Si", "No") & oImpresora.gPrnSaltoLinea
rtf.Text = rtf.Text & Space(10) & "REFINANCIADO      : " & IIf(chkRefinanciado.value = 1, "Si", "No") & oImpresora.gPrnSaltoLinea

End Sub
Private Sub Encabezado()

rtf.Text = rtf.Text & "@" & Space(4) & gsNomCmac & Space(85) & "Fecha : " & Format(gdFecSis + Time, "dd/mm/yyyy hh:mm") & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & gsNomAge & Space(96) & "Usuario: " & gsCodUser & "  " & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & CentrarCadena("R E S U M E N   G E N E R A L   D E   C R E D I T O", 150) & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnBoldON & "Crédito Nº :" & Trim(lblCodCred.Text) + oImpresora.gPrnBoldOFF & Space(10) & "Tipo : " & CadDerecha(lbltipoCredito, 50) & Space(2) & "Estado :" & lblEstado & oImpresora.gPrnSaltoLinea
lnNumLineas = lnNumLineas + 1
rtf.Text = rtf.Text & Space(5) & oImpresora.gPrnSaltoLinea
End Sub

Private Sub Form_Activate()
AbreConexion
End Sub

Private Sub Form_Load()
AbreConexion
ReDim MatPagos(0)
ReDim MatGastos(0)
ContCuotas = 0
ContGastos = 0
ReDim MatSaldos(0)
nSaldos = 0
SSTab1.Tab = 0
Me.lblCodCred.EnabledAge = True
If lbIni = False Then
    Me.KeyPreview = True
    lblCodCred.Enabled = True
    lblCodCred.Text = Right(gsCodAge, 2)
End If
End Sub

Private Function TotalDeudaLibre() As Double
Dim Dias As Integer
    Dias = gdFecSis - CDate(Format(dFecUltPago, "dd/mm/yyyy"))
    TotalDeudaLibre = InteresReal(nTasaInt / 100, Dias) * nSaldoCap + vIntPend + nSaldoCap
End Function

Private Function TotalInteresaFechaLibre() As Double
Dim Dias As Integer
    Dias = gdFecSis - CDate(Format(dFecUltPago, "dd/mm/yyyy"))
    TotalInteresaFechaLibre = InteresReal(nTasaInt / 100, Dias) * nSaldoCap + vIntPend
End Function

Private Sub DatosGenerales(psCodCta As String, dbConexion As DConecta)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim lblproxCuota  As String
Dim TipCuota As String
lblCodCred.Enabled = False

Sql = "SELECT Credito.nIntPend,Credito.dFecUltPago, Credito.cTipCuota, Credito.cCodCta, Credito.cEstado, " _
    & "FuenteIngreso.cRazonSocial AS DescFte," _
    & "FuenteIngreso.cNumFuente AS CodFte, Credito.dAsignacion, Credito.nMontoSol, Credito.nCuotasSol, " _
    & "Credito.nPlazoSol, Credito.cNumFuente, Credito.cCondCre, " _
    & "Credito.cDestCre, Credito.cCodLinCred, Credito.nMontoSug, " _
    & "Credito.nCuotasSug, Credito.nPlazoSug, Credito.nGraciaSug, Credito.nCuotaSug, " _
    & "Credito.nTasaInt, Credito.dFecApr, Credito.nMontoApr, Credito.nIntApr, Credito.nCuotasApr, " _
    & "Credito.nPlazoApr, Credito.nGraciaApr, Credito.cTipGraciaApr, Credito.nCuotaApr, Credito.ctipCuota, Credito.cperiodo, " _
    & "Credito.cCodAnalista, Credito.cApoderado, Credito.cCauRech, Credito.dFecVig, " _
    & "Credito.cTipoDesemb, Credito.nMontoDesemb, Credito.nNroProxDesemb, Credito.dFecUltDesemb, " _
    & "Credito.nSaldoCap, Credito.nCapPag, Credito.nIntComPag, Credito.nIntMorPag, " _
    & "Credito.nGastoPag, Credito.nDiasAtraso, Credito.nIntMorCal, Credito.nNroProxCuota, " _
    & "Credito.dFecUltPago, Credito.cMetLiquid, Credito.cDescarAuto, Credito.cFlagProtesto, " _
    & "Credito.cLibAmort, Credito.dCancelado, Credito.cComenta," _
    & "Credito.dJudicial, Credito.cNota1,Credito.cNota2, Credito.nDiAtrAcu, Credito.cCodInst, " _
    & "Credito.cCodModular, Credito.cCalifica, Credito.nDiaFijo,Credito.cRefinan,Credito.nNroRepro  " _
    & "FROM Credito INNER JOIN " _
    & "FuenteIngreso ON Credito.cNumFuente = FuenteIngreso.cNumFuente " _
    & "WHERE Credito.cCodCta='" & Trim(lsCodCred) & "'"
    
Screen.MousePointer = 11
Set rs = dbConexion.CargaRecordSet(Sql)
If RSVacio(rs) Then
    MsgBox "Datos no encontrados", vbInformation, "Aviso"
    Screen.MousePointer = 0
Else
    DoEvents
    If rs!cEstado = "F" Or rs!cEstado = "G" Then
       If IsNull(rs!dFecUltPago) Then
          dFecUltPago = rs!dfecUltDesemb
       Else
          dFecUltPago = IIf(rs!dfecUltDesemb > rs!dFecUltPago, rs!dfecUltDesemb, rs!dFecUltPago)
       End If
    Else
       dFecUltPago = rs!dAsignacion
    End If
  ' IIf(IsNull(RS!dFecUltPago), IIf(IsNull(RS!dfecUltDesemb), RS!dAsignacion, IIf(RS!dfecUltDesemb > RS!dFecUltPago, RS!dfecUltDesemb, RS!dFecUltPago)), RS!dFecUltPago)
    nTasaInt = IIf(IsNull(rs!nTasaInt), 0, rs!nTasaInt)
    nSaldoCap = IIf(IsNull(rs!nSaldoCap), 0, rs!nSaldoCap)
    vIntPend = IIf(IsNull(rs!nIntPend), 0, rs!nIntPend)
    vIntMorCal = IIf(IsNull(rs!nIntMorCal), 0, rs!nIntMorCal)
    TipCuota = IIf(IsNull(rs!ctipcuota), 0, rs!ctipcuota)
    CargaClientes Trim(lsCodCred), dbConexion
    lbltipoCredito = AbrevProd(Mid(lsCodCred, 3, 3), False)
    lblfechsolicitud = Format(IIf(IsNull(rs!dAsignacion), "", rs!dAsignacion), "dd/mm/yyyy")
    lblmontosolicitud = Format(IIf(IsNull(rs!nMontoSol), 0, rs!nMontoSol), "#0.00")
    lblcuotasSolicitud = Format(IIf(IsNull(rs!nCuotasSol), 0, rs!nCuotasSol), "#0")
    lblPlazoSolicitud = Format(IIf(IsNull(rs!nPlazoSol), 0, rs!nPlazoSol), "#0")
    lblfuente = Trim(IIf(IsNull(rs!DescFte), "", rs!DescFte))
    lblcondicion = Tablacod("38", rs!cCondCre)
    lbldestino = Tablacod("39", rs!cDestCre)
    lblLinea = LineaCredito(Trim(IIf(IsNull(rs!cCodLinCred), "", rs!cCodLinCred)))
    lblfechasugerida = Format(IIf(IsNull(rs!dAsignacion), "", rs!dAsignacion), "dd/mm/yyyy")
    lblMontosugerido = Format(IIf(IsNull(rs!nMontoSug), 0, rs!nMontoSug), "#0.00")
    lblcuotasugerida = Format(IIf(IsNull(rs!nCuotasSug), 0, rs!nCuotasSug), "#0")
    lblPlazoSugerido = Format(IIf(IsNull(rs!nPlazoSug), 0, rs!nPlazoSug), "#0")
    lblGraciasugerida = Format(IIf(IsNull(rs!nGraciaSug), 0, rs!nGraciaSug), "#0")
    lblmontoCuotaSugerida = Format(IIf(IsNull(rs!nCuotaSug), 0, rs!nCuotaSug), "#0.00")
    lblfechaAprobado = Format(IIf(IsNull(rs!dFecApr), "", rs!dFecApr), "dd/mm/yyyy")
    lbltasainteres = Format(IIf(IsNull(rs!nTasaInt), 0, rs!nTasaInt), "#0.00")
    lblcuotasAprobado = Format(IIf(IsNull(rs!nCuotasApr), 0, rs!nCuotasApr), "#0")
    lblPlazoAprobado = Format(IIf(IsNull(rs!nplazoapr), 0, rs!nplazoapr), "#0")
    lblGraciaAprobada = Format(IIf(IsNull(rs!nGraciaApr), 0, rs!nGraciaApr), "#0")
    Select Case rs!cTipGraciaApr
        Case "0"
          If lblGraciaAprobada <> "0" Then
            lblTipoGraciaApr = "Primera Cuota"
          End If
        Case "1"
          lblTipoGraciaApr = "Ultima Cuota"
        Case "2"
          lblTipoGraciaApr = "Prorrateada"
        Case "3"
          lblTipoGraciaApr = "Configurada"
        Case Else
          lblTipoGraciaApr = ""
    End Select
    lblmontoCuotaAprobada = Format(IIf(IsNull(rs!nCuotaApr), 0, rs!nCuotaApr), "#0.00")
    lblMontoAprobado = Format(IIf(IsNull(rs!nMontoApr), 0, rs!nMontoApr), "#0.00")
    Me.txtComenta = IIf(IsNull(rs!cComenta), "", rs!cComenta)
    
    If Not IsNull(rs!dfecUltDesemb) Then FecUltPago = Format(rs!dfecUltDesemb, "dd/mm/yyyy")
    lblEstado = Tablacod("26", rs!cEstado, True)
    'Tasa de Interes
    lsEstado = Trim(rs!cEstado)
    lsEstadoCredito = Trim(rs!cEstado)
    If Trim(rs!cEstado) >= "F" Then
        If rs!cPeriodo = "2" Then
            Periodo = 30
            TasaInt = Format(InteresReal(Format(rs!nTasaInt / 100, "#0.000"), 30) * 100, "#0.00")
        Else
            Periodo = IIf(IsNull(rs!nplazoapr), 0, rs!nplazoapr)
            TasaInt = Format(InteresReal(Format(rs!nTasaInt / 100, "#0.000"), IIf(IsNull(rs!nplazoapr), 0, rs!nplazoapr)) * 100, "#0.00")
        End If
    Else
        Periodo = 0
        TasaInt = 0
    End If
    If IsNull(rs!cTipoDesemb) Then
        lbltipoDesembolso = ""
    Else
        lbltipoDesembolso = IIf(Trim(rs!cTipoDesemb) = "T", "TOTAL", "PARCIAL")
    End If
    lblmontoDesembolsado = Format(IIf(IsNull(rs!nMontoDesemb), 0, rs!nMontoDesemb), "#0.00")
    lnNroRepro = IIf(IsNull(rs!nNroRepro), 0, rs!nNroRepro)
    Select Case Trim(IIf(IsNull(rs!ctipcuota), "", rs!ctipcuota))
           Case "1"
                lbltipocuota = "CUOTA FIJA"
           Case "2"
                lbltipocuota = "DECRECIENTE"
           Case "3"
                lbltipocuota = "CRECIENTE"
           Case "4"
                lbltipocuota = "CUOTA LIBRE"
    End Select
    Select Case Trim(IIf(IsNull(rs!cPeriodo), "", rs!cPeriodo))
        Case "1"
            lblperiodo = "PERIODO FIJO"
        Case "2"
            lblperiodo = "FECHA FIJA"
        Case "3"
            lblperiodo = "LIBRE"
    End Select
    lblanalista = Usuario(Trim(IIf(IsNull(rs!cCodAnalista), "", rs!cCodAnalista)), dbConexion)
    lblapoderado = Usuario(Trim(IIf(IsNull(rs!cApoderado), "", rs!cApoderado)), dbConexion, "0001")
    If Trim(IIf(IsNull(rs!cEstado), "", rs!cEstado)) = "L" Then
        lblrechazo = Tablacod("27", Trim(IIf(IsNull(rs!cCauRech), "", rs!cCauRech)))
    Else
        lblrechazo = ""
    End If
    lblmetodoLiquidacion = Trim(IIf(IsNull(rs!cMetLiquid), "GMIC", rs!cMetLiquid))
    lblfechavigencia = IIf(IsNull(rs!dFecVig), "", Format(rs!dFecVig, "dd/mm/yyyy"))
    lblProximoDesembolso = Format(IIf(IsNull(rs!nNroProxDesemb), 0, rs!nNroProxDesemb), "#0")
    lblMontoProxDesem = DatoProximo(psCodCta, lblProximoDesembolso, dbConexion, "1")
    lblMontoProxDesem = IIf(Len(lblMontoProxDesem) = 0, "0.00", lblMontoProxDesem)
    lblfechaproxdesem.Caption = Format(IIf(IsNull(rs!dfecUltDesemb), "", rs!dfecUltDesemb), "dd/mm/yyyy")
    lbltotalDesembolso = TotalDesembolso(psCodCta, dbConexion)
    lblcapitalpagado = Format(IIf(IsNull(rs!nCapPag), 0, rs!nCapPag), "#0.00")
    
    lblintcompPag = Format(IIf(IsNull(rs!nintcompag), 0, rs!nintcompag), "#0.00")
    lblIntMorPag = Format(IIf(IsNull(rs!nIntMorPag), 0, rs!nIntMorPag), "#0.00")
    lblGastopagado = Format(IIf(IsNull(rs!nGastoPag), 0, rs!nGastoPag), "#0.00")
    lblproxCuota = Format(IIf(IsNull(rs!nNroProxCuota), 0, rs!nNroProxCuota), "#0")
    chkProtesto.value = IIf(Val(IIf(IsNull(rs!cFlagProtesto), 0, rs!cFlagProtesto)) = 0, 0, 1)
    
    lblfechaCancelacion = Format(IIf(IsNull(rs!dCancelado), "", rs!dCancelado), "dd/mm/yyyy")
    lblfechajudicial = Format(IIf(IsNull(rs!dJudicial), "", rs!dJudicial), "dd/mm/yyyy")
    lblnota1 = Trim(IIf(IsNull(rs!cNota1), 0, rs!cNota1))
    lblnota2 = Trim(IIf(IsNull(rs!cNota2), 0, rs!cNota2))
    lsRefinanciado = Trim(IIf(IsNull(rs!cRefinan), "", rs!cRefinan))
    chkRefinanciado = IIf(Trim(lsRefinanciado) = "R", 1, 0)
    If lsRefinanciado = "R" Then
       Me.fraCreditosRefinanciado.Visible = True
       CargaRefinanciados psCodCta, dbConexion
    Else
       Me.fraCreditosRefinanciado.Visible = False
    End If
    
    CargaDesembolsos psCodCta, dbConexion
    CargaPagos psCodCta, CCur(lblmontoDesembolsado), dbConexion
    DatosPlanDesPag psCodCta, dbConexion
    '*********** Pagos Pendientes   ********
    Tipodesemb = Trim(IIf(IsNull(rs!cTipoDesemb), "", rs!cTipoDesemb))
    If Tipodesemb = "P" And TipCuota <> "4" Then
        CargaCuotas Trim(lsCodCred), dbConexion
        lblTotalFecha.Caption = TotalDeudaParcial(dbConexion)
        lblIntCompFecha = Format(TotalInteresaFecha, "#0.00")
    Else
        If rs!cEstado >= "F" And rs!cEstado <= "H" Then
            'CargaCuotas Trim(lsCodCred), dbConexion
            If TipCuota <> "4" Then
                If lblmetodoLiquidacion = "GMYC" Then
                    lblTotalFecha.Caption = Format(fgCalculaDeudaTotalaFechaGMYC(lsCodCred, dbConexion), "#0.00")
                    lblIntCompFecha = Format(fgCalculaInteresaFechaGMYC(lsCodCred, dbConexion), "#0.00")
                Else
                    lblTotalFecha.Caption = Format(fgCalculaDeudaTotalaFecha(lsCodCred, dbConexion), "#0.00")
                    lblIntCompFecha = Format(fgCalculaInteresaFecha(lsCodCred, dbConexion) + fgCalculaInteresReprogramados(lsCodCred, dbConexion), "#0.00")
                End If
            Else
                lblTotalFecha.Caption = Format(fgCalculaDeudaTotalaFechaCuotaLibre(lsCodCred, dbConexion), "#0.00")
                lblIntCompFecha = Format(fgCalculaInteresaFechaCuotaLibre(lsCodCred, dbConexion), "#0.00")
            End If
        Else
            lblTotalFecha.Caption = Format(0, "#0.00")
            lblIntCompFecha = Format(0, "#0.00")
        End If
    End If
    ' Imprime el Total
    Me.lblTotalFecha = Format(CDbl(Me.lblSaldoKFecha) + CDbl(Me.lblIntCompFecha) + CDbl(Me.lblIntMorFecha) + CDbl(Me.lblGastoFecha) + CDbl(Me.lblIntSuspFecha), "#0.00")
    
    '** Total Pendiente a la Fecha (Cuota Libre)
    If Me.lbltipocuota = "CUOTA LIBRE" Then ' Parche para mostrar IntComp a Fecha
       Me.lblInteresCuoPend = Format(CDbl(Me.lblIntCompFecha), "#0.00")
       Me.lblTotalCuoPend = Format(CDbl(Me.lblCapitalCuoPend) + CDbl(Me.lblIntCompFecha) + CDbl(Me.lblMoraCuoPend) + CDbl(Me.lblGastoCuoPend), "#0.00")
    End If

    
    ' ***** INSTITUCION **
    If Mid(lsCodCred, 3, 3) = "301" Then
        fraDescPlan.Visible = True
        lblinstitucion = Tablacod("55", Trim(IIf(IsNull(rs!cCodInst), "", rs!cCodInst)))
        lblmodular = Trim(IIf(IsNull(rs!cCodModular), "", rs!cCodModular))
    Else
        fraDescPlan.Visible = False
        lblinstitucion = ""
        lblmodular = ""
    End If
    chkCargoAuto.value = IIf(IIf(IsNull(rs!cDescarAuto), "N", Trim(rs!cDescarAuto)) = "N", 0, 1)

    '*************** GARANTIAS *******************
    Garantias Trim(lsCodCred), dbConexion
    
    'MsgBox "INTERES FECHA : " & fgCalculaInteresaFechaGMYC(lsCodCred, dbConexion) & Chr(13) & "  DEUDA FECHA : " & fgCalculaDeudaTotalaFechaGMYC(lsCodCred, dbConexion)
    
    DoEvents
End If
rs.Close
Set rs = Nothing
Screen.MousePointer = 0
Exit Sub
ErrorDatosGnrales:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description & Chr(13) & "Por Favor Consulte al Dpto. de Sisetmas", vbInformation, "Aviso"
    Screen.MousePointer = 0
End Sub

Private Sub Conexiones()
    Dim Agencia As String
    Dim Sql As String
    Dim lsFecha As String
    Dim j As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'On Error GoTo ErrorConexionPsta
    DoEvents
    Screen.MousePointer = 11
    Agencia = Trim(Mid(lsCodCred, 1, 2))
    oCon.AbreConexion 'Remota Agencia
    
    DatosGenerales lsCodCred, oCon
    
    Screen.MousePointer = 0
    
    Exit Sub
ErrorConexionPsta:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description & Chr(13) & "Por Favor Consulte al Dpto. de Sisetmas", vbInformation, "Aviso"
    Screen.MousePointer = 0
End Sub

Private Sub CargaClientes(psCodCred As String, poCon As DConecta)
Dim rsC As New ADODB.Recordset
Dim Sql As String
Dim Item As ListItem
    Sql = "SELECT Pe.cCodCta, P.cNomPers, Pe.cRelaCta, R.cCodTab, " _
        & " R.cNomTab AS DescRel, P.cCodPers " _
        & " FROM Credito C INNER JOIN " _
        & "PersCredito PE ON C.cCodCta = Pe.cCodCta INNER JOIN " _
        & "" & gcCentralPers & "Persona P ON Pe.cCodPers = P.cCodPers INNER JOIN " _
        & "" & gcCentralCom & "TablaCod R ON Pe.cRelaCta = R.cValor " _
        & "WHERE R.cCodTab LIKE '25__' AND Pe.cCodCta = '" & Trim(psCodCred) & " '"
Set rsC = poCon.CargaRecordSet(Sql)
If RSVacio(rsC) Then
    listaClientes.ListItems.Clear
Else
    listaClientes.ListItems.Clear
    Do While Not rsC.EOF
        Set Item = Me.listaClientes.ListItems.Add(, , rsC!cNomPers)
        Item.SubItems(1) = Trim(rsC!DescRel)
        Item.SubItems(2) = Trim(rsC!cCodPers)
        rsC.MoveNext
    Loop
End If
rsC.Close
Set rsC = Nothing
End Sub
Private Sub CargaDesembolsos(psCodCred As String, opCon As DConecta)
Dim Sql As String
Dim rsD As New ADODB.Recordset
Dim GastoCuota As String
Dim TotalPagado As String
Dim ItemD As ListItem  'item para desembolso
Dim ItemC As ListItem  'item para Cuotas pagadas
Dim lnMora As Long
Sql = "SELECT * FROM PlanDesPag " _
    & "WHERE PlanDesPag.cCodCta = '" & Trim(psCodCred) & "' AND PlanDesPag.cEstado = 'G' and cTipo='D' " _
    & "ORDER BY PlanDesPag.cNroCuo "

Set rsD = opCon.CargaRecordSet(Sql)

If RSVacio(rsD) Then
    listaDesembolsos.ListItems.Clear
Else
    listaDesembolsos.ListItems.Clear
    Do While Not rsD.EOF
        Set ItemD = listaDesembolsos.ListItems.Add(, , Format(rsD!dFecPag, "dd/mm/yyyy"))
        ItemD.SubItems(1) = Trim(rsD!cNroCuo)
        ItemD.SubItems(2) = Format(rsD!nCapital, "#0.00")
        ItemD.SubItems(3) = MontoGastos(Trim(rsD!cNroCuo), Trim(rsD!cTipo), Trim(psCodCred), opCon)
        rsD.MoveNext
    Loop
End If
rsD.Close
Set rsD = Nothing

End Sub
Private Sub CargaPagos(psCodCred As String, lnMontoDesemb As Currency, opCon As DConecta)
Dim Sql As String
Dim rsD As New ADODB.Recordset
Dim GastoCuota As String
Dim TotalPagado As String
Dim lsFechaPago As String
Dim ItemC As ListItem  'item para Cuotas pagadas
Dim lnMora As Long
Dim lnDiasMora As Long
Dim lnSaldoCap As Currency

Sql = "SELECT Kardex.cCodCta, " _
    & "Kardex.dFectra , Kardex.nNroCuota, " _
    & "Kardex.nDiasMora, Kardex.nMonTran, Kardex.nCapital, " _
    & "Kardex.nIntComp, Kardex.nIntMor, Kardex.nOtrGas, Kardex.nIntGraPag, " _
    & "Kardex.cCodOpe " _
    & "FROM Kardex " _
    & "WHERE (Kardex.cCodCta = '" & Trim(psCodCred) & "') AND " _
    & "(SUBSTRING(Kardex.cCodOpe,1,4) NOT IN ('" & Left("010102", 4) & "' ) AND KARDEX.CCODOPE NOT  IN ('019998')) " _
    & "ORDER BY KArdex.dFecTra "
    
lnSaldoCap = lnMontoDesemb
Set rsD = opCon.CargaRecordSet(Sql)
TotalPagado = 0
If RSVacio(rsD) Then
    ListaPagos.ListItems.Clear
Else
    ListaPagos.ListItems.Clear
    Do While Not rsD.EOF
        lnSaldoCap = lnSaldoCap - Format(IIf(IsNull(rsD!nCapital), 0, rsD!nCapital), "#0.00")  'Capital
        Set ItemC = ListaPagos.ListItems.Add(, , Format(rsD!dFecTra, "dd/mm/yyyy"))
        ItemC.SubItems(1) = IIf(IsNull(rsD!nNroCuota), "0", Trim(rsD!nNroCuota))  'numero de Cuota
        ItemC.SubItems(2) = Format(IIf(IsNull(rsD!nMonTran), 0, rsD!nMonTran), "#0.00")     'monto pagado
        ItemC.SubItems(3) = Format(IIf(IsNull(rsD!nCapital), 0, rsD!nCapital), "#0.00")    'Capital
        ItemC.SubItems(4) = Format(IIf(IsNull(rsD!nIntComp), 0, rsD!nIntComp) + IIf(IsNull(rsD!nIntGraPag), 0, rsD!nIntGraPag), "#0.00")  'Interés + Gracia
        ItemC.SubItems(5) = Format(IIf(IsNull(rsD!nIntMor), 0, rsD!nIntMor), "#0.00")   'mora
        ItemC.SubItems(6) = Format(IIf(IsNull(rsD!nOtrGas), 0, rsD!nOtrGas), "#0.00")     'Gastos
        ItemC.SubItems(7) = Format(IIf(IsNull(rsD!nDiasMora), 0, rsD!nDiasMora), "#0")  'Dias de Mora
        ItemC.SubItems(8) = Format(lnSaldoCap, "#0.00")  'Dias de Mora
        lnSaldoCap = lnSaldoCap
        rsD.MoveNext
    Loop
End If
rsD.Close
Set rsD = Nothing
End Sub

Private Function DatoProximo(psCodCta As String, lsCuota As String, opCon As DConecta, Optional psDato As String = "1", Optional psTipo As String = "D") As String
Dim Sql As String
Dim rsP As New ADODB.Recordset
Dim lsCadena As String
Sql = "SELECT * FROM PlanDesPag WHERE cCodCta='" & Trim(psCodCta) & "' " _
     & "AND cTipo='" & Trim(psTipo) & "' AND cEstado='P' and cNroCuo=" & Trim(lsCuota)

    Set rsP = opCon.CargaRecordSet(Sql)
    If RSVacio(rsP) Then
        lsCadena = ""
    Else
        Select Case Trim(psDato)
            Case "1" 'Monto
                lsCadena = Format(IIf(IsNull(rsP!nCapital), 0, rsP!nCapital), "#0.00")
            Case "2" 'Fecha de vencimiento
                lsCadena = Format(IIf(IsNull(rsP!dFecVenc), 0, rsP!dFecVenc), "dd/mm/yyyy")
        End Select
    End If
    rsP.Close
    Set rsP = Nothing
    DatoProximo = lsCadena
End Function
Private Function TotalDesembolso(psCodCta As String, opCon As DConecta) As String
Dim Sql As String
Dim rsP As New ADODB.Recordset

Sql = "SELECT SUM(nCapital) AS Sumatotal " _
    & "FROM Plandespag WHERE cCodCta='" & Trim(psCodCta) & "' AND cTipo='D' "
    
Set rsP = opCon.CargaRecordSet(Sql)
If RSVacio(rsP) Then
    TotalDesembolso = ""
Else
    TotalDesembolso = Format(IIf(IsNull(rsP!SumaTotal), 0, rsP!SumaTotal), "#0.00")
End If
rsP.Close
Set rsP = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub


Private Sub lblCodCred_keypressEnter()
'If Len(lblCodCred.Text) = 12 Then
'    lsCodCred = lblCodCred.Text
'    If Mid(lsCodCred, 1, 2) = Right(gsCodAge, 2) Then
'        DatosGenerales lsCodCred, dbCmact
'    Else
'        If AbreConeccion(lsCodCred, True) = True Then
'            DatosGenerales lsCodCred, dbCmactN
'            CierraConeccion
'        End If
'    End If
'Else
'    MsgBox "Cuenta de Crédito no Válida", vbInformation, "Aviso"
'    lblCodCred.Enfoque 1
'End If
End Sub

Private Sub DatosPlanDesPag(psCodCta As String, opCon As DConecta)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim lsCuotaMinPend As String
Dim lsCuotaMaxMora As String
Dim Item As ListItem
Dim lnMora As Double
Dim lsFecha As String
Dim rsGas As New ADODB.Recordset

lsFecha = Format(gdFecSis, "mm/dd/yyyy")
'--************Numero de cuotas en Mora Cuota Max y Cuota Min en Mora
Sql = "SELECT Min(cNroCuo) AS MinCuota, Max(cNroCuo) AS MaxCuota, " _
   & " Count(cCodCta) AS NumCuotas, cCodCta " _
   & " FROM PlanDesPag WHERE cEstado='P' AND cTipo ='C' And cCodcta='" & psCodCta & "' " _
   & " GROUP BY cCodCta "

Set rs = opCon.CargaRecordSet(Sql)
If Not RSVacio(rs) Then
    lsCuotaMinPend = IIf(IsNull(rs!MinCuota), 0, rs!MinCuota)
Else
    lsCuotaMinPend = 0
End If
rs.Close
Set rs = Nothing
'--****  Lista de Cuotas por Cubrir (Todas)
Sql = "SELECT cCodCta, cNroCuo, dFecVenc, Datediff(Day,dFecVenc,'" & lsFecha & "') AS DiasMora," _
    & "nCapital, nCapPag, nInteres, nIntComPag, nMora, nIntMorPag, nIntGra, nIntGraPag, cTipo " _
    & "FROM PlanDesPag " _
    & "WHERE cCodCta='" & psCodCta & "' AND  cEstado='P' AND cTipo = 'C' "

lnMora = 0
Me.lblCapitalCuoPend = "0.00"
Me.lblInteresCuoPend = "0.00"
Me.lblMoraCuoPend = "0.00"
Me.lblGastoCuoPend = "0.00"

lstCuotasPend.ListItems.Clear
Set rs = opCon.CargaRecordSet(Sql)
If Not RSVacio(rs) Then
    
    ' Si Gracia Prorateada
    If Me.lblTipoGraciaApr = "Prorrateada" Then
       Me.lblIntGraciaApr = Format(IIf(IsNull(rs!nIntGra), 0, rs!nIntGra), "#0.00")
    Else
       Me.lblIntGraciaApr = TotalIntGracia(psCodCta, opCon)
    End If

    Do While Not rs.EOF
        Set Item = lstCuotasPend.ListItems.Add(, , IIf(IsNull(rs!cNroCuo), 0, rs!cNroCuo))
        Item.SubItems(1) = IIf(IsNull(rs!dFecVenc), "", Format(rs!dFecVenc, "dd/mm/yyyy"))
        Item.SubItems(2) = Format(IIf(IsNull(rs!nCapital), 0, rs!nCapital) - IIf(IsNull(rs!nCapPag), 0, rs!nCapPag), "#0.00")
        Item.SubItems(3) = Format(IIf(IsNull(rs!nInteres), 0, rs!nInteres) - IIf(IsNull(rs!nintcompag), 0, rs!nintcompag) + _
                                  IIf(IsNull(rs!nIntGra), 0, rs!nIntGra) - IIf(IsNull(rs!nIntGraPag), 0, rs!nIntGraPag), "#0.00")
        Item.SubItems(4) = Format(IIf(IsNull(rs!nMora), 0, rs!nMora) - IIf(IsNull(rs!nIntMorPag), 0, rs!nIntMorPag), "#0.00")
        If lsEstadoCredito <> "R" Then
            lnMora = lnMora + Format(Item.SubItems(4), "#0.00")
            If Me.lbltipocuota = "CUOTA LIBRE" Then
               If Val(lblcuotasAprobado) = Val(rs!cNroCuo) And rs!cTipo = "C" Then
                  Item.SubItems(4) = Format(IIf(IsNull(vIntMorCal), 0, vIntMorCal), "#0.00")
               Else ' si no es la ultima
                  Item.SubItems(4) = Format(0, "#0")
               End If
            End If
            Item.SubItems(5) = IIf(IsNull(rs!DiasMora), 0, Format(rs!DiasMora, "#0"))
            'End If
        Else
            Item.SubItems(4) = 0
            Item.SubItems(5) = 0
        End If
        Item.SubItems(6) = IIf(IsNull(rs!cTipo), 0, rs!cTipo)
        
        '*********
        If rs!DiasMora > 0 Then  ' Cuotas Pendiente
           Me.lblCapitalCuoPend = Format(CDbl(Me.lblCapitalCuoPend) + CDbl(Item.SubItems(2)), "#0.00")
           Me.lblInteresCuoPend = Format(CDbl(Me.lblInteresCuoPend) + CDbl(Item.SubItems(3)), "#0.00")
           Me.lblMoraCuoPend = Format(CDbl(Me.lblMoraCuoPend) + CDbl(Item.SubItems(4)), "#0.00")
           '--****** Gastos Cuotas Pendientes en mora
           Sql = "SELECT cCodCta, SUM(nMonNeg-nMonPag) AS TotalGastos " _
                & " FROM PlanGastos " _
                & " WHERE cCodCta='" & psCodCta & "' AND cEstado='P' " _
                & " AND cAplicado='C' AND nNroCuo =" & rs!cNroCuo _
                & " GROUP BY cCodCta "
        
           Set rsGas = opCon.CargaRecordSet(Sql)
           If Not RSVacio(rsGas) Then
              If IsNull(rsGas!TotalGastos) Then
                Me.lblGastoCuoPend = Format(rsGas!TotalGastos, "#0.00")
              Else
                Me.lblGastoCuoPend = "0.00"
              End If
           Else
              Me.lblGastoCuoPend = "0.00"
           End If
           rsGas.Close
           Set rsGas = Nothing
           
        End If
        '*********
        rs.MoveNext
    Loop
    ' Agrega el Interes Pendiente de Cuota Libre
    If Me.lbltipocuota = "CUOTA LIBRE" Then
       Me.lblInteresCuoPend = Format(vIntPend, "#0.00")
    Else
       Me.lblInteresCuoPend = Format(CDbl(Me.lblInteresCuoPend), "#0.00")
    End If
End If
rs.Close
Set rs = Nothing

Me.lblTotalCuoPend = Format(CDbl(Me.lblCapitalCuoPend) + CDbl(Me.lblInteresCuoPend) + CDbl(Me.lblMoraCuoPend) + CDbl(Me.lblGastoCuoPend), "#0.00")

'--******  TOTAL A PAGAR - SEG CALENDARIO
Sql = "SELECT SUM(nCapital) - SUM(nCapPag) AS Capital, SUM(nInteres) - SUM(nIntComPag) AS InteresComp, " _
      & " SUM(nIntGra) - SUM(nIntGraPag) AS InteresGracia, SUM(nMora)- SUM(nIntMorPag) AS InteresMora " _
      & " From PlanDesPag WHERE cEstado='P' and cTipo = 'C' and cCodCta='" & psCodCta & "'"
      
Set rs = opCon.CargaRecordSet(Sql)
If Not RSVacio(rs) Then
    If Not IsNull(rs!Capital) Then
        Me.lblSaldoKCalend = IIf(IsNull(rs!Capital), "0.00", Format(rs!Capital, "#0.00"))
        Me.lblIntCompCalend = IIf(IsNull(rs!InteresComp), "0.00", Format(rs!InteresComp + IIf(IsNull(rs!InteresGracia), 0, rs!InteresGracia) + vIntPend, "#0.00"))
        'Me.lblIntMorCalend = IIf(IsNull(rs!InteresMora), "0.00", Format(rs!InteresMora, "#0.00"))
        If Me.lbltipocuota = "CUOTA LIBRE" Then
            Me.lblIntMorCalend = Format(IIf(IsNull(vIntMorCal), 0, vIntMorCal), "#0.00")
        Else
            Me.lblIntMorCalend = IIf(IsNull(rs!InteresMora), "0.00", Format(rs!InteresMora, "#0.00"))
        End If
        
        'A la Fecha
        Me.lblSaldoKFecha = IIf(IsNull(rs!Capital), "0.00", Format(rs!Capital, "#0.00"))
        'Me.lblIntMorFecha = IIf(IsNull(rs!InteresMora), "0.00", Format(rs!InteresMora, "#0.00"))
        If Me.lbltipocuota = "CUOTA LIBRE" Then
            Me.lblIntMorFecha = Format(IIf(IsNull(vIntMorCal), 0, vIntMorCal), "#0.00")
        Else
            Me.lblIntMorFecha = IIf(IsNull(rs!InteresMora), "0.00", Format(rs!InteresMora, "#0.00"))
        End If
        
    Else
         Me.lblSaldoKCalend = "0.00"
         Me.lblIntCompCalend = "0.00"
         Me.lblIntMorCalend = "0.00"
         'A la Fecha
         Me.lblSaldoKFecha = "0.00"
         Me.lblIntMorFecha = "0.00"
    End If
End If
rs.Close
Set rs = Nothing

'****** Gastos Totales Crédito de Cuotas Pendientes
Sql = "SELECT cCodCta, Sum(nMonNeg-nMonPag) AS TotalGastos " _
    & "FROM PlanGastos " _
    & "WHERE cCodCta='" & psCodCta & "' AND cEstado='P' " _
    & "AND cAplicado='C' GROUP BY cCodCta "

Set rs = opCon.CargaRecordSet(Sql)
If Not RSVacio(rs) Then
    Me.lblGastoCalend = IIf(IsNull(rs!TotalGastos), "0.00", Format(rs!TotalGastos, "#0.00"))
    Me.lblGastoFecha = IIf(IsNull(rs!TotalGastos), "0.00", Format(rs!TotalGastos, "#0.00"))
Else
    Me.lblGastoCalend = "0.00"
    Me.lblGastoFecha = "0.00"
End If
rs.Close
Set rs = Nothing

'--******* Int Suspenso
Sql = "SELECT (nIntCom + nIntMor+ nGasto+ nIntGraPag) AS IntSusp, cCodCtaAnt FROM Refinanc " _
        & " WHERE cCodCta='" & psCodCta & "'"

Dim lnIntSuspTemp As Single
Dim lnTipoCambioFijo As Single

lnTipoCambioFijo = ObtieneTipoCambioFijo()

Set rs = opCon.CargaRecordSet(Sql)
If Not RSVacio(rs) Then
   Do While Not rs.EOF
        'lnIntSuspTemp = lnIntSuspTemp & IIf(IsNull(RS!IntSusp), 0, IIf(Mid(psCodCta, 6, 1) = Mid(RS!cCodCtaAnt, 6, 1), RS!IntSusp, IIf(Mid(psCodCta, 6, 1) = "1", RS!IntSusp * lnTipoCambioFijo, RS!IntSusp / lnTipoCambioFijo)))
        If IsNull(rs!IntSusp) Then
            lnIntSuspTemp = lnIntSuspTemp + 0
        ElseIf Mid(psCodCta, 6, 1) = Mid(rs!ccodctaant, 6, 1) Then
            lnIntSuspTemp = lnIntSuspTemp + rs!IntSusp
        ElseIf Mid(psCodCta, 6, 1) = "1" Then
            lnIntSuspTemp = lnIntSuspTemp + rs!IntSusp * lnTipoCambioFijo
        Else
            lnIntSuspTemp = lnIntSuspTemp + rs!IntSusp / lnTipoCambioFijo
        End If
        
        
        'Me.lblIntSuspFecha = IIf(IsNull(RS!IntSusp), 0, RS!IntSusp)
        'lblKpend.Caption = IIf(IsNull(RS!nSaldoCap), "0.00", Format(RS!nSaldoCap, "#0.00"))
        'lblIntPend.Caption = IIf(IsNull(RS!TotalIntPend), "0.00", Format(RS!TotalIntPend, "#0.00"))
        'If lnNroRepro > 0 And CCur(lblKpend) = 0 Then
        '    lblIntPend = Me.lblTotalPagar
        'End If
      rs.MoveNext
   Loop
   Me.lblIntSusPendCalend = Format(lnIntSuspTemp, "#0.00")
   Me.lblIntSuspFecha = Format(lnIntSuspTemp, "#0.00")
Else
    Me.lblIntSusPendCalend = "0.00"
    Me.lblIntSuspFecha = "0.00"
End If
rs.Close
Set rs = Nothing

Me.lblTotalCalend = Format(CDbl(Me.lblSaldoKCalend) + CDbl(Me.lblIntCompCalend) + CDbl(Me.lblIntMorCalend) + CDbl(Me.lblGastoCalend) + CDbl(Me.lblIntSusPendCalend), "#0.00")

End Sub

Private Sub Limpiar()
lblCodCred.Enabled = True
lblanalista = ""
lblapoderado = ""
lblcapitalpagado = "0.00"
Me.lblIntCompFecha = "0.00"

Me.lblTotalFecha = "0.00"

lblCodCred.Text = Right(gsCodAge, 2)
Me.lblcondicion = ""
Me.lblcuotasAprobado = "0"
Me.lblcuotasSolicitud = "0"
Me.lblcuotasugerida = "0"

Me.lbldestino = ""
Me.lblEstado = ""
Me.lblfechaAprobado = ""


Me.lblfechaCancelacion = ""
Me.lblfechajudicial = ""
Me.lblfechaproxdesem = ""
Me.lblfechasugerida = ""
Me.lblfechavigencia = ""
Me.lblfechsolicitud = ""
Me.lblfuente = ""
'Me.lblGastoaPagar = "0.00"
Me.lblGastopagado = "0.00"
'Me.lblGastoPend = "0.00"
Me.lblGraciaAprobada = "0"
Me.lblGraciasugerida = "0"
Me.lblTipoGraciaApr = ""
Me.lblIntGraciaApr = "0"
Me.txtComenta = ""
Me.lblinstitucion = ""
Me.lblintcompPag = "0.00"
Me.lblIntMorPag = "0.00"
'Me.lblIntPend = "0.00"
'Me.lblKpend = "0.00"
Me.lblLinea = ""
Me.lblmetodoLiquidacion = ""
Me.lblmodular = ""
Me.lblMontoAprobado = "0.00"
Me.lblmontoCuotaAprobada = "0.00"
Me.lblmontoCuotaSugerida = "0.00"
Me.lblmontoDesembolsado = "0.00"
'Me.lblMontoPago = "0.00"
Me.lblMontoProxDesem = "0.00"
Me.lblmontosolicitud = "0.00"
Me.lblMontosugerido = "0.00"
'Me.lblMoraPend = "0.00"
Me.lblnota1 = ""
Me.chkCargoAuto.value = 0
Me.chkProtesto.value = 0
Me.chkRefinanciado.value = 0
Me.lblnota2 = ""
Me.lblperiodo = ""
Me.lblPlazoAprobado = "0.00"
Me.lblPlazoSolicitud = "0.00"
Me.lblPlazoSugerido = "0.00"
Me.lblProximoDesembolso = ""
Me.lblrechazo = ""
Me.lbltasainteres = "0.00"
Me.lbltipoCredito = ""
Me.lbltipocuota = ""
Me.lbltipoDesembolso = ""
Me.lbltotalDesembolso = "0.00"
'Me.lbltotalMora = "0.00"
'Me.lblTotalPagar = "0.00"
Me.lblSaldoKCalend = "0.00"
Me.lblSaldoKFecha = "0.00"
Me.lblIntCompCalend = "0.00"
Me.lblIntCompFecha = "0.00"
Me.lblIntMorCalend = "0.00"
Me.lblIntMorFecha = "0.00"
Me.lblIntSusPendCalend = "0.00"
Me.lblIntSuspFecha = "0.00"
Me.lblTotalCalend = "0.00"
Me.lblTotalFecha = "0.00"
Me.listaClientes.ListItems.Clear
Me.lstCuotasPend.ListItems.Clear
Me.ListaPagos.ListItems.Clear
Me.listaDesembolsos.ListItems.Clear
Me.ListaPagos.ListItems.Clear
Me.lstgarantias.ListItems.Clear
Me.lstRefinanciados.ListItems.Clear
ReDim MatPagos(0)
ReDim MatGastos(0)
ContCuotas = 0
ContGastos = 0
ReDim MatSaldos(0)
nSaldos = 0
End Sub
Sub CargaGastos(CodCta As String, Conexion As DConecta)
Dim RegGastos As New ADODB.Recordset
Dim SQL1 As String
    SQL1 = "SELECT cCodGas,cEstado,nNroCuo,(nMonNeg - nMonPag) AS Gasto FROM PlanGastos " & _
           " where cCodCta='" & CodCta & "' and cAplicado='C' and cEstado='P'" & _
           " Order By nNroCuo"
    ContGastos = 0
    Set RegGastos = Conexion.CargaRecordSet(SQL1)
    Do While Not RegGastos.EOF
        ContGastos = ContGastos + 1
        ReDim Preserve MatGastos(ContGastos)
        MatGastos(ContGastos - 1).NumCuota = RegGastos!nNroCuo
        MatGastos(ContGastos - 1).MontoGasto = Format(IIf(IsNull(RegGastos!Gasto), 0, RegGastos!Gasto), "#0.00")
        MatGastos(ContGastos - 1).Estado = RegGastos!cEstado
        MatGastos(ContGastos - 1).Gastopag = 0
        MatGastos(ContGastos - 1).Codgasto = RegGastos!cCodGas
        MatGastos(ContGastos - 1).Modificado = False
        RegGastos.MoveNext
    Loop
    RegGastos.Close
End Sub
Function GastoCuota(NumC As Integer) As Double
Dim i As Integer
Dim Acum As Double
    Acum = 0
    For i = 0 To ContGastos - 1
        If NumC = MatGastos(i).NumCuota Then
            Acum = Acum + MatGastos(i).MontoGasto
        End If
    Next i
    GastoCuota = Acum
End Function

Sub SaldoCapitalesCuota(ByVal CodCta As String, Conexion As DConecta)
Dim RCred As New ADODB.Recordset
Dim SQL1 As String
Dim i As Integer
Dim Monto As Double
    nSaldos = 1
    ReDim MatSaldos(nSaldos)
    SQL1 = "SELECT Credito.nCuotasApr,Credito.nMontoDesemb,Plandespag.nCapital FROM Credito Inner Join PlanDespag " & _
        " On Credito.cCodCta = Plandespag.cCodCta WHERE Plandespag.cTipo='C' and Credito.cCodCta='" & CodCta & "' ORDER BY Plandespag.cNroCuo ASC"
    Set RCred = Conexion.CargaRecordSet(SQL1)
    If Not RCred.BOF And Not RCred.EOF Then
        Monto = Format(IIf(IsNull(RCred!nMontoDesemb), 0, RCred!nMontoDesemb), "#0.00")
        MatSaldos(0) = IIf(IsNull(RCred!nMontoDesemb), 0, RCred!nMontoDesemb)
    End If
    Do While Not RCred.EOF
        Monto = Monto - Format(RCred!nCapital, "#0.00")
        nSaldos = nSaldos + 1
        ReDim Preserve MatSaldos(nSaldos)
        MatSaldos(nSaldos - 1) = Monto
        RCred.MoveNext
    Loop
    
    RCred.Close
    Set RCred = Nothing
End Sub

Sub CargaCuotas(ByVal CodCta As String, Conexion As DConecta)
Dim RegCuotas  As New ADODB.Recordset
Dim SQL1 As String
     ContCuotas = 0
    'Realiza Carga de Datos
        SQL1 = "SELECT * FROM PlanDesPag WHERE cCodCta='" & CodCta & "' and cTipo = 'C' and cEstado = 'P'" & _
               " ORDER BY cNroCuo"
        Call CargaGastos(CodCta, Conexion)
        Set RegCuotas = Conexion.CargaRecordSet(SQL1)
        Do While Not RegCuotas.EOF
            If RegCuotas!cEstado = "P" Then
                ContCuotas = ContCuotas + 1
                ReDim Preserve MatPagos(ContCuotas)
                MatPagos(ContCuotas - 1).NumCuota = RegCuotas!cNroCuo
                MatPagos(ContCuotas - 1).FecVenc = RegCuotas!dFecVenc
                MatPagos(ContCuotas - 1).Capital = Format(RegCuotas!nCapital - RegCuotas!nCapPag, "#0.00")
                MatPagos(ContCuotas - 1).Interes = Format(RegCuotas!nInteres - RegCuotas!nintcompag, "#0.00")
                MatPagos(ContCuotas - 1).MontoCuota = Format(MatPagos(ContCuotas - 1).Capital + MatPagos(ContCuotas - 1).Interes, "#0.00")
                MatPagos(ContCuotas - 1).Gasto = Format(GastoCuota(RegCuotas!cNroCuo), "#0.00")
                MatPagos(ContCuotas - 1).Mora = Format(IIf(IsNull(RegCuotas!nMora), 0, RegCuotas!nMora) - IIf(IsNull(RegCuotas!nIntMorPag), 0, RegCuotas!nIntMorPag), "#0.00")
'                If ContCuotas = 1 And MatPagos(ContCuotas - 1).Mora > 0 Then
'                    If Not CobrarMora(MatPagos(ContCuotas - 1).FecVenc, 3, gdFecSis) Then
'                        MatPagos(ContCuotas - 1).Mora = 0
'                    End If
'                End If
                MatPagos(ContCuotas - 1).Total = Format(MatPagos(ContCuotas - 1).MontoCuota + MatPagos(ContCuotas - 1).Gasto + MatPagos(ContCuotas - 1).Mora, "#0.00")
                MatPagos(ContCuotas - 1).CapPag = 0
                MatPagos(ContCuotas - 1).IntPag = Format(RegCuotas!nintcompag, "#0.00")
                MatPagos(ContCuotas - 1).MoraPag = 0
                MatPagos(ContCuotas - 1).Estado = RegCuotas!cEstado
                MatPagos(ContCuotas - 1).Modificado = False
            End If
            RegCuotas.MoveNext
        Loop
        RegCuotas.Close
        FecPagPend = MatPagos(0).FecVenc
        'Cargar la Ultima Cuota
        SQL1 = "SELECT * FROM PlanDesPag WHERE cCodCta='" & CodCta & "' and cTipo = 'C' and cEstado = 'G'" & _
               " ORDER BY cNroCuo DESC"
        Set RegCuotas = Conexion.CargaRecordSet(SQL1)
        If Not RegCuotas.BOF And Not RegCuotas.EOF Then
            FecUltPago = RegCuotas!dFecVenc
        End If
        RegCuotas.Close
        'Carga los saldos de capitales de las cuotas
        Call SaldoCapitalesCuota(CodCta, Conexion)
End Sub
Function TotalDeudaParcial(Conexion As DConecta) As Double
Dim SQL1 As String
Dim RC As New ADODB.Recordset
Dim Interes As Double
    If Tipodesemb = "P" And gdFecSis <= MatPagos(0).FecVenc Then
        SQL1 = "SELECT * FROM PlandesPag WHERE cCodCta='" & lsCodCred & "' and cTipo = 'D' and cEstado='G' ORDER BY cNroCuo ASC"
        Set RC = Conexion.CargaRecordSet(SQL1)
        Interes = 0
        Do While Not RC.EOF
            Interes = Interes + IntPerDias(TasaInt, gdFecSis - RC!dFecVenc, Periodo) * RC!nCapital
            RC.MoveNext
        Loop
        RC.Close
        If Interes <= MatPagos(0).Interes Then
            Interes = Format(Interes, "#0.00")
        Else
            Interes = 0
        End If
        
        TotalDeudaParcial = MatPagos(0).Total - MatPagos(0).Interes + Interes
        MatPagos(0).Interes = Interes
    Else
        TotalDeudaParcial = TotalDeuda
    End If
Set RC = Nothing
End Function

Function TotalDeuda() As Double
Dim Monto, IntAPagar, MontoIAPagar As Double

Dim i, NumDias As Integer
    TotalInteresaFecha = 0
    Monto = 0
    For i = 0 To ContCuotas - 1
        'si es una cuota adelantada al pago pendiente
        If MatPagos(i).FecVenc > FecPagPend Then
            'Verificar si es un pago adelantado
            If gdFecSis < MatPagos(i).FecVenc Then
                If gdFecSis > MatPagos(i - 1).FecVenc Then
                    NumDias = gdFecSis - MatPagos(i - 1).FecVenc
                    IntAPagar = IntPerDias(TasaInt, NumDias, Periodo)
                    MontoIAPagar = MatSaldos(MatPagos(i).NumCuota - 1) * IntAPagar
                    
                    'Si monto de interes a pagar es menor que lo que ya pago entonces
                    'ese dinero del cliente que sobra se da por perdido
                    If MontoIAPagar < MatPagos(i).IntPag Then
                        MontoIAPagar = 0#
                        TotalInteresaFecha = TotalInteresaFecha + MontoIAPagar
                    Else
                        MontoIAPagar = MontoIAPagar - MatPagos(i).IntPag
                        
                    End If
                    Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes + MontoIAPagar
                    MatPagos(i).Interes = MontoIAPagar
                    TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                Else
                    Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes
                    MatPagos(i).Interes = 0#
                    TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                End If
            Else
                Monto = Monto + MatPagos(i).Total
                TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
            End If
        Else
            'Preguntar si el Usuario esta al dia
            If gdFecSis < MatPagos(i).FecVenc Then
                NumDias = gdFecSis - FecUltPago
                If NumDias < 0 Then
                    MontoIAPagar = 0
                    'Si monto de interes a pagar es menor que lo que ya pago entonces
                    'ese dinero del cliente que sobra se da por perdido
                    If MontoIAPagar < MatPagos(i).IntPag Then
                        MontoIAPagar = 0#
                    Else
                        MontoIAPagar = MontoIAPagar - MatPagos(i).IntPag
                    End If
                    Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes + MontoIAPagar
                    MatPagos(i).Interes = MontoIAPagar
                    TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                Else
                    IntAPagar = IntPerDias(TasaInt, NumDias, Periodo)
                    MontoIAPagar = MatSaldos(MatPagos(i).NumCuota - 1) * IntAPagar
                    'Si monto de interes a pagar es menor que lo que ya pago entonces
                    'ese dinero del cliente que sobra se da por perdido
                    If MontoIAPagar < MatPagos(i).IntPag Then
                        MontoIAPagar = 0#
                    Else
                        MontoIAPagar = MontoIAPagar - MatPagos(i).IntPag
                    End If
                    Monto = Monto + MatPagos(i).Total - MatPagos(i).Interes + MontoIAPagar
                    MatPagos(i).Interes = MontoIAPagar
                    TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
                End If
            Else
                Monto = Monto + MatPagos(i).Total
                TotalInteresaFecha = TotalInteresaFecha + MatPagos(i).Interes
            End If
        End If
    Next i
    TotalDeuda = Monto
End Function
Private Sub Garantias(lsCodCred As String, opCon As DConecta)
Dim rsg As New ADODB.Recordset
Dim Sql As String
Dim Item As ListItem

Sql = "SELECT Garantias.cNumGarant, Garantias.cCodPers, " _
        & "Garantias.cTipoGarant, Garantias.cDesGarant, " _
        & "Garantias.cDocGarant, Garantias.cNumDoc, " _
        & "Garantias.cMoneda, Garantias.nMontoxGrav, " _
        & "GarantCred.cCodCta, GarantCred.nMontoGrava, " _
        & "Garantias.cEstado " _
        & "FROM Garantias INNER JOIN " _
        & "GarantCred ON " _
        & "Garantias.cNumGarant = GarantCred.cNumGarant " _
        & "WHERE (GarantCred.cCodCta = '" & lsCodCred & "') "

Set rsg = opCon.CargaRecordSet(Sql)

lstgarantias.ListItems.Clear
If Not RSVacio(rsg) Then
    Do While Not rsg.EOF
        Set Item = Me.lstgarantias.ListItems.Add(, , rsg!cNumGarant)
        Item.SubItems(1) = IIf((IsNull(rsg!cTipoGarant) Or Len(Trim(rsg!cTipoGarant)) = 0), "", Tablacod("24", rsg!cTipoGarant, True))
        Item.SubItems(2) = IIf(IsNull(rsg!cDesGarant), "", Trim(rsg!cDesGarant))
        Item.SubItems(3) = IIf((IsNull(rsg!cDocGarant) Or Len(Trim(rsg!cDocGarant)) = 0), "", Tablacod("49", rsg!cDocGarant, True))
        Item.SubItems(4) = IIf(IsNull(rsg!cNumDoc), "", Trim(rsg!cNumDoc))
        Item.SubItems(5) = IIf((IsNull(rsg!cMoneda) Or Len(Trim(rsg!cMoneda)) = 0), "", Tablacod("03", rsg!cMoneda, True))
        Item.SubItems(6) = IIf(IsNull(rsg!nMontoGrava), "0.00", Format(rsg!nMontoGrava, "#0.00"))
        Item.SubItems(7) = IIf(IsNull(rsg!nMontoxGrav), "0.00", Format(rsg!nMontoxGrav, "#0.00"))
        Item.SubItems(8) = IIf(IsNull(rsg!cEstado), "", Trim(rsg!cEstado))
        rsg.MoveNext
    Loop
End If
rsg.Close
Set rsg = Nothing
End Sub

Private Sub listaClientes_DblClick()
If Me.listaClientes.ListItems.Count > 0 Then
   ' frmClientes.Inicia Me, "", False, True, Me.listaClientes.SelectedItem.SubItems(2)
End If
End Sub

Private Sub ListaClientes_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Me.listaClientes.ListItems.Count > 0 Then
'        frmClientes.Inicia Me, "", False, True, Me.listaClientes.SelectedItem.SubItems(2)
'    End If
'End If
End Sub

Private Function TotalIntGracia(pCodCta As String, pConex As DConecta) As Double
Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = "SELECT Sum(nIntGra) AS IntGracia " & _
    " FROM PlanDesPag WHERE cTipo ='C' AND cCodcta='" & pCodCta & "'"
Set rs = pConex.CargaRecordSet(Sql)
If Not RSVacio(rs) Then
   TotalIntGracia = rs!IntGracia
Else
   TotalIntGracia = 0
End If
rs.Close
Set rs = Nothing

End Function


Private Function ObtieneTipoCambioFijo() As Single
Dim Sql As String
Dim rs As New ADODB.Recordset

  ObtieneTipoCambioFijo = gnTipCambio
End Function


Private Sub CargaRefinanciados(psCodCred As String, opCon As DConecta)
Dim Sql As String
Dim rsR As New ADODB.Recordset
'Dim GastoCuota As String
'Dim TotalPagado As String
Dim ItemR As ListItem  'item para desembolso
'Dim ItemC As ListItem  'item para Cuotas pagadas
'Dim lnMora As Long
Sql = "SELECT cCodCtaAnt,nMontoRef,(nIntCom + nIntMor+ nGasto+ nIntGraPag) AS IntSusp " & _
      " FROM Refinanc " & _
      " WHERE cCodCta='" & psCodCred & "'"
Set rsR = opCon.CargaRecordSet(Sql)

If RSVacio(rsR) Then
    lstRefinanciados.ListItems.Clear
Else
    lstRefinanciados.ListItems.Clear
    Do While Not rsR.EOF
        Set ItemR = lstRefinanciados.ListItems.Add(, , rsR!ccodctaant)
        ItemR.SubItems(1) = Format(rsR!nMontoRef, "#0.00")
        ItemR.SubItems(2) = Format(rsR!IntSusp, "#0.00")
        rsR.MoveNext
    Loop
End If
rsR.Close
Set rsR = Nothing

End Sub

Private Function AbrevProd(lsProducto As String, Optional lbAbreviatura As Boolean = True) As String
    Dim Tipo As String
    Dim Sql As String
    Dim SQL2 As String
    Dim Abrev As String
    Dim lsCodTab As String
    Dim rs As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion 'Remota "07"
    Tipo = Mid(Trim(lsProducto), 1, 1)
    Sql = "SELECT cCodTab,cNomTab,cAbrev FROM " & gcCentralCom & "TablaCod " & "WHERE cCodTab LIKE '6_' AND cValor='" & Tipo & "'"
   
    Set rs = oCon.CargaRecordSet(Sql)
    If Not RSVacio(rs) Then
        If lbAbreviatura Then
            Abrev = Trim(IIf(IsNull(rs!cAbrev), "", rs!cAbrev))
        Else
            Abrev = Trim(IIf(IsNull(rs!cNomtab), "", rs!cNomtab))
        End If
        lsCodTab = Trim(rs!cCodtab)
    End If
    rs.Close
    Set rs = Nothing
    lsCodTab = lsCodTab + "__"

    Sql = "SELECT cCodTab,cNomTab,cAbrev FROM " & gcCentralCom & "TABLACOD WHERE CCODTAB LIKE '" & Trim(lsCodTab) & "' AND CVALOR='" & Trim(Mid(lsProducto, 2, 2)) & "'"
    Set rs = oCon.CargaRecordSet(Sql)
    If Not RSVacio(rs) Then
        If lbAbreviatura Then
            Abrev = Abrev + Trim(IIf(IsNull(rs!cAbrev), "", rs!cAbrev))
        Else
            Abrev = Abrev + Space(1) + Trim(IIf(IsNull(rs!cNomtab), "", rs!cNomtab))
        End If
    End If
    rs.Close
    Set rs = Nothing
    AbrevProd = Abrev
End Function

Private Function Tablacod(psCodTab As String, psValor As String, Optional pbNombre As Boolean = True) As String
    Dim Sql As String
    Dim rsA As New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion 'Remota "07"
    
    Sql = "Select * from " & gcCentralCom & "Tablacod where cCodtab like '" & Trim(psCodTab) & "__'  and cValor='" & Trim(psValor) & "'"

    Set rsA = oCon.CargaRecordSet(Sql)
    If RSVacio(rsA) Then
    Else
        If pbNombre Then
            Tablacod = Trim(rsA!cNomtab)
        Else
            Tablacod = Trim(IIf(IsNull(rsA!cAbrev), "", rsA!cAbrev))
        End If
    End If
    rsA.Close
    Set rsA = Nothing
End Function

Private Function LineaCredito(psCodLinea As String) As String
    Dim rsLin As New ADODB.Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion 'Remota "07"
    
    Sql = "Select * from " & gcCentralCom & "lineacredito where cCodLinCred='" & Trim(psCodLinea) & "'"
    Set rsLin = oCon.CargaRecordSet(Sql)
    If RSVacio(rsLin) Then
        LineaCredito = ""
    Else
        LineaCredito = Trim(rsLin!cDesLinCred)
    End If
    rsLin.Close
    Set rsLin = Nothing
End Function

Private Function InteresReal(ByVal Ti As Double, ByVal Periodo As Integer) As Double
    InteresReal = ((1 + Ti) ^ (Periodo / 30)) - 1
End Function

Private Function Usuario(psCodAnalista As String, opCon As DConecta, Optional psGrupo As String = "0003") As String
    Dim rsUsu As New ADODB.Recordset
    Dim Sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion 'Remota "07"

    Sql = "SELECT DISTINCT U.cCodUsu, U.cNomUsu, GU.cCodGrp  FROM " & gcCentralCom & "Usuario U INNER JOIN " & gcCentralCom & "GrupoUsu GU ON U.cCodUsu = GU.cCodUsu  WHERE GU.cCodGrp = '" & Trim(psGrupo) & "' AND U.cCodUsu='" & Trim(psCodAnalista) & "'"
  
    Set rsUsu = oCon.CargaRecordSet(Sql)
    If RSVacio(rsUsu) Then
        Usuario = ""
    Else
        Usuario = Trim(rsUsu!cNomUsu)
    End If
    rsUsu.Close
    Set rsUsu = Nothing
End Function

Private Function fgCalculaDeudaTotalaFechaGMYC(ByVal pCodCta As String, poCon As DConecta) As Double
Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim lnSaldoCap As Double
Dim lnMoraTotal As Double
Dim lnGastoTotal As Double
Dim lnTotalDeuda As Double

    lsSQL = "SELECT nSaldoCap, nIntMorCal FROM Credito WHERE cCodCta = '" & pCodCta & "'"
    Set R = poCon.CargaRecordSet(lsSQL)
    
    If Not R.BOF And Not R.EOF Then
        lnSaldoCap = R!nSaldoCap
        lnMoraTotal = IIf(IsNull(R!nIntMorCal), 0, R!nIntMorCal)
    End If
    R.Close
    Set R = Nothing
    lnGastoTotal = fgCalculaGastoTotal(pCodCta, poCon)
    fgCalculaDeudaTotalaFechaGMYC = CDbl(Format(lnSaldoCap + fgCalculaInteresaFechaGMYC(pCodCta, poCon) + lnMoraTotal + lnGastoTotal, "#0.00"))

End Function

'Private Function fgCalculaInteresaFechaGMYC(ByVal pCodCta As String, poCon As DConecta) As Double
'Dim R As New ADODB.Recordset
'Dim lsSQL As String
'
'
'Dim cMetLiqAnt As String
'Dim NumDias As Integer
'Dim lnNroProxCuota As Integer
'Dim lsMetLiq As String
'Dim ldFecVig As Date
'Dim ldFecUltCuotaPag As Date
'Dim lnIntPagCuotaPend As Double
'Dim ldFecUltPago As Date
'Dim lsMetLiqAnt As String
'Dim lnTasaInt As Double
'Dim lnSaldoCap As Double
'
'
'    lsSQL = " SELECT nNroProxCuota, cMetLiquid, dFecUltPago, dFecUltDesemb, dFecVig, nSaldoCap, nTasaInt, nIntMorCal FROM Credito Where cCodCta = '" & pCodCta & "'"
'    R.Open lsSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
'
'    lnNroProxCuota = R!nNroProxCuota
'    lsMetLiq = R!cMetLiquid
'    ldFecVig = CDate(Format(R!dFecVig, "dd/mm/yyyy"))
'    lnSaldoCap = R!nSaldoCap
'    lnTasaInt = R!nTasaInt
'    If IsNull(R!dFecUltPago) Then
'        ldFecUltPago = RegCli!dfecUltDesemb
'    Else
'        If R!dFecUltPago >= R!dfecUltDesemb Then
'            ldFecUltPago = R!dFecUltPago
'        Else
'            ldFecUltPago = R!dfecUltDesemb
'        End If
'    End If
'    R.Close
'
'    lsMetLiqAnt = fgMetLiqUltPago(pCodCta, pConexion)
'    lnIntPagCuotaPend = 0
'    If lsMetLiqAnt <> lsMetLiq Then
'        If lnNroProxCuota > 1 Then
'            sSQL = "SELECT dFecVenc,cNroCuo, nIntComPag FROM PlandesPag WHERE cTipo='C' And cCodCta = '" & pCodCta & "' And (cNroCuo = " & Trim(Str(lnNroProxCuota - 1)) & " or cNrocuo = " & Trim(Str(lnNroProxCuota)) & ") ORDER BY cNrocuo"
'            R.Open sSQL, pConexion, adOpenStatic, adLockReadOnly, adCmdText
'            Do While Not R.EOF
'                If R!cNroCuo = (lnNroProxCuota - 1) Then
'                    ldFecUltCuotaPag = R!dFecVenc
'                End If
'                If R!cNroCuo = NroProxCuota Then
'                    lnIntPagCuotaPend = R!nintcompag
'                End If
'                R.MoveNext
'            Loop
'            R.Close
'        Else
'            lnIntPagCuotaPend = 0 'Arturo
'            ldFecUltCuotaPag = ldFecVig
'        End If
'        NumDias = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(ldFecUltCuotaPag, "dd/mm/yyyy"))
'        fgCalculaInteresaFechaGMYC = (IntPerDias(lnTasaInt, NumDias, 30) * lnSaldoCap) - lnIntPagCuotaPend
'    Else
'        NumDias = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(ldFecUltPago, "dd/mm/yyyy"))
'        fgCalculaInteresaFechaGMYC = (IntPerDias(lnTasaInt, NumDias, 30) * lnSaldoCap)
'    End If
'    Set R = Nothing
'    'Adicionar los intereses generados por una Reprogramacion
'
'    fgCalculaInteresaFechaGMYC = CDbl(Format(fgCalculaInteresaFechaGMYC + fgCalculaInteresReprogramados(pCodCta, pConexion), "#0.00"))
'
'End Function


 Private Function fgCalculaDeudaTotalaFecha(ByVal pCodCta As String, poCon As DConecta) As Double
Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim lnSaldoCap As Double
Dim lnMoraTotal As Double
Dim lnGastoTotal As Double
Dim lnTotalDeuda As Double
  
    lsSQL = "SELECT nSaldoCap, nIntMorCal FROM Credito WHERE cCodCta = '" & pCodCta & "'"
    
    Set R = poCon.CargaRecordSet(lsSQL)
    If Not R.BOF And Not R.EOF Then
        lnSaldoCap = R!nSaldoCap
        lnMoraTotal = IIf(IsNull(R!nIntMorCal), 0, R!nIntMorCal)
    End If
    R.Close
    Set R = Nothing
    lnGastoTotal = fgCalculaGastoTotal(pCodCta, poCon)
    fgCalculaDeudaTotalaFecha = CDbl(Format(lnSaldoCap + fgCalculaInteresaFecha(pCodCta, poCon) + lnMoraTotal + fgCalculaInteresReprogramados(pCodCta, poCon) + lnGastoTotal, "#0.00"))
    
    

End Function


Private Function fgCalculaDeudaTotalaFechaCuotaLibre(ByVal pCodCta As String, poCon As DConecta) As Double
Dim R As New ADODB.Recordset
Dim lsSQL As String
Dim lnIntPend As Double
Dim lnSaldoCap As Double
Dim lnMoraTotal As Double
Dim lnGastoTotal As Double
Dim lnInteresFecha As Double

    lsSQL = " SELECT nIntPend, nSaldoCap, nIntMorCal FROM Credito Where cCodCta = '" & pCodCta & "'"
    Set R = poCon.CargaRecordSet(lsSQL)
        lnIntPend = IIf(IsNull(R!nIntPend), 0, R!nIntPend)
        lnSaldoCap = IIf(IsNull(R!nSaldoCap), 0, R!nSaldoCap)
        lnMoraTotal = IIf(IsNull(R!nIntMorCal), 0, R!nIntMorCal)
    R.Close

    lnGastoTotal = fgCalculaGastoTotal(pCodCta, poCon)
    lnInteresFecha = fgCalculaInteresaFechaCuotaLibre(pCodCta, poCon)

    fgCalculaDeudaTotalaFechaCuotaLibre = CDbl(Format(lnSaldoCap + lnIntPend + lnInteresFecha + lnMoraTotal + lnGastoTotal, "#0.00"))

    Set R = Nothing
    
End Function

Private Function fgCalculaInteresaFecha(ByVal pCodCta As String, ByVal poCon As DConecta) As Double

Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim ldFecUltDes As Date
Dim lnNroProxCuota As Integer
Dim lnDiasTrans As Integer
'Dim nInteres As Double
Dim lnGraciaApr As Integer
Dim ldFecVenGracia As Date
Dim lnMontoApr As Double
Dim lnTasa As Double
Dim lnInteresFecha As Double
Dim lnPeriodo As Integer
Dim lnSaldoCap As Double
Dim ldFecVenCuota As Date
Dim lnIntPagFecha As Double
Dim lbCreditoAdelantado As Boolean
Dim lnIntTemp As Double
Dim lnCuotaCorresp As Integer
Dim ldFecCuotaCorresp As Date


    'Obteniendo fecha de Desmbolso y Nro de Proxima Cuota
    lsSQL = "SELECT nTasaInt, nMontoApr, nGraciaApr, nNroProxCuota, dFecUltDesemb FROM Credito WHERE cCodCta = '" & pCodCta & "'"
    Set R = poCon.CargaRecordSet(lsSQL)
        If Not R.BOF And Not R.EOF Then
            lnMontoApr = R!nMontoApr
            lnNroProxCuota = R!nNroProxCuota
            ldFecUltDes = Format(R!dfecUltDesemb, "dd/mm/yyyy")
            lnGraciaApr = R!nGraciaApr
            If lnGraciaApr > 0 Then
                ldFecVenGracia = ldFecUltDes + lnGraciaApr
            End If
            lnTasa = R!nTasaInt
        End If
    R.Close
    
    lbCreditoAdelantado = False
    lnIntPagFecha = 0
    lnInteresFecha = 0
    lnDiasTrans = gdFecSis - ldFecUltDes
    'Si es la primera cuota
    If lnNroProxCuota <= 1 And lnGraciaApr > 0 And lnDiasTrans < lnGraciaApr Then
        'Calculo de Dias Transcurridos y Si todavia no vence la gracia
         lnDiasTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(ldFecUltDes, "dd/mm/yyyy"))
         lnInteresFecha = IntPerDias(lnTasa, lnDiasTrans, 30) * lnMontoApr
    Else
        'Halla intereses de Gracia Ganados
        lsSQL = "SELECT Sum(nIntGra - nIntGraPag) AS IntTot FROM Plandespag WHERE cCodCta = '" & pCodCta & "' And cTipo='C'"
        Set R = poCon.CargaRecordSet(lsSQL)
            lnInteresFecha = lnInteresFecha + CDbl(Format(IIf(IsNull(R!IntTot), 0, R!IntTot), "#0.00"))
        R.Close
        
        lnSaldoCap = lnMontoApr
        If lnGraciaApr > 0 Then
            ldFecVenCuota = ldFecVenGracia
        Else
            ldFecVenCuota = ldFecUltDes
        End If
        
        lnCuotaCorresp = -1
        lsSQL = "SELECT * from Plandespag WHERE cCodCta = '" & pCodCta & "' And cTipo='C'"
        Set R = poCon.CargaRecordSet(lsSQL)
            Do While Not R.EOF
                If gdFecSis >= R!dFecVenc Then
                    ldFecVenCuota = Format(R!dFecVenc, "dd/mm/yyyy")
                    lnSaldoCap = lnSaldoCap - R!nCapital
                    'Capital mayor a cero para no tocar reprogramados ni intereses de gracia  a la primmera cuota o a la final
                    If R!cEstado = "P" And R!nCapital > 0 Then
                        lnInteresFecha = lnInteresFecha + (R!nInteres - R!nintcompag)
                    End If
                Else
                    lnCuotaCorresp = R!cNroCuo
                    If R!cEstado = "G" Then
                        lbCreditoAdelantado = True
                    Else
                        lbCreditoAdelantado = False
                    End If
                    ldFecCuotaCorresp = CDate(Format(R!dFecVenc, "dd/mm/yyyy"))
                    lnIntPagFecha = R!nintcompag
                    Exit Do
                End If
                R.MoveNext
            Loop
        R.Close
        If lnCuotaCorresp <> -1 And lbCreditoAdelantado = False Then
            lnDiasTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(ldFecVenCuota, "dd/mm/yyyy"))
            lnPeriodo = ldFecCuotaCorresp - ldFecVenCuota
            lnIntTemp = (IntPerDias(lnTasa, lnDiasTrans, 30) * lnSaldoCap) - lnIntPagFecha
            If lnIntTemp > 0 Then
               lnInteresFecha = lnInteresFecha + lnIntTemp
            End If
        End If
    End If
Set R = Nothing

fgCalculaInteresaFecha = CDbl(Format(lnInteresFecha, "#0.00"))
End Function


'** Halla los Intereses Reprogramados    ******
'**********************************************
Private Function fgCalculaInteresReprogramados(ByVal pCodCta As String, ByVal poCon As DConecta) As Double

Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim lnCuotaFinal As Integer

   lsSQL = "SELECT nCuotasApr FROM Credito Where cCodCta = '" & pCodCta & "'"
   Set R = poCon.CargaRecordSet(lsSQL)
   lnCuotaFinal = R!nCuotasApr
   R.Close
   fgCalculaInteresReprogramados = 0
   lsSQL = "Select * from PlandesPag Where cCodCta = '" & pCodCta & "' And cTipo='C' And cNroCuo > " & Str(Trim(lnCuotaFinal))
   Set R = poCon.CargaRecordSet(lsSQL)
      Do While Not R.EOF
         fgCalculaInteresReprogramados = fgCalculaInteresReprogramados + CDbl(Format(R!nInteres - R!nintcompag, "#0.00"))
         R.MoveNext
      Loop
   R.Close
   Set R = Nothing
End Function

Private Function fgCalculaInteresaFechaCuotaLibre(ByVal pCodCta As String, ByVal pConexion As DConecta)

Dim R As New ADODB.Recordset
Dim lsSQL As String
Dim lnTasaInt As Double
Dim lnSaldoCapital As Double
Dim ldFecUltPago As Date
Dim ldFecUltCuoCal As Date
Dim lnDias As Integer
Dim lnInteresFecha As Double

    lsSQL = " SELECT nTasaInt, cEstado, dFecUltDesemb,dFecUltPago, dAsignacion, nSaldoCap  FROM Credito Where cCodCta = '" & pCodCta & "'"
    Set R = pConexion.CargaRecordSet(lsSQL)
        lnTasaInt = R!nTasaInt
        lnSaldoCapital = R!nSaldoCap
    If R!cEstado = "F" Then
       If IsNull(R!dFecUltPago) Then
         ldFecUltPago = R!dfecUltDesemb
       Else
         ldFecUltPago = IIf(R!dfecUltDesemb > R!dFecUltPago, R!dfecUltDesemb, R!dFecUltPago)
       End If
    Else
        ldFecUltPago = R!dAsignacion
    End If
    R.Close
    
    lsSQL = "SELECT dFecVenc FROM Plandespag WHERE cCodCta = '" & pCodCta & "' AND cTipo ='C' ORDER BY dFecVenc DESC "
    Set R = pConexion.CargaRecordSet(lsSQL)
    If Not R.BOF And Not R.EOF Then
        ldFecUltCuoCal = R!dFecVenc
    Else
        ldFecUltCuoCal = gdFecSis
    End If
    R.Close
    Set R = Nothing
    If gdFecSis > ldFecUltCuoCal Then
        If ldFecUltPago < ldFecUltCuoCal Then
            lnDias = ldFecUltCuoCal - CDate(Format(ldFecUltPago, "dd/mm/yyyy"))
            lnInteresFecha = 0 'InteresReal(lnTasaInt / 100, Dias) * lnSaldoCapital
        Else
            lnInteresFecha = 0
        End If
    Else
        lnDias = gdFecSis - CDate(Format(ldFecUltPago, "dd/mm/yyyy"))
        lnInteresFecha = InteresReal(lnTasaInt / 100, lnDias) * lnSaldoCapital
    End If
    lnInteresFecha = CDbl(Format(lnInteresFecha, "#0.00"))
    
    fgCalculaInteresaFechaCuotaLibre = lnInteresFecha
    
End Function

Private Function MontoGastos(psCuota As String, psAplicado As String, psCodCred As String, dbConexion As DConecta) As String
Dim Sql As String
Dim rsg As New ADODB.Recordset

Sql = "SELECT SUM(PlanGastos.nMonPag) AS SumaGastos FROM PlanGastos INNER JOIN PlanDesPag ON PlanGastos.cCodCta = PlanDesPag.cCodCta AND PlanGastos.cAplicado = PlanDesPag.cTipo AND PlanGastos.nNroCuo = PlanDesPag.cNroCuo WHERE PlanGastos.nNroCuo =" & Trim(psCuota) & "  AND PlanGastos.cAplicado = '" & Trim(psAplicado) & "' AND PlanGastos.cEstado = 'G' AND PlanGastos.cCodCta = '" & Trim(psCodCred) & "' "
Set rsg = dbConexion.CargaRecordSet(Sql)
If RSVacio(rsg) Then
    MontoGastos = Format(0, "#0.00")
Else
    MontoGastos = Format(IIf(IsNull(rsg!SumaGastos), 0, rsg!SumaGastos), "#0.00")
End If
rsg.Close
Set rsg = Nothing
End Function


Private Function fgCalculaGastoTotal(ByVal pCodCta As String, opCon As DConecta) As Double
Dim lsSQL As String
Dim R As New ADODB.Recordset
Dim lnGastoTotal As Double
    
    lsSQL = "SELECT SUM(nMonNeg - nMonPag) AS Gasto FROM PlanGastos WHERE cCodCta='" & pCodCta & "' and cAplicado='C' and cEstado='P'"
    Set R = opCon.CargaRecordSet(lsSQL)
    If Not R.BOF And Not R.EOF Then
        lnGastoTotal = IIf(IsNull(R!Gasto), 0, R!Gasto)
    Else
        lnGastoTotal = 0
    End If
    R.Close
    Set R = Nothing
fgCalculaGastoTotal = lnGastoTotal
End Function

Private Function IntPerDias(ByVal inter As Double, ByVal DiasTrans As Integer, ByVal Periodo As Double) As Double
    IntPerDias = ((1 + inter / 100) ^ (DiasTrans / Periodo)) - 1
End Function

Private Function CadDerecha(psCadena As String, lsTam As Integer) As String
    CadDerecha = Format(psCadena, "!" & String(lsTam, "@"))
End Function

Public Function fgCalculaInteresaFechaGMYC(ByVal pCodCta As String, pConexion As DConecta) As Double
Dim R As New ADODB.Recordset
Dim lsSQL As String


Dim cMetLiqAnt As String
Dim NumDias As Integer
Dim lnNroProxCuota As Integer
Dim lsMetLiq As String
Dim ldFecVig As Date
Dim ldFecUltCuotaPag As Date
Dim lnIntPagCuotaPend As Double
Dim ldFecUltPago As Date
Dim lsMetLiqAnt As String
Dim lnTasaInt As Double
Dim lnSaldoCap As Double
Dim sSQL As String
    
    lsSQL = " SELECT nNroProxCuota, cMetLiquid, dFecUltPago, dFecUltDesemb, dFecVig, " & _
            " nSaldoCap, nTasaInt, nIntMorCal " & _
            " FROM Credito Where cCodCta = '" & pCodCta & "'"
    Set R = pConexion.CargaRecordSet(lsSQL)
    
    lnNroProxCuota = R!nNroProxCuota
    lsMetLiq = R!cMetLiquid
    ldFecVig = CDate(Format(R!dFecVig, "dd/mm/yyyy"))
    lnSaldoCap = R!nSaldoCap
    lnTasaInt = R!nTasaInt
    If IsNull(R!dFecUltPago) Then
        ldFecUltPago = R!dfecUltDesemb
    Else
        If R!dFecUltPago >= R!dfecUltDesemb Then
            ldFecUltPago = R!dFecUltPago
        Else
            ldFecUltPago = R!dfecUltDesemb
        End If
    End If
    R.Close
    
    'lsMetLiqAnt = fgMetLiqUltPago(pCodCta, pConexion)
    lnIntPagCuotaPend = 0
    If lsMetLiqAnt <> lsMetLiq Then
        If lnNroProxCuota > 1 Then
            sSQL = "SELECT dFecVenc,cNroCuo, nIntComPag FROM PlandesPag " & _
                   "WHERE cTipo='C' And cCodCta = '" & pCodCta & "' And (cNroCuo = " & Trim(Str(lnNroProxCuota - 1)) & " or cNrocuo = " & Trim(Str(lnNroProxCuota)) & ") ORDER BY cNrocuo"
            Set R = pConexion.CargaRecordSet(sSQL)
            Do While Not R.EOF
                If R!cNroCuo = (lnNroProxCuota - 1) Then
                    ldFecUltCuotaPag = R!dFecVenc
                End If
                'If R!cNroCuo = NroProxCuota Then
                '    lnIntPagCuotaPend = R!nintcompag
                'End If '
                R.MoveNext
            Loop
            R.Close
        Else
            lnIntPagCuotaPend = 0 'Arturo
            ldFecUltCuotaPag = ldFecVig
        End If
        NumDias = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(ldFecUltCuotaPag, "dd/mm/yyyy"))
        fgCalculaInteresaFechaGMYC = (IntPerDias(lnTasaInt, NumDias, 30) * lnSaldoCap) - lnIntPagCuotaPend
    Else
        NumDias = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(ldFecUltPago, "dd/mm/yyyy"))
        fgCalculaInteresaFechaGMYC = (IntPerDias(lnTasaInt, NumDias, 30) * lnSaldoCap)
    End If
    Set R = Nothing
    'Adicionar los intereses generados por una Reprogramacion
    
    fgCalculaInteresaFechaGMYC = CDbl(Format(fgCalculaInteresaFechaGMYC + fgCalculaInteresReprogramados(pCodCta, pConexion), "#0.00"))

End Function


